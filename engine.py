"""Core calculations for Company X slow-stock review case study."""
from __future__ import annotations
import io, re, math
import pandas as pd
import numpy as np

AGE_COLS = ['Qty 0-30 days','Qty 31-60 days','Qty 61-90 days','Qty 91-180 days','Qty 180+ days']
AGED_COLS = ['Qty 91-180 days','Qty 180+ days']
REQUIRED = {
 'Item_Master': ['SKU','Item Name','Category','Master Pack Qty','Unit Cost (Rs)','Alternate Item (name)','Alternate Flag','Status'],
 'Stock_Aging': ['SKU'] + AGE_COLS,
 'Sales_Lines': ['Month','CustomerID','CustomerName','Industry','SKU','Qty','Value (Rs)'],
 'Customers': ['CustomerID','CustomerName','Industry','Area']
}

def clean_sku(x):
    if pd.isna(x): return ''
    return str(x).strip().upper()

def norm_text(x):
    if pd.isna(x): return ''
    return re.sub(r'\s+', ' ', str(x).strip()).casefold()

def _read_uploaded(uploaded, sheet_aliases):
    """Read an xlsx workbook or separate xlsx/csv files into standard sheet names."""
    data = {}
    if isinstance(uploaded, dict):
        for name, obj in uploaded.items():
            key = sheet_aliases.get(str(name).lower().replace('.xlsx','').replace('.csv',''))
            if key:
                if hasattr(obj, 'getvalue'): raw = obj.getvalue()
                elif isinstance(obj, bytes): raw = obj
                else: raw = obj
                if str(name).lower().endswith('.csv'):
                    data[key] = pd.read_csv(io.BytesIO(raw) if isinstance(raw, bytes) else raw)
                else:
                    data[key] = pd.read_excel(io.BytesIO(raw) if isinstance(raw, bytes) else raw)
        return data
    raw = uploaded.getvalue() if hasattr(uploaded, 'getvalue') else uploaded
    xls = pd.ExcelFile(io.BytesIO(raw) if isinstance(raw, bytes) else raw)
    for sh in xls.sheet_names:
        key = sheet_aliases.get(sh.lower().strip())
        if key: data[key] = pd.read_excel(xls, sheet_name=sh)
    return data

def load_files(files):
    aliases = {
      'item_master':'Item_Master','1_item_master':'Item_Master',
      'stock_aging':'Stock_Aging','2_stock_aging':'Stock_Aging',
      'sales_lines':'Sales_Lines','3_sales_lines':'Sales_Lines',
      'customers':'Customers','4_customers':'Customers',
      'readme':'','0_readme':'','sheet1':''
    }
    frames = {}
    iterable = files.items() if isinstance(files, dict) else [(getattr(f,'name',f'upload_{i}.xlsx'),f) for i,f in enumerate(files)]
    for name, file in iterable:
        low = str(name).lower()
        raw = file.getvalue() if hasattr(file, 'getvalue') else file
        if low.endswith('.csv'):
            key = next((standard for alias,standard in aliases.items() if standard and (low == alias or low.startswith(alias) or alias in low)),None)
            if key: frames[key] = pd.read_csv(io.BytesIO(raw) if isinstance(raw,bytes) else raw)
            continue
        try:
            xf = pd.ExcelFile(io.BytesIO(raw) if isinstance(raw,bytes) else raw)
            found_sheet = False
            for sh in xf.sheet_names:
                key = aliases.get(sh.lower().strip())
                if key:
                    frames[key] = pd.read_excel(xf, sheet_name=sh); found_sheet=True
            if found_sheet: continue
            key = next((standard for alias,standard in aliases.items() if standard and (low == alias or low.startswith(alias) or alias in low)),None)
            if key and xf.sheet_names:
                frames[key] = pd.read_excel(xf, sheet_name=xf.sheet_names[0])
        except Exception:
            key = next((standard for alias,standard in aliases.items() if standard and (low == alias or low.startswith(alias) or alias in low)),None)
            if key: raise ValueError(f'Could not read {name} as an Excel or CSV file.')
    missing = [k for k in REQUIRED if k not in frames]
    if missing: raise ValueError('Missing required data sheet(s): ' + ', '.join(missing) + '. Upload the combined workbook or all four files.')
    return frames

def analyse(item_master, stock_aging, sales_lines, customers):
    """Return item results, customer targets, diagnostics and exportable detail tables."""
    issues=[]
    master=item_master.copy(); stock=stock_aging.copy(); sales=sales_lines.copy(); cust=customers.copy()
    # Validate required columns before calculations.
    for name, df in [('Item_Master',master),('Stock_Aging',stock),('Sales_Lines',sales),('Customers',cust)]:
        missing=[c for c in REQUIRED[name] if c not in df.columns]
        if missing: raise ValueError(f'{name} is missing columns: {", ".join(missing)}')
        df.columns=[str(c).strip() for c in df.columns]
    # Normalize IDs and text for matching.
    for df in [master,stock,sales]: df['SKU_raw']=df['SKU'].astype(str)
    master['SKU']=master['SKU'].map(clean_sku); stock['SKU']=stock['SKU'].map(clean_sku); sales['SKU']=sales['SKU'].map(clean_sku)
    for df in [sales,cust]: df['CustomerID']=df['CustomerID'].astype(str).str.strip().str.upper()
    master['Item Name']=master['Item Name'].astype(str).str.strip()
    master['Status_clean']=master['Status'].astype(str).str.strip().str.lower()
    master['Alternate Flag_clean']=master['Alternate Flag'].fillna('').astype(str).str.strip().str.upper()
    # Diagnostics for raw formatting before proceeding.
    raw_sku_count = int(((sales['SKU_raw'].astype(str) != sales['SKU_raw'].astype(str).str.strip()) | (sales['SKU_raw'].astype(str) != sales['SKU_raw'].astype(str).str.upper())).sum())
    if raw_sku_count:
        issues.append({'Severity':'Medium','Problem':f'{raw_sku_count} sales rows have SKU whitespace and/or lowercase formatting','Handling':'Trimmed whitespace and converted SKU values to uppercase before joins.','Source fix':'Standardize SKU format in the source export; store SKU as a validated uppercase key.'})
    # Drop exact duplicate sales rows; flag count as extra repeated rows.
    duplicate_count=int(sales.drop(columns=['SKU_raw']).duplicated().sum())
    if duplicate_count:
        issues.append({'Severity':'High','Problem':f'{duplicate_count} exact duplicate sales rows (duplicate records appear in 6 pairs)','Handling':'Removed exact duplicate rows before sales aggregation to avoid double-counting; rows are identified by exact match across all sales fields.','Source fix':'Investigate duplicate export/ETL events and add a unique transaction-line ID so true repeated purchases can be distinguished from duplicate extracts.'})
        sales=sales.drop_duplicates(subset=['Month','CustomerID','CustomerName','Industry','SKU','Qty','Value (Rs)']).copy()
    # Month format consistency diagnostic; preserve source rows, month is not needed for 6-month sum.
    month_raw=sales['Month'].dropna().astype(str).str.strip()
    month_formats=month_raw.map(lambda x: 'ISO YYYY-MM' if re.fullmatch(r'\d{4}-\d{2}',x) else ('Mon-YYYY' if re.fullmatch(r'[A-Za-z]{3}-\d{4}',x) else 'Other')).nunique()
    if month_formats>1:
        # Parse both source formats into an internal month key; the source label remains available for audit.
        month_series=sales['Month'].astype(str).str.strip()
        parsed=pd.to_datetime(month_series,format='%b-%Y',errors='coerce')
        parsed=parsed.fillna(pd.to_datetime(month_series+'-01',format='%Y-%m-%d',errors='coerce'))
        sales['_MonthKey']=parsed.dt.to_period('M').astype(str)
        issues.append({'Severity':'Low','Problem':'Month field uses inconsistent formats (e.g. Apr-2026 and 2026-06)','Handling':'Parsed supported formats to an internal YYYY-MM month key for validation; six-month totals use all valid case-period rows.','Source fix':'Export all months in one standard ISO format (YYYY-MM) and validate against Apr–Sep 2026.'})
    # Numeric validation.
    for c in AGE_COLS: stock[c]=pd.to_numeric(stock[c],errors='coerce')
    for c in ['Master Pack Qty','Unit Cost (Rs)']: master[c]=pd.to_numeric(master[c],errors='coerce')
    sales['Qty']=pd.to_numeric(sales['Qty'],errors='coerce'); sales['Value (Rs)']=pd.to_numeric(sales['Value (Rs)'],errors='coerce')
    negative_mask=stock[AGE_COLS].lt(0).any(axis=1)
    if negative_mask.any():
        skus=', '.join(stock.loc[negative_mask,'SKU'].astype(str).tolist())
        issues.append({'Severity':'High','Problem':f'Negative stock quantity in aging bucket(s) for {skus}','Handling':'For stock totals and aging calculations, negative bucket values are floored at zero and the row is flagged. This avoids negative stock artificially reducing aged value, but requires source confirmation.','Source fix':'Reconcile the affected SKU aging ledger and correct the export; do not automatically treat negative inventory as physical stock.'})
        stock[AGE_COLS]=stock[AGE_COLS].clip(lower=0)
    # Duplicate master keys and missing numeric essentials.
    dup_master=master[master['SKU'].duplicated(keep=False)]
    if not dup_master.empty:
        issues.append({'Severity':'High','Problem':f'Duplicate SKU keys in Item_Master: {", ".join(dup_master.SKU.unique())}','Handling':'Stopped analysis because a unique item master key is required.','Source fix':'Resolve duplicate master records before running the review.'})
        raise ValueError('Duplicate SKU keys in Item_Master: '+', '.join(dup_master.SKU.unique()))
    if master[['Master Pack Qty','Unit Cost (Rs)']].isna().any().any() or (master['Master Pack Qty']<=0).any() or (master['Unit Cost (Rs)']<0).any():
        issues.append({'Severity':'High','Problem':'Missing/invalid master pack quantity or unit cost','Handling':'Stopped analysis because classification and aged value would be unreliable.','Source fix':'Populate positive pack quantity and non-negative unit cost for every item.'})
        raise ValueError('Invalid or missing Master Pack Qty / Unit Cost in Item_Master.')
    # Item-name-to-SKU map handles alternate item references stored by name.
    name_to_sku={norm_text(r['Item Name']):r['SKU'] for _,r in master.iterrows()}
    switch_map={}
    for _,r in master.iterrows():
        if r['Alternate Flag_clean']=='SWITCH':
            alt=name_to_sku.get(norm_text(r['Alternate Item (name)']))
            if not alt:
                issues.append({'Severity':'High','Problem':f"SWITCH alternate item could not be matched for {r['SKU']}",'Handling':'No consolidation performed for this pair; flagged for review.','Source fix':'Use alternate SKU as a key instead of item name and ensure it exists in Item_Master.'})
            else: switch_map[r['SKU']]=alt
    # Report non-SWITCH alternate flags and unresolved alternate names.
    non_switch=master[(master['Alternate Flag_clean']!='') & (master['Alternate Flag_clean']!='SWITCH')]
    for _,r in non_switch.iterrows():
        issues.append({'Severity':'Medium','Problem':f"{r['SKU']} has Alternate Flag '{r['Alternate Flag_clean']}' and an alternate item name but is not marked SWITCH",'Handling':'No SKU consolidation is applied because only SWITCH triggers rule R1.','Source fix':'Confirm whether the flag is intentional and document allowed flag values (e.g. SWITCH, EXCLUDE, blank).'})
    blank_flag_alt=master[master['Alternate Item (name)'].notna() & (master['Alternate Flag_clean']=='')]
    for _,r in blank_flag_alt.iterrows():
        issues.append({'Severity':'Medium','Problem':f"{r['SKU']} has an alternate item name but a blank Alternate Flag",'Handling':'No switch consolidation is applied unless the flag is explicitly SWITCH, as required by R1.','Source fix':'Confirm whether the alternate reference is informational or a true pack-size switch; populate the controlled flag accordingly.'})
    for _,r in master[master['Alternate Item (name)'].notna()].iterrows():
        if norm_text(r['Alternate Item (name)']) not in name_to_sku:
            issues.append({'Severity':'Medium','Problem':f"Alternate item name for {r['SKU']} does not exactly match an Item_Master name after whitespace normalization",'Handling':'Alternate reference is not used unless Alternate Flag is SWITCH; SWITCH mismatches are flagged separately.','Source fix':'Use alternate SKU codes and enforce referential integrity.'})
    # Inactive stock alert; excluded from active item analysis.
    inactive=master[master['Status_clean']!='active'].merge(stock,on='SKU',how='left')
    for _,r in inactive.iterrows():
        qty=float(pd.to_numeric(r[AGE_COLS],errors='coerce').fillna(0).sum()) if all(c in r for c in AGE_COLS) else 0
        if qty>0:
            issues.append({'Severity':'High','Problem':f"Inactive item {r['SKU']} still holds {qty:,.0f} units of stock",'Handling':'Excluded from item classification and Push targets; surfaced in the inactive-stock alert.','Source fix':'Confirm whether stock should be written off, transferred, returned, or the item reactivated.'})
    # Missing customer industry.
    cust['Industry']=cust['Industry'].replace(r'^\s*$',np.nan,regex=True)
    missing_ind=cust['Industry'].isna()
    if missing_ind.any():
        ids=', '.join(cust.loc[missing_ind,'CustomerID'].astype(str).tolist())
        issues.append({'Severity':'Medium','Problem':f'Missing industry for customer(s): {ids}','Handling':'Such customers are never assigned bucket 2 (Similar business), as required by R6. They may still qualify for bucket 1 or 3.','Source fix':'Maintain a required, controlled industry field in the customer master.'})
    # Unmatched SKUs/customers.
    known_skus=set(master.SKU)
    unknown_stock=sorted(set(stock.SKU)-known_skus); unknown_sales=sorted(set(sales.SKU)-known_skus)
    if unknown_stock:
        issues.append({'Severity':'High','Problem':'Stock rows reference unknown SKUs: '+', '.join(unknown_stock),'Handling':'Excluded unknown SKUs from item calculations and listed in diagnostics.','Source fix':'Repair SKU references or add missing items to the item master.'})
    if unknown_sales:
        issues.append({'Severity':'High','Problem':'Sales rows reference unknown SKUs after normalization: '+', '.join(unknown_sales),'Handling':'Excluded unmatched SKUs from item sales aggregation and listed in diagnostics.','Source fix':'Standardize SKU references and reconcile sales export against Item_Master.'})
    cust['CustomerName']=cust['CustomerName'].astype(str).str.strip()
    cust['Area']=cust['Area'].astype(str).str.strip()
    # Customer master is authoritative for industry; fill from sales only if master has no industry.
    sales_industry=sales[['CustomerID','Industry']].copy(); sales_industry['Industry']=sales_industry['Industry'].replace(r'^\s*$',np.nan,regex=True)
    industry_lookup=sales_industry.dropna().drop_duplicates('CustomerID').set_index('CustomerID')['Industry'].to_dict()
    cust['Industry_effective']=cust['Industry']
    cust.loc[cust['Industry_effective'].isna(),'Industry_effective']=cust.loc[cust['Industry_effective'].isna(),'CustomerID'].map(industry_lookup)
    # Normalize each sale to the new SKU when old SKU has SWITCH.
    sales['Canonical SKU']=sales['SKU'].map(lambda s: switch_map.get(s,s))
    stock['Canonical SKU']=stock['SKU'].map(lambda s: switch_map.get(s,s))
    # Aggregate stock buckets after switch mapping.
    stock_active=stock[stock['SKU'].isin(known_skus)].copy()
    stock_agg=stock_active.groupby('Canonical SKU',as_index=False)[AGE_COLS].sum()
    sales_valid=sales[sales['SKU'].isin(known_skus)].copy()
    sales_agg=sales_valid.groupby('Canonical SKU',as_index=False).agg(**{'6M Sales Qty':('Qty','sum'),'6M Sales Value (Rs)':('Value (Rs)','sum')})
    # Base active master items, keep only new SKU record when a switch partner maps into it.
    active=master[master['Status_clean']=='active'].copy()
    old_switch_skus=set(switch_map.keys())
    active=active[~active['SKU'].isin(old_switch_skus)].copy()
    active=active.rename(columns={'SKU':'Canonical SKU'})
    # If target alternate inactive or missing, warn; only active canonical items can be reported.
    items=active.merge(stock_agg,on='Canonical SKU',how='left').merge(sales_agg,on='Canonical SKU',how='left')
    for c in AGE_COLS: items[c]=items[c].fillna(0)
    for c in ['6M Sales Qty','6M Sales Value (Rs)']: items[c]=items[c].fillna(0)
    items['Total Stock Qty']=items[AGE_COLS].sum(axis=1)
    items['Aged Qty']=items[AGED_COLS].sum(axis=1)
    items['Aged Value (Rs)']=items['Aged Qty']*items['Unit Cost (Rs)']
    items['Average Monthly Sales Qty']=items['6M Sales Qty']/6
    items['Months to Clear']=np.where(items['Average Monthly Sales Qty']>0,items['Total Stock Qty']/items['Average Monthly Sales Qty'],np.inf)
    def classify(r):
        if r['6M Sales Qty'] < r['Master Pack Qty']: return 'Red List'
        if r['Aged Qty'] <= 0 or r['Months to Clear'] < 3: return 'No Action'
        return 'Push'
    items['Class']=items.apply(classify,axis=1)
    items['Switch Partner SKUs']=items['Canonical SKU'].map(lambda k:', '.join(sorted([old for old,new in switch_map.items() if new==k])))
    items['Item Name']=items['Item Name'].astype(str)
    # Sales customer values across six months, deduplicated exact rows and normalized SKU.
    customer_totals=sales_valid.groupby('CustomerID',as_index=False).agg(**{'6M Customer Purchase Value (Rs)':('Value (Rs)','sum')})
    cust2=cust.merge(customer_totals,on='CustomerID',how='left')
    cust2['6M Customer Purchase Value (Rs)']=cust2['6M Customer Purchase Value (Rs)'].fillna(0)
    eligible=set(cust2.loc[cust2['6M Customer Purchase Value (Rs)']>100000,'CustomerID'])
    cust_by_id=cust2.set_index('CustomerID').to_dict('index')
    # Purchase history in canonical SKU terms.
    buyer_sales=sales_valid.groupby(['Canonical SKU','CustomerID'],as_index=False).agg(**{'Item Qty Bought':('Qty','sum'),'Item Sales Value (Rs)':('Value (Rs)','sum')})
    buyers_by_sku={sku:set(g.CustomerID) for sku,g in buyer_sales.groupby('Canonical SKU')}
    cust_skus={cid:set(g['Canonical SKU']) for cid,g in sales_valid.groupby('CustomerID')}
    # For industry similarity, use effective customer master industry.
    industry_by_id={r['CustomerID']:r['Industry_effective'] for _,r in cust2.iterrows()}
    item_category=items.set_index('Canonical SKU')['Category'].to_dict()
    target_rows=[]
    for _,item in items[items['Class']=='Push'].iterrows():
        sku=item['Canonical SKU']; buyers=buyers_by_sku.get(sku,set())
        bought_industries={industry_by_id.get(cid) for cid in buyers if pd.notna(industry_by_id.get(cid)) and str(industry_by_id.get(cid)).strip()}
        cat=item['Category']
        for cid in eligible:
            if cid not in cust_by_id: continue
            c=cust_by_id[cid]
            if cid in buyers: bucket=1; reason='Bought this item / switch partner'
            elif pd.notna(c.get('Industry_effective')) and str(c.get('Industry_effective')).strip() and c.get('Industry_effective') in bought_industries: bucket=2; reason='Same industry as an item buyer'
            elif any((s != sku and item_category.get(s)==cat) for s in cust_skus.get(cid,set())): bucket=3; reason='Bought another item in same category'
            else: continue
            target_rows.append({'SKU':sku,'Item Name':item['Item Name'],'Category':cat,'CustomerID':cid,'CustomerName':c['CustomerName'],'Industry':c.get('Industry_effective'),'Area':c['Area'],'Bucket':bucket,'Bucket Label':{1:'1 – Bought before',2:'2 – Similar business',3:'3 – Same category'}[bucket],'Reason':reason,'6M Customer Purchase Value (Rs)':c['6M Customer Purchase Value (Rs)']})
    targets=pd.DataFrame(target_rows)
    if not targets.empty:
        targets=targets.sort_values(['SKU','Bucket','6M Customer Purchase Value (Rs)','CustomerID'],ascending=[True,True,False,True]).reset_index(drop=True)
        targets['Call Rank']=targets.groupby('SKU').cumcount()+1
    # Item summary.
    class_order=['Red List','No Action','Push']
    counts=items['Class'].value_counts().reindex(class_order,fill_value=0).rename_axis('Class').reset_index(name='Item Count')
    push=items[items['Class']=='Push'].copy()
    top_push=push.sort_values('Aged Value (Rs)',ascending=False).head(10)[['Canonical SKU','Item Name','Category','Aged Qty','Unit Cost (Rs)','Aged Value (Rs)','Total Stock Qty','6M Sales Qty','Months to Clear']]
    # Inactive stock alerts table.
    inactive_stock=master[master['Status_clean']!='active'].merge(stock,on='SKU',how='left')
    if not inactive_stock.empty:
        inactive_stock['Total Stock Qty']=inactive_stock[AGE_COLS].fillna(0).sum(axis=1)
        inactive_stock=inactive_stock[inactive_stock['Total Stock Qty']>0][['SKU','Item Name','Status','Total Stock Qty']+AGE_COLS]
    else: inactive_stock=pd.DataFrame(columns=['SKU','Item Name','Status','Total Stock Qty']+AGE_COLS)
    # Format display column name and ensure sorted class lists.
    items=items.rename(columns={'Canonical SKU':'SKU'})
    items=items[['SKU','Item Name','Category','Master Pack Qty','Unit Cost (Rs)','Switch Partner SKUs']+AGE_COLS+['Total Stock Qty','Aged Qty','Aged Value (Rs)','6M Sales Qty','6M Sales Value (Rs)','Average Monthly Sales Qty','Months to Clear','Class']]
    return {
      'items':items,'targets':targets,'counts':counts,'top_push':top_push,
      'issues':pd.DataFrame(issues,columns=['Severity','Problem','Handling','Source fix']),
      'inactive_stock':inactive_stock,'push_aged_value':float(push['Aged Value (Rs)'].sum()),
      'switch_map':switch_map,'sales_clean':sales_valid,'customer_totals':customer_totals,
      'summary':{'active_items_analysed':len(items),'red_list':int((items['Class']=='Red List').sum()),'no_action':int((items['Class']=='No Action').sum()),'push':int((items['Class']=='Push').sum()),'push_aged_value':float(push['Aged Value (Rs)'].sum())}
    }

def read_workbook(uploaded_file):
    """Load the standard combined workbook uploaded through Streamlit."""
    raw=uploaded_file.getvalue() if hasattr(uploaded_file,'getvalue') else uploaded_file
    xf=pd.ExcelFile(io.BytesIO(raw) if isinstance(raw,bytes) else raw)
    aliases={'item_master':'Item_Master','stock_aging':'Stock_Aging','sales_lines':'Sales_Lines','customers':'Customers'}
    data={}
    for sh in xf.sheet_names:
        key=aliases.get(sh.lower().strip())
        if key: data[key]=pd.read_excel(xf,sheet_name=sh)
    missing=[k for k in REQUIRED if k not in data]
    if missing: raise ValueError('Workbook missing sheets: '+', '.join(missing))
    return data

def excel_export(results):
    buf=io.BytesIO()
    with pd.ExcelWriter(buf,engine='openpyxl') as writer:
        results['items'].to_excel(writer,index=False,sheet_name='Item Review')
        results['counts'].to_excel(writer,index=False,sheet_name='Class Summary')
        results['top_push'].to_excel(writer,index=False,sheet_name='Top 10 Push')
        results['targets'].to_excel(writer,index=False,sheet_name='Customer Targets')
        results['issues'].to_excel(writer,index=False,sheet_name='Data Issues')
        results['inactive_stock'].to_excel(writer,index=False,sheet_name='Inactive Stock')
    buf.seek(0); return buf.getvalue()
