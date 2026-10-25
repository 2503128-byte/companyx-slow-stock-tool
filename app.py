import streamlit as st
import pandas as pd
from pathlib import Path
from engine import analyse, load_files, excel_export

st.set_page_config(page_title='Company X | Slow-Stock Review', page_icon='📦', layout='wide')

st.markdown('''<style>
.block-container {padding-top: 1.2rem; padding-bottom: 2rem; max-width: 1400px;}
[data-testid="stMetric"] {background: rgba(120,120,120,.07); padding: 14px 16px; border-radius: 12px;}
.small-note {font-size: .9rem; opacity: .8;}
@media (max-width: 700px) {.block-container {padding-left: .7rem; padding-right: .7rem;}}
</style>''', unsafe_allow_html=True)

st.title('📦 Slow-Stock Review')
st.caption('Company X · Monthly inventory action tool · Synthetic case-study data')
st.write('Upload the monthly workbook or the four separate exports. The tool checks data quality, consolidates pack-size switches, classifies active items, recommends customer targets, and creates an Excel export.')

SAMPLE = Path(__file__).parent / 'sample_data' / 'Case_Data_Technology_and_Data_Management_All_Sheets.xlsx'
with st.sidebar:
    st.header('Run a review')
    use_sample = st.checkbox('Load included sample data', value=True, help='Use this to see the case study working immediately. Uncheck to upload your own monthly files.')
    uploads = st.file_uploader('Upload workbook or four exports', type=['xlsx','xls','csv'], accept_multiple_files=True, disabled=use_sample)
    st.divider()
    st.markdown('**Required data**')
    st.caption('Item_Master · Stock_Aging · Sales_Lines · Customers')
    st.caption('The combined workbook is easiest. For separate files, retain the sheet names or filenames from the case pack.')

try:
    if use_sample:
        with open(SAMPLE, 'rb') as f:
            frames = load_files({'Case_Data_Technology_and_Data_Management_All_Sheets.xlsx': f.read()})
    elif uploads:
        frames = load_files({f.name: f for f in uploads})
    else:
        st.info('Choose “Load included sample data” or upload the four required data exports in the sidebar.')
        st.stop()
    result = analyse(frames['Item_Master'], frames['Stock_Aging'], frames['Sales_Lines'], frames['Customers'])
except Exception as exc:
    st.error(f'Could not run the review: {exc}')
    st.stop()

summary = result['summary']
# Keep the hierarchy simple and usable on mobile.
t1,t2,t3,t4 = st.columns(4)
t1.metric('Items analysed', f"{summary['active_items_analysed']}")
t2.metric('🔴 Red List', f"{summary['red_list']}")
t3.metric('No Action', f"{summary['no_action']}")
t4.metric('Push', f"{summary['push']}")
st.metric('Total aged value in Push items', f"₹{summary['push_aged_value']:,.0f}")

with st.expander('What the classifications mean', expanded=False):
    st.markdown('- **Red List:** six-month sales quantity is less than one master pack. No customer targets are generated.\n- **No Action:** no aged stock, or months to clear is below 3.\n- **Push:** remaining active items that have aged stock and at least 3 months to clear.')

st.subheader('Data-quality checks')
if result['issues'].empty:
    st.success('No data-quality issues detected by the current checks.')
else:
    st.warning(f"{len(result['issues'])} data-quality issue(s) detected. Review before acting on recommendations.")
    st.dataframe(result['issues'], use_container_width=True, hide_index=True)
if not result['inactive_stock'].empty:
    st.error('Inactive items with stock — excluded from classification and customer targets.')
    st.dataframe(result['inactive_stock'], use_container_width=True, hide_index=True)

st.subheader('Top 10 Push items by aged value')
top = result['top_push'].copy()
if not top.empty:
    top['Aged Value (Rs)'] = top['Aged Value (Rs)'].map(lambda x: f'₹{x:,.0f}')
    top['Months to Clear'] = top['Months to Clear'].map(lambda x: f'{x:.1f}')
st.dataframe(top, use_container_width=True, hide_index=True)

st.subheader('Item review')
items = result['items'].copy()
classes = ['All'] + [x for x in ['Red List','No Action','Push'] if x in items['Class'].unique()]
cats = ['All'] + sorted(items['Category'].dropna().astype(str).unique().tolist())
f1,f2,f3 = st.columns([1,1,2])
selected_class = f1.selectbox('Filter by class', classes)
selected_category = f2.selectbox('Filter by category', cats)
search = f3.text_input('Search SKU or item name', placeholder='e.g. SKU1005')
filtered = items.copy()
if selected_class != 'All': filtered = filtered[filtered['Class'] == selected_class]
if selected_category != 'All': filtered = filtered[filtered['Category'] == selected_category]
if search.strip():
    term = search.strip().lower()
    filtered = filtered[filtered['SKU'].str.lower().str.contains(term) | filtered['Item Name'].str.lower().str.contains(term)]
st.caption(f'Showing {len(filtered)} of {len(items)} analysed items.')
st.dataframe(filtered, use_container_width=True, hide_index=True)

st.subheader('Customer targets for a Push item')
push_skus = sorted(items.loc[items['Class']=='Push','SKU'].tolist())
if push_skus:
    selected_sku = st.selectbox('Choose a Push SKU', push_skus, index=push_skus.index('SKU1005') if 'SKU1005' in push_skus else 0)
    target = result['targets']
    target = target[target['SKU'] == selected_sku].copy() if not target.empty else target
    if target.empty:
        st.info('No eligible customers met the targeting rules for this item.')
    else:
        st.caption('Targets are ordered by bucket (1 → 2 → 3), then by total six-month customer purchase value, highest first. Only customers above ₹1,00,000 are shown.')
        mobile_cols = ['Call Rank','CustomerName','Area','Industry','Bucket Label','6M Customer Purchase Value (Rs)','Reason']
        mobile = target[mobile_cols].copy()
        mobile['6M Customer Purchase Value (Rs)'] = mobile['6M Customer Purchase Value (Rs)'].map(lambda x: f'₹{x:,.0f}')
        st.dataframe(mobile, use_container_width=True, hide_index=True)
        # Optional quick sales message helper; deterministic template, no external AI/key.
        with st.expander('Draft a short WhatsApp message'):
            names = target['CustomerName'].tolist()
            customer_choice = st.selectbox('Customer', names, key='message_customer')
            item_row = items[items['SKU']==selected_sku].iloc[0]
            message = f"Hello {customer_choice}, we wanted to check whether you have a requirement for {item_row['Item Name']} ({item_row['SKU']}). We can share availability and pricing details if useful. Would you like us to follow up?"
            st.text_area('Editable message', value=message, height=110)

st.subheader('Export')
st.write('Download an Excel workbook containing the item review, class summary, top Push items, customer targets, data issues, and inactive-stock alerts.')
st.download_button('⬇️ Download review results (.xlsx)', data=excel_export(result), file_name='CompanyX_Slow_Stock_Review.xlsx', mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
st.caption('Important: this tool flags source-data anomalies rather than silently treating them as resolved. Confirm high-severity issues with the data owner before taking inventory or sales action.')
