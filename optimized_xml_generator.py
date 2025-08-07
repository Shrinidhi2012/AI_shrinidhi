import streamlit as st
import pandas as pd
import zipfile
import io
import re
import oracledb
from datetime import datetime

st.set_page_config(page_title="XML Generator Tool", layout="centered")
st.title("📄 Report XML Generator")
st.markdown("Upload the input Excel file and click the button to generate XML files in a ZIP archive.")

input_file = st.file_uploader("Upload Input Excel File", type=["xlsx"])

# Oracle connection (replace with your credentials)
oracledb.init_oracle_client()
conn = oracledb.connect(
    user="your_user",
    password="your_password",
    dsn="your_host:your_port/your_service"
)

if input_file:
    input_df = pd.read_excel(input_file)
    input_df['report'] = input_df['report'].astype(str).str.strip().str.lower()
    input_df['fund'] = input_df['fund'].astype(str).str.strip()

    reports = tuple(input_df['report'].unique())
    funds = tuple(input_df['fund'].unique())

    # Optimized: Get only matching rows from DB with vectorized filters
    report_pattern_sql = '|'.join([re.escape(r) for r in reports])
    fund_pattern_sql = '|'.join([re.escape(f) for f in funds])

    query = f"""
        SELECT RS_REPORT, RS_PARAMETERS, RS_FORMAT, RS_START, RS_STATUS, RS_ENGINE
        FROM KIRA_STAR.REPORT_STATISTICS
        WHERE LOWER(RS_ENGINE) = 'actuate'
          AND LOWER(RS_STATUS) = 'succeeded'
          AND REGEXP_LIKE(LOWER(RS_REPORT), :report_pattern, 'i')
          AND REGEXP_LIKE(RS_PARAMETERS, :fund_pattern, 'i')
    """

    cursor = conn.cursor()
    cursor.execute(query, {
        "report_pattern": report_pattern_sql,
        "fund_pattern": fund_pattern_sql
    })

    db_rows = cursor.fetchall()
    columns = [col[0] for col in cursor.description]
    db_df = pd.DataFrame(db_rows, columns=columns)
    cursor.close()
    conn.close()

    if db_df.empty:
        st.warning("No matches found in DB.")
    else:
        # Normalize for matching
        db_df['rs_report_norm'] = db_df['RS_REPORT'].str.lower().str.strip()
        db_df['rs_fund'] = db_df['RS_PARAMETERS'].str.extract(r'fund\s*:\s*([^\s;]+)', expand=False)
        db_df['RSDateTime'] = pd.to_datetime(db_df['RS_START'], errors='coerce')

        # Create ZIP in-memory
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w", zipfile.ZIP_DEFLATED) as zipf:
            for _, row in input_df.iterrows():
                rpt, fund, date = row['report'], row['fund'], row.get('date', None)
                matched = db_df[
                    (db_df['rs_report_norm'].str.contains(rpt)) &
                    (db_df['RS_PARAMETERS'].str.contains(re.escape(fund), case=False, na=False))
                ].copy()

                if matched.empty:
                    continue

                if pd.notna(date):
                    date = pd.to_datetime(date, errors='coerce')
                    matched = matched[matched['RSDateTime'].dt.date == date.date()]
                else:
                    max_date = matched['RSDateTime'].max()
                    matched = matched[matched['RSDateTime'].dt.date == max_date.date()]

                matched.sort_values(by='RSDateTime', ascending=False, inplace=True)
                for idx, m in matched.iterrows():
                    param_string = m['RS_PARAMETERS']
                    report_format = m['RS_FORMAT']
                    report_path = m['RS_REPORT']
                    xml_path = "/".join(report_path.split("/")[:-1])

                    # Generate <parameter> XML
                    parameters = [p.strip() for p in param_string.split(';') if ':' in p]
                    xml_params = ""
                    for p in parameters:
                        key, val = map(str.strip, p.split(':', 1))
                        xml_params += f'  <parameter name="{key}">\n    <value>{val}</value>\n  </parameter>\n'

                    xml_content = (
                        '<?xml version="1.0" encoding="UTF-8"?>\n'
                        f'<reportTestcase name="{rpt}" format="{report_format}" pfad="{xml_path}">\n'
                        f'{xml_params}</reportTestcase>'
                    )

                    date_str = m['RSDateTime'].strftime("%Y-%m-%d_%H-%M")
                    filename = f"{rpt}_{fund}_{date_str}_{idx+1}.xml"
                    zipf.writestr(filename, xml_content)

        st.success("✅ XMLs generated!")
        st.download_button(
            label="📦 Download All XMLs as ZIP",
            data=zip_buffer.getvalue(),
            file_name="generated_xmls.zip",
            mime="application/zip"
        )
