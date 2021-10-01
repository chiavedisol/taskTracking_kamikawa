import streamlit as st
import base64
import pandas as pd
import re
from io import BytesIO


def main():
    st.title("病院管理タスクトラッキング")
    adms = st.multiselect('何を基準にCSVファイルを作成しますか',
                          ['届出', '実務', '規定', '責任管理者', '研修', '責任者'])

    csv_file = st.file_uploader("Upload CSV or EXCEL",
                                type=['csv', 'xlsx', 'xls'])
    if st.button("Process"):
        if csv_file is not None:
            file_details = {
                "Filename": csv_file.name,
                "FileType": csv_file.type,
                "FileSize": csv_file.size
            }
            st.write(file_details)

            task_df = pd.read_excel(csv_file, sheet_name=None)
            sheet_names = list(task_df.keys())
            for adm in adms:
                final_report = []
                for sheet in sheet_names:
                    final_report.extend(df_from_sheet(task_df, sheet, adm))
                final_df = pd.io.json.json_normalize(final_report)
                final_df = final_df.reindex(
                    columns=['シート名', '大項目', '中項目', 'チェック', '種別'])
                st.write(final_df)
                download_csv(final_df, adm)


def df_from_sheet(df, sheet, adm):
    df = df[sheet]
    export_list = []
    daikomoku = df.iloc[:, 1]
    tyukomoku = df.iloc[:, 8]
    check = df.iloc[:, 12]
    kanrisha = df.iloc[:, 17]
    for i in range(len(kanrisha)):
        if not pd.isnull(kanrisha[i]):
            if adm in kanrisha[i]:
                middle = tyukomoku[i]
                cntdwn = 0
                while pd.isnull(middle):
                    middle = tyukomoku[i - cntdwn]
                    cntdwn += 1
                large = daikomoku[i]
                cntdwn = 0
                while pd.isnull(large) or (not re.match(r'[0-9]', large)):
                    large = daikomoku[i - cntdwn]
                    cntdwn += 1
                result = {
                    'シート名': sheet,
                    '大項目': large,
                    '中項目': middle,
                    'チェック': check[i],
                    '種別': adm
                }
                export_list.append(result)
    return export_list


def to_excel(df):
    output = BytesIO()
    writer = pd.ExcelWriter(output, engine='xlsxwriter')
    df.to_excel(writer, index=False, sheet_name='Sheet1', float_format="%.2f")
    writer.save()
    processed_data = output.getvalue()
    return processed_data


def download_csv(df, adm):
    xlsx = to_excel(df)
    b64 = base64.b64encode(xlsx)
    href = f'<a href="data:application/octet-stream;base64,{b64.decode()}" download="{adm}_集計.xlsx">Download</a>'
    st.markdown(f"{adm}に関するCSVをダウンロードする {href}", unsafe_allow_html=True)


if __name__ == '__main__':
    main()
