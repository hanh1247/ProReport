import streamlit as st
import pandas as pd
import time
import io
from datetime import date

st.set_page_config(page_title="Tool Xử Lý File", layout="centered")

st.title(" Tool Xử Lý Báo cáo Số dư TKLK Chứng khoán")
st.write("Vui lòng tải file lên, hệ thống sẽ xử lý và trả lại file kết quả.")

# 1. Widget Upload File
uploaded = st.file_uploader("Chọn file từ máy tính của bạn",type=["xls", "xlsx"],accept_multiple_files=True)
if uploaded :
   st.success("Đã tải file lên thành công!")
 # try:
 #     df = pd.read_excel(uploaded)
 # except:
  #    df = pd.read_csv(uploaded)

    # 2. Nút bấm Xử lý
if st.button("Bắt đầu xử lý"):
  with st.spinner('Đang xử lý, vui lòng đợi...'):
            # --- ĐOẠN NÀY BẠN CHÈN CODE LOGIC XỬ LÝ CỦA BẠN VÀO ---
    time.sleep(2) # Giả lập thời gian xử lý
    for file in uploaded:
      name_file_uploaded = file
      print(f"File: {file}")
      df_0 = pd.read_excel(file, sheet_name= 0, skiprows =11)
      sheets = pd.ExcelFile(file).sheet_names
      len_sheets= len(sheets)
    for i in range(1,len_sheets):
      df_1 = pd.read_excel(file,sheet_name=i)
      df = pd.concat([df_0, df_1], axis=0, ignore_index=True)
        # df.columns
    del_col=[ "Unnamed: 0","Unnamed: 1", "Unnamed: 2", "Unnamed: 3",
       "Tổng cộng", "Unnamed: 5", "Unnamed: 6", "Unnamed: 7", "Unnamed: 8",
       "Unnamed: 9", "1"]
    for i in del_col:
      df = df.drop(i, axis=1)
    df = df[(df["Số ĐKNSH"]!="Số ĐKNSH") | (df["Số ĐKNSH"]!="Tổng cộng")]
    
    def fillnullvn(x):
       if "VN0000" not in x:
          return ""
       else:
          return x
    # df['STT'] = df['STT'].astype(str)
    # df['STT'] = df['STT'].apply(fillnullvn)  
    
    for i in range (0, len(df)):
      if df["STT"][i] == "":
        df["STT"][i] = df["STT"][i-1]

    df=df[df["Họ và tên"].notna()].reset_index(drop=True)
    df["STT"]= df["STT"].str.split(" - ").str[0].str.strip()
    df= df.rename(columns={'STT': 'Stock code'})
    df["Số lượng"] = pd.to_numeric(df["Số lượng"]
                                          .astype(str)
                                          .str.strip()
                                          .str.replace('.', '', regex=False)
                                          .str.replace(',', '.', regex=False),
                                         errors="coerce" )
    # coerce = ép về NaN nếu convert không được
    df= df[df["Số lượng"].notna() ]
    df["Số lượng_"] = df["Số lượng"].apply(
         lambda x: f"{x:,.0f}" if pd.notna(x) else "")
    df_export = df[[col for col in df.columns if not col.endswith("Số lượng")]]
    df_result = df_export.copy()

    # ---------------------------------------------------
    st.success("Đã xử lý xong!")

    buffer = io.BytesIO()
    df_result.to_excel(buffer, index=False)
    buffer.seek(0)

    # lấy tên file chuẩn
    file_name = name_file_uploaded.name
    clean_name = file_name.replace(".xls", "").replace(".xlsx", "")
    today = date.today()

    st.download_button(
        "Download Excel",
        data=buffer,
        file_name=f"Processed_{clean_name}_{today}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )
