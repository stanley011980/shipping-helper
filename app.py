import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="出貨單轉換小幫手", page_icon="📦", layout="centered")

st.title("📦 出貨單自動轉換小幫手")
st.write("請上傳原始的出貨單 (.xls 或 .xlsx)，系統會自動剔除非黑豆漿的倉儲課、重新計算合計，並匯出為乾淨的 6 欄 Excel 檔案。")

# 支援多檔案同時上傳
uploaded_files = st.file_uploader("請選擇或拖曳出貨單檔案", type=['xls', 'xlsx'], accept_multiple_files=True)

if uploaded_files:
    for file in uploaded_files:
        try:
            # 讀取原始 Excel 檔案
            df = pd.read_excel(file, header=None)
            
            out_data = []
            # 建立目標 6 欄表頭
            out_data.append(["移出單位", "產品代號", "產品名稱", "移入單位", "數量", "單位"])
            
            current_out_unit = ""
            current_prod_code = ""
            current_prod_name = ""
            current_sum = 0.0
            current_unit_str = ""
            
            for idx, row in df.iterrows():
                # 跳過前 10 行非資料表頭區
                if idx < 10:
                    continue
                    
                col2 = str(row[2]).strip() if pd.notna(row[2]) else ""
                col4 = str(row[4]).strip() if pd.notna(row[4]) else ""
                col6 = str(row[6]).strip() if pd.notna(row[6]) else ""
                col8 = str(row[8]).strip() if pd.notna(row[8]) else ""
                col10 = str(row[10]).strip() if pd.notna(row[10]) else ""
                col11 = str(row[11]).strip() if pd.notna(row[11]) else ""
                
                # 追蹤當前的合併儲存格狀態
                if col2 not in ["", "nan", "None", "[移出單位]"]:
                    current_out_unit = col2
                if col4 not in ["", "nan", "None", "[產品代號]"]:
                    current_prod_code = col4
                if col6 not in ["", "nan", "None", "[ 產品名稱 ]"]:
                    current_prod_name = col6
                    
                if col8 in ["", "nan", "None", "[ 移入單位 ]"]:
                    continue
                    
                if col8 == "產品合計":
                    # 寫入重新計算後的總和
                    out_data.append([
                        current_out_unit,
                        current_prod_code,
                        current_prod_name,
                        "產品合計",
                        round(current_sum, 4) if current_sum % 1 != 0 else int(current_sum),
                        current_unit_str
                    ])
                    current_sum = 0.0
                else:
                    # 倉儲課過濾邏輯：不含黑豆漿就跳過
                    if col8 == "倉儲課" and "黑豆漿" not in current_prod_name:
                        continue 
                        
                    # 數字轉型與累加
                    try:
                        qty = float(col10)
                        current_sum += qty
                    except:
                        qty = col10
                        
                    if col11 not in ["", "nan", "None"]:
                        current_unit_str = col11
                        
                    out_data.append([
                        current_out_unit,
                        current_prod_code,
                        current_prod_name,
                        col8,
                        qty if type(qty) != float or qty % 1 != 0 else int(qty),
                        col11 if col11 not in ["nan", "None"] else ""
                    ])
            
            processed_df = pd.DataFrame(out_data)
            
            # 將 DataFrame 轉為 Excel 二進位串流
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                processed_df.to_excel(writer, index=False, header=False)
            processed_data = output.getvalue()
            
            # 顯示成功提示與獨立下載按鈕
            st.success(f"✅ {file.name} 處理完成！")
            st.download_button(
                label=f"⬇️ 下載 {file.name.split('.')[0]}_新格式.xlsx",
                data=processed_data,
                file_name=f"處理完成_{file.name.split('.')[0]}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            )
        except Exception as e:
            st.error(f"❌ 處理檔案 {file.name} 時發生錯誤，請確認格式是否正確。")
