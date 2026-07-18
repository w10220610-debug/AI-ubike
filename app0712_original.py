import streamlit as st
import pandas as pd

# 設定網頁標題與版面
st.set_page_config(page_title="臺東 Ubike 智慧調度決策系統", layout="centered")

st.title("🚚 臺東 Ubike 智慧調度工具")
st.write("已優化調度報表輸出邏輯，確保分析結果正常顯示。")

# ==========================================
# 1. 站點資料庫
# ==========================================

D2_DATABASE = {
    "平日配置": {
        "夜班配置": {
            "卑南鄉老人文康中心": {"bike": 3, "ebike": 0},
            "瑞源車站": {"bike": 2, "ebike": 4},
            "永安農產品銷售中心": {"bike": 4, "ebike": 0},
            "龍田大草坪": {"bike": 2, "ebike": 6},
            "鹿野車站": {"bike": 2, "ebike": 10},
            "福鹿山休閒農莊": {"bike": 1, "ebike": 6},
            "鹿野地區農會": {"bike": 2, "ebike": 2},
            "龍田老人會館": {"bike": 2, "ebike": 2},
            "永安環教中心": {"bike": 4, "ebike": 0},
            "關山車站": {"bike": 15, "ebike": 10},
            "米國學校": {"bike": 3, "ebike": 3},
            "關山工商": {"bike": 3, "ebike": 4},
            "關山鎮圖書館": {"bike": 4, "ebike": 2},
            "關山國小": {"bike": 4, "ebike": 3},
            "關山慈濟醫院": {"bike": 4, "ebike": 2},
            "池上車站": {"bike": 47, "ebike": 13},
            "池上郵局": {"bike": 13, "ebike": 12},
            "池上國中": {"bike": 8, "ebike": 10},
            "客家文化園區": {"bike": 10, "ebike": 5},
            "池上大坡池": {"bike": 4, "ebike": 21},
            "池上三號運動公園": {"bike": 7, "ebike": 8},
            "日暉國際渡假村": {"bike": 14, "ebike": 8}
        },
        "早班配置": {
            "卑南鄉老人文康中心": {"bike": 3, "ebike": 0},
            "瑞源車站": {"bike": 3, "ebike": 0},
            "永安農產品銷售中心": {"bike": 3, "ebike": 0},
            "龍田大草坪": {"bike": 4, "ebike": 0},
            "鹿野車站": {"bike": 2, "ebike": 6},
            "福鹿山休閒農莊": {"bike": 3, "ebike": 3},
            "鹿野地區農會": {"bike": 3, "ebike": 0},
            "龍田老人會館": {"bike": 4, "ebike": 0},
            "永安環教中心": {"bike": 3, "ebike": 0},
            "關山車站": {"bike": 4, "ebike": 10},
            "米國學校": {"bike": 2, "ebike": 2},
            "關山工商": {"bike": 2, "ebike": 4},
            "關山鎮圖書館": {"bike": 2, "ebike": 2},
            "關山國小": {"bike": 2, "ebike": 3},
            "關山慈濟醫院": {"bike": 2, "ebike": 2},
            "池上車站": {"bike": 20, "ebike": 8},
            "池上郵局": {"bike": 9, "ebike": 3},
            "池上國中": {"bike": 5, "ebike": 2},
            "客家文化園區": {"bike": 7, "ebike": 3},
            "池上大坡池": {"bike": 2, "ebike": 10},
            "池上三號運動公園": {"bike": 5, "ebike": 3},
            "日暉國際渡假村": {"bike": 6, "ebike": 4}
        },
        "晚班配置": {
            "卑南鄉老人文康中心": {"bike": 3, "ebike": 0},
            "瑞源車站": {"bike": 3, "ebike": 0},
            "永安農產品銷售中心": {"bike": 3, "ebike": 0},
            "龍田大草坪": {"bike": 4, "ebike": 0},
            "鹿野車站": {"bike": 2, "ebike": 6},
            "福鹿山休閒農莊": {"bike": 3, "ebike": 3},
            "鹿野地區農會": {"bike": 3, "ebike": 0},
            "龍田老人會館": {"bike": 4, "ebike": 0},
            "永安環教中心": {"bike": 3, "ebike": 0},
            "關山車站": {"bike": 4, "ebike": 10},
            "米國學校": {"bike": 2, "ebike": 2},
            "關山工商": {"bike": 2, "ebike": 5},
            "關山鎮圖書館": {"bike": 2, "ebike": 2},
            "關山國小": {"bike": 2, "ebike": 2},
            "關山慈濟醫院": {"bike": 2, "ebike": 2},
            "池上車站": {"bike": 20, "ebike": 5},
            "池上郵局": {"bike": 8, "ebike": 3},
            "池上國中": {"bike": 5, "ebike": 2},
            "客家文化園區": {"bike": 5, "ebike": 3},
            "池上大坡池": {"bike": 3, "ebike": 10},
            "池上三號運動公園": {"bike": 5, "ebike": 3},
            "日暉國際渡假村": {"bike": 6, "ebike": 4}
        }
    },
    "假日配置": {
        "夜班配置": {
            "卑南鄉老人文康中心": {"bike": 3, "ebike": 0},
            "瑞源車站": {"bike": 2, "ebike": 4},
            "永安農產品銷售中心": {"bike": 4, "ebike": 0},
            "龍田大草坪": {"bike": 2, "ebike": 6},
            "鹿野車站": {"bike": 2, "ebike": 10},
            "福鹿山休閒農莊": {"bike": 1, "ebike": 6},
            "鹿野地區農會": {"bike": 2, "ebike": 2},
            "龍田老人會館": {"bike": 2, "ebike": 2},
            "永安環教中心": {"bike": 4, "ebike": 0},
            "關山車站": {"bike": 15, "ebike": 10},
            "米國學校": {"bike": 3, "ebike": 3},
            "關山工商": {"bike": 3, "ebike": 4},
            "關山鎮圖書館": {"bike": 4, "ebike": 2},
            "關山國小": {"bike": 4, "ebike": 3},
            "關山慈濟醫院": {"bike": 4, "ebike": 2},
            "池上車站": {"bike": 47, "ebike": 13},
            "池上郵局": {"bike": 13, "ebike": 12},
            "池上國中": {"bike": 8, "ebike": 10},
            "客家文化園區": {"bike": 10, "ebike": 5},
            "池上大坡池": {"bike": 4, "ebike": 21},
            "池上三號運動公園": {"bike": 7, "ebike": 8},
            "日暉國際渡假村": {"bike": 14, "ebike": 8}
        },
        "早班配置": {
            "卑南鄉老人文康中心": {"bike": 3, "ebike": 0},
            "瑞源車站": {"bike": 4, "ebike": 0},
            "永安農產品銷售中心": {"bike": 4, "ebike": 0},
            "龍田大草坪": {"bike": 4, "ebike": 0},
            "鹿野車站": {"bike": 2, "ebike": 6},
            "福鹿山休閒農莊": {"bike": 4, "ebike": 3},
            "鹿野地區農會": {"bike": 4, "ebike": 0},
            "龍田老人會館": {"bike": 4, "ebike": 0},
            "永安環教中心": {"bike": 4, "ebike": 0},
            "關山車站": {"bike": 4, "ebike": 10},
            "米國學校": {"bike": 2, "ebike": 4},
            "關山工商": {"bike": 2, "ebike": 6},
            "關山鎮圖書館": {"bike": 2, "ebike": 4},
            "關山國小": {"bike": 2, "ebike": 4},
            "關山慈濟醫院": {"bike": 2, "ebike": 2},
            "池上車站": {"bike": 20, "ebike": 8},
            "池上郵局": {"bike": 9, "ebike": 3},
            "池上國中": {"bike": 5, "ebike": 2},
            "客家文化園區": {"bike": 5, "ebike": 3},
            "池上大坡池": {"bike": 2, "ebike": 10},
            "池上三號運動公園": {"bike": 5, "ebike": 3},
            "日暉國際渡假村": {"bike": 6, "ebike": 4}
        },
        "晚班配置": {
            "卑南鄉老人文康中心": {"bike": 3, "ebike": 0},
            "瑞源車站": {"bike": 4, "ebike": 0},
            "永安農產品銷售中心": {"bike": 4, "ebike": 0},
            "龍田大草坪": {"bike": 4, "ebike": 0},
            "鹿野車站": {"bike": 2, "ebike": 6},
            "福鹿山休閒農莊": {"bike": 4, "ebike": 3},
            "鹿野地區農會": {"bike": 4, "ebike": 0},
            "龍田老人會館": {"bike": 4, "ebike": 0},
            "永安環教中心": {"bike": 4, "ebike": 0},
            "關山車站": {"bike": 8, "ebike": 10},
            "米國學校": {"bike": 2, "ebike": 4},
            "關山工商": {"bike": 2, "ebike": 6},
            "關山鎮圖書館": {"bike": 2, "ebike": 4},
            "關山國小": {"bike": 2, "ebike": 4},
            "關山慈濟醫院": {"bike": 2, "ebike": 2},
            "池上車站": {"bike": 20, "ebike": 5},
            "池上郵局": {"bike": 8, "ebike": 3},
            "池上國中": {"bike": 5, "ebike": 2},
            "客家文化園區": {"bike": 5, "ebike": 3},
            "池上大坡池": {"bike": 3, "ebike": 10},
            "池上三號運動公園": {"bike": 5, "ebike": 3},
            "日暉國際渡假村": {"bike": 6, "ebike": 4}
        }
    }
}

D3_DATABASE = {
    "平日配置": {
        "夜班配置": {
            "富山護漁區": {"bike": 3, "ebike": 0},
            "成功漁港": {"bike": 7, "ebike": 0},
            "成功鎮公所": {"bike": 4, "ebike": 3},
            "成功豆花": {"bike": 4, "ebike": 3},
            "東管處停車場": {"bike": 3, "ebike": 0},
            "三仙台": {"bike": 7, "ebike": 0},
            "成功海濱公園": {"bike": 4, "ebike": 5},
            "都蘭舊糖廠": {"bike": 4, "ebike": 4},
            "東河消防隊": {"bike": 3, "ebike": 0},
            "成功消防隊": {"bike": 4, "ebike": 2}
        },
        "早班配置": {
            "富山護漁區": {"bike": 3, "ebike": 0},
            "成功漁港": {"bike": 7, "ebike": 0},
            "成功鎮公所": {"bike": 4, "ebike": 3},
            "成功豆花": {"bike": 4, "ebike": 3},
            "東管處停車場": {"bike": 3, "ebike": 0},
            "三仙台": {"bike": 7, "ebike": 0},
            "成功海濱公園": {"bike": 4, "ebike": 5},
            "都蘭舊糖廠": {"bike": 4, "ebike": 4},
            "東河消防隊": {"bike": 3, "ebike": 0},
            "成功消防隊": {"bike": 4, "ebike": 2}
        },
        "晚班配置": {
            "富山護漁區": {"bike": 3, "ebike": 0},
            "成功漁港": {"bike": 7, "ebike": 0},
            "成功鎮公所": {"bike": 4, "ebike": 3},
            "成功豆花": {"bike": 4, "ebike": 3},
            "東管處停車場": {"bike": 3, "ebike": 0},
            "三仙台": {"bike": 7, "ebike": 0},
            "成功海濱公園": {"bike": 4, "ebike": 5},
            "都蘭舊糖廠": {"bike": 4, "ebike": 4},
            "東河消防隊": {"bike": 3, "ebike": 0},
            "成功消防隊": {"bike": 4, "ebike": 2}
        }
    },
    "假日配置": {
        "夜班配置": {
            "富山護漁區": {"bike": 3, "ebike": 0},
            "成功漁港": {"bike": 7, "ebike": 0},
            "成功鎮公所": {"bike": 4, "ebike": 3},
            "成功豆花": {"bike": 4, "ebike": 3},
            "東管處停車場": {"bike": 3, "ebike": 0},
            "三仙台": {"bike": 7, "ebike": 0},
            "成功海濱公園": {"bike": 4, "ebike": 5},
            "都蘭舊糖廠": {"bike": 4, "ebike": 4},
            "東河消防隊": {"bike": 3, "ebike": 0},
            "成功消防隊": {"bike": 4, "ebike": 2}
        },
        "早班配置": {
            "富山護漁區": {"bike": 3, "ebike": 0},
            "成功漁港": {"bike": 8, "ebike": 0},
            "成功鎮公所": {"bike": 5, "ebike": 3},
            "成功豆花": {"bike": 5, "ebike": 3},
            "東管處停車場": {"bike": 3, "ebike": 0},
            "三仙台": {"bike": 10, "ebike": 0},
            "成功海濱公園": {"bike": 4, "ebike": 5},
            "都蘭舊糖廠": {"bike": 6, "ebike": 4},
            "東河消防隊": {"bike": 3, "ebike": 0},
            "成功消防隊": {"bike": 4, "ebike": 2}
        },
        "晚班配置": {
            "富山護漁區": {"bike": 3, "ebike": 0},
            "成功漁港": {"bike": 8, "ebike": 0},
            "成功鎮公所": {"bike": 5, "ebike": 3},
            "成功豆花": {"bike": 5, "ebike": 3},
            "東管處停車場": {"bike": 3, "ebike": 0},
            "三仙台": {"bike": 10, "ebike": 0},
            "成功海濱公園": {"bike": 4, "ebike": 5},
            "都蘭舊糖廠": {"bike": 6, "ebike": 4},
            "東河消防隊": {"bike": 3, "ebike": 0},
            "成功消防隊": {"bike": 4, "ebike": 2}
        }
    }
}

# ==========================================
# 2. 選單與資料處理 (修正新版寬度參數問題)
# ==========================================

col_s1, col_s2, col_s3 = st.columns(3)
with col_s1: selected_route = st.selectbox("🗺️ 路線", ["D2 縱谷線", "D3 海線"])
with col_s2: selected_config = st.selectbox("📅 配置", ["平日配置", "假日配置"])
with col_s3: selected_shift = st.selectbox("⏰ 班別", ["夜班配置", "早班配置", "晚班配置"])

base_config = D2_DATABASE[selected_config][selected_shift] if selected_route == "D2 縱谷線" else D3_DATABASE[selected_config][selected_shift]

st.subheader("✏️ 請輸入現場現況")
df = pd.DataFrame([{"場站名稱": s, "一般現況": v["bike"], "電輔現況": v["ebike"], "合約一般": v["bike"], "合約電輔": v["ebike"]} for s, v in base_config.items()])

# 將原本的 use_container_width 改為新版相容寫法 width="100%" (或直接用預設寬度避免新版錯誤)
edited_df = st.data_editor(df, hide_index=True, column_config={
    "場站名稱": st.column_config.TextColumn(disabled=True),
    "合約一般": st.column_config.NumberColumn(disabled=True),
    "合約電輔": st.column_config.NumberColumn(disabled=True)
})

# ==========================================
# 3. 調度分析邏輯 (改為獨立區塊集中輸出，確保渲染成功)
# ==========================================

if st.button("🚀 產出分析報表"):
    st.markdown("---")
    st.subheader("📋 調度分析結果")
    
    results = []
    for _, row in edited_df.iterrows():
        b_diff = int(row["一般現況"] - row["合約一般"])
        e_diff = int(row["電輔現況"] - row["合約電輔"])
        
        if b_diff != 0 or e_diff != 0:
            status_str = ""
            if b_diff != 0: 
                status_str += f"一般車 {'多 ' + str(b_diff) if b_diff > 0 else '缺 ' + str(abs(b_diff))}"
            if e_diff != 0:
                if status_str: status_str += " | "
                status_str += f"電輔車 {'多 ' + str(e_diff) if e_diff > 0 else '缺 ' + str(abs(e_diff))}"
                
            results.append({
                "場站名稱": row["場站名稱"],
                "調度狀況說明": status_str
            })
            
    if results:
        # 使用表格統整輸出，在新版 Streamlit 上最穩定
        result_df = pd.DataFrame(results)
        st.dataframe(result_df, hide_index=True)
    else:
        st.success("✨ 所有場站數量皆符合合約配置，目前不需要調度！")