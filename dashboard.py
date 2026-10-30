import streamlit as st
import pandas as pd
import os
from datetime import datetime
from io import BytesIO
import glob
import sqlite3


# =========================
# 업무관리 DB
# =========================

conn = sqlite3.connect(
    "dashboard.db",
    check_same_thread=False
)

cursor = conn.cursor()

cursor.execute("""
CREATE TABLE IF NOT EXISTS tasks (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    category TEXT,
    task TEXT,
    completed INTEGER DEFAULT 0
)
""")

conn.commit()



def get_latest_file(folder_path):

    files = (
        glob.glob(os.path.join(folder_path, "*.xlsx"))
        + glob.glob(os.path.join(folder_path, "*.xls"))
    )

    # 엑셀 임시파일 제외
    files = [
        f for f in files
        if not os.path.basename(f).startswith("~$")
    ]

    if not files:
        return None

    return max(files, key=os.path.getctime)

# ==================================================
# 기본 설정
# ==================================================

st.set_page_config(
    page_title="SCM 통합 대시보드",
    page_icon="📦",
    layout="wide"
)

from streamlit_autorefresh import st_autorefresh

st_autorefresh(
    interval=30000,
    key="datarefresh"
)

st.markdown("""
<style>

/* 메뉴 제목 */
section[data-testid="stSidebar"] label[data-testid="stWidgetLabel"] p {
    font-size: 40px !important;
    font-weight: bold !important;
}

/* 라디오 메뉴 글씨 */
section[data-testid="stSidebar"] div[role="radiogroup"] label p {
    font-size: 25px !important;
    font-weight: 400 !important;
}

</style>
""", unsafe_allow_html=True)

# ==================================================
# 품절 파일 로드
# ==================================================

oos_folder = r"C:\Users\195348\Desktop\python\DATA\품절율"

oos_file = get_latest_file(oos_folder)

if oos_file:
    oos_df = pd.read_excel(oos_file)
else:
    oos_df = pd.DataFrame()


# ==================================================
# 메뉴
# ==================================================
menu = st.sidebar.radio(
    "메뉴 선택",
    [
        "모니터링",
        "품절 분석",
        "업무정리",
        "재고분석",
        "과다재고",
        "대량주문분석",
        "상품상태확인",
        "비가용존 확인",
        "코드조회",
        "관리SKU",
        "판매량분석"
    ]
)

# ==================================================
# 모니터링
# ==================================================

if menu == "모니터링":

    st.title("📡 자동화 모니터링")

    # =========================
    # 금일 재고 현황
    # =========================

    try:

        재고폴더 = r"C:\Users\195348\Desktop\python\DATA\상품별 재고현황"

        재고파일목록 = sorted(
            glob.glob(os.path.join(재고폴더, "*.csv")),
            reverse=True
        )

        if len(재고파일목록) > 0:
 
            최신파일 = 재고파일목록[0]

            try:
                df_stock = pd.read_csv(
                    최신파일,
                    encoding="utf-8-sig",
                    dtype=str
                )
            except:
                df_stock = pd.read_csv(
                    최신파일,
                    encoding="cp949",
                    dtype=str
                )

            df_stock["총재고"] = pd.to_numeric(
                df_stock["총재고"],
                errors="coerce"
            ).fillna(0)

            df_stock["판매가"] = pd.to_numeric(
                df_stock["판매가"],
                errors="coerce"
            ).fillna(0)

            df_stock["재고금액"] = (
                df_stock["총재고"]
                * df_stock["판매가"]
            )

            df_stock["물류분류"] = (
                df_stock["물류분류"]
                .fillna("기타")
                .astype(str)
                .str[:3]
            )

            # 재고가 있는 상품만 SKU 계산
            sku_df = df_stock[df_stock["총재고"] > 0]
            
            sku_count = (
                sku_df
                .groupby("물류분류")["물류상품ID"]
                .nunique()
                .reset_index(name="재고SKU")
            )
            
            # 전체 집계
            summary = (
                df_stock
                .groupby("물류분류")
                .agg(
                    재고수량=("총재고", "sum"),
                    재고금액=("재고금액", "sum")
                )
                .reset_index()
            )
            
            # SKU 합치기
            summary = summary.merge(
                sku_count,
                on="물류분류",
                how="left"
            )
            
            summary["재고SKU"] = summary["재고SKU"].fillna(0).astype(int)
            
            summary = summary[
                ["물류분류", "재고SKU", "재고수량", "재고금액"]
            ]
            
            summary.columns = [
                "물류분류",
                "재고SKU",
                "재고수량",
                "재고금액"
            ]


            total_row = pd.DataFrame({
                "물류분류": ["합계"],
                "재고SKU": [summary["재고SKU"].sum()],
                "재고수량": [summary["재고수량"].sum()],
                "재고금액": [summary["재고금액"].sum()]
            })

            summary = pd.concat(
                [summary, total_row],
                ignore_index=True
            )

            st.subheader("📦 금일 재고 현황")

            st.dataframe(
                summary,
                use_container_width=True,
                hide_index=True
            )

        else:
            st.warning("상품별재고현황 파일이 없습니다.")

    except Exception as e:
        st.error(f"재고 현황 조회 오류 : {e}")

    st.divider()

    # =========================
    # 파일 모니터링
    # =========================

    CHECK_FILES = {

        "상품별 재고현황": (
            r"C:\Users\195348\Desktop\python\DATA\상품별 재고현황",
            "상품별재고현황"
        ),

        "위치별 재고현황": (
            r"C:\Users\195348\Desktop\python\DATA\위치별 재고현황",
            "위치별재고현황"
        ),

        "상품정보관리": (
            r"C:\Users\195348\Desktop\python\DATA\상품정보관리",
            "상품정보관리"
        ),

        "발주내역(17시)": (
            r"C:\Users\195348\Desktop\python\DATA\발주내역",
            "발주내역"
        )

    }

    rows = []

    today_str = datetime.today().strftime("%Y-%m-%d")

    for name, (folder, keyword) in CHECK_FILES.items():

        status = "❌ 없음"
        found_file = "-"

        if os.path.exists(folder):

            files = os.listdir(folder)

            for file in files:

                if (
                    today_str in file
                    and keyword in file
                ):

                    status = "✅ 정상"
                    found_file = file
                    break

        rows.append(
            {
                "항목": name,
                "상태": status,
                "확인파일": found_file
            }
        )

    monitor_df = pd.DataFrame(rows)

    st.dataframe(
        monitor_df,
        use_container_width=True,
        hide_index=True
    )

    st.divider()

    정상건수 = (
        monitor_df["상태"] == "✅ 정상"
    ).sum()

    확인건수 = (
        monitor_df["상태"] == "⚠ 확인필요"
    ).sum()

    오류건수 = (
        monitor_df["상태"] == "❌ 없음"
    ).sum()

    c1, c2, c3 = st.columns(3)

    c1.metric(
        "✅ 정상",
        정상건수
    )

    c2.metric(
        "⚠ 확인 필요",
        확인건수
    )

    c3.metric(
        "❌ 미확인",
        오류건수
    )

# ==================================================
# 과다재고
# ==================================================

elif menu == "과다재고":

    import pandas as pd
    import streamlit as st
    import os
    import glob

    from io import BytesIO

    st.title("📦 과다발주 분석 대시보드")

    # =====================
    # 최신 CSV 찾기
    # =====================
    def get_latest_file(folder):

        files = glob.glob(
            os.path.join(folder, "*.csv")
        )

        if not files:
            return None

        return max(
            files,
            key=os.path.getmtime
        )

    
    # =====================
    # CSV 읽기
    # =====================
    def read_csv_auto(file_path):

        try:
            return pd.read_csv(
                file_path,
                encoding="cp949"
            )

        except:

            try:
                return pd.read_csv(
                    file_path,
                    encoding="utf-8-sig"
                )

            except:

                return pd.read_csv(
                    file_path,
                    encoding="utf-8"
                )


    ORDER_FOLDER = r"C:\Users\195348\Desktop\python\DATA\발주내역"
    STOCK_FOLDER = r"C:\Users\195348\Desktop\python\DATA\상품별 재고현황"

    order_file = get_latest_file(ORDER_FOLDER)
    stock_file = get_latest_file(STOCK_FOLDER)

    st.write(
        "📄 최신 발주파일 :",
        os.path.basename(order_file)
        if order_file else "-"
    )

    st.write(
        "📦 최신 재고파일 :",
        os.path.basename(stock_file)
        if stock_file else "-"
    )

    if not order_file:
        st.error("발주파일이 없습니다.")
        st.stop()

    if not stock_file:
        st.error("재고파일이 없습니다.")
        st.stop()

    # =====================
    # 발주파일
    # =====================
    order_df = read_csv_auto(order_file)

    order_df.columns = (
        order_df.columns.astype(str)
        .str.strip()
    )

    order_df["오프라인상품번호"] = (
        order_df["오프라인상품번호"]
        .astype(str)
        .str.replace(".0", "", regex=False)
        .str.strip()
        .str.zfill(13)
    )

    order_df["발주"] = pd.to_numeric(
        order_df["발주"],
        errors="coerce"
    ).fillna(0)

    st.caption(
        f"발주 데이터 : {len(order_df):,}건"
    )

    order_group = (
        order_df
        .groupby(
            "오프라인상품번호",
            as_index=False
        )
        .agg(
            발주량=("발주", "sum"),

            확정자=(
                "확정자",
                lambda x: ",".join(
                    sorted(
                        set(
                            str(i).strip()
                            for i in x
                            if pd.notna(i)
                            and str(i).strip()
                        )
                    )
                )
            )
        )
    )
    
    # =====================
    # 재고파일
    # =====================
    stock_df = read_csv_auto(stock_file)

    stock_df.columns = (
        stock_df.columns.astype(str)
        .str.strip()
    )

    stock_df["오프라인상품번호"] = (
        stock_df["오프라인상품번호"]
        .astype(str)
        .str.replace(".0", "", regex=False)
        .str.strip()
        .str.zfill(13)
    )

    stock_df["재고일수"] = pd.to_numeric(
        stock_df["SSG 재고일수 (14일 기준)"],
        errors="coerce"
    ).fillna(0)

    st.caption(
        f"재고 데이터 : {len(stock_df):,}건"
    )

    # =====================
    # 병합
    # =====================
    result = stock_df.merge(
        order_group,
        on="오프라인상품번호",
        how="left"
    )

    result["발주량"] = (
        result["발주량"]
        .fillna(0)
    )

    result["확정자"] = (
        result["확정자"]
        .fillna("")
    )

    # =====================
    # 과다발주 조건
    # =====================
    over_df = result[
        (result["재고일수"] >= 50)
        &
        (result["발주량"] >= 1)
    ].copy()

    over_df["위험도"] = (
        over_df["재고일수"]
        * over_df["발주량"]
    )

    # =====================
    # KPI
    # =====================
    c1, c2, c3, c4 = st.columns(4)

    c1.metric(
        "과다발주 SKU",
        len(over_df)
    )

    c2.metric(
        "과다발주 발주수량",
        int(
            over_df["발주량"].sum()
        )
        if len(over_df) > 0
        else 0
    )

    c3.metric(
        "평균 재고일수",
        round(
            over_df["재고일수"].mean(),
            1
        )
        if len(over_df) > 0
        else 0
    )

    c4.metric(
        "총 위험도",
        f"{int(over_df['위험도'].sum()):,}"
        if len(over_df) > 0
        else "0"
    )

    st.divider()

    # =====================
    # TOP10
    # =====================
    st.subheader(
        "🔥 과다발주 TOP10"
    )

    top10 = (
        over_df
        .sort_values(
            "위험도",
            ascending=False
        )
        .head(10)
    )

    st.dataframe(
        top10[
            [
                "센터명",
                "오프라인상품번호",
                "상품명",
                "재고일수",
                "발주량",
                "확정자",
                "위험도"
            ]
        ],
        use_container_width=True
    )

    st.divider()

    # =====================
    # 전체 과다발주
    # =====================
    st.subheader(
        "📋 전체 과다발주"
    )

    st.dataframe(
        over_df[
            [
                "센터명",
                "오프라인상품번호",
                "상품명",
                "재고일수",
                "발주량",
                "확정자",
                "위험도"
            ]
        ],
        use_container_width=True,
        height=700
    )

    # =====================
    # 다운로드
    # =====================
    csv_data = over_df.to_csv(
        index=False,
        encoding="utf-8-sig"
    )

    st.download_button(
        label="📥 과다발주 CSV 다운로드",
        data=csv_data,
        file_name="과다발주.csv",
        mime="text/csv"
    )



elif menu == "상품상태확인":

    st.subheader("🚨 상태불량 상품 현황")

    import glob
    import os
    if st.button("🔄 데이터 새로고침"):
        st.cache_data.clear()
        st.rerun()

    @st.cache_data
    def load_data():

        # 상품정보관리 파일
        상품파일목록 = glob.glob(
            r"C:\Users\195348\Desktop\python\DATA\상품정보관리\*.csv"
        )

        상품파일 = max(
            상품파일목록,
            key=os.path.getctime
        )

        상품_df = pd.read_csv(
            상품파일,
            encoding="utf-8-sig",
            dtype=str
        )

        # 상품별 재고현황 파일
        재고파일목록 = glob.glob(
            r"C:\Users\195348\Desktop\python\DATA\상품별 재고현황\*.csv"
        )

        재고파일 = max(
            재고파일목록,
            key=os.path.getctime
        )

        재고_df = pd.read_csv(
            재고파일,
            encoding="utf-8-sig",
            dtype=str
        )

        return 상품_df, 재고_df

    상품_df, 재고_df = load_data()

    # 키 컬럼 정리
    상품_df["오프라인상품번호"] = (
        상품_df["오프라인상품번호"]
        .astype(str)
        .str.strip()
    )

    재고_df["오프라인상품번호"] = (
        재고_df["오프라인상품번호"]
        .astype(str)
        .str.strip()
    )

    # 판매중 제외
    상태_df = 상품_df[
        ~상품_df["몰판매상태"]
        .fillna("")
        .str.contains(r"(^|,)판매중(,|$)", regex=True)
    ].copy()


    # 재고정보 병합
    상태_df = 상태_df.merge(
        재고_df[
            [
                "오프라인상품번호",
                "총재고",
                "입하대기재고",
                "불량재고",
                "판매가능수량",
                "판매가"
            ]
        ],
        on="오프라인상품번호",
        how="left"
    )

    # 여기부터 추가
    상태_df["총재고"] = (
        상태_df["총재고"]
        .astype(str)
        .str.replace(",", "", regex=False)
    )
    
    상태_df["총재고"] = pd.to_numeric(
        상태_df["총재고"],
        errors="coerce"
    ).fillna(0)
    
    상태_df = 상태_df[
        상태_df["총재고"] > 0
    ]
    # 여기까지 추가
    상태_df["판매가능수량"] = pd.to_numeric(
        상태_df["판매가능수량"],
        errors="coerce"
    ).fillna(0)
    
    상태_df["입하대기재고"] = pd.to_numeric(
        상태_df["입하대기재고"],
        errors="coerce"
    ).fillna(0)
    
    상태_df["판매가_y"] = pd.to_numeric(
        상태_df["판매가_y"],
        errors="coerce"
    ).fillna(0)
    
    
    상태_df["재고금액"] = (
        상태_df["판매가_y"]
        * (
            상태_df["판매가능수량"]
        )
    )

    상태_df = 상태_df.rename(
        columns={
            "판매가_y": "판매가"
        }
    )

    상태_df["재고금액"] = 상태_df["재고금액"].round(0)
    

    

    # KPI
    c1, c2, c3 = st.columns(3)

    c1.metric(
        "상태불량 상품 수",
        f"{len(상태_df):,}"
    )

    c2.metric(
        "대표MD 수",
        상태_df["대표MD명"].nunique()
    )

    c3.metric(
        "SSG팀 수",
        상태_df["SSG팀"].nunique()
    )
    
    # ===== 추가 =====
    
    상태집계 = (
        상태_df.groupby("몰판매상태")
        .size()
        .reset_index(name="상품수")
        .sort_values("상품수", ascending=False)
    )
    
    md집계 = (
        상태_df.groupby("대표MD명")
        .size()
        .reset_index(name="상품수")
        .sort_values("상품수", ascending=False)
    )
    
    팀집계 = (
        상태_df.groupby("SSG팀")
        .size()
        .reset_index(name="상품수")
        .sort_values("상품수", ascending=False)
    )
    
    # ===== 가로 배치 =====
    
    st.subheader("📊 현황 분석")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("### 📋 상태불량 유형")
        st.dataframe(상태집계, use_container_width=True, hide_index=True)
    
    with col2:
        st.markdown("### 👤 대표MD별")
        st.dataframe(md집계, use_container_width=True, hide_index=True)
    
    with col3:
        st.markdown("### 🏢 SSG팀별")
        st.dataframe(팀집계, use_container_width=True, hide_index=True)
      # 상세조회
    st.subheader("🔍 상세 조회")

    선택MD = st.selectbox(
        "대표MD명 선택",
        ["전체"] + sorted(
            상태_df["대표MD명"]
            .fillna("미지정")
            .astype(str)
            .unique()
            .tolist()
        )
    )

    선택팀 = st.selectbox(
        "SSG팀 선택",
        ["전체"] + sorted(
            상태_df["SSG팀"]
            .fillna("미지정")
            .astype(str)
            .unique()
            .tolist()
        )
    )

    조회_df = 상태_df.copy()

    조회_df = 조회_df[
        조회_df["재고금액"] > 0
    ]

    if 선택MD != "전체":
        조회_df = 조회_df[
            조회_df["대표MD명"] == 선택MD
        ]

    if 선택팀 != "전체":
        조회_df = 조회_df[
            조회_df["SSG팀"] == 선택팀
        ]
    
    상세컬럼 = [
        "오프라인상품번호",
        "상품명",
        "몰판매상태",
        "대표MD명",
        "SSG팀",
        "총재고",
        "입하대기재고",
        "불량재고",
        "판매가능수량",
        "판매가",
        "재고금액"
    ]
    
    st.dataframe(
        조회_df[상세컬럼],
        use_container_width=True,
        hide_index=True
    )
   
#비가용존 확인
elif menu == "비가용존 확인":

    import glob
    st.header("📦 위치별 재고 현황")
    
    # ==========================
    # 위치별 재고현황 파일 읽기
    # ==========================
    
    위치파일목록 = glob.glob(
        r"C:\Users\195348\Desktop\python\DATA\위치별 재고현황\*.csv"
    )
    
    위치파일 = max(
        위치파일목록,
        key=os.path.getmtime
    )
    
    위치_df = pd.read_csv(
        위치파일,
        encoding="utf-8-sig",
        dtype=str
    )
    
    # ==========================
    # 상품별재고현황 파일 읽기
    # ==========================
    
    상품파일목록 = glob.glob(
        r"C:\Users\195348\Desktop\python\DATA\상품별 재고현황\*.csv"
    )
    
    상품파일 = max(
        상품파일목록,
        key=os.path.getmtime
    )
    
    상품_df = pd.read_csv(
        상품파일,
        encoding="utf-8-sig",
        dtype=str
    )
    
    # ==========================
    # 컬럼명 확인 후 수정
    # ==========================
    
    위치_df.columns = 위치_df.columns.str.strip()
    상품_df.columns = 상품_df.columns.str.strip()


    위치_df = 위치_df.rename(
    columns={
    "오프라인상품ID": "오프라인상품번호"
    }
    )

    # 층(B1층)
    위치_df = 위치_df[
        위치_df["층"] == "B1층"
    ]
    
    # 제외 존
    # 제외할 존 목록
    제외존 = [
        "DRY 누락반품존",
        "DRY입하존",
        "DRY고가상품불량보관존"
    ]
    
    # WET 시작 존 제외
    위치_df = 위치_df[
        ~위치_df["존"].str.startswith(
            "WET",
            na=False
        )
    ]
    
    # 특정 존 제외
    위치_df = 위치_df[
        ~위치_df["존"].isin(제외존)
    ]
    # ==========================
    # 판매가 컬럼 정리
    # ==========================
    
    if "판매가_y" in 상품_df.columns:
        상품_df = 상품_df.rename(
            columns={
                "판매가_y": "판매가"
            }
        )
    
    # ==========================
    # 숫자형 변환
    # ==========================
    
    위치_df["위치재고"] = pd.to_numeric(
        위치_df["위치재고"],
        errors="coerce"
    ).fillna(0)
    
    상품_df["판매가"] = pd.to_numeric(
        상품_df["판매가"],
        errors="coerce"
    ).fillna(0)
    
    # ==========================
    # 판매가 병합
    # ==========================
    
    위치_df = 위치_df.merge(
        상품_df[
            [
                "오프라인상품번호",
                "판매가"
            ]
        ],
        on="오프라인상품번호",
        how="left"
    )
    
    # ==========================
    # 금액 계산
    # ==========================
    
    위치_df["금액"] = (
        위치_df["위치재고"]
        * 상품_df["판매가"]
    )
    
    # ==========================
    # 존별 집계
    # ==========================
    
    존현황 = (
        위치_df
        .groupby("존")
        .agg(
            SKU수=("오프라인상품번호", "nunique"),
            위치재고=("위치재고", "sum"),
            금액=("금액", "sum")
        )
        .reset_index()
    )
    
    존현황 = 존현황.sort_values(
        "금액",
        ascending=False
    )
    
    # ==========================
    # KPI
    # ==========================
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.metric(
            "총 SKU",
            f"{존현황['SKU수'].sum():,}"
        )
    
    with col2:
        st.metric(
            "총 재고수량",
            f"{int(존현황['위치재고'].sum()):,}"
        )
    
    with col3:
        st.metric(
            "총 재고금액",
            f"{int(존현황['금액'].sum()):,}원"
        )
    
    # ==========================
    # 존별 현황
    # ==========================
    
    st.subheader("존별 재고 현황")
    
    st.dataframe(
        존현황.style.format({
            "SKU수": "{:,.0f}",
            "위치재고": "{:,.0f}",
            "금액": "{:,.0f}"
        }),
        use_container_width=True,
        hide_index=True
    )
    
    # ==========================
    # 존 선택
    # ==========================
    
    선택존 = st.selectbox(
        "존 선택",
        ["전체"] + sorted(
            위치_df["존"].dropna().unique()
        )
    )
    
    # ==========================
    # 상세조회
    # ==========================
    
    if 선택존 != "전체":
    
        상세_df = 위치_df[
            위치_df["존"] == 선택존
        ].copy()
    
        상세_df = 상세_df.sort_values(
            "금액",
            ascending=False
        )
    
        st.subheader(
            f"📋 {선택존} 상세 현황"
        )
    
        st.dataframe(
            상세_df,
            use_container_width=True,
            hide_index=True
        )
    
        csv = 상세_df.to_csv(
            index=False,
            encoding="utf-8-sig"
        )
    
        st.download_button(
            label="📥 상세내역 다운로드",
            data=csv,
            file_name=f"{선택존}_상세현황.csv",
            mime="text/csv"
        )
    
    
    
    
      
elif menu == "코드조회":

    import glob

    재고폴더 = r"C:\Users\195348\Desktop\python\DATA\상품별 재고현황"

    try:
        파일목록 = glob.glob(
            os.path.join(재고폴더, "*.csv")
        )

        최신파일 = max(
            파일목록,
            key=os.path.getctime
        )

        df = pd.read_csv(
            최신파일,
            dtype=str,
            encoding="utf-8-sig"
        )

    except Exception as e:
        st.error(f"재고파일 읽기 실패 : {e}")
        st.stop()

    st.subheader("🔍 상품 코드 조회")

    입력코드 = st.text_area(
        "코드를 여러 개 입력하세요",
        height=200
    )

    if st.button("조회"):

        조회목록 = [
            str(x).strip()
            for x in 입력코드.splitlines()
            if str(x).strip()
        ]

        결과 = []

        for _, row in df.iterrows():

            검색대상 = set()

            검색대상.add(
                str(row.get("물류상품ID", "")).strip()
            )

            검색대상.add(
                str(row.get("오프라인상품번호", "")).strip()
            )

            연동ID = str(
                row.get("연동상품ID", "")
            )

            for x in 연동ID.split(","):
                if x.strip():
                    검색대상.add(x.strip())

            바코드 = str(
                row.get("상품바코드", "")
            )

            for x in 바코드.split(","):
                if x.strip():
                    검색대상.add(x.strip())

            if any(
                code in 검색대상
                for code in 조회목록
            ):

                결과.append(
                    {
                        "물류상품ID": row.get("물류상품ID", ""),
                        "오프라인상품번호": row.get("오프라인상품번호", ""),
                        "상품명": row.get("상품명", ""),
                        "총재고": row.get("총재고", ""),
                        "입하대기재고": row.get("입하대기재고", ""),
                        "판매가능수량": row.get("판매가능수량", "")
                    }
                )

        결과df = pd.DataFrame(결과)

        st.success(
            f"{len(결과df):,}건 조회"
        )

        st.dataframe(
            결과df,
            use_container_width=True,
            hide_index=True
        )


elif menu == "관리SKU":
    st.subheader("📦 관리 SKU 현황")

    import pandas as pd

    # =========================
    # 파일 로드
    # =========================
    import glob
    import os
    관리파일목록 = glob.glob(
        r"C:\Users\195348\Desktop\python\DATA\관리sku\관리sku.csv"
    )
    관리파일 = max(
        관리파일목록,
        key=os.path.getmtime
    )
    
    관리_df = pd.read_csv(
        관리파일,
        encoding="utf-8-sig",
        dtype=str
    )

    상품파일목록 = glob.glob(
        r"C:\Users\195348\Desktop\python\DATA\상품별 재고현황\*.csv"
    )
    상품파일 = max(
        상품파일목록,
        key=os.path.getmtime
    )
    
    상품_df = pd.read_csv(
        상품파일,
        encoding="utf-8-sig",
        dtype=str
    )


    위치파일목록 = glob.glob(
        r"C:\Users\195348\Desktop\python\DATA\위치별 재고현황\*.csv"
    )
    위치파일 = max(
        위치파일목록,
        key=os.path.getmtime
    )
    
    위치_df = pd.read_csv(
        위치파일,
        encoding="utf-8-sig",
        dtype=str
    )

    # =========================
    # 컬럼명 공백 제거
    # =========================
    관리_df.columns = 관리_df.columns.str.strip()
    상품_df.columns = 상품_df.columns.str.strip()
    위치_df.columns = 위치_df.columns.str.strip()
    
    # Unnamed 컬럼 제거
    관리_df = 관리_df.loc[
        :,
        ~관리_df.columns.str.contains("^Unnamed")
    ]
    
    # 오프라인상품번호가 없는 행 제거
    관리_df = 관리_df.dropna(
        subset=["오프라인상품번호"]
    )
    
    관리_df["오프라인상품번호"] = (
        관리_df["오프라인상품번호"]
        .astype(str)
        .str.strip()
    )
    
    관리_df = 관리_df[
        관리_df["오프라인상품번호"] != ""
    ]

    # =========================
    # 숫자형 변환
    # =========================
    숫자컬럼 = [
        "총재고",
        "입하대기재고",
        "불량재고",
        "판매가능수량",
        "판매가"
    ]

    for col in 숫자컬럼:
        if col in 상품_df.columns:
            상품_df[col] = pd.to_numeric(
                상품_df[col],
                errors="coerce"
            ).fillna(0)

    # =========================
    # 상품별재고현황 매칭
    # =========================
    관리현황_df = 관리_df.merge(
        상품_df[
            [
                "오프라인상품번호",
                "총재고",
                "입하대기재고",
                "불량재고",
                "판매가능수량",
                "판매가"
            ]
        ],
        how="left",
        on="오프라인상품번호"
    )

    관리현황_df["재고금액"] = (
        관리현황_df["총재고"]
        * 관리현황_df["판매가"]
    )
# =========================
# 구분 필터
# =========================
    선택구분 = st.selectbox(
        "구분 선택",
        ["전체"] +
        sorted(
            관리현황_df["구분"]
            .dropna()
            .unique()
        )
    )
    
    if 선택구분 != "전체":
        관리현황_df = 관리현황_df[
            관리현황_df["구분"] == 선택구분
        ]
    
    # =========================
    # KPI
    # =========================
    col1, col2, col3 = st.columns(3)
    
    col1.metric(
        "관리 SKU",
        f"{len(관리현황_df):,}"
    )
    
    col2.metric(
        "총 재고",
        f"{관리현황_df['총재고'].sum():,.0f}"
    )
    
    col3.metric(
        "총 재고금액",
        f"{관리현황_df['재고금액'].sum():,.0f}원"
    )
    # =========================
    # 관리 SKU 현황
    # =========================
    st.dataframe(
        관리현황_df,
        use_container_width=True,
        hide_index=True
    )

    # =========================
    # SKU 선택
    # =========================
    선택상품 = st.selectbox(
        "상세조회 상품",
        관리현황_df["상품명"].fillna("")
    )

    if 선택상품:

        선택번호 = (
            관리현황_df.loc[
                관리현황_df["상품명"] == 선택상품,
                "오프라인상품번호"
            ]
            .iloc[0]
        )

        상세위치_df = 위치_df[
            위치_df["오프라인상품ID"]
            .astype(str)
            == str(선택번호)
        ].copy()

        st.markdown("### 📍 위치별 재고현황")

        st.dataframe(
            상세위치_df,
            use_container_width=True,
            hide_index=True
        )

elif menu == "판매량분석":

    st.header("📊 판매량 분석")

    # ==========================
    # 파일 읽기
    # ==========================
    df = pd.read_csv(
        r"C:\Users\195348\Desktop\python\DATA\올랩\주문수량.csv",
        sep="\t",
        dtype=str,
        encoding="cp949"
    )

    # 상품코드 컬럼명
    상품코드컬럼 = df.columns[0]

    # 날짜 컬럼
    date_cols = df.columns[1:]

    # ==========================
    # 상품 선택
    # ==========================
    상품목록 = sorted(df[상품코드컬럼].astype(str).unique())

    상품입력 = st.text_area(
        "상품코드 입력",
        height=200,
        placeholder="""0000001454918
    """
    )

    선택상품 = [
        x.strip()
        for x in 상품입력.split("\n")
        if x.strip()
    ]

    st.write(f"입력 상품수 : {len(선택상품):,}")

    if 선택상품:

        선택데이터 = df[
            df[상품코드컬럼].astype(str).isin(선택상품)
        ]

        # ==========================
        # 행 → 열 변환
        # ==========================
        melt_df = 선택데이터.melt(
            id_vars=[상품코드컬럼],
            value_vars=date_cols,
            var_name="일자",
            value_name="판매량"
        )

        melt_df["판매량"] = pd.to_numeric(
            melt_df["판매량"],
            errors="coerce"
        ).fillna(0)

        melt_df["일자"] = (
            melt_df["일자"]
            .astype(str)
            .str.extract(r'(\d{2}\.\d{2}\.\d{2})')[0]
        )
        
        melt_df["일자"] = pd.to_datetime(
            melt_df["일자"],
            format="%y.%m.%d",
            errors="coerce"
        )

        melt_df = melt_df.dropna(subset=["일자"])


        if melt_df.empty:
            st.error("날짜 컬럼을 인식하지 못했습니다.")
            st.stop()

        # ==========================
        # 기간 선택
        # ==========================
        min_date = melt_df["일자"].min()
        max_date = melt_df["일자"].max()

        col1, col2 = st.columns(2)

        with col1:
            시작일 = st.date_input(
                "시작일",
                min_date
            )

        with col2:
            종료일 = st.date_input(
                "종료일",
                max_date
            )

        결과 = melt_df[
            (melt_df["일자"] >= pd.to_datetime(시작일))
            &
            (melt_df["일자"] <= pd.to_datetime(종료일))
        ]

        # ==========================
        # KPI
        # ==========================
        총판매 = int(결과["판매량"].sum())

        일평균 = round(
            결과.groupby("일자")["판매량"].sum().mean(),
            1
        )

        판매일수 = int(
            (결과["판매량"] > 0).sum()
        )

        최고일 = (
            결과.groupby("일자")["판매량"]
            .sum()
            .idxmax()
        )

        c1, c2, c3, c4 = st.columns(4)

        c1.metric("총 판매량", f"{총판매:,}")
        c2.metric("일평균", f"{일평균:,}")
        c3.metric("판매 발생일", 판매일수)

        if pd.notna(최고일):
            c4.metric(
                "최대 판매일",
                최고일.strftime("%Y-%m-%d")
            )

        st.divider()

        # ==========================
        # 일자별 판매 추이
        # ==========================
        st.subheader("📈 일자별 판매 추이")

        일별 = (
            결과.groupby("일자")["판매량"]
            .sum()
            .reset_index()
        )

        st.line_chart(
            일별.set_index("일자")
        )

        # ==========================
        # 상품별 판매량
        # ==========================
        st.subheader("🏆 상품별 판매량")

        상품별 = (
            결과.groupby(상품코드컬럼)["판매량"]
            .sum()
            .sort_values(ascending=False)
        )

        st.bar_chart(상품별)

        # ==========================
        # 상세 데이터
        # ==========================
        st.subheader("📋 상세 데이터")

        st.dataframe(
            결과.sort_values("일자"),
            use_container_width=True
        )

elif menu == "대량주문분석":

    st.header("📦 전일 대량주문 고객 TOP100")

    uploaded_file = st.file_uploader(
        r"C:\Users\195348\Desktop\python\DATA\과다주문분석",
        type=["xlsx", "xls", "csv"]
    )

    if uploaded_file:

    # 파일 읽기
        if uploaded_file.name.endswith(".csv"):
            df_order = pd.read_csv(
                uploaded_file,
                header=3
            )
            st.write(uploaded_file.name)
        else:
            df_order = pd.read_excel(
                uploaded_file,
                header=3
            )
    
    # 컬럼명 공백 제거
        df_order.columns = (
            df_order.columns
            .astype(str)
            .str.strip()
        )
    
    # 숫자 변환

        for col in ["최근1일", "최근2일", "최근3일"]:
            df_order[col] = pd.to_numeric(
                df_order[col],
                errors="coerce"
            ).fillna(0)

        # 주문번호 생성
        df_order["주문번호"] = (
            df_order["주문번호_앞10자리"].astype(str)
            + "-"
            + df_order["주문번호_뒤4자리"].astype(str)
        )

        df_order["합계"] = (
            df_order["최근1일"]
            + df_order["최근2일"]
            + df_order["최근3일"]
        )
        
        # 대분류 선택
        categories = st.multiselect(
            "대분류 선택",
            options=sorted(
                df_order["표준카테대분류"].dropna().unique()
            )
        )
        
        if categories:
            df_order = df_order[
                df_order["표준카테대분류"].isin(categories)
            ]

        # 상품ID 다중 입력
        product_ids_text = st.text_area(
            "상품ID 입력 (여러 개 가능)",
            height=120,
            placeholder="""0000000000000
        """
        )
        
        if product_ids_text.strip():
        
            # 줄바꿈 기준 분리
            product_ids = [
                x.strip()
                for x in product_ids_text.splitlines()
                if x.strip()
            ]
        
            df_order = df_order[
                df_order["이마트오프라인상품ID"]
                .astype(str)
                .isin(product_ids)
            ]

        # 전일 구매건 추출
        result = (
            df_order[
                [
                    "표준카테대분류",
                    "주문번호",
                    "이마트오프라인상품ID",
                    "이마트오프라인상품",
                    "최근1일",
                    "최근2일",
                    "최근3일",
                    "합계"
                ]
            ]
            .sort_values(
                "합계",
                ascending=[False]
            )
            .head(100)
            .copy()
        )

        result.insert(
            0,
            "순위",
            range(1, len(result) + 1)
        )

        st.dataframe(
            result,
            use_container_width=True,
            hide_index=True
        )
# ==================================================
# 재고분석
# ==================================================
elif menu == "재고분석":

    st.title("📊 재고자산 관제탑 V1")

    # ==================================================
    # 상품별 재고현황
    # ==================================================

    stock_folder = r"C:\Users\195348\Desktop\python\DATA\상품별 재고현황"

    stock_file = sorted(
        glob.glob(os.path.join(stock_folder, "*.csv")),
        reverse=True
    )[0]

    try:
        item_df = pd.read_csv(
            stock_file,
            encoding="utf-8-sig",
            dtype=str
        )
    except:
        item_df = pd.read_csv(
            stock_file,
            encoding="cp949",
            dtype=str
        )

    # ==================================================
    # 위치별 재고현황
    # ==================================================

    loc_folder = r"C:\Users\195348\Desktop\python\DATA\위치별 재고현황"

    loc_file = sorted(
        glob.glob(os.path.join(loc_folder, "*.csv")),
        reverse=True
    )[0]

    try:
        loc_df = pd.read_csv(
            loc_file,
            encoding="utf-8-sig",
            dtype=str
        )
    except:
        loc_df = pd.read_csv(
            loc_file,
            encoding="cp949",
            dtype=str
        )

    # ==================================================
    # B1층 제외
    # ==================================================

    loc_df = loc_df[
        ~loc_df["층"].astype(str).str.contains(
            "B1",
            na=False
        )
    ].copy()

    # ==================================================
    # JOIN KEY
    # ==================================================

    item_df["오프라인상품번호"] = (
        item_df["오프라인상품번호"]
        .astype(str)
        .str.replace(".0", "", regex=False)
        .str.strip()
    )

    loc_df["오프라인상품ID"] = (
        loc_df["오프라인상품ID"]
        .astype(str)
        .str.replace(".0", "", regex=False)
        .str.strip()
    )

    # ==================================================
    # JOIN
    # ==================================================

    master_df = pd.merge(
        loc_df,
        item_df[
            [
                "오프라인상품번호",
                "상품명",
                "총재고",
                "판매가",
                "무판매일수",
                "SSG 재고일수 (14일 기준)",
                "SSG 재고일수 (85일 기준)",
                "물류분류",
                "물류상품등급"
            ]
        ],
        left_on="오프라인상품ID",
        right_on="오프라인상품번호",
        how="left"
    )

    # ==================================================
    # 숫자형 변환
    # ==================================================

    number_cols = [
        "위치재고",
        "총재고",
        "판매가",
        "무판매일수",
        "SSG 재고일수 (14일 기준)",
        "SSG 재고일수 (85일 기준)",
        "유통기한잔여일수"
    ]

    for col in number_cols:

        master_df[col] = pd.to_numeric(
            master_df[col],
            errors="coerce"
        ).fillna(0)

    # ==================================================
    # 재고금액
    # ==================================================

    master_df["재고금액"] = (
        master_df["위치재고"]
        * master_df["판매가"]
    )

    # ==================================================
    # KPI
    # ==================================================

    total_amount = master_df["재고금액"].sum()

    total_sku = (
        master_df["오프라인상품ID"]
        .nunique()
    )

    over_amount = (
        master_df.loc[
            master_df["SSG 재고일수 (14일 기준)"] >= 90,
            "재고금액"
        ]
        .sum()
    )

    dead_amount = (
        master_df.loc[
            master_df["무판매일수"] >= 90,
            "재고금액"
        ]
        .sum()
    )

    danger_amount = (
        master_df.loc[
            master_df["유통기한잔여일수"] <= 30,
            "재고금액"
        ]
        .sum()
    )

    c1, c2, c3, c4, c5 = st.columns(5)

    with c1:
        st.metric(
            "총재고금액",
            f"{total_amount:,.0f}원"
        )

    with c2:
        st.metric(
            "SKU",
            f"{total_sku:,}"
        )

    with c3:
        st.metric(
            "과재고금액",
            f"{over_amount:,.0f}원"
        )

    with c4:
        st.metric(
            "무판매금액",
            f"{dead_amount:,.0f}원"
        )

    with c5:
        st.metric(
            "폐기위험금액",
            f"{danger_amount:,.0f}원"
        )

    st.divider()

    # ==================================================
    # 탭
    # ==================================================

    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
        [
            "재고금액 TOP20",
            "재고수량 TOP20",
            "과재고",
            "무판매",
            "폐기위험",
            "존별 재고자산"
        ]
    )

    # ==================================================
    # 재고금액 TOP20
    # ==================================================

    with tab1:

        top_amount = (
            master_df
            .groupby(
                ["오프라인상품ID", "상품명"],
                as_index=False
            )
            .agg(
                재고수량=("위치재고", "sum"),
                재고금액=("재고금액", "sum")
            )
            .sort_values(
                "재고금액",
                ascending=False
            )
            .head(20)
        )

        st.dataframe(
            top_amount,
            use_container_width=True,
            hide_index=True
        )

    # ==================================================
    # 재고수량 TOP20
    # ==================================================

    with tab2:

        top_qty = (
            master_df
            .groupby(
                ["오프라인상품ID", "상품명"],
                as_index=False
            )
            .agg(
                재고수량=("위치재고", "sum"),
                재고금액=("재고금액", "sum")
            )
            .sort_values(
                "재고수량",
                ascending=False
            )
            .head(20)
        )

        st.dataframe(
            top_qty,
            use_container_width=True,
            hide_index=True
        )

    # ==================================================
    # 과재고
    # ==================================================

    with tab3:

        over_df = (
            master_df[
                master_df[
                    "SSG 재고일수 (14일 기준)"
                ] >= 90
            ]
            .sort_values(
                "재고금액",
                ascending=False
            )
        )

        st.dataframe(
            over_df[
                [
                    "상품명",
                    "SSG 재고일수 (14일 기준)",
                    "위치재고",
                    "재고금액",
                    "존",
                    "보관위치(셀)"
                ]
            ],
            use_container_width=True,
            hide_index=True
        )

    # ==================================================
    # 무판매
    # ==================================================

    with tab4:

        dead_df = (
            master_df[
                master_df["무판매일수"] >= 90
            ]
            .sort_values(
                "재고금액",
                ascending=False
            )
        )

        st.dataframe(
            dead_df[
                [
                    "상품명",
                    "무판매일수",
                    "위치재고",
                    "재고금액",
                    "존",
                    "보관위치(셀)"
                ]
            ],
            use_container_width=True,
            hide_index=True
        )

    # ==================================================
    # 폐기위험
    # ==================================================

    with tab5:

        danger_df = (
            master_df[
                master_df["유통기한잔여일수"] <= 30
            ]
            .sort_values(
                "재고금액",
                ascending=False
            )
        )

        st.dataframe(
            danger_df[
                [
                    "상품명",
                    "유통기한잔여일수",
                    "위치재고",
                    "재고금액",
                    "존",
                    "보관위치(셀)"
                ]
            ],
            use_container_width=True,
            hide_index=True
        )

    # ==================================================
    # 존별 재고자산
    # ==================================================

    with tab6:

        zone_df = (
            master_df
            .groupby(
                "존",
                as_index=False
            )
            .agg(
                SKU=("오프라인상품ID", "nunique"),
                재고수량=("위치재고", "sum"),
                재고금액=("재고금액", "sum")
            )
            .sort_values(
                "재고금액",
                ascending=False
            )
        )

        st.dataframe(
            zone_df,
            use_container_width=True,
            hide_index=True
        )

        zone = st.selectbox(
            "존 선택",
            zone_df["존"].tolist()
        )

        zone_detail = (
            master_df[
                master_df["존"] == zone
            ]
            .sort_values(
                "재고금액",
                ascending=False
            )
        )

        st.dataframe(
            zone_detail[
                [
                    "상품명",
                    "위치재고",
                    "재고금액",
                    "SSG 재고일수 (14일 기준)",
                    "유통기한잔여일수",
                    "보관위치(셀)"
                ]
            ],
            use_container_width=True,
            hide_index=True
        )



elif menu == "업무정리":

    st.title("📝 업무정리")

    tabs = st.tabs([
        "일일 업무",
        "주간 업무",
        "단발성 업무"
    ])

    categories = {
        "일일 업무": tabs[0],
        "주간 업무": tabs[1],
        "단발성 업무": tabs[2]
    }

    for category, tab in categories.items():

        with tab:

            st.subheader(category)

            # 신규 업무 추가
            new_task = st.text_input(
                f"{category} 추가",
                key=f"input_{category}"
            )

            if st.button(
                f"{category} 등록",
                key=f"btn_{category}"
            ):

                if new_task:

                    cursor.execute("""
                        INSERT INTO tasks
                        (category, task)
                        VALUES (?, ?)
                    """, (category, new_task))

                    conn.commit()
                    st.rerun()

            # 업무 조회
            task_df = pd.read_sql(
                f"""
                SELECT *
                FROM tasks
                WHERE category='{category}'
                ORDER BY id DESC
                """,
                conn
            )

            total = len(task_df)

            completed = len(
                task_df[
                    task_df["completed"] == 1
                ]
            )

            percent = (
                completed / total * 100
                if total > 0 else 0
            )

            st.metric(
                "완료율",
                f"{percent:.1f}%"
            )

            st.progress(percent / 100)

            st.divider()

            for _, row in task_df.iterrows():

                col1, col2 = st.columns([8, 1])

                with col1:

                    checked = st.checkbox(
                        row["task"],
                        value=bool(row["completed"]),
                        key=f"task_{row['id']}"
                    )

                    if checked != bool(row["completed"]):

                        cursor.execute("""
                            UPDATE tasks
                            SET completed=?
                            WHERE id=?
                        """,
                        (
                            int(checked),
                            row["id"]
                        ))

                        conn.commit()
                        st.rerun()

                with col2:

                    if st.button(
                        "❌",
                        key=f"del_{row['id']}"
                    ):

                        cursor.execute("""
                            DELETE FROM tasks
                            WHERE id=?
                        """,
                        (row["id"],)
                        )

                        conn.commit()
                        st.rerun()
elif menu == "품절 분석":

    st.header("📉 품절 분석 대시보드")

    df = oos_df.copy()

    st.write(df.columns.tolist())
    st.stop()
    # =========================
    # 데이터 정리
    # =========================
    numeric_cols = ["운영시간수", "품절시간수", "품절율"]

    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors="coerce")

    df = df.fillna(0)

    # 품절율이 0~1 형태이면 100 곱하기
    if "품절율" in df.columns:
        if df["품절율"].max() <= 1:
            df["품절율"] = df["품절율"] * 100

    # =========================
    # KPI
    # =========================

    total_oos_rate = (
        df["품절시간수"].sum()
        / df["운영시간수"].sum()
        * 100
    ) if df["운영시간수"].sum() > 0 else 0

    product_cnt = df["오프라인상품번호"].nunique()

    vendor_cnt = df["공급업체명"].nunique()

    avg_oos_rate = df["품절율"].mean()

    # TOP 상품
    top_item_name = "-"

    item_tmp = (
        df.groupby(
            ["오프라인상품번호", "오프라인상품명"],
            as_index=False
        )
        .agg({
            "품절시간수": "sum",
            "운영시간수": "sum"
        })
    )

    item_tmp["품절율"] = (
        item_tmp["품절시간수"]
        / item_tmp["운영시간수"]
        * 100
    )

    if not item_tmp.empty:
        top_item_name = item_tmp.sort_values(
            "품절율",
            ascending=False
        ).iloc[0]["오프라인상품명"]

    # TOP 업체
    top_vendor_name = "-"

    vendor_tmp = (
        df.groupby(
            "공급업체명",
            as_index=False
        )
        .agg({
            "품절시간수": "sum",
            "운영시간수": "sum"
        })
    )

    vendor_tmp["품절율"] = (
        vendor_tmp["품절시간수"]
        / vendor_tmp["운영시간수"]
        * 100
    )

    if not vendor_tmp.empty:
        top_vendor_name = vendor_tmp.sort_values(
            "품절율",
            ascending=False
        ).iloc[0]["공급업체명"]

    # =========================
    # KPI 카드
    # =========================

    c1, c2, c3, c4, c5, c6 = st.columns(6)

    c1.metric(
        "전체 품절율",
        f"{total_oos_rate:.2f}%"
    )

    c2.metric(
        "품절 상품수",
        f"{product_cnt:,}"
    )

    c3.metric(
        "품절 업체수",
        f"{vendor_cnt:,}"
    )

    c4.metric(
        "평균 품절율",
        f"{avg_oos_rate:.2f}%"
    )

    c5.metric(
        "최고 품절 상품",
        top_item_name
    )

    c6.metric(
        "최고 품절 업체",
        top_vendor_name
    )

    st.markdown("---")

    # =========================
    # 탭 구성
    # =========================

    tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs(
        [
            "TOP 품절 상품",
            "TOP 품절 업체",
            "미납 의심",
            "품절 사유",
            "카테고리 분석",
            "담당자 분석"
        ]
    )

    # ====================================================
    # TOP 품절 상품
    # ====================================================

    with tab1:

        top_item = (
            df.groupby(
                ["오프라인상품번호", "오프라인상품명"],
                as_index=False
            )
            .agg({
                "품절시간수": "sum",
                "운영시간수": "sum"
            })
        )

        top_item["품절율"] = (
            top_item["품절시간수"]
            / top_item["운영시간수"]
            * 100
        )

        top_item = top_item.sort_values(
            "품절율",
            ascending=False
        ).head(20)

        st.subheader("TOP20 품절 상품")

        st.dataframe(
            top_item,
            use_container_width=True
        )

    # ====================================================
    # TOP 품절 업체
    # ====================================================

    with tab2:

        top_vendor = (
            df.groupby(
                "공급업체명",
                as_index=False
            )
            .agg({
                "품절시간수": "sum",
                "운영시간수": "sum"
            })
        )

        top_vendor["품절율"] = (
            top_vendor["품절시간수"]
            / top_vendor["운영시간수"]
            * 100
        )

        top_vendor = top_vendor.sort_values(
            "품절율",
            ascending=False
        ).head(20)

        st.subheader("TOP20 품절 업체")

        st.dataframe(
            top_vendor,
            use_container_width=True
        )

    # ====================================================
    # 미납 의심
    # ====================================================

    with tab3:

        col1, col2 = st.columns(2)

        with col1:

            st.subheader("미납 의심 업체")

            vendor_risk = (
                df.groupby(
                    "공급업체명"
                )
                .agg(
                    상품수=("오프라인상품번호", "nunique"),
                    평균품절율=("품절율", "mean")
                )
                .reset_index()
            )

            vendor_risk = vendor_risk[
                (vendor_risk["상품수"] >= 5)
                &
                (vendor_risk["평균품절율"] >= 20)
            ]

            vendor_risk = vendor_risk.sort_values(
                "평균품절율",
                ascending=False
            )

            st.dataframe(
                vendor_risk,
                use_container_width=True
            )

        with col2:

            st.subheader("미납 의심 상품")

            risk_item = (
                df.groupby(
                    ["오프라인상품번호", "오프라인상품명"]
                )
                .agg(
                    평균품절율=("품절율", "mean")
                )
                .reset_index()
            )

            risk_item = risk_item[
                risk_item["평균품절율"] >= 50
            ]

            risk_item = risk_item.sort_values(
                "평균품절율",
                ascending=False
            )

            st.dataframe(
                risk_item,
                use_container_width=True
            )

    # ====================================================
    # 품절 사유
    # ====================================================

    with tab4:

        st.subheader("품절 사유 분석")

        reason_df = (
            df.groupby("품절사유")
            .size()
            .reset_index(name="건수")
            .sort_values(
                "건수",
                ascending=False
            )
        )

        st.dataframe(
            reason_df,
            use_container_width=True
        )

        st.bar_chart(
            reason_df.set_index("품절사유")["건수"]
        )

    # ====================================================
    # 카테고리 분석
    # ====================================================

    with tab5:

        st.subheader("대분류별 품절율")

        cat_df = (
            df.groupby(
                "오프라인대분류명",
                as_index=False
            )
            .agg(
                운영시간=("운영시간수", "sum"),
                품절시간=("품절시간수", "sum")
            )
        )

        cat_df["품절율"] = (
            cat_df["품절시간"]
            / cat_df["운영시간"]
            * 100
        )

        cat_df = cat_df.sort_values(
            "품절율",
            ascending=False
        )

        st.dataframe(
            cat_df,
            use_container_width=True
        )

        st.bar_chart(
            cat_df.set_index(
                "오프라인대분류명"
            )["품절율"]
        )

    # ====================================================
    # 담당자 분석
    # ====================================================

    with tab6:

        st.subheader("담당자별 품절율")

        manager_df = (
            df.groupby(
                "오프라인담당명",
                as_index=False
            )
            .agg(
                운영시간=("운영시간수", "sum"),
                품절시간=("품절시간수", "sum")
            )
        )

        manager_df["품절율"] = (
            manager_df["품절시간"]
            / manager_df["운영시간"]
            * 100
        )

        manager_df = manager_df.sort_values(
            "품절율",
            ascending=False
        )

        st.dataframe(
            manager_df,
            use_container_width=True
        )

        st.bar_chart(
            manager_df.set_index(
                "오프라인담당명"
            )["품절율"]
        )
