import io
from pathlib import Path

import pandas as pd
import streamlit as st


# =========================
# 基本設定
# =========================
st.set_page_config(
    page_title="向精神薬チェックツール",
    page_icon="📦",
    layout="wide",
)

APP_TITLE = "向精神薬チェックツール"
OUTPUT_SHEET_NAME = "向精神クエリ"
MASTER_CSV_PATH = Path("masters/product_master.csv")
GUIDE_DIR = Path("guide")

OUTPUT_COLUMNS = [
    "商品コード",
    "商品名",
    "包装単位",
    "商品マスタ.記号",
    "数",
]

# 今回の対象
DEFAULT_TARGET_KEYWORDS = [
    "向精神",
    "生活改善薬",
    "毒薬",
]

GUIDE_TEXTS = {
    "step1": [
        "現在在庫検索を開きます。",
        "条件を指定して検索します。",
        "在庫データをダウンロードします。",
    ],
    "step2": [
        "在庫データをアップロードします。",
        "資料を作成します。",
        "完成ファイルをダウンロードします。",
    ],
}


# =========================
# デザイン
# =========================
st.markdown(
    """
<style>
.main .block-container {
    padding-top: 2rem;
    padding-bottom: 3rem;
    max-width: 1180px;
}
.step-card {
    background: #ffffff;
    border: 1px solid #e5e7eb;
    border-radius: 18px;
    padding: 20px 22px;
    margin: 14px 0 10px 0;
    box-shadow: 0 6px 18px rgba(15, 23, 42, 0.06);
}
.step-title {
    font-size: 1.15rem;
    font-weight: 800;
    color: #111827;
    margin-bottom: 8px;
}
.step-note {
    color: #4b5563;
    line-height: 1.7;
}
.small-note {
    color: #6b7280;
    font-size: 0.92rem;
}
.ok-box {
    background: #ecfdf5;
    border: 1px solid #a7f3d0;
    color: #065f46;
    border-radius: 14px;
    padding: 12px 14px;
    margin: 10px 0;
}
.warn-box {
    background: #fffbeb;
    border: 1px solid #fde68a;
    color: #92400e;
    border-radius: 14px;
    padding: 12px 14px;
    margin: 10px 0;
}
.error-box {
    background: #fef2f2;
    border: 1px solid #fecaca;
    color: #991b1b;
    border-radius: 14px;
    padding: 12px 14px;
    margin: 10px 0;
}
.admin-box {
    background: #f8fafc;
    border: 1px dashed #94a3b8;
    color: #334155;
    border-radius: 14px;
    padding: 12px 14px;
    margin: 10px 0;
}
.guide-slider-card {
    max-width: 860px;
    margin: 12px auto 24px auto;
    padding: 14px 18px 16px 18px;
    border-radius: 18px;
    border: 1px solid #e5e7eb;
    background: #ffffff;
    box-shadow: 0 6px 18px rgba(15, 23, 42, 0.08);
}
.guide-nav {
    max-width: 720px;
    margin: 0 auto 8px auto;
}
.guide-page {
    text-align: center;
    font-weight: 700;
    color: #374151;
    line-height: 2rem;
}
.guide-image {
    display: flex;
    justify-content: center;
}
.guide-caption {
    text-align: center;
    color: #6b7280;
    margin-top: 6px;
}
</style>
""",
    unsafe_allow_html=True,
)


# =========================
# 共通関数
# =========================
def read_csv_safely(file_or_path) -> pd.DataFrame:
    """CSVを文字コード違いに強く読み込む。Excel由来のCP932にも対応。"""
    last_error = None
    for enc in ["utf-8-sig", "utf-8", "cp932", "shift_jis"]:
        try:
            if hasattr(file_or_path, "seek"):
                file_or_path.seek(0)
            return pd.read_csv(file_or_path, dtype=str, encoding=enc)
        except Exception as e:
            last_error = e
    raise ValueError(f"CSVを読み込めませんでした。文字コードまたは形式を確認してください。詳細: {last_error}")


def normalize_code(value) -> str:
    """商品コードの比較用。Excelで数値扱いされても文字列として揃える。"""
    if pd.isna(value):
        return ""
    text = str(value).strip()
    if text.endswith(".0"):
        text = text[:-2]
    return text


def find_column(df: pd.DataFrame, candidates: list[str]) -> str | None:
    """候補名から実在する列名を探す。"""
    normalized_map = {str(col).strip(): col for col in df.columns}
    for name in candidates:
        if name in normalized_map:
            return normalized_map[name]
    return None


def read_table(uploaded_file) -> pd.DataFrame:
    """CSV/Excelを読み込む。Excelは先頭シートを読む。"""
    name = uploaded_file.name.lower()
    if name.endswith(".csv"):
        return read_csv_safely(uploaded_file)
    return pd.read_excel(uploaded_file, sheet_name=0, dtype=str)


def normalize_master(master: pd.DataFrame) -> pd.DataFrame:
    """商品マスタを 商品コード・記号 の2列に整える。"""
    master = master.copy()
    master.columns = [str(c).strip() for c in master.columns]

    code_col = find_column(master, ["商品コード", "商品CD", "商品cd", "品目コード"])
    mark_col = find_column(master, ["記号", "商品マスタ.記号", "区分", "分類"])

    if code_col is None or mark_col is None:
        raise ValueError(
            "商品マスタに必要な列がありません。必要な列は「商品コード」と「記号」です。"
        )

    result = master[[code_col, mark_col]].copy()
    result.columns = ["商品コード", "記号"]

    result["商品コード"] = result["商品コード"].map(normalize_code)
    result["記号"] = result["記号"].fillna("").astype(str).str.strip()

    result = result[result["商品コード"] != ""]
    result = result[result["記号"] != ""]
    result = result.drop_duplicates(subset=["商品コード"], keep="first").reset_index(drop=True)

    return result


def load_master(master_upload=None) -> pd.DataFrame:
    """
    商品マスタを読む。
    優先順位：
    1. 管理者用タブで今回反映したマスタ
    2. masters/product_master.csv
    """
    if master_upload is not None:
        master = read_table(master_upload)
        return normalize_master(master)

    if "active_master_df" in st.session_state:
        return st.session_state["active_master_df"].copy()

    if MASTER_CSV_PATH.exists():
        master = read_csv_safely(MASTER_CSV_PATH)
        return normalize_master(master)

    raise FileNotFoundError(
        "商品マスタが見つかりません。管理者用タブで商品マスタCSVを作成するか、masters/product_master.csv を置いてください。"
    )


def filter_master(master: pd.DataFrame, selected_marks: list[str]) -> pd.DataFrame:
    """選択された記号だけ商品マスタを絞る。"""
    if not selected_marks:
        return master.iloc[0:0].copy()
    return master[master["記号"].isin(selected_marks)].reset_index(drop=True)


def df_to_csv_bytes(df: pd.DataFrame) -> bytes:
    """Excelで開きやすいUTF-8 BOM付きCSVにする。"""
    return df.to_csv(index=False, encoding="utf-8-sig").encode("utf-8-sig")


def save_master_csv(df: pd.DataFrame) -> None:
    """ローカルのmasters/product_master.csvへ保存する。"""
    MASTER_CSV_PATH.parent.mkdir(parents=True, exist_ok=True)
    df.to_csv(MASTER_CSV_PATH, index=False, encoding="utf-8-sig")


def build_output(source: pd.DataFrame, master: pd.DataFrame, target_keywords: list[str]) -> pd.DataFrame:
    """
    元データと商品マスタを照合して、出力用5列を作る。
    並び順ロジックは後から追加する前提。
    """
    source = source.copy()
    source.columns = [str(c).strip() for c in source.columns]

    code_col = find_column(source, ["商品コード", "商品CD", "商品cd", "品目コード"])
    name_col = find_column(source, ["商品名", "品名", "商品名称", "正式名称"])
    pack_col = find_column(source, ["包装単位", "包装", "規格", "包装規格"])
    qty_col = find_column(source, ["数", "数量", "在庫数", "現在庫数", "在庫数量"])

    missing = []
    if code_col is None:
        missing.append("商品コード")
    if name_col is None:
        missing.append("商品名")
    if pack_col is None:
        missing.append("包装単位")
    if qty_col is None:
        missing.append("数/数量/在庫数")

    if missing:
        raise ValueError(
            "元データに必要な列がありません："
            + "、".join(missing)
            + "。列名を確認してください。"
        )

    work = source[[code_col, name_col, pack_col, qty_col]].copy()
    work.columns = ["商品コード", "商品名", "包装単位", "数"]
    work["商品コード"] = work["商品コード"].map(normalize_code)

    work["数"] = pd.to_numeric(work["数"], errors="coerce").fillna(0)

    grouped = (
        work.groupby(["商品コード", "商品名", "包装単位"], as_index=False, dropna=False)["数"]
        .sum()
    )

    merged = grouped.merge(master, on="商品コード", how="left")
    merged["記号"] = merged["記号"].fillna("").astype(str).str.strip()

    if target_keywords:
        pattern = "|".join(target_keywords)
        merged = merged[merged["記号"].str.contains(pattern, na=False)]
    else:
        merged = merged[merged["記号"] != ""]

    output = merged[["商品コード", "商品名", "包装単位", "記号", "数"]].copy()
    output.columns = OUTPUT_COLUMNS

    # 並び順は後でロジック追加予定。今は暫定。
    output = output.sort_values(["商品マスタ.記号", "商品コード"], kind="stable").reset_index(drop=True)

    return output


def to_excel_bytes(output_df: pd.DataFrame) -> bytes:
    """出力Excelをメモリ上に作る。"""
    buffer = io.BytesIO()

    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        output_df.to_excel(writer, index=False, sheet_name=OUTPUT_SHEET_NAME)

        ws = writer.book[OUTPUT_SHEET_NAME]

        widths = {
            "A": 16,
            "B": 42,
            "C": 14,
            "D": 20,
            "E": 10,
        }
        for col, width in widths.items():
            ws.column_dimensions[col].width = width

        ws.auto_filter.ref = ws.dimensions
        ws.freeze_panes = "A2"

        for cell in ws["E"][1:]:
            cell.number_format = "0"

    buffer.seek(0)
    return buffer.getvalue()


def show_step(title: str, body: str):
    st.markdown(
        f"""
<div class="step-card">
  <div class="step-title">{title}</div>
  <div class="step-note">{body}</div>
</div>
""",
        unsafe_allow_html=True,
    )


def get_guide_images(prefix: str) -> list[Path]:
    """guide フォルダから指定プレフィックス画像を順序付きで取得する。"""
    if not GUIDE_DIR.exists():
        return []
    return sorted(
        [p for p in GUIDE_DIR.glob(f"{prefix}_*.png") if p.is_file()],
        key=lambda p: p.name,
    )


def render_guide_slider(prefix: str, descriptions: list[str]) -> None:
    """前へ/次へで切り替える簡易スライドUIを表示する。"""
    images = get_guide_images(prefix)
    if not images:
        return

    state_key = f"{prefix}_slide_index"
    if state_key not in st.session_state:
        st.session_state[state_key] = 0

    total = len(images)
    current_index = st.session_state[state_key]
    current_index = max(0, min(current_index, total - 1))
    st.session_state[state_key] = current_index

    st.markdown("<div class='guide-slider-card'>", unsafe_allow_html=True)
    st.markdown("<div class='guide-nav'>", unsafe_allow_html=True)

    nav_left, nav_pos, nav_right = st.columns([1.2, 0.8, 1.2], gap="small")
    with nav_left:
        if st.button(
            "← 前の画像",
            key=f"{prefix}_prev",
            use_container_width=False,
            disabled=(current_index == 0),
        ):
            st.session_state[state_key] = max(0, st.session_state[state_key] - 1)
            st.rerun()

    with nav_pos:
        st.markdown(
            f"<div class='guide-page'>{current_index + 1} / {total}</div>",
            unsafe_allow_html=True,
        )

    with nav_right:
        if st.button(
            "次の画像 →",
            key=f"{prefix}_next",
            use_container_width=False,
            disabled=(current_index == total - 1),
        ):
            st.session_state[state_key] = min(total - 1, st.session_state[state_key] + 1)
            st.rerun()

    st.markdown("</div>", unsafe_allow_html=True)
    st.markdown("<div class='guide-image'>", unsafe_allow_html=True)
    st.image(str(images[current_index]), width=720)
    st.markdown("</div>", unsafe_allow_html=True)

    if current_index < len(descriptions):
        st.markdown(
            f"<div class='guide-caption'>{descriptions[current_index]}</div>",
            unsafe_allow_html=True,
        )

    st.markdown("</div>", unsafe_allow_html=True)


# =========================
# 画面
# =========================
st.title(APP_TITLE)

st.caption(
    "在庫データを取り込み、対象商品のチェック用資料を作成します。"
)

tab_make, tab_admin = st.tabs(["在庫資料作成", "管理者用"])

with tab_make:
    show_step(
        "ステップ1：在庫データをダウンロード",
        "いつもの手順で在庫データをダウンロードしてください。"
        "<br>あとでここに操作画像やGIFを入れます。"
    )
    render_guide_slider("step1", GUIDE_TEXTS.get("step1", []))

    show_step(
        "ステップ2：在庫データをアップロード",
        "ダウンロードした在庫データを下のアップロード欄に入れてください。"
    )
    render_guide_slider("step2", GUIDE_TEXTS.get("step2", []))

    source_file = st.file_uploader(
        "在庫データをアップロード",
        type=["xlsx", "xls", "csv"],
        help="商品コード・商品名・包装単位・数/数量/在庫数 が入っているファイルを選んでください。",
        key="source_file",
    )

    master_file = None
    target_keywords = DEFAULT_TARGET_KEYWORDS.copy()

    st.markdown("---")

    if source_file is None:
        st.info("まず在庫データをアップロードしてください。")
    else:
        try:
            source_df = read_table(source_file)
            master_df = load_master(master_file)

            st.markdown('<div class="ok-box">ファイルを読み込みました。</div>', unsafe_allow_html=True)

            with st.expander("読み込み内容を確認する", expanded=False):
                st.write("在庫データ")
                st.dataframe(source_df.head(20), use_container_width=True)

            if st.button("資料を作成する", type="primary", use_container_width=True):
                output_df = build_output(source_df, master_df, target_keywords)

                if output_df.empty:
                    st.markdown(
                        '<div class="warn-box">対象データが0件でした。商品コードと商品マスタの記号を確認してください。</div>',
                        unsafe_allow_html=True,
                    )
                    st.stop()

                excel_bytes = to_excel_bytes(output_df)

                st.markdown(
                    f'<div class="ok-box">資料を作成しました。対象件数：{len(output_df)}件</div>',
                    unsafe_allow_html=True,
                )

                st.dataframe(output_df, use_container_width=True, height=420)

                show_step(
                    "ステップ3：完成ファイルをダウンロード",
                    "下のボタンから完成したExcelファイルをダウンロードしてください。"
                )

                st.download_button(
                    label="完成ファイルをダウンロード",
                    data=excel_bytes,
                    file_name="向精神薬_生活改善薬_毒薬_在庫資料.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

        except Exception as e:
            st.markdown(
                f'<div class="error-box">エラー：{e}</div>',
                unsafe_allow_html=True,
            )

with tab_admin:
    show_step(
        "管理者用：商品マスタ更新",
        "商品マスタが更新された時だけ使う画面です。"
        "<br>整形前の商品一覧Excelをアップロードすると、対象の記号だけを抽出して、アプリ用の product_master.csv 形式に整えます。"
        "<br>作成したマスタで、アプリ内の masters/product_master.csv を上書きできます。"
    )

    st.markdown(
        '<div class="warn-box">通常作業者はこのタブを触らなくてOKです。商品マスタ更新時だけ使用してください。</div>',
        unsafe_allow_html=True,
    )

    raw_master_file = st.file_uploader(
        "整形前の商品一覧Excel/CSVをアップロード",
        type=["xlsx", "xls", "csv"],
        key="raw_master_file",
    )

    if raw_master_file is None:
        st.info("整形前の商品一覧ファイルをアップロードしてください。")
    else:
        try:
            raw_master_df = read_table(raw_master_file)
            normalized_master = normalize_master(raw_master_df)

            available_marks = sorted([x for x in normalized_master["記号"].dropna().unique().tolist() if str(x).strip() != ""])

            default_selected = [
                mark for mark in available_marks
                if ("向精神" in mark) or (mark == "生活改善薬") or (mark == "毒薬")
            ]

            selected_marks = st.multiselect(
                "抽出する記号を選んでください",
                options=available_marks,
                default=default_selected,
            )

            product_master = filter_master(normalized_master, selected_marks)

            c1, c2, c3 = st.columns(3)
            c1.metric("整形前行数", f"{len(raw_master_df):,}")
            c2.metric("記号あり件数", f"{len(normalized_master):,}")
            c3.metric("抽出後件数", f"{len(product_master):,}")

            with st.expander("記号別件数を見る", expanded=True):
                count_df = (
                    normalized_master["記号"]
                    .value_counts()
                    .rename_axis("記号")
                    .reset_index(name="件数")
                )
                st.dataframe(count_df, use_container_width=True, height=260)

            st.write("抽出後プレビュー")
            st.dataframe(product_master.head(100), use_container_width=True, height=360)

            csv_bytes = df_to_csv_bytes(product_master)

            col_a, col_b = st.columns(2)

            with col_a:
                st.download_button(
                    label="product_master.csv をダウンロード",
                    data=csv_bytes,
                    file_name="product_master.csv",
                    mime="text/csv",
                    use_container_width=True,
                )

            with col_b:
                if st.button("このマスタで更新する", type="primary", use_container_width=True):
                    save_master_csv(product_master)
                    st.session_state["active_master_df"] = product_master.copy()
                    st.success("アプリ内の masters/product_master.csv を上書きしました。在庫資料作成タブでこのマスタを使用できます。")

            st.markdown(
                '<div class="small-note">「このマスタで更新する」を押すと、アプリ内の masters/product_master.csv を上書きします。更新後は在庫資料作成タブでそのまま使えます。</div>',
                unsafe_allow_html=True,
            )

        except Exception as e:
            st.markdown(
                f'<div class="error-box">エラー：{e}</div>',
                unsafe_allow_html=True,
            )
