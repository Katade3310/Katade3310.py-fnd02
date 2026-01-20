import pandas as pd
import datetime as dt
from dateutil.relativedelta import relativedelta
import msoffcrypto
import io
import re
import plotly.express as px
import streamlit as st

#ファイルを開いて情報をゲット-----------------------------
path = r"C:\Users\1634106\OneDrive - トヨタ自動車株式会社\デスクトップ\ショートカット\シャシーDX_Portal - ドキュメント\01_DB\管理用\ログ管理.xlsm"
pw = "#4gc4gc"

with open(path,"rb") as file: #ファイルを開く

    file_kari = msoffcrypto.OfficeFile(file) #復号化
    file_kari.load_key(password=pw) #鍵で開ける

    with io.BytesIO() as open_file:  # メモリ上に仮想ファイルを作る
        file_kari.decrypt(open_file) #復号化したデータを仮想ファイルに保存
        open_file.seek(0) #復号化したデータの上?最初?に移動
        df_query = pd.read_excel(open_file,sheet_name="query",index_col="ID",engine="openpyxl") #復号化したデータの読取
        df_app = pd.read_excel(open_file,sheet_name="アプリ毎",engine="openpyxl") #復号化したデータの読取

#App設定---------------------------------------------
tools_all = []
for v in df_app.iloc[4:, 0]:
    if v == "":
        break
    tools_all.append(v)

#部署--------------------------------------------------
#部署コード列の付与（KC/GC/TC/LC/MVC を含んでいればそのコード、無ければ OTHER）
codes = ["KC", "GC", "TC", "LC", "MVC"]
pattern = "(" + "|".join(codes) + ")"  # "(KC|GC|TC|LC|MVC)"
df_query1 = df_query.copy()
df_query1["部署コード"] = (
    df_query1["部署"].astype(str)
    .str.extract(pattern, flags=re.I, expand=False)  # 大文字小文字無視で抽出
    .str.upper()
    .fillna("OTHER")
)

#表示------------------------------------------------------
st.header("DXチームApp 使用状況ダッシュボード")
with st.sidebar:
    start_user = st.date_input("検索開始日を指定してください",value=dt.date.today() - relativedelta(years=1) + relativedelta(months=1))
    end_user = st.date_input("検索終了日を指定してください",value=dt.date.today())
    
    # 部署選択
    busyo_all = ["KC", "GC", "TC", "LC", "MVC"]  # 5本固定
    option1 = st.multiselect("部署を選択してください", options=busyo_all, default=busyo_all)
    if len(option1) == 0:
        st.warning("少なくとも1つの部署を選択してください。")
        st.stop()

    # App 選択
    option2 = st.multiselect(
        "Appを選択してください",
        options=tools_all,
        default=["SUツール", "オプション表_新旧比較結果", "車両仕様反映", "車両仕様紐付", "部品表登録"],
    )
    if len(option2) == 0:
        st.warning("少なくとも1つのAppを選択してください。")
        st.stop()

#日付設定-----------------------------------------------------
start = pd.Timestamp(start_user)
end = pd.Timestamp(end_user) + pd.Timedelta(days=1) - pd.Timedelta(nanoseconds=1)

#情報を加工-------------------------------------------------
#グラフ1
# 期間×App×部署でフィルタ（部署は5本固定のいずれか）
mask_date = df_query1["日時"].between(start, end)
mask_app = df_query1["App"].isin(option2)
mask_busyo = df_query1["部署コード"].isin(option1)  # 5本の中から選ばれたもの
df1 = df_query1.loc[mask_date & mask_app & mask_busyo].copy()

# 年月
df1["年月"] = df1["日時"].dt.to_period("M")
# 月レンジ
start_pm = start.to_period("M")
end_pm = end.to_period("M")
month_list = pd.period_range(start_pm, end_pm, freq="M", name="年月")

# 「月×部署×App」の完全格子を作って0埋め（各月に必ず5本）
apps = option2  # 選択App
busyo = option1 # 選択部署
#↓月×部署×Appの組み合わせを全部作っている
full_idx = pd.MultiIndex.from_product([month_list, busyo, apps], names=["年月", "部署コード", "App"])
# 実データ集計 → 0埋め
counts = (
    df1.groupby(["年月", "部署コード", "App"])
       .size()
       .reindex(full_idx, fill_value=0)
       .reset_index(name="回数")
)

# x軸：複合カテゴリ（「月×部署」を横並び）
counts["x_key"] = counts["年月"].astype(str) + "\n" + counts["部署コード"]
x_order = [f"{m}\n{d}" for m in month_list.astype(str).tolist() for d in busyo]#順番通りの組み合わせ（日付→部署）
# countsのx_keyをカテゴリ型に変換し、全カテゴリをx_order順に並び替え
counts["x_key"] = pd.Categorical(counts["x_key"], categories=x_order,ordered=True)

# 欠損しているx_keyの組み合わせを0で補完
# 各Appごとに全x_keyが存在するよう完全格子を再構築
full_x_app_idx = pd.MultiIndex.from_product([x_order, apps],names=["x_key", "App"])
counts_complete = (
    counts.set_index(["x_key", "App"])["回数"]
    .reindex(full_x_app_idx, fill_value=0)
    .reset_index()
)

# counts_complete にx_keyを分割して、年月と部署コード列を追加
counts_complete["年月"] = counts_complete["x_key"].str.split("\n").str[0]
counts_complete["部署コード_表示"] = counts_complete["x_key"].str.split("\n").str[1]

# Appの色固定　下記は現時点のTop5
color_map = {
    "車両仕様反映": "#1f77b4",
    "車両仕様紐付": "#ff7f0e",
    "オプション表_新旧比較結果": "#2ca02c",
    "SUツール": "#d62728",
    "部品表登録": "#9467bd",
}

# サブタイトル
st.markdown("---")
st.subheader("📊全期間の部署別App使用状況")

# 全期間の部署×Appの合計を集計
total_by_dept_app = (
    counts_complete.groupby(["部署コード_表示", "App"])["回数"]
    .sum()
    .reset_index()
)

# 部署ごとの合計（棒の上に表示する用）
total_by_dept = total_by_dept_app.groupby("部署コード_表示")["回数"].sum()

# グラフ作成
fig_total = px.bar(
    total_by_dept_app,
    x="部署コード_表示",
    y="回数",
    color="App",
    height=400,
    color_discrete_map=color_map,
    labels={
        "部署コード_表示": "部署",
        "回数": "回数 (回)",
        "App": "App"
    },
    category_orders={
        "部署コード_表示": busyo,
        "App": apps
    }
)

fig_total.update_layout(
    barmode="stack",
    showlegend=True,
    margin=dict(t=40, b=60),
    xaxis_title="",
    yaxis_title="回数 (回)"
)

# 各部署の棒の上に合計値を表示
for dept in busyo:
    if dept in total_by_dept.index and total_by_dept[dept] > 0:
        fig_total.add_annotation(
            text=f"<b>{int(total_by_dept[dept])}</b>",
            x=dept,
            y=total_by_dept[dept],
            showarrow=False,
            yshift=10
        )

# グラフ表示
st.plotly_chart(fig_total, use_container_width=True)

#---------------------------------------------------------
# グラフ1
st.markdown("### 📊 月別詳細")

# 表示可能な最大月数を設定（1〜3ヶ月の範囲）
num_months = len(month_list)
max_display_months = min(num_months, 3)  # 期間が短ければそれに合わせる

if "start_month_idx" not in st.session_state:
    # 初期値: 最新3ヶ月を表示(最後から3ヶ月前)
    st.session_state.start_month_idx = max(0, num_months - max_display_months)

# 表示する3ヶ月分のデータ
start_idx = st.session_state.start_month_idx
end_idx = min(start_idx + max_display_months, num_months)

# start_idxが大きすぎる場合の調整
if end_idx - start_idx < max_display_months and num_months >= max_display_months:
    start_idx = num_months - max_display_months
    end_idx = num_months

display_months = month_list[start_idx:end_idx]

# 表示対象の月でフィルタリング
counts_display = counts_complete[counts_complete["年月"].isin(display_months.astype(str))]

# 表示月数がない場合
num_display = len(display_months)
if num_display == 0:
    st.warning("表示可能な月がありません。")
    st.stop()

# 横最大3列
wrap = num_display
height = 350

# グラフ1：月ごとにファセット分割
fig = px.bar(
    counts_display,
    x="部署コード_表示",
    y="回数",
    color="App",
    facet_col="年月",
    facet_col_wrap=wrap,
    height=height,
    color_discrete_map=color_map,
    labels={
        "部署コード_表示": "",  # 「部署」ラベルを削除
        "回数": "回数 (回)", 
        "App": "App"
    },
    category_orders={
        "部署コード_表示": busyo,
        "年月": [str(m) for m in display_months]
    }
)

fig.update_layout(
    barmode="stack",
    showlegend=True,
    margin=dict(t=60, b=60),  # 上下の余白を調整
)

# 月ごとグラフ（アノテーション）更新
annotations = list(fig.layout.annotations)
    # テキストを整形
for annotation in annotations:
    annotation.update(
        text=annotation.text.split("=")[-1],
        font=dict(size=11, color="#333"),
        y=-0.15,
        yanchor="top"
    )

fig.layout.annotations = annotations

# 各棒の上に合計値を表示
# 全グラフでラベル表示
fig.update_xaxes(showticklabels=True)

# グラブ1表示
st.plotly_chart(fig, use_container_width=True)

# ボタン
col_left, col_center, col_right = st.columns([1, 2, 2])

with col_left:
    # 前の月へ移動
    if st.button("◀ 前の月", disabled=(st.session_state.start_month_idx <= 0)):
        st.session_state.start_month_idx -= 1
        st.rerun()

with col_right:
    # 次の月へ移動
    max_start_idx = max(0, num_months - max_display_months)
    if st.button("次の月 ▶", disabled=(st.session_state.start_month_idx >= max_start_idx)):
        st.session_state.start_month_idx += 1
        st.rerun()

#------------------------------------------------------------
# グラフ2：月×App の折れ線(全期間表示を維持)
st.markdown("---")
st.subheader("📈 全期間のApp別使用推移")

df_query_tool = df_query1.loc[mask_date & mask_app & mask_busyo].copy()
df_query_counts = (
    df_query_tool.assign(年月=df_query_tool["日時"].dt.to_period("M"))
                .groupby(["年月", "App"])
                .size()
                .unstack(fill_value=0)
                .reindex(month_list)
                .reset_index()
)

#グラフ2を表示-----------------------------------------------
df_query_counts["年月"] = df_query_counts["年月"].astype(str)
y_cols = [c for c in df_query_counts.columns if c != "年月"]  # ← tools を上書きしない
fig2 = px.line(df_query_counts, x="年月", y=y_cols, markers=True, labels={"年月": "年月", "value": "回数（回）"}, color_discrete_map=color_map)
st.plotly_chart(fig2, use_container_width=True)

