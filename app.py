import os
import streamlit as st
import pandas as pd
import io
import json

from openai import OpenAI

st.set_page_config(page_title="売上ミニ分析", page_icon="📊")
st.title("📊 売上ミニ分析（身内用）")
tab1, tab2 = st.tabs(["📊 売上分析", "🧾 Excel作成"])
with tab1:
   
# APIキー取得
    api_key = os.getenv("OPENAI_API_KEY")
    if not api_key:
        st.error("OPENAI_API_KEY が未設定です")
        st.stop()


    client = OpenAI(api_key=api_key)

    st.write("売上データを入力するか、Excelをアップロードしてください。")

    # テキスト入力
    text_data = st.text_area("データ入力", height=200)

    # Excelアップロード
    uploaded_file = st.file_uploader("Excelファイルをアップロード", type=["xlsx"])

    data = ""

    if uploaded_file is not None:
        df = pd.read_excel(uploaded_file)
        st.write("📊 アップロードデータ")
        st.dataframe(df)
        data = df.to_csv(index=False)
    else:
        data = text_data

if st.button("分析する"):
    if not data:
        st.warning("データを入力するか、Excelをアップロードしてください。")
    else:
        with st.spinner("分析中..."):
            try:
                response = client.responses.create(
                    model="gpt-4.1-mini",
                    input=f"""
以下の売上データを分析してください。

・全体の傾向（2〜3行）
・重要な気づき3つ
・次にやるべきこと3つ

データ:
{data}
"""
                )

                st.subheader("📊 分析結果")
                st.write(response.output_text)

            except Exception as e:
                st.error(f"エラーが発生しました: {e}")





with tab2:
    st.subheader("📄 Excel作成（複数シート対応）")

    spec = st.text_area(
        "作りたいExcelの内容を文章で書いてください（シート分けもOK）",
        height=240,
        placeholder="例：売上・経費でシート分け..."
    )

    if st.button("Excel作成（複数シート）"):
        if not spec.strip():
            st.warning("内容を入力してね")
        else:
            with st.spinner("複数シートの表を作成中..."):
                try:
                    prompt = f"""
あなたはExcel作成アシスタントです。
ユーザーの指示から、複数シートの表データをJSONのみで返してください。

形式:
{{
  "sheets": [
    {{
      "name": "シート名",
      "columns": ["列1", "列2"],
      "rows": [
        ["値1", "値2"]
      ]
    }}
  ]
}}

ユーザー指示:
{spec}
"""

                    res = client.responses.create(
                        model="gpt-4.1-mini",
                        input=prompt
                    )

                    raw = res.output[0].content[0].text.strip()

                    start = raw.find("{")
                    end = raw.rfind("}")
                    if start == -1 or end == -1:
                        raise ValueError("JSONが見つかりませんでした")

                    obj = json.loads(raw[start:end+1])
                    sheets = obj.get("sheets", [])

                    out = io.BytesIO()

                    with pd.ExcelWriter(out, engine="openpyxl") as writer:
                        for s in sheets:
                            name = str(s.get("name", "Sheet"))[:31]
                            cols = s.get("columns", [])
                            rows = s.get("rows", [])

                            df2 = pd.DataFrame(rows, columns=cols)
                            df2.to_excel(writer, index=False, sheet_name=name)

                    out.seek(0)

                    st.download_button(
                        label="📥 Excel（複数シート .xlsx）をダウンロード",
                        data=out,
                        file_name="generated_multi_sheet.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                    )

                except Exception as e:
                    st.error(f"作成失敗: {e}")
