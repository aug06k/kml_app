import streamlit as st
import xml.etree.ElementTree as ET
import pandas as pd
import re
import os
import tempfile

def run():
    st.title("WPP Height Converter")
    st.write("Excel (`.xlsm`) の `height` 値を `.wpp` ファイルに適用します。")

    # ファイルアップロード
    xlsm_file = st.file_uploader("Excelファイル（.xlsm）を選択", type=["xlsm"])
    wpp_file = st.file_uploader("WPPファイルを選択", type=["wpp"])

    if xlsm_file and wpp_file:
        # ファイルの保存
        with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsm") as tmp_xlsm:
            tmp_xlsm.write(xlsm_file.getbuffer())
            xlsm_path = tmp_xlsm.name

        with tempfile.NamedTemporaryFile(delete=False, suffix=".wpp") as tmp_wpp:
            tmp_wpp.write(wpp_file.getbuffer())
            wpp_path = tmp_wpp.name

        st.success("ファイルを処理しました！")

        # Excelデータを読み込む
        def load_height_data(file_path):
            df = pd.read_excel(file_path, sheet_name=None)
            first_sheet = list(df.keys())[0]
            df = df[first_sheet]
            df = df.iloc[5:, [0, 5]]  # A6以降とF6以降を取得
            df.columns = ["WP", "height"]
            df["WP"] = df["WP"].apply(lambda x: int(re.sub(r"\D", "", str(x))) if pd.notna(x) else None)
            return dict(zip(df["WP"], df["height"]))

        # WPPファイルの更新
        def update_wpp_heights(wpp_file, height_dict):
            tree = ET.parse(wpp_file)
            root = tree.getroot()

            for waypoint in root.findall(".//waypoint"):
                id_tag = waypoint.find("ID")
                height_tag = waypoint.find("height")

                if id_tag is not None and height_tag is not None:
                    waypoint_id = int(id_tag.text)
                    if waypoint_id in height_dict and not pd.isna(height_dict[waypoint_id]):
                        height_tag.text = str(height_dict[waypoint_id])

            # 修正後のWPPを一時ファイルとして保存
            with tempfile.NamedTemporaryFile(delete=False, suffix=".wpp") as tmp_output_wpp:
                tree.write(tmp_output_wpp.name, encoding="utf-8", xml_declaration=True)
                return tmp_output_wpp.name

        # Excelからheightデータ取得
        height_data = load_height_data(xlsm_path)
        
        # WPPを更新
        updated_wpp_path = update_wpp_heights(wpp_path, height_data)

        # ダウンロードボタンの表示
        with open(updated_wpp_path, "rb") as f:
            st.download_button(
                label="更新されたWPPをダウンロード",
                data=f,
                file_name="updated.wpp",
                mime="application/xml"
            )

        st.success("処理が完了しました！")
