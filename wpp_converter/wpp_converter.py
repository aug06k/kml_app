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

            # WP列を処理（数字のみ抽出、数値以外は無視）
            def extract_wp_number(value):
                match = re.fullmatch(r"\D*(\d+)\D*", str(value))  # 完全に数字のみの部分を取得
                return int(match.group(1)) if match else None  # 数値が見つからなければNone

            df["WP"] = df["WP"].apply(extract_wp_number)
            df = df.dropna(subset=["WP"])  # WPがNoneの行を削除

            return dict(zip(df["WP"].astype(int), df["height"]))
            
            st.write("読み込んだ height データ:", height_dict)  # 取得データを表示

        return height_dict


        # WPPファイルの更新
        def update_wpp_heights(wpp_file, height_dict):
            tree = ET.parse(wpp_file)
            root = tree.getroot()

            for waypoint in root.findall(".//waypoint"):
                id_tag = waypoint.find("ID")
                height_tag = waypoint.find("height")

                if id_tag is not None and height_tag is not None:
                    try:
                        waypoint_id = int(id_tag.text)  # 数字のみのID以外はエラー
                        if waypoint_id in height_dict and not pd.isna(height_dict[waypoint_id]):
                            height_tag.text = str(height_dict[waypoint_id])
                    except ValueError:
                        continue  # IDが数字以外を含む場合はスキップ

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
