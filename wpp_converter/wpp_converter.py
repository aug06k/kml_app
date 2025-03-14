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

        # Excelデータを読み込む関数
        def load_height_data(file_path):
            df = pd.read_excel(file_path, sheet_name=None)
            first_sheet = list(df.keys())[0]
            df = df[first_sheet]
            df = df.iloc[5:, [0, 5]]  # A6以降とF6以降を取得
            df.columns = ["WP", "height"]

            # WP列を数値に変換し、変換できないものはスキップ
            def extract_wp_number(value):
                match = re.search(r"\d+", str(value))  # 数字を抽出
                return int(match.group()) if match else None  # 数値が見つからなければNone

            df["WP"] = df["WP"].apply(extract_wp_number)
            df = df.dropna(subset=["WP"])  # WPがNone（変換不可）の行は削除

            height_dict = dict(zip(df["WP"].astype(int), df["height"]))

            # **デバッグ用: 取得したデータを表示**
            st.write("読み込んだ height データ:", height_dict)

            return height_dict

        # **height_dict を先に定義**
        height_dict = load_height_data(xlsm_path)  # Excelから height データ取得

        # WPPファイルの更新関数
        def update_wpp_heights(wpp_file, height_dict):
            tree = ET.parse(wpp_file)
            root = tree.getroot()

            all_ids = []  # WPP内のIDリストを確認するためのリスト
            
            for waypoint in root.findall(".//waypoint"):
                id_tag = waypoint.find("ID")
                height_tag = waypoint.find("height")

                if id_tag is not None and height_tag is not None:
                    try:
                        waypoint_id = int(id_tag.text)
                        all_ids.append(waypoint_id)  # IDリストに追加
                        if waypoint_id in height_dict and not pd.isna(height_dict[waypoint_id]):
                            height_tag.text = str(height_dict[waypoint_id])
                    except ValueError:
                        continue  # IDが数値でない場合はスキップ
            
            st.write("WPP 内の全 ID:", all_ids)  # WPP 内の ID リストをデバッグ表示

            # 修正後のWPPを一時ファイルとして保存
            with tempfile.NamedTemporaryFile(delete=False, suffix=".wpp") as tmp_output_wpp:
                tree.write(tmp_output_wpp.name, encoding="utf-8", xml_declaration=True)
                return tmp_output_wpp.name

        # WPPを更新
        updated_wpp_path = update_wpp_heights(wpp_path, height_dict)

        # ダウンロードボタンの表示
        with open(updated_wpp_path, "rb") as f:
            st.download_button(
                label="更新されたWPPをダウンロード",
                data=f,
                file_name="updated.wpp",
                mime="application/xml"
            )

        st.success("処理が完了しました！")
