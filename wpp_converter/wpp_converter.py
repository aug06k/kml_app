import xml.etree.ElementTree as ET
import pandas as pd
import re
import streamlit as st
import os

# アップロードファイルの保存先
UPLOAD_DIR = "wpp_converter/uploads"
os.makedirs(UPLOAD_DIR, exist_ok=True)

# Streamlit UI
st.title("WPP Height Converter")
st.write("Excel (`.xlsm`) の `height` 値を `.wpp` ファイルに適用します。")

# ファイルアップロード
xlsm_file = st.file_uploader("Excelファイル（.xlsm）を選択", type=["xlsm"])
wpp_file = st.file_uploader("WPPファイルを選択", type=["wpp"])

if xlsm_file and wpp_file:
    # 一時保存
    xlsm_path = os.path.join(UPLOAD_DIR, "input.xlsm")
    wpp_path = os.path.join(UPLOAD_DIR, "input.wpp")
    
    with open(xlsm_path, "wb") as f:
        f.write(xlsm_file.getbuffer())
    with open(wpp_path, "wb") as f:
        f.write(wpp_file.getbuffer())

    # Excelデータの読み込み
    def load_height_data(file_path):
        df = pd.read_excel(file_path, sheet_name=None)
        first_sheet = list(df.keys())[0]
        df = df[first_sheet]
        df = df.iloc[5:, [0, 5]]
        df.columns = ["WP", "height"]
        df["WP"] = df["WP"].apply(lambda x: int(re.sub(r"\D", "", str(x))) if pd.notna(x) else None)
        return dict(zip(df["WP"], df["height"]))

    # WPPファイルの更新
    def update_wpp_heights(wpp_file, height_dict, output_file):
        tree = ET.parse(wpp_file)
        root = tree.getroot()
        for waypoint in root.findall(".//waypoint"):
            id_tag = waypoint.find("ID")
            height_tag = waypoint.find("height")
            if id_tag is not None and height_tag is not None:
                waypoint_id = int(id_tag.text)
                if waypoint_id in height_dict and not pd.isna(height_dict[waypoint_id]):
                    height_tag.text = str(height_dict[waypoint_id])
        tree.write(output_file, encoding="utf-8", xml_declaration=True)
        return output_file

    height_data = load_height_data(xlsm_path)
    updated_wpp_path = os.path.join(UPLOAD_DIR, "updated.wpp")
    update_wpp_heights(wpp_path, height_data, updated_wpp_path)

    with open(updated_wpp_path, "rb") as f:
        st.download_button("更新されたWPPをダウンロード", f, "updated.wpp", mime="application/xml")
