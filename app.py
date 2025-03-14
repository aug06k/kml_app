import streamlit as st

st.sidebar.title("アプリ選択")
app_selection = st.sidebar.radio("アプリを選択", ["メインアプリ", "WPP Converter"])

if app_selection == "メインアプリ":
    st.title("メインアプリ")
    st.write("ここに既存のアプリのコードを記述")

elif app_selection == "WPP Converter":
    from wpp_converter import wpp_converter
    wpp_converter  # Streamlitが `wpp_converter.py` を実行

import xml.etree.ElementTree as ET
import tempfile

# タイトル
st.title("KML変換アプリ")
st.write("KMLファイルをアップロードすると、KMLにWP番号を付けたものを出力します。")

# KMLファイルのアップロード
uploaded_file = st.file_uploader("KMLファイルを選択してください", type=["kml"])

def process_kml(kml_content):
    tree = ET.ElementTree(ET.fromstring(kml_content))
    root = tree.getroot()
    namespace = {"kml": "http://www.opengis.net/kml/2.2"}

    # Documentノードを取得
    document = root.find("kml:Document", namespace)

    # 最初の LineString の座標を取得
    coordinates_elem = root.find(".//kml:Placemark/kml:LineString/kml:coordinates", namespace)

    if coordinates_elem is None:
        st.error("Error: No LineString coordinates found in the first section.")
        return None

    # 座標を取得
    coordinates_text = coordinates_elem.text.strip()
    coord_lines = coordinates_text.split("\n")

    # 各座標の Placemark を作成
    placemark_index = 1
    for line in coord_lines:
        parts = line.strip().split(",")
        if len(parts) == 3:
            lon, lat, alt = map(float, parts)

            # Placemarkを作成
            placemark = ET.Element("Placemark")

            name = ET.SubElement(placemark, "name")
            name.text = str(placemark_index)  # 連番を振る

            visibility = ET.SubElement(placemark, "visibility")
            visibility.text = "1"

            point = ET.SubElement(placemark, "Point")

            altitude_mode = ET.SubElement(point, "altitudeMode")
            altitude_mode.text = "absolute"

            coordinates = ET.SubElement(point, "coordinates")
            coordinates.text = f"{lon},{lat},{alt}"  # 10m上に配置

            # Document ノードに追加
            document.append(placemark)

            placemark_index += 1  # 番号を増やす

    # 修正後のKMLを一時ファイルとして保存
    temp_file = tempfile.NamedTemporaryFile(delete=False, suffix=".kml")
    tree.write(temp_file.name, encoding="utf-8", xml_declaration=True)

    return temp_file.name

# ファイルがアップロードされたら処理
if uploaded_file is not None:
    st.write("ファイルを処理中...")

    # ファイルを読み込み
    kml_content = uploaded_file.getvalue().decode("utf-8")

    # KMLを処理
    output_kml_path = process_kml(kml_content)

    if output_kml_path:
        # ダウンロードボタンを表示
        with open(output_kml_path, "rb") as f:
            st.download_button(
                label="修正済みKMLをダウンロード",
                data=f,
                file_name="updated_kml.kml",
                mime="application/vnd.google-earth.kml+xml"
            )

        st.success("処理が完了しました！")
