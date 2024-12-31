import streamlit as st
import json
from excel_metadata_extractor import ExcelMetadataExtractor
import pandas as pd
import traceback

def display_json_tree(data, key_prefix=""):
    """Display JSON data in an expandable tree format"""
    if isinstance(data, dict):
        for key, value in data.items():
            new_key = f"{key_prefix}/{key}" if key_prefix else key
            if isinstance(value, (dict, list)):
                with st.expander(f"🔍 {key}"):
                    display_json_tree(value, new_key)
            else:
                st.text(f"{key}: {value}")
    elif isinstance(data, list):
        for i, item in enumerate(data):
            new_key = f"{key_prefix}[{i}]"
            if isinstance(item, (dict, list)):
                with st.expander(f"📑 Item {i+1}"):
                    display_json_tree(item, new_key)
            else:
                st.text(f"- {item}")

def display_region_info(region):
    """Display region information in a structured format"""
    try:
        st.write(f"📍 Region Type: {region['regionType']}")
        st.write(f"📏 Range: {region['range']}")

        # Shape specific information
        if region['regionType'] == 'shape':
            st.markdown("##### 🔷 Shape Information")
            cols = st.columns(2)
            with cols[0]:
                if 'shape_type' in region and region['shape_type']: #Added condition to check for shape_type
                    st.metric("Shape Type", region.get('shape_type', 'Unknown').title())
                if region.get('name'):
                    st.text(f"Name: {region['name']}")
            with cols[1]:
                if region.get('description'):
                    st.text(f"Description: {region['description']}")

            # Display shape text content if available
            if 'text_content' in region:
                st.markdown("##### 📝 Text Content")
                st.text(region['text_content'])

        elif region['regionType'] in ['image', 'smartart', 'chart']:
            st.markdown("##### 🖼️ Drawing Information")

            # 基本情報の表示
            cols = st.columns(2)
            with cols[0]:
                st.metric("Type", region['type'].title())
                if region.get('name'):
                    st.text(f"Name: {region['name']}")
            with cols[1]:
                if region.get('description'):
                    st.text(f"Description: {region['description']}")

            # 座標情報の表示
            if 'coordinates' in region:
                st.markdown("**Position:**")
                coords = region['coordinates']
                st.text(f"From: Column {coords['from']['col']}, Row {coords['from']['row']}")
                st.text(f"To: Column {coords['to']['col']}, Row {coords['to']['row']}")

            # 図形タイプ別の特殊情報
            if region['type'] == 'image':
                if 'image_ref' in region:
                    st.markdown("**Image Details:**")
                    st.text(f"Reference: {region['image_ref']}")
            elif region['type'] == 'smartart':
                if 'diagram_type' in region:
                    st.markdown("**SmartArt Details:**")
                    st.text(f"Diagram Type: {region['diagram_type']}")
            elif region['type'] == 'chart':
                if 'chart_ref' in region:
                    st.markdown("**Chart Details:**")
                    st.text(f"Chart Reference: {region['chart_ref']}")

        elif region['regionType'] == 'table':
            st.markdown("##### 📊 Table Information")

            # Display header structure information
            if 'headerStructure' in region:
                st.markdown("**Header Structure:**")
                cols = st.columns(3)
                with cols[0]:
                    header_type = region['headerStructure'].get('headerType', 'Unknown')
                    st.metric("Header Type", header_type.title())
                with cols[1]:
                    header_range = region['headerStructure'].get('headerRange', 'N/A')
                    st.metric("Header Range", header_range)
                with cols[2]:
                    has_merged = region['headerStructure'].get('mergedCells', False)
                    st.metric("Has Merged Cells", "Yes" if has_merged else "No")

            # Display sample cells if available
            if 'sampleCells' in region and region['sampleCells']:
                st.markdown("**Sample Data:**")
                sample_data = []
                for row in region['sampleCells']:
                    row_data = {f"Column {cell['col']}": cell['value'] for cell in row}
                    sample_data.append(row_data)
                if sample_data:
                    st.dataframe(pd.DataFrame(sample_data))

        elif region['regionType'] == 'text':
            st.markdown("##### 📝 Text Information")
            if 'content' in region:
                st.text_area("Content", region['content'], height=100)
            if 'classification' in region:
                st.write("Classification:", region['classification'])
            if 'importance' in region:
                st.write("Importance:", region['importance'])

        # Display merged cells if available
        if 'mergedCells' in region and region['mergedCells']:
            st.markdown("##### 🔀 Merged Cells")
            for merged in region['mergedCells']:
                st.text(f"Range: {merged['range']} - Value: {merged.get('value', 'N/A')}")

    except Exception as e:
        st.error(f"Error displaying region info: {str(e)}")
        st.error(f"Region data: {json.dumps(region, indent=2)}")
        st.error(f"Stack trace:\n{traceback.format_exc()}")

def main():
    st.set_page_config(
        page_title="Excel Metadata Extractor",
        page_icon="📊",
        layout="wide"
    )

    st.title("📊 Excel Metadata Extractor")
    st.markdown("""
    Upload an Excel file to extract and view its metadata including:
    - File properties (name, size, creation date, etc.)
    - Sheet information (dimensions, protection status, etc.)
    - Detected regions (tables, text blocks, charts)
    - AI-powered analysis of content and structure
    """)

    uploaded_file = st.file_uploader("Choose an Excel file", type=['xlsx', 'xlsm'])

    if uploaded_file is not None:
        try:
            # Extract metadata
            extractor = ExcelMetadataExtractor(uploaded_file)
            metadata = extractor.extract_all_metadata()

            # Display sections
            st.header("📑 Extracted Metadata")

            # File Properties Section
            with st.expander("📌 File Properties", expanded=True):
                st.json(metadata["fileProperties"])

            # Worksheets Section
            for sheet_idx, sheet in enumerate(metadata["worksheets"]):
                st.subheader(f"📚 Sheet: {sheet['sheetName']}")

                # Sheet metrics
                cols = st.columns(3)
                with cols[0]:
                    st.metric("Rows", sheet["rowCount"])
                with cols[1]:
                    st.metric("Columns", sheet["columnCount"])
                with cols[2]:
                    st.metric("Merged Cells", len(sheet.get("mergedCells", [])))

                # Merged Cells (at same level as other sections)
                if sheet.get("mergedCells"):
                    st.markdown("##### 🔀 Merged Cells")
                    st.code("\n".join(sheet["mergedCells"]))

                # Regions (at same level as other sections)
                if "regions" in sheet and sheet["regions"]:
                    st.markdown("##### 📍 Detected Regions")
                    for region in sheet["regions"]:
                        try:
                            with st.expander(f"{region['regionType'].title()} Region - {region['range']}"):
                                display_region_info(region)
                        except Exception as e:
                            st.error(f"Error processing region: {str(e)}\nRegion data: {json.dumps(region, indent=2)}")
                            st.error(f"Stack trace:\n{traceback.format_exc()}")

                st.markdown("---")  # Add separator between sheets

            # Raw JSON View
            with st.expander("🔍 Raw JSON Data"):
                st.json(metadata)

            # Download button for JSON
            json_str = json.dumps(metadata, indent=2)
            st.download_button(
                label="Download JSON",
                data=json_str,
                file_name=f"{uploaded_file.name}_metadata.json",
                mime="application/json"
            )

        except Exception as e:
            st.error(f"Error processing file: {str(e)}")
            st.error(f"Detailed error:\n{traceback.format_exc()}")
            st.error("Please make sure you've uploaded a valid Excel file.")

if __name__ == "__main__":
    main()