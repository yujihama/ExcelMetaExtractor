import streamlit as st
import json
from excel_metadata_extractor import ExcelMetadataExtractor
import pandas as pd

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
    st.write(f"📍 Region Type: {region['regionType']}")
    st.write(f"📏 Range: {region['range']}")

    if region['regionType'] == 'table':
        with st.expander("📊 Table Structure"):
            st.json(region['headerStructure'])
            if 'notes' in region and region['notes']:
                st.write("📝 Notes:", region['notes'])

    elif region['regionType'] == 'text':
        with st.expander("📝 Text Content"):
            st.write("Content:", region['content'])
            st.write("Classification:", region['classification'])
            st.write("Importance:", region['importance'])
            if 'summary' in region:
                st.write("Summary:", region['summary'])
            if 'keyPoints' in region and region['keyPoints']:
                st.write("Key Points:")
                for point in region['keyPoints']:
                    st.write(f"• {point}")

    elif region['regionType'] == 'chart':
        with st.expander("📈 Chart Information"):
            st.write("Chart Type:", region['chartType'])
            st.write("Purpose:", region['purpose'])
            if 'dataRelations' in region and region['dataRelations']:
                st.write("Data Relations:")
                for relation in region['dataRelations']:
                    st.write(f"• {relation}")
            if 'suggestedUsage' in region:
                st.write("Suggested Usage:", region['suggestedUsage'])

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
            for sheet in metadata["worksheets"]:
                with st.expander(f"📚 Sheet: {sheet['sheetName']}", expanded=True):
                    cols = st.columns(3)
                    with cols[0]:
                        st.metric("Rows", sheet["rowCount"])
                    with cols[1]:
                        st.metric("Columns", sheet["columnCount"])
                    with cols[2]:
                        st.metric("Merged Cells", len(sheet.get("mergedCells", [])))

                    # Display regions
                    if "regions" in sheet and sheet["regions"]:
                        st.subheader("📍 Detected Regions")
                        for region in sheet["regions"]:
                            with st.expander(f"{region['regionType'].title()} Region - {region['range']}"):
                                display_region_info(region)

                    if sheet.get("mergedCells"):
                        with st.expander("🔀 Merged Cells"):
                            st.code("\n".join(sheet["mergedCells"]))

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
            st.error("Please make sure you've uploaded a valid Excel file.")

if __name__ == "__main__":
    main()