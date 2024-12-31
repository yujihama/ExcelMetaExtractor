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

        if region['regionType'] == 'table':
            st.markdown("##### 📊 Table Structure")
            st.json(region['headerStructure'])
            if 'notes' in region and region['notes']:
                st.write("📝 Notes:", region['notes'])

        elif region['regionType'] == 'text':
            st.markdown("##### 📝 Text Content")
            st.write("Content:", region['content'])
            st.write("Classification:", region.get('classification', 'Unknown'))
            st.write("Importance:", region.get('importance', 'Unknown'))
            if 'summary' in region:
                st.write("Summary:", region['summary'])
            if 'keyPoints' in region and region['keyPoints']:
                st.write("Key Points:")
                for point in region['keyPoints']:
                    st.write(f"• {point}")

        elif region['regionType'] == 'chart':
            st.markdown("##### 📈 Chart Information")
            st.write("Chart Type:", region.get('chartType', 'Unknown'))
            st.write("Purpose:", region.get('purpose', 'Unknown'))
            if 'dataRelations' in region and region['dataRelations']:
                st.write("Data Relations:")
                for relation in region['dataRelations']:
                    st.write(f"• {relation}")
            if 'suggestedUsage' in region:
                st.write("Suggested Usage:", region['suggestedUsage'])

    except Exception as e:
        st.error(f"Error displaying region info: {str(e)}\nRegion data: {json.dumps(region, indent=2)}")
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