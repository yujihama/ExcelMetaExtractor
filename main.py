
import os
import json
import math
import streamlit as st
import pandas as pd
import traceback
from datetime import datetime
import zipfile
import xml.etree.ElementTree as ET
from openpyxl import load_workbook
import openpyxl.cell.cell
from openpyxl.utils import get_column_letter, column_index_from_string
from typing import Dict, Any, List, Optional, Tuple
from pathlib import Path
import tempfile
from openai import OpenAI
import re

class OpenAIHelper:
    def __init__(self):
        self.client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))
        self.model = "gpt-4o"

    def analyze_region_type(self, region_data: str) -> Dict[str, Any]:
        data = json.loads(region_data)
        sample_data = {
            "cells": data["cells"][:5],
            "mergedCells": data.get("mergedCells", [])[:3]
        }

        prompt = f"""Analyze the following Excel region sample data and determine:
1. The type of region (table, text, chart, image)
2. Any special characteristics or patterns
3. The purpose or meaning of the content

Region sample data (first few rows/cells):
{json.dumps(sample_data, indent=2)}

Respond in JSON format with the following structure:
{{
    "regionType": "table" or "text" or "chart" or "image",
    "characteristics": [string],
    "purpose": string,
    "confidence": number
}}
"""
        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[{"role": "user", "content": prompt}],
                response_format={"type": "json_object"},
                max_tokens=1000
            )
            return response.choices[0].message.content
        except Exception as e:
            print(f"Error in analyze_region_type: {str(e)}")
            return json.dumps({
                "regionType": "unknown",
                "characteristics": [],
                "purpose": "Error in analysis",
                "confidence": 0
            })

    def analyze_table_structure(self, cells_data: str) -> Dict[str, Any]:
        data = json.loads(cells_data)
        if isinstance(data, list):
            sample_data = data[:5]
        else:
            sample_data = data

        prompt = f"""Analyze the following Excel cells sample data and determine the header structure:
1. Identify if any rows are headers based on their content and format
2. Determine if it's a single or multiple header structure
3. List the row indices (0-based) that are headers

Consider these factors:
- Headers often contain column titles or category names
- Headers may use different formatting or cell types
- Headers are typically at the top of the table
- Headers should have meaningful text content

Sample data (first few rows):
{json.dumps(sample_data, indent=2)}

Respond in JSON format with the following structure:
{{
    "headerType": "single" or "multiple" or "none",
    "headerRows": [row_indices],
    "confidence": number,
    "reasoning": string
}}
"""
        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[{"role": "user", "content": prompt}],
                response_format={"type": "json_object"},
                max_tokens=1000
            )
            return response.choices[0].message.content
        except Exception as e:
            print(f"Error in analyze_table_structure: {str(e)}")
            return json.dumps({
                "isTable": False,
                "headerType": "none",
                "headerRowsCount": 0,
                "confidence": 0
            })

    def analyze_text_block(self, text_content: str) -> Dict[str, Any]:
        sample_text = text_content[:1000]

        prompt = f"""Analyze the following text content sample from an Excel sheet and determine:
1. The type of content (meeting notes, comments, descriptions, etc.)
2. Its importance level
3. A concise summary
4. Key points or takeaways

Text sample:
{sample_text}

Respond in JSON format with the following structure:
{{
    "contentType": string,
    "importance": "high" or "medium" or "low",
    "summary": string,
    "keyPoints": [string],
    "confidence": number
}}
"""
        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[{"role": "user", "content": prompt}],
                response_format={"type": "json_object"},
                max_tokens=1000
            )
            return response.choices[0].message.content
        except Exception as e:
            print(f"Error in analyze_text_block: {str(e)}")
            return json.dumps({
                "contentType": "unknown",
                "importance": "low",
                "summary": "Error in analysis",
                "keyPoints": [],
                "confidence": 0
            })

    def analyze_chart(self, chart_elements: str) -> Dict[str, Any]:
        sample_elements = chart_elements[:1000]

        prompt = f"""Analyze the following chart elements sample from Excel and determine:
1. The type of visualization
2. Its purpose or what it's trying to communicate
3. Related data references
4. Suggested ways to use or interpret the chart

Chart elements sample:
{sample_elements}

Respond in JSON format with the following structure:
{{
    "chartType": string,
    "purpose": string,
    "dataRelations": [string],
    "suggestedUsage": string,
    "confidence": number
}}
"""
        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[{"role": "user", "content": prompt}],
                response_format={"type": "json_object"},
                max_tokens=1000
            )
            return response.choices[0].message.content
        except Exception as e:
            print(f"Error in analyze_chart: {str(e)}")
            return json.dumps({
                "chartType": "unknown",
                "purpose": "Error in analysis",
                "dataRelations": [],
                "suggestedUsage": "Error in analysis",
                "confidence": 0
            })

    def analyze_merged_cells(self, merged_cells_data: str) -> Dict[str, Any]:
        sample_data = merged_cells_data[:1000]

        prompt = f"""Analyze the following merged cells data sample and determine:
1. The pattern or purpose of the cell merging
2. Any hierarchy or structure it might represent
3. Potential implications for data organization

Merged cells data sample:
{sample_data}

Respond in JSON format with the following structure:
{{
    "pattern": string,
    "purpose": string,
    "structureType": "header" or "grouping" or "formatting" or "other",
    "implications": [string],
    "confidence": number
}}
"""
        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[{"role": "user", "content": prompt}],
                response_format={"type": "json_object"},
                max_tokens=1000
            )
            return response.choices[0].message.content
        except Exception as e:
            print(f"Error in analyze_merged_cells: {str(e)}")
            return json.dumps({
                "pattern": "unknown",
                "purpose": "Error in analysis",
                "structureType": "other",
                "implications": [],
                "confidence": 0
            })

class ExcelMetadataExtractor:

    def __init__(self, file_obj):
        self.file_obj = file_obj
        self.workbook = load_workbook(file_obj, data_only=True)
        self.openai_helper = OpenAIHelper()
        self.MAX_CELLS_PER_ANALYSIS = 100
        self.ns = {
            'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
            'xdr': 'http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing',
            'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
            'c': 'http://schemas.openxmlformats.org/drawingml/2006/chart',
            'sp': 'http://schemas.openxmlformats.org/spreadsheetml/2006/main',
            'pr': 'http://schemas.openxmlformats.org/package/2006/relationships'
        }

    # ExcelMetadataExtractorクラスの残りのメソッドはそのまま...
    # (元のexcel_metadata_extractor.pyの内容をここに続ける)

def display_json_tree(data, key_prefix=""):
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
    try:
        st.markdown("### Region Information")
        st.write(f"Region Type: {region['regionType']}")
        st.write(f"Range: {region['range']}")

        if region['regionType'] == 'shape':
            st.markdown("### Shape Information")
            cols = st.columns(2)
            with cols[0]:
                if 'shape_type' in region and region['shape_type']:
                    st.metric("Shape Type", region.get('shape_type', 'Unknown').title())
                if region.get('name'):
                    st.text(f"Name: {region['name']}")
            with cols[1]:
                if region.get('description'):
                    st.text(f"Description: {region['description']}")

            if 'text_content' in region:
                st.markdown("### Text Content")
                st.text(region['text_content'])

        elif region['regionType'] in ['image', 'smartart', 'chart']:
            st.markdown("Drawing Information")
            cols = st.columns(2)
            with cols[0]:
                st.metric("Type", region['type'].title())
                if region.get('name'):
                    st.text(f"Name: {region['name']}")
            with cols[1]:
                if region.get('description'):
                    st.text(f"Description: {region['description']}")

            if 'coordinates' in region:
                st.markdown("Position:")
                coords = region['coordinates']
                st.text(f"From: Column {coords['from']['col']}, Row {coords['from']['row']}")
                st.text(f"To: Column {coords['to']['col']}, Row {coords['to']['row']}")

            if region['type'] == 'image':
                if 'image_ref' in region:
                    st.markdown("Image Details:")
                    st.text(f"Reference: {region['image_ref']}")
            elif region['type'] == 'smartart':
                if 'diagram_type' in region:
                    st.markdown("SmartArt Details:")
                    st.text(f"Diagram Type: {region['diagram_type']}")
            elif region['type'] == 'chart':
                if 'chart_ref' in region:
                    st.markdown("Chart Details:")
                    st.text(f"Chart Reference: {region['chart_ref']}")

        elif region['regionType'] == 'table':
            st.markdown("### Table Information")
            if 'headerStructure' in region:
                st.markdown("#### Header Structure")
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
                
                if 'sampleCells' in region and len(region['sampleCells']) > 0:
                    st.markdown("#### Header Columns")
                    header_row = region['sampleCells'][0]
                    for cell in header_row:
                        if cell['value']:
                            st.text(f"Column {get_column_letter(cell['col'])}: {cell['value']}")

        elif region['regionType'] == 'text':
            st.markdown("Text Information")
            if 'content' in region:
                st.text_area("Content", region['content'], height=100)
            if 'classification' in region:
                st.write("Classification:", region['classification'])
            if 'importance' in region:
                st.write("Importance:", region['importance'])

        if 'mergedCells' in region and region['mergedCells']:
            st.markdown("Merged Cells")
            for merged in region['mergedCells']:
                st.text(f"Range: {merged['range']} - Value: {merged.get('value', 'N/A')}")

    except Exception as e:
        st.error(f"Error displaying region info: {str(e)}")
        st.error(f"Region data: {json.dumps(region, indent=2)}")
        st.error(f"Stack trace:\n{traceback.format_exc()}")

def main():
    st.set_page_config(page_title="Excel Metadata Extractor",
                       page_icon="📊",
                       layout="wide")

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
            extractor = ExcelMetadataExtractor(uploaded_file)
            metadata = extractor.extract_all_metadata()

            st.header("📑 Extracted Metadata")

            with st.expander("📌 File Properties", expanded=True):
                st.json(metadata["fileProperties"])

            for sheet_idx, sheet in enumerate(metadata["worksheets"]):
                st.subheader(f"📚 Sheet: {sheet['sheetName']}")

                cols = st.columns(3)
                with cols[0]:
                    st.metric("Rows", sheet["rowCount"])
                with cols[1]:
                    st.metric("Columns", sheet["columnCount"])
                with cols[2]:
                    st.metric("Merged Cells", len(sheet.get("mergedCells", [])))

                if sheet.get("mergedCells"):
                    st.markdown("##### 🔀 Merged Cells")
                    st.code("\n".join(sheet["mergedCells"]))

                if "regions" in sheet and sheet["regions"]:
                    st.markdown("##### 📍 Detected Regions")
                    for region in sheet["regions"]:
                        try:
                            with st.expander(f"{region['regionType'].title()} Region - {region['range']}"):
                                display_region_info(region)
                        except Exception as e:
                            st.error(f"Error processing region: {str(e)}\nRegion data: {json.dumps(region, indent=2)}")
                            st.error(f"Stack trace:\n{traceback.format_exc()}")

                st.markdown("---")

            with st.expander("🔍 Raw JSON Data"):
                st.json(metadata)

            json_str = json.dumps(metadata, indent=2)
            st.download_button(label="Download JSON",
                               data=json_str,
                               file_name=f"{uploaded_file.name}_metadata.json",
                               mime="application/json")

        except Exception as e:
            st.error(f"Error processing file: {str(e)}")
            st.error(f"Detailed error:\n{traceback.format_exc()}")
            st.error("Please make sure you've uploaded a valid Excel file.")

if __name__ == "__main__":
    main()
