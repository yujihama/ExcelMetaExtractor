import os
from openai import OpenAI
from typing import Dict, Any, Union, List
import json
import streamlit as st


class OpenAIHelper:

    def __init__(self):
        self.client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))
        self.model = "gpt-4o"

    def summarize_region(self, region: Dict[str, Any]) -> str:
        """Generate a summary for a region based on its content"""
        try:
            if region["regionType"] == "table":
                # テーブル用のサマリー生成
                cells = region.get("sampleCells", [])
                header_structure = region.get("headerStructure", {})
                prompt = f"""以下のExcelテーブル領域の内容を簡潔に要約してください:
                ヘッダー構造: {json.dumps(header_structure, ensure_ascii=False)}
                データサンプル: {json.dumps(cells[:2], ensure_ascii=False)}
                """
            else:
                # その他の領域用のサマリー生成
                prompt = f"""以下のExcel領域の内容を簡潔に要約してください:
                領域タイプ: {region["regionType"]}
                範囲: {region["range"]}
                内容: {json.dumps(region, ensure_ascii=False)[:200]}
                """

            response = self.client.chat.completions.create(model=self.model,
                                                           messages=[{
                                                               "role":
                                                               "user",
                                                               "content":
                                                               prompt
                                                           }],
                                                           max_tokens=150)
            return response.choices[0].message.content
        except Exception as e:
            print(f"Error generating summary: {str(e)}")
            return "サマリーの生成に失敗しました"

    def analyze_region_type(self, region_data: str) -> Dict[str, Any]:
        """Analyze region type using LLM with size limits"""
        data = json.loads(region_data)
        sample_data = {
            "cells": data["cells"][:5],
            "mergedCells": data.get("mergedCells", [])[:3]
        }

        prompt = f"""Analyze the following Excel region sample data and determine:
1. The type of region (table, text, chart, image)
2. If it contains a table title or document heading
3. The purpose or meaning of the content, considering Japanese text patterns

Region sample data (first few rows/cells):
{json.dumps(sample_data, indent=2)}

Consider Japanese text patterns like:
- Table titles (〇〇一覧, △△表, □□リスト)
- Section headings (大項目, 中項目, 小項目)
- Data categories (区分, 分類, 種別)

Respond in JSON format:
{{
    "regionType": "table" or "text" or "chart" or "image",
    "title": {{
        "detected": boolean,
        "content": string or null,
        "row": number or null
    }},
    "characteristics": [string],
    "purpose": string,
    "confidence": number
}}
"""
        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[{
                    "role": "user",
                    "content": prompt
                }],
                response_format={"type": "json_object"},
                max_tokens=1000)
            with st.expander("🔍 Analyzed Region Data"):
                st.write(prompt)
                st.write(response.choices[0].message.content)
            return response.choices[0].message.content
        except Exception as e:
            print(f"Error in analyze_region_type: {str(e)}")
            return json.dumps({
                "regionType": "unknown",
                "title": {
                    "detected": False,
                    "content": None,
                    "row": None
                },
                "characteristics": [],
                "purpose": "Error in analysis",
                "confidence": 0
            })

    def analyze_table_structure_with_hints(self, data: str) -> Dict[str, Any]:
        """Analyze table structure using LLM with hints about potential header rows"""
        try:
            prompt = f"""与えられたテーブルのサンプルデータとヘッダー行の候補から、以下を分析してください:
1. テーブルのヘッダー構造（単一行のヘッダー/複合ヘッダー）
2. ヘッダー行の正確な位置（複合ヘッダーの場合は複数行回答）
3. 各列の種類と意味

データ:
{data}

ヘッダーの判断基準:
- 一覧表やマスタ等の表題
- 列見出しの階層構造
- データ分類や単位の記載
- 結合セルの使用パターン

JSON形式で返答してください:
{{
    "headerType": "single" または "multiple" または "none",
    "headerRows": [行番号のリスト],
    "confidence": 0-1の数値
}}"""

            response = self.client.chat.completions.create(
                model=self.model,
                messages=[{
                    "role": "user",
                    "content": prompt
                }],
                response_format={"type": "json_object"})
            with st.expander("🔍 Analyzed Table Structure"):
                st.write(prompt)
                st.write(response.choices[0].message.content)

            return response.choices[0].message.content
        except Exception as e:
            print(f"Error in analyze_table_structure_with_hints: {str(e)}")
            return json.dumps({
                "headerType": "none",
                "headerRows": [],
                "confidence": 0
            })

    def analyze_table_structure(self, cells_data: str,
                                merged_cells) -> Dict[str, Any]:
        """Analyze table structure using LLM with size limits"""
        # data = json.loads(cells_data)
        # if isinstance(data, list):
        #     sample_data = data[:5]
        # else:
        #     sample_data = data

        prompt = f"""Analyze the following Excel cells sample data and determine:
1. Title row detection (例: 売上実績表, 商品マスタ一覧)
2. Header structure (single/multiple header rows)

ヘッダーの判断基準:
- 一覧表やマスタ等の表題
- 列見出しの階層構造
- データ分類や単位の記載
- 結合セルの使用
- 合計行や総計、小計の行はヘッダーに含めないこと

Sample data(Refer to the rows and columns (row and col) for accurate interpretation of the structure):

{cells_data}

また、以下のセルは結合されているのでヘッダー検知の参考にしてください。
{merged_cells}

Respond in JSON format:
{{
    "titleRow": {{
        "detected": boolean,
        "content": string or null,
        "row": number or null
    }},
    "headerStructure": {{
        "type": "single" or "multiple" or "none",
        "rows": [row_indices]
        "reason": if you ansewered "multiple", please explain why
    }},
    "confidence": number
}}
"""
        try:
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[{
                    "role": "user",
                    "content": prompt,
                }],
                response_format={"type": "json_object"},
                temperature=0)
            #max_tokens=1000)
            with st.expander("🔍 Analyzed Table Structure"):
                st.write(prompt)
                st.write(response.choices[0].message.content)
            return response.choices[0].message.content
        except Exception as e:
            print(f"Error in analyze_table_structure: {str(e)}")
            return json.dumps({
                "titleRow": {
                    "detected": False,
                    "content": None,
                    "row": None
                },
                "headerStructure": {
                    "type": "none",
                    "rows": [],
                    "hierarchy": None
                },
                "columns": [],
                "confidence": 0
            })

    def analyze_text_block(self, text_content: str) -> Dict[str, Any]:
        """Analyze text block content using LLM with size limits"""
        sample_text = text_content[:1000]

        prompt = f"""Analyze the following text content from an Excel sheet:
1. Type of content (議事録, コメント, 説明文など)
2. Importance level
3. Key points or takeaways
4. Consider Japanese text patterns and business terms

Text sample:
{sample_text}

Respond in JSON format:
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
                messages=[{
                    "role": "user",
                    "content": prompt
                }],
                response_format={"type": "json_object"},
                max_tokens=1000)
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
        """Analyze chart elements using LLM with size limits"""
        sample_elements = chart_elements[:1000]

        prompt = f"""Analyze the following chart elements from Excel:
1. Type of visualization
2. Purpose and key message
3. Data relationships
4. Consider Japanese chart titles and labels

Chart elements:
{sample_elements}

Respond in JSON format:
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
                messages=[{
                    "role": "user",
                    "content": prompt
                }],
                response_format={"type": "json_object"},
                max_tokens=1000)
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
        """Analyze merged cells pattern using LLM with size limits"""
        sample_data = merged_cells_data[:1000]

        prompt = f"""Analyze the following merged cells pattern:
1. Purpose of merging (見出し, グループ化, 整形など)
2. Structural implications
3. Consider Japanese organizational patterns

Merged cells data:
{sample_data}

Respond in JSON format:
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
                messages=[{
                    "role": "user",
                    "content": prompt
                }],
                response_format={"type": "json_object"},
                max_tokens=1000)
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
