import os
from openai import OpenAI
from typing import Dict, Any, Union, List
import json
import streamlit as st
from dotenv import load_dotenv


class OpenAIHelper:

    def __init__(self):
        load_dotenv()
        self.client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))
        self.model = "gpt-4o"

    def summarize_region(self, region: Dict[str, Any]) -> str:
        """Generate a summary for a region based on its content"""
        try:
            if region["regionType"] == "table":
                # テーブル用のサマリー生成
                cells = region.get("sampleCells", [])
                header_structure = region.get("headerStructure", {})
                prompt = f"""以下のExcelテーブル領域が何について記載されているか簡潔に説明してください:
                ヘッダー構造: {json.dumps(header_structure, ensure_ascii=False)}
                データサンプル: {json.dumps(cells[:2], ensure_ascii=False)}
                """
            elif region["regionType"] == "chart":
                # チャート用のサマリー生成
                prompt = f"""以下のグラフが何について記載されているか簡潔に説明してください:
                グラフタイプ: {region["chartType"]}
                データ範囲: {region["series"][0]["data_range"]}
                内容: {region["chart_data_json"]}
                """
            elif region["regionType"] == "image":
                # 画像用のサマリー生成（GPT-4の分析結果を含む）
                gpt4o_analysis = region.get("gpt4o_analysis", {})
                prompt = f"""以下の画像について簡潔に説明してください:
                画像の種類: {gpt4o_analysis.get("imageType", "不明")}
                内容: {gpt4o_analysis.get("content", "不明")}
                特徴: {", ".join(gpt4o_analysis.get("features", []))}
                位置: {region["range"]}
                名前: {region.get("name", "")}
                説明: {region.get("description", "")}
                """
            elif region["regionType"] == "shape":
                # 図形用のサマリー生成
                prompt = f"""以下のExcelの図形が何について記載されているか簡潔に説明してください:
                内容: {json.dumps(region, ensure_ascii=False)}
                """
            else:
                # その他の領域用のサマリー生成
                prompt = f"""以下のExcel領域が何について記載されているか簡潔に説明してください:
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
                                                           max_tokens=1000)
            # with st.expander("LLM_Summary"):
            #     st.write(prompt)
            #     st.write(response.choices[0].message.content)

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
- Table titles (一覧表, 集計表, リスト)
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
                max_tokens=2000)
            # with st.expander("🔍 Analyzed Region Data"):
            #     st.write(prompt)
            #     st.write(response.choices[0].message.content)
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
            # with st.expander("🔍 Analyzed Table Structure"):
            #     st.write(prompt)
            #     st.write(response.choices[0].message.content)
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

    def generate_sheet_summary(self, sheet_data: Dict[str, Any]) -> str:
        """Generate a summary for an entire sheet using LLM with region summaries already available."""
        try:
            regions = sheet_data.get('regions', [])
            region_summaries = []

            for region in regions:
                if "summary" in region:
                    region_type = region.get("regionType", "unknown")
                    region_range = region.get("range", "")
                    summary = region.get("summary", "")
                    region_summaries.append(
                        f"{region_type} ({region_range}): {summary}")

            prompt = f"""以下はExcelシートから抽出された情報です。各領域の情報をよく理解したうえで記載されていることを客観的に推測を交えずに説明してください。:
シート名: {sheet_data.get('sheetName', '')}
検出された領域数: {len(regions)}

各領域の要約:
{"\n".join(region_summaries)}

以下の点に注目して要約してください:
- シートの主な目的や内容
- 含まれる主要なテーブルや図形
- 各領域の関係性

特に以下にはよく注意してください:
- 表/グラフの見た目や内容の特徴をしっかりとらえて要約してください。
- sheetに含まれていない情報は含めないでください。
- 推測で記載しないでください。
"""
            response = self.client.chat.completions.create(model=self.model,
                                                           messages=[{
                                                               "role":
                                                               "user",
                                                               "content":
                                                               prompt
                                                           }],
                                                           max_tokens=2000)

            return response.choices[0].message.content
        except Exception as e:
            print(f"Error generating sheet summary: {str(e)}")
            return "シートのサマリー生成に失敗しました"

    def analyze_image_with_gpt4o(self, base64_image: str) -> Dict[str, Any]:
        """GPT-4oを使用して画像を分析"""
        try:
            prompt = """この画像について以下の点を分析してください：
1. 画像の種類類（グラフ、図表、写真など）
2. 主な内容や目的
3. 特徴的な要素

以下の形式でJSON形式で回答してください：
{
    "imageType": "graph/table/photo/other",
    "content": "画像の内容の説明",
    "features": ["特徴1", "特徴2", ...]
}
"""
            response = self.client.chat.completions.create(
                model=self.model,
                messages=[{
                    "role":
                    "user",
                    "content": [{
                        "type": "text",
                        "text": prompt
                    }, {
                        "type": "image_url",
                        "image_url": {
                            "url": f"data:image/png;base64,{base64_image}"
                        }
                    }]
                }],
                response_format={"type": "json_object"},
                max_tokens=2000)
            return json.loads(response.choices[0].message.content)
        except Exception as e:
            print(f"Error analyzing image with GPT-4o: {str(e)}")
            return {
                "imageType": "unknown",
                "content": "画像分析に失敗しました",
                "features": []
            }
