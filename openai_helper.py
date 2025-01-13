import os
import json
import traceback
from typing import Dict, Any, Union, List
from openai import OpenAI
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
                cells = region.get("sampleCells", [])
                header_structure = region.get("headerStructure", {})
                prompt = ("以下のExcelテーブル領域が何について記載されているか簡潔に説明してください:\n"
                          "ヘッダー構造: %s\n"
                          "データサンプル: %s") % (
                              json.dumps(header_structure, ensure_ascii=False),
                              json.dumps(cells[:2], ensure_ascii=False))
            elif region["regionType"] == "chart":
                prompt = ("以下のグラフが何について記載されているか簡潔に説明してください:\n"
                          "グラフタイプ: %s\n"
                          "データ範囲: %s\n"
                          "内容: %s") % (region['chartType'],
                                       region['series'][0]['data_range'],
                                       region['chart_data_json'])
            elif region["regionType"] == "image":
                gpt4o_analysis = region.get("gpt4o_analysis", {})
                prompt = ("以下の画像について簡潔に説明してください:\n"
                          "画像の種類: %s\n"
                          "内容: %s\n"
                          "特徴: %s\n"
                          "位置: %s\n"
                          "名前: %s\n"
                          "説明: %s") % (gpt4o_analysis.get('imageType', '不明'),
                                       gpt4o_analysis.get('content', '不明'),
                                       ', '.join(
                                           gpt4o_analysis.get('features', [])),
                                       region['range'], region.get('name', ''),
                                       region.get('description', ''))
            elif region["regionType"] == "shape":
                prompt = ("以下のExcelの図形が何について記載されているか簡潔に説明してください:\n"
                          "内容: %s") % json.dumps(region, ensure_ascii=False)
            else:
                prompt = ("以下のExcel領域が何について記載されているか簡潔に説明してください:\n"
                          "領域タイプ: %s\n"
                          "範囲: %s\n"
                          "内容: %s") % (region['regionType'], region['range'],
                                       json.dumps(region,
                                                  ensure_ascii=False)[:200])

            response = self.client.chat.completions.create(model="gpt-4o",
                                                           messages=[{
                                                               "role":
                                                               "user",
                                                               "content":
                                                               prompt
                                                           }],
                                                           max_tokens=1000)
            return response.choices[0].message.content
        except Exception as e:
            print(f"Error generating summary: {str(e)}")
            return "サマリーの生成に失敗しました"

    def analyze_region_type(self, region_data: str) -> Dict[str, Any]:
        """Analyze region type using LLM with size limits"""
        try:
            data = json.loads(region_data)
            sample_data = {
                "cells": data["cells"][:5],
                "mergedCells": data.get("mergedCells", [])[:3]
            }

            prompt = """
Analyze the following Excel region sample data and determine:
1. The type of region (table, text, chart, image)
2. If it contains a table title or document heading
3. The purpose or meaning of the content, considering Japanese text patterns

Region sample data (first few rows/cells):
{data}

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
""".format(data=json.dumps(sample_data, indent=2))

            response = self.client.chat.completions.create(
                model="gpt-4o",
                messages=[{
                    "role": "user",
                    "content": prompt
                }],
                response_format={"type": "json_object"},
                max_tokens=2000)
            return json.loads(response.choices[0].message.content)
        except Exception as e:
            print(f"Error in analyze_region_type: {str(e)}")
            return {
                "regionType": "unknown",
                "title": {
                    "detected": False,
                    "content": None,
                    "row": None
                },
                "characteristics": [],
                "purpose": "Error in analysis",
                "confidence": 0
            }

    def analyze_table_structure(self, cells_data: str,
                                merged_cells) -> Dict[str, Any]:
        """Analyze table structure using LLM with size limits"""
        prompt = """
Analyze the following Excel cells sample data and determine:
1. Title row detection (例: 売上実績表, 商品マスタ一覧)
2. Header structure (single/multiple header rows)

ヘッダーの判断基準:
- 一覧表やマスタ等の表題
- 列見出しの階層構造
- データ分類や単位の記載
- 結合セルの使用
- 合計行や総計、小計の行はヘッダーに含めないこと

Sample data(Refer to the rows and columns (row and col) for accurate interpretation of the structure):

{cells}

また、以下のセルは結合されているのでヘッダー検知の参考にしてください。
{merged}

Respond in JSON format:
{{
    "titleRow": {{
        "detected": boolean,
        "content": string or null,
        "row": number or null
    }},
    "headerStructure": {{
        "type": "single" or "multiple" or "none",
        "rows": [row_indices],
        "reason": string
    }},
    "confidence": number
}}
""".format(cells=cells_data, merged=merged_cells)

        try:
            response = self.client.chat.completions.create(
                model="gpt-4o",
                messages=[{
                    "role": "user",
                    "content": prompt
                }],
                response_format={"type": "json_object"},
                temperature=0)
            return json.loads(response.choices[0].message.content)
        except Exception as e:
            print(f"Error in analyze_table_structure: {str(e)}")
            return {
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
            }

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
                        "%s (%s): %s" % (region_type, region_range, summary))

            prompt = """以下はExcelシートから抽出された情報です。各領域の情報をよく理解したうえで記載されていることを客観的に推測を交えずに説明してください。:
シート名: %s
検出された領域数: %s

各領域の要約:
%s

以下の点に注目して要約してください:
- シートの主な目的や内容
- 含まれる主要なテーブルや図形
- 各領域の関係性

特に以下にはよく注意してください:
- 表/グラフの見た目や内容の特徴をしっかりとらえて要約してください。
- sheetに含まれていない情報は含めないでください。
- 推測で記載しないでください。
""" % (sheet_data.get('sheetName',
                      ''), len(regions), "\n".join(region_summaries))
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
        """GPT-4 Vision APIを使用して画像を分析"""
        try:
            prompt = """
この画像について以下の点を分析してください：
1. 画像の種類（グラフ、図表、写真など）
2. 主な内容や目的
3. 特徴的な要素

以下の形式でJSON形式で回答してください：
{
    "imageType": "graph/table/photo/other",
    "content": "画像の内容の説明",
    "features": ["特徴1", "特徴2", ...]
}
"""
            try:
                # APIリクエストのデバッグ情報
                print("\nSending request to GPT-4 Vision API...")
                print(f"Image data length: {len(base64_image)}")

                response = self.client.chat.completions.create(
                    model="gpt-4o",
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
                    max_tokens=500,
                    response_format={"type": "json_object"})

                # APIレスポンスのデバッグ情報
                print("\nGPT-4 Vision API Response:")
                print(f"Response status: Success")
                print(
                    f"Response content: {response.choices[0].message.content}")

                # レスポンスのパース
                result = json.loads(response.choices[0].message.content)

                # 結果の検証
                if not isinstance(result, dict):
                    raise ValueError("API response is not a dictionary")

                required_keys = ["imageType", "content", "features"]
                missing_keys = [
                    key for key in required_keys if key not in result
                ]
                if missing_keys:
                    raise ValueError(
                        f"Missing required keys in API response: {missing_keys}"
                    )

                return result

            except json.JSONDecodeError as json_error:
                print(f"\nJSON Decode Error: {str(json_error)}")
                print(
                    f"Raw response content: {response.choices[0].message.content}"
                )
                raise

            except Exception as api_error:
                print(f"\nAPI Error: {str(api_error)}")
                print(f"Error type: {type(api_error)}")
                if hasattr(api_error, 'response'):
                    print(f"API error response: {api_error.response}")
                raise

        except Exception as e:
            print(f"\nError in analyze_image_with_gpt4o: {str(e)}")
            print(f"Error type: {type(e)}")
            print(f"Stack trace:\n{traceback.format_exc()}")

            return {
                "imageType": "unknown",
                "content": f"画像分析に失敗しました: {str(e)}",
                "features": []
            }
