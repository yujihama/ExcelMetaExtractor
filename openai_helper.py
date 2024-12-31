import os
from openai import OpenAI
from typing import Dict, Any, Union, List
import json

class OpenAIHelper:
    def __init__(self):
        self.client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))
        # the newest OpenAI model is "gpt-4o" which was released May 13, 2024.
        # do not change this unless explicitly requested by the user
        self.model = "gpt-4o"

    def analyze_region_type(self, region_data: str) -> Dict[str, Any]:
        """Analyze region type using LLM with size limits"""
        # データサイズを制限
        data = json.loads(region_data)
        sample_data = {
            "cells": data["cells"][:5],  # 最初の5行のみ
            "mergedCells": data.get("mergedCells", [])[:3]  # 最初の3個の結合セルのみ
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
                max_tokens=1000  # トークン数を制限
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
        """Analyze table structure using LLM with size limits"""
        # データを制限
        data = json.loads(cells_data)
        if isinstance(data, list):
            sample_data = data[:5]  # 最初の5行のみ
        else:
            sample_data = data

        prompt = f"""Analyze the following Excel cells sample data and determine:
1. If this is a table structure
2. The type of headers (single/multiple rows)
3. The number of header rows

Sample data (first few rows):
{json.dumps(sample_data, indent=2)}

Respond in JSON format with the following structure:
{{
    "isTable": boolean,
    "headerType": "single" or "multiple" or "none",
    "headerRowsCount": number,
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
            print(f"Error in analyze_table_structure: {str(e)}")
            return json.dumps({
                "isTable": False,
                "headerType": "none",
                "headerRowsCount": 0,
                "confidence": 0
            })

    def analyze_text_block(self, text_content: str) -> Dict[str, Any]:
        """Analyze text block content using LLM with size limits"""
        # データサイズを制限
        sample_text = text_content[:1000] #最初の1000文字のみ

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
        """Analyze chart elements using LLM with size limits"""
        # データサイズを制限
        sample_elements = chart_elements[:1000] # 最初の1000文字のみ

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
        """Analyze merged cells pattern using LLM with size limits"""
        # データサイズを制限
        sample_data = merged_cells_data[:1000] # 最初の1000文字のみ

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