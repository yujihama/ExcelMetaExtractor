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
        """Analyze region type using LLM"""
        prompt = f"""Analyze the following Excel region data and determine:
1. The type of region (table, text, chart, image)
2. Any special characteristics or patterns
3. The purpose or meaning of the content

Region data:
{region_data}

Respond in JSON format with the following structure:
{{
    "regionType": "table" or "text" or "chart" or "image",
    "characteristics": [string],
    "purpose": string,
    "confidence": number,
    "chartElements": object (only for chart type)
}}
"""
        response = self.client.chat.completions.create(
            model=self.model,
            messages=[{"role": "user", "content": prompt}],
            response_format={"type": "json_object"}
        )
        return response.choices[0].message.content

    def analyze_table_structure(self, cells_data: str) -> Dict[str, Any]:
        """Analyze table structure using LLM"""
        prompt = f"""Analyze the following Excel cells data and determine:
1. If this is a table structure
2. The type of headers (single/multiple rows)
3. The number of header rows
4. Any special patterns or purposes
5. Your confidence in this analysis

Data:
{cells_data}

Respond in JSON format with the following structure:
{{
    "isTable": boolean,
    "headerType": "single" or "multiple" or "none",
    "headerRowsCount": number,
    "purpose": string,
    "specialPatterns": [string],
    "confidence": number
}}
"""
        response = self.client.chat.completions.create(
            model=self.model,
            messages=[{"role": "user", "content": prompt}],
            response_format={"type": "json_object"}
        )
        return response.choices[0].message.content

    def analyze_text_block(self, text_content: str) -> Dict[str, Any]:
        """Analyze text block content using LLM"""
        prompt = f"""Analyze the following text content from an Excel sheet and determine:
1. The type of content (meeting notes, comments, descriptions, etc.)
2. Its importance level
3. A concise summary
4. Key points or takeaways

Text:
{text_content}

Respond in JSON format with the following structure:
{{
    "contentType": string,
    "importance": "high" or "medium" or "low",
    "summary": string,
    "keyPoints": [string],
    "confidence": number
}}
"""
        response = self.client.chat.completions.create(
            model=self.model,
            messages=[{"role": "user", "content": prompt}],
            response_format={"type": "json_object"}
        )
        return response.choices[0].message.content

    def analyze_chart(self, chart_elements: str) -> Dict[str, Any]:
        """Analyze chart elements using LLM"""
        prompt = f"""Analyze the following chart elements from Excel and determine:
1. The type of visualization
2. Its purpose or what it's trying to communicate
3. Related data references
4. Suggested ways to use or interpret the chart

Chart elements:
{chart_elements}

Respond in JSON format with the following structure:
{{
    "chartType": string,
    "purpose": string,
    "dataRelations": [string],
    "suggestedUsage": string,
    "confidence": number
}}
"""
        response = self.client.chat.completions.create(
            model=self.model,
            messages=[{"role": "user", "content": prompt}],
            response_format={"type": "json_object"}
        )
        return response.choices[0].message.content

    def analyze_merged_cells(self, merged_cells_data: str) -> Dict[str, Any]:
        """Analyze merged cells pattern using LLM"""
        prompt = f"""Analyze the following merged cells data and determine:
1. The pattern or purpose of the cell merging
2. Any hierarchy or structure it might represent
3. Potential implications for data organization

Merged cells data:
{merged_cells_data}

Respond in JSON format with the following structure:
{{
    "pattern": string,
    "purpose": string,
    "structureType": "header" or "grouping" or "formatting" or "other",
    "implications": [string],
    "confidence": number
}}
"""
        response = self.client.chat.completions.create(
            model=self.model,
            messages=[{"role": "user", "content": prompt}],
            response_format={"type": "json_object"}
        )
        return response.choices[0].message.content