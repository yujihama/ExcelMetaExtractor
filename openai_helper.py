import os
from openai import OpenAI
from typing import Dict, Any, List

class OpenAIHelper:
    def __init__(self):
        self.client = OpenAI(api_key=os.environ.get("OPENAI_API_KEY"))
        # the newest OpenAI model is "gpt-4o" which was released May 13, 2024.
        # do not change this unless explicitly requested by the user
        self.model = "gpt-4o"

    def analyze_table_structure(self, cells_data: str) -> Dict[str, Any]:
        """Analyze table structure using LLM"""
        prompt = f"""Analyze the following Excel cells data and determine:
1. If this is a table structure
2. The type of headers (single/multiple rows)
3. Any special patterns or purposes

Data:
{cells_data}

Respond in JSON format with the following structure:
{{
    "isTable": boolean,
    "headerType": "single" or "multiple" or "none",
    "headerRowsCount": number,
    "purpose": string,
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
3. Any key points or summaries

Text:
{text_content}

Respond in JSON format with the following structure:
{{
    "contentType": string,
    "importance": "high" or "medium" or "low",
    "summary": string,
    "keyPoints": [string]
}}
"""
        response = self.client.chat.completions.create(
            model=self.model,
            messages=[{"role": "user", "content": prompt}],
            response_format={"type": "json_object"}
        )
        return response.choices[0].message.content

    def analyze_chart(self, chart_elements: Dict[str, Any]) -> Dict[str, Any]:
        """Analyze chart elements using LLM"""
        prompt = f"""Analyze the following chart elements from Excel and determine:
1. The type of visualization
2. Its purpose
3. Related data references

Chart elements:
{chart_elements}

Respond in JSON format with the following structure:
{{
    "chartType": string,
    "purpose": string,
    "dataRelations": [string],
    "suggestedUsage": string
}}
"""
        response = self.client.chat.completions.create(
            model=self.model,
            messages=[{"role": "user", "content": prompt}],
            response_format={"type": "json_object"}
        )
        return response.choices[0].message.content
