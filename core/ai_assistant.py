"""AI Assistant module for intelligent column analysis.
Supports Local Smart Heuristic AI as well as Cloud LLMs (Google Gemini & OpenAI).
"""

import json
import os
import re
import urllib.request
import urllib.error
from typing import Dict, Any, Optional, List

CONFIG_FILE = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "config.json")


def load_config() -> Dict[str, Any]:
    """Loads application config from config.json."""
    defaults = {
        "ai_provider": "local",  # 'local', 'gemini', 'openai'
        "gemini_api_key": "",
        "gemini_model": "gemini-1.5-flash",
        "openai_api_key": "",
        "openai_model": "gpt-4o-mini",
        "jpeg_quality": 95,
        "duplicate_mode": "once",  # 'once' (save once, no suffix _), 'overwrite'
        "subfolders_per_sheet": True,
        "theme": "dark"
    }
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                saved = json.load(f)
                defaults.update(saved)
        except Exception:
            pass
    return defaults


def save_config(cfg: Dict[str, Any]) -> bool:
    """Saves application config to config.json."""
    try:
        with open(CONFIG_FILE, "w", encoding="utf-8") as f:
            json.dump(cfg, f, indent=2)
        return True
    except Exception:
        return False


def call_gemini_api(api_key: str, model: str, prompt: str) -> Optional[str]:
    """Calls Gemini REST API without extra third-party SDK dependencies."""
    if not api_key:
        raise ValueError("Gemini API key is required.")
        
    model = model or "gemini-1.5-flash"
    url = f"https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent?key={api_key}"
    
    payload = {
        "contents": [{
            "parts": [{"text": prompt}]
        }],
        "generationConfig": {
            "temperature": 0.1,
            "maxOutputTokens": 800
        }
    }
    
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=data,
        headers={"Content-Type": "application/json"}
    )
    
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            body = resp.read().decode("utf-8")
            res_json = json.loads(body)
            candidates = res_json.get("candidates", [])
            if candidates:
                parts = candidates[0].get("content", {}).get("parts", [])
                if parts:
                    return parts[0].get("text", "")
    except urllib.error.HTTPError as e:
        err_msg = e.read().decode("utf-8", errors="ignore")
        raise RuntimeError(f"Gemini API Error ({e.code}): {err_msg}")
    except Exception as e:
        raise RuntimeError(f"Failed to connect to Gemini API: {e}")
        
    return None


def call_openai_api(api_key: str, model: str, prompt: str) -> Optional[str]:
    """Calls OpenAI REST API without extra dependencies."""
    if not api_key:
        raise ValueError("OpenAI API key is required.")
        
    model = model or "gpt-4o-mini"
    url = "https://api.openai.com/v1/chat/completions"
    
    payload = {
        "model": model,
        "messages": [
            {"role": "system", "content": "You are an expert Excel data structure analyzer."},
            {"role": "user", "content": prompt}
        ],
        "temperature": 0.1,
        "max_tokens": 500
    }
    
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(
        url,
        data=data,
        headers={
            "Content-Type": "application/json",
            "Authorization": f"Bearer {api_key}"
        }
    )
    
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            body = resp.read().decode("utf-8")
            res_json = json.loads(body)
            choices = res_json.get("choices", [])
            if choices:
                return choices[0].get("message", {}).get("content", "")
    except urllib.error.HTTPError as e:
        err_msg = e.read().decode("utf-8", errors="ignore")
        raise RuntimeError(f"OpenAI API Error ({e.code}): {err_msg}")
    except Exception as e:
        raise RuntimeError(f"Failed to connect to OpenAI API: {e}")
        
    return None


def analyze_columns_with_ai(
    sheet_name: str,
    headers: List[str],
    sample_rows: List[List[Any]],
    provider: str = "local",
    api_key: str = "",
    model_name: str = ""
) -> Dict[str, Any]:
    """Uses LLM (or Local AI) to identify Design Number and Image columns."""
    
    # If local, return signal indicating caller should use heuristic engine
    if provider == "local" or not api_key:
        return {
            "success": False,
            "provider": "local",
            "message": "Local Smart Heuristic AI is active."
        }

    # Prepare compact table representation
    rows_text = []
    header_str = " | ".join([f"Col {i} ({h})" for i, h in enumerate(headers)])
    rows_text.append(f"HEADERS: {header_str}")
    for idx, r in enumerate(sample_rows[:6]):
        r_str = " | ".join([str(val) if val is not None else "" for val in r[:len(headers)]])
        rows_text.append(f"ROW {idx+1}: {r_str}")
        
    prompt = f"""You are analyzing an Excel spreadsheet sheet named '{sheet_name}'.
The user needs to extract product photos from this sheet and save them as JPG named with their Design Number / Product Code.
The two columns ('Design Number' and 'Image') can be located anywhere in the sheet.

Here is a sample of the sheet columns and first rows:
{chr(10).join(rows_text)}

Identify:
1. Which column (0-indexed index) contains the Design Number / Item Code / Style Number?
2. Which column (0-indexed index) contains the Image / Picture / Photo?

Return your answer strictly as a valid JSON object matching this schema:
{{
  "design_col_index": <int, 0-indexed column>,
  "design_col_name": "<string, column name or header>",
  "image_col_index": <int, 0-indexed column>,
  "image_col_name": "<string, column name or header>",
  "confidence": <float between 0.0 and 1.0>,
  "reasoning": "<concise explanation>"
}}
Do not include any markdown formatting or extra text outside the JSON.
"""

    if provider == "gemini":
        raw_text = call_gemini_api(api_key, model_name, prompt)
    elif provider == "openai":
        raw_text = call_openai_api(api_key, model_name, prompt)
    else:
        raise ValueError(f"Unknown AI provider: {provider}")

    if not raw_text:
        raise RuntimeError("AI returned an empty response.")

    # Clean markdown json code fences if present
    clean_text = raw_text.strip()
    clean_text = re.sub(r'^```json\s*', '', clean_text, flags=re.IGNORECASE)
    clean_text = re.sub(r'^```\s*', '', clean_text)
    clean_text = re.sub(r'\s*```$', '', clean_text)

    try:
        parsed = json.loads(clean_text)
        parsed["success"] = True
        parsed["provider"] = provider
        return parsed
    except json.JSONDecodeError:
        # Try extracting JSON object using regex
        match = re.search(r'\{.*\}', clean_text, re.DOTALL)
        if match:
            parsed = json.loads(match.group(0))
            parsed["success"] = True
            parsed["provider"] = provider
            return parsed
        raise RuntimeError(f"Could not parse AI response as JSON: {raw_text[:200]}")
