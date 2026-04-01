"""
pip install requests
"""
import requests

OLLAMA_URL = "http://localhost:11434"

DEFAULT_PROMPT = """Summarize the following document section concisely, preserving the key ideas:

{document}"""

PRESET_PROMPTS = {
    "Summarize":           DEFAULT_PROMPT,
    "Key points":          "Extract the main key points from the following text as a concise bullet list:\n\n{document}",
    "Critical analysis":   "Provide a brief critical analysis of the following text, noting strengths and weaknesses:\n\n{document}",
    "Translate to Spanish":"Translate the following text to Spanish:\n\n{document}",
    "Translate to English":"Translate the following text to English:\n\n{document}",
}


def list_models() -> list[str]:
    try:
        r = requests.get(f"{OLLAMA_URL}/api/tags", timeout=5)
        r.raise_for_status()
        return [m["name"] for m in r.json().get("models", [])]
    except Exception:
        return []


def run_task(text: str, prompt_template: str, model: str) -> str:
    prompt = prompt_template.replace("{document}", text)
    r = requests.post(
        f"{OLLAMA_URL}/api/generate",
        json={"model": model, "prompt": prompt, "stream": False},
        timeout=120,
    )
    r.raise_for_status()
    return r.json()["response"].strip()
