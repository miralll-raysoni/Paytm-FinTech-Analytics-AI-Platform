import importlib.util
import json
import os
import re
from pathlib import Path

# Load DISCLOSURE_SNIPPETS dynamically from local directory
PART_C_DIR = Path(
    "/Users/miralraysoni/Downloads/BitSoM - Fintech and AI/Final Assessment/Part C"
)


def load_module_from_path(module_name: str, file_path: Path):
    if not file_path.exists():
        raise FileNotFoundError(f"Required seed file not found: {file_path}")
    spec = importlib.util.spec_from_file_location(module_name, file_path)
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module


disclosure_mod = load_module_from_path(
    "disclosure_snippets", PART_C_DIR / "disclosure_snippets.py"
)
DISCLOSURE_SNIPPETS = disclosure_mod.DISCLOSURE_SNIPPETS


def validate_schema(data: dict) -> bool:
    """Validates that extracted output conforms to required JSON schema."""
    if not isinstance(data, dict):
        return False

    required_keys = {"risk_flags", "hedging_detected", "sentiment"}
    if not required_keys.issubset(data.keys()):
        return False

    if not isinstance(data["risk_flags"], list):
        return False

    if not isinstance(data["hedging_detected"], bool):
        return False

    if data["sentiment"] not in {"confident", "cautious", "neutral"}:
        return False

    return True


def extract_signals_mock(snippet: str) -> dict:
    """Deterministic keyword/regex mock logic for baseline grading."""
    text_lower = snippet.lower()

    # 1. Identify Risk Flags
    risk_flags = []
    if "litigation" in text_lower:
        risk_flags.append("litigation exposure")
    if "regulatory" in text_lower or "compliance" in text_lower:
        risk_flags.append("regulatory exposure")
    if (
        "customer" in text_lower
        or "revenue" in text_lower
        or "percent of total" in text_lower
    ):
        if re.search(r"\d+\s*percent", text_lower) or "account for" in text_lower:
            risk_flags.append("customer concentration risk")

    # 2. Identify Hedging Phrases
    hedging_keywords = ["assuming", "cautiously", "visibility"]
    hedging_detected = any(kw in text_lower for kw in hedging_keywords)

    # 3. Classify Sentiment
    if "confident" in text_lower or "approved" in text_lower:
        sentiment = "confident"
    elif hedging_detected:
        sentiment = "cautious"
    else:
        sentiment = "neutral"

    return {
        "risk_flags": risk_flags,
        "hedging_detected": hedging_detected,
        "sentiment": sentiment,
    }


def extract_signals(snippet: str) -> dict:
    mock_llm_flag = os.getenv("MOCK_LLM", "1")

    # Graded Mock Mode Baseline Path
    if mock_llm_flag == "1" or mock_llm_flag.lower() == "true":
        return extract_signals_mock(snippet)

    # Optional MOCK_LLM=0 Extension Path
    try:
        from groq import Groq

        client = Groq(api_key=os.getenv("GROQ_API_KEY"))

        system_prompt = (
            "You are a financial NLP expert. Analyze corporate disclosure text and respond ONLY with a JSON object matching this exact schema:\n"
            "{\n"
            '  "risk_flags": ["list of identified risk phrasings or keywords like litigation, regulatory, customer concentration"],\n'
            '  "hedging_detected": true/false,\n'
            '  "sentiment": "confident" | "cautious" | "neutral"\n'
            "}"
        )

        for attempt in range(2):  # Try once, retry once on validation failure
            response = client.chat.completions.create(
                model="llama-3.3-70b-versatile",
                messages=[
                    {"role": "system", "content": system_prompt},
                    {"role": "user", "content": f"Snippet: {snippet}"},
                ],
                temperature=0.0,
                response_format={"type": "json_object"},
            )

            data = json.loads(response.choices[0].message.content)
            if validate_schema(data):
                return data

        # Fall back to mock result if validation fails twice
        return extract_signals_mock(snippet)

    except Exception:
        # Fall back to mock result on API failure
        return extract_signals_mock(snippet)


if __name__ == "__main__":
    print("=========================================================================================")
    print("                DISCLOSURE STRUCTURED SIGNAL EXTRACTION RESULTS                         ")
    print("=========================================================================================")

    for snippet in DISCLOSURE_SNIPPETS:
        doc_id = snippet.split(":")[0].strip()
        result = extract_signals(snippet)

        print(f"[{doc_id}]")
        print(f"  Snippet          : {snippet}")
        print(f"  Risk Flags       : {result['risk_flags']}")
        print(f"  Hedging Detected : {result['hedging_detected']}")
        print(f"  Sentiment        : '{result['sentiment']}'")
        print("-" * 89)