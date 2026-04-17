import json
import os
import re
import sys
import urllib.request
import urllib.error


def _read_file(path: str) -> str:
    with open(path, "r", encoding="utf-8") as f:
        return f.read()


def _post_json(url: str, headers: dict, payload: dict) -> dict:
    data = json.dumps(payload).encode("utf-8")
    req = urllib.request.Request(url, data=data, headers=headers, method="POST")
    with urllib.request.urlopen(req, timeout=120) as resp:
        raw = resp.read().decode("utf-8")
        try:
            return json.loads(raw)
        except json.JSONDecodeError:
            return {"_raw_text": raw}


def _clean_text(text: str) -> str:
    return text.strip()


def main():
    input_path = sys.argv[1] if len(sys.argv) > 1 else "PCI_44_54.txt"
    output_txt = sys.argv[2] if len(sys.argv) > 2 else "output/PCI_44_54_structured.txt"

    base_url = os.getenv("ANTHROPIC_BASE_URL")
    api_key = os.getenv("ANTHROPIC_AUTH_TOKEN")
    model = os.getenv("ANTHROPIC_DEFAULT_SONNET_MODEL", "claude-3-5-sonnet-latest")

    if not base_url or not api_key:
        raise SystemExit("Missing ANTHROPIC_BASE_URL or ANTHROPIC_AUTH_TOKEN in environment.")

    text = _read_file(input_path)

    system = (
        "You are a strict information extraction engine. "
        "Return ONLY the requested TXT format with no extra text. "
        "If a field is missing, use 'N/A'."
    )

    user = (
        "Convert the following PCI text into TXT blocks, one block per requirement (MaYeuCau).\n\n"
        "Required TXT format for each block:\n"
        "Mã Yêu cầu: \"<code>\"\n"
        "Defined Approach Requirements: <text>\n"
        "Customized Approach Objective: <text>\n"
        "Applicability Notes: <text or N/A>\n"
        "Defined Approach Testing Procedures:\n"
        "  \"<test a>\": <text>\n"
        "  \"<test b>\": <text>\n"
        "Guidance - Purpose: <text or N/A>\n"
        "Guidance - Good Practice: <text or N/A>\n"
        "Guidance - Further Information: <text or N/A>\n"
        "Guidance - Definitions: <text or N/A>\n"
        "Guidance - Examples: <text or N/A>\n"
        "--\n\n"
        "Rules:\n"
        "- Keep content verbatim from the text, but you may normalize whitespace.\n"
        "- If a requirement has multiple testing procedures with letters (a, b, c), output each on its own line.\n"
        "- If any guidance section is missing, use N/A.\n"
        "- Return ONLY TXT in the exact format above.\n\n"
        "Source text:\n"
        f"""{text}"""
    )

    url = base_url.rstrip("/") + "/v1/messages"
    headers = {
        "Content-Type": "application/json",
        "x-api-key": api_key,
        "anthropic-version": "2023-06-01",
    }

    payload = {
        "model": model,
        "max_tokens": 4000,
        "temperature": 0,
        "system": system,
        "messages": [
            {"role": "user", "content": user}
        ],
    }

    try:
        resp = _post_json(url, headers, payload)
    except urllib.error.HTTPError as e:
        body = e.read().decode("utf-8", errors="ignore")
        raise SystemExit(f"HTTPError {e.code}: {body}")

    # Anthropic response format (fallbacks handled)
    if "_raw_text" in resp:
        text_out = resp["_raw_text"].strip()
    else:
        content = resp.get("content", [])
        if content:
            text_out = "".join([c.get("text", "") for c in content if c.get("type") == "text"]).strip()
        else:
            # Try OpenAI-compatible format if gateway differs
            choices = resp.get("choices", [])
            if choices and "message" in choices[0]:
                text_out = choices[0]["message"].get("content", "").strip()
            else:
                raise SystemExit(f"Unexpected response: {resp}")

    text_out = _clean_text(text_out)
    if not text_out:
        debug_path = "output/PCI_44_54_llm_raw_response.txt"
        os.makedirs(os.path.dirname(debug_path), exist_ok=True)
        with open(debug_path, "w", encoding="utf-8") as f:
            f.write(text_out)
        raise SystemExit(
            "LLM response is empty. "
            f"Raw response saved to {debug_path}."
        )

    os.makedirs(os.path.dirname(output_txt), exist_ok=True)
    with open(output_txt, "w", encoding="utf-8") as f:
        f.write(text_out)

    print(f"Wrote TXT output to {output_txt}")


if __name__ == "__main__":
    main()
