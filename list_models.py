import json
import urllib.error
import urllib.parse
import urllib.request

from gemini_secret_store import resolve_api_key


def main() -> None:
    api_key = resolve_api_key()
    if not api_key:
        raise SystemExit(
            "No Gemini API key found. Set GEMINI_API_KEY/GOOGLE_API_KEY or run "
            "'py -3.12 manage_gemini_key.py --set' once."
        )

    url = (
        "https://generativelanguage.googleapis.com/v1beta/models"
        f"?key={urllib.parse.quote(api_key)}"
    )
    request = urllib.request.Request(url, method="GET")

    try:
        with urllib.request.urlopen(request, timeout=60) as response:
            payload = json.loads(response.read().decode("utf-8"))
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="ignore")
        raise SystemExit(f"HTTP {exc.code}: {detail}") from exc
    except urllib.error.URLError as exc:
        raise SystemExit(f"Request failed: {exc}") from exc

    print("Available models:\n")
    for model in payload.get("models", []):
        print(model.get("name", "<unknown>"))


if __name__ == "__main__":
    main()
