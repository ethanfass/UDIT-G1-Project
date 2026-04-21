import argparse
import getpass
import json
import urllib.error
import urllib.parse
import urllib.request

from gemini_secret_store import (
    SECRET_PATH,
    clear_local_api_key,
    local_api_key_exists,
    mask_api_key,
    resolve_api_key_with_source,
    save_local_api_key,
)


def test_api_key(api_key: str) -> tuple[bool, str]:
    url = (
        "https://generativelanguage.googleapis.com/v1beta/models"
        f"?key={urllib.parse.quote(api_key)}"
    )
    request = urllib.request.Request(url, method="GET")
    try:
        with urllib.request.urlopen(request, timeout=30) as response:
            payload = json.loads(response.read().decode("utf-8"))
        count = len(payload.get("models", []))
        return True, f"API key is valid. Visible models: {count}"
    except urllib.error.HTTPError as exc:
        detail = exc.read().decode("utf-8", errors="ignore")
        return False, f"HTTP {exc.code}: {detail}"
    except urllib.error.URLError as exc:
        return False, f"Request failed: {exc}"


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Manage the locally encrypted Gemini API key for this project."
    )
    parser.add_argument("--set", action="store_true", help="Store or replace the local API key.")
    parser.add_argument("--clear", action="store_true", help="Delete the locally stored API key.")
    parser.add_argument("--status", action="store_true", help="Show whether a local API key is configured.")
    parser.add_argument("--test", action="store_true", help="Validate the currently resolved API key against Gemini.")
    args = parser.parse_args()

    if not (args.set or args.clear or args.status or args.test):
        parser.print_help()
        return

    if args.status:
        resolved_key, source = resolve_api_key_with_source()
        print(f"Resolved key source: {source}")
        print(f"Resolved key preview: {mask_api_key(resolved_key)}")
        if local_api_key_exists():
            print(f"Local encrypted Gemini API key found at: {SECRET_PATH}")
        else:
            print("No local encrypted Gemini API key is configured.")

    if args.clear:
        if clear_local_api_key():
            print("Deleted the local encrypted Gemini API key.")
        else:
            print("No local encrypted Gemini API key was found.")

    if args.set:
        api_key = getpass.getpass("Enter Gemini API key: ").strip()
        if not api_key:
            raise SystemExit("No API key entered.")
        path = save_local_api_key(api_key)
        print(f"Stored encrypted Gemini API key at: {path}")
        print("It is encrypted with your current Windows user account via DPAPI.")

    if args.test:
        api_key, source = resolve_api_key_with_source()
        if not api_key:
            raise SystemExit("No API key found to test.")
        ok, message = test_api_key(api_key)
        print(f"Testing key from: {source}")
        print(f"Key preview: {mask_api_key(api_key)}")
        print(message)
        if not ok:
            raise SystemExit(1)


if __name__ == "__main__":
    main()
