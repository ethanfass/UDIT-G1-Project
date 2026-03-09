from google import genai
import os

api_key = os.getenv("GEMINI_API_KEY") or os.getenv("GOOGLE_API_KEY")
if not api_key:
    raise SystemExit("No API key set in GEMINI_API_KEY or GOOGLE_API_KEY")

client = genai.Client(api_key=api_key)

print("Available models:\n")
for m in client.models.list():
    print(m.name)
