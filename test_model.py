import tomllib, time
from google import genai
from google.genai import types

key = tomllib.load(open(".streamlit/secrets.toml", "rb"))["GEMINI_API_KEY"]
c = genai.Client(api_key=key, http_options=types.HttpOptions(timeout=20000))

for model in ["gemini-3.5-flash", "gemini-3.6-flash", "gemini-3.7-flash",
              "gemini-3.5-flash-lite", "gemini-3.1-flash-lite", "gemini-2.5-flash-lite"]:
    t = time.time()
    try:
        r = c.models.generate_content(model=model, contents="Say hi in one word.")
        print(model, "OK", round(time.time() - t, 1), "s:", r.text.strip())
    except Exception as e:
        print(model, "ERROR", round(time.time() - t, 1), "s:", str(e)[:150])