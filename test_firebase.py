"""
Перевірка підключення до Firebase крок за кроком.

Запуск (з активним venv, у папці english-assistant):
    python test_firebase.py

Скрипт нічого не змінює у ваших чатах: він пише й одразу видаляє один тестовий документ.
Ключі в вивід не друкуються.
"""
import sys
import tomllib


def step(n, text):
    print(f"{n}) {text}")


def fail(e, hint=""):
    print(f"   ❌ {type(e).__name__}: {str(e)[:300]}")
    if hint:
        print(f"   👉 {hint}")
    sys.exit(1)


# 1. Читаємо секрети
step(1, "Reading .streamlit/secrets.toml")
try:
    with open(".streamlit/secrets.toml", "rb") as f:
        secrets = tomllib.load(f)
except Exception as e:
    fail(e, "Run this script from the english-assistant folder, and check that secrets.toml is valid TOML.")

info = secrets.get("firebase")
if not info:
    print("   ❌ No [firebase] section found in secrets.toml")
    sys.exit(1)
print("   ✅ [firebase] section found")

# 2. Імпорт бібліотек
step(2, "Importing Google libraries")
try:
    from google.oauth2 import service_account
    from google.cloud import firestore
except Exception as e:
    fail(e, "Try: pip install --upgrade google-cloud-firestore google-auth")
print("   ✅ Libraries imported")

# 3. Розбір ключа
step(3, "Parsing the service account key")
try:
    creds = service_account.Credentials.from_service_account_info(dict(info))
except Exception as e:
    fail(e, "The key in secrets.toml looks damaged. Re-run convert_key.py and paste the result again.")
print(f"   ✅ Key parsed, project: {creds.project_id}")

# 4. Клієнт
step(4, "Creating Firestore client")
try:
    db = firestore.Client(credentials=creds, project=creds.project_id)
except Exception as e:
    fail(e)
print("   ✅ Client created")

# 5. Запис, читання, видалення
step(5, "Writing a test document")
try:
    ref = db.collection("_selftest").document("ping")
    ref.set({"ok": True})
except Exception as e:
    hint = ""
    low = str(e).lower()
    if "not found" in low or "does not exist" in low:
        hint = "The database is not created yet: Firebase console > Firestore Database > Create database."
    elif "permission" in low or "403" in low or "has not been used" in low:
        hint = "Open the link from the error message and press Enable (Cloud Firestore API), wait a minute, try again."
    fail(e, hint)
print("   ✅ Write OK")

step(6, "Reading and deleting it")
try:
    print("   ✅ Read OK:", ref.get().to_dict())
    ref.delete()
except Exception as e:
    fail(e)
print("   ✅ Cleanup OK")

print("\nEverything works. Firebase is connected correctly ✅")