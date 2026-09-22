from dotenv import load_dotenv
from google import genai

load_dotenv()
client = genai.Client()

print("Consultando modelos habilitados en tu cuenta...\n")
for m in client.models.list():
    if "generateContent" in m.supported_actions:
        print(f"Modelo disponible: {m.name}")