from flask import Flask, jsonify, request, send_from_directory
import win32com.client
from flask_cors import CORS
import os
import pythoncom
import locale

app = Flask(__name__, static_folder='.')
CORS(app)  # Allow CORS for frontend communication

# Initialize COM for the application
pythoncom.CoInitialize()

# Serve the main page
@app.route("/")
def home():
    return send_from_directory(app.static_folder, 'page.html')

# Serve static files (CSS, JS, etc.)
@app.route('/<path:filename>')
def serve_static(filename):
    return send_from_directory(app.static_folder, filename)

# API to handle text-to-speech
@app.route("/api/speak", methods=["POST"])
def speak_text():
    try:
        pythoncom.CoInitialize()
        data = request.get_json()
        text = data.get("text", "")
        voice = data.get("voice", "default")

        if not text:
            return jsonify({"error": "No text provided"}), 400

        speak = win32com.client.Dispatch("SAPI.SpVoice")

        # Set selected voice
        if voice != "default":
            voices = speak.GetVoices()
            for v in voices:
                if v.GetDescription() == voice:
                    speak.Voice = v
                    break

        speak.Speak(text)

        pythoncom.CoUninitialize()
        return jsonify({"message": "Text spoken successfully!"})

    except Exception as e:
        try:
            pythoncom.CoUninitialize()
        except:
            pass
        return jsonify({"error": str(e)}), 500

# API to get available voices dynamically with accents/languages
@app.route("/api/voices", methods=["GET"])
def get_voices():
    try:
        pythoncom.CoInitialize()
        speak = win32com.client.Dispatch("SAPI.SpVoice")
        voices = speak.GetVoices()
        voice_list = []

        for voice in voices:
            voice_name = voice.GetDescription()

            # Extract language and gender attributes
            language_attr = voice.GetAttribute("Language")
            gender_attr = voice.GetAttribute("Gender")

            # Convert hex language code (like 409) to readable locale
            try:
                language_code = int(language_attr, 16)
                language_name = locale.windows_locale.get(language_code, "Unknown")
            except:
                language_name = language_attr

            voice_info = {
                "name": voice_name,
                "accent": language_name,
                "gender": gender_attr if gender_attr else "Unknown"
            }
            voice_list.append(voice_info)

        # Sort by accent then gender
        voice_list.sort(key=lambda x: (x["accent"], x["gender"]))

        pythoncom.CoUninitialize()
        return jsonify({"voices": voice_list})

    except Exception as e:
        try:
            pythoncom.CoUninitialize()
        except:
            pass
        return jsonify({"error": str(e)}), 500


if __name__ == "__main__":
    print("Starting Text Reader application...")
    print("Please open http://localhost:8080 in your web browser")
    app.run(host="127.0.0.1", port=8080, debug=True)
