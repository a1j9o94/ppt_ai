from flask import Flask, request, send_file, render_template, jsonify
import os
from dotenv import load_dotenv
from ppt_creation_agent import create_presentation_from_prompt

# load environment variables
load_dotenv()

app = Flask(__name__)

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        topic = request.form.get('topic')
        if topic:
            try:
                # Create the presentation
                result = create_presentation_from_prompt(topic)
                file_path = result.get('file_path')
                
                if file_path and os.path.exists(file_path):
                    return send_file(file_path, as_attachment=True)
                else:
                    return jsonify({"error": "Failed to create presentation"}), 500
            except Exception as e:
                return jsonify({"error": str(e)}), 500
        else:
            return jsonify({"error": "No topic provided"}), 400
    return render_template('index.html')

@app.route('/download/<filename>')
def download(filename):
    return send_file(os.path.join(os.getcwd(), "output", filename), as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
