import os
import time
import traceback

from flask import Flask, request, render_template, send_from_directory, redirect

from model import getImgsFromPDF, getTextFromImgs, getPPTFromImgText, send_email_with_ppt, send_feedback_email

app = Flask(__name__)
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'
TEXT_FOLDER = 'text'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs(TEXT_FOLDER, exist_ok=True)


@app.route('/')
def home():
    return render_template('index.html')


@app.route('/send_email', methods=['POST'])
def send_email():
    recipient_email = request.form['recipient_email']

    if not recipient_email:
        return "Invalid email address."

    try:
        if 'file_upload' not in request.files:
            return 'No file part'

        file = request.files['file_upload']

        uploaded_file = request.files['file_upload']

        if uploaded_file.filename == '':
            return 'No selected file'

        pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        file.save(pdf_path)

        for folder in [OUTPUT_FOLDER, TEXT_FOLDER]:
            for filename in os.listdir(folder):
                file_path = os.path.join(folder, filename)
                os.remove(file_path)

        getImgsFromPDF(pdf_path, OUTPUT_FOLDER)

        getTextFromImgs(OUTPUT_FOLDER, TEXT_FOLDER)

        ppt_path = os.path.join(TEXT_FOLDER, 'extracted_text_presentation.pptx')
        getPPTFromImgText(OUTPUT_FOLDER, ppt_path)

        send_email_with_ppt(recipient_email, ppt_path)

        render_template("message.html")
        time.sleep(3)

        return redirect('/')
    except Exception as e:
        print(e)
        return redirect('/')


@app.route('/upload', methods=['POST'])
def upload_file():
    try:
        if request.method == 'POST':
            if 'file' not in request.files:
                return 'No file part'
            file = request.files['file']
            if file.filename == '':
                return 'No selected file'
            if file:
                pdf_path = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
                file.save(pdf_path)

                for folder in [OUTPUT_FOLDER, TEXT_FOLDER]:
                    for filename in os.listdir(folder):
                        file_path = os.path.join(folder, filename)
                        os.remove(file_path)

                getImgsFromPDF(pdf_path, OUTPUT_FOLDER)

                getTextFromImgs(OUTPUT_FOLDER, TEXT_FOLDER)

                ppt_path = os.path.join(TEXT_FOLDER, 'TextPPT.pptx')
                getPPTFromImgText(OUTPUT_FOLDER, ppt_path)

                return send_from_directory(TEXT_FOLDER, 'TextPPT.pptx', as_attachment=True)
            else:
                return 'Text extraction failed or no text found.'
    except Exception as e:
        traceback.print_exc()
        return str(e), 400


@app.route('/send_feedback', methods=['POST'])
def send_feedback():
    try:
        feedback_text = request.form['feedback_text']

        send_feedback_email("codesteinsprojectmail@gmail.com", feedback_text)

        return redirect('/')
    except Exception as e:
        return f"Feedback could not be sent. Error: {str(e)}"


if __name__ == '__main__':
    app.run(debug=True)
