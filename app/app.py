from flask import Flask, render_template, request, redirect, url_for, send_file, session
from PyPDF2 import PdfReader
import spacy
import pandas as pd
import matplotlib.pyplot as plt
import secrets
import os


app = Flask(__name__)

secret_key = secrets.token_hex(16)  # Generate a 32-character hexadecimal secret key
app.secret_key = secret_key

nlp = spacy.load("./spacy")

from spacy.lang.en.stop_words import STOP_WORDS

def preprocessing(sentence):
    stopwords    = list(STOP_WORDS)
    doc          = nlp(sentence)
    clean_tokens = []
    
    for token in doc:
        if token.text not in stopwords and token.pos_ != 'PUNCT' and token.pos_ != 'SYM' and \
            token.pos_ != 'SPACE':
                clean_tokens.append(token.lemma_.lower().strip())
                
    return " ".join(clean_tokens)

def get_info(text):
    
    doc = nlp(text)
    
    skills = []
    experiences = []
    certificates = []
    contacts = []
    
    for ent in doc.ents:
        if ent.label_ == 'SKILL':
            skills.append(ent.text)
        elif ent.label_ == 'EXPERIENCE':
            experiences.append(ent.text)
        elif ent.label_ == 'CERTIFICATE':
            certificates.append(ent.text)    
        elif ent.label_ == 'CONTACT':
            contacts.append(ent.text)    
    
    skills = list(set(skills))
    experiences = list(set(experiences))
    certificates = list(set(certificates))
    contacts = list(set(contacts))

    skills = ", ".join(skills)
    experiences = ", ".join(experiences)
    certificates = ", ".join(certificates)
    contacts = ", ".join(contacts)

    return [skills], [experiences], [certificates], [contacts]


@app.route("/", methods=["GET", "POST"])
def upload_pdf():
    if request.method == "POST":
        if request.files:
            pdf = request.files["pdf"]
            if pdf.filename.endswith(".pdf"):
                # Save the uploaded file in the same directory
                upload_dir = os.path.join(app.root_path, "uploads")
                if not os.path.exists(upload_dir):
                    os.makedirs(upload_dir)
                pdf_path = os.path.join(upload_dir, pdf.filename)
                pdf.save(pdf_path)
                reader = PdfReader(pdf)
                text = ""
                for page in reader.pages:
                    text += page.extract_text()
                text = preprocessing(text)
                skills, experiences, certificates, contacts = get_info(text)
                session['extracted_info'] = (skills, experiences, certificates, contacts)
                return render_template("results.html", skills=skills, experiences=experiences, certificates=certificates, contacts=contacts)
    return render_template("upload.html")

@app.route("/download_excel")
def download_excel():
    extracted_info = session.get('extracted_info')
    if extracted_info:
        # Save the Excel file in the same directory as the uploaded file
        upload_dir = os.path.join(app.root_path, "uploads")
        excel_path = os.path.join(upload_dir, "output.xlsx")

        skills, experiences, certificates,contacts = extracted_info
        df = pd.DataFrame({'Skill': skills, 'Experience': experiences, 'Certificate': certificates, 'Contact':contacts})
        df.to_excel(excel_path, index=False)
        return send_file(excel_path, as_attachment=True)
    else:
        return "No data found for download."

@app.route("/download_image")
def download_image():
    extracted_info = session.get('extracted_info')
    if extracted_info:
        # Save the image in the same directory as the uploaded file
        upload_dir = os.path.join(app.root_path, "uploads")
        image_path = os.path.join(upload_dir, "output.png")
        skills, experiences, certificates,contacts = extracted_info
        df = pd.DataFrame({'Skill': skills, 'Experience': experiences, 'Certificate': certificates, 'Contact':contacts})
        df['Skill'] = df['Skill'].str.replace(', ', '\n')
        df['Experience'] = df['Experience'].str.replace(', ', '\n')
        df['Certificate'] = df['Certificate'].str.replace(', ', '\n')
        df['Contact'] = df['Contact'].str.replace(', ', '\n')

        # Calculate the required width for each column based on the maximum text length
        column_widths = [max(df[col].apply(lambda x: len(str(x)))) for col in df.columns]

        # Plot DataFrame with larger font size and multiline text
        plt.figure(figsize=(sum(column_widths) * 0.1, (df.shape[0]+1) * 6))
        # Plot table
        table = plt.table(cellText=df.values, colLabels=df.columns, loc='center', cellLoc='center')

        # Adjust cell sizes to fit content
        for (i, j), cell in table.get_celld().items():
            if (i == 0):  
                cell.set_text_props(fontsize=12, fontweight='bold') 
                cell.set_height(0.05)  
            else:  
                cell.set_height(0.5)  

        # Adjust font size
        table.auto_set_font_size(False)
        table.set_fontsize(12)  # Set font size here, adjust as needed

        plt.axis('off')  # Turn off axis
        plt.savefig(image_path)  # Save as PNG image 
        plt.close()  # Close the plot to prevent displaying it
        return send_file(image_path, as_attachment=True)
    else:
        return "No data found for download."

if __name__ == "__main__":
    app.run(debug=True)