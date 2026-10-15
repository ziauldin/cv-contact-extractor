# 📇 CV Contact Extractor

A small desktop tool I built at Future Connect Training & Recruitment to speed up data entry from candidate CVs. You pick a batch of PDF or Word files, and it pulls out each candidate's name, phone number, email, and UK postcode, then saves everything to one CSV.

![Python](https://img.shields.io/badge/Python-3776AB?style=flat-square&logo=python&logoColor=white)
![spaCy](https://img.shields.io/badge/spaCy-09A3D5?style=flat-square&logo=spacy&logoColor=white)
![pandas](https://img.shields.io/badge/pandas-150458?style=flat-square&logo=pandas&logoColor=white)
![Tkinter](https://img.shields.io/badge/Tkinter_GUI-3776AB?style=flat-square&logoColor=white)

<!-- Add a screenshot: ![App](screenshots/app.png) -->

## ✨ Features
- 🖱️ Tkinter window ("Smart Document Data Extractor") to select several `.pdf` and `.docx` files at once
- 📄 Text extraction with `pdfplumber` for PDFs and `python-docx` for Word files
- 🎯 Two-pass extraction: first looks for the contact section of the CV, then falls back to the full text
- 🧠 Names found with targeted patterns, then spaCy's `PERSON` entities (`en_core_web_lg`)
- ☎️ Regex-based phone, email, and UK postcode detection, with validation and formatting (UK numbers as `07xxx xxx xxx`)
- 💾 Results saved to `output.csv` with columns `filename, name, phone, email, postcode`

## 🚀 Run locally

```bash
git clone https://github.com/ziauldin/cv-contact-extractor.git
cd cv-contact-extractor
pip install -r requirements.txt
python -m spacy download en_core_web_lg
python bot.py
```

## 🛠️ Tech stack
Python · Tkinter · spaCy · pdfplumber · python-docx · pandas · regex
