---
title: Workshop Evaluation Report Generator
emoji: 📊
colorFrom: blue
colorTo: cyan
sdk: gradio
app_file: app.py
pinned: false
---

# Workshop Evaluation Report Generator

A web interface that uses preinstalled evaluation and attendance PowerPoint templates.

Upload a survey results PDF, survey responses Excel workbook and either an attendance Excel sheet or an attendance photo. Review attendance before generating the editable PowerPoint report.

**Privacy:** Workshop participant names, comments and scanned signatures may be confidential. Create the Space as **Private** unless you have permission to make the content public. Never publish uploaded reports or user data in the repository.

## Deploy in Hugging Face Spaces

1. Sign in to https://huggingface.co and open **Spaces → Create new Space**.
2. Choose **Gradio** and **Private** visibility. Choose CPU hosting appropriate for your data volume.
3. Upload all these files (including the entire `templates/` directory) through **Files → Add file → Upload files**, or push using Git.
4. Wait for the Space status to show **Running**. Open the URL displayed on the Space page.
5. Upload your three workshop data files into the web interface. The PPTX templates are preloaded.

The URL will typically be `https://huggingface.co/spaces/YOUR_USERNAME/YOUR_SPACE_NAME` with an app page at `https://YOUR_USERNAME-YOUR_SPACE_NAME.hf.space`. These are patterns, **not an already published link**.

## Run on another server

Install Python 3.11+, `pip install -r requirements.txt`, install the OS package `tesseract-ocr`, then run `python app.py`. The app listens on `0.0.0.0:${PORT:-7860}`.

## Notes

- Image recognition is a draft and requires human confirmation.
- The app copies the original PPTX styles from `templates/`.
- On services that sleep when idle, the first visit after inactivity can be slower.
- Access to a private Space requires login/permissions; for custom client-facing authentication or strong data-handling controls, deploy in a dedicated secured environment.
