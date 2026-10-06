# PPT Generator

A web app built with Python and Streamlit that creates a PowerPoint presentation on any topic. Enter a topic and a slide count, and download the finished `.pptx` file.

## Features
- Enter any presentation topic and choose 3 to 10 slides
- Free mode: pulls real content from Wikipedia (no API key needed)
- AI mode: paste an Anthropic API key in the sidebar and Claude writes the slides
- Builds the slides with python-pptx and offers a one-click download

## Tech Stack
Python, Streamlit, python-pptx, Requests, Anthropic API

## How to Run
```bash
pip install -r requirements.txt
streamlit run app.py
```
