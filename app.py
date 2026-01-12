from flask import Flask, request, send_file, render_template, session, redirect, url_for, flash, after_this_request
import os
import pandas as pd
import re
from datetime import datetime
import tempfile
import matplotlib.pyplot as plt
from openpyxl import Workbook
from openpyxl.drawing.image import Image as ExcelImage
from collections import defaultdict
import hashlib
import openai
from openai import RateLimitError, APIError
import time
import random
from dotenv import load_dotenv
from docx import Document
import sys


app = Flask(__name__)

sys.stdout.reconfigure(encoding='utf-8')

# Load environment variables from .env file
load_dotenv()

app.secret_key = os.getenv('FLASK_SECRET_KEY', 'fallback_secret_key')
cache_dir = os.getenv('REPORT_CACHE_DIR', 'report_cache')
os.makedirs(cache_dir, exist_ok=True)
openai_last_request_ts = 0.0

# Set the OpenAI API key from the environment variable
openai.api_key = os.getenv("OPENAI_API_KEY")

if not openai.api_key:
    raise ValueError("OpenAI API key not set. Please configure it in the environment variable.")


def parse_whatsapp_chat(file):
    # Accept both desktop and mobile exports with flexible date/time formats.
    desktop_pattern = (
        r'^\[(\d{1,2})[./](\d{1,2})[./](\d{2,4}), '
        r'(\d{1,2}:\d{2})(?::(\d{2}))?\] (.*?)(?::\s)(.*)$'
    )
    mobile_pattern = (
        r'^(\d{1,2})/(\d{1,2})/(\d{2,4}), '
        r'(\d{1,2}:\d{2})(?:\s?(AM|PM))? - (.*)$'
    )
    chat_data = []
    current_message = None

    # Read and decode the file
    for line in file:
        line = line.decode('utf-8').strip()
        if not line:
            continue
        # Strip common directionality marks/BOM that may prefix exported lines.
        line = line.lstrip('\ufeff\u200e\u200f\u202a\u202b\u202c\u202d\u202e\u2066\u2067\u2068\u2069')

        date_time_obj = None
        sender = None
        message = None

        match = re.match(desktop_pattern, line)
        if match:
            day = match.group(1)
            month = match.group(2)
            year = match.group(3)
            time_str = match.group(4)
            seconds = match.group(5) or '00'
            sender = match.group(6)
            message = match.group(7)

            if len(year) == 2:
                year = f"20{year}"

            date_time_str = f'{day}.{month}.{year} {time_str}:{seconds}'
            try:
                date_time_obj = datetime.strptime(date_time_str, '%d.%m.%Y %H:%M:%S')
            except ValueError as ve:
                print(f"Date parsing error: {ve} for line: {line}")
                continue
        else:
            match = re.match(mobile_pattern, line)
            if match:
                month = match.group(1)
                day = match.group(2)
                year = match.group(3)
                time_str = match.group(4)
                am_pm = match.group(5)
                rest = match.group(6)

                if len(year) == 2:
                    year = f"20{year}"

                if am_pm:
                    time_str = f"{time_str} {am_pm}"
                    time_format = '%m/%d/%Y %I:%M %p'
                else:
                    time_format = '%m/%d/%Y %H:%M'

                date_time_str = f'{month}/{day}/{year} {time_str}'
                try:
                    date_time_obj = datetime.strptime(date_time_str, time_format)
                except ValueError as ve:
                    print(f"Date parsing error: {ve} for line: {line}")
                    continue

                sender, sep, message = rest.partition(': ')
                if not sep:
                    sender = 'System'
                    message = rest

        if date_time_obj is not None:
            chat_data.append({
                'Date': date_time_obj.date(),
                'Time': date_time_obj.time(),
                'Datetime': date_time_obj,
                'Sender': sender,
                'Message': message
            })
            current_message = len(chat_data) - 1
        elif current_message is not None:
            # Handle multi-line messages
            chat_data[current_message]['Message'] += '\n' + line

    return pd.DataFrame(chat_data)


@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        file = request.files.get('fileup')
        if file and file.filename.endswith('.txt'):
            # Parse the uploaded file
            parsed_chat = parse_whatsapp_chat(file)

            # Check if parsed_chat is empty
            if parsed_chat.empty:
                flash("Parsed chat is empty. Please check the chat format and try again.", 'error')
                return redirect(url_for('index'))

            # Calculate average reply times for each sender
            avg_reply_times = defaultdict(list)
            previous_sender = None
            previous_time = None

            for i, row in parsed_chat.iterrows():
                current_sender = row['Sender']
                current_time = row['Datetime']

                if previous_sender and previous_sender != current_sender:
                    time_difference = (current_time - previous_time).total_seconds() / 60  # Convert to minutes
                    avg_reply_times[previous_sender].append(time_difference)

                previous_sender = current_sender
                previous_time = current_time

            avg_reply_times = {sender: sum(times) / len(times) for sender, times in avg_reply_times.items() if times}

            # Create a temporary Excel file
            output_path = tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx').name
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                parsed_chat.to_excel(writer, sheet_name='Parsed Chat', index=False)
                message_summary = parsed_chat['Sender'].value_counts().reset_index()
                message_summary.columns = ['Sender', 'Total Messages']
                message_summary.to_excel(writer, sheet_name='Message Summary', index=False)
                avg_reply_df = pd.DataFrame(list(avg_reply_times.items()), columns=['Sender', 'Average Reply Time (mins)'])
                avg_reply_df.to_excel(writer, sheet_name='Avg Reply Time', index=False)

            # Store the file path for download
            session['output_file'] = output_path
            session['file_downloaded'] = False  # Track download status

            flash("File processed successfully.", "success")
            return redirect(url_for('index'))

    return render_template('index.html')

@app.route('/download_report', methods=['GET'])
def download_report():
    report_file = session.get('report_file')
    if not report_file or not os.path.exists(report_file):
        flash("No report available to download.", "error")
        return redirect(url_for('index'))

    # Delete the report file after sending it to the user
    @after_this_request
    def remove_report_file(response):
        try:
            os.remove(report_file)
            session.pop('report_file', None)  # Clear the file from the session
        except Exception as e:
            print(f"Error deleting file: {str(e).encode('utf-8', errors='ignore')}")
        return response

    return send_file(report_file, as_attachment=True, download_name="psychological_report.txt")

@app.route('/download_psychological_report', methods=['GET'])
def download_psychological_report():
    report_file = session.get('report_file')
    if not report_file or not os.path.exists(report_file):
        flash("No psychological report available to download.", "error")
        return redirect(url_for('index'))

    # Delete the report file after sending it to the user
    @after_this_request
    def remove_report_file(response):
        try:
            os.remove(report_file)
            session.pop('report_file', None)  # Clear the file from the session
        except Exception as e:
            print(f"Error deleting file: {str(e)}")
        return response

    return send_file(report_file, as_attachment=True, download_name="psychological_analysis_report.docx")

@app.route('/download', methods=['GET'])
def download():
    output_file = session.get('output_file')
    if not output_file or not os.path.exists(output_file):
        flash("No file available to download.", "error")
        return redirect(url_for('index'))

    if session.get('file_downloaded'):
        flash("The file has already been downloaded.", "error")
        return redirect(url_for('index'))

    session['file_downloaded'] = True

    @after_this_request
    def remove_file(response):
        try:
            os.remove(output_file)
            session.pop('output_file', None)  # Clear session after file is removed
            session.pop('file_downloaded', None)  # Reset download state
            flash("Download complete. The session has been reset.", "success")
        except Exception as e:
            print(f"Error deleting file: {str(e).encode('utf-8', errors='ignore')}")
        return response

    return send_file(output_file, as_attachment=True, download_name="parsed_chat_with_reply_times.xlsx")

# Route for processing CSV/Excel files and generating the psychological report
@app.route('/upload_csv', methods=['POST'])
def upload_csv():
    file = request.files.get('file_csv')
    if file and (file.filename.endswith('.csv') or file.filename.endswith('.xlsx')):
        # Load the uploaded CSV or Excel file into a DataFrame
        if file.filename.endswith('.csv'):
            df = pd.read_csv(file)
        else:
            df = pd.read_excel(file)

        # Check if the file is empty
        if df.empty:
            flash("Uploaded file is empty. Please check the file and try again.", 'error')
            return redirect(url_for('index'))

        # Generate the psychological analysis report
        psychological_analysis = generate_psychological_report(df)

        # Generate a Word document with the analysis
        report_path = generate_word_report(psychological_analysis, "psychological_report.docx")

        # Store the report path for download
        session['report_file'] = report_path

        flash("Psychological analysis report generated successfully.", "success")
        return redirect(url_for('index'))

    flash("Invalid file format. Please upload a CSV or Excel file.", 'error')
    return redirect(url_for('index'))

# Function to generate a psychological report
def call_openai_with_backoff(prompt, model_name, min_request_interval):
    global openai_last_request_ts
    retry_attempts = 5
    base_delay = 2
    max_delay = 60

    for attempt in range(retry_attempts):
        now = time.time()
        elapsed = now - openai_last_request_ts
        if elapsed < min_request_interval:
            wait_for = min_request_interval - elapsed
            print(f"Waiting {wait_for:.2f}s to respect RPM limit.")
            time.sleep(wait_for)
        try:
            response = openai.chat.completions.create(
                model=model_name,
                messages=[
                    {"role": "system", "content": "You are a helpful assistant."},
                    {"role": "user", "content": prompt}
                ],
                temperature=0
            )
            openai_last_request_ts = time.time()
            return response.choices[0].message.content
        except RateLimitError:
            delay = min_request_interval
            print(
                "Rate limit exceeded. Retrying in "
                f"{delay:.2f} seconds... (Attempt {attempt + 1}/{retry_attempts})"
            )
            time.sleep(delay)
        except APIError as e:
            return None

    return None


def generate_psychological_report(chat_data):
    # Convert the chat data to a readable format
    chat_text = ""
    for index, row in chat_data.iterrows():
        chat_text += f"[{row['Date']} {row['Time']}] {row['Sender']}: {row['Message']}\n"

    model_name = "gpt-4o-mini"
    rpm_limit = float(os.getenv("OPENAI_RPM_LIMIT", "3"))
    min_request_interval = 60.0 / max(rpm_limit, 1.0)
    chunk_max_chars = int(os.getenv("CHAT_CHUNK_MAX_CHARS", "8000"))

    cache_key_input = f"{model_name}:{chunk_max_chars}:{chat_text}"
    cache_key = hashlib.sha256(cache_key_input.encode('utf-8')).hexdigest()
    cache_path = os.path.join(cache_dir, f"{cache_key}.txt")
    if os.path.exists(cache_path):
        with open(cache_path, "r", encoding="utf-8") as cached:
            return cached.read()

    lines = chat_text.splitlines()
    chunks = []
    current_lines = []
    current_len = 0
    for line in lines:
        line_len = len(line) + 1
        if current_lines and current_len + line_len > chunk_max_chars:
            chunks.append("\n".join(current_lines))
            current_lines = [line]
            current_len = line_len
        else:
            current_lines.append(line)
            current_len += line_len
    if current_lines:
        chunks.append("\n".join(current_lines))

    total_parts = len(chunks)
    print(f"Prepared {total_parts} chat chunk(s) for report generation.")
    if total_parts > 1 and rpm_limit < 10:
        min_request_interval = 60.0
        print("Chunked mode enabled; waiting 60s before first OpenAI request.")
        time.sleep(min_request_interval)

    if total_parts == 1:
        prompt = f"""
        You will receive a WhatsApp chat snippet (part 1/1). Treat this as the full content.
        Act as an Industrial/Organizational Psychology expert with a specialization in Human
        Resource Management and Management. Analyze the chat data and generate a psychological
        analysis report covering emotions, relationships, psychological conditions, and
        communication patterns.

        WhatsApp Chat Data:
        {chunks[0]}

        Provide a comprehensive analysis of the individuals' psychological states, emotions,
        and relationships.
        """
        content = call_openai_with_backoff(prompt, model_name, min_request_interval)
        if content is None:
            return "Error: Rate limit exceeded after multiple attempts. Please try again later."
        with open(cache_path, "w", encoding="utf-8") as cached:
            cached.write(content)
        return content

    partial_summaries = []
    for idx, chunk in enumerate(chunks, start=1):
        print(f"Submitting chunk {idx}/{total_parts} to OpenAI.")
        prompt = f"""
        You will receive a WhatsApp chat snippet (part {idx}/{total_parts}).
        Act as an Industrial/Organizational Psychology expert with a specialization in Human
        Resource Management and Management. Summarize key psychological signals, emotions,
        relationships, and communication patterns found in this snippet. Return a concise
        bullet list of findings and notable quotes with speaker names.

        WhatsApp Chat Data:
        {chunk}
        """
        summary = call_openai_with_backoff(prompt, model_name, min_request_interval)
        if summary is None:
            return "Error: Rate limit exceeded after multiple attempts. Please try again later."
        partial_summaries.append(summary)

    print("Submitting final combined summary request to OpenAI.")
    combined_prompt = f"""
    You received {total_parts} partial summaries from a WhatsApp chat. Combine them into a
    single comprehensive psychological analysis report. Focus on emotions, relationships,
    psychological conditions, and communication patterns across the full conversation.

    Partial Summaries:
    {"\n\n".join(partial_summaries)}
    """

    content = call_openai_with_backoff(combined_prompt, model_name, min_request_interval)
    if content is None:
        return "Error: Rate limit exceeded after multiple attempts. Please try again later."
    with open(cache_path, "w", encoding="utf-8") as cached:
        cached.write(content)
    return content


def generate_word_report(analysis_text, file_name):
    # Function to create a Word document with the analysis
    from docx import Document
    doc = Document()
    doc.add_heading('Psychological Analysis Report', 0)

    doc.add_paragraph(analysis_text)

    # Save the document as a temporary file
    word_file = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    doc.save(word_file.name)

    return word_file.name

if __name__ == '__main__':
    host = os.getenv('FLASK_RUN_HOST', '127.0.0.1')
    port = int(os.getenv('FLASK_RUN_PORT', '5000'))
    app.run(debug=True, host=host, port=port)
