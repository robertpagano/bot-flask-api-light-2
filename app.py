from flask import Flask, request, redirect, url_for, flash, jsonify, send_file, render_template
from werkzeug.exceptions import Forbidden, HTTPException, NotFound, RequestTimeout, Unauthorized
import numpy as np
import pandas as pd
import pickle as p
import json
from summarization.textsummarization import bert_sum
from docx import Document
from docx.shared import Inches
from linkcheck import flag_private_urls, flag_private_urls_to_dict

app = Flask(__name__)

## Testing out HTML
@app.route('/')
def form():
  return render_template('index.html')

## Handling Errors
@app.errorhandler(NotFound)
def page_not_found_handler(e: HTTPException):
    return render_template('404.html'), 404


@app.errorhandler(Unauthorized)
def unauthorized_handler(e: HTTPException):
    return render_template('401.html'), 401


@app.errorhandler(Forbidden)
def forbidden_handler(e: HTTPException):
    return render_template('403.html'), 403


@app.errorhandler(RequestTimeout)
def request_timeout_handler(e: HTTPException):
    return render_template('408.html'), 408
    
## summarizer using form data as an input, returns text summary
@app.route('/api/v1/resources/text/', methods=['POST'])
def summarize_from_text():
    
    data = request.form["data"]
    summary = bert_sum(data)

    return summary

## summarizer that pulls text from word doc, returns text summary
@app.route('/api/v1/resources/document/summary', methods=['GET', 'POST'])
def summarize_from_file():
    
    f = request.files['data']
    f.save('datafile.docx')
    document = Document('datafile.docx')
    text =''
    for para in document.paragraphs:
        text+=para.text
    
    summary = bert_sum(text)

    return summary

## this function just takes in a word file, and returns the word file
@app.route('/api/v1/resources/document/docx', methods=['GET', 'POST'])
def return_document():
    f = request.files['data']
    f.save('datafile.docx')

    return send_file('datafile.docx', attachment_filename='test.docx')

## this takes in a word file, summarizes text, adds summary to end, and returns document. Used downstream in transform route below
def transform():
    data = 'datafile.docx'
    document = Document(data)
    new = Document()
    text =''
    for para in document.paragraphs:
        text+=para.text
    new.add_paragraph(bert_sum(text))
    return new

## shows html interface, for now users can submit a document, and it will add a summary to the end
@app.route('/transform', methods=["GET","POST"])
def transform_view():
    f = request.files['data_file']
    f.save('datafile.docx')
    if not f:
        return "Please upload a word document"
    result = transform()
    result.save('result.docx')
    return send_file('result.docx', attachment_filename='new_file.docx')

## this takes in a word file, scrapes for links, checks links, and returns an excel file with the results of link checker
@app.route('/api/v1/resources/document/links/excel', methods=['POST'])
def check_links_to_excel():
    
    data = request.files["link"]
    data.save('linkfile.docx')
    results = flag_private_urls('linkfile.docx')
    results.to_excel('links.xlsx')
    
    return send_file('links.xlsx')

## this takes in a word file, scrapes for links, checks links, and returns a json object of the table of results of link checker
@app.route('/api/v1/resources/document/links/json', methods=['POST'])
def check_links_to_json():
    
    data = request.files["link"]
    data.save('linkfile.docx')
    results = flag_private_urls_to_dict('linkfile.docx')
    
    return jsonify(results)

if __name__ == '__main__':
    app.run(debug=True)
