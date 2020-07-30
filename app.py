from flask import Flask, request, redirect, url_for, flash, jsonify, send_file, render_template, request
from werkzeug.exceptions import Forbidden, HTTPException, NotFound, RequestTimeout, Unauthorized
import numpy as np
import pandas as pd
import pickle as p
import json

from summarization.textsummarization import bert_sum, bert_sum_dynamic

from docx import Document
from docx.shared import Inches

from docx_tools import word_doc_gen

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

## summarizer that pulls text from word doc, returns text summary with a dynamic length
@app.route('/api/v1/resources/document/summary/dynamic', methods=['GET', 'POST'])
def summarize_from_file_dynamic_length():
    
    f = request.files['data']
    f.save('datafile.docx')
    document = Document('datafile.docx')
    text =''
    for para in document.paragraphs:
        text+=para.text
    
    summary = bert_sum_dynamic(text)

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

## this takes in multiple word files, and then will create master summary and article files
## i think we will need to tweak the function and also will need to take in some variables like:
#### month
#### section
#### title

@app.route('/api/v1/resources/document/docbuilder', methods=["POST"])
def build_docs():
    uploaded_files = request.files.getlist("file[]")
    paths = request.form.getlist("paths[]") #work

    # print(uploaded_files)
    

    # docx_list = []
    docx_dict = {}

    # {filename: {
    ## month:
    ## section:
    ## document object:
    ## article title:
    # }}

    for document in uploaded_files:
        #make each into a docx object
        filename = document.filename # + '.docx'# great this works
        print(filename)
        document.save(filename)
        doc = Document(filename)

        docx_dict.update({filename: doc})
        ## docx_list.append(doc)
        ## make a dict like "filename: doc object" (cant return dict not json serializable? maybe jsonify it)
        ## i think that's fine, will just use this dict for the function moving forward
        ## <class 'werkzeug.datastructures.FileStorage'>

        ## THIS NEEDS TO TAKE IN TWO ARRAYS: ONE FOR FILES AND ONE FOR PATHS

    print(docx_dict)    

    paths_list = []
    
    for path in paths:
        paths_list.append(path)
    
    print(paths_list)

    return ''



@app.route("/upload", methods=["POST"])
def upload():
    uploaded_files = request.files.to_dict()
    for file in uploaded_files:
        print(uploaded_files[file].filename)

    return ""

# images = request.files.to_dict() #convert multidict to dict
# for image in images:     #image will be the key 
#     print(images[image])        #this line will print value for the image key
#     file_name = images[image].filename
#     images[image].save(some_destination)

if __name__ == '__main__':
    app.run(debug=True)
