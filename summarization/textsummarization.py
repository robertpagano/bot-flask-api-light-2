# %%
from summarizer import Summarizer
# from TextRankSummarizer import get_summary
import unicodedata
import pandas as pd
import numpy as np
import re
import math

#this is the model for Bert extractor summarization
model = Summarizer()

def bert_sum(text, ratio=0.2, model=model):
    """
    text : Text to be summarized
    ratio : Length of summary (value from 0-1.0) relative to original text
    """
    return model(text, ratio=ratio)

def bert_sum_dynamic(text, model=model):
    """
    text: text to be summarized

    this will read in a text and dynamically change the ratio of the summary based on 
    the original article's length
    """

    text_length = len(text.split())

    if text_length <= 350:
        ratio = 0.2

    elif 350 < text_length <= 700:
        ratio = 0.5

    else:
        ratio = 350 / text_length

    return model(text, ratio=ratio)
