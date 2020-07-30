#%%
import docxpy
import docx2txt
import unicodedata

# %%
text = docxpy.process('data/article one multiple topics.docx')

# %%
doc = docxpy.DOCReader('data/articletwo.docx')

# %%
text = doc.process()

# %%
import docx2txt


# %%
text = docx2txt.process('data/articletwo.docx')

# %%
from textsummarization import *
# %%
methods = ['bert_sum']
for i in methods:
    print(i)
    print(extract_sum(text, 0.5, i))
# %%
import torch
import json 
from transformers import T5Tokenizer, T5ForConditionalGeneration, T5Config
model = T5ForConditionalGeneration.from_pretrained('t5-small')
tokenizer = T5Tokenizer.from_pretrained('t5-small')
device = torch.device('cpu')
# text ="""
# Microsoft experienced another successful month of positive and valuable results across analyst and press engagements. The Microsoft Biz Apps Summit generated several reports, including Futurum Research, Lopez Research (in Forbes.com), and IDC. In Futurum’s report, “Microsoft Biz Apps Summit: Enabling Rapid Digital Transformation,” the analyst comments, “The event was well done. In fact, it was one of the best online events I have attended. . .” and, “Speed and agility have long been a topic of digital transformation, but the COVID-19 pandemic has arguably taken a decade of transformation and consolidated it into a fraction of that time. . . Microsoft, during its Business Applications Summit, was able to demonstrate the ability with its Dynamics 365, Power Platform and Teams solutions– showing how the company has fared so well over the past several years and throughout this difficult pandemic.”  Lopez Research’s “From Hand Sanitizers To Tacos: Microsoft Shows Us How Companies Are Using Technology To Thrive During A Pandemic,” highlighted several Microsoft client examples and noted, “What\'s different today is that vendors, such as Microsoft, are delivering more comprehensive portfolios. . .Microsoft made its solutions more powerful and easier to use.” IDC’s report, “Microsoft Business Applications Summit: Transforming and Empowering Digital Enterprises,” commented, “Microsoft Dynamics 365 appears to be making progress with a "land and expand" approach to enterprise applications and accelerating upmarket from SMB to large enterprise, such as Chipotle and Ikea.” In other positive coverage, Nucleus Research rates Dynamics 365 Enterprise ERP as a “Leader,” and places Business Central in the "Facilitator" category its ERP Technology Value Matrix 2020. Forrester released it’s The Forrester New WaveTM: SaaS Marketplaces, Q2 2020 report highlighting how SaaS continues to grow in popularity across most categories of software and maintains its appeal for business-led, under-the-radar buying as well as sanctioned and IT-led spend. Microsoft’s commercial marketplaces received a “strong performer” rating, improving over the prior 2018 report. Amazon and Salesforce are the only Leaders.Gartner released a number of reports recently. In their Market Guide for Online Fraud Detection report, Gartner highlights Dynamics 365 Fraud Protection as a market disruptor noting, “With an established market presence, mature distribution channels and broad cross-selling opportunities, the potential for market disruption by both Amazon and Microsoft is real over a period of several years.” In Gartner’s Market Share Analysis: Supply Chain Management Software, Worldwide, 2019 Gartner highlights the substantial growth in the supply chain management market at 8.6% in 2019.  Gartner also notes that Cloud revenue grew 2.5 times faster than the overall market, accounting for nearly 34% of the market as all leading vendors transitioned their new product strategy to cloud.Gartner also produced reports coving HCM technology (Human Capital Management) spending in 2020 and 2021 as well as a Gartner Peer Insights publication polling customers on the digital experience of various platforms. In the later, Microsoft received a “Customers’ Choice” designation along with Adobe and Salesforce.Constellation Research published It’s Time to Think Differently About Business Apps—And the Future of Customer-Facing Work highlighting the need to focus on what businesses need to achieve with their Business Applications, as opposed to focusing on the applications technical capabilities which may or may not directly impact impact customer experience and retention. Visit Analyst Resource Central for up-to-date press and analyst coverage reports.
# """
#%%
preprocess_text = text.strip().replace("\n","")
t5_prepared_Text = "summarize: "+preprocess_text
print ("original text preprocessed: \n", preprocess_text)
tokenized_text = tokenizer.encode(t5_prepared_Text, return_tensors="pt").to(device)
 
# summmarize 
summary_ids = model.generate(tokenized_text,
                                    num_beams=4,
                                    no_repeat_ngram_size=2,
                                    min_length=125,
                                    max_length=400,
                                    early_stopping=True)
output = tokenizer.decode(summary_ids[0], skip_special_tokens=True)
print ("\n\nSummarized text: \n",output)
# Summarized output from above ::::::::::
# the us has over 637,000 confirmed Covid-19 cases and over 30,826 deaths. 
# president Donald Trump predicts some states will reopen the country in april, he said. 
# "we'll be the comeback kids, all of us," the president says.


# %%
