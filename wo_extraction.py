#!/usr/bin/env python
# coding: utf-8

# In[1]:


from IPython.lib.display import join

import pandas as pd
import pdfplumber
import re
import os

from dateutil import parser


# In[12]:


# Define all necessary variables

email_pattern = r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b'
end_characters = ['.', ':', ';']
start_characters = ['-', ' -']
address_to_remove_haart = ['Dulwich', 'London,', 'SE22 9DQ']
address_to_remove_seq = ['Cumbria House', 'c/o London LSC', '16-20 Hockliffe Street', 'Leighton Buzzard', 'Bedfordshire', 'LU7 1GN', 'England']

folder_path = "WorkOrders"

append_mode = False


# In[15]:


# Specific routine for cleaning address in Sequence work orders

def clean_sequence_address(seq_add):
    seq_add = seq_add.split('Invoice To:')
    seq1 = seq_add[0].strip()
    seq2 = seq_add[1].strip()

    for word in address_to_remove_seq:
        
        seq2 = seq2.replace(word,'')

    seq2 = seq2.split('\n')
    seq2.pop(0)

    for line in seq2:
        
        if line.strip() != '':
            seq1 = seq1 + ', ' + line.strip()

    return seq1


# In[16]:


# Check whether extraction file should be completely refreshed or appended with latest records

if append_mode:

    # Load existing records into dataframes
    df = pd.read_csv(folder_path + '/Extraction/extracted_data.csv', encoding='ISO-8859-1')
else:

    # Initialize an empty dataframe
    columns = ["Date", "Date Raw", "Ref", "Contact", "Agency", "PM", "Email", "Address", "Postcode", "Work Description", "Work Cleaned", "Address_match"]
    df = pd.DataFrame(columns=columns)


# In[20]:


def extract_info(text):

    agency_wo = ''
    date_raw = ''

    if 'Kinleigh' in text:
        agency_wo = 'KFH'

        # Use regex to extract the relevant information
        date_pattern = re.compile(r" Date: (.*?)\n", re.DOTALL)
        contact_pattern = re.compile(r"Landlord:(.*?)\nProperty:", re.DOTALL)
        address_pattern = re.compile(r"Property:(.*?)\n", re.DOTALL)
        pm_pattern = re.compile(r"Requested By:(.*?)\n", re.DOTALL)

        if 'Quote Request' in text:
            ref_pattern = re.compile(r"Quote No:(.*?)\n", re.DOTALL)
            job_pattern = re.compile(r"Proposed Works to be priced:\n(.*?)Quote Required by", re.DOTALL)
        else:
            ref_pattern = re.compile(r"Order No: (.*?)\n", re.DOTALL)
            job_pattern = re.compile(r"Works Description:\n(.*?)Invoice To:", re.DOTALL)

        date_wo = re.search(date_pattern, text).group(1).strip().replace('\n', '')
        contact_wo = re.search(contact_pattern, text).group(1).split('Tel:')[0].strip().replace('\n', ' ')
        ref_wo = re.search(ref_pattern, text).group(1).strip().replace('\n', '')
        address_wo = re.search(address_pattern, text).group(1).strip()
        pm_wo = re.search(pm_pattern, text).group(1).strip()
        job_wo = re.search(job_pattern, text).group(1).strip()

        # Format the date
        date_raw = date_wo
        date_obj = parser.parse(date_wo)
        date_wo = date_obj.strftime("%d/%m/%Y")

    elif 'ackroyd' in text:
        agency_wo = 'Stirling Ackroyd'

        # Use regex to extract the relevant information
        date_pattern = re.compile(r" Date: (.*?)\n", re.DOTALL)
        contact_pattern = re.compile(r"Landlord:(.*?)\n", re.DOTALL)
        ref_pattern = re.compile(r"Reference:(.*?)\n", re.DOTALL)
        address_pattern = re.compile(r"Property:(.*?)\n", re.DOTALL)
        pm_pattern = re.compile(r"Property Manager\n(.*?)\n", re.DOTALL)
        job_pattern = re.compile(r"Work Order Details Estimated Cost\n(.*?)Please advise if cost", re.DOTALL)

        date_wo = re.search(date_pattern, text).group(1).strip().replace('\n', '')
        contact_wo = re.search(contact_pattern, text).group(1).strip().replace('\n', '')
        ref_wo = re.search(ref_pattern, text).group(1).strip().replace('\n', '')
        address_wo = re.search(address_pattern, text).group(1).strip()
        pm_wo = re.search(pm_pattern, text).group(1).strip()
        job_wo = re.search(job_pattern, text).group(1).strip()

        pm_wo = pm_wo.split()[-2:]
        pm_wo = ' '.join(pm_wo)

    elif 'Sequence' in text:
        agency_wo = 'Sequence'

        # Use regex to extract the relevant information
        date_pattern = re.compile(r"DATE: (.*?)\n", re.DOTALL)
        contact_pattern = re.compile(r"Invoice To:(.*?)\n", re.DOTALL)
        ref_pattern = re.compile(r"JOB REFERENCE:(.*?)\n", re.DOTALL)
        address_pattern = re.compile(r"Address: (.*?)Our Ref: ", re.DOTALL)
        pm_pattern = re.compile(r"Signed:(.*?)Property Co-ordinator", re.DOTALL)
        pm_pattern2 = re.compile(r"Signed:(.*?)Team Leader", re.DOTALL)
        job_pattern = re.compile(r"Job Description:(.*?)Pick Up Keys From:", re.DOTALL)

        date_wo = re.search(date_pattern, text).group(1).strip().replace('\n', '')
        contact_wo = re.search(contact_pattern, text).group(1).strip().replace('\n', '')
        ref_wo = re.search(ref_pattern, text).group(1).strip().replace('\n', '')
        address_wo = re.search(address_pattern, text).group(1).strip()

        if re.search(pm_pattern, text):
            pm_wo = re.search(pm_pattern, text).group(1).strip()

        if re.search(pm_pattern2, text):
            pm_wo = re.search(pm_pattern2, text).group(1).strip()

        job_wo = re.search(job_pattern, text).group(1).split('The tenant has advised that')[0].strip()

        address_wo = clean_sequence_address(address_wo)

        emails = [pm_wo.replace(' ', '.') + '@sequencehome.co.uk']

    elif 'haart' in text:
        agency_wo = 'Haart'
        my_add = ''

        # Use regex to extract the relevant information
        date_pattern = re.compile(r"DATE: (.*?)\n", re.DOTALL)
        contact_pattern = re.compile(r"Invoice To:(.*?)\n", re.DOTALL)
        ref_pattern = re.compile(r"Our Ref: (.*?)Invoice To:", re.DOTALL)
        address_pattern = re.compile(r"Address: (.*?)Contact:", re.DOTALL)
        pm_pattern = re.compile(r"Contact Point:(.*?) Manager\n", re.DOTALL)
        job_pattern = re.compile(r"Job Description:(.*?)Health and Safety -", re.DOTALL)

        date_wo = re.search(date_pattern, text).group(1).strip().replace('\n', '')
        contact_wo = re.search(contact_pattern, text).group(1).strip().replace('\n', '')
        ref_wo = re.search(ref_pattern, text).group(1).strip().replace('\n', '')
        address_wo = re.search(address_pattern, text).group(1).replace('\n', ', ').strip().rstrip(',')
        pm_wo = re.search(pm_pattern, text).group(1).replace('Lettings','').replace('Property','').strip()
        job_wo = re.search(job_pattern, text).group(1).strip()

        for word in address_to_remove_haart:
            address_wo = address_wo.replace(word, ',')

        for word in address_wo.split(','):
            if word.strip() != '': my_add = my_add + word.strip() + ', '

        address_wo = my_add

    elif 'Foxtons' in text:
        agency_wo = 'Foxtons'

        # CLean text for possible irregularities
        text = text.replace('Invoice No:\n', '')
        text = text.replace('Comment:', 'Problem Reported:')
        text = text.replace('Queried By:', 'Property Manager:')

        # Use regex to extract the relevant information
        date_pattern = re.compile(r"Date: (.*?)\n", re.DOTALL)
        contact_pattern = re.compile(r"Landlord:(.*?)\n", re.DOTALL)
        ref_pattern = re.compile(r"Order No:(.*?)Job Date:", re.DOTALL)
        address_pattern = re.compile(r"Re: (.*?)Contact:", re.DOTALL)
        pm_pattern = re.compile(r"Property Manager: (.*?)020", re.DOTALL)
        pm_pattern2 = re.compile(r"Property Manager: (.*?)\(", re.DOTALL)
        job_pattern = re.compile(r"Problem Reported:(.*?)Property Manager: ", re.DOTALL)

        date_wo = re.search(date_pattern, text).group(1).strip().replace('\n', '')
        contact_wo = re.search(contact_pattern, text).group(1).strip().replace('\n', '')
        ref_wo = re.search(ref_pattern, text).group(1).strip().replace('\n', '')
        address_wo = re.search(address_pattern, text).group(1).replace('\n', ', ').strip()

        if re.search(pm_pattern, text):
            pm_wo = re.search(pm_pattern, text).group(1).strip()

        if re.search(pm_pattern2, text):
            pm_wo = re.search(pm_pattern2, text).group(1).strip()

        job_wo = re.search(job_pattern, text).group(1).strip()

        # Format the date
        date_raw = date_wo
        date_obj = parser.parse(date_wo)
        date_wo = date_obj.strftime("%d/%m/%Y")

        if len(address_wo.split('Property:')) > 1:
            address_wo = address_wo.split('Property:')[1].strip()

    else:
        print(text)
        address_wo = text
        
    return {
        "Date": date_wo,
        "DateRaw": date_raw,
        "Ref": ref_wo,
        "Contact": contact_wo,
        "Agency": agency_wo,
        "PM": pm_wo,
        "Address": address_wo,
        "Work": job_wo
    }


# In[21]:


def clean_work_description (text):
    
    # Final cleaning of work description
    job_final = ""
    job_prep = text
    job_prep = re.sub(r'\d\.', '\nX]', job_prep)
    job_prep = re.sub(r'\d\)', '\nX]', job_prep)
    job_prep = re.sub(r'\.([A-Z])', r'\n\1', job_prep)
    job_prep = re.sub(r'\. ([A-Z])', r'\n\1', job_prep)
    job_prep = job_prep.replace('\n - ', '\nX]').replace('\n- ', '\nX]').replace('\n-', '\nX]')

    for line in job_prep.split('\n'):
        clean_line = line.strip()

        if clean_line != '':

            if clean_line[-1] in end_characters:

                if (clean_line[0].isupper() or clean_line[0].isdigit()):
                    job_final = job_final + '\n' + clean_line + '\n'
                else:
                    job_final = job_final + clean_line + '\n'

            elif job_final != '' and (clean_line[0].isupper() or clean_line[0].isdigit()):
                job_final = job_final + '\n' + clean_line + ' '
            else:
                if job_final == '':
                    job_final = clean_line
                else:
                    job_final = job_final + ' ' + clean_line

    return job_final.replace('\n\n','\n').replace('X] ','').replace('X]','')

    


# In[22]:


# Iterate through all the PDF files in the folder

for file_name in os.listdir(folder_path):
    if file_name.endswith('.pdf'):

        file_path = os.path.join(folder_path, file_name)

        # Extract text using PDFplumber
        with pdfplumber.open(file_path) as pdf:
            txt = ''
            desc_final = ''
            for page in pdf.pages:
                txt += page.extract_text()

            txt = txt.replace('•', '')
            
            emails = re.findall(email_pattern, txt)
            
            extracted_wo = extract_info(txt)
            
            work_cleaned = clean_work_description(extracted_wo["Work"])

            # Final PM filtering
            email_wo = ''

            for email in emails:
                if agency_wo.replace(' ', '').lower() in email:
                    
                    if email != 'llm@kfh.co.uk':
                        email_wo = email

            data = {
                "Date": extracted_wo["Date"],
                "Date Raw": extracted_wo["DateRaw"],
                "Ref": extracted_wo["Ref"],
                "Contact": extracted_wo["Contact"],
                "Agency": extracted_wo["Agency"],
                "PM": extracted_wo["PM"],
                "Email": email_wo,
                "Address": extracted_wo["Address"],
                "Postcode": ' '.join(extracted_wo["Address"].split()[-2:]).rstrip(','),
                "Work Description": extracted_wo["Work"],
                "Work Cleaned": work_cleaned,
                "Address_match": ''.join(e for e in extracted_wo["Address"] if e.isalnum()).lower()
            }

            df = df.append(data, ignore_index=True)

            # print(df)


# In[23]:


# Save or utilize the dataframe as needed

df.to_csv(folder_path + '/Extraction/extracted_data.csv', index=False, encoding='utf-8')


# In[ ]:




