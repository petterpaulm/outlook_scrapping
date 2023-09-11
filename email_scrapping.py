import os
import pandas as pd
import win32com.client

def save_attachments(subject_keyword, path='attachments'):
    # Interact with Microsoft Outlook and extract Excel attachments
    outlook = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    inbox = outlook.GetDefaultFolder(6) # "6" refers to the inbox

    messages = inbox.Items
    messages = messages.Restrict(f"[Subject] LIKE '%{subject_keyword}%'")
    
    if not os.path.exists(path):
        os.makedirs(path)
    
    for message in messages:
        for attachment in message.Attachments:
            if attachment.FileName.endswith('.xlsx') or attachment.FileName.endswith('.xls'):
                attachment.SaveAsFile(os.path.join(path, attachment.FileName))
    
    return path

def compile_dataframes(path='attachments'):
    # Use pandas to read Excel files and compile them into a single DataFrame
    all_data = []
    
    for root, dirs, files in os.walk(path):
        for file in files:
            if file.endswith('.xlsx') or file.endswith('.xls'):
                df = pd.read_excel(os.path.join(root, file))
                all_data.append(df)
    
    # Compile all data into a single DataFrame
    combined_df = pd.concat(all_data, ignore_index=True)
    
    return combined_df

# Usage
subject_keyword = 'your_subject_keyword' # Replace with the keyword to find in the email subject
save_attachments(subject_keyword)
df = compile_dataframes()
print(df)
