# Excel Processor - CSV Output Only
import pandas as pd
from pathlib import Path
import re
import ast
from datetime import datetime

def col_letter_to_index(letter):
    """Convert Excel column letter to 0-based index"""
    result = 0
    for char in letter:
        result = result * 26 + (ord(char.upper()) - ord('A') + 1)
    return result - 1

def extract_required_columns(df):
    """Extract required columns"""
    all_cols = df.columns.tolist()
    
    required_positions = {
        'A': 'unique_buyer_id',
        'B': 'searched_location',
        'C': 'name',
        'E': 'address',
        'APD': 'first_name',
        'APE': 'last_name',
        'APG': 'type',
        'APH': 'score',
        'API': 'bnh_score',
        'APJ': 'officer_data_1_name',
        'APX': 'linked_customers',
        'AVM': 'buyer_office_address'
    }
    
    selected_columns = []
    col_mapping = {}
    
    for pos, field_name in required_positions.items():
        idx = col_letter_to_index(pos)
        if idx < len(all_cols):
            col_name = all_cols[idx]
            selected_columns.append(col_name)
            col_mapping[col_name] = field_name
    
    return selected_columns, col_mapping

def clean_address_column(df):
    """Clean the address column"""
    if 'address' in df.columns:
        df['address'] = df['address'].apply(lambda x: re.sub(r'\\x[0-9A-Fa-f]{2,4}', '', str(x)) if pd.notna(x) else x)
        df['address'] = df['address'].apply(lambda x: re.sub(r'\\[rnt]', ' ', str(x)) if pd.notna(x) else x)
        df['address'] = df['address'].apply(lambda x: re.sub(r'\s+', ' ', str(x)).strip() if pd.notna(x) else x)
    return df

def consolidate_data(df):
    """Consolidate duplicates by unique_buyer_id"""
    if 'unique_buyer_id' not in df.columns:
        return df
    return df.groupby('unique_buyer_id', dropna=False).first().reset_index()

def extract_buyer_office_data(df):
    """Extract Buyer Office data with categorized phones - SKIP RELATIVES"""
    if 'buyer_office_address' not in df.columns:
        return df
    
    max_entries = 10
    max_mobile = 10
    max_residential = 10
    max_other = 10
    
    all_data = []
    
    for idx, row in df.iterrows():
        val = row['buyer_office_address']
        entry_data = {
            'names': [],
            'mobile_phones': [],
            'residential_phones': [],
            'other_phones': [],
            'addresses': [],
            'emails': []
        }
        
        if pd.notna(val):
            val_str = str(val)
            try:
                if val_str.strip().startswith('['):
                    parsed = ast.literal_eval(val_str)
                    
                    if isinstance(parsed, list):
                        for item in parsed:
                            if isinstance(item, dict):
                                full_address = item.get('full', '')
                                if full_address and full_address not in entry_data['addresses']:
                                    entry_data['addresses'].append(full_address)
                                
                                skiptraces = item.get('skiptraces')
                                if skiptraces and isinstance(skiptraces, list):
                                    for skiptrace in skiptraces:
                                        if isinstance(skiptrace, dict):
                                            full_response = skiptrace.get('full_response', [])
                                            
                                            if isinstance(full_response, list):
                                                for response in full_response:
                                                    if isinstance(response, dict):
                                                        names_list = response.get('names', [])
                                                        if isinstance(names_list, list):
                                                            for name_obj in names_list:
                                                                if isinstance(name_obj, dict):
                                                                    firstname = name_obj.get('firstname', '')
                                                                    lastname = name_obj.get('lastname', '')
                                                                    full_name = f"{firstname} {lastname}".strip()
                                                                    if full_name and full_name not in entry_data['names']:
                                                                        entry_data['names'].append(full_name)
                                                        
                                                        phones = response.get('phones', [])
                                                        if isinstance(phones, list):
                                                            for phone_obj in phones:
                                                                if isinstance(phone_obj, dict):
                                                                    phone_num = phone_obj.get('phonenumber', '')
                                                                    phone_type = phone_obj.get('phonetype', '').lower()
                                                                    
                                                                    if phone_num:
                                                                        if 'mobile' in phone_type:
                                                                            if phone_num not in entry_data['mobile_phones']:
                                                                                entry_data['mobile_phones'].append(phone_num)
                                                                        elif 'residential' in phone_type or 'land line' in phone_type:
                                                                            if phone_num not in entry_data['residential_phones']:
                                                                                entry_data['residential_phones'].append(phone_num)
                                                                        else:
                                                                            if phone_num not in entry_data['other_phones']:
                                                                                entry_data['other_phones'].append(phone_num)
                                                        
                                                        emails = response.get('emails', [])
                                                        if isinstance(emails, list):
                                                            for email_obj in emails:
                                                                if isinstance(email_obj, dict):
                                                                    email = email_obj.get('email', '')
                                                                    if email and email not in entry_data['emails']:
                                                                        entry_data['emails'].append(email)
            except:
                pass
        
        entry_data['names'] = entry_data['names'][:max_entries]
        entry_data['mobile_phones'] = entry_data['mobile_phones'][:max_mobile]
        entry_data['residential_phones'] = entry_data['residential_phones'][:max_residential]
        entry_data['other_phones'] = entry_data['other_phones'][:max_other]
        entry_data['addresses'] = entry_data['addresses'][:max_entries]
        entry_data['emails'] = entry_data['emails'][:max_entries]
        
        all_data.append(entry_data)
    
    for i in range(max_entries):
        df[f'buyer_office_{i+1}_name'] = [data['names'][i] if i < len(data['names']) else '' for data in all_data]
    
    for i in range(max_mobile):
        df[f'phone_mobile_{i+1}'] = [data['mobile_phones'][i] if i < len(data['mobile_phones']) else '' for data in all_data]
    
    for i in range(max_residential):
        df[f'phone_residential_{i+1}'] = [data['residential_phones'][i] if i < len(data['residential_phones']) else '' for data in all_data]
    
    for i in range(max_other):
        df[f'phone_other_{i+1}'] = [data['other_phones'][i] if i < len(data['other_phones']) else '' for data in all_data]
    
    for i in range(max_entries):
        df[f'buyer_office_{i+1}_address'] = [data['addresses'][i] if i < len(data['addresses']) else '' for data in all_data]
    
    for i in range(max_entries):
        df[f'buyer_office_{i+1}_email'] = [data['emails'][i] if i < len(data['emails']) else '' for data in all_data]
    
    return df

def extract_officer_data(df):
    """Extract and split Officer data"""
    if 'officer_data_1_name' not in df.columns:
        return df
    
    max_officers = 5
    all_officers_data = []
    
    for idx, row in df.iterrows():
        val = row['officer_data_1_name']
        officers = []
        
        if pd.notna(val):
            val_str = str(val)
            
            if val_str.startswith('[') or val_str.startswith('{'):
                try:
                    parsed = ast.literal_eval(val_str)
                    
                    if isinstance(parsed, list):
                        for officer in parsed[:max_officers]:
                            if isinstance(officer, dict):
                                officer_data = {
                                    'name': officer.get('name', ''),
                                    'address': '',
                                    'phone': '',
                                    'email': ''
                                }
                                
                                address = officer.get('address', {})
                                if isinstance(address, dict):
                                    officer_data['address'] = address.get('full', '')
                                
                                phones = officer.get('phones', [])
                                if isinstance(phones, list) and phones:
                                    for phone_obj in phones:
                                        if isinstance(phone_obj, dict):
                                            pnum = phone_obj.get('phonenumber', '')
                                            if pnum:
                                                officer_data['phone'] = pnum
                                                break
                                
                                emails = officer.get('emails', [])
                                if isinstance(emails, list) and emails:
                                    for email_obj in emails:
                                        if isinstance(email_obj, dict):
                                            email = email_obj.get('email', '')
                                            if email:
                                                officer_data['email'] = email
                                                break
                                
                                officers.append(officer_data)
                    
                    elif isinstance(parsed, dict):
                        officer_data = {
                            'name': parsed.get('name', ''),
                            'address': '',
                            'phone': '',
                            'email': ''
                        }
                        
                        address = parsed.get('address', {})
                        if isinstance(address, dict):
                            officer_data['address'] = address.get('full', '')
                        
                        phones = parsed.get('phones', [])
                        if isinstance(phones, list) and phones:
                            for phone_obj in phones:
                                if isinstance(phone_obj, dict):
                                    pnum = phone_obj.get('phonenumber', '')
                                    if pnum:
                                        officer_data['phone'] = pnum
                                        break
                        
                        emails = parsed.get('emails', [])
                        if isinstance(emails, list) and emails:
                            for email_obj in emails:
                                if isinstance(email_obj, dict):
                                    email = email_obj.get('email', '')
                                    if email:
                                        officer_data['email'] = email
                                        break
                        
                        officers.append(officer_data)
                except:
                    pass
        
        all_officers_data.append(officers)
    
    for i in range(max_officers):
        df[f'officer_{i+1}_name'] = [officers[i]['name'] if i < len(officers) else '' for officers in all_officers_data]
        df[f'officer_{i+1}_address'] = [officers[i]['address'] if i < len(officers) else '' for officers in all_officers_data]
        df[f'officer_{i+1}_phone'] = [officers[i]['phone'] if i < len(officers) else '' for officers in all_officers_data]
        df[f'officer_{i+1}_email'] = [officers[i]['email'] if i < len(officers) else '' for officers in all_officers_data]
    
    return df

def extract_emails_phones(df):
    """Extract phone numbers and emails from ALL columns"""
    phone_pattern = re.compile(r'[\+\(]?[1-9][0-9 .\-\(\)]{8,}[0-9]')
    email_pattern = re.compile(r'\b[A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+\.[A-Z|a-z]{2,}\b')
    
    all_phones_list = []
    all_emails_list = []
    max_phones = 10
    max_emails = 5
    
    for idx, row in df.iterrows():
        phones_found = set()
        emails_found = set()
        
        for col in df.columns:
            val = row[col]
            if pd.notna(val):
                val_str = str(val)
                
                matches = phone_pattern.findall(val_str)
                for match in matches:
                    cleaned = re.sub(r'[^\d\+]', '', match)
                    if len(cleaned) >= 10:
                        phones_found.add(cleaned)
                
                email_matches = email_pattern.findall(val_str)
                for email in email_matches:
                    emails_found.add(email.upper())
        
        phones_list = sorted(phones_found)[:max_phones]
        emails_list = sorted(emails_found)[:max_emails]
        
        all_phones_list.append(phones_list)
        all_emails_list.append(emails_list)
    
    for i in range(max_phones):
        df[f'phone_number_{i+1}'] = [phones[i] if i < len(phones) else '' for phones in all_phones_list]
    
    df['phone_number_all'] = [', '.join(phones) if phones else '' for phones in all_phones_list]
    
    for i in range(max_emails):
        df[f'email_{i+1}'] = [emails[i] if i < len(emails) else '' for emails in all_emails_list]
    
    df['email_all'] = [', '.join(emails) if emails else '' for emails in all_emails_list]
    
    return df

def process_excel(input_file, output_file=None):
    """Main processing function"""
    print("Reading Excel file...")
    df_header = pd.read_excel(input_file, engine='openpyxl', nrows=0)
    selected_cols, col_mapping = extract_required_columns(df_header)
    
    if not selected_cols:
        raise Exception("Could not extract required columns")
    
    df = pd.read_excel(input_file, engine='openpyxl', usecols=selected_cols, dtype=str)
    print(f"Loaded {len(df)} rows")
    
    print("Renaming columns...")
    df = df.rename(columns=col_mapping)
    
    print("Cleaning address column...")
    df = clean_address_column(df)
    
    print("Consolidating duplicates...")
    df = consolidate_data(df)
    
    print("Extracting Buyer Office data...")
    df = extract_buyer_office_data(df)
    
    print("Extracting Officer data...")
    df = extract_officer_data(df)
    
    print("Extracting emails and phone numbers...")
    df = extract_emails_phones(df)
    
    print("Cleaning final dataset...")
    df = df[df['unique_buyer_id'].notna()]
    df = df[df['unique_buyer_id'] != '']
    df = df.drop_duplicates(subset=['unique_buyer_id'], keep='first')
    
    df['timestamp_of_import'] = datetime.utcnow().isoformat()
    df['timestamp_of_last_update'] = datetime.utcnow().isoformat()
    
    if output_file is None:
        output_file = Path(input_file).parent / f"{Path(input_file).stem}_processed.csv"
    
    df.to_csv(output_file, index=False, encoding='utf-8-sig')
    print(f"✓ CSV saved: {output_file}")
    print(f"✓ Records: {len(df)}")
    print(f"✓ Columns: {len(df.columns)}")
    
    return output_file

if __name__ == '__main__':
    input_excel = input("Enter Excel file path: ")
    process_excel(input_excel)