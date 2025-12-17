# desktop_app.py - Excel Processor v4.5 - WITH SUPABASE INTEGRATION
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import numpy as np
from pathlib import Path
import re
import json
import ast
from datetime import datetime
import requests


# SUPABASE CONFIGURATION
SUPABASE_URL = "https://muogezwetflataumekeg.supabase.co"
SUPABASE_KEY = "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJpc3MiOiJzdXBhYmFzZSIsInJlZiI6Im11b2dlendldGZsYXRhdW1la2VnIiwicm9sZSI6InNlcnZpY2Vfcm9sZSIsImlhdCI6MTc1OTkzNjUzOCwiZXhwIjoyMDc1NTEyNTM4fQ.ADVuaw1cxGlRTWa6BvEa35DL9onmPmSQuV5eMfKFGa0"
SUPABASE_TABLE = "Records"


class ExcelProcessorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Processor v4.5 - Supabase Integration")
        self.root.geometry("900x700")
        self.root.resizable(True, True)
        self.root.minsize(800, 600)
        
        self.root.configure(bg='#f8fafc')
        
        # Variables
        self.selected_file = None
        self.processing = False
        
        # Configure style
        self.setup_styles()
        
        # Create UI
        self.create_widgets()
        
    def setup_styles(self):
        style = ttk.Style()
        style.theme_use('clam')
        
        self.colors = {
            'primary': '#6366f1',
            'primary_dark': '#4f46e5',
            'success': '#10b981',
            'error': '#ef4444',
            'warning': '#f59e0b',
            'info': '#3b82f6',
            'bg_light': '#f8fafc',
            'bg_card': '#ffffff',
            'text_primary': '#1e293b',
            'text_secondary': '#64748b',
            'border': '#e2e8f0'
        }
        
        style.configure('Title.TLabel', 
                       font=('Segoe UI', 20, 'bold'), 
                       foreground=self.colors['primary'],
                       background=self.colors['bg_light'])
        
        style.configure('Subtitle.TLabel', 
                       font=('Segoe UI', 10), 
                       foreground=self.colors['text_secondary'],
                       background=self.colors['bg_light'])
        
        style.configure('Info.TLabel', 
                       font=('Segoe UI', 9), 
                       foreground=self.colors['text_secondary'],
                       background=self.colors['bg_card'])
        
        style.configure('Card.TFrame', 
                       background=self.colors['bg_card'],
                       relief='flat')
        
        style.configure('TFrame', background=self.colors['bg_light'])
        
        style.configure('Upload.TButton', 
                       font=('Segoe UI', 10, 'bold'),
                       padding=12,
                       background=self.colors['primary'])
        
        style.map('Upload.TButton',
                 background=[('active', self.colors['primary_dark'])])
        
        style.configure('Process.TButton', 
                       font=('Segoe UI', 10, 'bold'),
                       padding=12,
                       background=self.colors['success'])
        
        style.map('Process.TButton',
                 background=[('active', '#059669')])
        
        style.configure('Card.TLabelframe',
                       background=self.colors['bg_card'],
                       borderwidth=1,
                       relief='solid')
        
        style.configure('Card.TLabelframe.Label',
                       font=('Segoe UI', 10, 'bold'),
                       foreground=self.colors['text_primary'],
                       background=self.colors['bg_card'])
        
    def create_widgets(self):
        canvas = tk.Canvas(self.root, bg='#f8fafc', highlightthickness=0)
        scrollbar = ttk.Scrollbar(self.root, orient="vertical", command=canvas.yview)
        scrollable_frame = ttk.Frame(canvas, style='TFrame')
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        def _on_mousewheel(event):
            canvas.yview_scroll(int(-1*(event.delta/120)), "units")
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
        
        main_frame = ttk.Frame(scrollable_frame, style='TFrame', padding=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Header Section
        header_frame = ttk.Frame(main_frame, style='TFrame')
        header_frame.pack(fill=tk.X, pady=(0, 15))
        
        title_label = ttk.Label(header_frame, 
                               text="📊 Excel Processor v4.5", 
                               style='Title.TLabel')
        title_label.pack()
        
        subtitle_label = ttk.Label(header_frame, 
                                   text="Phone Type Categorization + Auto-Upload to Supabase", 
                                   style='Subtitle.TLabel')
        subtitle_label.pack(pady=(3, 0))
        
        db_frame = ttk.Frame(header_frame, style='TFrame')
        db_frame.pack(pady=(8, 0))
        
        db_label = ttk.Label(db_frame,
                            text="🎯 Saves CSV + Uploads to Supabase Database",
                            font=('Segoe UI', 9),
                            foreground=self.colors['success'],
                            background=self.colors['bg_light'])
        db_label.pack()
        
        range_label = ttk.Label(header_frame, 
                               text="Categorized phones by type + Relatives SKIPPED", 
                               font=('Segoe UI', 8),
                               foreground=self.colors['text_secondary'],
                               background=self.colors['bg_light'])
        range_label.pack(pady=(3, 0))
        
        # Upload Card
        upload_card = self.create_card(main_frame, "📁 Step 1: Select Your Excel File")
        upload_card.pack(fill=tk.X, pady=(0, 12))
        
        self.upload_btn = ttk.Button(upload_card, 
                                     text="Browse Excel File", 
                                     command=self.browse_file, 
                                     style='Upload.TButton')
        self.upload_btn.pack(fill=tk.X, padx=15, pady=(8, 10))
        
        self.file_label = ttk.Label(upload_card, 
                                    text="No file selected", 
                                    style='Info.TLabel',
                                    font=('Segoe UI', 9))
        self.file_label.pack(padx=15, pady=(0, 10))
        
        # Process Card
        process_card = self.create_card(main_frame, "🚀 Step 2: Process & Generate CSV")
        process_card.pack(fill=tk.X, pady=(0, 12))
        
        self.process_btn = ttk.Button(process_card, 
                                      text="Process & Generate CSV", 
                                      command=self.start_processing, 
                                      style='Process.TButton', 
                                      state='disabled')
        self.process_btn.pack(fill=tk.X, padx=15, pady=(8, 10))
        
        # Progress Section
        progress_card = ttk.Frame(main_frame, style='Card.TFrame')
        progress_card.pack(fill=tk.X, pady=(0, 12))
        
        self.progress_label = ttk.Label(progress_card, 
                                        text="", 
                                        style='Info.TLabel',
                                        font=('Segoe UI', 9))
        self.progress_label.pack(pady=(8, 3))
        
        self.progress_bar = ttk.Progressbar(progress_card, 
                                           mode='indeterminate')
        self.progress_bar.pack(fill=tk.X, padx=15, pady=(0, 8))
        
        # Console Output Card
        console_card = self.create_card(main_frame, "💻 Console Output")
        console_card.pack(fill=tk.BOTH, expand=True)
        
        console_frame = ttk.Frame(console_card, style='Card.TFrame')
        console_frame.pack(fill=tk.BOTH, expand=True, padx=15, pady=(8, 10))
        
        scrollbar_console = ttk.Scrollbar(console_frame)
        scrollbar_console.pack(side=tk.RIGHT, fill=tk.Y)
        
        self.results_text = tk.Text(console_frame, 
                                    height=15, 
                                    wrap=tk.WORD, 
                                    font=('Consolas', 9),
                                    bg='#0f172a',
                                    fg='#e2e8f0',
                                    insertbackground='#60a5fa',
                                    relief=tk.FLAT,
                                    padx=12,
                                    pady=12,
                                    borderwidth=0,
                                    yscrollcommand=scrollbar_console.set)
        self.results_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_console.config(command=self.results_text.yview)
        
        self.results_text.tag_config('success', foreground='#34d399')
        self.results_text.tag_config('error', foreground='#f87171')
        self.results_text.tag_config('info', foreground='#60a5fa')
        self.results_text.tag_config('warning', foreground='#fbbf24')
        self.results_text.tag_config('header', foreground='#a78bfa', font=('Consolas', 10, 'bold'))
        self.results_text.tag_config('highlight', foreground='#22d3ee')
        self.results_text.tag_config('critical', foreground='#ff6b9d', font=('Consolas', 9, 'bold'))
        
        welcome_msg = f"╔══════════════════════════════════════════════════════════════╗\n"
        welcome_msg += f"║  Excel Processor v4.5 - WITH SUPABASE INTEGRATION           ║\n"
        welcome_msg += f"╠══════════════════════════════════════════════════════════════╣\n"
        welcome_msg += f"║  ✓ Extracts all required columns                            ║\n"
        welcome_msg += f"║  ✓ Categorizes phones: Mobile, Residential, Land Line       ║\n"
        welcome_msg += f"║  ✓ Format: phone_mobile_1, phone_residential_1, etc.        ║\n"
        welcome_msg += f"║  ✓ SKIPS all 'relatives' sections completely                ║\n"
        welcome_msg += f"║  ✓ Extracts Buyer Office: names, phones, addresses, emails  ║\n"
        welcome_msg += f"║  ✓ Splits Officer data: name, address, phone, email         ║\n"
        welcome_msg += f"║  ✓ Finds emails from ALL columns                            ║\n"
        welcome_msg += f"║  ✓ AUTO-UPLOADS to Supabase database                        ║\n"
        welcome_msg += f"╚══════════════════════════════════════════════════════════════╝\n\n"
        welcome_msg += "> Ready to process your Excel files.\n"
        welcome_msg += "> Select a file to begin...\n"
        
        self.results_text.insert('1.0', welcome_msg, 'info')
        self.results_text.config(state='disabled')
        
    def create_card(self, parent, title):
        card = ttk.LabelFrame(parent, 
                             text=title,
                             style='Card.TLabelframe',
                             padding=8)
        return card
        
    def browse_file(self):
        filename = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xls"), ("All Files", "*.*")]
        )
        
        if filename:
            self.selected_file = filename
            file_path = Path(filename)
            self.file_label.config(text=f"✓ Selected: {file_path.name}")
            self.process_btn.config(state='normal')
            self.log_result(f"\n> File selected: {file_path.name}\n", 'success')
    
    def start_processing(self):
        if not self.selected_file or self.processing:
            return
        
        self.processing = True
        self.process_btn.config(state='disabled')
        self.upload_btn.config(state='disabled')
        self.progress_bar.start(10)
        self.progress_label.config(text="⏳ Processing file...")
        
        import threading
        thread = threading.Thread(target=self.process_file)
        thread.daemon = True
        thread.start()
    
    def process_file(self):
        try:
            self.log_result("\n" + "═"*62 + "\n", 'header')
            self.log_result("  STARTING EXCEL PROCESSING v4.5\n", 'header')
            self.log_result("═"*62 + "\n\n", 'header')
            
            # STEP 1: Read Excel file
            self.update_progress("📂 Reading Excel file...")
            self.log_result(f"[1/9] Reading file: {Path(self.selected_file).name}\n", 'info')
            
            df_header = pd.read_excel(self.selected_file, engine='openpyxl', nrows=0)
            selected_cols, col_mapping = self.extract_required_columns(df_header)
            
            if not selected_cols:
                raise Exception("Could not extract required columns")
            
            self.log_result(f"      ✓ Found {len(selected_cols)} required columns\n", 'success')
            
            df = pd.read_excel(self.selected_file, engine='openpyxl', usecols=selected_cols, dtype=str)
            self.log_result(f"      ✓ Loaded {len(df)} rows\n\n", 'success')
            
            # STEP 2: Rename columns
            self.update_progress("🏷️ Renaming columns...")
            self.log_result("[2/9] Renaming columns...\n", 'info')
            df_renamed = df.rename(columns=col_mapping)
            self.log_result(f"      ✓ Columns renamed\n\n", 'success')
            
            # STEP 3: Clean address column
            self.update_progress("🧹 Cleaning address column...")
            self.log_result("[3/9] Cleaning address column...\n", 'info')
            df_renamed = self.clean_address_column(df_renamed)
            self.log_result(f"      ✓ Address column cleaned\n\n", 'success')
            
            # STEP 4: Consolidate duplicates
            self.update_progress("🔄 Consolidating duplicates...")
            self.log_result("[4/9] Consolidating by unique_buyer_id...\n", 'info')
            df_consolidated = self.consolidate_data(df_renamed)
            duplicates_removed = len(df_renamed) - len(df_consolidated)
            self.log_result(f"      ✓ {len(df_consolidated)} unique records\n", 'success')
            self.log_result(f"      ✓ Removed {duplicates_removed} duplicates\n\n", 'success')
            
            # STEP 5: Extract Buyer Office data
            self.update_progress("📞 Extracting Buyer Office data...")
            self.log_result("[5/9] Extracting Buyer Office (categorized phones, skip relatives)...\n", 'info')
            df_consolidated = self.extract_buyer_office_data_with_names(df_consolidated)
            
            # STEP 6: Extract Officer Data
            self.update_progress("👤 Extracting Officer data...")
            self.log_result("[6/9] Extracting Officer data...\n", 'info')
            df_consolidated = self.extract_and_split_officer_data(df_consolidated)
            
            # STEP 7: Extract emails
            self.update_progress("📧 Extracting emails...")
            self.log_result("[7/9] Extracting emails from all columns...\n", 'info')
            df_consolidated = self.extract_all_phone_numbers_and_emails(df_consolidated)
            self.log_result(f"      ✓ Phone numbers & emails extracted\n\n", 'success')
            
            # STEP 8: Clean duplicates
            self.update_progress("🧹 Cleaning duplicates...")
            self.log_result("[8/9] Removing duplicate unique_buyer_ids...\n", 'info')
            before_clean = len(df_consolidated)
            df_consolidated = df_consolidated[df_consolidated['unique_buyer_id'].notna()]
            df_consolidated = df_consolidated[df_consolidated['unique_buyer_id'] != '']
            df_consolidated = df_consolidated.drop_duplicates(subset=['unique_buyer_id'], keep='first')
            after_clean = len(df_consolidated)
            removed = before_clean - after_clean
            self.log_result(f"      ✓ Removed {removed} duplicate/empty IDs\n", 'success')
            self.log_result(f"      ✓ Final dataset: {after_clean} unique records\n\n", 'success')
            
            # Add system fields
            now = datetime.utcnow().isoformat()
            df_consolidated['timestamp_of_import'] = now
            df_consolidated['timestamp_of_last_update'] = now
            
            # STEP 9: Save CSV
            self.update_progress("💾 Saving CSV...")
            self.log_result("[9/9] Saving CSV & uploading to Supabase...\n", 'info')
            
            output_path = Path(self.selected_file).parent / f"{Path(self.selected_file).stem}_processed.csv"
            df_consolidated.to_csv(output_path, index=False, encoding='utf-8-sig')
            self.log_result(f"      ✓ Saved: {output_path.name}\n\n", 'success')
            
            # Upload to Supabase
            self.update_progress("☁️ Uploading to Supabase...")
            success_count, error_count = self.upload_to_supabase(df_consolidated)
            
            if error_count > 0:
                self.log_result(f"      ⚠ Uploaded {success_count} records, {error_count} failed\n\n", 'warning')
            else:
                self.log_result(f"      ✓ Successfully uploaded {success_count} records to Supabase\n\n", 'success')
            
            # Summary
            self.log_result("═"*62 + "\n", 'header')
            self.log_result("  PROCESSING COMPLETE\n", 'header')
            self.log_result("═"*62 + "\n", 'header')
            self.log_result(f"  Original rows:        {len(df)}\n")
            self.log_result(f"  Final rows:           {after_clean}\n")
            self.log_result(f"  Final columns:        {len(df_consolidated.columns)}\n")
            self.log_result(f"  Duplicates removed:   {duplicates_removed + removed}\n")
            self.log_result(f"  Supabase uploaded:    {success_count}\n")
            self.log_result(f"  Upload failed:        {error_count}\n")
            self.log_result("═"*62 + "\n\n", 'header')
            
            # List all columns
            self.log_result("📋 ALL COLUMNS IN OUTPUT:\n", 'critical')
            for i, col in enumerate(df_consolidated.columns, 1):
                self.log_result(f"   {i:3}. {col}\n", 'info')
            
            self.log_result(f"\n✅ Processing complete!\n", 'success')
            self.log_result(f"   📁 CSV: {output_path}\n", 'success')
            self.log_result(f"   ☁️  Supabase: {success_count} records uploaded\n", 'success')
            self.update_progress("✅ Complete!")
            
            self.root.after(0, lambda: messagebox.showinfo("✅ Success", 
                f"Processing complete!\n\n"
                f"✓ {after_clean} records processed\n"
                f"✓ {len(df_consolidated.columns)} columns\n"
                f"✓ CSV: {output_path.name}\n"
                f"✓ Supabase: {success_count} uploaded"))
            
        except Exception as e:
            error_msg = f"Error: {str(e)}"
            self.log_result(f"\n❌ {error_msg}\n", 'error')
            import traceback
            self.log_result(f"\n{traceback.format_exc()}\n", 'error')
            self.root.after(0, lambda: messagebox.showerror("❌ Error", error_msg))
        
        finally:
            self.processing = False
            self.root.after(0, self.reset_ui)
    
    def update_progress(self, message):
        self.root.after(0, lambda: self.progress_label.config(text=message))
    
    def log_result(self, message, tag=''):
        def update():
            self.results_text.config(state='normal')
            if tag:
                self.results_text.insert(tk.END, message, tag)
            else:
                self.results_text.insert(tk.END, message)
            self.results_text.see(tk.END)
            self.results_text.config(state='disabled')
        self.root.after(0, update)
    
    def reset_ui(self):
        self.progress_bar.stop()
        self.progress_label.config(text="")
        self.process_btn.config(state='normal')
        self.upload_btn.config(state='normal')
    
    # ===================================================================
    # CORE PROCESSING FUNCTIONS
    # ===================================================================
    
    def col_letter_to_index(self, letter):
        """Convert Excel column letter to 0-based index"""
        result = 0
        for char in letter:
            result = result * 26 + (ord(char.upper()) - ord('A') + 1)
        return result - 1
    
    def extract_required_columns(self, df):
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
            idx = self.col_letter_to_index(pos)
            if idx < len(all_cols):
                col_name = all_cols[idx]
                selected_columns.append(col_name)
                col_mapping[col_name] = field_name
                self.log_result(f"      ✓ {pos}: '{col_name}' → {field_name}\n", 'info')
            else:
                self.log_result(f"      ⚠ {pos} not found\n", 'warning')
        
        return selected_columns, col_mapping
    
    def clean_address_column(self, df):
        """Clean the address column by removing garbage characters"""
        if 'address' in df.columns:
            df['address'] = df['address'].apply(lambda x: re.sub(r'\\x[0-9A-Fa-f]{2,4}', '', str(x)) if pd.notna(x) else x)
            df['address'] = df['address'].apply(lambda x: re.sub(r'\\[rnt]', ' ', str(x)) if pd.notna(x) else x)
            df['address'] = df['address'].apply(lambda x: re.sub(r'\s+', ' ', str(x)).strip() if pd.notna(x) else x)
            self.log_result(f"      ✓ Cleaned garbage characters from address\n", 'success')
        return df
    
    def consolidate_data(self, df):
        """Consolidate duplicates by unique_buyer_id"""
        if 'unique_buyer_id' not in df.columns:
            return df
        return df.groupby('unique_buyer_id', dropna=False).first().reset_index()
    
    def extract_buyer_office_data_with_names(self, df):
        """Extract names, phones (CATEGORIZED), addresses, emails - SKIP RELATIVES"""
        if 'buyer_office_address' not in df.columns:
            return df
        
        self.log_result("[ENHANCED] Extracting Buyer Office with categorized phones & SKIP relatives...\n", 'info')
        
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
                    # Use ast.literal_eval for Python literal strings (single quotes)
                    if val_str.strip().startswith('['):
                        parsed = ast.literal_eval(val_str)
                        
                        if isinstance(parsed, list):
                            for item in parsed:
                                if isinstance(item, dict):
                                    # Get address
                                    full_address = item.get('full', '')
                                    if full_address and full_address not in entry_data['addresses']:
                                        entry_data['addresses'].append(full_address)
                                    
                                    # Process skiptraces
                                    skiptraces = item.get('skiptraces')
                                    if skiptraces and isinstance(skiptraces, list):
                                        for skiptrace in skiptraces:
                                            if isinstance(skiptrace, dict):
                                                full_response = skiptrace.get('full_response', [])
                                                
                                                if isinstance(full_response, list):
                                                    for response in full_response:
                                                        if isinstance(response, dict):
                                                            # Extract names (ONLY from 'names', NOT relatives)
                                                            names_list = response.get('names', [])
                                                            if isinstance(names_list, list):
                                                                for name_obj in names_list:
                                                                    if isinstance(name_obj, dict):
                                                                        firstname = name_obj.get('firstname', '')
                                                                        lastname = name_obj.get('lastname', '')
                                                                        full_name = f"{firstname} {lastname}".strip()
                                                                        if full_name and full_name not in entry_data['names']:
                                                                            entry_data['names'].append(full_name)
                                                            
                                                            # Extract CATEGORIZED phones (NOT from relatives)
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
                                                            
                                                            # Extract emails (NOT from relatives)
                                                            emails = response.get('emails', [])
                                                            if isinstance(emails, list):
                                                                for email_obj in emails:
                                                                    if isinstance(email_obj, dict):
                                                                        email = email_obj.get('email', '')
                                                                        if email and email not in entry_data['emails']:
                                                                            entry_data['emails'].append(email)
                                                            
                                                            # CRITICAL: DO NOT PROCESS relatives
                                                            # relatives = response.get('relatives', [])  # IGNORED
                
                except Exception as e:
                    self.log_result(f"         ⚠ Parse error row {idx}: {str(e)}\n", 'warning')
            
            # Limit to max
            entry_data['names'] = entry_data['names'][:max_entries]
            entry_data['mobile_phones'] = entry_data['mobile_phones'][:max_mobile]
            entry_data['residential_phones'] = entry_data['residential_phones'][:max_residential]
            entry_data['other_phones'] = entry_data['other_phones'][:max_other]
            entry_data['addresses'] = entry_data['addresses'][:max_entries]
            entry_data['emails'] = entry_data['emails'][:max_entries]
            
            all_data.append(entry_data)
        
        # Create name columns
        for i in range(max_entries):
            col_name = f'buyer_office_{i+1}_name'
            df[col_name] = [data['names'][i] if i < len(data['names']) else '' for data in all_data]
        
        # Create MOBILE phone columns
        for i in range(max_mobile):
            col_name = f'phone_mobile_{i+1}'
            df[col_name] = [data['mobile_phones'][i] if i < len(data['mobile_phones']) else '' for data in all_data]
        
        # Create RESIDENTIAL phone columns
        for i in range(max_residential):
            col_name = f'phone_residential_{i+1}'
            df[col_name] = [data['residential_phones'][i] if i < len(data['residential_phones']) else '' for data in all_data]
        
        # Create OTHER phone columns
        for i in range(max_other):
            col_name = f'phone_other_{i+1}'
            df[col_name] = [data['other_phones'][i] if i < len(data['other_phones']) else '' for data in all_data]
        
        # Create address columns
        for i in range(max_entries):
            col_name = f'buyer_office_{i+1}_address'
            df[col_name] = [data['addresses'][i] if i < len(data['addresses']) else '' for data in all_data]
        
        # Create email columns
        for i in range(max_entries):
            col_name = f'buyer_office_{i+1}_email'
            df[col_name] = [data['emails'][i] if i < len(data['emails']) else '' for data in all_data]
        
        # Statistics
        total_mobile = sum(len(data['mobile_phones']) for data in all_data)
        total_residential = sum(len(data['residential_phones']) for data in all_data)
        total_other = sum(len(data['other_phones']) for data in all_data)
        
        self.log_result(f"         ✓ Created {max_entries} name columns\n", 'success')
        self.log_result(f"         ✓ Created {max_mobile} mobile phone columns\n", 'success')
        self.log_result(f"         ✓ Created {max_residential} residential phone columns\n", 'success')
        self.log_result(f"         ✓ Created {max_other} other phone columns\n", 'success')
        self.log_result(f"         ✓ Created {max_entries} address columns\n", 'success')
        self.log_result(f"         ✓ Created {max_entries} email columns\n", 'success')
        self.log_result(f"         ℹ Total mobile: {total_mobile}\n", 'info')
        self.log_result(f"         ℹ Total residential: {total_residential}\n", 'info')
        self.log_result(f"         ℹ Total other: {total_other}\n", 'info')
        self.log_result(f"         ✓ RELATIVES SKIPPED\n\n", 'success')
        
        return df
    
    def extract_and_split_officer_data(self, df):
        """Extract and split Officer data"""
        if 'officer_data_1_name' not in df.columns:
            return df
        
        self.log_result("[ENHANCED] Splitting Officer data...\n", 'info')
        
        max_officers = 5
        all_officers_data = []
        
        for idx, row in df.iterrows():
            val = row['officer_data_1_name']
            officers = []
            
            if pd.notna(val):
                val_str = str(val)
                
                if val_str.startswith('[') or val_str.startswith('{'):
                    try:
                        # Use ast.literal_eval for Python literal strings
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
        
        # Create columns for each officer
        for i in range(max_officers):
            df[f'officer_{i+1}_name'] = [officers[i]['name'] if i < len(officers) else '' for officers in all_officers_data]
            df[f'officer_{i+1}_address'] = [officers[i]['address'] if i < len(officers) else '' for officers in all_officers_data]
            df[f'officer_{i+1}_phone'] = [officers[i]['phone'] if i < len(officers) else '' for officers in all_officers_data]
            df[f'officer_{i+1}_email'] = [officers[i]['email'] if i < len(officers) else '' for officers in all_officers_data]
        
        total_officers = sum(1 for officers in all_officers_data if officers)
        self.log_result(f"         ✓ Created {max_officers} officer sets\n", 'success')
        self.log_result(f"         ✓ Each: name, address, phone, email\n", 'success')
        self.log_result(f"         ✓ Found in {total_officers} rows\n\n", 'success')
        
        return df
    
    def extract_all_phone_numbers_and_emails(self, df):
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
                    
                    # Find phone numbers
                    matches = phone_pattern.findall(val_str)
                    for match in matches:
                        cleaned = re.sub(r'[^\d\+]', '', match)
                        if len(cleaned) >= 10:
                            phones_found.add(cleaned)
                    
                    # Find emails
                    email_matches = email_pattern.findall(val_str)
                    for email in email_matches:
                        emails_found.add(email.upper())
            
            phones_list = sorted(phones_found)[:max_phones]
            emails_list = sorted(emails_found)[:max_emails]
            
            all_phones_list.append(phones_list)
            all_emails_list.append(emails_list)
        
        # Create phone columns
        for i in range(max_phones):
            df[f'phone_number_{i+1}'] = [phones[i] if i < len(phones) else '' for phones in all_phones_list]
        
        df['phone_number_all'] = [', '.join(phones) if phones else '' for phones in all_phones_list]
        
        # Create email columns
        for i in range(max_emails):
            df[f'email_{i+1}'] = [emails[i] if i < len(emails) else '' for emails in all_emails_list]
        
        df['email_all'] = [', '.join(emails) if emails else '' for emails in all_emails_list]
        
        phones_with_data = sum(1 for phones in all_phones_list if phones)
        emails_with_data = sum(1 for emails in all_emails_list if emails)
        
        self.log_result(f"         ✓ Created {max_phones} phone columns\n", 'success')
        self.log_result(f"         ✓ Created {max_emails} email columns\n", 'success')
        self.log_result(f"         ✓ Phones in {phones_with_data} rows\n", 'success')
        self.log_result(f"         ✓ Emails in {emails_with_data} rows\n", 'success')
        
        return df
    
    def upload_to_supabase(self, df):
        """Upload data to Supabase using REST API with upsert"""
        headers = {
            "apikey": SUPABASE_KEY,
            "Authorization": f"Bearer {SUPABASE_KEY}",
            "Content-Type": "application/json",
            "Prefer": "resolution=merge-duplicates"
        }
        
        url = f"{SUPABASE_URL}/rest/v1/{SUPABASE_TABLE}"
        
        # Convert DataFrame to list of dicts
        records = df.to_dict('records')
        
        # Clean the data for Supabase - convert empty strings to None (NULL)
        cleaned_records = []
        for record in records:
            cleaned_record = {}
            for key, value in record.items():
                # Convert NaN, empty strings, and 'nan' strings to None (NULL in database)
                if pd.isna(value):
                    cleaned_record[key] = None
                elif isinstance(value, str):
                    # Empty string or whitespace becomes None
                    if value.strip() == '' or value.lower() == 'nan':
                        cleaned_record[key] = None
                    else:
                        cleaned_record[key] = value
                else:
                    cleaned_record[key] = value
            cleaned_records.append(cleaned_record)
        
        success_count = 0
        error_count = 0
        batch_size = 100
        
        total_batches = (len(cleaned_records) + batch_size - 1) // batch_size
        
        for i in range(0, len(cleaned_records), batch_size):
            batch = cleaned_records[i:i + batch_size]
            batch_num = (i // batch_size) + 1
            
            try:
                self.log_result(f"         → Uploading batch {batch_num}/{total_batches} ({len(batch)} records)...\n", 'info')
                
                response = requests.post(url, headers=headers, json=batch, timeout=30)
                
                if response.status_code in [200, 201]:
                    success_count += len(batch)
                    self.log_result(f"         ✓ Batch {batch_num} uploaded successfully\n", 'success')
                else:
                    error_count += len(batch)
                    self.log_result(f"         ✗ Batch {batch_num} failed: {response.status_code} - {response.text[:200]}\n", 'error')
            
            except Exception as e:
                error_count += len(batch)
                self.log_result(f"         ✗ Batch {batch_num} error: {str(e)}\n", 'error')
        
        return success_count, error_count


def main():
    """Main entry point"""
    root = tk.Tk()
    app = ExcelProcessorApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()