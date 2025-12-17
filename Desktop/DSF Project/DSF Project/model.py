import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
from sklearn.model_selection import train_test_split
from sklearn.preprocessing import StandardScaler
from sklearn.ensemble import RandomForestRegressor, GradientBoostingRegressor
from sklearn.neighbors import KNeighborsRegressor
from sklearn.neural_network import MLPRegressor
from sklearn.metrics import r2_score, mean_squared_error
from sklearn.impute import SimpleImputer
import numpy as np
import pandas as pd
import pickle
import os
import warnings
import threading
warnings.filterwarnings('ignore')


class LeadScoringApp:
    def __init__(self, root):
        self.root = root
        self.root.title("🤖 Lead Scoring System")
        self.root.geometry("900x700")
        self.root.configure(bg='#f0f0f0')
        
        self.training_file = r"C:\Users\aliza\Desktop\DSF Project\DSF Project\Philadelphia_and_Pittsburgh_cleaned.csv"
        self.model_file = "lead_scoring_model.pkl"
        self.is_trained = os.path.exists(self.model_file)
        
        # Personal data columns that should never be predicted
        self.excluded_columns = [
            'name', 'email', 'phone', 'address', 'owner', 'officer', 
            'contact', 'first_name', 'last_name', 'full_name', 'street',
            'city', 'state', 'zip', 'postal', 'mail', 'buyer', 'seller'
        ]
        
        self.setup_ui()
        
    def setup_ui(self):
        title_frame = tk.Frame(self.root, bg='#2c3e50', height=80)
        title_frame.pack(fill='x', pady=(0, 20))
        
        title_label = tk.Label(
            title_frame,
            text="🤖 Lead Scoring & Prediction System",
            font=('Arial', 24, 'bold'),
            bg='#2c3e50',
            fg='white'
        )
        title_label.pack(pady=20)
        
        main_frame = tk.Frame(self.root, bg='#f0f0f0')
        main_frame.pack(fill='both', expand=True, padx=20, pady=10)
        
        left_frame = tk.LabelFrame(
            main_frame,
            text="📚 Step 1: Train Model",
            font=('Arial', 12, 'bold'),
            bg='white',
            padx=15,
            pady=15
        )
        left_frame.pack(side='left', fill='both', expand=True, padx=(0, 10))
        
        tk.Label(
            left_frame,
            text="Training Data:",
            font=('Arial', 10),
            bg='white'
        ).pack(anchor='w', pady=(0, 5))
        
        self.training_path_label = tk.Label(
            left_frame,
            text=self.training_file if len(self.training_file) < 50 
                 else "..." + self.training_file[-47:],
            font=('Arial', 9),
            bg='#e8e8e8',
            anchor='w',
            relief='sunken',
            padx=10,
            pady=8
        )
        self.training_path_label.pack(fill='x', pady=(0, 10))
        
        btn_frame1 = tk.Frame(left_frame, bg='white')
        btn_frame1.pack(fill='x', pady=(0, 15))
        
        self.browse_train_btn = tk.Button(
            btn_frame1,
            text="📂 Browse Training File",
            command=self.browse_training_file,
            bg='#3498db',
            fg='white',
            font=('Arial', 10, 'bold'),
            relief='flat',
            cursor='hand2',
            padx=15,
            pady=8
        )
        self.browse_train_btn.pack(side='left', padx=(0, 5))
        
        self.train_btn = tk.Button(
            btn_frame1,
            text="🚀 Train Model",
            command=self.train_model,
            bg='#27ae60',
            fg='white',
            font=('Arial', 10, 'bold'),
            relief='flat',
            cursor='hand2',
            padx=15,
            pady=8
        )
        self.train_btn.pack(side='left')
        
        self.train_status = tk.Label(
            left_frame,
            text="✅ Model Ready" if self.is_trained else "⚠️ Model Not Trained",
            font=('Arial', 10, 'bold'),
            bg='white',
            fg='#27ae60' if self.is_trained else '#e74c3c'
        )
        self.train_status.pack(pady=(0, 10))
        
        tk.Label(
            left_frame,
            text="Training Log:",
            font=('Arial', 10, 'bold'),
            bg='white'
        ).pack(anchor='w', pady=(10, 5))
        
        self.train_log = scrolledtext.ScrolledText(
            left_frame,
            height=12,
            font=('Consolas', 9),
            bg='#1e1e1e',
            fg='#00ff00',
            relief='sunken'
        )
        self.train_log.pack(fill='both', expand=True)
        
        right_frame = tk.LabelFrame(
            main_frame,
            text="🔮 Step 2: Predict Leads",
            font=('Arial', 12, 'bold'),
            bg='white',
            padx=15,
            pady=15
        )
        right_frame.pack(side='right', fill='both', expand=True, padx=(10, 0))
        
        tk.Label(
            right_frame,
            text="Select Excel/CSV File:",
            font=('Arial', 10),
            bg='white'
        ).pack(anchor='w', pady=(0, 5))
        
        path_entry_frame = tk.Frame(right_frame, bg='white')
        path_entry_frame.pack(fill='x', pady=(0, 10))
        
        self.predict_path_entry = tk.Entry(
            path_entry_frame,
            font=('Arial', 9),
            relief='sunken',
            bd=2
        )
        self.predict_path_entry.pack(side='left', fill='x', expand=True, padx=(0, 5))
        self.predict_path_entry.insert(0, "Paste file path here or browse...")
        self.predict_path_entry.config(fg='gray')
        
        def on_entry_click(event):
            if self.predict_path_entry.get() == "Paste file path here or browse...":
                self.predict_path_entry.delete(0, tk.END)
                self.predict_path_entry.config(fg='black')
        
        def on_focusout(event):
            if self.predict_path_entry.get() == "":
                self.predict_path_entry.insert(0, "Paste file path here or browse...")
                self.predict_path_entry.config(fg='gray')
            else:
                path = self.predict_path_entry.get().strip().strip('"').strip("'")
                if os.path.exists(path) and (path.endswith('.csv') or path.endswith('.xlsx') or path.endswith('.xls')):
                    self.predict_file = path
                    self.predict_path_label.config(
                        text=os.path.basename(path),
                        fg='#27ae60'
                    )
        
        self.predict_path_entry.bind('<FocusIn>', on_entry_click)
        self.predict_path_entry.bind('<FocusOut>', on_focusout)
        
        use_path_btn = tk.Button(
            path_entry_frame,
            text="✓",
            command=lambda: on_focusout(None),
            bg='#27ae60',
            fg='white',
            font=('Arial', 10, 'bold'),
            relief='flat',
            cursor='hand2',
            width=3
        )
        use_path_btn.pack(side='left')
        
        self.predict_path_label = tk.Label(
            right_frame,
            text="No file selected",
            font=('Arial', 9, 'italic'),
            bg='white',
            anchor='w',
            fg='#95a5a6'
        )
        self.predict_path_label.pack(fill='x', pady=(0, 10))
        
        btn_frame2 = tk.Frame(right_frame, bg='white')
        btn_frame2.pack(fill='x', pady=(0, 10))
        
        self.browse_predict_btn = tk.Button(
            btn_frame2,
            text="📂 Browse File",
            command=self.browse_predict_file,
            bg='#9b59b6',
            fg='white',
            font=('Arial', 10, 'bold'),
            relief='flat',
            cursor='hand2',
            padx=15,
            pady=8
        )
        self.browse_predict_btn.pack(side='left', padx=(0, 5))
        
        self.drop_area_btn = tk.Button(
            btn_frame2,
            text="📥 Help",
            command=self.show_drop_instructions,
            bg='#8e44ad',
            fg='white',
            font=('Arial', 10, 'bold'),
            relief='flat',
            cursor='hand2',
            padx=15,
            pady=8
        )
        self.drop_area_btn.pack(side='left')
        
        self.predict_btn = tk.Button(
            right_frame,
            text="🎯 Predict Leads",
            command=self.predict_leads,
            bg='#e67e22',
            fg='white',
            font=('Arial', 12, 'bold'),
            relief='flat',
            cursor='hand2',
            padx=20,
            pady=12,
            state='disabled' if not self.is_trained else 'normal'
        )
        self.predict_btn.pack(fill='x', pady=(0, 15))
        
        tk.Label(
            right_frame,
            text="Prediction Results:",
            font=('Arial', 10, 'bold'),
            bg='white'
        ).pack(anchor='w', pady=(10, 5))
        
        self.predict_log = scrolledtext.ScrolledText(
            right_frame,
            height=12,
            font=('Consolas', 9),
            bg='#1e1e1e',
            fg='#00ff00',
            relief='sunken'
        )
        self.predict_log.pack(fill='both', expand=True)
        
        self.progress_frame = tk.Frame(self.root, bg='#f0f0f0')
        self.progress_frame.pack(fill='x', padx=20, pady=(10, 5))
        
        self.progress = ttk.Progressbar(
            self.progress_frame,
            mode='indeterminate',
            length=300
        )
        self.progress.pack(side='left', fill='x', expand=True)
        
        self.progress_label = tk.Label(
            self.progress_frame,
            text="Ready",
            font=('Arial', 9),
            bg='#f0f0f0'
        )
        self.progress_label.pack(side='left', padx=(10, 0))
        
        footer = tk.Label(
            self.root,
            text="💡 Browse, Paste path, or type directly • Supports CSV and Excel files",
            font=('Arial', 9),
            bg='#f0f0f0',
            fg='#7f8c8d'
        )
        footer.pack(pady=(5, 10))
        
    def browse_training_file(self):
        filename = filedialog.askopenfilename(
            title="Select Training Data CSV",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            self.training_file = filename
            display_name = filename if len(filename) < 50 else "..." + filename[-47:]
            self.training_path_label.config(text=display_name)
    
    def show_drop_instructions(self):
        messagebox.showinfo(
            "File Input Help",
            "You can input files in 3 ways:\n\n"
            "1. 📂 Click 'Browse File' button\n"
            "2. ⌨️ Paste file path in text box\n"
            "3. 📝 Type path directly\n\n"
            "Supported formats:\n"
            "• Excel (.xlsx, .xls)\n"
            "• CSV (.csv)"
        )
    
    def browse_predict_file(self):
        filename = filedialog.askopenfilename(
            title="Select File for Prediction",
            filetypes=[
                ("All supported", "*.xlsx;*.xls;*.csv"),
                ("Excel files", "*.xlsx"),
                ("Excel files (old)", "*.xls"),
                ("CSV files", "*.csv"),
                ("All files", "*.*")
            ]
        )
        if filename:
            self.predict_file = filename
            display_name = filename if len(filename) < 50 else "..." + filename[-47:]
            self.predict_path_label.config(text=display_name, fg='#2c3e50')
            
    def log_train(self, message):
        self.train_log.insert(tk.END, message + "\n")
        self.train_log.see(tk.END)
        self.root.update()
        
    def log_predict(self, message):
        self.predict_log.insert(tk.END, message + "\n")
        self.predict_log.see(tk.END)
        self.root.update()
    
    def is_personal_data_column(self, col_name):
        col_lower = col_name.lower()
        return any(excluded in col_lower for excluded in self.excluded_columns)
    
    def engineer_features(self, df):
        df_eng = df.copy()
        
        phone_cols = [col for col in df.columns if 'phone' in col.lower()]
        email_cols = [col for col in df.columns if 'email' in col.lower()]
        office_cols = [col for col in df.columns if 'buyer_office' in col.lower() and 'name' in col.lower()]
        officer_cols = [col for col in df.columns if 'officer' in col.lower() and 'name' in col.lower()]
        
        df_eng['phone_count'] = df[phone_cols].notna().sum(axis=1)
        df_eng['email_count'] = df[email_cols].notna().sum(axis=1)
        df_eng['office_count'] = df[office_cols].notna().sum(axis=1)
        df_eng['officer_count'] = df[officer_cols].notna().sum(axis=1)
        
        if 'address' in df.columns:
            df_eng['has_address'] = df['address'].notna().astype(int)
        
        df_eng['contact_richness'] = (
            (df_eng['phone_count'] * 10) + 
            (df_eng['email_count'] * 15) + 
            (df_eng['office_count'] * 20) + 
            (df_eng['officer_count'] * 25)
        )
        df_eng['contact_richness'] = df_eng['contact_richness'].clip(upper=100)
        
        return df_eng
        
    def train_model(self):
        def train_thread():
            try:
                self.train_btn.config(state='disabled')
                self.browse_train_btn.config(state='disabled')
                self.progress.start()
                self.progress_label.config(text="Training...")
                
                self.train_log.delete(1.0, tk.END)
                self.log_train("="*60)
                self.log_train("🚀 STARTING MODEL TRAINING")
                self.log_train("="*60)
                
                self.log_train(f"\n📁 Loading: {self.training_file}")
                df_train = pd.read_csv(self.training_file)
                self.log_train(f"✅ Loaded: {df_train.shape[0]} rows × {df_train.shape[1]} columns")
                
                self.log_train("\n🔧 Engineering features...")
                df_train = self.engineer_features(df_train)
                self.log_train("✅ Created: phone_count, email_count, office_count")
                self.log_train("           officer_count, contact_richness")
                
                target_col = "score"
                y_train_full = df_train[target_col].copy()
                
                self.log_train(f"\n🎯 Target Statistics (Score):")
                self.log_train(f"   Mean: {y_train_full.mean():,.2f}")
                self.log_train(f"   Median: {y_train_full.median():,.2f}")
                self.log_train(f"   25th Percentile: {y_train_full.quantile(0.25):,.2f}")
                self.log_train(f"   75th Percentile: {y_train_full.quantile(0.75):,.2f}")
                
                SCORE_THRESHOLD = y_train_full.quantile(0.70)
                self.log_train(f"\n📊 Threshold (70th Percentile): {SCORE_THRESHOLD:,.2f}")
                self.log_train(f"   Score >= {SCORE_THRESHOLD:,.0f} → WARM LEAD 🔥 (Top 30%)")
                self.log_train(f"   Score < {SCORE_THRESHOLD:,.0f} → COLD LEAD ❄️ (Bottom 70%)")
                
                priority_features = ['bnh_score', 'phone_count', 'email_count', 
                                   'office_count', 'officer_count', 'contact_richness']
                
                numerical_cols = df_train.select_dtypes(include=['int64', 'float64']).columns.tolist()
                if target_col in numerical_cols:
                    numerical_cols.remove(target_col)
                
                # Filter out personal data columns
                numerical_cols = [col for col in numerical_cols 
                                 if not self.is_personal_data_column(col)]
                
                correlations = {}
                for col in numerical_cols:
                    if df_train[col].notna().sum() > 0 and col not in priority_features:
                        corr = df_train[col].corr(y_train_full)
                        if not np.isnan(corr) and abs(corr) > 0.01:
                            correlations[col] = abs(corr)
                
                sorted_corr = sorted(correlations.items(), key=lambda x: x[1], reverse=True)
                additional_features = [feat for feat, _ in sorted_corr[:10]]
                
                selected_features = [f for f in priority_features if f in df_train.columns] + additional_features
                
                self.log_train(f"\n✅ Selected {len(selected_features)} features (excluding personal data)")
                
                X_train_full = df_train[selected_features].copy()
                imputer = SimpleImputer(strategy="median")
                X_train_full = pd.DataFrame(
                    imputer.fit_transform(X_train_full),
                    columns=selected_features
                )
                
                X_train, X_test, y_train, y_test = train_test_split(
                    X_train_full, y_train_full, test_size=0.2, random_state=42
                )
                
                scaler = StandardScaler()
                X_train_scaled = scaler.fit_transform(X_train)
                X_test_scaled = scaler.transform(X_test)
                
                self.log_train(f"\n🤖 Training Models...")
                
                models = {}
                results = {}
                
                self.log_train("   [1/4] Gradient Boosting...")
                gb = GradientBoostingRegressor(n_estimators=100, max_depth=5, 
                                              learning_rate=0.1, random_state=42)
                gb.fit(X_train_scaled, y_train)
                y_pred_gb = gb.predict(X_test_scaled)
                r2_gb = r2_score(y_test, y_pred_gb)
                rmse_gb = np.sqrt(mean_squared_error(y_test, y_pred_gb))
                models['Gradient Boosting'] = gb
                results['Gradient Boosting'] = {'r2': r2_gb, 'rmse': rmse_gb, 'needs_scaling': True}
                self.log_train(f"       R² = {r2_gb:.4f}, RMSE = {rmse_gb:,.0f}")
                
                self.log_train("   [2/4] Random Forest...")
                rf = RandomForestRegressor(n_estimators=100, max_depth=15, 
                                          min_samples_split=10, random_state=42, n_jobs=-1)
                rf.fit(X_train, y_train)
                y_pred_rf = rf.predict(X_test)
                r2_rf = r2_score(y_test, y_pred_rf)
                rmse_rf = np.sqrt(mean_squared_error(y_test, y_pred_rf))
                models['Random Forest'] = rf
                results['Random Forest'] = {'r2': r2_rf, 'rmse': rmse_rf, 'needs_scaling': False}
                self.log_train(f"       R² = {r2_rf:.4f}, RMSE = {rmse_rf:,.0f}")
                
                self.log_train("   [3/4] Neural Network...")
                nn = MLPRegressor(hidden_layer_sizes=(100, 50), activation='relu', 
                                 solver='adam', max_iter=500, random_state=42, 
                                 early_stopping=True, validation_fraction=0.1)
                nn.fit(X_train_scaled, y_train)
                y_pred_nn = nn.predict(X_test_scaled)
                r2_nn = r2_score(y_test, y_pred_nn)
                rmse_nn = np.sqrt(mean_squared_error(y_test, y_pred_nn))
                models['Neural Network'] = nn
                results['Neural Network'] = {'r2': r2_nn, 'rmse': rmse_nn, 'needs_scaling': True}
                self.log_train(f"       R² = {r2_nn:.4f}, RMSE = {rmse_nn:,.0f}")
                
                self.log_train("   [4/4] K-Nearest Neighbors...")
                knn = KNeighborsRegressor(n_neighbors=10, weights='distance', 
                                         algorithm='auto', n_jobs=-1)
                knn.fit(X_train_scaled, y_train)
                y_pred_knn = knn.predict(X_test_scaled)
                r2_knn = r2_score(y_test, y_pred_knn)
                rmse_knn = np.sqrt(mean_squared_error(y_test, y_pred_knn))
                models['KNN'] = knn
                results['KNN'] = {'r2': r2_knn, 'rmse': rmse_knn, 'needs_scaling': True}
                self.log_train(f"       R² = {r2_knn:.4f}, RMSE = {rmse_knn:,.0f}")
                
                best_model_name = max(results.items(), key=lambda x: x[1]['r2'])[0]
                best_model = models[best_model_name]
                best_r2 = results[best_model_name]['r2']
                needs_scaling = results[best_model_name]['needs_scaling']
                
                self.log_train(f"\n🏆 Best Model: {best_model_name}")
                self.log_train(f"   R² Score: {best_r2:.4f}")
                self.log_train(f"   RMSE: {results[best_model_name]['rmse']:,.2f}")
                
                y_pred_test = best_model.predict(X_test_scaled if needs_scaling else X_test)
                test_warm_actual = (y_test >= SCORE_THRESHOLD).sum()
                test_warm_predicted = (y_pred_test >= SCORE_THRESHOLD).sum()
                
                self.log_train(f"\n📊 Test Set Classification:")
                self.log_train(f"   Actual Warm: {test_warm_actual} ({test_warm_actual/len(y_test)*100:.1f}%)")
                self.log_train(f"   Predicted Warm: {test_warm_predicted} ({test_warm_predicted/len(y_test)*100:.1f}%)")
                
                model_artifacts = {
                    'model': best_model,
                    'model_name': best_model_name,
                    'scaler': scaler,
                    'imputer': imputer,
                    'selected_features': selected_features,
                    'needs_scaling': needs_scaling,
                    'score_threshold': SCORE_THRESHOLD,
                    'threshold_percentile': 70
                }
                
                with open(self.model_file, 'wb') as f:
                    pickle.dump(model_artifacts, f)
                
                self.log_train(f"\n💾 Model saved: {self.model_file}")
                self.log_train("\n✅ TRAINING COMPLETED!")
                self.log_train("="*60)
                
                self.is_trained = True
                self.train_status.config(text="✅ Model Ready", fg='#27ae60')
                self.predict_btn.config(state='normal')
                
                messagebox.showinfo("Success", 
                    f"Model trained!\n\n"
                    f"Best Model: {best_model_name}\n"
                    f"Threshold: {SCORE_THRESHOLD:,.0f}\n"
                    f"Top 30% = WARM LEADS\n"
                    f"Bottom 70% = COLD LEADS")
                
            except Exception as e:
                self.log_train(f"\n❌ ERROR: {str(e)}")
                messagebox.showerror("Error", f"Training failed:\n{str(e)}")
            
            finally:
                self.progress.stop()
                self.progress_label.config(text="Ready")
                self.train_btn.config(state='normal')
                self.browse_train_btn.config(state='normal')
        
        threading.Thread(target=train_thread, daemon=True).start()
        
    def predict_leads(self):
        if not hasattr(self, 'predict_file'):
            messagebox.showwarning("No File", "Please select a file first!")
            return
            
        def predict_thread():
            try:
                self.predict_btn.config(state='disabled')
                self.browse_predict_btn.config(state='disabled')
                self.progress.start()
                self.progress_label.config(text="Predicting...")
                
                self.predict_log.delete(1.0, tk.END)
                self.log_predict("="*60)
                self.log_predict("🔮 STARTING PREDICTION")
                self.log_predict("="*60)
                
                self.log_predict(f"\n📂 Loading: {self.predict_file}")
                if self.predict_file.endswith('.csv'):
                    df_new = pd.read_csv(self.predict_file)
                else:
                    df_new = pd.read_excel(self.predict_file)
                self.log_predict(f"✅ Loaded: {df_new.shape[0]} rows × {df_new.shape[1]} columns")
                
                self.log_predict("\n🔧 Engineering features...")
                df_new = self.engineer_features(df_new)
                
                with open(self.model_file, 'rb') as f:
                    artifacts = pickle.load(f)
                
                model = artifacts['model']
                scaler = artifacts['scaler']
                imputer = artifacts['imputer']
                selected_features = artifacts['selected_features']
                needs_scaling = artifacts['needs_scaling']
                SCORE_THRESHOLD = artifacts['score_threshold']
                
                self.log_predict(f"🤖 Using: {artifacts['model_name']}")
                self.log_predict(f"📊 Threshold: {SCORE_THRESHOLD:,.0f}")
                
                missing_features = [f for f in selected_features if f not in df_new.columns]
                if missing_features:
                    self.log_predict(f"⚠️  {len(missing_features)} features missing (filling with zeros)")
                    for feat in missing_features:
                        df_new[feat] = 0
                
                X_new = df_new[selected_features].copy()
                X_new = pd.DataFrame(
                    imputer.transform(X_new),
                    columns=selected_features
                )
                
                if needs_scaling:
                    X_new = scaler.transform(X_new)
                
                self.log_predict("\n🎯 Making predictions...")
                predicted_scores = model.predict(X_new)
                
                lead_types = ['WARM LEAD' if score >= SCORE_THRESHOLD else 'COLD LEAD' 
                              for score in predicted_scores]
                
                df_new['predicted_score'] = predicted_scores
                df_new['lead_type'] = lead_types
                
                score_range = predicted_scores.max() - predicted_scores.min()
                if score_range > 0:
                    normalized_scores = (predicted_scores - predicted_scores.min()) / score_range * 100
                    df_new['confidence_score'] = normalized_scores.round(1)
                else:
                    df_new['confidence_score'] = 50.0
                
                warm_count = sum(1 for lt in lead_types if 'WARM' in lt)
                cold_count = len(lead_types) - warm_count
                
                self.log_predict(f"\n📊 RESULTS:")
                self.log_predict(f"   Total Leads: {len(df_new)}")
                self.log_predict(f"   Warm Leads 🔥: {warm_count} ({warm_count/len(df_new)*100:.1f}%)")
                self.log_predict(f"   Cold Leads ❄️: {cold_count} ({cold_count/len(df_new)*100:.1f}%)")
                self.log_predict(f"\n   Score Statistics:")
                self.log_predict(f"   Avg: {predicted_scores.mean():,.2f}")
                self.log_predict(f"   Min: {predicted_scores.min():,.2f}")
                self.log_predict(f"   Max: {predicted_scores.max():,.2f}")
                
                output_filename = self.predict_file.rsplit('.', 1)[0] + '_with_predictions.csv'
                df_new.to_csv(output_filename, index=False)
                
                self.log_predict(f"\n💾 Saved: {output_filename}")
                self.log_predict("\n✅ COMPLETED!")
                self.log_predict("="*60)
                
                result = messagebox.askyesno(
                    "Success",
                    f"Predictions completed!\n\n"
                    f"📊 {warm_count} Warm ({warm_count/len(df_new)*100:.1f}%) | "
                    f"{cold_count} Cold ({cold_count/len(df_new)*100:.1f}%)\n\n"
                    f"Saved to:\n{output_filename}\n\n"
                    f"Open folder?"
                )
                
                if result:
                    folder = os.path.dirname(output_filename)
                    os.startfile(folder)
                
            except Exception as e:
                import traceback
                self.log_predict(f"\n❌ ERROR: {str(e)}")
                messagebox.showerror("Error", f"Failed:\n{str(e)}")
            
            finally:
                self.progress.stop()
                self.progress_label.config(text="Ready")
                self.predict_btn.config(state='normal')
                self.browse_predict_btn.config(state='normal')
        
        threading.Thread(target=predict_thread, daemon=True).start()


if __name__ == "__main__":
    root = tk.Tk()
    app = LeadScoringApp(root)
    root.mainloop()