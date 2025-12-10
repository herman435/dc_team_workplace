# back_app.py - COMBINED VERSION WITH ALL FEATURES
from flask import Flask, request, jsonify, send_file, render_template
from flask_cors import CORS
import pandas as pd
import os
from datetime import datetime
from werkzeug.utils import secure_filename

app = Flask(__name__)
CORS(app)

# Configuration
UPLOAD_FOLDER = 'uploads'
OUTPUT_FOLDER = 'output'
ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'csv'}
MAX_CONTENT_LENGTH = 100 * 1024 * 1024

# Create necessary folders
for folder in [UPLOAD_FOLDER, OUTPUT_FOLDER]:
    os.makedirs(folder, exist_ok=True)

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def read_excel_file(filepath):
    """Read Excel or CSV file"""
    try:
        if filepath.endswith('.csv'):
            return pd.read_csv(filepath)
        else:
            return pd.read_excel(filepath, engine='openpyxl')
    except Exception as e:
        raise Exception(f"Error reading file {filepath}: {str(e)}")

def read_excel_sheet(filepath, sheet_name):
    """Read specific sheet from Excel file"""
    try:
        if filepath.endswith('.csv'):
            return pd.read_csv(filepath)
        else:
            return pd.read_excel(filepath, sheet_name=sheet_name, engine='openpyxl')
    except Exception as e:
        raise Exception(f"Error reading sheet '{sheet_name}' from {filepath}: {str(e)}")

def save_dataframes_to_excel(dfs_dict, filepath):
    """Save multiple DataFrames to Excel"""
    try:
        with pd.ExcelWriter(filepath, engine='openpyxl') as writer:
            for sheet_name, df in dfs_dict.items():
                df.to_excel(writer, sheet_name=sheet_name, index=False)
    except Exception as e:
        raise Exception(f"Error saving file {filepath}: {str(e)}")

def load_sales_category_mapping(filepath):
    """Load Sales Category mapping"""
    try:
        df = read_excel_file(filepath)
        if 'Tenant' not in df.columns or 'Sales Category' not in df.columns:
            raise Exception("Sales Category file must have 'Tenant' and 'Sales Category' columns")
        
        mapping = pd.Series(
            df['Sales Category'].astype(str).str.strip().values,
            index=df['Tenant'].astype(str).str.strip().values
        ).to_dict()
        return mapping
    except Exception as e:
        raise Exception(f"Error loading Sales Category file: {str(e)}")

def add_sales_category(df, sales_category_mapping):
    """Add Sales Category to dataframe"""
    if 'ShopName' not in df.columns:
        raise Exception("Market report must have 'ShopName' column")
    
    if 'Sales Category' not in df.columns:
        df['Sales Category'] = ''
    
    df['Sales Category'] = df['ShopName'].fillna('').astype(str).str.strip().map(
        lambda x: sales_category_mapping.get(x, '')
    )
    
    return df

def get_client_short_name(client_name):
    """Extract short name from client name"""
    if pd.isna(client_name):
        return 'Unknown'
    
    short_name_map = {
        'SJS': 'SJS', 'KTMSS': 'KTMSS', 'HKWDA': 'HKWDA',
        'PLK': 'PLK', 'YOT': 'YOT', 'SIS': 'SIS'
    }
    
    client_upper = str(client_name).upper()
    for key in short_name_map:
        if key in client_upper:
            return short_name_map[key]
    
    return str(client_name).split()[0] if client_name else 'Unknown'

def generate_coupon_purchased_table(weekly_df):
    """Generate coupon purchased by client and value"""
    if weekly_df.empty or 'client_name' not in weekly_df.columns:
        return None
    
    try:
        # Check if CouponValue exists, if not use coupon_amount
        value_col = 'CouponValue' if 'CouponValue' in weekly_df.columns else 'coupon_amount'
        
        if value_col not in weekly_df.columns:
            print(f"Warning: Neither CouponValue nor coupon_amount found in columns")
            return None
        
        # Group by client and coupon value, count occurrences
        coupon_data = weekly_df.groupby(['client_name', value_col]).size().unstack(fill_value=0)
        
        # Get short names for clients
        coupon_data.index = [get_client_short_name(c) for c in coupon_data.index]
        
        # Reset index to make client names a column
        coupon_data_reset = coupon_data.reset_index()
        coupon_data_reset.columns.name = None
        coupon_data_reset.rename(columns={'index': 'Client'}, inplace=True)
        
        # Convert to HTML table
        table_html = coupon_data_reset.to_html(classes='table table-striped border', index=False)
        
        # Also return as dict for JSON
        table_dict = coupon_data.to_dict('index')
        
        return {
            'title': 'No. of Coupon Purchased',
            'html': table_html,
            'data': table_dict,
            'columns': coupon_data_reset.columns.tolist()
        }
    except Exception as e:
        print(f"Error generating coupon_purchased table: {e}")
        import traceback
        traceback.print_exc()
        return None

def generate_distributed_coupon_table(weekly_df):
    """Generate distributed coupon (used vs unused) by client"""
    if weekly_df.empty or 'client_name' not in weekly_df.columns:
        return None
    
    try:
        # Count total coupons per client
        total_by_client = weekly_df.groupby('client_name').size()
        
        # Count used coupons (all records in weekly are used)
        used_by_client = weekly_df.groupby('client_name').size()
        
        # Create dataframe
        distributed = pd.DataFrame({
            'Client': [get_client_short_name(c) for c in total_by_client.index],
            'Used': used_by_client.values,
            'Unused': 0,  # All are used in this context
            'Total': used_by_client.values
        })
        
        table_html = distributed.to_html(classes='table table-striped border', index=False)
        table_dict = distributed.to_dict('records')
        
        return {
            'title': 'No. of Distributed Coupon',
            'html': table_html,
            'data': table_dict,
            'columns': distributed.columns.tolist()
        }
    except Exception as e:
        print(f"Error generating distributed_coupon table: {e}")
        return None

def generate_used_market_table(weekly_df, daily_df):
    """Generate used market table for both weekly and all records"""
    results = {}
    
    # Determine which column to use for Market (MarketCode or NameChinese)
    market_col_weekly = None
    market_col_daily = None
    
    if 'Market' in weekly_df.columns:
        market_col_weekly = 'Market'
    elif 'NameChinese' in weekly_df.columns:
        market_col_weekly = 'NameChinese'
    elif 'MarketCode' in weekly_df.columns:
        market_col_weekly = 'MarketCode'
    
    if 'Market' in daily_df.columns:
        market_col_daily = 'Market'
    elif 'NameChinese' in daily_df.columns:
        market_col_daily = 'NameChinese'
    elif 'MarketCode' in daily_df.columns:
        market_col_daily = 'MarketCode'
    
    # Weekly report
    if not weekly_df.empty and market_col_weekly:
        try:
            market_usage = weekly_df[market_col_weekly].value_counts().sort_values(ascending=False)
            df = pd.DataFrame({
                'Market': market_usage.index,
                'No. of Coupon Used': market_usage.values
            })
            table_html = df.to_html(classes='table table-striped border', index=False)
            table_dict = df.to_dict('records')
            results['weekly'] = {
                'title': 'Used Market (Weekly)',
                'html': table_html,
                'data': table_dict,
                'columns': df.columns.tolist()
            }
        except Exception as e:
            print(f"Error generating weekly used_market table: {e}")
            import traceback
            traceback.print_exc()
    
    # All records
    if not daily_df.empty and market_col_daily:
        try:
            market_usage = daily_df[market_col_daily].value_counts().sort_values(ascending=False)
            df = pd.DataFrame({
                'Market': market_usage.index,
                'No. of Coupon Used': market_usage.values
            })
            table_html = df.to_html(classes='table table-striped border', index=False)
            table_dict = df.to_dict('records')
            results['all'] = {
                'title': 'Used Market (All Records)',
                'html': table_html,
                'data': table_dict,
                'columns': df.columns.tolist()
            }
        except Exception as e:
            print(f"Error generating all used_market table: {e}")
            import traceback
            traceback.print_exc()
    
    return results if results else None

def generate_top_merchant_table(weekly_df, daily_df):
    """Generate top merchant table for both weekly and all records"""
    results = {}
    
    # Determine which column to use for Merchant
    merchant_col_weekly = 'ShopName' if 'ShopName' in weekly_df.columns else None
    merchant_col_daily = 'ShopName' if 'ShopName' in daily_df.columns else None
    
    # Weekly report
    if not weekly_df.empty and merchant_col_weekly:
        try:
            top_merchants = weekly_df[merchant_col_weekly].value_counts().head(15).sort_values(ascending=False)
            df = pd.DataFrame({
                'Merchant': top_merchants.index,
                'No. of Coupon Used': top_merchants.values
            })
            table_html = df.to_html(classes='table table-striped border', index=False)
            table_dict = df.to_dict('records')
            results['weekly'] = {
                'title': 'Top Merchant (Weekly)',
                'html': table_html,
                'data': table_dict,
                'columns': df.columns.tolist()
            }
        except Exception as e:
            print(f"Error generating weekly top_merchant table: {e}")
            import traceback
            traceback.print_exc()
    
    # All records
    if not daily_df.empty and merchant_col_daily:
        try:
            top_merchants = daily_df[merchant_col_daily].value_counts().head(15).sort_values(ascending=False)
            df = pd.DataFrame({
                'Merchant': top_merchants.index,
                'No. of Coupon Used': top_merchants.values
            })
            table_html = df.to_html(classes='table table-striped border', index=False)
            table_dict = df.to_dict('records')
            results['all'] = {
                'title': 'Top Merchant (All Records)',
                'html': table_html,
                'data': table_dict,
                'columns': df.columns.tolist()
            }
        except Exception as e:
            print(f"Error generating all top_merchant table: {e}")
            import traceback
            traceback.print_exc()
    
    return results if results else None

def generate_merchant_category_table(weekly_df, daily_df):
    """Generate merchant category table for both weekly and all records"""
    results = {}
    
    # Weekly report
    if not weekly_df.empty and 'Sales Category' in weekly_df.columns:
        try:
            category_usage = weekly_df['Sales Category'].value_counts().sort_values(ascending=False)
            df = pd.DataFrame({
                'Merchant Category': category_usage.index,
                'No. of Coupon Used': category_usage.values
            })
            table_html = df.to_html(classes='table table-striped border', index=False)
            table_dict = df.to_dict('records')
            results['weekly'] = {
                'title': 'Merchant Category (Weekly)',
                'html': table_html,
                'data': table_dict,
                'columns': df.columns.tolist()
            }
        except Exception as e:
            print(f"Error generating weekly merchant_category table: {e}")
    
    # All records
    if not daily_df.empty and 'Sales Category' in daily_df.columns:
        try:
            category_usage = daily_df['Sales Category'].value_counts().sort_values(ascending=False)
            df = pd.DataFrame({
                'Merchant Category': category_usage.index,
                'No. of Coupon Used': category_usage.values
            })
            table_html = df.to_html(classes='table table-striped border', index=False)
            table_dict = df.to_dict('records')
            results['all'] = {
                'title': 'Merchant Category (All Records)',
                'html': table_html,
                'data': table_dict,
                'columns': df.columns.tolist()
            }
        except Exception as e:
            print(f"Error generating all merchant_category table: {e}")
    
    return results if results else None

def generate_pivot_worksheet(daily_df):
    """Generate Pivot worksheet with coupon count grouped by Market and Tenant"""
    try:
        # Determine market column
        market_col = None
        if 'NameChinese' in daily_df.columns:
            market_col = 'NameChinese'
        elif 'Market' in daily_df.columns:
            market_col = 'Market'
        elif 'MarketCode' in daily_df.columns:
            market_col = 'MarketCode'
        
        if not market_col or 'ShopName' not in daily_df.columns:
            print("Cannot generate pivot worksheet - missing required columns")
            return None
        
        # Group by Market and Tenant, count records
        tenant_data = daily_df.groupby([market_col, 'ShopName']).size().reset_index(name='Count_of_coupon_uid')
        
        # Replace "MARAE LIMITED" with "U士多"
        tenant_data['ShopName'] = tenant_data['ShopName'].apply(
            lambda x: str(x).replace("MARAE LIMITED", "U士多") if "MARAE LIMITED" in str(x) else x
        )
        
        # Calculate total per market
        market_totals = daily_df.groupby(market_col).size().reset_index(name='Count_of_coupon_uid')
        
        # Build the final dataframe with market totals followed by tenants
        result_rows = []
        
        for market in sorted(market_totals[market_col].unique()):
            # Add market total row
            market_total = market_totals[market_totals[market_col] == market]['Count_of_coupon_uid'].values[0]
            result_rows.append({
                'Market': market,
                'Tenant': '',  # Empty for market total row
                'Count_of_coupon_uid': market_total
            })
            
            # Add tenant rows for this market, sorted by count descending
            market_tenants = tenant_data[tenant_data[market_col] == market].sort_values(
                'Count_of_coupon_uid', ascending=False
            )
            
            for _, row in market_tenants.iterrows():
                result_rows.append({
                    'Market': '',  # Empty for tenant rows
                    'Tenant': row['ShopName'],
                    'Count_of_coupon_uid': row['Count_of_coupon_uid']
                })
        
        pivot_data = pd.DataFrame(result_rows)
        
        print(f"✓ Generated Pivot worksheet with {len(pivot_data)} rows")
        return pivot_data
        
    except Exception as e:
        print(f"Error generating pivot worksheet: {e}")
        import traceback
        traceback.print_exc()
        return None

def generate_market_worksheets(daily_df, sales_category_mapping):
    """Generate individual worksheets for each market with tenant analysis"""
    market_sheets = {}
    
    # Determine market column
    market_col = None
    if 'NameChinese' in daily_df.columns:
        market_col = 'NameChinese'
    elif 'Market' in daily_df.columns:
        market_col = 'Market'
    elif 'MarketCode' in daily_df.columns:
        market_col = 'MarketCode'
    
    if not market_col or 'ShopName' not in daily_df.columns:
        print("Cannot generate market worksheets - missing required columns")
        return market_sheets
    
    try:
        # Get unique markets
        markets = daily_df[market_col].dropna().unique()
        
        for market in markets:
            # Filter data for this market ONLY
            market_data = daily_df[daily_df[market_col] == market].copy()
            
            if market_data.empty:
                continue
            
            # Left side: Data of Coupon Used (by tenant with total count for THIS MARKET)
            # Count coupons used per tenant in this market
            tenant_usage = market_data['ShopName'].value_counts().sort_values(ascending=False)
            
            left_data = []
            for tenant_name, count in tenant_usage.items():
                # Replace "MARAE LIMITED" with "U士多"
                display_name = tenant_name
                if "MARAE LIMITED" in str(tenant_name):
                    display_name = str(tenant_name).replace("MARAE LIMITED", "U士多")
                left_data.append([display_name, count])
            
            # Right side: Merchant List (Accept Food Coupon)
            # Get ALL unique tenants from sales_category_mapping that exist in this market
            # We need to check all tenants that appear in the daily_df for this market
            all_tenants_in_market = set(daily_df[daily_df[market_col] == market]['ShopName'].dropna().unique())
            
            # Get tenants that have used coupons in this market
            tenants_with_usage_in_market = set(market_data['ShopName'].dropna().unique())
            
            # Build merchant list - include ALL tenants from sales_category_mapping
            merchant_rows = []
            for tenant in sorted(all_tenants_in_market):
                has_coupon = 'Y' if tenant in tenants_with_usage_in_market else 'N'
                
                # Replace "MARAE LIMITED" with "U士多"
                display_tenant = str(tenant)
                if "MARAE LIMITED" in display_tenant:
                    display_tenant = display_tenant.replace("MARAE LIMITED", "U士多")
                
                merchant_rows.append({
                    'Market': market,
                    'Tenant': display_tenant,
                    'Have Coupon Used': has_coupon
                })
            
            # Create the final dataframe
            # Combine left and right data side by side
            max_rows = max(len(left_data), len(merchant_rows)) + 2  # +2 for headers
            
            combined_data = []
            for i in range(max_rows):
                row = {}
                
                # Left columns - Data of Coupon Used
                if i == 0:
                    row['Data of Coupon Used'] = 'Data of Coupon Used'
                    row[''] = ''
                elif i == 1:
                    row['Data of Coupon Used'] = 'Tenant'
                    row[''] = 'Count'
                elif i - 2 < len(left_data):
                    row['Data of Coupon Used'] = left_data[i-2][0]
                    row[''] = left_data[i-2][1]
                else:
                    row['Data of Coupon Used'] = ''
                    row[''] = ''
                
                # Empty column for spacing
                row['  '] = ''
                
                # Right columns - Merchant List (Accept Food Coupon)
                if i == 0:
                    row['Merchant List (Accept Food Coupon)'] = 'Merchant List (Accept Food Coupon)'
                    row['   '] = ''
                    row['    '] = ''
                    row['No. of "N"'] = ''
                elif i == 1:
                    row['Merchant List (Accept Food Coupon)'] = 'Market'
                    row['   '] = 'Tenant'
                    row['    '] = 'Have Coupon Used'
                    n_count = sum(1 for m in merchant_rows if m['Have Coupon Used'] == 'N')
                    row['No. of "N"'] = n_count
                elif i - 2 < len(merchant_rows):
                    merchant = merchant_rows[i-2]
                    row['Merchant List (Accept Food Coupon)'] = merchant['Market']
                    row['   '] = merchant['Tenant']
                    row['    '] = merchant['Have Coupon Used']
                    row['No. of "N"'] = ''
                else:
                    row['Merchant List (Accept Food Coupon)'] = ''
                    row['   '] = ''
                    row['    '] = ''
                    row['No. of "N"'] = ''
                
                combined_data.append(row)
            
            market_df = pd.DataFrame(combined_data)
            
            # Clean market name for sheet name (max 31 chars, no special chars)
            sheet_name = str(market)[:31]
            sheet_name = sheet_name.replace('/', '_').replace('\\', '_').replace('*', '_').replace('[', '_').replace(']', '_').replace(':', '_').replace('?', '_')
            
            market_sheets[sheet_name] = market_df
            
        print(f"✓ Generated {len(market_sheets)} market worksheets")
        
    except Exception as e:
        print(f"Error generating market worksheets: {e}")
        import traceback
        traceback.print_exc()
    
    return market_sheets

def generate_summary_tables(weekly_df, daily_df):
    """Generate all summary tables"""
    tables_data = {}
    
    print(f"Weekly DF shape: {weekly_df.shape}, Daily DF shape: {daily_df.shape}")
    print(f"Weekly columns: {weekly_df.columns.tolist()}")
    print(f"Daily columns: {daily_df.columns.tolist()}")
    
    # These only need weekly data
    try:
        result = generate_coupon_purchased_table(weekly_df)
        if result:
            tables_data['coupon_purchased'] = result
            print("✓ Generated coupon_purchased table")
        else:
            print("✗ Failed to generate coupon_purchased table")
    except Exception as e:
        print(f"Error generating coupon_purchased: {e}")
        import traceback
        traceback.print_exc()
    
    try:
        result = generate_distributed_coupon_table(weekly_df)
        if result:
            tables_data['distributed_coupon'] = result
            print("✓ Generated distributed_coupon table")
        else:
            print("✗ Failed to generate distributed_coupon table")
    except Exception as e:
        print(f"Error generating distributed_coupon: {e}")
        import traceback
        traceback.print_exc()
    
    # These need both weekly and daily data
    try:
        result = generate_used_market_table(weekly_df, daily_df)
        if result:
            tables_data['used_market'] = result
            print(f"✓ Generated used_market table with keys: {result.keys()}")
        else:
            print("✗ Failed to generate used_market table")
    except Exception as e:
        print(f"Error generating used_market: {e}")
        import traceback
        traceback.print_exc()
    
    try:
        result = generate_top_merchant_table(weekly_df, daily_df)
        if result:
            tables_data['top_merchant'] = result
            print(f"✓ Generated top_merchant table with keys: {result.keys()}")
        else:
            print("✗ Failed to generate top_merchant table")
    except Exception as e:
        print(f"Error generating top_merchant: {e}")
        import traceback
        traceback.print_exc()
    
    try:
        result = generate_merchant_category_table(weekly_df, daily_df)
        if result:
            tables_data['merchant_category'] = result
            print(f"✓ Generated merchant_category table with keys: {result.keys()}")
        else:
            print("✗ Failed to generate merchant_category table")
    except Exception as e:
        print(f"Error generating merchant_category: {e}")
        import traceback
        traceback.print_exc()
    
    print(f"Final tables_data keys: {tables_data.keys()}")
    
    return tables_data

@app.route('/process-reports', methods=['POST'])
def process_reports():
    """Main report processing endpoint"""
    try:
        # Validation
        if 'market_reports' not in request.files or 'summary' not in request.files or 'sales_category' not in request.files:
            return jsonify({'error': 'All three files are required'}), 400
        
        market_files = request.files.getlist('market_reports')
        summary_file = request.files['summary']
        sales_category_file = request.files['sales_category']
        
        if not market_files or market_files[0].filename == '' or not summary_file.filename or not sales_category_file.filename:
            return jsonify({'error': 'All files must be selected'}), 400
        
        # Save and load sales category mapping
        sales_category_filename = secure_filename(sales_category_file.filename)
        sales_category_path = os.path.join(UPLOAD_FOLDER, f'temp_sales_category_{sales_category_filename}')
        sales_category_file.save(sales_category_path)
        
        try:
            sales_category_mapping = load_sales_category_mapping(sales_category_path)
        except Exception as e:
            if os.path.exists(sales_category_path):
                os.remove(sales_category_path)
            return jsonify({'error': f'Sales Category file error: {str(e)}'}), 400
        
        # Save and read summary file
        summary_filename = secure_filename(summary_file.filename)
        summary_path = os.path.join(UPLOAD_FOLDER, f'temp_summary_{summary_filename}')
        summary_file.save(summary_path)
        
        try:
            daily_finance_df = read_excel_sheet(summary_path, 'Daily Finance Report')
        except Exception as e:
            if os.path.exists(summary_path):
                os.remove(summary_path)
            if os.path.exists(sales_category_path):
                os.remove(sales_category_path)
            return jsonify({'error': f'Summary file error: {str(e)}'}), 400
        
        previous_daily_count = len(daily_finance_df)
        
        # Process market files
        all_market_data = []
        uploaded_market_files = []
        
        for file in market_files:
            if not file or not allowed_file(file.filename):
                continue
            
            filename = secure_filename(file.filename)
            filepath = os.path.join(UPLOAD_FOLDER, f'temp_market_{filename}')
            file.save(filepath)
            
            try:
                df = read_excel_file(filepath)
                df = add_sales_category(df, sales_category_mapping)
                all_market_data.append(df)
                uploaded_market_files.append(filename)
            except Exception as e:
                print(f"Error processing market file {filename}: {e}")
            finally:
                if os.path.exists(filepath):
                    os.remove(filepath)
        
        if not all_market_data:
            if os.path.exists(summary_path):
                os.remove(summary_path)
            if os.path.exists(sales_category_path):
                os.remove(sales_category_path)
            return jsonify({'error': 'No valid market reports processed'}), 400
        
        # Combine data
        combined_market_df = pd.concat(all_market_data, ignore_index=True)
        new_records_count = len(combined_market_df)
        
        # Append to daily finance report FIRST (before generating tables)
        final_daily_finance_df = pd.concat([daily_finance_df, combined_market_df], ignore_index=True)
        
        # Generate summary tables with BOTH dataframes
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        tables_data = generate_summary_tables(combined_market_df, final_daily_finance_df)
        
        # Save output Excel file with multiple sheets
        output_filename = f'combined_report_{timestamp}.xlsx'
        output_path = os.path.join(OUTPUT_FOLDER, output_filename)
        
        dfs_to_save = {
            'Daily Finance Report': final_daily_finance_df,
            'Weekly': combined_market_df
        }
        
        # Add Pivot worksheet
        pivot_df = generate_pivot_worksheet(final_daily_finance_df)
        if pivot_df is not None:
            dfs_to_save['Pivot'] = pivot_df
        
        # Add market worksheets
        dfs_to_save.update(generate_market_worksheets(final_daily_finance_df, sales_category_mapping))
        
        save_dataframes_to_excel(dfs_to_save, output_path)
        
        # Cleanup temp files
        if os.path.exists(summary_path):
            os.remove(summary_path)
        if os.path.exists(sales_category_path):
            os.remove(sales_category_path)
        
        return jsonify({
            'message': f'Successfully processed {len(uploaded_market_files)} market report(s)',
            'output_filename': output_filename,
            'file_path': output_filename,
            'total_records': len(final_daily_finance_df),
            'previous_records': previous_daily_count,
            'new_records': new_records_count,
            'market_files_processed': uploaded_market_files,
            'tables': tables_data
        }), 200
        
    except Exception as e:
        import traceback
        print(f"Server error: {traceback.format_exc()}")
        return jsonify({'error': f'Server error: {str(e)}'}), 500

@app.route('/download/<filename>', methods=['GET'])
def download_file(filename):
    try:
        filepath = os.path.join(OUTPUT_FOLDER, filename)
        if not os.path.exists(filepath):
            return jsonify({'error': 'File not found'}), 404
        
        return send_file(
            filepath,
            as_attachment=True,
            download_name=filename,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
        )
    except Exception as e:
        return jsonify({'error': str(e)}), 500

@app.route('/health', methods=['GET'])
def health_check():
    return jsonify({'status': 'running'}), 200

@app.route('/')
def index():
    return render_template('index.html')

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)