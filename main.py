import streamlit as st
import pandas as pd
import io
import csv
from datetime import datetime
import numpy as np
import traceback
import sys

# Page configuration
st.set_page_config(
    page_title="Finance Automation Platform",
    page_icon="💰",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS styling
st.markdown("""
<style>
    .main-header {
        font-size: 32px;
        font-weight: bold;
        color: #1f2937;
        margin-bottom: 10px;
    }
    .subtitle {
        font-size: 16px;
        color: #6b7280;
        margin-bottom: 30px;
    }
    .metric-card {
        background: white;
        padding: 20px;
        border-radius: 8px;
        box-shadow: 0 1px 3px rgba(0,0,0,0.1);
        text-align: center;
    }
    .metric-value {
        font-size: 32px;
        font-weight: bold;
        color: #1f2937;
    }
    .metric-label {
        font-size: 12px;
        color: #6b7280;
        margin-top: 5px;
    }
    .success-msg {
        color: #10b981;
        font-family: monospace;
        font-size: 12px;
        padding: 2px 0;
    }
    .error-msg {
        color: #dc2626;
        font-family: monospace;
        font-size: 12px;
        padding: 2px 0;
    }
    .info-msg {
        color: #3b82f6;
        font-family: monospace;
        font-size: 12px;
        padding: 2px 0;
    }
    .warning-msg {
        color: #f59e0b;
        font-family: monospace;
        font-size: 12px;
        padding: 2px 0;
    }
    .console-box {
        background: #1e1e1e;
        color: #d4d4d4;
        padding: 15px;
        border-radius: 8px;
        font-family: 'Courier New', monospace;
        font-size: 12px;
        max-height: 350px;
        overflow-y: auto;
        line-height: 1.6;
    }
    .error-console-box {
        background: #2d1818;
        color: #ff6b6b;
        padding: 15px;
        border-radius: 8px;
        font-family: 'Courier New', monospace;
        font-size: 11px;
        max-height: 300px;
        overflow-y: auto;
        border: 1px solid #ff4444;
        line-height: 1.5;
    }
    .stDownloadButton button {
        background-color: #10b981;
        color: white;
    }
    .upload-section {
        border: 2px dashed #d1d5db;
        border-radius: 8px;
        padding: 20px;
        text-align: center;
        background: #f9fafb;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'processed_data' not in st.session_state:
    st.session_state.processed_data = None
if 'uploaded_file_name' not in st.session_state:
    st.session_state.uploaded_file_name = None
if 'console_logs' not in st.session_state:
    st.session_state.console_logs = []
if 'error_logs' not in st.session_state:
    st.session_state.error_logs = []
if 'original_data' not in st.session_state:
    st.session_state.original_data = None
if 'script_executed' not in st.session_state:
    st.session_state.script_executed = False

def log_to_console(message, msg_type='info'):
    """Add messages to console log"""
    timestamp = datetime.now().strftime("%H:%M:%S")
    icon = {'info': '📝', 'success': '✅', 'error': '❌', 'warning': '⚠️'}.get(msg_type, 'ℹ️')
    st.session_state.console_logs.append(f"[{timestamp}] {icon} {message}")
    if len(st.session_state.console_logs) > 100:
        st.session_state.console_logs = st.session_state.console_logs[-100:]

def log_error(error_msg, traceback_str=None):
    """Add error to error console"""
    timestamp = datetime.now().strftime("%H:%M:%S")
    error_entry = f"[{timestamp}] ❌ ERROR: {error_msg}"
    if traceback_str:
        error_entry += f"\n\nTraceback:\n{traceback_str}"
    st.session_state.error_logs.append(error_entry)
    if len(st.session_state.error_logs) > 20:
        st.session_state.error_logs = st.session_state.error_logs[-20:]

def process_uploaded_file(uploaded_file):
    """Process uploaded CSV or Excel file"""
    try:
        log_to_console(f"Processing file: {uploaded_file.name}", 'info')
        
        if uploaded_file.name.endswith('.csv'):
            df = pd.read_csv(uploaded_file)
            log_to_console(f"✓ CSV file loaded successfully", 'success')
        elif uploaded_file.name.endswith(('.xlsx', '.xls')):
            df = pd.read_excel(uploaded_file)
            log_to_console(f"✓ Excel file loaded successfully", 'success')
        else:
            log_to_console(f"Unsupported file type: {uploaded_file.name}", 'error')
            log_error(f"Unsupported file type: {uploaded_file.name}")
            return None
        
        log_to_console(f"Rows: {len(df)}, Columns: {len(df.columns)}", 'info')
        log_to_console(f"Column names: {', '.join(df.columns.tolist()[:5])}{'...' if len(df.columns) > 5 else ''}", 'info')
        return df
    except Exception as e:
        error_msg = f"Error processing file: {str(e)}"
        log_to_console(error_msg, 'error')
        log_error(error_msg, traceback.format_exc())
        return None

def execute_python_script(df, script):
    """Execute Python script on dataframe with enhanced library support"""
    try:
        log_to_console("=" * 50, 'info')
        log_to_console("🚀 Starting script execution...", 'info')
        log_to_console("=" * 50, 'info')
        
        # Convert DataFrame to CSV string for script processing
        input_csv = df.to_csv(index=False)
        log_to_console(f"Input data prepared: {len(df)} rows", 'info')
        
        # Create execution environment with all necessary libraries
        exec_globals = {
            'input_csv': input_csv,
            'input_df': df.copy(),  # Provide dataframe directly too
            'io': io,
            'csv': csv,
            'pd': pd,
            'np': np,
            'datetime': datetime,
        }
        
        # Try to import optional visualization libraries
        try:
            import matplotlib
            matplotlib.use('Agg')  # Non-interactive backend
            import matplotlib.pyplot as plt
            exec_globals['plt'] = plt
            exec_globals['matplotlib'] = matplotlib
            log_to_console("✓ Matplotlib loaded", 'info')
        except ImportError:
            log_to_console("⚠️ Matplotlib not available", 'warning')
        
        try:
            import seaborn as sns
            exec_globals['sns'] = sns
            log_to_console("✓ Seaborn loaded", 'info')
        except ImportError:
            log_to_console("⚠️ Seaborn not available", 'warning')
        
        try:
            from sklearn import preprocessing, model_selection, metrics
            exec_globals['sklearn'] = __import__('sklearn')
            exec_globals['preprocessing'] = preprocessing
            exec_globals['model_selection'] = model_selection
            exec_globals['metrics'] = metrics
            log_to_console("✓ Scikit-learn loaded", 'info')
        except ImportError:
            log_to_console("⚠️ Scikit-learn not available", 'warning')
        
        log_to_console("Executing user script...", 'info')
        
        # Execute the script
        exec(script, exec_globals)
        
        log_to_console("✓ Script executed without errors", 'success')
        
        # Get output_csv or output_df from executed script
        if 'output_csv' in exec_globals:
            output_csv = exec_globals['output_csv']
            result_df = pd.read_csv(io.StringIO(output_csv))
            log_to_console("✓ Output retrieved via 'output_csv'", 'success')
            log_to_console(f"Output: {len(result_df)} rows, {len(result_df.columns)} columns", 'success')
            log_to_console("=" * 50, 'success')
            log_to_console("✅ EXECUTION COMPLETED SUCCESSFULLY", 'success')
            log_to_console("=" * 50, 'success')
            return result_df
        elif 'output_df' in exec_globals:
            result_df = exec_globals['output_df']
            log_to_console("✓ Output retrieved via 'output_df'", 'success')
            log_to_console(f"Output: {len(result_df)} rows, {len(result_df.columns)} columns", 'success')
            log_to_console("=" * 50, 'success')
            log_to_console("✅ EXECUTION COMPLETED SUCCESSFULLY", 'success')
            log_to_console("=" * 50, 'success')
            return result_df
        else:
            error_msg = "Script must set 'output_csv' or 'output_df' variable"
            log_to_console(f"❌ {error_msg}", 'error')
            log_error(error_msg)
            return None
            
    except Exception as e:
        error_msg = f"Script execution failed: {str(e)}"
        log_to_console("=" * 50, 'error')
        log_to_console(f"❌ {error_msg}", 'error')
        log_to_console("=" * 50, 'error')
        log_error(error_msg, traceback.format_exc())
        return None

def create_dynamic_chart(df):
    """Create interactive chart from dataframe"""
    if df is None or len(df) == 0:
        return None
    
    try:
        # Find numeric columns
        numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
        
        if not numeric_cols:
            return None
        
        # Try to find date/time column for x-axis
        date_cols = [col for col in df.columns if any(term in col.lower() for term in ['date', 'time', 'timestamp', 'day', 'month', 'year'])]
        
        if date_cols:
            try:
                x_data = pd.to_datetime(df[date_cols[0]], errors='coerce')
                if x_data.isna().all():
                    x_data = list(range(len(df)))
                    x_label = 'Row Index'
                else:
                    x_label = date_cols[0]
            except:
                x_data = list(range(len(df)))
                x_label = 'Row Index'
        else:
            x_data = list(range(len(df)))
            x_label = 'Row Index'
        
        # Create figure
        fig = go.Figure()
        
        # Plot up to 3 numeric columns
        colors = ['#f59e0b', '#10b981', '#3b82f6', '#dc2626', '#6366f1']
        plot_limit = min(len(df), 100)  # Show up to 100 points
        
        for idx, col in enumerate(numeric_cols[:3]):
            fig.add_trace(go.Scatter(
                x=x_data[:plot_limit],
                y=df[col][:plot_limit],
                mode='lines+markers',
                name=col.capitalize(),
                line=dict(color=colors[idx], width=2),
                marker=dict(size=6)
            ))
        
        fig.update_layout(
            title=f'Data Visualization ({plot_limit} points)',
            xaxis_title=x_label,
            yaxis_title='Value',
            height=400,
            hovermode='x unified',
            template='plotly_white',
            legend=dict(orientation="h", yanchor="bottom", y=1.02, xanchor="right", x=1)
        )
        
        return fig
    except Exception as e:
        log_error(f"Error creating chart: {str(e)}", traceback.format_exc())
        return None

def calculate_metrics(df):
    """Calculate dynamic metrics from dataframe"""
    if df is None or len(df) == 0:
        return None
    
    try:
        total_records = len(df)
        numeric_cols = df.select_dtypes(include=[np.number]).columns.tolist()
        
        if numeric_cols:
            first_col = numeric_cols[0]
            values = df[first_col].dropna()
            
            if len(values) > 0:
                avg_value = values.mean()
                max_value = values.max()
                min_value = values.min()
                std_value = values.std()
                
                return {
                    'metric1': f"{avg_value:.2f}",
                    'metric1_label': f'Avg {first_col}',
                    'metric2': f"{max_value:.2f}",
                    'metric2_label': f'Max {first_col}',
                    'metric3': f"{min_value:.2f}",
                    'metric3_label': f'Min {first_col}',
                    'metric4': f"{total_records:,}",
                    'metric4_label': 'Total Records'
                }
        
        return {
            'metric1': 'N/A',
            'metric1_label': 'Average',
            'metric2': 'N/A',
            'metric2_label': 'Maximum',
            'metric3': 'N/A',
            'metric3_label': 'Minimum',
            'metric4': f"{total_records:,}",
            'metric4_label': 'Total Records'
        }
    except Exception as e:
        log_error(f"Error calculating metrics: {str(e)}", traceback.format_exc())
        return None

# Main Layout
st.markdown('<p class="main-header">💰 Finance Automation Platform</p>', unsafe_allow_html=True)
st.markdown('<p class="subtitle">Upload CSV/Excel → Paste Python Script → Download Processed Output</p>', unsafe_allow_html=True)

# Create two columns
left_col, right_col = st.columns([1, 2])

with left_col:
    # Upload Section
    st.markdown("### 📤 Step 1: Upload Document")
    
    uploaded_file = st.file_uploader(
        "Choose CSV or Excel file",
        type=['csv', 'xlsx', 'xls'],
        help="Maximum file size: 200MB"
    )
    
    if uploaded_file:
        # Check if this is a new file
        if st.session_state.uploaded_file_name != uploaded_file.name:
            log_to_console(f"New file detected: {uploaded_file.name}", 'info')
            df = process_uploaded_file(uploaded_file)
            if df is not None:
                st.session_state.original_data = df
                st.session_state.uploaded_file_name = uploaded_file.name
                st.session_state.processed_data = None
                st.session_state.script_executed = False
    
    # Display uploaded file info
    if st.session_state.uploaded_file_name:
        st.success(f"✅ File loaded: **{st.session_state.uploaded_file_name}**")
        
        if st.session_state.original_data is not None:
            col1, col2 = st.columns(2)
            with col1:
                st.metric("Rows", len(st.session_state.original_data))
            with col2:
                st.metric("Columns", len(st.session_state.original_data.columns))
        
        if st.button("🗑️ Remove File & Reset", use_container_width=True):
            st.session_state.uploaded_file_name = None
            st.session_state.processed_data = None
            st.session_state.original_data = None
            st.session_state.console_logs = []
            st.session_state.error_logs = []
            st.session_state.script_executed = False
            log_to_console("Application reset", 'info')
            st.rerun()
    
    st.markdown("---")
    
    # Script Section
    st.markdown("### 🐍 Step 2: Paste Python Script")
    st.caption("Available libraries: pandas (pd), numpy (np), matplotlib (plt), seaborn (sns), sklearn")
    
    default_script = """# Example 1: Using CSV string (input_csv)
import io, csv

reader = csv.DictReader(io.StringIO(input_csv))
rows = []
for r in reader:
    # Add your processing logic here
    amt = float(r.get('amount', 0) or 0)
    r['amount_plus_tax'] = f"{amt * 1.2:.2f}"
    rows.append(r)

headers = list(rows[0].keys()) if rows else []
buf = io.StringIO()
writer = csv.writer(buf)
writer.writerow(headers)
for r in rows:
    writer.writerow([r.get(h, '') for h in headers])

output_csv = buf.getvalue()

# Example 2: Using DataFrame (input_df)
# output_df = input_df.copy()
# output_df['new_column'] = output_df['column_name'] * 2
"""
    
    code_editor = st.text_area(
        "Python Code Editor",
        value=default_script,
        height=320,
        help="Your script must define 'output_csv' (string) or 'output_df' (DataFrame)"
    )
    
    col1, col2 = st.columns(2)
    with col1:
        run_disabled = st.session_state.original_data is None
        if st.button("🚀 Run Script", type="primary", use_container_width=True, disabled=run_disabled):
            if st.session_state.original_data is not None:
                st.session_state.console_logs = []  # Clear previous logs
                st.session_state.error_logs = []
                result = execute_python_script(st.session_state.original_data, code_editor)
                if result is not None:
                    st.session_state.processed_data = result
                    st.session_state.script_executed = True
                    st.rerun()
            else:
                st.error("Please upload a file first")
    
    with col2:
        if st.button("🔄 Clear Logs", use_container_width=True):
            st.session_state.console_logs = []
            st.session_state.error_logs = []
            log_to_console("Logs cleared", 'info')
            st.rerun()
    
    st.markdown("---")
    
    # Download Section
    st.markdown("### 📥 Step 3: Download Output")
    
    if st.session_state.processed_data is not None:
        col1, col2 = st.columns(2)
        
        with col1:
            csv_data = st.session_state.processed_data.to_csv(index=False)
            st.download_button(
                label="📄 Download CSV",
                data=csv_data,
                file_name=f"output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
                mime="text/csv",
                use_container_width=True
            )
        
        with col2:
            # Create Excel file
            buffer = io.BytesIO()
            with pd.ExcelWriter(buffer, engine='openpyxl') as writer:
                st.session_state.processed_data.to_excel(writer, index=False, sheet_name='Output')
            excel_data = buffer.getvalue()
            
            st.download_button(
                label="📊 Download Excel",
                data=excel_data,
                file_name=f"output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        
        st.success("✅ Output ready for download!")
    else:
        st.info("💡 Run the script to generate downloadable output")

with right_col:
    # Progress/Error Console (as requested by Mia)
    st.markdown("### 📟 Progress & Error Console")
    
    tab1, tab2 = st.tabs(["📋 Execution Log", "🐛 Error Details"])
    
    with tab1:
        if st.session_state.console_logs:
            console_html = '<div class="console-box">'
            for log in st.session_state.console_logs[-30:]:  # Show last 30 logs
                console_html += f"{log}<br>"
            console_html += '</div>'
            st.markdown(console_html, unsafe_allow_html=True)
        else:
            st.info("📝 Execution logs will appear here. Upload a file and run a script to get started.")
    
    with tab2:
        if st.session_state.error_logs:
            error_html = '<div class="error-console-box">'
            for error in st.session_state.error_logs[-5:]:
                error_html += f"{error}<br><br>{'=' * 60}<br><br>"
            error_html += '</div>'
            st.markdown(error_html, unsafe_allow_html=True)
        else:
            st.success("✅ No errors detected")
    
    st.markdown("---")
    
    # Chart Section
    st.markdown("### 📊 Data Visualization")
    
    if st.session_state.processed_data is not None:
        chart = create_dynamic_chart(st.session_state.processed_data)
        if chart:
            st.plotly_chart(chart, use_container_width=True)
        else:
            st.info("📈 No numeric columns found for visualization")
    elif st.session_state.original_data is not None:
        st.info("📤 Run the script to visualize processed data")
    else:
        st.info("📤 Upload a file to see data visualization")
    
    st.markdown("---")
    
    # Data Preview
    st.markdown("### 📋 Data Preview")
    
    if st.session_state.processed_data is not None:
        st.markdown("**Processed Data Output:**")
        preview_rows = st.slider("Preview rows:", 3, 20, 5, key="preview_slider")
        st.dataframe(st.session_state.processed_data.head(preview_rows), use_container_width=True)
        
        # Data summary
        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("Total Rows", len(st.session_state.processed_data))
        with col2:
            st.metric("Total Columns", len(st.session_state.processed_data.columns))
        with col3:
            memory_kb = st.session_state.processed_data.memory_usage(deep=True).sum() / 1024
            st.metric("Memory", f"{memory_kb:.1f} KB")
    elif st.session_state.original_data is not None:
        st.markdown("**Original Data (Input):**")
        st.dataframe(st.session_state.original_data.head(5), use_container_width=True)
    else:
        st.info("📋 Upload a file to preview data")
    
    st.markdown("---")
    
    # Metrics Section
    st.markdown("### 📈 Key Metrics")
    
    metrics = calculate_metrics(st.session_state.processed_data)
    
    if metrics:
        col1, col2, col3, col4 = st.columns(4)
        
        with col1:
            st.markdown('<div class="metric-card">', unsafe_allow_html=True)
            st.markdown(f'<div class="metric-value">{metrics["metric1"]}</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="metric-label">{metrics["metric1_label"]}</div>', unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
        
        with col2:
            st.markdown('<div class="metric-card">', unsafe_allow_html=True)
            st.markdown(f'<div class="metric-value">{metrics["metric2"]}</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="metric-label">{metrics["metric2_label"]}</div>', unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
        
        with col3:
            st.markdown('<div class="metric-card">', unsafe_allow_html=True)
            st.markdown(f'<div class="metric-value">{metrics["metric3"]}</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="metric-label">{metrics["metric3_label"]}</div>', unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
        
        with col4:
            st.markdown('<div class="metric-card">', unsafe_allow_html=True)
            st.markdown(f'<div class="metric-value">{metrics["metric4"]}</div>', unsafe_allow_html=True)
            st.markdown(f'<div class="metric-label">{metrics["metric4_label"]}</div>', unsafe_allow_html=True)
            st.markdown('</div>', unsafe_allow_html=True)
    else:
        st.info("📊 Metrics will appear after processing data")

# Footer
st.markdown("---")
col1, col2 = st.columns([3, 1])
with col1:
    st.caption("Finance Automation Platform v1.0 • Built with Streamlit")
with col2:
    if st.session_state.script_executed:
        st.success("✅ Ready")
    else:

        st.info("⏳ Waiting")
