import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import io
import csv
from datetime import datetime
import numpy as np
import traceback
import sys

# Page configuration
st.set_page_config(
    page_title="Finance Automation",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS styling
st.markdown("""
<style>
    /* Hide Streamlit branding and icons */
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
    .main-header {
        font-size: 96px;
        font-weight: 900;
        color: #1f2937;
        margin-bottom: 30px;
        letter-spacing: -1.5px;
    }
    
    /* Grey background for left panel */
    [data-testid="column"]:first-child {
        background-color: #f3f4f6;
        padding: 25px;
        border-radius: 8px;
    }
    
    .section-header {
        font-size: 18px;
        font-weight: 600;
        color: #1f2937;
        margin-bottom: 15px;
        margin-top: 20px;
    }
    
    .console-box {
        background: #1e1e1e;
        color: #d4d4d4;
        padding: 15px;
        border-radius: 8px;
        font-family: 'Courier New', monospace;
        font-size: 12px;
        max-height: 400px;
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
        max-height: 400px;
        overflow-y: auto;
        border: 1px solid #ff4444;
        line-height: 1.5;
    }
    
    .stDownloadButton button {
        background-color: #10b981;
        color: white;
    }
    
    /* Tab styling to match Upload Document size */
    .stTabs [data-baseweb="tab-list"] {
        gap: 2px;
    }
    
    .stTabs [data-baseweb="tab"] {
        font-size: 18px;
        font-weight: 600;
        padding: 12px 24px;
    }
    
    /* Hide file uploader drag and drop text */
    [data-testid="stFileUploader"] section div {
        display: none;
    }
    
    [data-testid="stFileUploader"] section {
        padding: 0;
    }
    
    /* Remove extra padding from file uploader */
    .uploadedFile {
        margin-top: 10px;
    }
    
    /* Center empty state */
    .empty-state {
        display: flex;
        align-items: center;
        justify-content: center;
        height: 400px;
        color: #6b7280;
        font-size: 16px;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'processed_data' not in st.session_state:
    st.session_state.processed_data = None
if 'uploaded_files' not in st.session_state:
    st.session_state.uploaded_files = []
if 'console_logs' not in st.session_state:
    st.session_state.console_logs = []
if 'error_logs' not in st.session_state:
    st.session_state.error_logs = []
if 'original_data' not in st.session_state:
    st.session_state.original_data = None
if 'script_executed' not in st.session_state:
    st.session_state.script_executed = False
if 'combined_data' not in st.session_state:
    st.session_state.combined_data = None
if 'file_uploader_key' not in st.session_state:
    st.session_state.file_uploader_key = 0

def log_to_console(message, msg_type='info'):
    """Add messages to console log"""
    timestamp = datetime.now().strftime("%H:%M:%S")
    icon = {'info': 'INFO', 'success': 'SUCCESS', 'error': 'ERROR', 'warning': 'WARNING'}.get(msg_type, 'INFO')
    st.session_state.console_logs.append(f"[{timestamp}] [{icon}] {message}")
    if len(st.session_state.console_logs) > 50:  # Reduced from 100
        st.session_state.console_logs = st.session_state.console_logs[-50:]

def log_error(error_msg, traceback_str=None):
    """Add error to error console"""
    timestamp = datetime.now().strftime("%H:%M:%S")
    error_entry = f"[{timestamp}] ERROR: {error_msg}"
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
            # Try reading CSV with different encodings
            try:
                df = pd.read_csv(uploaded_file, encoding='utf-8')
            except UnicodeDecodeError:
                uploaded_file.seek(0)  # Reset file pointer
                df = pd.read_csv(uploaded_file, encoding='latin-1')
            log_to_console(f"CSV file loaded successfully", 'success')
        elif uploaded_file.name.endswith(('.xlsx', '.xls')):
            # Try reading Excel with openpyxl engine (more robust)
            try:
                df = pd.read_excel(uploaded_file, engine='openpyxl')
                log_to_console(f"Excel file loaded successfully (openpyxl)", 'success')
            except Exception as e1:
                log_to_console(f"Trying alternative Excel reader...", 'warning')
                uploaded_file.seek(0)  # Reset file pointer
                try:
                    # Try with xlrd for older .xls files
                    df = pd.read_excel(uploaded_file, engine='xlrd')
                    log_to_console(f"Excel file loaded successfully (xlrd)", 'success')
                except Exception as e2:
                    # Last resort - try without specifying engine
                    uploaded_file.seek(0)
                    df = pd.read_excel(uploaded_file)
                    log_to_console(f"Excel file loaded successfully", 'success')
        else:
            log_to_console(f"Unsupported file type: {uploaded_file.name}", 'error')
            log_error(f"Unsupported file type: {uploaded_file.name}")
            return None
        
        # Clean column names (strip whitespace)
        df.columns = df.columns.str.strip()
        
        log_to_console(f"Rows: {len(df)}, Columns: {len(df.columns)}", 'info')
        log_to_console(f"Column names: {', '.join(df.columns.tolist()[:5])}{'...' if len(df.columns) > 5 else ''}", 'info')
        return df
    except Exception as e:
        error_msg = f"Error processing file: {str(e)}"
        log_to_console(error_msg, 'error')
        log_error(error_msg, traceback.format_exc())
        return None

def combine_uploaded_files(uploaded_files):
    """Combine multiple uploaded files into one dataframe"""
    try:
        if not uploaded_files:
            return None
        
        all_dfs = []
        for uploaded_file in uploaded_files:
            df = process_uploaded_file(uploaded_file)
            if df is not None:
                all_dfs.append(df)
        
        if not all_dfs:
            return None
        
        if len(all_dfs) == 1:
            combined_df = all_dfs[0]
        else:
            # Concatenate all dataframes
            combined_df = pd.concat(all_dfs, ignore_index=True)
            log_to_console(f"Combined {len(all_dfs)} files into one dataset", 'success')
        
        log_to_console(f"Total rows: {len(combined_df)}, Total columns: {len(combined_df.columns)}", 'info')
        return combined_df
    except Exception as e:
        error_msg = f"Error combining files: {str(e)}"
        log_to_console(error_msg, 'error')
        log_error(error_msg, traceback.format_exc())
        return None

def execute_python_script(df, script):
    """Execute Python script on dataframe with enhanced library support"""
    try:
        log_to_console("=" * 50, 'info')
        log_to_console("Starting script execution...", 'info')
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
            log_to_console("Matplotlib loaded", 'info')
        except ImportError:
            log_to_console("Matplotlib not available", 'warning')
        
        try:
            import seaborn as sns
            exec_globals['sns'] = sns
            log_to_console("Seaborn loaded", 'info')
        except ImportError:
            log_to_console("Seaborn not available", 'warning')
        
        try:
            from sklearn import preprocessing, model_selection, metrics
            exec_globals['sklearn'] = __import__('sklearn')
            exec_globals['preprocessing'] = preprocessing
            exec_globals['model_selection'] = model_selection
            exec_globals['metrics'] = metrics
            log_to_console("Scikit-learn loaded", 'info')
        except ImportError:
            log_to_console("Scikit-learn not available", 'warning')
        
        log_to_console("Executing user script...", 'info')
        
        # Execute the script
        exec(script, exec_globals)
        
        log_to_console("Script executed without errors", 'success')
        
        # Get output_csv or output_df from executed script
        if 'output_csv' in exec_globals:
            output_csv = exec_globals['output_csv']
            result_df = pd.read_csv(io.StringIO(output_csv))
            log_to_console("Output retrieved via 'output_csv'", 'success')
            log_to_console(f"Output: {len(result_df)} rows, {len(result_df.columns)} columns", 'success')
            log_to_console("=" * 50, 'success')
            log_to_console("EXECUTION COMPLETED SUCCESSFULLY", 'success')
            log_to_console("=" * 50, 'success')
            return result_df
        elif 'output_df' in exec_globals:
            result_df = exec_globals['output_df']
            log_to_console("Output retrieved via 'output_df'", 'success')
            log_to_console(f"Output: {len(result_df)} rows, {len(result_df.columns)} columns", 'success')
            log_to_console("=" * 50, 'success')
            log_to_console("EXECUTION COMPLETED SUCCESSFULLY", 'success')
            log_to_console("=" * 50, 'success')
            return result_df
        else:
            error_msg = "Script must set 'output_csv' or 'output_df' variable"
            log_to_console(f"{error_msg}", 'error')
            log_error(error_msg)
            return None
            
    except Exception as e:
        error_msg = f"Script execution failed: {str(e)}"
        log_to_console("=" * 50, 'error')
        log_to_console(f"{error_msg}", 'error')
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
            title=f'Output Data Visualization',
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

# Main Layout
st.title("Finance Automation")
st.markdown("<br>", unsafe_allow_html=True)

# Create two columns
left_col, right_col = st.columns([1, 2])

with left_col:
    # Upload Section
    st.markdown('<p class="section-header">Upload Files</p>', unsafe_allow_html=True)
    
    uploaded_files = st.file_uploader(
        "Upload Files",
        type=['csv', 'xlsx', 'xls'],
        accept_multiple_files=True,
        help="Maximum file size: 200MB per file",
        label_visibility="collapsed",
        key=f"file_uploader_{st.session_state.file_uploader_key}"
    )
    
    if uploaded_files:
        st.success(f"{len(uploaded_files)} file(s) uploaded")
        for file in uploaded_files:
            st.text(f"• {file.name}")
        
        # Process files
        combined_df = combine_uploaded_files(uploaded_files)
        if combined_df is not None:
            st.session_state.combined_data = combined_df
            st.session_state.uploaded_files = uploaded_files
            
            col1, col2 = st.columns(2)
            with col1:
                st.metric("Total Rows", len(st.session_state.combined_data))
            with col2:
                st.metric("Total Columns", len(st.session_state.combined_data.columns))
        
        if st.button("Remove All Files", use_container_width=True):
            # Reset all session state
            st.session_state.uploaded_files = []
            st.session_state.processed_data = None
            st.session_state.combined_data = None
            st.session_state.console_logs = []
            st.session_state.error_logs = []
            st.session_state.script_executed = False
            # Increment key to force file uploader to reset
            st.session_state.file_uploader_key += 1
            log_to_console("All files removed and application reset", 'info')
            st.rerun()
    
    st.markdown("---")
    
    # Script Section
    st.markdown('<p class="section-header">Paste Python Script</p>', unsafe_allow_html=True)
    
    code_editor = st.text_area(
        "Python Script",
        value="",
        height=300,
        placeholder="# Paste your Python script here\n# Your uploaded file is already loaded - use input_df\n# Example:\n#   df = input_df.copy()\n#   # ... process df ...\n#   output_df = df  # Required!\n\n# Available variables:\n#   - input_df (DataFrame) - Your uploaded data\n#   - input_csv (string) - CSV format\n# Required output:\n#   - output_df (DataFrame) or output_csv (string)",
        help="Your script must define 'output_csv' (string) or 'output_df' (DataFrame)",
        label_visibility="collapsed"
    )
    
    col1, col2 = st.columns(2)
    with col1:
        run_disabled = st.session_state.combined_data is None or not code_editor.strip()
        
        if st.button("Run Script", type="primary", use_container_width=True, disabled=run_disabled):
            if st.session_state.combined_data is not None and code_editor.strip():
                # Clear previous execution data
                st.session_state.console_logs = []
                st.session_state.error_logs = []
                st.session_state.processed_data = None
                st.session_state.script_executed = False
                
                result = execute_python_script(st.session_state.combined_data, code_editor)
                if result is not None:
                    st.session_state.processed_data = result
                    st.session_state.script_executed = True
                st.rerun()
            else:
                if st.session_state.combined_data is None:
                    st.error("Please upload file(s) first")
                else:
                    st.error("Please paste a Python script")
    
    with col2:
        if st.button("Clear Logs", use_container_width=True):
            st.session_state.console_logs = []
            st.session_state.error_logs = []
            log_to_console("Logs cleared", 'info')
            st.rerun()
    
    st.markdown("---")
    
    # Download Section
    st.markdown('<p class="section-header">Download Output</p>', unsafe_allow_html=True)
    
    if st.session_state.processed_data is not None:
        col1, col2 = st.columns(2)
        
        with col1:
            csv_data = st.session_state.processed_data.to_csv(index=False)
            st.download_button(
                label="Download CSV",
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
                label="Download Excel",
                data=excel_data,
                file_name=f"output_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )
        
        st.success("Output ready for download!")
    else:
        st.info("Run the script to generate output")

with right_col:
    # Tabs at the top - same size as section headers
    if st.session_state.script_executed and st.session_state.processed_data is not None:
        tab1, tab2, tab3 = st.tabs(["Output", "Execution Log", "Error Details"])
        
        with tab1:
            st.markdown("### Output Visualization")
            
            # Chart
            chart = create_dynamic_chart(st.session_state.processed_data)
            if chart:
                st.plotly_chart(chart, use_container_width=True)
            else:
                st.info("No numeric columns found for visualization")
            
            st.markdown("---")
            
            # Data Preview
            st.markdown("### Processed Data Preview")
            preview_rows = st.slider("Preview rows:", 5, 50, 10, key="preview_slider")
            st.dataframe(st.session_state.processed_data.head(preview_rows), use_container_width=True)
            
            # Summary metrics
            col1, col2, col3 = st.columns(3)
            with col1:
                st.metric("Total Rows", len(st.session_state.processed_data))
            with col2:
                st.metric("Total Columns", len(st.session_state.processed_data.columns))
            with col3:
                memory_kb = st.session_state.processed_data.memory_usage(deep=True).sum() / 1024
                st.metric("Memory", f"{memory_kb:.1f} KB")
        
        with tab2:
            if st.session_state.console_logs:
                console_html = '<div class="console-box">'
                for log in st.session_state.console_logs:
                    console_html += f"{log}<br>"
                console_html += '</div>'
                st.markdown(console_html, unsafe_allow_html=True)
            else:
                st.info("No execution logs available")
        
        with tab3:
            if st.session_state.error_logs:
                error_html = '<div class="error-console-box">'
                for error in st.session_state.error_logs:
                    error_html += f"{error}<br><br>{'=' * 60}<br><br>"
                error_html += '</div>'
                st.markdown(error_html, unsafe_allow_html=True)
            else:
                st.success("No errors detected")
    else:
        # Empty state before script execution
        st.markdown('<div class="empty-state">Upload files and run a script to see output visualization</div>', unsafe_allow_html=True)






