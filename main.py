import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import plotly.express as px
import matplotlib.pyplot as plt
import io
import csv
from datetime import datetime
import numpy as np
import traceback
import warnings
import sys
import base64
import zipfile
from PIL import Image

warnings.filterwarnings('ignore')

# Page configuration
st.set_page_config(
    page_title="Finance Automation",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# Custom CSS styling
st.markdown("""
<style>
    #MainMenu {visibility: hidden;}
    footer {visibility: hidden;}
    header {visibility: hidden;}
    
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
        white-space: pre-wrap;
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
        white-space: pre-wrap;
    }
    
    .stDownloadButton button {
        background-color: #10b981;
        color: white;
    }
    
    .footer-text {
        text-align: center;
        color: #6b7280;
        font-size: 13px;
        padding: 20px 0 10px 0;
        margin-top: 30px;
        border-top: 1px solid #e5e7eb;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state variables
def initialize_session_state():
    """Initialize all session state variables with default values."""
    defaults = {
        'processed_data': None,
        'captured_figures': [],
        'console_logs': [],
        'error_logs': [],
        'script_executed': False,
        'input_files': [],
        'file_uploader_key': 0,
        'current_file_names': []
    }
    
    for key, value in defaults.items():
        if key not in st.session_state:
            st.session_state[key] = value

initialize_session_state()

# Logging functions
def log_to_console(message, msg_type='info'):
    """Add timestamped message to console logs."""
    timestamp = datetime.now().strftime("%H:%M:%S")
    icons = {'info': '[INFO]', 'success': '[SUCCESS]', 'error': '[ERROR]', 'warning': '[WARNING]'}
    icon = icons.get(msg_type, '[INFO]')
    st.session_state.console_logs.append(f"[{timestamp}] {icon} {message}")
    
    # Keep only last 100 logs
    if len(st.session_state.console_logs) > 100:
        st.session_state.console_logs = st.session_state.console_logs[-100:]

def log_error(error_msg, traceback_str=None):
    """Add timestamped error message with optional traceback to error logs."""
    timestamp = datetime.now().strftime("%H:%M:%S")
    error_entry = f"[{timestamp}] ERROR: {error_msg}"
    if traceback_str:
        error_entry += f"\n\nTraceback:\n{traceback_str}"
    st.session_state.error_logs.append(error_entry)
    
    # Keep only last 20 errors
    if len(st.session_state.error_logs) > 20:
        st.session_state.error_logs = st.session_state.error_logs[-20:]

# File processing
@st.cache_data(show_spinner=False)
def process_uploaded_file(file_bytes, file_name):
    """Process a single uploaded CSV or Excel file into a DataFrame."""
    try:
        if file_name.endswith('.csv'):
            try:
                df = pd.read_csv(io.BytesIO(file_bytes), encoding='utf-8')
            except UnicodeDecodeError:
                df = pd.read_csv(io.BytesIO(file_bytes), encoding='latin-1')
        elif file_name.endswith(('.xlsx', '.xls')):
            try:
                df = pd.read_excel(io.BytesIO(file_bytes), engine='openpyxl')
            except Exception:
                try:
                    df = pd.read_excel(io.BytesIO(file_bytes), engine='xlrd')
                except Exception:
                    df = pd.read_excel(io.BytesIO(file_bytes))
        else:
            return None
        
        # Clean column names
        df.columns = df.columns.str.strip()
        return df
    except Exception as e:
        log_error(f"Failed to process {file_name}: {str(e)}")
        return None

def process_multiple_files(uploaded_files):
    """Process multiple uploaded files and return list of DataFrames with metadata."""
    try:
        if not uploaded_files:
            return []
        
        log_to_console(f"Processing {len(uploaded_files)} file(s)...", 'info')
        all_dfs = []
        
        for uploaded_file in uploaded_files:
            file_bytes = uploaded_file.read()
            df = process_uploaded_file(file_bytes, uploaded_file.name)
            if df is not None:
                all_dfs.append({'name': uploaded_file.name, 'data': df})
                log_to_console(f"Loaded: {uploaded_file.name} ({len(df)} rows, {len(df.columns)} cols)", 'success')
        
        return all_dfs
    except Exception as e:
        error_msg = f"Error processing files: {str(e)}"
        log_to_console(error_msg, 'error')
        log_error(error_msg, traceback.format_exc())
        return []

# Capture utilities
class PrintCapture:
    """Redirect print statements to console log."""
    def __init__(self, log_func):
        self.log_func = log_func
        
    def write(self, text):
        if text.strip():
            self.log_func(text.strip(), 'info')
    
    def flush(self):
        pass

class FigureCapture:
    """Capture matplotlib and plotly figures during script execution."""
    def __init__(self):
        self.figures = []
    
    def capture_matplotlib(self):
        """Capture current matplotlib figure as base64 PNG."""
        try:
            fig = plt.gcf()
            if fig.get_axes():
                buf = io.BytesIO()
                fig.savefig(buf, format='png', dpi=150, bbox_inches='tight')
                buf.seek(0)
                img_str = base64.b64encode(buf.read()).decode()
                self.figures.append(('matplotlib', img_str))
                log_to_console(f"Captured matplotlib figure", 'success')
                plt.close(fig)
        except Exception as e:
            log_error(f"Error capturing matplotlib figure: {str(e)}")
    
    def capture_plotly(self, fig):
        """Capture plotly figure object."""
        try:
            self.figures.append(('plotly', fig))
            log_to_console(f"Captured plotly figure", 'success')
        except Exception as e:
            log_error(f"Error capturing plotly figure: {str(e)}")

def create_charts_zip(figures):
    """Create a ZIP file containing all captured charts as PNG images."""
    try:
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
            for idx, (fig_type, fig_data) in enumerate(figures):
                try:
                    if fig_type == 'matplotlib':
                        # Matplotlib figure is already base64 encoded PNG
                        img_bytes = base64.b64decode(fig_data)
                        zip_file.writestr(f"chart_{idx + 1}.png", img_bytes)
                        log_to_console(f"Added matplotlib chart {idx + 1} to ZIP", 'success')
                    elif fig_type == 'plotly':
                        # Convert plotly figure to PNG using kaleido
                        try:
                            img_bytes = fig_data.to_image(format='png', width=1200, height=800, engine='kaleido')
                            zip_file.writestr(f"chart_{idx + 1}.png", img_bytes)
                            log_to_console(f"Added plotly chart {idx + 1} to ZIP", 'success')
                        except Exception as plotly_error:
                            # Fallback: save as HTML if PNG export fails
                            html_str = fig_data.to_html(include_plotlyjs='cdn')
                            zip_file.writestr(f"chart_{idx + 1}.html", html_str.encode('utf-8'))
                            log_to_console(f"Added plotly chart {idx + 1} as HTML (PNG export unavailable)", 'warning')
                except Exception as chart_error:
                    log_error(f"Error adding chart {idx + 1} to ZIP: {str(chart_error)}")
                    continue
        
        zip_buffer.seek(0)
        zip_data = zip_buffer.getvalue()
        
        if len(zip_data) > 0:
            log_to_console(f"Charts ZIP created successfully ({len(zip_data)} bytes)", 'success')
            return zip_data
        else:
            log_error("ZIP file is empty - no charts were added")
            return None
            
    except Exception as e:
        log_error(f"Error creating charts ZIP: {str(e)}", traceback.format_exc())
        return None

# Script execution
def execute_python_script(input_files_list, script):
    """Execute user's Python script with input files and capture outputs."""
    try:
        log_to_console("=" * 60, 'info')
        log_to_console("Starting script execution...", 'info')
        log_to_console(f"Input files: {len(input_files_list)}", 'info')
        for idx, file_info in enumerate(input_files_list):
            log_to_console(f"   File {idx}: {file_info['name']} ({len(file_info['data'])} rows)", 'info')
        log_to_console("=" * 60, 'info')
        
        fig_capture = FigureCapture()
        
        # Custom show functions to capture visualizations
        def custom_plt_show(*args, **kwargs):
            fig_capture.capture_matplotlib()
        
        original_plotly_show = go.Figure.show
        def custom_plotly_show(self, *args, **kwargs):
            fig_capture.capture_plotly(self)
        
        # Prepare input variables for script
        input_files = [file_info['data'].copy() for file_info in input_files_list]
        input_df = input_files[0] if input_files else pd.DataFrame()
        
        # Create execution environment
        exec_globals = {
            'input_files': input_files,
            'input_df': input_df,
            'pd': pd,
            'np': np,
            'datetime': datetime,
            'io': io,
            'csv': csv,
            'plt': plt,
            'go': go,
            'px': px,
            'matplotlib': plt,
        }
        
        exec_locals = {}
        
        # Redirect stdout and patch visualization functions
        old_stdout = sys.stdout
        sys.stdout = PrintCapture(log_to_console)
        
        original_plt_show = plt.show
        plt.show = custom_plt_show
        go.Figure.show = custom_plotly_show
        
        try:
            # Execute the script
            exec(script, exec_globals, exec_locals)
            
            # Capture any remaining matplotlib figures
            if plt.get_fignums():
                fig_capture.capture_matplotlib()
            
        finally:
            # Restore original functions
            sys.stdout = old_stdout
            plt.show = original_plt_show
            go.Figure.show = original_plotly_show
        
        # Extract output_df if defined
        output_df = None
        if 'output_df' in exec_locals:
            output_df = exec_locals['output_df']
        elif 'output_df' in exec_globals:
            output_df = exec_globals['output_df']
        
        # Validate output
        if output_df is None:
            log_to_console("No output_df defined - showing visualizations only", 'warning')
            if not fig_capture.figures:
                error_msg = "Script must define 'output_df' or create visualizations"
                log_to_console(error_msg, 'error')
                log_error(error_msg, "No output_df or visualizations found")
                return None, []
        
        if output_df is not None and not isinstance(output_df, pd.DataFrame):
            error_msg = f"output_df must be a DataFrame, got {type(output_df)}"
            log_to_console(error_msg, 'error')
            log_error(error_msg)
            return None, []
        
        # Log success
        log_to_console("=" * 60, 'success')
        if output_df is not None:
            log_to_console(f"Output data: {len(output_df)} rows × {len(output_df.columns)} columns", 'success')
        if fig_capture.figures:
            log_to_console(f"Captured {len(fig_capture.figures)} visualization(s)", 'success')
        log_to_console("=" * 60, 'success')
        
        return output_df, fig_capture.figures
            
    except SyntaxError as e:
        error_msg = f"Syntax Error on line {e.lineno}: {str(e)}"
        log_to_console(error_msg, 'error')
        log_error(error_msg, traceback.format_exc())
        return None, []
    except NameError as e:
        error_msg = f"Name Error: {str(e)}"
        log_to_console(error_msg, 'error')
        log_error(error_msg, traceback.format_exc())
        return None, []
    except Exception as e:
        error_msg = f"Execution failed: {str(e)}"
        log_to_console(error_msg, 'error')
        log_error(error_msg, traceback.format_exc())
        return None, []
    finally:
        sys.stdout = old_stdout
        plt.close('all')

def safe_reset_session():
    """Safely reset all session state variables without iteration errors."""
    keys_to_reset = {
        'processed_data': None,
        'captured_figures': [],
        'console_logs': [],
        'error_logs': [],
        'script_executed': False,
        'input_files': [],
        'current_file_names': []
    }
    
    for key, default_value in keys_to_reset.items():
        st.session_state[key] = default_value
    
    st.session_state.file_uploader_key += 1

# ===== MAIN UI =====

st.title("Finance Automation")
st.markdown("<br>", unsafe_allow_html=True)

left_col, right_col = st.columns([1, 2])

# LEFT COLUMN: File upload, script editor, download
with left_col:
    # File Upload Section
    st.markdown('<p class="section-header">Upload Files</p>', unsafe_allow_html=True)
    
    uploaded_files = st.file_uploader(
        "Upload Files",
        type=['csv', 'xlsx', 'xls'],
        accept_multiple_files=True,
        help="Upload one or more CSV/Excel files. Maximum file size: 200MB per file",
        label_visibility="collapsed",
        key=f"file_uploader_{st.session_state.file_uploader_key}"
    )
    
    if uploaded_files:
        file_names = [f.name for f in uploaded_files]
        current_files = st.session_state.get('current_file_names', [])
        
        # Process files only if they changed
        if file_names != current_files:
            st.session_state.current_file_names = file_names
            processed_files = process_multiple_files(uploaded_files)
            if processed_files:
                st.session_state.input_files = processed_files
        
        st.success(f"{len(uploaded_files)} file(s) uploaded")
        for idx, file in enumerate(uploaded_files):
            if idx < len(st.session_state.input_files):
                file_data = st.session_state.input_files[idx]['data']
                st.text(f"• {file.name} ({len(file_data)} rows, {len(file_data.columns)} cols)")
        
        if st.button("Remove All Files", use_container_width=True):
            safe_reset_session()
            st.rerun()
    else:
        if st.session_state.input_files:
            safe_reset_session()
    
    st.markdown("---")
    
    # Python Script Editor Section
    st.markdown('<p class="section-header">Python Script</p>', unsafe_allow_html=True)
    
    default_script = ""
    
    code_editor = st.text_area(
        "Python Script",
        value=default_script,
        height=400,
        help="Write Python code. Access files via input_files[0], input_files[1], etc.",
        label_visibility="collapsed"
    )
    
    # Run and Clear Buttons
    col1, col2 = st.columns(2)
    with col1:
        run_disabled = not st.session_state.input_files or not code_editor.strip()
        
        if st.button("Run Script", type="primary", use_container_width=True, disabled=run_disabled):
            if st.session_state.input_files and code_editor.strip():
                # Clear previous execution data
                st.session_state.console_logs = []
                st.session_state.error_logs = []
                st.session_state.processed_data = None
                st.session_state.captured_figures = []
                st.session_state.script_executed = False
                
                with st.spinner('Executing script...'):
                    result, figures = execute_python_script(st.session_state.input_files, code_editor)
                    st.session_state.processed_data = result
                    st.session_state.captured_figures = figures
                    st.session_state.script_executed = True
                st.rerun()
    
    with col2:
        if st.button("Clear Logs", use_container_width=True):
            st.session_state.console_logs = []
            st.session_state.error_logs = []
            st.rerun()
    
    st.markdown("---")
    
    # Download Output Section
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
        st.info("Run script to generate downloadable output")
    
    # Download Charts Section
    if st.session_state.captured_figures:
        st.markdown("---")
        st.markdown('<p class="section-header">Export Charts</p>', unsafe_allow_html=True)
        
        zip_data = create_charts_zip(st.session_state.captured_figures)
        if zip_data:
            st.download_button(
                label="Download All Charts (ZIP)",
                data=zip_data,
                file_name=f"charts_{datetime.now().strftime('%Y%m%d_%H%M%S')}.zip",
                mime="application/zip",
                use_container_width=True
            )
            st.success(f"{len(st.session_state.captured_figures)} chart(s) ready!")

# RIGHT COLUMN: Output, data preview, logs
with right_col:
    if st.session_state.input_files or st.session_state.script_executed:
        if st.session_state.script_executed:
            tabs = st.tabs(["Output", "Input Data", "Execution Log", "Error Details"])
        else:
            tabs = st.tabs(["Input Data", "Execution Log", "Error Details"])
        
        # TAB: Output (only after execution)
        if st.session_state.script_executed:
            with tabs[0]:
                has_output = st.session_state.processed_data is not None or st.session_state.captured_figures
                
                if not has_output:
                    st.warning("No output generated. Script must define 'output_df' or create visualizations.")
                else:
                    # Display visualizations first
                    if st.session_state.captured_figures:
                        st.markdown("### Visualizations")
                        st.markdown(f"*{len(st.session_state.captured_figures)} chart(s) generated*")
                        st.markdown("")
                        
                        for idx, (fig_type, fig_data) in enumerate(st.session_state.captured_figures):
                            if fig_type == 'matplotlib':
                                st.image(f"data:image/png;base64,{fig_data}", use_container_width=True)
                            elif fig_type == 'plotly':
                                st.plotly_chart(fig_data, use_container_width=True)
                            
                            if idx < len(st.session_state.captured_figures) - 1:
                                st.markdown("<br>", unsafe_allow_html=True)
                        
                        if st.session_state.processed_data is not None:
                            st.markdown("---")
                    
                    # Display processed data table
                    if st.session_state.processed_data is not None:
                        st.markdown("### Processed Data")
                        preview_rows = st.slider("Preview rows:", 5, 50, 10, key="output_preview_slider")
                        st.dataframe(st.session_state.processed_data.head(preview_rows), use_container_width=True)
                        
                        st.markdown("")
                        col1, col2, col3 = st.columns(3)
                        with col1:
                            st.metric("Total Rows", len(st.session_state.processed_data))
                        with col2:
                            st.metric("Total Columns", len(st.session_state.processed_data.columns))
                        with col3:
                            memory_kb = st.session_state.processed_data.memory_usage(deep=True).sum() / 1024
                            st.metric("Memory", f"{memory_kb:.1f} KB")
            
            # TAB: Input Data (after execution)
            with tabs[1]:
                st.markdown("### Input Data (Before Processing)")
                
                if len(st.session_state.input_files) == 1:
                    file_info = st.session_state.input_files[0]
                    st.info(f"**File:** {file_info['name']}")
                    preview_rows = st.slider("Preview rows:", 5, 50, 10, key="input_preview_after_exec")
                    st.dataframe(file_info['data'].head(preview_rows), use_container_width=True)
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        st.metric("Total Rows", len(file_info['data']))
                    with col2:
                        st.metric("Total Columns", len(file_info['data'].columns))
                else:
                    file_options = [f"File {idx}: {f['name']}" for idx, f in enumerate(st.session_state.input_files)]
                    selected = st.selectbox("Select file to preview:", file_options, key="file_selector_after_exec")
                    
                    if selected:
                        file_idx = int(selected.split(":")[0].replace("File ", ""))
                        file_info = st.session_state.input_files[file_idx]
                        
                        preview_rows = st.slider("Preview rows:", 5, 50, 10, key=f"input_preview_multi_{file_idx}")
                        st.dataframe(file_info['data'].head(preview_rows), use_container_width=True)
                        
                        col1, col2 = st.columns(2)
                        with col1:
                            st.metric("Total Rows", len(file_info['data']))
                        with col2:
                            st.metric("Total Columns", len(file_info['data'].columns))
            
            # TAB: Execution Log
            with tabs[2]:
                if st.session_state.console_logs:
                    console_html = '<div class="console-box">'
                    console_html += '\n'.join(st.session_state.console_logs)
                    console_html += '</div>'
                    st.markdown(console_html, unsafe_allow_html=True)
                else:
                    st.info("No execution logs yet.")
            
            # TAB: Error Details
            with tabs[3]:
                if st.session_state.error_logs:
                    error_html = '<div class="error-console-box">'
                    error_html += '\n\n'.join(st.session_state.error_logs)
                    error_html += '</div>'
                    st.markdown(error_html, unsafe_allow_html=True)
                else:
                    st.success("No errors detected")
        
        # Before execution - only show Input Data, Logs, Errors
        else:
            # TAB: Input Data (before execution)
            with tabs[0]:
                st.markdown("### Input Data Preview")
                
                if len(st.session_state.input_files) == 1:
                    file_info = st.session_state.input_files[0]
                    st.info(f"**File:** {file_info['name']} | Run the script to process this data.")
                    preview_rows = st.slider("Preview rows:", 5, 50, 10, key="input_preview_before_exec")
                    st.dataframe(file_info['data'].head(preview_rows), use_container_width=True)
                    
                    col1, col2 = st.columns(2)
                    with col1:
                        st.metric("Total Rows", len(file_info['data']))
                    with col2:
                        st.metric("Total Columns", len(file_info['data'].columns))
                else:
                    st.info("Multiple files uploaded. Run the script to process them.")
                    file_options = [f"File {idx}: {f['name']}" for idx, f in enumerate(st.session_state.input_files)]
                    selected = st.selectbox("Select file to preview:", file_options, key="file_selector_before_exec")
                    
                    if selected:
                        file_idx = int(selected.split(":")[0].replace("File ", ""))
                        file_info = st.session_state.input_files[file_idx]
                        
                        preview_rows = st.slider("Preview rows:", 5, 50, 10, key=f"input_preview_single_{file_idx}")
                        st.dataframe(file_info['data'].head(preview_rows), use_container_width=True)
                        
                        col1, col2 = st.columns(2)
                        with col1:
                            st.metric("Total Rows", len(file_info['data']))
                        with col2:
                            st.metric("Total Columns", len(file_info['data'].columns))
            
            # TAB: Execution Log
            with tabs[1]:
                if st.session_state.console_logs:
                    console_html = '<div class="console-box">'
                    console_html += '\n'.join(st.session_state.console_logs)
                    console_html += '</div>'
                    st.markdown(console_html, unsafe_allow_html=True)
                else:
                    st.info("No execution logs yet. Run a script to see logs here.")
            
            # TAB: Error Details
            with tabs[2]:
                if st.session_state.error_logs:
                    error_html = '<div class="error-console-box">'
                    error_html += '\n\n'.join(st.session_state.error_logs)
                    error_html += '</div>'
                    st.markdown(error_html, unsafe_allow_html=True)
                else:
                    st.success("No errors detected")
else:
        st.markdown('<div style="display: flex; align-items: center; justify-content: center; height: 400px; color: #6b7280; font-size: 16px;">Upload files to get started</div>', unsafe_allow_html=True)
