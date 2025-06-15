import streamlit as st
import pandas as pd
from io import BytesIO
import time
from ai_spec_processor import AISpecProcessor
import os
from datetime import datetime

# Page configuration
st.set_page_config(
    page_title="AI Spec Sheet Generator",
    page_icon="👟",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Custom CSS for better styling
st.markdown("""
<style>
    .main-header {
        font-size: 3rem;
        font-weight: bold;
        color: #1f77b4;
        text-align: center;
        margin-bottom: 2rem;
    }
    .upload-section {
        background-color: #f0f2f6;
        padding: 2rem;
        border-radius: 10px;
        margin: 1rem 0;
    }
    .success-box {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        color: #155724;
        padding: 1rem;
        border-radius: 5px;
        margin: 1rem 0;
    }
    .error-box {
        background-color: #f8d7da;
        border: 1px solid #f5c6cb;
        color: #721c24;
        padding: 1rem;
        border-radius: 5px;
        margin: 1rem 0;
    }
    .info-box {
        background-color: #d1ecf1;
        border: 1px solid #bee5eb;
        color: #0c5460;
        padding: 1rem;
        border-radius: 5px;
        margin: 1rem 0;
    }
</style>
""", unsafe_allow_html=True)

# Initialize session state
if 'processor' not in st.session_state:
    st.session_state.processor = None
if 'processing_result' not in st.session_state:
    st.session_state.processing_result = None
if 'files_uploaded' not in st.session_state:
    st.session_state.files_uploaded = False

# Header
st.markdown('<h1 class="main-header">👟 AI-Powered Spec Sheet Generator</h1>', unsafe_allow_html=True)

# Sidebar
st.sidebar.markdown("## ⚙️ Configuration")

# API Key input
api_key = st.sidebar.text_input(
    "OpenAI API Key",
    type="password",
    placeholder="sk-proj-...",
    help="Your OpenAI API key for AI-enhanced material matching"
)

# AI Settings
st.sidebar.markdown("### 🤖 AI Settings")
use_ai = st.sidebar.checkbox("Enable AI Enhancement", value=True, help="Use GPT-4.1 for smart material matching")
confidence_threshold = st.sidebar.slider("AI Confidence Threshold", 0.5, 1.0, 0.7, 0.05)

# Initialize processor if API key is provided
if api_key and not st.session_state.processor:
    try:
        st.session_state.processor = AISpecProcessor(api_key)
        st.sidebar.success("✅ AI Processor initialized!")
    except Exception as e:
        st.sidebar.error(f"❌ Error initializing AI: {str(e)}")

# Main content
if not api_key:
    st.markdown("""
    <div class="info-box">
        <h3>🚀 Welcome to the AI Spec Sheet Generator!</h3>
        <p>This tool uses advanced AI to automatically generate spec sheets from your development sample logs.</p>
        <p><strong>Please enter your OpenAI API key in the sidebar to get started.</strong></p>
    </div>
    """, unsafe_allow_html=True)
    
    st.markdown("""
    ## 📋 How it works:
    1. **Upload your files**: Development Sample Log, Spec Template, and Simplified BOM
    2. **AI Processing**: Our AI extracts and standardizes material names
    3. **Download results**: Get a complete workbook with one spec sheet per sample
    
    ## 🎯 Features:
    - **Smart Material Matching**: AI understands synonyms, abbreviations, and variations
    - **Complex Text Parsing**: Handles messy supplier descriptions
    - **Fallback System**: Exact → Fuzzy → AI matching for maximum accuracy
    - **Batch Processing**: Handle hundreds of samples at once
    """)

else:
    # File upload section
    st.markdown("## 📁 Upload Your Files")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        st.markdown("""
        <div style="
            border: 3px dashed #1f77b4;
            border-radius: 15px;
            padding: 2rem;
            text-align: center;
            background: linear-gradient(135deg, #f8f9ff 0%, #e8f4fd 100%);
            margin: 1rem 0;
            transition: all 0.3s ease;
        ">
            <div style="font-size: 3rem; margin-bottom: 1rem;">📊</div>
            <h3 style="color: #1f77b4; margin-bottom: 1rem;">Development Sample Log</h3>
            <p style="color: #666; margin-bottom: 1.5rem;">Drag & drop your Excel file here<br/>or click to browse</p>
        </div>
        """, unsafe_allow_html=True)
        
        dev_log_file = st.file_uploader(
            "Upload Excel file",
            type=['xlsx', 'xls'],
            key="dev_log",
            help="The file containing your sample data (one row per sample)",
            label_visibility="collapsed"
        )
        if dev_log_file:
            st.success(f"✅ {dev_log_file.name}")
            # Preview
            try:
                preview_df = pd.read_excel(dev_log_file)
                st.write(f"📏 Shape: {preview_df.shape}")
                st.write("🔍 First few rows:")
                st.dataframe(preview_df.head(2), use_container_width=True)
            except Exception as e:
                st.error(f"Error previewing file: {e}")
    
    with col2:
        st.markdown("""
        <div style="
            border: 3px dashed #28a745;
            border-radius: 15px;
            padding: 2rem;
            text-align: center;
            background: linear-gradient(135deg, #f8fff8 0%, #e8f5e8 100%);
            margin: 1rem 0;
            transition: all 0.3s ease;
        ">
            <div style="font-size: 3rem; margin-bottom: 1rem;">📝</div>
            <h3 style="color: #28a745; margin-bottom: 1rem;">Spec Template</h3>
            <p style="color: #666; margin-bottom: 1.5rem;">Drag & drop your template here<br/>or click to browse</p>
        </div>
        """, unsafe_allow_html=True)
        
        template_file = st.file_uploader(
            "Upload Excel template",
            type=['xlsx', 'xls'],
            key="template",
            help="Your blank spec sheet template",
            label_visibility="collapsed"
        )
        if template_file:
            st.success(f"✅ {template_file.name}")
    
    with col3:
        st.markdown("""
        <div style="
            border: 3px dashed #ffc107;
            border-radius: 15px;
            padding: 2rem;
            text-align: center;
            background: linear-gradient(135deg, #fffef8 0%, #fff8e1 100%);
            margin: 1rem 0;
            transition: all 0.3s ease;
        ">
            <div style="font-size: 3rem; margin-bottom: 1rem;">🔍</div>
            <h3 style="color: #ffc107; margin-bottom: 1rem;">Simplified BOM</h3>
            <p style="color: #666; margin-bottom: 1.5rem;">Drag & drop your BOM here<br/>or click to browse</p>
        </div>
        """, unsafe_allow_html=True)
        
        bom_file = st.file_uploader(
            "Upload BOM Excel file",
            type=['xlsx', 'xls'],
            key="bom",
            help="Your standardized material vocabulary",
            label_visibility="collapsed"
        )
        if bom_file:
            st.success(f"✅ {bom_file.name}")
            # Load BOM into processor
            if st.session_state.processor:
                bom_buffer = BytesIO(bom_file.getvalue())
                if st.session_state.processor.load_bom(bom_buffer):
                    st.success(f"🎯 Loaded {len(st.session_state.processor.materials)} materials")
                else:
                    st.error("Failed to load BOM")
    
    # Processing section
    if dev_log_file and template_file and bom_file and st.session_state.processor:
        st.markdown("---")
        st.markdown("## 🚀 Process Spec Sheets")
        
        col1, col2 = st.columns([3, 1])
        
        with col1:
            st.markdown("""
            **Ready to process!** The AI will:
            - Parse complex material descriptions
            - Match materials to your standardized BOM
            - Generate one spec sheet per sample
            """)
        
        with col2:
            if st.button("🎯 Generate Spec Sheets", type="primary", use_container_width=True):
                with st.spinner("🤖 AI is processing your files..."):
                    try:
                        # Prepare file buffers
                        dev_buffer = BytesIO(dev_log_file.getvalue())
                        template_buffer = BytesIO(template_file.getvalue())
                        
                        # Process
                        result = st.session_state.processor.process_spec_sheets(
                            dev_buffer, template_buffer
                        )
                        
                        st.session_state.processing_result = result
                        
                    except Exception as e:
                        st.error(f"Processing failed: {str(e)}")
        
        # Results section
        if st.session_state.processing_result:
            result = st.session_state.processing_result
            
            st.markdown("---")
            st.markdown("## 📊 Processing Results")
            
            if result.success:
                st.markdown('<div class="success-box">', unsafe_allow_html=True)
                st.markdown(f"### ✅ Successfully processed {result.samples_processed}/{result.total_samples} samples!")
                st.markdown('</div>', unsafe_allow_html=True)
                
                # Statistics
                col1, col2, col3, col4 = st.columns(4)
                with col1:
                    st.metric("🎯 Exact Matches", result.matches_by_method.get('exact_matches', 0))
                with col2:
                    st.metric("🔍 Fuzzy Matches", result.matches_by_method.get('fuzzy_matches', 0))
                with col3:
                    st.metric("🤖 AI Matches", result.matches_by_method.get('ai_matches', 0))
                with col4:
                    st.metric("❓ No Matches", result.matches_by_method.get('no_matches', 0))
                
                # Download section
                if result.output_file:
                    st.markdown("### 📥 Download Your Results")
                    
                    filename = f"AI_Generated_Spec_Sheets_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                    
                    st.download_button(
                        label="📁 Download Spec Sheets",
                        data=result.output_file,
                        file_name=filename,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary",
                        use_container_width=True
                    )
                
                # Errors section
                if result.errors:
                    st.markdown("### ⚠️ Processing Warnings")
                    for error in result.errors:
                        st.warning(error)
            
            else:
                st.markdown('<div class="error-box">', unsafe_allow_html=True)
                st.markdown("### ❌ Processing Failed")
                for error in result.errors:
                    st.error(error)
                st.markdown('</div>', unsafe_allow_html=True)

# Footer
st.markdown("---")
st.markdown("""
<div style="text-align: center; color: #666; padding: 2rem;">
    Made with ❤️ using Streamlit and OpenAI GPT-4.1 | 
    <a href="#" style="color: #1f77b4;">Need help?</a>
</div>
""", unsafe_allow_html=True) 