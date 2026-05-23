import streamlit as st
import pandas as pd
import plotly.express as px
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
import io

# 1. Page Configuration & Theme Styling
st.set_page_config(page_title="EXTRA HD DATA ANALYZER", layout="wide")

# Advanced Custom CSS for 3D Neumorphic / Skeuomorphic Look
st.markdown("""
    <style>
    /* Background Gradient */
    .stApp {
        background: linear-gradient(135deg, #f0f4f8 0%, #e2ebf0 100%);
    }
    
    /* 3D Soft Embossed Headers */
    h1 {
        color: #003366 !important;
        font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
        font-weight: 800;
        text-shadow: 2px 2px 4px rgba(0, 0, 0, 0.1), -1px -1px 0px rgba(255, 255, 255, 0.9);
    }
    h2, h3 { 
        color: #003366 !important; 
        font-family: 'Segoe UI', sans-serif;
        font-weight: 600;
    }
    
    /* 3D Glass & Neumorphic Cards for Data & File Uploader */
    div[data-testid="stFileUploader"], .stDataFrame, div[data-testid="stMetricBlock"] {
        background: #f0f4f8 !important;
        border-radius: 16px !important;
        box-shadow: 9px 9px 16px rgba(163, 177, 198, 0.6), -9px -9px 16px rgba(255, 255, 255, 0.8) !important;
        border: 1px solid rgba(255, 255, 255, 0.5) !important;
        padding: 20px !important;
    }

    /* Tabs Styling 3D buttons effect */
    button[data-baseweb="tab"] {
        background: #f0f4f8 !important;
        border-radius: 10px 10px 0px 0px !important;
        box-shadow: 3px -3px 6px rgba(163, 177, 198, 0.4), -3px -3px 6px rgba(255, 255, 255, 0.7) !important;
        border: none !important;
        margin-right: 5px !important;
        color: #003366 !important;
        font-weight: bold !important;
    }
    button[data-baseweb="tab"][aria-selected="true"] {
        background: #003366 !important;
        color: #ffffff !important;
        box-shadow: inset 2px 2px 5px rgba(0,0,0,0.3) !important;
    }

    /* 3D Hover Effects on Buttons */
    div.stButton > button {
        background: linear-gradient(145deg, #003366, #002244) !important;
        color: white !important;
        border-radius: 12px !important;
        border: none !important;
        box-shadow: 5px 5px 10px rgba(163, 177, 198, 0.6), -5px -5px 10px rgba(255, 255, 255, 0.8) !important;
        transition: all 0.2s ease-in-out !important
