#!/usr/bin/env python
# -*- coding: utf-8 -*-

r"""
Subterranean Detection App (Enterprise-Grade UI)
-------------------------------------------------
For Option A, data is read from "folder_a" (relative to main.py) and for Option B, from a hard-coded path:
    C:\Users\ilioumbas\Documents\GitHub\Sub
Ensure that for Option B, your area folders (e.g. "7") are located directly in that folder with their
subfolders (e.g. "7_grid", "7_jpgs", etc.).
"""

import os
import glob
import re
from datetime import datetime, date
import xml.etree.ElementTree as ET

import numpy as np
import pandas as pd
import rasterio
import streamlit as st
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots

from rasterio.errors import NotGeoreferencedWarning
import warnings
warnings.filterwarnings("ignore", category=NotGeoreferencedWarning)

# Global debug flag. Set to True to show extra debug output.
DEBUG = False

def debug(*args, **kwargs):
    if DEBUG:
        st.write(*args, **kwargs)

# -------------------------------------------------------------------------
# Streamlit page configuration
# -------------------------------------------------------------------------
st.set_page_config(
    page_title="Subterranean Detection Characteristics",
    layout="wide",
    initial_sidebar_state="expanded"
)

# -----------------------------------------------------------------------------
# Inject custom CSS
# -----------------------------------------------------------------------------
def inject_custom_css():
    custom_css = """
    <link href="https://fonts.googleapis.com/css?family=Roboto:400,500,700&display=swap" rel="stylesheet">
    <style>
        html, body, [class*="css"] { font-family: 'Roboto', sans-serif; }
        .block-container { background: #0d0d0d; color: #e0e0e0; padding: 1rem; }
        .sidebar .sidebar-content { background: #1b1b1b; border: none; }
        .card { background: #1e1e1e; padding: 2rem; border-radius: 12px;
                box-shadow: 0 4px 8px rgba(0,0,0,0.6); margin-bottom: 2rem; }
        .header-title { color: #ffca28; margin-bottom: 1rem; font-size: 1.75rem; text-align: center; }
        .nav-section { padding: 1rem; background: #262626; border-radius: 8px; margin-bottom: 1rem; }
        .nav-section h4 { margin: 0; color: #ffca28; font-weight: 500; }
        .stButton button { background-color: #3949ab; color: #fff; border-radius: 8px; padding: 10px 20px; border: none;
                           box-shadow: 0 3px 6px rgba(0,0,0,0.3); transition: background-color 0.3s ease; }
        .stButton button:hover { background-color: #5c6bc0; }
        .plotly-graph-div { border: 1px solid #333; border-radius: 8px; }
    </style>
    """
    st.markdown(custom_css, unsafe_allow_html=True)

inject_custom_css()

# -----------------------------------------------------------------------------
# Helper functions for file and date handling
# -----------------------------------------------------------------------------
def extract_date_from_filename(filename: str):
    basename = os.path.basename(filename)
    debug("Extracting date from filename:", basename)
    match = re.search(r'(\d{4})[_-](\d{2})[_-](\d{2})', basename)
    if not match:
        match = re.search(r'(\d{4})(\d{2})(\d{2})', basename)
    if match:
        year, month, day = match.groups()
        try:
            date_obj = datetime(int(year), int(month), int(day))
            day_of_year = date_obj.timetuple().tm_yday
            return day_of_year, date_obj
        except Exception as e:
            debug("Error converting date:", e)
            return None, None
    return None, None

def load_lake_shape_from_xml(xml_file: str, bounds: tuple = None, xml_width: float = 518.0, xml_height: float = 505.0):
    debug("Loading outline from:", xml_file)
    try:
        tree = ET.parse(xml_file)
        root = tree.getroot()
        points = []
        for point_elem in root.findall("point"):
            x_str = point_elem.get("x")
            y_str = point_elem.get("y")
            if x_str is None or y_str is None:
                continue
            points.append([float(x_str), float(y_str)])
        if not points:
            st.warning("No points found in XML:", xml_file)
            return None
        if bounds is not None:
            minx, miny, maxx, maxy = bounds
            transformed_points = []
            for x_xml, y_xml in points:
                x_geo = minx + (x_xml / xml_width) * (maxx - minx)
                y_geo = maxx - (y_xml / xml_height) * (maxx - miny)
                transformed_points.append([x_geo, y_geo])
            points = transformed_points
        if points and (points[0] != points[-1]):
            points.append(points[0])
        debug("Loaded", len(points), "points.")
        return {"type": "Polygon", "coordinates": [points]}
    except Exception as e:
        st.error(f"Error loading outline from {xml_file}: {e}")
        return None

def read_image(file_path: str, lake_shape: dict = None):
    debug("Reading image from:", file_path)
    with rasterio.open(file_path) as src:
        img = src.read(1).astype(np.float32)
        profile = src.profile.copy()
        profile.update(dtype="float32")
        no_data_value = src.nodata
        if no_data_value is not None:
            img = np.where(img == no_data_value, np.nan, img)
        img = np.where(img == 0, np.nan, img)
        if lake_shape is not None:
            from rasterio.features import geometry_mask
            poly_mask = geometry_mask([lake_shape], transform=src.transform, invert=False, out_shape=img.shape)
            img = np.where(~poly_mask, img, np.nan)
    return img, profile

def load_data(input_folder: str, shapefile_name="shapefile.xml"):
    debug("Loading data from folder:", input_folder)
    if not os.path.exists(input_folder):
        raise Exception(f"Folder does not exist: {input_folder}")
    shapefile_path_xml = os.path.join(input_folder, shapefile_name)
    shapefile_path_txt = os.path.join(input_folder, "shapefile.txt")
    lake_shape = None
    if os.path.exists(shapefile_path_xml):
        shape_file = shapefile_path_xml
    elif os.path.exists(shapefile_path_txt):
        shape_file = shapefile_path_txt
    else:
        shape_file = None
        debug("No XML outline found in folder", input_folder)
    all_tif_files = sorted(glob.glob(os.path.join(input_folder, "*.tif")))
    tif_files = [fp for fp in all_tif_files if os.path.basename(fp).lower() != "mask.tif"]
    if not tif_files:
        raise Exception("No GeoTIFF files found.")
    with rasterio.open(tif_files[0]) as src:
        bounds = src.bounds
    if shape_file is not None:
        lake_shape = load_lake_shape_from_xml(shape_file, bounds=bounds)
    images, days, date_list = [], [], []
    for file_path in tif_files:
        day_of_year, date_obj = extract_date_from_filename(file_path)
        if day_of_year is None:
            continue
        img, _ = read_image(file_path, lake_shape=lake_shape)
        images.append(img)
        days.append(day_of_year)
        date_list.append(date_obj)
    if not images:
        raise Exception("No valid images found.")
    stack = np.stack(images, axis=0)
    return stack, np.array(days), date_list

# -----------------------------------------------------------------------------
# get_data_folder: For Option A, use folder_a; for Option B, hard-code the path.
# -----------------------------------------------------------------------------
def get_data_folder(waterbody: str, index: str) -> str:
    if st.session_state.get("method_option", "Option A") == "Option A":
        base_dir = os.path.dirname(os.path.abspath(__file__))
        base_folder = os.path.join(base_dir, "folder_a")
    else:
        # Hard-coded path for Option B.
        base_folder = r"C:\Users\ilioumbas\Documents\GitHub\Sub"
    
    debug("Base folder being used:", base_folder)
    if not os.path.exists(base_folder):
        st.error(f"Base folder does not exist: {base_folder}")
        return None

    waterbody_folder = os.path.join(base_folder, waterbody)
    debug("Looking for area folder at:", waterbody_folder)
    if not os.path.exists(waterbody_folder):
        st.error(f"Area folder not found: {waterbody_folder}")
        return None

    if index == "Χλωροφύλλη":
        data_folder = os.path.join(waterbody_folder, "Chlorophyll")
    elif index == "Burned Areas":
        data_folder = os.path.join(waterbody_folder, "Burned Areas")
    elif index == "Πραγματικό" and st.session_state.get("method_option") != "Option A":
        data_folder = os.path.join(waterbody_folder, "Pragmatiko")
    else:
        data_folder = os.path.join(waterbody_folder, index)
    
    debug("Data folder resolved to:", data_folder)
    if not os.path.exists(data_folder):
        st.error(f"Data folder does not exist: {data_folder}")
        return None
    return data_folder

# -----------------------------------------------------------------------------
# UI Functions
# -----------------------------------------------------------------------------
def run_intro_page():
    with st.container():
        st.markdown('<div class="card">', unsafe_allow_html=True)
        col_logo, col_text = st.columns([1, 3])
        with col_logo:
            base_dir = os.path.dirname(os.path.abspath(__file__))
            logo_path = os.path.join(base_dir, "logo.jpg")
            if os.path.exists(logo_path):
                st.image(logo_path, width=250)
            else:
                debug("Logo not found.")
        with col_text:
            st.markdown("<h2 class='header-title'>Subterranean Detection Characteristics</h2>", unsafe_allow_html=True)
            st.markdown("<p style='text-align: center; font-size: 1.1rem;'>This detection application uses remote sensing tools. Select the settings from the sidebar and explore the data.</p>", unsafe_allow_html=True)
        st.markdown('</div>', unsafe_allow_html=True)

def run_custom_ui():
    st.sidebar.markdown("<div class='nav-section'><h4>Analysis Settings</h4></div>", unsafe_allow_html=True)
    base_dir = os.path.dirname(os.path.abspath(__file__))

    method_option = st.sidebar.selectbox("Select Methodology", ["Option A", "Option B"], key="method_option")
    if method_option == "Option A":
        chosen_dir = os.path.join(base_dir, "folder_a")
    else:
        # Hard-coded path for Option B.
        chosen_dir = r"C:\Users\ilioumbas\Documents\GitHub\Sub"
    
    st.write(f"**Data will be read from:** {chosen_dir}")
    debug("Chosen directory:", chosen_dir)
    
    # Debug: List entire directory tree for chosen_dir.
    st.write("### DEBUG: Directory Tree for chosen_dir:")
    for root, dirs, files in os.walk(chosen_dir):
        st.write(f"**{root}**")
        st.write("   dirs:", dirs)
        st.write("   files:", files)
    
    if not os.path.exists(chosen_dir):
        st.error("No valid mother folders found with required subfolders.")
        return

    if method_option == "Option A":
        area_options = sorted(
            [d for d in os.listdir(chosen_dir) if os.path.isdir(os.path.join(chosen_dir, d))]
        )
    else:
        area_options = sorted(
            [d for d in os.listdir(chosen_dir) if os.path.isdir(os.path.join(chosen_dir, d))]
        )
        if not area_options:
            st.warning("No subdirectories found in Option B path; using default area list.")
            area_options = ["7"]
    
    st.write("### DEBUG: Found subdirectories:", area_options)
    
    area = st.sidebar.selectbox("Select Area", area_options, key="waterbody_choice")
    index = st.sidebar.selectbox("Select Index",
                                 ["Πραγματικό", "Χλωροφύλλη", "CDOM", "Colour", "Burned Areas"],
                                 key="index_choice")
    analysis = st.sidebar.selectbox("Select Analysis Type",
                                    ["Subterranean Processing", "Subterranean Quality Dashboard"],
                                    key="analysis_choice")
    st.sidebar.markdown(f"""
    <div style="padding: 0.5rem; background:#262626; border-radius:5px; margin-top:1rem;">
        <strong>Methodology:</strong> {method_option}<br>
        <strong>Area:</strong> {area}<br>
        <strong>Index:</strong> {index}<br>
        <strong>Analysis:</strong> {analysis}
    </div>
    """, unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# Core Processing Functions
# -----------------------------------------------------------------------------
def run_lake_processing_app(waterbody: str, index: str):
    with st.container():
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.title(f"Subterranean Processing ({waterbody} - {index})")
        data_folder = get_data_folder(waterbody, index)
        if data_folder is None:
            st.error("Data folder for the selected area/index does not exist.")
            st.stop()
        input_folder = os.path.join(data_folder, "GeoTIFFs")
        try:
            STACK, DAYS, DATES = load_data(input_folder)
        except Exception as e:
            st.error(f"Data loading error: {e}")
            st.stop()
        if not DATES:
            st.error("No date information available.")
            st.stop()
        # Replace the following placeholder with your full lake processing logic.
        st.write("Running Lake Processing... (insert your processing logic here)")

def run_water_quality_dashboard(waterbody: str, index: str):
    with st.container():
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.title(f"Subterranean Quality Dashboard ({waterbody} - {index})")
        data_folder = get_data_folder(waterbody, index)
        if data_folder is None:
            st.error("Data folder for the selected area/index does not exist.")
            st.stop()
        # Replace the following placeholder with your full dashboard logic.
        st.write("Running Water Quality Dashboard... (insert your dashboard logic here)")

# -----------------------------------------------------------------------------
# Main entry point
# -----------------------------------------------------------------------------
def main():
    debug("Entered main()")
    run_intro_page()
    run_custom_ui()
    
    wb = st.session_state.get("waterbody_choice", None)
    idx = st.session_state.get("index_choice", None)
    analysis = st.session_state.get("analysis_choice", None)
    debug("Selections: Area =", wb, "Index =", idx, "Analysis =", analysis)
    
    if wb is not None and idx is not None:
        if analysis == "Subterranean Processing":
            run_lake_processing_app(wb, idx)
        elif analysis == "Subterranean Quality Dashboard":
            run_water_quality_dashboard(wb, idx)
        else:
            st.info("Please select an analysis type.")
    else:
        st.warning("No available data for this combination.")

if __name__ == "__main__":
    from multiprocessing import freeze_support
    freeze_support()
    main()
