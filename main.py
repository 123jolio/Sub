#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Water Quality App (Enterprise-Grade UI)
-----------------------------------------
Φιλικό, επαγγελματικό περιβάλλον ανάλυσης δορυφορικών δεδομένων υδάτων.
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

import matplotlib.pyplot as plt
import matplotlib.colors as mcolors

# Global debug flag (set to True for debugging output)
DEBUG = False

def debug(*args, **kwargs):
    if DEBUG:
        st.write(*args, **kwargs)

# ---------------------------------------------------
# Εξατομίκευση CSS & Animation για Pro Look
# ---------------------------------------------------
def inject_custom_css():
    custom_css = """
    <link href="https://fonts.googleapis.com/css?family=Roboto:400,500,700&display=swap" rel="stylesheet">
    <style>
        html, body, [class*="css"] {
            font-family: 'Roboto', sans-serif;
        }
        .block-container {
            background: #161b22; /* Darker background for a more modern feel */
            color: #e0e0e0;
            padding: 1.2rem; /* Slightly more padding */
        }
        .sidebar .sidebar-content {
            background: #23272f; /* Sidebar background */
            border: none;
        }
        .card {
            background: #1a1a1d; /* Card background - very dark gray */
            padding: 2rem 2.5rem; /* More padding inside cards */
            border-radius: 16px; /* More rounded corners */
            box-shadow: 0 4px 16px rgba(0,0,0,0.25); /* Softer shadow */
            margin-bottom: 2rem;
            animation: fadein 1.5s; /* Fade-in animation for cards */
        }
        @keyframes fadein { 0% {opacity:0;} 100%{opacity:1;} }
        .header-title {
            color: #ffd600; /* Gold color for titles */
            margin-bottom: 1rem;
            font-size: 2rem; /* Larger title font */
            text-align: center;
            letter-spacing: 0.5px;
            font-weight: 700; /* Bolder title */
        }
        .nav-section {
            padding: 1rem;
            background: #2c2f36; /* Nav section background in sidebar */
            border-radius: 10px;
            margin-bottom: 1.2rem;
        }
        .nav-section h4 {
            margin: 0;
            color: #ffd600; /* Gold for nav section titles */
            font-weight: 500;
        }
        .stButton button {
            background-color: #009688; /* Teal button color */
            color: #fff;
            border-radius: 8px;
            padding: 10px 20px;
            border: none;
            box-shadow: 0 3px 6px rgba(0,0,0,0.12);
            font-size: 1.05rem; /* Slightly larger button font */
            transition: background-color 0.2s;
        }
        .stButton button:hover {
            background-color: #26a69a; /* Lighter teal on hover */
        }
        .plotly-graph-div {
            border: 1px solid #23272f; /* Border for Plotly graphs */
            border-radius: 10px;
        }
        .legend { /* Custom class for legends if needed */
            font-size: 0.95rem;
            color: #ffd600;
        }
        .footer {
            text-align:center;
            color:gray;
            font-size:0.9rem;
            padding:1.3rem 0 0.1rem 0;
        }
    </style>
    """
    st.markdown(custom_css, unsafe_allow_html=True)

# Call CSS injection at the start
# inject_custom_css() # This will be called in main() or when script runs

# ---------------------------------------------------
# Footer Branding
# ---------------------------------------------------
def render_footer():
    st.markdown("""
        <hr>
        <div class='footer'>
            © 2025 EYATH SA • Powered by OpenAI & Streamlit | Contact: <a href='mailto:ilioumbas@eyath.gr'>ilioumbas@eyath.gr</a>
        </div>
    """, unsafe_allow_html=True)

# ---------------------------------------------------
# Καλωσόρισμα και οδηγίες
# ---------------------------------------------------
def run_intro_page_new(): # Renamed to avoid conflict if an old one exists
    with st.container():
        st.markdown('<div class="card">', unsafe_allow_html=True)
        col_logo, col_text = st.columns([1, 3]) # Adjusted column ratio
        with col_logo:
            # Determine base_dir safely for Streamlit Cloud compatibility
            base_dir = os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else os.getcwd()
            logo_path = os.path.join(base_dir, "logo.jpg") # Assuming logo.jpg is in the same directory
            if os.path.exists(logo_path):
                st.image(logo_path, width=180, output_format="auto", caption="EYATH Water Quality") # Adjusted width
            else:
                st.markdown("💧") # Fallback if logo not found
        with col_text:
            st.markdown("""
                <h2 class='header-title'>🚀 Καλωσορίσατε στην Εφαρμογή Ανάλυσης Υδάτων EYATH</h2>
                <p style='font-size:1.15rem;text-align:center'>
                Εξερευνήστε τα δεδομένα ποιότητας με ευκολία.<br>
                Επιλέξτε τι θέλετε να δείτε από το πλάι και απολαύστε δυναμικά, διαδραστικά γραφήματα!
                </p>
                """, unsafe_allow_html=True)
            with st.expander("🔰 Οδηγίες Χρήσης", expanded=False):
                st.write("""
                    - Επιλέξτε υδάτινο σώμα, δείκτη και ανάλυση στην πλαϊνή μπάρα.
                    - Περιηγηθείτε στις καρτέλες με τα διαγράμματα.
                    - Ανεβάστε το δικό σας KML για custom σημεία δειγματοληψίας.
                    - Όλα τα δεδομένα & εικόνες μένουν τοπικά.
                """)
        st.markdown('</div>', unsafe_allow_html=True)

# ---------------------------------------------------
# Sidebar Navigation
# ---------------------------------------------------
def run_custom_ui_new(): # Renamed
    st.sidebar.markdown("<div class='nav-section'><h4>🛠️ Επιλογές Ανάλυσης</h4></div>", unsafe_allow_html=True)
    st.sidebar.info("❔ Επιλέξτε τις ρυθμίσεις σας και προχωρήστε στα αποτελέσματα!")
    # Προσαρμοσμένες επιλογές για το Γαδουρά
    waterbody = st.sidebar.selectbox("🌊 Υδάτινο σώμα", ["Γαδουρά"], key="waterbody_choice")
    index = st.sidebar.selectbox("🔬 Δείκτης", ["Πραγματικό", "Χλωροφύλλη", "Θολότητα"], key="index_choice") # Προστέθηκε Θολότητα
    analysis = st.sidebar.selectbox(
        "📊 Είδος Ανάλυσης",
        [
            "Επιφανειακή Αποτύπωση", # "Lake Processing"
            "Προφίλ ποιότητας και στάθμης", # "Water Quality Dashboard"
            "Eργαλεία Πρόβλεψης και έγκαιρης ενημέρωσης" # Νέα επιλογή
        ],
        key="analysis_choice"
    )
    st.sidebar.markdown(
        f"""<div style="padding: 0.7rem; background:#2c2f36; border-radius:8px; margin-top:1.2rem;">
        <strong>🌊 Υδάτινο σώμα:</strong> {waterbody}<br>
        <strong>🔬 Δείκτης:</strong> {index}<br>
        <strong>📊 Ανάλυση:</strong> {analysis}
        </div>""",
        unsafe_allow_html=True
    )
# --------------------------------------------------------------------------
# Parsing KML -> sampling points
# --------------------------------------------------------------------------
def parse_sampling_kml(kml_source) -> list:
    try:
        # Κάνε reset το pointer αν είναι file uploader
        if hasattr(kml_source, "seek"):
            kml_source.seek(0)
        # Διαβάζει από uploaded file ή path string
        if hasattr(kml_source, "read"): # For uploaded file objects
            tree = ET.parse(kml_source)
        else: # For file paths (strings)
            tree = ET.parse(str(kml_source)) # Ensure it's a string for ET.parse

        root = tree.getroot()
        ns = {'kml': 'http://www.opengis.net/kml/2.2'} # KML namespace
        points = []
        # Ψάχνει για LineString που περιέχει τα σημεία
        for ls in root.findall('.//kml:LineString', ns):
            coords_text = ls.find('kml:coordinates', ns).text
            if coords_text:
                coords = coords_text.strip().split()
                for i, coord_str in enumerate(coords):
                    lon_str, lat_str, *_ = coord_str.split(',') # Αγνοεί το υψόμετρο
                    points.append((f"Point {i+1}", float(lon_str), float(lat_str)))
        return points
    except Exception as e:
        st.error(f"Σφάλμα ανάλυσης KML: {e}")
        return []

# --------------------------------------------------------------------------
# Core sampling analyzer
# --------------------------------------------------------------------------
# analyze_sampling function from the user's latest code
def analyze_sampling(sampling_points, first_image_data, first_transform,
                     images_folder, lake_height_path, selected_points, # selected_points should be list of point names
                     lower_thresh=0, upper_thresh=255, # These seem unused in current version of analyze_sampling
                     date_min=None, date_max=None): # Date filters for processing
    results_colors = {name: [] for name, _, _ in sampling_points}
    results_mg = {name: [] for name, _, _ in sampling_points}
    
    # Ensure selected_points is a list of names for easy lookup
    if not isinstance(selected_points, list):
        selected_points = [] # Or handle error

    # Iterate through GeoTIFFs
    for filename in sorted(os.listdir(images_folder)):
        if not filename.lower().endswith(('.tif', '.tiff')):
            continue
        
        match_date = re.search(r'(\d{4}_\d{2}_\d{2})', filename) # Flexible date regex
        if not match_date:
            continue
        date_str = match_date.group(1)
        try:
            date_obj = datetime.strptime(date_str, '%Y_%m_%d')
        except ValueError:
            continue

        # Apply date filtering
        if date_min and date_obj.date() < date_min:
            continue
        if date_max and date_obj.date() > date_max:
            continue

        image_path = os.path.join(images_folder, filename)
        with rasterio.open(image_path) as src:
            if src.count < 3: # Need at least 3 bands for RGB
                continue
            
            # Process only selected sampling points relevant to this image
            for point_name, lon, lat in sampling_points:
                if point_name not in selected_points: # Filter by selected_points
                    continue

                # Convert geographic to pixel coordinates
                col, row = (~src.transform) * (lon, lat)
                col, row = int(col), int(row)

                # Check if pixel is within image bounds
                if not (0 <= col < src.width and 0 <= row < src.height):
                    continue
                
                window = rasterio.windows.Window(col, row, 1, 1)
                # Assuming bands 1,2,3 are R,G,B like. Adjust if necessary.
                r_val = src.read(1, window=window)[0,0]
                g_val = src.read(2, window=window)[0,0]
                b_val = src.read(3, window=window)[0,0]
                
                # Calculate mg value (example, ensure this formula is correct for your data)
                # Using the map_rgb_to_mg logic: (g / 255.0) * mg_factor (default 2.0)
                mg_calculated = (g_val / 255.0) * 2.0 if g_val is not None else 0.0 
                
                results_mg[point_name].append((date_obj, mg_calculated))
                # Store normalized RGB for color display
                results_colors[point_name].append((date_obj, (r_val/255.0, g_val/255.0, b_val/255.0)))

    # GeoTIFF figure with sampling points
    # Normalize first_image_data if it's not already 0-1 for px.imshow
    # Assuming first_image_data is (bands, height, width) and needs scaling
    scaled_first_image = first_image_data.astype(np.float32)
    for i in range(scaled_first_image.shape[0]):
        band = scaled_first_image[i,:,:]
        min_val, max_val = np.nanpercentile(band, 2), np.nanpercentile(band, 98)
        if max_val > min_val:
            scaled_first_image[i,:,:] = np.clip((band - min_val) / (max_val - min_val), 0, 1)
        else:
            scaled_first_image[i,:,:] = 0 # Or some other default for flat bands

    rgb_display_bg = np.transpose(scaled_first_image, (1,2,0)) # to H,W,C

    fig_geo = px.imshow(rgb_display_bg, title='Εικόνα GeoTIFF με επιλεγμένα σημεία δειγματοληψίας')
    for name, lon, lat in sampling_points:
        if name not in selected_points: continue
        col, row = (~first_transform) * (lon, lat) # Use the transform of the background image
        fig_geo.add_trace(go.Scatter(x=[col], y=[row], mode='markers+text', text=[name], textposition="top right",
                                     marker=dict(color='yellow', size=10, symbol='cross'), name=name))
    fig_geo.update_xaxes(visible=False)
    fig_geo.update_yaxes(visible=False)
    fig_geo.update_layout(width=900, height=600, showlegend=True) # showlegend might be useful

    # Lake height data
    try:
        df_h = pd.read_excel(lake_height_path)
        # Assuming first column is Date, second is Height
        df_h.rename(columns={df_h.columns[0]: 'Date', df_h.columns[1]: 'Height'}, inplace=True)
        df_h['Date'] = pd.to_datetime(df_h['Date'])
        df_h.sort_values('Date', inplace=True)
    except Exception as e:
        debug(f"Could not read or process lake height Excel: {e}")
        df_h = pd.DataFrame(columns=['Date','Height']) # Empty dataframe

    # Pixel colors plot + lake height
    fig_colors = make_subplots(specs=[[{'secondary_y':True}]])
    
    point_name_to_y_idx = {name_tuple[0]: i for i, name_tuple in enumerate(sampling_points)}

    for name_tuple in sampling_points: # Iterate through all defined sampling points for consistent y-axis
        point_name = name_tuple[0]
        if point_name not in selected_points: # Only plot selected points
            continue
        
        y_idx = point_name_to_y_idx[point_name]
        
        data_for_point = sorted(results_colors.get(point_name, []), key=lambda x: x[0])
        
        if data_for_point:
            dates_plot, colors_plot_tuples = zip(*data_for_point)
            colors_rgb_strings = [f"rgb({int(c[0]*255)},{int(c[1]*255)},{int(c[2]*255)})" for c in colors_plot_tuples]
            fig_colors.add_trace(go.Scatter(x=list(dates_plot), y=[y_idx]*len(dates_plot), mode='markers',
                                            marker=dict(color=colors_rgb_strings, size=12, line=dict(width=1, color='DarkSlateGrey')), 
                                            name=point_name), secondary_y=False)
    
    if not df_h.empty:
        fig_colors.add_trace(go.Scatter(x=df_h['Date'], y=df_h['Height'],
                                        mode='lines', name='Στάθμη Λίμνης', line=dict(color='cyan', width=2)), secondary_y=True)
    
    fig_colors.update_layout(title='Χρώματα Pixel (από εικόνες) & Στάθμη Λίμνης',
                             xaxis_title='Ημερομηνία',
                             yaxis_title='Σημεία Δειγματοληψίας (ID)',
                             yaxis2_title='Στάθμη Λίμνης (m)',
                             showlegend=True,
                             legend_title_text='Legend')
    fig_colors.update_yaxes(tickvals=list(point_name_to_y_idx.values()), 
                            ticktext=[name for name in point_name_to_y_idx.keys() if name in selected_points], 
                            secondary_y=False)


    # Mean mg plot
    all_mg_values_by_date = {}
    for point_name_key in selected_points: # Iterate only over selected points for MG calculation
        mg_data_list = results_mg.get(point_name_key, [])
        for d_obj, mg_val in mg_data_list:
            all_mg_values_by_date.setdefault(d_obj, []).append(mg_val)
            
    sorted_dates_mg = sorted(all_mg_values_by_date.keys())
    mean_mg_values = [np.mean(all_mg_values_by_date[d]) if all_mg_values_by_date[d] else np.nan for d in sorted_dates_mg]
    
    fig_mg = go.Figure()
    if sorted_dates_mg:
        fig_mg.add_trace(go.Scatter(x=sorted_dates_mg, y=mean_mg_values, mode='lines+markers',
                                    marker=dict(color=mean_mg_values, 
                                                colorscale='Viridis', 
                                                colorbar=dict(title='mg/m³'), 
                                                size=10, symbol='diamond'),
                                    line=dict(color='rgba(120,120,120,0.7)'))) # Light gray line
    fig_mg.update_layout(title='Μέση τιμή mg/m³ (από επιλεγμένα σημεία) στην πορεία του χρόνου',
                         xaxis_title='Ημερομηνία', yaxis_title='Μέση τιμή mg/m³')

    # Dual plot: Lake Height & Mean mg/m³
    fig_dual = make_subplots(specs=[[{'secondary_y':True}]])
    if not df_h.empty:
        fig_dual.add_trace(go.Scatter(x=df_h['Date'], y=df_h['Height'], name='Στάθμη Λίμνης',
                                      mode='lines', line=dict(color='deepskyblue')), secondary_y=False)
    if sorted_dates_mg:
        fig_dual.add_trace(go.Scatter(x=sorted_dates_mg, y=mean_mg_values, name='Μέση τιμή mg/m³', 
                                      mode='lines+markers', line=dict(color='orange'),
                                      marker=dict(symbol='star', size=8)), secondary_y=True)
    fig_dual.update_layout(title='Στάθμη Λίμνης & Μέση τιμή mg/m³',
                           xaxis_title='Ημερομηνία',
                           yaxis_title='Στάθμη Λίμνης (m)',
                           yaxis2_title='Μέση τιμή mg/m³')

    return fig_geo, fig_dual, fig_colors, fig_mg, results_colors, results_mg, df_h
# -----------------------------------------------------------------------------
# Helper Function: Create Chlorophyll‑a Legend Figure
# -----------------------------------------------------------------------------
def create_chl_legend_figure():
    levels = [0, 6, 12, 20, 30, 50]
    colors = ["#496FF2", "#82D35F", "#FEFD05", "#FD0004", "#8E2026", "#D97CF5"]
    cmap = mcolors.LinearSegmentedColormap.from_list("ChlLegend", 
                                                      list(zip(np.linspace(0, 1, len(levels)), colors)))
    norm = mcolors.Normalize(vmin=levels[0], vmax=levels[-1])
    fig, ax = plt.subplots(figsize=(6, 1.5)) # Adjusted for better aspect ratio
    fig.subplots_adjust(bottom=0.5)
    sm = plt.cm.ScalarMappable(cmap=cmap, norm=norm)
    sm.set_array([]) # Critical for ScalarMappable
    cbar = fig.colorbar(sm, cax=ax, orientation="horizontal", ticks=levels)
    cbar.ax.set_xticklabels([str(l) for l in levels])
    cbar.set_label("Chlorophyll‑a concentration (mg/m³)")
    plt.tight_layout() # Ensure everything fits
    return fig

def create_chl_legend_figure_vertical():
    levels = [0, 6, 12, 20, 30, 50] # Example levels
    colors = ["#496FF2", "#82D35F", "#FEFD05", "#FD0004", "#8E2026", "#D97CF5"] # Example colors
    cmap = mcolors.LinearSegmentedColormap.from_list(
        "ChlLegendVertical", # Unique name
        list(zip(np.linspace(0, 1, len(levels)), colors))
    )
    norm = mcolors.Normalize(vmin=levels[0], vmax=levels[-1])
    
    fig, ax = plt.subplots(figsize=(1.2, 6)) # Narrow and tall
    fig.subplots_adjust(left=0.4, right=0.6, top=0.95, bottom=0.05) # Adjust margins
    
    sm = plt.cm.ScalarMappable(cmap=cmap, norm=norm)
    sm.set_array([]) # Important for ScalarMappable to work correctly
    
    cbar = fig.colorbar(sm, cax=ax, orientation="vertical", ticks=levels)
    cbar.ax.set_yticklabels([str(l) for l in levels]) # Set Y-axis tick labels
    cbar.set_label("Chlorophyll‑a (mg/m³)", rotation=270, labelpad=15) # Rotate label
    
    # plt.tight_layout() # May not be ideal with manual subplot_adjust
    return fig
# -----------------------------------------------------------------------------
# Βοηθητική Συνάρτηση για Επιλογή φακέλου δεδομένων
# -----------------------------------------------------------------------------
def get_data_folder(waterbody: str, index: str) -> str:
    try:
        base_dir = os.path.dirname(os.path.abspath(__file__)) if "__file__" in globals() else os.getcwd()
    except NameError: # Fallback for environments where __file__ is not defined
        base_dir = os.getcwd()
    debug("DEBUG: Τρέχων φάκελος:", base_dir)
    
    waterbody_map = {
        "Γαδουρά": "Gadoura"
        # Add other waterbodies here if needed
    }
    waterbody_folder_name = waterbody_map.get(waterbody)
    if not waterbody_folder_name:
        st.error(f"Δεν βρέθηκε αντιστοίχιση φακέλου για το υδάτινο σώμα: {waterbody}")
        return None

    # Map index to subfolder name, case-sensitive or however your folders are named
    # Example: "Χλωροφύλλη" maps to "Chlorophyll" folder
    index_folder_map = {
        "Χλωροφύλλη": "Chlorophyll",
        "Πραγματικό": "Pragmatiko", # Or "Πραγματικό" if folder name is in Greek
        "Θολότητα": "Turbidity"   # Or "Θολότητα"
    }
    index_folder_name = index_folder_map.get(index, index) # Fallback to index name if no specific mapping

    data_folder = os.path.join(base_dir, waterbody_folder_name, index_folder_name)
    
    debug("DEBUG: Ο φάκελος δεδομένων επιλύθηκε σε:", data_folder)
    if not os.path.exists(data_folder):
        st.error(f"Ο φάκελος δεν υπάρχει: {data_folder}")
        return None
    return data_folder
# -----------------------------------------------------------------------------
# Εξαγωγή ημερομηνίας από όνομα αρχείου
# -----------------------------------------------------------------------------
def extract_date_from_filename(filename: str): # Already defined, ensure it's used consistently
    basename = os.path.basename(filename)
    # Try YYYY_MM_DD or YYYY-MM-DD first
    match = re.search(r'(\d{4})[_-](\d{2})[_-](\d{2})', basename)
    if not match: # Then try YYYYMMDD
        match = re.search(r'(\d{4})(\d{2})(\d{2})', basename)
    
    if match:
        year, month, day = match.groups()
        try:
            date_obj = datetime(int(year), int(month), int(day))
            day_of_year = date_obj.timetuple().tm_yday
            return day_of_year, date_obj
        except ValueError as e:
            debug(f"DEBUG: Σφάλμα μετατροπής ημερομηνίας για {basename}: {e}")
            return None, None
    return None, None
# -----------------------------------------------------------------------------
# Βοηθητικές Συναρτήσεις για Εξαγωγή Δεδομένων και Επεξεργασία Εικόνας
# -----------------------------------------------------------------------------
# load_lake_shape_from_xml, read_image, load_data are from the older user script.
# If they are needed for "Επιφανειακή Αποτύπωση" (Lake Processing), they should be here.
# For now, I'll assume they are part of run_lake_processing_app or similar if that mode is complex.
# If run_lake_processing_app is simple or a placeholder, these might not be immediately used.

def load_lake_shape_from_xml(xml_file: str, bounds: tuple = None,
                             xml_width: float = 518.0, xml_height: float = 505.0):
    debug("DEBUG: Φόρτωση περιγράμματος από:", xml_file)
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
            st.warning("Δεν βρέθηκαν σημεία στο XML:", xml_file)
            return None
        if bounds is not None: # Transform points if bounds are given
            minx, miny, maxx, maxy = bounds
            transformed_points = []
            for x_xml, y_xml in points:
                x_geo = minx + (x_xml / xml_width) * (maxx - minx)
                y_geo = maxy - (y_xml / xml_height) * (maxy - miny) # Y is often inverted
                transformed_points.append([x_geo, y_geo])
            points = transformed_points
        if points and (points[0] != points[-1]): # Close polygon if not closed
            points.append(points[0])
        debug("DEBUG: Φορτώθηκαν", len(points), "σημεία.")
        return {"type": "Polygon", "coordinates": [points]}
    except Exception as e:
        st.error(f"Σφάλμα φόρτωσης περιγράμματος από {xml_file}: {e}")
        return None

def read_image(file_path: str, lake_shape: dict = None): # For single band processing
    debug("DEBUG: Ανάγνωση εικόνας από:", file_path)
    with rasterio.open(file_path) as src:
        img = src.read(1).astype(np.float32) # Read first band
        profile = src.profile.copy()
        profile.update(dtype="float32") # Ensure profile matches float32 type
        
        no_data_value = src.nodata
        if no_data_value is not None:
            img = np.where(img == no_data_value, np.nan, img)
        # img = np.where(img == 0, np.nan, img) # Optional: treat 0 as NaN
        
        if lake_shape is not None:
            from rasterio.features import geometry_mask # Import here to keep it local
            # Ensure mask is for valid geometries and transform is correct
            poly_mask = geometry_mask([lake_shape], 
                                      transform=src.transform, 
                                      invert=False, # Mask cells outside polygon
                                      out_shape=img.shape)
            img = np.where(~poly_mask, img, np.nan) # Keep data inside polygon (invert=False means True where outside)
    return img, profile

def load_data(input_folder: str, shapefile_name="shapefile.xml"): # For Lake Processing
    debug("DEBUG: load_data καλεσμένη με:", input_folder)
    if not os.path.exists(input_folder):
        st.error(f"Ο φάκελος δεν υπάρχει: {input_folder}")
        raise FileNotFoundError(f"Ο φάκελος δεν υπάρχει: {input_folder}") # Raise error to stop
        
    shapefile_path_xml = os.path.join(input_folder, shapefile_name)
    # Allow for shapefile.txt as well if it's an alternative format you use
    # shapefile_path_txt = os.path.join(input_folder, "shapefile.txt") 
    
    lake_shape = None
    if os.path.exists(shapefile_path_xml):
        # Need to get bounds for load_lake_shape_from_xml if it transforms
        # This requires opening a sample TIF first
        all_tif_files_temp = sorted(glob.glob(os.path.join(input_folder, "*.tif")))
        sample_tif_for_bounds = next((f for f in all_tif_files_temp if os.path.basename(f).lower() != "mask.tif"), None)
        if sample_tif_for_bounds:
            with rasterio.open(sample_tif_for_bounds) as src_temp:
                bounds_temp = src_temp.bounds
            lake_shape = load_lake_shape_from_xml(shapefile_path_xml, bounds=bounds_temp)
        else:
            debug("DEBUG: Δεν βρέθηκε sample TIF για φόρτωση ορίων για το shapefile.")
    else:
        debug("DEBUG: Δεν βρέθηκε XML/TXT περιγράμματος στον φάκελο", input_folder)

    all_tif_files = sorted(glob.glob(os.path.join(input_folder, "*.tif")))
    # Filter out any mask.tif file if it exists
    tif_files = [fp for fp in all_tif_files if os.path.basename(fp).lower() != "mask.tif"]
    
    if not tif_files:
        st.error("Δεν βρέθηκαν GeoTIFF αρχεία στον φάκελο.")
        raise FileNotFoundError("Δεν βρέθηκαν GeoTIFF αρχεία.")

    images, days_of_year_list, date_obj_list = [], [], []
    for file_path in tif_files:
        day_of_year, date_obj = extract_date_from_filename(file_path)
        if day_of_year is None or date_obj is None:
            debug(f"Παράλειψη αρχείου λόγω αποτυχίας εξαγωγής ημερομηνίας: {file_path}")
            continue
        
        img, _ = read_image(file_path, lake_shape=lake_shape) # Assumes read_image handles single band
        images.append(img)
        days_of_year_list.append(day_of_year)
        date_obj_list.append(date_obj)
        
    if not images:
        st.error("Δεν βρέθηκαν έγκυρες εικόνες για επεξεργασία.")
        raise ValueError("Δεν φορτώθηκαν εικόνες.")
        
    image_stack = np.stack(images, axis=0)
    return image_stack, np.array(days_of_year_list), date_obj_list
# -----------------------------------------------------------------------------
# Επεξεργασία Λίμνης (Lake Processing) - "Επιφανειακή Αποτύπωση"
# -----------------------------------------------------------------------------
def run_lake_processing_app(waterbody: str, index: str): # Placeholder, needs full implementation
    with st.container():
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.title(f"Επιφανειακή Αποτύπωση ({waterbody} - {index})")

        data_folder = get_data_folder(waterbody, index)
        if data_folder is None:
            # Error already shown by get_data_folder
            st.stop()

        input_folder_geotiffs = os.path.join(data_folder, "GeoTIFFs") # Assuming GeoTIFFs are in a subfolder
        if not os.path.exists(input_folder_geotiffs):
            st.error(f"Ο υποφάκελος 'GeoTIFFs' δεν βρέθηκε στο {data_folder}")
            st.stop()
            
        try:
            # Note: load_data expects shapefile in input_folder_geotiffs if used
            STACK, DAYS, DATES = load_data(input_folder_geotiffs) 
        except Exception as e:
            st.error(f"Σφάλμα φόρτωσης δεδομένων για Επιφανειακή Αποτύπωση: {e}")
            st.stop()

        if not DATES or STACK is None: # Check if DATES is empty or STACK is None
            st.error("Δεν φορτώθηκαν δεδομένα ή πληροφορίες ημερομηνίας.")
            st.stop()
        
        st.info(f"Φορτώθηκαν {STACK.shape[0]} εικόνες για την περίοδο από {min(DATES).date()} έως {max(DATES).date()}.")
        
        # --- Filters from Sidebar (Example based on previous version) ---
        min_date_data = min(DATES)
        max_date_data = max(DATES)
        
        st.sidebar.header(f"Φίλτρα ({waterbody} - Επιφ. Αποτύπωση)")
        threshold_val_range = st.sidebar.slider("Εύρος τιμών pixel (0-255)", 0, 255, (10, 200), key="thresh_lake_proc")
        
        # Date range sliders
        selected_date_range_lake = st.sidebar.slider(
            "Επιλογή περιόδου", 
            min_value=min_date_data.date(), 
            max_value=max_date_data.date(),
            value=(min_date_data.date(), max_date_data.date()), 
            key="date_range_lake_proc"
        )
        start_date_dt, end_date_dt = selected_date_range_lake
        start_datetime = datetime.combine(start_date_dt, datetime.min.time())
        end_datetime = datetime.combine(end_date_dt, datetime.max.time())


        # Filter STACK and DATES based on selected_date_range_lake
        # This needs to be done carefully if DATES contains datetime objects
        date_indices_to_keep = [
            i for i, d_obj in enumerate(DATES) 
            if start_datetime <= d_obj <= end_datetime
        ]

        if not date_indices_to_keep:
            st.warning("Δεν υπάρχουν δεδομένα για την επιλεγμένη περίοδο.")
            st.stop()

        stack_time_filtered = STACK[date_indices_to_keep, :, :]
        dates_time_filtered = [DATES[i] for i in date_indices_to_keep]
        # DAYS_time_filtered = DAYS[date_indices_to_keep] # if DAYS is also needed

        # Apply threshold
        lower_thresh_val, upper_thresh_val = threshold_val_range
        pixels_in_range_mask = np.logical_and(stack_time_filtered >= lower_thresh_val, stack_time_filtered <= upper_thresh_val)
        
        # --- Example Plots (based on previous version's logic) ---
        
        # 1. "Ημέρες σε Εύρος" (Days in Range)
        days_pixel_in_range = np.nansum(pixels_in_range_mask, axis=0)
        fig_days_in_range = px.imshow(days_pixel_in_range, color_continuous_scale="viridis",
                                      title="Ημέρες που κάθε pixel ήταν εντός του επιλεγμένου εύρους τιμών")
        st.plotly_chart(fig_days_in_range, use_container_width=True)

        # 2. "Μέσο Δείγμα Εικόνας" (Average Sample Image after filtering)
        # Apply mask to stack before mean: where pixels_in_range_mask is False, set to NaN
        stack_for_mean = np.where(pixels_in_range_mask, stack_time_filtered, np.nan)
        average_image_filtered = np.nanmean(stack_for_mean, axis=0)
        
        fig_avg_image = px.imshow(average_image_filtered, color_continuous_scale="jet",
                                  title="Μέση εικόνα (pixels εντός εύρους και περιόδου)")
        st.plotly_chart(fig_avg_image, use_container_width=True)


        st.info("Η λειτουργία 'Επιφανειακή Αποτύπωση' είναι υπό ανάπτυξη. Εμφανίζονται βασικά παραδείγματα.")
        st.markdown('</div>', unsafe_allow_html=True)


# --------------------------------------------------------------------------
# Helper function for image processing in dashboard
# --------------------------------------------------------------------------
def process_and_enhance_geotiff_for_display(image_path_to_process):
    try:
        with rasterio.open(image_path_to_process) as src:
            if src.count >= 3: # Needs at least 3 bands for RGB
                # Assuming bands 1,2,3 are R,G,B. Adjust if necessary.
                # e.g., for Sentinel-2 True Color: use bands [4,3,2] for [R,G,B]
                img_bands_raw = src.read([1, 2, 3]) 

                scaled_bands_for_rgb = []
                for i in range(img_bands_raw.shape[0]): # Process each of the 3 bands
                    band_data = img_bands_raw[i, :, :].astype(np.float32)
                    nodata_val = src.nodatavals[i] if src.nodatavals and src.nodatavals[i] is not None else None
                    
                    band_data_for_percentile = band_data.copy() # Work on a copy
                    if nodata_val is not None:
                        if not np.isnan(nodata_val):
                            band_data_for_percentile[band_data_for_percentile == nodata_val] = np.nan
                        # If nodata_val is already NaN, it's fine for nanpercentile

                    min_p, max_p = np.nanpercentile(band_data_for_percentile, [2, 98])
                    
                    if max_p <= min_p: # Handle flat bands or all-NaN bands after masking
                        band_stretched = np.zeros_like(band_data, dtype=np.uint8)
                    else:
                        band_stretched = (band_data - min_p) / (max_p - min_p)
                        band_stretched = np.clip(band_stretched, 0, 1)
                        band_stretched = (band_stretched * 255).astype(np.uint8)
                    
                    # Set original NoData pixels to black (0) in the 8-bit image
                    if nodata_val is not None:
                        if not np.isnan(nodata_val):
                             band_stretched[band_data == nodata_val] = 0 
                        else:
                             band_stretched[np.isnan(band_data)] = 0 # If original nodata was NaN

                    scaled_bands_for_rgb.append(band_stretched)

                if len(scaled_bands_for_rgb) == 3:
                    img_rgb_8bit = np.transpose(np.stack(scaled_bands_for_rgb, axis=0), (1, 2, 0))
                else: return None

                R, G, B = img_rgb_8bit[:, :, 0], img_rgb_8bit[:, :, 1], img_rgb_8bit[:, :, 2]
                
                intensity_min_thresh, intensity_max_thresh = 160, 230
                max_channel_difference = 40
                
                pale_intensity_mask = (R >= intensity_min_thresh) & (R <= intensity_max_thresh) & \
                                      (G >= intensity_min_thresh) & (G <= intensity_max_thresh) & \
                                      (B >= intensity_min_thresh) & (B <= intensity_max_thresh)
                
                rgb_max_ch_vals = np.maximum(np.maximum(R, G), B)
                rgb_min_ch_vals = np.minimum(np.minimum(R, G), B)
                low_saturation_mask = (rgb_max_ch_vals - rgb_min_ch_vals) < max_channel_difference
                
                final_anomaly_mask = pale_intensity_mask & low_saturation_mask
                
                img_enhanced = img_rgb_8bit.copy()
                img_enhanced[final_anomaly_mask] = [255, 255, 0] # Highlight with Yellow
                return img_enhanced

            elif src.count == 1: # Grayscale for single-band
                band_data = src.read(1).astype(np.float32)
                nodata_val = src.nodata if src.nodata is not None else None
                band_data_for_percentile = band_data.copy()
                if nodata_val is not None:
                    if not np.isnan(nodata_val): band_data_for_percentile[band_data_for_percentile == nodata_val] = np.nan
                
                min_p, max_p = np.nanpercentile(band_data_for_percentile, [2, 98])
                if max_p <= min_p: band_8bit = np.zeros_like(band_data, dtype=np.uint8)
                else:
                    band_stretched = np.clip((band_data - min_p) / (max_p - min_p), 0, 1)
                    band_8bit = (band_stretched * 255).astype(np.uint8)
                if nodata_val is not None: # Set NoData to black
                    if not np.isnan(nodata_val): band_8bit[band_data == nodata_val] = 0
                    else: band_8bit[np.isnan(band_data)] = 0
                return band_8bit
            return None # Not 3-band or 1-band
    except Exception as e:
        st.error(f"Σφάλμα επεξεργασίας εικόνας {os.path.basename(image_path_to_process)}: {e}")
        return None
# -----------------------------------------------------------------------------
# Πίνακας Ποιότητας Ύδατος - "Προφίλ ποιότητας και στάθμης"
# -----------------------------------------------------------------------------
def run_water_quality_dashboard(waterbody: str, index: str): # Main focus for image enhancement
    with st.container():
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.title(f"Προφίλ Ποιότητας και Στάθμης ({waterbody} - {index})")

        data_folder = get_data_folder(waterbody, index)
        if data_folder is None: st.stop()

        images_folder = os.path.join(data_folder, "GeoTIFFs")
        if not os.path.exists(images_folder):
            st.error(f"Ο φάκελος GeoTIFFs δεν βρέθηκε στο {data_folder}")
            st.stop()
            
        lake_height_path = os.path.join(data_folder, "lake height.xlsx") # Used by analyze_sampling
        sampling_kml_path = os.path.join(data_folder, "sampling.kml") # For default points

        # Prepare list of GeoTIFFs and their dates for image selection carousel
        tif_files_list = [f for f in os.listdir(images_folder) if f.lower().endswith(('.tif', '.tiff'))]
        available_dates_for_carousel = {}
        for f_name in tif_files_list:
            _, date_obj = extract_date_from_filename(f_name) # Use your existing robust function
            if date_obj:
                available_dates_for_carousel[str(date_obj.date())] = f_name
        
        # Background GeoTIFF for fig_geo (from analyze_sampling)
        # Needs first_image_data and first_transform
        # Let's pick the most recent image as default background, or the first if none selected
        first_image_data, first_transform = None, None
        if available_dates_for_carousel:
            # Default background: most recent image from the carousel list
            default_bg_date_str = sorted(available_dates_for_carousel.keys())[-1]
            default_bg_filename = available_dates_for_carousel[default_bg_date_str]
            default_bg_path = os.path.join(images_folder, default_bg_filename)
            try:
                with rasterio.open(default_bg_path) as src_bg:
                    if src_bg.count >= 3:
                        first_image_data = src_bg.read([1,2,3]) # Assuming R,G,B like
                        first_transform = src_bg.transform
                    else:
                        st.warning(f"Default background {default_bg_filename} has <3 bands.")
            except Exception as e:
                st.error(f"Error loading default background image: {e}")
        
        if first_image_data is None or first_transform is None: # Fallback if loading default failed
            st.error("Δεν ήταν δυνατή η φόρτωση μιας βασικής εικόνας GeoTIFF για το υπόβαθρο των σημείων. Ορισμένα διαγράμματα ενδέχεται να μην λειτουργούν.")
            # Allow to proceed but fig_geo might be empty or error out in analyze_sampling

        # Sidebar filters for date range, specific to this dashboard
        st.sidebar.markdown("---")
        st.sidebar.subheader(f"Φίλτρα για '{index}' ({waterbody})")
        date_filter_min = st.sidebar.date_input("Από ημερομηνία:", date(2015,1,1), key=f"date_min_{index}")
        date_filter_max = st.sidebar.date_input("Έως ημερομηνία:", date.today(), key=f"date_max_{index}")


        # Tabs for Default and Upload Sampling
        tab_names = ["Δειγματοληψία (Προεπιλογή)", "Δειγματοληψία (Ανέβασμα KML)"]
        sampling_tabs = st.tabs(tab_names)

        for i, tab_mode in enumerate(["Default", "Upload"]):
            with sampling_tabs[i]:
                st.subheader(f"Ανάλυση με {tab_mode} σημεία")
                current_sampling_points = []
                unique_key_suffix = f"_{tab_mode.lower()}"

                if tab_mode == "Default":
                    if os.path.exists(sampling_kml_path):
                        current_sampling_points = parse_sampling_kml(sampling_kml_path)
                    else:
                        st.warning(f"Δεν βρέθηκε το προεπιλεγμένο αρχείο KML: {sampling_kml_path}")
                else: # Upload mode
                    uploaded_kml_file = st.file_uploader("Ανεβάστε το KML αρχείο σας", type="kml", key=f"kml_upload{unique_key_suffix}")
                    if uploaded_kml_file:
                        current_sampling_points = parse_sampling_kml(uploaded_kml_file)
                
                if not current_sampling_points:
                    st.info(f"Δεν έχουν οριστεί/φορτωθεί σημεία δειγματοληψίας για τη λειτουργία {tab_mode}.")
                    continue # Skip to next tab if no points

                all_point_names = [pt[0] for pt in current_sampling_points]
                selected_analysis_points = st.multiselect(
                    "Επιλογή σημείων για ανάλυση:", 
                    options=all_point_names, 
                    default=all_point_names, 
                    key=f"select_pts{unique_key_suffix}"
                )

                if st.button(f"Εκτέλεση Ανάλυσης ({tab_mode})", key=f"run_analysis{unique_key_suffix}"):
                    if not selected_analysis_points:
                        st.error("Παρακαλώ επιλέξτε τουλάχιστον ένα σημείο για ανάλυση.")
                    elif first_image_data is None or first_transform is None:
                         st.error("Σφάλμα: Δεν έχει οριστεί βασική εικόνα GeoTIFF (first_image_data).")
                    else:
                        with st.spinner("Επεξεργασία δεδομένων... Παρακαλώ περιμένετε."):
                            session_results_key = f"results{unique_key_suffix}"
                            st.session_state[session_results_key] = analyze_sampling(
                                sampling_points=current_sampling_points, # Pass full list for y-axis consistency
                                first_image_data=first_image_data,
                                first_transform=first_transform,
                                images_folder=images_folder,
                                lake_height_path=lake_height_path,
                                selected_points=selected_analysis_points, # Pass selected for filtering inside
                                date_min=date_filter_min, # Pass date filters
                                date_max=date_filter_max
                            )
                
                session_results_key = f"results{unique_key_suffix}"
                if session_results_key in st.session_state and st.session_state[session_results_key]:
                    res_fig_geo, res_fig_dual, res_fig_colors, res_fig_mg, _, _, _ = st.session_state[session_results_key]
                    
                    # Nested tabs for results
                    result_tabs = st.tabs([
                        "🗺️ GeoTIFF & Σημεία", "🖼️ Επιλογή Εικόνας (Επεξεργ.)", 
                        "🎬 Video/GIF", "📊 Χρώματα Pixel & Στάθμη", 
                        "🧪 Μέσο mg/m³", "📈 Διπλά Διαγράμματα"
                    ])
                    
                    with result_tabs[0]: # GeoTIFF & Points
                        st.plotly_chart(res_fig_geo, use_container_width=True, key=f"geo{unique_key_suffix}")
                        if index == "Χλωροφύλλη":
                            st.pyplot(create_chl_legend_figure())


                    with result_tabs[1]: # Enhanced Image Selection
                        st.markdown("#### Επιλογή και Εμφάνιση Επεξεργασμένης Εικόνας GeoTIFF")
                        if available_dates_for_carousel:
                            sorted_carousel_dates = sorted(available_dates_for_carousel.keys())
                            
                            img_idx_key = f"img_idx{unique_key_suffix}"
                            if img_idx_key not in st.session_state:
                                st.session_state[img_idx_key] = len(sorted_carousel_dates) - 1 # Default to most recent

                            # Image Navigation Buttons
                            cols_nav = st.columns([1,8,1])
                            if cols_nav[0].button("⬅️ Προηγούμενη", key=f"prev_img{unique_key_suffix}"):
                                st.session_state[img_idx_key] = max(0, st.session_state[img_idx_key] - 1)
                            if cols_nav[2].button("Επόμενη ➡️", key=f"next_img{unique_key_suffix}"):
                                st.session_state[img_idx_key] = min(len(sorted_carousel_dates)-1, st.session_state[img_idx_key] + 1)
                            
                            selected_date_str = cols_nav[1].selectbox(
                                "Επιλογή Ημερομηνίας Εικόνας:", 
                                options=sorted_carousel_dates, 
                                index=st.session_state[img_idx_key],
                                key=f"sel_date_img{unique_key_suffix}"
                            )
                            st.session_state[img_idx_key] = sorted_carousel_dates.index(selected_date_str) # Update index

                            img_filename = available_dates_for_carousel[selected_date_str]
                            img_path = os.path.join(images_folder, img_filename)

                            st.markdown(f"**Εμφανίζεται η εικόνα για:** `{selected_date_str}` (`{img_filename}`)")
                            enhanced_img_array = process_and_enhance_geotiff_for_display(img_path)
                            if enhanced_img_array is not None:
                                st.image(enhanced_img_array, caption=f"Επεξεργασμένη: {selected_date_str}", use_column_width=True)
                            else: # Fallback if processing fails
                                st.warning("Η επεξεργασία απέτυχε. Εμφανίζεται η αρχική εικόνα.")
                                if os.path.exists(img_path):
                                     st.image(img_path, caption=f"Αρχική: {selected_date_str}", use_column_width=True)
                                else:
                                     st.error("Το αρχείο εικόνας δεν βρέθηκε.")

                        else:
                            st.info("Δεν υπάρχουν διαθέσιμες εικόνες στον φάκελο GeoTIFFs για επιλογή.")
                        if index == "Χλωροφύλλη":
                            st.pyplot(create_chl_legend_figure())


                    with result_tabs[2]: # Video/GIF
                        # Your existing video_path logic - assuming video_path is defined globally or passed
                        video_path = None # Placeholder, define how video_path is determined
                        possible_video_paths = [
                            os.path.join(data_folder, "timelapse.mp4"),
                            os.path.join(images_folder, "timelapse.mp4"), # Common place
                            os.path.join(data_folder, "animation.gif"),
                            os.path.join(images_folder, "animation.gif"),
                        ]
                        for p_vid in possible_video_paths:
                            if os.path.exists(p_vid):
                                video_path = p_vid
                                break
                        if video_path:
                            if video_path.endswith(".mp4"): st.video(video_path)
                            else: st.image(video_path)
                        else: st.info("Δεν βρέθηκε αρχείο Video/GIF.")
                        if index == "Χλωροφύλλη":
                            st.pyplot(create_chl_legend_figure())
                            

                    with result_tabs[3]: # Pixel Colors & Lake Height
                        cols_chart_legend = st.columns([0.85, 0.15]) # 비율 조정
                        with cols_chart_legend[0]:
                            res_fig_colors.update_layout(height=500) # Adjust height if needed
                            st.plotly_chart(res_fig_colors, use_container_width=True, key=f"colors{unique_key_suffix}")
                        with cols_chart_legend[1]:
                            if index == "Χλωροφύλλη":
                                st.pyplot(create_chl_legend_figure_vertical())
                                
                    with result_tabs[4]: # Mean MG
                        st.plotly_chart(res_fig_mg, use_container_width=True, key=f"mg{unique_key_suffix}")
                        
                    with result_tabs[5]: # Dual Plot
                        st.plotly_chart(res_fig_dual, use_container_width=True, key=f"dual{unique_key_suffix}")
        
        st.info("Τέλος Πίνακα Ποιότητας Ύδατος.")
        st.markdown('</div>', unsafe_allow_html=True)

# --------------------------------------------------------------------------
# Predictive Tools with on-the-fly re-calculation for all indices
# --------------------------------------------------------------------------
def run_predictive_tools(waterbody: str, current_selected_index: str): # Pass current_selected_index
    with st.container():
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.title(f"Εργαλεία Πρόβλεψης & Έγκαιρης Ενημέρωσης ({waterbody})")

        chart_options = ["GeoTIFF Background", "Pixel Colors & Lake Height", "Mean mg/m³", "Lake Height & Mean mg/m³"]
        selected_charts_to_recalc = st.multiselect(
            "Επιλέξτε διαγράμματα για επανυπολογισμό με νέα φίλτρα:",
            options=chart_options,
            default=chart_options,
            key="predict_charts_select"
        )

        st.markdown("##### Παράμετροι Φιλτραρίσματος για Επανυπολογισμό")
        col_filter1, col_filter2 = st.columns(2)
        with col_filter1:
            # Date range for re-calculation
            predict_date_min = st.date_input("Ημερομηνία Από:", value=date(2015, 1, 1), key="predict_date_min")
        with col_filter2:
            predict_date_max = st.date_input("Ημερομηνία Έως:", value=date.today(), key="predict_date_max")

        sampling_type_predict = st.radio("Σύνολο Σημείων Δειγματοληψίας:", ["Default KML", "Upload New KML"], key="predict_kml_type")
        
        uploaded_kml_predict = None
        if sampling_type_predict == "Upload New KML":
            uploaded_kml_predict = st.file_uploader("Ανεβάστε KML για πρόβλεψη:", type="kml", key="predict_kml_upload")

        # Iterate through all available indices for recalculation
        indices_to_process = ["Χλωροφύλλη", "Πραγματικό", "Θολότητα"] # As defined in sidebar
        
        if st.button("Επανυπολογισμός Διαγραμμάτων", key="predict_recalc_button"):
            for idx_predict in indices_to_process:
                with st.expander(f"Αποτελέσματα για Δείκτη: {idx_predict}", expanded=(idx_predict == current_selected_index)):
                    
                    data_folder_predict = get_data_folder(waterbody, idx_predict)
                    if not data_folder_predict:
                        st.error(f"Δεν βρέθηκε φάκελος δεδομένων για τον δείκτη {idx_predict}.")
                        continue

                    images_folder_predict  = os.path.join(data_folder_predict, "GeoTIFFs")
                    lake_height_xl_predict = os.path.join(data_folder_predict, "lake height.xlsx")
                    default_kml_path_predict = os.path.join(data_folder_predict, "sampling.kml")

                    kml_source_predict = None
                    if sampling_type_predict == "Default KML":
                        if os.path.exists(default_kml_path_predict):
                            kml_source_predict = default_kml_path_predict
                        else:
                            st.error(f"Το προεπιλεγμένο KML δεν βρέθηκε για {idx_predict}: {default_kml_path_predict}")
                            continue
                    elif uploaded_kml_predict: # Upload New KML and file is uploaded
                        kml_source_predict = uploaded_kml_predict
                    else: # Upload New KML but no file uploaded
                        st.warning(f"Παρακαλώ ανεβάστε ένα αρχείο KML για τον δείκτη {idx_predict} εάν επιλέξατε 'Upload New KML'.")
                        continue
                    
                    sampling_points_predict = parse_sampling_kml(kml_source_predict)
                    if not sampling_points_predict:
                        st.warning(f"Δεν βρέθηκαν σημεία δειγματοληψίας από το KML για τον δείκτη {idx_predict}.")
                        continue

                    # Load a representative GeoTIFF for background in fig_geo
                    # (analyze_sampling needs first_image_data & first_transform)
                    first_img_pred, first_tr_pred = None, None
                    tif_files_pred = sorted(glob.glob(os.path.join(images_folder_predict, "*.tif"))) # Look for .tif
                    tif_files_pred.extend(sorted(glob.glob(os.path.join(images_folder_predict, "*.tiff")))) # Add .tiff

                    # Filter tif_files_pred by date_min, date_max to pick a relevant background
                    relevant_tifs_for_bg = []
                    for tf_path in tif_files_pred:
                        _, tf_date_obj = extract_date_from_filename(os.path.basename(tf_path))
                        if tf_date_obj and (predict_date_min <= tf_date_obj.date() <= predict_date_max):
                            relevant_tifs_for_bg.append(tf_path)
                    
                    bg_tif_path_predict = relevant_tifs_for_bg[-1] if relevant_tifs_for_bg else (tif_files_pred[-1] if tif_files_pred else None)


                    if bg_tif_path_predict and os.path.exists(bg_tif_path_predict):
                        with rasterio.open(bg_tif_path_predict) as src_pred_bg:
                            if src_pred_bg.count >= 3:
                                first_img_pred = src_pred_bg.read([1,2,3])
                                first_tr_pred = src_pred_bg.transform
                            else:
                                st.error(f"Η εικόνα GeoTIFF {os.path.basename(bg_tif_path_predict)} δεν έχει 3 κανάλια.")
                                continue
                    else:
                        st.error(f"Δεν βρέθηκε κατάλληλη εικόνα GeoTIFF για υπόβαθρο για τον δείκτη {idx_predict}.")
                        continue
                    
                    try:
                        with st.spinner(f"Επεξεργασία {idx_predict}..."):
                            pred_fig_geo, pred_fig_dual, pred_fig_colors, pred_fig_mg, _, _, _ = analyze_sampling(
                                sampling_points=sampling_points_predict,
                                first_image_data=first_img_pred,
                                first_transform=first_tr_pred,
                                images_folder=images_folder_predict,
                                lake_height_path=lake_height_xl_predict,
                                selected_points=[pt[0] for pt in sampling_points_predict], # Analyze all parsed points
                                date_min=predict_date_min, # Pass new date filters
                                date_max=predict_date_max
                            )
                    except Exception as e_anls:
                        st.error(f"Σφάλμα κατά την ανάλυση για τον δείκτη {idx_predict}: {e_anls}")
                        continue

                    # Display selected charts
                    if "GeoTIFF Background" in selected_charts_to_recalc:
                        st.subheader(f"GeoTIFF & Σημεία ({idx_predict})")
                        st.plotly_chart(pred_fig_geo, use_container_width=True, key=f"pred_geo_{idx_predict}")
                    if "Pixel Colors & Lake Height" in selected_charts_to_recalc:
                        st.subheader(f"Χρώματα Pixel & Στάθμη ({idx_predict})")
                        st.plotly_chart(pred_fig_colors, use_container_width=True, key=f"pred_colors_{idx_predict}")
                    if "Mean mg/m³" in selected_charts_to_recalc:
                        st.subheader(f"Μέσο mg/m³ ({idx_predict})")
                        st.plotly_chart(pred_fig_mg, use_container_width=True, key=f"pred_mg_{idx_predict}")
                    if "Lake Height & Mean mg/m³" in selected_charts_to_recalc:
                        st.subheader(f"Στάθμη & Μέσο mg/m³ ({idx_predict})")
                        st.plotly_chart(pred_fig_dual, use_container_width=True, key=f"pred_dual_{idx_predict}")
                        
        st.markdown('</div>', unsafe_allow_html=True)


# -----------------------------------------------------------------------------
# Main Entry Point
# -----------------------------------------------------------------------------
def main():
    # Inject custom CSS - should be called once
    inject_custom_css() 
    
    run_intro_page_new() # Use the new intro page
    run_custom_ui_new()  # Use the new UI for sidebar

    waterbody_selected = st.session_state.get("waterbody_choice")
    index_selected = st.session_state.get("index_choice")
    analysis_selected = st.session_state.get("analysis_choice")

    if waterbody_selected == "Γαδουρά": # Assuming only Gadoura is configured for now
        if analysis_selected == "Επιφανειακή Αποτύπωση":
            # This corresponds to the old "Lake Processing"
            run_lake_processing_app(waterbody_selected, index_selected)
        elif analysis_selected == "Προφίλ ποιότητας και στάθμης":
            # This corresponds to the "Water Quality Dashboard"
            run_water_quality_dashboard(waterbody_selected, index_selected)
        elif analysis_selected == "Eργαλεία Πρόβλεψης και έγκαιρης ενημέρωσης":
            run_predictive_tools(waterbody_selected, index_selected) # Pass current index for default expansion
        else:
            st.info("Παρακαλώ επιλέξτε ένα έγκυρο είδος ανάλυσης από την πλαϊνή μπάρα.")
    # elif analysis_selected == "Burned Areas": # Kept for reference if needed later
    #     if waterbody_selected == "Γαδουρά":
    #         # run_burned_areas() # This function was a placeholder
    #         st.info("Η ανάλυση καμένων περιοχών δεν είναι ακόμα διαθέσιμη.")
    #     else:
    #         st.warning("Η ανάλυση 'Burned Areas' είναι προς το παρόν διαθέσιμη μόνο για το υδάτινο σώμα Γαδουρά.")
    else:
        st.warning(f"Δεν υπάρχουν διαθέσιμα δεδομένα ή αναλύσεις για τον επιλεγμένο συνδυασμό: {waterbody_selected} / {index_selected} / {analysis_selected}.")

    render_footer()

if __name__ == "__main__":
    main()
