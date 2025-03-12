#!/usr/bin/env python
# -*- coding: utf-8 -*-

"""
Subterranean Detection App (Enterprise-Grade UI)
-------------------------------------------------
For Option A, data is read from "folder_a" and for Option B, from "folder_b".
All file paths are constructed absolutely using the location of this script.
Make sure your folder structure is:
  Subterra_2/
      main.py
      folder_a/   <- contains area folders for Option A
      folder_b/   <- contains area folders (e.g., "7", etc.) for Option B
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
# get_data_folder: Build absolute paths using base_dir and chosen methodology.
# -----------------------------------------------------------------------------
def get_data_folder(waterbody: str, index: str) -> str:
    base_dir = os.path.dirname(os.path.abspath(__file__))
    selected_method = st.session_state.get("method_option", "Option A")
    if selected_method == "Option A":
        base_folder = os.path.join(base_dir, "folder_a")
    else:
        base_folder = os.path.join(base_dir, "folder_b")
    
    debug("Base folder being used:", base_folder)
    if not os.path.exists(base_folder):
        st.error("No valid mother folders found with required subfolders.")
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
    elif index == "Πραγματικό" and selected_method != "Option A":
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
    
    # Select methodology; value stored in session_state.
    method_option = st.sidebar.selectbox("Select Methodology", ["Option A", "Option B"], key="method_option")
    if method_option == "Option A":
        chosen_dir = os.path.join(base_dir, "folder_a")
    else:
        chosen_dir = os.path.join(base_dir, "folder_b")
    
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

    # Gather immediate subdirectories.
    area_options = sorted(
        [d for d in os.listdir(chosen_dir) if os.path.isdir(os.path.join(chosen_dir, d))]
    )
    st.write("### DEBUG: Found subdirectories:", area_options)
    
    if method_option == "Option B" and not area_options:
        st.warning("No subdirectories found in folder_b; using default area list.")
        area_options = ["Κορώνεια", "Πολυφύτου", "Γαδουρά", "Αξιός"]

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
# Core Processing Functions (Lake Processing & Dashboard)
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

        # Basic filters from sidebar.
        min_date = min(DATES)
        max_date = max(DATES)
        unique_years = sorted({d.year for d in DATES if d is not None})

        st.sidebar.header(f"Filters (Subterranean Processing: {waterbody})")
        threshold_range = st.sidebar.slider("Pixel Value Range", 0, 255, (0, 255), key="thresh_lp")
        broad_date_range = st.sidebar.slider("General Date Range", min_value=min_date, max_value=max_date,
                                             value=(min_date, max_date), key="broad_date_lp")
        refined_date_range = st.sidebar.slider("Refined Date Range", min_value=min_date, max_value=max_date,
                                               value=(min_date, max_date), key="refined_date_lp")
        display_option = st.sidebar.radio("Display Mode", options=["Thresholded", "Original"], index=0, key="display_lp")

        st.sidebar.markdown("### Select Months")
        month_options = {i: datetime(2000, i, 1).strftime('%B') for i in range(1, 13)}
        if "selected_months" not in st.session_state:
            st.session_state.selected_months = list(month_options.keys())
        selected_months = st.sidebar.multiselect("Months",
                                                 options=list(month_options.keys()),
                                                 format_func=lambda x: month_options[x],
                                                 default=st.session_state.selected_months,
                                                 key="months_lp")
        st.session_state.selected_years = unique_years
        selected_years = st.sidebar.multiselect("Years", options=unique_years,
                                                default=unique_years,
                                                key="years_lp")

        start_dt, end_dt = refined_date_range
        selected_indices = [i for i, d in enumerate(DATES)
                            if start_dt <= d <= end_dt and d.month in selected_months and d.year in selected_years]

        if not selected_indices:
            st.error("No data for the selected date range/months/years.")
            st.stop()

        stack_filtered = STACK[selected_indices, :, :]
        days_filtered = np.array(DAYS)[selected_indices]
        filtered_dates = np.array(DATES)[selected_indices]

        lower_thresh, upper_thresh = threshold_range
        in_range = np.logical_and(stack_filtered >= lower_thresh, stack_filtered <= upper_thresh)

        # "Days in Range" chart.
        days_in_range = np.nansum(in_range, axis=0)
        fig_days = px.imshow(days_in_range, color_continuous_scale="plasma",
                             title="Chart: Days in Range", labels={"color": "Days in Range"})
        fig_days.update_layout(width=800, height=600)
        st.plotly_chart(fig_days, use_container_width=True, key="fig_days")
        with st.expander("Explanation: Days in Range"):
            st.write("This chart shows how many days each pixel falls within the selected pixel value range. Adjust the slider to see changes.")

        tick_vals = [1, 32, 60, 91, 121, 152, 182, 213, 244, 274, 305, 335, 365]
        tick_text = ["1 (Jan)", "32 (Feb)", "60 (Mar)", "91 (Apr)",
                     "121 (May)", "152 (Jun)", "182 (Jul)", "213 (Aug)",
                     "244 (Sep)", "274 (Oct)", "305 (Nov)", "335 (Dec)", "365 (Dec)"]

        days_array = days_filtered.reshape((-1, 1, 1))
        sum_days = np.nansum(days_array * in_range, axis=0)
        count_in_range = np.nansum(in_range, axis=0)
        mean_day = np.divide(sum_days, count_in_range,
                             out=np.full(sum_days.shape, np.nan),
                             where=(count_in_range != 0))
        fig_mean = px.imshow(mean_day, color_continuous_scale="RdBu",
                             title="Chart: Mean Day of Occurrence", labels={"color": "Mean Day"})
        fig_mean.update_layout(width=800, height=600)
        fig_mean.update_layout(coloraxis_colorbar=dict(tickmode='array', tickvals=tick_vals, ticktext=tick_text))
        st.plotly_chart(fig_mean, use_container_width=True, key="fig_mean")
        with st.expander("Explanation: Mean Day of Occurrence"):
            st.write("This chart displays the mean day of occurrence for pixels within the selected range.")

        if display_option.lower() == "thresholded":
            filtered_stack = np.where(in_range, stack_filtered, np.nan)
        else:
            filtered_stack = stack_filtered

        average_sample_img = np.nanmean(filtered_stack, axis=0)
        if not np.all(np.isnan(average_sample_img)):
            avg_min = float(np.nanmin(average_sample_img))
            avg_max = float(np.nanmax(average_sample_img))
        else:
            avg_min, avg_max = 0, 0

        fig_sample = px.imshow(average_sample_img, color_continuous_scale="jet",
                               range_color=[avg_min, avg_max],
                               title="Chart: Mean Sample Image", labels={"color": "Pixel Value"})
        fig_sample.update_layout(width=800, height=600)
        st.plotly_chart(fig_sample, use_container_width=True, key="fig_sample")
        with st.expander("Explanation: Mean Sample Image"):
            st.write("This chart shows the mean pixel value after applying the filter.")

        filtered_day_of_year = np.array([d.timetuple().tm_yday for d in filtered_dates])
        def nanargmax_or_nan(arr):
            return np.nan if np.all(np.isnan(arr)) else np.nanargmax(arr)
        max_index = np.apply_along_axis(nanargmax_or_nan, 0, filtered_stack)
        time_max = np.full(max_index.shape, np.nan, dtype=float)
        valid_mask = ~np.isnan(max_index)
        max_index_int = np.zeros_like(max_index, dtype=int)
        max_index_int[valid_mask] = max_index[valid_mask].astype(int)
        max_index_int[valid_mask] = np.clip(max_index_int[valid_mask], 0, len(filtered_day_of_year) - 1)
        time_max[valid_mask] = filtered_day_of_year[max_index_int[valid_mask]]
        fig_time = px.imshow(time_max, color_continuous_scale="RdBu",
                             range_color=[1, 365],
                             title="Chart: Time of Maximum Occurrence", labels={"color": "Day"})
        fig_time.update_layout(width=800, height=600)
        fig_time.update_layout(coloraxis_colorbar=dict(tickmode='array', tickvals=tick_vals, ticktext=tick_text))
        st.plotly_chart(fig_time, use_container_width=True, key="fig_time")
        with st.expander("Explanation: Time of Maximum Occurrence"):
            st.write("This chart shows the day when each pixel reached its maximum value.")

        st.header("Analysis Maps")
        col1, col2 = st.columns(2)
        with col1:
            st.plotly_chart(fig_days, use_container_width=True, key="fig_days_2")
        with col2:
            st.plotly_chart(fig_mean, use_container_width=True, key="fig_mean_2")

        st.header("Sample Image Analysis")
        col3, col4 = st.columns(2)
        with col3:
            st.plotly_chart(fig_sample, use_container_width=True, key="fig_sample_2")
        with col4:
            st.plotly_chart(fig_time, use_container_width=True, key="fig_time_2")

        # Additional Annual Analysis: Monthly Distribution of Days in Range
        st.header("Additional Annual Analysis: Monthly Distribution of Days in Range")
        stack_full_in_range = (STACK >= lower_thresh) & (STACK <= upper_thresh)
        monthly_days_in_range = {}
        for m in range(1, 13):
            indices_m = [i for i, d in enumerate(DATES) if d is not None and d.month == m]
            if indices_m:
                monthly_days_in_range[m] = np.sum(stack_full_in_range[indices_m, :, :], axis=0)
            else:
                monthly_days_in_range[m] = None

        months_to_display = [m for m in range(1, 13) if m in selected_months]
        months_in_order = sorted(months_to_display)
        if 3 in months_in_order:
            months_in_order = list(range(3, 13)) + [m for m in months_in_order if m < 3]
            seen = set()
            months_in_order = [x for x in months_in_order if not (x in seen or seen.add(x))]
        num_cols = 3
        cols = st.columns(num_cols)
        for idx, m in enumerate(months_in_order):
            col_index = idx % num_cols
            img = monthly_days_in_range[m]
            month_name = datetime(2000, m, 1).strftime('%B')
            if img is not None:
                fig_month = px.imshow(
                    img,
                    color_continuous_scale="plasma",
                    title=month_name,
                    labels={"color": "Days in Range"}
                )
                fig_month.update_layout(width=500, height=400, margin=dict(l=0, r=0, t=30, b=0))
                fig_month.update_coloraxes(showscale=False)
                cols[col_index].plotly_chart(fig_month, use_container_width=False)
            else:
                cols[col_index].info(f"No data for {month_name}")
            if (idx + 1) % num_cols == 0 and (idx + 1) < len(months_in_order):
                cols = st.columns(num_cols)
        with st.expander("Explanation: Monthly Distribution of Days in Range"):
            st.write("For each month, this chart shows how many days each pixel falls within the selected value range.")

        # Additional Annual Analysis: Annual Distribution of Days in Range
        st.header("Additional Annual Analysis: Annual Distribution of Days in Range")
        unique_years_full = sorted({d.year for d in DATES if d is not None})
        years_to_display = [y for y in unique_years_full if y in selected_years]
        if not years_to_display:
            st.error("No valid years available after filtering.")
            st.stop()
        stack_full_in_range = (STACK >= lower_thresh) & (STACK <= upper_thresh)
        yearly_days_in_range = {}
        for year in years_to_display:
            indices_y = [i for i, d in enumerate(DATES) if d.year == year]
            if indices_y:
                yearly_days_in_range[year] = np.sum(stack_full_in_range[indices_y, :, :], axis=0)
            else:
                yearly_days_in_range[year] = None
        num_cols = 3
        cols = st.columns(num_cols)
        for idx, year in enumerate(years_to_display):
            col_index = idx % num_cols
            img = yearly_days_in_range[year]
            if img is not None:
                fig_year = px.imshow(
                    img,
                    color_continuous_scale="plasma",
                    title=f"Year: {year}",
                    labels={"color": "Days in Range"}
                )
                fig_year.update_layout(width=500, height=400, margin=dict(l=0, r=0, t=30, b=0))
                fig_year.update_coloraxes(showscale=False)
                cols[col_index].plotly_chart(fig_year, use_container_width=False)
            else:
                cols[col_index].info(f"No data for {year}")
            if (idx + 1) % num_cols == 0 and (idx + 1) < len(years_to_display):
                cols = st.columns(num_cols)
        with st.expander("Explanation: Annual Distribution of Days in Range"):
            st.write("For each year, this chart shows how many days each pixel falls within the selected value range.")

        st.info("End of Subterranean Processing.")
        st.markdown('</div>', unsafe_allow_html=True)

def run_water_quality_dashboard(waterbody: str, index: str):
    with st.container():
        st.markdown('<div class="card">', unsafe_allow_html=True)
        st.title(f"Subterranean Quality Dashboard ({waterbody} - {index})")
        data_folder = get_data_folder(waterbody, index)
        if data_folder is None:
            st.error("Data folder for the selected area/index does not exist.")
            st.stop()
        images_folder = os.path.join(data_folder, "GeoTIFFs")
        lake_height_path = os.path.join(data_folder, "lake height.xlsx")
        sampling_kml_path = os.path.join(data_folder, "sampling.kml")
        possible_video = [
            os.path.join(data_folder, "timelapse.mp4"),
            os.path.join(data_folder, "Sentinel-2_L1C-202307221755611-timelapse.gif"),
            os.path.join(images_folder, "Sentinel-2_L1C-202307221755611-timelapse.gif")
        ]
        video_path = None
        for v in possible_video:
            if os.path.exists(v):
                video_path = v
                break

        st.sidebar.header(f"Dashboard Settings ({waterbody} - Dashboard)")
        x_start = st.sidebar.date_input("Start Date", date(2015, 1, 1), key="wq_start")
        x_end = st.sidebar.date_input("End Date", date(2026, 12, 31), key="wq_end")
        x_start_dt = datetime.combine(x_start, datetime.min.time())
        x_end_dt = datetime.combine(x_end, datetime.min.time())

        tif_files = [f for f in os.listdir(images_folder) if f.lower().endswith('.tif')]
        available_dates = {}
        for filename in tif_files:
            match = re.search(r'(\d{4})[_-]?(\d{2})[_-]?(\d{2})', filename)
            if match:
                year, month, day = match.groups()
                date_str = f"{year}_{month}_{day}"
                try:
                    date_obj = datetime.strptime(date_str, '%Y_%m_%d').date()
                    available_dates[str(date_obj)] = filename
                except Exception as e:
                    debug("Error extracting date from", filename, ":", e)
                    continue

        if available_dates:
            sorted_dates = sorted(available_dates.keys())
            selected_bg_date = st.selectbox("Select background date", sorted_dates, key="wq_bg")
        else:
            selected_bg_date = None
            st.warning("No GeoTIFF images with date found.")

        if selected_bg_date is not None:
            bg_filename = available_dates[selected_bg_date]
            bg_path = os.path.join(images_folder, bg_filename)
            if os.path.exists(bg_path):
                with rasterio.open(bg_path) as src:
                    if src.count >= 3:
                        first_image_data = src.read([1, 2, 3])
                        first_transform = src.transform
                    else:
                        st.error("The selected GeoTIFF does not have at least 3 bands.")
                        st.stop()
            else:
                st.error(f"GeoTIFF background not found: {bg_path}")
                st.stop()
        else:
            st.error("No valid background date selected.")
            st.stop()

        def parse_sampling_kml(kml_file) -> list:
            try:
                tree = ET.parse(kml_file)
                root = tree.getroot()
                namespace = {'kml': 'http://www.opengis.net/kml/2.2'}
                points = []
                for linestring in root.findall('.//kml:LineString', namespace):
                    coord_text = linestring.find('kml:coordinates', namespace).text.strip()
                    coords = coord_text.split()
                    for idx, coord in enumerate(coords):
                        lon_str, lat_str, *_ = coord.split(',')
                        points.append((f"Point {idx+1}", float(lon_str), float(lat_str)))
                return points
            except Exception as e:
                st.error("Error parsing KML:", e)
                return []

        def geographic_to_pixel(lon: float, lat: float, transform) -> tuple:
            inverse_transform = ~transform
            col, row = inverse_transform * (lon, lat)
            return int(col), int(row)

        def map_rgb_to_mg(r: float, g: float, b: float, mg_factor: float = 2.0) -> float:
            return (g / 255.0) * mg_factor

        def mg_to_color(mg: float) -> str:
            scale = [
                (0.00, "#0000ff"), (0.02, "#0007f2"), (0.04, "#0011de"),
                (0.06, "#0017d0"), (1.98, "#80007d"), (2.00, "#800080")
            ]
            if mg <= scale[0][0]:
                color = scale[0][1]
            elif mg >= scale[-1][0]:
                color = scale[-1][1]
            else:
                for i in range(len(scale) - 1):
                    low_mg, low_color = scale[i]
                    high_mg, high_color = scale[i+1]
                    if low_mg <= mg <= high_mg:
                        t = (mg - low_mg) / (high_mg - low_mg)
                        low_rgb = tuple(int(low_color[j:j+2], 16) for j in (1, 3, 5))
                        high_rgb = tuple(int(high_color[j:j+2], 16) for j in (1, 3, 5))
                        rgb = tuple(int(low_rgb[k] + (high_rgb[k] - low_rgb[k]) * t) for k in range(3))
                        return f"rgb({rgb[0]},{rgb[1]},{rgb[2]})"
            rgb = tuple(int(color[j:j+2], 16) for j in (1, 3, 5))
            return f"rgb({rgb[0]},{rgb[1]},{rgb[2]})"

        def analyze_sampling(sampling_points: list, first_image_data, first_transform,
                             images_folder: str, lake_height_path: str, selected_points: list = None):
            results_colors = {name: [] for name, _, _ in sampling_points}
            results_mg = {name: [] for name, _, _ in sampling_points}
            for filename in sorted(os.listdir(images_folder)):
                if filename.lower().endswith(('.tif', '.tiff')):
                    match = re.search(r'(\d{4}_\d{2}_\d{2})', filename)
                    if not match:
                        continue
                    date_str = match.group(1)
                    try:
                        date_obj = datetime.strptime(date_str, '%Y_%m_%d')
                    except ValueError:
                        continue
                    image_path = os.path.join(images_folder, filename)
                    with rasterio.open(image_path) as src:
                        transform = src.transform
                        width, height = src.width, src.height
                        if src.count < 3:
                            continue
                        for name, lon, lat in sampling_points:
                            col, row = geographic_to_pixel(lon, lat, transform)
                            if 0 <= col < width and 0 <= row < height:
                                window = rasterio.windows.Window(col, row, 1, 1)
                                r = src.read(1, window=window)[0, 0]
                                g = src.read(2, window=window)[0, 0]
                                b = src.read(3, window=window)[0, 0]
                                mg_value = map_rgb_to_mg(r, g, b)
                                results_mg[name].append((date_obj, mg_value))
                                pixel_color = (r / 255, g / 255, b / 255)
                                results_colors[name].append((date_obj, pixel_color))
            rgb_image = first_image_data.transpose((1, 2, 0)) / 255.0
            fig_geo = px.imshow(rgb_image, title='GeoTIFF Image with Sampling Points')
            for name, lon, lat in sampling_points:
                col, row = geographic_to_pixel(lon, lat, first_transform)
                fig_geo.add_trace(go.Scatter(x=[col], y=[row], mode='markers',
                                             marker=dict(color='red', size=8), name=name))
            fig_geo.update_xaxes(visible=False)
            fig_geo.update_yaxes(visible=False)
            fig_geo.update_layout(width=900, height=600, showlegend=True)
            try:
                lake_data = pd.read_excel(lake_height_path)
                lake_data['Date'] = pd.to_datetime(lake_data.iloc[:, 0])
                lake_data.sort_values('Date', inplace=True)
            except Exception as e:
                st.error(f"Error reading depth file: {e}")
                lake_data = pd.DataFrame()
            scatter_traces = []
            point_names = list(results_colors.keys())
            if selected_points is not None:
                point_names = [p for p in point_names if p in selected_points]
            for idx, name in enumerate(point_names):
                data_list = results_colors[name]
                if not data_list:
                    continue
                data_list.sort(key=lambda x: x[0])
                dates = [d for d, _ in data_list]
                colors = [f"rgb({int(c[0]*255)},{int(c[1]*255)},{int(c[2]*255)})" for _, c in data_list]
                scatter_traces.append(go.Scatter(x=dates, y=[idx] * len(dates),
                                                 mode='markers',
                                                 marker=dict(color=colors, size=10),
                                                 name=name))
            fig_colors = make_subplots(specs=[[{"secondary_y": True}]])
            for trace in scatter_traces:
                fig_colors.add_trace(trace, secondary_y=False)
            if not lake_data.empty:
                trace_height = go.Scatter(
                    x=lake_data['Date'],
                    y=lake_data[lake_data.columns[1]],
                    name='Depth', mode='lines', line=dict(color='blue', width=2)
                )
                fig_colors.add_trace(trace_height, secondary_y=True)
            fig_colors.update_layout(title='Pixel Colors and Depth Over Time',
                                     xaxis_title='Date',
                                     yaxis_title='Sampling Points',
                                     showlegend=True)
            fig_colors.update_yaxes(title_text="Depth", secondary_y=True)
            all_dates_dict = {}
            for data_list in results_mg.values():
                for date_obj, mg_val in data_list:
                    all_dates_dict.setdefault(date_obj, []).append(mg_val)
            sorted_dates = sorted(all_dates_dict.keys())
            avg_mg = [np.mean(all_dates_dict[d]) for d in sorted_dates]
            fig_mg = go.Figure()
            fig_mg.add_trace(go.Scatter(
                x=sorted_dates,
                y=avg_mg,
                mode='markers',
                marker=dict(color=avg_mg, colorscale='Viridis', reversescale=True,
                            colorbar=dict(title='mg/m³'), size=10),
                name='Mean mg/m³'
            ))
            fig_mg.update_layout(title='Mean mg/m³ Over Time',
                                 xaxis_title='Date', yaxis_title='mg/m³',
                                 showlegend=False)
            fig_dual = make_subplots(specs=[[{"secondary_y": True}]])
            if not lake_data.empty:
                fig_dual.add_trace(go.Scatter(
                    x=lake_data['Date'],
                    y=lake_data[lake_data.columns[1]],
                    name='Depth', mode='lines'
                ), secondary_y=False)
            fig_dual.add_trace(go.Scatter(
                x=sorted_dates,
                y=avg_mg,
                name='Mean mg/m³',
                mode='markers',
                marker=dict(color=avg_mg, colorscale='Viridis', reversescale=True,
                            colorbar=dict(title='mg/m³'), size=10)
            ), secondary_y=True)
            fig_dual.update_layout(title='Depth and Mean mg/m³ Over Time',
                                   xaxis_title='Date', showlegend=True)
            fig_dual.update_yaxes(title_text="Depth", secondary_y=False)
            fig_dual.update_yaxes(title_text="mg/m³", secondary_y=True)
            return fig_geo, fig_dual, fig_colors, fig_mg, results_colors, results_mg, lake_data

        # Two tabs for sampling
        if "default_results" not in st.session_state:
            st.session_state.default_results = None
        if "upload_results" not in st.session_state:
            st.session_state.upload_results = None

        tab_names = ["Sampling 1 (Default)", "Sampling 2 (Upload)"]
        sampling_tabs = st.tabs(tab_names)

        # Tab 1: Default Sampling
        with sampling_tabs[0]:
            st.header("Analysis for Sampling 1 (Default)")
            default_sampling_points = []
            if os.path.exists(sampling_kml_path):
                default_sampling_points = parse_sampling_kml(sampling_kml_path)
            else:
                st.warning("Sampling file (sampling.kml) not found.")
            point_names = [name for name, _, _ in default_sampling_points]
            selected_points = st.multiselect("Select points for mg/m³ analysis",
                                             options=point_names,
                                             default=point_names,
                                             key="default_points")
            if st.button("Run Analysis (Default)", key="default_run"):
                with st.spinner("Running analysis..."):
                    st.session_state.default_results = analyze_sampling(
                        default_sampling_points,
                        first_image_data,
                        first_transform,
                        images_folder,
                        lake_height_path,
                        selected_points
                    )
            if st.session_state.default_results is not None:
                results = st.session_state.default_results
                if isinstance(results, tuple) and len(results) == 7:
                    fig_geo, fig_dual, fig_colors, fig_mg, results_colors, results_mg, lake_data = results
                else:
                    st.error("Result formatting error. Please rerun the analysis.")
                    st.stop()
                nested_tabs = st.tabs(["GeoTIFF", "Image Selection", "Video/GIF", "Pixel Colors", "Mean mg", "Dual Charts", "Detailed mg Analysis"])
                with nested_tabs[0]:
                    st.plotly_chart(fig_geo, use_container_width=True, key="default_fig_geo")
                with nested_tabs[1]:
                    st.header("Image Selection")
                    tif_files = [f for f in os.listdir(images_folder) if f.lower().endswith('.tif')]
                    available_dates = {}
                    for filename in tif_files:
                        match = re.search(r'(\d{4}_\d{2}_\d{2})', filename)
                        if match:
                            date_str = match.group(1)
                            try:
                                date_obj = datetime.strptime(date_str, '%Y_%m_%d').date()
                                available_dates[str(date_obj)] = filename
                            except Exception as e:
                                debug("Error extracting date from", filename, ":", e)
                                continue
                    if available_dates:
                        sorted_dates = sorted(available_dates.keys())
                        if 'current_image_index' not in st.session_state:
                            st.session_state.current_image_index = 0
                        col_prev, col_select, col_next = st.columns([1, 3, 1])
                        with col_prev:
                            if st.button("<< Previous"):
                                st.session_state.current_image_index = max(0, st.session_state.current_image_index - 1)
                        with col_next:
                            if st.button("Next >>"):
                                st.session_state.current_image_index = min(len(sorted_dates) - 1, st.session_state.current_image_index + 1)
                        with col_select:
                            selected_date = st.selectbox("Select date", sorted_dates, index=st.session_state.current_image_index)
                            st.session_state.current_image_index = sorted_dates.index(selected_date)
                        current_date = sorted_dates[st.session_state.current_image_index]
                        st.write(f"Selected Date: {current_date}")
                        image_filename = available_dates[current_date]
                        image_path = os.path.join(images_folder, image_filename)
                        if os.path.exists(image_path):
                            st.image(image_path, caption=f"Image for {current_date}", use_container_width=True)
                        else:
                            st.error("Image not found.")
                    else:
                        st.info("No images found with a date in the folder.")
                with nested_tabs[2]:
                    if video_path is not None:
                        if video_path.endswith(".mp4"):
                            st.video(video_path, key="default_video")
                        else:
                            st.image(video_path)
                    else:
                        st.info("No timelapse video found.")
                with nested_tabs[3]:
                    st.plotly_chart(fig_colors, use_container_width=True, key="default_fig_colors")
                with nested_tabs[4]:
                    st.plotly_chart(fig_mg, use_container_width=True, key="default_fig_mg")
                with nested_tabs[5]:
                    st.plotly_chart(fig_dual, use_container_width=True, key="default_fig_dual")
                with nested_tabs[6]:
                    selected_detail_point = st.selectbox("Select point for detailed mg analysis",
                                                         options=list(results_mg.keys()),
                                                         key="default_detail")
                    if selected_detail_point:
                        mg_data = results_mg[selected_detail_point]
                        if mg_data:
                            mg_data_sorted = sorted(mg_data, key=lambda x: x[0])
                            dates_mg = [d for d, _ in mg_data_sorted]
                            mg_values = [val for _, val in mg_data_sorted]
                            detail_colors = [mg_to_color(val) for val in mg_values]
                            fig_detail = go.Figure()
                            fig_detail.add_trace(go.Scatter(
                                x=dates_mg, y=mg_values, mode='lines+markers',
                                marker=dict(color=detail_colors, size=10),
                                line=dict(color="gray"),
                                name=selected_detail_point
                            ))
                            fig_detail.update_layout(title=f"Detailed mg analysis for {selected_detail_point}",
                                                     xaxis_title="Date", yaxis_title="mg/m³")
                            st.plotly_chart(fig_detail, use_container_width=True, key="default_fig_detail")
                        else:
                            st.info("No mg data available for this point.")
        # Tab 2: Upload Sampling
        with sampling_tabs[1]:
            st.header("Analysis for Upload Sampling")
            uploaded_file = st.file_uploader("Upload a KML file for new sampling points", type="kml", key="upload_kml")
            if uploaded_file is not None:
                try:
                    new_sampling_points = parse_sampling_kml(uploaded_file)
                except Exception as e:
                    st.error(f"Error processing uploaded file: {e}")
                    new_sampling_points = []
                point_names = [name for name, _, _ in new_sampling_points]
                selected_points = st.multiselect("Select points for mg/m³ analysis",
                                                 options=point_names,
                                                 default=point_names,
                                                 key="upload_points")
                if st.button("Run Analysis (Upload)", key="upload_run"):
                    with st.spinner("Running analysis..."):
                        st.session_state.upload_results = analyze_sampling(
                            new_sampling_points,
                            first_image_data,
                            first_transform,
                            images_folder,
                            lake_height_path,
                            selected_points
                        )
                if st.session_state.upload_results is not None:
                    results = st.session_state.upload_results
                    if isinstance(results, tuple) and len(results) == 7:
                        fig_geo, fig_dual, fig_colors, fig_mg, results_colors, results_mg, lake_data = results
                    else:
                        st.error("Result formatting error (Upload). Please rerun the analysis.")
                        st.stop()
                    nested_tabs = st.tabs(["GeoTIFF", "Image Selection", "Video/GIF", "Pixel Colors", "Mean mg", "Dual Charts", "Detailed mg Analysis"])
                    with nested_tabs[0]:
                        st.plotly_chart(fig_geo, use_container_width=True, key="upload_fig_geo")
                    with nested_tabs[1]:
                        st.header("Image Selection")
                        tif_files = [f for f in os.listdir(images_folder) if f.lower().endswith('.tif')]
                        available_dates = {}
                        for filename in tif_files:
                            match = re.search(r'(\d{4}_\d{2}_\d{2})', filename)
                            if match:
                                date_str = match.group(1)
                                try:
                                    date_obj = datetime.strptime(date_str, '%Y_%m_%d').date()
                                    available_dates[str(date_obj)] = filename
                                except Exception as e:
                                    debug("Error extracting date from", filename, ":", e)
                                    continue
                        if available_dates:
                            sorted_dates = sorted(available_dates.keys())
                            if 'current_upload_image_index' not in st.session_state:
                                st.session_state.current_upload_image_index = 0
                            col_prev, col_select, col_next = st.columns([1, 3, 1])
                            with col_prev:
                                if st.button("<< Previous", key="upload_prev"):
                                    st.session_state.current_upload_image_index = max(0, st.session_state.current_upload_image_index - 1)
                            with col_next:
                                if st.button("Next >>", key="upload_next"):
                                    st.session_state.current_upload_image_index = min(len(sorted_dates) - 1, st.session_state.current_upload_image_index + 1)
                            with col_select:
                                selected_date = st.selectbox("Select date", sorted_dates, index=st.session_state.current_upload_image_index)
                                st.session_state.current_upload_image_index = sorted_dates.index(selected_date)
                            current_date = sorted_dates[st.session_state.current_upload_image_index]
                            st.write(f"Selected Date: {current_date}")
                            image_filename = available_dates[current_date]
                            image_path = os.path.join(images_folder, image_filename)
                            if os.path.exists(image_path):
                                st.image(image_path, caption=f"Image for {current_date}", use_container_width=True)
                            else:
                                st.error("Image not found.")
                        else:
                            st.info("No images found with a date in the folder.")
                    with nested_tabs[2]:
                        if video_path is not None:
                            if video_path.endswith(".mp4"):
                                st.video(video_path, key="upload_video")
                            else:
                                st.image(video_path)
                        else:
                            st.info("No Video/GIF file found.")
                    with nested_tabs[3]:
                        st.plotly_chart(fig_colors, use_container_width=True, key="upload_fig_colors")
                    with nested_tabs[4]:
                        st.plotly_chart(fig_mg, use_container_width=True, key="upload_fig_mg")
                    with nested_tabs[5]:
                        st.plotly_chart(fig_dual, use_container_width=True, key="upload_fig_dual")
                    with nested_tabs[6]:
                        selected_detail_point = st.selectbox("Select point for detailed mg analysis",
                                                             options=list(results_mg.keys()),
                                                             key="upload_detail")
                        if selected_detail_point:
                            mg_data = results_mg[selected_detail_point]
                            if mg_data:
                                mg_data_sorted = sorted(mg_data, key=lambda x: x[0])
                                dates_mg = [d for d, _ in mg_data_sorted]
                                mg_values = [val for _, val in mg_data_sorted]
                                detail_colors = [mg_to_color(val) for val in mg_values]
                                fig_detail = go.Figure()
                                fig_detail.add_trace(go.Scatter(
                                    x=dates_mg, y=mg_values, mode='lines+markers',
                                    marker=dict(color=detail_colors, size=10),
                                    line=dict(color="gray"),
                                    name=selected_detail_point
                                ))
                                fig_detail.update_layout(title=f"Detailed mg analysis for {selected_detail_point}",
                                                         xaxis_title="Date", yaxis_title="mg/m³")
                                st.plotly_chart(fig_detail, use_container_width=True, key="upload_fig_detail")
                            else:
                                st.info("No mg data available for this point.", key="upload_no_mg")
            else:
                st.info("Please upload a KML file for new sampling points.")

        st.info("End of Subterranean Quality Dashboard.")
        st.markdown('</div>', unsafe_allow_html=True)

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
