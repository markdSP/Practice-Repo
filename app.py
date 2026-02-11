"""
Python to Browser App - Training Tool
A simple Streamlit application that reads Excel data, processes it, and displays a bar chart.
"""

import streamlit as st
import pandas as pd
import plotly.express as px
from pathlib import Path

# Page configuration
st.set_page_config(
    page_title="Excel to Chart App",
    page_icon="📊",
    layout="wide"
)

# App title
st.title("📊 Excel Data Visualization Tool")
st.markdown("---")

# Sidebar instructions
with st.sidebar:
    st.header("📖 Instructions")
    st.markdown("""
    1. The app automatically loads **data.xlsx**
    2. Modify the Excel file with your data
    3. Click **Reload Data** to refresh
    4. View your updated chart!
    
    ### Expected Excel Format:
    - Column 1: Categories (text)
    - Column 2: Values (numbers)
    """)

# File path
excel_file = Path("sales_report.xlsx")

# Function to load data
@st.cache_data
def load_data(file_path):
    """Load data from Excel file"""
    try:
        df = pd.read_excel(file_path)
        return df, None
    except FileNotFoundError:
        return None, "❌ Error: data.xlsx not found! Please ensure the file exists in the project folder."
    except Exception as e:
        return None, f"❌ Error loading file: {str(e)}"

# Reload button
col1, col2, col3 = st.columns([1, 1, 3])
with col1:
    if st.button("🔄 Reload Data", use_container_width=True):
        st.cache_data.clear()
        st.rerun()

# Load the data
df, error = load_data(excel_file)

if error:
    st.error(error)
    st.stop()

# Display data info
st.subheader("📋 Data Preview")
col1, col2 = st.columns(2)
with col1:
    st.metric("Total Rows", len(df))
with col2:
    st.metric("Total Columns", len(df.columns))

# Show the dataframe
st.dataframe(df, use_container_width=True, height=200)

# Data processing section
st.markdown("---")
st.subheader("🔧 Data Processing")

# Get column names dynamically
columns = df.columns.tolist()

if len(columns) >= 2:
    # Let user select which columns to use
    col1, col2 = st.columns(2)
    with col1:
        category_col = st.selectbox("Select Category Column", columns, index=0)
    with col2:
        value_col = st.selectbox("Select Value Column", columns, index=1)
    
    # Process the data
    try:
        # Create a clean dataframe for plotting
        plot_df = df[[category_col, value_col]].copy()
        
        # Remove any rows with missing values
        plot_df = plot_df.dropna()
        
        # Convert value column to numeric
        plot_df[value_col] = pd.to_numeric(plot_df[value_col], errors='coerce')
        plot_df = plot_df.dropna()
        
        # Display processed data stats
        st.success(f"✅ Processed {len(plot_df)} valid rows")
        
        # Visualization section
        st.markdown("---")
        st.subheader("📊 Bar Chart Visualization")
        
        # Chart customization options
        col1, col2 = st.columns(2)
        with col1:
            chart_title = st.text_input("Chart Title", f"{value_col} by {category_col}")
        with col2:
            color_theme = st.selectbox("Color Theme", 
                                      ["Plotly", "Viridis", "Plasma", "Blues", "Reds", "Greens"])
        
        # Create the bar chart
        fig = px.bar(
            plot_df,
            x=category_col,
            y=value_col,
            title=chart_title,
            labels={category_col: category_col, value_col: value_col},
            color=value_col,
            color_discrete_sequence=px.colors.qualitative.Plotly if color_theme == "Plotly" else None,
            template="plotly_white"
        )
        
        # Apply color theme if not default
        if color_theme != "Plotly":
            color_map = {
                "Viridis": px.colors.sequential.Viridis,
                "Plasma": px.colors.sequential.Plasma,
                "Blues": px.colors.sequential.Blues,
                "Reds": px.colors.sequential.Reds,
                "Greens": px.colors.sequential.Greens
            }
            colors = color_map.get(color_theme, px.colors.qualitative.Plotly)
            fig.update_traces(marker_color=colors[0:len(plot_df)])
        
        # Update layout for better appearance
        fig.update_layout(
            height=500,
            showlegend=False,
            hovermode='x unified'
        )
        
        # Display the chart
        st.plotly_chart(fig, use_container_width=True)
        
        # Display summary statistics
        st.markdown("---")
        st.subheader("📈 Summary Statistics")
        
        col1, col2, col3, col4 = st.columns(4)
        with col1:
            st.metric("Mean", f"{plot_df[value_col].mean():.2f}")
        with col2:
            st.metric("Median", f"{plot_df[value_col].median():.2f}")
        with col3:
            st.metric("Max", f"{plot_df[value_col].max():.2f}")
        with col4:
            st.metric("Min", f"{plot_df[value_col].min():.2f}")
            
    except Exception as e:
        st.error(f"❌ Error processing data: {str(e)}")
        st.info("💡 Tip: Make sure your value column contains numeric data!")
else:
    st.warning("⚠️ The Excel file needs at least 2 columns to create a chart.")

# Footer
st.markdown("---")
st.markdown("""
<div style='text-align: center; color: #666;'>
    <p>Built with Streamlit 🎈 | Modify data.xlsx and reload to see changes</p>
</div>
""", unsafe_allow_html=True)
