# Biomechanics Graph App

This repository contains a Tkinter-based GUI application for plotting and analysing biomechanics data stored in Excel or CSV files.

## Usage

1. Install requirements:
   ```bash
   pip install pandas matplotlib scipy
   ```
2. Run the application:
   ```bash
   python bio_graph_app.py
   ```

The original Jupyter notebook is provided as `アオキ編集中-完成.ipynb`. It was converted to the standalone script `bio_graph_app.py` for easier execution.

## New features

- Multiple files can be loaded at once. Each file appears as a selectable sheet.
- Graph presets now support descriptions and tags and can be exported/imported as JSON.
- Double clicking the graph allows adding annotations which are drawn with arrows.
- Some internal functions were refactored into modules under `biograph/`.
