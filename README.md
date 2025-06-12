# Biomechanics Graph App

This repository contains a Tkinter-based GUI application for plotting and analyzing biomechanics data stored in Excel or CSV files.

## Usage

1. Install requirements:
   ```bash
   pip install -r requirements.txt
   ```
2. Run the menu script to choose whether to integrate files or start the GUI:
   ```bash
   python main_menu.py
   ```

   When you select the integration option you will be prompted for multiple CSV
   or Excel files. The selected files are concatenated and saved as a new CSV
   file.

   You can still launch the GUI directly if desired:
   ```bash
   python bio_graph_app.py
   ```

To exit the application, use **File → 終了** from the menu bar or close the window. The app now confirms before closing.

Use the "CSVに保存..." button in the data table window to export processed data to a CSV file. If no sliced data is available, you'll be notified instead of saving an empty file.

Presets now store the selected X/Y columns in addition to display settings so that you can easily reapply axis selections.

The original Jupyter notebook is provided as `アオキ編集中-完成.ipynb`. It was converted to the standalone script `bio_graph_app.py` for easier execution.

## License

This project is licensed under the MIT License. See [LICENSE](LICENSE) for details.
