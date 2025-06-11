#!/usr/bin/env python
# coding: utf-8



import tkinter as tk
from tkinter import ttk, filedialog, messagebox, simpledialog
import pandas as pd
import matplotlib
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import matplotlib.lines as mlines
import re
import numpy as np
import platform
import sqlite3
import json
import os
from scipy.integrate import cumulative_trapezoid

# --- Matplotlibの日本語フォント設定 ---
try:
    if platform.system() == 'Windows':
        font_candidates = ['Meiryo', 'MS Gothic', 'Yu Gothic', 'TakaoPGothic', 'IPAexGothic', 'sans-serif']
    elif platform.system() == 'Darwin':
        font_candidates = ['Hiragino Sans', 'IPAexGothic', 'TakaoPGothic', 'AppleGothic', 'sans-serif']
    else:
        font_candidates = ['IPAexGothic', 'TakaoPGothic', 'VL Gothic', 'Noto Sans CJK JP', 'DejaVu Sans', 'sans-serif']
    found_font = False
    for font_name in font_candidates:
        try:
            matplotlib.font_manager.fontManager.findfont(font_name, fallback_to_default=False)
            matplotlib.rcParams['font.family'] = font_name
            found_font = True; break
        except: continue
    if not found_font: matplotlib.rcParams['font.family'] = 'sans-serif'
    matplotlib.rcParams['axes.unicode_minus'] = False
except Exception as e:
    print(f"Error setting Matplotlib font properties globally: {e}")


class AboutAppWindow(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("アプリケーション情報")
        self.geometry("600x450") 
        self.transient(master) 
        self.grab_set() 

        notebook = ttk.Notebook(self)
        notebook.pack(expand=True, fill="both", padx=10, pady=10)

        info_tab = ttk.Frame(notebook)
        notebook.add(info_tab, text="バージョン/作成者")
        self.create_info_tab(info_tab)

        features_tab = ttk.Frame(notebook)
        notebook.add(features_tab, text="機能一覧")
        self.create_features_tab(features_tab)

        history_tab = ttk.Frame(notebook)
        notebook.add(history_tab, text="修正履歴")
        self.create_history_tab(history_tab)

        close_button = ttk.Button(self, text="閉じる", command=self.destroy)
        close_button.pack(pady=10)

    def create_info_tab(self, parent_frame):
        app_name_label = ttk.Label(parent_frame, text="バイオメカニクス グラフ表示アプリ", font=("TkDefaultFont", 14, "bold"))
        app_name_label.pack(pady=10)
        
        version_label = ttk.Label(parent_frame, text="バージョン: 1.0.0 (2025-05-23)") # バージョンと日付を更新
        version_label.pack(pady=5)
        
        author_text = "作成者: 青木ビクター達哉\n(筑波大学大学院 野球コーチング論研究室)"
        author_label = ttk.Label(parent_frame, text=author_text, justify=tk.CENTER)
        author_label.pack(pady=10)
        
    def create_features_tab(self, parent_frame):
        features_text_frame = ttk.Frame(parent_frame)
        features_text_frame.pack(expand=True, fill="both", padx=5, pady=5)

        features_text = tk.Text(features_text_frame, wrap="word", height=15, width=70)
        features_text.pack(side=tk.LEFT, expand=True, fill="both")
        
        features_content = """
        ■ 主な機能 (Ver 1.0.0)
        - Excelファイル(.xlsx, .xls)からのデータ読み込み
        - シート選択、X軸・Y軸（複数可）データ列選択
        - グラフ描画（線グラフ）
        - グラフタイトル、軸ラベル、凡例名編集
        - グラフ外観カスタマイズ
            - プロットエリア背景色、図全体の背景色
            - グリッドの表示/非表示、色、スタイル、太さ
            - 基本フォントサイズ調整（タイトル、軸ラベル、凡例、目盛り）
        - データ範囲指定（開始行、終了行）
        - イベントマーカー（縦線）の追加・編集・削除（最大5本）
        - グラフのアスペクト比、凡例表示位置の選択
        - 描画されたグラフの画像保存（PNG, PDF）
        - プリセット機能（設定の保存・読み込み・削除）
        - データテーブル表示
            - スライスデータ（現在グラフ表示中のデータ）のテーブル表示とクリップボードコピー
            - イベントマーカー位置での各Y軸の値表示
            - 選択Y軸の基本統計量表示
        - 自動イベント検出 (最大値・最小値プロット)
        - データ処理 (選択Y軸データの微分・積分計算とグラフ追加)
        - グラフのインタラクティブ操作
            - マウスホイールによるズームイン・ズームアウト
            - データカーソル（マウスオーバーで座標と値を表示）
            - 凡例クリックによるグラフ線の表示/非表示切り替え
            - Matplotlibナビゲーションツールバーによる操作
        - コントロールパネルのスクロール機能
        - アプリケーション情報ウィンドウ（バージョン、機能、履歴）
        """
        features_text.insert(tk.END, features_content.strip())
        features_text.config(state="disabled") 

        vsb = ttk.Scrollbar(features_text_frame, orient="vertical", command=features_text.yview)
        features_text.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")

    def create_history_tab(self, parent_frame):
        history_text_frame = ttk.Frame(parent_frame)
        history_text_frame.pack(expand=True, fill="both", padx=5, pady=5)

        history_text = tk.Text(history_text_frame, wrap="word", height=15, width=70)
        history_text.pack(side=tk.LEFT, expand=True, fill="both")
        
        history_content = """
        ■ Ver 1.0.0 (2025-05-23)
        - 初期リリース。
        """
        history_text.insert(tk.END, history_content.strip())
        history_text.config(state="disabled")

        vsb = ttk.Scrollbar(history_text_frame, orient="vertical", command=history_text.yview)
        history_text.configure(yscrollcommand=vsb.set)
        vsb.pack(side="right", fill="y")


class DataOutputWindow(tk.Toplevel):
    def __init__(self, master, app_instance):
        super().__init__(master)
        self.title("データテーブル出力")
        self.geometry("800x600")
        self.app = app_instance 
        if self.app.sliced_df is None or self.app.sliced_df.empty:
            messagebox.showwarning("データなし", "表示するデータがありません。グラフを描画してください。", parent=self)
            self.destroy(); return
        self.notebook = ttk.Notebook(self)
        self.notebook.pack(expand=True, fill="both", padx=10, pady=10)

        self.sliced_data_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.sliced_data_tab, text="スライスデータ")
        self.create_sliced_data_table(self.sliced_data_tab)

        self.marker_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.marker_tab, text="イベントマーカー値")
        self.create_marker_values_table(self.marker_tab)

        self.stats_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.stats_tab, text="Y軸データ統計量")
        self.create_statistics_table_ui(self.stats_tab)

        self.protocol("WM_DELETE_WINDOW", self.on_close)

    def on_close(self):
        self.app.data_output_window = None; self.destroy()

    def create_sliced_data_table(self, parent_frame):
        table_frame = ttk.Frame(parent_frame)
        table_frame.pack(expand=True, fill="both", padx=5, pady=5)

        if self.app.sliced_df is None or self.app.sliced_df.empty:
            ttk.Label(table_frame, text="表示するデータがありません").pack(padx=10, pady=10)
            return

        columns = self.app.sliced_df.columns.tolist()
        self.sliced_data_tree = ttk.Treeview(table_frame, columns=columns, show="headings")
        for col in columns:
            self.sliced_data_tree.heading(col, text=col)
            self.sliced_data_tree.column(col, width=100, anchor="center", minwidth=50)

        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.sliced_data_tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.sliced_data_tree.xview)
        self.sliced_data_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)

        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self.sliced_data_tree.pack(expand=True, fill="both")

        for index, row in self.app.sliced_df.iterrows():
            formatted_row = [f"{val:.3f}" if isinstance(val, (int, float)) else str(val) for val in row]
            self.sliced_data_tree.insert("", "end", values=formatted_row)

        copy_button = ttk.Button(parent_frame, text="テーブル内容をコピー", command=lambda: self.copy_treeview_to_clipboard(self.sliced_data_tree))
        copy_button.pack(pady=5)

        export_button = ttk.Button(parent_frame, text="CSVに保存...", command=self.export_sliced_data_to_csv)
        export_button.pack(pady=5)

    def copy_treeview_to_clipboard(self, treeview):
        try:
            header = '\t'.join(treeview['columns']) + '\n'
            items_data = []
            for child_id in treeview.get_children():
                values = treeview.item(child_id, 'values')
                items_data.append('\t'.join(map(str, values)))
            text_to_copy = header + '\n'.join(items_data)
            
            self.clipboard_clear()
            self.clipboard_append(text_to_copy)
            self.update() 
            messagebox.showinfo("コピー完了", "テーブルの内容をクリップボードにコピーしました。", parent=self)
        except Exception as e:
            messagebox.showerror("コピー失敗", f"クリップボードへのコピー中にエラーが発生しました:\n{e}", parent=self)

    def export_sliced_data_to_csv(self):
        """Export the displayed sliced data table to a CSV file."""
        if self.app.sliced_df is None or self.app.sliced_df.empty:
            messagebox.showwarning("データなし", "書き出すスライスデータがありません。", parent=self)
            return

        file_path = filedialog.asksaveasfilename(
            title="スライスデータをCSV保存",
            defaultextension=".csv",
            filetypes=(("CSVファイル", "*.csv"), ("すべてのファイル", "*.*"))
        )
        if not file_path:
            return
        try:
            self.app.sliced_df.to_csv(file_path, index=False)
            messagebox.showinfo("成功", f"スライスデータを {file_path} に保存しました。", parent=self)
        except Exception as e:
            messagebox.showerror("保存失敗", f"CSV保存中にエラーが発生しました:\n{e}", parent=self)


    def create_marker_values_table(self, parent_frame):
        table_frame = ttk.Frame(parent_frame)
        table_frame.pack(expand=True, fill="both", padx=5, pady=5)
        columns = ["マーカー名", "X座標"]
        selected_y_cols_original = [self.app.y_axis_listbox.get(i) for i in self.app.y_axis_listbox.curselection()]
        y_col_display_names = [self.app.legend_label_vars.get(original_name, tk.StringVar(value=original_name)).get() for original_name in selected_y_cols_original]
        columns.extend(y_col_display_names)

        self.marker_tree = ttk.Treeview(table_frame, columns=columns, show="headings")
        for col in columns:
            self.marker_tree.heading(col, text=col)
            self.marker_tree.column(col, width=100, anchor="center")

        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.marker_tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.marker_tree.xview)
        self.marker_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self.marker_tree.pack(expand=True, fill="both")
        self.populate_marker_values_table()

    def populate_marker_values_table(self):
        for item in self.marker_tree.get_children():
            self.marker_tree.delete(item)

        if self.app.sliced_df is None or self.app.sliced_df.empty or not self.app.vline_configs:
            self.marker_tree.insert("", "end", values=("イベントマーカーが設定されていません",) + ("",) * (len(self.marker_tree["columns"]) - 1))
            return

        selected_x_col = self.app.x_axis_var.get()
        if not selected_x_col:
            self.marker_tree.insert("", "end", values=("X軸が選択されていません",) + ("",) * (len(self.marker_tree["columns"]) - 1))
            return

        selected_y_cols_original = [self.app.y_axis_listbox.get(i) for i in self.app.y_axis_listbox.curselection()]
        if not selected_y_cols_original:
            self.marker_tree.insert("", "end", values=("Y軸が選択されていません",) + ("",) * (len(self.marker_tree["columns"]) - 1))
            return

        for vline_config in self.app.vline_configs:
            marker_name = vline_config['name_var'].get() or "(名称なし)"
            x_coord_str = vline_config['x_var'].get()
            row_values = [marker_name]

            if not x_coord_str:
                row_values.extend(["N/A"] * (1 + len(selected_y_cols_original)))
                self.marker_tree.insert("", "end", values=tuple(row_values))
                continue
            try:
                x_coord = float(x_coord_str)
                row_values.append(f"{x_coord:.3f}")

                if selected_x_col not in self.app.sliced_df.columns or not pd.api.types.is_numeric_dtype(self.app.sliced_df[selected_x_col]):
                    row_values.extend(["X軸非数値/存在せず"] * len(selected_y_cols_original))
                    self.marker_tree.insert("", "end", values=tuple(row_values))
                    continue

                valid_x_series = self.app.sliced_df[selected_x_col].dropna()
                if valid_x_series.empty:
                    row_values.extend(["X軸NaNのみ"] * len(selected_y_cols_original))
                    self.marker_tree.insert("", "end", values=tuple(row_values))
                    continue
                
                closest_idx = (valid_x_series - x_coord).abs().idxmin()

                for y_col_original in selected_y_cols_original:
                    if y_col_original in self.app.sliced_df.columns and closest_idx in self.app.sliced_df.index:
                        y_val = self.app.sliced_df.loc[closest_idx, y_col_original]
                        row_values.append(f"{y_val:.3f}" if pd.notnull(y_val) and isinstance(y_val, (int, float)) else "N/A")
                    else:
                        row_values.append("データなし")
            except ValueError:
                row_values.append("X座標不正")
                row_values.extend(["N/A"] * len(selected_y_cols_original))
            except KeyError as e:
                print(f"KeyError in populate_marker_values_table: {e}")
                row_values.append(f"{x_coord:.3f}" if 'x_coord' in locals() else "X座標不明")
                row_values.extend(["列エラー"] * len(selected_y_cols_original))
            except Exception as e:
                print(f"Error populating marker table row: {e}")
                row_values.append(f"{x_coord:.3f}" if 'x_coord' in locals() else "X座標不明")
                row_values.extend(["エラー"] * len(selected_y_cols_original))
            self.marker_tree.insert("", "end", values=tuple(row_values))


    def create_statistics_table_ui(self, parent_frame):
        checkbox_frame = ttk.Frame(parent_frame)
        checkbox_frame.pack(pady=5, padx=5, fill="x")
        self.stat_vars = {} 
        self.stat_items = {
            "最大値": "max", "最小値": "min", "平均値": "mean",
            "標準偏差": "std", "中央値": "median",
            "最大値時のX座標": "idxmax_x", "最小値時のX座標": "idxmin_x"
        }
        ttk.Label(checkbox_frame, text="表示する統計項目:").pack(side=tk.LEFT, padx=(0, 10))
        for display_name in self.stat_items.keys():
            var = tk.BooleanVar(value=True)
            self.stat_vars[display_name] = var 
            cb = ttk.Checkbutton(checkbox_frame, text=display_name, variable=var, command=self.populate_statistics_table)
            cb.pack(side=tk.LEFT, padx=2)

        table_frame = ttk.Frame(parent_frame)
        table_frame.pack(expand=True, fill="both", padx=5, pady=5)
        self.stats_tree = ttk.Treeview(table_frame, show="headings")
        vsb = ttk.Scrollbar(table_frame, orient="vertical", command=self.stats_tree.yview)
        hsb = ttk.Scrollbar(table_frame, orient="horizontal", command=self.stats_tree.xview)
        self.stats_tree.configure(yscrollcommand=vsb.set, xscrollcommand=hsb.set)
        vsb.pack(side="right", fill="y")
        hsb.pack(side="bottom", fill="x")
        self.stats_tree.pack(expand=True, fill="both")
        self.populate_statistics_table()


    def populate_statistics_table(self):
        for item in self.stats_tree.get_children():
            self.stats_tree.delete(item)

        if self.stats_tree["columns"]:
             self.stats_tree["columns"] = ()


        if self.app.sliced_df is None or self.app.sliced_df.empty:
            self.stats_tree["columns"] = ("メッセージ",)
            self.stats_tree.heading("メッセージ", text="メッセージ")
            self.stats_tree.insert("", "end", values=("表示するデータがありません",))
            return

        selected_y_cols_original = [self.app.y_axis_listbox.get(i) for i in self.app.y_axis_listbox.curselection()]
        if not selected_y_cols_original:
            self.stats_tree["columns"] = ("メッセージ",)
            self.stats_tree.heading("メッセージ", text="メッセージ")
            self.stats_tree.insert("", "end", values=("Y軸が選択されていません",))
            return

        selected_x_col = self.app.x_axis_var.get()
        active_stats_display_names = [name for name, var in self.stat_vars.items() if var.get()]

        if not active_stats_display_names:
            self.stats_tree["columns"] = ("メッセージ",)
            self.stats_tree.heading("メッセージ", text="メッセージ")
            self.stats_tree.insert("", "end", values=("表示する統計項目が選択されていません",))
            return

        columns = ["Y軸データ系列"] + active_stats_display_names
        self.stats_tree["columns"] = columns

        for col_name in columns:
            self.stats_tree.heading(col_name, text=col_name)
            self.stats_tree.column(col_name, width=120, anchor="center", minwidth=60)


        for y_col_original in selected_y_cols_original:
            y_col_display_name = self.app.legend_label_vars.get(y_col_original, tk.StringVar(value=y_col_original)).get()
            row_values = [y_col_display_name]

            if y_col_original not in self.app.sliced_df.columns:
                row_values.extend(["列なし"] * len(active_stats_display_names))
                self.stats_tree.insert("", "end", values=tuple(row_values))
                continue
            
            y_series = self.app.sliced_df[y_col_original]

            if not pd.api.types.is_numeric_dtype(y_series):
                row_values.extend(["非数値データ"] * len(active_stats_display_names))
                self.stats_tree.insert("", "end", values=tuple(row_values))
                continue

            y_series_numeric = y_series.dropna()
            if y_series_numeric.empty:
                row_values.extend(["NaNのみ"] * len(active_stats_display_names))
                self.stats_tree.insert("", "end", values=tuple(row_values))
                continue

            for stat_display_name in active_stats_display_names:
                stat_key = self.stat_items.get(stat_display_name)
                val = "N/A"
                try:
                    if stat_key == "max": val = f"{y_series_numeric.max():.3f}"
                    elif stat_key == "min": val = f"{y_series_numeric.min():.3f}"
                    elif stat_key == "mean": val = f"{y_series_numeric.mean():.3f}"
                    elif stat_key == "std": val = f"{y_series_numeric.std():.3f}"
                    elif stat_key == "median": val = f"{y_series_numeric.median():.3f}"
                    elif stat_key == "idxmax_x" and selected_x_col and selected_x_col in self.app.sliced_df.columns:
                        idx_max = y_series_numeric.idxmax()
                        if pd.notnull(idx_max) and idx_max in self.app.sliced_df.index:
                            x_at_max = self.app.sliced_df.loc[idx_max, selected_x_col]
                            val = f"{x_at_max:.3f}" if pd.notnull(x_at_max) and isinstance(x_at_max, (int, float)) else "N/A"
                    elif stat_key == "idxmin_x" and selected_x_col and selected_x_col in self.app.sliced_df.columns:
                        idx_min = y_series_numeric.idxmin()
                        if pd.notnull(idx_min) and idx_min in self.app.sliced_df.index:
                            x_at_min = self.app.sliced_df.loc[idx_min, selected_x_col]
                            val = f"{x_at_min:.3f}" if pd.notnull(x_at_min) and isinstance(x_at_min, (int, float)) else "N/A"
                except Exception as e:
                    print(f"Error calculating stat {stat_key} for {y_col_original}: {e}")
                    val = "計算エラー"
                row_values.append(val)
            self.stats_tree.insert("", "end", values=tuple(row_values))


class BioGraphApp:
    def __init__(self, master):
        self.master = master
        master.title("バイオメカニクス グラフ表示アプリ")
        master.geometry("1000x800")

        self.create_menu(master)

        self.db_path = os.path.join(os.getcwd(), "biograph_presets.db")
        self.db_conn = None
        self.init_database()

        self.df = None; self.sliced_df = None; self.sheet_names = []; self.column_names = []
        self.df_dict = {}; self.vline_configs = []; self.current_fig = None
        self.data_output_window = None
        self.plotted_lines = {}
        self.tooltip_annotation = None

        self.aspect_ratios = {"デフォルト (6:4)": (6, 4), "4:3": (4, 3), "16:9": (16, 9), "1:1 (正方形)": (1, 1), "3:4 (縦長)": (3, 4)}
        self.default_figure_width_inches = 6
        self.legend_label_vars = {}
        self.y_legend_entries_frame = None
        self.vline_colors = ["black", "red", "blue", "green", "orange", "purple", "gray", "cyan", "magenta", "brown"]
        self.vline_linestyles = {"実線": "-", "破線": "--", "点線": ":", "一点鎖線": "-."}
        self.vline_linewidths = [0.5, 1.0, 1.5, 2.0, 2.5, 3.0, 4.0, 5.0]
        self.legend_locations = {"自動": "best", "右上": "upper right", "左上": "upper left", "右下": "lower right", "左下": "lower left", "右": "right", "中央左": "center left", "中央右": "center right", "下中央": "lower center", "上中央": "upper center", "中央": "center"}

        self._applying_preset = False
        self.loaded_preset_settings = None

        self.detect_maxima_var = tk.BooleanVar(value=False)
        self.detect_minima_var = tk.BooleanVar(value=False)

        self.plot_bg_color_choices = ["white", "lightgray", "ivory", "lightcyan", "whitesmoke", "gainsboro"]
        self.figure_bg_color_choices = ["white", "lightgray", "whitesmoke", "gainsboro", "#F0F0F0"]
        self.grid_color_choices = ["lightgray", "gray", "darkgray", "black", "red", "blue"]
        self.grid_linestyle_choices = {"実線": "-", "破線": "--", "点線": ":", "一点鎖線": "-."}
        self.grid_linewidth_choices = [0.5, 0.8, 1.0, 1.2, 1.5, 2.0]
        self.fontsize_choices = [8, 9, 10, 11, 12, 14, 16, 18, 20]


        self.plot_bg_color_var = tk.StringVar(value="white")
        self.figure_bg_color_var = tk.StringVar(value="#F0F0F0")
        self.grid_visible_var = tk.BooleanVar(value=True)
        self.grid_color_var = tk.StringVar(value="lightgray")
        self.grid_linestyle_var = tk.StringVar(value="-") 
        self.grid_linewidth_var = tk.DoubleVar(value=0.8)
        self.global_fontsize_var = tk.IntVar(value=10) # 基本フォントサイズ


        main_paned_window = ttk.PanedWindow(master, orient=tk.HORIZONTAL)
        main_paned_window.pack(fill=tk.BOTH, expand=True)
        
        controls_container_frame = ttk.Frame(main_paned_window, padding=0)
        main_paned_window.add(controls_container_frame, weight=1)

        self.controls_canvas = tk.Canvas(controls_container_frame)
        self.controls_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        scrollbar = ttk.Scrollbar(controls_container_frame, orient=tk.VERTICAL, command=self.controls_canvas.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        self.controls_canvas.configure(yscrollcommand=scrollbar.set)
        
        self.scrollable_controls_frame = ttk.Frame(self.controls_canvas, padding=5)
        self.canvas_frame_id = self.controls_canvas.create_window((0, 0), window=self.scrollable_controls_frame, anchor="nw")

        def _on_frame_configure(event):
            self.controls_canvas.configure(scrollregion=self.controls_canvas.bbox("all"))
            # キャンバス自体の幅も内部フレームに合わせる (横スクロールバーが不要な場合)
            # self.controls_canvas.itemconfig(self.canvas_frame_id, width=event.width)

        self.scrollable_controls_frame.bind("<Configure>", _on_frame_configure)

        def _on_mouse_scroll_canvas(event):
            # Windows/macOSではevent.delta, Linuxではevent.num
            if event.num == 4 or (hasattr(event, 'delta') and event.delta > 0): # Linux上スクロール or Win/Mac上スクロール
                self.controls_canvas.yview_scroll(-1, "units")
            elif event.num == 5 or (hasattr(event, 'delta') and event.delta < 0): # Linux下スクロール or Win/Mac下スクロール
                self.controls_canvas.yview_scroll(1, "units")
        
        # Canvasと内部フレームにマウスホイールイベントをバインド
        self.controls_canvas.bind("<MouseWheel>", _on_mouse_scroll_canvas) # Windows, macOS
        self.controls_canvas.bind("<Button-4>", _on_mouse_scroll_canvas)   # Linux scroll up
        self.controls_canvas.bind("<Button-5>", _on_mouse_scroll_canvas)   # Linux scroll down
        # 内部フレームにもバインドして、ウィジェット上でホイールしてもスクロールするようにする
        # (ただし、Listboxなどは自身のスクロールが優先される場合がある)
        self._bind_mousewheel_recursive(self.scrollable_controls_frame, _on_mouse_scroll_canvas)


        # --- コントロールパネルの各セクション (親を self.scrollable_controls_frame に変更) ---
        preset_frame = ttk.LabelFrame(self.scrollable_controls_frame, text="プリセット管理")
        preset_frame.pack(padx=5, pady=5, fill="x")
        ttk.Label(preset_frame, text="プリセット:").pack(side=tk.LEFT, padx=5, pady=5)
        self.preset_var = tk.StringVar()
        self.preset_combobox = ttk.Combobox(preset_frame, textvariable=self.preset_var, state="readonly", width=20)
        self.preset_combobox.pack(side=tk.LEFT, padx=5, pady=5)
        self.preset_combobox.bind("<<ComboboxSelected>>", self.load_settings_from_preset)
        self.save_preset_button = ttk.Button(preset_frame, text="現在の設定を保存...", command=self.save_current_settings_as_preset)
        self.save_preset_button.pack(side=tk.LEFT, padx=5, pady=5)
        self.delete_preset_button = ttk.Button(preset_frame, text="選択中プリセットを削除", command=self.delete_selected_preset)
        self.delete_preset_button.pack(side=tk.LEFT, padx=5, pady=5)


        file_frame = ttk.LabelFrame(self.scrollable_controls_frame, text="1. ファイル選択")
        file_frame.pack(padx=5, pady=5, fill="x")
        self.file_path_label = ttk.Label(file_frame, text="ファイルが選択されていません")
        self.file_path_label.pack(side=tk.LEFT, padx=5, pady=5, expand=True, fill="x")
        self.select_file_button = ttk.Button(file_frame, text="ファイルを選択...", command=lambda: self.load_excel_file_interactive())
        self.select_file_button.pack(side=tk.RIGHT, padx=5, pady=5)

        sheet_frame = ttk.LabelFrame(self.scrollable_controls_frame, text="2. シート選択")
        sheet_frame.pack(padx=5, pady=5, fill="x")
        ttk.Label(sheet_frame, text="シート名:").pack(side=tk.LEFT, padx=5, pady=5)
        self.sheet_var = tk.StringVar()
        self.sheet_dropdown = ttk.Combobox(sheet_frame, textvariable=self.sheet_var, state="disabled", postcommand=self.update_sheet_dropdown_options)
        self.sheet_dropdown.pack(side=tk.LEFT, padx=5, pady=5, expand=True, fill="x")
        self.sheet_dropdown.bind("<<ComboboxSelected>>", self.on_sheet_selected)

        axis_legend_frame = ttk.LabelFrame(self.scrollable_controls_frame, text="3. 軸選択と凡例名編集")
        axis_legend_frame.pack(padx=5, pady=5, fill="x")
        x_axis_sub_frame = ttk.Frame(axis_legend_frame)
        x_axis_sub_frame.pack(pady=2, fill="x")
        ttk.Label(x_axis_sub_frame, text="X軸データ列:").pack(side=tk.LEFT, padx=5)
        self.x_axis_var = tk.StringVar()
        self.x_axis_listbox = ttk.Combobox(x_axis_sub_frame, textvariable=self.x_axis_var, state="disabled", postcommand=self.update_xaxis_options)
        self.x_axis_listbox.pack(side=tk.LEFT, padx=5, expand=True, fill="x")
        self.x_axis_listbox.bind("<<ComboboxSelected>>", self.on_x_axis_selected)
        y_axis_outer_frame = ttk.Frame(axis_legend_frame)
        y_axis_outer_frame.pack(pady=2, fill="both", expand=True)
        y_list_frame = ttk.Frame(y_axis_outer_frame)
        y_list_frame.pack(side=tk.LEFT, fill="both", expand=True, padx=(0,2))
        ttk.Label(y_list_frame, text="Y軸データ列 (複数可):").pack(anchor="w")
        self.y_axis_listbox = tk.Listbox(y_list_frame, selectmode=tk.MULTIPLE, exportselection=False, state="disabled", height=5)
        self.y_axis_listbox.pack(side=tk.LEFT, fill="both", expand=True)
        y_scrollbar = ttk.Scrollbar(y_list_frame, orient=tk.VERTICAL, command=self.y_axis_listbox.yview)
        y_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.y_axis_listbox.config(yscrollcommand=y_scrollbar.set)
        self.y_axis_listbox.bind("<<ListboxSelect>>", self.on_y_axis_selected)
        self.y_legend_entries_frame = ttk.Frame(y_axis_outer_frame)
        self.y_legend_entries_frame.pack(side=tk.RIGHT, fill="both", expand=True, padx=(2,0))
        ttk.Label(self.y_legend_entries_frame, text="凡例名編集:").pack(anchor="w")


        range_frame = ttk.LabelFrame(self.scrollable_controls_frame, text="4. データ範囲指定 (オプション)")
        range_frame.pack(padx=5, pady=5, fill="x")
        ttk.Label(range_frame, text="開始行 (1から):").pack(side=tk.LEFT, padx=5, pady=2)
        self.start_row_var = tk.StringVar()
        self.start_row_entry = ttk.Entry(range_frame, textvariable=self.start_row_var, width=7)
        self.start_row_entry.pack(side=tk.LEFT, padx=5, pady=2)
        vcmd_start = (self.start_row_entry.register(self.validate_numeric_input), '%P')
        self.start_row_entry.config(validate='key', validatecommand=vcmd_start)
        ttk.Label(range_frame, text="終了行:").pack(side=tk.LEFT, padx=5, pady=2)
        self.end_row_var = tk.StringVar()
        self.end_row_entry = ttk.Entry(range_frame, textvariable=self.end_row_var, width=7)
        self.end_row_entry.pack(side=tk.LEFT, padx=5, pady=2)
        vcmd_end = (self.end_row_entry.register(self.validate_numeric_input), '%P')
        self.end_row_entry.config(validate='key', validatecommand=vcmd_end)
        ttk.Label(range_frame, text="(空欄で全範囲)").pack(side=tk.LEFT, padx=5, pady=2)

        self.marker_outer_frame = ttk.LabelFrame(self.scrollable_controls_frame, text="5. イベントマーカー設定 (最大5本)")
        self.marker_outer_frame.pack(padx=5, pady=5, fill="x")
        self.add_vline_button = ttk.Button(self.marker_outer_frame, text="縦線マーカーを追加", command=self.add_vline_entry_ui)
        self.add_vline_button.pack(pady=2)
        self.vline_entries_container = ttk.Frame(self.marker_outer_frame)
        self.vline_entries_container.pack(fill="x", expand=True, pady=2)

        data_proc_frame = ttk.LabelFrame(self.scrollable_controls_frame, text="6. データ処理")
        data_proc_frame.pack(padx=5, pady=5, fill="x")
        self.diff_button = ttk.Button(data_proc_frame, text="選択Y軸を微分", command=self.differentiate_selected_y, state="disabled")
        self.diff_button.pack(side=tk.LEFT, padx=5, pady=2)
        self.integ_button = ttk.Button(data_proc_frame, text="選択Y軸を積分", command=self.integrate_selected_y, state="disabled")
        self.integ_button.pack(side=tk.LEFT, padx=5, pady=2)

        event_detection_frame = ttk.LabelFrame(self.scrollable_controls_frame, text="7. 自動イベント検出")
        event_detection_frame.pack(padx=5, pady=5, fill="x")
        self.detect_maxima_checkbox = ttk.Checkbutton(event_detection_frame, text="最大値をプロット (赤)", variable=self.detect_maxima_var, command=self.trigger_redraw_if_possible)
        self.detect_maxima_checkbox.pack(side=tk.LEFT, padx=5, pady=2)
        self.detect_minima_checkbox = ttk.Checkbutton(event_detection_frame, text="最小値をプロット (青)", variable=self.detect_minima_var, command=self.trigger_redraw_if_possible)
        self.detect_minima_checkbox.pack(side=tk.LEFT, padx=5, pady=2)

        display_settings_frame = ttk.LabelFrame(self.scrollable_controls_frame, text="8. グラフの表示設定")
        display_settings_frame.pack(padx=5, pady=5, fill="x")
        
        # フォントサイズ設定
        font_size_frame = ttk.Frame(display_settings_frame)
        font_size_frame.pack(fill="x", padx=5, pady=2)
        ttk.Label(font_size_frame, text="基本フォントサイズ:").pack(side=tk.LEFT, padx=5)
        self.fontsize_combo = ttk.Combobox(font_size_frame, textvariable=self.global_fontsize_var, values=self.fontsize_choices, state="disabled", width=5)
        self.fontsize_combo.pack(side=tk.LEFT, padx=5)
        self.fontsize_combo.bind("<<ComboboxSelected>>", lambda e: self.trigger_redraw_if_possible())

        plot_bg_frame = ttk.Frame(display_settings_frame)
        plot_bg_frame.pack(fill="x", padx=5, pady=2)
        ttk.Label(plot_bg_frame, text="プロット背景色:").pack(side=tk.LEFT, padx=5)
        self.plot_bg_color_combo = ttk.Combobox(plot_bg_frame, textvariable=self.plot_bg_color_var, values=self.plot_bg_color_choices, state="disabled", width=12)
        self.plot_bg_color_combo.pack(side=tk.LEFT, padx=5)
        self.plot_bg_color_combo.bind("<<ComboboxSelected>>", lambda e: self.trigger_redraw_if_possible())

        figure_bg_frame = ttk.Frame(display_settings_frame)
        figure_bg_frame.pack(fill="x", padx=5, pady=2)
        ttk.Label(figure_bg_frame, text="図全体の背景色:").pack(side=tk.LEFT, padx=5)
        self.figure_bg_color_combo = ttk.Combobox(figure_bg_frame, textvariable=self.figure_bg_color_var, values=self.figure_bg_color_choices, state="disabled", width=12)
        self.figure_bg_color_combo.pack(side=tk.LEFT, padx=5)
        self.figure_bg_color_combo.bind("<<ComboboxSelected>>", lambda e: self.trigger_redraw_if_possible())

        grid_settings_frame = ttk.Frame(display_settings_frame)
        grid_settings_frame.pack(fill="x", padx=5, pady=2)
        self.grid_visible_checkbox = ttk.Checkbutton(grid_settings_frame, text="グリッドを表示", variable=self.grid_visible_var, command=self.on_grid_visibility_change)
        self.grid_visible_checkbox.pack(side=tk.LEFT, padx=5)
        self.grid_visible_checkbox.config(state="disabled")

        self.grid_color_label = ttk.Label(grid_settings_frame, text="色:")
        self.grid_color_label.pack(side=tk.LEFT, padx=(10,0))
        self.grid_color_combo = ttk.Combobox(grid_settings_frame, textvariable=self.grid_color_var, values=self.grid_color_choices, state="disabled", width=8)
        self.grid_color_combo.pack(side=tk.LEFT, padx=2)
        self.grid_color_combo.bind("<<ComboboxSelected>>", lambda e: self.trigger_redraw_if_possible())

        self.grid_style_label = ttk.Label(grid_settings_frame, text="スタイル:")
        self.grid_style_label.pack(side=tk.LEFT, padx=(10,0))
        self.grid_style_combo = ttk.Combobox(grid_settings_frame, textvariable=self.grid_linestyle_var, values=list(self.grid_linestyle_choices.keys()), state="disabled", width=6)
        self.grid_style_combo.pack(side=tk.LEFT, padx=2)
        self.grid_style_combo.bind("<<ComboboxSelected>>", lambda e: self.trigger_redraw_if_possible())
        
        self.grid_linewidth_label = ttk.Label(grid_settings_frame, text="太さ:")
        self.grid_linewidth_label.pack(side=tk.LEFT, padx=(10,0))
        self.grid_linewidth_combo = ttk.Combobox(grid_settings_frame, textvariable=self.grid_linewidth_var, values=self.grid_linewidth_choices, state="disabled", width=4)
        self.grid_linewidth_combo.pack(side=tk.LEFT, padx=2)
        self.grid_linewidth_combo.bind("<<ComboboxSelected>>", lambda e: self.trigger_redraw_if_possible())
        
        title_frame = ttk.Frame(display_settings_frame)
        title_frame.pack(fill="x", padx=5, pady=2)
        ttk.Label(title_frame, text="グラフタイトル:").pack(side=tk.LEFT, padx=5)
        self.graph_title_var = tk.StringVar()
        self.graph_title_entry = ttk.Entry(title_frame, textvariable=self.graph_title_var, state="disabled")
        self.graph_title_entry.pack(side=tk.LEFT, padx=5, expand=True, fill="x")

        xlabel_frame = ttk.Frame(display_settings_frame)
        xlabel_frame.pack(fill="x", padx=5, pady=2)
        ttk.Label(xlabel_frame, text="X軸ラベル:").pack(side=tk.LEFT, padx=5)
        self.x_axis_label_var = tk.StringVar()
        self.x_axis_label_entry = ttk.Entry(xlabel_frame, textvariable=self.x_axis_label_var, state="disabled")
        self.x_axis_label_entry.pack(side=tk.LEFT, padx=5, expand=True, fill="x")

        ylabel_frame = ttk.Frame(display_settings_frame)
        ylabel_frame.pack(fill="x", padx=5, pady=2)
        ttk.Label(ylabel_frame, text="Y軸ラベル:").pack(side=tk.LEFT, padx=5)
        self.y_axis_label_var = tk.StringVar()
        self.y_axis_label_entry = ttk.Entry(ylabel_frame, textvariable=self.y_axis_label_var, state="disabled")
        self.y_axis_label_entry.pack(side=tk.LEFT, padx=5, expand=True, fill="x")
        self.y_axis_label_var.set("値")

        aspect_ratio_frame = ttk.Frame(display_settings_frame)
        aspect_ratio_frame.pack(fill="x", padx=5, pady=2)
        ttk.Label(aspect_ratio_frame, text="アスペクト比:").pack(side=tk.LEFT, padx=5)
        self.aspect_ratio_var = tk.StringVar()
        self.aspect_ratio_dropdown = ttk.Combobox(aspect_ratio_frame, textvariable=self.aspect_ratio_var, values=list(self.aspect_ratios.keys()), state="disabled", width=15)
        self.aspect_ratio_dropdown.pack(side=tk.LEFT, padx=5)
        self.aspect_ratio_dropdown.bind("<<ComboboxSelected>>", self.on_aspect_ratio_selected)
        self.aspect_ratio_var.set(list(self.aspect_ratios.keys())[0])

        legend_loc_frame = ttk.Frame(display_settings_frame)
        legend_loc_frame.pack(fill="x", padx=5, pady=2)
        ttk.Label(legend_loc_frame, text="凡例の表示位置:").pack(side=tk.LEFT, padx=5)
        self.legend_loc_var = tk.StringVar()
        self.legend_loc_dropdown = ttk.Combobox(legend_loc_frame, textvariable=self.legend_loc_var, values=list(self.legend_locations.keys()), state="disabled", width=15)
        self.legend_loc_dropdown.pack(side=tk.LEFT, padx=5)
        self.legend_loc_dropdown.bind("<<ComboboxSelected>>", self.on_legend_loc_selected)
        self.legend_loc_var.set(list(self.legend_locations.keys())[0])
        
        self.on_grid_visibility_change()

        action_frame = ttk.Frame(self.scrollable_controls_frame)
        action_frame.pack(padx=5, pady=10, fill="x")
        self.draw_graph_button = ttk.Button(action_frame, text="グラフ描画", command=self.draw_graph, state="disabled")
        self.draw_graph_button.pack(side=tk.LEFT, padx=2, pady=5, expand=True)
        self.save_graph_button = ttk.Button(action_frame, text="グラフを保存...", command=self.save_graph, state="disabled")
        self.save_graph_button.pack(side=tk.LEFT, padx=2, pady=5, expand=True)
        self.create_table_button = ttk.Button(action_frame, text="データテーブルを作成...", command=self.show_data_table_window, state="disabled")
        self.create_table_button.pack(side=tk.LEFT, padx=2, pady=5, expand=True)
        
        self.graph_display_outer_frame = ttk.Frame(main_paned_window, padding=5)
        main_paned_window.add(self.graph_display_outer_frame, weight=3)

        self.graph_display_frame = ttk.LabelFrame(self.graph_display_outer_frame, text="グラフ表示エリア")
        self.graph_display_frame.pack(padx=5, pady=5, fill="both", expand=True)
        self.initial_graph_label = ttk.Label(self.graph_display_frame, text="ファイルと軸を選択して「グラフ描画」ボタンを押してください。")
        self.initial_graph_label.pack(padx=20, pady=20, expand=True)
        self.canvas_widget = None
        self.toolbar = None

        self.load_presets_to_combobox()
        self.master.protocol("WM_DELETE_WINDOW", self.on_app_close)
        
        self.show_about_window()

    def _bind_mousewheel_recursive(self, widget, callback):
        """指定されたウィジェットとその全ての子ウィジェットにマウスホイールイベントをバインドする"""
        widget.bind("<MouseWheel>", callback)
        widget.bind("<Button-4>", callback)
        widget.bind("<Button-5>", callback)
        for child in widget.winfo_children():
            # ListboxやEntryなど、独自のスクロールやテキスト編集を持つものは除外
            if not isinstance(child, (tk.Listbox, ttk.Entry, ttk.Combobox, ttk.Scrollbar, tk.Text)):
                self._bind_mousewheel_recursive(child, callback)


    def on_grid_visibility_change(self):
        if hasattr(self, 'grid_color_combo'): 
            is_visible = self.grid_visible_var.get()
            state = tk.NORMAL if is_visible else tk.DISABLED
            # grid_visible_checkbox が有効な場合のみ、従属するウィジェットの状態を変更
            if self.grid_visible_checkbox.instate(['!disabled']):
                self.grid_color_combo.config(state=state if is_visible else tk.DISABLED)
                self.grid_style_combo.config(state=state if is_visible else tk.DISABLED)
                self.grid_linewidth_combo.config(state=state if is_visible else tk.DISABLED)
            else: # grid_visible_checkbox 自体が無効なら、他も全て無効
                self.grid_color_combo.config(state=tk.DISABLED)
                self.grid_style_combo.config(state=tk.DISABLED)
                self.grid_linewidth_combo.config(state=tk.DISABLED)

        self.trigger_redraw_if_possible()


    def show_about_window(self):
        self._about_window = AboutAppWindow(self.master)
        self.master.wait_window(self._about_window)

    def create_menu(self, master):
        """Create the menubar with a File→終了 option."""
        menubar = tk.Menu(master)
        file_menu = tk.Menu(menubar, tearoff=0)
        file_menu.add_command(label="終了", command=self.on_app_close)
        menubar.add_cascade(label="ファイル", menu=file_menu)
        master.config(menu=menubar)

    def init_database(self):
        self.db_conn = sqlite3.connect(self.db_path)
        cursor = self.db_conn.cursor()
        cursor.execute("CREATE TABLE IF NOT EXISTS presets (id INTEGER PRIMARY KEY AUTOINCREMENT, name TEXT UNIQUE NOT NULL, settings TEXT NOT NULL)")
        self.db_conn.commit()

    def on_app_close(self):
        if not messagebox.askokcancel("終了確認", "アプリを終了しますか？", parent=self.master):
            return
        if self.db_conn:
            self.db_conn.close()
        if self.data_output_window and self.data_output_window.winfo_exists():
            self.data_output_window.destroy()
        if hasattr(self, '_about_window') and self._about_window and self._about_window.winfo_exists():
            self._about_window.destroy()
        self.master.destroy()

    def load_presets_to_combobox(self):
        if not self.db_conn: return
        cursor = self.db_conn.cursor()
        cursor.execute("SELECT name FROM presets ORDER BY name")
        presets = [row[0] for row in cursor.fetchall()]
        self.preset_combobox['values'] = presets
        self.preset_var.set("")
        if not presets: self.preset_combobox['values'] = []

    def collect_current_settings(self):
        settings = {
            'legend_labels': {orig_name: var.get() for orig_name, var in self.legend_label_vars.items()},
            'vline_markers': [{'name': item['name_var'].get(), 'color': item['color_var'].get(), 'linewidth': item['linewidth_var'].get()} for item in self.vline_configs],
            'graph_title': self.graph_title_var.get(),
            'x_axis_label': self.x_axis_label_var.get(),
            'y_axis_label': self.y_axis_label_var.get(),
            'x_axis_column': self.x_axis_var.get(),
            'y_axis_columns': [self.y_axis_listbox.get(i) for i in self.y_axis_listbox.curselection()],
            'aspect_ratio': self.aspect_ratio_var.get(),
            'legend_location': self.legend_loc_var.get(),
            'plot_bg_color': self.plot_bg_color_var.get(),
            'figure_bg_color': self.figure_bg_color_var.get(),
            'grid_visible': self.grid_visible_var.get(),
            'grid_color': self.grid_color_var.get(),
            'grid_linestyle': self.grid_linestyle_var.get(),
            'grid_linewidth': self.grid_linewidth_var.get(),
            'global_fontsize': self.global_fontsize_var.get()
        }
        return settings

    def save_current_settings_as_preset(self):
        preset_name = simpledialog.askstring("プリセット名", "プリセット名を入力してください:", parent=self.master)
        if not preset_name: return
        settings_dict = self.collect_current_settings(); settings_json = json.dumps(settings_dict)
        if not self.db_conn: self.init_database()
        cursor = self.db_conn.cursor()
        try:
            cursor.execute("SELECT id FROM presets WHERE name = ?", (preset_name,))
            existing = cursor.fetchone()
            if existing:
                if not messagebox.askyesno("上書き確認", f"プリセット '{preset_name}' は既に存在します。上書きしますか？", parent=self.master): return
                cursor.execute("UPDATE presets SET settings = ? WHERE name = ?", (settings_json, preset_name))
            else:
                cursor.execute("INSERT INTO presets (name, settings) VALUES (?, ?)", (preset_name, settings_json))
            self.db_conn.commit()
            messagebox.showinfo("成功", f"プリセット '{preset_name}' を保存しました。", parent=self.master)
            self.load_presets_to_combobox(); self.preset_var.set(preset_name)
        except sqlite3.Error as e: messagebox.showerror("データベースエラー", f"プリセットの保存に失敗しました: {e}", parent=self.master)

    def load_settings_from_preset(self, event=None):
        preset_name = self.preset_var.get()
        if not preset_name: return

        self._applying_preset = True
        if not self.db_conn: self.init_database()
        cursor = self.db_conn.cursor()
        cursor.execute("SELECT settings FROM presets WHERE name = ?", (preset_name,))
        result = cursor.fetchone()

        if result:
            settings_json = result[0]
            try:
                self.loaded_preset_settings = json.loads(settings_json)
                self.apply_settings_from_loaded_preset()
                messagebox.showinfo("成功", f"プリセット '{preset_name}' を読み込みました。\nファイルと軸を再選択してください。", parent=self.master)
            except json.JSONDecodeError:
                messagebox.showerror("エラー", "プリセットデータの読み込みに失敗しました (JSON形式エラー)。", parent=self.master)
                self.loaded_preset_settings = None
            except Exception as e:
                messagebox.showerror("エラー", f"プリセットの適用中にエラーが発生しました: {e}", parent=self.master)
                self.loaded_preset_settings = None
        else:
            messagebox.showerror("エラー", f"プリセット '{preset_name}' が見つかりません。", parent=self.master)
            self.loaded_preset_settings = None
        self._applying_preset = False


    def apply_settings_from_loaded_preset(self):
        if not self.loaded_preset_settings:
            return

        settings_dict = self.loaded_preset_settings

        if self.canvas_widget:
            self.canvas_widget.get_tk_widget().destroy(); self.canvas_widget = None
        if self.toolbar:
            self.toolbar.destroy(); self.toolbar = None
        if self.initial_graph_label:
            self.initial_graph_label.destroy(); self.initial_graph_label = None
        self.initial_graph_label = ttk.Label(self.graph_display_frame, text="ファイルと軸を選択して「グラフ描画」ボタンを押してください。")
        self.initial_graph_label.pack(padx=20, pady=20, expand=True)
        self.current_fig = None
        self.sliced_df = None

        self.file_path_label.config(text="ファイルが選択されていません")
        self.sheet_var.set("")
        self.sheet_dropdown.config(state="disabled", values=[])
        self.x_axis_var.set("")
        self.x_axis_listbox.config(state="disabled", values=[])
        self.y_axis_listbox.delete(0, tk.END)
        self.y_axis_listbox.config(state="disabled")
        self.clear_legend_entries_ui()

        self.start_row_var.set("")
        self.end_row_var.set("")

        for config_item in self.vline_configs:
            config_item['widgets_frame'].destroy()
        self.vline_configs.clear()
        self.add_vline_button.config(state="normal")

        saved_vlines = settings_dict.get('vline_markers', [])
        for vline_data in saved_vlines:
            if len(self.vline_configs) < 5:
                self.add_vline_entry_ui()
                last_config = self.vline_configs[-1]
                last_config['x_var'].set("")
                last_config['name_var'].set(vline_data.get('name', ""))
                last_config['color_var'].set(vline_data.get('color', self.vline_colors[0]))
                last_config['linewidth_var'].set(vline_data.get('linewidth', 1.5))
        
        self.graph_title_var.set(settings_dict.get('graph_title', ""))
        self.x_axis_label_var.set(settings_dict.get('x_axis_label', ""))
        self.y_axis_label_var.set(settings_dict.get('y_axis_label', "値"))

        aspect_ratio_to_set = settings_dict.get('aspect_ratio', list(self.aspect_ratios.keys())[0])
        if aspect_ratio_to_set in self.aspect_ratios:
            self.aspect_ratio_var.set(aspect_ratio_to_set)

        legend_loc_to_set = settings_dict.get('legend_location', list(self.legend_locations.keys())[0])
        if legend_loc_to_set in self.legend_locations:
            self.legend_loc_var.set(legend_loc_to_set)

        self.plot_bg_color_var.set(settings_dict.get('plot_bg_color', 'white'))
        self.figure_bg_color_var.set(settings_dict.get('figure_bg_color', '#F0F0F0'))
        self.grid_visible_var.set(settings_dict.get('grid_visible', True))
        self.grid_color_var.set(settings_dict.get('grid_color', 'lightgray'))
        saved_grid_linestyle_display = settings_dict.get('grid_linestyle', '実線') # 表示名で取得
        # 辞書からMatplotlibスタイルを探す
        found_style = False
        for display_name, style_str in self.grid_linestyle_choices.items():
            if display_name == saved_grid_linestyle_display:
                self.grid_linestyle_var.set(style_str) # StringVarにはMatplotlibのスタイル文字列をセット
                found_style = True
                break
        if not found_style:
             self.grid_linestyle_var.set('-') # 見つからなければデフォルト

        self.grid_linewidth_var.set(settings_dict.get('grid_linewidth', 0.8))
        self.global_fontsize_var.set(settings_dict.get('global_fontsize', 10))
        
        self.graph_title_entry.config(state="disabled")
        self.x_axis_label_entry.config(state="disabled")
        self.y_axis_label_entry.config(state="disabled")
        self.aspect_ratio_dropdown.config(state="disabled")
        self.legend_loc_dropdown.config(state="disabled")
        self.plot_bg_color_combo.config(state="disabled")
        self.figure_bg_color_combo.config(state="disabled")
        self.grid_visible_checkbox.config(state="disabled")
        self.fontsize_combo.config(state="disabled")
        self.on_grid_visibility_change()


        self.draw_graph_button.config(state="disabled")
        self.save_graph_button.config(state="disabled")
        self.create_table_button.config(state="disabled")
        self.diff_button.config(state="disabled") 
        self.integ_button.config(state="disabled")


    def delete_selected_preset(self):
        preset_name = self.preset_var.get()
        if not preset_name: messagebox.showwarning("未選択", "削除するプリセットが選択されていません。", parent=self.master); return
        if messagebox.askyesno("削除確認", f"プリセット '{preset_name}' を本当に削除しますか？\nこの操作は元に戻せません。", parent=self.master):
            if not self.db_conn: self.init_database()
            cursor = self.db_conn.cursor()
            try:
                cursor.execute("DELETE FROM presets WHERE name = ?", (preset_name,))
                self.db_conn.commit()
                if cursor.rowcount > 0: messagebox.showinfo("成功", f"プリセット '{preset_name}' を削除しました。", parent=self.master)
                else: messagebox.showwarning("失敗", f"プリセット '{preset_name}' の削除に失敗したか、見つかりませんでした。", parent=self.master)
                self.load_presets_to_combobox(); self.preset_var.set("")
            except sqlite3.Error as e: messagebox.showerror("データベースエラー", f"プリセットの削除に失敗しました: {e}", parent=self.master)

    def load_excel_file_interactive(self):
        filepath = filedialog.askopenfilename(title="Excelファイルを選択", filetypes=(("Excelファイル", "*.xlsx *.xls"), ("すべてのファイル", "*.*")))
        if filepath: self.load_excel_file(filepath=filepath)

    def load_excel_file(self, filepath=None):
        if filepath is None: return
        if not self._applying_preset:
            self.loaded_preset_settings = None

        try:
            xls = pd.ExcelFile(filepath)
            self.sheet_names = xls.sheet_names; self.df_dict = {name: xls.parse(name) for name in self.sheet_names}
            self.file_path_label.config(text=filepath)
            self.sheet_dropdown.config(state="readonly"); self.sheet_var.set("")
            self.x_axis_listbox.config(state="disabled"); self.x_axis_var.set("")
            self.y_axis_listbox.config(state="disabled"); self.y_axis_listbox.delete(0, tk.END)
            self.draw_graph_button.config(state="disabled"); self.save_graph_button.config(state="disabled")
            self.create_table_button.config(state="disabled")
            self.diff_button.config(state="disabled")
            self.integ_button.config(state="disabled")
            self.start_row_var.set(""); self.end_row_var.set("")
            for config_item in self.vline_configs: config_item['widgets_frame'].destroy()
            self.vline_configs.clear(); self.add_vline_button.config(state="normal")
            self.current_fig = None; self.sliced_df = None
            self.reset_display_settings_inputs()

            if self.data_output_window and self.data_output_window.winfo_exists():
                self.data_output_window.destroy()
                self.data_output_window = None

            if self.sheet_names:
                self.sheet_var.set(self.sheet_names[0])
                self.on_sheet_selected(None)
            else:
                messagebox.showwarning("警告", "選択されたExcelファイルにシートがありません。", parent=self.master)
                self.sheet_dropdown.config(state="disabled")

        except Exception as e:
            messagebox.showerror("エラー", f"ファイルの読み込みに失敗しました ({filepath}):\n{e}", parent=self.master)
            self.file_path_label.config(text="ファイルが選択されていません")
            self.current_fig = None; self.sliced_df = None
            self.reset_display_settings_inputs()
            if self.data_output_window and self.data_output_window.winfo_exists():
                self.data_output_window.destroy()
                self.data_output_window = None
            self.loaded_preset_settings = None

    def update_sheet_dropdown_options(self):
        self.sheet_dropdown['values'] = self.sheet_names if self.sheet_names else []

    def on_sheet_selected(self, event):
        selected_sheet_name = self.sheet_var.get()
        if not self._applying_preset:
             self.loaded_preset_settings = None

        if not selected_sheet_name or selected_sheet_name not in self.df_dict:
            self.x_axis_listbox.config(state="disabled"); self.x_axis_var.set("")
            self.y_axis_listbox.config(state="disabled"); self.y_axis_listbox.delete(0, tk.END)
            self.draw_graph_button.config(state="disabled"); self.create_table_button.config(state="disabled")
            self.diff_button.config(state="disabled")
            self.integ_button.config(state="disabled")
            self.reset_display_settings_inputs()
            self.current_fig = None; self.sliced_df = None
            return

        self.df = self.df_dict[selected_sheet_name].copy()
        self.column_names = self.df.columns.tolist()

        if self.column_names:
            self.x_axis_listbox.config(state="readonly"); self.x_axis_var.set("")
            self.y_axis_listbox.config(state="normal"); self.y_axis_listbox.delete(0, tk.END)
            for col_name in self.column_names:
                self.y_axis_listbox.insert(tk.END, col_name)
            self.draw_graph_button.config(state="disabled"); self.create_table_button.config(state="disabled")
            self.diff_button.config(state="normal")
            self.integ_button.config(state="normal")
            self.reset_display_settings_inputs(state="disabled") 
            self.aspect_ratio_dropdown.config(state="readonly")
            self.legend_loc_dropdown.config(state="readonly")
            self.plot_bg_color_combo.config(state="readonly")
            self.figure_bg_color_combo.config(state="readonly")
            self.grid_visible_checkbox.config(state="normal")
            self.fontsize_combo.config(state="readonly")
            self.on_grid_visibility_change()

            self.clear_legend_entries_ui()
            self.apply_preset_legend_labels_if_needed()
            self.apply_preset_axis_selections_if_needed()

        else:
            messagebox.showwarning("警告", f"シート '{selected_sheet_name}' に列がありません。", parent=self.master)
            self.x_axis_listbox.config(state="disabled"); self.x_axis_var.set("")
            self.y_axis_listbox.config(state="disabled"); self.y_axis_listbox.delete(0, tk.END)
            self.draw_graph_button.config(state="disabled"); self.create_table_button.config(state="disabled")
            self.diff_button.config(state="disabled")
            self.integ_button.config(state="disabled")
            self.reset_display_settings_inputs()
            self.current_fig = None; self.sliced_df = None

    def on_x_axis_selected(self, event):
        selected_x = self.x_axis_var.get()
        if selected_x:
            if not self._applying_preset: self.x_axis_label_var.set(selected_x) # プリセット適用中でなければラベルを更新
            self.x_axis_label_entry.config(state="normal")
            # 軸選択後に表示設定を有効化
            self.aspect_ratio_dropdown.config(state="readonly")
            self.legend_loc_dropdown.config(state="readonly")
            self.plot_bg_color_combo.config(state="readonly")
            self.figure_bg_color_combo.config(state="readonly")
            self.grid_visible_checkbox.config(state="normal")
            self.fontsize_combo.config(state="readonly")
            self.on_grid_visibility_change()


            if self.y_axis_listbox.curselection():
                if not self._applying_preset: self.update_default_graph_title()
                self.graph_title_entry.config(state="normal")
                self.y_axis_label_entry.config(state="normal")
                self.draw_graph_button.config(state="normal")
                self.create_table_button.config(state="normal")
                if not self._applying_preset: self.update_legend_entries_ui()
                self.trigger_redraw_if_possible()
            else:
                if not self._applying_preset: self.graph_title_var.set("")
                self.graph_title_entry.config(state="disabled")
                self.y_axis_label_entry.config(state="normal") 
                self.draw_graph_button.config(state="disabled")
                self.create_table_button.config(state="disabled")
                if not self._applying_preset: self.clear_legend_entries_ui()
        else:
            if not self._applying_preset: self.x_axis_label_var.set("")
            self.x_axis_label_entry.config(state="disabled")
            if not self._applying_preset: self.graph_title_var.set("")
            self.graph_title_entry.config(state="disabled")
            self.draw_graph_button.config(state="disabled")
            self.create_table_button.config(state="disabled")
            self.aspect_ratio_dropdown.config(state="disabled")
            self.legend_loc_dropdown.config(state="disabled")
            self.plot_bg_color_combo.config(state="disabled")
            self.figure_bg_color_combo.config(state="disabled")
            self.grid_visible_checkbox.config(state="disabled")
            self.fontsize_combo.config(state="disabled")
            self.on_grid_visibility_change()

            if not self._applying_preset: self.clear_legend_entries_ui()

    def on_y_axis_selected(self, event=None):
        if not self._applying_preset: self.update_legend_entries_ui()

        if self.x_axis_var.get() and self.y_axis_listbox.curselection():
            if not self._applying_preset: self.update_default_graph_title()
            self.graph_title_entry.config(state="normal")
            self.x_axis_label_entry.config(state="normal")
            self.y_axis_label_entry.config(state="normal")
            self.aspect_ratio_dropdown.config(state="readonly")
            self.legend_loc_dropdown.config(state="readonly")
            self.plot_bg_color_combo.config(state="readonly")
            self.figure_bg_color_combo.config(state="readonly")
            self.grid_visible_checkbox.config(state="normal")
            self.fontsize_combo.config(state="readonly")
            self.on_grid_visibility_change()

            self.draw_graph_button.config(state="normal")
            self.create_table_button.config(state="normal")
        elif self.x_axis_var.get(): # X軸は選択されているがY軸が未選択
            if not self._applying_preset: self.graph_title_var.set("")
            self.graph_title_entry.config(state="disabled")
            self.draw_graph_button.config(state="disabled")
            self.create_table_button.config(state="disabled")

    def update_legend_entries_ui(self):
        for widget in list(self.y_legend_entries_frame.winfo_children())[1:]:
            widget.destroy()

        selected_y_indices = self.y_axis_listbox.curselection()
        if not selected_y_indices:
            return

        new_legend_vars = {}
        preset_legend_labels = self.loaded_preset_settings.get('legend_labels', {}) if self.loaded_preset_settings else {}
        
        for i in selected_y_indices:
            original_col_name = self.y_axis_listbox.get(i)
            
            if original_col_name in self.legend_label_vars:
                 label_var = self.legend_label_vars[original_col_name]
            else:
                default_legend_name = preset_legend_labels.get(original_col_name, original_col_name)
                label_var = tk.StringVar(value=default_legend_name)

            new_legend_vars[original_col_name] = label_var

            entry_row_frame = ttk.Frame(self.y_legend_entries_frame)
            entry_row_frame.pack(fill="x", pady=1)
            display_name = original_col_name[:20] + '...' if len(original_col_name) > 20 else original_col_name
            ttk.Label(entry_row_frame, text=f"{display_name}:").pack(side=tk.LEFT, padx=(0,2))
            entry = ttk.Entry(entry_row_frame, textvariable=label_var, width=15)
            entry.pack(side=tk.LEFT, expand=True, fill="x")
        self.legend_label_vars = new_legend_vars

    def apply_preset_legend_labels_if_needed(self):
        if not self.loaded_preset_settings or 'legend_labels' not in self.loaded_preset_settings:
            return

        preset_legend_labels = self.loaded_preset_settings['legend_labels']
        missing_cols_in_current_sheet = []
        for original_col_name, legend_name in preset_legend_labels.items():
            if original_col_name not in self.column_names:
                missing_cols_in_current_sheet.append(original_col_name)
            else:
                 if original_col_name in self.legend_label_vars:
                     self.legend_label_vars[original_col_name].set(legend_name)
                 else:
                     self.legend_label_vars[original_col_name] = tk.StringVar(value=legend_name)


        if missing_cols_in_current_sheet:
            messagebox.showwarning("プリセット凡例警告",
                                   f"プリセットには以下の列の凡例名設定がありましたが、現在のシートには存在しません:\n{', '.join(missing_cols_in_current_sheet)}\nこれらの凡例設定は無視されます。",
                                   parent=self.master)

    def apply_preset_axis_selections_if_needed(self):
        """Apply axis selections saved in a preset if the columns exist."""
        if not self.loaded_preset_settings:
            return

        preset_x = self.loaded_preset_settings.get('x_axis_column')
        preset_y_cols = self.loaded_preset_settings.get('y_axis_columns', [])

        self._applying_preset = True
        if preset_x in self.column_names:
            self.x_axis_var.set(preset_x)
            self.on_x_axis_selected(None)
        valid_indices = []
        for col in preset_y_cols:
            if col in self.column_names:
                valid_indices.append(self.column_names.index(col))
        if valid_indices:
            self.y_axis_listbox.selection_clear(0, tk.END)
            for idx in valid_indices:
                self.y_axis_listbox.selection_set(idx)
            self.on_y_axis_selected(None)
        self._applying_preset = False


    def on_aspect_ratio_selected(self, event): self.trigger_redraw_if_possible()
    def on_legend_loc_selected(self, event): self.trigger_redraw_if_possible()
    def trigger_redraw_if_possible(self):
        if self.x_axis_var.get() and self.y_axis_listbox.curselection() and self.df is not None: self.draw_graph()
    def update_default_graph_title(self):
        selected_x = self.x_axis_var.get(); selected_y_indices = self.y_axis_listbox.curselection()
        if selected_x and selected_y_indices:
            selected_y_cols = [self.y_axis_listbox.get(i) for i in selected_y_indices]
            title = f"{', '.join(selected_y_cols)} vs {selected_x}"; self.graph_title_var.set(title)
    def update_xaxis_options(self):
        self.x_axis_listbox['values'] = self.column_names if self.column_names else []

    def get_figure_size(self):
        selected_ratio_name = self.aspect_ratio_var.get()
        ratio_w, ratio_h = self.aspect_ratios.get(selected_ratio_name, self.aspect_ratios[list(self.aspect_ratios.keys())[0]])
        fig_width = self.default_figure_width_inches; fig_height = fig_width * (ratio_h / ratio_w)
        return (fig_width, fig_height)

    def on_mouse_scroll(self, event):
        if event.inaxes and self.current_fig:
            ax = self.current_fig.gca()
            cur_xlim = ax.get_xlim()
            cur_ylim = ax.get_ylim()
            xdata = event.xdata
            ydata = event.ydata

            if xdata is None or ydata is None:
                return

            zoom_factor = 1.1
            
            if event.button == 'up':
                scale_factor = 1 / zoom_factor
            elif event.button == 'down':
                scale_factor = zoom_factor
            else:
                return

            new_width = (cur_xlim[1] - cur_xlim[0]) * scale_factor
            new_height = (cur_ylim[1] - cur_ylim[0]) * scale_factor

            relx = (cur_xlim[1] - xdata)/(cur_xlim[1] - cur_xlim[0]) if (cur_xlim[1] - cur_xlim[0]) != 0 else 0.5
            rely = (cur_ylim[1] - ydata)/(cur_ylim[1] - cur_ylim[0]) if (cur_ylim[1] - cur_ylim[0]) != 0 else 0.5


            ax.set_xlim([xdata - new_width * (1-relx), xdata + new_width * (relx)])
            ax.set_ylim([ydata - new_height * (1-rely), ydata + new_height * (rely)])
            self.canvas_widget.draw_idle()

    def draw_graph(self):
        selected_x_column = self.x_axis_var.get()
        selected_y_indices = self.y_axis_listbox.curselection()
        if not selected_x_column: messagebox.showwarning("警告", "X軸データ列を選択してください.", parent=self.master); return
        if not selected_y_indices: messagebox.showwarning("警告", "Y軸データ列を1つ以上選択してください.", parent=self.master); return
        if self.df is None: messagebox.showerror("エラー", "データが読み込まれていません.", parent=self.master); return

        graph_title = self.graph_title_var.get(); x_label = self.x_axis_label_var.get(); y_label = self.y_axis_label_var.get()
        plot_labels = {original_name: str_var.get() for original_name, str_var in self.legend_label_vars.items()}
        selected_y_columns_original = [self.y_axis_listbox.get(i) for i in selected_y_indices]
        start_row_str, end_row_str = self.start_row_var.get(), self.end_row_var.get()
        current_df = self.df
        selected_legend_loc_name = self.legend_loc_var.get()
        legend_loc_code = self.legend_locations.get(selected_legend_loc_name, 'best')
        base_fontsize = self.global_fontsize_var.get()


        try:
            row_slice = slice(None)
            if start_row_str and end_row_str:
                start_idx, end_idx = int(start_row_str)-1, int(end_row_str)
                if start_idx < 0: start_idx = 0
                if end_idx > len(current_df): end_idx = len(current_df)
                if start_idx >= end_idx: messagebox.showwarning("警告", "開始行が終了行と同じか後です。全範囲を表示します。", parent=self.master); row_slice = slice(None)
                else: row_slice = slice(start_idx, end_idx)
            elif start_row_str:
                start_idx = int(start_row_str)-1
                if start_idx < 0: start_idx = 0
                if start_idx >= len(current_df): messagebox.showwarning("警告", "開始行がデータ範囲を超えています。全範囲を表示します。", parent=self.master); row_slice = slice(None)
                else: row_slice = slice(start_idx, None)
            elif end_row_str:
                end_idx = int(end_row_str)
                if end_idx <= 0 : messagebox.showwarning("警告", "終了行が不正です。全範囲を表示します。", parent=self.master); row_slice = slice(None)
                else:
                    if end_idx > len(current_df): end_idx = len(current_df)
                    row_slice = slice(None, end_idx)
            self.sliced_df = current_df.iloc[row_slice].copy()
            if self.sliced_df.empty: messagebox.showwarning("警告", "指定された行範囲にデータがありません。", parent=self.master); self.sliced_df = None; return
        except ValueError: messagebox.showerror("エラー", "開始行または終了行には数値を入力してください。", parent=self.master); self.sliced_df = None; return
        except Exception as e: messagebox.showerror("エラー", f"データ範囲の処理中にエラー: {e}", parent=self.master); self.sliced_df = None; return

        if self.canvas_widget: self.canvas_widget.get_tk_widget().destroy()
        if self.toolbar: self.toolbar.destroy(); self.toolbar = None
        if self.initial_graph_label: self.initial_graph_label.destroy(); self.initial_graph_label = None

        fig_size = self.get_figure_size()
        self.current_fig = Figure(figsize=fig_size, dpi=100)
        self.current_fig.patch.set_facecolor(self.figure_bg_color_var.get())
        ax = self.current_fig.add_subplot(111); ax.clear()
        ax.set_facecolor(self.plot_bg_color_var.get())
        self.plotted_lines.clear()

        try:
            x_data = self.sliced_df[selected_x_column]
            for y_col_original in selected_y_columns_original:
                y_data = self.sliced_df[y_col_original]
                if not pd.api.types.is_numeric_dtype(x_data): messagebox.showerror("エラー", f"X軸の列 '{selected_x_column}' は数値データではありません。", parent=self.master); self.sliced_df=None; return
                if not pd.api.types.is_numeric_dtype(y_data): messagebox.showwarning("警告", f"Y軸の列 '{y_col_original}' は数値データではありません。スキップします。", parent=self.master); continue
                legend_name_to_use = plot_labels.get(y_col_original, y_col_original)
                line, = ax.plot(x_data, y_data, label=legend_name_to_use)
                self.plotted_lines[legend_name_to_use] = line

            if self.detect_maxima_var.get():
                for y_col_original in selected_y_columns_original:
                    if y_col_original in self.sliced_df and pd.api.types.is_numeric_dtype(self.sliced_df[y_col_original]):
                        y_series = self.sliced_df[y_col_original].dropna()
                        if not y_series.empty:
                            try:
                                idx_max = y_series.idxmax()
                                if pd.notnull(idx_max) and idx_max in self.sliced_df.index:
                                     x_at_max = self.sliced_df.loc[idx_max, selected_x_column]
                                     y_at_max = y_series.loc[idx_max]
                                     ax.scatter(x_at_max, y_at_max, color='red', marker='o', s=50, zorder=5)
                            except Exception as e_max:
                                print(f"Error plotting maxima for {y_col_original}: {e_max}")


            if self.detect_minima_var.get():
                for y_col_original in selected_y_columns_original:
                    if y_col_original in self.sliced_df and pd.api.types.is_numeric_dtype(self.sliced_df[y_col_original]):
                        y_series = self.sliced_df[y_col_original].dropna()
                        if not y_series.empty:
                            try:
                                idx_min = y_series.idxmin()
                                if pd.notnull(idx_min) and idx_min in self.sliced_df.index:
                                    x_at_min = self.sliced_df.loc[idx_min, selected_x_column]
                                    y_at_min = y_series.loc[idx_min]
                                    ax.scatter(x_at_min, y_at_min, color='blue', marker='o', s=50, zorder=5)
                            except Exception as e_min:
                                print(f"Error plotting minima for {y_col_original}: {e_min}")


            for vline_config_item in self.vline_configs:
                x_val_str = vline_config_item['x_var'].get(); name_val = vline_config_item['name_var'].get()
                color_val = vline_config_item['color_var'].get(); linewidth_val = vline_config_item['linewidth_var'].get()
                if x_val_str:
                    try:
                        x_coord = float(x_val_str)
                        ax.axvline(x=x_coord, color=color_val, linewidth=linewidth_val, linestyle='--')
                        if name_val:
                            y_min, y_max = ax.get_ylim(); text_y_position = y_min + (y_max - y_min) * 0.9
                            x_min, x_max = ax.get_xlim(); text_x_offset = (x_max - x_min) * 0.01
                            ax.text(x_coord + text_x_offset, text_y_position, name_val, color=color_val, fontsize=base_fontsize -1, ha='left', va='center') # マーカー名もフォントサイズ適用
                    except ValueError: messagebox.showwarning("警告", f"マーカーのX座標 '{x_val_str}' は数値である必要があります。", parent=self.master)
                    except Exception as e_v: messagebox.showwarning("警告", f"マーカー '{name_val}' の描画中にエラー: {e_v}", parent=self.master)

            # フォントサイズ適用
            ax.set_title(graph_title, fontsize=base_fontsize + 2)
            ax.set_xlabel(x_label, fontsize=base_fontsize)
            ax.set_ylabel(y_label, fontsize=base_fontsize)
            ax.tick_params(axis='x', labelsize=base_fontsize -1)
            ax.tick_params(axis='y', labelsize=base_fontsize -1)


            if selected_y_columns_original and ax.get_lines():
                legend = ax.legend(loc=legend_loc_code, fontsize=base_fontsize -1) # 凡例にもフォントサイズ適用
                if legend:
                    for legline, legtext in zip(legend.get_lines(), legend.get_texts()):
                        label_of_legtext = legtext.get_text()
                        if label_of_legtext in self.plotted_lines:
                            original_line = self.plotted_lines[label_of_legtext]
                            legline.set_picker(5)
                            if original_line.get_visible():
                                legline.set_alpha(1.0)
                            else:
                                legline.set_alpha(0.2)

            if self.grid_visible_var.get():
                grid_linestyle_key = self.grid_linestyle_var.get() # これは表示名（例：「実線」）
                grid_linestyle_str = self.grid_linestyle_choices.get(grid_linestyle_key, '-') # Matplotlibスタイルへ変換
                ax.grid(True, color=self.grid_color_var.get(), linestyle=grid_linestyle_str, linewidth=self.grid_linewidth_var.get())
            else:
                ax.grid(False)
            
            self.current_fig.tight_layout()

            self.canvas_widget = FigureCanvasTkAgg(self.current_fig, master=self.graph_display_frame)
            self.canvas_widget.draw()
            self.canvas_widget.get_tk_widget().pack(side=tk.TOP, fill=tk.BOTH, expand=True)

            if self.toolbar: self.toolbar.destroy()
            self.toolbar = NavigationToolbar2Tk(self.canvas_widget, self.graph_display_frame)
            self.toolbar.update()

            self.current_fig.canvas.mpl_connect('motion_notify_event', self.on_mouse_motion)
            self.current_fig.canvas.mpl_connect('pick_event', self.on_legend_pick)
            self.current_fig.canvas.mpl_connect('scroll_event', self.on_mouse_scroll)


            self.save_graph_button.config(state="normal"); self.create_table_button.config(state="normal")
        except KeyError as e: messagebox.showerror("エラー", f"選択された列が見つかりません: {e}", parent=self.master); self.current_fig=None; self.sliced_df=None; self.save_graph_button.config(state="disabled"); self.create_table_button.config(state="disabled")
        except Exception as e: messagebox.showerror("エラー", f"グラフの描画中にエラーが発生しました:\n{e}", parent=self.master); self.current_fig=None; self.sliced_df=None; self.save_graph_button.config(state="disabled"); self.create_table_button.config(state="disabled")

    def on_mouse_motion(self, event):
        if self.current_fig and event.inaxes == self.current_fig.gca():
            ax = self.current_fig.gca()
            if self.tooltip_annotation:
                self.tooltip_annotation.remove()
                self.tooltip_annotation = None

            min_dist_sq = float('inf')
            closest_line_info = None

            for line in ax.get_lines():
                if not line.get_visible() or not hasattr(line, 'get_xdata'): continue

                xdata, ydata = line.get_data()
                if len(xdata) == 0: continue

                for i in range(len(xdata)):
                    point_display_coords = ax.transData.transform_point((xdata[i], ydata[i]))
                    dist_sq = (point_display_coords[0] - event.x)**2 + (point_display_coords[1] - event.y)**2
                    if dist_sq < min_dist_sq:
                        min_dist_sq = dist_sq
                        closest_line_info = (line, xdata[i], ydata[i])

            if closest_line_info and min_dist_sq < 20**2:
                line, x_val, y_val = closest_line_info
                text = f"{line.get_label()}\nX: {x_val:.3f}\nY: {y_val:.3f}"
                self.tooltip_annotation = ax.annotate(text,
                                                      xy=(x_val, y_val),
                                                      xytext=(10, 10),
                                                      textcoords="offset points",
                                                      bbox=dict(boxstyle="round,pad=0.4", fc="lightyellow", alpha=0.8, ec="gray"),
                                                      arrowprops=dict(arrowstyle="->", connectionstyle="arc3,rad=.2", color='gray'))
                if self.canvas_widget: self.canvas_widget.draw_idle()
        elif self.tooltip_annotation:
            self.tooltip_annotation.remove()
            self.tooltip_annotation = None
            if self.canvas_widget: self.canvas_widget.draw_idle()


    def on_legend_pick(self, event):
        leg_artist = event.artist
        if not self.current_fig: return
        ax = self.current_fig.gca()
        legend = ax.get_legend()
        if not legend: return

        clicked_legend_label = None
        try:
            handles = legend.legendHandles if hasattr(legend, 'legendHandles') else legend.legend_handles
            for i, handle in enumerate(handles):
                if handle is leg_artist:
                    clicked_legend_label = legend.get_texts()[i].get_text()
                    break
        except AttributeError:
             for i, leg_line_in_legend in enumerate(legend.get_lines()):
                if leg_line_in_legend is leg_artist:
                    clicked_legend_label = legend.get_texts()[i].get_text()
                    break

        if clicked_legend_label and clicked_legend_label in self.plotted_lines:
            original_line_to_toggle = self.plotted_lines[clicked_legend_label]
            visible = not original_line_to_toggle.get_visible()
            original_line_to_toggle.set_visible(visible)
            
            if visible:
                leg_artist.set_alpha(1.0)
            else:
                leg_artist.set_alpha(0.2)
            if self.canvas_widget: self.canvas_widget.draw_idle()


    def save_graph(self):
        if self.current_fig is None: messagebox.showwarning("警告", "保存するグラフがありません。", parent=self.master); return
        file_path = filedialog.asksaveasfilename(title="グラフを保存", defaultextension=".png", filetypes=(("PNGファイル", "*.png"), ("PDFファイル", "*.pdf"), ("すべてのファイル", "*.*")))
        if not file_path: return
        try:
            self.current_fig.savefig(file_path, bbox_inches='tight', dpi=300)
            messagebox.showinfo("成功", f"グラフを {file_path} に保存しました。", parent=self.master)
        except Exception as e: messagebox.showerror("エラー", f"グラフの保存に失敗しました:\n{e}", parent=self.master)

    def show_data_table_window(self):
        if self.data_output_window is None or not self.data_output_window.winfo_exists():
            if self.sliced_df is not None and not self.sliced_df.empty:
                self.data_output_window = DataOutputWindow(self.master, self)
                self.data_output_window.grab_set()
            else:
                messagebox.showwarning("データなし", "表示するテーブルデータがありません。まずグラフを描画してください。", parent=self.master)
        else:
            self.data_output_window.lift()

    def validate_numeric_input(self, P):
        if P == "" or (P.isdigit() and int(P) > 0): return True
        return False

    def add_vline_entry_ui(self):
        if len(self.vline_configs) >= 5: messagebox.showinfo("情報", "縦線マーカーは最大5本までです。", parent=self.master); self.add_vline_button.config(state="disabled"); return
        marker_frame = ttk.Frame(self.vline_entries_container); marker_frame.pack(fill="x", pady=2)
        line_num = len(self.vline_configs) + 1
        ttk.Label(marker_frame, text=f"線{line_num}:").pack(side=tk.LEFT, padx=(0,2))
        ttk.Label(marker_frame, text="X座標:").pack(side=tk.LEFT, padx=(0,2)); x_var = tk.StringVar(); x_entry = ttk.Entry(marker_frame, textvariable=x_var, width=6); x_entry.pack(side=tk.LEFT, padx=(0,5))
        ttk.Label(marker_frame, text="名称:").pack(side=tk.LEFT, padx=(0,2)); name_var = tk.StringVar(); name_entry = ttk.Entry(marker_frame, textvariable=name_var, width=10); name_entry.pack(side=tk.LEFT, padx=(0,5))
        ttk.Label(marker_frame, text="色:").pack(side=tk.LEFT, padx=(0,2)); color_var = tk.StringVar(value=self.vline_colors[line_num % len(self.vline_colors)]); color_combo = ttk.Combobox(marker_frame, textvariable=color_var, values=self.vline_colors, width=7, state="readonly"); color_combo.pack(side=tk.LEFT, padx=(0,5))
        ttk.Label(marker_frame, text="太さ:").pack(side=tk.LEFT, padx=(0,2)); linewidth_var = tk.DoubleVar(value=1.5); linewidth_combo = ttk.Combobox(marker_frame, textvariable=linewidth_var, values=self.vline_linewidths, width=4, state="readonly"); linewidth_combo.pack(side=tk.LEFT, padx=(0,5))
        remove_button = ttk.Button(marker_frame, text="削除", width=5, command=lambda mf=marker_frame: self.remove_vline_entry_ui(mf)); remove_button.pack(side=tk.RIGHT, padx=2)
        config_item = {'widgets_frame': marker_frame, 'x_var': x_var, 'name_var': name_var, 'color_var': color_var, 'linewidth_var': linewidth_var}
        self.vline_configs.append(config_item)
        if len(self.vline_configs) >= 5: self.add_vline_button.config(state="disabled")

    def remove_vline_entry_ui(self, marker_frame_to_remove):
        marker_frame_to_remove.destroy()
        item_to_delete = next((item for item in self.vline_configs if item['widgets_frame'] == marker_frame_to_remove), None)
        if item_to_delete: self.vline_configs.remove(item_to_delete)
        self.add_vline_button.config(state="normal")
        for i, config in enumerate(self.vline_configs):
            first_label_in_row = config['widgets_frame'].winfo_children()[0]
            if isinstance(first_label_in_row, ttk.Label):
                first_label_in_row.config(text=f"線{i + 1}:")

    def reset_display_settings_inputs(self, state="disabled"):
        self.graph_title_var.set(""); self.x_axis_label_var.set(""); self.y_axis_label_var.set("値")
        self.graph_title_entry.config(state=state); self.x_axis_label_entry.config(state=state); self.y_axis_label_entry.config(state=state)
        self.aspect_ratio_dropdown.config(state=state); self.aspect_ratio_var.set(list(self.aspect_ratios.keys())[0])
        self.legend_loc_dropdown.config(state=state); self.legend_loc_var.set(list(self.legend_locations.keys())[0])
        
        self.plot_bg_color_var.set("white")
        self.figure_bg_color_var.set("#F0F0F0")
        self.grid_visible_var.set(True)
        self.grid_color_var.set("lightgray")
        self.grid_linestyle_var.set("-")
        self.grid_linewidth_var.set(0.8)
        self.global_fontsize_var.set(10) # フォントサイズもリセット
        
        self.plot_bg_color_combo.config(state=state)
        self.figure_bg_color_combo.config(state=state)
        self.grid_visible_checkbox.config(state=state)
        self.fontsize_combo.config(state=state)
        self.on_grid_visibility_change()

        self.clear_legend_entries_ui()

    def clear_legend_entries_ui(self):
        for widget in self.y_legend_entries_frame.winfo_children():
            if widget not in [self.y_legend_entries_frame.winfo_children()[0]]:
                widget.destroy()
        if not self.y_axis_listbox.curselection():
             self.legend_label_vars.clear()

    def _process_data(self, operation_type):
        selected_indices = self.y_axis_listbox.curselection()
        if not selected_indices:
            messagebox.showwarning("未選択", "処理対象のY軸データ列を1つ選択してください。", parent=self.master)
            return
        if len(selected_indices) > 1:
            messagebox.showwarning("複数選択", "処理対象のY軸データ列は1つだけ選択してください。", parent=self.master)
            return
        
        selected_index = selected_indices[0]
        y_col_original = self.y_axis_listbox.get(selected_index)
        x_col = self.x_axis_var.get()

        if not x_col:
            messagebox.showwarning("X軸未選択", "計算にはX軸（通常は時間）の選択が必要です。", parent=self.master)
            return
        if self.df is None:
            messagebox.showerror("データなし", "データが読み込まれていません。", parent=self.master)
            return
        if x_col not in self.df.columns or y_col_original not in self.df.columns:
            messagebox.showerror("列エラー", "選択されたX軸またはY軸の列がデータフレームに存在しません。", parent=self.master)
            return

        x_data = self.df[x_col].dropna()
        y_data = self.df[y_col_original].dropna()
        
        common_index = x_data.index.intersection(y_data.index)
        x_data = x_data.loc[common_index]
        y_data = y_data.loc[common_index]

        if not pd.api.types.is_numeric_dtype(x_data) or not pd.api.types.is_numeric_dtype(y_data):
            messagebox.showerror("データ型エラー", "X軸とY軸のデータは数値である必要があります。", parent=self.master)
            return
        if len(x_data) < 2:
            messagebox.showwarning("データ不足", "計算には少なくとも2つのデータポイントが必要です。", parent=self.master)
            return

        try:
            if operation_type == 'diff':
                result_data = np.gradient(y_data.to_numpy(), x_data.to_numpy())
                suffix = '_deriv'
                op_name = '微分'
            elif operation_type == 'integ':
                result_data = cumulative_trapezoid(y_data.to_numpy(), x_data.to_numpy(), initial=0)
                suffix = '_integ'
                op_name = '積分'
            else:
                return

            new_col_name = f"{y_col_original}{suffix}"
            count = 1
            while new_col_name in self.df.columns:
                count += 1
                new_col_name = f"{y_col_original}{suffix}{count}"

            self.df[new_col_name] = pd.Series(result_data, index=common_index)
            self.df[new_col_name] = self.df[new_col_name].reindex(self.df.index)


            self.update_column_lists_ui(new_col_name, y_col_original, op_name)
            messagebox.showinfo("処理完了", f"'{y_col_original}' の{op_name}計算を行い、\n'{new_col_name}' として追加しました。", parent=self.master)

        except ImportError:
             messagebox.showerror("ライブラリエラー", "積分機能には 'SciPy' ライブラリが必要です。\n'pip install scipy' でインストールしてください。", parent=self.master)
        except Exception as e:
            messagebox.showerror("計算エラー", f"{op_name}処理中にエラーが発生しました:\n{e}", parent=self.master)


    def differentiate_selected_y(self):
        self._process_data('diff')

    def integrate_selected_y(self):
        self._process_data('integ')

    def update_column_lists_ui(self, new_col_name, original_col_name, op_name):
        self.column_names = self.df.columns.tolist()

        current_x = self.x_axis_var.get()
        self.x_axis_listbox['values'] = self.column_names
        if current_x in self.column_names:
            self.x_axis_var.set(current_x)
        else:
            self.x_axis_var.set("")

        self.y_axis_listbox.insert(tk.END, new_col_name)

        original_legend_name_var = self.legend_label_vars.get(original_col_name)
        if original_legend_name_var:
            original_legend_name = original_legend_name_var.get()
        else:
            original_legend_name = original_col_name
        
        new_legend_name = f"{original_legend_name} ({op_name})"
        self.legend_label_vars[new_col_name] = tk.StringVar(value=new_legend_name)
        
        if self.y_axis_listbox.curselection() and new_col_name in [self.y_axis_listbox.get(i) for i in self.y_axis_listbox.curselection()]:
            self.update_legend_entries_ui()

if __name__ == '__main__':
    root = tk.Tk()
    app = BioGraphApp(root)
    root.mainloop()
