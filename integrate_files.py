import tkinter as tk
from tkinter import filedialog, messagebox
import os
import pandas as pd
import numpy as np


def integrate_files():
    """Integrate TXT files from a folder into a single Excel file."""
    # --- 1. 入力フォルダの選択 ---
    root = tk.Tk()
    root.withdraw()
    input_folder_path = filedialog.askdirectory(
        title="統合するテキストファイルが含まれるフォルダを選択してください"
    )

    if not input_folder_path:
        messagebox.showinfo("情報", "フォルダが選択されませんでした。処理を中止します。")
        return

    # --- 2. 出力Excelファイルの指定 ---
    output_excel_path = filedialog.asksaveasfilename(
        title="統合結果を保存するExcelファイル名を指定してください",
        defaultextension=".xlsx",
        filetypes=[("Excelファイル", "*.xlsx"), ("すべてのファイル", "*.*")],
    )

    if not output_excel_path:
        messagebox.showinfo("情報", "出力ファイル名が指定されませんでした。処理を中止します。")
        return

    all_data_xyz_list = []
    file_names_for_header = []
    first_file_frames = None

    try:
        file_list = [
            f
            for f in os.listdir(input_folder_path)
            if os.path.isfile(os.path.join(input_folder_path, f))
        ]

        if not file_list:
            messagebox.showinfo("情報", "選択されたフォルダにファイルが見つかりませんでした。")
            return

        for file_name in sorted(file_list):
            file_path = os.path.join(input_folder_path, file_name)

            try:
                with open(file_path, "r", encoding="utf-8") as f:
                    lines = f.readlines()
            except UnicodeDecodeError:
                try:
                    with open(file_path, "r", encoding="shift-jis") as f:
                        lines = f.readlines()
                except Exception as e:
                    print(
                        f"ファイル '{file_name}' の読み込みに失敗しました (エンコーディングエラー): {e}。スキップします。"
                    )
                    continue
            except Exception as e:
                print(f"ファイル '{file_name}' の読み込み中にエラーが発生しました: {e}。スキップします。")
                continue

            if len(lines) < 4:
                print(f"ファイル '{file_name}' は基本的なデータ形式が異なります (4行未満)。スキップします。")
                continue

            base_file_name = os.path.splitext(file_name)[0]

            data_rows_for_current_file = []
            current_file_frames = []

            for line_number, line in enumerate(lines[3:], start=4):
                cleaned_line = line.strip()
                if cleaned_line.startswith("-"):
                    cleaned_line = cleaned_line[1:].strip()

                if not cleaned_line:
                    continue

                parts = cleaned_line.split("\t")

                try:
                    if len(parts) > 0:
                        frame = int(parts[0])
                    else:
                        print(
                            f"ファイル '{file_name}' の {line_number}行目: フレーム番号が見つかりません。この行をスキップします。"
                        )
                        continue
                except ValueError:
                    print(
                        f"ファイル '{file_name}' の {line_number}行目: フレーム番号 '{parts[0]}' が数値に変換できません。この行をスキップします。"
                    )
                    continue

                x_val, y_val, z_val = np.nan, np.nan, np.nan

                if len(parts) > 1:
                    try:
                        x_val = float(parts[1])
                    except ValueError:
                        print(
                            f"ファイル '{file_name}' の {line_number}行目: X値 ('{parts[1]}') が数値に変換できません。NaNとして扱います。"
                        )

                if len(parts) > 2:
                    try:
                        y_val = float(parts[2])
                    except ValueError:
                        print(
                            f"ファイル '{file_name}' の {line_number}行目: Y値 ('{parts[2]}') が数値に変換できません。NaNとして扱います。"
                        )

                if len(parts) > 3:
                    try:
                        z_val = float(parts[3])
                    except ValueError:
                        print(
                            f"ファイル '{file_name}' の {line_number}行目: Z値 ('{parts[3]}') が数値に変換できません。NaNとして扱います。"
                        )

                current_file_frames.append(frame)
                data_rows_for_current_file.append({"X": x_val, "Y": y_val, "Z": z_val})

            if not data_rows_for_current_file:
                print(
                    f"ファイル '{file_name}' から有効なデータ行（フレーム番号含む）を抽出できませんでした。このファイルはスキップします。"
                )
                continue

            file_names_for_header.append(base_file_name)
            all_data_xyz_list.append(pd.DataFrame(data_rows_for_current_file))

            if first_file_frames is None and current_file_frames:
                first_file_frames = current_file_frames
            elif (
                current_file_frames
                and first_file_frames is not None
                and (
                    len(first_file_frames) != len(current_file_frames)
                    or first_file_frames != current_file_frames
                )
            ):
                messagebox.showwarning(
                    "警告",
                    f"ファイル '{file_name}' のフレーム数またはフレーム番号の並びが、基準となる最初のファイルと異なります。\n"
                    f"基準フレーム ({len(first_file_frames)}行) に合わせて処理を続行します。",
                )
            elif not current_file_frames and first_file_frames is not None:
                messagebox.showwarning(
                    "警告",
                    f"ファイル '{file_name}' から有効なフレーム番号を抽出できませんでした。\n"
                    f"このファイルのXYZデータは、基準フレーム ({len(first_file_frames)}行) に合わせて処理されますが、不整合が生じる可能性があります。",
                )

        if not all_data_xyz_list or first_file_frames is None:
            messagebox.showinfo(
                "情報", "処理対象となる有効なデータ（フレーム情報含む）を持つファイルが見つかりませんでした。Excelファイルは作成されません。"
            )
            return

        output_df = pd.DataFrame({"Frame": first_file_frames})

        for i, df_xyz in enumerate(all_data_xyz_list):
            base_name = file_names_for_header[i]

            if len(df_xyz) == len(first_file_frames):
                df_xyz_adjusted = df_xyz.reset_index(drop=True)
            elif len(df_xyz) > len(first_file_frames):
                df_xyz_adjusted = df_xyz.iloc[: len(first_file_frames)].reset_index(drop=True)
            else:
                temp_df = pd.DataFrame(index=range(len(first_file_frames)))
                temp_df = temp_df.merge(df_xyz.reset_index(drop=True), left_index=True, right_index=True, how="left")
                df_xyz_adjusted = temp_df[["X", "Y", "Z"]]

            output_df[f"{base_name}_X"] = df_xyz_adjusted["X"]
            output_df[f"{base_name}_Y"] = df_xyz_adjusted["Y"]
            output_df[f"{base_name}_Z"] = df_xyz_adjusted["Z"]

        output_df.to_excel(output_excel_path, index=False)
        messagebox.showinfo("成功", f"ファイルの統合が完了しました。\n出力先: {output_excel_path}")

    except Exception as e:
        messagebox.showerror("エラー", f"処理中に予期せぬエラーが発生しました:\n{e}")
        import traceback

        print(f"詳細エラー: {traceback.format_exc()}")


if __name__ == "__main__":
    integrate_files()
