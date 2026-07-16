from itertools import count
import shutil
import pandas as pd
import tkinter as tk
from pyparsing import col
import win32com.client as win32
from tkinter import filedialog, messagebox,ttk
import ctypes
ctypes.windll.shcore.SetProcessDpiAwareness(1)  # 解决 Windows DPI 缩放导致的模糊问题
import os
import re
import time


OUTPUT_ROW_HEIGHT = 24.95


# 读取 Excel 文件
def read_excel_file(file_path, header_row, sheet_name=0):
    df = pd.read_excel(file_path, sheet_name=sheet_name, header=header_row)

    df.columns = ['部位', '编号', '长', '宽', '厚', '片数', '平方数','小面加工','箱号','备注']

    text_columns = ['部位', '编号', '小面加工', '箱号', '备注']
    for col in text_columns:
        df[col] = df[col].apply(
            lambda x: None if pd.isna(x) or str(x).strip().lower() == 'nan' else str(x).strip()
        )

    for col in ['长', '宽', '片数', '平方数']:
        df[col] = pd.to_numeric(df[col], errors='coerce')

    df['厚'] = df['厚'].apply(normalize_thickness_input)

    return df
def get_sheet_names(file_path):
    excel_file = pd.ExcelFile(file_path)
    return excel_file.sheet_names

def normalize_thickness_input(value):
    if value is None:
        return ""

    if pd.isna(value):
        return ""

    if isinstance(value, (int, float)):
        num = float(value)
        if num.is_integer():
            return str(int(num))
        return str(num).rstrip("0").rstrip(".")

    text = str(value).strip()

    if not text or text.lower() == "nan":
        return ""

    text = (
        text.replace("～", "~")
            .replace("－", "-")
            .replace("—", "-")
            .replace("–", "-")
    )

    if re.fullmatch(r"\d+\.0", text):
        text = text[:-2]

    return text

def open_multi_sheet_dialog(root, sheet_names):
    dialog = tk.Toplevel(root)
    dialog.title("多 sheet 参数设置")
    dialog.geometry("1160x420")
    dialog.grab_set()

    result = {}
    row_vars = {}

    title_label = tk.Label(dialog, text="请为每个 sheet 设置参数", font=("Arial", 11, "bold"))
    title_label.pack(pady=10)

    table_frame = tk.Frame(dialog)
    table_frame.pack(fill="both", expand=True, padx=20, pady=10)

    headers = [
        "是否转换", "Sheet名", "纹路方向", "尺寸单位", "计价方式",
        "Header行", "变更厚度", "厚度", "重算片数"
    ]
    for col, text in enumerate(headers):
        tk.Label(
            table_frame,
            text=text,
            font=("Arial", 10, "bold")
        ).grid(row=0, column=col, padx=10, pady=8, sticky="w")

    for i, sheet_name in enumerate(sheet_names, start=1):
        enabled_var = tk.BooleanVar(value=True)
        grain_var = tk.StringVar(value="无纹路")
        unit_var = tk.StringVar(value="mm")
        pricing_var = tk.StringVar(value="平方数")
        header_var = tk.StringVar(value="1")
        override_thickness_var = tk.BooleanVar(value=False)
        thickness_var = tk.StringVar(value="")
        recalc_piece_var = tk.BooleanVar(value=True)

        tk.Checkbutton(table_frame, variable=enabled_var).grid(
            row=i, column=0, padx=10, pady=8
        )

        tk.Label(
            table_frame,
            text=sheet_name,
            anchor="w",
            width=16
        ).grid(row=i, column=1, padx=10, pady=8, sticky="w")

        grain_combo = ttk.Combobox(
            table_frame,
            textvariable=grain_var,
            values=["无纹路", "左右向", "上下向"],
            state="readonly",
            width=10
        )
        grain_combo.grid(row=i, column=2, padx=10, pady=8)

        unit_combo = ttk.Combobox(
            table_frame,
            textvariable=unit_var,
            values=["mm", "cm"],
            state="readonly",
            width=8
        )
        unit_combo.grid(row=i, column=3, padx=10, pady=8)

        pricing_combo = ttk.Combobox(
            table_frame,
            textvariable=pricing_var,
            values=["平方数", "米"],
            state="readonly",
            width=10
        )
        pricing_combo.grid(row=i, column=4, padx=10, pady=8)

        header_entry = tk.Entry(
            table_frame,
            textvariable=header_var,
            width=10
        )
        header_entry.grid(row=i, column=5, padx=10, pady=8)

        thickness_entry = tk.Entry(
            table_frame,
            textvariable=thickness_var,
            width=8,
            state="disabled"
        )
        thickness_entry.grid(row=i, column=7, padx=10, pady=8)

        def toggle_thickness_entry(entry=thickness_entry, var=override_thickness_var):
            entry.config(state="normal" if var.get() else "disabled")

        tk.Checkbutton(
            table_frame,
            variable=override_thickness_var,
            command=toggle_thickness_entry
        ).grid(row=i, column=6, padx=10, pady=8)

        tk.Checkbutton(table_frame, variable=recalc_piece_var).grid(
            row=i, column=8, padx=10, pady=8
        )

        row_vars[sheet_name] = {
            "enabled_var": enabled_var,
            "grain_var": grain_var,
            "unit_var": unit_var,
            "pricing_var": pricing_var,
            "header_var": header_var,
            "override_thickness_var": override_thickness_var,
            "thickness_var": thickness_var,
            "recalc_piece_var": recalc_piece_var
        }

    def confirm():
        try:
            for sheet_name, vars_dict in row_vars.items():
                header_row = int(vars_dict["header_var"].get().strip())
                if header_row < 0:
                    raise ValueError(f"{sheet_name} 的 header 不能小于 0")

                override_thickness = vars_dict["override_thickness_var"].get()

                result[sheet_name] = {
                    "enabled": vars_dict["enabled_var"].get(),
                    "grain_axis": vars_dict["grain_var"].get(),
                    "unit": vars_dict["unit_var"].get(),
                    "pricing_mode": vars_dict["pricing_var"].get(),
                    "header_row": header_row,
                    "override_thickness": override_thickness,
                    "thickness": (
                        normalize_thickness_input(vars_dict["thickness_var"].get())
                        if override_thickness else None
                    ),
                    "recalc_piece_count": vars_dict["recalc_piece_var"].get()
                }

            dialog.destroy()

        except Exception as e:
            messagebox.showwarning("提示", f"参数输入有误：\n{e}", parent=dialog)

    def cancel():
        result.clear()
        dialog.destroy()

    btn_frame = tk.Frame(dialog)
    btn_frame.pack(pady=12)

    tk.Button(btn_frame, text="确定", width=12, command=confirm).pack(side="left", padx=10)
    tk.Button(btn_frame, text="取消", width=12, command=cancel).pack(side="left", padx=10)

    dialog.wait_window()
    return result

def get_correct_count(code):
    code = str(code).strip()

    if not code or code.lower() == "nan":
        return None

    code = (
        code.replace("～", "~")
            .replace("〜", "~")
            .replace("∼", "~")
            .replace("˜", "~")
            .replace("﹏", "~")
            .replace("－", "-")
            .replace("—", "-")
            .replace("–", "-")
    )

    if "~" in code:
        tilde_index = code.find("~")

        left_dash = code.rfind("-", 0, tilde_index)
        right_dash = code.find("-", tilde_index)

        start = left_dash + 1 if left_dash != -1 else 0
        end = right_dash if right_dash != -1 else len(code)

        range_part = code[start:end].strip()

        match = re.fullmatch(
            r"([A-Za-z]*)(\d+)([A-Za-z]*)\s*~\s*([A-Za-z]*)(\d+)([A-Za-z]*)",
            range_part
        )

        if match:
            prefix1, num1, suffix1, prefix2, num2, suffix2 = match.groups()

            if prefix1 and prefix2 and prefix1 != prefix2:
                print(f"⚠️ 编号范围前缀不一致: {code}")
                return None

            if suffix1 and suffix2 and suffix1 != suffix2:
                print(f"⚠️ 编号范围后缀不一致: {code}")
                return None

            a = int(num1)
            b = int(num2)

            return b - a + 1 if b >= a else None

        print(f"⚠️ 编号范围无法识别: {code}")
        return None

    if re.search(r"\d", code):
        return 1

    return None


# 片数处理
def fix_piece_count(row):
    code = row['编号']
    original_count = row['片数']

    code_str = '' if pd.isna(code) else str(code).strip()

    if '备用料' in code_str:
        return original_count

    correct_n = get_correct_count(code_str)

    if correct_n is None:
        print(f"⚠️ 编号异常: {code_str}")
        return original_count

    return correct_n


def calculate_piece_count(df):
    df['片数'] = df.apply(fix_piece_count, axis=1)
    return df

def calc_area(row, unit, pricing_mode):
    length = row['长']
    width = row['宽']
    count = row['片数']

    if pd.notna(length) and pd.notna(count):
        if pricing_mode == "米":
            if unit == "mm":
                return length * count / 1000
            elif unit == "cm":
                return length * count / 100

    if pd.notna(length) and pd.notna(width) and pd.notna(count):
        if pricing_mode == "平方数":
            if unit == "mm":
                return length * width * count / 1000000
            elif unit == "cm":
                return length * width * count / 10000

    return None

def calculate_area(df, unit, pricing_mode):
    df['平方数'] = df.apply(lambda row: calc_area(row, unit, pricing_mode), axis=1)
    df['平方数'] = df['平方数'].round(3)
    return df

def get_grain_direction(length, width, grain_axis):
    if pd.isna(length) or pd.isna(width):
        return None
    
    if grain_axis == '无纹路':
        return None

    if grain_axis == '左右向':
        if length >= width:
            return '横'
        else:
            return '竖'

    elif grain_axis == '上下向':
        if width >= length:
            return '横'
        else:
            return '竖'

    return None

def calculate_grain(df, grain_axis):
    if grain_axis == '无纹路':
        if '纹路' in df.columns:
            df = df.drop(columns=['纹路'])
        return df

    df['纹路'] = df.apply(
        lambda row: get_grain_direction(row['长'], row['宽'], grain_axis),
        axis=1
    )
    return df

# 长宽统一
def normalize_length_width(df):
    df = df.copy()

    original_length = df['长'].copy()
    original_width = df['宽'].copy()

    mask = original_width > original_length

    df.loc[mask, '长'] = original_width[mask]
    df.loc[mask, '宽'] = original_length[mask]

    print("已调换的行数：", int(mask.sum()))
    print("还有多少行是 宽 > 长：", int((df['宽'] > df['长']).sum()))

    return df

# 小面加工相关的函数
def get_shape_info(shape):
    return {
        "name": shape.Name,
        "type": getattr(shape, "AutoShapeType", None),
        "left": shape.Left,
        "top": shape.Top,
        "width": shape.Width,
        "height": shape.Height,
        "center_x": shape.Left + shape.Width / 2,
        "center_y": shape.Top + shape.Height / 2,
        "right": shape.Left + shape.Width,
        "bottom": shape.Top + shape.Height
    }

# 判断一个点是否在图形范围内
def is_point_in_shape_range(x, y, shape_info):
    return (
        shape_info["left"] <= x <= shape_info["right"]
        and shape_info["top"] <= y <= shape_info["bottom"]
    )

# 判断两个图形是否属于同一符号
def is_same_symbol(shape_info_1, shape_info_2):
    x1, y1 = shape_info_1["center_x"], shape_info_1["center_y"]
    x2, y2 = shape_info_2["center_x"], shape_info_2["center_y"]

    cond1 = is_point_in_shape_range(x1, y1, shape_info_2)
    cond2 = is_point_in_shape_range(x2, y2, shape_info_1)

    return cond1 or cond2

# 清洗图形名称
def clean_shape_name(name):
    if not name:
        return ""

    name = str(name).strip()

    # 排除底板/整组
    if (
        name.startswith("SmallFaceBlock_")
        or name.startswith("BaseRect_")
        or name == "SmallFaceBlock"
        or name == "BaseRect"
    ):
        return ""

    direction_pattern = (
        r"top_left|top_right|bottom_left|bottom_right|"
        r"left_top|right_top|left_bottom|right_bottom|"
        r"lefttop|righttop|leftbottom|rightbottom|"
        r"topleft|topright|bottomleft|bottomright|"
        r"top|bottom|left|right"
    )

    name = re.sub(
        rf"^.*?_(?=(?:{direction_pattern})_H\d+_\d+_)",
        "",
        name,
        flags=re.IGNORECASE
    )

    name = re.sub(
        rf"^(?:{direction_pattern})_",
        "",
        name,
        flags=re.IGNORECASE
    )

    name = re.sub(r"^H\d+_\d+_", "", name, flags=re.IGNORECASE)

    name = re.sub(r"_(oval|rectangle|triangle|line)$", "", name, flags=re.IGNORECASE)

    name = re.sub(r"_child\d*$", "", name, flags=re.IGNORECASE)

    if "_child" in name:
        name = name.split("_child")[0]

    name = re.sub(r"[_\s]+\d+$", "", name)

    name = re.sub(r"\s+", "", name).strip("_")

    if name.lower() in {"oval", "rectangle", "triangle", "line"}:
        return ""

    return name
# 构建列名
def build_group_column_name(group):
    cleaned_names = []

    ignore_names = {"", "oval", "rectangle", "triangle", "line"}

    for item in group:
        cleaned = clean_shape_name(item["name"])
        if cleaned and cleaned.lower() not in ignore_names:
            cleaned_names.append(cleaned)

    cleaned_names = sorted(set(cleaned_names))
    return "+".join(cleaned_names)

# 判断图形相对于矩形的位置
def get_shape_side_obj(shape_info, rect, tolerance=2):
    rect_left = rect["left"]
    rect_top = rect["top"]
    rect_right = rect["left"] + rect["width"]
    rect_bottom = rect["top"] + rect["height"]

    shape_center_x = shape_info["center_x"]
    shape_center_y = shape_info["center_y"]

    if shape_center_x < rect_left - tolerance:
        return "左"
    elif shape_center_x > rect_right + tolerance:
        return "右"
    elif shape_center_y < rect_top - tolerance:
        return "上"
    elif shape_center_y > rect_bottom + tolerance:
        return "下"
    else:
        return "重叠/未知"

# 获取图形相对于矩形的位置集合
def get_group_sides(group, rect):
    sides = set()

    for item in group:
        side = get_shape_side_obj(item, rect)
        if side not in ["重叠/未知", None]:
            sides.add(side)

    return sides

def get_group_size(group, rect, row_data):
    sides = get_group_sides(group, rect)

    total_size = 0

    for side in sides:
        if side in ["左", "右"]:
            if pd.notna(row_data["宽"]):
                total_size += row_data["宽"]
        elif side in ["上", "下"]:
            if pd.notna(row_data["长"]):
                total_size += row_data["长"]

    return total_size

def split_symbol_groups(shp):
    if not hasattr(shp, "GroupItems"):
        return None, []

    group_items = shp.GroupItems
    rect_candidates = []
    others = []

    for j in range(1, group_items.Count + 1):
        sub = group_items.Item(j)
        sub_info = get_shape_info(sub)

        if getattr(sub, "AutoShapeType", None) == 1:
            rect_candidates.append((sub, sub_info))
        else:
            others.append(sub_info)

    if not rect_candidates:
        return None, []

    # 面积最大的 rectangle 当底板
    rect_sub, rect_info = max(
        rect_candidates,
        key=lambda x: x[1]["width"] * x[1]["height"]
    )

    for sub, sub_info in rect_candidates:
        if sub is not rect_sub:
            others.append(sub_info)

    used = set()
    symbol_groups = []

    for i in range(len(others)):
        if i in used:
            continue

        current_group = [others[i]]
        used.add(i)

        for j in range(i + 1, len(others)):
            if j in used:
                continue

            if is_same_symbol(others[i], others[j]):
                current_group.append(others[j])
                used.add(j)

        symbol_groups.append(current_group)

    return rect_info, symbol_groups

def excel_row_to_df_index(excel_row, header_row):
    return excel_row - (header_row + 2)


def extract_small_face_shapes(file_path,sheet_name):
    excel = None
    wb = None

    try:
        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(os.path.abspath(file_path))
        ws = wb.Worksheets(sheet_name)

        results = []

        for i in range(1, ws.Shapes.Count + 1):
            shp = ws.Shapes.Item(i)

            if shp.TopLeftCell.Column != 8:
                continue

            rect, symbol_groups = split_symbol_groups(shp)

            if rect is None or not symbol_groups:
                continue

            results.append({
                "row": shp.TopLeftCell.Row,
                "col": shp.TopLeftCell.Column,
                "address": shp.TopLeftCell.Address,
                "rect": {
                    "left": rect["left"],
                    "top": rect["top"],
                    "width": rect["width"],
                    "height": rect["height"]
                },
                "symbol_groups": symbol_groups
            })

        results = sorted(results, key=lambda x: (x["row"], x["col"]))
        return results

    finally:
        if wb is not None:
            wb.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()
            
def extract_small_face_records(df, file_path, header_row, sheet_name):
    groups = extract_small_face_shapes(file_path, sheet_name)

    if not groups:
        print("未检测到小面加工图形")
        return [], set()

    records = []
    all_types = set()

    for g in groups:
        excel_row = g["row"]
        df_idx = excel_row_to_df_index(excel_row, header_row)

        if df_idx < 0 or df_idx >= len(df):
            continue

        rect = g["rect"]
        symbol_groups = g["symbol_groups"]
        row_data = df.loc[df_idx]

        for group in symbol_groups:
            col_name = build_group_column_name(group)
            if not col_name:
                continue
            
            print(f"[小面加工识别] Excel第{excel_row}行 -> {col_name}")

            group_size_value = get_group_size(group, rect, row_data)
            sides = get_group_sides(group, rect)
            edge_count = len(sides)

            count = row_data["片数"]
            if pd.isna(count) or count <= 0:
                count = 1

            records.append({
                "df_idx": df_idx,
                "col_name": col_name,
                "single_size": group_size_value,
                "edge_count": edge_count,
                "piece_count": count,
                "sides": set(sides)
            })

            all_types.add(col_name)

    return records, all_types

# 计算小面加工
def calculate_by_rule(single_size, edge_count, piece_count, rule, unit):
    if unit == "mm":
        meter_divisor = 1000
        threshold = 300
    elif unit == "cm":
        meter_divisor = 100
        threshold = 30
    else:
        return 0, None

    if rule == "个":
        return edge_count * piece_count, "个"

    elif rule == "米":
        if single_size <= threshold:
            return edge_count * piece_count, "个"
        else:
            return round(single_size * piece_count / meter_divisor, 3), "米"

    return 0, None

def apply_small_face_rules(df, records, pricing_rules,unit):
    df = df.copy()
    created_small_face_cols = set()

    for record in records:
        df_idx = record["df_idx"]
        col_name = record["col_name"]

        rule = pricing_rules.get(col_name, "米")

        value, actual_unit = calculate_by_rule(
            record["single_size"],
            record["edge_count"],
            record["piece_count"],
            rule,
            unit
        )

        if actual_unit is None:
            continue

        output_col = f"{col_name}({actual_unit})"
        created_small_face_cols.add(output_col)

        if output_col not in df.columns:
            df[output_col] = None

        current_value = df.at[df_idx, output_col]
        if pd.isna(current_value) or current_value == "":
            current_value = 0

        new_value = current_value + value

        if actual_unit == "个":
            new_value = int(round(new_value))
        else:
            new_value = round(new_value, 3)

        df.at[df_idx, output_col] = new_value

    # ===== 最终清洗 =====
    for col in created_small_face_cols:
        cleaned_values = []

        for val in df[col]:
            if pd.isna(val) or val == "" or val == 0 or val == 0.0:
                cleaned_values.append(None)
                continue

            if col.endswith("(个)"):
                cleaned_values.append(int(round(float(val))))
            else:
                cleaned_values.append(round(float(val), 3))

        if col.endswith("(个)"):
            df[col] = pd.Series(cleaned_values, dtype="object")
        else:
            df[col] = cleaned_values

    return df

def calculate_small_face(df, file_path, unit, header_row, sheet_name,pricing_rules):
    records, _ = extract_small_face_records(df, file_path, header_row, sheet_name)

    if not records:
        return df
    
    df = apply_small_face_rules(df, records, pricing_rules, unit)
    return df

# 获取默认输出路径
def get_default_output_path(input_path, output_dir="output"):
    # 当前 py 文件所在目录：.../version2
    current_dir = os.path.dirname(os.path.abspath(__file__))

    # 项目根目录：.../CAD_excel_transform
    project_root = os.path.dirname(current_dir)

    # 根目录下的 output 文件夹
    output_dir = os.path.join(project_root, output_dir)

    os.makedirs(output_dir, exist_ok=True)

    base_name = os.path.splitext(os.path.basename(input_path))[0]

    i = 0
    while True:
        if i == 0:
            output_file = f"{base_name}_output.xlsx"
        else:
            output_file = f"{base_name}_output_{i}.xlsx"

        output_path = os.path.join(output_dir, output_file)

        if not os.path.exists(output_path):
            return output_path

        i += 1

# 合计
def add_total_row(df):
    df = df.copy()

    if "部位" in df.columns:
        df = df[df["部位"].astype(str).str.strip() != "合计"].copy()

    total_row = {col: None for col in df.columns}
    total_row["部位"] = "合计"

    sum_cols = []

    for col in ["片数", "平方数"]:
        if col in df.columns:
            sum_cols.append(col)

    for col in df.columns:
        if isinstance(col, str) and (col.endswith("(米)") or col.endswith("(个)")):
            sum_cols.append(col)

    sum_cols = list(dict.fromkeys(sum_cols))

    for col in sum_cols:
        numeric_series = pd.to_numeric(df[col], errors="coerce")
        total = numeric_series.sum(skipna=True)

        if pd.isna(total):
            total_row[col] = None
        elif col == "片数":
            total_row[col] = int(round(total))
        elif col == "平方数":
            total_row[col] = round(total, 3)
        elif isinstance(col, str) and col.endswith("(个)"):
            total_row[col] = int(round(total))
        elif isinstance(col, str) and col.endswith("(米)"):
            total_row[col] = round(total, 3)

    df = pd.concat([df, pd.DataFrame([total_row])], ignore_index=True)
    return df

def get_unit_divisor(unit, pricing_mode):
    if pricing_mode == "平方数":
        return 1000000 if unit == "mm" else 10000
    elif pricing_mode == "米":
        return 1000 if unit == "mm" else 100
    return None


def build_area_formula(excel_row, unit, pricing_mode):
    divisor = get_unit_divisor(unit, pricing_mode)
    if divisor is None:
        return ""

    if pricing_mode == "平方数":
        return f"=C{excel_row}*D{excel_row}*F{excel_row}/{divisor}"
    elif pricing_mode == "米":
        return f"=C{excel_row}*F{excel_row}/{divisor}"

    return ""


def build_small_face_formula_from_record(record, excel_row, unit):
    sides = record.get("sides", set())

    if unit == "mm":
        divisor = 1000
        threshold = 300
    elif unit == "cm":
        divisor = 100
        threshold = 30
    else:
        return ""

    parts = []

    for side in sorted(sides):
        if side in ["上", "下"]:
            parts.append(f"C{excel_row}*F{excel_row}/{divisor}")
        elif side in ["左", "右"]:
            parts.append(f"D{excel_row}*F{excel_row}/{divisor}")

    if not parts:
        return ""

    single_size = record.get("single_size", 0)
    if single_size <= threshold:
        return ""

    return "=" + "+".join(parts)

def rgb_to_bgr_int(r, g, b):
    return r + (g << 8) + (b << 16)

def export_excel(data, output_path):
    if isinstance(data, pd.DataFrame):
        data.to_excel(output_path, index=False)

    elif isinstance(data, dict):
        with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
            for sheet_name, df in data.items():
                safe_sheet_name = str(sheet_name)[:31]  # Excel sheet 名最长 31 个字符
                df.to_excel(writer, sheet_name=safe_sheet_name, index=False)

    else:
        raise ValueError("导出数据格式不支持")

    print("已输出到：", output_path)
    return output_path

# 导出为生产单格式
def export_to_original_format(
    df,
    input_path,
    output_path,
    header_row,
    sheet_name,
    unit,
    pricing_mode,
    pricing_rules=None,
    copy_file=True
):
    excel = None
    wb = None

    try:
        if copy_file:
            copy_template_file(input_path, output_path)

        excel = win32.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(os.path.abspath(output_path))
        ws = wb.Worksheets(sheet_name)
        
        pricing_rules = pricing_rules or {}

        df_no_total = df.copy()
        if "部位" in df_no_total.columns:
            df_no_total = df_no_total[
                df_no_total["部位"].astype(str).str.strip() != "合计"
            ].copy()

        small_face_records, _ = extract_small_face_records(
            df_no_total,
            input_path,
            header_row,
            sheet_name
        )

        # {(df_idx, "circle_xxx(米)"): [record1, record2, ...]}
        small_face_formula_map = {}

        for record in small_face_records:
            rule = pricing_rules.get(record["col_name"], "米")
            value, actual_unit = calculate_by_rule(
                record["single_size"],
                record["edge_count"],
                record["piece_count"],
                rule,
                unit
            )

            if actual_unit != "米":
                continue

            output_col = f"{record['col_name']}(米)"
            key = (record["df_idx"], output_col)
            small_face_formula_map.setdefault(key, []).append(record)

        header_excel_row = header_row + 1
        data_start_row = header_row + 2
        header_top_row = header_excel_row - 1

        last_used_cell = ws.Cells.Find(
            What="*",
            After=ws.Cells(1, 1),
            LookAt=1,              # xlPart
            LookIn=-4123,          # xlFormulas
            SearchOrder=1,         # xlByRows
            SearchDirection=2,     # xlPrevious
            MatchCase=False
        )

        if last_used_cell is not None:
            template_last_row = last_used_cell.Row
        else:
            template_last_row = data_start_row

        if template_last_row >= data_start_row:
            ws.Range(f"A{data_start_row}:G{template_last_row}").ClearContents()
            ws.Range(f"I{data_start_row}:J{template_last_row}").ClearContents()

        ws.Range(f"K{header_excel_row}:XFD{template_last_row}").ClearContents()

        fixed_mapping = {
            "部位": 1,
            "编号": 2,
            "长": 3,
            "宽": 4,
            "厚": 5,
            "片数": 6,
            "平方数": 7,
            "箱号": 9,
            "备注": 10
        }

        for excel_row, (df_idx, row) in enumerate(df.iterrows(), start=data_start_row):
            is_total_row = str(row.get("部位", "")).strip() == "合计"

            for col_name, excel_col in fixed_mapping.items():
                if col_name not in df.columns:
                    continue

                cell = ws.Cells(excel_row, excel_col)
                value = row[col_name]

                if col_name == "平方数" and not is_total_row:
                    formula = build_area_formula(excel_row, unit, pricing_mode)
                    cell.Formula = formula if formula else ""
                    continue

                if pd.isna(value):
                    value = ""

                cell.Value = value

        fixed_cols = list(fixed_mapping.keys()) + ["小面加工"]
        extra_cols = [col for col in df.columns if col not in fixed_cols]

        extra_start_col = 11  # K
        extra_end_col = extra_start_col + len(extra_cols) - 1 if extra_cols else 10
        last_data_row = data_start_row + len(df) - 1

        for row_num in range(header_top_row, last_data_row + 1):
            ws.Rows(row_num).RowHeight = OUTPUT_ROW_HEIGHT

        for excel_col, col_name in enumerate(extra_cols, start=extra_start_col):
            ws.Cells(header_excel_row, excel_col).Value = col_name

        for excel_col, col_name in enumerate(extra_cols, start=extra_start_col):
            for offset, value in enumerate(df[col_name], start=0):
                excel_row = data_start_row + offset
                df_idx = df.index[offset]
                row = df.iloc[offset]
                is_total_row = str(row.get("部位", "")).strip() == "合计"
                cell = ws.Cells(excel_row, excel_col)

                if col_name == "纹路":
                    cell.Value = "" if pd.isna(value) else value
                    continue

                if isinstance(col_name, str) and col_name.endswith("(米)") and not is_total_row:
                    records_for_cell = small_face_formula_map.get((df_idx, col_name), [])

                    formula_parts = []
                    for record in records_for_cell:
                        one_formula = build_small_face_formula_from_record(record, excel_row, unit)
                        if one_formula:
                            formula_parts.append(one_formula.lstrip("="))

                    if formula_parts:
                        cell.Formula = "=" + "+".join(formula_parts)
                    else:
                        cell.Value = "" if pd.isna(value) else value
                    continue

                if pd.isna(value):
                    value = ""
                cell.Value = value

        if extra_cols:
            format_range = ws.Range(
                ws.Cells(header_top_row, extra_start_col),
                ws.Cells(last_data_row, extra_end_col)
            )

            format_range.Font.Name = "黑体"
            format_range.HorizontalAlignment = -4108
            format_range.VerticalAlignment = -4108

            for col in range(extra_start_col, extra_end_col + 1):
                ws.Columns(col).ColumnWidth = 10

            for row_num in range(header_top_row, last_data_row + 1):
                ws.Rows(row_num).RowHeight = OUTPUT_ROW_HEIGHT

            header_range = ws.Range(
                ws.Cells(header_top_row, extra_start_col),
                ws.Cells(header_excel_row, extra_end_col)
            )
            header_range.Font.Bold = True

            if "纹路" in extra_cols:
                grain_col = extra_start_col + extra_cols.index("纹路")

                grain_header_range = ws.Range(
                    ws.Cells(header_top_row, grain_col),
                    ws.Cells(header_excel_row, grain_col)
                )

                grain_header_range.UnMerge()
                grain_header_range.Merge()
                ws.Cells(header_top_row, grain_col).Value = "纹路"

                grain_header_range.Font.Name = "黑体"
                grain_header_range.Font.Bold = True
                grain_header_range.HorizontalAlignment = -4108
                grain_header_range.VerticalAlignment = -4108

                red_color = rgb_to_bgr_int(255, 0, 0)
                green_color = rgb_to_bgr_int(0, 176, 80)

                for excel_row in range(data_start_row, last_data_row + 1):
                    cell = ws.Cells(excel_row, grain_col)
                    value = str(cell.Value).strip() if cell.Value is not None else ""

                    if value == "横":
                        cell.Font.Color = red_color
                    elif value == "竖":
                        cell.Font.Color = green_color
                    else:
                        cell.Font.ColorIndex = 1

        if str(ws.Cells(last_data_row, 1).Value).strip() == "合计":
            ws.Cells(last_data_row, 6).Formula = f"=SUM(F{data_start_row}:F{last_data_row - 1})"

            ws.Cells(last_data_row, 7).Formula = f"=SUM(G{data_start_row}:G{last_data_row - 1})"

            for excel_col, col_name in enumerate(extra_cols, start=extra_start_col):
                if col_name == "纹路":
                    continue

                if isinstance(col_name, str) and (col_name.endswith("(米)") or col_name.endswith("(个)")):
                    addr_text = ws.Cells(1, excel_col).GetAddress(False, False)
                    match = re.match(r"([A-Z]+)", addr_text)
                    if not match:
                        continue

                    col_letter = match.group(1)
                    ws.Cells(last_data_row, excel_col).Formula = (
                        f"=SUM({col_letter}{data_start_row}:{col_letter}{last_data_row - 1})"
                    )

        if str(ws.Cells(last_data_row, 1).Value).strip() == "合计":
            total_full_range = ws.Range(
                ws.Cells(last_data_row, 1),
                ws.Cells(last_data_row, extra_end_col if extra_cols else 10)
            )
            total_full_range.Font.Name = "黑体"
            total_full_range.Font.Bold = True
            total_full_range.HorizontalAlignment = -4108
            total_full_range.VerticalAlignment = -4108
            total_full_range.RowHeight = OUTPUT_ROW_HEIGHT
            
        full_end_col = 10

        full_range = ws.Range(
            ws.Cells(header_top_row, 1),
            ws.Cells(last_data_row, full_end_col)
        )

        full_range.Borders.LineStyle = 1
        full_range.Borders.Weight = 2
        
        if template_last_row > last_data_row:
            ws.Rows(f"{last_data_row + 1}:{template_last_row}").Delete()
            
        clear_end_col = extra_end_col if extra_cols else 10

        if template_last_row > last_data_row:
            ws.Range(
                ws.Cells(last_data_row + 1, 1),
                ws.Cells(template_last_row, clear_end_col)
            ).Clear()
        
        wb.Save()
        print(f"已输出生产单格式文件：{output_path} / {sheet_name}")
        return output_path

    finally:
        if wb is not None:
            wb.Close(SaveChanges=True)
        if excel is not None:
            excel.Quit()

# 主流程函数
def process_one_sheet(
    input_path,
    grain_axis,
    unit,
    pricing_mode,
    thickness,
    root,
    header_row,
    sheet_name,
    pricing_rules=None,
    override_thickness=False,
    recalc_piece_count=True
):
    df = read_excel_file(input_path, header_row, sheet_name=sheet_name)

    df = df[
        df["编号"].notna() &
        (df["编号"].astype(str).str.strip() != "")
    ].copy()

    if override_thickness:
        df["厚"] = normalize_thickness_input(thickness)
    else:
        df["厚"] = df["厚"].apply(normalize_thickness_input)

    print(f"\n===== [{sheet_name}] 长宽统一前 =====")
    print(df[["编号", "长", "宽", "厚", "片数", "平方数"]].head(10))

    if recalc_piece_count:
        df = calculate_piece_count(df)

    df = calculate_grain(df, grain_axis)
    df = normalize_length_width(df)

    print(f"\n===== [{sheet_name}] 长宽统一后 =====")
    print(df[["编号", "长", "宽", "厚", "片数", "平方数"]].head(10))

    df = calculate_area(df, unit, pricing_mode)
    df = calculate_small_face(df, input_path, unit, header_row, sheet_name, pricing_rules)
    df = add_total_row(df)
    return df

def ask_small_face_rules(root, all_types):
    dialog = tk.Toplevel(root)
    dialog.title("选择小面加工计价方式")
    dialog.geometry("520x320")
    dialog.grab_set()

    result = {}
    combo_vars = {}

    tk.Label(dialog, text="请为每种小面加工选择计价方式：", font=("Arial", 11, "bold")).pack(pady=10)

    main_frame = tk.Frame(dialog)
    main_frame.pack(fill="both", expand=True, padx=15, pady=10)

    for col_name in sorted(all_types):
        row_frame = tk.Frame(main_frame)
        row_frame.pack(fill="x", pady=5)

        tk.Label(row_frame, text=col_name, anchor="w", width=28).pack(side="left")

        var = tk.StringVar(value="米")
        combo = ttk.Combobox(
            row_frame,
            textvariable=var,
            values=["米", "个"],
            state="readonly",
            width=18
        )
        combo.pack(side="left", padx=8)

        combo_vars[col_name] = var

    def confirm():
        for col_name, var in combo_vars.items():
            result[col_name] = var.get()
        dialog.destroy()

    def cancel():
        result.clear()
        dialog.destroy()

    btn_frame = tk.Frame(dialog)
    btn_frame.pack(pady=12)

    tk.Button(btn_frame, text="确定", width=12, command=confirm).pack(side="left", padx=10)
    tk.Button(btn_frame, text="取消", width=12, command=cancel).pack(side="left", padx=10)

    dialog.wait_window()
    return result

def ask_multi_sheet_small_face_rules(root, sheet_types_map):
    dialog = tk.Toplevel(root)
    dialog.title("多 sheet 小面加工计价方式")
    dialog.geometry("760x500")
    dialog.grab_set()

    result = {}
    combo_vars = {}

    tk.Label(
        dialog,
        text="请分别为每个 sheet 的小面加工类型选择计价方式：",
        font=("Arial", 11, "bold")
    ).pack(pady=10)

    notebook = ttk.Notebook(dialog)
    notebook.pack(fill="both", expand=True, padx=10, pady=10)

    for sheet_name, all_types in sheet_types_map.items():
        tab = tk.Frame(notebook)
        notebook.add(tab, text=sheet_name)

        combo_vars[sheet_name] = {}

        if not all_types:
            tk.Label(tab, text="这个 sheet 没有检测到小面加工类型").pack(pady=20)
            continue

        for col_name in sorted(all_types):
            row_frame = tk.Frame(tab)
            row_frame.pack(fill="x", padx=15, pady=6)

            tk.Label(row_frame, text=col_name, anchor="w", width=30).pack(side="left")

            var = tk.StringVar(value="米")
            combo = ttk.Combobox(
                row_frame,
                textvariable=var,
                values=["米", "个"],
                state="readonly",
                width=12
            )
            combo.pack(side="left", padx=8)

            combo_vars[sheet_name][col_name] = var

    def confirm():
        for sheet_name, type_vars in combo_vars.items():
            result[sheet_name] = {}
            for col_name, var in type_vars.items():
                result[sheet_name][col_name] = var.get()
        dialog.destroy()

    def cancel():
        result.clear()
        dialog.destroy()

    btn_frame = tk.Frame(dialog)
    btn_frame.pack(pady=12)

    tk.Button(btn_frame, text="确定", width=12, command=confirm).pack(side="left", padx=10)
    tk.Button(btn_frame, text="取消", width=12, command=cancel).pack(side="left", padx=10)

    dialog.wait_window()
    return result

def copy_template_file(input_path, output_path):
    shutil.copy2(input_path, output_path)
    print(f"已复制原文件到：{output_path}")
    return output_path

# GUI

def start_gui():
    root = tk.Tk()
    root.title("Excel Transform 工具")
    root.geometry("1100x650")

    input_path = {"value": None}
    df_result = {"value": None}
    processed_sheet_params = {"value": None}
    show_all_var = tk.BooleanVar(value=False)
    grain_var = tk.StringVar(value="无纹路")
    unit_var = tk.StringVar(value="mm")
    pricing_mode_var = tk.StringVar(value="平方数")
    thickness_var = tk.StringVar(value="")
    override_thickness_var = tk.BooleanVar(value=False)
    header_var = tk.StringVar(value="1")
    recalc_piece_var = tk.BooleanVar(value=True)
    tree_widget = {"value": None}
    current_dir = os.path.dirname(os.path.abspath(__file__))
    project_root = os.path.dirname(current_dir)

    input_dir = os.path.join(project_root, "input")
    output_dir = os.path.join(project_root, "output")

    # 左侧
    left_frame = tk.Frame(root, width=250, bd=1, relief="solid")
    left_frame.pack(side="left", fill="y", padx=10, pady=10)

    # 右侧
    right_frame = tk.Frame(root, bd=1, relief="solid")
    right_frame.pack(side="right", expand=True, fill="both", padx=10, pady=10)

    def select_file():
        file_path = filedialog.askopenfilename(
            title="选择 Excel 文件",
            initialdir=input_dir,
            filetypes=[("Excel files", "*.xlsx *.xls")]
        )
        if file_path:
            input_path["value"] = file_path
            file_label.config(text=f"已选择文件：\n{file_path}")
            status_label.config(text="文件已选择，等待处理")

    def process_file():
        if not input_path["value"]:
            messagebox.showwarning("提示", "请先选择文件")
            return

        try:
            status_label.config(text="⏳ 正在转换，请稍候…")
            btn_process.config(state="disabled")
            root.update_idletasks()

            start_time = time.time()

            # 单
            if mode_var.get() == "single":
                try:
                    header_row = int(header_var.get().strip())
                    if header_row < 0:
                        raise ValueError
                except:
                    messagebox.showwarning(
                        "提示",
                        "header 必须输入 0 或正整数\n输入数据开始行-2"
                    )
                    status_label.config(text="请先输入正确的 header")
                    return
                
                sheet_names = get_sheet_names(input_path["value"])
                first_sheet_name = sheet_names[0]
                
                pricing_rules = {}

                df_tmp = read_excel_file(
                    input_path["value"],
                    header_row,
                    sheet_name=first_sheet_name
                )

                _, all_types = extract_small_face_records(
                    df_tmp,
                    input_path["value"],
                    header_row,
                    first_sheet_name
                )

                if all_types:
                    pricing_rules = ask_small_face_rules(root, all_types)
                    if not pricing_rules:
                        status_label.config(text="已取消转换")
                        return

    
                df = process_one_sheet(
                    input_path["value"],
                    grain_var.get(),
                    unit_var.get(),
                    pricing_mode_var.get(),
                    normalize_thickness_input(thickness_var.get()),
                    root,
                    header_row,
                    sheet_name = first_sheet_name,
                    pricing_rules=pricing_rules,
                    override_thickness=override_thickness_var.get(),
                    recalc_piece_count=recalc_piece_var.get()
                )

                end_time = time.time()
                duration = round(end_time - start_time, 3)

                df_result["value"] = df
                processed_sheet_params["value"] = {
                    first_sheet_name: {
                        "enabled": True,
                        "grain_axis": grain_var.get(),
                        "unit": unit_var.get(),
                        "header_row": header_row,
                        "pricing_mode": pricing_mode_var.get(),
                        "override_thickness": override_thickness_var.get(),
                        "thickness": (
                            normalize_thickness_input(thickness_var.get())
                            if override_thickness_var.get() else None
                        ),
                        "pricing_rules": pricing_rules,
                        "recalc_piece_count": recalc_piece_var.get()
                    }
                }
                status_label.config(text=f"✅ 单sheet转换完成，共 {len(df)} 行（含合计），用时 {duration} 秒")
                update_preview()

            # 多
            else:
                sheet_names = get_sheet_names(input_path["value"])
                sheet_params = open_multi_sheet_dialog(root, sheet_names)

                if not sheet_params:
                    status_label.config(text="已取消转换")
                    return

                sheet_types_map = {}

                for sheet_name, params in sheet_params.items():
                    if not params["enabled"]:
                        continue

                    df_tmp = read_excel_file(
                        input_path["value"],
                        params["header_row"],
                        sheet_name=sheet_name
                    )

                    _, all_types = extract_small_face_records(
                        df_tmp,
                        input_path["value"],
                        params["header_row"],
                        sheet_name=sheet_name
                    )

                    sheet_types_map[sheet_name] = all_types
                    
                pricing_rules_by_sheet = ask_multi_sheet_small_face_rules(root, sheet_types_map)
                
                if not pricing_rules_by_sheet:
                    status_label.config(text="已取消转换")
                    return

                results = {}

                for sheet_name, params in sheet_params.items():
                    if not params["enabled"]:
                        continue
                    
                    sheet_rules = pricing_rules_by_sheet.get(sheet_name, {})
                    params["pricing_rules"] = sheet_rules

                    df = process_one_sheet(
                        input_path["value"],
                        params["grain_axis"],
                        params["unit"],
                        params["pricing_mode"],
                        params["thickness"],
                        root,
                        params["header_row"],
                        sheet_name=sheet_name,
                        pricing_rules=sheet_rules,
                        override_thickness=params.get("override_thickness", False),
                        recalc_piece_count=params.get("recalc_piece_count", True)
                    )

                    results[sheet_name] = df

                df_result["value"] = results
                processed_sheet_params["value"] = sheet_params
                if len(results) == 0:
                    status_label.config(text="⚠️ 没有选中任何需要处理的sheet")
                    return
                
                status_label.config(text=f"✅ 多sheet转换完成，共处理 {len(results)} 个sheet")
                
                update_preview()

        except Exception as e:
            msg = str(e)

            if "已取消" in msg:
                status_label.config(text="已取消转换")
                messagebox.showinfo("提示", msg)
            else:
                status_label.config(text="转换失败")
                messagebox.showerror("错误", f"处理失败：\n{e}")

        finally:
            btn_process.config(state="normal")
            

    def update_preview():
    # 清空旧内容
        for widget in tree_frame.winfo_children():
            widget.destroy()

        data = df_result["value"]

        if data is None:
            empty_label = tk.Label(tree_frame, text="目前没有可预览的数据")
            empty_label.pack()
            return

        # 单
        if isinstance(data, pd.DataFrame):
            if show_all_var.get():
                preview_df = data
            else:
                preview_df = data.head(20)

            preview_df = preview_df.where(pd.notna(preview_df), '')
            columns = list(preview_df.columns)

            tree = ttk.Treeview(tree_frame, columns=columns, show='headings')
            tree_widget["value"] = tree

            for col in columns:
                tree.heading(col, text=col)

                if len(col) >= 15:
                    col_width = 150
                elif len(col) >= 8:
                    col_width = 100
                else:
                    col_width = 80

                tree.column(col, anchor='center', width=col_width, minwidth=60, stretch=False)

            for _, row in preview_df.iterrows():
                tree.insert('', 'end', values=list(row))

            scrollbar_y = tk.Scrollbar(tree_frame, orient="vertical", command=tree.yview)
            tree.configure(yscrollcommand=scrollbar_y.set)

            scrollbar_x = tk.Scrollbar(tree_frame, orient="horizontal", command=tree.xview)
            tree.configure(xscrollcommand=scrollbar_x.set)

            scrollbar_y.pack(side="right", fill="y")
            scrollbar_x.pack(side="bottom", fill="x")
            tree.pack(side="left", fill="both", expand=True)

        # 多
        elif isinstance(data, dict):
            notebook = ttk.Notebook(tree_frame)
            notebook.pack(fill="both", expand=True)

            for sheet_name, df in data.items():
                tab_frame = tk.Frame(notebook)
                notebook.add(tab_frame, text=sheet_name)

                if show_all_var.get():
                    preview_df = df
                else:
                    preview_df = df.head(20)

                preview_df = preview_df.where(pd.notna(preview_df), '')
                columns = list(preview_df.columns)

                tree = ttk.Treeview(tab_frame, columns=columns, show='headings')

                for col in columns:
                    tree.heading(col, text=col)

                    if len(col) >= 15:
                        col_width = 150
                    elif len(col) >= 8:
                        col_width = 100
                    else:
                        col_width = 80

                    tree.column(col, anchor='center', width=col_width, minwidth=60, stretch=False)

                for _, row in preview_df.iterrows():
                    tree.insert('', 'end', values=list(row))

                scrollbar_y = tk.Scrollbar(tab_frame, orient="vertical", command=tree.yview)
                tree.configure(yscrollcommand=scrollbar_y.set)

                scrollbar_x = tk.Scrollbar(tab_frame, orient="horizontal", command=tree.xview)
                tree.configure(xscrollcommand=scrollbar_x.set)

                scrollbar_y.pack(side="right", fill="y")
                scrollbar_x.pack(side="bottom", fill="x")
                tree.pack(side="left", fill="both", expand=True)

    def export_file():
        if df_result["value"] is None:
            messagebox.showwarning("提示", "请先处理数据")
            return

        default_path = get_default_output_path(input_path["value"])

        save_path = filedialog.asksaveasfilename(
            title="保存处理后的 Excel 文件",
            initialdir=output_dir,
            initialfile=os.path.basename(default_path),
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )

        if save_path:
            try:
                export_excel(df_result["value"], save_path)
                messagebox.showinfo("完成", f"文件已导出：\n{save_path}")
                status_label.config(text="导出完成")
            except Exception as e:
                messagebox.showerror("错误", f"导出失败：\n{e}")
                
    def export_original_format_file():
        if df_result["value"] is None:
            messagebox.showwarning("提示", "请先处理数据")
            return
        
        if processed_sheet_params["value"] is None:
            messagebox.showwarning("提示", "缺少处理参数，请先重新执行一次转换")
            return

        current_dir = os.path.dirname(os.path.abspath(__file__))
        project_root = os.path.dirname(current_dir)
        output_dir = os.path.join(project_root, "output")
        os.makedirs(output_dir, exist_ok=True)

        default_path = get_default_output_path(input_path["value"])

        save_path = filedialog.asksaveasfilename(
            title="保存生产单格式文件",
            initialdir=output_dir,
            initialfile=os.path.basename(default_path),
            defaultextension=".xlsx",
            filetypes=[("Excel files", "*.xlsx")]
        )

        if not save_path:
            return
        
        try:
            # 单sheet
            if isinstance(df_result["value"], pd.DataFrame):
                sheet_params = processed_sheet_params["value"]
                sheet_name = list(sheet_params.keys())[0]
                header_row = sheet_params[sheet_name]["header_row"]

                params = sheet_params[sheet_name]

                export_to_original_format(
                    df_result["value"],
                    input_path["value"],
                    save_path,
                    header_row,
                    sheet_name,
                    unit=params["unit"],
                    pricing_mode=params["pricing_mode"],
                    pricing_rules=params.get("pricing_rules", {})
                )
            
            # 多sheet
            elif isinstance(df_result["value"], dict):
                copy_template_file(input_path["value"], save_path)

                for sheet_name, df in df_result["value"].items():
                    params = processed_sheet_params["value"].get(sheet_name)

                    if not params or not params.get("enabled", False):
                        continue

                    header_row = params["header_row"]

                    export_to_original_format(
                        df,
                        input_path["value"],
                        save_path,
                        header_row,
                        sheet_name,
                        unit=params["unit"],
                        pricing_mode=params["pricing_mode"],
                        pricing_rules=params.get("pricing_rules", {}),
                        copy_file=False
                    )
                    
            else:
                raise ValueError("处理结果格式不支持")
            
            messagebox.showinfo("完成", f"生产单格式文件已导出：\n{save_path}")
            status_label.config(text="生产单格式导出完成")

        except Exception as e:
            messagebox.showerror("错误", f"导出失败：\n{e}")

    # ===== 左侧控件 =====
    title_label = tk.Label(left_frame, text="Excel Transform", font=("Arial", 14, "bold"))
    title_label.pack(pady=15)

    btn_select = tk.Button(left_frame, text="选择文件", width=20, command=select_file)
    btn_select.pack(pady=10)

    file_label = tk.Label(left_frame, text="尚未选择文件", wraplength=220, justify="left")
    file_label.pack(pady=10)

    btn_process = tk.Button(left_frame, text="开始转换", width=20, command=process_file)
    btn_process.pack(pady=10)

    btn_export = tk.Button(left_frame, text="导出文件", width=20, command=export_file)
    btn_export.pack(pady=10)
    
    btn_export_original = tk.Button(
        left_frame,
        text="导出生产单格式",
        width=20,
        command=export_original_format_file
    )
    btn_export_original.pack(pady=10)

    # ===== 右侧控件 =====
    preview_title = tk.Label(right_frame, text="状态信息", font=("Arial", 12, "bold"))
    preview_title.pack(pady=10)


    status_label = tk.Label(
    right_frame,
    text="请先选择文件",
    anchor="w",
    justify="left"
)
    status_label.pack(fill="x", padx=10, pady=10)

    params_frame = tk.Frame(right_frame)
    params_frame.pack(fill="x", padx=10, pady=10)

    mode_var = tk.StringVar(value="single")

    row1 = tk.Frame(params_frame)
    row1.pack(fill="x", pady=5)

    grain_block = tk.Frame(row1)
    grain_block.pack(side="left", padx=(0, 20))

    tk.Label(grain_block, text="纹路方向：").pack(anchor="w")
    grain_combo = ttk.Combobox(
        grain_block,
        textvariable=grain_var,
        values=["无纹路", "左右向", "上下向"],
        state="readonly",
        width=10
    )
    grain_combo.pack(anchor="w")

    unit_block = tk.Frame(row1)
    unit_block.pack(side="left", padx=(0, 20))

    tk.Label(unit_block, text="尺寸单位：").pack(anchor="w")
    unit_combo = ttk.Combobox(
        unit_block,
        textvariable=unit_var,
        values=["mm", "cm"],
        state="readonly",
        width=8
    )
    unit_combo.pack(anchor="w")

    pricing_block = tk.Frame(row1)
    pricing_block.pack(side="left", padx=(0, 20))

    tk.Label(pricing_block, text="计价方式：").pack(anchor="w")
    pricing_mode_combo = ttk.Combobox(
        pricing_block,
        textvariable=pricing_mode_var,
        values=["平方数", "米"],
        state="readonly",
        width=10
    )
    pricing_mode_combo.pack(anchor="w")


    row2 = tk.Frame(params_frame)
    row2.pack(fill="x", pady=5)

    header_block = tk.Frame(row2)
    header_block.pack(side="left", padx=(0, 20))

    tk.Label(header_block, text="表头所在行（header）：").pack(anchor="w")
    header_entry = tk.Entry(header_block, textvariable=header_var, width=12)
    header_entry.pack(anchor="w")

    header_hint = tk.Label(
        header_block,
        text="例：Excel第3行填 1；Excel第30行填 28",
        fg="gray"
    )
    header_hint.pack(anchor="w")

    thickness_block = tk.Frame(row2)
    thickness_block.pack(side="left", padx=(0, 20), anchor="n")

    override_thickness_check = tk.Checkbutton(
        thickness_block,
        text="变更厚度",
        variable=override_thickness_var,
        command=lambda: thickness_entry.config(
            state="normal" if override_thickness_var.get() and mode_var.get() == "single" else "disabled"
        )
    )
    override_thickness_check.pack(anchor="w")

    tk.Label(thickness_block, text="厚度：").pack(anchor="w")
    thickness_entry = tk.Entry(
        thickness_block,
        textvariable=thickness_var,
        width=12,
        state="disabled"
    )
    thickness_entry.pack(anchor="w")
    
    row2_5 = tk.Frame(params_frame)
    row2_5.pack(fill="x", pady=5)

    recalc_piece_check = tk.Checkbutton(
        row2_5,
        text="按编号重算片数",
        variable=recalc_piece_var
    )
    recalc_piece_check.pack(side="left")


    def update_mode_ui():
        if mode_var.get() == "single":
            grain_combo.config(state="readonly")
            unit_combo.config(state="readonly")
            header_entry.config(state="normal")
            pricing_mode_combo.config(state="readonly")
            override_thickness_check.config(state="normal")

            if override_thickness_var.get():
                thickness_entry.config(state="normal")
            else:
                thickness_entry.config(state="disabled")
        else:
            grain_combo.config(state="disabled")
            unit_combo.config(state="disabled")
            header_entry.config(state="disabled")
            pricing_mode_combo.config(state="disabled")
            override_thickness_check.config(state="disabled")
            thickness_entry.config(state="disabled")

    row3 = tk.Frame(params_frame)
    row3.pack(fill="x", pady=8)

    tk.Label(row3, text="处理模式：").pack(side="left", padx=(0, 10))

    tk.Radiobutton(
        row3,
        text="单 sheet",
        variable=mode_var,
        value="single",
        command=update_mode_ui
    ).pack(side="left", padx=(0, 10))

    tk.Radiobutton(
        row3,
        text="多 sheet",
        variable=mode_var,
        value="multi",
        command=update_mode_ui
    ).pack(side="left", padx=(0, 20))

    preview_check = tk.Checkbutton(
        row3,
        text="显示全部",
        variable=show_all_var,
        command=update_preview
    )
    preview_check.pack(side="left")

    update_mode_ui()

    tree_frame = tk.Frame(right_frame)
    tree_frame.pack(fill="both", expand=True, padx=10, pady=10)
    
    def refresh_tree_scroll(event=None):
        tree = tree_widget["value"]
        if tree is not None:
            try:
                tree.configure(displaycolumns=tree["columns"])
                tree.update_idletasks()
                tree.xview_moveto(tree.xview()[0])
                tree.yview_moveto(tree.yview()[0])
            except:
                pass

    root.bind("<Configure>", refresh_tree_scroll)

    root.mainloop()

if __name__ == "__main__":
    start_gui()