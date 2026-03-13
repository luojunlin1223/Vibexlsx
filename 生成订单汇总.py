"""
从 Sheet A（订单明细）自动生成 Sheet B（按产品线×国家汇总）。
带完整 GUI 界面。
"""

import sys
import os
import threading
from collections import defaultdict
import openpyxl
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

# ========== 固定配置 ==========

COUNTRY_COL_MAP = {
    "India": 2,
    "Singapore": 3,
    "Vietnam": 4,
    "Philippines": 5,
    "Indonesia": 6,
    "Thailand": 7,
    "Malaysia": 8,
    "Japan": 9,
    "Korea": 10,
    "New Zealand": 11,
    "Australia": 12,
    "Other": 14,
}
CIS_COL = 13
TOTAL_COL = 15

CIS_COUNTRIES = {
    "Russia", "Kazakhstan", "Uzbekistan", "Turkmenistan", "Tajikistan",
    "Kyrgyzstan", "Belarus", "Ukraine", "Moldova", "Armenia",
    "Azerbaijan", "Georgia",
}

GROUP_HEADERS = {
    2: "Flexible Packaging",
    9: "Product Integrity&Material Test",
    17: "Food&Beverage",
}

PRODUCT_ROW_MAP = {
    "SI Permeation": 3,
    "Oxysense": 4,
    "TMI": 5,
    "Buchel": 6,
    "Technidyne": 7,
    "Rayran": 8,
    "TME": 10,
    "SI Process-Semiconductor": 11,
    "SI Process-others": 12,
    "SpecMetrix": 13,
    "UTS": 14,
    "TQC Sheen": 15,
    "C&W": 16,
    "CMC-KUHNKE": 18,
    "Quality By Vision": 19,
    "Eagle Vision": 20,
    "XTS": 21,
    "Steinfurth": 22,
    "Torus": 23,
    "SI Headspace": 24,
    "3rd party": 25,
    "Service": 26,
}

TOTAL_ROW = 27
DATA_ROW_START = 3
DATA_ROW_END = 26

NUM_FMT = '_ * #,##0.00_ ;_ * \\-#,##0.00_ ;_ * "-"??_ ;_ @_ '
NUM_FMT_TOTAL = "0.000000"

COL_HEADERS = [
    "India", "Singapore", "Vietnam", "Philippines", "Indonesia",
    "Thailand", "Malaysia", "Japan", "Korea", "New Zealand",
    "Australia", "CIS Countries", "Other", "APAC(Total)",
]

# 列号 → 国家名（用于日志显示）
COL_TO_COUNTRY = {v: k for k, v in COUNTRY_COL_MAP.items()}
COL_TO_COUNTRY[CIS_COL] = "CIS Countries"


def get_country_col(country_name):
    if not country_name:
        return None
    country_name = country_name.strip()
    if country_name in COUNTRY_COL_MAP:
        return COUNTRY_COL_MAP[country_name]
    if country_name in CIS_COUNTRIES:
        return CIS_COL
    return COUNTRY_COL_MAP["Other"]


def read_sheet_a(ws):
    agg = defaultdict(float)
    warnings = []
    row_count = 0

    for row in ws.iter_rows(min_row=2, max_row=ws.max_row, values_only=False):
        product = row[14].value   # O列
        country = row[11].value   # L列
        q5 = row[16].value        # Q列
        value = row[22].value     # W列

        if not product or value is None:
            continue

        # 跳过公式字符串等非数值
        if not isinstance(value, (int, float)):
            warnings.append(f"第{row[0].row}行 W列不是数值，已跳过")
            continue

        row_count += 1

        if q5 and str(q5).strip() == "Parts":
            product = "Service"
        else:
            product = str(product).strip()

        col = get_country_col(country)
        if col is None:
            continue

        if product not in PRODUCT_ROW_MAP:
            warnings.append(f"未知产品线 '{product}'，已跳过")
            continue

        agg[(product, col)] += value

    return agg, warnings, row_count


def write_sheet_b(wb, agg):
    if "Sheet B" in wb.sheetnames:
        del wb["Sheet B"]

    ws = wb.create_sheet("Sheet B")

    font_bold = Font(bold=True)
    font_normal = Font(bold=False)
    align_center = Alignment(horizontal="center")
    align_right = Alignment(horizontal="right")

    for i, header in enumerate(COL_HEADERS):
        cell = ws.cell(row=1, column=i + 2, value=header)
        cell.alignment = align_center
        if header == "APAC(Total)":
            cell.font = font_bold

    for row_num, title in GROUP_HEADERS.items():
        cell = ws.cell(row=row_num, column=1, value=title)
        cell.font = font_bold

    for product, row_num in PRODUCT_ROW_MAP.items():
        cell = ws.cell(row=row_num, column=1, value=product)
        if product in ("3rd party", "Service"):
            cell.font = font_bold
        else:
            cell.font = font_normal

    for (product, col), value in agg.items():
        row_num = PRODUCT_ROW_MAP[product]
        cell = ws.cell(row=row_num, column=col, value=value)
        cell.number_format = NUM_FMT
        cell.alignment = align_right

    b_letter = get_column_letter(2)
    n_letter = get_column_letter(14)
    for row_num in range(DATA_ROW_START, DATA_ROW_END + 1):
        if row_num in GROUP_HEADERS:
            continue
        formula = f"=SUM({b_letter}{row_num}:{n_letter}{row_num})"
        cell = ws.cell(row=row_num, column=TOTAL_COL, value=formula)
        if row_num == 25:
            cell.font = font_bold
            cell.number_format = NUM_FMT_TOTAL
        else:
            cell.number_format = NUM_FMT
        cell.alignment = align_right

    total_cell_a = ws.cell(row=TOTAL_ROW, column=1, value="Total")
    total_cell_a.font = font_bold
    for col in range(2, TOTAL_COL + 1):
        col_letter = get_column_letter(col)
        formula = f"=SUM({col_letter}{DATA_ROW_START}:{col_letter}{DATA_ROW_END})"
        cell = ws.cell(row=TOTAL_ROW, column=col, value=formula)
        cell.font = font_bold
        cell.number_format = NUM_FMT_TOTAL
        cell.alignment = align_right

    ws.column_dimensions["A"].width = 28
    for col in range(2, TOTAL_COL + 1):
        ws.column_dimensions[get_column_letter(col)].width = 14
    ws.column_dimensions[get_column_letter(TOTAL_COL)].width = 16


class App:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("订单汇总生成器")
        self.root.geometry("700x520")
        self.root.resizable(False, False)

        self.input_path = tk.StringVar()
        self.output_path = tk.StringVar()

        self._build_ui()

    def _build_ui(self):
        root = self.root

        # ---------- 标题 ----------
        title_label = tk.Label(root, text="本月已收到订单汇总", font=("Microsoft YaHei", 16, "bold"))
        title_label.pack(pady=(18, 4))

        subtitle = tk.Label(root, text="从 Sheet A 自动生成 Sheet B 汇总表", font=("Microsoft YaHei", 10), fg="#666666")
        subtitle.pack(pady=(0, 14))

        # ---------- 输入文件 ----------
        frame_in = tk.LabelFrame(root, text="输入文件", font=("Microsoft YaHei", 10), padx=12, pady=8)
        frame_in.pack(fill="x", padx=20, pady=(0, 8))

        entry_in = tk.Entry(frame_in, textvariable=self.input_path, font=("Microsoft YaHei", 9), state="readonly")
        entry_in.pack(side="left", fill="x", expand=True, ipady=3)

        btn_browse_in = tk.Button(frame_in, text="浏览...", font=("Microsoft YaHei", 9), command=self._browse_input, width=8)
        btn_browse_in.pack(side="right", padx=(8, 0))

        # ---------- 输出文件 ----------
        frame_out = tk.LabelFrame(root, text="输出文件", font=("Microsoft YaHei", 10), padx=12, pady=8)
        frame_out.pack(fill="x", padx=20, pady=(0, 8))

        entry_out = tk.Entry(frame_out, textvariable=self.output_path, font=("Microsoft YaHei", 9), state="readonly")
        entry_out.pack(side="left", fill="x", expand=True, ipady=3)

        btn_browse_out = tk.Button(frame_out, text="浏览...", font=("Microsoft YaHei", 9), command=self._browse_output, width=8)
        btn_browse_out.pack(side="right", padx=(8, 0))

        # ---------- 生成按钮 ----------
        self.btn_run = tk.Button(root, text="生 成", font=("Microsoft YaHei", 12, "bold"),
                                 bg="#4CAF50", fg="white", activebackground="#45a049",
                                 width=14, height=1, command=self._run)
        self.btn_run.pack(pady=12)

        # ---------- 日志区域 ----------
        frame_log = tk.LabelFrame(root, text="处理日志", font=("Microsoft YaHei", 10), padx=12, pady=8)
        frame_log.pack(fill="both", expand=True, padx=20, pady=(0, 16))

        self.log_text = tk.Text(frame_log, font=("Consolas", 9), height=10, state="disabled", wrap="word")
        scrollbar = tk.Scrollbar(frame_log, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scrollbar.set)
        scrollbar.pack(side="right", fill="y")
        self.log_text.pack(fill="both", expand=True)

    def _log(self, msg, tag=None):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", msg + "\n", tag)
        self.log_text.see("end")
        self.log_text.configure(state="disabled")
        self.root.update_idletasks()

    def _log_clear(self):
        self.log_text.configure(state="normal")
        self.log_text.delete("1.0", "end")
        self.log_text.configure(state="disabled")

    def _browse_input(self):
        path = filedialog.askopenfilename(
            title="选择订单汇总 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
        )
        if path:
            self.input_path.set(path)
            # 自动设置输出路径
            base, ext = os.path.splitext(path)
            self.output_path.set(f"{base}_output{ext}")

    def _browse_output(self):
        path = filedialog.asksaveasfilename(
            title="保存输出文件",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
        )
        if path:
            self.output_path.set(path)

    def _run(self):
        input_file = self.input_path.get()
        output_file = self.output_path.get()

        if not input_file:
            messagebox.showwarning("提示", "请先选择输入文件！")
            return
        if not output_file:
            messagebox.showwarning("提示", "请先设置输出文件路径！")
            return

        self._log_clear()
        self.btn_run.configure(state="disabled", text="处理中...")

        thread = threading.Thread(target=self._run_worker, args=(input_file, output_file), daemon=True)
        thread.start()

    def _schedule(self, func, *args):
        """在主线程中安全地执行 UI 操作。"""
        self.root.after(0, func, *args)

    def _reset_button(self):
        self.btn_run.configure(state="normal", text="生 成")

    def _run_worker(self, input_file, output_file):
        try:
            self._schedule(self._log, f"读取文件: {os.path.basename(input_file)}")

            wb = openpyxl.load_workbook(input_file, data_only=True)

            if "Sheet A" not in wb.sheetnames:
                self._schedule(self._log, "[错误] 找不到 'Sheet A' 工作表！")
                self._schedule(messagebox.showerror, "错误", "找不到 'Sheet A' 工作表！\n请确保输入文件包含名为 'Sheet A' 的工作表。")
                return

            ws_a = wb["Sheet A"]
            self._schedule(self._log, "正在解析 Sheet A 数据...")

            agg, warnings, row_count = read_sheet_a(ws_a)

            self._schedule(self._log, f"共读取 {row_count} 条有效数据行")
            self._schedule(self._log, f"汇总为 {len(agg)} 个产品×国家组合:")
            self._schedule(self._log, "")

            for (product, col), value in sorted(agg.items()):
                country_name = COL_TO_COUNTRY.get(col, f"列{col}")
                self._schedule(self._log, f"  {product:25s} | {country_name:15s} | {value:,.2f}")

            self._schedule(self._log, "")

            for w in set(warnings):
                self._schedule(self._log, f"[警告] {w}")

            self._schedule(self._log, "正在生成 Sheet B...")
            write_sheet_b(wb, agg)

            self._schedule(self._log, f"正在保存: {os.path.basename(output_file)}")
            wb.save(output_file)

            self._schedule(self._log, "")
            self._schedule(self._log, f"完成！输出文件: {output_file}")
            self._schedule(messagebox.showinfo, "完成", f"处理完成！\n\n输出文件：\n{output_file}")

        except Exception as e:
            self._schedule(self._log, f"\n[错误] {e}")
            self._schedule(messagebox.showerror, "错误", f"处理过程中出错：\n{e}")

        finally:
            self._schedule(self._reset_button)

    def run(self):
        self.root.mainloop()


def main():
    if len(sys.argv) >= 2:
        input_file = sys.argv[1]
        output_file = sys.argv[2] if len(sys.argv) >= 3 else None
        if output_file is None:
            base, ext = os.path.splitext(input_file)
            output_file = f"{base}_output{ext}"
        try:
            wb = openpyxl.load_workbook(input_file, data_only=True)
            if "Sheet A" not in wb.sheetnames:
                print("错误：找不到 'Sheet A' 工作表")
                sys.exit(1)
            agg, warnings, row_count = read_sheet_a(wb["Sheet A"])
            print(f"读取 {row_count} 条数据，汇总为 {len(agg)} 个组合")
            for w in set(warnings):
                print(f"[警告] {w}")
            write_sheet_b(wb, agg)
            wb.save(output_file)
            print(f"输出文件: {output_file}")
        except Exception as e:
            print(f"错误：{e}")
            sys.exit(1)
    else:
        App().run()


if __name__ == "__main__":
    main()
