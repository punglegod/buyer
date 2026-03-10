import openpyxl
import os
import re
from copy import copy
import tkinter as tk
from tkinter import messagebox

# ================= 配置区域 =================
TEMPLATE_FILE = "抛单.xlsx"
SUPPLIER_MAPPING = {
    "台州xx科技有限公司": "xx"
}


def get_supplier_short(full_name):
    if full_name in SUPPLIER_MAPPING:
        return SUPPLIER_MAPPING[full_name]
    short = re.sub(r'[（\(].*?[）\)]', '', full_name)
    short = re.sub(r'(成都|布艺|有限责任公司|有限公司|国际|实业|股份|科技|工业|贸易|中心|集团|布业|塑胶|LIMITED|LTD|CO\.|INC\.)',
                   '', short, flags=re.IGNORECASE)
    return short.strip()[:4]


# ================= 核心处理逻辑 =================
def process_order(input_text, log_func):
    orders_by_supplier = {}
    lines = input_text.strip().split('\n')

    # ====== 第一阶段：扫描并分组 ======
    for line in lines:
        line = re.sub(r'\s+', ' ', line).strip()
        if not line: continue

        dept, order_id, date, supplier = "", "", "", ""

        dept_match = re.search(r'(固定产品研发中心|功能品类发展部)', line)
        if dept_match: dept = dept_match.group(1)

        order_match = re.search(r'(CP[A-Z0-9]{10,})', line)
        if order_match: order_id = order_match.group(1)

        # 【核心修复】解决多个相同日期导致的定位错乱问题
        dates = re.findall(r'\d{4}[-/]\d{1,2}[-/]\d{1,2}', line)
        if dates:
            date = dates[-1]  # 取最后一个日期
            last_idx = line.rfind(date)  # 从右往左找这个日期的精确位置
            if last_idx != -1:
                # 截取该日期后面的所有字符作为供应商
                supplier = line[last_idx + len(date):].strip()

        item = {'mat_code': '', 'mat_desc': '', 'status': '', 'unit': '', 'qty': ''}
        mat_code_match = re.search(r'(\d{2}\.\d{2}\.[A-Z0-9\-]+)', line)
        if mat_code_match: item['mat_code'] = mat_code_match.group(1)

        sqq_match = re.search(r'([A-Z0-9]{2,3})\s*([\u4e00-\u9fa5]+)\s*(\d+)', line)
        if sqq_match:
            item['status'] = sqq_match.group(1)
            item['unit'] = sqq_match.group(2)
            item['qty'] = sqq_match.group(3)

        if item['mat_code'] and item['status']:
            desc_pattern = re.escape(item['mat_code']) + r'\s*(.*?)\s*' + re.escape(item['status'])
            desc_match = re.search(desc_pattern, line)
            if desc_match: item['mat_desc'] = desc_match.group(1).strip()

        if supplier and item['mat_code']:
            if supplier not in orders_by_supplier:
                orders_by_supplier[supplier] = {'dept': dept, 'order_id': order_id, 'date': date, 'items': []}
            orders_by_supplier[supplier]['items'].append(item)

    if not orders_by_supplier:
        log_func("❌ 错误：未能识别到有效订单，请检查粘贴的格式是否正确。")
        return False

    # ====== 第二阶段：生成 Excel ======
    for supplier, order_data in orders_by_supplier.items():
        log_func(f"\n📦 开始处理供应商: {supplier} (共 {len(order_data['items'])} 条物料)")

        try:
            wb = openpyxl.load_workbook(TEMPLATE_FILE)
            ws = wb.active
        except FileNotFoundError:
            log_func(f"❌ 致命错误：找不到模板文件 '{TEMPLATE_FILE}'")
            log_func("👉 请确保模板文件与本程序放在同一个文件夹内！")
            return False

        r_gui, r_zhao = None, None
        for r in range(4, 15):
            cell_val = str(ws.cell(row=r, column=10).value or "")
            if "桂小强" in cell_val: r_gui = r
            if "赵必保" in cell_val: r_zhao = r

        dept_val = order_data['dept']
        if "固定" in dept_val and r_gui:
            if r_zhao:
                ws.cell(row=r_gui, column=10).value = ws.cell(row=r_zhao, column=10).value
                ws.cell(row=r_zhao, column=10).value = ""
            else:
                ws.cell(row=r_gui, column=10).value = ""
        elif "功能" in dept_val and r_zhao:
            ws.cell(row=r_zhao, column=10).value = ""

        for i, item in enumerate(order_data['items']):
            current_row = 5 + i
            ws.cell(row=current_row, column=1).value = order_data['order_id']
            ws.cell(row=current_row, column=2).value = item['mat_code']
            ws.cell(row=current_row, column=3).value = item['mat_desc']
            ws.cell(row=current_row, column=4).value = item['status']
            ws.cell(row=current_row, column=5).value = item['unit']
            ws.cell(row=current_row, column=6).value = int(item['qty']) if item['qty'].isdigit() else item['qty']
            ws.cell(row=current_row, column=7).value = order_data['date']
            ws.cell(row=current_row, column=8).value = supplier

            if current_row > 5:
                if ws.row_dimensions[5].height is not None:
                    ws.row_dimensions[current_row].height = ws.row_dimensions[5].height
                for col in range(1, 9):
                    source_cell = ws.cell(row=5, column=col)
                    target_cell = ws.cell(row=current_row, column=col)
                    if source_cell.has_style:
                        target_cell.font = copy(source_cell.font)
                        target_cell.border = copy(source_cell.border)
                        target_cell.fill = copy(source_cell.fill)
                        target_cell.number_format = copy(source_cell.number_format)
                        target_cell.alignment = copy(source_cell.alignment)

            log_func(f"   ✅ 写入成功: {item['mat_code']} ({item['qty']}{item['unit']})")

        sup_short = get_supplier_short(supplier)
        new_filename = f"{order_data['order_id']}-{sup_short}-段腾飞.xlsx"
        try:
            wb.save(new_filename)
            log_func(f"🎉 文件已生成: {new_filename}")
        except PermissionError:
            log_func(f"❌ 错误：文件 {new_filename} 正在被其他程序占用！")
            log_func("👉 请先在 Excel 中关闭它，然后重试。")
            return False

    return True


# ================= GUI 界面构建 =================
class App:
    def __init__(self, root):
        self.root = root
        self.root.title("采购订单自动拆单神器")
        self.root.geometry("750x650")
        self.root.configure(bg="#f3f4f6")

        self.font_title = ("Microsoft YaHei", 16, "bold")
        self.font_main = ("Microsoft YaHei", 10)

        self.setup_ui()

    def setup_ui(self):
        title_frame = tk.Frame(self.root, bg="#3b82f6", pady=15)
        title_frame.pack(fill=tk.X)
        tk.Label(title_frame, text="📦 采购抛单 智能拆解生成器", font=self.font_title, bg="#3b82f6", fg="white").pack()

        main_frame = tk.Frame(self.root, bg="#f3f4f6", padx=20, pady=10)
        main_frame.pack(fill=tk.BOTH, expand=True)

        tk.Label(main_frame, text="请在此处粘贴订单文本（支持多行/多供应商混贴）：", font=self.font_main,
                 bg="#f3f4f6").pack(anchor=tk.W, pady=(10, 5))
        self.text_input = tk.Text(main_frame, height=10, font=("Consolas", 10), relief=tk.FLAT, bd=1)
        self.text_input.pack(fill=tk.X, pady=5)
        self.text_input.config(highlightbackground="#d1d5db", highlightcolor="#3b82f6", highlightthickness=1)

        btn_frame = tk.Frame(main_frame, bg="#f3f4f6")
        btn_frame.pack(fill=tk.X, pady=15)

        self.run_btn = tk.Button(btn_frame, text="⚡ 一键拆单生成", font=("Microsoft YaHei", 12, "bold"), bg="#10b981",
                                 fg="white",
                                 activebackground="#059669", activeforeground="white", relief=tk.FLAT, cursor="hand2",
                                 command=self.run_process, padx=20, pady=5)
        self.run_btn.pack(side=tk.LEFT)

        tk.Button(btn_frame, text="🗑️ 清空输入框", font=self.font_main, relief=tk.FLAT, cursor="hand2",
                  command=lambda: self.text_input.delete(1.0, tk.END)).pack(side=tk.LEFT, padx=15)

        tk.Button(btn_frame, text="📁 打开输出文件夹", font=self.font_main, relief=tk.FLAT, cursor="hand2", bg="#e5e7eb",
                  command=self.open_folder).pack(side=tk.RIGHT)

        tk.Label(main_frame, text="运行日志：", font=self.font_main, bg="#f3f4f6").pack(anchor=tk.W, pady=(10, 5))
        self.text_log = tk.Text(main_frame, height=12, font=("Consolas", 10), bg="#1f2937", fg="#10b981",
                                relief=tk.FLAT, padx=10, pady=10)
        self.text_log.pack(fill=tk.BOTH, expand=True)
        self.text_log.config(state=tk.DISABLED)

        self.status_bar = tk.Label(self.root, text=" 准备就绪 | 确保模板文件与本程序在同一目录下", bd=1,
                                   relief=tk.SUNKEN, anchor=tk.W, bg="#e5e7eb", font=("Microsoft YaHei", 9))
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    def log(self, message):
        self.text_log.config(state=tk.NORMAL)
        self.text_log.insert(tk.END, message + "\n")
        self.text_log.see(tk.END)
        self.text_log.config(state=tk.DISABLED)
        self.root.update()

    def open_folder(self):
        try:
            os.startfile(os.getcwd())
        except Exception as e:
            self.log(f"⚠️ 无法自动打开文件夹: {e}")

    def run_process(self):
        content = self.text_input.get(1.0, tk.END).strip()
        if not content:
            messagebox.showwarning("提示", "您还没有粘贴任何订单内容哦！")
            return

        self.text_log.config(state=tk.NORMAL)
        self.text_log.delete(1.0, tk.END)
        self.text_log.config(state=tk.DISABLED)

        self.log("⏳ 开始解析数据...\n" + "-" * 40)
        self.run_btn.config(state=tk.DISABLED, text="处理中...")
        self.root.update()

        success = process_order(content, self.log)

        self.log("-" * 40)
        if success:
            self.log("✨ 所有任务处理完毕！您可以点击右上角的【打开输出文件夹】查看文件。")
            self.status_bar.config(text=" ✅ 处理完成")
            messagebox.showinfo("成功", "拆单完成！生成的文件已保存在当前目录下。")
        else:
            self.log("⚠️ 处理过程中出现异常，请检查上述日志。")
            self.status_bar.config(text=" ⚠️ 处理失败")

        self.run_btn.config(state=tk.NORMAL, text="⚡ 一键拆单生成")


if __name__ == "__main__":
    root = tk.Tk()
    app = App(root)

    root.mainloop()
