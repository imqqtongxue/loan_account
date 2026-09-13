
import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import json
from datetime import datetime
from decimal import Decimal, InvalidOperation
import os
from openpyxl import Workbook

DATA_FILE = "loan_data.json"


class LoanManager:
    def __init__(self, root):
        self.root = root
        self.root.title("个人借款记账工具 (分借款人版)")
        self.root.geometry("1100x650")
        self.root.resizable(True, True)
        self.data = self.load_data()
        self.current_borrower = None

        # ==========顶部栏==========
        top_frame = tk.Frame(root)
        top_frame.pack(fill="x", padx=8, pady=6)

        self.btn_new_bor = tk.Button(top_frame, text="新建借款人", command=self.manage_borrower_popup, width=10)
        self.btn_new_bor.pack(side="left")

        tk.Label(top_frame, text="当前借款人：", font=("微软雅黑", 10)).pack(side="left", padx=(12, 3))
        self.borrower_cb = ttk.Combobox(top_frame, state="readonly", width=12)
        self.borrower_cb.pack(side="left")
        self.borrower_cb.bind("<<ComboboxSelected>>", self.on_select_borrower)

        tk.Label(top_frame, text="当前总欠款：", font=("微软雅黑", 10)).pack(side="left", padx=(20, 3))
        self.balance_label = tk.Label(top_frame, text="0.00", font=("微软雅黑", 11, "bold"), fg="red")
        self.balance_label.pack(side="left")

        self.btn_export = tk.Button(top_frame, text="导出流水XLSX", command=self.export_excel)
        self.btn_export.pack(side="left", padx=(20, 0))

        self.time_label = tk.Label(top_frame, text="系统时间：", font=("微软雅黑", 10))
        self.time_label.pack(side="right")

        # ==========录入区域==========
        input_frame = tk.Frame(root)
        input_frame.pack(fill="x", padx=8, pady=4)

        # 第一行录入控件
        tk.Label(input_frame, text="金额(正数)：").grid(row=0, column=0, sticky="w", pady=3)
        self.amount_entry = tk.Entry(input_frame, width=12)
        self.amount_entry.grid(row=0, column=1, padx=4)

        tk.Label(input_frame, text="收支类型：").grid(row=0, column=2, sticky="w", padx=(8, 0))
        self.type_cb = ttk.Combobox(input_frame, values=["借出(-)", "还款(+)"], state="readonly", width=11)
        self.type_cb.set("借出(-)")
        self.type_cb.grid(row=0, column=3, padx=4)

        tk.Label(input_frame, text="转账方式：").grid(row=0, column=4, sticky="w", padx=(8, 0))
        self.transfer_cb = ttk.Combobox(input_frame, values=["微信", "支付宝"], state="readonly", width=10)
        self.transfer_cb.set("微信")
        self.transfer_cb.grid(row=0, column=5, padx=4)

        # 第二行录入控件
        tk.Label(input_frame, text="借款事由：").grid(row=1, column=0, sticky="w", pady=3)
        self.reason_cb = ttk.Combobox(input_frame, values=["一般借钱", "代购物/消费", "a钱"], state="readonly", width=12)
        self.reason_cb.set("一般借钱")
        self.reason_cb.grid(row=1, column=1, padx=4)

        tk.Label(input_frame, text="补充信息：").grid(row=1, column=2, sticky="w", padx=(8, 0))
        self.desc_entry = tk.Entry(input_frame, width=24)
        self.desc_entry.grid(row=1, column=3, columnspan=2, padx=4)

        # 第三行时间选择
        tk.Label(input_frame, text="年份：").grid(row=2, column=0, sticky="w", pady=3)
        self.year_cb = ttk.Combobox(input_frame, values=[str(y) for y in range(2020, 2027)], state="readonly", width=6)
        self.year_cb.set("2026")
        self.year_cb.grid(row=2, column=1, padx=4)

        tk.Label(input_frame, text="月份：").grid(row=2, column=2, sticky="w", padx=(8, 0))
        self.month_cb = ttk.Combobox(input_frame, values=[str(m) for m in range(1, 13)], state="readonly", width=6)
        self.month_cb.set("9")
        self.month_cb.grid(row=2, column=3, padx=4)

        tk.Label(input_frame, text="日期：").grid(row=2, column=4, sticky="w", padx=(8, 0))
        self.day_entry = tk.Entry(input_frame, width=6)
        self.day_entry.insert(0, "13")
        self.day_entry.grid(row=2, column=5, padx=4)

        tk.Label(input_frame, text="小时：").grid(row=2, column=6, sticky="w", padx=(8, 0))
        self.hour_entry = tk.Entry(input_frame, width=6)
        self.hour_entry.insert(0, "12")
        self.hour_entry.grid(row=2, column=7, padx=4)

        # ==========功能按钮行==========
        btn_frame = tk.Frame(root)
        btn_frame.pack(pady=6)
        tk.Button(btn_frame, text="保存本条记录", command=self.save_record, width=12).pack(side="left", padx=10)
        tk.Button(btn_frame, text="删除选中记录", command=self.delete_selected_record, width=12).pack(side="left")

        # ==========流水表格==========
        cols = ("trade_time", "amount", "transfer", "reason", "input_time", "desc")
        self.tree = ttk.Treeview(root, columns=cols, show="headings")
        style = ttk.Style()
        style.configure("Treeview", anchor="center")

        self.tree.heading("trade_time", text="交易时间")
        self.tree.heading("amount", text="金额")
        self.tree.heading("transfer", text="转账方式")
        self.tree.heading("reason", text="事由类型")
        self.tree.heading("input_time", text="录入时间")
        self.tree.heading("desc", text="补充详情")

        self.tree.column("trade_time", width=160, anchor="center")
        self.tree.column("amount", width=70, anchor="center")
        self.tree.column("transfer", width=80, anchor="center")
        self.tree.column("reason", width=100, anchor="center")
        self.tree.column("input_time", width=160, anchor="center")
        self.tree.column("desc", width=220, anchor="center")

        scroll_bar = ttk.Scrollbar(root, orient="vertical", command=self.tree.yview)
        self.tree.configure(yscrollcommand=scroll_bar.set)
        self.tree.pack(side="left", fill="both", expand=True, padx=8)
        scroll_bar.pack(side="right", fill="y")

        self.refresh_borrower_combo()
        self.update_ui_timer()

    # 自动消失提示函数
    def show_tip(self, msg, duration=2000):
        tip_label = tk.Label(self.root, text=msg, bg="#222222", fg="white", font=("微软雅黑", 10))
        tip_label.place(relx=0.5, rely=0.4, anchor="center")
        self.root.after(duration, tip_label.destroy)

    # =====借款人弹窗=====
    def manage_borrower_popup(self):
        win = tk.Toplevel(self.root)
        win.title("管理借款人")
        win.geometry("300x210")
        win.transient(self.root)
        win.grab_set()
        tk.Label(win, text="新增借款人姓名：").pack(pady=5)
        name_input = tk.Entry(win, width=12)
        name_input.pack()

        def add_bor():
            name = name_input.get().strip()
            if not name:
                messagebox.showwarning("提示", "请填写姓名", parent=win)
                return
            if name in self.data:
                messagebox.showinfo("提示", "借款人已存在", parent=win)
                return
            self.data[name] = []
            self.save_json()
            self.refresh_borrower_combo()
            name_input.delete(0, tk.END)
            # [修复Bug7] 新增成功后自动关闭弹窗
            win.destroy()

        def del_bor():
            sel_name = self.borrower_cb.get()
            if not sel_name:
                messagebox.showwarning("提示", "先选中借款人", parent=win)
                return
            if messagebox.askyesno("确认删除", f"删除【{sel_name}】及全部流水，无法恢复？", parent=win):
                del self.data[sel_name]
                self.save_json()
                self.refresh_borrower_combo()
                # [修复Bug7] 删除成功后自动关闭弹窗
                win.destroy()

        btn_box = tk.Frame(win)
        btn_box.pack(pady=10)
        tk.Button(btn_box, text="新增", command=add_bor).pack(side="left", padx=6)
        tk.Button(btn_box, text="删除选中借款人", fg="red", command=del_bor).pack(side="left", padx=6)
        tk.Button(win, text="关闭", command=win.destroy).pack(pady=4)

    # [修复Bug6] JSON返回值类型未校验 + [修复Bug8] 空JSON文件导致数据丢失
    def load_data(self):
        if os.path.exists(DATA_FILE):
            try:
                file_size = os.path.getsize(DATA_FILE)
                if file_size == 0:
                    # [修复Bug8] 空文件不返回空字典吞掉数据，而是提示用户
                    messagebox.showwarning("警告", "数据文件为空，将创建空白数据文件")
                    return {}
                with open(DATA_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                # [修复Bug6] 校验JSON根节点是否为字典类型
                if not isinstance(data, dict):
                    messagebox.showwarning("警告", "数据文件格式异常(非字典结构)，将创建空白数据")
                    return {}
                return data
            except json.JSONDecodeError:
                messagebox.showwarning("警告", "数据文件JSON格式损坏，将创建空白数据文件")
                return {}
            except Exception as e:
                # [修复Bug5] 避免裸except捕获系统异常
                messagebox.showwarning("警告", f"读取数据文件异常: {e}，将创建空白数据")
                return {}
        return {}

    def save_json(self):
        try:
            with open(DATA_FILE, "w", encoding="utf-8") as f:
                json.dump(self.data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            # [修复Bug5] 避免裸except
            messagebox.showerror("错误", f"保存数据失败: {e}")

    def refresh_borrower_combo(self):
        names = list(self.data.keys())
        self.borrower_cb["values"] = names
        if names:
            self.borrower_cb.set(names[0])
            self.on_select_borrower(None)
        else:
            self.borrower_cb.set("")
            self.current_borrower = None
            self.refresh_table()
        self.calc_and_show_balance()

    def on_select_borrower(self, event):
        self.current_borrower = self.borrower_cb.get()
        self.refresh_table()
        self.calc_and_show_balance()

    def refresh_table(self):
        for row in self.tree.get_children():
            self.tree.delete(row)
        if not self.current_borrower:
            return
        records = self.data[self.current_borrower]
        # 过滤有效记录，同时保留原始索引
        valid_with_idx = [(i, r) for i, r in enumerate(records) if isinstance(r, dict) and r.get("trade_time")]
        valid_with_idx.sort(key=lambda x: x[1]["trade_time"])  # 按交易时间升序

        for orig_idx, item in valid_with_idx:
            self.tree.insert(
                "",
                tk.END,
                values=(
                    item["trade_time"],
                    item["amount"],
                    item["transfer_method"],
                    item["reason"],
                    item["input_time"],
                    item.get("desc", ""),
                ),
                # [修复Bug1] 将原始索引存入tags，删除时通过tags定位原始数据列表中的正确位置
                tags=(orig_idx,),
            )

    # [修复Bug10] 使用Decimal避免浮点精度累积误差
    def calc_and_show_balance(self):
        total = Decimal("0.00")
        if self.current_borrower:
            for rec in self.data[self.current_borrower]:
                try:
                    total += Decimal(str(rec["amount"]))
                except (InvalidOperation, TypeError, ValueError):
                    continue
        self.balance_label.config(text=f"{total:.2f}")

    def save_record(self):
        if not self.current_borrower:
            messagebox.showerror("错误", "请先选择借款人！")
            return
        try:
            year = self.year_cb.get().strip()
            month = self.month_cb.get().strip()
            day = self.day_entry.get().strip()
            hour = self.hour_entry.get().strip()
            # [修复Bug4] 日期时间输入无校验 — 用datetime验证合法性
            trade_time_str = f"{year}-{month.zfill(2)}-{day.zfill(2)} {hour.zfill(2)}:00:00"
            datetime.strptime(trade_time_str, "%Y-%m-%d %H:%M:%S")
        except ValueError:
            messagebox.showerror("错误", "日期时间格式不正确！请检查年/月/日/小时是否为有效数字。")
            return

        # [修复Bug2] 金额符号反转 — 校验金额必须为正数
        try:
            money = float(self.amount_entry.get().strip())
        except ValueError:
            messagebox.showerror("错误", "金额必须填写数字！")
            return

        if money <= 0:
            messagebox.showerror("错误", "金额必须为正数！")
            return

        type_text = self.type_cb.get()
        if type_text == "借出(-)":
            final_amount = -money
        else:
            final_amount = money
        record = {
            "trade_time": trade_time_str,
            "amount": f"{final_amount:.2f}",
            "transfer_method": self.transfer_cb.get(),
            "reason": self.reason_cb.get(),
            "desc": self.desc_entry.get().strip(),
            "input_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }
        self.data[self.current_borrower].append(record)
        self.save_json()
        self.refresh_table()
        self.calc_and_show_balance()
        # 清空输入框
        self.amount_entry.delete(0, tk.END)
        self.desc_entry.delete(0, tk.END)
        # 自动消失提示，2秒后消失
        self.show_tip("流水保存成功")

    # [修复Bug1] 删除记录索引错乱 — 通过tags中的原始索引定位，而非Treeview行号
    def delete_selected_record(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("提示", "表格选中一行流水！")
            return
        if not messagebox.askyesno("确认", "确定删除本条流水？"):
            return
        # 从tree item的tags中获取原始数据列表索引
        tags = self.tree.item(sel[0])["tags"]
        if tags:
            orig_idx = tags[0]
            records = self.data[self.current_borrower]
            # 双重校验：确保索引未越界且记录未被破坏
            if 0 <= orig_idx < len(records) and isinstance(records[orig_idx], dict):
                del records[orig_idx]
                self.save_json()
                self.refresh_table()
                self.calc_and_show_balance()
            else:
                messagebox.showerror("错误", "记录索引异常，请刷新后重试")
        else:
            messagebox.showerror("错误", "无法定位记录，请刷新后重试")

    # [修复Bug3] Excel导出金额为文本格式 + [修复Bug9] 导出文件路径不明确
    def export_excel(self):
        if not self.current_borrower or len(self.data[self.current_borrower]) == 0:
            messagebox.showwarning("提示", "没有流水数据可以导出！")
            return
        wb = Workbook()
        ws = wb.active
        ws.title = "流水记录"
        headers = ["交易时间", "金额", "转账方式", "事由类型", "录入时间", "补充详情"]
        ws.append(headers)
        for rec in self.data[self.current_borrower]:
            # [修复Bug3] 金额转为数值格式，便于用户在Excel中做SUM/AVERAGE计算
            amount_val = rec["amount"]
            try:
                amount_val = float(amount_val)
            except (ValueError, TypeError):
                pass  # 非数字则保留原字符串
            row_data = [
                rec["trade_time"],
                amount_val,
                rec["transfer_method"],
                rec["reason"],
                rec["input_time"],
                rec.get("desc", ""),
            ]
            ws.append(row_data)
        # [修复Bug9] 使用文件对话框让用户选择保存位置
        save_name = filedialog.asksaveasfilename(
            title="选择保存位置",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx")],
            initialfile=f"{self.current_borrower}_流水.xlsx",
        )
        if save_name:
            wb.save(save_name)
            messagebox.showinfo("导出成功", f"文件已保存：{save_name}")

    def update_ui_timer(self):
        now = datetime.now().strftime("系统时间：%Y-%m-%d %H:%M:%S")
        self.time_label.config(text=now)
        self.root.after(1000, self.update_ui_timer)


if __name__ == "__main__":
    window = tk.Tk()
    app = LoanManager(window)
    window.mainloop()