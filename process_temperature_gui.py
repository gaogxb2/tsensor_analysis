#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
基于 tkinter 的桌面 GUI，用于包装 process_temperature_data.py 的温度处理能力。

功能：
- 选择数据文件（默认 data/data1.txt）
- 选择模板文件（默认 template/template.xlsx）
- 选择输出目录（默认 result）
- 点击“开始生成”后执行处理，并在界面中实时显示日志
"""

import json
import os
import threading
import traceback
from pathlib import Path

import tkinter as tk
from tkinter import filedialog, messagebox

from openpyxl import Workbook

from process_temperature_data import (
    parse_data_file,
    read_template_mapping,
    calculate_average_temps,
    write_title,
    write_block_to_excel,
    apply_color_scale,
)


def run_pipeline(data_file: str, template_file: str, output_dir: str, logger=None) -> Path:
    """
    温度处理流水线，对外提供可复用接口。

    Args:
        data_file: 数据文件路径
        template_file: 模板 Excel 路径
        output_dir: 输出目录路径
        logger: 可选日志回调函数，形如 logger(str)

    Returns:
        生成的结果 Excel 文件路径 Path 对象
    """

    def log(msg: str):
        if logger is not None:
            logger(msg)
        else:
            print(msg)

    data_path = Path(data_file)
    template_path = Path(template_file)
    output_dir_path = Path(output_dir)

    output_dir_path.mkdir(parents=True, exist_ok=True)
    output_file = output_dir_path / "result.xlsx"

    log("开始处理温度数据...")

    # 1. 解析数据文件
    log("1. 解析数据文件...")
    blocks, block_titles = parse_data_file(str(data_path))
    log(f"   找到 {len(blocks)} 个测试块")

    # 2. 读取模板映射
    log("2. 读取模板映射...")
    mapping, max_template_row, max_template_col = read_template_mapping(str(template_path))
    log(f"   模板大小: {max_template_row} 行 x {max_template_col} 列")
    log(f"   找到 {len(mapping)} 个通道位置")

    # 3. 计算均值
    log("3. 计算平均温度...")
    avg_temps = calculate_average_temps(blocks)
    log(f"   计算了 {len(avg_temps)} 个通道的平均值")

    # 4. 创建新 Excel
    log("4. 生成 Excel 文件...")
    wb = Workbook()
    ws = wb.active
    ws.title = "result"

    # 收集所有温度值，用于条件格式
    all_temps = []

    # 4.1 先写入均值图标题和数据（最上面）
    log("   写入均值数据图...")
    # 写入均值图标题
    write_title(ws, "Average Temperature Map", row=1, max_col=max_template_col)

    # 写入均值图数据（从第 2 行开始）
    current_row = 2
    for (template_row, template_col), chnl in mapping.items():
        excel_row = current_row + template_row - 1
        excel_col = template_col
        if chnl in avg_temps:
            temp = avg_temps[chnl]
            ws.cell(row=excel_row, column=excel_col, value=temp)
            all_temps.append(temp)

    # 4.2 写入各个测试块（均值图下方，留空行）
    # 均值图结束行 = current_row + max_template_row - 1
    # 下一个位置 = 均值图结束行 + 2（留一个空行）
    current_row = current_row + max_template_row + 1

    for i, (block, title) in enumerate(zip(blocks, block_titles), 1):
        log(f"   写入测试块 {i} (标题: {title})...")

        # 写入块标题
        write_title(ws, f"Block {i} (#####{title}#####)", row=current_row, max_col=max_template_col)

        # 写入块数据（从标题下一行开始）
        data_start_row = current_row + 1
        block_end_row = write_block_to_excel(ws, block, mapping, data_start_row)

        # 收集温度值
        for temp in block.values():
            all_temps.append(temp)

        # 下一个块从当前块结束行 + 1 个空行开始
        current_row = block_end_row + 1

    # 5. 应用条件格式
    if all_temps:
        log("5. 应用条件格式...")
        min_temp = min(all_temps)
        max_temp = max(all_temps)
        log(f"   温度范围: {min_temp} ~ {max_temp}")
        apply_color_scale(ws, min_temp, max_temp)

    # 6. 保存文件
    log(f"6. 保存文件到: {output_file}")
    wb.save(str(output_file))

    log("处理完成！")
    return output_file


class TemperatureGUI(tk.Tk):
    """Tsensor 温度数据处理工具 GUI"""

    def __init__(self):
        super().__init__()
        self.title("Tsensor 温度数据处理工具")
        self.geometry("720x480")

        # 默认路径
        self.default_data_file = "data/data1.txt"
        self.default_template_file = "template/template.xlsx"
        self.default_output_dir = "result"

        # 配置文件路径（保存在用户主目录）
        self._config_file = Path.home() / ".tsensor_config.json"

        # 加载保存的路径配置
        config = self._load_config()
        
        # 变量（使用保存的路径或默认路径）
        self.data_file_var = tk.StringVar(value=config.get("data_file", self.default_data_file))
        self.template_file_var = tk.StringVar(value=config.get("template_file", self.default_template_file))
        self.output_dir_var = tk.StringVar(value=config.get("output_dir", self.default_output_dir))

        self._build_ui()

    # =========================
    # 界面构建
    # =========================
    def _build_ui(self):
        # 顶部标题
        title_label = tk.Label(
            self,
            text="Tsensor 温度数据处理工具",
            font=("Helvetica", 16, "bold"),
        )
        title_label.pack(pady=10)

        # 表单区域
        form_frame = tk.Frame(self)
        form_frame.pack(fill=tk.X, padx=20, pady=10)

        # 数据文件
        self._add_file_row(
            parent=form_frame,
            row=0,
            label_text="数据文件：",
            var=self.data_file_var,
            filetypes=[("文本文件", "*.txt"), ("所有文件", "*.*")],
            select_dir=False,
        )

        # 模板文件
        self._add_file_row(
            parent=form_frame,
            row=1,
            label_text="模板文件：",
            var=self.template_file_var,
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
            select_dir=False,
        )

        # 输出目录
        self._add_file_row(
            parent=form_frame,
            row=2,
            label_text="输出目录：",
            var=self.output_dir_var,
            filetypes=None,
            select_dir=True,
        )

        # 控制区域
        control_frame = tk.Frame(self)
        control_frame.pack(fill=tk.X, padx=20, pady=10)

        self.start_button = tk.Button(
            control_frame,
            text="开始生成",
            width=12,
            command=self.on_start_clicked,
        )
        self.start_button.pack(side=tk.LEFT)

        open_output_btn = tk.Button(
            control_frame,
            text="打开输出目录",
            width=12,
            command=self.open_output_dir,
        )
        open_output_btn.pack(side=tk.LEFT, padx=10)

        # 日志区域
        log_frame = tk.LabelFrame(self, text="处理日志", padx=5, pady=5)
        log_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(0, 20))

        self.log_text = tk.Text(
            log_frame,
            wrap=tk.WORD,
            height=12,
        )
        self.log_text.pack(fill=tk.BOTH, expand=True, side=tk.LEFT)

        scrollbar = tk.Scrollbar(log_frame, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.configure(yscrollcommand=scrollbar.set)

    def _add_file_row(
        self,
        parent: tk.Frame,
        row: int,
        label_text: str,
        var: tk.StringVar,
        filetypes,
        select_dir: bool = False,
    ):
        """共用的一行：标签 + 输入框 + 浏览按钮"""
        label = tk.Label(parent, text=label_text, anchor="e", width=10)
        label.grid(row=row, column=0, sticky="e", pady=5)

        entry = tk.Entry(parent, textvariable=var)
        entry.grid(row=row, column=1, sticky="we", padx=5, pady=5)

        parent.grid_columnconfigure(1, weight=1)

        def on_browse():
            initial_dir = None
            current = var.get().strip()
            if current:
                initial_dir = str(Path(current).expanduser().resolve().parent)

            if select_dir:
                selected = filedialog.askdirectory(
                    title="选择输出目录",
                    initialdir=initial_dir or ".",
                )
            else:
                selected = filedialog.askopenfilename(
                    title="选择文件",
                    initialdir=initial_dir or ".",
                    filetypes=filetypes or [("所有文件", "*.*")],
                )

            if selected:
                var.set(selected)
                # 保存配置
                self._save_config()

        button = tk.Button(parent, text="浏览...", command=on_browse, width=8)
        button.grid(row=row, column=2, padx=5, pady=5)

    # =========================
    # 配置管理
    # =========================
    def _load_config(self) -> dict:
        """从配置文件读取保存的路径"""
        if not self._config_file.exists():
            return {}
        
        try:
            with open(self._config_file, 'r', encoding='utf-8') as f:
                config = json.load(f)
                # 验证配置格式
                if isinstance(config, dict):
                    return config
        except (json.JSONDecodeError, IOError) as e:
            # 配置文件损坏或读取失败，返回空配置
            print(f"读取配置文件失败: {e}")
        
        return {}
    
    def _save_config(self):
        """保存当前路径到配置文件"""
        config = {
            "data_file": self.data_file_var.get().strip(),
            "template_file": self.template_file_var.get().strip(),
            "output_dir": self.output_dir_var.get().strip(),
        }
        
        try:
            with open(self._config_file, 'w', encoding='utf-8') as f:
                json.dump(config, f, indent=2, ensure_ascii=False)
        except IOError as e:
            # 保存失败不影响程序运行，只打印错误
            print(f"保存配置文件失败: {e}")

    # =========================
    # 事件与业务逻辑
    # =========================
    def append_log(self, msg: str):
        """在日志窗口追加一行文本（主线程调用）"""
        if not msg.endswith("\n"):
            msg += "\n"
        self.log_text.insert(tk.END, msg)
        self.log_text.see(tk.END)

    def thread_safe_log(self, msg: str):
        """从工作线程中安全地写日志"""
        self.after(0, self.append_log, msg)

    def on_start_clicked(self):
        """点击“开始生成”按钮"""
        data_file = self.data_file_var.get().strip()
        template_file = self.template_file_var.get().strip()
        output_dir = self.output_dir_var.get().strip()

        if not data_file:
            messagebox.showwarning("提示", "请先选择数据文件。")
            return
        if not template_file:
            messagebox.showwarning("提示", "请先选择模板文件。")
            return
        if not output_dir:
            messagebox.showwarning("提示", "请先选择输出目录。")
            return

        # 清空旧日志
        self.log_text.delete("1.0", tk.END)
        self.append_log("开始执行处理流水线...\n")

        # 禁用按钮，防止重复点击
        self.start_button.config(state=tk.DISABLED)

        thread = threading.Thread(
            target=self._run_pipeline_in_thread,
            args=(data_file, template_file, output_dir),
            daemon=True,
        )
        thread.start()

    def _run_pipeline_in_thread(self, data_file: str, template_file: str, output_dir: str):
        """后台线程中执行处理逻辑"""
        try:
            output_path = run_pipeline(
                data_file=data_file,
                template_file=template_file,
                output_dir=output_dir,
                logger=self.thread_safe_log,
            )
        except Exception as e:
            err_msg = f"处理过程中发生错误: {e}"
            tb = traceback.format_exc()
            self.thread_safe_log(err_msg)
            self.thread_safe_log(tb)
            self.after(
                0,
                lambda: messagebox.showerror("错误", err_msg),
            )
        else:
            success_msg = f"处理完成，结果已保存到：{output_path}"
            self.thread_safe_log(success_msg)
            self.after(
                0,
                lambda: messagebox.showinfo("完成", success_msg),
            )
        finally:
            # 恢复按钮可用
            self.after(0, lambda: self.start_button.config(state=tk.NORMAL))

    def open_output_dir(self):
        """打开输出目录"""
        path_str = self.output_dir_var.get().strip() or self.default_output_dir
        path = Path(path_str).expanduser()
        if not path.exists():
            messagebox.showwarning("提示", f"输出目录不存在：{path}")
            return

        try:
            # macOS 使用 open
            import subprocess

            subprocess.run(["open", str(path)], check=False)
        except Exception as e:
            messagebox.showerror("错误", f"无法打开输出目录：{e}")


def main():
    app = TemperatureGUI()
    app.mainloop()


if __name__ == "__main__":
    main()

