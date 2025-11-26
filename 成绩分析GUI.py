#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
成绩分析系统
每个学科独立参数设置 | 一键生成报告
"""

import customtkinter as ctk
from tkinter import filedialog, messagebox
import threading
import os
from 成绩分析核心 import SubjectConfig, SubjectManager, GradeAnalysisCore

# 设置CustomTkinter主题
ctk.set_appearance_mode("dark")
ctk.set_default_color_theme("blue")


class ModernGradeAnalysisGUI:

    
    def __init__(self):
        self.root = ctk.CTk()
        self.root.title("成绩分析系统")
        self.root.geometry("1500x900")
        
        # 学科管理器
        self.subject_manager = SubjectManager()
        
        # 变量
        self.excel_file = ""
        self.output_file = "统计分析结果.docx"
        self.excel_output_file = "统计数据.xlsx"
        self.subject_widgets = {}  # 存储每个学科的控件
        
        # 创建界面
        self._create_ui()
        
    def _create_ui(self):
        """创建用户界面"""
        # 主容器
        self.root.grid_columnconfigure(1, weight=1)
        self.root.grid_rowconfigure(0, weight=1)
        
        # 左侧边栏
        self._create_sidebar()
        
        # 右侧主内容区
        self._create_main_content()
        
    def _create_sidebar(self):
        """创建左侧边栏"""
        sidebar = ctk.CTkFrame(self.root, width=320, corner_radius=0)
        sidebar.grid(row=0, column=0, rowspan=2, sticky="nsew")
        sidebar.grid_rowconfigure(6, weight=1)
        
        # 标题
        title = ctk.CTkLabel(
            sidebar,
            text="🎓 成绩分析系统",
            font=ctk.CTkFont(size=26, weight="bold")
        )
        title.grid(row=0, column=0, padx=20, pady=(30, 5))
        
        subtitle = ctk.CTkLabel(
            sidebar,
            text="v5.0",
            font=ctk.CTkFont(size=13),
            text_color=("gray70", "gray30")
        )
        subtitle.grid(row=1, column=0, padx=20, pady=(0, 30))
        
        # 文件选择区域
        file_frame = ctk.CTkFrame(sidebar, fg_color="transparent")
        file_frame.grid(row=2, column=0, padx=20, pady=10, sticky="ew")
        
        file_label = ctk.CTkLabel(
            file_frame,
            text="📁 Excel文件",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        file_label.pack(anchor="w", pady=(0, 8))
        
        self.file_entry = ctk.CTkEntry(
            file_frame,
            placeholder_text="选择成绩Excel文件...",
            height=40
        )
        self.file_entry.pack(fill="x", pady=(0, 8))
        
        browse_btn = ctk.CTkButton(
            file_frame,
            text="📂 浏览文件",
            command=self._browse_excel,
            height=40,
            fg_color="#4a9eff",
            hover_color="#3a8eef",
            font=ctk.CTkFont(size=13, weight="bold")
        )
        browse_btn.pack(fill="x", pady=(0, 8))
        
        detect_btn = ctk.CTkButton(
            file_frame,
            text="🔍 自动识别学科",
            command=self._auto_detect_subjects,
            height=40,
            fg_color="#34c759",
            hover_color="#24b749",
            font=ctk.CTkFont(size=13, weight="bold")
        )
        detect_btn.pack(fill="x")
        
        # 分隔线
        separator = ctk.CTkFrame(sidebar, height=2, fg_color=("gray80", "gray20"))
        separator.grid(row=3, column=0, padx=20, pady=20, sticky="ew")
        
        # 操作说明
        info_label = ctk.CTkLabel(
            sidebar,
            text="💡 使用说明",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        info_label.grid(row=4, column=0, padx=20, pady=(0, 10), sticky="w")
        
        info_text = ctk.CTkTextbox(sidebar, height=150, fg_color=("gray90", "gray10"))
        info_text.grid(row=5, column=0, padx=20, pady=(0, 10), sticky="ew")
        info_text.insert("1.0", 
            "1. 点击'浏览文件'选择Excel\n"
            "2. 点击'自动识别学科'\n"
            "3. 在右侧调整各学科参数\n"
            "   - 满分值\n"
            "   - 合格线百分比\n"
            "   - 优秀线百分比\n"
            "4. 点击'一键生成报告'\n"
            "5. 选择保存位置\n"
            "6. 等待生成完成"
        )
        info_text.configure(state="disabled")
        
        # 底部按钮区域
        self.export_btn = ctk.CTkButton(
            sidebar,
            text="📄 生成Word报告",
            command=self._generate_report,
            height=50,
            font=ctk.CTkFont(size=16, weight="bold"),
            fg_color="#4a9eff",
            hover_color="#3a8eef"
        )
        self.export_btn.grid(row=7, column=0, padx=20, pady=(20, 10), sticky="ew")

        # Excel导出按钮
        self.excel_btn = ctk.CTkButton(
            sidebar,
            text="📊 导出Excel数据",
            command=self._export_excel,
            height=50,
            font=ctk.CTkFont(size=16, weight="bold"),
            fg_color="#34c759",
            hover_color="#24b749"
        )
        self.excel_btn.grid(row=8, column=0, padx=20, pady=(0, 20), sticky="ew")

        # 进度条
        self.progress = ctk.CTkProgressBar(sidebar, height=8)
        self.progress.grid(row=9, column=0, padx=20, pady=(0, 10), sticky="ew")
        self.progress.set(0)

        self.status_label = ctk.CTkLabel(
            sidebar,
            text="✅ 就绪",
            font=ctk.CTkFont(size=12)
        )
        self.status_label.grid(row=10, column=0, padx=20, pady=(0, 20))

    def _create_main_content(self):
        """创建主内容区"""
        main_frame = ctk.CTkFrame(self.root, corner_radius=0, fg_color="transparent")
        main_frame.grid(row=0, column=1, sticky="nsew", padx=20, pady=20)
        main_frame.grid_columnconfigure(0, weight=1)
        main_frame.grid_rowconfigure(1, weight=1)

        # 顶部标题
        header_frame = ctk.CTkFrame(main_frame, fg_color="transparent")
        header_frame.grid(row=0, column=0, sticky="ew", pady=(0, 20))

        header = ctk.CTkLabel(
            header_frame,
            text="📊 学科参数设置",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        header.pack(side="left")

        # 添加学科按钮
        add_btn = ctk.CTkButton(
            header_frame,
            text="➕ 手动添加学科",
            command=self._add_subject_manually,
            height=35,
            fg_color="#34c759",
            hover_color="#24b749",
            font=ctk.CTkFont(size=13, weight="bold")
        )
        add_btn.pack(side="right", padx=10)

        # 学科列表容器（可滚动）
        self.subjects_container = ctk.CTkScrollableFrame(
            main_frame,
            label_text="",
            fg_color=("gray95", "gray10")
        )
        self.subjects_container.grid(row=1, column=0, sticky="nsew")
        self.subjects_container.grid_columnconfigure(0, weight=1)

        # 提示信息
        self.hint_label = ctk.CTkLabel(
            self.subjects_container,
            text="👈 请先选择Excel文件并点击'自动识别学科'\n或点击右上角'手动添加学科'",
            font=ctk.CTkFont(size=16),
            text_color=("gray60", "gray40")
        )
        self.hint_label.grid(row=0, column=0, pady=100)

    def _browse_excel(self):
        """浏览Excel文件"""
        filename = filedialog.askopenfilename(
            title="选择Excel文件",
            filetypes=[("Excel文件", "*.xlsx *.xls"), ("所有文件", "*.*")]
        )
        if filename:
            self.excel_file = filename
            self.file_entry.delete(0, "end")
            self.file_entry.insert(0, os.path.basename(filename))
            self.status_label.configure(text=f"✅ 已选择: {os.path.basename(filename)}")

    def _auto_detect_subjects(self):
        """自动识别学科"""
        if not self.excel_file:
            messagebox.showwarning("警告", "请先选择Excel文件！")
            return

        self.status_label.configure(text="🔍 正在识别学科...")
        self.progress.set(0.3)

        success, result = self.subject_manager.auto_detect_from_excel(self.excel_file)

        if success:
            self._refresh_subject_list()
            self.status_label.configure(text=f"✅ 成功识别 {len(result)} 个学科")
            self.progress.set(1.0)
            messagebox.showinfo("成功", f"成功识别 {len(result)} 个学科！\n\n请在右侧调整各学科参数。")
            self.progress.set(0)
        else:
            self.status_label.configure(text="❌ 识别失败")
            self.progress.set(0)
            messagebox.showerror("错误", f"识别失败: {result}")

    def _add_subject_manually(self):
        """手动添加学科"""
        dialog = ctk.CTkInputDialog(
            text="请输入学科名称:",
            title="添加学科"
        )
        subject_name = dialog.get_input()

        if not subject_name:
            return

        dialog2 = ctk.CTkInputDialog(
            text=f"请输入'{subject_name}'的满分:",
            title="设置满分"
        )
        max_score_str = dialog2.get_input()

        if not max_score_str:
            return

        try:
            max_score = float(max_score_str)
            config = SubjectConfig(subject_name, max_score)
            if self.subject_manager.add_subject(config):
                self._refresh_subject_list()
                messagebox.showinfo("成功", f"已添加学科: {subject_name}")
            else:
                messagebox.showwarning("警告", f"学科'{subject_name}'已存在！")
        except ValueError:
            messagebox.showerror("错误", "请输入有效的数字！")

    def _refresh_subject_list(self):
        """刷新学科列表"""
        # 清空容器
        for widget in self.subjects_container.winfo_children():
            widget.destroy()

        self.subject_widgets.clear()

        subjects = self.subject_manager.get_subjects()

        if not subjects:
            self.hint_label = ctk.CTkLabel(
                self.subjects_container,
                text="👈 请先选择Excel文件并点击'自动识别学科'\n或点击右上角'手动添加学科'",
                font=ctk.CTkFont(size=16),
                text_color=("gray60", "gray40")
            )
            self.hint_label.grid(row=0, column=0, pady=100)
            return

        # 为每个学科创建卡片
        for idx, subject in enumerate(subjects):
            self._create_subject_card(idx, subject)

    def _create_subject_card(self, idx, subject):
        """创建学科参数卡片"""
        # 卡片容器 - 玻璃拟态效果
        card = ctk.CTkFrame(
            self.subjects_container,
            corner_radius=15,
            fg_color=("white", "gray20"),
            border_width=1,
            border_color=("gray80", "gray30")
        )
        card.grid(row=idx, column=0, padx=15, pady=12, sticky="ew")
        card.grid_columnconfigure(1, weight=1)

        # 学科名称和删除按钮
        header_frame = ctk.CTkFrame(card, fg_color="transparent")
        header_frame.grid(row=0, column=0, columnspan=3, padx=25, pady=(20, 15), sticky="ew")

        name_label = ctk.CTkLabel(
            header_frame,
            text=f"📚 {subject.name}",
            font=ctk.CTkFont(size=18, weight="bold")
        )
        name_label.pack(side="left")

        delete_btn = ctk.CTkButton(
            header_frame,
            text="🗑️ 删除",
            width=80,
            height=28,
            command=lambda: self._delete_subject(subject.name),
            fg_color="#e74c3c",
            hover_color="#c0392b",
            font=ctk.CTkFont(size=12)
        )
        delete_btn.pack(side="right")

        # 满分设置
        max_score_label = ctk.CTkLabel(
            card,
            text="满分:",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        max_score_label.grid(row=1, column=0, padx=(25, 10), pady=8, sticky="w")

        max_score_entry = ctk.CTkEntry(card, width=100, height=35, font=ctk.CTkFont(size=14))
        max_score_entry.insert(0, str(subject.max_score))
        max_score_entry.grid(row=1, column=1, padx=10, pady=8, sticky="w")

        max_score_unit = ctk.CTkLabel(card, text="分", font=ctk.CTkFont(size=13))
        max_score_unit.grid(row=1, column=2, padx=(0, 25), pady=8, sticky="w")

        # 合格线百分比设置
        pass_frame = ctk.CTkFrame(card, fg_color="transparent")
        pass_frame.grid(row=2, column=0, columnspan=3, padx=25, pady=8, sticky="ew")
        pass_frame.grid_columnconfigure(1, weight=1)

        pass_label = ctk.CTkLabel(
            pass_frame,
            text="合格线百分比:",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        pass_label.grid(row=0, column=0, sticky="w")

        pass_value_label = ctk.CTkLabel(
            pass_frame,
            text=f"{subject.pass_percent}%",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color="#4a9eff"
        )
        pass_value_label.grid(row=0, column=2, padx=15, sticky="e")

        pass_slider = ctk.CTkSlider(
            pass_frame,
            from_=0,
            to=100,
            number_of_steps=100,
            height=20,
            command=lambda v, lbl=pass_value_label: lbl.configure(text=f"{int(v)}%")
        )
        pass_slider.set(subject.pass_percent)
        pass_slider.grid(row=0, column=1, padx=15, sticky="ew")

        # 优秀线百分比设置
        excel_frame = ctk.CTkFrame(card, fg_color="transparent")
        excel_frame.grid(row=3, column=0, columnspan=3, padx=25, pady=8, sticky="ew")
        excel_frame.grid_columnconfigure(1, weight=1)

        excel_label = ctk.CTkLabel(
            excel_frame,
            text="优秀线百分比:",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        excel_label.grid(row=0, column=0, sticky="w")

        excel_value_label = ctk.CTkLabel(
            excel_frame,
            text=f"{subject.excellence_percent}%",
            font=ctk.CTkFont(size=16, weight="bold"),
            text_color="#34c759"
        )
        excel_value_label.grid(row=0, column=2, padx=15, sticky="e")

        excel_slider = ctk.CTkSlider(
            excel_frame,
            from_=0,
            to=100,
            number_of_steps=100,
            height=20,
            command=lambda v, lbl=excel_value_label: lbl.configure(text=f"{int(v)}%")
        )
        excel_slider.set(subject.excellence_percent)
        excel_slider.grid(row=0, column=1, padx=15, sticky="ew")

        # 保存按钮
        save_btn = ctk.CTkButton(
            card,
            text="💾 保存设置",
            width=120,
            height=38,
            command=lambda: self._save_subject_config(
                subject.name,
                max_score_entry,
                pass_slider,
                excel_slider
            ),
            fg_color="#34c759",
            hover_color="#24b749",
            font=ctk.CTkFont(size=14, weight="bold")
        )
        save_btn.grid(row=4, column=0, columnspan=3, padx=25, pady=(15, 20))

        # 存储控件引用
        self.subject_widgets[subject.name] = {
            'max_score': max_score_entry,
            'pass_slider': pass_slider,
            'excel_slider': excel_slider
        }

    def _delete_subject(self, name):
        """删除学科"""
        if messagebox.askyesno("确认删除", f"确定要删除学科 {name} 吗？"):
            self.subject_manager.remove_subject(name)
            self._refresh_subject_list()
            messagebox.showinfo("成功", f"已删除学科: {name}")

    def _save_subject_config(self, name, max_entry, pass_slider, excel_slider):
        """保存学科配置"""
        try:
            max_score = float(max_entry.get())
            pass_percent = int(pass_slider.get())
            excellence_percent = int(excel_slider.get())

            new_config = SubjectConfig(name, max_score, pass_percent, excellence_percent)
            self.subject_manager.update_subject(name, new_config)

            messagebox.showinfo("成功", f"✅ {name} 配置已保存！\n\n满分: {max_score}\n合格线: {pass_percent}%\n优秀线: {excellence_percent}%")
        except ValueError:
            messagebox.showerror("错误", "请输入有效的数字！")

    def _generate_report(self):
        """生成报告"""
        if not self.excel_file:
            messagebox.showwarning("警告", "请先选择Excel文件！")
            return

        if not self.subject_manager.get_subjects():
            messagebox.showwarning("警告", "请先识别或添加学科！")
            return

        # 选择输出文件
        output_file = filedialog.asksaveasfilename(
            title="保存报告",
            defaultextension=".docx",
            filetypes=[("Word文档", "*.docx"), ("所有文件", "*.*")],
            initialfile="成绩统计分析.docx"
        )

        if not output_file:
            return

        self.output_file = output_file

        # 在后台线程中生成报告
        thread = threading.Thread(target=self._generate_report_thread)
        thread.daemon = True
        thread.start()

    def _generate_report_thread(self):
        """后台生成报告"""
        try:
            self.export_btn.configure(state="disabled", text="⏳ 生成中...")

            def update_progress(value, text):
                self.progress.set(value)
                self.status_label.configure(text=f"⏳ {text}")

            update_progress(0.1, "加载数据...")

            # 创建分析核心
            core = GradeAnalysisCore(self.excel_file, self.subject_manager)

            if not core.load_data():
                messagebox.showerror("错误", "加载数据失败！")
                return

            update_progress(0.3, "计算统计数据...")
            core.calculate_statistics()

            update_progress(0.6, "生成Word报告...")
            core.generate_word_report(self.output_file, update_progress)

            update_progress(1.0, "完成！")

            messagebox.showinfo("成功", f"✅ 报告已生成！\n\n保存位置:\n{self.output_file}")

        except Exception as e:
            messagebox.showerror("错误", f"生成报告失败：\n{str(e)}")
            import traceback
            traceback.print_exc()
        finally:
            self.export_btn.configure(state="normal", text="📄 生成Word报告")
            self.progress.set(0)
            self.status_label.configure(text="✅ 就绪")

    def _export_excel(self):
        """导出Excel数据"""
        if not self.excel_file:
            messagebox.showwarning("警告", "请先选择Excel文件！")
            return

        if not self.subject_manager.get_subjects():
            messagebox.showwarning("警告", "请先识别或添加学科！")
            return

        # 选择输出文件
        output_file = filedialog.asksaveasfilename(
            title="保存Excel数据",
            defaultextension=".xlsx",
            filetypes=[("Excel文件", "*.xlsx"), ("所有文件", "*.*")],
            initialfile="成绩统计数据.xlsx"
        )

        if not output_file:
            return

        self.excel_output_file = output_file

        # 在后台线程中导出
        thread = threading.Thread(target=self._export_excel_thread)
        thread.daemon = True
        thread.start()

    def _export_excel_thread(self):
        """后台导出Excel"""
        try:
            self.excel_btn.configure(state="disabled", text="⏳ 导出中...")
            self.status_label.configure(text="⏳ 加载数据...")
            self.progress.set(0.2)

            # 创建分析核心
            core = GradeAnalysisCore(self.excel_file, self.subject_manager)

            if not core.load_data():
                messagebox.showerror("错误", "加载数据失败！")
                return

            self.status_label.configure(text="⏳ 计算统计数据...")
            self.progress.set(0.5)
            core.calculate_statistics()

            self.status_label.configure(text="⏳ 导出Excel...")
            self.progress.set(0.8)
            core.export_to_excel(self.excel_output_file)

            self.progress.set(1.0)
            self.status_label.configure(text="✅ 完成！")

            messagebox.showinfo("成功", f"✅ Excel数据已导出！\n\n保存位置:\n{self.excel_output_file}\n\n包含内容:\n• 各科目统计\n• 总分统计\n• 原始数据\n• 整体分析\n• 科目分析\n• 学校分析\n• 改进建议")

        except Exception as e:
            messagebox.showerror("错误", f"导出Excel失败：\n{str(e)}")
            import traceback
            traceback.print_exc()
        finally:
            self.excel_btn.configure(state="normal", text="📊 导出Excel数据")
            self.progress.set(0)
            self.status_label.configure(text="✅ 就绪")

    def run(self):
        """运行应用"""
        self.root.mainloop()


if __name__ == "__main__":
    app = ModernGradeAnalysisGUI()
    app.run()

