#!/usr/bin/env python3
"""
CSV转Excel工具 - 批量转换CSV为Excel
"""
import sys, os, tkinter as tk
from pathlib import Path
from tkinter import filedialog, messagebox
import tkinter as tk

try:
    import pandas as pd
    HAS_PANDAS = True
except ImportError:
    HAS_PANDAS = False

class App:
    def __init__(self, root):
        self.root = root
        root.title("CSV转Excel工具 v1.0")
        root.geometry("550x400")
        self.files = []
        self.build_ui()
    
    def build_ui(self):
        f = tk.Frame(self.root, bg="#1976d2", height=50)
        f.pack(fill="x")
        tk.Label(f, text="📄 CSV → Excel", font=("Arial",14,"bold"),
                 fg="white", bg="#1976d2").pack(pady=12)
        
        main = tk.Frame(self.root, padx=15, pady=10)
        main.pack(fill="both", expand=True)
        
        bf = tk.Frame(main)
        bf.pack(fill="x", pady=5)
        tk.Button(bf, text="添加CSV文件", command=self.add_files,
                  bg="#1976d2", fg="white", padx=12).pack(side="left", padx=5)
        tk.Button(bf, text="清空", command=self.clear,
                  padx=12).pack(side="left", padx=5)
        
        self.lb = tk.Listbox(main, font=("Consolas",10), bg="#e3f2fd", height=8)
        self.lb.pack(fill="both", expand=True, pady=10)
        
        tk.Button(main, text="开始转换", command=self.convert,
                  bg="#4caf50", fg="white", font=("Arial",10,"bold"),
                  padx=20).pack(pady=10)
        
        self.status = tk.Label(main, text="请添加CSV文件",
                               font=("Arial",10), fg="gray")
        self.status.pack()
    
    def add_files(self):
        fs = filedialog.askopenfilenames(title="选择CSV文件",
             filetypes=[("CSV","*.csv *.tsv")])
        for f in fs:
            if f not in self.files:
                self.files.append(f)
                self.lb.insert("end", Path(f).name)
        self.status.config(text=f"已添加 {len(self.files)} 个文件")
    
    def clear(self):
        self.files.clear()
        self.lb.delete(0, "end")
    
    def convert(self):
        if not self.files:
            messagebox.showwarning("提示", "请先添加CSV文件")
            return
        if not HAS_PANDAS:
            messagebox.showerror("缺少依赖", "请运行：pip install pandas openpyxl")
            return
        
        out_dir = filedialog.askdirectory(title="选择输出目录")
        if not out_dir: return
        
        ok = 0
        for f in self.files:
            try:
                df = pd.read_csv(f)
                out = str(Path(out_dir) / (Path(f).stem + ".xlsx"))
                df.to_excel(out, index=False)
                ok += 1
            except Exception as e:
                print(f"转换失败 {f}: {e}")
        
        messagebox.showinfo("完成", f"成功转换 {ok}/{len(self.files)} 个文件")
        self.status.config(text=f"✅ 完成：{ok} 个文件")

if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
