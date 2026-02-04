import os
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg


class DataAnalyzerApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Corporate Data Analyzer")
        self.root.geometry("1200x750")

        self.file_path = None
        self.df = None
        self.report_df = None

        self.figure = None
        self.chart_canvas = None

        self._build_ui()

    # ================= UI ================= #
    def _build_ui(self):

        # -------- File Selection -------- #
        file_frame = ttk.LabelFrame(self.root, text="File Selection")
        file_frame.pack(fill="x", padx=10, pady=5)

        self.file_label = ttk.Label(file_frame, text="No file selected")
        self.file_label.pack(side="left", padx=5)

        ttk.Button(file_frame, text="Browse",
                   command=self.browse_file).pack(side="left", padx=5)
        ttk.Button(file_frame, text="Read", command=self.read_file).pack(
            side="left", padx=5)

        # -------- Dataset Info -------- #
        info_frame = ttk.LabelFrame(self.root, text="Dataset Information")
        info_frame.pack(fill="x", padx=10, pady=5)

        self.info_text = tk.Text(info_frame, height=5, state="disabled")
        self.info_text.pack(fill="x", padx=5, pady=5)

        # -------- Report Builder -------- #
        report_frame = ttk.LabelFrame(self.root, text="Report Builder")
        report_frame.pack(fill="x", padx=10, pady=5)

        ttk.Label(report_frame, text="Group By Column").grid(
            row=0, column=0, padx=5)
        ttk.Label(report_frame, text="Aggregation").grid(
            row=0, column=1, padx=5)
        ttk.Label(report_frame, text="Value Column").grid(
            row=0, column=2, padx=5)

        self.group_cb = ttk.Combobox(report_frame, state="readonly")
        self.agg_cb = ttk.Combobox(
            report_frame,
            values=["sum", "mean", "max", "min", "count", "median"],
            state="readonly"
        )
        self.value_cb = ttk.Combobox(report_frame, state="readonly")

        self.group_cb.grid(row=1, column=0, padx=5)
        self.agg_cb.grid(row=1, column=1, padx=5)
        self.value_cb.grid(row=1, column=2, padx=5)

        ttk.Button(report_frame, text="Preview Report",
                   command=self.preview_report).grid(row=1, column=3, padx=10)
        ttk.Button(report_frame, text="Export Report",
                   command=self.export_report).grid(row=1, column=4, padx=5)

        # -------- Main Split: LEFT table | RIGHT chart -------- #
        main_frame = ttk.Frame(self.root)
        main_frame.pack(fill="both", expand=True, padx=10, pady=5)

        # -------- LEFT: Report Preview -------- #
        left_frame = ttk.LabelFrame(main_frame, text="Report Preview")
        left_frame.pack(side="left", fill="both", expand=True, padx=(0, 5))

        self.tree = ttk.Treeview(left_frame, show="headings")
        self.tree.pack(fill="both", expand=True)

        # -------- RIGHT: Chart Builder + Preview -------- #
        right_frame = ttk.LabelFrame(main_frame, text="Chart Builder")
        right_frame.pack(side="right", fill="both", expand=True, padx=(5, 0))

        ttk.Label(right_frame, text="Chart Type").pack(anchor="w", padx=5)
        self.chart_cb = ttk.Combobox(
            right_frame,
            values=["Bar", "Column", "Line", "Pie"],
            state="readonly"
        )
        self.chart_cb.current(0)
        self.chart_cb.pack(anchor="w", padx=5, pady=2)

        btn_frame = ttk.Frame(right_frame)
        btn_frame.pack(anchor="w", padx=5, pady=5)

        ttk.Button(btn_frame, text="Preview Chart",
                   command=self.preview_chart).pack(side="left", padx=5)
        ttk.Button(btn_frame, text="Export Chart",
                   command=self.export_chart).pack(side="left", padx=5)

        self.chart_preview = ttk.Frame(right_frame)
        self.chart_preview.pack(fill="both", expand=True, padx=5, pady=5)

    # ================= FILE ================= #
    def browse_file(self):
        self.file_path = filedialog.askopenfilename(
            filetypes=[("CSV Files", "*.csv"), ("Excel Files", "*.xlsx *.xls")]
        )
        if self.file_path:
            self.file_label.config(text=os.path.basename(self.file_path))

    def read_file(self):
        if not self.file_path:
            messagebox.showerror("Error", "Select a file first.")
            return

        try:
            if self.file_path.endswith(".csv"):
                self.df = pd.read_csv(self.file_path)
            else:
                self.df = pd.read_excel(self.file_path)
        except Exception as e:
            messagebox.showerror("Error", str(e))
            return

        self._clean_text_columns()
        self._show_dataset_info()
        self._populate_dropdowns()

    # ================= DATA CLEANING ================= #
    def _clean_text_columns(self):
        text_cols = self.df.select_dtypes(include=["object", "string"]).columns
        for col in text_cols:
            self.df[col] = (
                self.df[col]
                .astype(str)
                .str.strip()
                .str.title()
            )

    # ================= INFO ================= #
    def _show_dataset_info(self):
        info = (
            f"Rows: {self.df.shape[0]}\n"
            f"Columns: {self.df.shape[1]}\n"
            f"Column Names:\n{', '.join(self.df.columns)}"
        )
        self.info_text.config(state="normal")
        self.info_text.delete("1.0", tk.END)
        self.info_text.insert(tk.END, info)
        self.info_text.config(state="disabled")

    def _populate_dropdowns(self):
        self.group_cb["values"] = self.df.select_dtypes(
            include=["object", "string"]
        ).columns.tolist()

        allowed_values = [c for c in ["Quantity",
                                      "Unit_Price", "Discount"] if c in self.df.columns]
        self.value_cb["values"] = allowed_values

    # ================= REPORT ================= #
    def preview_report(self):
        if not all([self.group_cb.get(), self.agg_cb.get(), self.value_cb.get()]):
            messagebox.showerror("Error", "Select all report options.")
            return

        self.report_df = (
            self.df.groupby(self.group_cb.get())[self.value_cb.get()]
            .agg(self.agg_cb.get())
            .reset_index()
            .sort_values(by=self.value_cb.get(), ascending=False)
        )

        self.tree.delete(*self.tree.get_children())
        self.tree["columns"] = list(self.report_df.columns)

        for col in self.report_df.columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150)

        for _, row in self.report_df.iterrows():
            self.tree.insert("", tk.END, values=list(row))

        self.chart_cb.current(0)

    def export_report(self):
        if self.report_df is None:
            messagebox.showerror("Error", "No report to export.")
            return

        path = os.path.splitext(self.file_path)[0] + "_report.xlsx"
        self.report_df.to_excel(path, index=False)
        messagebox.showinfo("Success", f"Report exported:\n{path}")

    # # ================= CHART ================= #
    def preview_chart(self):
        if self.report_df is None:
            messagebox.showerror("Error", "Generate report first.")
            return

        chart_type = self.chart_cb.get()

        if self.figure:
            plt.close(self.figure)

        self.figure = plt.Figure(figsize=(8, 5), dpi=100)
        ax = self.figure.add_subplot(111)

        labels = self.report_df.iloc[:, 0]
        values = self.report_df.iloc[:, 1]

        if chart_type == "Bar":
            ax.barh(labels, values)
            ax.invert_yaxis()

        elif chart_type == "Column":
            x = range(len(labels))
            ax.bar(x, values)
            ax.set_xticks(x)
            ax.set_xticklabels(labels)

        elif chart_type == "Line":
            x = range(len(labels))
            ax.plot(x, values, marker="o")
            ax.set_xticks(x)
            ax.set_xticklabels(labels)

        elif chart_type == "Pie":
            if len(labels) > 8:
                messagebox.showwarning(
                    "Too Many Categories",
                    "Pie charts work best with 8 or fewer categories."
                )
                return

            ax.pie(values, labels=labels, autopct="%1.1f%%", startangle=90)

            ax.axis("equal")  # Makes pie a circle

        ax.set_title(f"{chart_type} Chart")
        self.figure.tight_layout()

        if self.chart_canvas:
            self.chart_canvas.get_tk_widget().destroy()

        self.chart_canvas = FigureCanvasTkAgg(
            self.figure, master=self.chart_preview)
        self.chart_canvas.draw()
        self.chart_canvas.get_tk_widget().pack(fill="both", expand=True)

    def export_chart(self):
        if not self.figure:
            messagebox.showerror("Error", "No chart to export.")
            return

        path = os.path.splitext(self.file_path)[0] + "_chart.png"
        self.figure.savefig(path)
        messagebox.showinfo("Success", f"Chart exported:\n{path}")


# ================= MAIN ================= #
if __name__ == "__main__":
    root = tk.Tk()
    app = DataAnalyzerApp(root)
    root.mainloop()
