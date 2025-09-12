# general_plotter.py
# General-purpose plotting GUI for CSV/Excel data (Tkinter + Matplotlib + pandas)
# - Clear plot types (Single, Single + Right Y, Subplot Grid, Subplot Grid + Right Y, Facet Grid, Facet Grid + Right Y)
# - Consolidated Plot tab with labeled sections
# - Contextual controls (right-axis & facet appear only when relevant)
# - Live preview (debounced), export PNG/SVG, settings persistence

import os, json, sys
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from matplotlib.ticker import MultipleLocator, AutoMinorLocator, AutoLocator
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg

# ---------- Global styling ----------
plt.rcParams["font.family"] = "serif"
plt.rcParams["font.serif"] = ["Times New Roman"] + plt.rcParams["font.serif"]

SUBPLOT_TITLE_FONTSIZE = 14
LABEL_FONTSIZE = 12
SUPTITLE_FONTSIZE = 18
SUPTITLE_Y = 0.975
TICK_LABELSIZE = 11

LINESTYLE_DEFAULT = "-"
MARKER_DEFAULT = "."
LINEWIDTH_DEFAULT = 1.0
MARKERSIZE_DEFAULT = 2.0

SETTINGS_FILE = "general_plotter_settings.json"


def _safe_float(v, default=None):
    try:
        return float(v)
    except Exception:
        return default


class GeneralPlotter(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("General Plotter")
        self.minsize(1200, 840)
        self.geometry(self._load_window_geom() or "1260x920+50+50")

        # State
        self.df: pd.DataFrame | None = None
        self.file_path = ""
        self.sheet_names: list[str] = []
        self.current_sheet = tk.StringVar()
        self.columns: list[str] = []

        # Persistent settings
        self.settings = self._load_settings()

        # ---- Tk Vars (seed from settings) ----
        self.title_text = tk.StringVar(value=self.settings.get("title_text", ""))
        self.suptitle_text = tk.StringVar(value=self.settings.get("suptitle_text", ""))

        # New explicit plot types (back-compat mapping handled below)
        default_plot_type = self.settings.get("plot_type") or self._map_old_mode(
            self.settings.get("plot_mode")
        )
        self.plot_type = tk.StringVar(value=default_plot_type or "single")

        self.legend_loc = tk.StringVar(
            value=self.settings.get("legend_loc", "lower center")
        )
        self.plot_kind = tk.StringVar(
            value=self.settings.get("plot_kind", "line")
        )  # line | scatter

        # Axis ranges & ticks
        self.auto_x = tk.BooleanVar(value=self.settings.get("auto_x", True))
        self.auto_y_left = tk.BooleanVar(value=self.settings.get("auto_y_left", True))
        self.auto_y_right = tk.BooleanVar(value=self.settings.get("auto_y_right", True))

        self.x_min = tk.StringVar(value=str(self.settings.get("x_min", "")))
        self.x_max = tk.StringVar(value=str(self.settings.get("x_max", "")))
        self.y_left_min = tk.StringVar(value=str(self.settings.get("y_left_min", "")))
        self.y_left_max = tk.StringVar(value=str(self.settings.get("y_left_max", "")))
        self.y_right_min = tk.StringVar(value=str(self.settings.get("y_right_min", "")))
        self.y_right_max = tk.StringVar(value=str(self.settings.get("y_right_max", "")))

        self.auto_time_ticks = tk.BooleanVar(
            value=self.settings.get("auto_time_ticks", True)
        )
        self.auto_y_ticks = tk.BooleanVar(value=self.settings.get("auto_y_ticks", True))
        self.xmaj = tk.StringVar(value=str(self.settings.get("x_major_tick", "")))
        self.xminr = tk.StringVar(value=str(self.settings.get("x_minor_tick", "")))
        self.ymaj = tk.StringVar(value=str(self.settings.get("y_major_tick", "")))
        self.yminr = tk.StringVar(value=str(self.settings.get("y_minor_tick", "")))

        # Style per-all
        self.linewidth = tk.DoubleVar(
            value=self.settings.get("linewidth", LINEWIDTH_DEFAULT)
        )
        self.markersize = tk.DoubleVar(
            value=self.settings.get("markersize", MARKERSIZE_DEFAULT)
        )

        # Grids
        self.ncols = tk.IntVar(value=int(self.settings.get("ncols", 2)))

        # Column choices
        self.right_axis_series = tk.StringVar(
            value=self.settings.get("right_axis_series", "None")
        )
        self.facet_col_saved = self.settings.get("facet_col", "None")

        # Live preview debounce
        self._preview_after_id = None
        self._preview_debounce_ms = 250

        # UI
        self._build_ui()

        # Reload last-used file if present
        last_path = self.settings.get("last_file_path", "")
        if last_path and os.path.exists(last_path):
            self._load_file(last_path, auto=True)

        # Save settings on window close
        self.protocol("WM_DELETE_WINDOW", self._on_close)

    # ---------- Back-compat ----------
    @staticmethod
    def _map_old_mode(old_mode: str | None) -> str | None:
        # Old “plot_mode” values: single | by_y | facet  -> map to new explicit types
        if not old_mode:
            return None
        return {"single": "single", "by_y": "grid", "facet": "facet"}.get(
            old_mode, "single"
        )

    # ---------- Persistence ----------
    def _load_settings(self):
        try:
            if os.path.exists(SETTINGS_FILE):
                with open(SETTINGS_FILE, "r") as f:
                    return json.load(f)
        except Exception:
            pass
        return {}

    def _save_settings(self):
        payload = dict(self.settings)
        payload.update(
            {
                "title_text": self.title_text.get(),
                "suptitle_text": self.suptitle_text.get(),
                "plot_type": self.plot_type.get(),
                "legend_loc": self.legend_loc.get(),
                "plot_kind": self.plot_kind.get(),
                "auto_x": self.auto_x.get(),
                "auto_y_left": self.auto_y_left.get(),
                "auto_y_right": self.auto_y_right.get(),
                "x_min": self.x_min.get(),
                "x_max": self.x_max.get(),
                "y_left_min": self.y_left_min.get(),
                "y_left_max": self.y_left_max.get(),
                "y_right_min": self.y_right_min.get(),
                "y_right_max": self.y_right_max.get(),
                "auto_time_ticks": self.auto_time_ticks.get(),
                "auto_y_ticks": self.auto_y_ticks.get(),
                "x_major_tick": self.xmaj.get(),
                "x_minor_tick": self.xminr.get(),
                "y_major_tick": self.ymaj.get(),
                "y_minor_tick": self.yminr.get(),
                "linewidth": self.linewidth.get(),
                "markersize": self.markersize.get(),
                "ncols": self.ncols.get(),
                "x_col": self.cb_x.get() if hasattr(self, "cb_x") else "",
                "facet_col": (
                    self.cb_facet.get() if hasattr(self, "cb_facet") else "None"
                ),
                "y_cols": self._get_selected_y(),
                "right_axis_series": self.right_axis_series.get(),
            }
        )
        if self.file_path:
            payload["last_file_path"] = self.file_path
        try:
            payload["last_sheet_name"] = self.current_sheet.get()
        except Exception:
            pass
        try:
            with open(SETTINGS_FILE, "w") as f:
                json.dump(payload, f, indent=2)
        except Exception:
            pass
        self.settings = payload

    def _load_window_geom(self):
        try:
            g = self.settings.get("window_geometry", "")
            return g if g else None
        except Exception:
            return None

    def _remember_geom(self, *_):
        if self.state() == "normal":
            self.settings["window_geometry"] = self.geometry()

    def _on_close(self):
        try:
            self._save_settings()
        finally:
            try:
                plt.close("all")
            except Exception:
                pass
            self.destroy()

    # ---------- UI ----------
    def _build_ui(self):
        self.bind("<Configure>", self._remember_geom)

        nb = ttk.Notebook(self)
        nb.pack(fill="both", expand=True, padx=10, pady=10)

        self.tab_data = ttk.Frame(nb)
        self.tab_plot = ttk.Frame(nb)

        nb.add(self.tab_data, text="Data")
        nb.add(self.tab_plot, text="Plot")

        self._build_tab_data()
        self._build_tab_plot()

        # Bottom buttons
        bottom = ttk.Frame(self)
        bottom.pack(fill="x", padx=10, pady=(0, 10))
        ttk.Button(
            bottom, text="Render Plot(s) to New Window", command=self.render_plots
        ).pack(side="left")
        ttk.Button(bottom, text="Export PNG/SVG", command=self.export_plots).pack(
            side="left", padx=8
        )
        ttk.Button(bottom, text="Save Settings", command=self._save_settings).pack(
            side="left", padx=8
        )
        ttk.Button(
            bottom, text="Close All Figures", command=lambda: plt.close("all")
        ).pack(side="left", padx=8)

    def _build_tab_data(self):
        f = self.tab_data
        for i in range(4):
            f.grid_columnconfigure(i, weight=1)

        ttk.Label(f, text="Data file (CSV or Excel):").grid(
            row=0, column=0, sticky="w", padx=6, pady=6
        )
        self.e_path = ttk.Entry(f)
        self.e_path.grid(row=0, column=1, columnspan=2, sticky="ew", padx=6, pady=6)
        ttk.Button(f, text="Browse…", command=self._browse_file).grid(
            row=0, column=3, sticky="e", padx=6, pady=6
        )

        ttk.Button(
            f,
            text="Load / Refresh",
            command=lambda: self._load_file(self.e_path.get().strip()),
        ).grid(row=1, column=3, sticky="e", padx=6, pady=6)

        ttk.Label(f, text="Sheet (Excel only):").grid(
            row=1, column=0, sticky="w", padx=6, pady=6
        )
        self.cb_sheet = ttk.Combobox(
            f,
            values=self.sheet_names,
            textvariable=self.current_sheet,
            state="readonly",
        )
        self.cb_sheet.grid(row=1, column=1, sticky="ew", padx=6, pady=6)
        ttk.Button(f, text="Load Sheet", command=self._load_sheet).grid(
            row=1, column=2, sticky="w", padx=6, pady=6
        )

        self.lbl_status = ttk.Label(f, text="No data loaded.")
        self.lbl_status.grid(row=2, column=0, columnspan=4, sticky="w", padx=6, pady=10)

    def _build_tab_plot(self):
        # Split into left controls / right preview
        f = self.tab_plot
        f.grid_columnconfigure(0, weight=1)
        f.grid_columnconfigure(1, weight=2)
        f.grid_rowconfigure(0, weight=1)

        left = ttk.Frame(f)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        right = ttk.Frame(f)
        right.grid(row=0, column=1, sticky="nsew")
        right.grid_rowconfigure(1, weight=1)
        right.grid_columnconfigure(0, weight=1)

        # ---------- Controls (left) ----------
        # Section: Plot Type (with explanation)
        lf_type = ttk.Labelframe(left, text="Plot Type")
        lf_type.pack(fill="x", padx=6, pady=(6, 4))
        ttk.Label(lf_type, text="Choose a layout:").grid(
            row=0, column=0, sticky="w", padx=6, pady=4
        )
        self.cb_plot_type = ttk.Combobox(
            lf_type,
            state="readonly",
            width=28,
            textvariable=self.plot_type,
            values=[
                "single",
                "single_right",
                "grid",
                "grid_right",
                "facet",
                "facet_right",
            ],
        )
        self.cb_plot_type.grid(row=0, column=1, sticky="w", padx=6, pady=4)
        lf_type.grid_columnconfigure(1, weight=1)
        self.lbl_type_desc = ttk.Label(
            lf_type,
            foreground="#333",
            text=self._plot_type_description(self.plot_type.get()),
            wraplength=360,
            justify="left",
        )
        self.lbl_type_desc.grid(
            row=1, column=0, columnspan=2, sticky="w", padx=6, pady=(0, 6)
        )

        # Section: Columns
        lf_cols = ttk.Labelframe(left, text="Columns")
        lf_cols.pack(fill="x", padx=6, pady=4)

        ttk.Label(lf_cols, text="X Column (Required)").grid(
            row=0, column=0, sticky="w", padx=6, pady=4
        )
        self.cb_x = ttk.Combobox(lf_cols, values=self.columns, state="readonly")
        self.cb_x.grid(row=0, column=1, sticky="ew", padx=6, pady=4)
        lf_cols.grid_columnconfigure(1, weight=1)

        ttk.Label(lf_cols, text="Y Columns (one or more)").grid(
            row=1, column=0, sticky="nw", padx=6, pady=4
        )
        self.lb_y = tk.Listbox(
            lf_cols, selectmode="extended", exportselection=False, height=8
        )
        self.lb_y.grid(row=1, column=1, sticky="nsew", padx=6, pady=4)

        # Right axis (contextual)
        self.lf_right = ttk.Labelframe(
            left, text="Secondary Right Y (for *_right types)"
        )
        self.lf_right.pack(fill="x", padx=6, pady=4)
        ttk.Label(
            self.lf_right, text="Choose ONE series to plot on the right axis:"
        ).pack(side="left", padx=6, pady=6)
        self.cb_right = ttk.Combobox(
            self.lf_right,
            values=["None"],
            textvariable=self.right_axis_series,
            state="readonly",
            width=28,
        )
        self.cb_right.pack(side="left", padx=6, pady=6)

        # Facet (contextual)
        self.lf_facet = ttk.Labelframe(
            left, text="Facet (split data by a category column)"
        )
        self.lf_facet.pack(fill="x", padx=6, pady=4)
        ttk.Label(self.lf_facet, text="Facet by (category column)").grid(
            row=0, column=0, sticky="w", padx=6, pady=4
        )
        self.cb_facet = ttk.Combobox(
            self.lf_facet, values=["None"], state="readonly", width=28
        )
        self.cb_facet.grid(row=0, column=1, sticky="ew", padx=6, pady=4)
        ttk.Label(
            self.lf_facet,
            text="Tip: Faceting makes one subplot per unique category value.",
        ).grid(row=1, column=0, columnspan=2, sticky="w", padx=6, pady=(0, 6))

        # Section: Titles
        lf_titles = ttk.Labelframe(left, text="Titles")
        lf_titles.pack(fill="x", padx=6, pady=4)
        ttk.Label(lf_titles, text="Title").grid(
            row=0, column=0, sticky="w", padx=6, pady=4
        )
        ttk.Entry(lf_titles, textvariable=self.title_text).grid(
            row=0, column=1, sticky="ew", padx=6, pady=4
        )
        ttk.Label(lf_titles, text="Suptitle").grid(
            row=1, column=0, sticky="w", padx=6, pady=4
        )
        ttk.Entry(lf_titles, textvariable=self.suptitle_text).grid(
            row=1, column=1, sticky="ew", padx=6, pady=4
        )
        lf_titles.grid_columnconfigure(1, weight=1)

        # Section: Style & Layout
        lf_style = ttk.Labelframe(left, text="Style & Layout")
        lf_style.pack(fill="x", padx=6, pady=4)
        ttk.Label(lf_style, text="Kind").grid(
            row=0, column=0, sticky="w", padx=6, pady=4
        )
        ttk.Combobox(
            lf_style,
            values=["line", "scatter"],
            textvariable=self.plot_kind,
            state="readonly",
            width=10,
        ).grid(row=0, column=1, sticky="w", padx=6, pady=4)
        ttk.Label(lf_style, text="Linewidth").grid(
            row=0, column=2, sticky="e", padx=6, pady=4
        )
        ttk.Entry(lf_style, textvariable=self.linewidth, width=8).grid(
            row=0, column=3, sticky="w", padx=6, pady=4
        )
        ttk.Label(lf_style, text="Markersize").grid(
            row=0, column=4, sticky="e", padx=6, pady=4
        )
        ttk.Entry(lf_style, textvariable=self.markersize, width=8).grid(
            row=0, column=5, sticky="w", padx=6, pady=4
        )
        ttk.Label(lf_style, text="Grid columns (grids/facets)").grid(
            row=1, column=0, sticky="w", padx=6, pady=4
        )
        ttk.Spinbox(lf_style, from_=1, to=6, textvariable=self.ncols, width=6).grid(
            row=1, column=1, sticky="w", padx=6, pady=4
        )
        ttk.Label(lf_style, text="Legend location").grid(
            row=1, column=2, sticky="e", padx=6, pady=4
        )
        ttk.Combobox(
            lf_style,
            values=[
                "best",
                "upper right",
                "upper left",
                "lower left",
                "lower right",
                "right",
                "center left",
                "center right",
                "lower center",
                "upper center",
                "center",
            ],
            textvariable=self.legend_loc,
            state="readonly",
            width=14,
        ).grid(row=1, column=3, columnspan=3, sticky="w", padx=6, pady=4)

        # Section: Axis Ranges
        lf_ranges = ttk.Labelframe(left, text="Axis Ranges")
        lf_ranges.pack(fill="x", padx=6, pady=4)
        ttk.Checkbutton(
            lf_ranges,
            text="Auto X",
            variable=self.auto_x,
            command=self._sync_range_states,
        ).grid(row=0, column=0, sticky="w", padx=6, pady=4)
        ttk.Label(lf_ranges, text="X min").grid(
            row=0, column=1, sticky="e", padx=6, pady=4
        )
        self.e_xmin = ttk.Entry(lf_ranges, textvariable=self.x_min, width=10)
        self.e_xmin.grid(row=0, column=2, sticky="w", padx=6, pady=4)
        ttk.Label(lf_ranges, text="X max").grid(
            row=0, column=3, sticky="e", padx=6, pady=4
        )
        self.e_xmax = ttk.Entry(lf_ranges, textvariable=self.x_max, width=10)
        self.e_xmax.grid(row=0, column=4, sticky="w", padx=6, pady=4)

        ttk.Checkbutton(
            lf_ranges,
            text="Auto Left Y",
            variable=self.auto_y_left,
            command=self._sync_range_states,
        ).grid(row=1, column=0, sticky="w", padx=6, pady=4)
        ttk.Label(lf_ranges, text="Yₗ min").grid(
            row=1, column=1, sticky="e", padx=6, pady=4
        )
        self.e_ylmin = ttk.Entry(lf_ranges, textvariable=self.y_left_min, width=10)
        self.e_ylmin.grid(row=1, column=2, sticky="w", padx=6, pady=4)
        ttk.Label(lf_ranges, text="Yₗ max").grid(
            row=1, column=3, sticky="e", padx=6, pady=4
        )
        self.e_ylmax = ttk.Entry(lf_ranges, textvariable=self.y_left_max, width=10)
        self.e_ylmax.grid(row=1, column=4, sticky="w", padx=6, pady=4)

        ttk.Checkbutton(
            lf_ranges,
            text="Auto Right Y",
            variable=self.auto_y_right,
            command=self._sync_range_states,
        ).grid(row=2, column=0, sticky="w", padx=6, pady=4)
        ttk.Label(lf_ranges, text="Yᵣ min").grid(
            row=2, column=1, sticky="e", padx=6, pady=4
        )
        self.e_yrmin = ttk.Entry(lf_ranges, textvariable=self.y_right_min, width=10)
        self.e_yrmin.grid(row=2, column=2, sticky="w", padx=6, pady=4)
        ttk.Label(lf_ranges, text="Yᵣ max").grid(
            row=2, column=3, sticky="e", padx=6, pady=4
        )
        self.e_yrmax = ttk.Entry(lf_ranges, textvariable=self.y_right_max, width=10)
        self.e_yrmax.grid(row=2, column=4, sticky="w", padx=6, pady=4)

        # Section: Ticks
        lf_ticks = ttk.Labelframe(left, text="Ticks")
        lf_ticks.pack(fill="x", padx=6, pady=(4, 8))
        ttk.Checkbutton(
            lf_ticks, text="Auto X ticks", variable=self.auto_time_ticks
        ).grid(row=0, column=0, sticky="w", padx=6, pady=4)
        ttk.Label(lf_ticks, text="X major").grid(
            row=0, column=1, sticky="e", padx=6, pady=4
        )
        ttk.Entry(lf_ticks, textvariable=self.xmaj, width=8).grid(
            row=0, column=2, sticky="w", padx=6, pady=4
        )
        ttk.Label(lf_ticks, text="X minor").grid(
            row=0, column=3, sticky="e", padx=6, pady=4
        )
        ttk.Entry(lf_ticks, textvariable=self.xminr, width=8).grid(
            row=0, column=4, sticky="w", padx=6, pady=4
        )

        ttk.Checkbutton(
            lf_ticks, text="Auto Left Y ticks", variable=self.auto_y_ticks
        ).grid(row=1, column=0, sticky="w", padx=6, pady=4)
        ttk.Label(lf_ticks, text="Y major").grid(
            row=1, column=1, sticky="e", padx=6, pady=4
        )
        ttk.Entry(lf_ticks, textvariable=self.ymaj, width=8).grid(
            row=1, column=2, sticky="w", padx=6, pady=4
        )
        ttk.Label(lf_ticks, text="Y minor").grid(
            row=1, column=3, sticky="e", padx=6, pady=4
        )
        ttk.Entry(lf_ticks, textvariable=self.yminr, width=8).grid(
            row=1, column=4, sticky="w", padx=6, pady=4
        )

        # ---------- Preview (right) ----------
        ttk.Label(
            right, text="Live Preview", font=("Times New Roman", 13, "bold")
        ).grid(row=0, column=0, sticky="w", padx=6, pady=(6, 0))
        self.preview_fig = plt.Figure(figsize=(6.6, 4.8))
        self.preview_canvas = FigureCanvasTkAgg(self.preview_fig, master=right)
        self.preview_widget = self.preview_canvas.get_tk_widget()
        self.preview_widget.grid(row=1, column=0, sticky="nsew", padx=6, pady=6)
        ttk.Button(right, text="Refresh Preview", command=self.draw_preview).grid(
            row=2, column=0, sticky="e", padx=6, pady=(0, 6)
        )

        # Bind change events to auto-refresh preview (debounced) + contextual sections
        self.cb_plot_type.bind("<<ComboboxSelected>>", self._on_plot_type_change)
        self._wire_live_preview(left)
        self._on_plot_type_change()  # set initial visibility/enabled state
        self._sync_range_states()

    # Helpers for plot type description & contextual UI
    def _plot_type_description(self, key: str) -> str:
        d = {
            "single": "Single axes: plots all selected Y series against X on the same (left) Y-axis.",
            "single_right": "Single axes + Right Y: same as Single, plus ONE selected series on a secondary (right) Y-axis.",
            "grid": "Subplot grid: one subplot per Y series.",
            "grid_right": "Subplot grid + Right Y: each subplot shows its Y series on the left axis; the chosen right-axis series overlays on the right axis in each subplot (if present).",
            "facet": "Facet grid: choose a category column; creates one subplot per unique category. All selected Y series are drawn in each facet.",
            "facet_right": "Facet grid + Right Y: same as Facet, with ONE series plotted on a secondary right axis in each facet.",
        }
        return d.get(key, "")

    def _on_plot_type_change(self, *_):
        key = self.plot_type.get()
        self.lbl_type_desc.configure(text=self._plot_type_description(key))
        # Show/hide right-axis & facet sections
        right_needed = key.endswith("_right")
        facet_needed = key.startswith("facet")
        self.lf_right.pack_forget()
        self.lf_facet.pack_forget()
        if right_needed:
            self.lf_right.pack(fill="x", padx=6, pady=4)
        if facet_needed:
            self.lf_facet.pack(fill="x", padx=6, pady=4)
        # Enable/disable right-axis entries' state
        state = "normal" if right_needed else "disabled"
        for w in self.lf_right.winfo_children():
            try:
                w.configure(state=state)
            except Exception:
                pass
        # Make sure preview updates
        self._schedule_preview()

    def _wire_live_preview(self, root):
        def schedule(*_):
            self._schedule_preview()

        # recurse and bind
        self._bind_recursive(root, schedule)
        # explicit listbox selection
        self.lb_y.bind("<<ListboxSelect>>", schedule)

    def _bind_recursive(self, widget, callback):
        if isinstance(widget, ttk.Combobox):
            widget.bind("<<ComboboxSelected>>", callback)
        cls = widget.winfo_class().lower()
        if "entry" in cls:
            widget.bind("<FocusOut>", callback)
            widget.bind("<Return>", callback)
        if isinstance(widget, ttk.Spinbox):
            widget.bind("<FocusOut>", callback)
            widget.bind("<Return>", callback)
        if isinstance(widget, ttk.Checkbutton) or isinstance(widget, ttk.Radiobutton):
            try:
                widget.configure(command=callback)
            except Exception:
                pass
        for child in widget.winfo_children():
            self._bind_recursive(child, callback)

    def _schedule_preview(self):
        if self._preview_after_id:
            self.after_cancel(self._preview_after_id)
        self._preview_after_id = self.after(
            self._preview_debounce_ms, self.draw_preview
        )

    # ---------- Data loading ----------
    def _browse_file(self):
        path = filedialog.askopenfilename(
            title="Select Data File",
            filetypes=[("Data files", "*.xlsx *.xls *.csv"), ("All files", "*.*")],
        )
        if not path:
            return
        self.e_path.delete(0, tk.END)
        self.e_path.insert(0, path)

    def _load_file(self, path: str, auto=False):
        if not path or not os.path.exists(path):
            if not auto:
                messagebox.showerror("Missing file", "Select a valid CSV/Excel file.")
            return
        self.file_path = path
        self.e_path.delete(0, tk.END)
        self.e_path.insert(0, path)

        # Sheets (Excel)
        self.sheet_names = []
        if path.lower().endswith((".xlsx", ".xls")):
            try:
                xls = pd.ExcelFile(path, engine="openpyxl")
                self.sheet_names = xls.sheet_names
            except Exception as e:
                messagebox.showerror("Excel error", f"Could not read workbook: {e}")
                return
            self.cb_sheet.configure(values=self.sheet_names)
            last_sheet = self.settings.get("last_sheet_name", "")
            if last_sheet in self.sheet_names:
                self.current_sheet.set(last_sheet)
            elif self.sheet_names:
                self.current_sheet.set(self.sheet_names[0])
            self._load_sheet()
        else:
            try:
                self.df = pd.read_csv(path)
            except Exception as e:
                messagebox.showerror("CSV error", f"Could not read CSV: {e}")
                return
            self._post_load_dataframe()

        self.settings["last_file_path"] = path
        self._save_settings()
        self._schedule_preview()

    def _load_sheet(self):
        if not self.file_path or not self.current_sheet.get():
            messagebox.showerror("Missing info", "Pick a file and a sheet.")
            return
        try:
            self.df = pd.read_excel(
                self.file_path, sheet_name=self.current_sheet.get(), engine="openpyxl"
            )
        except Exception as e:
            messagebox.showerror("Load error", f"Could not load sheet: {e}")
            return
        self._post_load_dataframe()
        self._schedule_preview()

    def _post_load_dataframe(self):
        self.columns = list(self.df.columns)
        self.cb_x.configure(values=self.columns)
        # Y list
        self.lb_y.delete(0, tk.END)
        for c in self.columns:
            self.lb_y.insert(tk.END, c)
        # Facet choices
        self.cb_facet.configure(values=["None"] + self.columns)
        if self.facet_col_saved in (["None"] + self.columns):
            self.cb_facet.set(self.facet_col_saved)
        else:
            self.cb_facet.set("None")
        # Restore X/Y selections where possible
        prev_x = self.settings.get("x_col", "")
        if prev_x in self.columns:
            self.cb_x.set(prev_x)
        elif self.columns:
            self.cb_x.set(self.columns[0])
        prev_y = self.settings.get("y_cols", [])
        for idx, col in enumerate(self.columns):
            if col in prev_y:
                self.lb_y.selection_set(idx)
        # Right-axis choices mirror Y (+ None)
        self.cb_right.configure(values=["None"] + self.columns)
        if self.right_axis_series.get() not in (["None"] + self.columns):
            self.right_axis_series.set("None")

        self.lbl_status.configure(
            text=f"Loaded: {os.path.basename(self.file_path)}"
            + (
                f" | Sheet: {self.current_sheet.get()}"
                if self.current_sheet.get()
                else ""
            )
            + f" | Rows: {len(self.df)} | Cols: {len(self.columns)}"
        )

    def _get_selected_y(self):
        return [self.lb_y.get(i) for i in self.lb_y.curselection()]

    # ---------- Plotting helpers ----------
    def _apply_ticks(self, ax, set_x=True):
        if set_x:
            if self.auto_time_ticks.get():
                ax.xaxis.set_major_locator(AutoLocator())
                ax.xaxis.set_minor_locator(AutoMinorLocator())
            else:
                xm = _safe_float(self.xmaj.get())
                xmn = _safe_float(self.xminr.get())
                if xm:
                    ax.xaxis.set_major_locator(MultipleLocator(xm))
                if xmn:
                    ax.xaxis.set_minor_locator(MultipleLocator(xmn))
        if self.auto_y_ticks.get():
            ax.yaxis.set_major_locator(AutoLocator())
            ax.yaxis.set_minor_locator(AutoMinorLocator())
        else:
            ym = _safe_float(self.ymaj.get())
            ymn = _safe_float(self.yminr.get())
            if ym:
                ax.yaxis.set_major_locator(MultipleLocator(ym))
            if ymn:
                ax.yaxis.set_minor_locator(MultipleLocator(ymn))
        ax.minorticks_on()
        ax.tick_params(axis="both", which="major", labelsize=TICK_LABELSIZE)

    def _apply_limits(self, ax, which="left", set_x=True):
        if set_x and not self.auto_x.get():
            xmin = _safe_float(self.x_min.get())
            xmax = _safe_float(self.x_max.get())
            if xmin is not None and xmax is not None:
                ax.set_xlim(xmin, xmax)
        if which == "left":
            if not self.auto_y_left.get():
                ymin = _safe_float(self.y_left_min.get())
                ymax = _safe_float(self.y_left_max.get())
                if ymin is not None and ymax is not None:
                    ax.set_ylim(ymin, ymax)
        else:
            if not self.auto_y_right.get():
                ymin = _safe_float(self.y_right_min.get())
                ymax = _safe_float(self.y_right_max.get())
                if ymin is not None and ymax is not None:
                    ax.set_ylim(ymin, ymax)

    def _plot_series(self, ax, x, y, label):
        if self.plot_kind.get() == "scatter":
            return ax.scatter(x, y, s=max(1, self.markersize.get()), label=label)
        (h,) = ax.plot(
            x,
            y,
            linestyle=LINESTYLE_DEFAULT,
            marker=MARKER_DEFAULT,
            linewidth=max(0.1, self.linewidth.get()),
            markersize=max(0.1, self.markersize.get()),
            label=label,
        )
        return h

    # ---------- Live preview ----------
    def draw_preview(self):
        self._preview_after_id = None
        self.preview_fig.clf()

        if self.df is None:
            ax = self.preview_fig.add_subplot(111)
            ax.text(
                0.5,
                0.5,
                "Load data on the Data tab.",
                ha="center",
                va="center",
                fontsize=12,
            )
            self.preview_canvas.draw()
            return

        xcol = self.cb_x.get()
        ycols = self._get_selected_y()
        if not xcol or not ycols:
            ax = self.preview_fig.add_subplot(111)
            ax.text(
                0.5,
                0.5,
                "Select X and at least one Y.",
                ha="center",
                va="center",
                fontsize=12,
            )
            self.preview_canvas.draw()
            return

        # Thin huge datasets for speed
        df = self.df
        if len(df) > 20000:
            df = df.iloc[:: max(1, len(df) // 20000)].copy()

        x = pd.to_numeric(df[xcol], errors="coerce")
        data = {col: pd.to_numeric(df[col], errors="coerce") for col in ycols}
        ptype = self.plot_type.get()
        legend_loc = self.legend_loc.get()

        if ptype in ("single", "single_right"):
            ax = self.preview_fig.add_subplot(111)
            for col in ycols:
                self._plot_series(ax, x, data[col], col)
            if ptype.endswith("right"):
                r = self.right_axis_series.get()
                if r == "None" or r not in df.columns:
                    ax.text(
                        0.5,
                        0.1,
                        "Select a Right-Y series.",
                        ha="center",
                        transform=ax.transAxes,
                        fontsize=10,
                    )
                else:
                    ax2 = ax.twinx()
                    self._plot_series(ax2, x, pd.to_numeric(df[r], errors="coerce"), r)
                    self._apply_limits(ax2, which="right", set_x=False)
                    self._apply_ticks(ax2, set_x=False)
                    ax2.set_ylabel(
                        r, fontsize=LABEL_FONTSIZE, rotation=-90, labelpad=12
                    )

            ax.set_xlabel(xcol, fontsize=LABEL_FONTSIZE)
            ax.set_ylabel(
                ", ".join([c for c in ycols if c != self.right_axis_series.get()])
                or "Y",
                fontsize=LABEL_FONTSIZE,
            )
            self._apply_ticks(ax)
            self._apply_limits(ax, which="left")
            ax.set_title(
                self.title_text.get() or "Preview", fontsize=SUBPLOT_TITLE_FONTSIZE
            )
            self.preview_fig.suptitle(
                self.suptitle_text.get(), fontsize=SUPTITLE_FONTSIZE, y=SUPTITLE_Y
            )
            ax.legend(
                loc=legend_loc, fontsize=LABEL_FONTSIZE - 1, ncol=min(3, len(ycols))
            )

        elif ptype in ("grid", "grid_right"):
            n = len(ycols)
            ncols = max(1, int(self.ncols.get()))
            nrows = int(np.ceil(n / ncols))
            axs = self.preview_fig.subplots(nrows=nrows, ncols=ncols)
            axs = np.array(axs).reshape(-1)
            right_series = (
                self.right_axis_series.get() if ptype.endswith("right") else None
            )

            for i, col in enumerate(ycols):
                ax = axs[i]
                self._plot_series(ax, x, data[col], col)
                if right_series and right_series in df.columns:
                    ax2 = ax.twinx()
                    self._plot_series(
                        ax2,
                        x,
                        pd.to_numeric(df[right_series], errors="coerce"),
                        right_series,
                    )
                    self._apply_limits(ax2, which="right", set_x=False)
                    self._apply_ticks(ax2, set_x=False)
                ax.set_title(col, fontsize=SUBPLOT_TITLE_FONTSIZE)
                ax.set_xlabel(xcol, fontsize=LABEL_FONTSIZE)
                ax.set_ylabel(col, fontsize=LABEL_FONTSIZE)
                self._apply_ticks(ax)
                self._apply_limits(ax, which="left")
                ax.legend(loc="best", fontsize=LABEL_FONTSIZE - 1)

            for j in range(n, len(axs)):
                axs[j].set_visible(False)
            self.preview_fig.tight_layout(rect=[0, 0, 1, 0.94])
            self.preview_fig.suptitle(
                self.suptitle_text.get() or self.title_text.get(),
                fontsize=SUPTITLE_FONTSIZE,
                y=SUPTITLE_Y,
            )

        else:  # facet / facet_right
            facet_col = self.cb_facet.get()
            if facet_col in (None, "", "None"):
                ax = self.preview_fig.add_subplot(111)
                ax.text(
                    0.5,
                    0.5,
                    "Choose a Facet column (category) for facet plots.",
                    ha="center",
                    va="center",
                    fontsize=12,
                )
                self.preview_canvas.draw()
                return
            cats = pd.Series(df[facet_col]).astype("category")
            levels = list(cats.cat.categories)
            if not levels:
                ax = self.preview_fig.add_subplot(111)
                ax.text(
                    0.5,
                    0.5,
                    f"Column '{facet_col}' has no categories.",
                    ha="center",
                    va="center",
                    fontsize=12,
                )
                self.preview_canvas.draw()
                return

            n = len(levels)
            ncols = max(1, int(self.ncols.get()))
            nrows = int(np.ceil(n / ncols))
            axs = self.preview_fig.subplots(nrows=nrows, ncols=ncols)
            axs = np.array(axs).reshape(-1)
            right_series = (
                self.right_axis_series.get() if ptype.endswith("right") else None
            )

            for i, lvl in enumerate(levels):
                ax = axs[i]
                mask = cats == lvl
                x_sub = pd.to_numeric(df.loc[mask, xcol], errors="coerce")
                for col in ycols:
                    y_sub = pd.to_numeric(df.loc[mask, col], errors="coerce")
                    self._plot_series(ax, x_sub, y_sub, col)
                if right_series and right_series in df.columns:
                    ax2 = ax.twinx()
                    y_r = pd.to_numeric(df.loc[mask, right_series], errors="coerce")
                    self._plot_series(ax2, x_sub, y_r, right_series)
                    self._apply_limits(ax2, which="right", set_x=False)
                    self._apply_ticks(ax2, set_x=False)
                ax.set_title(f"{facet_col} = {lvl}", fontsize=SUBPLOT_TITLE_FONTSIZE)
                ax.set_xlabel(xcol, fontsize=LABEL_FONTSIZE)
                ax.set_ylabel(", ".join(ycols), fontsize=LABEL_FONTSIZE)
                self._apply_ticks(ax)
                self._apply_limits(ax, which="left")
                ax.legend(loc="best", fontsize=LABEL_FONTSIZE - 1)

            for j in range(n, len(axs)):
                axs[j].set_visible(False)
            self.preview_fig.tight_layout(rect=[0, 0, 1, 0.94])
            self.preview_fig.suptitle(
                self.suptitle_text.get() or self.title_text.get(),
                fontsize=SUPTITLE_FONTSIZE,
                y=SUPTITLE_Y,
            )

        self.preview_canvas.draw()

    # ---------- Limits/ticks state ----------
    def _sync_range_states(self):
        for widget, flag in [(self.e_xmin, self.auto_x), (self.e_xmax, self.auto_x)]:
            widget.configure(state="disabled" if flag.get() else "normal")
        for widget, flag in [
            (self.e_ylmin, self.auto_y_left),
            (self.e_ylmax, self.auto_y_left),
        ]:
            widget.configure(state="disabled" if flag.get() else "normal")
        for widget, flag in [
            (self.e_yrmin, self.auto_y_right),
            (self.e_yrmax, self.auto_y_right),
        ]:
            widget.configure(state="disabled" if flag.get() else "normal")

    # ---------- Full render/export ----------
    def render_plots(self):
        if self.df is None:
            messagebox.showerror("No data", "Load a file/sheet first.")
            return
        xcol = self.cb_x.get()
        ycols = self._get_selected_y()
        if not xcol:
            messagebox.showerror("Missing X", "Select an X column.")
            return
        if not ycols:
            messagebox.showerror("Missing Y", "Select at least one Y column.")
            return

        ptype = self.plot_type.get()
        x = pd.to_numeric(self.df[xcol], errors="coerce")
        data = {col: pd.to_numeric(self.df[col], errors="coerce") for col in ycols}
        legend_loc = self.legend_loc.get()
        figs = []

        def apply_common(ax):
            ax.set_xlabel(xcol, fontsize=LABEL_FONTSIZE)
            self._apply_ticks(ax)
            self._apply_limits(ax, which="left")
            ax.legend(
                loc=legend_loc if ptype.startswith("single") else "best",
                fontsize=LABEL_FONTSIZE,
                ncol=min(3, len(ycols)),
            )

        if ptype in ("single", "single_right"):
            fig, ax = plt.subplots(figsize=(11, 8.5))
            fig.subplots_adjust(left=0.075, right=0.92, bottom=0.14, top=0.91)
            for col in ycols:
                self._plot_series(ax, x, data[col], col)

            right_choice = None
            if ptype.endswith("right"):
                right_choice = self.right_axis_series.get()
                if (
                    not right_choice
                    or right_choice == "None"
                    or right_choice not in self.df.columns
                ):
                    messagebox.showwarning(
                        "Right Y", "Select a valid series for the right Y-axis."
                    )
                else:
                    ax2 = ax.twinx()
                    self._plot_series(
                        ax2,
                        x,
                        pd.to_numeric(self.df[right_choice], errors="coerce"),
                        right_choice,
                    )
                    self._apply_limits(ax2, which="right", set_x=False)
                    self._apply_ticks(ax2, set_x=False)
                    ax2.set_ylabel(
                        right_choice, fontsize=LABEL_FONTSIZE, rotation=-90, labelpad=15
                    )

            ax.set_ylabel(
                ", ".join([c for c in ycols if c != right_choice]) or "Y",
                fontsize=LABEL_FONTSIZE,
            )
            ax.set_title(self.title_text.get(), fontsize=SUBPLOT_TITLE_FONTSIZE)
            fig.suptitle(
                self.suptitle_text.get(), fontsize=SUPTITLE_FONTSIZE, y=SUPTITLE_Y
            )
            apply_common(ax)
            figs.append(fig)

        elif ptype in ("grid", "grid_right"):
            n = len(ycols)
            ncols = max(1, int(self.ncols.get()))
            nrows = int(np.ceil(n / ncols))
            fig, axes = plt.subplots(
                nrows=nrows, ncols=ncols, figsize=(12, max(6, 3 * nrows))
            )
            axes = np.array(axes).reshape(-1)
            fig.subplots_adjust(
                left=0.07, right=0.95, bottom=0.08, top=0.9, hspace=0.35, wspace=0.25
            )

            right_choice = (
                self.right_axis_series.get() if ptype.endswith("right") else None
            )
            for i, col in enumerate(ycols):
                ax = axes[i]
                self._plot_series(ax, x, data[col], col)
                if right_choice and right_choice in self.df.columns:
                    ax2 = ax.twinx()
                    self._plot_series(
                        ax2,
                        x,
                        pd.to_numeric(self.df[right_choice], errors="coerce"),
                        right_choice,
                    )
                    self._apply_limits(ax2, which="right", set_x=False)
                    self._apply_ticks(ax2, set_x=False)
                ax.set_title(col, fontsize=SUBPLOT_TITLE_FONTSIZE)
                ax.set_ylabel(col, fontsize=LABEL_FONTSIZE)
                apply_common(ax)

            for j in range(len(ycols), len(axes)):
                axes[j].set_visible(False)
            fig.suptitle(
                self.suptitle_text.get() or self.title_text.get(),
                fontsize=SUPTITLE_FONTSIZE,
                y=SUPTITLE_Y,
            )
            figs.append(fig)

        else:  # facet / facet_right
            facet_col = self.cb_facet.get()
            if facet_col in (None, "", "None"):
                messagebox.showerror(
                    "Facet missing", "Choose a Facet column for facet plots."
                )
                return
            cats = pd.Series(self.df[facet_col]).astype("category")
            levels = list(cats.cat.categories)
            if not levels:
                messagebox.showerror(
                    "No categories", f"Column '{facet_col}' has no categories."
                )
                return

            n = len(levels)
            ncols = max(1, int(self.ncols.get()))
            nrows = int(np.ceil(n / ncols))
            fig, axes = plt.subplots(
                nrows=nrows, ncols=ncols, figsize=(13, max(6, 3 * nrows))
            )
            axes = np.array(axes).reshape(-1)
            fig.subplots_adjust(
                left=0.07, right=0.95, bottom=0.08, top=0.9, hspace=0.35, wspace=0.25
            )

            right_choice = (
                self.right_axis_series.get() if ptype.endswith("right") else None
            )
            for i, lvl in enumerate(levels):
                ax = axes[i]
                mask = cats == lvl
                x_sub = pd.to_numeric(self.df.loc[mask, xcol], errors="coerce")
                for col in ycols:
                    y_sub = pd.to_numeric(self.df.loc[mask, col], errors="coerce")
                    self._plot_series(ax, x_sub, y_sub, col)
                if right_choice and right_choice in self.df.columns:
                    ax2 = ax.twinx()
                    y_r = pd.to_numeric(
                        self.df.loc[mask, right_choice], errors="coerce"
                    )
                    self._plot_series(ax2, x_sub, y_r, right_choice)
                    self._apply_limits(ax2, which="right", set_x=False)
                    self._apply_ticks(ax2, set_x=False)
                ax.set_title(f"{facet_col} = {lvl}", fontsize=SUBPLOT_TITLE_FONTSIZE)
                ax.set_ylabel(", ".join(ycols), fontsize=LABEL_FONTSIZE)
                apply_common(ax)

            for j in range(n, len(axes)):
                axes[j].set_visible(False)
            fig.suptitle(
                self.suptitle_text.get() or self.title_text.get(),
                fontsize=SUPTITLE_FONTSIZE,
                y=SUPTITLE_Y,
            )
            figs.append(fig)

        # Show non-blocking
        for fig in figs:
            try:
                fig.canvas.manager.set_window_title("General Plotter – Figure")
            except Exception:
                pass
        plt.show(block=False)
        self._save_settings()

    def export_plots(self):
        if not plt.get_fignums():
            messagebox.showerror("Nothing to export", "Render a plot first.")
            return
        outdir = filedialog.askdirectory(title="Choose export folder")
        if not outdir:
            return
        base = self.title_text.get().strip() or "plot"
        base = (
            "".join(ch for ch in base if ch.isalnum() or ch in (" ", "_", "-"))
            .strip()
            .replace(" ", "_")
        )
        for i, num in enumerate(plt.get_fignums(), 1):
            fig = plt.figure(num)
            png = os.path.join(outdir, f"{base}_{i:02d}.png")
            svg = os.path.join(outdir, f"{base}_{i:02d}.svg")
            try:
                fig.savefig(png, dpi=300, bbox_inches="tight")
                fig.savefig(svg, bbox_inches="tight")
            except Exception as e:
                messagebox.showerror("Export error", f"Could not save figure {i}: {e}")
                return
        messagebox.showinfo(
            "Export complete", f"Saved {len(plt.get_fignums())} figure(s) to:\n{outdir}"
        )


if __name__ == "__main__":
    app = GeneralPlotter()
    try:
        app.draw_preview()
    except Exception:
        pass
    app.mainloop()
    try:
        plt.close("all")
    except Exception:
        pass
    sys.exit(0)
