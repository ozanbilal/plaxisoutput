
import re
import threading
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.scrolledtext import ScrolledText
from types import SimpleNamespace

import export_plaxis_data as core


class PlaxisExportApp(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("PLAXIS Export Tool")
        self.geometry("1120x760")
        self.minsize(980, 680)
        self.busy = False
        self.busy_widgets = []
        self.phase_label_to_name = {}
        self.pile_label_to_name = {}
        self.plate_label_to_name = {}
        self.api_curvepoint_ids = {}
        self._setup_styles()
        self._build_ui()

    def _setup_styles(self):
        style = ttk.Style(self)
        try:
            style.configure("Vertical.TScrollbar", arrowsize=18, width=18)
            style.configure("Horizontal.TScrollbar", arrowsize=18, width=18)
        except Exception:
            pass

    def _build_ui(self):
        root = ttk.Frame(self)
        root.pack(fill="both", expand=True, padx=8, pady=8)
        root.columnconfigure(0, weight=3)
        root.columnconfigure(1, weight=2)
        root.rowconfigure(0, weight=1)

        left = ttk.Frame(root)
        left.grid(row=0, column=0, sticky="nsew", padx=(0, 8))
        left.columnconfigure(0, weight=1)
        left.rowconfigure(0, weight=1)
        self._build_multiphase_tab(left)

        side = ttk.Frame(root)
        side.grid(row=0, column=1, sticky="nsew")
        side.columnconfigure(0, weight=1)
        side.rowconfigure(1, weight=1)

        quick = ttk.LabelFrame(side, text="Quick Actions")
        quick.grid(row=0, column=0, sticky="ew")
        self.run_active_btn = ttk.Button(quick)
        self.run_active_btn.pack(fill="x", padx=6, pady=6)
        self.busy_widgets.append(self.run_active_btn)
        self.run_active_btn.configure(
            text="Run Node Spectrum Analysis",
            command=self.run_node_multiphase_export,
        )

        log_frame = ttk.LabelFrame(side, text="Log")
        log_frame.grid(row=1, column=0, sticky="nsew", pady=(8, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        self.log_box = ScrolledText(log_frame, height=10, wrap="word")
        self.log_box.grid(row=0, column=0, sticky="nsew", padx=6, pady=6)
        self.log_box.configure(state="disabled")

    def _build_multiphase_tab(self, parent):
        self.hist_host = tk.StringVar(value="localhost")
        self.hist_port = tk.StringVar(value="10000")
        self.hist_password = tk.StringVar(value="")
        self.hist_out_struct = tk.StringVar(
            value=r"C:\Users\PC\OneDrive\Desktop\PLAXIS_multiphase_structural_results.xlsx"
        )
        self.hist_out_node = tk.StringVar(
            value=r"C:\Users\PC\OneDrive\Desktop\PLAXIS_multiphase_node_results.xlsx"
        )
        self.hist_out_stress_strain = tk.StringVar(
            value=r"C:\Users\PC\OneDrive\Desktop\PLAXIS_multiphase_stress_strain_output.xlsx"
        )
        self.phase_regex_x = tk.StringVar(value=r"^DD2_X_.*")
        self.phase_regex_y = tk.StringVar(value=r"^DD2_Y_.*")
        self.hist_result_type = tk.StringVar(value="Soil.Ax")
        self.hist_stress_result_type = tk.StringVar(value="Soil.Sigxy")
        self.hist_strain_result_type = tk.StringVar(value="Soil.Gamxy")
        self.hist_time_col = tk.StringVar(value="DynamicTime")
        self.hist_damping = tk.StringVar(value="0.05")
        self.hist_period_start = tk.StringVar(value="0.01")
        self.hist_period_end = tk.StringVar(value="3.00")
        self.hist_period_step = tk.StringVar(value="0.01")
        self.hist_plot_dpi = tk.StringVar(value="180")
        self.hist_save_phase_timehistory = tk.BooleanVar(value=False)
        self.plate_group1_merge_single = tk.BooleanVar(value=False)
        self.plate_group2_merge_single = tk.BooleanVar(value=False)

        canvas = tk.Canvas(parent, highlightthickness=0)
        scroll = tk.Scrollbar(
            parent,
            orient="vertical",
            command=canvas.yview,
            width=24,
            activebackground="#9A9A9A",
            bg="#CFCFCF",
            troughcolor="#F2F2F2",
        )
        scroll.pack(side="right", fill="y")
        canvas.pack(side="left", fill="both", expand=True)
        canvas.configure(yscrollcommand=scroll.set)
        inner = ttk.Frame(canvas)
        window_id = canvas.create_window((0, 0), window=inner, anchor="nw")

        def _on_inner_configure(_e):
            canvas.configure(scrollregion=canvas.bbox("all"))

        def _on_canvas_configure(e):
            canvas.itemconfigure(window_id, width=e.width)

        inner.bind("<Configure>", _on_inner_configure)
        canvas.bind("<Configure>", _on_canvas_configure)

        api = ttk.LabelFrame(inner, text="API Settings")
        api.pack(fill="x", padx=8, pady=8)
        api.columnconfigure(1, weight=1)
        self._add_labeled_entry(api, 0, "Host:", self.hist_host)
        self._add_labeled_entry(api, 1, "Port:", self.hist_port)
        self._add_labeled_entry(api, 2, "Password:", self.hist_password, show="*")
        self._add_labeled_entry(api, 3, "Structural output:", self.hist_out_struct)
        browse_struct_out = ttk.Button(
            api,
            text="Browse",
            command=lambda: self._pick_file(self.hist_out_struct, save=True),
        )
        browse_struct_out.grid(row=3, column=2, padx=6, pady=4, sticky="w")
        self.busy_widgets.append(browse_struct_out)

        self._add_labeled_entry(api, 4, "Node output:", self.hist_out_node)
        browse_node_out = ttk.Button(
            api,
            text="Browse",
            command=lambda: self._pick_file(self.hist_out_node, save=True),
        )
        browse_node_out.grid(row=4, column=2, padx=6, pady=4, sticky="w")
        self.busy_widgets.append(browse_node_out)

        self._add_labeled_entry(api, 5, "Stress-strain output:", self.hist_out_stress_strain)
        browse_stress_out = ttk.Button(
            api,
            text="Browse",
            command=lambda: self._pick_file(self.hist_out_stress_strain, save=True),
        )
        browse_stress_out.grid(row=5, column=2, padx=6, pady=4, sticky="w")
        self.busy_widgets.append(browse_stress_out)

        analysis = ttk.LabelFrame(inner, text="Node Spectrum Settings")
        analysis.pack(fill="x", padx=8, pady=(0, 8))
        analysis.columnconfigure(1, weight=1)
        self._add_labeled_entry(analysis, 0, "Accel result type:", self.hist_result_type)
        self._add_labeled_entry(analysis, 1, "Tauxy result type:", self.hist_stress_result_type)
        self._add_labeled_entry(analysis, 2, "Gamxy result type:", self.hist_strain_result_type)
        self._add_labeled_entry(analysis, 3, "Time column:", self.hist_time_col)
        self._add_labeled_entry(analysis, 4, "Damping ratio:", self.hist_damping)
        self._add_labeled_entry(analysis, 5, "Period start (s):", self.hist_period_start)
        self._add_labeled_entry(analysis, 6, "Period end (s):", self.hist_period_end)
        self._add_labeled_entry(analysis, 7, "Period step (s):", self.hist_period_step)
        self._add_labeled_entry(analysis, 8, "PNG DPI:", self.hist_plot_dpi)
        save_hist_chk = ttk.Checkbutton(
            analysis,
            text="Save node time histories as TXT files",
            variable=self.hist_save_phase_timehistory,
        )
        save_hist_chk.grid(row=9, column=0, columnspan=2, sticky="w", padx=6, pady=(2, 6))
        self.busy_widgets.append(save_hist_chk)

        phase_box = ttk.LabelFrame(inner, text="Phase Selection (Regex + Manual)")
        phase_box.pack(fill="both", padx=8, pady=(0, 8))
        phase_box.columnconfigure(0, weight=1)
        phase_box.columnconfigure(1, weight=1)

        top = ttk.Frame(phase_box)
        top.grid(row=0, column=0, columnspan=2, sticky="ew", padx=6, pady=6)
        top.columnconfigure(1, weight=1)
        top.columnconfigure(3, weight=1)
        ttk.Label(top, text="X regex:").grid(row=0, column=0, sticky="w")
        x_entry = ttk.Entry(top, textvariable=self.phase_regex_x)
        x_entry.grid(row=0, column=1, sticky="ew", padx=(4, 10))
        ttk.Label(top, text="Y regex:").grid(row=0, column=2, sticky="w")
        y_entry = ttk.Entry(top, textvariable=self.phase_regex_y)
        y_entry.grid(row=0, column=3, sticky="ew", padx=(4, 10))
        load_phase_btn = ttk.Button(top, text="Load Phases", command=self.load_phases)
        load_phase_btn.grid(row=0, column=4, sticky="e")
        self.busy_widgets.extend([x_entry, y_entry, load_phase_btn])

        x_frame = ttk.LabelFrame(phase_box, text="X Direction Phases")
        y_frame = ttk.LabelFrame(phase_box, text="Y Direction Phases")
        x_frame.grid(row=1, column=0, sticky="nsew", padx=(6, 3), pady=(0, 6))
        y_frame.grid(row=1, column=1, sticky="nsew", padx=(3, 6), pady=(0, 6))
        x_frame.columnconfigure(0, weight=1)
        x_frame.rowconfigure(1, weight=1)
        y_frame.columnconfigure(0, weight=1)
        y_frame.rowconfigure(1, weight=1)

        x_actions = ttk.Frame(x_frame)
        x_actions.grid(row=0, column=0, sticky="ew", padx=6, pady=(6, 4))
        x_sel_all = ttk.Button(x_actions, text="Select All", command=self.select_all_x_phases)
        x_sel_all.pack(side="left")
        x_clear = ttk.Button(x_actions, text="Clear", command=self.clear_x_phases)
        x_clear.pack(side="left", padx=(6, 0))

        y_actions = ttk.Frame(y_frame)
        y_actions.grid(row=0, column=0, sticky="ew", padx=6, pady=(6, 4))
        y_sel_all = ttk.Button(y_actions, text="Select All", command=self.select_all_y_phases)
        y_sel_all.pack(side="left")
        y_clear = ttk.Button(y_actions, text="Clear", command=self.clear_y_phases)
        y_clear.pack(side="left", padx=(6, 0))

        self.x_phase_list = tk.Listbox(x_frame, selectmode="extended", height=7)
        self.x_phase_list.grid(row=1, column=0, sticky="nsew", padx=6, pady=(0, 6))
        self._style_listbox(self.x_phase_list)
        self._attach_vertical_scrollbar(x_frame, self.x_phase_list, row=1, column=1)
        self.y_phase_list = tk.Listbox(y_frame, selectmode="extended", height=7)
        self.y_phase_list.grid(row=1, column=0, sticky="nsew", padx=6, pady=(0, 6))
        self._style_listbox(self.y_phase_list)
        self._attach_vertical_scrollbar(y_frame, self.y_phase_list, row=1, column=1)
        self.busy_widgets.extend(
            [x_sel_all, x_clear, y_sel_all, y_clear, self.x_phase_list, self.y_phase_list]
        )
        struct_box = ttk.LabelFrame(inner, text="Structural Object Selection")
        struct_box.pack(fill="both", padx=8, pady=(0, 8))
        struct_box.columnconfigure(0, weight=1)
        struct_box.columnconfigure(1, weight=1)
        struct_box.columnconfigure(2, weight=1)

        struct_top = ttk.Frame(struct_box)
        struct_top.grid(row=0, column=0, columnspan=3, sticky="ew", padx=6, pady=6)
        load_struct_btn = ttk.Button(
            struct_top, text="Load Structural Objects", command=self.load_structural_objects
        )
        load_struct_btn.pack(side="left")
        self.busy_widgets.append(load_struct_btn)

        pile_frame = ttk.LabelFrame(struct_box, text="Piles (EmbeddedBeams)")
        p1_frame = ttk.LabelFrame(struct_box, text="Plate Group 1")
        p2_frame = ttk.LabelFrame(struct_box, text="Plate Group 2")
        pile_frame.grid(row=1, column=0, sticky="nsew", padx=(6, 3), pady=(0, 6))
        p1_frame.grid(row=1, column=1, sticky="nsew", padx=3, pady=(0, 6))
        p2_frame.grid(row=1, column=2, sticky="nsew", padx=(3, 6), pady=(0, 6))

        for frame in (pile_frame, p1_frame, p2_frame):
            frame.columnconfigure(0, weight=1)
            frame.columnconfigure(1, weight=0)
            frame.rowconfigure(1, weight=1)

        pile_actions = ttk.Frame(pile_frame)
        pile_actions.grid(row=0, column=0, sticky="ew", padx=6, pady=(6, 4))
        pile_sel_all = ttk.Button(
            pile_actions, text="Select All", command=lambda: self._select_all_listbox(self.pile_list)
        )
        pile_sel_all.pack(side="left")
        pile_clear = ttk.Button(
            pile_actions, text="Clear", command=lambda: self._clear_listbox(self.pile_list)
        )
        pile_clear.pack(side="left", padx=(6, 0))

        p1_actions = ttk.Frame(p1_frame)
        p1_actions.grid(row=0, column=0, sticky="ew", padx=6, pady=(6, 4))
        p1_sel_all = ttk.Button(
            p1_actions,
            text="Select All",
            command=lambda: self._select_all_listbox(self.plate_group1_list),
        )
        p1_sel_all.pack(side="left")
        p1_clear = ttk.Button(
            p1_actions, text="Clear", command=lambda: self._clear_listbox(self.plate_group1_list)
        )
        p1_clear.pack(side="left", padx=(6, 0))
        p1_merge_chk = ttk.Checkbutton(
            p1_actions,
            text="Merge as single profile",
            variable=self.plate_group1_merge_single,
        )
        p1_merge_chk.pack(side="left", padx=(8, 0))

        p2_actions = ttk.Frame(p2_frame)
        p2_actions.grid(row=0, column=0, sticky="ew", padx=6, pady=(6, 4))
        p2_sel_all = ttk.Button(
            p2_actions,
            text="Select All",
            command=lambda: self._select_all_listbox(self.plate_group2_list),
        )
        p2_sel_all.pack(side="left")
        p2_clear = ttk.Button(
            p2_actions, text="Clear", command=lambda: self._clear_listbox(self.plate_group2_list)
        )
        p2_clear.pack(side="left", padx=(6, 0))
        p2_merge_chk = ttk.Checkbutton(
            p2_actions,
            text="Merge as single profile",
            variable=self.plate_group2_merge_single,
        )
        p2_merge_chk.pack(side="left", padx=(8, 0))

        self.pile_list = tk.Listbox(pile_frame, selectmode="extended", height=7)
        self.pile_list.grid(row=1, column=0, sticky="nsew", padx=6, pady=(0, 6))
        self._style_listbox(self.pile_list)
        self._attach_vertical_scrollbar(pile_frame, self.pile_list, row=1, column=1)
        self.plate_group1_list = tk.Listbox(p1_frame, selectmode="extended", height=7)
        self.plate_group1_list.grid(row=1, column=0, sticky="nsew", padx=6, pady=(0, 6))
        self._style_listbox(self.plate_group1_list)
        self._attach_vertical_scrollbar(p1_frame, self.plate_group1_list, row=1, column=1)
        self.plate_group2_list = tk.Listbox(p2_frame, selectmode="extended", height=7)
        self.plate_group2_list.grid(row=1, column=0, sticky="nsew", padx=6, pady=(0, 6))
        self._style_listbox(self.plate_group2_list)
        self._attach_vertical_scrollbar(p2_frame, self.plate_group2_list, row=1, column=1)
        self.busy_widgets.extend(
            [
                pile_sel_all,
                pile_clear,
                p1_sel_all,
                p1_clear,
                p2_sel_all,
                p2_clear,
                p1_merge_chk,
                p2_merge_chk,
                self.pile_list,
                self.plate_group1_list,
                self.plate_group2_list,
            ]
        )

        nodes = ttk.LabelFrame(inner, text="CurvePoints (for Node Spectrum Analysis)")
        nodes.pack(fill="both", padx=8, pady=(0, 8))
        nodes.columnconfigure(0, weight=1)
        nodes.columnconfigure(1, weight=0)
        nodes.rowconfigure(1, weight=1)

        node_actions = ttk.Frame(nodes)
        node_actions.grid(row=0, column=0, sticky="ew", padx=6, pady=6)
        load_nodes_btn = ttk.Button(
            node_actions, text="Load CurvePoints", command=self.load_api_curvepoints
        )
        load_nodes_btn.pack(side="left", padx=(0, 6))
        sel_all_nodes_btn = ttk.Button(
            node_actions, text="Select All", command=self.select_all_api_nodes
        )
        sel_all_nodes_btn.pack(side="left", padx=(0, 6))
        clear_nodes_btn = ttk.Button(
            node_actions, text="Clear", command=self.clear_api_node_selection
        )
        clear_nodes_btn.pack(side="left")
        self.busy_widgets.extend([load_nodes_btn, sel_all_nodes_btn, clear_nodes_btn])

        self.api_nodes_list = tk.Listbox(nodes, selectmode="extended", height=8)
        self.api_nodes_list.grid(row=1, column=0, sticky="nsew", padx=6, pady=(0, 6))
        self._style_listbox(self.api_nodes_list)
        self._attach_vertical_scrollbar(nodes, self.api_nodes_list, row=1, column=1)
        self.busy_widgets.append(self.api_nodes_list)

        run_frame = ttk.LabelFrame(inner, text="Run")
        run_frame.pack(fill="x", padx=8, pady=(0, 10))
        run_struct_btn = ttk.Button(
            run_frame, text="Run Structural Output", command=self.run_structural_moment_export
        )
        run_struct_btn.pack(side="left", padx=6, pady=6)
        run_node_btn = ttk.Button(
            run_frame, text="Run Node Spectrum Analysis", command=self.run_node_multiphase_export
        )
        run_node_btn.pack(side="left", padx=6, pady=6)
        run_stress_btn = ttk.Button(
            run_frame, text="Run Stress-Strain Output", command=self.run_node_stress_strain_export
        )
        run_stress_btn.pack(side="left", padx=6, pady=6)
        self.busy_widgets.extend([run_struct_btn, run_node_btn, run_stress_btn])

    def _add_labeled_entry(self, parent, row, label, var, show=None):
        ttk.Label(parent, text=label).grid(row=row, column=0, sticky="w", padx=6, pady=4)
        entry = ttk.Entry(parent, textvariable=var, show=show if show else "")
        entry.grid(row=row, column=1, sticky="ew", padx=6, pady=4)
        self.busy_widgets.append(entry)
        return entry

    def _pick_file(self, var, save):
        if save:
            path = filedialog.asksaveasfilename(
                defaultextension=".xlsx",
                filetypes=[("Excel file", "*.xlsx"), ("CSV file", "*.csv"), ("All files", "*.*")],
            )
        else:
            path = filedialog.askopenfilename(
                filetypes=[
                    ("Data files", "*.xlsx *.xls *.xlsm *.csv"),
                    ("All files", "*.*"),
                ]
            )
        if path:
            var.set(path)

    def _style_listbox(self, lb):
        lb.configure(
            bg="white",
            fg="black",
            selectbackground="#0A64AD",
            selectforeground="white",
            disabledforeground="#5A5A5A",
            exportselection=False,
            activestyle="none",
        )

    def _attach_vertical_scrollbar(self, parent, listbox, row, column):
        scroll = tk.Scrollbar(parent, orient="vertical", command=listbox.yview, width=16)
        scroll.grid(row=row, column=column, sticky="ns", pady=(0, 6), padx=(0, 6))
        listbox.configure(yscrollcommand=scroll.set)
        self.busy_widgets.append(scroll)
        return scroll

    def _set_busy(self, busy):
        self.busy = busy
        state = "disabled" if busy else "normal"
        for widget in self.busy_widgets:
            try:
                if isinstance(widget, tk.Listbox):
                    # Keep listboxes readable/scrollable while background tasks run.
                    widget.configure(state="normal")
                    continue
                widget.configure(state=state)
            except Exception:
                pass

    def _run_background(self, func):
        if self.busy:
            return
        self._set_busy(True)

        def worker():
            try:
                func()
            except Exception as exc:
                err = str(exc).strip() or repr(exc)
                self.after(0, lambda: messagebox.showerror("Error", err))
                self.after(0, lambda: self.log(f"ERROR: {err}"))
            finally:
                self.after(0, lambda: self._set_busy(False))

        threading.Thread(target=worker, daemon=True).start()

    def log(self, message):
        self.log_box.configure(state="normal")
        self.log_box.insert("end", f"{message}\n")
        self.log_box.see("end")
        self.log_box.configure(state="disabled")

    def _log_async(self, message):
        self.after(0, lambda m=message: self.log(m))

    def _candidate_ports(self):
        ports = []
        try:
            current = int(self.hist_port.get().strip())
            ports.append(current)
        except Exception:
            pass
        for p in (10000, 10001, 10002, 10003):
            if p not in ports:
                ports.append(p)
        return ports

    def _is_retryable_port_error(self, exc):
        msg = (str(exc) or repr(exc)).lower()
        tokens = [
            "httpconnectionpool",
            "max retries exceeded",
            "failed to establish a new connection",
            "winerror 10061",
            "reply code is different from what was sent",
            "no active project",
            "requested attribute 'resulttypes' is not present",
            "requested attribute 'phases' is not present",
            "not plaxis output api",
            "resulttypes missing",
            "request is missing",
            "decryption",
        ]
        return any(token in msg for token in tokens)

    def _load_with_fallback(self, fetcher, action_text):
        host = self.hist_host.get().strip()
        password = self.hist_password.get().strip()
        if not password:
            raise RuntimeError("Password is required for Output API connection.")
        ports = self._candidate_ports()
        last_error = None
        for port in ports:
            try:
                self._log_async(f"{action_text} via port {port}...")
                records = fetcher(host=host, port=port, password=password)
                return records, port
            except Exception as exc:
                last_error = exc
                self._log_async(f"Port {port} failed: {str(exc).strip() or repr(exc)}")
                continue
        raise RuntimeError(
            f"{action_text} failed on all candidate ports. "
            f"Last error: {str(last_error).strip() or repr(last_error)}"
        )

    def _select_all_listbox(self, lb):
        lb.selection_set(0, "end")

    def _clear_listbox(self, lb):
        lb.selection_clear(0, "end")

    def load_phases(self):
        def task():
            records, used_port = self._load_with_fallback(core.list_phases_api, "Load phases")
            x_regex = self.phase_regex_x.get().strip()
            y_regex = self.phase_regex_y.get().strip()
            try:
                rx = re.compile(x_regex) if x_regex else re.compile(r".*")
                ry = re.compile(y_regex) if y_regex else re.compile(r".*")
            except re.error as exc:
                raise RuntimeError(f"Invalid regex: {exc}")

            def ui_update():
                self.hist_port.set(str(used_port))
                self.phase_label_to_name = {}
                self.x_phase_list.delete(0, "end")
                self.y_phase_list.delete(0, "end")
                x_selected = 0
                y_selected = 0
                for idx, rec in enumerate(records):
                    label = f"{int(rec.get('index', idx + 1)):02d} | {rec.get('name', '')}"
                    name = str(rec.get("name", "")).strip()
                    self.phase_label_to_name[label] = name
                    self.x_phase_list.insert("end", label)
                    self.y_phase_list.insert("end", label)
                    if rx.search(name):
                        self.x_phase_list.selection_set(idx)
                        x_selected += 1
                    if ry.search(name):
                        self.y_phase_list.selection_set(idx)
                        y_selected += 1
                self.log(
                    f"Loaded {len(records)} phases (port {used_port}). "
                    f"Auto-selected X={x_selected}, Y={y_selected}."
                )

            self.after(0, ui_update)

        self._run_background(task)

    def select_all_x_phases(self):
        self._select_all_listbox(self.x_phase_list)
        self.log("All X phases selected.")

    def clear_x_phases(self):
        self._clear_listbox(self.x_phase_list)
        self.log("X phase selection cleared.")

    def select_all_y_phases(self):
        self._select_all_listbox(self.y_phase_list)
        self.log("All Y phases selected.")

    def clear_y_phases(self):
        self._clear_listbox(self.y_phase_list)
        self.log("Y phase selection cleared.")

    def load_structural_objects(self):
        def task():
            records, used_port = self._load_with_fallback(
                core.list_structural_objects_api, "Load structural objects"
            )
            piles = records.get("embedded_beams", [])
            plates = records.get("plates", [])

            def ui_update():
                self.hist_port.set(str(used_port))
                self.pile_label_to_name = {}
                self.plate_label_to_name = {}
                self.pile_list.delete(0, "end")
                self.plate_group1_list.delete(0, "end")
                self.plate_group2_list.delete(0, "end")

                for rec in piles:
                    label = str(rec.get("label") or rec.get("name") or "").strip()
                    name = str(rec.get("name") or "").strip()
                    if not label or not name:
                        continue
                    self.pile_label_to_name[label] = name
                    self.pile_list.insert("end", label)

                for rec in plates:
                    label = str(rec.get("label") or rec.get("name") or "").strip()
                    name = str(rec.get("name") or "").strip()
                    if not label or not name:
                        continue
                    self.plate_label_to_name[label] = name
                    self.plate_group1_list.insert("end", label)
                    self.plate_group2_list.insert("end", label)

                self.log(
                    f"Loaded structural objects (port {used_port}): "
                    f"piles={len(self.pile_label_to_name)}, plates={len(self.plate_label_to_name)}."
                )

            self.after(0, ui_update)

        self._run_background(task)

    def _unique_display_label(self, base_label):
        label = base_label
        n = 2
        while label in self.api_curvepoint_ids:
            label = f"{base_label} ({n})"
            n += 1
        return label

    def load_api_curvepoints(self):
        def task():
            records, used_port = self._load_with_fallback(core.list_curve_points_api, "Load CurvePoints")

            def ui_update():
                self.hist_port.set(str(used_port))
                self.api_nodes_list.delete(0, "end")
                self.api_curvepoint_ids = {}
                for rec in records:
                    base = (
                        f"{rec['index']:02d} | {rec['node_name']} | "
                        f"x={rec['x']:.2f} y={rec['y']:.2f}"
                    )
                    if rec["data_from"]:
                        base = f"{base} | {rec['data_from']}"
                    label = self._unique_display_label(base)
                    self.api_curvepoint_ids[label] = rec["id"]
                    self.api_nodes_list.insert("end", label)
                self.api_nodes_list.yview_moveto(0.0)
                self.log(f"Loaded {len(records)} CurvePoints (port {used_port}).")

            self.after(0, ui_update)

        self._run_background(task)

    def select_all_api_nodes(self):
        self.api_nodes_list.selection_set(0, "end")
        self.log("All CurvePoints selected.")

    def clear_api_node_selection(self):
        self.api_nodes_list.selection_clear(0, "end")
        self.log("CurvePoint selection cleared.")

    def _selected_phase_names(self, listbox):
        out = []
        for idx in listbox.curselection():
            label = listbox.get(idx)
            name = self.phase_label_to_name.get(label)
            if name:
                out.append(name)
        return out

    def _selected_object_names(self, listbox, mapping):
        out = []
        for idx in listbox.curselection():
            label = listbox.get(idx)
            name = mapping.get(label)
            if name:
                out.append(name)
        return out

    def _phase_direction_warnings(self, x_phase_names, y_phase_names):
        warnings = []
        x_regex = self.phase_regex_x.get().strip()
        y_regex = self.phase_regex_y.get().strip()
        try:
            rx = re.compile(x_regex) if x_regex else None
        except re.error:
            rx = None
        try:
            ry = re.compile(y_regex) if y_regex else None
        except re.error:
            ry = None

        if rx and x_phase_names and all(not rx.search(name) for name in x_phase_names):
            warnings.append(
                "Warning: selected X phases do not match X regex. Check X list selection."
            )
        if ry and y_phase_names and all(not ry.search(name) for name in y_phase_names):
            warnings.append(
                "Warning: selected Y phases do not match Y regex. Check Y list selection."
            )
        return warnings

    def run_structural_moment_export(self):
        def task():
            if not self.hist_password.get().strip():
                raise RuntimeError("Password is required for Output API connection.")

            x_phase_names = self._selected_phase_names(self.x_phase_list)
            y_phase_names = self._selected_phase_names(self.y_phase_list)
            if not x_phase_names and not y_phase_names:
                raise RuntimeError("Select at least one X or Y phase.")
            for msg in self._phase_direction_warnings(x_phase_names, y_phase_names):
                self._log_async(msg)

            pile_names = self._selected_object_names(self.pile_list, self.pile_label_to_name)
            plate_g1_names = self._selected_object_names(
                self.plate_group1_list, self.plate_label_to_name
            )
            plate_g2_names = self._selected_object_names(
                self.plate_group2_list, self.plate_label_to_name
            )
            if not pile_names and not plate_g1_names and not plate_g2_names:
                raise RuntimeError("Select at least one pile or plate object.")

            host = self.hist_host.get().strip()
            password = self.hist_password.get().strip()
            ports = self._candidate_ports()
            last_error = None
            for port in ports:
                args = SimpleNamespace(
                    host=host,
                    port=port,
                    password=password,
                    x_phase_names=x_phase_names,
                    y_phase_names=y_phase_names,
                    embedded_beam_names=pile_names,
                    plate_group1_names=plate_g1_names,
                    plate_group2_names=plate_g2_names,
                    plate_group1_merge_single_profile=bool(
                        self.plate_group1_merge_single.get()
                    ),
                    plate_group2_merge_single_profile=bool(
                        self.plate_group2_merge_single.get()
                    ),
                    plot_dpi=int(float(self.hist_plot_dpi.get().strip())),
                    out=self.hist_out_struct.get().strip(),
                )
                try:
                    self._log_async(f"Structural output via port {port}...")
                    core.run_structural_moment_export(
                        args, logger=lambda msg: self.after(0, lambda m=msg: self.log(m))
                    )
                    self.after(0, lambda p=port: self.hist_port.set(str(p)))
                    return
                except Exception as exc:
                    last_error = exc
                    err_text = str(exc).strip() or repr(exc)
                    self._log_async(f"Port {port} structural run failed: {err_text}")
                    if not self._is_retryable_port_error(exc):
                        raise RuntimeError(err_text) from exc
                    continue

            raise RuntimeError(
                "Structural output failed on all candidate ports. "
                f"Last error: {str(last_error).strip() or repr(last_error)}"
            )

        self._run_background(task)

    def run_node_multiphase_export(self):
        def task():
            if not self.hist_password.get().strip():
                raise RuntimeError("Password is required for Output API connection.")

            x_phase_names = self._selected_phase_names(self.x_phase_list)
            y_phase_names = self._selected_phase_names(self.y_phase_list)
            if not x_phase_names and not y_phase_names:
                raise RuntimeError("Select at least one X or Y phase.")
            for msg in self._phase_direction_warnings(x_phase_names, y_phase_names):
                self._log_async(msg)

            selected_indices = self.api_nodes_list.curselection()
            selected_labels = [self.api_nodes_list.get(i) for i in selected_indices]
            selected_ids = [
                self.api_curvepoint_ids[label]
                for label in selected_labels
                if label in self.api_curvepoint_ids
            ]

            host = self.hist_host.get().strip()
            password = self.hist_password.get().strip()
            ports = self._candidate_ports()
            last_error = None
            for port in ports:
                args = SimpleNamespace(
                    host=host,
                    port=port,
                    password=password,
                    x_phase_names=x_phase_names,
                    y_phase_names=y_phase_names,
                    curvepoint_id=selected_ids,
                    result_type=self.hist_result_type.get().strip(),
                    time_col=self.hist_time_col.get().strip(),
                    damping=float(self.hist_damping.get().strip()),
                    period_start=float(self.hist_period_start.get().strip()),
                    period_end=float(self.hist_period_end.get().strip()),
                    period_step=float(self.hist_period_step.get().strip()),
                    plot_dpi=int(float(self.hist_plot_dpi.get().strip())),
                    save_node_timehistory_subfolders=bool(
                        self.hist_save_phase_timehistory.get()
                    ),
                    out=self.hist_out_node.get().strip(),
                )
                try:
                    self._log_async(f"Node spectrum analysis via port {port}...")
                    core.run_node_multiphase_spectrum_export(
                        args, logger=lambda msg: self.after(0, lambda m=msg: self.log(m))
                    )
                    self.after(0, lambda p=port: self.hist_port.set(str(p)))
                    return
                except Exception as exc:
                    last_error = exc
                    err_text = str(exc).strip() or repr(exc)
                    self._log_async(f"Port {port} node run failed: {err_text}")
                    if not self._is_retryable_port_error(exc):
                        raise RuntimeError(err_text) from exc
                    continue

            raise RuntimeError(
                "Node spectrum analysis failed on all candidate ports. "
                f"Last error: {str(last_error).strip() or repr(last_error)}"
            )

        self._run_background(task)

    def run_node_stress_strain_export(self):
        def task():
            if not self.hist_password.get().strip():
                raise RuntimeError("Password is required for Output API connection.")

            x_phase_names = self._selected_phase_names(self.x_phase_list)
            y_phase_names = self._selected_phase_names(self.y_phase_list)
            if not x_phase_names and not y_phase_names:
                raise RuntimeError("Select at least one X or Y phase.")
            for msg in self._phase_direction_warnings(x_phase_names, y_phase_names):
                self._log_async(msg)

            selected_indices = self.api_nodes_list.curselection()
            selected_labels = [self.api_nodes_list.get(i) for i in selected_indices]
            selected_ids = [
                self.api_curvepoint_ids[label]
                for label in selected_labels
                if label in self.api_curvepoint_ids
            ]

            host = self.hist_host.get().strip()
            password = self.hist_password.get().strip()
            ports = self._candidate_ports()
            last_error = None
            for port in ports:
                args = SimpleNamespace(
                    host=host,
                    port=port,
                    password=password,
                    x_phase_names=x_phase_names,
                    y_phase_names=y_phase_names,
                    curvepoint_id=selected_ids,
                    stress_result_type=self.hist_stress_result_type.get().strip(),
                    strain_result_type=self.hist_strain_result_type.get().strip(),
                    time_col=self.hist_time_col.get().strip(),
                    plot_dpi=int(float(self.hist_plot_dpi.get().strip())),
                    stress_strain_out=self.hist_out_stress_strain.get().strip(),
                )
                try:
                    self._log_async(f"Stress-strain analysis via port {port}...")
                    core.run_node_stress_strain_export(
                        args, logger=lambda msg: self.after(0, lambda m=msg: self.log(m))
                    )
                    self.after(0, lambda p=port: self.hist_port.set(str(p)))
                    return
                except Exception as exc:
                    last_error = exc
                    err_text = str(exc).strip() or repr(exc)
                    self._log_async(f"Port {port} stress-strain run failed: {err_text}")
                    if not self._is_retryable_port_error(exc):
                        raise RuntimeError(err_text) from exc
                    continue

            raise RuntimeError(
                "Stress-strain analysis failed on all candidate ports. "
                f"Last error: {str(last_error).strip() or repr(last_error)}"
            )

        self._run_background(task)

    def run_curve_api_export(self):
        self.run_node_multiphase_export()


def main():
    app = PlaxisExportApp()
    app.mainloop()


if __name__ == "__main__":
    main()
