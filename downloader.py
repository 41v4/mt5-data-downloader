import datetime
import os
import tkinter as tk
from tkinter import filedialog, ttk, messagebox

import pandas as pd
import pytz

import MetaTrader5 as mt5


class MainApplication(tk.Frame):
    def __init__(self, parent, *args, **kwargs):
        super().__init__(parent, *args, **kwargs)
        self.parent = parent
        self.create_widgets()

    def create_widgets(self):
        self.l_frame = tk.Frame(self)
        self.l_frame.grid(row=0, column=0, sticky="ne")
        sep0 = ttk.Separator(self.l_frame, orient=tk.VERTICAL)
        sep0.grid(row=0, column=1, rowspan=23, padx=0, pady=5, sticky="ns")

        conn_button = tk.Button(self.l_frame, text="Connect", command=self.create_conn)
        conn_button.grid(row=0, column=0)

        self.conn_status_var = tk.StringVar()
        self.conn_status_var.set("Disconnected")
        conn_status = tk.Label(
            self.l_frame, fg="gray40", textvariable=self.conn_status_var
        )
        conn_status.grid(row=1, column=0)

        sep1 = ttk.Separator(self.l_frame, orient=tk.HORIZONTAL)
        sep1.grid(row=2, column=0, padx=5, pady=5, sticky="ew")

        symbols_label = tk.Label(self.l_frame, text="Select symbol:")
        symbols_label.grid(row=3, column=0)
        self.symbols_combobox = ttk.Combobox(self.l_frame, width=10)
        self.symbols_combobox.grid(row=4, column=0)

        sep2 = ttk.Separator(self.l_frame, orient=tk.HORIZONTAL)
        sep2.grid(row=5, column=0, padx=5, pady=5, sticky="ew")

        date_label = tk.Label(self.l_frame, text="Date selection")
        date_label.grid(row=6, column=0)

        from_date_label = tk.Label(self.l_frame, fg="gray40", text="From:")
        from_date_label.grid(row=7, column=0)
        self.from_date_entry = tk.Entry(self.l_frame, width=12)
        self.from_date_entry.grid(row=8, column=0)

        to_date_label = tk.Label(self.l_frame, fg="gray40", text="To:")
        to_date_label.grid(row=9, column=0)
        self.to_date_entry = tk.Entry(self.l_frame, width=12)
        self.to_date_entry.grid(row=10, column=0)

        self.sep_d_var = tk.IntVar()
        self.sep_d_c = tk.Checkbutton(
            self.l_frame,
            text="Sep. days",
            variable=self.sep_d_var,
            onvalue=1,
            offvalue=0,
        )
        self.sep_d_c.grid(row=11, column=0)

        sep3 = ttk.Separator(self.l_frame, orient=tk.HORIZONTAL)
        sep3.grid(row=12, column=0, padx=5, pady=5, sticky="ew")

        tf_label = tk.Label(self.l_frame, text="Select timeframe:")
        tf_label.grid(row=13, column=0, padx=5)
        self.tf = ["Ticks", "M1", "M5", "M15", "M30", "H1", "H4", "D1", "W1", "MN"]
        self.tf_combobox = ttk.Combobox(self.l_frame, width=10, values=self.tf)
        self.tf_combobox.grid(row=14, column=0)

        sep4 = ttk.Separator(self.l_frame, orient=tk.HORIZONTAL)
        sep4.grid(row=15, column=0, padx=5, pady=5, sticky="ew")

        s_dir_label = tk.Label(self.l_frame, text="Save directory:")
        s_dir_label.grid(row=16, column=0)

        chosen_s_dir = tk.StringVar()
        self.s_dir_entry = tk.Entry(self.l_frame, textvariable=chosen_s_dir, width=12)
        self.s_dir_entry.grid(row=17, column=0)

        def select_dir():
            dir_selection = tk.filedialog.askdirectory()
            chosen_s_dir.set(dir_selection)

        s_dir_button = tk.Button(self.l_frame, text="Select dir.", command=select_dir)
        s_dir_button.grid(row=18, column=0)

        s_format_label = tk.Label(self.l_frame, text="Select format:")
        s_format_label.grid(row=19, column=0)
        self.s_formats = ["CSV", "XLSX"]
        self.s_format_combobox = ttk.Combobox(
            self.l_frame, width=10, values=self.s_formats
        )
        self.s_format_combobox.grid(row=20, column=0)

        sep5 = ttk.Separator(self.l_frame, orient=tk.HORIZONTAL)
        sep5.grid(row=21, column=0, padx=5, pady=5, sticky="ew")

        start_button = tk.Button(self.l_frame, text="Start", command=self.start)
        start_button.grid(row=22, column=0)

        self.r_frame = tk.Frame(self)
        self.r_frame.grid(row=0, column=2, sticky="ne")

        logging_label = tk.Label(self.r_frame, text="Logging:")
        logging_label.pack()

        self.logging_listbox = tk.Listbox(self.r_frame, height=23, width=41)
        self.logging_listbox.pack(padx=5)

        clear_button = tk.Button(self.r_frame, text="Clear", command=self.clear)
        clear_button.pack()

    def create_conn(self):
        if mt5.initialize():
            self.conn_status_var.set("Connected")
            self.get_all_symbols()
        else:
            messagebox.showerror("Connection error", "Connection initialization failed")

    def get_all_symbols(self):
        self.symbols = [s.name for s in mt5.symbols_get()]
        self.symbols_combobox["values"] = self.symbols
        self.symbols_combobox.current(0)

    def validate_symbol(self):
        selected_symbol = self.symbols_combobox.get()
        if selected_symbol not in self.symbols:
            return "Incorrect symbol"

    def validate_date(self):
        from_date = self.from_date_entry.get()
        to_date = self.to_date_entry.get()
        try:
            datetime.datetime.strptime(from_date, "%Y-%m-%d")
            datetime.datetime.strptime(to_date, "%Y-%m-%d")
        except ValueError:
            return "Incorrect date format (should be YYYY-MM-DD)"

    def validate_tf(self):
        selected_tf = self.tf_combobox.get()
        if selected_tf not in self.tf:
            return "Incorrect timeframe"

    def validate_dir(self):
        dir_entry = self.s_dir_entry.get()
        if not os.path.exists(dir_entry):
            return "Path does not exists"

    def validate_s_format(self):
        s_format = self.s_format_combobox.get()
        if s_format not in self.s_formats:
            return "Incorrect save format"

    def check_errors(self):
        validations = [
            self.validate_symbol(),
            self.validate_date(),
            self.validate_tf(),
            self.validate_dir(),
            self.validate_s_format(),
        ]
        errors = [f for f in validations if f is not None]
        return errors

    def get_from_date_obj(self):
        from_date = self.from_date_entry.get()
        from_date_strp = datetime.datetime.strptime(from_date, "%Y-%m-%d")
        return from_date_strp

    def get_to_date_obj(self):
        to_date = self.to_date_entry.get()
        to_date_strp = datetime.datetime.strptime(to_date, "%Y-%m-%d")
        return to_date_strp

    def get_from_date_vals(self, from_date):
        year = from_date.year
        month = from_date.month
        day = from_date.day
        return year, month, day

    def get_to_date_vals(self, to_date):
        year = to_date.year
        month = to_date.month
        day = to_date.day
        return year, month, day

    def get_ticks(self, symbol, utc_from, utc_to):
        ticks = mt5.copy_ticks_range(symbol, utc_from, utc_to, mt5.COPY_TICKS_ALL)
        ticks_df = pd.DataFrame(ticks)
        ticks_df["time"] = pd.to_datetime(ticks_df["time"], unit="s")
        return ticks_df

    def get_bars(self, symbol, tf, utc_from, utc_to):
        tf_to_mt = {
            "M1": mt5.TIMEFRAME_M1,
            "M5": mt5.TIMEFRAME_M5,
            "M15": mt5.TIMEFRAME_M15,
            "M30": mt5.TIMEFRAME_M30,
            "H1": mt5.TIMEFRAME_H1,
            "H4": mt5.TIMEFRAME_H4,
            "D1": mt5.TIMEFRAME_D1,
            "W1": mt5.TIMEFRAME_W1,
            "MN": mt5.TIMEFRAME_MN1,
        }
        bars = mt5.copy_rates_range(symbol, tf_to_mt[tf], utc_from, utc_to)
        bars_df = pd.DataFrame(bars)
        bars_df["time"] = pd.to_datetime(bars_df["time"], unit="s")
        return bars_df

    def get_sep_days(self):
        all_days = []
        wknd = [6, 7]
        from_date_obj = self.get_from_date_obj()
        to_date_obj = self.get_to_date_obj()
        delta = datetime.timedelta(days=1)
        while from_date_obj < to_date_obj:
            all_days.append(from_date_obj)
            from_date_obj += delta
        all_days = [day for day in all_days if not day.isoweekday() in wknd]
        return all_days

    def save_data(self, data, filename):
        selected_s_format = self.s_format_combobox.get()
        path_w_filename = os.path.join(self.s_dir_entry.get(), filename)
        if selected_s_format == "CSV":
            data.to_csv(path_w_filename + ".csv", index=False)
        elif selected_s_format == "XLSX":
            data.to_excel(path_w_filename + ".xlsx", index=False)

    def pull_data(self):
        sep_days_state = self.sep_d_var.get()
        selected_symbol = self.symbols_combobox.get()
        selected_tf = self.tf_combobox.get()
        delta = datetime.timedelta(days=1)
        timezone = pytz.timezone("Etc/UTC")
        if sep_days_state:
            all_days = self.get_sep_days()
            for from_day in all_days:
                to_day = from_day + delta
                from_y, from_m, from_d = self.get_from_date_vals(from_day)
                to_y, to_m, to_d = self.get_to_date_vals(to_day)
                utc_from = datetime.datetime(from_y, from_m, from_d, tzinfo=timezone)
                utc_to = datetime.datetime(to_y, to_m, to_d, tzinfo=timezone)
                if selected_tf == "Ticks":
                    data = self.get_ticks(selected_symbol, utc_from, utc_to)
                else:
                    data = self.get_bars(selected_symbol, selected_tf, utc_from, utc_to)
                joined_date = "-".join([str(from_y), str(from_m), str(from_d)])
                filename = "_".join([joined_date, selected_tf, selected_symbol])
                self.save_data(data, filename)
                current_time = datetime.datetime.now().strftime("%H:%M:%S")
                logging_msg = f"{current_time} // Downloaded {filename}"
                self.logging_listbox.insert("end", logging_msg)
        else:
            from_date = self.get_from_date_obj()
            to_date = self.get_to_date_obj()
            from_y, from_m, from_d = self.get_from_date_vals(from_date)
            to_y, to_m, to_d = self.get_to_date_vals(to_date)
            utc_from = datetime.datetime(from_y, from_m, from_d, tzinfo=timezone)
            utc_to = datetime.datetime(to_y, to_m, to_d, tzinfo=timezone)
            if selected_tf == "Ticks":
                data = self.get_ticks(selected_symbol, utc_from, utc_from)
            else:
                data = self.get_bars(selected_symbol, selected_tf, utc_from, utc_to)
            joined_from_date = "-".join([str(from_y), str(from_m), str(from_d)])
            joined_to_date = "-".join([str(to_y), str(to_m), str(to_d)])
            joined_date_range = "_".join([joined_from_date, joined_to_date])
            filename = "_".join([joined_date_range, selected_tf, selected_symbol])
            self.save_data(data, filename)
            current_time = datetime.datetime.now().strftime("%H:%M:%S")
            logging_msg = f"{current_time} // Downloaded {filename}"
            self.logging_listbox.insert("end", logging_msg)

    def start(self):
        errors = self.check_errors()
        if errors:
            messagebox.showerror("Errors", "\n".join(errors))
        else:
            self.pull_data()

    def clear(self):
        self.logging_listbox.delete(0, tk.END)


if __name__ == "__main__":
    root = tk.Tk()
    root.title("MT5 data downloader")
    MainApplication(root).pack(side="top", fill="both", expand=True)
    root.mainloop()
