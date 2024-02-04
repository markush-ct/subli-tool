import ttkbootstrap as ttkb
from ttkbootstrap.constants import *
from ttkbootstrap.dialogs.dialogs import Messagebox
from tkinter.filedialog import askdirectory
from tkinter.filedialog import askopenfilename
import pathlib
from sublitool import SubliTool


class SubliToolGui(ttkb.Frame):
    def __init__(self, master):
        super().__init__(master, padding=15)
        self.pack(fill=BOTH, expand=YES)

        self.base_dir = pathlib.Path().absolute().as_posix()
        self.lineup_var = ttkb.StringVar(value=f"{self.base_dir}/lineup.csv")
        self.destination_var = ttkb.StringVar(value=f"{self.base_dir}/export")
        self.config_var = ttkb.StringVar(value="Select config file")
        self.config = "jersey"

        option_text = "Complete the form to begin automation"
        self.option_lf = ttkb.Labelframe(self, text=option_text, padding=15)
        self.option_lf.pack(fill=X, expand=YES, anchor=N)

        self.create_lineup_row()
        self.create_destination_row()
        self.create_config_row()
        self.progressbar = ttkb.Progressbar(
            master=self,
            mode=INDETERMINATE,
            bootstyle=(STRIPED, SUCCESS)
        )
        self.progressbar.pack(fill=X, expand=YES, pady=15)
        self.create_run_button()

    def create_lineup_row(self):
        """Add lineup row to labelframe"""
        lineup_row = ttkb.Frame(self.option_lf)
        lineup_row.pack(fill=X, expand=YES)
        lineup_lbl = ttkb.Label(lineup_row, text="Lineup", width=8)
        lineup_lbl.pack(side=LEFT, padx=(15, 0))
        lineup_ent = ttkb.Entry(lineup_row, textvariable=self.lineup_var)
        lineup_ent.pack(side=LEFT, fill=X, expand=YES, padx=5)
        browse_btn = ttkb.Button(
            master=lineup_row,
            text="Browse",
            command=self.on_file_browse,
            width=8
        )
        browse_btn.pack(side=LEFT, padx=5)

    def create_destination_row(self):
        """Add destination row to labelframe"""
        destination_row = ttkb.Frame(self.option_lf)
        destination_row.pack(fill=X, expand=YES, pady=15)
        destination_lbl = ttkb.Label(
            destination_row, text="Destination", width=8)
        destination_lbl.pack(side=LEFT, padx=(15, 0))
        destination_ent = ttkb.Entry(
            destination_row, textvariable=self.destination_var)
        destination_ent.pack(side=LEFT, fill=X, expand=YES, padx=5)
        browse_btn = ttkb.Button(
            master=destination_row,
            text="Browse",
            command=self.on_dir_browse,
            width=8
        )
        browse_btn.pack(side=LEFT, padx=5)

    def create_config_row(self):
        """Add config row to labelframe"""
        config_row = ttkb.Frame(self.option_lf)
        config_row.pack(fill=X, expand=YES)
        config_lbl = ttkb.Label(config_row, text="Config", width=8)
        config_lbl.pack(side=LEFT, padx=(15, 0))
        config_menu = ttkb.Menubutton(
            config_row, text="Select config file", textvariable=self.config_var)
        config_menu.pack(side=LEFT, fill=X, expand=YES, padx=5)
        config_menu_items = ttkb.Menu(config_menu)
        items = ['jersey', 'short']
        for item in items:
            config_menu_items.add_radiobutton(
                label=item.title(), command=lambda item=item: self.on_config_change(item))
        config_menu['menu'] = config_menu_items

    def create_run_button(self):
        run_row = ttkb.Frame(self)
        run_row.pack(fill=BOTH, expand=YES)
        self.run_btn = ttkb.Button(
            master=run_row,
            text="Automate it!",
            command=self.on_run_click,
            width=8
        )
        self.run_btn.pack(side=RIGHT, padx=5)

    def on_file_browse(self):
        """Callback for file browse"""
        path = askopenfilename(title="Select lineup file", filetypes=[
                               ("CSV File", "*.csv")], initialdir=self.base_dir,
                               initialfile=self.lineup_var)
        if path:
            self.lineup_var.set(path)

    def on_dir_browse(self):
        """Callback for directory browse"""
        path = askdirectory(title="Select destination",
                            initialdir=self.destination_var,
                            mustexist=True)
        if path:
            self.destination_var.set(path)

    def on_config_change(self, item):
        self.config = item
        self.config_var.set(item.title())

    def on_run_click(self):
        self.run_btn.configure(state="disabled")
        self.progressbar.start(10)
        st = SubliTool(self.lineup_var.get(),
                       self.destination_var.get(), self.config)
        mb = Messagebox()
        mb.ok(st.run(), "Info!", True)
        self.run_btn.configure(state="enabled")


if __name__ == '__main__':

    app = ttkb.Window("Markush Subli Tool", "superhero",
                      size=(800, 300), resizable=(True, False), maxsize=(900, 300))
    app.position_center()
    SubliToolGui(app)
    app.mainloop()
