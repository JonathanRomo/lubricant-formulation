import tkinter as tk
from tkinter import ttk

COLORS = {"background": "#f2f4f7", "surface": "#ffffff", "primary": "#e6002d", "primary_dark": "#b80024", "text": "#17202a", "muted": "#607080", "border": "#d9e0e7"}

ACCENT_COLORS = {
    "Rojo": ("#e6002d", "#b80024"),
    "Azul": ("#1765c1", "#0f4f9a"),
    "Verde": ("#16826a", "#106854"),
    "Índigo": ("#6152c7", "#4a3da3"),
    "Petróleo": ("#087f8c", "#066671"),
}

def configure(root: tk.Misc, preferences: dict | None = None) -> None:
    preferences = preferences or {"appearance": "Claro", "accent": "Rojo", "scale": 1.0, "density": "Cómoda"}
    dark = preferences.get("appearance") == "Oscuro"
    primary, primary_dark = ACCENT_COLORS.get(preferences.get("accent"), ACCENT_COLORS["Rojo"])
    COLORS.update({"background": "#141c2a" if dark else "#f2f4f7", "surface": "#202c3d" if dark else "#ffffff", "text": "#f2f6fc" if dark else "#17202a", "muted": "#b9c9dc" if dark else "#526579", "border": "#60758f" if dark else "#a9bacb", "primary": primary, "primary_dark": primary_dark})
    field = "#101b2a" if dark else "#ffffff"
    field_border = "#60758f" if dark else "#91a8bf"
    results = "#27384d" if dark else "#f8fafc"
    heading = "#304763" if dark else "#e5edf6"
    hover = "#344b66" if dark else "#e4edf7"
    root.configure(background=COLORS["background"])
    style = ttk.Style(root)
    style.theme_use("clam")
    size = 9 if preferences.get("density") == "Compacta" else 10
    style.configure(".", font=("Segoe UI", size), foreground=COLORS["text"])
    style.configure("TFrame", background=COLORS["surface"])
    style.configure("TLabel", background=COLORS["surface"], foreground=COLORS["text"])
    style.configure("App.TFrame", background=COLORS["background"])
    style.configure("Card.TFrame", background=COLORS["surface"])
    style.configure("Title.TLabel", background=COLORS["background"], foreground=COLORS["text"], font=("Segoe UI", 22, "bold"))
    style.configure("Subtitle.TLabel", background=COLORS["background"], foreground=COLORS["muted"], font=("Segoe UI", 10))
    style.configure("CardTitle.TLabel", background=COLORS["surface"], foreground=COLORS["text"], font=("Segoe UI", 14, "bold"))
    style.configure("CardBody.TLabel", background=COLORS["surface"], foreground=COLORS["muted"])
    theory_text = '#bdbdbd' if dark else '#606060'
    style.configure('Theory.TLabel',background=COLORS['surface'],foreground=theory_text)
    style.configure('TheoryTitle.TLabel',background=COLORS['surface'],foreground=theory_text,
                    font=('Segoe UI',20,'bold'))
    style.configure('Theory.Treeview',background=COLORS['surface'],fieldbackground=COLORS['surface'],
                    foreground=theory_text,rowheight=38,font=('Segoe UI',10),
                    stripebackground=results)
    style.configure('Theory.Treeview.Heading',background=COLORS['surface'],foreground=theory_text,
                    font=('Segoe UI',10,'bold'))
    style.configure("Primary.TButton", background=COLORS["primary"], foreground="white", padding=(14, 8), borderwidth=0)
    style.map("Primary.TButton", background=[("pressed", COLORS["primary_dark"]), ("active", COLORS["primary_dark"]), ("disabled", COLORS["surface"])], foreground=[("disabled", COLORS["muted"])])
    style.configure("TButton", background=COLORS["surface"], foreground=COLORS["text"], bordercolor=COLORS["border"], padding=(10, 7), borderwidth=1)
    style.map("TButton", background=[("pressed", COLORS["primary_dark"]), ("active", COLORS["primary"]), ("disabled", COLORS["surface"])], foreground=[("active", "white"), ("pressed", "white"), ("disabled", COLORS["muted"])])
    style.configure("TCheckbutton", background=COLORS["surface"], foreground=COLORS["text"], padding=(2, 1))
    style.map("TCheckbutton", background=[("active", COLORS["surface"]), ("disabled", COLORS["surface"])], foreground=[("disabled", COLORS["muted"])])
    style.configure("Menu.TButton", background=COLORS["surface"], foreground=COLORS["text"], padding=(12, 10), borderwidth=0)
    style.map("Menu.TButton", background=[("active", COLORS["primary"]), ("pressed", COLORS["primary"])], foreground=[("active", "white"), ("pressed", "white")])
    style.configure("TEntry", fieldbackground=field, foreground=COLORS["text"], insertcolor=COLORS["text"], bordercolor=field_border, lightcolor=field_border, darkcolor=field_border, borderwidth=1)
    style.map("TEntry", bordercolor=[("focus", COLORS["primary"])], lightcolor=[("focus", COLORS["primary"])], darkcolor=[("focus", COLORS["primary"])])
    style.configure("TCombobox", fieldbackground=field, foreground=COLORS["text"], background=COLORS["surface"], arrowcolor=COLORS["text"], bordercolor=field_border, lightcolor=field_border, darkcolor=field_border)
    style.map("TCombobox", fieldbackground=[("readonly", field)], foreground=[("readonly", COLORS["text"])], bordercolor=[("focus", COLORS["primary"])])
    style.configure("Treeview", rowheight=30, background=results, fieldbackground=results, foreground=COLORS["text"], bordercolor=COLORS["border"])
    style.configure("Treeview.Heading", font=("Segoe UI", 10, "bold"), background=heading, foreground=COLORS["text"], bordercolor=COLORS["border"], relief="flat")
    style.map("Treeview.Heading", background=[("active", hover)])
    style.map("Treeview", background=[("selected", COLORS["primary"])], foreground=[("selected", "white")])
    style.configure("TNotebook", background=COLORS["surface"], borderwidth=0, tabmargins=(0, 0, 0, 0))
    style.configure("TNotebook.Tab", background=COLORS["surface"], foreground=COLORS["text"], padding=(14, 7), bordercolor=COLORS["border"])
    style.map("TNotebook.Tab", background=[("selected", COLORS["primary"]), ("active", hover)], foreground=[("selected", "white"), ("active", COLORS["text"])])
    style.configure("TScale", background=COLORS["surface"], troughcolor=COLORS["border"], bordercolor=COLORS["border"])
