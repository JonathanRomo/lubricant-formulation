from ui.main_window import RonnaApp
from database import initialize

if __name__ == "__main__":
    initialize()
    RonnaApp().mainloop()
