from ui.zatracap import ZabbixAutomationApp
import customtkinter as ctk

if __name__ == "__main__":
    root = ctk.CTk()
    app = ZabbixAutomationApp(root)
    root.mainloop()