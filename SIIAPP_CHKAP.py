import customtkinter as ctk
import logging
import modules.gui_frame as gf
import modules.login as lg
import modules.login_frame as logf
import modules.tab_view as tb
 
class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        self.geometry("1000x600")
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(0, weight=1)
        self.login_frame = logf.LoginFrame(master=self)
        self.login_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")

    def show_app_frame(self, user_access):
        # Replace the login frame with the main app frame
        self.login_frame.destroy()
        self.geometry("1000x600")
        self.my_frame = tb.MyTabView(master=self)
        self.my_frame.grid(row=0, column=0, padx=20, pady=20, sticky="nsew")

        # Disable tabs and set active tab based on access configuration
        if user_access:
            # Disable tabs
            if "disabled_tabs" in user_access:
                logging.info(f"Disabling tabs: {user_access['disabled_tabs']}")
                for tab in user_access["disabled_tabs"]:
                    if tab in self.my_frame._segmented_button._buttons_dict:
                        self.my_frame._segmented_button._buttons_dict[tab].configure(state=ctk.DISABLED)
                    else:
                        logging.warning(f"Tab '{tab}' not found in segmented button dictionary.")

            # Set active tab
            if "active_tab" in user_access:
                active_tab = user_access.get("active_tab")
                if active_tab and active_tab in self.my_frame.tab_configs:
                    # For your custom TabView, use .set() method
                    self.my_frame.set(active_tab)
                    logging.info(f"Set active tab to: {active_tab}")
                else:
                    logging.warning(f"Could not set active tab: {active_tab}")
app = App()
app.title("Checklist Aprobaciones")
app.iconbitmap(tb.icon_path)
app.mainloop()
