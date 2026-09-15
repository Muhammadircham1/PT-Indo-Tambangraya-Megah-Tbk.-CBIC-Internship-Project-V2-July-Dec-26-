import customtkinter as ctk
import tkinter.filedialog as filedialog
import tkinter.messagebox as messagebox
from PIL import Image
import os
import subprocess
import threading

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        # Window Config
        self.title("Automation System")
        self.geometry("900x600")
        
        # Determine paths
        self.base_dir = os.path.dirname(os.path.abspath(__file__))
        self.logo_path = os.path.join(self.base_dir, "images.png")
        self.bg_path = os.path.join(self.base_dir, "bg_mine.jpg")
        
        # Load Images
        self.logo_image = ctk.CTkImage(light_image=Image.open(self.logo_path),
                                       dark_image=Image.open(self.logo_path),
                                       size=(150, 80))
                                       
        self.large_logo_image = ctk.CTkImage(light_image=Image.open(self.logo_path),
                                             dark_image=Image.open(self.logo_path),
                                             size=(300, 150))
        
        try:
            bg_pil = Image.open(self.bg_path)
            self.bg_image = ctk.CTkImage(light_image=bg_pil, dark_image=bg_pil, size=(700, 600))
        except Exception as e:
            print("Background image not found:", e)
            self.bg_image = None

        # Configure Grid Layout
        self.grid_rowconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        # ------------------- SIDEBAR -------------------
        self.sidebar_frame = ctk.CTkFrame(self, width=200, corner_radius=0, fg_color="#333333")
        self.sidebar_frame.grid(row=0, column=0, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(6, weight=1)

        self.sidebar_logo = ctk.CTkLabel(self.sidebar_frame, image=self.logo_image, text="")
        self.sidebar_logo.grid(row=0, column=0, padx=20, pady=(30, 0))
        
        self.sidebar_title = ctk.CTkLabel(self.sidebar_frame, text="PT Indo Tambangraya Megah Tbk", font=ctk.CTkFont(size=11, weight="bold"))
        self.sidebar_title.grid(row=1, column=0, padx=20, pady=(5, 30))

        self.home_button = ctk.CTkButton(self.sidebar_frame, text="Home", fg_color="transparent", 
                                         hover_color="#444444", text_color="#DDDDDD", height=40, font=ctk.CTkFont(size=14, weight="bold"),
                                         command=self.show_home)
        self.home_button.grid(row=2, column=0, padx=20, pady=10)

        self.monthly_button = ctk.CTkButton(self.sidebar_frame, text="Monthly Margin RPA", fg_color="transparent", 
                                           hover_color="#444444", text_color="#DDDDDD", height=40, font=ctk.CTkFont(size=13, weight="bold"),
                                           command=self.show_monthly_margin)
        self.monthly_button.grid(row=3, column=0, padx=20, pady=10)

        # ------------------- MAIN CONTENT CONTAINER -------------------
        self.main_container = ctk.CTkFrame(self, corner_radius=0)
        self.main_container.grid(row=0, column=1, sticky="nsew")
        self.main_container.grid_rowconfigure(0, weight=1)
        self.main_container.grid_columnconfigure(0, weight=1)
        
        # Variables to store paths
        self.margin_paths = {"Master Data": ctk.StringVar(), "Summary Loading": ctk.StringVar(), "Profitability": ctk.StringVar()}
        
        self.show_home() # Show home by default

    def clear_main_container(self):
        for widget in self.main_container.winfo_children():
            widget.destroy()
            
        # Add Background back if needed
        if self.bg_image:
            self.bg_label = ctk.CTkLabel(self.main_container, image=self.bg_image, text="")
            self.bg_label.place(x=0, y=0, relwidth=1, relheight=1)
            
        # Reset sidebar colors
        for btn in [self.home_button, self.monthly_button]:
            btn.configure(fg_color="transparent")

    def show_home(self):
        self.clear_main_container()
        self.home_button.configure(fg_color="#3B597C")
        
        center_frame = ctk.CTkFrame(self.main_container, fg_color="transparent")
        center_frame.place(relx=0.5, rely=0.45, anchor="center")

        main_logo = ctk.CTkLabel(center_frame, image=self.large_logo_image, text="")
        main_logo.pack(pady=(0, 10))

        main_title = ctk.CTkLabel(center_frame, text="PT Indo Tambangraya Megah Tbk", 
                                  font=ctk.CTkFont(size=24, weight="bold"), text_color="white")
        main_title.pack(pady=5)

        main_subtitle = ctk.CTkLabel(center_frame, text="Welcome to the Automation System", 
                                     font=ctk.CTkFont(size=28, weight="bold"), text_color="white")
        main_subtitle.pack(pady=20)
        
    def create_file_row(self, parent, label_text, string_var):
        row_frame = ctk.CTkFrame(parent, fg_color="transparent")
        row_frame.pack(fill="x", pady=3, padx=20)
        
        label = ctk.CTkLabel(row_frame, text=label_text, width=100, anchor="w", text_color="white", font=ctk.CTkFont(weight="bold"))
        label.pack(side="left", padx=5)
        
        entry = ctk.CTkEntry(row_frame, textvariable=string_var, width=400, fg_color="white", text_color="black")
        entry.pack(side="left", padx=5)
        
        def browse():
            filepath = filedialog.askopenfilename()
            if filepath:
                string_var.set(filepath)
                
        btn = ctk.CTkButton(row_frame, text="Browse", width=80, command=browse, fg_color="#6c8ebf", hover_color="#4f71a1")
        btn.pack(side="left", padx=5)

    def show_monthly_margin(self):
        self.clear_main_container()
        self.monthly_button.configure(fg_color="#3B597C")
        
        form_frame = ctk.CTkFrame(self.main_container, fg_color="#222222", corner_radius=15)
        form_frame.place(relx=0.5, rely=0.5, anchor="center")
        
        title = ctk.CTkLabel(form_frame, text="Monthly Margin RPA Data", font=ctk.CTkFont(size=20, weight="bold"), text_color="white")
        title.pack(pady=(20, 15))
        
        self.create_file_row(form_frame, "Master Data", self.margin_paths["Master Data"])
        self.create_file_row(form_frame, "Summary Loading", self.margin_paths["Summary Loading"])
        self.create_file_row(form_frame, "Profitability", self.margin_paths["Profitability"])
        
        self.status_label = ctk.CTkLabel(form_frame, text="", font=ctk.CTkFont(size=12), text_color="#00FF00")
        self.status_label.pack(pady=(5,0))
        
        submit_btn = ctk.CTkButton(form_frame, text="Run RPA", width=150, height=40, command=self.run_monthly_margin, fg_color="#6c8ebf", hover_color="#4f71a1")
        submit_btn.pack(pady=20)
        
    def run_monthly_margin(self):
        self.status_label.configure(text="Running Monthly Margin RPA...", text_color="#FFFF00")
        self.update()
        
        def task():
            try:
                cmd = ["python", "monthly_margin_rpa.py"]
                master = self.margin_paths["Master Data"].get()
                summary = self.margin_paths["Summary Loading"].get()
                profitability = self.margin_paths["Profitability"].get()
                
                if master: cmd.extend(["--master", master])
                if summary: cmd.extend(["--summary", summary])
                if profitability: cmd.extend(["--profitability", profitability])
                
                result = subprocess.run(cmd, capture_output=True, text=True, cwd=self.base_dir)
                if result.returncode == 0:
                    if "BUSINESS RULE EXCEPTION" in result.stdout:
                        self.status_label.configure(text="Validation Error. Check Terminal/Logs.", text_color="#FF5555")
                        self.after(0, lambda: messagebox.showerror("Validation Error", "BUSINESS RULE EXCEPTION occurred.\nPlease check the terminal/logs."))
                    else:
                        self.status_label.configure(text="RPA Process Completed Successfully!", text_color="#00FF00")
                        self.after(0, lambda: messagebox.showinfo("Success", "RPA Process Completed Successfully!"))
                else:
                    self.status_label.configure(text="Error occurred during RPA process.", text_color="#FF0000")
                    self.after(0, lambda: messagebox.showerror("Error", "Error occurred during RPA process."))
                print(result.stdout)
                print(result.stderr)
            except Exception as e:
                self.status_label.configure(text=f"Failed to start RPA: {str(e)}", text_color="#FF0000")
                self.after(0, lambda err=str(e): messagebox.showerror("Failed", f"Failed to start RPA:\n{err}"))
        
        threading.Thread(target=task).start()

if __name__ == "__main__":
    ctk.set_appearance_mode("dark")
    ctk.set_default_color_theme("blue")
    
    app = App()
    app.mainloop()
