import customtkinter as ctk
from PIL import Image, ImageTk
import tkinter as tk
import os
from PRCPayrollSystem.Main.excelImportPage import ExcelImportPage
from PRCPayrollSystem.Main.controller import AppController
from PRCPayrollSystem.Main.resource_utils import resource_path

ctk.set_appearance_mode("light")
ctk.set_default_color_theme("blue")

#ready (ata)
#for the gradient frame 
class GradientFrame(ctk.CTkFrame):
    def __init__(self, master, color1, color2, **kwargs):
        super().__init__(master, **kwargs)
        self.color1 = color1
        self.color2 = color2
        self.canvas = tk.Canvas(self, highlightthickness=0)
        self.canvas.pack(fill="both", expand=True)
        self.bind("<Configure>", self._draw_gradient)

    def _draw_gradient(self, event=None):
        self.canvas.delete("gradient")
        width = self.winfo_width()
        height = self.winfo_height()
        limit = height
        (r1, g1, b1) = self.winfo_rgb(self.color1)
        (r2, g2, b2) = self.winfo_rgb(self.color2)
        r_ratio = float(r2 - r1) / limit
        g_ratio = float(g2 - g1) / limit
        b_ratio = float(b2 - b1) / limit
        for i in range(limit):
            nr = int(r1 + (r_ratio * i))
            ng = int(g1 + (g_ratio * i))
            nb = int(b1 + (b_ratio * i))
            color = f'#{nr//256:02x}{ng//256:02x}{nb//256:02x}'
            self.canvas.create_line(0, i, width, i, tags=("gradient",), fill=color)
        self.canvas.lower("gradient")
#display for the main menu
class MainMenu(ctk.CTkFrame):
    def __init__(self, parent, controller):
        super().__init__(parent)
        self.controller = controller
        self.configure(fg_color="white")
        self.gradient = GradientFrame(self, color1="#ffffff", color2="#b3d1fa")
        self.gradient.pack(fill="both", expand=True)
        logo_path = resource_path("Components/PRClogo.png")
        logo_img = Image.open(logo_path).resize((100, 100), Image.LANCZOS)
        self.logo = ImageTk.PhotoImage(logo_img)
        self.logo_label = tk.Label(self.gradient.canvas, image=self.logo, bg="white", borderwidth=0)
        self.logo_label.place(x=20, y=20)
        self.title_label = ctk.CTkLabel(self.gradient.canvas, text="Payroll System", font=("Montserrat", 36, "normal"), text_color="#000000", bg_color="white")
        self.title_label.place(relx=0.5, y=60, anchor="center")
        button_width = 350
        button_height = 60
        button_font = ("Montserrat", 18, "normal")
        button_pad = 25
        start_y = 150
        btns = [
            ("Import Employee Information", "#0a47b1", lambda: controller.show_frame("ImportEmployeePage")),
            ("Generate Payslip/Reports", "#0a47b1", lambda: controller.show_frame("ExcelImportPage")),
            ("History", "#0a47b1", lambda: controller.show_frame("HistoryPage")),
            ("Exit", "#d0021b", self.exit_app)
        ]
        for i, (text, color, cmd) in enumerate(btns):
            btn = ctk.CTkButton(
                self.gradient,
                text=text,
                fg_color=color,
                hover_color="#003580" if color == "#0a47b1" else "#a30010",
                text_color="white",
                font=button_font,
                corner_radius=30,
                width=button_width,
                height=button_height,
                command=cmd,
                bg_color="white"
            )
            btn.place(relx=0.5, y=start_y + i * (button_height + button_pad), anchor="center")

    def exit_app(self):
        self.controller.destroy()

#starterr (main)
if __name__ == "__main__":
    app = AppController()
    app.mainloop()