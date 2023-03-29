import tkinter
import tkinter.messagebox
import customtkinter
import pandas as pd
import win32com.client as win32
import datetime
from datetime import date, datetime

customtkinter.set_appearance_mode("System")  # Modes: "System" (standard), "Dark", "Light"
customtkinter.set_default_color_theme("dark-blue")  # Themes: "blue" (standard), "green", "dark-blue"


class App(customtkinter.CTk):

    def __init__(self):
        super().__init__()

        # configure window
        self.title("Sistema da Logística")
        self.geometry(f"{1200}x{760}")

        # configure grid layout (4x4)
        self.grid_columnconfigure(1, weight=4)
        self.grid_columnconfigure((2, 3), weight=0)
        self.grid_rowconfigure((0), weight=1)

        # create sidebar frame with widgets
        self.sidebar_frame = customtkinter.CTkFrame(self, width=140, corner_radius=0)
        self.sidebar_frame.grid(row=0, column=0, rowspan=4, sticky="nsew")
        self.sidebar_frame.grid_rowconfigure(4, weight=1)
        self.logo_label = customtkinter.CTkLabel(self.sidebar_frame, text="Logística", font=customtkinter.CTkFont(size=20, weight="bold"))
        self.logo_label.grid(row=0, column=0, padx=20, pady=(20, 10))
        self.sidebar_button_1 = customtkinter.CTkButton(self.sidebar_frame, text='Notas atrasadas', command=notas_atrasadas)
        self.sidebar_button_1.grid(row=1, column=0, padx=20, pady=10)
        self.sidebar_button_2 = customtkinter.CTkButton(self.sidebar_frame, text='Entregas de hoje', command=self.sidebar_button_event)
        self.sidebar_button_2.grid(row=2, column=0, padx=20, pady=10)
        self.sidebar_button_3 = customtkinter.CTkButton(self.sidebar_frame, text='Salvar EDI', command=self.sidebar_button_event)
        self.sidebar_button_3.grid(row=3, column=0, padx=20, pady=10)
        
        self.appearance_mode_label = customtkinter.CTkLabel(self.sidebar_frame, text="Aparência:", anchor="w")
        self.appearance_mode_label.grid(row=5, column=0, padx=20, pady=(10, 0))
        self.appearance_mode_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame, values=["Light", "Dark",],
                                                                       command=self.change_appearance_mode_event)
        self.appearance_mode_optionemenu.grid(row=6, column=0, padx=20, pady=(0, 10))
        self.scaling_label = customtkinter.CTkLabel(self.sidebar_frame, text="Escala de zoom:", anchor="w")
        self.scaling_label.grid(row=7, column=0, padx=20, pady=(5, 0))
        self.scaling_optionemenu = customtkinter.CTkOptionMenu(self.sidebar_frame, values=["80%", "90%", "100%", "110%", "120%"],
                                                                        command=self.change_scaling_event)
        self.scaling_optionemenu.grid(row=8, column=0, padx=20, pady=(0, 20))

                # create textbox
        self.textbox = customtkinter.CTkTextbox(self, width=250)
        self.textbox.grid(row=0, column=1, padx=(10, 0), pady=(20, 0), sticky="nsew")

      
        # create scrollable frame
        self.scrollable_frame = customtkinter.CTkScrollableFrame(self, label_text=data_atual)
        self.scrollable_frame.grid(row=0, column=4, padx=(10, 0), pady=(20, 0), sticky="nsew")
        self.scrollable_frame.grid_columnconfigure(0, weight=1)
      
      
    def change_appearance_mode_event(self, new_appearance_mode: str):
        customtkinter.set_appearance_mode(new_appearance_mode)

    def change_scaling_event(self, new_scaling: str):
        new_scaling_float = int(new_scaling.replace("%", "")) / 100
        customtkinter.set_widget_scaling(new_scaling_float)

    def sidebar_button_event(self):
        var_text = ("sidebar_button click")
        print("sidebar_button click")

#lógica antiga
data_atual = date.today()

tabela_em_transito = pd.read_excel('Pedido em transito.xlsx')

# visualizar a base de dados
pd.set_option('display.max_columns', None)

global lista_notas 

lista_notas = tabela_em_transito['Nota']

def notas_atrasadas():
        print ("atrasado")

if __name__ == "__main__":
    app = App()
    app.mainloop()