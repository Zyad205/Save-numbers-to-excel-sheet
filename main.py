import customtkinter as ctk
from openpyxl import load_workbook

LIGHT_GREY = "#969696"


def create_button(app, func):
    btn = ctk.CTkButton(app, text="Add number", command=func, text_color_disabled="#FFFFFF")
    btn.place(relx=0.5, rely=0.5, anchor="center")



def create_entry(app, var):
    entry = ctk.CTkEntry(app, font=("", 18), textvariable=var, text_color="#FFFFFF")
    entry.place(relx=0.5, rely=0.1, anchor="center", relwidth=0.6)



class SuccessLabel(ctk.CTkLabel):
    def __init__(self, app):
        super().__init__(
            app,
            text="Number successfully saved",
            font=("Arial", 15),
            width=180,
            bg_color=LIGHT_GREY,
            text_color="#FFFFFF")
        
        self.place(anchor="nw", relx=1, rely=.21, relwidth=0.4, relheight=0.1)

        self.position = 1000
        
        self.appear()

    def appear(self):
        while self.position > 580:
    
            self.position -= 1
            self.update()

            self.place(
                anchor="nw",
                relx=self.position / 1000,
                rely=.21,
                relwidth=0.4,
                relheight=0.1)
            
        self.after(500, self.disappear)

    def disappear(self):
        while self.position < 1000:
            
            self.position += 1.2
            self.update()

            self.place(
                anchor="nw",
                relx=self.position / 1000,
                rely=.21,
                relwidth=0.4,
                relheight=0.1)


class FailLabel(ctk.CTkLabel):
    def __init__(self, app):
        super().__init__(
            app,
            text="Invalid Number",
            font=("Arial", 15),
            bg_color="#FF0000",
            text_color="#FFFFFF")
        
        self.place(anchor="sw", relx=0.05, rely=0, relwidth=0.2, relheight=0.1)

        self.position = 0
        
        self.appear()

    def appear(self):
        while self.position < 120:
    
            self.position += 1
            self.update()

            self.place(
                anchor="sw",
                relx=0.05,
                rely=self.position / 1000,
                relwidth=0.2,
                relheight=0.1)
            
        self.after(500, self.disappear)

    def disappear(self):
        while self.position > 0:
            
            self.position -= 1.2
            self.update()

            self.place(
                anchor="sw",
                relx=0.05,
                rely=self.position / 1000,
                relwidth=0.2,
                relheight=0.1)
        

class App(ctk.CTk):
    def __init__(self):
        super().__init__()

        create_button(self, self.add_num)

        self.entry_var = ctk.StringVar()

        self.minsize(500, 300)
      
        create_entry(self, self.entry_var)


        self.wb = load_workbook("nums.xlsx")

        self.mainloop()

    def add_num(self):
        num = self.entry_var.get()
        
        if num.isnumeric():
            ws = self.wb.worksheets[0]
            ws.append([num])
            self.entry_var.set("")
            self.wb.save("nums.xlsx")
            self.update()
            SuccessLabel(self)
        else:
            FailLabel(self)


App()
