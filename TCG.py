import tkinter as tk
from tkinter import filedialog, simpledialog
from PIL import Image, ImageTk
import TCG_format
import TCG_generate

class TCGApp:
    def __init__(self, root):
        self.root = root
        self.root.geometry('700x265')
        self.root.title('TCG')
        self.root.config(bg='linen')

        # State Variables
        self.system_list = ""
        self.base_name = ""
        self.sim_loc = ""
        self.runs = 0
        
        self.options = ['Template 1', 'Template 2']
        self.clicked = tk.StringVar(self.root)
        self.clicked.set(self.options[0])

        self.easter_egg_img = None
        self.error_img = None

        self.setup_ui()

    def setup_ui(self):
        # Titles  
        tk.Label(self.root, text='Test Case Generator', font=('Arial Bold', 20), bg='linen').pack()
        tk.Label(self.root, text='(TCG)', font=('Arial Bold', 15), bg='linen').pack()

        # Template Dropdown  
        self.drop = tk.OptionMenu(self.root, self.clicked, *self.options, command=self.set_template)
        self.drop.place(x=65, y=6)
        self.dropLab = tk.Label(self.root, text='Template: ', font=('Arial', 10), bg='linen')
        self.dropLab.place(x=10, y=10)

        # System List Location  
        self.loc = tk.Button(self.root, text='System List Location', font=('Arial', 10), 
                             command=self.get_list_location, background='light slate gray', 
                             activebackground='light slate blue')
        self.loc.place(x=25, y=100)
        self.locLab = tk.Label(self.root, text='', width=50, font=('Arial', 10), bg='dim gray')
        self.locLab.place(x=180, y=104)

        # Output File Base Name  
        self.nameButton = tk.Button(self.root, text='Set Output File Base Name', font=('Arial', 10), 
                                    command=self.get_base_name, background='light slate gray', 
                                    activebackground='light slate blue')
        self.nameButton.place(x=10, y=150)
        self.nameLab = tk.Label(self.root, text='', width=50, font=('Arial', 10), bg='dim gray')
        self.nameLab.place(x=195, y=154)

        # Start Button & Status Labels  
        self.start_button = tk.Button(self.root, text='Generate Test Case', font=('Arial', 10), 
                                      command=self.run, background='light slate gray', 
                                      activebackground='light slate blue')
        self.start_button.place(x=280, y=200)

        self.startLab = tk.Label(self.root, text='No System List Location Set', bg='red', fg='white', font=('Arial', 10))
        self.startLab2 = tk.Label(self.root, text='No Base Name Set', bg='red', fg='white', font=('Arial', 10))
        self.runLab = tk.Label(self.root, text='Run complete!', bg='spring green', font=('Arial', 10))

    # Methods  

    def get_list_location(self):
        selected_file = filedialog.askopenfilename()
        
        if selected_file:
            self.system_list = selected_file
            
            if self.clicked.get() != 'Template 1':
                TCG_generate.get_elnots(self.system_list, out_list=True)
                
            self.loc.config(bg='spring green')
            self.locLab.config(text=self.system_list, bg='light gray')
            
            if self.startLab.winfo_exists():
                self.startLab.destroy()

    def set_sim_location(self, pop):
        selected_dir = filedialog.askdirectory()
        if selected_dir:
            self.sim_loc = selected_dir
            if hasattr(self, 'SIM_location_btn') and self.SIM_location_btn.winfo_exists():
                self.SIM_location_btn.config(bg='spring green')
        
        # Keep popup on top
        if pop.winfo_exists():
            pop.lift()

    def get_base_name(self):
        name_input = simpledialog.askstring('Base Name Input', 
                                            'Input a base name for output files, enclose in quotes if name has spaces',
                                            parent=self.root)
        if name_input:
            self.base_name = name_input
            self.nameButton.config(bg='spring green')
            self.nameLab.config(text=self.base_name, bg='light gray')
            
            if self.startLab2.winfo_exists():
                self.startLab2.destroy()

    def dog(self):
        dog_window = tk.Toplevel(self.root)
        try:
            img = Image.open('easter_egg.jpg')
            self.easter_egg_img = ImageTk.PhotoImage(img) 
            dog_label = tk.Label(dog_window, image=self.easter_egg_img, borderwidth=0, highlightthickness=0)
            dog_label.pack()
        except Exception:
            tk.Label(dog_window, text="Dog meme missing! (easter_egg.jpg not found)").pack()

    def set_template(self, selection):
        if selection == 'Template 2':
            pop = tk.Toplevel(self.root)
            pop.lift()
            pop.geometry('525x400')
            pop.wm_title('Preliminary Steps for Other Simulator-based Test Case')
            
            intro = tk.Label(pop, text='Currently unable to access restricted simulator library via this program.\n  The following steps need to be taken to generate test case:', font=('Arial Bold', 12))
            intro.pack(pady=10)

            step_one = tk.Label(pop, text='1. Input system list below.  This operation will output a text file\n in the same location as your system list containing all identifiers present.', font=('Arial', 10), fg='Red')
            step_one.pack(pady=5)

            identifier_list_button = tk.Button(pop, text='System List Location', font=('Arial', 10), 
                                               command=lambda: [self.get_list_location(), pop.lift()], 
                                               background='light slate gray', activebackground='light slate blue')
            identifier_list_button.pack(pady=5)

            step_two = tk.Label(pop, text='2. Instructions to download local copy of simulator library for specified identifier', font=('Arial', 10), fg='Red')
            step_two.pack(pady=5)

            step_three = tk.Label(pop, text='3. Use the button below to provide the location of the downloaded simulation files\n Then feel free to click Exit', font=('Arial', 10), fg='Red')
            step_three.pack(pady=5)

            self.SIM_location_btn = tk.Button(pop, text='Downloaded simulation library location', font=('Arial', 10), 
                                              command=lambda: self.set_sim_location(pop), 
                                              background='light slate gray', activebackground='light slate blue')
            self.SIM_location_btn.pack(pady=5)

            exit_button = tk.Button(pop, text='Exit', font=('Arial', 10), command=pop.destroy, background='light slate gray', activebackground='light slate blue')
            exit_button.place(x=250, y=350)

            dog_button = tk.Button(pop, text=r'¯\_(ツ)_/¯', font=('Arial', 6), command=self.dog, background='light slate gray', activebackground='light slate blue')
            dog_button.place(x=480, y=380)

    def run(self):
        template = self.clicked.get()
        
        # Validation  
        if not self.system_list:
            if not self.startLab.winfo_exists():
                self.startLab = tk.Label(self.root, text='No System List Location Set', bg='red', fg='white', font=('Arial', 10))
            self.startLab.place(x=420, y=204)

        if not self.base_name:
            if not self.startLab2.winfo_exists():
                self.startLab2 = tk.Label(self.root, text='No Base Name Set', bg='red', fg='white', font=('Arial', 10))
            self.startLab2.place(x=150, y=204)

        if not template:
            error_window = tk.Toplevel(self.root)
            try:
                img = Image.open('system_error.jpg')
                self.error_img = ImageTk.PhotoImage(img)
                error_label = tk.Label(error_window, image=self.error_img, borderwidth=0, highlightthickness=0)
                error_label.pack()
            except Exception:
                tk.Label(error_window, text="System error! (system_error.jpg not found)").pack()
            return

        # Execution  
        if self.system_list and self.base_name and template:
            test_case, summaries, available = TCG_generate.generate_test_case(self.system_list, self.base_name, self.sim_loc, template)
            TCG_format.format_test_case(test_case, summaries, available, template)
            
            # Update Run Counter
            if self.runs > 0:
                self.runLab.config(text=f'Run complete... Again! (x{self.runs + 1})')
                self.runLab.place(x=265, y=235)
            else:
                self.runLab.place(x=298, y=235)
            self.runs += 1

            # Reset Visuals
            self.loc.config(bg='light slate gray')
            self.nameButton.config(bg='light slate gray')
            self.locLab.config(text='', bg='dim gray')
            self.nameLab.config(text='', bg='dim gray')


if __name__ == "__main__":
    main_window = tk.Tk()
    app = TCGApp(main_window)
    main_window.mainloop()
