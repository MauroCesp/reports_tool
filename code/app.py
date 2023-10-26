
try:
    import customtkinter

    from customtkinter import CTkToplevel, CTkLabel, CTkButton



    from tkinter import filedialog
    """customtkinter.set_appearance_mode("dark")"""
    """customtkinter.set_appearance_mode("system")-default"""
    customtkinter.set_appearance_mode("dark")



    import tkinter as tk
    from tkinter import ttk


    from tkinter import messagebox

    import os

    # Importo toda la informacion de mi controllador
    from controller import *

    # Importo toda la informacion de mi controllador CONSORTIA
    from controller_consortia import *

    # Esta libreria ayuda a extraer archivos zip
    import zipfile

    # Esta libraria me ayuda a trabajar ocn varios porcesos al mismo tiempo.
    import threading

    # Get the directory where the executable is located
    base_path = os.path.dirname(os.path.abspath(__file__))

    # Construct the path to the theme file
    theme_path = os.path.join(base_path, 'customtkinter', 'assets', 'themes', 'blue.json')

except ModuleNotFoundError as err:
    print('Opssss... Looks like there is an error importing the package', err)


class MyCheckboxFrame(customtkinter.CTkFrame):
    """My checkbox frame"""
    def __init__(self, master, title, values):
        """Init function"""
        super().__init__(master)

        self.grid_columnconfigure(0, weight=1)
        """But the values of the checkboxes in the MyCheckboxFrame are hardcoded in the code right now"""
        # To make the MyCheckboxFrame class more dynamically usable, we pass a list of string values to the MyCheckboxFrame, which will be the text values of the checkboxes in the frame. Now the number of checkboxes is also arbitrary

        self.values = values
        #--------- TITLES --------------
        # Note that column 0 inside of the frame is now configure to have a weight of 1, so that the label spans the whole frame with its sticky value of 'ew'. 
        # For the CTkLabel we passed an fg_color and corner_radius argument because the label is 'transparent' by default and has a corner_radius of 0. 
        # ALso note that the grid row position is now i+1 because of the title label in the first row.
        self.title = title

        self.checkboxes = []

        self.title = customtkinter.CTkLabel(self, text=self.title,
                                            fg_color="gray40", corner_radius=6, height = 50)
        self.title.grid(row=0, column=0, padx=10, pady=(10, 0), sticky="ew")

        for i, value in enumerate(self.values):
            checkbox = customtkinter.CTkCheckBox(self, text=value)
            checkbox.grid(row=i+1, column=0, padx=10, pady=(10, 0), sticky="w")
            self.checkboxes.append(checkbox)
    # ------------ GET -----------
    # Este metodo lo utilizo para poder pasar los valores de los check boxes a la aplicacion APP
    # En el boton de submit llamo con el objeto SELF al 'check frame' y el metodo GET()
    def get(self):

        checked_checkboxes = []

        # Aqui lo que hago es adjuntar el texto de cada checkbox
        for checkbox in self.checkboxes:
            if checkbox.get() == 1:
                checked_checkboxes.append(checkbox.cget("text"))
        return checked_checkboxes

        # -------- CLASE PRINCIPAL ---------------

class App(customtkinter.CTk):

    def __init__(self):
        
        """Init function"""
        super().__init__()

        self.title("Reports App")
        self.geometry("600x320")

        # Width de las columans se maneja aqui.
        # Add weight to the columns to distribute them equally
        self.grid_columnconfigure(0, weight=1)
        self.grid_columnconfigure(1, weight=1)

        self.grid_rowconfigure(0, weight=1)

        self.configure(bg_color="white")
        # ----------SWITCH ---------------
        switch_var = customtkinter.StringVar(value="on")

        self.switch_1 = customtkinter.CTkSwitch(self,text=" ZIP files.",
                                                variable=switch_var, onvalue="on", offvalue="off")

        self.switch_1.grid(row=0, column=0, padx=10,pady=(10, 0), sticky="nsew")
        # ----------IMAGE ---------------
        # Create an object of tkinter ImageTk


        # ----------FRAME1 - WEBSTATS---------------
        self.checkbox_frame_1 = MyCheckboxFrame(self, "Webstats",
                                                values=["Journals", "Books", "Databases"])

        self.checkbox_frame_1.grid(row=1, column=0, padx=10, pady=(10, 0), sticky="nsew")

        # ----------FRAME2 - WEBSTATS CONSORTIA---------------
        self.checkbox_frame_2 = MyCheckboxFrame(self, "Webstats - CONSORTIA -",
                                                values=["Journals - consortia",
                                                "Books - consortia", "Databases - consortia"])

        self.checkbox_frame_2.grid(row=1, column=1, padx=10, pady=(10, 0), sticky="nsew")

        # ----------BUTTONS---------------
        self.button = customtkinter.CTkButton(self, text="Create",command=self.button_callback)
        self.button.grid(row=7, column=0, padx=10, pady=10)

        self.button2 = customtkinter.CTkButton(self, text="Quit",command=self.quit,
                                                fg_color='Dark Red', hover_color= "Dark Gray")
        self.button2.grid(row=7, column=1, padx=10, pady=10)



        self.product_title = ''
    # ----------- BOTON ------------
    def button_callback(self):
        """Call back function"""
        # Esto es para que la ventana quede de frente
        self.lower()

        # Asi llamamos al metodo que nos regresa los valores de los check boxes
        # Aqui es donde voy a definir que hacer.
        webstats_one = self.checkbox_frame_1.get()

        webstats_consortia = self.checkbox_frame_2.get()

        switch = self.switch_1.get()

        # Es solo para averiguar que me viene por la
        #print('----------------------------- SWITCH-------------------', switch)
        try:
            ## -------------  ONLY ONE TYPE OF REPORT AT THE TIME --------------------
            if len(webstats_one) >= 1 and len(webstats_consortia) >= 1:
                # En caso de que no haya seleccionado ningun reporte
                raise Exception("You can only select one report type at the time.\n\n It looks like you have currentlly selected Webstats and also Consortia reports.")
            # Verifico cunatos elementos me vienen en el array
            # SOlo puede venir uno           
            if len(webstats_one) == 1 and len(webstats_consortia) == 0:
                ## Utilizo el 1 para reportes normales
                self.t = 1
                ## -------------  IF WEBSTATS --------------------
                if len(webstats_one) == 1:
                    for i in webstats_one:
                        if i == 'Journals':
                            product_title = i
                            self.open_toplevel(product_title, switch, self.t)

                        elif i == 'Books':
                            product_title = i
                            self.open_toplevel(product_title, switch, self.t)

                        elif i == 'Databases':
                            product_title = i
                            self.open_toplevel(product_title, switch, self.t)

            elif len(webstats_one) == 0 and len(webstats_consortia) == 1:

                # Utilizo el 2 para reportes consortia
                self.t = 2
                ## -------------  IF CONSORTIA --------------------
                if len(webstats_consortia) == 1:

                    for i in webstats_consortia:

                        if i == 'Journals - consortia':
                            product_title = i
                            self.open_toplevel(product_title, switch, self.t)

                        elif i == 'Books - consortia':
                            product_title = i
                            self.open_toplevel(product_title, switch, self.t)

                        elif i == 'Databases - consortia':
                            product_title = i
                            self.open_toplevel(product_title, switch, self.t)

            else:
             # En caso de que no haya seleccionado ningun reporte
                raise Exception("You haven´t selected any report at all or you have selected more than one at the time.\n\n  Your current selection: \n\n - {} Webstats \n\n - {} Consortia Reports.".format(len(webstats_one), len(webstats_consortia)))


        except Exception as e:
            # Este es el nombre de la exception
            exception_type = type(e).__name__

            messagebox.showerror(exception_type,f' This is a - {exception_type}: \n\n{str(e)}\n\n In case you need further assitance please contact support: mauro.cespedes@wolterskluwer.com.')
    #---------- TOP LEVEL WINDOW -------------------
    def open_toplevel(self, product_title, switch, type):
        """Top level window"""
        product_type = product_title

        switch1 = switch

        type = type
        # ------WEBSTATS---------------
        if type == 1:
            # create window if its None or destroyed
            toplevel_window = ToplevelWindow(product_type, switch1)

            toplevel_window.attributes("-topmost", 1)
            toplevel_window.lift()

        # ------CONSORTIA---------------
        else:

            MembersNumber(product_type, switch1)
    #---------- ALERT -------------------------
    def show_alert(self):
        """Show Alerts"""
        message = "This is an alert message!"
        alert_window = AlertWindow(message)
        alert_window.lift()
        alert_window.attributes("-topmost", True)

class ToplevelWindow(customtkinter.CTkToplevel):
    """CLASE TOP LEVEL WEBSTATS"""
    """Aqui tengo que pasar el product tile que obtengo de la funcion que llama esta clase.
    # Necesito agregar la variable al constructor para poder utilizarla."""

    def __init__(self, product_type, switch1, *args, **kwargs):
        super().__init__(*args, **kwargs)
        # Ahora ya puedo utilizar la variabel para personalizar cada ventana que se abra.
        self.title(product_type)

        self.switch = switch1
        self.geometry("400x200")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        # ----------FRAME---------------
        self.browse_button = customtkinter.CTkButton(self, text="Browse Files",
                                                    command=self.browse_files)

        self.browse_button.grid(row=2, column=0, padx=10, pady=10, sticky="w")
        # Esto es para guardar el nombre del archivo que seleccionamos.
        self.file_label = customtkinter.CTkLabel(self, text="Select your file:",
                                                fg_color="black", wraplength=300)
        self.file_label.grid(row=1, column=0, columnspan = 2, padx=10, pady=10, sticky="w")
        #--------- BUTTONS --------------
        ## COn el load wiith data lo que hago es pasar lo parametros a la nesva ventana.
        self.button = customtkinter.CTkButton(self, text="Create", command=self.open_loading_window)
        self.button.grid(row=3, column=0, padx=10, pady=10, sticky="ew")
        ## COn el destroy quito la ventana en la que estoy sin destruir las que estan atras.
        self.quit_button = customtkinter.CTkButton(self, text="Quit", command=self.destroy)
        self.quit_button.grid(row=3, column=1, padx=10, pady=10, sticky="ew")
    # ----------- BOTON ------------
    def button_callback(self):
    # Asi llamamos al metodo que nos regresa los valores de los check boxes
    # Aqui es donde voy a definir que hacer.
        print("checkbox_frame_1:", self.checkbox_frame_1.get())
        print("checkbox_frame_2:", self.checkbox_frame_2.get())

    # Function for opening the
    # file explorer window
    def browse_files(self):
        try:
            #----------  SEND WINDOWS TO THE BACK
            self.lower()
            # Use the filedialog.askopenfilename() function to open the file browser dialog
            self.file_path = filedialog.askopenfilename()
            current_directory = os.getcwd()
            ## ------------ ZIP ON ----------------------
            if self.switch == 'on':
                ##-------ZIP FILES ------------------
                with zipfile.ZipFile(self.file_path, 'r') as zip_ref:

                    for file_name in zip_ref.namelist():
                        if file_name.endswith('.xlsx'):
                            zip_ref.extract(file_name, current_directory)
                            ## Convierto el self path en la ruta del archivo descomprimido.
                            self.file_path = file_name
            # Display the selected file path in a label or perform any other desired action
            self.file_label.configure(text=self.file_path)
            self.lift()

        except zipfile.BadZipFile:
            messagebox.showerror('Invalid ZIP file',f' This is a invalid ZIP file \n\n The file you submited could not be opened.. \n\n Please make sure to check the ZIP option if you want to use ZIP files.')
            LoadingWindow.destroy(self)

        except FileNotFoundError:
            messagebox.showerror('File Not Found ',f' The file you submited could not be found.. \n\n Please make sure to check the ZIP option if you want to use ZIP files.')
            LoadingWindow.destroy(self)


    def open_loading_window(self):

        title = self.title()
        browse_button_result = self.file_path
        #---------- SEND TO THE BACK --------------
        self.lower()

        # Create and display the loading window
        # He creado un aclase nueva par la consortia.
        LoadingWindow(title, browse_button_result)
        #loading_window.wait_window()
        self.destroy()

class ToplevelWindowConsortia(customtkinter.CTkToplevel):

    #  Aqui tengo que pasar el product tile que obtengo de la funcion que llama esta clase.
    # Necesito agregar la variable al constructor para poder utilizarla.
    def __init__(self, product_type,  switch1, res, *args, **kwargs):
        super().__init__(*args, **kwargs)

        # Ahora ya puedo utilizar la variabel para personalizar cada ventana que se abra.
        self.title(product_type)

        self.switch = switch1

        # Esta variable me guarda el numero de flidialogs que quiero abrir.
        self.memb = res

        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)

        self.attributes("-topmost", 1)


        my_frame = customtkinter.CTkScrollableFrame(self, width=400,
                                                    height=400, corner_radius=0,
                                                    fg_color="transparent")
                      
        my_frame.grid(row=0, column=0, columnspan = 3, sticky="nsew")

        # ----------FRAME---------------

        # Esto es para que podamos almacenar los labels
        self.file_labels = []  # Store the labels in a list



        for i in range(int(self.memb)):


            self.browse_button = customtkinter.CTkButton(my_frame, text=f"File{i}",
                                                        command=lambda i=i: self.browse_files(i))

            self.browse_button.grid(row=i, column=1, columnspan = 3, padx=10, pady=10, sticky="ew")


            self.file_labels.append(customtkinter.CTkLabel(my_frame,
                                                            text="Select your file:",fg_color="black", wraplength=300))

            # Incluimos el resultado del path en cada label.
            self.file_labels[i].grid(row=i, column=0, padx=10, pady=10, sticky="ew")
             # Display the selected file path in a label or perform any other desired action



        #--------- BUTTONS --------------
        ## COn el load wiith data lo que hago es pasar lo parametros a la nesva ventana.
        self.button = customtkinter.CTkButton(my_frame, text="Create", command=self.button_callback)
        self.button.grid(row=self.memb+2, column=0, padx=10, pady=10, sticky="ew")

        ## COn el destroy quito la ventana en la que estoy sin destruir las que estan atras.
        self.quit_button = customtkinter.CTkButton(my_frame, text="Quit",
                                                    command=self.destroy, fg_color='Dark Red', hover_color= "Dark Gray")

        self.quit_button.grid(row=self.memb+2, column=1, columnspan = 3, padx=10, pady=10, sticky="ew")



    def button_callback(self):
        file_paths = [label.cget("text") for label in self.file_labels]
        self.open_loading_window(file_paths)



    def browse_files(self, index):
        try:
            #----------  SEND WINDOWS TO THE BACK
            # Send window to the back
            self.lower()
            #self.file_label.configure(text=self.browse_button.get())
        

            file_labels = []  # Store the labels in a list
            # ------------ UPLOAD MILTIPLE FILES

            # Ask for file path
            file_paths = filedialog.askopenfilenames()

            for i, file_path in enumerate(file_paths):

                ## ------------ ZIP ON ----------------------
                if self.switch == 'on'and file_path.endswith('.zip'):
                    current_directory = os.getcwd()

                    ##-------ZIP FILES ------------------
                    with zipfile.ZipFile(file_path, 'r') as zip_ref:

                        for file_name in zip_ref.namelist():
                            if file_name.endswith('.xlsx'):
                                zip_ref.extract(file_name, current_directory)
                                ## Convierto el self path en la ruta del archivo descomprimido.
                                file_path = file_name

                self.file_labels[index].configure(text=file_path)
                self.file_labels[index].update()


            self.lift()
   
        except zipfile.BadZipFile:
            messagebox.showerror('Invalid ZIP file',f' This is a invalid ZIP file \n\n The file you submited could not be opened.. \n\n Please make sure to check the ZIP option if you want to use ZIP files.')
            LoadingWindow.destroy(self)

        except FileNotFoundError:
            messagebox.showerror('File Not Found ',f' The file you submited could not be found.. \n\n Please make sure to check the ZIP option if you want to use ZIP files.')
            LoadingWindow.destroy(self)


    def open_loading_window(self, file_paths):

        title = self.title()
        browse_button_results = file_paths

        #---------- SEND TO THE BACK --------------
        self.lower()

        # Create and display the loading window
        # He creado un aclase nueva par la consortia.
        LoadingWindowConsortia(title, browse_button_results)
        #loading_window.wait_window()
        self.destroy()

class MembersNumber(customtkinter.CTkToplevel):

    #  Aqui tengo que pasar el product tile que obtengo de la funcion que llama esta clase.
    # Necesito agregar la variable al constructor para poder utilizarla.
    def __init__(self, product_type, switch1, *args, **kwargs):
        super().__init__(*args, **kwargs)

        self.geometry("300x200")
        self.grid_columnconfigure(0, weight=1)
        self.grid_rowconfigure(0, weight=1)
        self.title('Select Consortia Members')

        self.attributes("-topmost", 1)
        self.lift()

        self.pt = product_type
        self.sw = switch1

        self.input_label = customtkinter.CTkLabel(self, text="Enter number of customers:")
        self.input_label.pack(padx=10, pady=10)
        self.input_entry = customtkinter.CTkEntry(self)
        self.input_entry.pack(padx=10, pady=10)
        self.ok_button = customtkinter.CTkButton(self, text="OK", command=self.ok_button_callback)
        self.ok_button.pack(side=tk.LEFT, padx=10, pady=10)
        self.cancel_button = customtkinter.CTkButton(self, text="Cancel", command=self.cancel_button_callback)
        self.cancel_button.pack(side=tk.LEFT, padx=10, pady=10)


    # ----------- BOTON ------------
    def ok_button_callback(self):

        res = int(self.input_entry.get())
        self.destroy()
        # create window if its None or destroyed
        ToplevelWindowConsortia(self.pt, self.sw, res)


    def cancel_button_callback(self):
        self.destroy()

# Create a custom top-level window for the alert message
class AlertWindow(CTkToplevel):
    def __init__(self, message):
        super().__init__()

        self.title("Alert")
        self.geometry("300x150")

        self.label = CTkLabel(self, text=message)
        self.label.pack(padx=20, pady=20)

        self.ok_button = CTkButton(self, text="OK", command=self.destroy)
        self.ok_button.pack(pady=10)

class LoadingWindow(customtkinter.CTkToplevel):
    def __init__(self, title, browse_button_result):
        super().__init__()

        self.geometry("300x300")

        # Use wm_attributes from Tkinter to set the window attributes
        self.wm_attributes("-topmost", True)
        # Use the underlying Tkinter window to raise the CustomTkinter window
        #self.tk.call('wm', 'attributes', '.', '-topmost', True)



        # Add widgets or modify the content of the window as needed
        style = ttk.Style(self)
        style.configure("Custom.Horizontal.TProgressbar", thickness=80)

        # Create a progress bar
        self.progress_bar = ttk.Progressbar(self, orient="horizontal", mode="indeterminate", style="Custom.Horizontal.TProgressbar", length=200)
        self.progress_bar.pack(pady=10)
        # Start the progress bar animation
        self.progress_bar.start()
  
        # Center the loading window on the screen
        LoadingWindow.update_idletasks(self)


        # Identifica el directorio en donde vive nuestra aplicacion
        # Es decir el directorio desde dond ecorro el main
        working_directory = os.getcwd()


        # Create a separate thread for the loading process
        loading_thread = threading.Thread(target=LoadingWindow.execute_loading_process, args=(self, title, browse_button_result, working_directory))
        loading_thread.start()

     

        #------------ EXCETIONS HANDLEING

        # Voy a generar un pop up para que le indique al usuario el tipo de error
        # Para no estrar dentro de las clases espero el error aqui y puclico un pop up


        # Method to execute the loading process
    def execute_loading_process(self, title, browse_button_result, working_directory):      
        try:

            if title == 'Journals':

                journal = Journal(browse_button_result, working_directory)
                journal.new_column()
                journal.save()
                journal.clean_folder()

                # Para que destruya todos los pop ups
                LoadingWindow.destroy(self)

                # Ask the user if they want to open the file
                answer = messagebox.askyesno("READY!!!", "Do you want to open this file?")

                if answer == True:
                    journal.load()
                    self.destroy()
                    LoadingWindow.destroy(self)
                else:
                    print(answer)
                    self.destroy()
                    LoadingWindow.destroy(self)
            elif title == 'Books':
                book = Book(browse_button_result, working_directory)
                book.new_column()
                book.save()
                book.clean_folder()
                LoadingWindow.destroy(self)

                # Ask the user if they want to open the file
                answer = messagebox.askyesno("READY!!!", "Do you want to open this file?")
                if answer == True:
                    book.load()
                    self.destroy()
                    LoadingWindow.destroy(self)
                else:
                    print(answer)
                    self.destroy()
                    LoadingWindow.destroy(self)

            elif title == 'Databases':
                db = Database(browse_button_result, working_directory)
                db.new_column()
                db.save_db()
                db.clean_folder()
                LoadingWindow.destroy(self)

                # Ask the user if they want to open the file
                answer = messagebox.askyesno("READY!!!", "Do you want to open this file?")
                if answer == True:
                    db.load('db')
                    self.destroy()
                    LoadingWindow.destroy(self)
                else:
                    print(answer)
                    self.destroy()
                    LoadingWindow.destroy(self)
            
            #----- IF NO SELECTION ------------------------
            else:
                print('Error no me llego ningun nombre')
     
                answer = messagebox.askyesno("Please select a report type.",  title="Ok")

                if answer=="OK":
                    print ("You pressed OK")
                LoadingWindow.destroy(self)        
     
        except Exception as e:
            # Este es el nombre de la exception
            exception_type = type(e).__name__

            # Por cada tipo de exception hago un mensaje diferente.
            if isinstance(e, ValueError):

                messagebox.showerror(exception_type, f' This is a -  {exception_type} -  exception.\n\n You are trying up upload an Excel file with incorrect columns for the report requested.\n\n Usually this happens when the type of report and the file uploaded dont match.')
                LoadingWindow.destroy(self)

            elif isinstance(e, FileNotFoundError):
                messagebox.showerror(exception_type,f' This is a - {exception_type} - exception \n\n The file you submited could not be found or no file has been selected.. \n\n Please make sure to select a valid file taken from Webstats, that matches the type of report selected.')
                LoadingWindow.destroy(self)

            elif isinstance(e, KeyError):
                messagebox.showerror(exception_type,f' This is a - {exception_type} - exception: \n\n The file you selected doesn´t have the : - {str(e)} - column. \n\n Usually this happens when the type of report and the file uploaded dont match. \n\n Please make sure to select a file taken from Webstats, that matches the type of report selected.')
                LoadingWindow.destroy(self)
            else:
                messagebox.showerror(exception_type,f' This is a - {exception_type} - exception: \n\n {str(e)} \n\n Please contact support.')
                LoadingWindow.destroy(self)    

#--------- LOADING CLASE CONSORTIA --------------
class LoadingWindowConsortia(customtkinter.CTkToplevel):
    def __init__(self, title, browse_button_results):
        super().__init__()
        self.ttitle = title
        self.the_paths = browse_button_results

        for path in browse_button_results:
            print(path)

        self.geometry("300x300")
        # Use wm_attributes from Tkinter to set the window attributes       
        self.wm_attributes("-topmost", True)

        # Add widgets or modify the content of the window as needed
        style = ttk.Style(self)
        style.configure("Custom.Horizontal.TProgressbar", thickness=80)

        # Create a progress bar
        self.progress_bar = ttk.Progressbar(self, orient="horizontal", mode="indeterminate", style="Custom.Horizontal.TProgressbar", length=200)
        self.progress_bar.pack(pady=10)
       # Start the progress bar animation
        self.progress_bar.start()

        self.loading_done = False
        self.cpu_label = ttk.Label(self, text="CPU: ")
        self.cpu_label.pack()
        self.ram_label = ttk.Label(self, text="RAM: ")
        self.ram_label.pack()
        # Center the loading window on the screen
        LoadingWindow.update_idletasks(self)

        # Identifica el directorio en donde vive nuestra aplicacion
        # Es decir el directorio desde dond ecorro el main
        working_directory = os.getcwd()

        # Create a separate thread for the loading process
        loading_thread = threading.Thread(target=LoadingWindowConsortia.execute_loading_process, args=(self, working_directory))
        loading_thread.start()


        #------------ EXCETIONS HANDLEING

        # Voy a generar un pop up para que le indique al usuario el tipo de error
        # Para no estrar dentro de las clases espero el error aqui y puclico un pop up


        # Method to execute the loading process
    def execute_loading_process(self, working_directory):
        try:
      
            if self.ttitle == 'Journals - consortia':
                journal = ReportC(self.the_paths)

                #journal = ReportC([i for i in browse_button_results], i)

                # Para que destruya todos los pop ups
                LoadingWindow.destroy(self)

                # Ask the user if they want to open the file
                answer = messagebox.askyesno("READY!!!", "Do you want to open this file?")

                if answer == True:
                    journal.load()
                    self.destroy()
                    LoadingWindow.destroy(self)
                else:
                    print(answer)
                    self.destroy()
                    LoadingWindow.destroy(self)

            elif self.ttitle == 'Books - consortia':
                book = ReportC(self.the_paths) 
                book.new_column()
                book.save()
                book.clean_folder()
                LoadingWindow.destroy(self)

                # Ask the user if they want to open the file
                answer = messagebox.askyesno("READY!!!", "Do you want to open this file?")
                if answer == True:
                    book.load()
                    self.destroy()
                    LoadingWindow.destroy(self)
                else:
                    print(answer)
                    self.destroy()
                    LoadingWindow.destroy(self)

            elif self.ttitle == 'Databases - consortia':
                db = ReportC(self.the_paths)
                db.new_column()
                db.save_db()
                db.clean_folder()
                LoadingWindow.destroy(self)

                # Ask the user if they want to open the file
                answer = messagebox.askyesno("READY!!!", "Do you want to open this file?")
                if answer == True:
                    db.load('db')
                    self.destroy()
                    LoadingWindow.destroy(self)
                else:
                    print(answer)
                    self.destroy()
                    LoadingWindow.destroy(self)

            else:
                print('Error no me llego ningun nombre')
   
                answer = messagebox.askyesno("Please select a report type.",  title="Ok")

                if answer=="OK":
                    print ("You pressed OK")
                LoadingWindow.destroy(self)   

                if self.loading_done:
                    return


        except Exception as e:

            exception_type = type(e).__name__
            if isinstance(e, ValueError):
           
                messagebox.showerror(exception_type, f' This is a -  {exception_type} -  exception.\n\n You are trying up upload an Excel file with incorrect columns for the report requested.\n\n Usually this happens when the type of report and the file uploaded dont match.')
                LoadingWindow.destroy(self)

            elif isinstance(e, FileNotFoundError):
                messagebox.showerror(exception_type,f' This is a - {exception_type} - exception \n\n The file you submited could not be found or no file has been selected.. \n\n Please make sure to select a valid file taken from Webstats, that matches the type of report selected.')
                LoadingWindow.destroy(self)

            elif isinstance(e, KeyError):
                messagebox.showerror(exception_type,f' This is a - {exception_type} - exception: \n\n The file you selected doesn´t have the : - {str(e)} - column. \n\n Usually this happens when the type of report and the file uploaded dont match. \n\n Please make sure to select a file taken from Webstats, that matches the type of report selected.')
                LoadingWindow.destroy(self)
            else:
                messagebox.showerror(exception_type,f' This is a - {exception_type} - exception: \n\n {str(e)} \n\n Please contact support.')
                LoadingWindow.destroy(self)



## When We upload the documentation we need to make sure to remove temporalitly these lineas

app = App()
app.mainloop()