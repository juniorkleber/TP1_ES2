from tkinter import Tk, Button, Entry, Label, ttk, Toplevel, Menu, messagebox, StringVar
import tkinter as tk

import pyodbc
import csv
import smtplib
from email.message import EmailMessage
import random
import string

# Variável global para armazenar o nome do usuário logado
logged_in_user = None

def connect():
    dadosConexao = ("Driver={SQLite3 ODBC Driver};Server=localhost;Database=Projeto.db")
    connection = pyodbc.connect(dadosConexao)
    return connection

class ProductInterface:
    def __init__(self, main_window):
        self.main_window = main_window
        self.setup_ui()
        self.products = []  # Adicionando um atributo para armazenar os produtos

    def setup_ui(self):
        self.create_treeview()  # Cria a visualização em árvore para exibir os produtos

    def create_treeview(self):
        self.treeview = ttk.Treeview(self.main_window, columns=("ID", "Nome", "Descricao", "Preco"), show="headings")
        
        self.treeview.heading("ID", text="ID")
        self.treeview.heading("Nome", text="Product Name")
        self.treeview.heading("Descricao", text="Product Description")
        self.treeview.heading("Preco", text="Product Price")
        
        self.treeview.column("#0", width=0, stretch="NO")
        self.treeview.column("ID", width=100)
        self.treeview.column("Nome", width=300)
        self.treeview.column("Descricao", width=500)
        self.treeview.column("Preco", width=200)
        
        self.treeview.grid(row=3, column=0, columnspan=10, sticky="NSEW")

        self.treeview.bind("<Double-1>", self.edit_product)

    def calculate_price_statistics(self):
        if not self.products:
            messagebox.showinfo("No Products", "No products available to calculate statistics.")
            return

        prices = [product[3] for product in self.products]

        average_price = sum(prices) / len(prices)
        max_price = max(prices)
        min_price = min(prices)

        messagebox.showinfo("Price Statistics", f"Average Price: {average_price:.2f}\nMaximum Price: {max_price}\nMinimum Price: {min_price}")

    def sort_products_by_price(self, order='asc'):
        if not self.products:
            messagebox.showinfo("No Products", "No products available to sort.")
            return

        reverse_order = False
        if order == 'desc':
            reverse_order = True

        self.products.sort(key=lambda x: x[3], reverse=reverse_order)
        self.list_products()  # Atualiza a Treeview com a lista de produtos ordenada

    def display_expensive_cheap(self):
        # Encontrar o produto mais caro e mais barato
        connection = connect()
        cursor = connection.cursor()

        cursor.execute("SELECT * FROM Produtos ORDER BY Preco DESC LIMIT 1")
        most_expensive = cursor.fetchone()

        cursor.execute("SELECT * FROM Produtos ORDER BY Preco ASC LIMIT 1")
        cheapest = cursor.fetchone()

        # Exibir informações na interface
        most_expensive_name = most_expensive[1] if most_expensive else "N/A"
        cheapest_name = cheapest[1] if cheapest else "N/A"

        # Criar rótulos para exibir o produto mais caro e mais barato
        self.most_expensive_label = Label(self.main_window, text=f"Product Most Expensive: {most_expensive_name}", font=("Arial", 12), fg="#4fb804")
        self.most_expensive_label.grid(row=0, column=11, padx=10, pady=10, sticky="NSEW")

        self.cheapest_label = Label(self.main_window, text=f"Product Cheapest: {cheapest_name}", font=("Arial", 12), fg="#b00202")
        self.cheapest_label.grid(row=2, column=11, padx=10, pady=10, sticky="NSEW")

    def clear_treeview(self):
        for i in self.treeview.get_children():
            self.delete_product_treeview()

    def list_products(self):
        for item in self.treeview.get_children():
            self.treeview.delete(item)

        connection = connect()
        cursor = connection.cursor()
        cursor.execute("SELECT * FROM Produtos")

        products = cursor.fetchall()

        for product in products:
            self.treeview.insert("", "end", values=(product[0], product[1], product[2], product[3]))

        cursor.close()

    def register_new_product(self):
        register_product_window = Toplevel(self.main_window)
        register_product_window.title("Register New Product")
        register_product_window.configure(bg="#eeeeee")

        width_window = 450
        height_window = 230

        width_screen = register_product_window.winfo_screenwidth()
        height_screen = register_product_window.winfo_screenheight()

        pos_x = (width_screen // 2) - (width_window // 2)
        pos_y = (height_screen // 2) - (height_window // 2)

        register_product_window.geometry('{}x{}+{}+{}'.format(width_window, height_window, pos_x, pos_y))

        border_style = {"borderwidth": 2, "relief": "groove"}

        Label(register_product_window, text="Name", font=("Arial", 12), bg="#eeeeee").grid(row=0, column=0, padx=10, pady=10, sticky="w")
        product_name_register = Entry(register_product_window, font=("Arial", 14), **border_style, bg="#eeeeee")
        product_name_register.grid(row=0, column=1, padx=10, pady=10)

        Label(register_product_window, text="Description", font=("Arial", 12), bg="#eeeeee").grid(row=1, column=0, padx=10, pady=10, sticky="w")
        product_description_register = Entry(register_product_window, font=("Arial", 14), **border_style, bg="#eeeeee")
        product_description_register.grid(row=1, column=1, padx=10, pady=10)

        Label(register_product_window, text="Price", font=("Arial", 12), bg="#eeeeee").grid(row=2, column=0, padx=10, pady=10, sticky="w")
        product_price_register = Entry(register_product_window, font=("Arial", 14), **border_style, bg="#eeeeee")
        product_price_register.grid(row=2, column=1, padx=10, pady=10, columnspan=2)

        for i in range(5):
            register_product_window.grid_rowconfigure(i, weight=1)

        for i in range(2):
            register_product_window.grid_columnconfigure(i, weight=1)

        def saveData():
            register_new_product_ = (product_name_register.get(), product_description_register.get(), product_price_register.get())

            connection = connect()

            cursor = connection.cursor()

            cursor.execute("INSERT INTO Produtos (Nome, Descricao, Preco) Values (?,?,?)", register_new_product_)
            connection.commit()

            print("Product registered successfully!")

            register_product_window.destroy()

            self.list_products()  # Chama a função list_products da classe ProductInterface

        btn_save_product = Button(register_product_window, text="Save", font=("Arial", 14), bg="#008000", fg="#ffffff", command=saveData)
        btn_save_product.grid(row=3,column=0, columnspan=2,padx=10,pady=10, sticky="NSEW")

        btn_cancel_product = Button(register_product_window, text="Cancel", font=("Arial", 14), bg="#FF0000", fg="#ffffff", command=register_product_window.destroy)
        btn_cancel_product.grid(row=4,column=0, columnspan=2,padx=10,pady=10, sticky="NSEW")

    def edit_product(self, event):
        selected_item = self.treeview.selection()[0]

        select_values = self.treeview.item(selected_item)['values']

        edit_product_window = Toplevel(self.main_window)
        edit_product_window.title("Edit Product")
        edit_product_window.configure(bg="#eeeeee")

        width_window = 500
        height_window = 200

        width_screen = edit_product_window.winfo_screenwidth()
        height_screen = edit_product_window.winfo_screenheight()

        pos_x = (width_screen // 2) - (width_window // 2)
        pos_y = (height_screen // 2) - (height_window // 2)

        edit_product_window.geometry('{}x{}+{}+{}'.format(width_window, height_window, pos_x, pos_y))

        border_style = {"borderwidth": 2, "relief": "groove"}

        Label(edit_product_window, text="Product Name", font=("Arial", 16), bg="#eeeeee").grid(row=0, column=0, padx=10, pady=10, sticky="w")
        product_name_edit = Entry(edit_product_window, font=("Arial", 16), **border_style, bg="#eeeeee", textvariable=StringVar(value=select_values[1]))
        product_name_edit.grid(row=0, column=1, padx=10, pady=10)

        Label(edit_product_window, text="Product Description", font=("Arial", 16), bg="#eeeeee").grid(row=1, column=0, padx=10, pady=10, sticky="w")
        product_description_edit = Entry(edit_product_window, font=("Arial", 16), **border_style, bg="#eeeeee", textvariable=StringVar(value=select_values[2]))
        product_description_edit.grid(row=1, column=1, padx=10, pady=10)

        Label(edit_product_window, text="Product Price", font=("Arial", 16), bg="#f5f5f5").grid(row=2, column=0, padx=10, pady=10, sticky="w")
        product_price_edit = Entry(edit_product_window, font=("Arial", 16), **border_style, bg="#f5f5f5", textvariable=StringVar(value=select_values[3]))
        product_price_edit.grid(row=2, column=1, padx=10, pady=10)

        for i in range(5):
            edit_product_window.grid_rowconfigure(i, weight=1)

        for i in range(2):
            edit_product_window.grid_columnconfigure(i, weight=1)
        
        def saveEdit():
            product = product_name_edit.get()
            new_description = product_description_edit.get()
            new_price = product_price_edit.get()

            self.treeview.item(selected_item, values=(select_values[0], product, new_description, new_price))

            connection = connect()

            cursor = connection.cursor()

            cursor.execute("UPDATE Produtos SET Nome = ?, Descricao = ?, Preco = ? WHERE ID = ?", (product, new_description, new_price, select_values[0]))

            connection.commit()

            print("Data Registered Successfully!")

            edit_product_window.destroy()

            self.list_products()

        btn_save_edit = Button(edit_product_window, text="Confirm", font=("Arial", 14), bg="#008000", fg="#ffffff", command=saveEdit)
        btn_save_edit.grid(row=4,column=1, padx=20,pady=20)
        btn_cancel_edit = Button(edit_product_window, text="Cancel", font=("Arial", 14), bg="#FF0000", fg="#ffffff", command=edit_product_window.destroy)
        btn_cancel_edit.grid(row=4,column=0, padx=20,pady=20)

    def delete_product_treeview(self):
        selected_item = self.treeview.selection()
        
        if selected_item:
            product_id = self.treeview.item(selected_item[0])['values'][0]
            
            self.treeview.delete(selected_item)
            
            connection = connect()

            cursor = connection.cursor()
            cursor.execute("DELETE FROM Produtos WHERE ID = ?", (product_id,))
            connection.commit()

        else:
            messagebox.showerror("Error", "Select a product to delete.")

    def filter_data(self, product_name, product_description):
        
        connection = connect()
        cursor = connection.cursor()

        if not product_name.get() and not product_description.get():
            self.list_products()
            return
        
        sql = "SELECT * FROM Produtos"
        params = []
        if product_name.get():
            sql += " WHERE Nome LIKE ?"
            params.append('%' + product_name.get() + '%') 
        
        if product_description.get():
            if product_name.get():
                sql += " AND "
            else:
                sql += " WHERE"
            sql += " Descricao LIKE ?"
            params.append('%' + product_description.get() + '%')

        cursor.execute(sql, tuple(params))
        product = cursor.fetchall()
        
        for i in self.treeview.get_children():
            self.treeview.delete(i)
            
        for data in product:
            self.treeview.insert('', 'end', values=(data[0], data[1],data[2],data[3]))

    def show_product_history(self):
        product_history_window = Toplevel(self.main_window)
        product_history_window.title("Product History")
        product_history_window.configure(bg="#eeeeee")

        width_window = 600
        height_window = 400

        width_screen = product_history_window.winfo_screenwidth()
        height_screen = product_history_window.winfo_screenheight()

        pos_x = (width_screen // 2) - (width_window // 2)
        pos_y = (height_screen // 2) - (height_window // 2)

        product_history_window.geometry('{}x{}+{}+{}'.format(width_window, height_window, pos_x, pos_y))

        product_history_treeview = ttk.Treeview(product_history_window, columns=("ID", "Nome", "Descricao", "Preco"), show="headings")

        product_history_treeview.heading("ID", text="ID")
        product_history_treeview.heading("Nome", text="Product Name")
        product_history_treeview.heading("Descricao", text="Product Description")
        product_history_treeview.heading("Preco", text="Product Price")

        product_history_treeview.column("#0", width=0, stretch="NO")
        product_history_treeview.column("ID", width=100)
        product_history_treeview.column("Nome", width=200)
        product_history_treeview.column("Descricao", width=300)
        product_history_treeview.column("Preco", width=100)

        product_history_treeview.grid(row=0, column=0, columnspan=10, sticky="NSEW")

        product_history_window.grid_rowconfigure(0, weight=1)  # Define a primeira linha para expandir
        product_history_window.grid_columnconfigure(0, weight=1)  # Define a primeira coluna para expandir

        connection = connect()
        cursor = connection.cursor()
        cursor.execute("SELECT * FROM Produtos")

        products = cursor.fetchall()

        for product in products:
            product_history_treeview.insert("", "end", values=(product[0], product[1], product[2], product[3]))

        cursor.close()

def recuperar_senha(email):
    nova_senha = ''.join(random.choices(string.ascii_letters + string.digits, k=10))
    
    server = smtplib.SMTP('smtp.gmail.com', 587)  # Atualize com as configurações do seu servidor SMTP
    server.starttls()
    server.login("seuemail@gmail.com", "suasenha")  # credenciais de e-mail

    msg = EmailMessage()
    msg.set_content(f"Sua nova senha é: {nova_senha}")

    msg['Subject'] = "Recuperação de Senha"
    msg['From'] = "seuemail@gmail.com"
    msg['To'] = email

    server.send_message(msg)
    server.quit()

def open_password_recovery():
    def submit_email():
        user_email = email_entry.get()
        recuperar_senha(user_email)
        #fechar a janela de recuperação de senha após enviar o e-mail

    recovery_window = Tk()
    recovery_window.title("Recuperar Senha")

    width_window = 420
    height_window = 100

    width_screen = recovery_window.winfo_screenwidth()
    height_screen = recovery_window.winfo_screenheight()

    pos_x  = (width_screen // 2) - (width_window // 2)
    pos_y  = (height_screen // 2) - (height_window // 2)

    recovery_window.geometry('{}x{}+{}+{}'.format(width_window, height_window, pos_x, pos_y))

    email_label = Label(recovery_window, text="Insira o e-mail associado à sua conta:", font="Arial 12")
    email_label.pack()

    email_entry = Entry(recovery_window)
    email_entry.pack()

    submit_button = Button(recovery_window, text="Enviar", command=recovery_window.destroy)
    submit_button.pack()

    recovery_window.mainloop()

def register_new_user(login_window):
    register_user_window = Toplevel(login_window)
    register_user_window.title("Sign up")
    register_user_window.configure(bg="#ADD8E6")

    width_window = 420
    height_window = 220

    width_screen = register_user_window.winfo_screenwidth()
    height_screen = register_user_window.winfo_screenheight()

    pos_x = (width_screen // 2) - (width_window // 2)
    pos_y = (height_screen // 2) - (height_window // 2)

    register_user_window.geometry('{}x{}+{}+{}'.format(width_window, height_window, pos_x, pos_y))

    title_lbl = Label(register_user_window, text="Create an account", font="Arial 14 bold", bg="#ADD8E6")
    title_lbl.grid(row=0, column=0, columnspan=2, pady=10)

    username_lbl = Label(register_user_window, text="Enter your username", font="Arial 12 bold", bg="#ADD8E6")
    username_lbl.grid(row=1, column=0, sticky="e")

    password_lbl = Label(register_user_window, text="Enter your password", font="Arial 12 bold", bg="#ADD8E6")
    password_lbl.grid(row=2, column=0, sticky="e")

    new_username_entry = Entry(register_user_window, font="Arial 14")
    new_username_entry.grid(row=1, column=1, pady=10, padx=10)

    new_password_entry = Entry(register_user_window, show="*", font="Arial 14")
    new_password_entry.grid(row=2, column=1, pady=10, padx=10)

    def save_user():
        new_user = new_username_entry.get()
        new_password = new_password_entry.get()

        connection = connect()

        cursor = connection.cursor()

        cursor.execute("INSERT INTO Usuarios (Nome, Senha) VALUES (?, ?)", (new_user, new_password))

        connection.commit()

        register_user_window.destroy()

    btn_save_cadastro = Button(register_user_window, text="Confirm", font=("Arial", 12), command=save_user)
    btn_save_cadastro.grid(row=3, column=1, columnspan=1, pady=10)

    btn_cancel_cadastro = Button(register_user_window, text="Cancel", font=("Arial", 12), command=register_user_window.destroy)
    btn_cancel_cadastro.grid(row=3, column=0, columnspan=1, pady=10)

    register_user_window.grid_columnconfigure(0, weight=1)
    register_user_window.grid_columnconfigure(1, weight=1)


def show_login_history(main_window):
    login_history_window = Toplevel(main_window)
    login_history_window.title("Login History")

    width_window = 600
    height_window = 700

    width_screen = login_history_window.winfo_screenwidth()
    height_screen = login_history_window.winfo_screenheight()

    pos_x = (width_screen // 2) - (width_window // 2)
    pos_y = (height_screen // 2) - (height_window // 2)

    login_history_window.geometry('{}x{}+{}+{}'.format(width_window, height_window, pos_x, pos_y))

    scrollbar = Scrollbar(login_history_window)
    scrollbar.pack(side="right", fill="y")

    login_list = Listbox(login_history_window, yscrollcommand=scrollbar.set)
    login_list.pack()

    # Exemplo de preenchimento da Listbox (simulado)
    login_data = [
        "test - 2023-01-01 10:00:00",
      ]

    for login in login_data:
        login_list.insert("end", login)

    scrollbar.config(command=login_list.yview)

    Label(login_history_window, text="Detailed login information").pack()

    login_history_window.mainloop()

def show_profile(main_window):
    profile_window = Toplevel(main_window)
    profile_window.title("Profile")

    width_window = 450
    height_window = 300

    width_screen = profile_window.winfo_screenwidth()
    height_screen = profile_window.winfo_screenheight()

    pos_x = (width_screen // 2) - (width_window // 2)
    pos_y = (height_screen // 2) - (height_window // 2)

    profile_window.geometry('{}x{}+{}+{}'.format(width_window, height_window, pos_x, pos_y))
    
    # Aqui, você exibe o nome do usuário logado
    if logged_in_user:
        label = Label(profile_window, text=f"Logged in as: {logged_in_user}")
        label.pack()

def validate_user(username, password):
    conection = pyodbc.connect("Driver={SQLite3 ODBC Driver};Server=localhost;Database=Projeto.db")
    cursor = conection.cursor()
    cursor.execute("SELECT * FROM Usuarios WHERE Nome = ? AND Senha = ?", (username, password))
    user = cursor.fetchone()
    
    if user:
        return True
    else:
        return False

def show_login_window():
    login_window = Tk()
    login_window.title("Sign in")

    def show_register_window():
        register_new_user(login_window)

    def verify_login():
        global logged_in_user  # Usada para acessar a variável global
        username = username_entry.get()
        password = password_entry.get()

        if validate_user(username, password):
            login_window.destroy()  # Feche a janela de login
            open_main_interface()  # Abra a nova janela

        else:
            msg_lbl = Label(login_window, text="Incorrect username or password.", fg="red")
            msg_lbl.grid(row=3,column=0,columnspan=2)

        logged_in_user = username  # Armazena o nome do usuário logado

    login_window.configure(bg="#00008B")
    width_window = 450
    height_window = 300

    width_screen = login_window.winfo_screenwidth()
    height_screen = login_window.winfo_screenheight()

    pos_x  = (width_screen // 2) - (width_window // 2)
    pos_y  = (height_screen // 2) - (height_window // 2)

    login_window.geometry('{}x{}+{}+{}'.format(width_window, height_window, pos_x, pos_y))

    title_lbl = Label(login_window, text="SysControl", font="Arial 20", bg="#00008B", fg="#ffffff")
    title_lbl.grid(row=0,column=0,columnspan=2, pady=20) 

    username_lbl = Label(login_window, text="Username", font="Arial 14 bold", bg="#00008B", fg="#ffffff")
    username_lbl.grid(row=1,column=0,sticky="NSEW") #NSEW

    password_lbl = Label(login_window, text="Password", font="Arial 14 bold", bg="#00008B", fg="#ffffff")
    password_lbl.grid(row=2,column=0,sticky="NSEW") #NSEW

    username_entry = Entry(login_window, font="Arial 14")
    username_entry.grid(row=1,column=1,pady=10)

    password_entry = Entry(login_window, show="*",font="Arial 14")
    password_entry.grid(row=2,column=1, pady=10)

    login_btn = Button(login_window, text="Create an account", font="Arial 12", bg="#eeeeee", command=show_register_window)
    login_btn.grid(row=4, column=0,columnspan=1, padx=20, pady=10, sticky="NSEW")

    login_btn = Button(login_window, text="Login", font="Arial 12", bg="#eeeeee", command=verify_login)
    login_btn.grid(row=4, column=1,columnspan=2, padx=20, pady=10, sticky="NSEW")

    exit_btn = Button(login_window, text="Exit", font="Arial 12", bg="#eeeeee", command=login_window.destroy)
    exit_btn.grid(row=5, column=0,columnspan=2, padx=20, pady=10, sticky="NSEW")

    # Adicionando o botão "Recuperar Senha"
    password_recovery_button = Button(login_window, text="Recuperar Senha", command=open_password_recovery)
    password_recovery_button.grid(row=6, column=0, columnspan=2, padx=20, pady=10, sticky="NSEW")

    for i in range(5):
        login_window.grid_rowconfigure(i, weight=1)
        login_window.grid_rowconfigure(i, weight=1)

    for i in range(2):
        login_window.grid_columnconfigure(i, weight=1)
    
    login_window.mainloop()

def export_to_csv(products):
    with open('product_list.csv', 'w', newline='') as file:
        writer = csv.writer(file)
        writer.writerow(["ID", "Nome", "Descricao", "Preco"])  # Cabeçalho do CSV
        for product in products:
            writer.writerow(product)

# Função para salvar o conteúdo do bloco de notas
def save_note():
    content = text_widget.get("1.0", tk.END)
    with open("note.txt", "w") as file:
        file.write(content)
    messagebox.showinfo("Nota", "Nota salva com sucesso!")

# Função para carregar o conteúdo do bloco de notas
def load_note():
    try:
        with open("note.txt", "r") as file:
            content = file.read()
        text_widget.delete("1.0", tk.END)
        text_widget.insert(tk.END, content)
    except FileNotFoundError:
        messagebox.showinfo("Nota", "Nenhuma nota encontrada!")

# Função para abrir a janela do bloco de notas
def open_note_window(main_window):
    note_window = Toplevel(main_window)
    note_window.title("Bloco de Notas")

    text_widget = tk.Text(note_window, width=60, height=15)
    text_widget.pack()

    scrollbar = ttk.Scrollbar(note_window, command=text_widget.yview)
    scrollbar.pack(side=tk.RIGHT, fill='y')

    text_widget['yscrollcommand'] = scrollbar.set

    frame = ttk.Frame(note_window)
    frame.pack()

    save_button = ttk.Button(frame, text="Salvar", command=save_note)
    save_button.pack(side=tk.LEFT)

    load_button = ttk.Button(frame, text="Carregar", command=load_note)
    load_button.pack(side=tk.LEFT)

def open_main_interface():

    connection = connect()

    cursor = connection.cursor()

    def new_product_action():
        product_interface.register_new_product()

    def show_historic_action():
        product_interface.show_product_history()

    def search_action():
        product_interface.search_product(product_name.get(), product_description.get())

    def delete_product_action():
        product_interface.delete_product_treeview()
    
    def calculate_price_action():
        product_interface.calculate_price_statistics()

    def login_historic_action():
        show_login_history(main_window)

    def login_current_action():
        show_profile(main_window)

    def notes_action():
        open_note_window(main_window)

    main_window = Tk()
    main_window.title("SysControl")
    main_window.configure(bg="#eeeeee")
    main_window.attributes("-fullscreen", True)

    menu_bar = Menu(main_window)
    main_window.configure(menu=menu_bar)

    menu_arquivo = Menu(menu_bar, tearoff=0)
    menu_bar.add_cascade(label="Menu", menu=menu_arquivo)

    menu_arquivo.add_command(label="Register",command=new_product_action)

    menu_arquivo.add_command(label="Exit", command=main_window.destroy)

    # Adicionando uma nova opção para o histórico de login
    menu_login = Menu(menu_bar, tearoff=0)
    menu_bar.add_cascade(label="Login", menu=menu_login)

    menu_login.add_command(label="Profile", command=login_current_action)
    menu_login.add_command(label="Login History", command=login_historic_action)

    menu_opcoes = tk.Menu(menu_bar, tearoff=0)
    menu_bar.add_cascade(label="Opções", menu=menu_opcoes)

    menu_opcoes.add_command(label="Bloco de Notas", command=notes_action)
    menu_opcoes.add_command(label="Sair", command=main_window.destroy)

    Label(main_window, text="Search by", font="Arial 14", bg="#eeeeee").grid(row=0,column=1, padx=10, pady=10)

    Label(main_window, text="Product Name: ", font="Arial 14", bg="#eeeeee").grid(row=0,column=2, padx=10, pady=10)
    product_name = Entry(main_window, font="Arial 14")
    product_name.grid(row=0,column=3, padx=10, pady=10)

    Label(main_window, text="Product Description: ", font="Arial 14", bg="#eeeeee").grid(row=0,column=5, padx=10, pady=10)
    product_description = Entry(main_window, font="Arial 14")
    product_description.grid(row=0,column=6, padx=10, pady=10)

    Label(main_window, text="All Products", font="Arial 18 bold", fg="black" , bg="#eeeeee").grid(row=2,column=0, columnspan=10, padx=10, pady=10)

    btn_save = Button(main_window, text="New Product", font="Arial 26", command=new_product_action)
    btn_save.grid(row=4,column=0, columnspan=4, sticky="NSEW", padx=20, pady=5)

    btn_delete = Button(main_window, text="Delete", font="Arial 26", command=delete_product_action)
    btn_delete.grid(row=4,column=4, columnspan=4, sticky="NSEW", padx=20, pady=5)

    btn_show_history = Button(main_window, text="Show History", font="Arial 16", command=show_historic_action)
    btn_show_history.grid(row=5, column=0, columnspan=8, sticky="NSEW", padx=20, pady=10)

    # Adicione o botão e a chamada da função na interface principal
    btn_export = Button(main_window, text="Export to CSV", font="Arial 12", command=lambda: export_to_csv(product_interface.products))
    btn_export.grid(row=6, column=0, columnspan=4, sticky="NSEW", padx=20, pady=5)

    btn_statistics = Button(main_window, text="Calculate Price Statistics", font="Arial 12", command=calculate_price_action)
    btn_statistics.grid(row=8, column=0, columnspan=4, sticky="NSEW", padx=20, pady=5)

    def sort_ascending():
        product_interface.sort_products_by_price('asc')

    def sort_descending():
        product_interface.sort_products_by_price('desc')

    btn_asc_order = Button(main_window, text="Sort Ascending", font="Arial 12", command=sort_ascending)
    btn_asc_order.grid(row=9,column=0, columnspan=4, sticky="NSEW", padx=20, pady=5)

    btn_desc_order = Button(main_window, text="Sort Descending", font="Arial 12", command=sort_descending)
    btn_desc_order.grid(row=9,column=4, columnspan=4, sticky="NSEW", padx=20, pady=5)

    connection = pyodbc.connect("Driver={SQLite3 ODBC Driver};Server=localhost;Database=Projeto.db")

    product_interface = ProductInterface(main_window)
    
    product_interface.create_treeview()
    
    product_interface.list_products()

    product_name.bind('<KeyRelease>', lambda e: product_interface.filter_data(product_name, product_description))
    product_description.bind('<KeyRelease>', lambda e: product_interface.filter_data(product_name, product_description))

    # Chamar a função após a criação da janela principal e a exibição da mesma
    main_window.after(100, product_interface.display_expensive_cheap)

    main_window.mainloop()
    connection.close()

show_login_window()