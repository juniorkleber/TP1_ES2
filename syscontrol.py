# Importa o módulo pyodbc para conexão com o banco de dados
import pyodbc

# Importa o módulo tkinter para criar interfaces gráficas
from tkinter import * 
from tkinter import ttk

#Função para realizar a conexão com o banco de dados para um função
#Para que não seja necessária uma nova conexão a cada interação
def connect_to_database():
    # Configuração da conexão com o banco de dados
    dadosConexao = ("Driver={SQLite3 ODBC Driver};Server=localhost;Database=Projeto.db")
    
    # Cria uma nova conexão com o banco de dados
    connection = pyodbc.connect(dadosConexao)

    return connection

class ProductInterface:
     # Inicializa a interface do produto com a janela principal e a conexão ao banco de dados
    def __init__(self, main_window):
        self.main_window = main_window
        self.setup_ui()

    # Configura a interface do usuário
    def setup_ui(self):
        self.create_treeview()  # Cria a visualização em árvore para exibir os produtos

    # Cria a TreeView para exibir os produtos e configura as colunas
    # A TreeView exibe os dados dos produtos e permite a edição deles
    def create_treeview(self):
        self.treeview = ttk.Treeview(self.main_window, columns=("ID", "Nome", "Descricao", "Preco"), show="headings")
        
        # Configuração dos cabeçalhos das colunas
        self.treeview.heading("ID", text="ID")
        self.treeview.heading("Nome", text="Product Name")
        self.treeview.heading("Descricao", text="Product Description")
        self.treeview.heading("Preco", text="Product Price")
        
        # Configuração das larguras das colunas
        self.treeview.column("#0", width=0, stretch="NO")
        self.treeview.column("ID", width=100)
        self.treeview.column("Nome", width=300)
        self.treeview.column("Descricao", width=500)
        self.treeview.column("Preco", width=200)
        
        self.treeview.grid(row=3, column=0, columnspan=10, sticky="NSEW")

        # Permite edição de dados com duplo clique
        self.treeview.bind("<Double-1>", self.edit_product)

    #limpo dados da treeview
    def clear_treeview(self):
        for i in self.treeview.get_children():
            # Remove linha por linha
            self.delete_product_treeview()

    # Exibe todos os produtos existentes no banco de dados na TreeView
    def list_products(self):
        for item in self.treeview.get_children():
            self.treeview.delete(item)

        connection = connect_to_database()

        # Realiza uma consulta SQL para selecionar todos os produtos
        cursor = connection.cursor()
        cursor.execute("SELECT * FROM Produtos")

        # Obtém os resultados da consulta
        products = cursor.fetchall()

        # Adiciona os produtos à TreeView
        for product in products:
            self.treeview.insert("", "end", values=(product[0], product[1], product[2], product[3]))

        # Fecha o cursor
        cursor.close()

    # Cria uma nova janela para inserir os dados de um novo produto
    def register_new_product(self):
        register_product_window = Toplevel(self.main_window)
        register_product_window.title("Register New Product")
        register_product_window.configure(bg="#eeeeee")

        # Define as dimensões da janela de cadastro.
        width_window = 450
        height_window = 230

        # Obtém o tamanho da tela do usuário.
        width_screen = register_product_window.winfo_screenwidth()
        height_screen = register_product_window.winfo_screenheight()

        # Centraliza a janela no meio da tela.
        pos_x = (width_screen // 2) - (width_window // 2)
        pos_y = (height_screen // 2) - (height_window // 2)

        register_product_window.geometry('{}x{}+{}+{}'.format(width_window, height_window, pos_x, pos_y))

        # Define um estilo de borda para widgets.
        border_style = {"borderwidth": 2, "relief": "groove"}

        # Cria rótulos e campos de entrada para o nome, descrição e preço do produto.
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

        # Esta função é chamada quando o botão "Salvar" na janela de cadastro é clicado.
        def saveData():
            # Coleta os valores inseridos nos campos de entrada.
            register_new_product_ = (product_name_register.get(), product_description_register.get(), product_price_register.get())

            connection = connect_to_database()

            # Cria um novo cursor para executar SQL na nova conexão
            cursor = connection.cursor()

            # Executa uma operação SQL para inserir os dados no banco de dados.
            cursor.execute("INSERT INTO Produtos (Nome, Descricao, Preco) Values (?,?,?)", register_new_product_)
            # Grava as alterações no banco de dados.
            connection.commit()

            print("Product registered successfully!")

            # Atualiza a lista de dados na interface.
            register_product_window.destroy()

            self.list_products()  # Chama a função list_products da classe ProductInterface

        #Botes para salvar ou cancelar a requisão de registro de algum item
        btn_save_product = Button(register_product_window, text="Save", font=("Arial", 14), bg="#008000", fg="#ffffff", command=saveData)
        btn_save_product.grid(row=3,column=0, columnspan=2,padx=10,pady=10, sticky="NSEW")

        btn_cancel_product = Button(register_product_window, text="Cancel", font=("Arial", 14), bg="#FF0000", fg="#ffffff", command=register_product_window.destroy)
        btn_cancel_product.grid(row=4,column=0, columnspan=2,padx=10,pady=10, sticky="NSEW")

    # Permite editar os detalhes de um produto selecionado na TreeView
    def edit_product(self, event):
        # Obtém o item selecionado na TreeView.
        selected_item = self.treeview.selection()[0]

        # Obtém os valores do item selecionado na TreeView.
        select_values = self.treeview.item(selected_item)['values']

        # Cria uma nova janela para a edição do produto.
        edit_product_window = Toplevel(self.main_window)
        edit_product_window.title("Edit Product")
        edit_product_window.configure(bg="#eeeeee")

        # Define as dimensões da janela de edição.
        width_window = 500
        height_window = 200

        # Obtém o tamanho da tela do usuário.
        width_screen = edit_product_window.winfo_screenwidth()
        height_screen = edit_product_window.winfo_screenheight()

        # Centraliza a janela no meio da tela.
        pos_x = (width_screen // 2) - (width_window // 2)
        pos_y = (height_screen // 2) - (height_window // 2)

        edit_product_window.geometry('{}x{}+{}+{}'.format(width_window, height_window, pos_x, pos_y))

        # Define um estilo de borda para os widgets.
        border_style = {"borderwidth": 2, "relief": "groove"}

        # Cria rótulos e campos de entrada para editar o nome, descrição e preço do produto.
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
        
        # Esta função é chamada quando o usuário clica no botão "Confirm" para salvar as edições feitas em um produto.
        def saveEdit():
            # Obtém os novos valores inseridos nos campos de edição.
            product = product_name_edit.get()
            new_description = product_description_edit.get()
            new_price = product_price_edit.get()

            # Atualiza os valores do item na TreeView com os novos valores.
            self.treeview.item(selected_item, values=(select_values[0], product, new_description, new_price))

            connection = connect_to_database()

            # Cria um novo cursor para executar SQL na nova conexão
            cursor = connection.cursor()

            # Executa uma operação SQL para atualizar o produto no banco de dados.
            cursor.execute("UPDATE Produtos SET Nome = ?, Descricao = ?, Preco = ? WHERE ID = ?", (product, new_description, new_price, select_values[0]))

            # Grava as alterações no banco de dados.
            connection.commit()

            print("Data Registered Successfully!")

            # Fecha a janela de edição.
            edit_product_window.destroy()

            # Atualiza a lista de dados na interface.
            self.list_products()

        #Botoes para salvar ou cancelar a requisão de registro de algum item
        btn_save_edit = Button(edit_product_window, text="Confirm", font=("Arial", 14), bg="#008000", fg="#ffffff", command=saveEdit)
        btn_save_edit.grid(row=4,column=1, padx=20,pady=20)
        btn_cancel_edit = Button(edit_product_window, text="Cancel", font=("Arial", 14), bg="#FF0000", fg="#ffffff", command=edit_product_window.destroy)
        btn_cancel_edit.grid(row=4,column=0, padx=20,pady=20)

    # Remove o produto selecionado na TreeView e na base de dados
    def delete_product_treeview(self):
        # Obtém o item selecionado na TreeView
        selected_item = self.treeview.selection()
        
        if selected_item:
            # Pega o ID do item selecionado
            product_id = self.treeview.item(selected_item[0])['values'][0]
            
            # Remove o item da TreeView
            self.treeview.delete(selected_item)
            
            connection = connect_to_database()

            # Remove o produto do banco de dados
            cursor = connection.cursor()
            cursor.execute("DELETE FROM Produtos WHERE ID = ?", (product_id,))
            connection.commit()

        else:
            # Se nenhum item estiver selecionado, exiba uma mensagem de erro
            messagebox.showerror("Error", "Select a product to delete.")

    # Filtra os produtos exibidos na TreeView com base no nome e descrição fornecidos
    # Se os campos estiverem vazios, exibe todos os produtos
    def filter_data(self, product_name, product_description):
        
        connection = connect_to_database()
        # Cria um novo cursor para executar SQL na nova conexão
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
        
        #limpo dados da treeview
        for i in self.treeview.get_children():
            #deleto linha por linha
            self.treeview.delete(i)
            
        #preenche preview com dados filtrados   
        for data in product:
            self.treeview.insert('', 'end', values=(data[0], data[1],data[2],data[3]))

def register_new_user(login_window):
    # Cria uma nova janela superior (Toplevel) para o cadastro.
    register_user_window = Toplevel(login_window)
    register_user_window.title("Sign up")
    register_user_window.configure(bg="#ADD8E6")

    # Define as dimensões da janela de cadastro.
    width_window = 420
    height_window = 220

    # Obtém o tamanho da tela do usuário.
    width_screen = register_user_window.winfo_screenwidth()
    height_screen = register_user_window.winfo_screenheight()

    # Centraliza a janela no meio da tela.
    pos_x = (width_screen // 2) - (width_window // 2)
    pos_y = (height_screen // 2) - (height_window // 2)

    register_user_window.geometry('{}x{}+{}+{}'.format(width_window, height_window, pos_x, pos_y))

    title_lbl = Label(register_user_window, text="Create an account", font="Arial 14 bold", bg="#ADD8E6")
    title_lbl.grid(row=0, column=0, columnspan=2, pady=10)

    username_lbl = Label(register_user_window, text="Enter your username", font="Arial 12 bold", bg="#ADD8E6")
    username_lbl.grid(row=1, column=0, sticky="e")

    password_lbl = Label(register_user_window, text="Enter your password", font="Arial 12 bold", bg="#ADD8E6")
    password_lbl.grid(row=2, column=0, sticky="e")

    # Cria campos de entrada para o nome de usuário e senha
    new_username_entry = Entry(register_user_window, font="Arial 14")
    new_username_entry.grid(row=1, column=1, pady=10, padx=10)

    new_password_entry = Entry(register_user_window, show="*", font="Arial 14")
    new_password_entry.grid(row=2, column=1, pady=10, padx=10)

    def save_user():
        new_user = new_username_entry.get()
        new_password = new_password_entry.get()

        connection = connect_to_database()

        # Cria um novo cursor para executar SQL na nova conexão
        cursor = connection.cursor()

        # Executa uma operação SQL para inserir os dados no banco de dados.
        cursor.execute("INSERT INTO Usuarios (Nome, Senha) VALUES (?, ?)", (new_user, new_password))

        # Grava as alterações no banco de dados.
        connection.commit()

        # Feche a janela de cadastro após a conclusão
        register_user_window.destroy()

    #Botoes da tela de registro de novo usuario
    btn_save_cadastro = Button(register_user_window, text="Confirm", font=("Arial", 12), command=save_user)
    btn_save_cadastro.grid(row=3, column=1, columnspan=1, pady=10)

    btn_cancel_cadastro = Button(register_user_window, text="Cancel", font=("Arial", 12), command=register_user_window.destroy)
    btn_cancel_cadastro.grid(row=3, column=0, columnspan=1, pady=10)

    # Define as colunas 0 e 1 para ocupar todo o espaço disponível
    register_user_window.grid_columnconfigure(0, weight=1)
    register_user_window.grid_columnconfigure(1, weight=1)

def validate_user(username, password):
    # Sua lógica para validar o usuário aqui
    # Retorne True se as credenciais forem válidas, caso contrário, retorne False
    conection = pyodbc.connect("Driver={SQLite3 ODBC Driver};Server=localhost;Database=Projeto.db")
    cursor = conection.cursor()
    cursor.execute("SELECT * FROM Usuarios WHERE Nome = ? AND Senha = ?", (username, password))
    user = cursor.fetchone()
    
    if user:
        return True
    else:
        return False

def show_login_window():
    # Cria a janela principal
    login_window = Tk()
    login_window.title("Sign in")

    def show_register_window():
        register_new_user(login_window)

    def verify_login():
        username = username_entry.get()
        password = password_entry.get()

        # Validação das credenciais
        if validate_user(username, password):
            login_window.destroy()  # Feche a janela de login
            open_main_interface()  # Abra a nova janela

        else:
            # Se a condição "if" anterior não for atendida, ou seja, se as credenciais de login forem incorretas
            # Cria um rótulo de aviso com texto "Nome de Usuário ou Senha incorretos" em vermelho
            msg_lbl = Label(login_window, text="Incorrect username or password.", fg="red")
            msg_lbl.grid(row=3,column=0,columnspan=2)

    # Configura janela principal
    login_window.configure(bg="#00008B")
    width_window = 450
    height_window = 300

    # Obtém o tamanho da tela
    width_screen = login_window.winfo_screenwidth()
    height_screen = login_window.winfo_screenheight()

    # Calcula a posição para centralizar a janela na tela
    pos_x  = (width_screen // 2) - (width_window // 2)
    pos_y  = (height_screen // 2) - (height_window // 2)

    # Define as dimensões e posição da janela principal
    login_window.geometry('{}x{}+{}+{}'.format(width_window, height_window, pos_x, pos_y))

    # Cria rótulos para título, nome de usuário e senha
    title_lbl = Label(login_window, text="SysControl", font="Arial 20", bg="#00008B", fg="#ffffff")
    title_lbl.grid(row=0,column=0,columnspan=2, pady=20) 

    username_lbl = Label(login_window, text="Username", font="Arial 14 bold", bg="#00008B", fg="#ffffff")
    username_lbl.grid(row=1,column=0,sticky="NSEW") #NSEW

    password_lbl = Label(login_window, text="Password", font="Arial 14 bold", bg="#00008B", fg="#ffffff")
    password_lbl.grid(row=2,column=0,sticky="NSEW") #NSEW

    # Cria campos de entrada para o nome de usuário e senha
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

    # Define a configuração de peso para as linhas e colunas da grade da tela de login
    for i in range(5):
        login_window.grid_rowconfigure(i, weight=1)
        login_window.grid_rowconfigure(i, weight=1)

    for i in range(2):
        login_window.grid_columnconfigure(i, weight=1)
    
    login_window.mainloop()

#Funçao que é responsavel pela abertura e gerencia da tela principal do programa
def open_main_interface():

    connection = connect_to_database()

    # Cria um novo cursor para executar SQL na nova conexão
    cursor = connection.cursor()

    # Defina uma função de ação para o botão "New Product"
    def new_product_action():
        product_interface.register_new_product()

    # Defina funções de ação para botões, como o botão de pesquisa
    def search_action():
        product_interface.search_product(product_name.get(), product_description.get())

    def delete_product_action():
        product_interface.delete_product_treeview()

    # Cria uma nova janela para a tela principal
    main_window = Tk()
    main_window.title("SysControl")
    main_window.configure(bg="#eeeeee")
    main_window.attributes("-fullscreen", True)

    #Configura Janela de atalhos
    menu_bar = Menu(main_window)
    main_window.configure(menu=menu_bar)

    menu_arquivo = Menu(menu_bar, tearoff=0)
    menu_bar.add_cascade(label="Menu", menu=menu_arquivo)

    #Cria opção menu Cadastrar
    menu_arquivo.add_command(label="Register",command=new_product_action)

    #Cria opção menu Sair
    menu_arquivo.add_command(label="Exit", command=main_window.destroy)

    # Cria rótulos e campos de entrada para nome e descrição do produto
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

    # Cria um botão "Deletar" na janela que chama a função delete() quando clicado
    btn_delete = Button(main_window, text="Delete", font="Arial 26", command=delete_product_action)
    btn_delete.grid(row=4,column=4, columnspan=4, sticky="NSEW", padx=20, pady=5)

    # Cria a conexão com o banco de dados
    connection = pyodbc.connect("Driver={SQLite3 ODBC Driver};Server=localhost;Database=Projeto.db")

    # Cria uma instância da classe ProductInterface
    product_interface = ProductInterface(main_window)
    
    # Chame o método para criar a TreeView
    product_interface.create_treeview()
    
    # Chame a função list_products na inicialização para preencher a TreeView
    product_interface.list_products()

    #Abre a instancia da função da barra de pesquisa
    product_name.bind('<KeyRelease>', lambda e: product_interface.filter_data(product_name, product_description))
    product_description.bind('<KeyRelease>', lambda e: product_interface.filter_data(product_name, product_description))

    # Mantém a janela principal aberta
    main_window.mainloop()

    # Feche a conexão com o banco de dados após fechar a janela principal
    connection.close()

# Execução da aplicação
show_login_window()