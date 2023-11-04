import pyodbc

# Importa o módulo tkinter para criar interfaces gráficas
from tkinter import *
from tkinter import ttk

class product:
     # Inicializa a interface do produto com a janela principal e a conexão ao banco de dados
    def __init__(self, main_window):
        self.main_window = main_window
        self.setup_ui()

    # Configura a interface do usuário
    def setup_ui(self):
        self.create()  # Cria a visualização em árvore para exibir os produtos

    # Cria a TreeView para exibir os produtos e configura as colunas
    # A TreeView exibe os dados dos produtos e permite a edição deles
    def create(self):
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
        self.treeview.bind("<Double-1>", self.edit)


    # Exibe todos os produtos existentes no banco de dados na TreeView
    def list(self):
        for item in self.treeview.get_children():
            self.treeview.delete(item)

        # Configuração da conexão com o banco de dados
        dadosConexao = ("Driver={SQLite3 ODBC Driver};Server=localhost;Database=Projeto.db")

        # Cria uma nova conexão com o banco de dados
        connection = pyodbc.connect(dadosConexao)

        # Cria um novo cursor para executar SQL na nova conexão
        cursor = connection.cursor()

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
    def register_p(self):
        window4 = Toplevel(self.main_window)
        window4.title("Register New Product")
        window4.configure(bg="#eeeeee")

        # Define as dimensões da janela de cadastro.
        width_window = 450
        height_window = 230

        # Obtém o tamanho da tela do usuário.
        width_screen = window4.winfo_screenwidth()
        height_screen = window4.winfo_screenheight()

        # Centraliza a janela no meio da tela.
        pos_x = (width_screen // 2) - (width_window // 2)
        pos_y = (height_screen // 2) - (height_window // 2)

        window4.geometry('{}x{}+{}+{}'.format(width_window, height_window, pos_x, pos_y))

        # Define um estilo de borda para widgets.
        border_style = {"borderwidth": 2, "relief": "groove"}

        # Cria rótulos e campos de entrada para o nome, descrição e preço do produto.
        Label(window4, text="Name", 
              font=("Arial", 12), bg="#eeeeee").grid(row=0, column=0, padx=10, pady=10, sticky="w")
        name = Entry(window4, font=("Arial", 14), **border_style, bg="#eeeeee")
        name.grid(row=0, column=1, padx=10, pady=10)

        Label(window4, text="Description", 
              font=("Arial", 12), bg="#eeeeee").grid(row=1, column=0, padx=10, pady=10, sticky="w")
        description = Entry(window4, font=("Arial", 14), **border_style, bg="#eeeeee")
        description.grid(row=1, column=1, padx=10, pady=10)

        Label(window4, text="Price", 
              font=("Arial", 12), bg="#eeeeee").grid(row=2, column=0, padx=10, pady=10, sticky="w")
        price = Entry(window4, font=("Arial", 14), **border_style, bg="#eeeeee")
        price.grid(row=2, column=1, padx=10, pady=10, columnspan=2)

        for i in range(5):
            window4.grid_rowconfigure(i, weight=1)

        for i in range(2):
            window4.grid_columnconfigure(i, weight=1)

        # Esta função é chamada quando o botão "Salvar" na janela de cadastro é clicado.
        def save():
            # Coleta os valores inseridos nos campos de entrada.
            register_new_product_ = (name.get(), description.get(), price.get())

            # Configuração da conexão com o banco de dados
            dadosConexao = ("Driver={SQLite3 ODBC Driver};Server=localhost;Database=Projeto.db")

            # Cria uma nova conexão com o banco de dados
            connection = pyodbc.connect(dadosConexao)

            # Cria um novo cursor para executar SQL na nova conexão
            cursor = connection.cursor()
            
            # Executa uma operação SQL para inserir os dados no banco de dados.
            cursor.execute("INSERT INTO Produtos (Nome, Descricao, Preco) Values (?,?,?)", register_new_product_)
            # Grava as alterações no banco de dados.
            connection.commit()

            print("Product registered successfully!")

            # Atualiza a lista de dados na interface.
            window4.destroy()

            self.list()  # Chama a função list_products da classe ProductInterface

        #Botes para salvar ou cancelar a requisão de registro de algum item
        btn1 = Button(window4, text="Save", font=("Arial", 14), bg="#008000", fg="#ffffff", command=save)
        btn1.grid(row=3,column=0, columnspan=2,padx=10,pady=10, sticky="NSEW")

        btn2 = Button(window4, text="Cancel", font=("Arial", 14), bg="#FF0000", fg="#ffffff", command=window4.destroy)
        btn2.grid(row=4,column=0, columnspan=2,padx=10,pady=10, sticky="NSEW")

    # Permite editar os detalhes de um produto selecionado na TreeView
    def edit(self, event):
        # Obtém o item selecionado na TreeView.
        selected_item = self.treeview.selection()[0]

        # Obtém os valores do item selecionado na TreeView.
        select_values = self.treeview.item(selected_item)['values']

        # Cria uma nova janela para a edição do produto.
        window3 = Toplevel(self.main_window)
        window3.title("Edit Product")
        window3.configure(bg="#eeeeee")

        # Define as dimensões da janela de edição.
        width_window = 500
        height_window = 200

        # Obtém o tamanho da tela do usuário.
        width_screen = window3.winfo_screenwidth()
        height_screen = window3.winfo_screenheight()

        # Centraliza a janela no meio da tela.
        pos_x = (width_screen // 2) - (width_window // 2)
        pos_y = (height_screen // 2) - (height_window // 2)

        window3.geometry('{}x{}+{}+{}'.format(width_window, height_window, pos_x, pos_y))

        # Define um estilo de borda para os widgets.
        border_style = {"borderwidth": 2, "relief": "groove"}

        # Cria rótulos e campos de entrada para editar o nome, descrição e preço do produto.
        Label(window3, text="Product Name", 
              font=("Arial", 16), bg="#eeeeee").grid(row=0, column=0, padx=10, pady=10, sticky="w")
        name = Entry(window3, font=("Arial", 16), **border_style, bg="#eeeeee", 
                     textvariable=StringVar(value=select_values[1]))
        name.grid(row=0, column=1, padx=10, pady=10)

        Label(window3, text="Product Description", 
              font=("Arial", 16), bg="#eeeeee").grid(row=1, column=0, padx=10, pady=10, sticky="w")
        description = Entry(window3, font=("Arial", 16), **border_style, bg="#eeeeee", 
                            textvariable=StringVar(value=select_values[2]))
        description.grid(row=1, column=1, padx=10, pady=10)

        Label(window3, text="Product Price", 
              font=("Arial", 16), bg="#f5f5f5").grid(row=2, column=0, padx=10, pady=10, sticky="w")
        price = Entry(window3, font=("Arial", 16), **border_style, bg="#f5f5f5", 
                      textvariable=StringVar(value=select_values[3]))
        price.grid(row=2, column=1, padx=10, pady=10)

        for i in range(5):
            window3.grid_rowconfigure(i, weight=1)

        for i in range(2):
            window3.grid_columnconfigure(i, weight=1)
        
        # Esta função é chamada quando o usuário clica no botão "Confirm" para salvar as edições feitas em um produto.
        def saveEdit():
            # Obtém os novos valores inseridos nos campos de edição.
            product = name.get()
            new_description = description.get()
            new_price = price.get()

            # Atualiza os valores do item na TreeView com os novos valores.
            self.treeview.item(selected_item, values=(select_values[0], product, new_description, new_price))

            # Configuração da conexão com o banco de dados
            dadosConexao = ("Driver={SQLite3 ODBC Driver};Server=localhost;Database=Projeto.db")

            # Cria uma nova conexão com o banco de dados
            connection = pyodbc.connect(dadosConexao)

            # Cria um novo cursor para executar SQL na nova conexão
            cursor = connection.cursor()

            # Executa uma operação SQL para atualizar o produto no banco de dados.
            cursor.execute("UPDATE Produtos SET Nome = ?, Descricao = ?, Preco = ? WHERE ID = ?", 
                                (product, new_description, new_price, select_values[0]))

            # Grava as alterações no banco de dados.
            connection.commit()

            print("Data Registered Successfully!")

            # Fecha a janela de edição.
            window3.destroy()

            # Atualiza a lista de dados na interface.
            self.list()

        #Botoes para salvar ou cancelar a requisão de registro de algum item
        btn1 = Button(window3, text="Confirm", font=("Arial", 14), bg="#008000", fg="#ffffff", command=saveEdit)
        btn1.grid(row=4,column=1, padx=20,pady=20)

        btn2 = Button(window3, text="Cancel", font=("Arial", 14), bg="#FF0000", fg="#ffffff", command=window3.destroy)
        btn2.grid(row=4,column=0, padx=20,pady=20)

    # Remove o produto selecionado na TreeView e na base de dados
    def delete(self):
        # Obtém o item selecionado na TreeView
        selected_item = self.treeview.selection()
        
        if selected_item:
            # Pega o ID do item selecionado
            product_id = self.treeview.item(selected_item[0])['values'][0]
            
            # Remove o item da TreeView
            self.treeview.delete(selected_item)
            
            # Configuração da conexão com o banco de dados
            dadosConexao = ("Driver={SQLite3 ODBC Driver};Server=localhost;Database=Projeto.db")

            # Cria uma nova conexão com o banco de dados
            connection = pyodbc.connect(dadosConexao)

            # Cria um novo cursor para executar SQL na nova conexão
            cursor = connection.cursor()

            # Remove o produto do banco de dados
            
            cursor.execute("DELETE FROM Produtos WHERE ID = ?", (product_id,))
            connection.commit()

        else:
            # Se nenhum item estiver selecionado, exiba uma mensagem de erro
            messagebox.showerror("Error", "Select a product to delete.")

    # Filtra os produtos exibidos na TreeView com base no nome e descrição fornecidos
    # Se os campos estiverem vazios, exibe todos os produtos
    def filter(self, x, y):

        # Configuração da conexão com o banco de dados
        dadosConexao = ("Driver={SQLite3 ODBC Driver};Server=localhost;Database=Projeto.db")

        # Cria uma nova conexão com o banco de dados
        connection = pyodbc.connect(dadosConexao)

        # Cria um novo cursor para executar SQL na nova conexão
        cursor = connection.cursor()

        if not x.get() and not y.get():
            self.list()
            return
        
        sql = "SELECT * FROM Produtos"
        params = []
        if x.get():
            sql += " WHERE Nome LIKE ?"
            params.append('%' + x.get() + '%') 
        
        if y.get():
            if x.get():
                sql += " AND "
            else:
                sql += " WHERE"
            sql += " Descricao LIKE ?"
            params.append('%' + y.get() + '%')

        cursor.execute(sql, tuple(params))
        product = cursor.fetchall()
        
        #limpo dados da treeview
        for i in self.treeview.get_children():
            #deleto linha por linha
            self.treeview.delete(i)
            
        #preenche preview com dados filtrados   
        for data in product:
            self.treeview.insert('', 'end', values=(data[0], data[1],data[2],data[3]))
    def calculate_price_statistics(self):
        if not self.products:
            messagebox.showinfo("No Products", "No products available to calculate statistics.")
            return

        prices = [product[3] for product in self.products]

        average_price = sum(prices) / len(prices)
        max_price = max(prices)
        min_price = min(prices)

        messagebox.showinfo("Price Statistics", f"Average Price: {average_price:.2f}\nMaximum Price: {max_price}\nMinimum Price: {min_price}")


def register(window1):
    # Cria uma nova janela superior (Toplevel) para o cadastro.
    window2 = Toplevel(window1)
    window2.title("Sign up")
    window2.configure(bg="#ADD8E6")

    # Define as dimensões da janela de cadastro.
    width_window = 420
    height_window = 220

    # Obtém o tamanho da tela do usuário.
    width_screen = window2.winfo_screenwidth()
    height_screen = window2.winfo_screenheight()

    # Centraliza a janela no meio da tela.
    pos_x = (width_screen // 2) - (width_window // 2)
    pos_y = (height_screen // 2) - (height_window // 2)

    window2.geometry('{}x{}+{}+{}'.format(width_window, height_window, pos_x, pos_y))

    lbl1 = Label(window2, text="Create an account", font="Arial 14 bold", bg="#ADD8E6")
    lbl1.grid(row=0, column=0, columnspan=2, pady=10)

    lbl2 = Label(window2, text="Enter your username", font="Arial 12 bold", bg="#ADD8E6")
    lbl2.grid(row=1, column=0, sticky="e")

    lbl3 = Label(window2, text="Enter your password", font="Arial 12 bold", bg="#ADD8E6")
    lbl3.grid(row=2, column=0, sticky="e")

    # Cria campos de entrada para o nome de usuário e senha
    entry1 = Entry(window2, font="Arial 14")
    entry1.grid(row=1, column=1, pady=10, padx=10)

    entry2 = Entry(window2, show="*", font="Arial 14")
    entry2.grid(row=2, column=1, pady=10, padx=10)

    def save_user():
        new_user = entry1.get()
        new_password = entry2.get()

        # Configuração da conexão com o banco de dados
        dadosConexao = ("Driver={SQLite3 ODBC Driver};Server=localhost;Database=Projeto.db")

        # Cria uma nova conexão com o banco de dados
        conexao = pyodbc.connect(dadosConexao)

        # Cria um novo cursor para executar SQL na nova conexão
        cursor = conexao.cursor()

        # Executa uma operação SQL para inserir os dados no banco de dados.
        cursor.execute("INSERT INTO Usuarios (Nome, Senha) VALUES (?, ?)", (new_user, new_password))

        # Grava as alterações no banco de dados.
        conexao.commit()

        # Feche a janela de cadastro após a conclusão
        window2.destroy()

    #Botoes da tela de registro de novo usuario
    btn1 = Button(window2, text="Confirm", font=("Arial", 12), command=save_user)
    btn1.grid(row=3, column=1, columnspan=1, pady=10)

    btn2 = Button(window2, text="Cancel", font=("Arial", 12), command=window2.destroy)
    btn2.grid(row=3, column=0, columnspan=1, pady=10)

    # Define as colunas 0 e 1 para ocupar todo o espaço disponível
    window2.grid_columnconfigure(0, weight=1)
    window2.grid_columnconfigure(1, weight=1)

def user(username, password):
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
    window1 = Tk()
    window1.title("Sign in")

    def show_register_window():
        register(window1)

    def verify_login():
        username = entry1.get()
        password = entry2.get()

        # Validação das credenciais
        if user(username, password):
            window1.destroy()  # Feche a janela de login
            open_main_interface()  # Abra a nova janela

        else:
            # Se a condição "if" anterior não for atendida, ou seja, se as credenciais de login forem incorretas
            # Cria um rótulo de aviso com texto "Nome de Usuário ou Senha incorretos" em vermelho
            msg_lbl = Label(window1, text="Incorrect username or password.", fg="red")
            msg_lbl.grid(row=3,column=0,columnspan=2)

    # Configura janela principal
    window1.configure(bg="#00008B")
    width_window = 450
    height_window = 300

    # Obtém o tamanho da tela
    width_screen = window1.winfo_screenwidth()
    height_screen = window1.winfo_screenheight()

    # Calcula a posição para centralizar a janela na tela
    pos_x  = (width_screen // 2) - (width_window // 2)
    pos_y  = (height_screen // 2) - (height_window // 2)

    # Define as dimensões e posição da janela principal
    window1.geometry('{}x{}+{}+{}'.format(width_window, height_window, pos_x, pos_y))

    # Cria rótulos para título, nome de usuário e senha
    lbl1 = Label(window1, text="SysControl", font="Arial 20", bg="#00008B", fg="#ffffff")
    lbl1.grid(row=0,column=0,columnspan=2, pady=20) 

    lbl2 = Label(window1, text="Username", font="Arial 14 bold", bg="#00008B", fg="#ffffff")
    lbl2.grid(row=1,column=0,sticky="NSEW") #NSEW

    lbl3 = Label(window1, text="Password", font="Arial 14 bold", bg="#00008B", fg="#ffffff")
    lbl3.grid(row=2,column=0,sticky="NSEW") #NSEW

    # Cria campos de entrada para o nome de usuário e senha
    entry1 = Entry(window1, font="Arial 14")
    entry1.grid(row=1,column=1,pady=10)

    entry2 = Entry(window1, show="*",font="Arial 14")
    entry2.grid(row=2,column=1, pady=10)

    btn1 = Button(window1, text="Create an account", font="Arial 12", bg="#eeeeee", command=show_register_window)
    btn1.grid(row=4, column=0,columnspan=1, padx=20, pady=10, sticky="NSEW")

    btn2 = Button(window1, text="Login", font="Arial 12", bg="#eeeeee", command=verify_login)
    btn2.grid(row=4, column=1,columnspan=2, padx=20, pady=10, sticky="NSEW")

    btn3 = Button(window1, text="Exit", font="Arial 12", bg="#eeeeee", command=window1.destroy)
    btn3.grid(row=5, column=0,columnspan=2, padx=20, pady=10, sticky="NSEW")

    # Define a configuração de peso para as linhas e colunas da grade da tela de login
    for i in range(5):
        window1.grid_rowconfigure(i, weight=1)
        window1.grid_rowconfigure(i, weight=1)

    for i in range(2):
        window1.grid_columnconfigure(i, weight=1)
    
    window1.mainloop()

#Funçao que é responsavel pela abertura e gerencia da tela principal do programa
def open_main_interface():
    # Configuração da conexão com o banco de dados
    dadosConexao = ("Driver={SQLite3 ODBC Driver};Server=localhost;Database=Projeto.db")

    # Cria uma nova conexão com o banco de dados
    connection = pyodbc.connect(dadosConexao)

    # Cria um novo cursor para executar SQL na nova conexão
    cursor = connection.cursor()

    # Defina uma função de ação para o botão "New Product"
    def action1():
        interface.register_p()

    # Defina funções de ação para botões, como o botão de pesquisa
    def action2():
        interface.search_product(name.get(), description.get())

    def action3():
        interface.delete()

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
    menu_arquivo.add_command(label="Register",command=action1)

    #Cria opção menu Sair
    menu_arquivo.add_command(label="Exit", command=main_window.destroy)

    # Cria rótulos e campos de entrada para nome e descrição do produto
    Label(main_window, text="Search by", font="Arial 14", bg="#eeeeee").grid(row=0,column=1, padx=10, pady=10)

    Label(main_window, text="Product Name: ", 
          font="Arial 14", bg="#eeeeee").grid(row=0,column=2, padx=10, pady=10)
    name = Entry(main_window, font="Arial 14")
    name.grid(row=0,column=3, padx=10, pady=10)

    Label(main_window, text="Product Description: ", 
          font="Arial 14", bg="#eeeeee").grid(row=0,column=5, padx=10, pady=10)
    description = Entry(main_window, font="Arial 14")
    description.grid(row=0,column=6, padx=10, pady=10)

    Label(main_window, text="All Products", 
          font="Arial 18 bold", fg="black" , bg="#eeeeee").grid(row=2,column=0, columnspan=10, padx=10, pady=10)

    btn_save = Button(main_window, text="New Product", font="Arial 26", command=action1)
    btn_save.grid(row=4,column=0, columnspan=4, sticky="NSEW", padx=20, pady=5)

    # Cria um botão "Deletar" na janela que chama a função delete() quando clicado
    btn_delete = Button(main_window, text="Delete", font="Arial 26", command=action3)
    btn_delete.grid(row=4,column=4, columnspan=4, sticky="NSEW", padx=20, pady=5)

    # Cria uma instância da classe ProductInterface
    interface = product(main_window)
    
    # Chame o método para criar a TreeView
    interface.create()
    
    # Chame a função list_products na inicialização para preencher a TreeView
    interface.list()

    #Abre a instancia da função da barra de pesquisa
    name.bind('<KeyRelease>', lambda e: interface.filter(name, description))
    description.bind('<KeyRelease>', lambda e: interface.filter(name, description))

    # Mantém a janela principal aberta
    main_window.mainloop()

    # Feche a conexão com o banco de dados após fechar a janela principal
    connection.close()

# Execução da aplicação
show_login_window()