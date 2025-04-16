import customtkinter as ctk
import tkcalendar as tkcal
from time import strftime
from openpyxl import Workbook, load_workbook
from tkinter import messagebox

class App(ctk.CTk):
    
    # Classe principal da aplicação, herda de ctk.CTk
    def __init__(self, fg_color=None, **kwargs):
        super().__init__(fg_color, **kwargs)
        # Configurações iniciais da interface
        self.layoutConfiguracao()  
        # Configuração de aparência
        self.alterarAparencia()
        # Configuração do sistema (frames, labels, tabs, etc.)
        self.sistema()
        # Configuração da hora do sistema
        self.horaSistema()
        

    def layoutConfiguracao(self):
        # Define o tema padrão para o aplicativo (azul)
        ctk.set_default_color_theme('blue')
        # Define o título da janela
        self.title(string='Cadastramento de Clientes')
        # Define o tamanho inicial da janela
        self.geometry(geometry_string='900x620')
        # Probibir redimensionamento
        self.resizable(False, False)
        # Configura a expansão de colunas para melhor ajuste no layout
        for i in range(5):
            self.grid_columnconfigure(i, weight=2)

    # Método para alterar cores com base no modo (claro ou escuro)
    def alterarAparencia(self):
        
        # Função anônima para definir a cor de fundo com base no modo
        modoCor = lambda modo: '#EBEBEB' if modo == 'Light' else '#242424'

        # Verifica widgets específicos (por tipo) dentro da interface
        def verificaoWidget(objetoCTk: object) -> list:
            return [widget for widget in self.winfo_children() if isinstance(widget, objetoCTk)]

        # Lógica para alterar o modo de aparência
        def alterar():
            # Verifica o estado do botão switch e define o modo
            modo = 'Light' if butaoAlterarCor.get() else 'Dark'
            self._set_appearance_mode(modo)  # Define o modo na interface
            
            # Define cores baseadas no modo atual
            nova_cores = '#000' if butaoAlterarCor.get() else '#fff'

            # Configura o botão switch
            butaoAlterarCor.configure(
                text=modo,
                font=('Roboto', 12),
                text_color=nova_cores,
                bg_color=modoCor(modo=modo),
                
            )
            
            # Altera a cor de todos os labels na interface
            for widgetLabel in verificaoWidget(ctk.CTkLabel):
                widgetLabel.configure(
                    text_color=nova_cores, 
                    bg_color=modoCor(modo=modo)
                )
            
            # Altera a cor do primeiro frame encontrado
            frameAjuste = verificaoWidget(ctk.CTkFrame)
            frameAjuste[0].configure(bg_color=modoCor(modo=modo))
            
            # Configura o TabView e seus widgets
            tabViewAjuste = verificaoWidget(ctk.CTkTabview)
            if tabViewAjuste:
                tabViewAjuste[0].configure(
                    fg_color='#fff' if modo == 'Light' else '#2b2e2c',
                    bg_color=modoCor(modo=modo)
                )
                
                # Itera por todos os labels dentro das abas do TabView
                for tabviewConfigLabel in tabViewAjuste:
                    for infoLabel in tabviewConfigLabel.winfo_children(): # infoLabel -> Aba1 e Aba2
                        for label in infoLabel.winfo_children(): # Dentro da Aba1 e Aba2,  verifico os labels
                            if isinstance(label, ctk.CTkLabel): # Se pertencer ao objeto CTkLabel ->  configura a cor de texto
                                label.configure(text_color=nova_cores)
                                
    
        # Cria o botão switch para alternar entre claro e escuro
        butaoAlterarCor = ctk.CTkSwitch(
            self, 
            text='Tema', 
            font=('Roboto', 12),  
            command=alterar,
            button_hover_color='#c4003b'
        )
        # Posiciona o botão switch
        butaoAlterarCor.grid(row=5, column=0, sticky='w', padx=5, pady=5)
        
    def horaSistema(self):
        
        def atualizarHora():
            horaAtual =strftime('%H:%M:%S')
            labelHora.configure(text=horaAtual)
            labelHora.after(ms=1000, func=atualizarHora) # Altera o horário a cada 1 segundo
            
        labelHora = ctk.CTkLabel( #Inicia com um horário
            self, 
            text=strftime('%H:%M:%S'), 
            font=('Coolvetica', 28),
            )
        
        labelHora.grid(row=2, column=0, padx=15, pady=5, sticky='n')
        atualizarHora()
 
    # Método que organiza os widgets do sistema
    def sistema(self):
        
        # Função criar arquivo
        def criar_arquivo(caminho_excel):
                try:
                    arquivoExcel = load_workbook(caminho_excel) 
                except FileNotFoundError:
                    arquivoExcel = Workbook()
                    planilhaFisica = arquivoExcel.active
                    planilhaFisica.title = 'PessoaFisica'
                    planilhaJuridica = arquivoExcel.create_sheet('PessoaJuridica')
                    
                    planilhaFisica['A1'] = labelNome.cget(attribute_name='text')
                    planilhaFisica['B1'] = labelSobrenome.cget(attribute_name='text')
                    planilhaFisica['C1'] = labelSexo.cget(attribute_name='text')
                    planilhaFisica['D1'] = labelEmail.cget(attribute_name='text')
                    planilhaFisica['E1'] = labelSenha.cget(attribute_name='text')
                    planilhaFisica['F1'] = 'Data(Abertura)'
                    planilhaFisica['G1'] = labelCpf.cget(attribute_name='text')
                    planilhaFisica['H1'] = labelTelefone.cget(attribute_name='text')
                    planilhaFisica['I1'] = labelEndereco.cget(attribute_name='text')
                    
                    planilhaJuridica['A1'] = labelNomeEmpresa.cget(attribute_name='text')
                    planilhaJuridica['B1'] = labelNomeFantasia.cget(attribute_name='text')
                    planilhaJuridica['C1'] = labelTipoEmpresa.cget(attribute_name='text')
                    planilhaJuridica['D1'] = labelEmailEmpresa.cget(attribute_name='text')
                    planilhaJuridica['E1'] = labelSenhaEmpresa.cget(attribute_name='text')
                    planilhaJuridica['F1'] = 'Data(Abertura)'
                    planilhaJuridica['G1'] = labelCnpj.cget(attribute_name='text')
                    planilhaJuridica['H1'] = labelTelefoneEmpresa.cget(attribute_name='text')
                    planilhaJuridica['I1'] = labelEnderecoEmpresa.cget(attribute_name='text')
                    
                    arquivoExcel.save(caminho_excel)
                
                return arquivoExcel
            
        # Frame principal do sistema
        frame = ctk.CTkFrame(
            self, 
            width=550,
            height=100,
            fg_color='#c4003b',  # Cor de fundo do frame
        )
        
        # Label com título do sistema
        labelSistema = ctk.CTkLabel(
            frame, 
            text='Sistema de Cadastro de Clientes', 
            text_color='#fff',  # Cor do texto
            font=('Coolvetica', 36),  # Fonte e tamanho
            bg_color='transparent',  # Cor de fundo transparente
            height=60
        )
        
        # Configuração da expansão da coluna para centralizar
        frame.grid_columnconfigure(0, weight=1)
        # Posiciona o título no frame
        labelSistema.grid(row=0, column=0, padx=10, pady=10)
        # Posiciona o frame no grid principal
        frame.grid(row=0, column=0, columnspan=6, padx=10, pady=10, sticky='nsew')

        # Label "Cadastra-se" abaixo do frame principal
        labelCadastro = ctk.CTkLabel(
            self, 
            text='Cadastra-se',
            text_color='#fff',  # Cor do texto
            font=('Coolvetica', 32),  # Fonte e tamanho
            bg_color='transparent',  # Cor de fundo
            anchor='w'  # Alinhamento à esquerda
        )
        # Posiciona o label no grid
        labelCadastro.grid(row=1, column=2, padx=10, pady=10, sticky='nw')
        
        # Cria o TabView para as abas de cadastro
        tabsView = ctk.CTkTabview(
            self, 
            width=800, 
            height=400, 
            border_width=2,  # Largura da borda
            segmented_button_selected_color='#c4003b',  # Cor da aba selecionada
            segmented_button_selected_hover_color='#730732',  # Cor ao passar o mouse
            border_color='#c4003b',  # Cor da borda
            segmented_button_fg_color='#000'
        )
        # Posiciona o TabView no grid
        tabsView.grid(row=2, column=0, columnspan=5, padx=5, pady=5)

        # Adiciona abas no TabView
        tabsView.add('Pessoa Física')
        tabsView.add('Pessoa Jurídica')
        
        # Adiciona labels dentro da aba "Pessoa Física"
        
        # Linha 0 e 1
        # Coluna 0 e 3
        # Nome
        labelNome = ctk.CTkLabel(tabsView.tab('Pessoa Física'), text='Nome', font=('Roboto', 16))
        labelNome.grid(row=0, column=0, padx=15, pady=5, sticky='nw')
        
        entradaNome = ctk.CTkEntry(tabsView.tab('Pessoa Física'), placeholder_text='Digite o seu Nome', width=250)
        entradaNome.grid(row=1, column=0, columnspan=2, padx=10, pady=5, sticky='nw')
        
        # Sobrenome
        labelSobrenome = ctk.CTkLabel(tabsView.tab('Pessoa Física'), text='Sobrenome', font=('Roboto', 16))
        labelSobrenome.grid(row=0, column=3, padx=15, pady=5, sticky='nw')
        
        entradaSobrenome = ctk.CTkEntry(tabsView.tab('Pessoa Física'), placeholder_text='Digite o seu Sobrenome', width=250)
        entradaSobrenome.grid(row=1, column=3, columnspan=2, padx=10, pady=5, sticky='nw')
        
        # Opção de Menu
        labelSexo = ctk.CTkLabel(tabsView.tab('Pessoa Física'), text='Sexo', font=('Roboto', 16))
        labelSexo.grid(row=0, column=5, padx=15, pady=5, sticky='nw')
        
        opcaoMenuSexo = ctk.CTkOptionMenu(
            tabsView.tab('Pessoa Física'), 
            values=['Masculino', 'Feminino', 'Outros'], 
            bg_color='transparent', 
            fg_color='#c4003b', 
            dropdown_hover_color='#c4003b'
            )
        
        opcaoMenuSexo.grid(row=1, column=5, padx=10, pady=5, sticky='nw')
        
        # Linha 2 e 3
        # Coluna 0 e 3
        # Email
        labelEmail = ctk.CTkLabel(tabsView.tab('Pessoa Física'), text='E-mail', font=('Roboto', 16))
        labelEmail.grid(row=2, column=0, padx=15, pady=5, sticky='nw')
        
        entradaEmail = ctk.CTkEntry(tabsView.tab('Pessoa Física'), placeholder_text='Digite o seu E-mail', width=250)
        entradaEmail.grid(row=3, column=0, columnspan=2, padx=10, pady=5, sticky='nw')
        
        # Senha
        labelSenha = ctk.CTkLabel(tabsView.tab('Pessoa Física'), text='Senha', font=('Roboto', 16))
        labelSenha.grid(row=2, column=3, padx=15, pady=5, sticky='nw')
        
        entradaSenha = ctk.CTkEntry(tabsView.tab('Pessoa Física'), placeholder_text='Digite a sua Senha', show='*', width=250)
        entradaSenha.grid(row=3, column=3, columnspan=2, padx=10, pady=5, sticky='nw')
        
        #data
        dataAbertura = ctk.CTkLabel(tabsView.tab('Pessoa Física'), text='Data(Abertura)', font=('Roboto', 16))
        dataAbertura.grid(row=2, column=5, padx=15, pady=5, sticky='nw')
        
        calendario = tkcal.DateEntry(tabsView.tab('Pessoa Física'), font=('Roboto', 12), width=15)
        calendario.grid(row=3, column=5, padx=15, pady=5, sticky='nw')
    
        # Linha 4 e 5
        # Coluna 0 e 3
        # CPF
        labelCpf = ctk.CTkLabel(tabsView.tab('Pessoa Física'), text='CPF', font=('Roboto', 16))
        labelCpf.grid(row=4, column=0, padx=15, pady=5, sticky='nw')
        
        entradaCPF = ctk.CTkEntry(tabsView.tab('Pessoa Física'), placeholder_text='Informe apenas dígitos', width=250)
        entradaCPF.grid(row=5, column=0, columnspan=2, padx=10, pady=5, sticky='nw')
        
        # telefone
        labelTelefone = ctk.CTkLabel(tabsView.tab('Pessoa Física'), text='Telefone', font=('Roboto', 16))
        labelTelefone.grid(row=4, column=3, padx=15, pady=5, sticky='nw')
        
        entradaTelefone= ctk.CTkEntry(tabsView.tab('Pessoa Física'), placeholder_text='Telefone', width=250)
        entradaTelefone.grid(row=5, column=3, columnspan=2, padx=10, pady=5, sticky='nw')
        
        # Endereço
        labelEndereco = ctk.CTkLabel(tabsView.tab('Pessoa Física'), text='Endereço', font=('Roboto', 16))
        labelEndereco.grid(row=4, column=5, padx=15, pady=5, sticky='nw')
        
        entradaEndereco = ctk.CTkEntry(tabsView.tab('Pessoa Física'), placeholder_text='Digite o seu Endereço', width=250)
        entradaEndereco.grid(row=5, column=5, columnspan=2, padx=10, pady=5, sticky='nw')
        
        # Pessoa física
        # Botões de Ação - Importante para executar as principais funcionalidades
        
        # Função limpar
        def limpar_camposFisica():
            for widgetTabview in self.winfo_children():
                if isinstance(widgetTabview, ctk.CTkTabview):
                    abaFisica = widgetTabview.winfo_children()[0]
                    for entry in abaFisica.winfo_children():
                        if isinstance(entry, ctk.CTkEntry):
                            entry.delete(0, ctk.END)
                            entradaSenha.configure(show='*') # Segurança -  Manter a proteção das informações inseridas no campo senha
                     
        # Botão Limpar campos               
        botaoLimparFisica = ctk.CTkButton(
            tabsView.tab('Pessoa Física'), 
            text='Limpar Campos', 
            font=('Roboto', 16),
            height=40,
            width=250,
            text_color='#fff',
            bg_color='transparent',
            fg_color='#c4003b',
            hover_color='#730732',
            command=limpar_camposFisica
            )
        
        botaoLimparFisica.grid(row=6, column=0, padx=15, pady=15, sticky='sw')
         
         # Função enviar dados
        def enviar_dadosFisica():
            
            def cadastrarClientes(caminhoArquivoExcel, *args):
                arquivoExcel = load_workbook(caminhoArquivoExcel)
                planilha = arquivoExcel['PessoaFisica'] # ativa  nossa planilha
                novaLinha = planilha.max_row # pega número da última linha
                
                colunas = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']
                for i, info in enumerate(args):
                    if i < len(colunas):  # Verifica se o número de argumentos não excede as colunas
                        planilha[f'{colunas[i]}{novaLinha + 1}'] = info
                    
                arquivoExcel.save(caminhoArquivoExcel)
         
            nome = entradaNome.get()
            sobrenome = entradaSobrenome.get()
            sexo = opcaoMenuSexo.get()
            email = entradaEmail.get()
            senha = entradaSenha.get()
            data = calendario.get_date()
            cpf = entradaCPF.get()
            telefone = entradaTelefone.get()
            endereco = entradaEndereco.get()

            # Chama a função para criar ou abrir o arquivo
            caminho_arquivo = 'cadastroClientes.xlsx'
            criar_arquivo(caminho_arquivo) # Validação que o nosso arquivo existe
            
            cadastrarClientes(caminho_arquivo, nome, sobrenome, sexo, email, senha, data, cpf, telefone, endereco)
            
            messagebox.showinfo(title='Mensagem sobre Cadastro', message='Cliente cadastrado com sucesso!')          
            
        # Botão Enviar
        botaoEnviarFisica = ctk.CTkButton(
            tabsView.tab('Pessoa Física'), 
            text='Enviar', 
            font=('Roboto', 16),
            height=40,
            width=250,
            text_color='#fff',
            bg_color='transparent',
            fg_color='#02ad10',
            hover_color='#730732',
            command=enviar_dadosFisica
            )
        
        botaoEnviarFisica.grid(row=6, column=3, padx=15, pady=15, sticky='sw')
    
        # Pessoa Jurídica TODO
        
         # Linha 0 e 1
        # Coluna 0 e 3
        # Nome Empresa
        labelNomeEmpresa = ctk.CTkLabel(tabsView.tab('Pessoa Jurídica'), text='Empresa', font=('Roboto', 16))
        labelNomeEmpresa.grid(row=0, column=0, padx=15, pady=5, sticky='nw')
        
        entradaNomeEmpresa = ctk.CTkEntry(tabsView.tab('Pessoa Jurídica'), placeholder_text='Nome Empresarial', width=250)
        entradaNomeEmpresa.grid(row=1, column=0, columnspan=2, padx=10, pady=5, sticky='nw')
        
        # TÍTULO DO ESTABELECIMENTO ( NOME FANTASIA)
        labelNomeFantasia = ctk.CTkLabel(tabsView.tab('Pessoa Jurídica'), text='Nome Fantasia', font=('Roboto', 16))
        labelNomeFantasia.grid(row=0, column=3, padx=15, pady=5, sticky='nw')
        
        entradaNomeFantasia = ctk.CTkEntry(tabsView.tab('Pessoa Jurídica'), placeholder_text='Informe o Título do Estabelecimento', width=250)
        entradaNomeFantasia.grid(row=1, column=3, columnspan=2, padx=10, pady=5, sticky='nw')
        
        # Opção de Menu
        labelTipoEmpresa = ctk.CTkLabel(tabsView.tab('Pessoa Jurídica'), text='Tipo de Empresa', font=('Roboto', 16))
        labelTipoEmpresa.grid(row=0, column=5, padx=15, pady=5, sticky='nw')
        
        opcaoMenuTipoEmpresa = ctk.CTkOptionMenu(
            tabsView.tab('Pessoa Jurídica'), 
            values=['Microempreendedor Individual (MEI)', 'Microempresa (ME)', 'Sociedade Anônima (S/A)'], 
            bg_color='transparent', 
            fg_color='#c4003b', 
            dropdown_hover_color='#c4003b'
            )
        
        opcaoMenuTipoEmpresa.grid(row=1, column=5, padx=10, pady=5, sticky='nw')
        
        # Linha 2 e 3
        # Coluna 0 e 3
        # Email Empresarial
        labelEmailEmpresa = ctk.CTkLabel(tabsView.tab('Pessoa Jurídica'), text='E-mail Empresarial', font=('Roboto', 16))
        labelEmailEmpresa.grid(row=2, column=0, padx=15, pady=5, sticky='nw')
        
        entradaEmailEmpresa = ctk.CTkEntry(tabsView.tab('Pessoa Jurídica'), placeholder_text='Digite o E-mail da Empresa ', width=250)
        entradaEmailEmpresa.grid(row=3, column=0, columnspan=2, padx=10, pady=5, sticky='nw')
        
        # Senha
        labelSenhaEmpresa= ctk.CTkLabel(tabsView.tab('Pessoa Jurídica'), text='Senha', font=('Roboto', 16))
        labelSenhaEmpresa.grid(row=2, column=3, padx=15, pady=5, sticky='nw')
        
        entradaSenhaEmpresa = ctk.CTkEntry(tabsView.tab('Pessoa Jurídica'), placeholder_text='Digite a sua Senha', show='*', width=250)
        entradaSenhaEmpresa.grid(row=3, column=3, columnspan=2, padx=10, pady=5, sticky='nw')
        
        #data
        dataAberturaEmpresa = ctk.CTkLabel(tabsView.tab('Pessoa Jurídica'), text='Data(Abertura)', font=('Roboto', 16))
        dataAberturaEmpresa.grid(row=2, column=5, padx=15, pady=5, sticky='nw')
        
        # Calendario 
        calendarioEmpresa = tkcal.DateEntry(tabsView.tab('Pessoa Jurídica'), font=('Roboto', 12), width=15)
        calendarioEmpresa.grid(row=3, column=5, padx=15, pady=5, sticky='nw')
    
        # Linha 4 e 5
        # Coluna 0 e 3
        # CNPJ
        labelCnpj = ctk.CTkLabel(tabsView.tab('Pessoa Jurídica'), text='CNPJ', font=('Roboto', 16))
        labelCnpj.grid(row=4, column=0, padx=15, pady=5, sticky='nw')
        
        entradaCnpj = ctk.CTkEntry(tabsView.tab('Pessoa Jurídica'), placeholder_text='Informe apenas dígitos', width=250)
        entradaCnpj.grid(row=5, column=0, columnspan=2, padx=10, pady=5, sticky='nw')
        
        # telefone
        labelTelefoneEmpresa = ctk.CTkLabel(tabsView.tab('Pessoa Jurídica'), text='Telefone', font=('Roboto', 16))
        labelTelefoneEmpresa.grid(row=4, column=3, padx=15, pady=5, sticky='nw')
        
        entradaTelefoneEmpresa= ctk.CTkEntry(tabsView.tab('Pessoa Jurídica'), placeholder_text='Telefone', width=250)
        entradaTelefoneEmpresa.grid(row=5, column=3, columnspan=2, padx=10, pady=5, sticky='nw')
        
        # Endereço
        labelEnderecoEmpresa = ctk.CTkLabel(tabsView.tab('Pessoa Jurídica'), text='Endereço', font=('Roboto', 16))
        labelEnderecoEmpresa.grid(row=4, column=5, padx=15, pady=5, sticky='nw')
        
        entradaEnderecoEmpresa = ctk.CTkEntry(tabsView.tab('Pessoa Jurídica'), placeholder_text='Digite o Endereço do Estabelecimento', width=250)
        entradaEnderecoEmpresa.grid(row=5, column=5, columnspan=2, padx=10, pady=5, sticky='nw')
        
        # Botões de Ação - Importante para executar as principais funcionalidades
        
        # Função Limpar
        def limpar_camposJuridica():
            for widgetTabview in self.winfo_children():
                if isinstance(widgetTabview, ctk.CTkTabview):
                    abaJuridica = widgetTabview.winfo_children()[1]
                    for entry in abaJuridica.winfo_children():
                        if isinstance(entry, ctk.CTkEntry):
                            entry.delete(0, ctk.END)
                            entradaSenhaEmpresa.configure(show='*') # Segurança -  Manter a proteção das informações inseridas no campo senha
                                    
        # Botão limmpar               
        botaoLimparJuridica = ctk.CTkButton(
            tabsView.tab('Pessoa Jurídica'), 
            text='Limpar Campos', 
            font=('Roboto', 16),
            height=40,
            width=250,
            text_color='#fff',
            bg_color='transparent',
            fg_color='#c4003b',
            hover_color='#730732',
            command=limpar_camposJuridica
            )
        
        botaoLimparJuridica.grid(row=6, column=0, padx=15, pady=15, sticky='sw')
    
        # Função Enviar dados
        def enviar_dadosJuridica():
        
            def cadastrarClientes(caminhoArquivoExcel, *args):
                arquivoExcel = load_workbook(caminhoArquivoExcel)
                planilha = arquivoExcel['PessoaJuridica'] # ativa  nossa planilha
                novaLinha = planilha.max_row # pega número da última linha
                
                colunas = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I']
                for i, info in enumerate(args):
                    if i < len(colunas):  # Verifica se o número de argumentos não excede as colunas
                        planilha[f'{colunas[i]}{novaLinha + 1}'] = info
                    
                arquivoExcel.save(caminhoArquivoExcel)
         
            nomeEmpresa = entradaNomeEmpresa.get()
            nomeFantasia = entradaNomeFantasia.get()
            nomeTipoEmpresa = opcaoMenuTipoEmpresa.get()
            emailEmpresa = entradaEmailEmpresa.get()
            senhaEmpresa = entradaSenhaEmpresa.get()
            dataEmpresa = calendarioEmpresa.get_date()
            cnpj = entradaCnpj.get()
            telefoneEmpresa = entradaTelefoneEmpresa.get()
            enderecoEmpresa = entradaEnderecoEmpresa.get()

            # Chama a função para criar ou abrir o arquivo
            caminho_arquivo = 'cadastroClientes.xlsx'
            criar_arquivo(caminho_arquivo) # Validação que o nosso arquivo existe
            
            cadastrarClientes(caminho_arquivo, nomeEmpresa, nomeFantasia, nomeTipoEmpresa,  emailEmpresa, senhaEmpresa, dataEmpresa, cnpj, telefoneEmpresa, enderecoEmpresa)
            
            messagebox.showinfo(title='Mensagem sobre Cadastro', message='Cliente cadastrado com sucesso!')         
        botaoEnviarJuridica = ctk.CTkButton(
            tabsView.tab('Pessoa Jurídica'), 
            text='Enviar', 
            font=('Roboto', 16),
            height=40,
            width=250,
            text_color='#fff',
            bg_color='transparent',
            fg_color='#02ad10',
            hover_color='#730732',
            command=enviar_dadosJuridica
            )
        
        botaoEnviarJuridica.grid(row=6, column=3, padx=15, pady=15, sticky='sw')
        
        
        # Botão de fechar
        def fechar():
            self.quit()
        
        for cadastro in ['Pessoa Física', 'Pessoa Jurídica']:
            botaoFechar = ctk.CTkButton(
                tabsView.tab(cadastro), 
                text='Fechar', 
                font=('Roboto', 16),
                height=30,
                width=100,
                text_color='#fff',
                bg_color='transparent',
                fg_color='#c4003b',
                hover_color='#730732',
                command=fechar
                )
            
            botaoFechar.grid(row=7, column=5, padx=10, pady=15, sticky='se')

# Inicializa o aplicativo
if __name__ == '__main__':
    app = App()
    app.mainloop()
