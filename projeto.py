# update for linguage detection
import os
import tkinter as tk
from tkinter import messagebox, ttk
import pandas as pd
from datetime import datetime, timedelta
import customtkinter as ctk
import matplotlib.pyplot as plt
import matplotlib.backends.backend_tkagg as tkagg
import winsound
import logging
from PIL import Image, ImageTk
import threading
from tkinter.font import Font
import sys

# Caminho seguro para salvar o log no mesmo local do .exe ou .py
if getattr(sys, 'frozen', False):
    app_path = os.path.dirname(sys.executable)
else:
    app_path = os.path.dirname(__file__)

log_path = os.path.join(app_path, "sistema_inventario.log")

import logging
logging.basicConfig(
    filename=log_path,
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    filemode='a'
)

class SistemaInventario:
    def __init__(self):
        ctk.set_appearance_mode("light")
        ctk.set_default_color_theme("blue")
        
        # Variáveis de controle
        self.codigo_registrado = {}
        self.todos_codigos_registrados = set()
        self.usuario_logado = None
        self.planta_selecionada = None
        self.pasta_projeto = os.path.dirname(os.path.abspath(__file__))
        
        # Bancos de dados
        self.arquivo_excel_campinas = os.path.join(self.pasta_projeto, "Campinas", "dados_produtos_campinas.xlsx")
        self.arquivo_excel_mafra = os.path.join(self.pasta_projeto, "MAFRA", "dados_produtos_mafra.xlsx")
        self.arquivo_excel_atual = None
        
        # Arquivos de produção/retrabalho
        self.arquivo_producao = os.path.join(self.pasta_projeto, "PRODUCAO.xlsx")
        self.arquivo_retrabalho = os.path.join(self.pasta_projeto, "RETRABALHO.xlsx")
        self.df_producao = None
        self.df_retrabalho = None
        
        # Elementos da interface
        self.root_menu = None
        self.tree = None
        self.etiqueta_palet_atual = None
        self.registros_temporarios = []
        self.tree_consulta = None
        self.frame_graficos = None
        
        # Novas variáveis para a visualização de leitura
        self.frame_visualizacao = None
        self.frame_colunas_codigos = [None, None]  # Frames para cada coluna
        self.labels_codigos = [[], []]  # Lista de labels para cada coluna
        self.codigos_a_ler = []  # Lista de códigos que precisam ser lidos
        self.codigos_lidos = set()  # Conjunto de códigos já lidos
        
        # Configurações
        self.leitura_continua = False
        self.auto_save_interval = 30
        self.tipo_etiqueta = None
        
        # Cache de dados
        self._data_cache = None
        self._last_update = None
        self._cache_expiry = timedelta(minutes=5)
        
        # Usuários e plantas
        self.usuarios = {
            "operador": {"senha": "123", "planta": "AMBAS"},
            "Operadormafra": {"senha": "456", "planta": "AMBAS"},
            "operadorCampinas": {"senha": "789", "planta": "AMBAS"},
            "operador1": {"senha": "111", "planta": "BANDAG"},
            "operador2": {"senha": "222", "planta": "MAFRA"}
        }
        
        self.inicializar_excel()
        self.carregar_dados_externos()
        self.tela_login()

    # ==============================================
    # MÉTODOS DE INICIALIZAÇÃO E CARREGAMENTO DE DADOS
    # ==============================================
    
    def carregar_dados_externos(self):
        try:
            if os.path.exists(self.arquivo_producao):
                self.df_producao = pd.read_excel(
                    self.arquivo_producao,
                    usecols=[0, 1],
                    header=None,
                    names=["PROGRESSIVA", "PALET"],
                    dtype=str
                )
                self.df_producao = self.df_producao.dropna()
                self.df_producao['PROGRESSIVA'] = self.df_producao['PROGRESSIVA'].astype(str).str.strip().str.upper()
                self.df_producao['PALET'] = self.df_producao['PALET'].astype(str).str.strip().str.upper()
            
            if os.path.exists(self.arquivo_retrabalho):
                self.df_retrabalho = pd.read_excel(
                    self.arquivo_retrabalho,
                    usecols=[0, 1],
                    header=None,
                    names=["PROGRESSIVA", "PALET"],
                    dtype=str
                )
                self.df_retrabalho = self.df_retrabalho.dropna()
                self.df_retrabalho['PROGRESSIVA'] = self.df_retrabalho['PROGRESSIVA'].astype(str).str.strip().str.upper()
                self.df_retrabalho['PALET'] = self.df_retrabalho['PALET'].astype(str).str.strip().str.upper()
                
        except Exception as e:
            logging.error(f"Erro ao carregar dados externos: {str(e)}")
            messagebox.showerror("Erro", f"Falha ao ler arquivos de produção/retrabalho: {str(e)}")

    def inicializar_excel(self):
        for planta, arquivo in [("Campinas", self.arquivo_excel_campinas), ("MAFRA", self.arquivo_excel_mafra)]:
            pasta = os.path.dirname(arquivo)
            if not os.path.exists(pasta):
                os.makedirs(pasta)
                
            if not os.path.exists(arquivo):
                colunas = ["Etiqueta Palet", "Codigo", "Tipo", "Repetida", 
                          "Data", "Hora", "Excel", "Usuario", "Local"]
                df = pd.DataFrame(columns=colunas)
                df.to_excel(arquivo, index=False, engine="openpyxl")

    def definir_arquivo_excel_atual(self):
        if self.planta_selecionada == "Campinas":
            self.arquivo_excel_atual = self.arquivo_excel_campinas
        elif self.planta_selecionada == "MAFRA":
            self.arquivo_excel_atual = self.arquivo_excel_mafra
        else:
            self.arquivo_excel_atual = None
            messagebox.showerror("Erro", "Planta não selecionada corretamente!")
        
        logging.info(f"Arquivo Excel atual definido para: {self.arquivo_excel_atual}")

    # ==============================================
    # MÉTODOS DE INTERFACE - LOGIN E TELA PRINCIPAL
    # ==============================================
    
    def tela_login(self):
        self.login_window = ctk.CTk()
        self.login_window.title("Multinacional - Login")
        self.login_window.geometry("800x600")
        self.login_window.configure(fg_color="white")
        
        frame_login = ctk.CTkFrame(self.login_window, fg_color="white")
        frame_login.pack(pady=50, padx=100, fill="both", expand=True)
        
        lbl_titulo = ctk.CTkLabel(
            frame_login, 
            text="Sistema Inventário", 
            font=ctk.CTkFont(size=28, weight="bold"),
            text_color="#002b5c"
        )
        lbl_titulo.pack(pady=(20, 10))
        
        lbl_subtitulo = ctk.CTkLabel(
            frame_login, 
            text="Multinacional", 
            font=ctk.CTkFont(size=16),
            text_color="#666666"
        )
        lbl_subtitulo.pack(pady=(0, 30))
        
        self.entry_usuario = ctk.CTkEntry(
            frame_login,
            placeholder_text="Usuário",
            width=300,
            height=40,
            font=ctk.CTkFont(size=14),
            corner_radius=10
        )
        self.entry_usuario.pack(pady=10)
        
        self.entry_senha = ctk.CTkEntry(
            frame_login,
            placeholder_text="Senha",
            show="*",
            width=300,
            height=40,
            font=ctk.CTkFont(size=14),
            corner_radius=10
        )
        self.entry_senha.pack(pady=10)
        
        self.combo_planta = ctk.CTkComboBox(
            frame_login,
            values=["Selecione..."],
            width=300,
            height=40,
            state="disabled",
            corner_radius=10
        )
        self.combo_planta.pack(pady=10)
        
        btn_login = ctk.CTkButton(
            frame_login,
            text="Entrar",
            command=self.validar_login,
            width=300,
            height=40,
            font=ctk.CTkFont(size=14, weight="bold"),
            fg_color="#002b5c",
            hover_color="#004080",
            corner_radius=10
        )
        btn_login.pack(pady=20)
        
        self.entry_usuario.bind("<KeyRelease>", self.atualizar_plantas_usuario)
        self.entry_senha.bind("<Return>", lambda e: self.validar_login())
        
        self.login_window.mainloop()

    def atualizar_plantas_usuario(self, event=None):
        usuario = self.entry_usuario.get().strip()
        
        if usuario in self.usuarios:
            planta_usuario = self.usuarios[usuario]["planta"]
            
            if planta_usuario == "AMBAS":
                self.combo_planta.configure(values=["Campinas", "MAFRA"], state="readonly")
                self.combo_planta.set("Selecione...")
            else:
                self.combo_planta.configure(values=[planta_usuario], state="readonly")
                self.combo_planta.set(planta_usuario)
        else:
            self.combo_planta.configure(values=["Selecione..."], state="disabled")

    def validar_login(self):
        usuario = self.entry_usuario.get().strip()
        senha = self.entry_senha.get().strip()
        planta = self.combo_planta.get()

        if usuario in self.usuarios and self.usuarios[usuario]["senha"] == senha:
            planta_usuario = self.usuarios[usuario]["planta"]
            
            if planta_usuario == "AMBAS" or planta_usuario == planta:
                self.usuario_logado = usuario
                self.planta_selecionada = planta
                self.definir_arquivo_excel_atual()
                winsound.Beep(1000, 200)
                logging.info(f"Usuário {usuario} logado na planta {planta}")
                self.login_window.destroy()
                self.tela_inicial()
            else:
                winsound.Beep(400, 500)
                messagebox.showerror("Erro", f"Usuário não tem acesso à planta {planta}!")
        else:
            winsound.Beep(400, 500)
            messagebox.showerror("Erro", "Credenciais inválidas!")
            self.entry_senha.delete(0, 'end')

    def tela_inicial(self):
        self.root_menu = tk.Tk()
        self.root_menu.title(f"Campinas - {self.usuario_logado} - {self.planta_selecionada}")
        self.root_menu.state('zoomed')
        
        header = tk.Frame(self.root_menu, bg="#002b5c", height=70)
        header.pack(fill="x")
        
        btn_voltar = tk.Button(
            header,
            text="← Sair",
            command=self.voltar_login,
            bg="#002b5c",
            fg="white",
            font=("Arial", 12),
            bd=0,
            relief="flat"
        )
        btn_voltar.pack(side="left", padx=10, pady=5)
        
        tk.Label(
            header,
            text=f"Multinacional - {self.planta_selecionada}",
            font=("Arial", 22, "bold"),
            fg="white",
            bg="#002b5c"
        ).pack(side="left", expand=True, pady=10)
        
        notebook = ttk.Notebook(self.root_menu)
        notebook.pack(fill="both", expand=True)
        
        self.criar_aba_leitura(notebook)
        self.criar_aba_consulta(notebook)
        self.criar_aba_analise(notebook)
        
        self.configurar_atalhos()
        self.agendar_auto_save()
        self.root_menu.mainloop()

    def voltar_login(self):
        if messagebox.askyesno("Confirmar", "Deseja realmente sair?"):
            self.root_menu.destroy()
            self.usuario_logado = None
            self.planta_selecionada = None
            self.tela_login()

    def configurar_atalhos(self):
        self.root_menu.bind('<F1>', lambda e: self.entry_palete.focus())
        self.root_menu.bind('<F2>', lambda e: self.entry_progressiva.focus())
        self.root_menu.bind('<F3>', lambda e: self.fechar_palete())
        self.root_menu.bind('<F5>', lambda e: self.limpar_campos())
        self.root_menu.bind('<F8>', lambda e: self.toggle_leitura_continua())

    # ==============================================
    # MÉTODOS DA ABA DE LEITURA
    # ==============================================
    
    def criar_aba_leitura(self, notebook):
        aba = tk.Frame(notebook, bg="white")
        notebook.add(aba, text="📷 Leitura")
        
        # Frame superior com controles
        top_frame = tk.Frame(aba, bg="white")
        top_frame.pack(fill="x", padx=5, pady=5)
        
        sidebar = tk.Frame(top_frame, bg="#f0f0f0", width=200)
        sidebar.pack(side="left", fill="y", padx=2, pady=2)
        
        tk.Label(sidebar, text="ETIQUETA PALET:", bg="#f0f0f0", font=("Arial", 10, "bold")).pack(pady=(10, 2))
        self.entry_palete = tk.Entry(sidebar, font=("Arial", 12), justify="center")
        self.entry_palete.pack(fill="x", padx=5, pady=2)
        self.entry_palete.bind("<KeyRelease>", lambda e: self.processar_etiqueta("PALETE"))
        
        tk.Label(sidebar, text="PROGRESSIVAS:", bg="#f0f0f0", font=("Arial", 10, "bold")).pack(pady=(10, 2))
        self.entry_progressiva = tk.Entry(sidebar, font=("Arial", 12), justify="center", state="disabled")
        self.entry_progressiva.pack(fill="x", padx=5, pady=2)
        self.entry_progressiva.bind("<KeyRelease>", lambda e: self.processar_etiqueta("PROGRESSIVA"))
        
        self.btn_fechar = tk.Button(
            sidebar,
            text="FECHAR PALETE",
            command=self.fechar_palete,
            bg="#002b5c",
            fg="white",
            font=("Arial", 10, "bold"),
            state="disabled"
        )
        self.btn_fechar.pack(fill="x", pady=10, padx=5)
        
        tk.Button(
            sidebar,
            text="LIMPAR CAMPOS ",
            command=self.limpar_campos,
            bg="#6c757d",
            fg="white",
            font=("Arial", 10, "bold")
        ).pack(fill="x", pady=5, padx=5)
        
        self.lbl_contador = tk.Label(
            sidebar,
            text="0",
            bg="#f0f0f0",
            font=("Arial", 16, "bold"),
            fg="#002b5c"
        )
        self.lbl_contador.pack(pady=20)
        
        # Treeview para exibir os registros
        tree_frame = tk.Frame(top_frame)
        tree_frame.pack(side="left", fill="both", expand=True)
        
        cols = ["Etiqueta Palet", "Código", "Tipo", "Repetida", "Data", "Hora", "Excel", "Usuário"]
        self.tree = ttk.Treeview(
            tree_frame,
            columns=cols,
            show="headings",
            height=10
        )
        
        for col in cols:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=120, anchor="center")
        
        # Cores mais suaves para as linhas
        self.tree.tag_configure('repetida', background='#ffffcc')  # Amarelo claro
        self.tree.tag_configure('nao_apontada', background='#ffe6e6')  # Vermelho claro
        self.tree.tag_configure('repetida_nao_apontada', background='#ffdd99')  # Laranja claro
        
        scroll = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree.yview)
        scroll.pack(side="right", fill="y")
        self.tree.configure(yscrollcommand=scroll.set)
        self.tree.pack(fill="both", expand=True)
        
        # Frame para a visualização dos códigos a serem lidos
        self.frame_visualizacao = tk.Frame(aba, bg="white")
        self.frame_visualizacao.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Cria as duas colunas para exibição dos códigos
        self.criar_visualizacao_codigos()

    def criar_visualizacao_codigos(self):
        """Cria a visualização dos códigos em duas colunas com estilo melhorado"""
        # Limpa o frame se já existir conteúdo
        for widget in self.frame_visualizacao.winfo_children():
            widget.destroy()
        
        # Configurações de estilo
        bg_color = "#f8f9fa"  # Cor de fundo mais clara
        fg_color = "#212529"  # Cor do texto mais escura
        font_size = 12  # Tamanho maior da fonte
        item_height = 50  # Altura maior para cada item
        item_width = 30  # Largura maior para cada item
        radius = 15  # Raio maior para bordas arredondadas
        
        # Frame principal que contém as duas colunas
        main_container = tk.Frame(self.frame_visualizacao, bg=bg_color)
        main_container.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Cria as duas colunas
        self.frame_colunas_codigos = [None, None]
        self.labels_codigos = [[], []]
        
        # Coluna esquerda
        col_left = tk.Frame(main_container, bg=bg_color)
        col_left.pack(side="left", fill="both", expand=True, padx=5)
        
        # Título da coluna esquerda
        lbl_left_title = tk.Label(
            col_left, 
            text="CÓDIGOS A LER", 
            font=("Arial", 12, "bold"), 
            bg="#002b5c", 
            fg="white",
            padx=10,
            pady=5
        )
        lbl_left_title.pack(fill="x", pady=(0, 10))
        
        # Canvas e scrollbar para coluna esquerda
        canvas_left = tk.Canvas(col_left, bg=bg_color, highlightthickness=0)
        scroll_left = ttk.Scrollbar(col_left, orient="vertical", command=canvas_left.yview)
        frame_left = tk.Frame(canvas_left, bg=bg_color)
        
        frame_left.bind(
            "<Configure>",
            lambda e: canvas_left.configure(scrollregion=canvas_left.bbox("all")))
        
        canvas_left.create_window((0, 0), window=frame_left, anchor="nw")
        canvas_left.configure(yscrollcommand=scroll_left.set)
        
        scroll_left.pack(side="right", fill="y")
        canvas_left.pack(side="left", fill="both", expand=True)
        
        self.frame_colunas_codigos[0] = frame_left
        
        # Coluna direita
        col_right = tk.Frame(main_container, bg=bg_color)
        col_right.pack(side="right", fill="both", expand=True, padx=5)
        
        # Título da coluna direita
        lbl_right_title = tk.Label(
            col_right, 
            text="CÓDIGOS LIDOS", 
            font=("Arial", 12, "bold"), 
            bg="#002b5c", 
            fg="white",
            padx=10,
            pady=5
        )
        lbl_right_title.pack(fill="x", pady=(0, 10))
        
        # Canvas e scrollbar para coluna direita
        canvas_right = tk.Canvas(col_right, bg=bg_color, highlightthickness=0)
        scroll_right = ttk.Scrollbar(col_right, orient="vertical", command=canvas_right.yview)
        frame_right = tk.Frame(canvas_right, bg=bg_color)
        
        frame_right.bind(
            "<Configure>",
            lambda e: canvas_right.configure(scrollregion=canvas_right.bbox("all")))
        
        canvas_right.create_window((0, 0), window=frame_right, anchor="nw")
        canvas_right.configure(yscrollcommand=scroll_right.set)
        
        scroll_right.pack(side="right", fill="y")
        canvas_right.pack(side="left", fill="both", expand=True)
        
        self.frame_colunas_codigos[1] = frame_right
        
        # Configura o bind do mouse wheel para ambas as colunas
        canvas_left.bind_all("<MouseWheel>", lambda e: canvas_left.yview_scroll(int(-1*(e.delta/120)), "units"))
        canvas_right.bind_all("<MouseWheel>", lambda e: canvas_right.yview_scroll(int(-1*(e.delta/120)), "units"))

    def atualizar_visualizacao_codigos(self):
        """Atualiza a visualização dos códigos nas colunas com estilo melhorado"""
        if not hasattr(self, 'frame_colunas_codigos') or not self.codigos_a_ler:
            return
        
        # Limpa os labels antigos
        for labels in self.labels_codigos:
            for label in labels:
                label.destroy()
        self.labels_codigos = [[], []]
        
        # Divide os códigos em duas colunas
        metade = len(self.codigos_a_ler) // 2
        if len(self.codigos_a_ler) % 2 != 0:
            metade += 1
        
        colunas_codigos = [
            self.codigos_a_ler[:metade],
            self.codigos_a_ler[metade:]
        ]
        
        # Configurações de estilo
        bg_color = "#f8f9fa"
        font_size = 12
        item_height = 2  # Altura em linhas
        item_width = 20  # Largura em caracteres
        radius = 25  # Raio para bordas arredondadas
        
        # Adiciona o palet como primeiro item se existir
        if hasattr(self, 'etiqueta_palet_atual') and self.etiqueta_palet_atual:
            lbl_palet = tk.Label(
                self.frame_colunas_codigos[0],
                text=f"PALET: {self.etiqueta_palet_atual}",
                font=("Arial", 12, "bold"),
                bg="#d4edda",  # Verde claro
                fg="#155724",   # Verde escuro
                relief="groove",
                padx=10,
                pady=5,
                width=item_width,
                height=item_height,
                borderwidth=2,
                highlightthickness=0,
                highlightbackground="#c3e6cb",
                highlightcolor="#c3e6cb"
            )
            lbl_palet.pack(fill="x", pady=5, padx=5, ipady=5)
            self.labels_codigos[0].append(lbl_palet)
        
        # Preenche as colunas com os códigos
        for col_idx, codigos in enumerate(colunas_codigos):
            for codigo in codigos:
                # Verifica se o código já foi lido
                lido = codigo in self.codigos_lidos
                
                # Configurações de estilo baseadas no status
                if lido:
                    bg = "#d4edda"  # Verde claro para códigos lidos
                    fg = "#155724"   # Verde escuro
                else:
                    bg = "#f8d7da"  # Vermelho claro para códigos faltantes
                    fg = "#721c24"   # Vermelho escuro
                
                lbl = tk.Label(
                    self.frame_colunas_codigos[col_idx],
                    text=codigo,
                    font=("Arial", font_size),
                    bg=bg,
                    fg=fg,
                    relief="groove",
                    padx=10,
                    pady=5,
                    width=item_width,
                    height=item_height,
                    borderwidth=2,
                    highlightthickness=0,
                    highlightbackground="#f5c6cb" if not lido else "#c3e6cb",
                    highlightcolor="#f5c6cb" if not lido else "#c3e6cb"
                )
                lbl.pack(fill="x", pady=5, padx=5, ipady=5)
                self.labels_codigos[col_idx].append(lbl)

    def determinar_tipo_etiqueta(self, codigo):
        """Determina se a etiqueta é CAMPINAS (C) ou MAFRA (U)"""
        if codigo.startswith('B'):
            return "Campinas"
        elif codigo.startswith('U'):
            return "MAFRA"
        return None

    def validar_formato_etiqueta(self, codigo, tipo_etq, tipo):
        if tipo_etq == "Campinas":
            if tipo == "PALETE":
                return (len(codigo) == 10 and codigo[1:6].isdigit() and codigo[6] == 'P' and codigo[7:].isdigit())
            else:
                return len(codigo) == 10 and codigo[1:].isdigit()
        elif tipo_etq == "MAFRA":
            if tipo == "PALETE":
                return (len(codigo) == 10 and codigo[1:6].isdigit() and codigo[6].isalpha() and codigo[7:].isdigit())
            else:
                return len(codigo) == 10 and codigo[1:].isdigit()
        return False

    def _codigo_ja_registrado_no_palete(self, codigo):
        """Verifica se o código já foi registrado no palete atual"""
        if not hasattr(self, 'etiqueta_palet_atual') or not self.etiqueta_palet_atual:
            return False
            
        for registro in self.registros_temporarios:
            if registro['Codigo'] == codigo:
                return True
        return False

    def _verificar_etiqueta_em_todas_plantas(self, codigo):
        """Verifica se a etiqueta existe em qualquer planta"""
        # Primeiro verifica nos registros temporários (ainda não salvos)
        for registro in self.registros_temporarios:
            if registro['Codigo'] == codigo:
                return True
        
        # Depois verifica nos arquivos Excel
        for arquivo in [self.arquivo_excel_campinas, self.arquivo_excel_mafra]:
            if os.path.exists(arquivo):
                try:
                    df = pd.read_excel(arquivo, engine='openpyxl')
                    if not df.empty and 'Codigo' in df.columns:
                        if codigo in df['Codigo'].values:
                            return True
                except Exception as e:
                    logging.error(f"Erro ao verificar etiqueta em {arquivo}: {str(e)}")
                    continue
        return False

    def processar_etiqueta(self, tipo):
        """Processa as etiquetas com validação e verificação de existência"""
        entry = self.entry_palete if tipo == "PALETE" else self.entry_progressiva
        codigo = entry.get().strip().upper()
        
        if len(codigo) < 10:
            return
            
        # Determina o tipo de etiqueta
        self.tipo_etiqueta = self.determinar_tipo_etiqueta(codigo)
        if self.tipo_etiqueta is None:
            winsound.Beep(400, 200)
            entry.delete(0, tk.END)
            return
            
        # Valida o formato
        if not self.validar_formato_etiqueta(codigo, self.planta_selecionada, tipo):
            winsound.Beep(400, 200)
            entry.delete(0, tk.END)
            return

        # Verifica se já foi registrado no palete atual
        if self._codigo_ja_registrado_no_palete(codigo):
            messagebox.showwarning("Aviso", f"Código {codigo} já foi registrado neste palete!")
            entry.delete(0, tk.END)
            winsound.Beep(400, 200)
            return

        # Verifica se está em PRODUÇÃO/RETRABALHO (para todos os tipos)
        origem = None
        if self.df_producao is not None:
            if tipo == "PALETE" and codigo in self.df_producao["PALET"].values:
                progressivas = self.df_producao[self.df_producao["PALET"] == codigo]["PROGRESSIVA"].tolist()
                if progressivas:
                    self._mostrar_progressivas_na_treeview(codigo, progressivas, "PRODUÇÃO")
                    self.entry_palete.config(state="disabled")
                    self.entry_progressiva.config(state="normal")
                    self.btn_fechar.config(state="normal")
                    self.entry_progressiva.focus()
                    entry.delete(0, tk.END)
                    winsound.Beep(1000, 100)
                    
                    # Atualiza a lista de códigos a serem lidos
                    self.codigos_a_ler = progressivas
                    self.codigos_lidos = set()
                    self.atualizar_visualizacao_codigos()
                    return
            elif codigo in self.df_producao["PROGRESSIVA"].values:
                origem = "PRODUÇÃO"
        
        if self.df_retrabalho is not None:
            if tipo == "PALETE" and codigo in self.df_retrabalho["PALET"].values:
                progressivas = self.df_retrabalho[self.df_retrabalho["PALET"] == codigo]["PROGRESSIVA"].tolist()
                if progressivas:
                    self._mostrar_progressivas_na_treeview(codigo, progressivas, "RETRABALHO")
                    self.entry_palete.config(state="disabled")
                    self.entry_progressiva.config(state="normal")
                    self.btn_fechar.config(state="normal")
                    self.entry_progressiva.focus()
                    entry.delete(0, tk.END)
                    winsound.Beep(1000, 100)
                    
                    # Atualiza a lista de códigos a serem lidos
                    self.codigos_a_ler = progressivas
                    self.codigos_lidos = set()
                    self.atualizar_visualizacao_codigos()
                    return
            elif codigo in self.df_retrabalho["PROGRESSIVA"].values:
                origem = "RETRABALHO"

        # Se não encontrou em PRODUÇÃO/RETRABALHO
        if not origem:
            if not messagebox.askyesno("Código não vinculado", f"Código {codigo} não está em PRODUÇÃO nem RETRABALHO.\nDeseja adicionar?", icon="question"):
                entry.delete(0, tk.END)
                return
            excel_status = "NÃO APONTADA"
            tipo_exibicao = tipo  # Mostra PALETE/PROGRESSIVA na coluna Tipo
        else:
            excel_status = "APONTADA"
            tipo_exibicao = origem  # Mostra PRODUÇÃO/RETRABALHO na coluna Tipo

        # Verifica se já existe no banco de dados (em qualquer planta)
        repetida = self._verificar_etiqueta_em_todas_plantas(codigo)
        
        # Determina o valor da etiqueta palet
        if tipo == "PALETE":
            etiqueta_palet = "-"  # Paletes não têm palet associado
            self.etiqueta_palet_atual = codigo  # Define o palete atual
        else:
            etiqueta_palet = self.etiqueta_palet_atual if hasattr(self, 'etiqueta_palet_atual') and self.etiqueta_palet_atual else "-"
        
        # Processa o registro
        if tipo == "PALETE":
            self.registrar_etiqueta(etiqueta_palet, codigo, tipo_exibicao, excel_status, repetida)
            self.entry_palete.config(state="disabled")
            self.entry_progressiva.config(state="normal")
            self.btn_fechar.config(state="normal")
            self.entry_progressiva.focus()
        else:
            self.registrar_etiqueta(etiqueta_palet, codigo, tipo_exibicao, excel_status, repetida)
            # Adiciona o código aos lidos
            self.codigos_lidos.add(codigo)
            self.atualizar_visualizacao_codigos()
        
        entry.delete(0, tk.END)
        winsound.Beep(1000 if tipo == "PALETE" else 800, 100)
        
        if self.leitura_continua and tipo == "PROGRESSIVA":
            self.root_menu.after(50, lambda: self.entry_progressiva.focus())

    def _mostrar_progressivas_na_treeview(self, palet, progressivas, origem):
        """Mostra progressivas relacionadas na treeview"""
        for item in self.tree.get_children():
            self.tree.delete(item)
        
        for prog in progressivas:
            # Verifica se a progressiva já existe em qualquer planta
            repetida = self._verificar_etiqueta_em_todas_plantas(prog)
            
            self.tree.insert("", "end", values=(
                palet,          # Etiqueta Palet
                prog,           # Código (progressiva)
                origem,         # Tipo (PRODUÇÃO/RETRABALHO)
                "SIM" if repetida else "NÃO",  # Repetida
                datetime.now().strftime("%d/%m/%Y"),
                datetime.now().strftime("%H:%M:%S"),
                "APONTADA",     # Excel
                self.usuario_logado
            ), tags=('repetida' if repetida else ''))
        
        self.lbl_contador.config(text=str(len(progressivas)))
        self.registros_temporarios = [{
            "Etiqueta Palet": palet,
            "Codigo": prog,
            "Tipo": origem,
            "Repetida": "SIM" if self._verificar_etiqueta_em_todas_plantas(prog) else "NÃO",
            "Data": datetime.now().strftime("%d/%m/%Y"),
            "Hora": datetime.now().strftime("%H:%M:%S"),
            "Excel": "APONTADA",
            "Usuario": self.usuario_logado,
            "Local": self.planta_selecionada
        } for prog in progressivas]

    def registrar_etiqueta(self, etiqueta_palet, codigo, tipo, excel_status, repetida):
        """Registra uma etiqueta na treeview com cores diferenciadas"""
        data = datetime.now().strftime("%d/%m/%Y")
        hora = datetime.now().strftime("%H:%M:%S")
        
        # Garante que os valores sejam consistentes
        status_repetida = "SIM" if repetida else "NÃO"
        etiqueta_palet = etiqueta_palet if etiqueta_palet else "-"
        
        # Determina as tags para coloração
        tags = []
        if repetida:
            tags.append('repetida')
        if excel_status == "NÃO APONTADA":
            tags.append('nao_apontada')
        
        # Se for ambos, usa uma tag especial
        if repetida and excel_status == "NÃO APONTADA":
            tags = ['repetida_nao_apontada']
        
        self.tree.insert("", "end", values=(
            etiqueta_palet,
            codigo,
            tipo,
            status_repetida,
            data,
            hora,
            excel_status,
            self.usuario_logado
        ), tags=tuple(tags))
        
        self.registros_temporarios.append({
            "Etiqueta Palet": etiqueta_palet,
            "Codigo": codigo,
            "Tipo": tipo,
            "Repetida": status_repetida,
            "Data": data,
            "Hora": hora,
            "Excel": excel_status,
            "Usuario": self.usuario_logado,
            "Local": self.planta_selecionada
        })
        
        if not repetida:
            self.lbl_contador.config(text=str(int(self.lbl_contador.cget("text")) + 1))

    def fechar_palete(self):
        """Fecha o palete atual e salva no banco de dados da planta atual"""
        # Verifica se há registros temporários primeiro
        if not self.registros_temporarios:
            messagebox.showwarning("Aviso", "Nenhuma etiqueta registrada para fechar o palete!")
            return
            
        # Obtém o palet atual do primeiro registro (se existir)
        palet_atual = self.registros_temporarios[0].get("Etiqueta Palet", "-")
        if palet_atual == "-":
            # Se não tem palet associado, verifica se foi lido um palet
            if hasattr(self, 'etiqueta_palet_atual') and self.etiqueta_palet_atual:
                palet_atual = self.etiqueta_palet_atual
            else:
                messagebox.showwarning("Aviso", "Nenhum palete selecionado para fechar!")
                return
            
        try:
            # Verifica se temos um arquivo Excel definido para a planta atual
            if not self.arquivo_excel_atual:
                messagebox.showerror("Erro", "Planta não selecionada ou arquivo de destino não definido!")
                return
                
            logging.info(f"Tentando salvar palete {palet_atual} no arquivo: {self.arquivo_excel_atual}")
                
            # Cria o diretório se não existir
            pasta_destino = os.path.dirname(self.arquivo_excel_atual)
            if not os.path.exists(pasta_destino):
                os.makedirs(pasta_destino)
                logging.info(f"Diretório criado: {pasta_destino}")
                
            # Carrega ou cria o DataFrame existente
            if os.path.exists(self.arquivo_excel_atual):
                try:
                    df_existente = pd.read_excel(self.arquivo_excel_atual, engine="openpyxl")
                except Exception as e:
                    logging.error(f"Erro ao ler arquivo existente: {str(e)}")
                    colunas = ["Etiqueta Palet", "Codigo", "Tipo", "Repetida", 
                              "Data", "Hora", "Excel", "Usuario", "Local"]
                    df_existente = pd.DataFrame(columns=colunas)
            else:
                colunas = ["Etiqueta Palet", "Codigo", "Tipo", "Repetida", 
                          "Data", "Hora", "Excel", "Usuario", "Local"]
                df_existente = pd.DataFrame(columns=colunas)
            
            # Verifica duplicatas antes de salvar
            codigos_existentes = set(df_existente['Codigo'].dropna().astype(str).unique())
            registros_novos = []
            
            for registro in self.registros_temporarios:
                codigo_str = str(registro['Codigo'])
                registro['Local'] = self.planta_selecionada
                registros_novos.append(registro)
            
            # Cria DataFrame com os novos registros
            df_novos = pd.DataFrame(registros_novos)
            
            # Combina com dados existentes
            df_final = pd.concat([df_existente, df_novos], ignore_index=True)
            
            # Tenta salvar no arquivo
            try:
                with pd.ExcelWriter(self.arquivo_excel_atual, engine='openpyxl', mode='w') as writer:
                    df_final.to_excel(writer, index=False)
                    
                # Atualiza o cache de dados
                self._data_cache = None
                self._last_update = None
                
                # Mensagem de confirmação sem mostrar duplicatas
                total_etiquetas = len(self.registros_temporarios)
                mensagem = f"Palete {palet_atual} ({self.planta_selecionada}) fechado com sucesso!\n"
                mensagem += f"Total de etiquetas: {total_etiquetas}"
                
                messagebox.showinfo("Palete Fechado", mensagem)
                logging.info(f"Palete fechado: {palet_atual} com {total_etiquetas} etiquetas")
                
                # Limpa os dados temporários
                self.limpar_dados_temporarios()
                
            except PermissionError:
                messagebox.showerror("Erro", "Arquivo Excel está aberto ou sem permissão de escrita!")
                logging.error("Erro de permissão ao tentar salvar o arquivo Excel")
            except Exception as e:
                messagebox.showerror("Erro", f"Falha ao salvar no arquivo: {str(e)}")
                logging.error(f"Erro ao salvar arquivo: {str(e)}")
            
        except Exception as e:
            logging.error(f"Erro ao fechar palete: {str(e)}", exc_info=True)
            messagebox.showerror("Erro", f"Falha ao fechar palete: {str(e)}")

    def limpar_dados_temporarios(self):
        self.registros_temporarios = []
        self.codigo_registrado = {}
        self.etiqueta_palet_atual = None
        self.tipo_etiqueta = None
        self.codigos_a_ler = []
        self.codigos_lidos = set()
        
        if hasattr(self, 'lbl_contador'):
            self.lbl_contador.config(text="0")
            
        if hasattr(self, 'tree'):
            for item in self.tree.get_children():
                self.tree.delete(item)
                
        if hasattr(self, 'entry_palete'):
            self.entry_palete.config(state="normal")
            self.entry_palete.delete(0, tk.END)
            self.entry_palete.focus()
            
        if hasattr(self, 'entry_progressiva'):
            self.entry_progressiva.config(state="disabled")
            self.entry_progressiva.delete(0, tk.END)
            
        if hasattr(self, 'btn_fechar'):
            self.btn_fechar.config(state="disabled")
        
        # Atualiza a visualização dos códigos
        self.atualizar_visualizacao_codigos()

    def limpar_campos(self):
        self.limpar_dados_temporarios()

    def toggle_leitura_continua(self):
        self.leitura_continua = not self.leitura_continua
        status = "ON" if self.leitura_continua else "OFF"
        messagebox.showinfo("Leitura Contínua", f"Modo leitura contínua: {status}")

    # ==============================================
    # MÉTODOS DA ABA DE CONSULTA
    # ==============================================
    
    def criar_aba_consulta(self, notebook):
        aba = tk.Frame(notebook)
        notebook.add(aba, text="📁 Consulta")
        
        main_frame = tk.Frame(aba)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        filtro_frame = tk.Frame(main_frame)
        filtro_frame.pack(fill="x", pady=5)
        
        tk.Label(filtro_frame, text="Código:").grid(row=0, column=0, padx=5)
        self.entrada_codigo = tk.Entry(filtro_frame, width=15)
        self.entrada_codigo.grid(row=0, column=1, padx=5)
        
        tk.Label(filtro_frame, text="Tipo:").grid(row=0, column=2, padx=5)
        self.combo_tipo = ttk.Combobox(filtro_frame, values=["Todos", "PALET", "PROGRESSIVA"], width=10)
        self.combo_tipo.grid(row=0, column=3, padx=5)
        self.combo_tipo.set("Todos")
        
        tk.Label(filtro_frame, text="Local:").grid(row=0, column=4, padx=5)
        self.combo_local = ttk.Combobox(filtro_frame, values=["Todos", "CAMPINAS", "MAFRA"], width=10)
        self.combo_local.grid(row=0, column=5, padx=5)
        self.combo_local.set("Todos")
        
        tk.Label(filtro_frame, text="Data Início (dd/mm/aaaa):").grid(row=0, column=6, padx=5)
        self.entrada_data_inicio = tk.Entry(filtro_frame, width=12)
        self.entrada_data_inicio.grid(row=0, column=7, padx=5)
        self.entrada_data_inicio.insert(0, "dd/mm/aaaa")
        self.entrada_data_inicio.bind("<FocusIn>", lambda e: self.limpar_placeholder(self.entrada_data_inicio, "dd/mm/aaaa"))
        
        tk.Label(filtro_frame, text="Data Fim (dd/mm/aaaa):").grid(row=0, column=8, padx=5)
        self.entrada_data_fim = tk.Entry(filtro_frame, width=12)
        self.entrada_data_fim.grid(row=0, column=9, padx=5)
        self.entrada_data_fim.insert(0, "dd/mm/aaaa")
        self.entrada_data_fim.bind("<FocusIn>", lambda e: self.limpar_placeholder(self.entrada_data_fim, "dd/mm/aaaa"))
        
        tk.Button(filtro_frame, text="Pesquisar", command=self.filtrar_dados, bg="#002b5c", fg="white").grid(row=0, column=10, padx=5)
        tk.Button(filtro_frame, text="Limpar", command=self.limpar_filtros, bg="#6c757d", fg="white").grid(row=0, column=12, padx=5)
        
        tree_frame = tk.Frame(main_frame)
        tree_frame.pack(fill="both", expand=True)
        
        cols = ["Etiqueta Palet", "Codigo", "Tipo", "Repetida", "Data", "Hora", "Excel", "Usuario", "Local"]
        self.tree_consulta = ttk.Treeview(
            tree_frame,
            columns=cols,
            show="headings",
            height=25
        )
        
        col_config = {
            "Etiqueta Palet": 120, "Codigo": 120, "Tipo": 80, 
            "Repetida": 80, "Data": 100, "Hora": 80, 
            "Excel": 100, "Usuario": 100, "Local": 80
        }
        
        for col in cols:
            self.tree_consulta.heading(col, text=col)
            self.tree_consulta.column(col, width=col_config.get(col, 120), anchor="center")
        
        # Configuração das cores de fundo mais suaves para a consulta
        self.tree_consulta.tag_configure('repetida', background='#ffffcc')  # Amarelo claro
        self.tree_consulta.tag_configure('nao_apontada', background='#ffe6e6')  # Vermelho claro
        self.tree_consulta.tag_configure('repetida_nao_apontada', background='#ffdd99')  # Laranja claro
        
        scroll_y = ttk.Scrollbar(tree_frame, orient="vertical", command=self.tree_consulta.yview)
        scroll_y.pack(side="right", fill="y")
        scroll_x = ttk.Scrollbar(tree_frame, orient="horizontal", command=self.tree_consulta.xview)
        scroll_x.pack(side="bottom", fill="x")
        self.tree_consulta.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        self.tree_consulta.pack(fill="both", expand=True)
        
        self.context_menu = tk.Menu(self.tree_consulta, tearoff=0)
        self.context_menu.add_command(label="Deletar Registro", command=self.deletar_registro)
        
        self.tree_consulta.bind("<Button-3>", self.mostrar_menu_contexto)
        if os.name == 'nt':
            self.tree_consulta.bind("<Button-2>", self.mostrar_menu_contexto)
        else:
            self.tree_consulta.bind("<Button-2>", self.mostrar_menu_contexto)
        
        self.lbl_info_consulta = tk.Label(
            tree_frame, 
            text="Preencha os filtros e clique em 'Pesquisar' para exibir os dados",
            font=("Arial", 10)
        )
        self.lbl_info_consulta.pack(pady=50)

    def mostrar_menu_contexto(self, event):
        item = self.tree_consulta.identify_row(event.y)
        if item:
            self.tree_consulta.selection_set(item)
            self.context_menu.post(event.x_root, event.y_root)

    def deletar_registro(self):
        selected_item = self.tree_consulta.selection()
        if not selected_item:
            return
            
        item_data = self.tree_consulta.item(selected_item, 'values')
        codigo = item_data[1]
        local = item_data[8]
        
        if not messagebox.askyesno("Confirmar", f"Tem certeza que deseja deletar o registro {codigo}?"):
            return
            
        try:
            arquivo_excel = (
                self.arquivo_excel_campinas if local == "CAMPINAS" 
                else self.arquivo_excel_mafra
            )
            
            df = pd.read_excel(arquivo_excel, engine='openpyxl')
            df = df[df['Codigo'] != codigo]
            
            with pd.ExcelWriter(arquivo_excel, engine='openpyxl', mode='w') as writer:
                df.to_excel(writer, index=False)
            
            self.todos_codigos_registrados.discard(codigo)
            
            self.tree_consulta.delete(selected_item)
            messagebox.showinfo("Sucesso", "Registro deletado com sucesso!")
            
        except Exception as e:
            logging.error(f"Erro ao deletar registro: {str(e)}")
            messagebox.showerror("Erro", f"Falha ao deletar: {str(e)}")

    def carregar_dados_consulta(self):
        if self._data_cache is not None and self._last_update is not None:
            if datetime.now() - self._last_update < self._cache_expiry:
                return self._data_cache.copy()
        
        dados = []
        self.todos_codigos_registrados = set()
        
        if os.path.exists(self.arquivo_excel_campinas):
            df_campinas = pd.read_excel(self.arquivo_excel_campinas, engine='openpyxl')
            df_campinas['Local'] = "CAMPINAS"
            dados.append(df_campinas)
            self.todos_codigos_registrados.update(df_campinas['Codigo'].dropna().unique())
        
        if os.path.exists(self.arquivo_excel_mafra):
            df_mafra = pd.read_excel(self.arquivo_excel_mafra, engine='openpyxl')
            df_mafra['Local'] = "MAFRA"
            dados.append(df_mafra)
            self.todos_codigos_registrados.update(df_mafra['Codigo'].dropna().unique())
        
        if not dados:
            return pd.DataFrame()
        
        df_final = pd.concat(dados, ignore_index=True)
        
        self._data_cache = df_final.copy()
        self._last_update = datetime.now()
        
        return df_final

    def filtrar_dados(self):
        try:
            self._mostrar_loading_consulta()
            
            threading.Thread(
                target=self._filtrar_dados_background,
                daemon=True
            ).start()
            
        except Exception as e:
            logging.error(f"Erro ao filtrar dados: {str(e)}")
            messagebox.showerror("Erro", f"Falha ao filtrar: {str(e)}")

    def _filtrar_dados_background(self):
        try:
            df = self.carregar_dados_consulta()
            
            codigo = self.entrada_codigo.get().strip().upper()
            tipo = self.combo_tipo.get()
            local = self.combo_local.get()
            data_ini_str = self.entrada_data_inicio.get()
            data_fim_str = self.entrada_data_fim.get()
            
            data_ini = self.validar_data(data_ini_str)
            data_fim = self.validar_data(data_fim_str)
            
            if data_ini is False or data_fim is False:
                return
                
            if not any([codigo, tipo != "Todos", local != "Todos", data_ini_str not in ["", "dd/mm/aaaa"], data_fim_str not in ["", "dd/mm/aaaa"]]):
                self.root_menu.after(0, lambda: messagebox.showwarning("Aviso", "Preencha pelo menos um filtro para pesquisar!"))
                return
            
            if df.empty:
                self.root_menu.after(0, lambda: self._atualizar_ui_consulta(pd.DataFrame(), "Nenhum dado encontrado nos arquivos"))
                return
            
            df['Etiqueta Palet'] = df['Etiqueta Palet'].fillna('-')
            df['DataHora'] = pd.to_datetime(df['Data'] + ' ' + df['Hora'], dayfirst=True, errors='coerce')
            
            if codigo:
                df = df[df['Codigo'].str.contains(codigo, case=False, na=False)]
            
            if tipo != "Todos":
                df = df[df['Tipo'] == tipo]
            
            if local != "Todos":
                df = df[df['Local'] == local]
            
            if data_ini and data_fim:
                mask = (df['DataHora'] >= data_ini) & (df['DataHora'] <= data_fim)
                df = df.loc[mask]
            elif data_ini:
                mask = (df['DataHora'] >= data_ini)
                df = df.loc[mask]
            elif data_fim:
                mask = (df['DataHora'] <= data_fim)
                df = df.loc[mask]
            
            df = df.sort_values(['DataHora', 'Tipo'], ascending=[False, True])
            
            self.root_menu.after(0, lambda: self._atualizar_ui_consulta(df))
            
        except Exception as e:
            logging.error(f"Erro ao filtrar dados em background: {str(e)}")
            self.root_menu.after(0, lambda: messagebox.showerror("Erro", f"Falha ao filtrar: {str(e)}"))

    def _mostrar_loading_consulta(self):
        for item in self.tree_consulta.get_children():
            self.tree_consulta.delete(item)
            
        if hasattr(self, 'lbl_info_consulta'):
            self.lbl_info_consulta.config(text="Carregando dados...")
            self.lbl_info_consulta.pack(pady=50)

    def _atualizar_ui_consulta(self, df, mensagem=None):
        self.tree_consulta.delete(*self.tree_consulta.get_children())
        
        if hasattr(self, 'lbl_info_consulta'):
            self.lbl_info_consulta.pack_forget()
        
        if mensagem:
            self.lbl_info_consulta.config(text=mensagem)
            self.lbl_info_consulta.pack(pady=50)
            return
            
        if df.empty:
            self.lbl_info_consulta.config(text="Nenhum resultado encontrado com os filtros aplicados")
            self.lbl_info_consulta.pack(pady=50)
            return
        
        for _, row in df.iterrows():
            tags = []
            if row['Repetida'] == "SIM" and row['Excel'] == "NÃO APONTADA":
                tags.append('repetida_nao_apontada')
            elif row['Repetida'] == "SIM":
                tags.append('repetida')
            elif row['Excel'] == "NÃO APONTADA":
                tags.append('nao_apontada')
            
            self.tree_consulta.insert("", "end", values=(
                row['Etiqueta Palet'],
                row['Codigo'],
                row['Tipo'],
                row['Repetida'],
                row['Data'],
                row['Hora'],
                row['Excel'],
                row['Usuario'],
                row['Local']
            ), tags=tuple(tags))

    def limpar_placeholder(self, entry, placeholder):
        if entry.get() == placeholder:
            entry.delete(0, tk.END)

    def validar_data(self, data_str):
        try:
            if data_str == "dd/mm/aaaa" or not data_str:
                return None
            return datetime.strptime(data_str, "%d/%m/%Y")
        except ValueError:
            self.root_menu.after(0, lambda: messagebox.showerror("Erro", "Formato de data inválido! Use dd/mm/aaaa"))
            return False

    def limpar_filtros(self):
        self.entrada_codigo.delete(0, tk.END)
        self.combo_tipo.set("Todos")
        self.combo_local.set("Todos")
        self.entrada_data_inicio.delete(0, tk.END)
        self.entrada_data_inicio.insert(0, "dd/mm/aaaa")
        self.entrada_data_fim.delete(0, tk.END)
        self.entrada_data_fim.insert(0, "dd/mm/aaaa")
        
        for item in self.tree_consulta.get_children():
            self.tree_consulta.delete(item)
            
        if hasattr(self, 'lbl_info_consulta'):
            self.lbl_info_consulta.config(text="Preencha os filtros e clique em 'Pesquisar' para exibir os dados")
            self.lbl_info_consulta.pack(pady=50)

    # ==============================================
    # MÉTODOS DA ABA DE ANÁLISE
    # ==============================================
    
    def criar_aba_analise(self, notebook):
        """Aba de análise de dados"""
        aba = tk.Frame(notebook)
        notebook.add(aba, text="📊 Análise")
        
        main_frame = tk.Frame(aba)
        main_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        controle_frame = tk.Frame(main_frame)
        controle_frame.pack(fill="x", pady=5)
        
        tk.Button(
            controle_frame,
            text="Atualizar Dados",
            command=self.atualizar_analise,
            bg="#002b5c",
            fg="white"
        ).pack(side="left", padx=5)
        
        # Adicionando filtro por período
        tk.Label(controle_frame, text="Período:").pack(side="left", padx=5)
        self.combo_periodo = ttk.Combobox(controle_frame, 
                                         values=["Últimos 7 dias", "Últimos 30 dias", "Últimos 90 dias", "Todo o período"],
                                         width=15)
        self.combo_periodo.pack(side="left", padx=5)
        self.combo_periodo.set("Todo o período")
        self.combo_periodo.bind("<<ComboboxSelected>>", lambda e: self.atualizar_analise())
        
        metricas_frame = tk.Frame(main_frame, bd=1, relief="solid")
        metricas_frame.pack(fill="x", pady=5)
        
        self.lbl_total = tk.Label(metricas_frame, text="Total: 0", font=("Arial", 10))
        self.lbl_total.pack(side="left", padx=10)
        
        self.lbl_palets = tk.Label(metricas_frame, text="Palets: 0", font=("Arial", 10))
        self.lbl_palets.pack(side="left", padx=10)
        
        self.lbl_progressivas = tk.Label(metricas_frame, text="Progressivas: 0", font=("Arial", 10))
        self.lbl_progressivas.pack(side="left", padx=10)
        
        self.lbl_repetidas = tk.Label(metricas_frame, text="Repetidas: 0", font=("Arial", 10))
        self.lbl_repetidas.pack(side="left", padx=10)
        
        self.lbl_unicas = tk.Label(metricas_frame, text="Únicas: 0", font=("Arial", 10))
        self.lbl_unicas.pack(side="left", padx=10)
        
        self.frame_graficos = tk.Frame(main_frame)
        self.frame_graficos.pack(fill="both", expand=True)
        
        # Adiciona label inicial
        self.lbl_grafico_inicial = tk.Label(
            self.frame_graficos,
            text="Clique em 'Atualizar Dados' para gerar os gráficos",
            font=("Arial", 12)
        )
        self.lbl_grafico_inicial.pack(pady=50)

    def carregar_dados_analise(self):
        """Carrega dados para análise do banco de dados da planta atual"""
        periodo = self.combo_periodo.get()
        
        # Verifica se temos dados em cache que ainda são válidos
        if self._data_cache is not None and self._last_update is not None:
            if datetime.now() - self._last_update < self._cache_expiry:
                df = self._data_cache.copy()
                
                # Aplica filtro de período no cache
                hoje = datetime.now()
                if periodo == "Últimos 7 dias":
                    df = df[df['DataHora'] >= (hoje - timedelta(days=7))]
                elif periodo == "Últimos 30 dias":
                    df = df[df['DataHora'] >= (hoje - timedelta(days=30))]
                elif periodo == "Últimos 90 dias":
                    df = df[df['DataHora'] >= (hoje - timedelta(days=90))]
                
                return df
        
        dados = []
        
        if self.arquivo_excel_atual and os.path.exists(self.arquivo_excel_atual):
            try:
                df = pd.read_excel(self.arquivo_excel_atual, engine='openpyxl')
                
                # Verifica se as colunas necessárias existem
                required_columns = ['Tipo', 'Repetida', 'Data', 'Hora']
                for col in required_columns:
                    if col not in df.columns:
                        df[col] = None  # Cria coluna vazia se não existir
                
                df['DataHora'] = pd.to_datetime(df['Data'] + ' ' + df['Hora'], dayfirst=True, errors='coerce')
                
                # Aplica filtro de período
                hoje = datetime.now()
                if periodo == "Últimos 7 dias":
                    df = df[df['DataHora'] >= (hoje - timedelta(days=7))]
                elif periodo == "Últimos 30 dias":
                    df = df[df['DataHora'] >= (hoje - timedelta(days=30))]
                elif periodo == "Últimos 90 dias":
                    df = df[df['DataHora'] >= (hoje - timedelta(days=90))]
                
                df['Local'] = self.planta_selecionada
                dados.append(df)
                
                # Atualiza cache
                self._data_cache = df.copy()
                self._last_update = datetime.now()
                
            except Exception as e:
                logging.error(f"Erro ao carregar dados para análise: {str(e)}")
                messagebox.showerror("Erro", f"Falha ao ler dados para análise: {str(e)}")
        
        if not dados:
            return pd.DataFrame()
        
        return pd.concat(dados, ignore_index=True)

    def atualizar_analise(self):
        """Atualiza a análise com os dados mais recentes"""
        # Mostra indicador de carregamento
        self._mostrar_loading_analise()
        
        # Inicia thread para processamento em segundo plano
        threading.Thread(
            target=self._atualizar_analise_background,
            daemon=True
        ).start()

    def _mostrar_loading_analise(self):
        """Mostra indicador de carregamento na análise"""
        for widget in self.frame_graficos.winfo_children():
            widget.destroy()
        
        self.loading_label = tk.Label(
            self.frame_graficos,
            text="Carregando dados...",
            font=("Arial", 12)
        )
        self.loading_label.pack(pady=50)

    def _atualizar_analise_background(self):
        """Processa os dados de análise em segundo plano"""
        try:
            df = self.carregar_dados_analise()
            
            # Atualiza a UI na thread principal
            self.root_menu.after(0, lambda: self._atualizar_ui_analise(df))
            
        except Exception as e:
            logging.error(f"Erro ao atualizar análise em background: {str(e)}")
            self.root_menu.after(0, lambda: messagebox.showerror("Erro", f"Falha ao atualizar análise: {str(e)}"))

    def _atualizar_ui_analise(self, df):
        """Atualiza a interface do usuário com os dados de análise"""
        try:
            # Remove o loading indicator
            if hasattr(self, 'loading_label'):
                self.loading_label.pack_forget()
            
            # Limpa o frame antes de criar novos gráficos
            for widget in self.frame_graficos.winfo_children():
                widget.destroy()
            
            if df.empty:
                lbl = tk.Label(
                    self.frame_graficos,
                    text="Nenhum dado disponível para análise",
                    font=("Arial", 12)
                )
                lbl.pack(pady=50)
                return
            
            # Verificação adicional de colunas necessárias
            required_columns = ['Tipo', 'Repetida', 'Data', 'Hora']
            missing_cols = [col for col in required_columns if col not in df.columns]
            
            if missing_cols:
                raise ValueError(f"Colunas faltando no Excel: {', '.join(missing_cols)}")
            
            total = len(df)
            palets = len(df[df['Tipo'] == 'PALET']) if 'Tipo' in df.columns else 0
            progressivas = len(df[df['Tipo'] == 'PROGRESSIVA']) if 'Tipo' in df.columns else 0
            repetidas = len(df[df['Repetida'] == 'SIM']) if 'Repetida' in df.columns else 0
            unicas = len(df[df['Repetida'] == 'NÃO']) if 'Repetida' in df.columns else 0
            
            self.lbl_total.config(text=f"Total: {total}")
            self.lbl_palets.config(text=f"Palets: {palets}")
            self.lbl_progressivas.config(text=f"Progressivas: {progressivas}")
            self.lbl_repetidas.config(text=f"Repetidas: {repetidas}")
            self.lbl_unicas.config(text=f"Únicas: {unicas}")
            
            # Cria os gráficos
            self.criar_graficos_analise(df)
            
        except Exception as e:
            logging.error(f"Erro ao gerar análise: {str(e)}")
            lbl = tk.Label(
                self.frame_graficos,
                text=f"Erro ao gerar análise: {str(e)}",
                font=("Arial", 12),
                fg="red"
            )
            lbl.pack(pady=50)

    def criar_graficos_analise(self, df):
        """Cria os gráficos de análise"""
        try:
            fig = plt.Figure(figsize=(12, 8), dpi=100)
            fig.suptitle(f"Análise de Dados - {self.planta_selecionada}\nPeríodo: {self.combo_periodo.get()}", 
                        fontsize=14)
            
            # Define color scheme
            colors = {
                'primary': '#002b5c',    # Dark blue
                'success': '#28a745',    # Green
                'danger': '#dc3545',     # Red
                'secondary': "#28a745",  # Green
                'warning': '#ffc107'     # Yellow
            }
            
            # Plot 1: Distribution by Type
            ax1 = fig.add_subplot(221)
            self._create_type_distribution_chart(ax1, df, colors)
            
            # Plot 2: Unique vs Repeated
            ax2 = fig.add_subplot(222)
            self._create_unique_vs_repeated_chart(ax2, df, colors)
            
            # Plot 3: Daily Registrations
            ax3 = fig.add_subplot(223)
            self._create_daily_registrations_chart(ax3, df, colors)
            
            # Plot 4: Top Users
            ax4 = fig.add_subplot(224)
            self._create_top_users_chart(ax4, df, colors)
            
            # Adjust layout and display
            fig.tight_layout()
            self._display_chart(fig)
            
        except Exception as e:
            self._handle_visualization_error(e)

    def _create_type_distribution_chart(self, ax, df, colors):
        """Creates the type distribution chart"""
        if 'Tipo' not in df.columns or df['Tipo'].empty:
            ax.text(0.5, 0.5, 'Sem dados de tipo', ha='center', va='center')
            ax.set_title("Distribuição por Tipo (Sem dados)", pad=10)
            return
        
        counts = df['Tipo'].value_counts()
        if counts.empty:
            ax.text(0.5, 0.5, 'Sem dados', ha='center', va='center')
            ax.set_title("Distribuição por Tipo (Sem dados)", pad=10)
            return
        
        # Map types to colors
        color_map = {
            'PALET': colors['primary'],
            'PROGRESSIVA': colors['secondary'],
            'PRODUÇÃO': colors['success'],
            'RETRABALHO': colors['warning']
        }
        bar_colors = [color_map.get(tipo, colors['secondary']) for tipo in counts.index]
        
        bars = counts.plot(kind='bar', ax=ax, color=bar_colors)
        ax.set_title("Distribuição por Tipo", pad=10)
        ax.set_xlabel("")
        ax.set_ylabel("Quantidade")
        
        # Add value labels
        for i, v in enumerate(counts):
            ax.text(i, v + 0.5, str(v), ha='center')

    def _create_unique_vs_repeated_chart(self, ax, df, colors):
        """Creates the unique vs repeated chart"""
        if 'Repetida' not in df.columns or df['Repetida'].empty:
            ax.text(0.5, 0.5, 'Sem dados de repetição', ha='center', va='center')
            ax.set_title("Únicas vs Repetidas (Sem dados)", pad=10)
            return
        
        counts = df['Repetida'].value_counts()
        if counts.empty:
            ax.text(0.5, 0.5, 'Sem dados', ha='center', va='center')
            ax.set_title("Únicas vs Repetidas (Sem dados)", pad=10)
            return
        
        # Map status to colors
        color_map = {
            'NÃO': colors['success'],  # Verde para únicas
            'SIM': colors['danger']    # Vermelho para repetidas
        }
        bar_colors = [color_map.get(rep, colors['secondary']) for rep in counts.index]
        
        bars = counts.plot(kind='bar', ax=ax, color=bar_colors)
        ax.set_title("Etiquetas Únicas vs Repetidas", pad=10)
        ax.set_xlabel("")
        ax.set_ylabel("Quantidade")
        
        # Add value labels
        for i, v in enumerate(counts):
            ax.text(i, v + 0.5, str(v), ha='center')

    def _create_daily_registrations_chart(self, ax, df, colors):
        """Creates the daily registrations line chart"""
        if 'DataHora' not in df.columns or df['DataHora'].empty:
            ax.text(0.5, 0.5, 'Sem dados de data', ha='center', va='center')
            ax.set_title("Registros por Dia (Sem dados)", pad=10)
            return
        
        try:
            df['Data'] = pd.to_datetime(df['Data'], dayfirst=True, errors='coerce')
            registros_por_dia = df.groupby(df['Data'].dt.date).size()
            
            if registros_por_dia.empty:
                ax.text(0.5, 0.5, 'Sem dados suficientes', ha='center', va='center')
                ax.set_title("Registros por Dia (Sem dados)", pad=10)
                return
            
            registros_por_dia.plot(
                ax=ax, 
                marker='o', 
                color=colors['primary'],
                linestyle='-',
                linewidth=2,
                markersize=8
            )
            ax.set_title("Registros por Dia", pad=10)
            ax.set_xlabel("Data")
            ax.set_ylabel("Quantidade")
            
            # Format x-axis to show dates
            ax.xaxis.set_major_formatter(plt.matplotlib.dates.DateFormatter('%d/%m'))
            
            # Add value labels
            for x, y in zip(registros_por_dia.index, registros_por_dia.values):
                ax.text(x, y + 0.5, str(y), ha='center')
                
        except Exception as e:
            logging.error(f"Erro ao criar gráfico diário: {str(e)}")
            ax.text(0.5, 0.5, 'Erro ao processar datas', ha='center', va='center')
            ax.set_title("Registros por Dia (Erro)", pad=10)

    def _create_top_users_chart(self, ax, df, colors):
        """Creates the top users horizontal bar chart"""
        if 'Usuario' not in df.columns or df['Usuario'].empty:
            ax.set_visible(False)  # Hide if no data
            return
        
        top_users = df['Usuario'].value_counts().head(5)
        if top_users.empty:
            ax.text(0.5, 0.5, 'Sem dados de usuários', ha='center', va='center')
            ax.set_title("Registros por Usuário (Sem dados)", pad=10)
            return
        
        bars = top_users.plot(kind='barh', ax=ax, color=colors['primary'])
        ax.set_title("Top 5 Usuários com Mais Registros", pad=10)
        ax.set_xlabel("Quantidade")
        ax.set_ylabel("Usuário")
        
        # Add value labels
        for i, v in enumerate(top_users):
            ax.text(v + 0.5, i, str(v), ha='left', va='center')

    def _display_chart(self, fig):
        """Displays the chart in the GUI"""
        # Clear previous chart if any
        for widget in self.frame_graficos.winfo_children():
            widget.destroy()
        
        # Create canvas and display
        canvas = tkagg.FigureCanvasTkAgg(fig, master=self.frame_graficos)
        canvas.draw()
        
        # Add toolbar for interactivity
        toolbar = tkagg.NavigationToolbar2Tk(canvas, self.frame_graficos)
        toolbar.update()
        
        # Pack components
        canvas.get_tk_widget().pack(fill="both", expand=True)
        toolbar.pack(fill="x")

    def _handle_visualization_error(self, error):
        """Handles errors during visualization"""
        logging.error(f"Erro na visualização: {str(error)}", exc_info=True)
        
        # Clear frame and show error message
        for widget in self.frame_graficos.winfo_children():
            widget.destroy()
        
        error_frame = tk.Frame(self.frame_graficos)
        error_frame.pack(expand=True, fill="both", padx=20, pady=20)
        
        tk.Label(
            error_frame,
            text="Erro ao gerar visualizações",
            font=("Arial", 12, "bold"),
            fg="red"
        ).pack(pady=5)
        
        tk.Label(
            error_frame,
            text=str(error),
            font=("Arial", 10),
            wraplength=400
        ).pack(pady=5)
        
        tk.Button(
            error_frame,
            text="Tentar Novamente",
            command=self.atualizar_analise,
            bg="#002b5c",
            fg="white"
        ).pack(pady=10)

    # ==============================================
    # MÉTODOS AUXILIARES
    # ==============================================
    
    def toggle_leitura_continua(self):
        self.leitura_continua = not self.leitura_continua
        status = "ON" if self.leitura_continua else "OFF"
        messagebox.showinfo("Leitura Contínua", f"Modo leitura contínua: {status}")

    def agendar_auto_save(self):
        """Agenda o próximo auto-save"""
        if self.root_menu:
            self.root_menu.after(self.auto_save_interval * 1000, self.agendar_auto_save)

if __name__ == "__main__":
    app = SistemaInventario()
