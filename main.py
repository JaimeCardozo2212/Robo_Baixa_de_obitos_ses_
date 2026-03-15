import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.service import Service
from time import sleep, time
import pyautogui as pg
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from tkinter.font import Font

class Aplicativo:
    def __init__(self, root):
        self.root = root
        self.root.title("Consulta de Unidades de Saúde")
        self.root.geometry("500x300")
        
        # Configurar cores modernas
        self.cores = {
            'bg_primary': '#f0f4f8',      # Fundo principal suave
            'bg_secondary': '#ffffff',     # Fundo dos elementos
            'accent': '#4a90e2',            # Azul moderno para botões
            'accent_hover': '#357abd',      # Azul mais escuro para hover
            'text_primary': '#2c3e50',      # Texto principal
            'text_secondary': '#7f8c8d',    # Texto secundário
            'success': '#27ae60',            # Verde para sucesso
            'border': '#e1e8ed'              # Cor de borda
        }
        
        self.root.configure(bg=self.cores['bg_primary'])
        
        self.arquivo_excel = None
        self.df = None
        self.navegador = None
        self.em_execucao = False
        
        # Configurar fontes
        self.fonte_titulo = Font(family="Helvetica", size=12, weight="bold")
        self.fonte_normal = Font(family="Helvetica", size=10)
        self.fonte_status = Font(family="Helvetica", size=9)
        
        # Criar cabeçalho
        self.criar_cabecalho()
        
        # Frame principal com sombra
        self.frame = tk.Frame(
            root, 
            bg=self.cores['bg_secondary'],
            padx=30, 
            pady=30,
            highlightbackground=self.cores['border'],
            highlightthickness=1
        )
        self.frame.pack(expand=True, fill=tk.BOTH, padx=30, pady=30)
        
        # Título
        self.lbl_titulo = tk.Label(
            self.frame, 
            text="Consulta de Unidades de Saúde",
            font=self.fonte_titulo,
            bg=self.cores['bg_secondary'],
            fg=self.cores['text_primary']
        )
        self.lbl_titulo.pack(pady=(0, 20))
        
        # Área de seleção de arquivo com ícone
        self.frame_arquivo = tk.Frame(self.frame, bg=self.cores['bg_secondary'])
        self.frame_arquivo.pack(fill=tk.X, pady=10)
        
        # Botão selecionar com estilo moderno
        self.btn_selecionar = tk.Button(
            self.frame_arquivo, 
            text="📁 Selecionar Arquivo Excel", 
            command=self.selecionar_arquivo,
            font=self.fonte_normal,
            bg=self.cores['accent'],
            fg='white',
            relief=tk.FLAT,
            cursor='hand2',
            height=2,
            width=25
        )
        self.btn_selecionar.pack(side=tk.LEFT, padx=(0, 10))
        
        # Label para mostrar o arquivo selecionado
        self.lbl_arquivo = tk.Label(
            self.frame_arquivo, 
            text="Nenhum arquivo selecionado",
            font=self.fonte_status,
            bg=self.cores['bg_secondary'],
            fg=self.cores['text_secondary'],
            wraplength=200
        )
        self.lbl_arquivo.pack(side=tk.LEFT, fill=tk.X, expand=True)
        
        # Frame para botão iniciar
        self.frame_botao = tk.Frame(self.frame, bg=self.cores['bg_secondary'])
        self.frame_botao.pack(fill=tk.X, pady=20)
        
        # Botão iniciar com estilo
        self.btn_iniciar = tk.Button(
            self.frame_botao, 
            text="▶ Iniciar Consulta", 
            command=self.iniciar_processo,
            state=tk.DISABLED,
            font=self.fonte_normal,
            bg=self.cores['text_secondary'],
            fg='white',
            relief=tk.FLAT,
            cursor='hand2',
            height=2,
            width=25
        )
        self.btn_iniciar.pack()
        
        # Frame de progresso
        self.frame_progresso = tk.Frame(self.frame, bg=self.cores['bg_secondary'])
        self.frame_progresso.pack(fill=tk.X, pady=10)
        self.frame_progresso.pack_forget()
        
        # Barra de progresso estilizada
        style = ttk.Style()
        style.theme_use('clam')
        style.configure(
            "Modern.Horizontal.TProgressbar",
            background=self.cores['accent'],
            troughcolor=self.cores['bg_primary'],
            bordercolor=self.cores['bg_secondary'],
            lightcolor=self.cores['accent'],
            darkcolor=self.cores['accent']
        )
        
        self.progresso = ttk.Progressbar(
            self.frame_progresso, 
            style="Modern.Horizontal.TProgressbar",
            orient=tk.HORIZONTAL, 
            length=400, 
            mode='determinate'
        )
        self.progresso.pack(pady=5)
        
        # Label para status
        self.lbl_status = tk.Label(
            self.frame_progresso, 
            text="",
            font=self.fonte_status,
            bg=self.cores['bg_secondary'],
            fg=self.cores['text_secondary']
        )
        self.lbl_status.pack()
        
        # Adicionar efeito hover nos botões
        self.btn_selecionar.bind("<Enter>", lambda e: self.on_enter(e, self.btn_selecionar))
        self.btn_selecionar.bind("<Leave>", lambda e: self.on_leave(e, self.btn_selecionar))
        self.btn_iniciar.bind("<Enter>", lambda e: self.on_enter(e, self.btn_iniciar))
        self.btn_iniciar.bind("<Leave>", lambda e: self.on_leave(e, self.btn_iniciar))
    
    def criar_cabecalho(self):
        # Frame para o cabeçalho
        cabecalho = tk.Frame(self.root, bg=self.cores['accent'], height=40)
        cabecalho.pack(fill=tk.X)
        cabecalho.pack_propagate(False)
        
        # Título do cabeçalho
        titulo = tk.Label(
            cabecalho,
            text="🔍 Consulta Automatizada",
            font=Font(family="Helvetica", size=10, weight="bold"),
            bg=self.cores['accent'],
            fg='white'
        )
        titulo.pack(side=tk.LEFT, padx=20, pady=10)
    
    def on_enter(self, event, botao):
        if botao['state'] != tk.DISABLED:
            if botao == self.btn_selecionar:
                botao.configure(bg=self.cores['accent_hover'])
            elif botao == self.btn_iniciar:
                botao.configure(bg=self.cores['success'])
    
    def on_leave(self, event, botao):
        if botao['state'] != tk.DISABLED:
            if botao == self.btn_selecionar:
                botao.configure(bg=self.cores['accent'])
            elif botao == self.btn_iniciar:
                botao.configure(bg=self.cores['accent'])
    
    def selecionar_arquivo(self):
        arquivo = filedialog.askopenfilename(
            title="Selecione a planilha de dados",
            filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")]
        )
        
        if arquivo:
            self.arquivo_excel = arquivo
            self.lbl_arquivo.config(
                text=f"📊 {os.path.basename(arquivo)}",
                fg=self.cores['success']
            )
            self.btn_iniciar.config(
                state=tk.NORMAL,
                bg=self.cores['accent']
            )
            self.mostrar_notificacao("Arquivo carregado com sucesso!", "sucesso")
    
    def mostrar_notificacao(self, mensagem, tipo="info"):
        cores = {
            "sucesso": self.cores['success'],
            "erro": "#e74c3c",
            "info": self.cores['accent']
        }
        
        self.lbl_status.config(
            text=mensagem,
            fg=cores.get(tipo, self.cores['text_secondary'])
        )
        self.root.update()
    
    def verificar_chromedriver(self):
        chromedriver_path = os.path.join(os.getcwd(), "chromedriver.exe")
        if not os.path.isfile(chromedriver_path):
            messagebox.showerror(
                "Erro",
                "ChromeDriver não encontrado na pasta do executável.\n"
                "Por favor, baixe o ChromeDriver compatível com a versão do seu Chrome "
                "e coloque o arquivo 'chromedriver.exe' na mesma pasta deste programa."
            )
            return None
        return chromedriver_path
    
    def buscar_distrito_cnes(self, endereco, indice):
        start_time = time()
        
        digitar_nome = self.navegador.find_element(By.XPATH, '//*[@id="address"]')
        digitar_nome.clear()
        digitar_nome.send_keys(endereco + Keys.ENTER)
        sleep(1.4)

        pg.click(x=767, y=578)
        sleep(0.6)
        unidade = self.navegador.find_element(By.XPATH, '//*[@id="dataTable"]/thead/tr/td/table/tbody/tr[1]/td').text
        cnes = self.navegador.find_element(By.XPATH, '//*[@id="dataTable"]/thead/tr/td/table/tbody/tr[2]/td/table/tbody/tr[1]/td[2]').text
        
        duration = int(time() - start_time)
        self.lbl_status.config(
            text=f"📌 Processando {indice}/{len(self.df)}: {endereco[:50]}..."
        )
        self.root.update()
        
        return unidade, cnes
    
    def iniciar_processo(self):
        if not self.arquivo_excel:
            self.mostrar_notificacao("Selecione um arquivo primeiro!", "erro")
            return
        
        if self.em_execucao:
            return
            
        self.em_execucao = True
        self.btn_selecionar.config(state=tk.DISABLED)
        self.btn_iniciar.config(state=tk.DISABLED)
        self.frame_progresso.pack()
        
        try:
            # Carregar os endereços do arquivo Excel
            self.df = pd.read_excel(self.arquivo_excel, engine='openpyxl')
            
            # Verificar se as colunas "unidade" e "cnes" existem
            if 'unidade' not in self.df.columns:
                self.df['unidade'] = None
            if 'cnes' not in self.df.columns:
                self.df['cnes'] = None
            
            # Verificar ChromeDriver
            chromedriver_path = self.verificar_chromedriver()
            if not chromedriver_path:
                return
                
            # Configurar o serviço do ChromeDriver
            service = Service(executable_path=chromedriver_path)
            
            # Iniciar o navegador
            self.navegador = webdriver.Chrome(service=service)
            self.navegador.get('https://inovasaude.joinville.br/mapa/mapa')
            self.navegador.maximize_window()
            sleep(5)
            
            inicio_programa = time()
            self.progresso['maximum'] = len(self.df)
            
            # Processar cada linha
            for index, row in self.df.iterrows():
                if not self.em_execucao:
                    break
                    
                endereco = row['usuarios_sistema.concat_Endereço do usuário']
                try:
                    unidade, cnes = self.buscar_distrito_cnes(endereco, index+1)
                    self.df.at[index, 'unidade'] = unidade
                    self.df.at[index, 'cnes'] = cnes
                    
                    # Atualizar progresso
                    self.progresso['value'] = index + 1
                    self.root.update()
                    
                    # Salvar periodicamente
                    if (index + 1) % 5 == 0:
                        self.df.to_excel(self.arquivo_excel, index=False, engine='openpyxl')
                except Exception as e:
                    print(f"Erro ao buscar dados para o endereço {endereco}: {e}")
            
            # Salvar final
            self.df.to_excel(self.arquivo_excel, index=False, engine='openpyxl')
            
            duracao_programa = int(time() - inicio_programa)
            minutos = duracao_programa // 60
            segundos = duracao_programa % 60
            
            messagebox.showinfo(
                "✅ Concluído",
                f"Processo finalizado com sucesso!\n\n"
                f"⏱️ Tempo total: {minutos}min {segundos}seg\n"
                f"📊 Registros processados: {len(self.df)}"
            )
            
        except Exception as e:
            messagebox.showerror(
                "❌ Erro", 
                f"Ocorreu um erro durante o processamento:\n{str(e)}"
            )
        finally:
            if hasattr(self, 'navegador') and self.navegador:
                self.navegador.quit()
            self.em_execucao = False
            self.btn_selecionar.config(state=tk.NORMAL)
            self.btn_iniciar.config(state=tk.NORMAL)
            self.lbl_status.config(text="✅ Processo concluído")
            self.progresso.pack_forget()

if __name__ == "__main__":
    root = tk.Tk()
    app = Aplicativo(root)
    root.mainloop()
