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

class Aplicativo:
    def __init__(self, root):
        self.root = root
        self.root.title("Consulta de Unidades de Saúde")
        self.root.geometry("400x200")
        
        self.arquivo_excel = None
        self.df = None
        self.navegador = None
        self.em_execucao = False
        
        # Frame principal
        self.frame = tk.Frame(root, padx=20, pady=20)
        self.frame.pack(expand=True, fill=tk.BOTH)
        
        # Botão para selecionar arquivo
        self.btn_selecionar = tk.Button(
            self.frame, 
            text="Selecionar Arquivo Excel", 
            command=self.selecionar_arquivo,
            height=2,
            width=20
        )
        self.btn_selecionar.pack(pady=10)
        
        # Label para mostrar o arquivo selecionado
        self.lbl_arquivo = tk.Label(self.frame, text="Nenhum arquivo selecionado")
        self.lbl_arquivo.pack(pady=5)
        
        # Botão para iniciar o processo
        self.btn_iniciar = tk.Button(
            self.frame, 
            text="Iniciar Consulta", 
            command=self.iniciar_processo,
            state=tk.DISABLED,
            height=2,
            width=20
        )
        self.btn_iniciar.pack(pady=10)
        
        # Barra de progresso
        self.progresso = ttk.Progressbar(
            self.frame, 
            orient=tk.HORIZONTAL, 
            length=300, 
            mode='determinate'
        )
        self.progresso.pack(pady=10)
        self.progresso.pack_forget()  # Esconde inicialmente
        
        # Label para status
        self.lbl_status = tk.Label(self.frame, text="")
        self.lbl_status.pack(pady=5)
    
    def selecionar_arquivo(self):
        arquivo = filedialog.askopenfilename(
            title="Selecione a planilha de dados",
            filetypes=[("Arquivos Excel", "*.xlsx"), ("Todos os arquivos", "*.*")]
        )
        
        if arquivo:
            self.arquivo_excel = arquivo
            self.lbl_arquivo.config(text=os.path.basename(arquivo))
            self.btn_iniciar.config(state=tk.NORMAL)
    
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
        self.lbl_status.config(text=f"Processando {indice}/{len(self.df)}: {endereco}")
        self.root.update()
        
        return unidade, cnes
    
    def iniciar_processo(self):
        if not self.arquivo_excel:
            messagebox.showerror("Erro", "Nenhum arquivo foi selecionado.")
            return
        
        if self.em_execucao:
            return
            
        self.em_execucao = True
        self.btn_selecionar.config(state=tk.DISABLED)
        self.btn_iniciar.config(state=tk.DISABLED)
        self.progresso.pack()
        
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
                if not self.em_execucao:  # Permite parar o processo se necessário
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
                    if (index + 1) % 5 == 0:  # Salva a cada 5 registros
                        self.df.to_excel(self.arquivo_excel, index=False, engine='openpyxl')
                except Exception as e:
                    print(f"Erro ao buscar dados para o endereço {endereco}: {e}")
            
            # Salvar final
            self.df.to_excel(self.arquivo_excel, index=False, engine='openpyxl')
            
            duracao_programa = int(time() - inicio_programa)
            messagebox.showinfo(
                "Concluído",
                f"Processo finalizado!\nTempo total: {int(duracao_programa/60)}min {duracao_programa%60}seg"
            )
            
        except Exception as e:
            messagebox.showerror("Erro", f"Ocorreu um erro durante o processamento:\n{str(e)}")
        finally:
            if hasattr(self, 'navegador') and self.navegador:
                self.navegador.quit()
            self.em_execucao = False
            self.btn_selecionar.config(state=tk.NORMAL)
            self.btn_iniciar.config(state=tk.NORMAL)
            self.lbl_status.config(text="Processo concluído")
            self.progresso.pack_forget()

if __name__ == "__main__":
    root = tk.Tk()
    app = Aplicativo(root)
    root.mainloop()
