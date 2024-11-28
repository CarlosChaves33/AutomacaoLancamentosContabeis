import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import pandas as pd
import json
import os
import datetime
import numpy as np
import decimal

class PlanoContasViewer(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Visualização de Planos de Contas")
        self.geometry("1200x600")
        self.setup_ui()
        
    def setup_ui(self):
        # Frame principal
        main_frame = ttk.Frame(self, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Lista de planos
        self.tree = ttk.Treeview(main_frame, columns=("arquivo_json", "empresa"), show="headings")
        self.tree.heading("arquivo_json", text="Arquivo JSON")
        self.tree.heading("empresa", text="Empresa")
        
        # Configurar larguras das colunas
        self.tree.column("arquivo_json", width=300)
        self.tree.column("empresa", width=500)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # Frame para botões
        btn_frame = ttk.Frame(main_frame)
        
        # Botões
        btn_novo = ttk.Button(btn_frame, text="Novo Plano de Contas", command=self.novo_plano)
        btn_visualizar = ttk.Button(btn_frame, text="Visualizar Detalhes", command=self.visualizar_detalhes)
        btn_atualizar = ttk.Button(btn_frame, text="Atualizar Plano", command=self.atualizar_plano)
        btn_excluir = ttk.Button(btn_frame, text="Excluir Plano", command=self.excluir_plano)
        
        # Layout
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        btn_frame.grid(row=1, column=0, pady=10)
        
        btn_novo.grid(row=0, column=0, padx=5)
        btn_visualizar.grid(row=0, column=1, padx=5)
        btn_atualizar.grid(row=0, column=2, padx=5)
        btn_excluir.grid(row=0, column=3, padx=5)
        
        # Configurar grid
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(0, weight=1)
        
        # Carregar planos
        self.carregar_planos()
    
    def carregar_planos(self):
        # Limpar itens existentes
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        if os.path.exists('data'):
            for arquivo in sorted(os.listdir('data')):
                if arquivo.endswith('.json'):
                    with open(os.path.join('data', arquivo), 'r', encoding='utf-8') as f:
                        dados = json.load(f)
                        self.tree.insert('', 'end', values=(arquivo, dados['empresa']))
    
    def novo_plano(self):
        filepath = filedialog.askopenfilename(
            title="Selecione o arquivo do Plano de Contas",
            filetypes=[("Excel files", "*.xls *.xlsx")],
            parent=self
        )
        
        if filepath:
            try:
                # Lê a primeira linha para pegar o nome da empresa
                df_empresa = pd.read_excel(filepath, nrows=1, header=None)
                
                # Encontra a coluna que contém "Empresa:" e o nome
                empresa = None
                empresa_encontrada = False
                
                # Itera sobre as colunas da primeira linha
                for col in df_empresa.columns:
                    valor = str(df_empresa.iloc[0, col]).strip()
                    
                    if empresa_encontrada and valor and valor != 'nan':
                        empresa = valor
                        break
                        
                    if "Empresa:" in valor:
                        empresa_encontrada = True
                
                if not empresa:
                    raise ValueError("Nome da empresa não encontrado no arquivo")
                
                # Verifica se já existe um plano de contas para esta empresa
                for arquivo in os.listdir('data'):
                    if arquivo.endswith('.json'):
                        with open(os.path.join('data', arquivo), 'r', encoding='utf-8') as f:
                            dados_existentes = json.load(f)
                            if dados_existentes['empresa'] == empresa:
                                if not tk.messagebox.askyesno(
                                    "Empresa Existente",
                                    f"Já existe um plano de contas para a empresa:\n{empresa}\n\nDeseja substituir?",
                                    parent=self
                                ):
                                    return
                
                # Processa o arquivo
                df = pd.read_excel(filepath, skiprows=3, header=None)
                df = df.dropna(axis=1, how='all')
                df = df.iloc[1:]
                
                dados = []
                for idx, row in df.iterrows():
                    codigo = row[0]
                    if pd.notna(codigo):
                        tipo = str(row[3]).strip() if pd.notna(row[3]) else ''
                        classificacao = str(row[7]).strip() if pd.notna(row[7]) else ''
                        
                        # Procura o nome em todas as colunas não vazias após a classificação
                        nome = None
                        for col in [col for col in df.columns if col > 7]:
                            valor = row[col]
                            if pd.notna(valor) and isinstance(valor, str) and valor.strip():
                                nome = valor.strip()
                                break
                        
                        if not nome:
                            continue
                            
                        try:
                            grau = None
                            for col in reversed(df.columns):
                                valor = row[col]
                                if pd.notna(valor) and isinstance(valor, (int, float)):
                                    grau = int(valor)
                                    break
                            if grau is None:
                                grau = 0
                        except (ValueError, TypeError):
                            grau = 0
                            
                        conta = {
                            'codigo': str(codigo).strip(),
                            'tipo': tipo,
                            'classificacao': classificacao,
                            'nome': nome,
                            'grau': grau
                        }
                        dados.append(conta)
                
                # Cria o dicionário final
                plano_contas = {
                    'empresa': empresa,
                    'contas': dados
                }
                
                # Cria diretório data se não existir
                os.makedirs('data', exist_ok=True)
                
                # Gera nome do arquivo JSON
                nome_base = os.path.splitext(os.path.basename(filepath))[0]
                json_filepath = os.path.join('data', f'{nome_base}.json')
                
                # Salva como JSON
                with open(json_filepath, 'w', encoding='utf-8') as f:
                    json.dump(plano_contas, f, ensure_ascii=False, indent=4)
                
                self.carregar_planos()  # Recarrega a lista
                tk.messagebox.showinfo("Sucesso", "Plano de contas adicionado com sucesso!", parent=self)
                
            except Exception as e:
                tk.messagebox.showerror(
                    "Erro",
                    f"Erro ao processar o arquivo:\n{str(e)}",
                    parent=self
                )
    
    def atualizar_plano(self):
        selecionado = self.tree.selection()
        if not selecionado:
            tk.messagebox.showwarning("Aviso", "Selecione um plano de contas para atualizar!", parent=self)
            return
            
        arquivo_json = self.tree.item(selecionado[0])['values'][0]
        
        filepath = filedialog.askopenfilename(
            title="Selecione o novo arquivo do Plano de Contas",
            filetypes=[("Excel files", "*.xls *.xlsx")],
            parent=self
        )
        
        if filepath:
            try:
                # Mesmo processo do novo_plano, mas mantendo o arquivo JSON existente
                # ... (código de processamento igual ao novo_plano)
                
                self.carregar_planos()
                tk.messagebox.showinfo("Sucesso", "Plano de contas atualizado com sucesso!", parent=self)
                
            except Exception as e:
                tk.messagebox.showerror(
                    "Erro",
                    f"Erro ao atualizar o arquivo:\n{str(e)}",
                    parent=self
                )
    
    def excluir_plano(self):
        selecionado = self.tree.selection()
        if not selecionado:
            tk.messagebox.showwarning("Aviso", "Selecione um plano de contas para excluir!", parent=self)
            return
            
        arquivo = self.tree.item(selecionado[0])['values'][0]
        empresa = self.tree.item(selecionado[0])['values'][1]
        
        if tk.messagebox.askyesno(
            "Confirmar Exclusão",
            f"Deseja realmente excluir o plano de contas da empresa:\n{empresa}?",
            parent=self
        ):
            try:
                os.remove(os.path.join('data', arquivo))
                self.carregar_planos()
                tk.messagebox.showinfo("Sucesso", "Plano de contas excluído com sucesso!", parent=self)
            except Exception as e:
                tk.messagebox.showerror(
                    "Erro",
                    f"Erro ao excluir o arquivo:\n{str(e)}",
                    parent=self
                )

    def visualizar_detalhes(self):
        selecionado = self.tree.selection()
        if not selecionado:
            return
            
        arquivo = self.tree.item(selecionado[0])['values'][0]
        caminho = os.path.join('data', arquivo)
        
        try:
            with open(caminho, 'r', encoding='utf-8') as f:
                dados = json.load(f)
                
            # Criar nova janela para mostrar os detalhes
            detalhes = tk.Toplevel(self)
            detalhes.title(f"Detalhes - {dados['empresa']}")
            detalhes.geometry("1200x600")
            
            # Criar Treeview para mostrar as contas
            tree = ttk.Treeview(detalhes, columns=("codigo", "tipo", "classificacao", "nome", "grau"), show="headings")
            tree.heading("codigo", text="Código")
            tree.heading("tipo", text="Tipo")
            tree.heading("classificacao", text="Classificação")
            tree.heading("nome", text="Nome")
            tree.heading("grau", text="Grau")
            
            # Configurar colunas
            tree.column("codigo", width=80)
            tree.column("tipo", width=50)
            tree.column("classificacao", width=100)
            tree.column("nome", width=800)
            tree.column("grau", width=50)
            
            # Adicionar scrollbar
            scrollbar = ttk.Scrollbar(detalhes, orient=tk.VERTICAL, command=tree.yview)
            tree.configure(yscrollcommand=scrollbar.set)
            
            # Layout
            tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
            scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
            
            # Configurar grid
            detalhes.columnconfigure(0, weight=1)
            detalhes.rowconfigure(0, weight=1)
            
            # Preencher dados
            for conta in dados['contas']:
                nome_tabulado = self.adicionar_tabulacao(conta['nome'], conta['classificacao'])
                tree.insert('', 'end', values=(
                    conta['codigo'],
                    conta['tipo'],
                    conta['classificacao'],
                    nome_tabulado,
                    conta['grau']
                ))
                
        except Exception as e:
            tk.messagebox.showerror("Erro", f"Erro ao abrir detalhes: {str(e)}", parent=self)
    
    def adicionar_tabulacao(self, nome, classificacao):
        niveis = len(classificacao.split('.')) - 1
        return '        ' * niveis + nome

class TelaSelecaoArquivos(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Seleção de Arquivos para Processamento")
        self.geometry("1000x600")
        
        # Configurações da janela
        self.extratos_vars = {}  # Dicionário para guardar as variáveis dos checkboxes
        self.plano_contas_var = tk.StringVar()
        self.setup_ui()
        self.carregar_arquivos()
    
    def setup_ui(self):
        # Frame principal com dois lados
        main_frame = ttk.Frame(self, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Frame esquerdo (Extratos)
        frame_extratos = ttk.LabelFrame(main_frame, text="Extratos Disponíveis", padding="5")
        frame_extratos.grid(row=0, column=0, padx=5, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Canvas e frame para scrolling dos extratos
        canvas_extratos = tk.Canvas(frame_extratos)
        scrollbar_extratos = ttk.Scrollbar(frame_extratos, orient=tk.VERTICAL, command=canvas_extratos.yview)
        self.frame_checkboxes = ttk.Frame(canvas_extratos)
        
        canvas_extratos.configure(yscrollcommand=scrollbar_extratos.set)
        
        # Frame direito (Planos de Contas)
        frame_planos = ttk.LabelFrame(main_frame, text="Plano de Contas", padding="5")
        frame_planos.grid(row=0, column=1, padx=5, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Canvas e frame para scrolling dos planos
        canvas_planos = tk.Canvas(frame_planos)
        scrollbar_planos = ttk.Scrollbar(frame_planos, orient=tk.VERTICAL, command=canvas_planos.yview)
        self.frame_radio = ttk.Frame(canvas_planos)
        
        canvas_planos.configure(yscrollcommand=scrollbar_planos.set)
        
        # Configurar canvas e scrollbars
        canvas_extratos.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar_extratos.grid(row=0, column=1, sticky=(tk.N, tk.S))
        canvas_extratos.create_window((0, 0), window=self.frame_checkboxes, anchor=tk.NW)
        
        canvas_planos.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar_planos.grid(row=0, column=1, sticky=(tk.N, tk.S))
        canvas_planos.create_window((0, 0), window=self.frame_radio, anchor=tk.NW)
        
        # Frame para botões
        frame_botoes = ttk.Frame(main_frame)
        frame_botoes.grid(row=1, column=0, columnspan=2, pady=10)
        
        # Botões
        self.btn_confirmar = ttk.Button(frame_botoes, text="Confirmar", command=self.confirmar_selecao)
        self.btn_confirmar.grid(row=0, column=0, padx=5)
        
        self.btn_cancelar = ttk.Button(frame_botoes, text="Cancelar", command=self.destroy)
        self.btn_cancelar.grid(row=0, column=1, padx=5)
        
        # Configurar grid weights
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(0, weight=1)
        frame_extratos.columnconfigure(0, weight=1)
        frame_extratos.rowconfigure(0, weight=1)
        frame_planos.columnconfigure(0, weight=1)
        frame_planos.rowconfigure(0, weight=1)
        
        # Configurar eventos de scroll
        self.frame_checkboxes.bind('<Configure>', lambda e: canvas_extratos.configure(scrollregion=canvas_extratos.bbox('all')))
        self.frame_radio.bind('<Configure>', lambda e: canvas_planos.configure(scrollregion=canvas_planos.bbox('all')))
    
    def carregar_arquivos(self):
        # Carregar extratos com checkboxes
        if os.path.exists('extratos'):
            for arquivo in sorted(os.listdir('extratos')):
                if arquivo.endswith('.json'):
                    var = tk.BooleanVar()
                    self.extratos_vars[arquivo] = var
                    ttk.Checkbutton(
                        self.frame_checkboxes,
                        text=arquivo,
                        variable=var
                    ).pack(anchor=tk.W, padx=5, pady=2)
        
        # Carregar planos de contas com radio buttons
        if os.path.exists('data'):
            for arquivo in sorted(os.listdir('data')):
                if arquivo.endswith('.json'):
                    with open(os.path.join('data', arquivo), 'r', encoding='utf-8') as f:
                        dados = json.load(f)
                        ttk.Radiobutton(
                            self.frame_radio,
                            text=f"{dados['empresa']} ({arquivo})",
                            value=arquivo,
                            variable=self.plano_contas_var
                        ).pack(anchor=tk.W, padx=5, pady=2)
    
    def confirmar_selecao(self):
        # Verificar seleção de extratos
        extratos_selecionados = [arquivo for arquivo, var in self.extratos_vars.items() if var.get()]
        if not extratos_selecionados:
            tk.messagebox.showwarning("Aviso", "Selecione pelo menos um extrato!", parent=self)
            return
        
        # Verificar seleção do plano de contas
        plano_selecionado = self.plano_contas_var.get()
        if not plano_selecionado:
            tk.messagebox.showwarning("Aviso", "Selecione um plano de contas!", parent=self)
            return
        
        # Guardar seleções como atributos da instância
        self.extratos_selecionados = extratos_selecionados
        self.plano_contas_selecionado = plano_selecionado
        
        # Fechar janela
        self.destroy()

class ExtratoViewer(tk.Toplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Visualização de Extratos")
        self.geometry("1200x600")  # Aumentei a largura para acomodar mais texto
        self.setup_ui()
        
    def setup_ui(self):
        # Frame principal
        main_frame = ttk.Frame(self, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Lista de extratos
        self.tree = ttk.Treeview(main_frame, columns=("arquivo_json", "arquivo_excel"), show="headings")
        self.tree.heading("arquivo_json", text="Arquivo JSON")
        self.tree.heading("arquivo_excel", text="Arquivo Excel Original")
        
        # Configurar larguras das colunas
        self.tree.column("arquivo_json", width=300)
        self.tree.column("arquivo_excel", width=500)
        
        # Scrollbar
        scrollbar = ttk.Scrollbar(main_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)
        
        # Frame para botões
        btn_frame = ttk.Frame(main_frame)
        
        # Botões
        btn_novo = ttk.Button(btn_frame, text="Novo Extrato", command=self.novo_extrato)
        btn_visualizar = ttk.Button(btn_frame, text="Visualizar Detalhes", command=self.visualizar_detalhes)
        btn_atualizar = ttk.Button(btn_frame, text="Atualizar Extrato", command=self.atualizar_extrato)
        btn_excluir = ttk.Button(btn_frame, text="Excluir Extrato", command=self.excluir_extrato)
        
        # Layout
        self.tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
        btn_frame.grid(row=1, column=0, pady=10)
        
        btn_novo.grid(row=0, column=0, padx=5)
        btn_visualizar.grid(row=0, column=1, padx=5)
        btn_atualizar.grid(row=0, column=2, padx=5)
        btn_excluir.grid(row=0, column=3, padx=5)
        
        # Configurar grid
        self.columnconfigure(0, weight=1)
        self.rowconfigure(0, weight=1)
        main_frame.columnconfigure(0, weight=1)
        main_frame.rowconfigure(0, weight=1)
        
        # Carregar extratos
        self.carregar_extratos()
        
    def carregar_extratos(self):
        # Limpar itens existentes
        for item in self.tree.get_children():
            self.tree.delete(item)
            
        if os.path.exists('extratos'):
            for arquivo in sorted(os.listdir('extratos')):
                if arquivo.endswith('.json'):
                    with open(os.path.join('extratos', arquivo), 'r', encoding='utf-8') as f:
                        dados = json.load(f)
                        nome_excel = dados['arquivo_origem']
                        self.tree.insert('', 'end', values=(arquivo, f"({nome_excel})"))

    def excluir_extrato(self):
        selecionado = self.tree.selection()
        if not selecionado:
            tk.messagebox.showwarning("Aviso", "Selecione um extrato para excluir!", parent=self)
            return
            
        arquivo = self.tree.item(selecionado[0])['values'][0]
        
        if tk.messagebox.askyesno(
            "Confirmar Exclusão",
            f"Deseja realmente excluir o extrato:\n{arquivo}?",
            parent=self
        ):
            try:
                os.remove(os.path.join('extratos', arquivo))
                self.carregar_extratos()  # Recarrega a lista
                tk.messagebox.showinfo("Sucesso", "Extrato excluído com sucesso!", parent=self)
            except Exception as e:
                tk.messagebox.showerror(
                    "Erro",
                    f"Erro ao excluir o arquivo:\n{str(e)}",
                    parent=self
                )
    
    def atualizar_extrato(self):
        selecionado = self.tree.selection()
        if not selecionado:
            tk.messagebox.showwarning("Aviso", "Selecione um extrato para atualizar!", parent=self)
            return
            
        arquivo_json = self.tree.item(selecionado[0])['values'][0]
        
        # Solicitar novo arquivo Excel
        filepath = filedialog.askopenfilename(
            title="Selecione o novo arquivo de Extrato",
            filetypes=[("Excel files", "*.xlsx")],
            parent=self
        )
        
        if filepath:
            try:
                # Lê o arquivo Excel
                df = pd.read_excel(filepath)
                
                # Lista de possíveis nomes para as colunas que queremos remover
                colunas_para_remover = [
                    'Código', 'Cod', 'Cod.', 'Codigo',
                    'Doc', 'Doc.', 'Documento', 'Nº Doc', 'N Doc', 'Num Doc',
                    'Saldo dia', 'Saldo Dia', 'Saldo do dia', 'Saldo'
                ]
                
                # Remove as colunas que não precisamos (ignorando case)
                colunas_atuais = df.columns.tolist()
                for coluna in colunas_atuais:
                    if any(remover.lower() in coluna.lower() for remover in colunas_para_remover):
                        df = df.drop(columns=[coluna])
                
                # Remove linhas totalmente vazias
                df = df.dropna(how='all')
                
                # Remove linhas com descrição vazia ou nula
                coluna_descricao = [col for col in df.columns if 'desc' in col.lower()][0]
                df = df.dropna(subset=[coluna_descricao])
                df = df[df[coluna_descricao].astype(str).str.strip() != '']
                
                # Converte DataFrame para dicionário, tratando valores especiais
                dados = []
                for _, row in df.iterrows():
                    linha = {}
                    for coluna in df.columns:
                        linha[str(coluna)] = self.converter_para_serializavel(row[coluna])
                    dados.append(linha)
                
                # Atualiza o arquivo JSON
                dados_extrato = {
                    'arquivo_origem': os.path.basename(filepath),
                    'data_processamento': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    'dados': dados
                }
                
                # Salva o arquivo atualizado
                with open(os.path.join('extratos', arquivo_json), 'w', encoding='utf-8') as f:
                    json.dump(dados_extrato, f, ensure_ascii=False, indent=4)
                
                self.carregar_extratos()  # Recarrega a lista
                tk.messagebox.showinfo("Sucesso", "Extrato atualizado com sucesso!", parent=self)
                
            except Exception as e:
                tk.messagebox.showerror(
                    "Erro",
                    f"Erro ao atualizar o arquivo:\n{str(e)}",
                    parent=self
                )
    
    def converter_para_serializavel(self, valor):
        """Converte valores para formatos serializáveis em JSON"""
        if pd.isna(valor):
            return None
        elif isinstance(valor, (pd.Timestamp, pd._libs.tslibs.timestamps.Timestamp)):
            return valor.strftime('%Y-%m-%d')
        elif isinstance(valor, datetime.datetime):
            return valor.strftime('%Y-%m-%d')
        elif isinstance(valor, datetime.date):
            return valor.strftime('%Y-%m-%d')
        elif isinstance(valor, (np.int64, np.int32, np.int16, np.int8)):
            return int(valor)
        elif isinstance(valor, (np.float64, np.float32)):
            return float(valor)
        elif isinstance(valor, decimal.Decimal):
            return str(valor)
        return valor

    def visualizar_detalhes(self):
        selecionado = self.tree.selection()
        if not selecionado:
            return
            
        arquivo = self.tree.item(selecionado[0])['values'][0]
        caminho = os.path.join('extratos', arquivo)
        
        try:
            with open(caminho, 'r', encoding='utf-8') as f:
                dados = json.load(f)
            
            # Criar nova janela para mostrar os detalhes
            detalhes = tk.Toplevel(self)
            detalhes.title(f"Detalhes - {dados['arquivo_origem']}")
            detalhes.geometry("1200x600")
            
            # Criar Treeview para mostrar os dados
            tree = ttk.Treeview(detalhes, show="headings")
            
            # Configurar colunas baseado nos dados
            if dados['dados']:
                colunas = list(dados['dados'][0].keys())
                tree["columns"] = colunas
                for col in colunas:
                    tree.heading(col, text=col)
                    tree.column(col, width=100)
            
            # Adicionar scrollbar
            scrollbar = ttk.Scrollbar(detalhes, orient=tk.VERTICAL, command=tree.yview)
            tree.configure(yscrollcommand=scrollbar.set)
            
            # Layout
            tree.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
            scrollbar.grid(row=0, column=1, sticky=(tk.N, tk.S))
            
            # Configurar grid
            detalhes.columnconfigure(0, weight=1)
            detalhes.rowconfigure(0, weight=1)
            
            # Preencher dados
            for linha in dados['dados']:
                tree.insert('', 'end', values=list(linha.values()))
                
        except Exception as e:
            tk.messagebox.showerror("Erro", f"Erro ao abrir detalhes: {str(e)}", parent=self)

    def novo_extrato(self):
        # Solicitar arquivo Excel
        filepath = filedialog.askopenfilename(
            title="Selecione o arquivo de Extrato",
            filetypes=[("Excel files", "*.xlsx")],
            parent=self
        )
        
        if filepath:
            try:
                # Lê o arquivo Excel
                df = pd.read_excel(filepath)
                
                # Lista de possíveis nomes para as colunas que queremos remover
                colunas_para_remover = [
                    'Código', 'Cod', 'Cod.', 'Codigo',
                    'Doc', 'Doc.', 'Documento', 'Nº Doc', 'N Doc', 'Num Doc',
                    'Saldo dia', 'Saldo Dia', 'Saldo do dia', 'Saldo'
                ]
                
                # Remove as colunas que não precisamos (ignorando case)
                colunas_atuais = df.columns.tolist()
                for coluna in colunas_atuais:
                    if any(remover.lower() in coluna.lower() for remover in colunas_para_remover):
                        df = df.drop(columns=[coluna])
                
                # Remove linhas totalmente vazias
                df = df.dropna(how='all')
                
                # Remove linhas com descrição vazia ou nula
                coluna_descricao = [col for col in df.columns if 'desc' in col.lower()][0]
                df = df.dropna(subset=[coluna_descricao])
                df = df[df[coluna_descricao].astype(str).str.strip() != '']
                
                # Converte DataFrame para dicionário, tratando valores especiais
                dados = []
                for _, row in df.iterrows():
                    linha = {}
                    for coluna in df.columns:
                        linha[str(coluna)] = self.converter_para_serializavel(row[coluna])
                    dados.append(linha)
                
                # Cria diretório extratos se não existir
                os.makedirs('extratos', exist_ok=True)
                
                # Extrai informações básicas
                dados_extrato = {
                    'arquivo_origem': os.path.basename(filepath),
                    'data_processamento': datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
                    'dados': dados
                }
                
                # Gera nome do arquivo JSON
                nome_base = os.path.splitext(os.path.basename(filepath))[0]
                json_filepath = os.path.join('extratos', f'{nome_base}.json')
                
                # Verifica se já existe um arquivo com esse nome
                if os.path.exists(json_filepath):
                    if not tk.messagebox.askyesno(
                        "Arquivo Existente",
                        f"Já existe um extrato com o nome {nome_base}.\nDeseja substituir?",
                        parent=self
                    ):
                        return
                
                # Salva como JSON
                with open(json_filepath, 'w', encoding='utf-8') as f:
                    json.dump(dados_extrato, f, ensure_ascii=False, indent=4)
                
                self.carregar_extratos()  # Recarrega a lista
                tk.messagebox.showinfo("Sucesso", "Extrato adicionado com sucesso!", parent=self)
                
            except Exception as e:
                tk.messagebox.showerror(
                    "Erro",
                    f"Erro ao processar o arquivo:\n{str(e)}",
                    parent=self
                )

class MainWindow:
    def __init__(self, master):
        self.master = master
        self.setup_ui()
    
    def setup_ui(self):
        # Configuração básica da janela
        self.master.geometry("800x600")
        
        # Frame principal
        self.main_frame = ttk.Frame(self.master, padding="10")
        self.main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))
        
        # Título
        self.title_label = ttk.Label(
            self.main_frame, 
            text="Sistema de Automação Contábil",
            font=("Arial", 16, "bold")
        )
        self.title_label.grid(row=0, column=0, pady=10)

        # Frame para botões
        self.btn_frame = ttk.Frame(self.main_frame)
        self.btn_frame.grid(row=1, column=0, pady=10)

        # Botão de visualização de planos
        self.view_button = ttk.Button(
            self.btn_frame,
            text="Visualizar Planos de Contas",
            command=self.abrir_visualizador
        )
        self.view_button.grid(row=0, column=0, padx=5)

        # Botão para visualizar extratos
        self.view_extratos_button = ttk.Button(
            self.btn_frame,
            text="Visualizar Extratos",
            command=self.abrir_visualizador_extratos
        )
        self.view_extratos_button.grid(row=0, column=1, padx=5)

        # Botão para gerar lançamentos
        self.lancamentos_button = ttk.Button(
            self.btn_frame,
            text="Gerar Lançamentos",
            command=self.abrir_selecao_arquivos
        )
        self.lancamentos_button.grid(row=0, column=2, padx=5)

    def abrir_visualizador(self):
        PlanoContasViewer(self.master)

    def abrir_visualizador_extratos(self):
        ExtratoViewer(self.master)

    def abrir_selecao_arquivos(self):
        tela_selecao = TelaSelecaoArquivos(self.master)
        tela_selecao.grab_set()  # Torna a janela modal
        self.master.wait_window(tela_selecao)  # Espera a janela ser fechada
        
        # Verificar seleções
        if hasattr(tela_selecao, 'extratos_selecionados') and hasattr(tela_selecao, 'plano_contas_selecionado'):
            if tela_selecao.extratos_selecionados and tela_selecao.plano_contas_selecionado:
                print("Extratos selecionados:", tela_selecao.extratos_selecionados)
                print("Plano de contas selecionado:", tela_selecao.plano_contas_selecionado)
                # Aqui você pode continuar com o processamento dos arquivos selecionados