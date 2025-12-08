import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import gspread
from google.oauth2.service_account import Credentials
import re
import json
import os
from datetime import datetime, timedelta
from tkinter import filedialog

ARQUIVO_PREFIXOS = "prefixos.json"

prefixos_padrao = {
    "TECLADO": "T",
    "MOUSE": "M",
    "MONITOR": "MN",
    "IMPRESSORA": "I"
}

colunas = [
    "Registro",
    "Descrição do Equipamento",
    "Marca/ Modelo",
    "Numero de Serie",
    "Data da Aquisição",
    "Condição",
    "Local Instalado",
    "Detalhes",
    "Ainda no Local de Instalação",
    "Novo Local de Instalação",
    "Data de Saida"
]

class ControleMateriaisApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Controle de Materiais de TI - SEMED")
        self.root.geometry("1200x700")
        self.root.configure(bg="#f0f0f0")
        
        self.planilha = None
        self.prefixos = self.carregar_prefixos()
        
        # Estilo
        self.style = ttk.Style()
        self.style.theme_use('clam')
        
        # Conectar à planilha
        self.conectar_planilha()
        
        # Criar interface
        self.criar_interface()
        
    def carregar_prefixos(self):
        if os.path.exists(ARQUIVO_PREFIXOS):
            with open(ARQUIVO_PREFIXOS, "r") as f:
                dados = json.load(f)
                return {**prefixos_padrao, **dados}
        return prefixos_padrao.copy()
    
    def salvar_prefixos(self):
        aprendidos = {k:v for k,v in self.prefixos.items() if k not in prefixos_padrao}
        with open(ARQUIVO_PREFIXOS, "w") as f:
            json.dump(aprendidos, f)
    
    def conectar_planilha(self):
        try:
            nome_planilha = "Controle de Materiais de TI - SEMED"
            credenciais_json = "credenciais.json"
            scope = [
                "https://spreadsheets.google.com/feeds",
                "https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive.file",
                "https://www.googleapis.com/auth/drive"
            ]
            creds = Credentials.from_service_account_file(credenciais_json, scopes=scope)
            cliente = gspread.authorize(creds)
            self.planilha = cliente.open(nome_planilha).sheet1
        except Exception as e:
            messagebox.showerror("Erro de Conexão", f"Não foi possível conectar à planilha:\n{str(e)}")
            self.root.destroy()
    
    def criar_interface(self):
        # Frame principal
        main_frame = tk.Frame(self.root, bg="#f0f0f0")
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Título
        titulo = tk.Label(main_frame, text="🖥️ Controle de Materiais de TI - SEMED", 
                         font=("Arial", 18, "bold"), bg="#f0f0f0", fg="#2c3e50")
        titulo.pack(pady=(0, 10))
        
        # Frame de botões
        btn_frame = tk.Frame(main_frame, bg="#f0f0f0")
        btn_frame.pack(fill=tk.X, pady=10)
        
        # Botões principais
        ttk.Button(btn_frame, text="📋 Visualizar Registros", 
                  command=self.visualizar_registros).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="➕ Inserir Novo Registro", 
                  command=self.abrir_janela_inserir).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="🔍 Buscar Registros", 
                  command=self.abrir_janela_buscar).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="✏️ Editar Registro", 
                  command=self.abrir_janela_editar).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="📄 Imprimir Relatório", 
                  command=self.abrir_janela_relatorio).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="🔄 Atualizar", 
                  command=self.visualizar_registros).pack(side=tk.LEFT, padx=5)
        
        # Frame para tabela
        tabela_frame = tk.Frame(main_frame, bg="white", relief=tk.SUNKEN, bd=2)
        tabela_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Scrollbars
        scroll_y = ttk.Scrollbar(tabela_frame, orient=tk.VERTICAL)
        scroll_x = ttk.Scrollbar(tabela_frame, orient=tk.HORIZONTAL)
        
        # Treeview para exibir dados
        self.tree = ttk.Treeview(tabela_frame, 
                                 yscrollcommand=scroll_y.set,
                                 xscrollcommand=scroll_x.set,
                                 selectmode='browse')
        
        scroll_y.config(command=self.tree.yview)
        scroll_x.config(command=self.tree.xview)
        
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        self.tree.pack(fill=tk.BOTH, expand=True)
        
        # Configurar colunas
        self.tree['columns'] = colunas
        self.tree.column("#0", width=0, stretch=tk.NO)
        
        for col in colunas:
            if col == "Registro":
                self.tree.column(col, anchor=tk.W, width=150, minwidth=100)
            else:
                self.tree.column(col, anchor=tk.W, width=120, minwidth=80)
            self.tree.heading(col, text=col, anchor=tk.W)
        
        # Status bar
        self.status_bar = tk.Label(main_frame, text="✅ Conectado à planilha", 
                                   bg="#2ecc71", fg="white", 
                                   font=("Arial", 10), relief=tk.SUNKEN, anchor=tk.W)
        self.status_bar.pack(fill=tk.X, pady=(10, 0))
        
        # Carregar dados iniciais
        self.visualizar_registros()
    
    def visualizar_registros(self):
        try:
            # Limpar tabela
            for item in self.tree.get_children():
                self.tree.delete(item)
            
            # Buscar dados
            dados = self.planilha.get_all_records()
            
            if not dados:
                self.status_bar.config(text="📭 Nenhum registro encontrado", bg="#e74c3c")
                return
            
            # Inserir dados na tabela
            for linha in dados:
                valores = [linha.get(col, "") for col in colunas]
                self.tree.insert("", tk.END, values=valores)
            
            self.status_bar.config(text=f"✅ {len(dados)} registros carregados", bg="#2ecc71")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar registros:\n{str(e)}")
            self.status_bar.config(text="❌ Erro ao carregar dados", bg="#e74c3c")
    
    def gerar_id_categoria(self, prefixo):
        dados = self.planilha.get_all_records()
        ids_categoria = [linha.get("Registro", "") for linha in dados if linha.get("Registro", "").startswith(prefixo)]
        numeros = [int(re.sub(r'\D', '', i)) for i in ids_categoria if i]
        proximo = max(numeros)+1 if numeros else 1
        return f"{prefixo}{proximo:04d}"
    
    def determinar_prefixo(self, descricao):
        descricao_upper = descricao.upper()
        for chave, val in self.prefixos.items():
            if chave in descricao_upper or any(p in descricao_upper for p in chave.split()):
                return val
        palavras = [c for c in descricao_upper if c.isalpha()]
        novo_prefixo = "".join(palavras[:2]).upper() if len(palavras)>=2 else "".join(palavras).upper()
        self.prefixos[descricao_upper] = novo_prefixo
        self.salvar_prefixos()
        return novo_prefixo
    
    def abrir_janela_inserir(self):
        janela = tk.Toplevel(self.root)
        janela.title("➕ Inserir Novo Registro")
        janela.geometry("600x700")
        janela.configure(bg="#f0f0f0")
        
        # Frame com scroll
        canvas = tk.Canvas(janela, bg="#f0f0f0")
        scrollbar = ttk.Scrollbar(janela, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg="#f0f0f0")
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Campos de entrada
        campos = {}
        
        # Descrição do equipamento (primeiro campo especial)
        tk.Label(scrollable_frame, text="Descrição do Equipamento *", 
                bg="#f0f0f0", font=("Arial", 10, "bold")).grid(row=0, column=0, sticky=tk.W, padx=10, pady=5)
        campos["Descrição do Equipamento"] = tk.Entry(scrollable_frame, width=40)
        campos["Descrição do Equipamento"].grid(row=0, column=1, padx=10, pady=5)
        
        # ID sugerido
        tk.Label(scrollable_frame, text="Registro (gerado automaticamente)", 
                bg="#f0f0f0", font=("Arial", 10)).grid(row=1, column=0, sticky=tk.W, padx=10, pady=5)
        lbl_id = tk.Label(scrollable_frame, text="", bg="#f0f0f0", fg="#3498db", font=("Arial", 10, "bold"))
        lbl_id.grid(row=1, column=1, sticky=tk.W, padx=10, pady=5)
        
        def atualizar_id(event=None):
            desc = campos["Descrição do Equipamento"].get().strip()
            if desc:
                prefixo = self.determinar_prefixo(desc)
                novo_id = self.gerar_id_categoria(prefixo)
                lbl_id.config(text=f"💡 {novo_id} (Prefixo: {prefixo})")
        
        campos["Descrição do Equipamento"].bind("<KeyRelease>", atualizar_id)
        
        # Demais campos
        row = 2
        for col in colunas[2:]:
            tk.Label(scrollable_frame, text=col, bg="#f0f0f0", font=("Arial", 10)).grid(
                row=row, column=0, sticky=tk.W, padx=10, pady=5)
            
            if col == "Condição":
                campos[col] = ttk.Combobox(scrollable_frame, width=37, 
                                          values=["Novo", "Bom", "Regular", "Ruim", "Quebrado"])
                campos[col].grid(row=row, column=1, padx=10, pady=5)
            elif col == "Ainda no Local de Instalação":
                campos[col] = ttk.Combobox(scrollable_frame, width=37, values=["Sim", "Não"])
                campos[col].grid(row=row, column=1, padx=10, pady=5)
            elif "Data" in col:
                frame_data = tk.Frame(scrollable_frame, bg="#f0f0f0")
                frame_data.grid(row=row, column=1, padx=10, pady=5, sticky=tk.W)
                campos[col] = tk.Entry(frame_data, width=30)
                campos[col].pack(side=tk.LEFT)
                tk.Button(frame_data, text="Hoje", 
                         command=lambda c=col: campos[c].insert(0, datetime.now().strftime("%d/%m/%Y"))).pack(side=tk.LEFT, padx=5)
            else:
                campos[col] = tk.Entry(scrollable_frame, width=40)
                campos[col].grid(row=row, column=1, padx=10, pady=5)
            row += 1
        
        def salvar():
            descricao = campos["Descrição do Equipamento"].get().strip()
            if not descricao:
                messagebox.showwarning("Atenção", "A descrição do equipamento é obrigatória!")
                return
            
            prefixo = self.determinar_prefixo(descricao)
            novo_id = self.gerar_id_categoria(prefixo)
            
            nova_linha = [novo_id, descricao]
            for col in colunas[2:]:
                valor = campos[col].get().strip() if col in campos else ""
                nova_linha.append(valor)
            
            try:
                self.planilha.append_row(nova_linha)
                messagebox.showinfo("Sucesso", f"✅ Registro {novo_id} inserido com sucesso!")
                janela.destroy()
                self.visualizar_registros()
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao inserir registro:\n{str(e)}")
        
        # Botões
        btn_frame = tk.Frame(scrollable_frame, bg="#f0f0f0")
        btn_frame.grid(row=row, column=0, columnspan=2, pady=20)
        
        ttk.Button(btn_frame, text="💾 Salvar", command=salvar).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="❌ Cancelar", command=janela.destroy).pack(side=tk.LEFT, padx=5)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
    
    def abrir_janela_buscar(self):
        janela = tk.Toplevel(self.root)
        janela.title("🔍 Buscar Registros")
        janela.geometry("900x500")
        janela.configure(bg="#f0f0f0")
        
        # Frame de busca
        busca_frame = tk.Frame(janela, bg="#f0f0f0")
        busca_frame.pack(fill=tk.X, padx=10, pady=10)
        
        tk.Label(busca_frame, text="Buscar em:", bg="#f0f0f0", font=("Arial", 10)).pack(side=tk.LEFT, padx=5)
        combo_coluna = ttk.Combobox(busca_frame, values=colunas, width=25)
        combo_coluna.current(0)
        combo_coluna.pack(side=tk.LEFT, padx=5)
        
        tk.Label(busca_frame, text="Valor:", bg="#f0f0f0", font=("Arial", 10)).pack(side=tk.LEFT, padx=5)
        entry_valor = tk.Entry(busca_frame, width=30)
        entry_valor.pack(side=tk.LEFT, padx=5)
        
        # Treeview para resultados
        resultado_frame = tk.Frame(janela, bg="white", relief=tk.SUNKEN, bd=2)
        resultado_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        scroll_y = ttk.Scrollbar(resultado_frame, orient=tk.VERTICAL)
        scroll_x = ttk.Scrollbar(resultado_frame, orient=tk.HORIZONTAL)
        
        tree_busca = ttk.Treeview(resultado_frame, 
                                  yscrollcommand=scroll_y.set,
                                  xscrollcommand=scroll_x.set)
        
        scroll_y.config(command=tree_busca.yview)
        scroll_x.config(command=tree_busca.xview)
        
        scroll_y.pack(side=tk.RIGHT, fill=tk.Y)
        scroll_x.pack(side=tk.BOTTOM, fill=tk.X)
        tree_busca.pack(fill=tk.BOTH, expand=True)
        
        tree_busca['columns'] = colunas
        tree_busca.column("#0", width=0, stretch=tk.NO)
        
        for col in colunas:
            tree_busca.column(col, anchor=tk.W, width=100)
            tree_busca.heading(col, text=col, anchor=tk.W)
        
        def buscar():
            for item in tree_busca.get_children():
                tree_busca.delete(item)
            
            coluna = combo_coluna.get()
            valor = entry_valor.get().strip()
            
            if not valor:
                messagebox.showwarning("Atenção", "Digite um valor para buscar!")
                return
            
            dados = self.planilha.get_all_records()
            encontrados = [linha for linha in dados if str(linha.get(coluna, "")).lower() == valor.lower()]
            
            if not encontrados:
                messagebox.showinfo("Resultado", "⚠ Nenhum registro encontrado.")
                return
            
            for linha in encontrados:
                valores = [linha.get(col, "") for col in colunas]
                tree_busca.insert("", tk.END, values=valores)
            
            messagebox.showinfo("Resultado", f"🔹 {len(encontrados)} registro(s) encontrado(s)!")
        
        ttk.Button(busca_frame, text="🔍 Buscar", command=buscar).pack(side=tk.LEFT, padx=5)
    
    def abrir_janela_editar(self):
        janela = tk.Toplevel(self.root)
        janela.title("✏️ Editar Registro")
        janela.geometry("700x600")
        janela.configure(bg="#f0f0f0")
        
        # Buscar registro
        tk.Label(janela, text="Digite o Registro (ID) para editar:", 
                bg="#f0f0f0", font=("Arial", 12, "bold")).pack(pady=10)
        
        entry_registro = tk.Entry(janela, width=30, font=("Arial", 12))
        entry_registro.pack(pady=5)
        
        campos_frame = tk.Frame(janela, bg="#f0f0f0")
        campos_edicao = {}
        
        def carregar_registro():
            for widget in campos_frame.winfo_children():
                widget.destroy()
            campos_edicao.clear()
            
            registro = entry_registro.get().strip()
            if not registro:
                messagebox.showwarning("Atenção", "Digite um registro válido!")
                return
            
            try:
                cel = self.planilha.find(registro, in_column=1)
                linha_num = cel.row
                
                dados_linha = self.planilha.row_values(linha_num)
                
                # Criar campos de edição
                row = 0
                for i, col in enumerate(colunas):
                    tk.Label(campos_frame, text=col, bg="#f0f0f0", 
                            font=("Arial", 10, "bold")).grid(row=row, column=0, sticky=tk.W, padx=10, pady=5)
                    
                    if col == "Registro":
                        lbl = tk.Label(campos_frame, text=dados_linha[i] if i < len(dados_linha) else "", 
                                      bg="#f0f0f0", fg="#3498db", font=("Arial", 10))
                        lbl.grid(row=row, column=1, sticky=tk.W, padx=10, pady=5)
                        campos_edicao[col] = {"widget": lbl, "col_idx": i+1, "tipo": "label"}
                    else:
                        entry = tk.Entry(campos_frame, width=40)
                        entry.insert(0, dados_linha[i] if i < len(dados_linha) else "")
                        entry.grid(row=row, column=1, padx=10, pady=5)
                        campos_edicao[col] = {"widget": entry, "col_idx": i+1, "linha": linha_num, "tipo": "entry"}
                    
                    row += 1
                
                def salvar_edicao():
                    try:
                        for col, info in campos_edicao.items():
                            if info["tipo"] == "entry":
                                novo_valor = info["widget"].get().strip()
                                self.planilha.update_cell(info["linha"], info["col_idx"], novo_valor)
                        
                        messagebox.showinfo("Sucesso", "✅ Registro atualizado com sucesso!")
                        janela.destroy()
                        self.visualizar_registros()
                    except Exception as e:
                        messagebox.showerror("Erro", f"Erro ao salvar edição:\n{str(e)}")
                
                btn_salvar = ttk.Button(campos_frame, text="💾 Salvar Alterações", command=salvar_edicao)
                btn_salvar.grid(row=row, column=0, columnspan=2, pady=20)
                
            except Exception as e:
                messagebox.showerror("Erro", f"Registro não encontrado:\n{str(e)}")
        
        ttk.Button(janela, text="🔍 Carregar Registro", command=carregar_registro).pack(pady=10)
        
        campos_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
    
    def abrir_janela_relatorio(self):
        janela = tk.Toplevel(self.root)
        janela.title("📄 Imprimir Relatório")
        janela.geometry("500x550")
        janela.configure(bg="#f0f0f0")
        
        # Título
        tk.Label(janela, text="Gerar Relatório de Materiais", 
                bg="#f0f0f0", font=("Arial", 14, "bold")).pack(pady=20)
        
        # Frame de opções
        opcoes_frame = tk.LabelFrame(janela, text="Selecione o Período", 
                                     bg="#f0f0f0", font=("Arial", 11, "bold"))
        opcoes_frame.pack(fill=tk.X, padx=20, pady=10)
        
        periodo_var = tk.StringVar(value="todos")
        
        tk.Radiobutton(opcoes_frame, text="📅 Todos os Registros", 
                      variable=periodo_var, value="todos", 
                      bg="#f0f0f0", font=("Arial", 10)).pack(anchor=tk.W, padx=20, pady=10)
        
        tk.Radiobutton(opcoes_frame, text="📆 Semanal (últimos 7 dias)", 
                      variable=periodo_var, value="semanal", 
                      bg="#f0f0f0", font=("Arial", 10)).pack(anchor=tk.W, padx=20, pady=10)
        
        tk.Radiobutton(opcoes_frame, text="📊 Mensal (últimos 30 dias)", 
                      variable=periodo_var, value="mensal", 
                      bg="#f0f0f0", font=("Arial", 10)).pack(anchor=tk.W, padx=20, pady=10)
        
        tk.Radiobutton(opcoes_frame, text="📈 Anual (último ano)", 
                      variable=periodo_var, value="anual", 
                      bg="#f0f0f0", font=("Arial", 10)).pack(anchor=tk.W, padx=20, pady=10)
        
        # Frame de formato
        formato_frame = tk.LabelFrame(janela, text="Formato de Saída", 
                                      bg="#f0f0f0", font=("Arial", 11, "bold"))
        formato_frame.pack(fill=tk.X, padx=20, pady=10)
        
        formato_var = tk.StringVar(value="txt")
        
        tk.Radiobutton(formato_frame, text="📝 Arquivo TXT", 
                      variable=formato_var, value="txt", 
                      bg="#f0f0f0", font=("Arial", 10)).pack(side=tk.LEFT, padx=20, pady=10)
        
        tk.Radiobutton(formato_frame, text="📋 Arquivo CSV", 
                      variable=formato_var, value="csv", 
                      bg="#f0f0f0", font=("Arial", 10)).pack(side=tk.LEFT, padx=20, pady=10)
        
        def gerar_relatorio():
            periodo = periodo_var.get()
            formato = formato_var.get()
            
            try:
                dados = self.planilha.get_all_records()
                dados_filtrados = []
                
                hoje = datetime.now()
                
                for linha in dados:
                    incluir = False
                    
                    if periodo == "todos":
                        incluir = True
                    else:
                        # Tentar pegar a data de aquisição
                        data_str = linha.get("Data da Aquisição", "")
                        if data_str:
                            try:
                                # Tentar formatos comuns de data
                                for fmt in ["%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y"]:
                                    try:
                                        data_registro = datetime.strptime(data_str, fmt)
                                        break
                                    except:
                                        continue
                                
                                diff_dias = (hoje - data_registro).days
                                
                                if periodo == "semanal" and diff_dias <= 7:
                                    incluir = True
                                elif periodo == "mensal" and diff_dias <= 30:
                                    incluir = True
                                elif periodo == "anual" and diff_dias <= 365:
                                    incluir = True
                            except:
                                pass
                    
                    if incluir:
                        dados_filtrados.append(linha)
                
                if not dados_filtrados:
                    messagebox.showwarning("Aviso", "Nenhum registro encontrado para o período selecionado.")
                    return
                
                # Solicitar local para salvar
                extensao = "txt" if formato == "txt" else "csv"
                arquivo = filedialog.asksaveasfilename(
                    defaultextension=f".{extensao}",
                    filetypes=[(f"Arquivo {extensao.upper()}", f"*.{extensao}"), ("Todos os arquivos", "*.*")],
                    initialfile=f"relatorio_{periodo}_{hoje.strftime('%Y%m%d')}.{extensao}"
                )
                
                if not arquivo:
                    return
                
                # Gerar relatório
                if formato == "txt":
                    self.gerar_relatorio_txt(arquivo, dados_filtrados, periodo)
                else:
                    self.gerar_relatorio_csv(arquivo, dados_filtrados, periodo)
                
                messagebox.showinfo("Sucesso", f"✅ Relatório gerado com sucesso!\n\nArquivo: {arquivo}\nRegistros: {len(dados_filtrados)}")
                janela.destroy()
                
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao gerar relatório:\n{str(e)}")
        
        def imprimir_direto():
            periodo = periodo_var.get()
            
            try:
                dados = self.planilha.get_all_records()
                dados_filtrados = []
                
                hoje = datetime.now()
                
                for linha in dados:
                    incluir = False
                    
                    if periodo == "todos":
                        incluir = True
                    else:
                        data_str = linha.get("Data da Aquisição", "")
                        if data_str:
                            try:
                                for fmt in ["%d/%m/%Y", "%d-%m-%Y", "%Y-%m-%d", "%d/%m/%y"]:
                                    try:
                                        data_registro = datetime.strptime(data_str, fmt)
                                        break
                                    except:
                                        continue
                                
                                diff_dias = (hoje - data_registro).days
                                
                                if periodo == "semanal" and diff_dias <= 7:
                                    incluir = True
                                elif periodo == "mensal" and diff_dias <= 30:
                                    incluir = True
                                elif periodo == "anual" and diff_dias <= 365:
                                    incluir = True
                            except:
                                pass
                    
                    if incluir:
                        dados_filtrados.append(linha)
                
                if not dados_filtrados:
                    messagebox.showwarning("Aviso", "Nenhum registro encontrado para o período selecionado.")
                    return
                
                # Criar janela de pré-visualização
                self.preview_e_imprimir(dados_filtrados, periodo)
                
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao preparar impressão:\n{str(e)}")
        
        # Botões
        btn_frame = tk.Frame(janela, bg="#f0f0f0")
        btn_frame.pack(pady=20)
        
        ttk.Button(btn_frame, text="🖨️ Imprimir", command=imprimir_direto).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="💾 Salvar Arquivo", command=gerar_relatorio).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="❌ Cancelar", command=janela.destroy).pack(side=tk.LEFT, padx=5)
    
    def gerar_relatorio_txt(self, arquivo, dados, periodo):
        with open(arquivo, 'w', encoding='utf-8') as f:
            f.write("="*80 + "\n")
            f.write("RELATÓRIO DE MATERIAIS DE TI - SEMED\n")
            f.write("="*80 + "\n\n")
            
            periodo_texto = {
                "todos": "Todos os Registros",
                "semanal": "Últimos 7 Dias",
                "mensal": "Últimos 30 Dias",
                "anual": "Último Ano"
            }
            
            f.write(f"Período: {periodo_texto.get(periodo, 'Todos')}\n")
            f.write(f"Data de Geração: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n")
            f.write(f"Total de Registros: {len(dados)}\n")
            f.write("\n" + "="*80 + "\n\n")
            
            for idx, linha in enumerate(dados, 1):
                f.write(f"REGISTRO #{idx}\n")
                f.write("-"*80 + "\n")
                for col in colunas:
                    valor = linha.get(col, "N/A")
                    f.write(f"{col:30s}: {valor}\n")
                f.write("\n")
            
            f.write("="*80 + "\n")
            f.write(f"FIM DO RELATÓRIO - {len(dados)} REGISTRO(S)\n")
            f.write("="*80 + "\n")
    
    def gerar_relatorio_csv(self, arquivo, dados, periodo):
        import csv
        
        with open(arquivo, 'w', newline='', encoding='utf-8-sig') as f:
            writer = csv.writer(f, delimiter=';')
            
            # Cabeçalho do relatório
            writer.writerow(["RELATÓRIO DE MATERIAIS DE TI - SEMED"])
            writer.writerow([f"Período: {periodo.upper()}"])
            writer.writerow([f"Data: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}"])
            writer.writerow([f"Total: {len(dados)} registros"])
            writer.writerow([])
            
            # Cabeçalhos das colunas
            writer.writerow(colunas)
            
            # Dados
            for linha in dados:
                valores = [linha.get(col, "") for col in colunas]
                writer.writerow(valores)
    
    def preview_e_imprimir(self, dados, periodo):
        janela_preview = tk.Toplevel(self.root)
        janela_preview.title("🖨️ Pré-visualização de Impressão")
        janela_preview.geometry("900x600")
        janela_preview.configure(bg="#ffffff")
        
        # Frame para o conteúdo
        content_frame = tk.Frame(janela_preview, bg="#ffffff")
        content_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        
        # Área de texto com scroll
        texto_scroll = scrolledtext.ScrolledText(content_frame, 
                                                  wrap=tk.WORD, 
                                                  width=100, 
                                                  height=30,
                                                  font=("Courier New", 9),
                                                  bg="#ffffff",
                                                  fg="#000000")
        texto_scroll.pack(fill=tk.BOTH, expand=True)
        
        # Gerar conteúdo do relatório
        periodo_texto = {
            "todos": "Todos os Registros",
            "semanal": "Últimos 7 Dias",
            "mensal": "Últimos 30 Dias",
            "anual": "Último Ano"
        }
        
        conteudo = "="*80 + "\n"
        conteudo += "RELATÓRIO DE MATERIAIS DE TI - SEMED\n"
        conteudo += "="*80 + "\n\n"
        conteudo += f"Período: {periodo_texto.get(periodo, 'Todos')}\n"
        conteudo += f"Data de Geração: {datetime.now().strftime('%d/%m/%Y %H:%M:%S')}\n"
        conteudo += f"Total de Registros: {len(dados)}\n"
        conteudo += "\n" + "="*80 + "\n\n"
        
        for idx, linha in enumerate(dados, 1):
            conteudo += f"REGISTRO #{idx}\n"
            conteudo += "-"*80 + "\n"
            for col in colunas:
                valor = linha.get(col, "N/A")
                conteudo += f"{col:30s}: {valor}\n"
            conteudo += "\n"
        
        conteudo += "="*80 + "\n"
        conteudo += f"FIM DO RELATÓRIO - {len(dados)} REGISTRO(S)\n"
        conteudo += "="*80 + "\n"
        
        texto_scroll.insert(tk.END, conteudo)
        texto_scroll.config(state=tk.DISABLED)
        
        def enviar_impressora():
            try:
                # Criar arquivo temporário
                import tempfile
                with tempfile.NamedTemporaryFile(mode='w', delete=False, suffix='.txt', encoding='utf-8') as temp:
                    temp.write(conteudo)
                    temp_path = temp.name
                
                # Tentar imprimir usando notepad (Windows)
                import subprocess
                if os.name == 'nt':  # Windows
                    # Usar notepad para imprimir
                    subprocess.run(['notepad', '/p', temp_path], check=False)
                    messagebox.showinfo("Impressão", "📄 Documento enviado para impressão!\n\nFeche a janela do Notepad após a impressão.")
                else:
                    # Para Linux/Mac, tentar lpr
                    subprocess.run(['lpr', temp_path], check=False)
                    messagebox.showinfo("Impressão", "✅ Documento enviado para a impressora padrão!")
                
                # Aguardar um pouco antes de deletar o arquivo
                self.root.after(5000, lambda: os.unlink(temp_path) if os.path.exists(temp_path) else None)
                
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao enviar para impressão:\n{str(e)}\n\nTente usar 'Salvar Arquivo' e imprimir manualmente.")
        
        # Botões
        btn_frame = tk.Frame(janela_preview, bg="#ffffff")
        btn_frame.pack(pady=10)
        
        ttk.Button(btn_frame, text="🖨️ Enviar para Impressora", 
                  command=enviar_impressora).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="❌ Fechar", 
                  command=janela_preview.destroy).pack(side=tk.LEFT, padx=5)

def main():
    root = tk.Tk()
    app = ControleMateriaisApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()
