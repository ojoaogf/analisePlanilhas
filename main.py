import pandas as pd
import matplotlib.pyplot as plt
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg
from matplotlib.figure import Figure
import datetime
import os
from reportlab.lib.pagesizes import letter, A4
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, Image
from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet
import io
import numpy as np

# Função para normalizar nomes de técnicos
def normalizar_nome(nome):
    if pd.isna(nome) or nome == "":
        return "Sem Técnico"
    
    # Remover espaços extras e converter para minúsculas para padronização
    nome = nome.strip().lower()
    
    # Capitalizar cada palavra para apresentação
    nome = ' '.join(word.capitalize() for word in nome.split())
    
    return nome

# Função para dividir nomes compostos separados por vírgula
def dividir_nomes_tecnicos(df):
    # Criar uma cópia do dataframe para não modificar o original
    df_expandido = df.copy()
    
    # Verificar se a coluna Técnico existe
    if "Técnico" not in df_expandido.columns:
        return df_expandido
    
    # Normalizar todos os nomes de técnicos
    df_expandido["Técnico"] = df_expandido["Técnico"].apply(normalizar_nome)
    
    # Criar um novo dataframe para armazenar as linhas expandidas
    linhas_expandidas = []
    
    # Processar cada linha do dataframe
    for _, row in df_expandido.iterrows():
        tecnico = row["Técnico"]
        
        # Verificar se o nome contém vírgula (múltiplos técnicos)
        if "," in tecnico:
            # Dividir os nomes e criar uma linha para cada técnico
            for nome in tecnico.split(","):
                nome_normalizado = normalizar_nome(nome)
                nova_linha = row.copy()
                nova_linha["Técnico"] = nome_normalizado
                linhas_expandidas.append(nova_linha)
        else:
            # Manter a linha original
            linhas_expandidas.append(row)
    
    # Criar um novo dataframe com as linhas expandidas
    return pd.DataFrame(linhas_expandidas)

def selecionar_arquivo():
    caminho_arquivo = filedialog.askopenfilename(
        filetypes=[("Planilhas Excel", "*.xlsx")],
        title="Selecione a planilha"
    )

    if not caminho_arquivo:
        return

    try:
        # Usar parse_dates para converter automaticamente colunas de data
        df = pd.read_excel(caminho_arquivo, parse_dates=["Data Início", "Data Vencimento"])
        
        # Verificar se as colunas necessárias existem
        colunas_necessarias = ["ID tarefa", "URL tarefa", "Projeto", "Atividade", 
                              "Data Início", "Data Vencimento", "Técnico"]
        
        colunas_faltantes = [col for col in colunas_necessarias if col not in df.columns]
        
        if colunas_faltantes:
            messagebox.showerror("Erro", f"A planilha não contém as seguintes colunas: {', '.join(colunas_faltantes)}")
            return
        
        # Exibir o dashboard
        exibir_dashboard(df)

    except Exception as e:
        messagebox.showerror("Erro ao processar", str(e))

def exibir_dashboard(df):
    # Criar uma nova janela para o dashboard
    janela_dashboard = tk.Toplevel()
    janela_dashboard.title("Dashboard de Métricas")
    janela_dashboard.geometry("1200x800")
    janela_dashboard.configure(bg=cor_fundo)
    
    # Adicionar cabeçalho
    frame_header = tk.Frame(janela_dashboard, bg=cor_destaque, height=60)
    frame_header.pack(fill="x")
    
    tk.Label(frame_header, text="Dashboard de Análise de Tarefas", 
            font=("Arial", 16, "bold"), bg=cor_destaque, fg="white").pack(pady=15)
    
    # Criar notebook (abas) com estilo
    style = ttk.Style()
    style.configure("TNotebook", background=cor_fundo, borderwidth=0)
    style.configure("TNotebook.Tab", background="#ddd", padding=[15, 5], font=('Arial', 10, 'bold'), foreground="#666666")
    style.map("TNotebook.Tab", 
             background=[("selected", "#e0e0e0")],  # Fundo cinza claro quando selecionado
             foreground=[("selected", "black")],    # Texto preto quando selecionado
             # Garantir que a aba selecionada não pareça desativada
             state=[("selected", "!disabled"), ("active", "!disabled")])
    
    # Remover o contorno de foco das abas
    style.layout("TNotebook.Tab", [
        ('Notebook.tab', {
            'sticky': 'nswe', 
            'children': [
                ('Notebook.padding', {
                    'side': 'top', 
                    'sticky': 'nswe',
                    'children': [
                        ('Notebook.label', {'side': 'top', 'sticky': ''})
                    ]
                })
            ]
        })
    ])
    notebook = ttk.Notebook(janela_dashboard)
    notebook.pack(fill="both", expand=True, padx=15, pady=15)
    
    # Aba 1: Tabela de dados
    tab_dados = ttk.Frame(notebook, style="TFrame")
    notebook.add(tab_dados, text="Dados")
    
    # Aba 2: Gráficos
    tab_graficos = ttk.Frame(notebook)
    notebook.add(tab_graficos, text="Gráficos")
    
    # Aba 3: Métricas
    tab_metricas = ttk.Frame(notebook)
    notebook.add(tab_metricas, text="Métricas")
    
    # Remover a criação da aba de intercorrências
    # tab_intercorrencias = ttk.Frame(notebook)
    # notebook.add(tab_intercorrencias, text="Intercorrências")
    
    # Configurar as abas
    configurar_aba_dados(tab_dados, df)
    configurar_aba_graficos(tab_graficos, df)
    configurar_aba_metricas(tab_metricas, df)
    # Remover a chamada para configurar_aba_intercorrencias
    # configurar_aba_intercorrencias(tab_intercorrencias, df)
    
    # Frame para botões de ação
    frame_acoes = tk.Frame(janela_dashboard, bg=cor_fundo, height=60)
    frame_acoes.pack(fill="x", padx=15, pady=10)
    
    # Botões para exportar
    btn_exportar_pdf = tk.Button(frame_acoes, text="Exportar para PDF", 
                               command=lambda: exportar_pdf(df),
                               font=("Arial", 11), bg=cor_destaque, fg="white",
                               padx=15, pady=8, borderwidth=0)
    btn_exportar_pdf.pack(side="right", padx=10)
    
    btn_voltar = tk.Button(frame_acoes, text="Voltar", 
                         command=janela_dashboard.destroy,
                         font=("Arial", 11), bg="#999", fg="white",
                         padx=15, pady=8, borderwidth=0)
    btn_voltar.pack(side="left", padx=10)
    
def configurar_aba_dados(tab, df):
    # Criar um frame com scrollbar
    frame = tk.Frame(tab, bg=cor_fundo)
    frame.pack(fill="both", expand=True, padx=15, pady=15)
    
    # Adicionar barra de pesquisa
    frame_pesquisa = tk.Frame(frame, bg=cor_fundo)
    frame_pesquisa.pack(fill="x", pady=10)
    
    tk.Label(frame_pesquisa, text="Pesquisar:", font=("Arial", 11), bg=cor_fundo).pack(side="left", padx=5)
    entrada_pesquisa = tk.Entry(frame_pesquisa, width=40, font=("Arial", 11), fg="black")
    entrada_pesquisa.pack(side="left", padx=5)
    
    def pesquisar():
        termo = entrada_pesquisa.get().lower()
        for item in tree.get_children():
            tree.delete(item)
            
        for _, row in df.iterrows():
            valores = []
            encontrado = False
            
            for col in colunas:
                if col in df.columns:
                    valor = row[col]
                    # Formatar datas
                    if isinstance(valor, (datetime.datetime, pd.Timestamp)):
                        valor = valor.strftime("%d/%m/%Y")
                    valor_str = str(valor) if not pd.isna(valor) else ""
                    valores.append(valor_str)
                    
                    if termo in valor_str.lower():
                        encontrado = True
                else:
                    valores.append("")
            
            if encontrado or termo == "":
                tree.insert("", "end", values=valores)
    
    btn_pesquisar = tk.Button(frame_pesquisa, text="Buscar", command=pesquisar,
                            font=("Arial", 10), bg=cor_destaque, fg="white",
                            padx=10, pady=2, borderwidth=0)
    btn_pesquisar.pack(side="left", padx=5)
    
    btn_limpar = tk.Button(frame_pesquisa, text="Limpar", 
                         command=lambda: [entrada_pesquisa.delete(0, tk.END), pesquisar()],
                         font=("Arial", 10), bg="#999", fg="white",
                         padx=10, pady=2, borderwidth=0)
    btn_limpar.pack(side="left", padx=5)
    
    # Frame para a tabela
    frame_tabela = tk.Frame(frame)
    frame_tabela.pack(fill="both", expand=True, pady=10)
    
    # Criar scrollbars
    scrollbar_y = tk.Scrollbar(frame_tabela)
    scrollbar_y.pack(side="right", fill="y")
    
    scrollbar_x = tk.Scrollbar(frame_tabela, orient="horizontal")
    scrollbar_x.pack(side="bottom", fill="x")
    
    # Configurar estilo da tabela
    style = ttk.Style()
    style.configure("Treeview", 
                   background="#f9f9f9",
                   foreground="black",
                   rowheight=25,
                   fieldbackground="#f9f9f9",
                   font=("Arial", 10))
    style.configure("Treeview.Heading", 
                   font=("Arial", 11, "bold"),
                   background="#e0e0e0",
                   foreground="black")
    style.map("Treeview", background=[("selected", "#bfbfbf")])
    
    # Criar Treeview (tabela)
    colunas = ["ID tarefa", "URL tarefa", "Projeto", "Atividade", 
               "Data Início", "Data Vencimento", "Técnico"]
    
    tree = ttk.Treeview(frame_tabela, columns=colunas, show="headings",
                        yscrollcommand=scrollbar_y.set,
                        xscrollcommand=scrollbar_x.set)
    
    # Configurar as scrollbars
    scrollbar_y.config(command=tree.yview)
    scrollbar_x.config(command=tree.xview)
    
    # Configurar cabeçalhos e colunas
    for col in colunas:
        tree.heading(col, text=col)
        tree.column(col, width=120, anchor="center")
    
    # Função para abrir URL quando clicada
    def abrir_url(event):
        item = tree.selection()[0]
        url_tarefa = tree.item(item, "values")[1]  # URL está na segunda coluna (índice 1)
        if url_tarefa and url_tarefa != "":
            import webbrowser
            webbrowser.open(url_tarefa)
    
    # Vincular evento de clique duplo à função de abrir URL
    tree.bind("<Double-1>", abrir_url)
    
    # Inserir dados
    for _, row in df.iterrows():
        valores = []
        for col in colunas:
            if col in df.columns:
                valor = row[col]
                # Formatar datas
                if isinstance(valor, (datetime.datetime, pd.Timestamp)):
                    valor = valor.strftime("%d/%m/%Y")
                # Formatar URL da tarefa para incluir o domínio completo
                if col == "URL tarefa" and not pd.isna(valor) and valor != "":
                    # Verificar se a URL já tem o protocolo (http:// ou https://)
                    if not valor.startswith(("http://", "https://")):
                        # Se não tiver, adicionar https://
                        valor = "https://" + valor
                valores.append(str(valor) if not pd.isna(valor) else "")
            else:
                valores.append("")
        
        tree.insert("", "end", values=valores)
    
    # Configurar estilo para links
    tree.tag_configure("link", foreground="blue")
    
    # Alterar o cursor quando passar sobre a coluna de URL
    def on_motion(event):
        item = tree.identify_row(event.y)
        column = tree.identify_column(event.x)
        if item and column == "#2":  # Coluna URL tarefa (segunda coluna)
            tree.config(cursor="hand2")
        else:
            tree.config(cursor="")
    
    tree.bind("<Motion>", on_motion)
    
    tree.pack(fill="both", expand=True)
    
    # Frame para botões
    frame_botoes = tk.Frame(frame, bg=cor_fundo)
    frame_botoes.pack(fill="x", pady=10)
    
    # Botão para exportar para Excel
    btn_exportar = tk.Button(frame_botoes, text="Exportar para Excel", 
                            command=lambda: exportar_excel(df[colunas]),
                            font=("Arial", 11), bg=cor_destaque, fg="white",
                            padx=15, pady=5, borderwidth=0)
    btn_exportar.pack(side="right", padx=10)
    
    # Contador de registros
    tk.Label(frame_botoes, text=f"Total de registros: {len(df)}", 
            font=("Arial", 11), bg=cor_fundo).pack(side="left", padx=10)

def configurar_aba_graficos(tab, df):
    # Criar frame para os gráficos
    frame = tk.Frame(tab, bg=cor_fundo)
    frame.pack(fill="both", expand=True, padx=20, pady=20)

    # Criar frames separados para cada gráfico
    frame_superior = tk.Frame(frame, bg=cor_fundo)
    frame_superior.pack(fill="x", expand=True, pady=10)
    
    # Recriar o frame inferior
    frame_inferior = tk.Frame(frame, bg=cor_fundo)
    frame_inferior.pack(fill="x", expand=True, pady=10)
    
    # Dividir o frame superior em duas colunas
    frame_sup_esq = tk.Frame(frame_superior, bg=cor_fundo)
    frame_sup_esq.pack(side="left", fill="both", expand=True, padx=10)
    
    frame_sup_dir = tk.Frame(frame_superior, bg=cor_fundo)
    frame_sup_dir.pack(side="right", fill="both", expand=True, padx=10)
    
    # Recriar os frames inferiores
    frame_inf_esq = tk.Frame(frame_inferior, bg=cor_fundo)
    frame_inf_esq.pack(side="left", fill="both", expand=True, padx=10)
    
    # Variáveis globais para os gráficos
    global canvas1, canvas2
    canvas1 = None
    canvas2 = None
    
    # Modificar a função para não usar filtragem
    def atualizar_graficos(dados_filtrados=df):
        global canvas1, canvas2
        
        # Processar os dados para normalizar e dividir nomes de técnicos
        dados_processados = dividir_nomes_tecnicos(dados_filtrados)
        
        # Limpar os frames dos gráficos
        for widget in frame_sup_esq.winfo_children():
            widget.destroy()
        
        for widget in frame_sup_dir.winfo_children():
            widget.destroy()
        
        # Gráfico 1: Tarefas por Projeto (superior esquerdo)
        # Recriar o canvas com scrollbar e usar um frame simples
        frame_grafico_projetos = tk.Frame(frame_sup_esq, bg=cor_fundo)
        frame_grafico_projetos.pack(fill="both", expand=True)
        
        # Comentado: Criar um canvas com scrollbar para o gráfico de projetos
        # canvas_frame = tk.Frame(frame_sup_esq, bg=cor_fundo)
        # canvas_frame.pack(fill="both", expand=True)
        
        # Comentado: Adicionar scrollbar vertical
        # scrollbar = tk.Scrollbar(canvas_frame, orient="vertical")
        # scrollbar.pack(side="right", fill="y")
        
        # Comentado: Criar canvas que será rolável
        # canvas = tk.Canvas(canvas_frame, bg=cor_fundo, yscrollcommand=scrollbar.set)
        # canvas.pack(side="left", fill="both", expand=True)
        
        # Comentado: Configurar scrollbar para controlar o canvas
        # scrollbar.config(command=canvas.yview)
        
        # Comentado: Frame dentro do canvas para conter o gráfico
        # grafico_frame = tk.Frame(canvas, bg=cor_fundo)
        # canvas.create_window((0, 0), window=grafico_frame, anchor="nw")
        
        # Calcular contagem de projetos
        contagem_projetos = dados_processados["Projeto"].value_counts().sort_values(ascending=False)
        
        # Limitar o número de projetos exibidos para manter o gráfico com tamanho similar ao de técnicos
        # Mostrar apenas os 12 projetos mais frequentes (similar ao gráfico de técnicos)
        if len(contagem_projetos) > 12:
            outros_projetos = contagem_projetos[12:].sum()
            contagem_projetos = contagem_projetos[:12]
            # Adicionar uma entrada para "Outros projetos"
            if outros_projetos > 0:
                contagem_projetos["Outros projetos"] = outros_projetos
        
        # Criar figura com tamanho fixo, similar ao gráfico de técnicos
        fig1 = Figure(figsize=(5, 4), dpi=100)
        ax1 = fig1.add_subplot(111)
        
        # Usar barras horizontais para melhor visualização, com a mesma cor do gráfico de técnicos
        contagem_projetos.plot(kind="barh", ax=ax1, color="#4682B4")
        
        # Configurar título e labels
        ax1.set_title("Tarefas por Projeto", fontsize=12, fontweight='bold')
        ax1.set_xlabel("Quantidade", fontsize=10)
        ax1.set_ylabel("")  # Remover label do eixo Y
        ax1.tick_params(axis='y', labelsize=8)
        
        # Adicionar linhas de grade horizontais para facilitar a leitura
        ax1.grid(True, axis='x', linestyle='--', alpha=0.7)
        
        # Adicionar os valores no final das barras
        for i, v in enumerate(contagem_projetos):
            ax1.text(v + 0.1, i, str(v), va='center', fontsize=8)
        
        # Remover bordas desnecessárias (comentado o contorno cinza)
        ax1.spines['top'].set_visible(False)
        ax1.spines['right'].set_visible(False)
        ax1.spines['left'].set_visible(False)  # Remover contorno em cinza
        ax1.spines['bottom'].set_visible(False)  # Remover contorno em cinza
        
        fig1.tight_layout(pad=2.0)
        
        # Remover o canvas anterior se existir
        if canvas1:
            canvas1.get_tk_widget().destroy()
        
        # Criar o widget do gráfico no frame
        canvas1 = FigureCanvasTkAgg(fig1, master=frame_grafico_projetos)
        canvas1.draw()
        canvas1.get_tk_widget().pack(fill="both", expand=True)
        
        # Comentado: Atualizar região de rolagem do canvas
        # grafico_frame.update_idletasks()
        # canvas.config(scrollregion=canvas.bbox("all"))
        
        # Gráfico 2: Tarefas por Técnico (superior direito)
        if "Técnico" in dados_processados.columns:
            fig2 = Figure(figsize=(5, 4), dpi=100)
            ax2 = fig2.add_subplot(111)
            
            # Contar tarefas por técnico com nomes normalizados
            contagem_tecnicos = dados_processados["Técnico"].value_counts().sort_values(ascending=False)
            
            # Usar barras horizontais para melhor visualização, como no exemplo
            contagem_tecnicos.plot(kind="barh", ax=ax2, color="#4682B4")  # Cor azul similar à imagem
            
            # Configurar título e labels
            ax2.set_title("Tarefas por Técnico", fontsize=12, fontweight='bold')
            ax2.set_xlabel("Quantidade", fontsize=10)
            ax2.set_ylabel("")  # Remover label do eixo Y
            ax2.tick_params(axis='y', labelsize=8)
            
            # Adicionar linhas de grade horizontais para facilitar a leitura
            ax2.grid(True, axis='x', linestyle='--', alpha=0.7)
            
            # Adicionar os valores no final das barras
            for i, v in enumerate(contagem_tecnicos):
                ax2.text(v + 0.1, i, str(v), va='center', fontsize=8)
            
            # Remover bordas desnecessárias
            ax2.spines['top'].set_visible(False)
            ax2.spines['right'].set_visible(False)
            
            fig2.tight_layout(pad=2.0)
            
            # Remover o canvas anterior se existir
            if canvas2:
                canvas2.get_tk_widget().destroy()
            
            canvas2 = FigureCanvasTkAgg(fig2, master=frame_sup_dir)
            canvas2.draw()
            canvas2.get_tk_widget().pack(fill="both", expand=True)
    
    # Inicializar os gráficos com todos os dados (sem filtragem)
    atualizar_graficos()

def configurar_aba_metricas(tab, df):
    # Criar frame para as métricas
    frame = tk.Frame(tab, bg=cor_fundo)
    frame.pack(fill="both", expand=True, padx=20, pady=20)
    
    # Calcular métricas
    total_tarefas = len(df)
    total_projetos = df["Projeto"].nunique()
    
    
    # Calcular média de dias por tarefa
    media_dias = "N/A"
    if "Data Início" in df.columns and "Data Vencimento" in df.columns:
        df_temp = df.copy()
        df_temp["Data Início"] = pd.to_datetime(df_temp["Data Início"])
        df_temp["Data Vencimento"] = pd.to_datetime(df_temp["Data Vencimento"])
        df_temp["Duração (dias)"] = (df_temp["Data Vencimento"] - df_temp["Data Início"]).dt.days
        media_dias = round(df_temp["Duração (dias)"].mean(), 1)
    
    # Encontrar o projeto com mais tarefas
    projeto_mais_tarefas = df["Projeto"].value_counts().idxmax()
    qtd_tarefas_projeto = df["Projeto"].value_counts().max()
    
    # Encontrar o técnico com mais tarefas
    if "Técnico" in df.columns:
        # Usar a função dividir_nomes_tecnicos para normalizar e dividir os nomes
        df_processado = dividir_nomes_tecnicos(df)
        
        # Lista de técnicos a serem excluídos da contagem
        tecnicos_excluidos = ["João Gabriel", "Isabella Cristina", "Paula Grippa"]
        
        # Filtrar o dataframe para excluir os técnicos da lista
        df_filtrado = df_processado[~df_processado["Técnico"].isin(tecnicos_excluidos)]
        
        # Verificar se ainda existem técnicos após a filtragem
        if len(df_filtrado) > 0 and df_filtrado["Técnico"].nunique() > 0:
            tecnico_mais_tarefas = df_filtrado["Técnico"].value_counts().idxmax()
            qtd_tarefas_tecnico = df_filtrado["Técnico"].value_counts().max()
        else:
            tecnico_mais_tarefas = "N/A"
            qtd_tarefas_tecnico = 0
    else:
        tecnico_mais_tarefas = "N/A"
        qtd_tarefas_tecnico = 0
    
    # Encontrar o dia com mais tarefas
    if "Data Início" in df.columns:
        df_temp = df.copy()
        df_temp["Data"] = pd.to_datetime(df_temp["Data Início"]).dt.date
        dia_mais_tarefas = df_temp["Data"].value_counts().idxmax()
        qtd_tarefas_dia = df_temp["Data"].value_counts().max()
        dia_formatado = dia_mais_tarefas.strftime("%d/%m/%Y") if hasattr(dia_mais_tarefas, "strftime") else str(dia_mais_tarefas)
    else:
        dia_formatado = "N/A"
        qtd_tarefas_dia = 0
    
    # Título com estilo
    tk.Label(frame, text="Métricas Principais", 
            font=("Arial", 18, "bold"), bg=cor_fundo, fg=cor_texto).pack(pady=15)
    
    # Criar um frame para organizar as métricas em cards
    frame_metricas = tk.Frame(frame, bg=cor_fundo)
    frame_metricas.pack(fill="both", expand=True, pady=20)
    
    # Configurar o grid para 3 colunas
    for i in range(3):
        frame_metricas.columnconfigure(i, weight=1)
    
    # Função para criar um card de métrica
    def criar_card_metrica(parent, titulo, valor, row, col, cor="#4CAF50"):
        # Cores para os cards
        cores = {
            0: "#F39C12",  # Laranja
            1: "#3498DB",  # Azul
            2: "#1ABC9C",  # Verde-água
            3: "#F39C12",  # Laranja
            4: "#3498DB",  # Azul
            5: "#1ABC9C"   # Verde-água
        }
        
        # Usar a cor baseada na posição
        cor_card = cores.get(col, cor)
        
        # Criar card com borda fina
        card = tk.Frame(parent, bg="white", relief=tk.SOLID, borderwidth=1)
        card.grid(row=row, column=col, padx=10, pady=10, sticky="nsew")
        
        # Barra colorida no topo
        barra = tk.Frame(card, bg=cor_card, height=3)
        barra.pack(fill="x")
        
        # Valor grande e destacado (centralizado)
        label_valor = tk.Label(card, text=str(valor), 
                font=("Arial", 24, "bold"), bg="white", fg=cor_card)
        label_valor.pack(pady=(20, 5), fill="x")
        label_valor.config(anchor="center", justify="center")
        
        # Título abaixo do valor (texto menor e cinza)
        label_titulo = tk.Label(card, text=titulo, 
                font=("Arial", 10), bg="white", fg="#666666")
        label_titulo.pack(pady=(0, 20), fill="x")
        label_titulo.config(anchor="center", justify="center")
        
        # Ícone à direita
        icone_frame = tk.Frame(card, bg="white")
        icone_frame.place(relx=1.0, rely=0, anchor="ne", x=-5, y=5)
        
        # Criar um canvas para o ícone
        icone = tk.Canvas(icone_frame, width=16, height=16, bg="white", 
                         highlightthickness=0)
        icone.pack()
        
        # Desenhar um retângulo arredondado como ícone
        icone.create_rectangle(2, 2, 14, 14, fill="white", outline=cor_card, width=1)
    
    # Criar os cards para as métricas importantes
    criar_card_metrica(frame_metricas, "Total de Tarefas", total_tarefas, 0, 0)
    criar_card_metrica(frame_metricas, "Total de Projetos", total_projetos, 0, 1)
    criar_card_metrica(frame_metricas, "Dia com Mais Tarefas", f"{dia_formatado}\n({qtd_tarefas_dia} tarefas)", 0, 2)
    
    # Segunda linha de cards
    criar_card_metrica(frame_metricas, "Projeto com Mais Tarefas", f"{projeto_mais_tarefas}\n({qtd_tarefas_projeto} tarefas)", 1, 0)
    criar_card_metrica(frame_metricas, "Técnico com Mais Tarefas", f"{tecnico_mais_tarefas}\n({qtd_tarefas_tecnico} tarefas)", 1, 1)
    criar_card_metrica(frame_metricas, "Erros", "0", 1, 2, "#E74C3C")  # Vermelho para erros

def exportar_excel(df, nome_arquivo="dados_exportados"):
    # Solicitar ao usuário onde salvar o arquivo
    caminho_arquivo = filedialog.asksaveasfilename(
        defaultextension=".xlsx",
        filetypes=[("Planilhas Excel", "*.xlsx")],
        initialfile=f"{nome_arquivo}.xlsx"
    )
    
    if not caminho_arquivo:
        return
    
    try:
        # Exportar para Excel
        df.to_excel(caminho_arquivo, index=False)
        messagebox.showinfo("Sucesso", f"Dados exportados com sucesso para {caminho_arquivo}")
    except Exception as e:
        messagebox.showerror("Erro ao exportar", str(e))

def exportar_pdf(df):
    caminho_salvar = filedialog.asksaveasfilename(
        defaultextension=".pdf",
        filetypes=[("Arquivos PDF", "*.pdf")],
        title="Salvar PDF como"
    )
    
    if not caminho_salvar:
        return
    
    try:
        # Criar o documento PDF
        doc = SimpleDocTemplate(caminho_salvar, pagesize=A4)
        elementos = []
        
        # Estilos
        estilos = getSampleStyleSheet()
        estilo_titulo = estilos["Heading1"]
        estilo_subtitulo = estilos["Heading2"]
        estilo_normal = estilos["Normal"]
        
        # Título
        elementos.append(Paragraph("Dashboard de Métricas", estilo_titulo))
        elementos.append(Spacer(1, 20))
        
        # Seção 1: Métricas Principais
        elementos.append(Paragraph("Métricas Principais", estilo_subtitulo))
        elementos.append(Spacer(1, 10))
        
        # Calcular métricas
        total_tarefas = len(df)
        total_projetos = df["Projeto"].nunique()

        
        # Calcular média de dias por tarefa
        media_dias = "N/A"
        if "Data Início" in df.columns and "Data Vencimento" in df.columns:
            df_temp = df.copy()
            df_temp["Data Início"] = pd.to_datetime(df_temp["Data Início"])
            df_temp["Data Vencimento"] = pd.to_datetime(df_temp["Data Vencimento"])
            df_temp["Duração (dias)"] = (df_temp["Data Vencimento"] - df_temp["Data Início"]).dt.days
            media_dias = round(df_temp["Duração (dias)"].mean(), 1)
        
        # Tabela de métricas
        dados_metricas = [
            ["Métrica", "Valor"],
            ["Total de Tarefas", str(total_tarefas)],
            ["Total de Projetos", str(total_projetos)],
            ["Média de Dias por Tarefa", str(media_dias)]
        ]
        
        if "Projeto" in df.columns:
            projeto_mais_tarefas = df["Projeto"].value_counts().idxmax()
            qtd_tarefas_projeto = df["Projeto"].value_counts().max()
            dados_metricas.append(["Projeto com Mais Tarefas", f"{projeto_mais_tarefas} ({qtd_tarefas_projeto} tarefas)"])
        
        if "Técnico" in df.columns:
            tecnico_mais_tarefas = df["Técnico"].value_counts().idxmax()
            qtd_tarefas_tecnico = df["Técnico"].value_counts().max()
            dados_metricas.append(["Técnico com Mais Tarefas", f"{tecnico_mais_tarefas} ({qtd_tarefas_tecnico} tarefas)"])
        
        tabela_metricas = Table(dados_metricas, colWidths=[300, 200])
        tabela_metricas.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (1, 0), 'CENTER'),
            ('FONTNAME', (0, 0), (1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (1, 0), 12),
            ('BOTTOMPADDING', (0, 0), (1, 0), 12),
            ('BACKGROUND', (0, 1), (1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        
        elementos.append(tabela_metricas)
        elementos.append(Spacer(1, 20))
        
        # Seção 2: Gráficos
        elementos.append(Paragraph("Gráficos", estilo_subtitulo))
        elementos.append(Spacer(1, 10))
        
        # Gráfico 1: Tarefas por Projeto
        fig1 = Figure(figsize=(8, 4))
        ax1 = fig1.add_subplot(111)
        contagem_projetos = df["Projeto"].value_counts()
        contagem_projetos.plot(kind="bar", ax=ax1)
        ax1.set_title("Tarefas por Projeto")
        ax1.set_ylabel("Quantidade")
        ax1.tick_params(axis='x', rotation=45)
        fig1.tight_layout()
        
        # Salvar o gráfico como imagem
        buf1 = io.BytesIO()
        fig1.savefig(buf1, format='png')
        buf1.seek(0)
        
        # Adicionar o gráfico ao PDF
        img1 = Image(buf1, width=450, height=250)
        elementos.append(img1)
        elementos.append(Spacer(1, 20))
        
        # Gráfico 2: Tarefas por Técnico
        if "Técnico" in df.columns:
            fig2 = Figure(figsize=(8, 4))
            ax2 = fig2.add_subplot(111)
            contagem_tecnicos = df["Técnico"].value_counts()
            contagem_tecnicos.plot(kind="bar", ax=ax2)
            ax2.set_title("Tarefas por Técnico")
            ax2.set_ylabel("Quantidade")
            ax2.tick_params(axis='x', rotation=45)
            fig2.tight_layout()
            
            # Salvar o gráfico como imagem
            buf2 = io.BytesIO()
            fig2.savefig(buf2, format='png')
            buf2.seek(0)
            
            # Adicionar o gráfico ao PDF
            img2 = Image(buf2, width=450, height=250)
            elementos.append(img2)
            elementos.append(Spacer(1, 20))
        
        # Seção 3: Tabela de Dados
        elementos.append(Paragraph("Dados das Tarefas", estilo_subtitulo))
        elementos.append(Spacer(1, 10))
        
        # Preparar dados para a tabela
        colunas = ["ID tarefa", "Projeto", "Atividade", "Data Início", "Data Vencimento", "Técnico"]
        dados_tabela = [colunas]  # Cabeçalho
        
        # Limitar a 20 linhas para não sobrecarregar o PDF
        for _, row in df.head(20).iterrows():
            linha = []
            for col in colunas:
                if col in df.columns:
                    valor = row[col]
                    # Formatar datas
                    if isinstance(valor, (datetime.datetime, pd.Timestamp)):
                        valor = valor.strftime("%d/%m/%Y")
                    linha.append(str(valor) if not pd.isna(valor) else "")
                else:
                    linha.append("")
            dados_tabela.append(linha)
        
        # Criar a tabela
        tabela_dados = Table(dados_tabela, colWidths=[60, 100, 150, 80, 80, 50])
        tabela_dados.setStyle(TableStyle([
            ('BACKGROUND', (0, 0), (-1, 0), colors.grey),
            ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
            ('ALIGN', (0, 0), (-1, 0), 'CENTER'),
            ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
            ('FONTSIZE', (0, 0), (-1, 0), 10),
            ('BOTTOMPADDING', (0, 0), (-1, 0), 12),
            ('BACKGROUND', (0, 1), (-1, -1), colors.beige),
            ('GRID', (0, 0), (-1, -1), 1, colors.black)
        ]))
        
        elementos.append(tabela_dados)
        
        # Adicionar nota de rodapé
        elementos.append(Spacer(1, 30))
        elementos.append(Paragraph(f"Relatório gerado em {datetime.datetime.now().strftime('%d/%m/%Y %H:%M:%S')}", estilo_normal))
        
        # Construir o PDF
        doc.build(elementos)
        
        messagebox.showinfo("Sucesso", f"PDF exportado com sucesso para:\n{caminho_salvar}")
        
    except Exception as e:
        messagebox.showerror("Erro ao exportar PDF", str(e))

# Interface principal
janela = tk.Tk()
janela.title("Analisador de Planilhas")
janela.geometry("500x400")  # Aumentar o tamanho da janela

# Configurar cores e estilos
cor_fundo = "#f0f0f0"
cor_destaque = "#4CAF50"
cor_texto = "#333333"
cor_texto_claro = "white"

janela.configure(bg=cor_fundo)

# Estilizar a interface principal
frame_principal = tk.Frame(janela, padx=30, pady=30, bg=cor_fundo)
frame_principal.pack(fill="both", expand=True)

# Logo ou ícone (pode ser substituído por uma imagem real)
frame_logo = tk.Frame(frame_principal, bg=cor_fundo, height=80)
frame_logo.pack(fill="x", pady=10)
tk.Label(frame_logo, text="📊", font=("Arial", 40), bg=cor_fundo, fg=cor_destaque).pack()

# Título com estilo melhorado
titulo = tk.Label(frame_principal, text="Analisador de Planilhas", 
                 font=("Arial", 22, "bold"), bg=cor_fundo, fg=cor_texto)
titulo.pack(pady=20)

# Descrição com estilo melhorado
descricao = tk.Label(frame_principal, 
                    text="Selecione uma planilha Excel para analisar e gerar um dashboard interativo com métricas e gráficos.", 
                    font=("Arial", 11), wraplength=400, bg=cor_fundo, fg=cor_texto)
descricao.pack(pady=20)

# Frame para botões
frame_botoes = tk.Frame(frame_principal, bg=cor_fundo)
frame_botoes.pack(pady=20)

# Botão estilizado com hover effect
estilo_botao = {"font": ("Arial", 12, "bold"), "bg": cor_destaque, "fg": cor_texto_claro, 
               "activebackground": "#45a049", "relief": tk.RAISED, "padx": 25, "pady": 12,
               "borderwidth": 0, "cursor": "hand2"}

botao = tk.Button(frame_botoes, text="Selecionar Planilha", command=selecionar_arquivo, **estilo_botao)
botao.pack(pady=10)

# Adicionar rodapé
rodape = tk.Label(frame_principal, text="© 2023 Analisador de Planilhas", 
                 font=("Arial", 8), bg=cor_fundo, fg="#999999")
rodape.pack(side="bottom", pady=10)

# Centralizar a janela na tela
largura_janela = 500
altura_janela = 400
largura_tela = janela.winfo_screenwidth()
altura_tela = janela.winfo_screenheight()
x = (largura_tela - largura_janela) // 2
y = (altura_tela - altura_janela) // 2
janela.geometry(f"{largura_janela}x{altura_janela}+{x}+{y}")

janela.mainloop()


def configurar_aba_intercorrencias(tab, df):
    # Criar um frame com scrollbar
    frame = tk.Frame(tab, bg=cor_fundo)
    frame.pack(fill="both", expand=True, padx=15, pady=15)
    
    # Adicionar título
    tk.Label(frame, text="Intercorrências e Erros", 
            font=("Arial", 14, "bold"), bg=cor_fundo).pack(pady=10)
    
    # Filtrar dados para mostrar apenas intercorrências/erros
    # Assumindo que as intercorrências são identificadas pela coluna "Atividade" contendo palavras-chave
    palavras_chave_erro = ["erro", "falha", "problema", "bug", "defeito", "intercorrência", "incidente"]
    
    # Função para verificar se uma atividade é uma intercorrência
    def eh_intercorrencia(atividade):
        if pd.isna(atividade):
            return False
        atividade = str(atividade).lower()
        return any(palavra in atividade for palavra in palavras_chave_erro)
    
    # Filtrar o DataFrame
    df_intercorrencias = df[df["Atividade"].apply(eh_intercorrencia)]
    
    # Frame para a tabela
    frame_tabela = tk.Frame(frame)
    frame_tabela.pack(fill="both", expand=True, pady=10)
    
    # Criar scrollbars
    scrollbar_y = tk.Scrollbar(frame_tabela)
    scrollbar_y.pack(side="right", fill="y")
    
    scrollbar_x = tk.Scrollbar(frame_tabela, orient="horizontal")
    scrollbar_x.pack(side="bottom", fill="x")
    
    # Configurar estilo da tabela
    style = ttk.Style()
    style.configure("Treeview", 
                   background="#f9f9f9",
                   foreground="black",
                   rowheight=25,
                   fieldbackground="#f9f9f9",
                   font=("Arial", 10))
    style.configure("Treeview.Heading", 
                   font=("Arial", 11, "bold"),
                   background="#e0e0e0",
                   foreground="black")
    style.map("Treeview", background=[("selected", "#bfbfbf")])
    
    # Criar Treeview (tabela)
    colunas = ["ID tarefa", "URL tarefa", "Projeto", "Atividade", 
               "Data Início", "Data Vencimento", "Técnico"]
    
    tree = ttk.Treeview(frame_tabela, columns=colunas, show="headings",
                        yscrollcommand=scrollbar_y.set,
                        xscrollcommand=scrollbar_x.set)
    
    # Configurar as scrollbars
    scrollbar_y.config(command=tree.yview)
    scrollbar_x.config(command=tree.xview)
    
    # Configurar cabeçalhos e colunas
    for col in colunas:
        tree.heading(col, text=col)
        tree.column(col, width=120, anchor="center")
    
    # Função para abrir URL quando clicada
    def abrir_url(event):
        item = tree.selection()[0]
        url_tarefa = tree.item(item, "values")[1]  # URL está na segunda coluna (índice 1)
        if url_tarefa and url_tarefa != "":
            import webbrowser
            webbrowser.open(url_tarefa)
    
    # Vincular evento de clique duplo à função de abrir URL
    tree.bind("<Double-1>", abrir_url)
    
    # Inserir dados
    for _, row in df_intercorrencias.iterrows():
        valores = []
        for col in colunas:
            if col in df.columns:
                valor = row[col]
                # Formatar datas
                if isinstance(valor, (datetime.datetime, pd.Timestamp)):
                    valor = valor.strftime("%d/%m/%Y")
                # Formatar URL da tarefa para incluir o domínio completo
                if col == "URL tarefa" and not pd.isna(valor) and valor != "":
                    # Verificar se a URL já tem o protocolo (http:// ou https://)
                    if not valor.startswith(("http://", "https://")):
                        # Se não tiver, adicionar https://
                        valor = "https://" + valor
                valores.append(str(valor) if not pd.isna(valor) else "")
            else:
                valores.append("")
        
        tree.insert("", "end", values=valores)
    
    # Configurar estilo para links
    tree.tag_configure("link", foreground="blue")
    
    # Alterar o cursor quando passar sobre a coluna de URL
    def on_motion(event):
        item = tree.identify_row(event.y)
        column = tree.identify_column(event.x)
        if item and column == "#2":  # Coluna URL tarefa (segunda coluna)
            tree.config(cursor="hand2")
        else:
            tree.config(cursor="")
    
    tree.bind("<Motion>", on_motion)
    
    tree.pack(fill="both", expand=True)
    
    # Frame para botões e informações
    frame_botoes = tk.Frame(frame, bg=cor_fundo)
    frame_botoes.pack(fill="x", pady=10)
    
    # Contador de registros
    tk.Label(frame_botoes, text=f"Total de intercorrências: {len(df_intercorrencias)}", 
            font=("Arial", 11), bg=cor_fundo).pack(side="left", padx=10)
    
    # Botão para exportar para Excel
    btn_exportar = tk.Button(frame_botoes, text="Exportar Intercorrências", 
                            command=lambda: exportar_excel(df_intercorrencias[colunas], "intercorrencias"),
                            font=("Arial", 11), bg=cor_destaque, fg="white",
                            padx=15, pady=5, borderwidth=0)
    btn_exportar.pack(side="right", padx=10)