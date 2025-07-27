import tkinter as tk
from tkinter import ttk, filedialog, messagebox
from tkinter.font import Font
import pandas as pd
import locale
import json


# Configurar locale para moeda brasileira
locale.setlocale(locale.LC_ALL, 'pt_BR.UTF-8')

def carregar_planilha():
    """Carregar a planilha de Excel."""
    global df
    file_path = filedialog.askopenfilename(
        title="Selecione a Planilha",
        filetypes=[("Arquivos Excel", "*.xlsx *.xls")]
    )
    if file_path:
        try:
            df = pd.read_excel(file_path)

            # Converter a coluna COD para texto
            if "COD" in df.columns:
                df["COD"] = df["COD"].astype(str)

            required_columns = ["COD", "PRODUTO", "VL. UNT.", "ENDEREÇO", "QTD"]
            for col in required_columns:
                if col not in df.columns:
                    raise ValueError(f"Coluna obrigatória '{col}' não encontrada na planilha.")

            if "CONTAGEM" not in df.columns:
                df["CONTAGEM"] = 0

            df["VL. ESTOQUE"] = 0.0
            df["DIF. ETQ"] = 0.0
            df["VL. DIF."] = 0.0

            atualizar_tabela()
            redimensionar_colunas()  # Ajusta a largura das colunas após carregar os dados
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar a planilha: {e}")

def formatar_moeda(valor):
    """Formata valores como moeda no padrão brasileiro, incluindo R$."""
    if isinstance(valor, (int, float)):
        if valor >= 1000 or valor <= -1000:  # Adiciona suporte para valores negativos
            return f"R$ {valor:,.2f}".replace(",", "X").replace(".", ",").replace("X", ".")
        else:
            return f"R$ {valor:,.2f}".replace(".", ",")
    return valor
def salvar_json():
    """Salva os dados da planilha no formato JSON."""
    if df is None:
        messagebox.showwarning("Aviso", "Nenhuma planilha foi carregada!")
        return

    save_path = filedialog.asksaveasfilename(
        title="Salvar como JSON",
        defaultextension=".json",
        filetypes=[("Arquivos JSON", "*.json")]
    )
    if save_path:
        try:
            # Converte o DataFrame para JSON
            df.to_json(save_path, orient="records", force_ascii=False, indent=4)
            messagebox.showinfo("Sucesso", "Dados salvos no formato JSON com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar o arquivo JSON: {e}")

def carregar_json():
    """Carrega os dados do inventário a partir de um arquivo JSON."""
    global df
    file_path = filedialog.askopenfilename(
        title="Selecione o Arquivo JSON",
        filetypes=[("Arquivos JSON", "*.json")]
    )
    if file_path:
        try:
            # Carregar o JSON e converter para DataFrame
            df = pd.read_json(file_path, orient="records")

            # Verificar se as colunas obrigatórias estão presentes
            required_columns = ["COD", "PRODUTO", "VL. UNT.", "ENDEREÇO", "QTD"]
            for col in required_columns:
                if col not in df.columns:
                    raise ValueError(f"Coluna obrigatória '{col}' não encontrada no arquivo JSON.")

            if "CONTAGEM" not in df.columns:
                df["CONTAGEM"] = 0

            df["VL. ESTOQUE"] = 0.0
            df["DIF. ETQ"] = 0.0
            df["VL. DIF."] = 0.0

            atualizar_tabela()
            redimensionar_colunas()
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao carregar o arquivo JSON: {e}")

def atualizar_tabela(filtro_codigo=None, filtro_endereco=None):
    """Atualiza a tabela exibida na aba Apuração com filtros opcionais."""
    for row in tabela.get_children():
        tabela.delete(row)

    df_filtrado = df
    if filtro_codigo:
        df_filtrado = df_filtrado[df_filtrado["COD"].astype(str).str.contains(filtro_codigo, na=False)]
    if filtro_endereco:
        df_filtrado = df_filtrado[df_filtrado["ENDEREÇO"].astype(str).str.contains(filtro_endereco, na=False)]

    for index, row in df_filtrado.iterrows():
        valores_formatados = [
            formatar_moeda(row["VL. UNT."]) if col == "VL. UNT." else
            formatar_moeda(row["VL. ESTOQUE"]) if col == "VL. ESTOQUE" else
            formatar_moeda(row["VL. DIF."]) if col == "VL. DIF." else
            row[col] for col in colunas[1:]
        ]
        tabela.insert("", "end", values=(index, *valores_formatados))

    redimensionar_colunas()
    formatar_coluna_vl_dif()  # Aplica a formatação condicional

def formatar_coluna_vl_dif():
    """Aplica formatação condicional à coluna VL. DIF."""
    vl_dif_index = colunas.index("VL. DIF.")  # Localiza o índice da coluna VL. DIF.

    # Removendo todas as tags para evitar conflitos anteriores
    for item in tabela.get_children():
        tabela.item(item, tags="")

    for item in tabela.get_children():
        valores = tabela.item(item, "values")
        try:
            vl_dif = float(
                str(valores[vl_dif_index])
                .replace("R$", "")
                .replace(".", "")
                .replace(",", ".")
            )  # Converte o valor

            if vl_dif > 0:  # Valor positivo
                tabela.tag_configure("positivo", background="green", foreground="white")
                tabela.item(item, tags=("positivo",))
            if vl_dif == 0:  # Valor positivo
                tabela.tag_configure("igual", background="white", foreground="black")
                tabela.item(item, tags=("igual",))
            elif vl_dif < 0:  # Valor negativo
                tabela.tag_configure("negativo", background="red", foreground="white")
                tabela.item(item, tags=("negativo",))
        except ValueError:
            # Ignora erros de conversão
            continue

def classificar_vl_dif(ascendente=True):
    """Classifica a tabela com base na coluna VL. DIF."""
    global df
    if df is None:
        messagebox.showwarning("Aviso", "Nenhum inventário foi carregado!")
        return

    try:
        # Ordena o DataFrame com base em VL. DIF.
        df = df.sort_values(by="VL. DIF.", ascending=ascendente).reset_index(drop=True)
        atualizar_tabela()
        messagebox.showinfo("Sucesso", f"VL. DIF. classificado em ordem {'crescente' if ascendente else 'decrescente'}!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao classificar VL. DIF.: {e}")


def redimensionar_colunas():
    """Ajusta automaticamente a largura de cada coluna com base no conteúdo."""
    fonte_padrao = Font()
    for col in colunas:
        tabela.column(col, width=fonte_padrao.measure(col))
        for item in tabela.get_children():
            valor = tabela.item(item, "values")[colunas.index(col)]
            largura = fonte_padrao.measure(str(valor))
            if tabela.column(col, "width") < largura:
                tabela.column(col, width=largura + 10)



def editar_valor(event):
    """Permite editar o valor da coluna CONTAGEM ao clicar em uma célula."""
    global entry_temporaria
    # Obtem o item selecionado
    item = tabela.identify_row(event.y)
    coluna = tabela.identify_column(event.x)

    # Verifica se clicou na coluna "CONTAGEM"
    if coluna != f"#{colunas.index('CONTAGEM') + 1}":  # Apenas permitir edição na coluna CONTAGEM
        return

    valores = tabela.item(item, "values")
    if not valores:
        return
    indice = int(valores[0])  # Obtém o índice da linha

    # Posiciona o campo de entrada (Entry) na célula
    x, y, width, height = tabela.bbox(item, coluna)
    entry_temporaria = tk.Entry(frame_dados)
    entry_temporaria.place(x=x, y=y, width=width, height=height)
    entry_temporaria.insert(0, valores[colunas.index("CONTAGEM")])
    entry_temporaria.focus()

    # Salvar automaticamente ao perder o foco ou pressionar Enter
    entry_temporaria.bind("<Return>", lambda e: salvar_valor(indice))
    entry_temporaria.bind("<FocusOut>", lambda e: salvar_valor(indice))


def salvar_valor(indice):
    """Salva o valor editado na coluna CONTAGEM."""
    global entry_temporaria
    try:
        novo_valor = float(entry_temporaria.get().replace(",", "."))
        df.at[indice, "CONTAGEM"] = novo_valor
        df.at[indice, "DIF. ETQ"] = novo_valor - df.at[indice, "QTD"]
        df.at[indice, "VL. ESTOQUE"] = df.at[indice, "VL. UNT."] * df.at[indice, "QTD"]
        df.at[indice, "VL. DIF."] = df.at[indice, "DIF. ETQ"] * df.at[indice, "VL. UNT."]
        atualizar_tabela()
        atualizar_resumo()
    except ValueError:
        messagebox.showerror("Erro", "Por favor, insira um valor válido.")
    finally:
        if entry_temporaria:
            entry_temporaria.destroy()
            entry_temporaria = None

def atualizar_resumo():
    """Atualiza o resumo exibido na aba Resumo."""
    for widget in frame_resumo.winfo_children():
        widget.destroy()

    # Calcular resumo
    total_estoque = df["VL. ESTOQUE"].sum()
    total_dif_neg = df[df["DIF. ETQ"] < 0]["VL. DIF."].sum()
    total_dif_pos = df[df["DIF. ETQ"] > 0]["VL. DIF."].sum()
    divergencia_absoluta = abs(total_dif_neg) + abs(total_dif_pos)



    # Calcular os novos totais
    total_itens_contados = len(df[df["CONTAGEM"] > 0])
    total_itens_negativos = len(df[df["DIF. ETQ"] < 0])
    total_itens_positivos = len(df[df["DIF. ETQ"] > 0])

    # Dados de resumo
    resumo_data = [
        ("ESTOQUE TOTAL", formatar_moeda(total_estoque)),
        ("TOTAL DE ITENS CONTADOS", f"{total_itens_contados}"),
        ("TOTAL DE ITENS NEGATIVOS", f"{total_itens_negativos}"),
        ("TOTAL DE ITENS POSITIVOS", f"{total_itens_positivos}"),
        ("TOTAL DIVERGÊNCIAS NEGATIVAS", formatar_moeda(total_dif_neg)),
        ("TOTAL DIVERGÊNCIAS POSITIVAS", formatar_moeda(total_dif_pos)),
        ("% DIVERGÊNCIA ABSOLUTA", f"{(divergencia_absoluta / total_estoque * 100):,.2f}%" if total_estoque != 0 else "0,00%"),
    ]

    # Criar tabela de resumo
    colunas_resumo = ["Descrição", "Valor"]
    tabela_resumo = ttk.Treeview(frame_resumo, columns=colunas_resumo, show="headings", height=len(resumo_data))
    tabela_resumo.pack(fill="both", expand=True, padx=10, pady=10)

    # Configurar colunas
    tabela_resumo.heading("Descrição", text="Descrição", anchor="w")
    tabela_resumo.heading("Valor", text="Valor", anchor="center")
    tabela_resumo.column("Descrição", anchor="w", width=200)  # Reduza a largura da coluna "Descrição"
    tabela_resumo.column("Valor", anchor="center", width=150)  # Mantenha ou aumente a largura da coluna "Valor"

    # Inserir dados na tabela
    for titulo, valor in resumo_data:
        tabela_resumo.insert("", "end", values=(titulo, valor))

    # Estilizar linhas da tabela
    style = ttk.Style()
    style.configure("Treeview", font=("Arial", 12), rowheight=30)
    style.configure("Treeview.Heading", font=("Arial", 14, "bold"))

def apurar_inventario():
    """Apura o inventário realizando os cálculos para cada item."""
    if df is None:
        messagebox.showwarning("Aviso", "Nenhum inventário foi carregado!")
        return

    try:
        # Recalcular valores para cada item
        df["DIF. ETQ"] = df["CONTAGEM"] - df["QTD"]
        df["VL. ESTOQUE"] = df["QTD"] * df["VL. UNT."]
        df["VL. DIF."] = df["DIF. ETQ"] * df["VL. UNT."]

        # Atualizar a tabela e o resumo
        atualizar_tabela()
        atualizar_resumo()
        messagebox.showinfo("Sucesso", "Inventário apurado com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao apurar o inventário: {e}")


def selecionar_contagens():
    """Seleciona todas as linhas da coluna CONTAGEM para edição."""
    if df is None:
        messagebox.showwarning("Aviso", "Nenhuma planilha foi carregada!")
        return

    try:
        # Preenche todas as linhas com um valor padrão (exemplo: 0 ou outro número)
        valor_padrao = simpledialog.askfloat(
            "Valor Padrão",
            "Insira o valor para preencher todas as contagens:",
            minvalue=0
        )
        if valor_padrao is None:
            return  # Caso o usuário cancele

        df["CONTAGEM"] = valor_padrao
        df["DIF. ETQ"] = df["CONTAGEM"] - df["QTD"]
        df["VL. ESTOQUE"] = df["VL. UNT."] * df["QTD"]
        df["VL. DIF."] = df["DIF. ETQ"] * df["VL. UNT."]

        atualizar_tabela()
        atualizar_resumo()
        messagebox.showinfo("Sucesso", "Contagens preenchidas com sucesso!")
    except Exception as e:
        messagebox.showerror("Erro", f"Erro ao preencher contagens: {e}")

def redimensionar_colunas():
    """Ajusta automaticamente a largura de cada coluna com base no conteúdo."""
    fonte_padrao = Font(family="TkDefaultFont")
    for col in colunas:
        largura_maxima = fonte_padrao.measure(col)  # Largura mínima com base no nome da coluna

        for item in tabela.get_children():
            valor = tabela.item(item, "values")[colunas.index(col)]
            largura_valor = fonte_padrao.measure(str(valor))
            largura_maxima = max(largura_maxima, largura_valor)

        tabela.column(col, width=largura_maxima + 10)  # Adiciona um pequeno espaçamento extra

def salvar_planilha():
    """Salva a planilha com os cálculos realizados."""
    if df is None:
        messagebox.showwarning("Aviso", "Nenhuma planilha foi carregada!")
        return

    save_path = filedialog.asksaveasfilename(
        title="Salvar Planilha",
        defaultextension=".xlsx",
        filetypes=[("Arquivos Excel", "*.xlsx")]
    )
    if save_path:
        try:
            df.to_excel(save_path, index=False)
            messagebox.showinfo("Sucesso", "Planilha salva com sucesso!")
        except Exception as e:
            messagebox.showerror("Erro", f"Erro ao salvar a planilha: {e}")


def buscar_por_codigo_endereco():
    """Filtra a tabela por Código de Produto ou Endereço."""
    codigo = entry_busca_codigo.get().strip()
    endereco = entry_busca_endereco.get().strip()
    atualizar_tabela(filtro_codigo=codigo, filtro_endereco=endereco)


# Configuração da interface Tkinter
root = tk.Tk()
root.title("GESTÃO DE ESTOQUES")
root.geometry("1300x1000")

# Título na parte superior
titulo_principal = tk.Label(root, text="GESTÃO DE ESTOQUE - CONTAGEM", font=("Arial", 20, "bold"))
titulo_principal.pack(side="top", pady=10)

df = None  # DataFrame para armazenar os dados carregados

# Menu de Navegação
menu_principal = tk.Menu(root)
root.config(menu=menu_principal)

menu_navegacao = tk.Menu(menu_principal, tearoff=0)
menu_principal.add_cascade(label="NAVEGAÇÃO", menu=menu_navegacao)
menu_navegacao.add_command(label="APURAÇÃO", command=lambda: frame_resumo.pack_forget() or frame_dados.pack(fill="both", expand=True))
menu_navegacao.add_command(label="RESUMO", command=lambda: frame_dados.pack_forget() or frame_resumo.pack(fill="both", expand=True))

# Botões superiores
frame_botoes = tk.Frame(root)
frame_botoes.pack(pady=10)

btn_carregar = tk.Button(frame_botoes, text="CARREGAR PLANILHA", command=carregar_planilha)
btn_carregar.grid(row=0, column=0, padx=10)

btn_salvar = tk.Button(frame_botoes, text="SALVAR PLANILHA", command=salvar_planilha)
btn_salvar.grid(row=1, column=0, padx=10)

# Funções para os botões de filtro
def filtrar_faltas():
    """Filtra itens com valores negativos na coluna VL.DIF."""
    df_filtrado = df[df["VL. DIF."] < 0]
    atualizar_tabela_com_filtro(df_filtrado)

def filtrar_sobras():
    """Filtra itens com valores positivos na coluna VL.DIF."""
    df_filtrado = df[df["VL. DIF."] > 0]
    atualizar_tabela_com_filtro(df_filtrado)

def mostrar_todos():
    """Mostra todos os itens, removendo qualquer filtro."""
    atualizar_tabela()

def atualizar_tabela_com_filtro(df_filtrado):
    """Atualiza a tabela com base em um DataFrame filtrado."""
    for row in tabela.get_children():
        tabela.delete(row)

    for index, row in df_filtrado.iterrows():
        valores_formatados = [
            formatar_moeda(row["VL. UNT."]) if col == "VL. UNT." else
            formatar_moeda(row["VL. ESTOQUE"]) if col == "VL. ESTOQUE" else
            formatar_moeda(row["VL. DIF."]) if col == "VL. DIF." else
            row[col] for col in colunas[1:]
        ]
        tabela.insert("", "end", values=(index, *valores_formatados))

# Adicionando os botões ao canto superior direito
frame_filtros = tk.Frame(root)
frame_filtros.place(relx=0.92, rely=0.01)  # Posicionado no canto superior direito

btn_faltas = tk.Button(
    frame_filtros,
    text="FALTAS",
    font=("Arial", 10, "bold"),
    bg="red",
    fg="white",
    command=filtrar_faltas
)
btn_faltas.pack(side="top", padx=10, pady=2)

btn_sobras = tk.Button(
    frame_filtros,
    text="SOBRAS",
    font=("Arial", 10, "bold"),
    bg="green",
    fg="white",
    command=filtrar_sobras
)
btn_sobras.pack(side="top", padx=5, pady=2)

btn_todos = tk.Button(
    frame_filtros,
    text="TODOS",
    font=("Arial", 10, "bold"),
    bg="darkgray",
    fg="black",
    command=mostrar_todos
)
btn_todos.pack(side="top", padx=5, pady=2)


tk.Label(frame_botoes, text="BUSCAR POR CÓDIGO:").grid(row=1, column=3, padx=5)
entry_busca_codigo = tk.Entry(frame_botoes)
entry_busca_codigo.grid(row=0, column=3, padx=5)

tk.Label(frame_botoes, text="BUSCAR POR ENDEREÇO:").grid(row=1, column=5, padx=5)
entry_busca_endereco = tk.Entry(frame_botoes)
entry_busca_endereco.grid(row=0, column=5, padx=5)

btn_buscar = tk.Button(frame_botoes, text="BUSCAR", command=buscar_por_codigo_endereco)
btn_buscar.grid(row=0, column=6, padx=10)

# Tabela - Aba Apuração
frame_dados = tk.Frame(root)
colunas = ["ÍNDICE", "COD", "PRODUTO", "VL. UNT.", "ENDEREÇO", "QTD", "CONTAGEM", "VL. ESTOQUE", "DIF. ETQ", "VL. DIF."]
tabela = ttk.Treeview(frame_dados, columns=colunas, show="headings", height=15)
tabela.pack(fill="both", expand=True)

style = ttk.Style()
style.configure("Treeview", rowheight=25)

for col in colunas:
    if col == "PRODUTO":
        tabela.heading(col, text=col.upper())
        tabela.column(col, anchor="w")  # Alinhar texto à esquerda
    else:
        tabela.heading(col, text=col.upper())
        tabela.column(col, anchor="center")

scrollbar_y = ttk.Scrollbar(frame_dados, orient="vertical", command=tabela.yview)
scrollbar_y.pack(side="right", fill="y")
tabela.configure(yscrollcommand=scrollbar_y.set)

tabela.bind("<Double-1>", editar_valor)  # Permitir edição ao clicar duas vezes

# Botões superiores (adicionando botão para salvar como JSON)
btn_salvar_json = tk.Button(frame_botoes, text="SALVAR CONTAGEM", command=salvar_json)
btn_salvar_json.grid(row=0, column=2, padx=10)

# Menu de navegação (adicionando opção para salvar como JSON)
menu_navegacao.add_command(label="SALVAR CONTAGEM", command=salvar_json)

# Botão para apura inventário de estoque
btn_apurar = tk.Button(frame_botoes, text="APURAR INVENTÁRIO", command=apurar_inventario)
btn_apurar.grid(row=0, column=7, padx=10)


# Botão para carregar JSON
btn_carregar_json = tk.Button(frame_botoes, text="CARREGAR JSON", command=carregar_json)
btn_carregar_json.grid(row=1, column=7, padx=10)

# Botão ordem crescente
btn_classificar_crescente = tk.Button(
    frame_botoes, text="VL. DIF. -", command=lambda: classificar_vl_dif(ascendente=True)
)
btn_classificar_crescente.grid(row=0, column=9, padx=5)

# Botão ordem decrescente
btn_classificar_decrescente = tk.Button(
    frame_botoes, text="VL. DIF. +", command=lambda: classificar_vl_dif(ascendente=False)
)
btn_classificar_decrescente.grid(row=1, column=9, padx=5)


# Aba Resumo
frame_resumo = tk.Frame(root)

# Inicializa na aba Apuração
frame_dados.pack(fill="both", expand=True)

# Rodapé na parte inferior
rodape = tk.Label(
    root,
    text="Created by: Ricardo Leffers Gomes\nIn: 01/01/2025\nVersion: 001",
    font=("Arial", 10, "italic"),
    anchor="w"
)
rodape.pack(side="bottom", pady=10)

# Rodar o aplicativo
root.mainloop()
