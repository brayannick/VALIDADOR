"""
Validador de Alunos
-------------------
Confere, a partir de uma planilha (XLSX ou CSV), se um novo cadastro
pode ser duplicado antes de ser inserido no sistema da instituição.

O programa é independente do sistema original: não acessa seu código
nem seu banco de dados e não grava nada na planilha.
"""

import json
import logging
import re
import sys
import unicodedata
import tkinter as tk

from datetime import datetime
from pathlib import Path
from tkinter import ttk, filedialog, messagebox

import pandas as pd
from rapidfuzz import fuzz


# -----------------------------------------
# 0. Configurações
# -----------------------------------------

DDD_PADRAO = "88"            # usado quando o telefone vem sem DDD
LIMIAR_PADRAO = 86           # similaridade mínima de nomes (%)
DIAS_ALERTA_BASE = 7         # avisa se a planilha for mais antiga que isto

NIVEIS = {"Forte": 3, "Conferir": 2, "Possível": 1}


def pasta_do_programa():
    """Pasta do script (ou do executável, quando gerado com PyInstaller)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).parent
    return Path(__file__).resolve().parent


ARQUIVO_CONFIG = pasta_do_programa() / "config.json"
ARQUIVO_LOG = pasta_do_programa() / "validador.log"

# O log guarda apenas eventos técnicos, nunca nomes ou telefones.
logging.basicConfig(
    filename=ARQUIVO_LOG,
    level=logging.INFO,
    format="%(asctime)s %(levelname)s %(message)s",
    encoding="utf-8",
)


def carregar_config():
    try:
        return json.loads(ARQUIVO_CONFIG.read_text(encoding="utf-8"))
    except (OSError, ValueError):
        return {}


def salvar_config(config):
    try:
        ARQUIVO_CONFIG.write_text(
            json.dumps(config, ensure_ascii=False, indent=2),
            encoding="utf-8",
        )
    except OSError:
        logging.warning("Não foi possível salvar o config.json")


# -----------------------------------------
# 1. Normalização dos dados
# -----------------------------------------

def texto_limpo(valor):
    """Converte qualquer valor em texto, tratando vazios e NaN."""
    if valor is None:
        return ""
    try:
        if pd.isna(valor):
            return ""
    except (TypeError, ValueError):
        pass
    return str(valor).strip()


def normalizar_nome(valor):
    texto = texto_limpo(valor).lower()

    texto = unicodedata.normalize("NFKD", texto)
    texto = "".join(
        caractere
        for caractere in texto
        if not unicodedata.combining(caractere)
    )

    return re.sub(r"\s+", " ", texto)


def normalizar_telefone(valor, ddd_padrao=""):
    """
    Padroniza um telefone brasileiro em (ddd, numero).

    - remove tudo que não for dígito;
    - remove o prefixo internacional (00) e o código do país (55);
    - remove o zero antes do DDD (ex.: 088...);
    - se vier sem DDD (8 ou 9 dígitos), usa o DDD padrão.

    Retorna ("", "") para vazio e (ddd, numero) nos demais casos.
    O número só é considerado válido com pelo menos 8 dígitos.
    """
    digitos = re.sub(r"\D", "", texto_limpo(valor))

    if not digitos:
        return "", ""

    if digitos.startswith("00"):
        digitos = digitos[2:]

    if len(digitos) in (12, 13) and digitos.startswith("55"):
        digitos = digitos[2:]

    if len(digitos) in (11, 12) and digitos.startswith("0"):
        digitos = digitos[1:]

    if len(digitos) in (8, 9):
        return ddd_padrao, digitos

    if len(digitos) in (10, 11):
        return digitos[:2], digitos[2:]

    # Formato desconhecido: guarda o que foi informado, sem DDD
    return "", digitos


def telefone_valido(ddd_numero):
    return len(ddd_numero[1]) >= 8


def formatar_telefone(ddd_numero):
    ddd, numero = ddd_numero
    if not numero:
        return ""
    if len(numero) == 9:
        corpo = f"{numero[:5]}-{numero[5:]}"
    elif len(numero) == 8:
        corpo = f"{numero[:4]}-{numero[4:]}"
    else:
        corpo = numero
    return f"({ddd}) {corpo}" if ddd else corpo


def telefones_coincidem(a, b):
    """
    Compara dois telefones (ddd, numero).

    Usa os 8 últimos dígitos, o que faz o número antigo de 8 dígitos
    coincidir com a versão atual de 9 dígitos (com o 9 na frente).
    Se os dois tiverem DDD, ele também precisa ser igual.
    """
    ddd_a, num_a = a
    ddd_b, num_b = b

    if len(num_a) < 8 or len(num_b) < 8:
        return False

    if num_a[-8:] != num_b[-8:]:
        return False

    if ddd_a and ddd_b and ddd_a != ddd_b:
        return False

    return True


# -----------------------------------------
# 2. Carregamento da base
# -----------------------------------------

def carregar_base(caminho, ddd_padrao=""):
    extensao = Path(caminho).suffix.lower()

    if extensao == ".xlsx":
        df = pd.read_excel(caminho, dtype=str)
    elif extensao == ".csv":
        df = pd.read_csv(
            caminho,
            dtype=str,
            sep=None,
            engine="python",
            encoding="utf-8-sig",
        )
    else:
        raise ValueError("Selecione um arquivo XLSX ou CSV.")

    # Ignora diferenças de maiúsculas nos cabeçalhos
    colunas = {
        str(coluna).strip().lower(): coluna
        for coluna in df.columns
    }

    if "nome" not in colunas or "telefone" not in colunas:
        raise ValueError(
            "A planilha precisa ter as colunas "
            "'nome' e 'telefone'."
        )

    df = df.rename(columns={
        colunas["nome"]: "nome",
        colunas["telefone"]: "telefone",
    })

    # Número da linha como aparece na planilha (cabeçalho = linha 1)
    df["linha"] = df.index + 2

    # Versões normalizadas, usadas apenas para comparação
    df["nome_norm"] = df["nome"].map(normalizar_nome)
    df["tel_norm"] = df["telefone"].map(
        lambda valor: normalizar_telefone(valor, ddd_padrao)
    )

    df = df[
        (df["nome_norm"] != "")
        | (df["tel_norm"].map(lambda t: t[1]) != "")
    ].copy()

    return df.reset_index(drop=True)


# -----------------------------------------
# 3. Comparação dos cadastros
# -----------------------------------------

def similaridade_nomes(nome_novo, nome_existente):
    """
    Retorna (similaridade, contido).

    similaridade: maior valor entre a comparação direta e a comparação
    que ignora a ordem das palavras (ex.: 'Silva Ana' x 'Ana Silva').

    contido: True quando todas as palavras do nome menor aparecem no
    nome maior (ex.: 'Maria Silva' dentro de 'Maria Oliveira Silva').
    Serve para a pesquisa por partes do nome.
    """
    if not nome_existente:
        return 0, False

    similaridade = max(
        fuzz.ratio(nome_novo, nome_existente),
        fuzz.token_sort_ratio(nome_novo, nome_existente),
    )

    palavras_novo = nome_novo.split()
    palavras_existente = nome_existente.split()

    contido = (
        nome_novo != nome_existente
        and min(len(palavras_novo), len(palavras_existente)) >= 2
        and fuzz.token_set_ratio(nome_novo, nome_existente) == 100
    )

    return similaridade, contido


def verificar_duplicidade(nome, telefone, base,
                          ddd_padrao="", limiar=LIMIAR_PADRAO):
    nome_novo = normalizar_nome(nome)
    tel_novo = normalizar_telefone(telefone, ddd_padrao)

    if not nome_novo:
        raise ValueError("Informe o nome do aluno.")

    if not tel_novo[1]:
        raise ValueError("Informe o telefone do aluno.")

    if not telefone_valido(tel_novo):
        raise ValueError(
            "Telefone inválido: informe pelo menos 8 dígitos "
            "(com DDD, se possível)."
        )

    resultados = []

    for aluno in base.itertuples(index=False):
        nome_existente = aluno.nome_norm
        tel_existente = aluno.tel_norm

        mesmo_nome = bool(nome_existente) and nome_novo == nome_existente
        mesmo_telefone = telefones_coincidem(tel_novo, tel_existente)

        similaridade, contido = similaridade_nomes(
            nome_novo, nome_existente
        )
        nome_parecido = similaridade >= limiar

        if mesmo_nome and mesmo_telefone:
            motivo = "Nome e telefone coincidem"
            nivel = "Forte"
        elif mesmo_telefone and (nome_parecido or contido):
            motivo = (
                "Telefone coincidente e nome semelhante "
                f"({similaridade:.0f}%)"
            )
            nivel = "Forte"
        elif mesmo_telefone:
            motivo = "Telefone coincidente (nome diferente)"
            nivel = "Conferir"
        elif mesmo_nome:
            motivo = "Nome coincidente (telefone diferente)"
            nivel = "Conferir"
        elif nome_parecido:
            motivo = f"Nome semelhante ({similaridade:.0f}%)"
            nivel = "Possível"
        elif contido:
            motivo = "Nome contido em outro cadastro"
            nivel = "Possível"
        else:
            continue

        resultados.append({
            "linha": aluno.linha,
            "nome": texto_limpo(aluno.nome),
            "telefone": texto_limpo(aluno.telefone),
            "motivo": motivo,
            "nivel": nivel,
            "similaridade": similaridade,
        })

    # Prioriza as correspondências mais relevantes
    resultados.sort(
        key=lambda item: (NIVEIS[item["nivel"]], item["similaridade"]),
        reverse=True,
    )

    return resultados


# -----------------------------------------
# 4. Interface gráfica
# -----------------------------------------

class Aplicacao:
    def __init__(self, janela):
        self.janela = janela
        self.janela.title("Validador de Alunos")
        self.janela.geometry("980x640")
        self.janela.minsize(820, 520)

        self.config = carregar_config()

        self.base = None
        self.caminho_base = None
        self.modificado_em = None   # data de modificação do arquivo
        self.carregado_em = None    # momento em que a base foi lida

        self.var_arquivo = tk.StringVar(value="Nenhuma base selecionada")
        self.var_info_base = tk.StringVar(value="")
        self.var_ddd = tk.StringVar(
            value=self.config.get("ddd_padrao", DDD_PADRAO)
        )
        self.var_limiar = tk.IntVar(
            value=self.config.get("limiar", LIMIAR_PADRAO)
        )
        self.var_status = tk.StringVar(
            value="Selecione a base para começar."
        )

        self.criar_interface()
        self.janela.protocol("WM_DELETE_WINDOW", self.fechar)

        # Reabre automaticamente a última base usada
        ultimo = self.config.get("ultima_base")
        if ultimo and Path(ultimo).exists():
            self.carregar(ultimo, silencioso=True)

    # ---------- construção da tela ----------

    def criar_interface(self):
        principal = ttk.Frame(self.janela, padding=16)
        principal.pack(fill="both", expand=True)

        ttk.Label(
            principal,
            text="Validador de cadastros de alunos",
            font=("Arial", 16, "bold"),
        ).pack(anchor="w", pady=(0, 12))

        # Base de referência
        quadro_base = ttk.LabelFrame(
            principal, text="Base de referência", padding=10
        )
        quadro_base.pack(fill="x", pady=(0, 12))

        linha_base = ttk.Frame(quadro_base)
        linha_base.pack(fill="x")

        ttk.Button(
            linha_base,
            text="Selecionar planilha",
            command=self.selecionar_base,
        ).pack(side="left", padx=(0, 8))

        self.botao_atualizar = ttk.Button(
            linha_base,
            text="Atualizar base",
            command=self.atualizar_base,
            state="disabled",
        )
        self.botao_atualizar.pack(side="left", padx=(0, 10))

        ttk.Label(
            linha_base, textvariable=self.var_arquivo
        ).pack(side="left", fill="x", expand=True)

        self.rotulo_info_base = tk.Label(
            quadro_base,
            textvariable=self.var_info_base,
            anchor="w",
            justify="left",
        )
        self.rotulo_info_base.pack(fill="x", pady=(6, 0))

        # Formulário
        formulario = ttk.LabelFrame(
            principal, text="Novo cadastro para verificar", padding=10
        )
        formulario.pack(fill="x", pady=(0, 12))

        ttk.Label(formulario, text="Nome:").grid(
            row=0, column=0, sticky="w", pady=5
        )
        self.campo_nome = ttk.Entry(formulario)
        self.campo_nome.grid(
            row=0, column=1, columnspan=3, sticky="ew", padx=10, pady=5
        )

        ttk.Label(formulario, text="Telefone:").grid(
            row=1, column=0, sticky="w", pady=5
        )
        self.campo_telefone = ttk.Entry(formulario)
        self.campo_telefone.grid(
            row=1, column=1, columnspan=3, sticky="ew", padx=10, pady=5
        )

        # Ajustes
        ajustes = ttk.Frame(formulario)
        ajustes.grid(row=2, column=0, columnspan=4, sticky="w", pady=(4, 0))

        ttk.Label(ajustes, text="DDD padrão:").pack(side="left")
        ttk.Entry(
            ajustes, textvariable=self.var_ddd, width=4
        ).pack(side="left", padx=(6, 20))

        ttk.Label(
            ajustes, text="Similaridade mínima de nomes (%):"
        ).pack(side="left")
        ttk.Spinbox(
            ajustes,
            from_=60, to=100, increment=1,
            textvariable=self.var_limiar,
            width=5,
        ).pack(side="left", padx=6)

        formulario.columnconfigure(1, weight=1)

        botoes = ttk.Frame(formulario)
        botoes.grid(
            row=3, column=0, columnspan=4, sticky="ew", pady=(10, 0)
        )

        ttk.Button(
            botoes,
            text="Verificar duplicidade",
            command=self.verificar,
        ).pack(side="left", fill="x", expand=True)

        ttk.Button(
            botoes, text="Limpar", command=self.limpar
        ).pack(side="left", padx=(8, 0))

        self.janela.bind("<Return>", lambda _evento: self.verificar())

        # Resultados
        quadro_resultados = ttk.LabelFrame(
            principal, text="Possíveis correspondências", padding=8
        )
        quadro_resultados.pack(fill="both", expand=True)

        colunas = ("linha", "nome", "telefone", "motivo", "nivel")

        self.tabela = ttk.Treeview(
            quadro_resultados, columns=colunas, show="headings"
        )

        configuracoes = {
            "linha": ("Linha", 55),
            "nome": ("Nome encontrado", 230),
            "telefone": ("Telefone", 140),
            "motivo": ("Motivo", 300),
            "nivel": ("Classificação", 95),
        }

        for coluna, (titulo, largura) in configuracoes.items():
            self.tabela.heading(coluna, text=titulo)
            self.tabela.column(coluna, width=largura, minwidth=50)

        self.tabela.tag_configure("Forte", background="#f8d7da")
        self.tabela.tag_configure("Conferir", background="#fff3cd")
        self.tabela.tag_configure("Possível", background="#e2e3e5")

        barra = ttk.Scrollbar(
            quadro_resultados, orient="vertical",
            command=self.tabela.yview,
        )
        self.tabela.configure(yscrollcommand=barra.set)

        self.tabela.pack(side="left", fill="both", expand=True)
        barra.pack(side="right", fill="y")

        ttk.Label(
            principal, textvariable=self.var_status, wraplength=920
        ).pack(anchor="w", pady=(8, 0))

    # ---------- base de dados ----------

    def ddd_atual(self):
        return re.sub(r"\D", "", self.var_ddd.get())[:2]

    def selecionar_base(self):
        caminho = filedialog.askopenfilename(
            title="Selecione a base de alunos",
            filetypes=[
                ("Planilhas e arquivos CSV", "*.xlsx *.csv"),
                ("Excel", "*.xlsx"),
                ("CSV", "*.csv"),
            ],
        )

        if caminho:
            self.carregar(caminho)

    def atualizar_base(self):
        if self.caminho_base:
            self.carregar(self.caminho_base)

    def carregar(self, caminho, silencioso=False):
        try:
            nova_base = carregar_base(caminho, self.ddd_atual())
        except PermissionError:
            self.falha_carga(
                "Não foi possível abrir o arquivo. "
                "Se ele estiver aberto no Excel, feche-o e tente de novo.",
                silencioso,
            )
            return
        except Exception as erro:
            logging.exception("Falha ao carregar a base")
            self.falha_carga(str(erro), silencioso)
            return

        self.base = nova_base
        self.caminho_base = caminho
        self.modificado_em = Path(caminho).stat().st_mtime
        self.carregado_em = datetime.now()

        self.var_arquivo.set(Path(caminho).name)
        self.botao_atualizar.configure(state="normal")
        self.atualizar_info_base()
        self.var_status.set(f"Base carregada: {len(self.base)} registros.")

        self.config["ultima_base"] = str(caminho)
        salvar_config(self.config)
        logging.info("Base carregada: %d registros", len(self.base))

    def falha_carga(self, mensagem, silencioso):
        self.base = None
        self.var_arquivo.set("Falha ao carregar a base")
        self.var_info_base.set("")
        if not silencioso:
            messagebox.showerror("Erro na base", mensagem)
        else:
            self.var_status.set(
                "A última base usada não pôde ser aberta. "
                "Selecione a planilha novamente."
            )

    def atualizar_info_base(self):
        """Mostra a data da exportação e avisa se a base estiver antiga."""
        exportada = datetime.fromtimestamp(self.modificado_em)
        dias = (datetime.now() - exportada).days

        texto = (
            f"{len(self.base)} registros  |  "
            f"arquivo modificado em {exportada:%d/%m/%Y %H:%M}  |  "
            f"lido em {self.carregado_em:%d/%m/%Y %H:%M}"
        )

        if dias >= DIAS_ALERTA_BASE:
            texto += (
                f"\nAtenção: a planilha tem {dias} dias. Atualize a "
                "exportação para não perder cadastros recentes."
            )
            self.rotulo_info_base.configure(fg="#b02a37")
        else:
            self.rotulo_info_base.configure(fg="#2e7d32")

        self.var_info_base.set(texto)

    def recarregar_se_modificada(self):
        """Se a planilha mudou no disco desde a leitura, recarrega."""
        if not self.caminho_base:
            return
        try:
            atual = Path(self.caminho_base).stat().st_mtime
        except OSError:
            return
        if atual != self.modificado_em:
            self.carregar(self.caminho_base)

    # ---------- ações ----------

    def verificar(self):
        if self.base is None:
            messagebox.showwarning(
                "Base necessária",
                "Selecione uma planilha válida primeiro.",
            )
            return

        self.recarregar_se_modificada()

        try:
            limiar = int(self.var_limiar.get())
        except (tk.TclError, ValueError):
            messagebox.showwarning(
                "Similaridade inválida",
                "Informe um número inteiro entre 60 e 100.",
            )
            return

        limiar = max(60, min(100, limiar))
        self.config["ddd_padrao"] = self.ddd_atual()
        self.config["limiar"] = limiar
        salvar_config(self.config)

        try:
            resultados = verificar_duplicidade(
                self.campo_nome.get(),
                self.campo_telefone.get(),
                self.base,
                ddd_padrao=self.ddd_atual(),
                limiar=limiar,
            )
        except ValueError as erro:
            messagebox.showwarning("Dados incompletos", str(erro))
            return

        self.limpar_tabela()

        for resultado in resultados:
            self.tabela.insert(
                "", "end",
                values=(
                    resultado["linha"],
                    resultado["nome"],
                    resultado["telefone"],
                    resultado["motivo"],
                    resultado["nivel"],
                ),
                tags=(resultado["nivel"],),
            )

        logging.info("Verificação concluída: %d correspondências",
                     len(resultados))

        if resultados:
            self.var_status.set(
                f"{len(resultados)} possível(is) correspondência(s). "
                "Confira no sistema original antes de cadastrar."
            )
        else:
            self.var_status.set(
                "Nenhuma correspondência encontrada pelas regras atuais. "
                "Isso não garante que o aluno seja novo: a planilha pode "
                "estar desatualizada."
            )

    def limpar_tabela(self):
        for item in self.tabela.get_children():
            self.tabela.delete(item)

    def limpar(self):
        self.campo_nome.delete(0, "end")
        self.campo_telefone.delete(0, "end")
        self.limpar_tabela()
        self.campo_nome.focus_set()
        if self.base is not None:
            self.var_status.set(f"Base carregada: {len(self.base)} registros.")

    def fechar(self):
        self.config["ddd_padrao"] = self.ddd_atual()
        salvar_config(self.config)
        self.janela.destroy()


# -----------------------------------------
# 4b. Comparação de duas planilhas de uma só vez
# -----------------------------------------

def comparar_planilhas(novos, base, ddd_padrao="", limiar=LIMIAR_PADRAO):
    """
    Compara uma planilha de novos cadastros com a base de uma só vez.

    Cada linha de 'novos' passa pela mesma verificação usada na pesquisa
    individual (verificar_duplicidade). Retorna um dicionário com:

    - "existentes": cadastros novos que têm correspondência na base,
      cada um com a lista de correspondências encontradas;
    - "novos": cadastros sem nenhuma correspondência na base;
    - "invalidos": linhas que não puderam ser verificadas (nome ou
      telefone ausente ou inválido), com o motivo.
    """
    existentes = []
    sem_correspondencia = []
    invalidos = []

    for novo in novos.itertuples(index=False):
        registro = {
            "linha": novo.linha,
            "nome": texto_limpo(novo.nome),
            "telefone": texto_limpo(novo.telefone),
        }

        try:
            correspondencias = verificar_duplicidade(
                novo.nome,
                novo.telefone,
                base,
                ddd_padrao=ddd_padrao,
                limiar=limiar,
            )
        except ValueError as erro:
            invalidos.append({**registro, "motivo": str(erro)})
            continue

        if correspondencias:
            existentes.append(
                {**registro, "correspondencias": correspondencias}
            )
        else:
            sem_correspondencia.append(registro)

    return {
        "existentes": existentes,
        "novos": sem_correspondencia,
        "invalidos": invalidos,
    }


class JanelaComparacao(tk.Toplevel):
    """Janela para comparar duas planilhas (base x novos cadastros)."""

    FILTROS = [
        ("Planilhas e arquivos CSV", "*.xlsx *.csv"),
        ("Excel", "*.xlsx"),
        ("CSV", "*.csv"),
    ]

    def __init__(self, app):
        super().__init__(app.janela)
        self.app = app
        self.title("Comparar duas planilhas")
        self.geometry("1020x640")
        self.minsize(860, 520)

        # A base já carregada na janela principal é sugerida por padrão
        self.caminho_base = app.caminho_base
        self.caminho_novos = None

        self.var_base = tk.StringVar(
            value=(
                Path(self.caminho_base).name
                if self.caminho_base
                else "Nenhuma planilha selecionada"
            )
        )
        self.var_novos = tk.StringVar(value="Nenhuma planilha selecionada")
        self.var_resumo = tk.StringVar(
            value="Selecione as duas planilhas e clique em Comparar."
        )

        self.criar_interface()
        self.transient(app.janela)
        self.focus_set()

    # ---------- construção da tela ----------

    def criar_interface(self):
        principal = ttk.Frame(self, padding=16)
        principal.pack(fill="both", expand=True)

        ttk.Label(
            principal,
            text="Comparar duas planilhas",
            font=("Arial", 16, "bold"),
        ).pack(anchor="w", pady=(0, 12))

        quadro = ttk.LabelFrame(principal, text="Planilhas", padding=10)
        quadro.pack(fill="x", pady=(0, 12))

        ttk.Label(quadro, text="Base (cadastros existentes):").grid(
            row=0, column=0, sticky="w", pady=5
        )
        ttk.Button(
            quadro, text="Selecionar...", command=self.selecionar_base
        ).grid(row=0, column=1, padx=10, pady=5)
        ttk.Label(quadro, textvariable=self.var_base).grid(
            row=0, column=2, sticky="w"
        )

        ttk.Label(quadro, text="Novos cadastros a verificar:").grid(
            row=1, column=0, sticky="w", pady=5
        )
        ttk.Button(
            quadro, text="Selecionar...", command=self.selecionar_novos
        ).grid(row=1, column=1, padx=10, pady=5)
        ttk.Label(quadro, textvariable=self.var_novos).grid(
            row=1, column=2, sticky="w"
        )

        ttk.Label(
            quadro,
            text=(
                "As duas planilhas precisam ter as colunas 'nome' e "
                "'telefone'. Usa o DDD padrão e a similaridade mínima "
                "definidos na janela principal."
            ),
        ).grid(row=2, column=0, columnspan=3, sticky="w", pady=(4, 0))

        quadro.columnconfigure(2, weight=1)

        botoes = ttk.Frame(principal)
        botoes.pack(fill="x", pady=(0, 12))

        ttk.Button(
            botoes, text="Comparar planilhas", command=self.comparar
        ).pack(side="left", fill="x", expand=True)

        ttk.Button(
            botoes, text="Fechar", command=self.destroy
        ).pack(side="left", padx=(8, 0))

        self.rotulo_resumo = tk.Label(
            principal,
            textvariable=self.var_resumo,
            anchor="w",
            justify="left",
            wraplength=940,
        )
        self.rotulo_resumo.pack(fill="x", pady=(0, 8))

        self.abas = ttk.Notebook(principal)
        self.abas.pack(fill="both", expand=True)

        self.tabela_existentes = self.criar_tabela(
            "Já existentes",
            {
                "linha_novo": ("Linha (novo)", 85),
                "nome_novo": ("Nome (novo)", 170),
                "telefone_novo": ("Telefone (novo)", 125),
                "linha_base": ("Linha (base)", 85),
                "nome_base": ("Nome na base", 170),
                "telefone_base": ("Telefone na base", 125),
                "motivo": ("Motivo", 250),
                "nivel": ("Classificação", 90),
            },
        )
        self.tabela_existentes.tag_configure("Forte", background="#f8d7da")
        self.tabela_existentes.tag_configure(
            "Conferir", background="#fff3cd"
        )
        self.tabela_existentes.tag_configure(
            "Possível", background="#e2e3e5"
        )

        self.tabela_novos = self.criar_tabela(
            "Não encontrados",
            {
                "linha": ("Linha", 70),
                "nome": ("Nome", 400),
                "telefone": ("Telefone", 200),
            },
        )

        self.tabela_invalidos = self.criar_tabela(
            "Não verificados",
            {
                "linha": ("Linha", 70),
                "nome": ("Nome", 250),
                "telefone": ("Telefone", 160),
                "motivo": ("Motivo", 400),
            },
        )

    def criar_tabela(self, titulo_aba, configuracoes):
        quadro = ttk.Frame(self.abas, padding=6)
        self.abas.add(quadro, text=titulo_aba)

        tabela = ttk.Treeview(
            quadro, columns=tuple(configuracoes), show="headings"
        )

        for coluna, (titulo, largura) in configuracoes.items():
            tabela.heading(coluna, text=titulo)
            tabela.column(coluna, width=largura, minwidth=50)

        barra = ttk.Scrollbar(
            quadro, orient="vertical", command=tabela.yview
        )
        tabela.configure(yscrollcommand=barra.set)

        tabela.pack(side="left", fill="both", expand=True)
        barra.pack(side="right", fill="y")

        return tabela

    # ---------- seleção das planilhas ----------

    def selecionar_base(self):
        caminho = filedialog.askopenfilename(
            parent=self,
            title="Selecione a base de cadastros existentes",
            filetypes=self.FILTROS,
        )
        if caminho:
            self.caminho_base = caminho
            self.var_base.set(Path(caminho).name)

    def selecionar_novos(self):
        caminho = filedialog.askopenfilename(
            parent=self,
            title="Selecione a planilha de novos cadastros",
            filetypes=self.FILTROS,
        )
        if caminho:
            self.caminho_novos = caminho
            self.var_novos.set(Path(caminho).name)

    # ---------- comparação ----------

    def carregar(self, caminho, rotulo, ddd):
        """Carrega uma planilha; em caso de erro, avisa e devolve None."""
        try:
            return carregar_base(caminho, ddd)
        except PermissionError:
            messagebox.showerror(
                "Erro na planilha",
                f"Não foi possível abrir a planilha {rotulo}. "
                "Se ela estiver aberta no Excel, feche-a e tente de novo.",
                parent=self,
            )
        except Exception as erro:
            logging.exception("Falha ao carregar a planilha %s", rotulo)
            messagebox.showerror(
                "Erro na planilha",
                f"Planilha {rotulo}: {erro}",
                parent=self,
            )
        return None

    def comparar(self):
        if not self.caminho_base or not self.caminho_novos:
            messagebox.showwarning(
                "Planilhas necessárias",
                "Selecione a base e a planilha de novos cadastros.",
                parent=self,
            )
            return

        try:
            limiar = int(self.app.var_limiar.get())
        except (tk.TclError, ValueError):
            messagebox.showwarning(
                "Similaridade inválida",
                "Informe, na janela principal, um número inteiro "
                "entre 60 e 100.",
                parent=self,
            )
            return

        limiar = max(60, min(100, limiar))
        ddd = self.app.ddd_atual()

        base = self.carregar(self.caminho_base, "da base", ddd)
        if base is None:
            return

        novos = self.carregar(
            self.caminho_novos, "de novos cadastros", ddd
        )
        if novos is None:
            return

        resultado = comparar_planilhas(
            novos, base, ddd_padrao=ddd, limiar=limiar
        )
        self.exibir(resultado, total=len(novos), tamanho_base=len(base))

        logging.info(
            "Comparação de planilhas: %d novos, %d existentes, "
            "%d sem correspondência, %d não verificados",
            len(novos),
            len(resultado["existentes"]),
            len(resultado["novos"]),
            len(resultado["invalidos"]),
        )

    def limpar_tabelas(self):
        for tabela in (
            self.tabela_existentes,
            self.tabela_novos,
            self.tabela_invalidos,
        ):
            for item in tabela.get_children():
                tabela.delete(item)

    def exibir(self, resultado, total, tamanho_base):
        self.limpar_tabelas()

        existentes = resultado["existentes"]
        novos = resultado["novos"]
        invalidos = resultado["invalidos"]

        for cadastro in existentes:
            for achado in cadastro["correspondencias"]:
                self.tabela_existentes.insert(
                    "", "end",
                    values=(
                        cadastro["linha"],
                        cadastro["nome"],
                        cadastro["telefone"],
                        achado["linha"],
                        achado["nome"],
                        achado["telefone"],
                        achado["motivo"],
                        achado["nivel"],
                    ),
                    tags=(achado["nivel"],),
                )

        for cadastro in novos:
            self.tabela_novos.insert(
                "", "end",
                values=(
                    cadastro["linha"],
                    cadastro["nome"],
                    cadastro["telefone"],
                ),
            )

        for cadastro in invalidos:
            self.tabela_invalidos.insert(
                "", "end",
                values=(
                    cadastro["linha"],
                    cadastro["nome"],
                    cadastro["telefone"],
                    cadastro["motivo"],
                ),
            )

        self.abas.tab(0, text=f"Já existentes ({len(existentes)})")
        self.abas.tab(1, text=f"Não encontrados ({len(novos)})")
        self.abas.tab(2, text=f"Não verificados ({len(invalidos)})")

        if total == 0:
            texto = "A planilha de novos cadastros está vazia."
            cor = "#b02a37"
            self.abas.select(1)
        elif existentes:
            texto = (
                f"{len(existentes)} de {total} cadastro(s) da planilha "
                "nova têm correspondência na base (veja a classificação "
                f"de cada um). {len(novos)} não foi(ram) encontrado(s)."
            )
            cor = "#b02a37"
            self.abas.select(0)
        else:
            texto = (
                f"Nenhum dos {total} cadastro(s) da planilha nova existe "
                "na base pelas regras atuais."
            )
            cor = "#2e7d32"
            self.abas.select(1)

        if invalidos:
            texto += (
                f"\n{len(invalidos)} linha(s) não puderam ser verificadas "
                "(nome ou telefone ausente ou inválido); veja a aba "
                "'Não verificados'."
            )

        texto += (
            f"\nBase com {tamanho_base} registros. Confira no sistema "
            "original antes de cadastrar: a planilha pode estar "
            "desatualizada."
        )

        self.rotulo_resumo.configure(fg=cor)
        self.var_resumo.set(texto)


class AplicacaoComparacao(Aplicacao):
    """
    Aplicação original acrescida do botão que abre a comparação
    de duas planilhas. Todo o restante é herdado sem alterações.
    """

    def criar_interface(self):
        super().criar_interface()
        self.janela_comparacao = None

        principal = self.janela.winfo_children()[0]
        ttk.Button(
            principal,
            text="Comparar duas planilhas...",
            command=self.abrir_comparacao,
        ).pack(anchor="e", pady=(8, 0))

    def abrir_comparacao(self):
        janela = self.janela_comparacao
        if janela is not None and janela.winfo_exists():
            janela.lift()
            janela.focus_set()
            return
        self.janela_comparacao = JanelaComparacao(self)


# -----------------------------------------
# 5. Inicialização
# -----------------------------------------

if __name__ == "__main__":
    janela = tk.Tk()
    aplicacao = AplicacaoComparacao(janela)
    janela.mainloop()