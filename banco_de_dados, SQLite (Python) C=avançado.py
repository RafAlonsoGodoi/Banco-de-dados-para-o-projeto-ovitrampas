"""
Monitoramento de Arboviroses — Brasilândia/MS
Banco de dados SQLite com exportação para Excel
"""

import sqlite3
import pandas as pd
from pathlib import Path
from datetime import date

# ─────────────────────────────────────────
# CONFIGURAÇÕES
# ─────────────────────────────────────────
DB_PATH    = Path("monitoramento_C=avançado.db")
XLSX_PATH  = Path("monitoramento_C=avançado.xlsx")


# ─────────────────────────────────────────
# DADOS INICIAIS
# Formato: (bairro, casos_possiveis, casos_confirmados)
# ─────────────────────────────────────────
DADOS_INICIAIS: list[tuple[str, int, int]] = [
    ("João Paulo da Silva",  0, 0),
    ("José Inácio",          0, 0),
    ("Nova Brasilândia",     0, 0),
    ("José Rodrigues",       0, 0),
    ("José Arara",           0, 0),
    ("Juvenal Uchôa",        0, 0),
    ("Flávio Derzi",         0, 0),
    ("João de Abreu",        0, 0),
    ("Centro",               0, 0),
    ("Isaac Honorato",       0, 0),
    ("São Domingos",         0, 0),
    ("Thomas de Almeida",    0, 0),
    ("Mão Amiga",            0, 0),
    ("Jardim Brasília",      0, 0),
    ("Jardim Camargo",       0, 0),
    ("Coqueiral",            0, 0),
    ("Parque das Araras",    0, 0),
    ("Vale Verde 1",         0, 0),
    ("Vale Verde 2",         0, 0),
    ("Imperial",             0, 0),
    ("Primavera",            0, 0),
]


# ─────────────────────────────────────────
# FUNÇÕES
# ─────────────────────────────────────────

def criar_tabelas(cursor: sqlite3.Cursor) -> None:
    """Cria as tabelas do banco se não existirem."""

    # Tabela principal de monitoramento
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS monitoramento (
            id                  INTEGER PRIMARY KEY AUTOINCREMENT,
            bairro              TEXT NOT NULL UNIQUE,
            casos_possiveis     INTEGER NOT NULL DEFAULT 0,
            casos_confirmados   INTEGER NOT NULL DEFAULT 0,
            atualizado_em       TEXT NOT NULL DEFAULT (date('now'))
        )
    """)

    # Tabela de histórico — registra cada alteração
    cursor.execute("""
        CREATE TABLE IF NOT EXISTS historico (
            id                  INTEGER PRIMARY KEY AUTOINCREMENT,
            bairro              TEXT NOT NULL,
            casos_possiveis     INTEGER NOT NULL,
            casos_confirmados   INTEGER NOT NULL,
            registrado_em       TEXT NOT NULL DEFAULT (date('now'))
        )
    """)


def inserir_bairros(cursor: sqlite3.Cursor, dados: list[tuple]) -> None:
    """Insere bairros — ignora se já existirem (INSERT OR IGNORE)."""
    cursor.executemany("""
        INSERT OR IGNORE INTO monitoramento (bairro, casos_possiveis, casos_confirmados)
        VALUES (?, ?, ?)
    """, dados)


def atualizar_casos(
    cursor: sqlite3.Cursor,
    bairro: str,
    casos_possiveis: int,
    casos_confirmados: int,
) -> None:
    """
    Atualiza os casos de um bairro e grava no histórico automaticamente.
    Lança ValueError se o bairro não existir.
    """
    cursor.execute("SELECT id FROM monitoramento WHERE bairro = ?", (bairro,))
    if not cursor.fetchone():
        raise ValueError(f"Bairro '{bairro}' não encontrado no banco.")

    hoje = date.today().isoformat()

    cursor.execute("""
        UPDATE monitoramento
        SET casos_possiveis   = ?,
            casos_confirmados = ?,
            atualizado_em     = ?
        WHERE bairro = ?
    """, (casos_possiveis, casos_confirmados, hoje, bairro))

    # Registrar no histórico
    cursor.execute("""
        INSERT INTO historico (bairro, casos_possiveis, casos_confirmados, registrado_em)
        VALUES (?, ?, ?, ?)
    """, (bairro, casos_possiveis, casos_confirmados, hoje))


def exibir_tabela(conexao: sqlite3.Connection) -> None:
    """Imprime a tabela de monitoramento formatada no terminal."""
    df = pd.read_sql_query(
        "SELECT bairro, casos_possiveis, casos_confirmados, atualizado_em "
        "FROM monitoramento ORDER BY bairro",
        conexao,
    )
    print("\nMONITORAMENTO DE ARBOVIROSES - BRASILANDIA/MS")
    print("=" * 65)
    print(df.to_string(index=False))
    print(f"\nTotal de bairros : {len(df)}")
    print(f"Casos possiveis  : {df['casos_possiveis'].sum()}")
    print(f"Casos confirmados: {df['casos_confirmados'].sum()}")
    print("=" * 65)


def exportar_excel(conexao: sqlite3.Connection, caminho: Path) -> None:
    """Exporta monitoramento e histórico para abas separadas no Excel."""
    df_monitor = pd.read_sql_query(
        "SELECT bairro, casos_possiveis, casos_confirmados, atualizado_em "
        "FROM monitoramento ORDER BY bairro",
        conexao,
    )
    df_historico = pd.read_sql_query(
        "SELECT bairro, casos_possiveis, casos_confirmados, registrado_em "
        "FROM historico ORDER BY registrado_em DESC, bairro",
        conexao,
    )

    with pd.ExcelWriter(caminho, engine="openpyxl") as writer:
        df_monitor.to_excel(writer, sheet_name="Monitoramento", index=False)
        df_historico.to_excel(writer, sheet_name="Histórico", index=False)

    print(f"\nExcel exportado -> {caminho.resolve()}")


# ─────────────────────────────────────────
# EXECUÇÃO PRINCIPAL
# ─────────────────────────────────────────

def main() -> None:
    # Usar context manager garante commit/rollback automático
    with sqlite3.connect(DB_PATH) as conexao:
        cursor = conexao.cursor()

        criar_tabelas(cursor)
        inserir_bairros(cursor, DADOS_INICIAIS)

        # ── Exemplo de atualização de casos ──────────────────────────
        atualizar_casos(cursor, "Centro",         casos_possiveis=8,  casos_confirmados=5)
        atualizar_casos(cursor, "Jardim Camargo", casos_possiveis=4,  casos_confirmados=2)
        atualizar_casos(cursor, "Primavera",      casos_possiveis=3,  casos_confirmados=1)
        # ─────────────────────────────────────────────────────────────

        exibir_tabela(conexao)
        exportar_excel(conexao, XLSX_PATH)

    print(f"\nBanco salvo -> {DB_PATH.resolve()}")


if __name__ == "__main__":
    main()
