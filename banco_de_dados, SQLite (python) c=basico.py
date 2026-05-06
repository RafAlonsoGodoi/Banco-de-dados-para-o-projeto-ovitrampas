import sqlite3
import pandas as pd

# conectar ao banco
conexao = sqlite3.connect("monitoramento_relacional.db")
cursor = conexao.cursor()

# apagar tabela antiga (se existir)
cursor.execute("DROP TABLE IF EXISTS monitoramento")

# criar tabela
cursor.execute("""
CREATE TABLE monitoramento (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    bairro TEXT,
    numero_casos INTEGER
)
""")

# dados de cada bairro
dados = [
    ("joao_paulo_da_silva", 1),
    ("jose_inacio", 1),
    ("nova_brasilandia", 1),
    ("jose_rodrigues", 1),
    ("jose_arara", 1),
    ("juvenal_uchoa", 1),
    ("flavio_derzi", 1),
    ("joao_de_abreu", 1),
    ("centro", 1),
    ("isac_honorato", 1),
    ("sao_domingos", 1),
    ("thomas_de_almeida", 1),
    ("mao_amiga", 1),
    ("jardim_brasilia", 1),
    ("jardim_camargo", 1),
    ("coqueiral", 1),
    ("parque_das_araras", 1),
    ("vale_verde_1", 1),
    ("vale_verde_2", 1),
    ("imperial", 1),
    ("primavera", 1)
]

# inserir dados
cursor.executemany("""
INSERT INTO monitoramento (bairro, numero_casos)
VALUES (?, ?)
""", dados)

#  corrigido: atualizar dados corretamente
cursor.execute("""
UPDATE monitoramento
SET numero_casos = ?
WHERE bairro = ?
""", (5, "centro"))

# mostrar dados no terminal
cursor.execute("SELECT * FROM monitoramento")
for linha in cursor.fetchall():
    print(linha)

# salvar alterações no banco
conexao.commit()

#  exportar para Excel
df = pd.read_sql_query("SELECT * FROM monitoramento", conexao)
df.to_excel("monitoramento_c=basico.xlsx", index=False)

print("Planilha Excel criada com sucesso!")

# fechar conexão
conexao.close()