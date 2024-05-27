import pandas as pd
from tkinter import filedialog
from tkinter import Tk
from sqlalchemy import create_engine
import os

def importar_planilhas_para_bd(planilhas, engine):
    # Criando um DataFrame vazio para armazenar dados de todas as planilhas
    df_concatenado = pd.DataFrame()

    for i, planilha in enumerate(planilhas, 1):
        # Lendo apenas as primeiras 16 colunas de cada planilha
        print(f"Lendo planilha {i}/{len(planilhas)}: {planilha}")
        df = pd.read_excel(planilha, usecols=range(16))
        df_concatenado = pd.concat([df_concatenado, df], ignore_index=True)

    # Importando o DataFrame combinado para o banco de dados
    print("Importando DataFrame para o banco de dados...")
    df_concatenado.to_sql('consulta', con=engine, if_exists='append', index=False)
    print("DataFrame importado com sucesso para a tabela 'consulta'.")

def main():
    # Configurando a janela de diálogo para seleção de arquivos
    root = Tk()
    root.withdraw()  # Ocultando a janela principal

    # Solicitando ao usuário a seleção de planilhas
    planilhas = filedialog.askopenfilenames(title="Selecione as planilhas", filetypes=[("Planilhas Excel", "*.xlsx")])

    # Definindo as informações de conexão diretamente
    usuario = '*'
    senha = '*'
    host = '*'
    nome_banco = '*'

    # Criando a URL de conexão com o banco de dados usando SQLAlchemy
    url_conexao = f"mysql://{usuario}:{senha}@{host}:3306/{nome_banco}"
    engine = create_engine(url_conexao)

    # Importando todas as planilhas para uma tabela única no banco de dados
    importar_planilhas_para_bd(planilhas, engine)

    # Fechando a conexão com o banco de dados
    engine.dispose()

if __name__ == "__main__":
    main()
