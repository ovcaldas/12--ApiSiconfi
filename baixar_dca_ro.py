import requests
import pandas as pd


URL = "https://apidatalake.tesouro.gov.br/ords/siconfi/tt/dca"
ID_ENTE = "11"
ANEXO = "DCA-Anexo I-E"
ANO_INICIAL = 2015
ANO_FINAL = 2025
ARQUIVO_SAIDA = "DCA_RO_2015_2025.xlsx"
NOME_ABA = "DCA_Anexo_IE"


def coletar_dados():
    dados_anos = []

    for ano in range(ANO_INICIAL, ANO_FINAL + 1):
        params = {
            "an_exercicio": ano,
            "no_anexo": ANEXO,
            "id_ente": ID_ENTE,
        }

        try:
            response = requests.get(URL, params=params, timeout=60)
            response.raise_for_status()
            dados = response.json()
            registros = dados.get("items", [])

            if not registros:
                print(f"{ano}: nenhum registro obtido.")
                continue

            df_ano = pd.DataFrame(registros)
            df_ano["ano_consulta"] = ano
            dados_anos.append(df_ano)
            print(f"{ano}: {len(df_ano)} registros")
        except requests.RequestException as erro:
            print(f"{ano}: erro na requisição: {erro}")
        except (ValueError, TypeError) as erro:
            print(f"{ano}: erro ao interpretar a resposta: {erro}")

    return dados_anos


def salvar_excel(dados_anos):
    if dados_anos:
        df_dca = pd.concat(dados_anos, ignore_index=True)
    else:
        df_dca = pd.DataFrame(columns=["ano_consulta"])

    df_dca.to_excel(ARQUIVO_SAIDA, sheet_name=NOME_ABA, index=False)
    print(f"Arquivo {ARQUIVO_SAIDA} criado com sucesso.")
    print(f"Total de registros consolidados: {len(df_dca)}")


if __name__ == "__main__":
    dados_anos = coletar_dados()
    salvar_excel(dados_anos)
