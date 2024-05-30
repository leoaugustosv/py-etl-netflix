import pandas as pd
import os
import glob
import openpyxl as op

# definindo caminhos de entrada e saída
folder_path = "src\\data\\raw"
output_path = os.path.join("src","data","ready","clean.xlsx")

# listar todos os arquivos do caminho
excel_files = glob.glob(os.path.join(folder_path, "*.xlsx"))

print(excel_files)

if not excel_files:
    print("Nenhum arquivo \".xlsx\" encontrado.")
else:
    df = []

    for file in excel_files:
        try:
            #ler arquivo excel
            df_temp = pd.read_excel(file)
            
            #pegar nome do arquivo base
            file_name = os.path.basename(file)

            
            #definindo valor na coluna location
            if "brasil" in file_name.lower():
                df_temp["Localidade"] = "BR"
            elif "france" in file_name.lower():
                df_temp["Localidade"] = "FR"
            elif "italian" in file_name.lower():
                df_temp["Localidade"] = "IT"

            #definindo valor na coluna campaign
            df_temp["Campanha"] = df_temp["utm_link"].str.extract(r"utm_campaign=(.*)")

            #acrescenta df tratado ao df definido anteriormente
            df.append(df_temp)

        except Exception as e:
            print(f"Erro ao ler arquivo {file}: ",e)

    

if df:

    # concatenar todas as subtabelas no dataframe temp em uma unica tabela
    result = pd.concat(df,ignore_index=True)

    # renomear colunas
    result = result.rename(columns={"sale_date":"Data da Venda","Customer ":"Cliente",
                                    "Contracted Plan":"Plano Contratado","Amount":"Valor",
                                    "utm_link":"Link","Age":"Idade"})

    # config do writer
    try:
        writer = pd.ExcelWriter(output_path,engine="xlsxwriter")
        result.to_excel(writer, index=False)
        
        #salvar arquivo
        writer._save()
    except Exception as e:
        print("Erro: ",e)
        


else:
    print("Dataframe vazio. Verifique o ETL e tente novamente!")
