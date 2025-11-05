import pandas as pd

def ler_planilha(caminho_arquivo):
    df = pd.read_excel(caminho_arquivo, engine='openpyxl', skiprows=7)
    df.columns = df.columns.str.strip()
    return df

def filtrar_presenciais(df):
    return df[df['Presencial'].str.strip().str.upper() == 'SIM']

def associar_biopsicossocial(df):
    # Converte Horário
    df['Horário'] = pd.to_datetime(df['Horário'], errors='coerce')

    tipos_principal = [
        'PERICIA_POR ESPECIALIDADE',
        'PERICIA MEDICA',
        'PERICIA EM JUNTA MEDICA',
        'PERICIA MEDICA COMPLEXA'
    ]

    # Identifica biopsicossociais
    df['is_biopsicossocial'] = df['Tipo de perícia médica'].str.upper().isin(['SOCIAL', 'PSICOLOGICA', 'FISIOTERAPEUTA'])

    # Cria uma cópia para não modificar original
    df_final = df.copy()

    # Para cada paciente, associa biopsicossocial ao profissional principal
    for paciente, grupo in df.groupby('Nome'):
        # Profissional principal do paciente
        principal = grupo[grupo['Tipo de perícia médica'].str.upper().isin(tipos_principal)]
        bio = grupo[grupo['is_biopsicossocial']]

        if not principal.empty and not bio.empty:
            # Associa todos os biopsicossociais ao primeiro profissional principal
            df_final.loc[bio.index, 'Profissional de Saúde'] = principal.iloc[0]['Profissional de Saúde']

    df_final = df_final.sort_values(['Profissional de Saúde', 'Horário']).reset_index(drop=True)
    df_final = df_final.drop(columns=['is_biopsicossocial'])
    return df_final

def transformar_colunas_com_linhas(df):
    df_final = pd.DataFrame()
    df_final['HORA'] = df['Horário'].dt.time
    df_final['MATRÍCULA'] = df['Matrícula']
    df_final['N     O     M     E'] = df['Nome']
    df_final['" ORGÃO\n DA SEGURANÇA"'] = df['Órgão/Empresa do servidor']
    df_final['CONTATO'] = df['Profissional de Saúde']
    df_final['ORDEM DE CHEGADA'] = ''
    df_final['C'] = ''
    df_final['NC'] = ''
    df_final['OBSERVAÇÃO'] = ''
    df_final['BIOPSICOSSOCIAL'] = df['Tipo de perícia médica']
    df_final['BENEFÍCIO'] = df['Benefício']

    resultado = []
    profissional_atual = None

    for _, row in df_final.iterrows():
        if profissional_atual is None:
            profissional_atual = row['CONTATO']
        elif row['CONTATO'] != profissional_atual:
            # Duas linhas em branco entre profissionais
            resultado.append(pd.Series({col: '' for col in df_final.columns}))
            resultado.append(pd.Series({col: '' for col in df_final.columns}))
            profissional_atual = row['CONTATO']
        resultado.append(row)

    df_final_linhas = pd.DataFrame(resultado)
    return df_final_linhas

def exportar_planilha(df, caminho_saida):
    df.to_excel(caminho_saida, index=False, engine='openpyxl')
    print(f"Planilha exportada com sucesso: {caminho_saida}")

def main():
    arquivo_entrada = "atendimento.xlsx"
    arquivo_saida = "atendimento_personalizado.xlsx"

    df = ler_planilha(arquivo_entrada)
    df = filtrar_presenciais(df)
    df = associar_biopsicossocial(df)
    df_final = transformar_colunas_com_linhas(df)
    exportar_planilha(df_final, arquivo_saida)

if __name__ == "__main__":
    main()
