#COD PARA LER PLANILHA DP HALLEN - valor
import pandas as pd

# === CONFIGURAÇÕES PERSONALIZÁVEIS ===
codigo_empresa = 'py -m uvicorn app:app --reload'              # Código da empresa (5 dígitos)
codigo_evento = '123'                 # Código do evento (3 dígitos)
referencia1 = '010426'                # Referência 1 (MMYYYY)
referencia2 = '300426'                # Referência 2 (MMYYYY)
arquivo_excel = 'EMPRESTIMO.xlsx'      # Nome do arquivo Excel de entrada
arquivo_saida = f'saida_evento_{codigo_evento}.txt'  # Nome do arquivo de saída

# === CONFIGURAÇÕES FIXAS ===
processo = 'F'                         # Processo sempre será 'F'
cnpj_empresa = ''                     # CNPJ não utilizado
pis_funcionario = ''                  # PIS não utilizado
departamento = '0000'                 # Código do departamento

# === LEITURA DA PLANILHA (primeira aba) ===
df = pd.read_excel(arquivo_excel)
df.columns = df.columns.str.strip().str.upper()

# Verificação mínima
if 'MATRICULA' not in df.columns or 'VALOR' not in df.columns:
    raise ValueError("A planilha deve conter as colunas 'MATRICULA' e 'VALOR'.")

# Verifica se existe a coluna 'TIPO'
usa_coluna_tipo = 'TIPO' in df.columns

# Se não houver coluna TIPO, pergunta para o usuário
if not usa_coluna_tipo:
    tipo_padrao = input("A planilha contém valores em horas (H) ou dinheiro (R$)? [H/R$]: ").strip().upper()
    if tipo_padrao not in ['H', 'R$']:
        raise ValueError("Você deve informar 'H' para horas ou 'R$' para dinheiro.")

# === CONSTRUÇÃO DO ARQUIVO TXT ===
linhas = []
for i, row in df.iterrows():
    sequencial = str(i + 1).zfill(6)
    matricula = str(row['MATRICULA']).zfill(6)

    # Define tipo da linha
    if usa_coluna_tipo:
        tipo_linha = str(row['TIPO']).strip().upper()
    else:
        tipo_linha = tipo_padrao

    # Conversão do valor conforme o tipo
    valor_str = str(row['VALOR']).strip()
    if tipo_linha == 'H':
        try:
            if ':' in valor_str:
                horas, minutos = map(int, valor_str.split(':'))
                total_minutos = horas * 60 + minutos
            else:
                total_minutos = int(float(valor_str) * 60)
            valor_evento = str(total_minutos).zfill(14)
        except Exception as e:
            raise ValueError(f"Erro ao processar valor de horas '{valor_str}' na linha {i+2}: {e}")
    else:
        try:
            valor_evento = f"{float(valor_str):014.2f}".replace('.', '').zfill(14)
        except Exception as e:
            raise ValueError(f"Erro ao processar valor monetário '{valor_str}' na linha {i+2}: {e}")

    linha = (
        sequencial +                      # 1-6
        codigo_empresa.zfill(5) +         # 7-11
        referencia1.zfill(6) +            # 12-17
        referencia2.zfill(6) +            # 18-23
        '000000' +                        # Faltas (24-29)
        '000000' +                        # Horas trabalhadas (30-35)
        '00' +                            # Dias úteis (36-37)
        codigo_evento.zfill(3) +          # Código do evento (38-40)
        valor_evento +                    # Valor (41-54)
        matricula +                       # Código funcionário (55-60)
        processo +                        # Processo fixo 'F' (61)
        cnpj_empresa.zfill(14) +          # CNPJ (62-75)
        pis_funcionario.zfill(11) +       # PIS (76-86)
        departamento.zfill(4) +           # Departamento (87-90)
        '0' * 14 +                        # CNPJ da operadora (91-104)
        '0000' +                          # Código do plano (105-108)
        '00000' +                         # Código do beneficiário (109-113)
        '0' * 11 +                        # CPF do beneficiário (114-124)
        'N' + 'N' + 'N' + 'N'             # Formas de apuração (125-128)
    )
    linhas.append(linha)

# === SALVAR O ARQUIVO TXT ===
with open(arquivo_saida, 'w', encoding='utf-8') as f:
    f.write('\n'.join(linhas))

print(f'\nArquivo "{arquivo_saida}" gerado com sucesso.')