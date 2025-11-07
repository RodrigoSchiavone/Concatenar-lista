import pandas as pd  

produtos_dtype = {
    'Código LEGO': 'int64',
    'Código M.Cassab': 'int64',
    'Nome': 'object',
    'Linha': 'object',
    'Sub-Linha': 'object',
    'Idade': 'object',
    'Status Item': 'object',
    'Mês Lanç.': 'object',
    'Data Saida': 'object',
    'data abertura Lojas lego': 'object',
    'Distr.': 'object',
    'Status Entrega': 'float64',
    'Pedido Quant.': 'float64',
    'Pedido R$': 'float64',
    'NIP': 'object',
    'ICMS 18%': 'object',
    'ICMS 4%': 'object',
    'Sug. Varejo': 'object',
    'Emb. Coletiva': 'object',
    'CÓDIGO DE BARRAS': 'object'
}
    

produtos = pd.read_excel("MATRIZ.xlsx",dtype=produtos_dtype)

colunas_para_float = ['NIP', 'ICMS 18% ', 'ICMS 4% ', 'Sug. Varejo ']

for col in colunas_para_float:
    try:
        produtos[col] = produtos[col].astype(str)
        
        produtos[col] = produtos[col].str.replace(r'[^\d,.-]', '', regex=True)

        produtos[col] = produtos[col].str.replace(',', '.', regex=False)
        
        produtos[col] = pd.to_numeric(produtos[col], errors='coerce')
        
    except KeyError:
        print(f"❌ Erro: Coluna '{col}' não encontrada no DataFrame.")
    except Exception as e:
        print(f"⚠️ Erro inesperado ao processar a coluna '{col}': {e}")


lojas_dtype = {
    'Cod':'int64',
    'Loja':'object'
}

try:
    lojas = pd.read_excel("Lojas.xlsx", dtype=lojas_dtype)
except FileNotFoundError:
    print("❌ Erro: Arquivo 'Lojas.xlsx' não encontrado. Usando dados de exemplo.")
    # Crie um DataFrame de lojas de exemplo para continuar
    data_lojas = {
        'Cod': [1, 2, 3],
        'Loja': ['Loja A', 'Loja B', 'Loja C']
    }
    lojas = pd.DataFrame(data_lojas)


# --- Combinação (Cross Join) ---
# O merge com how='cross' cria todas as combinações de linhas
tabela_combinada = lojas.merge(produtos, how='cross')

# --- Reorganização das Colunas ---
# Opcional, mas recomendado para ter 'Cod' e 'Loja' no início
colunas_lojas = ['Cod', 'Loja']
colunas_produtos = [col for col in tabela_combinada.columns if col not in colunas_lojas]

tabela_final = tabela_combinada[colunas_lojas + colunas_produtos]

# --- Salvando em um novo arquivo Excel ---
try:
    tabela_final.to_excel("Tabela_Produtos_por_Loja.xlsx", index=False)
    print("\n✅ Sucesso! O arquivo 'Tabela_Produtos_por_Loja.xlsx' foi criado.")
except Exception as e:
    print(f"\n❌ Erro ao salvar o arquivo Excel: {e}")

# Exemplo da estrutura final (apenas as primeiras colunas para ilustrar)
print("\nPrimeiras linhas da Tabela Final (exemplo):")
print(tabela_final[['Cod', 'Loja', 'Nome', 'Código LEGO']].head(6))


