import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.patches as patches
from matplotlib import rcParams

def criar_mapa_mesa(arquivo_excel, planilha_nome, coluna_nome, coluna_cadeira, arquivo_saida):
    # Configurar fonte
    rcParams['font.family'] = 'sans-serif'
    rcParams['font.sans-serif'] = ['Arial', 'DejaVu Sans']

    # Carregar dados do Excel
    try:
        df = pd.read_excel(arquivo_excel, sheet_name=planilha_nome)
    except Exception as e:
        print(f"Erro ao ler o arquivo Excel: {e}")
        return
    
    # Verificar colunas
    if coluna_nome not in df.columns or coluna_cadeira not in df.columns:
        print(f"Colunas '{coluna_nome}' ou '{coluna_cadeira}' não encontradas")
        return
    
    # Criar dicionário de cadeiras
    cadeiras = {}
    for _, row in df.iterrows():
        cadeira = str(row[coluna_cadeira]).strip().upper()
        nome = str(row[coluna_nome]).strip()
        cadeiras[cadeira] = nome
    
    # Configurar figura com menos espaço no topo
    fig, ax = plt.subplots(figsize=(22, 28))  # Reduzi a altura total
    ax.set_xlim(0, 100)
    ax.set_ylim(0, 130)  # Reduzi o limite superior
    ax.axis('off')
    
    # Título com menos padding (reduzi de 30 para 10)
    plt.title("Mapa de Mesa - FNI", fontsize=24, pad=10, weight='bold')  # Ajuste aqui
    
    # Função para adicionar cadeira
    def adicionar_cadeira(x, y, numero, nome):
        # Retângulo da cadeira
        cadeira = patches.Rectangle((x, y), 6, 3, linewidth=1.5,
                                  edgecolor='#333333', facecolor='#EEEEEE')
        ax.add_patch(cadeira)
        
        # Número da cadeira
        plt.text(x + 3, y + 2.3, numero, ha='center', va='center', 
                fontsize=10, weight='bold', color='#333333')
        
        # Nome do participante
        plt.text(x + 3, y + 1.3, nome, ha='center', va='center',
                fontsize=9, wrap=True, color='#222222')

    # ========== POSIÇÕES DAS CADEIRAS ==========
    # Coordenadas bases (ajustadas para cima)
    x_esquerda = 25
    x_direita = 65
    y_topo = 120  # Mudei de 130 para 120 (subi as cadeiras)
    y_base = 10

    # ========== CADEIRAS SUPERIORES (C01-C05) ==========
    for i in range(5):
        cadeira_num = f"C{i+1:02d}"
        nome = cadeiras.get(cadeira_num, "")
        x_pos = x_esquerda + (i * (x_direita - x_esquerda)/4)
        adicionar_cadeira(x_pos, y_topo, cadeira_num, nome)
    
    # ========== CADEIRAS LATERAIS ESQUERDA (C11-C55) ==========
    for i in range(45):
        cadeira_num = f"C{11+i:02d}"
        nome = cadeiras.get(cadeira_num, "")
        adicionar_cadeira(x_esquerda, y_topo-5 - i*2.3, cadeira_num, nome)  # Ajuste no espaçamento
    
    # ========== CADEIRAS LATERAIS DIREITA (C56-C100) ==========
    for i in range(45):
        cadeira_num = f"C{56+i:02d}"
        nome = cadeiras.get(cadeira_num, "")
        adicionar_cadeira(x_direita, y_topo-5 - i*2.3, cadeira_num, nome)  # Ajuste no espaçamento
    
    # ========== CADEIRAS INFERIORES (C06-C10) ==========
    for i in range(5):
        cadeira_num = f"C{i+6:02d}"
        nome = cadeiras.get(cadeira_num, "")
        x_pos = x_esquerda + (i * (x_direita - x_esquerda)/4)
        adicionar_cadeira(x_pos, y_base, cadeira_num, nome)
    
    # Salvar imagem
    plt.tight_layout()
    plt.savefig(arquivo_saida, dpi=150, bbox_inches='tight', facecolor='#FAFAFA')
    print(f"Mapa gerado com sucesso: '{arquivo_saida}'")

if __name__ == "__main__":
    criar_mapa_mesa(
        arquivo_excel="participantes.xlsx",
        planilha_nome="Participantes",
        coluna_nome="Participantes",
        coluna_cadeira="Cadeiras",
        arquivo_saida="mapa_mesa_ajustado.png"
    )