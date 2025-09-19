import gspread
from google.oauth2.service_account import Credentials

# Conectar à planilha
def conectar_planilha(credenciais_json):
    nome_planilha = "Controle de Materiais de TI - SEMED"  # nome correto da sua planilha
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/spreadsheets",
        "https://www.googleapis.com/auth/drive.file",
        "https://www.googleapis.com/auth/drive"
    ]
    creds = Credentials.from_service_account_file(credenciais_json, scopes=scope)
    cliente = gspread.authorize(creds)
    planilha = cliente.open(nome_planilha).sheet1
    return planilha

# Mapeamento de colunas
colunas = {
    "Registro": 1,
    "Descrição do Equipamento": 2,
    "Marca/ Modelo": 3,
    "Numero de Serie": 4,
    "Data da Aquisição": 5,
    "Condição": 6,
    "Local Instalado": 7,
    "Detalhes": 8,
    "Ainda no Local de Instalação": 9,
    "Novo Local de Instalação": 10,
    "Data de Saida": 11
}

def exibir_colunas():
    print("\nColunas disponíveis:")
    for nome in colunas.keys():
        print(" -", nome)

# Visualizar registros
def visualizar_registros(planilha):
    dados = planilha.get_all_records()
    if not dados:
        print("📭 Nenhum registro encontrado.")
        return
    print("\n📋 Registros atuais:")
    for linha in dados:
        print(linha)

# Alterar registros existentes
def alterar_registro(planilha):
    registro_alvo = input("\nDigite o valor do Registro que quer editar: ").strip()
    try:
        cel = planilha.find(registro_alvo, in_column=colunas["Registro"])
        linha = cel.row
    except:
        print("❌ Registro não encontrado.")
        return
    exibir_colunas()
    coluna_escolhida = input("Digite o nome exato da coluna que deseja alterar: ").strip()
    if coluna_escolhida not in colunas:
        print("❌ Coluna inválida.")
        return
    novo_valor = input(f"Digite o novo valor para {coluna_escolhida}: ").strip()
    planilha.update_cell(linha, colunas[coluna_escolhida], novo_valor)
    print(f"✅ Registro '{registro_alvo}' atualizado ({coluna_escolhida} → {novo_valor})")

# Inserir novo registro
def inserir_registro(planilha):
    print("\n📌 Inserir novo registro:")
    nova_linha = []
    for coluna in colunas.keys():
        valor = input(f"{coluna}: ").strip()
        nova_linha.append(valor)
    planilha.append_row(nova_linha)
    print("✅ Novo registro inserido com sucesso!")

# Buscar registros por coluna
def buscar_registros(planilha):
    exibir_colunas()
    coluna_escolhida = input("Digite o nome da coluna para buscar: ").strip()
    if coluna_escolhida not in colunas:
        print("❌ Coluna inválida.")
        return
    valor_busca = input(f"Digite o valor a buscar na coluna '{coluna_escolhida}': ").strip()
    dados = planilha.get_all_records()
    encontrados = [linha for linha in dados if str(linha[coluna_escolhida]) == valor_busca]
    if not encontrados:
        print("⚠ Nenhum registro encontrado.")
        return
    print("\n🔍 Resultados encontrados:")
    for linha in encontrados:
        print(linha)

# Menu interativo
def menu(planilha):
    while True:
        print("\n===== MENU =====")
        print("1 - Visualizar registros")
        print("2 - Alterar registro existente")
        print("3 - Inserir novo registro")
        print("4 - Buscar registros")
        print("5 - Sair")
        escolha = input("Escolha uma opção: ").strip()
        if escolha == "1":
            visualizar_registros(planilha)
        elif escolha == "2":
            alterar_registro(planilha)
        elif escolha == "3":
            inserir_registro(planilha)
        elif escolha == "4":
            buscar_registros(planilha)
        elif escolha == "5":
            print("Encerrando.")
            break
        else:
            print("Opção inválida. Tente novamente.")

# Principal
def main():
    credenciais_json = "credenciais.json"  # arquivo das credenciais do Google Cloud
    planilha = conectar_planilha(credenciais_json)
    print("🔗 Conectado à planilha com sucesso!")
    menu(planilha)

if __name__ == "__main__":
    main()
