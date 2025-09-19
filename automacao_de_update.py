import gspread
from google.oauth2.service_account import Credentials
from tabulate import tabulate
import re
import json
import os

ARQUIVO_PREFIXOS = "prefixos.json"

prefixos_padrao = {
    "TECLADO": "T",
    "MOUSE": "M",
    "MONITOR": "MN",
    "IMPRESSORA": "I"
}

def carregar_prefixos():
    if os.path.exists(ARQUIVO_PREFIXOS):
        with open(ARQUIVO_PREFIXOS, "r") as f:
            dados = json.load(f)
            return {**prefixos_padrao, **dados}
    return prefixos_padrao.copy()

def salvar_prefixos(prefixos):
    aprendidos = {k:v for k,v in prefixos.items() if k not in prefixos_padrao}
    with open(ARQUIVO_PREFIXOS, "w") as f:
        json.dump(aprendidos, f)

def conectar_planilha(credenciais_json):
    nome_planilha = "Controle de Materiais de TI - SEMED"
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

colunas = [
    "Registro",
    "Descrição do Equipamento",
    "Marca/ Modelo",
    "Numero de Serie",
    "Data da Aquisição",
    "Condição",
    "Local Instalado",
    "Detalhes",
    "Ainda no Local de Instalação",
    "Novo Local de Instalação",
    "Data de Saida"
]

def exibir_colunas():
    print("\nColunas disponíveis:")
    for nome in colunas:
        print(" -", nome)

def visualizar_registros(planilha):
    dados = planilha.get_all_records()
    if not dados:
        print("📭 Nenhum registro encontrado.")
        return
    print("\n📋 Registros atuais:")
    print(tabulate(dados, headers="keys", tablefmt="grid"))

def gerar_id_categoria(planilha, prefixo):
    dados = planilha.get_all_records()
    ids_categoria = [linha["Registro"] for linha in dados if linha["Registro"].startswith(prefixo)]
    numeros = [int(re.sub(r'\D', '', i)) for i in ids_categoria]
    proximo = max(numeros)+1 if numeros else 1
    return f"{prefixo}{proximo:04d}"

def determinar_prefixo(descricao, prefixos):
    descricao = descricao.upper()
    # Busca por correspondência parcial ou palavras-chave
    for chave, val in prefixos.items():
        if chave in descricao or any(p in descricao for p in chave.split()):
            return val
    # Se não encontrou, cria automático (primeiras letras das palavras)
    palavras = [c for c in descricao if c.isalpha()]
    novo_prefixo = "".join(palavras[:2]).upper() if len(palavras)>=2 else "".join(palavras).upper()
    prefixos[descricao] = novo_prefixo
    salvar_prefixos(prefixos)
    print(f"🔹 Novo prefixo aprendido: '{descricao}' → '{novo_prefixo}'")
    return novo_prefixo

def inserir_registro(planilha, prefixos):
    print("\n📌 Inserir novo registro:")
    descricao = input("Descrição do equipamento: ").strip()
    prefixo = determinar_prefixo(descricao, prefixos)
    novo_id = gerar_id_categoria(planilha, prefixo)
    
    # Sugestão automática
    print(f"💡 Sugestão automática: ID = {novo_id}, Prefixo = {prefixo}")
    confirmar = input("Deseja usar essa sugestão? (S/n): ").strip().lower()
    if confirmar == "n":
        prefixo = input("Digite o prefixo desejado: ").upper()
        novo_id = gerar_id_categoria(planilha, prefixo)

    nova_linha = [novo_id, descricao]
    for coluna in colunas[2:]:
        valor = input(f"{coluna}: ").strip()
        nova_linha.append(valor)
    planilha.append_row(nova_linha)
    print("✅ Registro inserido com sucesso!")

def buscar_registros(planilha):
    print("\n🔍 Buscar registros")
    exibir_colunas()
    coluna_escolhida = input("Digite o nome da coluna para buscar: ").strip()
    if coluna_escolhida not in colunas:
        print("❌ Coluna inválida.")
        return []
    valor_busca = input(f"Digite o valor a buscar na coluna '{coluna_escolhida}': ").strip()
    dados = planilha.get_all_records()
    encontrados = [linha for linha in dados if str(linha[coluna_escolhida]) == valor_busca]
    if not encontrados:
        print("⚠ Nenhum registro encontrado.")
        return []
    print("\n🔹 Resultados encontrados:")
    print(tabulate(encontrados, headers="keys", tablefmt="grid"))
    return encontrados

def editar_registro(planilha):
    encontrados = buscar_registros(planilha)
    if not encontrados:
        return
    registro_alvo = input("\nDigite o Registro que deseja editar: ").strip()
    try:
        cel = planilha.find(registro_alvo, in_column=1)
        linha = cel.row
    except:
        print("❌ Registro não encontrado.")
        return
    while True:
        print("\nDigite a coluna que deseja alterar ou 'sair' para finalizar:")
        exibir_colunas()
        coluna_editar = input("Coluna: ").strip()
        if coluna_editar.lower() == "sair":
            break
        if coluna_editar not in colunas:
            print("❌ Coluna inválida.")
            continue
        novo_valor = input(f"Novo valor para {coluna_editar}: ").strip()
        planilha.update_cell(linha, colunas.index(coluna_editar)+1, novo_valor)
        print(f"✅ Atualizado ({coluna_editar} → {novo_valor})")

def menu(planilha, prefixos):
    while True:
        print("\n===== MENU =====")
        print("1 - Visualizar registros")
        print("2 - Inserir novo registro")
        print("3 - Buscar e editar registros")
        print("4 - Sair")
        escolha = input("Escolha uma opção: ").strip()
        if escolha == "1":
            visualizar_registros(planilha)
        elif escolha == "2":
            inserir_registro(planilha, prefixos)
        elif escolha == "3":
            editar_registro(planilha)
        elif escolha == "4":
            print("Encerrando.")
            break
        else:
            print("Opção inválida. Tente novamente.")

def main():
    credenciais_json = "credenciais.json"
    planilha = conectar_planilha(credenciais_json)
    prefixos = carregar_prefixos()
    print("🔗 Conectado à planilha com sucesso!")
    menu(planilha, prefixos)

if __name__ == "__main__":
    main()
