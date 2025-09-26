# Controle de Materiais de TI - SEMED Power BI

Este repositório contém um script Python para automatizar o gerenciamento de registros de equipamentos de TI na planilha "Controle de Materiais de TI - SEMED" no Google Sheets, integrado com o Power BI.

## Objetivo
O script `automacao_de_update.py` permite:
- Visualizar registros de equipamentos.
- Inserir novos registros com IDs automáticos baseados em categorias.
- Buscar e editar registros na planilha.
- Gerenciar prefixos de categorias de equipamentos.

## Pré-requisitos
- Python 3.6+
- Bibliotecas Python:
  ```bash
  pip install gspread google-auth tabulate
  ```
- Conta de serviço do Google Cloud com acesso à API do Google Sheets.
- Arquivo de credenciais (`credenciais.json`) gerado no Google Cloud Console.

## Configuração
1. **Clone o repositório:**
   ```bash
   git clone https://github.com/ZAQU3O/Materiais-de-TI---SEMED-Power-BI.git
   cd Materiais-de-TI---SEMED-Power-BI
   ```

2. **Configure as credenciais:**
   - Crie uma conta de serviço no Google Cloud Console (projeto: `planilhaequipamentos-472621`).
   - Baixe a chave JSON e salve como `credenciais.json` no diretório do projeto.
   - **Nota:** O arquivo `credenciais.json` está no `.gitignore` para segurança. Nunca o commite.

3. **Habilite as APIs:**
   - No Google Cloud Console, habilite as APIs do Google Sheets e Google Drive.

4. **Instale as dependências:**
   ```bash
   pip install -r requirements.txt
   ```
   (Crie um `requirements.txt` com: `gspread`, `google-auth`, `tabulate`.)

## Uso
1. Execute o script:
   ```bash
   python automacao_de_update.py
   ```
2. Escolha uma opção no menu interativo:
   - **1 - Visualizar registros:** Mostra todos os registros da planilha.
   - **2 - Inserir novo registro:** Adiciona um equipamento com ID gerado automaticamente.
   - **3 - Buscar e editar registros:** Busca por registros e permite edição.
   - **4 - Sair:** Encerra o programa.

## Estrutura do projeto
- `automacao_de_update.py`: Script principal para interação com a planilha.
- `prefixos.json`: Armazena prefixos aprendidos para categorias de equipamentos.
- `.gitignore`: Ignora `credenciais.json` e outros arquivos sensíveis.

## Notas
- **Segurança:** Mantenha `credenciais.json` fora do controle de versão. Regenere a chave se comprometida.
- **Integração com Power BI:** A planilha pode ser conectada ao Power BI para visualização de dados.
- **Colaboradores:** Após alterações no histórico (ex.: `git push --force`), atualize cópias locais:
  ```bash
  git fetch origin
  git reset --hard origin/main
  ```

## Licença
Este projeto é de uso interno da SEMED. Contate o proprietário do repositório para permissões.