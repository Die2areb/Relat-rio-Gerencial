# Envio Automático de Relatório de Vendas Mensal

Este projeto em Python automatiza o envio de relatórios de vendas por e-mail no final de cada mês. Ele utiliza a biblioteca `pandas` para manipulação de dados e `win32com.client` para integração com o Microsoft Outlook. O script é projetado para ser executado automaticamente através do Agendador de Tarefas do Windows.

## Funcionalidades

- **Importação de Dados:** Importa dados de vendas a partir de um arquivo Excel (`Vendas.xlsx`).
- **Análise de Dados:** Calcula o faturamento, quantidade vendida e ticket médio por loja.
- **Geração de Relatório:** Cria um relatório HTML detalhado com os resultados da análise.
- **Envio de E-mail:** Envia automaticamente o relatório gerado para um endereço de e-mail especificado.

## Como Usar

### Pré-requisitos

- Python 3.x
- Biblioteca `pandas`
- Biblioteca `pywin32`
- Microsoft Outlook instalado e configurado

### Instalação

1. Clone o repositório:
   ```sh
   git clone https://github.com/seu_usuario/envio-relatorio-vendas-mensal.git
   cd envio-relatorio-vendas-mensal
