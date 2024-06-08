# Envio Automático de Relatório de Vendas Mensal

Este projeto em Python automatiza o envio de relatórios de vendas por e-mail no final de cada mês. Ele utiliza a biblioteca `pandas` para manipulação de dados e `win32com.client` para integração com o Microsoft Outlook. O script é projetado para ser executado automaticamente através do Agendador de Tarefas do Windows.

## Funcionalidades

- **Importação de Dados:** Importa dados de vendas a partir de um arquivo Excel (`Vendas.xlsx`).
- **Análise de Dados:** Calcula o faturamento, quantidade vendida e ticket médio por loja.
- **Geração de Relatório:** Cria um relatório HTML detalhado com os resultados da análise.
- **Envio de E-mail:** Envia automaticamente o relatório gerado para um endereço de e-mail especificado.

## Como Usar
Preparação do Script Python
Salvar o Script:
Copie o código Python fornecido anteriormente e salve-o em um arquivo chamado envia_relatorio.py em um diretório de sua preferência. Por exemplo: C:\Scripts\envia_relatorio.py.
Configuração do Agendador de Tarefas do Windows
Abrir o Agendador de Tarefas:

Pressione Win + R, digite taskschd.msc e pressione Enter.
Criar uma Nova Tarefa Básica:

No painel à direita, clique em "Criar Tarefa Básica".
Dê um nome para a tarefa, como "Enviar Relatório de Vendas", e clique em "Avançar".
Definir a Frequência:

Selecione "Mensalmente" e clique em "Avançar".
Configurar o Mês e o Dia:


#### Nota Importante
Certifique-se de que o Outlook esteja configurado corretamente e que você tenha permissão para enviar e-mails programaticamente usando o Outlook. O script depende do Outlook estar configurado na máquina onde ele está sendo executado.


 
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
