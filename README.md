# Automatizar Mensageria

Este projeto reúne uma central de automações para processos administrativos, com foco em faturamento, mensageria, Podio, Correios e relatórios operacionais.

A solução foi organizada em uma interface central, onde é possível executar diferentes robôs a partir de botões em uma tela gráfica.

## Objetivo

Automatizar tarefas repetitivas e manuais, como:

- geração de rascunhos no Outlook;
- criação de planilhas de rateio;
- lançamento de notas fiscais em portais internos;
- geração de relatórios de produtividade e encomendas;
- integração com sistemas como Podio, Agilis e plataformas internas.

## Estrutura do projeto

- app_central.py: interface principal em Tkinter com os botões de execução.
- robo_faturamento.py: automações relacionadas a faturamento e correios.
- Processos_simples/robo_juridico.py: automação para o processo jurídico no Podio.
- Processos_simples/robo_relatorio_correios.py: automação para gerar/baixar relatórios de produtividade no Agilis.
- Arquivos_rateio_regionais/: pasta utilizada como base para arquivos de rateio e exemplos de entrada/saída.
- config.py: configurações locais, como e-mail, senha e caminhos de pasta.
- config_exemplo.py: modelo de configuração para ser copiado para o arquivo real.
- produtividade.py: automação de extração de dados de produtividade (Podio, Agilis, Bússola/SAP) e geração de relatório preenchido.

## Pré-requisitos

Antes de executar, certifique-se de que o ambiente atende aos itens abaixo:

- Windows
- Python 3.10 ou superior
- Google Chrome instalado
- ChromeDriver compatível com a versão do Chrome
- Outlook instalado e logado
- Acesso aos sistemas usados pelos robôs (Podio, Agilis, portal interno, etc.)

### Bibliotecas Python necessárias

Instale as dependências abaixo, se ainda não estiverem disponíveis:

```bash
pip install pandas openpyxl PyPDF2 selenium pywinauto pywin32 holidays
```

## Configuração

1. Copie o arquivo de exemplo para o arquivo real:

```bash
copy config_exemplo.py config.py
```

2. Edite o arquivo config.py com:

- e-mail corporativo;
- senha ou credenciais de acesso;
- caminhos das pastas utilizadas pelo projeto, incluindo a nova `PASTA_PRODUTIVIDADE` para relatório de produtividade.

> Importante: mantenha o arquivo config.py com as credenciais seguras e não compartilhe esse conteúdo publicamente.

## Como executar

Na pasta do projeto, execute:

```bash
python app_central.py
```

A interface gráfica será aberta com os processos disponíveis. Ao clicar em um botão, o robô correspondente será executado em segundo plano.

### Cancelar processo ativo

A central agora oferece suporte universal de cancelamento para **todos os processos**. Um botão vermelho chamado "🛑 CANCELAR PROCESSO ATIVO" fica localizado na parte superior da interface e se habilita automaticamente quando qualquer automação está em execução.

**Características:**
- Funciona com todos os robôs (Faturamento, Produtividade, Jurídico, Correios, etc.)
- Ao clicar, exibe uma confirmação antes de interromper
- Encerra o processo ativo e todos os navegadores/janelas abertos por ele
- Ideal para parada de emergência ou ajustes rápidos
- Pode interromper o fluxo antes da conclusão completa

Cada processo é executado em um subprocesso isolado, garantindo que o cancelamento funcione de forma segura e responsiva, sem travar a interface central.

## Processos disponíveis

### Correios e Faturamento

- Relatório de Encomendas do Dia
- Rateio de Malote
- Faturamento 1: gerar rascunhos
- Faturamento 2: gerar planilha de rateio
- Faturamento 3: lançar NF no portal

### Podio e Mensageria

- Relatório Jurídico Montreal
- Inclusão rápida de correspondências no Podio

### Agilis e produtividade

- Gerar relatório de envio para Correios
- Gerar Produtividade (Podio/Agilis/SAP) — executa extração de dados e preenche o relatório final na pasta `Produtividade`
- Fechar chamados a vencer

### Outros sistemas

- Relatório de utilização Uber
- Faturamento transação ZMM180

## Observações importantes

- Alguns processos exigem arquivos específicos nas pastas de entrada, como planilhas Excel e boletos em PDF.
- O fluxo de faturamento depende de arquivos organizados dentro da pasta Arquivos_rateio_regionais.
- Alguns robôs usam autenticação com MFA, então pode ser necessário aprovar a ação manualmente no celular.
- Em caso de erro, verifique o console da interface, pois os logs são exibidos ali durante a execução.

## Boas práticas

- mantenha os arquivos de entrada bem organizados;
- confira o nome das planilhas antes de executar;
- revise os resultados após cada processo;
- teste os fluxos em ambiente controlado antes de usar com dados críticos.

## Resumo

Este repositório funciona como um hub central de automação administrativa, reunindo em um único ponto de acesso os processos mais comuns para execução rápida e padronizada.
