# Automatizar Mensageria - Versão 2.0

Este projeto agora funciona como um aplicativo executável para Windows, reunindo em uma única interface gráfica as automações administrativas mais usadas no dia a dia, com foco em faturamento, mensageria, Podio, Correios e relatórios operacionais.

A versão 2.0 foi transformada em um executável pronto para uso, facilitando a abertura e execução sem depender diretamente de um ambiente Python configurado na máquina.

## O que mudou na versão 2.0

- O projeto passou a ser distribuído como executável Windows (.exe).
- A interface central continua disponível para execução dos robôs a partir de botões na tela.
- O fluxo de uso ficou mais simples para usuários finais.
- O aplicativo mantém o mesmo funcionamento das automações, agora com execução mais prática.

## Objetivo

Automatizar tarefas repetitivas e manuais, como:

- geração de rascunhos no Outlook;
- criação de planilhas de rateio;
- lançamento de notas fiscais em portais internos;
- geração de relatórios de produtividade e encomendas;
- integração com sistemas como Podio, Agilis e plataformas internas.

## Estrutura do projeto

- app_central.py: interface principal em Tkinter com os botões de execução e o controle dos robôs.
- config.py: configurações locais, como e-mail, senha e caminhos de pasta.
- config_exemplo.py: modelo de configuração para ser copiado para o arquivo real.
- treinar_ia.py: referência com dicionários e mapeamentos usados para classificação/treinamento interno.
- robos/: scripts com as automações organizadas por tema.
- dist/app_central.exe: executável gerado para uso no Windows.

## Pré-requisitos

Antes de executar, certifique-se de que o ambiente atende aos itens abaixo:

- Windows
- Google Chrome instalado
- Outlook instalado e logado
- ChromeDriver compatível com a versão do Chrome
- Acesso aos sistemas usados pelos robôs (Podio, Agilis, portal interno, etc.)

## Configuração

1. Copie o arquivo de exemplo para o arquivo real:

```bash
copy config_exemplo.py config.py
```

2. Edite o arquivo config.py com:

- e-mail corporativo;
- senha ou credenciais de acesso;
- caminhos das pastas utilizadas pelo projeto.

> Importante: mantenha o arquivo config.py com as credenciais seguras e não compartilhe esse conteúdo publicamente.

## Como usar a versão 2.0

### Opção 1: executar o executável

Na pasta dist do projeto, execute:

```bash
dist\app_central.exe
```

A interface gráfica será aberta com os processos disponíveis. Ao clicar em um botão, o robô correspondente será executado.

### Opção 2: executar a versão em Python

Se quiser rodar diretamente pelo código-fonte:

```bash
python app_central.py
```

## Cancelar processo ativo

A central oferece suporte para cancelamento de processos em execução. Um botão vermelho chamado "🛑 CANCELAR PROCESSO ATIVO" fica localizado na parte superior da interface e se habilita automaticamente quando qualquer automação estiver em execução.

Características:
- funciona com todos os robôs;
- exibe confirmação antes de interromper;
- encerra o processo ativo e janelas abertas pelo robô;
- é útil para parada de emergência ou ajustes rápidos.

## Novas funcionalidades

- motor universal de processos canceláveis;
- melhor tratamento de logs e erros na interface;
- novo fluxo de Rateio de Malote com leitura de dados e consolidação automática;
- novos scripts adicionados ao repositório para diferentes cenários operacionais.

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
- Gerar Produtividade (Podio/Agilis/SAP)
- Fechar chamados a vencer

### Outros sistemas

- Relatório de utilização Uber
- Faturamento transação ZMM180

## Observações importantes

- Alguns processos exigem arquivos específicos nas pastas de entrada, como planilhas Excel e boletos em PDF.
- O fluxo de faturamento depende de arquivos organizados dentro da pasta arquivos/faturamento.
- Alguns robôs usam autenticação com MFA, então pode ser necessário aprovar a ação manualmente no celular.
- Em caso de erro, verifique o console da interface, pois os logs são exibidos ali durante a execução.

## Boas práticas

- mantenha os arquivos de entrada bem organizados;
- confira o nome das planilhas antes de executar;
- revise os resultados após cada processo;
- teste os fluxos em ambiente controlado antes de usar com dados críticos.

## Resumo

Este repositório funciona como um hub central de automação administrativa, reunindo em um único ponto de acesso os processos mais comuns para execução rápida e padronizada, agora com uma versão 2.0 executável para Windows.
