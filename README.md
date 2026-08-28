## Estrutura de documentação (v3.1)

Este projeto reúne as automações administrativas mais usadas no dia a dia em uma interface gráfica única para Windows. A versão 3.0 traz uma interface renovada, novos fluxos e mais robustez na execução dos robôs.

A partir da versão 3.1, a documentação dos robôs passa a ser modular.

A pasta `ajuda_robos/` contém arquivos individuais com instruções específicas para cada automação.

Exemplo:

ajuda_robos/
├── rateio_malote.md
├── rateio_agf.md
├── produtividade.md
├── fechar_chamados.md
└── ...

Cada arquivo pode conter:

- objetivo do robô;
- sistemas necessários;
- arquivos de entrada;
- passo a passo de utilização;
- requisitos especiais;
- configurações obrigatórias;
- solução de problemas;
- observações importantes.


## O que há de novo

- Aplicativo Windows executável (`dist/app_central.exe`).
- Interface central renovada com sidebar e abas Início / Robôs / Configurações / Ajuda.
- Sistema de documentação preparado para a versão 3.1.
- Ajuda contextual por robô utilizando arquivos independentes.
- Separação entre ajuda geral do Hub e documentação específica de cada automação.
- Barra de progresso visual durante a execução.
- Console integrado para logs em tempo real.

## Objetivo

Automatizar tarefas repetitivas como:

- geração de rascunhos de e-mail;
- preparação de planilhas de rateio;
- lançamento de notas fiscais em portais internos;
- extração de relatórios de produtividade e correios;
- integração com sistemas como Podio, Agilis e SAP.

## Central de Ajuda

A aba Ajuda possui duas funções:

### Ajuda Geral

Documentação sobre:

- credenciais;
- configurações;
- APIs;
- boas práticas;
- dúvidas frequentes;
- utilização geral do Hub.

### Ajuda dos Robôs (v3.1)

Cada robô possui sua própria documentação contextual.

Ao clicar no botão de ajuda do robô, o Hub exibe instruções específicas daquela automação, incluindo:

- objetivo;
- requisitos;
- arquivos necessários;
- configurações obrigatórias;
- alertas importantes;
- solução de problemas comuns.

Essa documentação é carregada a partir dos arquivos localizados na pasta `ajuda_robos/`.

## Estrutura do projeto

- `app_central.py`: interface principal com navegação, painel de robôs, console e controle de execução.
- `config.py`: lógica de carregamento/salvamento de credenciais e definição de pastas do projeto.
- `config_exemplo.py`: referência do arquivo de configuração.
- `treinar_ia.py`: material de apoio com dicionários e mapeamentos usados internamente.
- `robos/`: scripts de automação organizados por tema.
- ajuda_robos/: documentação individual dos robôs utilizada pelo sistema de ajuda contextual.
- `dist/`: executável gerado para Windows.
- `build/`: artefatos gerados pelo PyInstaller.

## Pré-requisitos

- Windows.
- Google Chrome instalado.
- Outlook instalado e logado.
- ChromeDriver compatível com a versão do Chrome.
- Acesso aos sistemas utilizados pelos robôs (Podio, Agilis, portais internos, SAP, etc.).
- Para executar pelo código-fonte: Python 3.x e dependências necessárias (`customtkinter`, bibliotecas de automação web, Excel, etc.).

## Configuração de credenciais

O app salva as credenciais em `config_mrv.json` no diretório do projeto quando executado pelo código-fonte, e em `dist/config_mrv.json` quando executado como o executável (`dist\app_central.exe`).

### Primeira execução

1. Execute o aplicativo:

```bash
python app_central.py
```

ou abra o executável:

```bash
dist\app_central.exe
```

2. Vá para a aba `⚙️ Configurações`.
3. Preencha:
   - E-mail MRV;
   - Senha MRV;
   - Senha Malote Web (se aplicável).
4. Clique em `Salvar Credenciais`.

> Importante: mantenha `config_mrv.json` seguro e não compartilhe suas credenciais publicamente.

## Como usar

### Executar o executável

Abra:

```bash
dist\app_central.exe
```

### Executar pelo código-fonte

No diretório do projeto:

```bash
python app_central.py
```

## Robôs disponíveis

### Correios & Faturamento

- Rateio de Malote (Centros de Custo)
- Faturamento 1: Gerar rascunhos
- Faturamento Completo: processo de e-mail para MRV Pag
- Cobrança de boletos de contratos: follow-up automático por rascunho no Outlook após X dias sem retorno

### Podio & Mensageria

- Relatório Jurídico Montreal
- Incluir correspondências rápidas

### Agilis & Chamados

- Gerar relatório de envio para Correios
- Gerar produtividade (Podio/Agilis/SAP)
- Fechar chamados a vencer

### Uber / SAP / Outros

- Uber 1: Atualizar responsáveis (SAP)
- Uber 2: Gerar relatórios e pastas
- Uber 3: Criar rascunhos de e-mail
- Faturamento Transação ZMM180

## Modo de execução e controle

- A barra de progresso indica o andamento do processo atual.
- O console integrado mostra logs e mensagens em tempo real.
- O botão `CANCELAR PROCESSO ATIVO` interrompe o processo em execução e ajuda a recuperar a interface.

## Observações importantes

## Atualizações (Resumo de mudanças recentes)

- **Configuração:** `config.py` — carregamento e salvamento de credenciais em `config_mrv.json`, caminhos dinâmicos para `Downloads` e `produtividade`, e variáveis de compatibilidade para os robôs.
- **Interface central:** `app_central.py` — nova interface com sidebar Início/Robôs/Config/Ajuda, botões para cada fluxo, barra de progresso visual, console em tempo real, botão de cancelamento de processo e import hooks para empacotamento com PyInstaller.
- **Novos fluxos:** inclusão de `Rateio AGF` e `Atualizar Macro de Contratos` na interface principal.
- **Malote / Correios:** `robos/robo_rateio_malote.py` e `robos/malote_web_scraper.py` — melhorias no fluxo de rateio de malote e integração de consulta/lookup dos Correios.
- **Faturamento:** `robos/robo_faturamento.py` — gerar rascunhos no Outlook e executar faturamento completo para MRV Pag.
- **Cobrança de boletos:** `robos/robo_cobrar_boleto.py` — follow-up automático por rascunho no Outlook para boletos em aberto.
- **Produtividade & Relatórios:** `robos/produtividade.py` e `robos/robo_relatorio_correios.py` — extração de relatórios Podio/Agilis/SAP e geração de exportações.
- **Mensageria & Inclusão:** `robos/robo_incluir_encomendas.py` e `robos/criar_rascunhos_uber.py` — criação de correspondências rápidas e rascunhos de e-mail para Uber.
- **Chamados / Agilis:** `robos/robo_fechar_chamados.py` — fechamento de chamados a vencer com maior robustez no login e no fluxo de status.
- **Jurídico & Contratos:** `robos/robo_juridico.py` e `robos/robo_macro_contratos.py` — processamento de documentos jurídicos e atualização de macros de contratos.
- **Uber / SAP:** `robos/robo_uber_relatorios.py` — geração de relatórios e pastas, além de atualização de responsáveis com base em dados SAP.
- **ZMM180 / OCR:** `robos/robo_zmm180.py` — automação com PyAutoGUI e suporte a OCR para processos ZMM180.
- **Mapeamentos IA:** `treinar_ia.py` — dicionários e mapeamentos usados pelos robôs.

Observação: itens acima refletem os scripts e a interface atual do projeto. Se quiser, posso também criar um `CHANGELOG.md` com entradas detalhadas, datas e versões por arquivo.

- Alguns robôs exigem arquivos de entrada específicos dentro da pasta `arquivos`.
- Verifique se as planilhas e PDFs estão nomeados corretamente antes de iniciar.
- Robôs que acessam portais podem pedir autorização de MFA no celular.
- Para processos SAP, deixe o sistema aberto e não mexa no mouse/teclado enquanto a automação estiver em execução.

## Boas práticas

- Mantenha os arquivos de entrada organizados.
- Confirme o nome das planilhas antes de executar.
- Revise os resultados após cada processo.
- Teste os fluxos em ambiente controlado antes de usar dados críticos.

## Resumo

Este repositório é um hub central de automação administrativa para Windows, com uma interface gráfica que reúne os principais robôs em um único ponto de controle.
