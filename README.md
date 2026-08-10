# Hub Central de Automações MRV - Versão 2.0

Este projeto reúne as automações administrativas mais usadas no dia a dia em uma interface gráfica única para Windows. A versão 2.0 oferece execução mais prática dos robôs, progresso visual em tempo real e controle de cancelamento de processos.

## O que há de novo

- Aplicativo Windows executável (`dist/app_central.exe`).
- Interface central com botões para cada robô.
- Barra de progresso visual durante a execução.
- Console integrado para logs em tempo real.
- Botão de cancelamento para parar o processo ativo.
- Tela de configurações para salvar credenciais dos robôs.

## Objetivo

Automatizar tarefas repetitivas como:

- geração de rascunhos de e-mail;
- preparação de planilhas de rateio;
- lançamento de notas fiscais em portais internos;
- extração de relatórios de produtividade e correios;
- integração com sistemas como Podio, Agilis e SAP.

## Estrutura do projeto

- `app_central.py`: interface principal com navegação, painel de robôs, console e controle de execução.
- `config.py`: lógica de carregamento/salvamento de credenciais e definição de pastas do projeto.
- `config_exemplo.py`: referência do arquivo de configuração.
- `treinar_ia.py`: material de apoio com dicionários e mapeamentos usados internamente.
- `robos/`: scripts de automação organizados por tema.
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

O app salva as credenciais em `config_mrv.json` no diretório do projeto.

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

- **Configuração:** `config.py` — caminhos dinâmicos, variável `SENHA_MALOTE` e carregamento de `config_mrv.json` para credenciais.
- **Interface central:** `app_central.py` — integração dos robôs ao empacotamento PyInstaller, barra de progresso, console integrado e controle de cancelamento.
- **Malote / Rateio:** `robos/robo_rateio_malote.py` (RATEIO MALOTE v15) e `robos/malote_web_scraper.py` (integração/lookup dos Correios) — melhorias em heurísticas de extração e cache, correções de valores órfãos e tratamento de exceções.
- **Faturamento:** `robos/robo_faturamento.py` — geração de rascunhos no Outlook e nova função end-to-end para processo de e-mail → MRV Pag; extrações inteligentes de valores e rastreios.
- **Produtividade & Relatórios:** `robos/produtividade.py` e `robos/robo_relatorio_correios.py` — extração automatizada de relatórios (Podio/Agilis), controles de login SSO e exportação para Excel.
- **Mensageria / Inclusão:** `robos/robo_incluir_encomendas.py` e `robos/criar_rascunhos_uber.py` — criação de rascunhos no Outlook, preenchimento automático de campos no Podio e rotinas para arquivos do Uber.
- **Chamados / Agilis:** `robos/robo_fechar_chamados.py` — melhorias no fluxo de login SSO, tratamento de MFA, robustez nas interações do Agilis e verificação prévia de status via API, fecha o chamado de acordo com a categoria do mesmo.
- **Jurídico & Contratos:** `robos/robo_juridico.py`, `robos/robo_macro_contratos.py` e `robos/robo_cobrar_boleto.py` — extração e processamento de documentos, cobrança de boletos e follow-up automático de boletos em aberto.
- **Uber / SAP:** `robos/robo_uber_relatorios.py` — processamento de planilhas do Uber, mapeamento de responsáveis e atualizações com dados do SAP.
- **ZMM180 / OCR:** `robos/robo_zmm180.py` — automação com PyAutoGUI e suporte a Tesseract OCR (configuração e detecção de idiomas).
- **Mapeamentos IA:** `treinar_ia.py` — dicionários e mapeamentos (pessoas e origens) usados por vários robôs.

Observação: os itens acima foram extraídos dos scripts na pasta `robos/` e dos arquivos principais do projeto. Se você quiser que eu gere um changelog mais formal (por exemplo, com datas, autores e versões por arquivo), eu posso criar um `CHANGELOG.md` com entradas detalhadas.

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
