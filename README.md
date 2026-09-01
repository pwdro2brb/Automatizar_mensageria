# Hub Central de Automações MRV - Versão 3.1

Este projeto reúne as automações administrativas mais utilizadas no dia a dia em uma interface gráfica única para Windows.

Este projeto reúne as automações administrativas mais usadas no dia a dia em uma interface gráfica única para Windows. A versão 3.0 traz uma interface renovada, novos fluxos e mais robustez na execução dos robôs.

A versão 3.1 amplia a Central de Ajuda, adiciona documentação contextual para cada robô e incorpora os novos processos de Rateio Uber Central e Rateio Uber Tradicional.

## Estrutura de documentação da versão 3.1

A pasta `ajuda_robos/` contém arquivos Markdown individuais com instruções específicas para cada automação.

Exemplo:

```text
ajuda_robos/
├── criar_rascunhos_uber.md
├── produtividade.md
├── robo_faturamento.md
├── robo_rateio_malote.md
├── robo_rateio_uber_central.md
├── robo_rateio_uber_tradicional.md
└── ...
```

Cada documento pode apresentar:

- objetivo do robô;
- sistemas necessários;
- arquivos de entrada;
- passo a passo de utilização;
- requisitos especiais;
- configurações obrigatórias;
- resultados esperados;
- pontos de conferência;
- solução de problemas;
- observações importantes.


## O que há de novo na versão 3.1

- Ajuda contextual específica para cada robô.
- Documentação modular por meio da pasta `ajuda_robos/`.
- Botão de ajuda disponível nos cards das automações.
- Separação entre a ajuda geral do Hub e as instruções específicas dos robôs.
- Novo robô Rateio Uber Tradicional.
- Novo robô Rateio Uber Central.
- Campos específicos para e-mail e senha da Uber na tela de Configurações.
- Login compartilhado para os processos do Uber Business.
- Suporte às diferentes etapas de autenticação da Uber.
- Suporte à autenticação por código SMS.
- Seleção automática da organização correspondente.
- Download automático do relatório CSV e da Nota de Débito em PDF.
- Busca dinâmica da Base de Centro de Custo mais recente.
- Validação automática dos centros de custo.
- Geração automática da planilha de rateio.
- Criação de rascunhos no Outlook quando são encontradas viagens que precisam de análise.
- Barra de progresso visual durante a execução.
- Console integrado para logs em tempo real.
- Botão de cancelamento do processo ativo.

## Objetivo

Automatizar tarefas repetitivas como:

- geração de rascunhos de e-mail;
- preparação de planilhas de rateio;
- lançamento de notas fiscais em portais internos;
- extração de relatórios de produtividade e correios;
- integração com sistemas como Podio, Agilis e SAP.

## Central de Ajuda

A Central de Ajuda possui duas funções principais.

### Ajuda geral

A opção `Ajuda`, disponível na barra lateral, apresenta orientações gerais sobre:

- primeira utilização do Hub;
- credenciais;
- configurações;
- APIs;
- boas práticas;
- cancelamento de processos;
- dúvidas frequentes;
- utilização geral do aplicativo.

### Ajuda específica dos robôs

Cada card de automação possui um botão `❓`.

Ao clicar nesse botão, o Hub abre uma janela com a documentação específica do robô selecionado.

A documentação pode incluir:

- objetivo do processo;
- requisitos;
- arquivos necessários;
- configurações obrigatórias;
- passo a passo;
- alertas importantes;
- resultados esperados;
- conferências necessárias;
- solução de problemas comuns.

Os documentos são carregados a partir da pasta `ajuda_robos/`.

Alguns arquivos de ajuda podem ser compartilhados entre mais de uma opção do Hub. Nesses casos, o aplicativo abre automaticamente

## Estrutura do projeto

- `app_central.py`: interface principal com navegação, painel de robôs, console, ajuda contextual e controle de execução.
- `config.py`: carregamento e salvamento de credenciais e definição de caminhos do projeto.
- `config_exemplo.py`: referência da estrutura de configuração.
- `treinar_ia.py`: dicionários e mapeamentos utilizados internamente.
- `robos/`: scripts de automação organizados por processo.
- `robos/uber_login.py`: fluxo compartilhado de autenticação no Uber Business.
- `robos/uber_downloads.py`: navegação e download compartilhados entre os rateios Uber.
- `ajuda_robos/`: documentação individual utilizada pelo sistema de ajuda contextual.
- `arquivos/`: arquivos locais utilizados pelos processos.
- `dist/`: executável gerado para Windows.
- `build/`: artefatos temporários gerados pelo PyInstaller.

## Pré-requisitos

- Windows.
- Google Chrome instalado.
- Outlook instalado e logado.
- ChromeDriver compatível com a versão do Chrome.
- Acesso aos sistemas utilizados pelos robôs (Podio, Agilis, portais internos, SAP, etc.).
- Para executar pelo código-fonte: Python 3.x e dependências necessárias (`customtkinter`, bibliotecas de automação web, Excel, etc.).

## Configuração de credenciais

O aplicativo salva as credenciais no arquivo `config_mrv.json`.

Quando executado pelo código-fonte, o arquivo fica no diretório do projeto. Quando executado pelo executável, o arquivo fica no diretório `dist`.

> Importante: o arquivo `config_mrv.json` pode conter senhas e chaves de acesso. Mantenha o arquivo seguro, não o compartilhe e não o inclua no repositório.

### Primeira execução

1. Execute o aplicativo pelo código-fonte:

```bash
python app_central.py
```

Ou abra o executável:

```text
dist\app_central.exe
```

2. Acesse a aba `⚙️ Configurações`.

3. Preencha os campos aplicáveis:

- E-mail MRV;
- Senha MRV;
- Senha Malote Web;
- Chave API Agilis;
- credenciais dos Correios;
- credenciais da API do Podio;
- E-mail Uber;
- Senha Uber.

4. Clique em `Salvar Todas as Credenciais`.

O código SMS da Uber não é armazenado. Quando solicitado, o código precisa ser informado manualmente no navegador.

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

- Rateio de Malote por centro de custo.
- Rateio AGF.
- Faturamento 1: Gerar Rascunhos.
- Faturamento 2: Processo Completo.
- Cobrança de boletos.

### Podio & Mensageria

- Relatório Jurídico Montreal.
- Incluir Correspondências Rápidas.

### Agilis & Chamados

- Gerar Relatório de Envio para os Correios.
- Gerar Produtividade.
- Fechar Chamados a Vencer.

### Uber / SAP / Contratos

- Rateio Uber Central.
- Rateio Uber Tradicional.
- Atualizar Macro de Contratos.
- Uber 1: Atualizar Responsáveis.
- Uber 2: Gerar Relatórios e Pastas.
- Uber 3: Criar Rascunhos de E-mail.
- Faturamento pela transação ZMM180.

## Modo de execução e controle

- A barra de progresso indica o andamento do processo atual.
- O console integrado mostra logs e mensagens em tempo real.
- O botão `CANCELAR PROCESSO ATIVO` interrompe o processo em execução e ajuda a recuperar a interface.

## Observações importantes

## Atualizações da versão 3.1

- **Ajuda contextual:** inclusão de documentação específica para cada automação por meio de arquivos Markdown.
- **Interface:** substituição do indicador de prioridade pelo botão de ajuda contextual.
- **Documentação modular:** criação da pasta `ajuda_robos/`.
- **Rateio Uber Central:** novo processo de download, validação de centros de custo, geração de rateio e criação de rascunhos.
- **Rateio Uber Tradicional:** novo processo de download, validação de centros de custo, geração de rateio e criação de rascunhos.
- **Autenticação Uber:** criação de fluxo compartilhado com suporte a SMS, senha e sessão previamente autenticada.
- **Downloads Uber:** criação de rotina compartilhada para download do CSV e da Nota de Débito.
- **Credenciais Uber:** inclusão de campos específicos para e-mail e senha na tela de Configurações.
- **Base de Centro de Custo:** localização automática do arquivo mais recente que contenha `BASE CENTRO DE CUSTO` no nome.
- **Empacotamento:** inclusão dos novos robôs e módulos compartilhados nos imports utilizados pelo PyInstaller.

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
