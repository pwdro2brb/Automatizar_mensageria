# Rateio Uber Central

## Objetivo

Este robô realiza o processamento mensal do Uber Central.

A automação acessa o Uber Business, baixa o relatório de transações em formato CSV e a Nota de Débito em PDF, valida os centros de custo e gera a planilha final de rateio.

Caso sejam encontradas viagens duplicadas que precisam de análise, o robô também cria um rascunho de e-mail no Outlook.

## 🛑 Passo 1: Verificação de Acesso

Antes de executar o robô, confirme se você possui cadastro como **Administrador** nas plataformas Uber Tradicional e Uber Central.

1. Entre em contato com a equipe da Mensageria e solicite a verificação do seu e-mail corporativo dentro da plataforma.
2. Caso seu acesso já esteja liberado como Administrador, prossiga com a execução.
3. Caso ainda não tenha acesso, solicite a liberação.

## ⚙️ Pré-requisitos

Antes de executar o processo, confirme os seguintes requisitos:

### Acesso à rede

O computador precisa estar conectado à rede corporativa para que o robô consiga acessar as pastas públicas da empresa.

### Base de Centro de Custo atualizada

A pasta dos Correios deve possuir um arquivo Excel contendo `BASE CENTRO DE CUSTO` no nome.

O robô procura automaticamente o arquivo mais recente que contenha esse texto no nome.

### Credenciais da Uber

O e-mail e a senha da Uber devem estar preenchidos na aba `Configurações` do Hub.

### Celular disponível

A Uber pode solicitar um código de autenticação enviado por SMS.

O código deve ser digitado manualmente no navegador. O Hub não salva o código recebido por SMS.

### Google Chrome

O Google Chrome deve estar instalado e disponível para a automação.

### Outlook aberto

Mantenha o Outlook aberto e conectado à conta corporativa.

O Outlook será utilizado para criar um rascunho quando o robô encontrar viagens duplicadas que precisam de análise.

## Como funciona a execução

1. O robô identifica o mês anterior.
2. Cria ou localiza a pasta correspondente ao período na rede.
3. Abre o Google Chrome.
4. Acessa o Uber Business.
5. Preenche automaticamente o e-mail configurado no Hub.
6. Trata as etapas de autenticação apresentadas pela Uber.
7. Caso seja solicitado, aguarda o preenchimento manual do código SMS.
8. Preenche automaticamente a senha configurada no Hub.
9. Seleciona a organização `Uber Central`.
10. Acessa a área de Pagamentos.
11. Baixa o relatório de transações em formato CSV.
12. Baixa a Nota de Débito em formato PDF.
13. Move os arquivos para a pasta pública do mês anterior.
14. Mantém o nome original da Nota de Débito em PDF.
15. Converte o relatório CSV para Excel.
16. Valida os valores da coluna `Observação de despesas do programa de vouchers` usando a Base de Centro de Custo.
17. Gera a planilha final de validação.
18. Cria a aba `Rateio` com os valores consolidados.
19. Cria um rascunho no Outlook quando identifica viagens duplicadas que precisam de análise.
20. Abre a pasta do resultado ao finalizar.

## Arquivos gerados

O processo gera os seguintes arquivos na pasta do período:

- Nota de Débito em PDF com o nome original do download;
- planilha bruta de utilização;
- planilha final de validação e rateio.

A planilha final segue este padrão:

```text
Validacao_Utilizacao_<MÊS>.xlsx
```

## O que validar ao final

### Pasta do período

Verifique se:

- a pasta do mês anterior foi criada no caminho do Uber Central;
- a Nota de Débito em PDF está na pasta;
- o PDF foi mantido com o nome original;
- a planilha bruta foi criada;
- a planilha final de validação foi criada.

### Planilha final

Abra o arquivo:

```text
Validacao_Utilizacao_<MÊS>.xlsx
```

Confira:

- se a coluna `Validação` foi criada;
- se a coluna `Validação` está imediatamente à direita da coluna `Observação de despesas do programa de vouchers`;
- se existem registros classificados como `Não localizado`;
- se os centros de custo foram identificados corretamente;
- se a aba `Rateio` foi criada;
- se os valores foram consolidados corretamente;
- se o valor total está de acordo com o documento da Uber.

### Rascunho de e-mail

Se o Outlook abrir um rascunho, significa que o robô encontrou viagens duplicadas que precisam de análise.

Antes de enviar:

1. Confira os IDs das viagens.
2. Confira os valores.
3. Verifique se existe realmente uma cobrança duplicada.
4. Revise os destinatários e o conteúdo da mensagem.
5. Envie o e-mail somente após a conferência.

Se nenhum rascunho for criado, não foram encontradas duplicidades que exigissem análise.

## Cuidados importantes

- Não feche o Chrome durante a execução.
- Digite o código SMS quando solicitado.
- Não compartilhe sua senha ou o código recebido.
- Não feche o Outlook durante a geração do rascunho.
- Confirme se o computador está conectado à rede corporativa.
- Revise a planilha final antes de utilizar o rateio.

## Possíveis problemas

### Organização não encontrada

Confirme se seu e-mail está cadastrado como Administrador na organização do Uber Central.

### Código SMS não recebido

Confira:

- se o telefone cadastrado na Uber está correto;
- se existe sinal no aparelho;
- se a Uber apresentou outra opção de autenticação;
- se a mensagem não foi bloqueada pelo aparelho.

### Download não concluído

Verifique se:

- o Chrome permitiu o download;
- a página de Pagamentos foi carregada;
- o menu de download foi aberto;
- o relatório está disponível para o período;
- existe espaço disponível no computador.

### Centro de custo não localizado

Confira:

- se o código foi preenchido corretamente na coluna `Observação de despesas do programa de vouchers`;
- se a Base de Centro de Custo está atualizada;
- se o centro de custo existe na base;
- se o robô selecionou a base mais recente.

### Rascunho não criado

Confirme se:

- o Outlook está aberto;
- a conta corporativa está conectada;
- foram encontradas viagens classificadas para verificação.