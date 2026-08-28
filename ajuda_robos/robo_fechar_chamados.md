# Fechar Chamados a Vencer

## Objetivo

Este robô acessa o Agilis e analisa os chamados disponíveis na caixa de Solicitações que vencem no dia atual.

A automação pode fechar todos os chamados do dia em uma única execução ou permanecer monitorando os chamados até que falte uma hora para o vencimento.

## Modos de execução

### Fechar todos de hoje

Nesse modo, o robô:

1. Acessa o Agilis.
2. Abre a caixa de Solicitações.
3. Localiza os chamados que vencem no dia atual.
4. Identifica a categoria e a subcategoria.
5. Seleciona a resposta padrão adequada.
6. Fecha os chamados encontrados.

Utilize essa opção quando for necessário concluir todos os chamados do dia imediatamente.

### Monitorar o dia todo

Nesse modo, o robô:

1. Acessa o Agilis.
2. Analisa periodicamente os chamados.
3. Verifica o horário de vencimento.
4. Aguarda até que falte uma hora para o vencimento.
5. Identifica a categoria e a subcategoria.
6. Insere a resposta padrão adequada.
7. Fecha o chamado.

Enquanto o monitoramento estiver ativo, mantenha o Hub aberto.

## Respostas padrão

O robô considera a categoria e a subcategoria do chamado para selecionar a resposta adequada.

Para que o código localize um modelo automaticamente, o nome do modelo de resposta deve corresponder ao texto exibido após o campo Categoria no Agilis.

Exemplo:

Administrativo - SOLICITAÇÃO DE ENVIO DE CORRESPONDÊNCIA

O texto precisa seguir o mesmo padrão utilizado pelo Agilis.

## Pré-requisitos

Antes da execução, confirme se:

- as credenciais MRV estão configuradas;
- o Agilis está acessível;
- os modelos de resposta estão cadastrados;
- os nomes dos modelos correspondem às categorias;
- a autenticação MFA está disponível, caso seja solicitada.

## Resultado esperado

Os chamados serão fechados com o modelo de resposta correspondente à categoria e à subcategoria identificadas.

## Possíveis problemas

### O modelo de resposta não foi localizado

Confira se o nome do modelo:

- corresponde ao texto da categoria no Agilis;
- utiliza a mesma grafia;
- mantém espaços, acentos e pontuação;
- não possui caracteres adicionais.

### Um chamado não foi fechado

Verifique se:

- o chamado vence no dia atual;
- o chamado está na caixa de Solicitações;
- a categoria possui um modelo correspondente;
- o status permite o fechamento;
- o Agilis não está aguardando MFA.

### O modo de monitoramento parou

Confira se:

- o Hub continua aberto;
- o processo não foi cancelado;
- o computador não entrou em suspensão;
- a conexão com o Agilis permanece ativa.