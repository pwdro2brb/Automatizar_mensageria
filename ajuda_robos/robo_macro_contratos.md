# Atualizar Macro de Contratos

## Objetivo

Este robô localiza e atualiza a macro utilizada no processo de cobrança de boletos de contratos.

A automação identifica a pasta pública mais recente, localiza o arquivo necessário e executa a atualização da macro.

O processo atualiza o conteúdo da macro sem remover seu funcionamento. Depois da atualização, o arquivo continua sendo uma planilha habilitada para macro e pode ser utilizado normalmente.

## Pré-requisitos

Para executar o robô, é necessário:

- possuir acesso à pasta pública;
- manter a pasta mais recente disponível;
- manter o arquivo da macro dentro da pasta esperada;
- fechar o arquivo no Excel antes da execução;
- possuir acesso aos arquivos utilizados na atualização.

## Preparação

Não é necessário editar ou preparar manualmente o arquivo antes da execução.

Antes de iniciar, apenas confirme se:

1. A pasta pública está acessível.
2. A pasta mais recente do processo existe.
3. O arquivo esperado está dentro da pasta.
4. A macro está fechada.
5. Nenhum outro usuário está bloqueando o arquivo para edição.

## Funcionamento

Durante a execução, o robô:

1. Acessa o caminho configurado da pasta pública.
2. Localiza a pasta mais recente.
3. Identifica o arquivo da macro de cobrança de boletos.
4. Abre o arquivo para atualização.
5. Atualiza as informações utilizadas pelo processo.
6. Preserva o formato e o funcionamento da macro.
7. Salva o arquivo atualizado.

## Resultado esperado

Ao final da execução, a macro de cobrança de boletos estará atualizada e continuará funcionando normalmente.

Depois do processamento, recomenda-se abrir o arquivo e conferir:

- se a atualização foi concluída;
- se os dados mais recentes estão disponíveis;
- se as fórmulas foram preservadas;
- se os recursos da macro continuam habilitados;
- se o arquivo foi salvo no local correto.

## Cuidados importantes

- Não renomeie o arquivo sem verificar se o robô depende do nome atual.
- Não mova o arquivo para outra pasta antes da execução.
- Não mantenha a macro aberta no Excel.
- Não altere a estrutura das abas usadas pela automação.
- Não salve o arquivo em um formato que remova as macros.

## Possíveis problemas

### A pasta mais recente não foi encontrada

Verifique se:

- a pasta pública está acessível;
- a pasta do período mais recente foi criada;
- o caminho da rede está disponível;
- o computador está conectado à rede corporativa.

### O arquivo não foi encontrado

Confirme se o arquivo está dentro da pasta mais recente e se não foi renomeado ou movido.

### O arquivo está bloqueado

Feche o Excel e verifique se o arquivo não está aberto por outro usuário.

### A macro deixou de funcionar

Confirme se o arquivo continua salvo no formato habilitado para macros e se o conteúdo não foi salvo como uma planilha comum.