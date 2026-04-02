# Guia do DataLogger EmbTech

O `DataLogger` foi pensado para uso de desenvolvimento, bancada e validação de placas eletrônicas.  
Ele permite transformar o terminal da página inicial em uma fonte estruturada de dados para Excel.

## O que ele faz

- Salva eventos do terminal em um arquivo `.xlsx`
- Extrai valores automaticamente com `regex`
- Organiza os dados por colunas
- Pode trabalhar em dois modos:
  - `Evento por linha`: cada mensagem vira uma nova linha
  - `Snapshot consolidado`: várias mensagens próximas no tempo atualizam a mesma linha
- Mantém um `preview` com amostras recentes para validar as regras

## Fluxo recomendado

1. Na página inicial, clique em `Configurar` ao lado do `DataLogger`
2. Escolha o arquivo Excel
3. Defina o nome da planilha
4. Escolha um `preset` próximo do seu caso
5. Ajuste as colunas e regras
6. Use `Simular no preview` com uma mensagem real do terminal
7. Ative o `DataLogger`

## Estrutura básica da planilha

As colunas base normalmente ficam assim:

- `A`: Data/Hora
- `B`: Tipo
- `C`: Porta
- `D`: Latência
- `E`: Mensagem

As colunas seguintes podem ser usadas para valores extraídos:

- `F`: AN1
- `G`: AN2
- `H`: Corrente
- `I`: Temperatura

Você pode reorganizar isso conforme o projeto.

## Presets

### Generico

Use quando quiser registrar números que aparecem em qualquer mensagem do terminal.

Bom para:

- debug rápido
- desenvolvimento inicial
- mensagens não padronizadas

### Sensores EmbTech

Use quando o terminal já traz sensores no estilo:

`AN1: 3mV`
`AN2: 5mV`
`IAC: 127mA`
`Sonda 1: 27`

Bom para:

- calibração
- monitoramento de sensores
- teste de estabilidade

### Estados Digitais

Use quando o firmware imprime estados em hexadecimal:

`Entradas: 0x0`
`Saidas: 0xF`

Bom para:

- validação de I/O
- testes de comando/resposta
- análise de estados digitais

## Modos de captura

### Evento por linha

Cada evento vira uma linha nova.

Use quando:

- você quer histórico completo
- quer ver a ordem dos eventos
- está analisando transientes e respostas do sistema

### Snapshot consolidado

Várias mensagens próximas atualizam a mesma linha.

Use quando:

- um conjunto de mensagens representa um mesmo estado da placa
- você quer uma linha por leitura consolidada
- está monitorando sensores em ciclos

## Como criar uma regra boa

Cada regra tem:

- `Cabeçalho`: nome da coluna no Excel
- `Coluna`: letra da coluna
- `Tipo`: quais eventos a regra considera
- `Porta`: filtro opcional por nome da porta
- `Regex`: expressão para encontrar o valor
- `Valor`: define se grava o `grupo 1`, o `match completo` ou a `mensagem completa`

### Exemplo 1

Mensagem:

```text
AN1: 3mV
```

Regex:

```regex
AN1:\s*([-+]?\d+(?:\.\d+)?)mV
```

Resultado:

- grava `3`

### Exemplo 2

Mensagem:

```text
Entradas: 0x0
```

Regex:

```regex
Entradas:\s*(0x[0-9A-Fa-f]+)
```

Resultado:

- grava `0x0`

## Preview inteligente

O bloco `Monitor e preview` serve para:

- colar mensagens reais do terminal
- testar regex antes de logar
- ver se a coluna escolhida faz sentido
- acompanhar o histórico recente de um valor numérico

### Simular no preview

Cole uma mensagem real e clique em `Simular no preview`.

Se a regra estiver correta:

- o valor aparece em `Última extração`
- a série numérica aparece no gráfico

## Boas práticas

- mantenha a mesma estrutura de colunas por projeto
- use `snapshot` para leituras periódicas de sensores
- use `evento por linha` para debug e investigação
- prefira nomes de cabeçalho curtos e claros
- valide a regex com mensagens reais
- crie um arquivo por sessão ou por placa quando necessário

## Exemplos de uso em engenharia

### Monitorar sensores

- corrente
- tensão
- temperatura
- entradas analógicas

### Validar barramentos e I/O

- estados de entrada
- estados de saída
- respostas do firmware
- flags de erro

### Desenvolvimento de firmware

- acompanhar mudanças de estado
- comparar antes/depois de comandos
- medir latência entre mensagens

## Quando usar filtro por porta

Use o campo `Porta` da regra quando:

- a página inicial estiver usando mais de uma porta
- houver principal e modbus ativos
- o mesmo terminal misturar mensagens de origens diferentes

## Limitações atuais

- o logger grava em Excel local `.xlsx`
- a extração é baseada em regex e mensagens do terminal
- o preview histórico mostra valores numéricos das regras

## Estratégia recomendada por maturidade

### Fase 1

- preset `Generico`
- evento por linha
- poucas regras

### Fase 2

- preset de sensores
- snapshot consolidado
- colunas por variável

### Fase 3

- planilha padrão por projeto
- regras refinadas por porta
- uso contínuo em bancada

## Resumo rápido

- `Configurar`: define arquivo, planilha, modo e regras
- `Simular no preview`: testa a extração
- `Ativar DataLogger`: começa a gravar
- `Snapshot`: melhor para leitura consolidada
- `Evento por linha`: melhor para debug detalhado
