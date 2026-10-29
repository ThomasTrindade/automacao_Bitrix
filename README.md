# automacao_Bitrix

# Automação: preencher nomes de usuários do Bitrix24 em planilha

## 1. Objetivo

Receber uma planilha `.xlsx` com uma coluna de IDs de usuário (`author_id`) e preencher, na coluna ao lado, o **nome completo** de cada pessoa, buscando os dados no Bitrix24 (`villela.bitrix24.com.br`).

Modelo da planilha:

| A (`author_id`) | B (`Nome`) |
|---|---|
| 29184 | *(preenchido pelo script)* |

- Linha 1: cabeçalho.
- Dados a partir da linha 2.
- O arquivo original **não é alterado**. O resultado sai em uma cópia chamada `<nome>_preenchido.xlsx`, na mesma pasta.

## 2. Como funciona

O script não automatiza o navegador. Ele usa a **API REST do Bitrix24** por meio de um **webhook de entrada**, o que é mais rápido e estável do que simular cliques.

Fluxo:

1. Lê o webhook (variável de ambiente `BITRIX_WEBHOOK` ou digitado na execução).
2. Copia a planilha de entrada para o arquivo de saída.
3. Para cada linha, lê o ID da coluna A.
4. Consulta `user.get` no Bitrix24 e monta `NAME + LAST_NAME`.
5. Grava o resultado na coluna B e salva.

### Regras de preenchimento

| Situação | Valor gravado na coluna B |
|---|---|
| ID encontrado entre usuários ativos | Nome e sobrenome |
| ID não encontrado entre os ativos | Busca entre os inativos (desligados) |
| ID não encontrado em nenhum | `Não encontrado` |
| ID `0` ou valor não numérico | `Sem usuário` |
| Linha com ID vazio | Ignorada |
| Falha na consulta | `Erro: <mensagem>` |

### Desempenho e limites

- IDs repetidos são consultados **uma única vez** (cache em memória).
- Há uma pausa de ~550 ms entre chamadas, para respeitar o limite do Bitrix24 de cerca de 2 requisições por segundo.

## 3. Arquivos

| Arquivo | Descrição |
|---|---|
| `preencher_nomes_bitrix_excel.ps1` | **Versão recomendada.** Usa o Microsoft Excel instalado (COM). Não depende de módulos externos. |
| `preencher_nomes_bitrix.ps1` | Versão alternativa. Usa o módulo PowerShell `ImportExcel` (não precisa do Excel instalado). |

### Por que existem duas versões

A versão com `ImportExcel` falhou neste ambiente com erro ao carregar a `EPPlus.dll` (`HRESULT 0x80131515`). Dois fatores contribuíram:

- A pasta Documentos do usuário está redirecionada para um servidor de rede, e o PowerShell não carrega DLLs de compartilhamentos de rede.
- Mesmo após mover o módulo para uma pasta local (`C:\PSModulos`), o erro continuou, provavelmente por política de segurança do computador.

A versão com Excel COM contorna o problema e é a indicada onde o Excel desktop estiver instalado.

## 4. Pré-requisitos

- Windows com PowerShell.
- **Versão Excel:** Microsoft Excel desktop instalado.
- **Versão ImportExcel:** acesso à internet para baixar o módulo na primeira execução.
- Webhook de entrada criado no Bitrix24 (ver seção 5).
- Planilha e scripts em uma pasta **local** (por exemplo `C:\Automacao`), não em compartilhamento de rede.

## 5. Configuração do webhook no Bitrix24

É necessário ser administrador do portal.

1. Acesse `https://villela.bitrix24.com.br/devops/section/standard-integrations/` (ou, no menu, Recursos para desenvolvedores → Outros → Webhook de entrada).
2. Crie um **Webhook de entrada**.
3. Em **permissões de acesso**, marque **Usuários (user)** (e, se existir, `user_basic`).
4. Salve e copie a URL pelo botão de copiar. Ela tem este formato:

   ```
   https://villela.bitrix24.com.br/rest/<ID-DO-USUARIO>/<CODIGO-DE-16-CARACTERES>/
   ```

Observações importantes:

- A URL deve terminar na barra depois do código, **sem** nome de método no final (como `profile.json` ou `user.option.set.json`). O script acrescenta `user.get.json` sozinho.
- O número que aparece na coluna **ID** da lista de webhooks é o ID do webhook, **não** o ID do usuário usado na URL.
- O webhook herda as permissões do usuário que o criou. Se essa conta for desativada, o webhook para de funcionar.

## 6. Como executar

Em uma janela do PowerShell:

```powershell
cd C:\Automacao
Set-ExecutionPolicy -Scope Process Bypass
$env:BITRIX_WEBHOOK = Read-Host "Cole a URL do webhook"
.\preencher_nomes_bitrix_excel.ps1 -Entrada ".\planilha.xlsx"
```

Detalhes:

- Ao executar `Read-Host`, cole a URL com o botão direito do mouse e dê Enter.
- A variável `BITRIX_WEBHOOK` vale só para a janela atual. Em uma janela nova, defina de novo.
- Feche a planilha no Excel antes de rodar.
- Teste primeiro com 3 ou 4 linhas cujos IDs você conheça.

### Parâmetros opcionais

| Parâmetro | Padrão | Função |
|---|---|---|
| `-Saida` | `<nome>_preenchido.xlsx` | Caminho do arquivo de saída |
| `-ColunaId` | `1` (A) | Coluna onde estão os IDs |
| `-ColunaNome` | `2` (B) | Coluna onde o nome será gravado |
| `-LinhaInicial` | `2` | Primeira linha de dados |

### Teste rápido do webhook (sem rodar o script)

```powershell
$u = $env:BITRIX_WEBHOOK.Trim().TrimEnd('/')
(Invoke-RestMethod "$u/user.get.json?filter[ID]=1").result | Select-Object NAME, LAST_NAME
```

Troque o `1` por um ID que exista no portal.

### Conferir a URL sem expor o código

```powershell
$p = ($env:BITRIX_WEBHOOK.Trim().TrimEnd('/') -split '/')
"Host: " + $p[2] + " | Usuario: " + $p[-2] + " | Tamanho do codigo: " + $p[-1].Length
```

O resultado esperado é o host do portal, um número como usuário e código com 16 caracteres.

## 7. Problemas encontrados e soluções

| Sintoma | Causa | Solução |
|---|---|---|
| `env:BITRIX_WEBHOOK não é reconhecido como cmdlet` | Faltou o `$` no comando | Usar `$env:BITRIX_WEBHOOK = "..."` |
| `Não foi possível carregar ... EPPlus.dll` (`0x80131515`) | Bloqueio de DLL (pasta de rede e/ou política de segurança) | Usar a versão `preencher_nomes_bitrix_excel.ps1` |
| `O módulo ImportExcel está em uso` ao desinstalar | Módulo carregado na sessão | Fechar todas as janelas do PowerShell e refazer em uma nova |
| `404 Não Localizado` | URL montada de forma errada | Usar a URL base terminando em `/`, sem nome de método |
| `401` com `INVALID_CREDENTIALS` | Variável com texto de exemplo (`NUMERO/CODIGO`) ou código errado | Colar a URL real do webhook |
| `Não é possível chamar um método em uma expressão de valor nulo` | Variável `BITRIX_WEBHOOK` vazia (janela nova) | Definir a variável de novo na janela atual |
| `401` com `insufficient_scope` | Webhook sem permissão de Usuários | Marcar **Usuários (user)** no webhook e salvar |
| `Classe não registrada` / `COM class factory` | Excel desktop não instalado | Instalar o Excel ou usar PowerShell 7 com `ImportExcel` |

Como ver o motivo exato de um `401`:

```powershell
$u = $env:BITRIX_WEBHOOK.Trim().TrimEnd('/')
try { Invoke-RestMethod "$u/user.get.json?filter[ID]=1" } catch { (New-Object IO.StreamReader($_.Exception.Response.GetResponseStream())).ReadToEnd() }
```

## 8. Segurança

- A URL do webhook funciona como uma **senha** do portal. Não a coloque dentro do script, não a envie por chat ou e-mail e não a salve em repositórios.
- Se a URL for exposta, apague o webhook no Bitrix24 e crie outro.
- Quando o webhook não for mais necessário, apague-o.
- Conceda ao webhook apenas a permissão de Usuários, que é a necessária para esta automação.

## 9. Limitações conhecidas

- Processa a **primeira aba** da planilha.
- Faz uma consulta por ID único, em sequência, então planilhas com milhares de IDs distintos levam alguns minutos.
- A versão com Excel COM exige o Excel desktop instalado e não roda em ambientes sem interface do Office.
- O script grava `Nome` + `Sobrenome` conforme cadastrado no Bitrix24. Se o cadastro estiver incompleto, o nome sai incompleto.

## 10. Evoluções possíveis

- Consultar vários IDs por chamada usando o método `batch` do Bitrix24, para acelerar planilhas grandes.
- Versão em Google Apps Script, caso a planilha passe a ficar no Google Sheets.
- Agendamento da execução, em plano que permita tarefas agendadas.
