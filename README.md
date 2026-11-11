# Bot de Relatório — Serviços Acessórios (TMS DHL)

Este bot faz o seguinte:
1. Entra no TMS, reaproveitando a sessão salva ou fazendo login sozinho com a senha guardada no Windows.
2. Escolhe o papel **Analista de Serviços Acessórios** e abre **Inclusão de Serviços Acessórios**, ambos por link direto.
3. Aplica os filtros: tipo **Receita de Documento**, status **Aguardando Aprovação** e **Aprovada**,
   **data inicial e final iguais ao dia anterior** e usuário solicitante **Gabriel Teixeira de Farias**.
4. Clica em **Pesquisar**, depois em **⋮ > Exportar para Excel**, espera o aviso de relatório gerado e baixa o arquivo.
5. Do arquivo exportado, mantém **só o serviço "Troca de Gelo por Volume - Cliente/Transportador"**
   (configurável em `planilha.filtrar_linhas`) e descarta as outras linhas. O arquivo original em `downloads/` não é alterado.
6. Grava as linhas na **planilha online (Excel no OneDrive/SharePoint)**, com exatamente as colunas do
   relatório do TMS (Viagem … Usuário), mantendo datas/horas e números. Cada lançamento é identificado por **Documento Logística Original +
   Inclusão (data/hora com segundos)**. Dois lançamentos para o mesmo documento viram duas linhas. Se uma linha
   já gravada vier de novo com outro status, o bot atualiza a linha em vez de duplicar.

## Instalação (uma vez, no PC com acesso ao TMS)
**Pré-requisito: Python 3.10 ou mais novo.** Para conferir, abra o Prompt de Comando e digite `python --version`.
Se não aparecer a versão, instale pelo [python.org](https://www.python.org/downloads/) e marque **"Add python.exe to PATH"**.
Sem permissão de administrador, peça à TI ou use o Python da Microsoft Store.

Na pasta do projeto:
```bat
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
playwright install msedge
```

### Verificar se o PC está pronto
```bat
.venv\Scripts\activate
python verificar_ambiente.py
```
O script confere Python, bibliotecas, navegador, credenciais, acesso ao TMS (VPN), planilha do OneDrive,
tarefa agendada e a última data processada. Ele mostra `[OK]`, `[AVISO]` ou `[FALHA]`, e para cada problema diz como resolver.

## Login sem passar sua senha para ninguém
- **Opção A (recomendada):** rode `python salvar_credenciais.py`. O script pede o usuário e a senha
  no terminal, e a senha não aparece enquanto você digita. Ela fica criptografada no
  **Gerenciador de Credenciais do Windows**, nunca num arquivo. Para remover: `python salvar_credenciais.py --apagar`.
- **Sessão salva:** na 1ª execução, com `headless: false`, a janela abre. Se houver MFA,
  aprove no celular. Os cookies ficam em `perfil_navegador/` e as próximas execuções já entram logadas.
- **Opção B:** se o keyring não funcionar, copie `.env.example` para `.env` e preencha. O arquivo fica em texto puro.

> `perfil_navegador/` e `.env` dão acesso à sua conta. Não compartilhe esses arquivos. Eles já estão no `.gitignore`.

## Planilha online (OneDrive/SharePoint)
O bot grava no arquivo da **pasta sincronizada do OneDrive** no seu PC, e o OneDrive envia para a nuvem.
1. No Explorer, localize a planilha. Ela precisa estar sincronizada; se aparecer o ícone de nuvem, clique
   com o botão direito e escolha "Sempre manter neste dispositivo".
2. Segure Shift, clique com o botão direito no arquivo e escolha **Copiar como caminho**.
3. Cole o caminho em `planilha.arquivo` no `config.yaml`, trocando `\` por `/`.
4. Se a planilha tiver mais de uma aba, informe o nome da aba em `planilha.aba`.

Proteções:
- se a planilha estiver aberta no Excel do PC, o bot espera até 10 minutos para ela ser fechada;
- antes de cada gravação, uma cópia vai para `backups/` (as últimas 30 ficam guardadas);
- se a planilha for uma Tabela do Excel, as linhas novas entram na tabela;
- se a planilha tiver gráficos, imagens ou tabelas dinâmicas, o bot **não grava**, porque esses itens seriam
  apagados. Faça os gráficos em outro arquivo que leia esta planilha.
- Quem estiver com a planilha aberta no navegador recebe as linhas novas quando o OneDrive sincronizar,
  normalmente em poucos segundos.

## Uso
```bat
python main.py                                   :: tudo que falta até ontem (padrão)
python main.py --data 10/11/2026                 :: um dia específico
python main.py --de 01/11/2026 --ate 10/11/2026  :: um período específico
python main.py --arquivo rel.xlsx                :: só alimenta a planilha com um arquivo já baixado
```
Os logs ficam em `logs/`. Em caso de erro, o bot salva um print e o HTML da tela no mesmo lugar.

### Recuperação de dias perdidos
O bot guarda em `estado.json` a última data processada com sucesso. Sem datas na linha de comando, ele busca
**de uma vez todo o período que falta até ontem**. Por exemplo, se o PC ficou desligado no fim de semana, a
execução de segunda traz sexta, sábado e domingo. Se já estiver em dia, o bot encerra em segundos, sem abrir o navegador.
- Se uma execução falhar, por exemplo com a VPN desconectada, o bot tenta de novo 3 vezes, com 20 minutos de intervalo
  (`execucao` no `config.yaml`). Se ainda assim falhar, o período fica pendente para a próxima execução.
- Uma execução recupera no máximo 31 dias (`filtros.max_dias_recuperacao`). Para períodos maiores, use `--de`/`--ate`.

## Como o bot navega
- O papel é escolhido abrindo direto `Roles/Select/K_ANALISTASERVICOSACESSORIOS`, e a tela pelo link do `Grid.aspx`.
- Tipo, datas, Pesquisar, ⋮ e Exportar usam os IDs reais do TMS, definidos em `config.yaml > seletores`.
- **Status** e **Usuário Solicitante** são caixas select2. O bot acha cada campo pelo bloco
  `<div data-field=... data-label=...>` cujo rótulo é "Status" ou contém "Solicitante". Antes de
  preencher, ele limpa as seleções anteriores. O usuário é digitado na busca e escolhido na lista.
- Depois de exportar, o bot acompanha a **barra de processos** (`#header_task_bar`). Quando surge o item
  "Exportar para Excel" com o marcador `process-success` ("Concluído"), o bot abre a barra e clica no item para baixar.
- A pesquisa de satisfação (Amplifique.me) que aparece por cima da tela fica escondida automaticamente,
  para não bloquear os cliques.

## Se algo falhar
O bot salva em `logs/` um **print e o HTML** da tela no momento do erro. Os pontos que podem precisar de ajuste no `config.yaml`:
- `seletores.status_campo` / `seletores.usuario_campo`: o data-field exato, por exemplo `div[data-field='STATUS']`;
- `seletores.processo_concluido`: o marcador de item pronto na barra de processos (padrão: `.process-success`).

## Agendar todo dia (0h01)
No PowerShell, dentro da pasta do projeto (não precisa ser administrador):
```bat
powershell -ExecutionPolicy Bypass -File .\agendar_tarefa.ps1
```
A tarefa criada:
- roda **todo dia às 0h01**;
- se o PC estava **desligado** nesse horário, roda **assim que ele for ligado**;
- tenta **acordar o PC da suspensão**. Para isso, ative "Permitir temporizadores de ativação" nas Opções de Energia;
- roda também 5 minutos depois que você entra no Windows, como reforço. Se já estiver em dia, não faz nada.

Outros comandos: `-Testar` roda a tarefa agora, `-Horario 06:30` muda o horário e `-Remover` apaga a tarefa.
A tarefa roda na sua sessão do Windows, que é o que o OneDrive precisa para sincronizar. O PC pode estar com a tela bloqueada.

## Testes
`pytest` roda os testes contra um TMS falso local (`tests/mock_tms.py`).
