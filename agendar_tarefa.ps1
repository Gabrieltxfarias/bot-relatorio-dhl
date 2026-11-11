<#
  Cria (ou atualiza) a tarefa do Agendador do Windows que roda o bot.

  Uso (PowerShell, na pasta do projeto, SEM precisar ser administrador):
      powershell -ExecutionPolicy Bypass -File .\agendar_tarefa.ps1
      powershell -ExecutionPolicy Bypass -File .\agendar_tarefa.ps1 -Horario 06:30
      powershell -ExecutionPolicy Bypass -File .\agendar_tarefa.ps1 -Testar     # roda agora
      powershell -ExecutionPolicy Bypass -File .\agendar_tarefa.ps1 -Remover

  A tarefa:
    - roda todo dia no horario (padrao 00:01);
    - se o PC estava desligado no horario, roda assim que ele for ligado;
    - tenta acordar o PC se ele estiver em suspensao;
    - roda tambem 5 min apos voce entrar no Windows (se ja estiver em dia, o bot
      encerra em segundos sem abrir o navegador);
    - o bot recupera sozinho os dias perdidos (estado.json).
#>
param(
    [string]$Horario = "00:01",
    [switch]$Remover,
    [switch]$Testar
)
$ErrorActionPreference = "Stop"
$Nome = "Bot Relatorio DHL - Servicos Acessorios"
$Pasta = Split-Path -Parent $MyInvocation.MyCommand.Path
$Bat = Join-Path $Pasta "executar_bot.bat"
$Usuario = "$env:USERDOMAIN\$env:USERNAME"

if ($Remover) {
    Unregister-ScheduledTask -TaskName $Nome -Confirm:$false
    Write-Host "Tarefa removida: $Nome"
    exit 0
}
if ($Testar) {
    Start-ScheduledTask -TaskName $Nome
    Write-Host "Tarefa iniciada. Acompanhe em: $Pasta\logs"
    exit 0
}

if (-not (Test-Path $Bat)) { throw "Nao encontrei $Bat" }
if (-not (Test-Path (Join-Path $Pasta ".venv\Scripts\python.exe"))) {
    Write-Warning "Ambiente .venv nao encontrado. Faca a instalacao do README antes do primeiro horario agendado."
}

$Acao = New-ScheduledTaskAction -Execute "cmd.exe" -Argument "/c `"$Bat`"" -WorkingDirectory $Pasta
$Diario = New-ScheduledTaskTrigger -Daily -At $Horario
$Logon = New-ScheduledTaskTrigger -AtLogOn -User $Usuario
$Logon.Delay = "PT5M"

$Config = New-ScheduledTaskSettingsSet `
    -StartWhenAvailable `
    -WakeToRun `
    -AllowStartIfOnBatteries `
    -DontStopIfGoingOnBatteries `
    -ExecutionTimeLimit (New-TimeSpan -Hours 3) `
    -MultipleInstances IgnoreNew

# Interactive: roda na sua sessao do Windows (necessario para o OneDrive
# sincronizar e para a janela de login aparecer, se a sessao do TMS expirar).
$Principal = New-ScheduledTaskPrincipal -UserId $Usuario -LogonType Interactive -RunLevel Limited
$Descricao = "Extrai o relatorio de Servicos Acessorios do TMS e alimenta a planilha do OneDrive. Projeto: $Pasta"

try {
    Register-ScheduledTask -TaskName $Nome -Action $Acao -Trigger @($Diario, $Logon) `
        -Settings $Config -Principal $Principal -Description $Descricao -Force | Out-Null
    $Gatilhos = "todo dia as $Horario e 5 min apos entrar no Windows"
} catch {
    Write-Warning "Nao foi possivel criar o gatilho de logon ($($_.Exception.Message)). Criando so o diario."
    Register-ScheduledTask -TaskName $Nome -Action $Acao -Trigger $Diario `
        -Settings $Config -Principal $Principal -Description $Descricao -Force | Out-Null
    $Gatilhos = "todo dia as $Horario (ou assim que o PC ligar, se estiver desligado no horario)"
}

Write-Host ""
Write-Host "Tarefa criada: $Nome"
Write-Host "Quando roda : $Gatilhos"
Write-Host "Logs        : $Pasta\logs"
Write-Host ""
Write-Host "Para o PC acordar da suspensao as $Horario, deixe ativado:"
Write-Host "  Painel de Controle > Opcoes de Energia > Alterar configuracoes do plano >"
Write-Host "  Configuracoes avancadas > Suspender > Permitir temporizadores de ativacao = Habilitar"
Write-Host ""
Write-Host "Teste agora com:  powershell -ExecutionPolicy Bypass -File .\agendar_tarefa.ps1 -Testar"
