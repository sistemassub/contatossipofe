<# :
@echo off
chcp 65001 >nul
title Instalar backup do Conecta SIPOFE
powershell -NoProfile -ExecutionPolicy Bypass -Command "$f='%~f0'; Invoke-Expression ([IO.File]::ReadAllText($f,[Text.Encoding]::UTF8))"
echo.
pause
exit /b
#>
$BACKUP_PS1 = @'
# Backup automático do Conecta SIPOFE
# Roda todo dia pelo Agendador de Tarefas do Windows. Não precisa abrir este arquivo.
$ErrorActionPreference = 'Stop'
$Sis  = $PSScriptRoot
$Base = Split-Path -Parent $Sis
$Bk   = Join-Path $Base 'backups'
$Pl   = Join-Path $Base 'planilhas'
$Url  = 'https://wlixkuxembfiotbmyrgk.supabase.co'
$Key  = 'sb_publishable_5kJA97mc8eDIbLdiEMiu0A_yoLrxx1_'
$Txt  = New-Object System.Text.UTF8Encoding($true)
$Json = New-Object System.Text.UTF8Encoding($false)
$Erro = Join-Path $Base 'BACKUP COM ERRO - LEIA.txt'
$Stat = Join-Path $Base 'ultimo-backup.txt'
function Log($m) { Add-Content -Path (Join-Path $Sis 'backup.log') -Value ('{0}  {1}' -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $m) -Encoding UTF8 }
function Csv($linhas, $nome) { if ($linhas) { $linhas | Export-Csv -Path (Join-Path $Pl $nome) -Delimiter ';' -NoTypeInformation -Encoding UTF8 } }
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
New-Item -ItemType Directory -Force -Path $Bk, $Pl | Out-Null

try {
  # 1. Entrar com o usuário salvo (senha criptografada pelo Windows para esta conta)
  $cf = Join-Path $Sis 'credencial.xml'
  if (-not (Test-Path $cf)) { throw 'O usuário do backup não está salvo. Rode o instalador de novo.' }
  $cred = Import-Clixml $cf
  $body = @{ email = ($cred.UserName.Trim().ToLower() + '@scpma.local'); password = $cred.GetNetworkCredential().Password } | ConvertTo-Json
  try {
    $tok = Invoke-RestMethod -Method Post -Uri "$Url/auth/v1/token?grant_type=password" -Headers @{ apikey = $Key } -ContentType 'application/json' -Body $body
  } catch { throw 'Não foi possível entrar com o usuário salvo. Se a senha do Hub SCPMA mudou, rode o instalador de novo.' }
  $h = @{ apikey = $Key; Authorization = 'Bearer ' + $tok.access_token }

  # 2. Baixar todos os registros, sem alterar nada
  $partes = @(); $lista = New-Object System.Collections.ArrayList; $off = 0
  do {
    $r = Invoke-WebRequest -UseBasicParsing -Headers $h -Uri "$Url/rest/v1/sipofe_docs?select=colecao,id,ordem,dados&order=colecao.asc,id.asc&limit=1000&offset=$off"
    $raw = [System.Text.Encoding]::UTF8.GetString($r.RawContentStream.ToArray()).Trim()
    $arr = $raw | ConvertFrom-Json
    $qt = ($arr | Measure-Object).Count
    if ($qt -gt 0) { $partes += $raw.Substring(1, $raw.Length - 2); foreach ($x in $arr) { [void]$lista.Add($x) } }
    $off += 1000
  } while ($qt -eq 1000)
  if ($lista.Count -eq 0) { throw 'O banco não devolveu nenhum registro. Confira se o usuário salvo tem acesso ao Conecta SIPOFE.' }

  # 3. Backup completo, no mesmo formato do botão "Backup completo (.json)" (serve para restaurar)
  $hoje = Get-Date -Format 'yyyy-MM-dd'
  $arq = Join-Path $Bk "conecta-sipofe-backup-$hoje.json"
  $gerado = (Get-Date).ToUniversalTime().ToString("yyyy-MM-dd'T'HH:mm:ss'Z'")
  [System.IO.File]::WriteAllText($arq, '{"app":"Conecta SIPOFE","gerado":"' + $gerado + '","rows":[' + ($partes -join ',') + ']}', $Json)

  # 4. Planilhas para consulta (abrem no Excel; sempre a versão mais recente)
  function Dados($c) { $lista | Where-Object { $_.colecao -eq $c } | ForEach-Object { $_.dados } }
  Csv (foreach ($o in (Dados 'orgao')) { foreach ($p in @($o.pessoas)) { if ($p) { [pscustomobject]@{
    'Sigla' = $o.sigla; 'Entidade' = $o.ent; 'Tipo de entidade' = $o.tipo; 'Nome' = $p.nome; 'Ponto focal' = $(if ($p.focal) { 'Sim' } else { '' })
    'Cargo' = $p.cargo; 'Lotação' = $p.lot; 'E-mail' = $p.email; 'Celular' = $p.cel; 'Nível' = $p.fc } } } }) 'Contatos SIPOFE.csv'
  Csv (foreach ($e in (Dados 'eixo')) { foreach ($g in @($e.orgaos)) { if ($g) { [pscustomobject]@{
    'Eixo' = $e.eixo; 'Órgão' = $g.org; 'Código' = $g.cod; 'Gestor' = $g.gestor
    'Representantes' = (@($g.reps) | Where-Object { $_ } | ForEach-Object { ($_.nome + $(if ($_.papel) { ' (' + $_.papel + ')' } else { '' }) + $(if ($_.email) { ' - ' + $_.email } else { '' })) }) -join ' | '
    'Outros e-mails' = (@($g.mails) -join ' | ') } } } }) 'Pontos focais GECEG.csv'
  Csv (foreach ($x in (Dados 'fc')) { [pscustomobject]@{
    'Servidor' = $x.nome; 'Órgão' = $x.org; 'Entidade' = $x.ent; 'Vínculo' = $x.vinc; 'Cód. vínculo' = $x.cod; 'Cargo' = $x.cargo
    'Lotação' = $x.lot; 'Nível' = $x.fc; 'Valor mensal' = $x.valor; 'E-mail' = $x.email; 'Celular' = $x.cel } }) 'Controle FCs e GRGs.csv'
  Csv (foreach ($e in (Dados 'entrevista')) { [pscustomobject]@{
    'Servidor' = $e.nome; 'Órgão' = $e.org; 'Unidade' = $e.uni; 'Processo SEI' = $e.sei; 'Grau' = $e.grau; 'Data de envio' = $e.env
    'Data da entrevista' = $e.ent; 'Habilitação' = $e.hab; 'Aprovação' = $e.apr; 'Resultado' = $(if ($e.res) { $e.res } else { 'Em andamento' }); 'Observações' = $e.obs } }) 'Entrevistas.csv'
  Csv (foreach ($a in (Dados 'visita')) { [pscustomobject]@{
    'Data' = $(if ($a.iso) { $a.iso } else { $a.dataRaw }); 'Horário' = $a.hor; 'Órgão' = $a.org; 'Situação da agenda' = $a.stA
    'Situação da visita' = $a.stV; 'Processo SEI' = $a.sei; 'Pauta' = $a.pauta; 'Encaminhamentos' = $a.enc } }) 'Agenda de visitas.csv'

  # 5. Guardar: todos os dos últimos 30 dias + o primeiro de cada mês mais antigo
  $lim = (Get-Date).AddDays(-30).ToString('yyyy-MM-dd'); $meses = @{}
  foreach ($f in (Get-ChildItem $Bk -Filter 'conecta-sipofe-backup-*.json' | Sort-Object Name)) {
    $d = $f.BaseName.Substring(22); $m = $d.Substring(0, 7)
    if ($d -ge $lim) { continue }
    if (-not $meses.ContainsKey($m)) { $meses[$m] = 1; continue }
    Remove-Item $f.FullName
  }

  $qtd = ($lista | Group-Object colecao | ForEach-Object { '{0}: {1}' -f $_.Name, $_.Count }) -join ', '
  [System.IO.File]::WriteAllText($Stat, ("Último backup: {0}`r`nRegistros: {1} ({2})`r`nArquivo: backups\conecta-sipofe-backup-{3}.json`r`n`r`nO backup roda sozinho todo dia às 13h (ou logo que o computador for ligado, se estiver desligado nesse horário)." -f (Get-Date -Format 'dd/MM/yyyy HH:mm'), $lista.Count, $qtd, $hoje), $Txt)
  if (Test-Path $Erro) { Remove-Item $Erro }
  Log ('OK - ' + $lista.Count + ' registros')
} catch {
  $msg = $_.Exception.Message
  Log ('ERRO - ' + $msg)
  [System.IO.File]::WriteAllText($Erro, ("O backup de {0} não foi feito.`r`n`r`nMotivo: {1}`r`n`r`nO que fazer: confira a internet e tente de novo com o arquivo 'Fazer backup agora'. Se a senha do Hub SCPMA mudou, rode o instalador outra vez." -f (Get-Date -Format 'dd/MM/yyyy HH:mm'), $msg), $Txt)
  exit 1
}
'@
$ErrorActionPreference = 'Stop'
$Url = 'https://wlixkuxembfiotbmyrgk.supabase.co'
$Key = 'sb_publishable_5kJA97mc8eDIbLdiEMiu0A_yoLrxx1_'
$Bom = New-Object System.Text.UTF8Encoding($true)
function Titulo($t) { Write-Host ''; Write-Host $t -ForegroundColor Cyan }
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
try {
  Titulo '== Backup automático do Conecta SIPOFE =='

  # 1. Encontrar a pasta "projetos"
  $docs = [Environment]::GetFolderPath('MyDocuments'); $desk = [Environment]::GetFolderPath('Desktop')
  $cands = @((Join-Path $HOME 'projetos'), (Join-Path $docs 'projetos'), (Join-Path $desk 'projetos'), (Join-Path $HOME 'OneDrive\projetos'), (Join-Path $HOME 'OneDrive\Documentos\projetos'), (Join-Path $HOME 'OneDrive\Documents\projetos'), 'C:\projetos', 'D:\projetos') | Select-Object -Unique
  $proj = $cands | Where-Object { Test-Path $_ -PathType Container } | Select-Object -First 1
  if ($proj) {
    Write-Host "Encontrei a pasta: $proj"
    $r = Read-Host 'Usar esta pasta? Digite S para sim ou N para escolher outra'
    if ($r -match '^[nN]') { $proj = $null }
  } else { Write-Host 'Não encontrei a pasta "projetos" automaticamente. Escolha a pasta na janela que vai abrir.' }
  if (-not $proj) {
    Add-Type -AssemblyName System.Windows.Forms
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.Description = 'Escolha a sua pasta "projetos". A pasta CONECTA SIPOFE será criada dentro dela.'
    if ($dlg.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { throw 'Nenhuma pasta foi escolhida.' }
    $proj = $dlg.SelectedPath
  }
  $base = Join-Path $proj 'CONECTA SIPOFE'; $sis = Join-Path $base 'sistema'
  New-Item -ItemType Directory -Force -Path $base, $sis, (Join-Path $base 'backups'), (Join-Path $base 'planilhas') | Out-Null
  Write-Host "Pasta pronta: $base" -ForegroundColor Green

  # 2. Usuário do Hub SCPMA (a senha fica criptografada pelo Windows, só esta conta do Windows consegue ler)
  Titulo 'Entre com o mesmo usuário e senha do Hub SCPMA (ex.: nome.sobrenome)'
  $c = Get-Credential -Message 'Usuário e senha do Hub SCPMA (os mesmos do Conecta SIPOFE)'
  if (-not $c) { throw 'O login foi cancelado.' }
  $u = $c.UserName.TrimStart('\').Trim()
  if ($u -match '@') { $u = $u.Split('@')[0] }
  $cred = New-Object System.Management.Automation.PSCredential($u, $c.Password)
  $body = @{ email = ($u.ToLower() + '@scpma.local'); password = $cred.GetNetworkCredential().Password } | ConvertTo-Json
  try {
    $tok = Invoke-RestMethod -Method Post -Uri "$Url/auth/v1/token?grant_type=password" -Headers @{ apikey = $Key } -ContentType 'application/json' -Body $body
  } catch { throw 'Usuário ou senha incorretos (ou sem internet). Rode o instalador de novo.' }
  $t = Invoke-WebRequest -UseBasicParsing -Uri "$Url/rest/v1/sipofe_docs?select=id&limit=1" -Headers @{ apikey = $Key; Authorization = 'Bearer ' + $tok.access_token }
  if ($t.Content.Trim() -eq '[]') { throw 'Este usuário não tem acesso aos dados do Conecta SIPOFE.' }
  $cred | Export-Clixml (Join-Path $sis 'credencial.xml')
  Write-Host 'Usuário conferido e salvo com segurança.' -ForegroundColor Green

  # 3. Programa de backup e atalhos
  [System.IO.File]::WriteAllText((Join-Path $sis 'backup.ps1'), $BACKUP_PS1, $Bom)
  [System.IO.File]::WriteAllText((Join-Path $base 'Fazer backup agora.cmd'), "@echo off`r`necho Fazendo o backup do Conecta SIPOFE...`r`npowershell -NoProfile -ExecutionPolicy Bypass -File `"%~dp0sistema\backup.ps1`"`r`necho.`r`ntype `"%~dp0ultimo-backup.txt`"`r`necho.`r`npause`r`n", [System.Text.Encoding]::ASCII)
  [System.IO.File]::WriteAllText((Join-Path $sis 'desinstalar.cmd'), "@echo off`r`nschtasks /Delete /TN `"Backup Conecta SIPOFE`" /F`r`necho O backup automatico foi desligado. Os arquivos de backup continuam na pasta.`r`npause`r`n", [System.Text.Encoding]::ASCII)
  $leia = @"
CONECTA SIPOFE · BACKUP AUTOMÁTICO

Todo dia às 13h (ou logo que o computador for ligado, se estiver desligado nesse horário), o Windows faz uma cópia de todos os dados do Conecta SIPOFE nesta pasta. Na hora do backup, uma janela preta pode piscar por um segundo: é normal.

O QUE TEM AQUI
- backups\        Cópias completas (.json), uma por dia. Ficam as dos últimos 30 dias e a primeira de cada mês mais antigo.
- planilhas\      Os dados atuais para consulta no Excel: contatos, pontos focais GECEG, controle de FCs, entrevistas e agenda.
- ultimo-backup.txt   Quando foi o último backup e quantos registros ele tem.
- Fazer backup agora.cmd   Faz um backup na hora, com dois cliques.
- sistema\        Arquivos do backup. Não precisa mexer.

COMO RESTAURAR
No Conecta SIPOFE, como administradora: tela inicial > Exportar > Importar backup (.json) e escolha o arquivo da pasta backups.

SE APARECER "BACKUP COM ERRO - LEIA.txt"
Abra o arquivo para ver o motivo. O mais comum é a senha do Hub SCPMA ter mudado: nesse caso, rode o instalador de novo.

PARA DESLIGAR O BACKUP AUTOMÁTICO
Dê dois cliques em sistema\desinstalar.cmd. Os arquivos já salvos continuam na pasta.
"@
  [System.IO.File]::WriteAllText((Join-Path $base 'LEIA-ME.txt'), $leia, $Bom)

  # 4. Agendar todo dia às 13h, para este usuário do Windows
  $act = New-ScheduledTaskAction -Execute 'powershell.exe' -Argument ('-NoProfile -ExecutionPolicy Bypass -WindowStyle Hidden -File "' + (Join-Path $sis 'backup.ps1') + '"')
  $trg = New-ScheduledTaskTrigger -Daily -At 13:00
  $set = New-ScheduledTaskSettingsSet -StartWhenAvailable -AllowStartIfOnBatteries -DontStopIfGoingOnBatteries -ExecutionTimeLimit (New-TimeSpan -Minutes 15)
  Register-ScheduledTask -TaskName 'Backup Conecta SIPOFE' -Description 'Cópia diária dos dados do Conecta SIPOFE na pasta projetos\CONECTA SIPOFE' -Action $act -Trigger $trg -Settings $set -Force | Out-Null
  Write-Host 'Backup agendado para todo dia às 13h.' -ForegroundColor Green

  # 5. Primeiro backup agora
  Titulo 'Fazendo o primeiro backup...'
  & powershell.exe -NoProfile -ExecutionPolicy Bypass -File (Join-Path $sis 'backup.ps1')
  if (Test-Path (Join-Path $base 'ultimo-backup.txt')) { Write-Host ''; Get-Content (Join-Path $base 'ultimo-backup.txt') -Encoding UTF8 | Write-Host; Write-Host ''; Write-Host 'Tudo pronto!' -ForegroundColor Green }
  else { Write-Host 'O primeiro backup não deu certo. Veja o arquivo "BACKUP COM ERRO - LEIA.txt" na pasta.' -ForegroundColor Yellow }
  Start-Process explorer.exe $base
} catch {
  Write-Host ''
  Write-Host ('Não foi possível concluir: ' + $_.Exception.Message) -ForegroundColor Red
}
