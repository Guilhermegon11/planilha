#requires -Version 5.1
param([switch]$SelfTest,[string]$WorkbookPath)
Set-StrictMode -Version 2.0
$ErrorActionPreference = 'Stop'
Import-Module (Join-Path $PSScriptRoot 'Juridico.Core.psm1') -Force
if ($SelfTest) { Invoke-CoreSelfTest; Write-Output 'SELFTEST OK - sem Excel, chave ou rede.'; exit 0 }
if ($env:OS -ne 'Windows_NT') { throw 'O assistente requer Windows e Excel desktop. Use -SelfTest para validar somente as funções puras.' }
Add-Type -AssemblyName System.Windows.Forms,System.Drawing,System.Security
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
[Windows.Forms.Application]::EnableVisualStyles()
$script:DataDir = Join-Path $env:LOCALAPPDATA 'JuridicoIA'
$null = New-Item -ItemType Directory -Force -Path $script:DataDir
$script:KeyPath = Join-Path $script:DataDir 'groq-key.dpapi'
$script:Wb = $null; $script:Excel = $null; $script:Central = $null
$script:Busy = $false; $script:Pending = @(); $script:Preview = @{}
$script:Ledger = @{}
$script:FinancialBefore = @{}
$script:CaseFamilies = @{}
$script:PrivateFonts=$null
$script:Context = $null; $script:BatchHash = ''; $script:Key = $null
$script:Cells = @{command='C6';connection='G6';case='C8';source='H8';office='C10';company='H10';consortium='C12';mode='C14';text='A17';message='A24';lastRun='H14'}
$script:Ranges = @{Processos='A10:U509';Valores='A10:K509';Andamentos='A10:F2009';'Ações'='A10:G1009'}
$script:Limits = @{Processos=509;Valores=509;Andamentos=2009;'Ações'=1009}
$script:Widths = @{Processos=21;Valores=11;Andamentos=6;'Ações'=7}

function Write-LocalLog([string]$Kind,[string]$Message) {
    # Não grava texto processual, resposta da API, URL livre, cabeçalhos ou chave.
    $safe = $Message -replace 'gsk_[A-Za-z0-9_-]+','[CHAVE REMOVIDA]'
    if ($safe.Length -gt 500) { $safe=$safe.Substring(0,500) }
    Add-Content -LiteralPath (Join-Path $script:DataDir 'assistente.log') -Encoding UTF8 -Value ('{0} [{1}] {2}' -f [datetime]::Now.ToString('s'),$Kind,$safe)
}
function Set-Status([string]$Message,[bool]$ErrorStatus=$false) {
    $script:StatusLabel.Text = $Message
    $script:StatusLabel.ForeColor = if($ErrorStatus){[Drawing.Color]::Firebrick}else{[Drawing.Color]::FromArgb(25,65,70)}
    if ($script:Central) { try { $script:Central.Range($script:Cells.message).Value2 = ConvertTo-ExcelLiteral $Message } catch {} }
}
function Read-Central([string]$Name) { return [string]$script:Central.Range($script:Cells[$Name]).Value2 }
function Write-Central([string]$Name,[object]$Value) { $script:Central.Range($script:Cells[$Name]).Value2 = ConvertTo-ExcelLiteral $Value }
function Read-Block([string]$Sheet) { return ,($script:Wb.Worksheets.Item($Sheet).Range($script:Ranges[$Sheet]).Value2) }
function Is-Blank([object]$Value) { return $null -eq $Value -or ($Value -is [string] -and $Value.Length -eq 0) }
function Get-ColumnNumber([string]$Letter) { return ([int][char]$Letter)-64 }
function Get-CaseIndex([object]$Block,[string]$Sheet) {
    $index=@{}
    for($r=1;$r -le $Block.GetLength(0);$r++) {
        $n=$Block[$r,1]
        if(Is-Blank $n) { continue }
        if($n -isnot [string]) { throw "Há identificador numérico em $Sheet, linha $($r+9). Converta a partir do número original para texto; não use o número arredondado." }
        $k=Get-CaseKey $n
        if($index.ContainsKey($k)){ throw "Identificador duplicado em $Sheet. Resolva a duplicação antes de usar a IA." }
        $index[$k]=[pscustomobject]@{Row=$r+9;Number=[string]$n}
    }
    return $index
}
function Get-EmptyRow([string]$Sheet,[object]$Block,[hashtable]$Reserved) {
    for($r=1;$r -le $Block.GetLength(0);$r++) {
        if($Reserved.ContainsKey("${Sheet}:$($r+9)")){continue}
        $empty=$true
        for($c=1;$c -le $script:Widths[$Sheet];$c++){if(-not(Is-Blank $Block[$r,$c])){$empty=$false;break}}
        if($empty){$Reserved["${Sheet}:$($r+9)"]=$true;return $r+9}
    }
    throw "A capacidade da aba $Sheet foi atingida. Amplie o modelo antes de continuar."
}
function Get-Context {
    $ctx=[ordered]@{}
    foreach($k in @('case','source','office','company','consortium')){$ctx[$k]=(Read-Central $k).Trim()}
    foreach($k in @('case','source','office','company','consortium')){if($ctx[$k].Length -gt 500){throw 'Um campo de contexto excede 500 caracteres. Use somente o nome ou identificador.'}}
    $ctx['text']=Read-Central 'text'
    if([string]::IsNullOrWhiteSpace($ctx.text)){throw 'Cole um texto ou importe um arquivo na Central IA.'}
    if($ctx.text.Length -gt 24000){throw 'O texto excede 24.000 caracteres. Divida a comunicação em lotes menores.'}
    if([string]::IsNullOrWhiteSpace($ctx.source)){$ctx.source='Texto informado na Central IA'}
    return [pscustomobject]$ctx
}
function Get-ContextHash([object]$Context) { return Get-TextHash ('campos-revisados-v1|' + ($Context | ConvertTo-Json -Compress -Depth 5)) }
function Get-SafeDisplay([object]$Value,[string]$Type='text') {
    if(Is-Blank $Value){return '(em branco)'}
    if($Type -eq 'date'){return ([datetime]::FromOADate([double]$Value)).ToString('dd/MM/yyyy')}
    if($Type -eq 'probability'){return ([double]$Value).ToString('P1',[Globalization.CultureInfo]::GetCultureInfo('pt-BR'))}
    if($Type -in @('money','balance')){return ([double]$Value).ToString('C2',[Globalization.CultureInfo]::GetCultureInfo('pt-BR'))}
    return [string]$Value
}
function Get-Key {
    if($script:Key){return $script:Key}
    if(Test-Path -LiteralPath $script:KeyPath){
        try{$script:Key=Get-Content -LiteralPath $script:KeyPath -Raw | ConvertTo-SecureString;return $script:Key}
        catch{throw 'Não foi possível ler a chave protegida deste usuário. Clique em Trocar chave.'}
    }
    Show-KeyDialog
    if(-not $script:Key){throw 'Informe uma chave Groq para analisar o texto.'}
    return $script:Key
}
function Show-KeyDialog {
    $dialog=New-Object Windows.Forms.Form
    $dialog.Text='Conectar à Groq';$dialog.Size=New-Object Drawing.Size(570,270);$dialog.StartPosition='CenterParent';$dialog.Font=$script:UiFont
    $info=New-Object Windows.Forms.Label;$info.SetBounds(20,18,515,65)
    $info.Text="Cole uma chave nova da Groq. Ela fica protegida para seu usuário do Windows, fora do Excel. A análise envia somente o texto e o contexto escolhidos."
    $box=New-Object Windows.Forms.TextBox;$box.SetBounds(20,94,515,28);$box.UseSystemPasswordChar=$true
    $save=New-Object Windows.Forms.Button;$save.Text='Salvar chave';$save.SetBounds(365,145,170,38)
    $save.Add_Click({
        if($box.Text.Trim().Length -lt 15){[Windows.Forms.MessageBox]::Show('Cole a chave completa.');return}
        try{
            $script:Key=ConvertTo-SecureString $box.Text.Trim() -AsPlainText -Force
            $script:Key | ConvertFrom-SecureString | Set-Content -LiteralPath $script:KeyPath -Encoding ASCII
            $box.Clear();$dialog.DialogResult=[Windows.Forms.DialogResult]::OK;$dialog.Close()
        }catch{[Windows.Forms.MessageBox]::Show('Não foi possível proteger a chave neste usuário do Windows.');$script:Key=$null}
    })
    $dialog.Controls.AddRange(@($info,$box,$save));$null=$dialog.ShowDialog($script:Form);$dialog.Dispose()
}
function Invoke-GroqExtraction([object]$Context) {
    $secure=Get-Key
    $ptr=[Runtime.InteropServices.Marshal]::SecureStringToBSTR($secure)
    try{$plain=[Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr)}finally{[Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr)}
    $body=@{model='openai/gpt-oss-20b';temperature=0;max_completion_tokens=12000;reasoning_effort='low';messages=@(
        @{role='system';content=(Get-ExtractionPrompt)},
        @{role='user';content=($Context | ConvertTo-Json -Depth 6 -Compress)}
    );response_format=@{type='json_schema';json_schema=@{name='atualizacao_juridica';strict=$true;schema=(Get-ExtractionSchema)}}} | ConvertTo-Json -Depth 40 -Compress
    try{
        for($attempt=1;$attempt -le 3;$attempt++){
            try{
                $result=Invoke-RestMethod -Method Post -Uri 'https://api.groq.com/openai/v1/chat/completions' -UserAgent 'ControleJuridico/4.0' -Headers @{Authorization="Bearer $plain"} -ContentType 'application/json; charset=utf-8' -Body ([Text.Encoding]::UTF8.GetBytes($body)) -TimeoutSec 80
                if(-not $result.choices -or $result.choices.Count -ne 1){throw 'A Groq não retornou uma resposta única.'}
                if($result.choices[0].finish_reason -ne 'stop'){throw 'A resposta da IA foi interrompida. Divida o texto e analise novamente.'}
                $data=$result.choices[0].message.content | ConvertFrom-Json
                $null=Test-ExtractionResult $data
                if($null -eq $data -or -not $data.PSObject.Properties['casos'] -or -not $data.PSObject.Properties['avisos']){throw 'Resposta da IA sem a estrutura esperada.'}
                if(@($data.casos).Count -gt 20){throw 'Foram identificados mais de 20 processos. Divida o lote.'}
                return $data
            }catch{
                $status=0
                if($_.Exception.PSObject.Properties['Response'] -and $_.Exception.Response){try{$status=[int]$_.Exception.Response.StatusCode}catch{}}
                if($status -eq 401){throw 'A Groq recusou a chave (401). Clique em Trocar chave e informe uma chave válida.'}
                if($status -eq 429 -and $attempt -lt 3){
                    $delay=$attempt*3
                    try{
                        $retry=[string]$_.Exception.Response.Headers['Retry-After'];$seconds=0
                        if([int]::TryParse($retry,[ref]$seconds)){$delay=[Math]::Max($delay,$seconds)}
                        else{$retryDate=[datetime]::MinValue;if([datetime]::TryParse($retry,[ref]$retryDate)){$delay=[Math]::Max($delay,[int][Math]::Ceiling(($retryDate.ToUniversalTime()-[datetime]::UtcNow).TotalSeconds))}}
                    }catch{}
                    if($delay -gt 30){throw 'A Groq solicitou uma espera superior a 30 segundos. Aguarde e tente novamente; nenhuma proposta foi aplicada.'}
                    Set-Status "Limite temporário da Groq. Nova tentativa $($attempt+1) de 3 em $delay segundos."
                    Start-Sleep -Seconds $delay;continue
                }
                if($status -eq 429){throw 'Limite da Groq atingido (429) após 3 tentativas. Aguarde e tente novamente.'}
                if($status -ge 400){throw "A Groq retornou erro HTTP $status. Nenhum dado foi aplicado. Verifique a conta, o modelo e a conexão."}
                if($_.Exception -is [Net.WebException]){throw 'Falha de conexão ou tempo esgotado na Groq. Nenhum dado foi aplicado.'}
                throw
            }
        }
    }finally{$plain=$null;$body=$null}
}
function Load-Ledger {
    $script:Ledger=@{}
    if(Test-Path -LiteralPath $script:LedgerPath){
        try{
            $obj=Get-Content -LiteralPath $script:LedgerPath -Raw | ConvertFrom-Json
            foreach($p in $obj.PSObject.Properties){$script:Ledger[$p.Name]=[string]$p.Value}
        }catch{throw 'O registro local de aplicações está ilegível. Restaure o arquivo de registro antes de aplicar para evitar duplicações.'}
    }
}
function Save-Ledger([hashtable]$Data) {
    $temp=$script:LedgerPath+'.tmp'
    $Data | ConvertTo-Json -Depth 8 -Compress | Set-Content -LiteralPath $temp -Encoding UTF8
    Move-Item -LiteralPath $temp -Destination $script:LedgerPath -Force
}
function New-Proposal([string]$Case,[string]$Operation,[string]$Sheet,[string]$Field,[string]$Column,[object]$Before,[object]$Value,[string]$Evidence,[string]$Warning,[string]$Type='text') {
    if($Operation -notin @('CreateCase','SetField','AppendHistory') -or $Sheet -eq 'Ações'){throw 'Operação indisponível nesta versão.'}
    $stable=if($Operation -eq 'AppendHistory'){[string]$Value[1]}
    else{($Value | ConvertTo-Json -Compress -Depth 8)}
    $signature=Get-TextHash ($script:BatchHash+'|'+(Get-CaseKey $Case)+'|'+$Operation+'|'+$Field+'|'+$stable)
    return [pscustomobject]@{Id=[guid]::NewGuid().ToString('N');Case=$Case;Key=(Get-CaseKey $Case);Operation=$Operation;Sheet=$Sheet;Field=$Field;Column=$Column;Before=$Before;Value=$Value;Evidence=$Evidence;Warning=$Warning;Type=$Type;Signature=$signature;Applied=$false;PreviewRow=0;NewCase=$false}
}
function Build-Proposals([object]$Data,[object]$Context) {
    $pb=Read-Block 'Processos';$vb=Read-Block 'Valores'
    $hb=Read-Block 'Andamentos'
    $pi=Get-CaseIndex $pb 'Processos';$vi=Get-CaseIndex $vb 'Valores'
    $existingNumbers=@($pi.Values | ForEach-Object {$_.Number})
    $script:CaseFamilies=@{}
    foreach($existing in $existingNumbers){
        $baseKey=Get-CaseBaseKey $existing
        if(-not $script:CaseFamilies.ContainsKey($baseKey)){$script:CaseFamilies[$baseKey]=@()}
        $script:CaseFamilies[$baseKey]+=(Get-CaseKey $existing)
    }
    $script:FinancialBefore=@{}
    foreach($k in $vi.Keys){$script:FinancialBefore[$k]=@($vb[($vi[$k].Row-9),10],$vb[($vi[$k].Row-9),11])}
    $out=New-Object System.Collections.Generic.List[object]
    $warnings=New-Object System.Collections.Generic.List[string]
    $seen=@{};$knownCount=@($Data.casos).Count
    foreach($item in @($Data.casos)){
        $number=[string]$item.numero
        if([string]::IsNullOrWhiteSpace($number)){
            if($knownCount -eq 1 -and $Context.case){$number=$Context.case}else{$warnings.Add('Registro sem identificador inequívoco ignorado.');continue}
        }
        $number=$number.Trim()
        if($number -match '(?i)^(?:processo|proc\.?|n[º°o]\.?|n[uú]mero)\s*[:#]?\s+' -or $number -match '[;,:]$'){
            $warnings.Add('A IA incluiu um rótulo ou pontuação no identificador. Escreva o número isolado na fonte ou selecione o processo e analise novamente.');continue
        }
        try{$resolved=Resolve-CatalogCase $number $Context.case $Context.text $existingNumbers $knownCount}
        catch{$warnings.Add($_.Exception.Message);continue}
        $number=$resolved.Number;$key=$resolved.Key;$isNew=$resolved.IsNew
        if($seen.ContainsKey($key)){throw 'A IA repetiu um processo no lote. Analise um trecho menor.'};$seen[$key]=$true
        if($isNew){
            if(-not $Context.office -or -not $Context.company){$warnings.Add("Processo $number não cadastrado: informe escritório e empresa e mantenha o identificador explícito no texto.");continue}
            $base=[ordered]@{A=$number;D=$Context.office;E=$Context.company;F=$Context.consortium}
            $p=New-Proposal $number 'CreateCase' 'Processos' 'Cadastrar processo e contexto' 'A' $null $base $number 'Novo cadastro. Confira o número, o escritório e a empresa.'
            $p.NewCase=$true;$out.Add($p)
        }else{$number=$pi[$key].Number}
        $fieldsSeen=@{}
        foreach($field in @($item.campos)){
            $name=[string]$field.campo
            if($fieldsSeen.ContainsKey($name)){throw 'A resposta contém o mesmo campo duas vezes para um processo.'};$fieldsSeen[$name]=$true
            if([string]::IsNullOrWhiteSpace([string]$field.valor)){continue}
            try{$v=ConvertTo-ValidatedField $name ([string]$field.valor) ([string]$field.evidencia) $Context.text}
            catch{$warnings.Add("Campo rejeitado para ${number}: $($_.Exception.Message)");continue}
            $before=$null
            if($v.Sheet -eq 'Processos' -and -not $isNew){$before=$pb[($pi[$key].Row-9),(Get-ColumnNumber $v.Column)]}
            if($v.Sheet -eq 'Valores' -and $vi.ContainsKey($key)){$before=$vb[($vi[$key].Row-9),(Get-ColumnNumber $v.Column)]}
            if(Test-ValueEqual $before $v.Value){continue}
            # Contexto do novo cadastro é escolhido pela pessoa; a IA não o substitui.
            if($isNew -and $v.Column -in @('D','E','F') -and $v.Sheet -eq 'Processos'){continue}
            $warning=if($v.Sheet -eq 'Valores'){'Confira valor, categoria e abrangência. A revisão financeira ficará datada de hoje.'}else{'Confira a evidência antes de aplicar.'}
            $p=New-Proposal $number 'SetField' $v.Sheet $v.Label $v.Column $before $v.Value ([string]$field.evidencia) $warning $v.Type
            $p.NewCase=$isNew;$out.Add($p)
        }
        if($item.data_andamento -and $item.resumo){
            try{
                if(-not(Test-SourceQuote $Context.text ([string]$item.evidencia_andamento))){throw 'Trecho de andamento ausente da fonte.'}
                $eventDate=Get-ExplicitDate ([string]$item.data_andamento) ([string]$item.evidencia_andamento)
                if(([string]$item.resumo).Length -gt 6000){throw 'Resumo muito longo.'}
                $values=@($number,$eventDate.ToOADate(),[string]$item.resumo,[datetime]::Now.ToOADate(),$Context.source,$Context.text)
                $p=New-Proposal $number 'AppendHistory' 'Andamentos' 'Acrescentar andamento' '' $null $values ([string]$item.evidencia_andamento) 'Acrescenta histórico; preserva o texto original em Observações.'
                $p.NewCase=$isNew;$out.Add($p)
            }catch{$warnings.Add("Andamento não proposto para ${number}: $($_.Exception.Message)")}
        }

    }
    $filtered=New-Object System.Collections.Generic.List[object]
    $batchSignatures=@{}
    foreach($p in $out){
        if($batchSignatures.ContainsKey($p.Signature)){
            $warnings.Add('Uma proposta repetida no mesmo lote foi removida.');continue
        }
        $batchSignatures[$p.Signature]=$p.Field
        $already=$false
        if($p.Operation -eq 'AppendHistory'){
            for($r=1;$r -le $hb.GetLength(0);$r++){
                if(-not(Is-Blank $hb[$r,1]) -and (Get-CaseKey ([string]$hb[$r,1])) -ceq $p.Key -and (Test-ValueEqual $hb[$r,2] $p.Value[1]) -and (Test-ValueEqual $hb[$r,5] $p.Value[4]) -and (Test-ValueEqual $hb[$r,6] $p.Value[5])){$already=$true;break}
            }

        }
        if($already){$warnings.Add("Registro já presente para $($p.Case); não será duplicado.");continue}
        if($script:Ledger.ContainsKey($p.Signature)){$warnings.Add("Proposta já registrada como aplicada para $($p.Case). Se restaurou um backup ou apagou o registro, revise o histórico antes de reprocessar.");continue}
        $filtered.Add($p)
    }
    $script:AnalysisWarnings=@($warnings.ToArray())
    return $filtered.ToArray()
}
function Show-Preview {
    if($script:Pending.Count -gt 200){throw 'Mais de 200 propostas. Divida o texto em lotes menores; nenhuma proposta foi aplicada.'}
    $script:Central.Range('A29:J228').ClearContents();$script:Preview=@{}
    $r=29
    foreach($p in $script:Pending){
        $p.PreviewRow=$r
        $proposed=if($p.Operation -eq 'CreateCase'){'Número: '+$p.Case+' | Escritório: '+$p.Value.D+' | Empresa: '+$p.Value.E+' | Consórcio: '+$p.Value.F}
        elseif($p.Operation -eq 'AppendHistory'){(Get-SafeDisplay $p.Value[1] 'date')+' | '+$p.Value[2]}
        else{Get-SafeDisplay $p.Value $p.Type}
        $values=@('Não',$p.Case,$p.Sheet,$p.Field,(Get-SafeDisplay $p.Before $p.Type),$proposed,$p.Evidence,$p.Warning,$p.Id,'')
        for($c=1;$c -le 10;$c++){$script:Central.Cells.Item($r,$c).Value2=ConvertTo-ExcelLiteral $values[$c-1]}
        $script:Preview[$p.Id]=@($values[1..8]);$r++
    }
    Write-Central 'lastRun' ([datetime]::Now.ToString('dd/MM/yyyy HH:mm'))
}
function Add-Write([System.Collections.Generic.List[object]]$List,[string]$Sheet,[int]$Row,[string]$Column,[object]$Value,[bool]$InitializeTotals=$false) {
    if($Sheet -eq 'Ações' -or ($Sheet -eq 'Processos' -and $Column -ceq 'R') -or ($Sheet -eq 'Valores' -and $Column -in @('F','I'))){throw 'Este campo foi retirado do controle e não aceita gravação pela integração.'}
    if($Row -lt 10 -or $Row -gt $script:Limits[$Sheet]){throw 'Linha de destino fora do contrato.'}
    $col=Get-ColumnNumber $Column
    $max=if($Sheet -eq 'Processos'){19}else{$script:Widths[$Sheet]}
    $totalsMetadata=$Sheet -eq 'Valores' -and $Column -ceq 'T' -and $InitializeTotals -and $Value -ceq 'Sim'
    if(($col -lt 1 -or $col -gt $max) -and -not $totalsMetadata){throw 'Coluna de destino fora do contrato.'}
    $address=$Column+$Row
    foreach($w in $List){if($w.Sheet -eq $Sheet -and $w.Address -eq $address){$w.Value=$Value;return}}
    $cell=$script:Wb.Worksheets.Item($Sheet).Range($address)
    if($totalsMetadata -and -not(Is-Blank $cell.Value2)){throw 'A inclusão nos totais já foi definida. O assistente não altera decisões existentes.'}
    if($cell.HasFormula){throw "O destino $Sheet!$address contém fórmula. Revise antes de aplicar."}
    $List.Add([pscustomobject]@{Sheet=$Sheet;Address=$address;Value=$Value;Before=$cell.Value2;Format=$cell.NumberFormat})
}
function New-Backup {
    $backup=Join-Path $script:BackupDir ('Controle_Juridico_'+[datetime]::Now.ToString('yyyyMMdd_HHmmss_fff')+'.xlsx')
    $script:Wb.SaveCopyAs($backup)
    return $backup
}
function Commit-Writes([System.Collections.Generic.List[object]]$Writes,[object[]]$Proposals) {
    if($Writes.Count -eq 0){return}
    if($script:Wb.ReadOnly){throw 'A planilha está em modo somente leitura.'}
    $backup=New-Backup
    $oldEvents=$script:Excel.EnableEvents;$oldScreen=$script:Excel.ScreenUpdating
    $ledgerBefore=$script:Ledger.Clone();$completed=New-Object System.Collections.Generic.List[object]
    try{
        $script:Excel.EnableEvents=$false;$script:Excel.ScreenUpdating=$false
        foreach($w in $Writes){
            $cell=$script:Wb.Worksheets.Item($w.Sheet).Range($w.Address)
            if($cell.HasFormula){throw 'Uma célula de destino passou a conter fórmula. Operação cancelada.'}
            if(-not(Test-ValueEqual $cell.Value2 $w.Before)){throw 'Uma célula de destino mudou durante a aplicação. Operação cancelada.'}
            $completed.Add($w)
            if($w.Address -match '^A\d+$'){$cell.NumberFormat='@'}
            $cell.Value2=ConvertTo-ExcelLiteral $w.Value
        }
        $script:Excel.Calculate();$script:Wb.Save()
        foreach($p in $Proposals){$script:Ledger[$p.Signature]=[datetime]::Now.ToString('s')}
        Save-Ledger $script:Ledger
        foreach($p in $Proposals){$p.Applied=$true}
        Write-LocalLog 'APLICADO' ("Células: $($Writes.Count); propostas: $($Proposals.Count); backup: "+[IO.Path]::GetFileName($backup))
    }catch{
        $reason=$_.Exception.Message;$rollbackOk=$true
        for($i=$completed.Count-1;$i -ge 0;$i--){try{
            $w=$completed[$i];$cell=$script:Wb.Worksheets.Item($w.Sheet).Range($w.Address)
            $cell.Value2=ConvertTo-ExcelLiteral $w.Before;$cell.NumberFormat=$w.Format
        }catch{$rollbackOk=$false}}
        $script:Ledger=$ledgerBefore
        try{Save-Ledger $script:Ledger;$script:Excel.Calculate();$script:Wb.Save()}catch{$rollbackOk=$false}
        if($rollbackOk){throw "Aplicação desfeita: $reason Backup: $backup"}
        throw "Não foi possível concluir a restauração. Feche sem salvar e restaure o backup: $backup"
    }finally{$script:Excel.EnableEvents=$oldEvents;$script:Excel.ScreenUpdating=$oldScreen}
}
function Apply-Proposals([object[]]$Selection,[bool]$CheckPreview=$true) {
    if($Selection.Count -eq 0){Set-Status 'Marque Sim na coluna Aplicar? para as propostas desejadas.';return}
    foreach($p in $Selection){if($p.Operation -notin @('CreateCase','SetField','AppendHistory') -or $p.Sheet -eq 'Ações'){throw 'A prévia contém uma operação retirada desta versão. Analise novamente.'}}
    if((Get-ContextHash (Get-Context)) -cne $script:BatchHash){throw 'O texto ou contexto mudou depois da prévia. Analise novamente.'}
    $blocks=@{};foreach($s in @('Processos','Valores','Andamentos')){$blocks[$s]=Read-Block $s}
    $pi=Get-CaseIndex $blocks.Processos 'Processos';$vi=Get-CaseIndex $blocks.Valores 'Valores'
    $reserved=@{};$newRows=@{};$valueRows=@{};$financialRows=@{};$writes=New-Object System.Collections.Generic.List[object]
    foreach($p in $Selection){
        if($p.Applied -or $script:Ledger.ContainsKey($p.Signature)){throw 'Uma proposta já foi aplicada. Analise novamente para atualizar a prévia.'}
        if($CheckPreview){
            $snapshot=$script:Preview[$p.Id]
            for($c=2;$c -le 9;$c++){
                if(-not(Test-ValueEqual $script:Central.Cells.Item($p.PreviewRow,$c).Value2 $snapshot[$c-2])){throw 'O conteúdo da prévia foi alterado. Somente a coluna Aplicar? aceita edição; analise novamente.'}
            }
        }
        $baseKey=Get-CaseBaseKey $p.Case
        $currentFamily=@($pi.Values | Where-Object {(Get-CaseBaseKey $_.Number) -ceq $baseKey} | ForEach-Object {Get-CaseKey $_.Number} | Sort-Object)
        $previousFamily=if($script:CaseFamilies.ContainsKey($baseKey)){@($script:CaseFamilies[$baseKey] | Sort-Object)}else{@()}
        if(($currentFamily -join "`n") -cne ($previousFamily -join "`n")){throw 'Os registros que compartilham este número mudaram depois da prévia. Analise novamente e confira a seleção com [ID].'}
        if($p.Operation -eq 'CreateCase'){
            if($currentFamily.Count){throw 'Este número já possui cadastro ou variantes com [ID]. Selecione o registro completo; a base não será cadastrada novamente.'}
            $row=Get-EmptyRow 'Processos' $blocks.Processos $reserved;$newRows[$p.Key]=$row
            foreach($col in $p.Value.Keys){Add-Write $writes 'Processos' $row $col $p.Value[$col]}
        }
    }
    foreach($p in $Selection){
        if($p.Operation -eq 'CreateCase'){continue}
        if($p.NewCase -and -not $newRows.ContainsKey($p.Key)){throw 'Aprove também o cadastro do processo novo antes de aplicar seus registros.'}
        if(-not $p.NewCase -and -not $pi.ContainsKey($p.Key)){throw 'O identificador original mudou ou foi removido após a prévia. Analise novamente.'}
        if(-not $p.NewCase -and $pi[$p.Key].Number -cne $p.Case){throw 'A escrita do identificador mudou depois da análise. Analise novamente.'}
        if($p.Operation -eq 'SetField'){
            if($p.Sheet -eq 'Processos'){$row=if($p.NewCase){$newRows[$p.Key]}else{$pi[$p.Key].Row}}
            else{
                if($vi.ContainsKey($p.Key)){
                    if($vi[$p.Key].Number -cne $p.Case){throw 'O número em Valores está escrito de modo diferente de Processos. Padronize o identificador antes de aplicar.'}
                    $row=$vi[$p.Key].Row
                }
                elseif($valueRows.ContainsKey($p.Key)){$row=$valueRows[$p.Key]}
                else{
                    $row=Get-EmptyRow 'Valores' $blocks.Valores $reserved;$valueRows[$p.Key]=$row;Add-Write $writes 'Valores' $row 'A' $p.Case
                    # T é metadado determinístico; nunca um campo escolhido pela IA.
                    if(Is-Blank $script:Wb.Worksheets.Item('Valores').Range('T'+$row).Value2){Add-Write $writes 'Valores' $row 'T' 'Sim' $true}
                }
                $financialRows[$p.Key]=$row
            }
            $current=$script:Wb.Worksheets.Item($p.Sheet).Range($p.Column+$row).Value2
            if(-not(Test-ValueEqual $current $p.Before)){throw "Conflito em $($p.Case), $($p.Field). O valor mudou desde a prévia; analise novamente."}
            Add-Write $writes $p.Sheet $row $p.Column $p.Value
        }elseif($p.Operation -eq 'AppendHistory'){
            $row=Get-EmptyRow $p.Sheet $blocks[$p.Sheet] $reserved
            for($c=0;$c -lt $p.Value.Count;$c++){Add-Write $writes $p.Sheet $row ([string][char](65+$c)) $p.Value[$c]}
        }else{throw 'Operação não permitida.'}
    }
    foreach($key in $financialRows.Keys){
        $row=$financialRows[$key]
        $meta=if($script:FinancialBefore.ContainsKey($key)){$script:FinancialBefore[$key]}else{@($null,$null)}
        foreach($column in @('J','K')){
            if(@($Selection | Where-Object {$_.Key -ceq $key -and $_.Sheet -eq 'Valores' -and $_.Column -eq $column}).Count){continue}
            if(@($script:Pending | Where-Object {$_.Applied -and $_.Key -ceq $key -and $_.Sheet -eq 'Valores' -and $_.Column -eq $column}).Count){continue}
            $offset=if($column -eq 'J'){0}else{1}
            $current=$script:Wb.Worksheets.Item('Valores').Range($column+$row).Value2
            if(-not(Test-ValueEqual $current $meta[$offset])){throw 'Os metadados financeiros mudaram depois da prévia. Analise novamente.'}
            $value=if($column -eq 'J'){[datetime]::Today.ToOADate()}else{$script:Context.source}
            Add-Write $writes 'Valores' $row $column $value
        }
    }
    Commit-Writes $writes $Selection
    # Cadastro e alterações podem ser aprovados em etapas na mesma sessão.
    foreach($created in @($Selection | Where-Object {$_.Operation -eq 'CreateCase'})){
        foreach($remaining in $script:Pending){if($remaining.Key -ceq $created.Key){$remaining.NewCase=$false}}
        $script:CaseFamilies[(Get-CaseBaseKey $created.Case)]=@($created.Key)
    }
    foreach($key in $financialRows.Keys){
        $row=$financialRows[$key]
        $script:FinancialBefore[$key]=@($script:Wb.Worksheets.Item('Valores').Range('J'+$row).Value2,$script:Wb.Worksheets.Item('Valores').Range('K'+$row).Value2)
    }
    foreach($p in $Selection){
        $script:Central.Cells.Item($p.PreviewRow,1).Value2='Não'
        $script:Central.Cells.Item($p.PreviewRow,10).Value2=[datetime]::Now.ToString('dd/MM/yyyy HH:mm:ss')
    }
    $script:Wb.Save()
    Set-Status "$($Selection.Count) proposta(s) aplicada(s). Backup e registro local criados. As demais continuam aguardando sua escolha."
}
function Analyze-Text {
    $script:Context=Get-Context;$script:BatchHash=Get-ContextHash $script:Context
    $cacheFile=Join-Path $script:CacheDir ($script:BatchHash+'.dpapi')
    if(Test-Path -LiteralPath $cacheFile){
        Set-Status 'Reutilizando a análise protegida deste mesmo texto e contexto; verificando o que já foi aplicado.'
        $cached=Get-Content -LiteralPath $cacheFile -Raw | ConvertTo-SecureString
        $ptr=[Runtime.InteropServices.Marshal]::SecureStringToBSTR($cached)
        try{$data=[Runtime.InteropServices.Marshal]::PtrToStringBSTR($ptr) | ConvertFrom-Json}finally{[Runtime.InteropServices.Marshal]::ZeroFreeBSTR($ptr)}
        $null=Test-ExtractionResult $data
    }else{
        Set-Status 'Analisando somente o texto e o contexto escolhidos. Aguarde a Groq...'
        $data=Invoke-GroqExtraction $script:Context
        $json=$data | ConvertTo-Json -Depth 30 -Compress
        ConvertTo-SecureString $json -AsPlainText -Force | ConvertFrom-SecureString | Set-Content -LiteralPath $cacheFile -Encoding ASCII
        $json=$null
    }
    $script:Pending=@(Build-Proposals $data $script:Context)
    if($script:Pending.Count -eq 0 -and $script:AnalysisWarnings.Count -and (Test-Path -LiteralPath $cacheFile)){Remove-Item -LiteralPath $cacheFile -Force}
    Show-Preview
    $message="$($script:Pending.Count) proposta(s). Confira os trechos, marque Sim nas desejadas e escolha Aplicar aprovados."
    if($script:AnalysisWarnings.Count){$message+=' '+($script:AnalysisWarnings -join ' ');if($message.Length -gt 3000){$message=$message.Substring(0,3000)}}
    Set-Status $message
    if((Read-Central 'mode') -eq 'Andamentos automáticos'){
        $automatic=@($script:Pending | Where-Object {$_.Operation -eq 'AppendHistory' -and -not $_.NewCase})
        if($automatic.Count){Apply-Proposals $automatic $true}
    }
}
function Test-LocalImportPackage([string]$Path) {
    if($Path.StartsWith('\\')){throw 'Copie a origem para uma pasta local antes de importar.'}
    Add-Type -AssemblyName System.IO.Compression.FileSystem
    $zip=[IO.Compression.ZipFile]::OpenRead($Path)
    try{
        if($zip.Entries.Count -gt 5000){throw 'Arquivo de origem com estrutura excessivamente grande.'}
        $expanded=[long]0
        foreach($entry in $zip.Entries){
            $expanded+=$entry.Length
            if($expanded -gt 100MB){throw 'Arquivo de origem descompactado excede o limite de segurança.'}
            if($entry.FullName -match '(?i)^xl/(connections\.xml|queryTables/|externalLinks/|vbaProject\.bin|embeddings/|macrosheets/|richData/|webextensions/)'){throw 'A origem contém vínculos, consultas, macros ou objetos externos. Exporte um TXT/CSV ou XLSX somente com valores.'}
            if($entry.FullName -match '(?i)^xl/worksheets/[^/]+\.xml$'){
                $stream=$entry.Open();$reader=New-Object IO.StreamReader($stream)
                try{$xml=$reader.ReadToEnd()}finally{$reader.Dispose();$stream.Dispose()}
                if($xml -match '<(?:[A-Za-z0-9_]+:)?f(?:\s|>)'){throw 'A origem contém fórmulas. Para importar sem cálculos ou chamadas externas, salve uma cópia somente com valores ou um CSV/TXT.'}
            }
            if($entry.FullName.EndsWith('.rels')){
                $stream=$entry.Open();$reader=New-Object IO.StreamReader($stream)
                try{$relationships=$reader.ReadToEnd()}finally{$reader.Dispose();$stream.Dispose()}
                if($relationships -match '(?i)TargetMode\s*=\s*["'']External["'']'){throw 'A origem contém relações externas. Exporte um TXT/CSV ou XLSX somente com valores.'}
            }
        }
    }finally{$zip.Dispose()}
}
function Import-Source {
    $pick=New-Object Windows.Forms.OpenFileDialog;$pick.Title='Importar comunicação recebida';$pick.Filter='Textos e planilhas (*.txt;*.csv;*.xlsx)|*.txt;*.csv;*.xlsx';$pick.Multiselect=$false
    if($pick.ShowDialog($script:Form) -ne [Windows.Forms.DialogResult]::OK){$pick.Dispose();return}
    $file=$pick.FileName;$pick.Dispose();$extension=[IO.Path]::GetExtension($file).ToLowerInvariant()
    if($file.StartsWith('\\')){throw 'Copie a origem para uma pasta local antes de importar.'}
    if([IO.Path]::GetFullPath($file) -eq [IO.Path]::GetFullPath($script:Wb.FullName)){throw 'Escolha um arquivo de origem diferente da planilha de controle.'}
    if((Get-Item -LiteralPath $file).Length -gt 10MB){throw 'Arquivo maior que 10 MB. Exporte apenas o trecho relevante.'}
    if($extension -in @('.txt','.csv')){$text=[IO.File]::ReadAllText($file,[Text.Encoding]::UTF8)}
    elseif($extension -eq '.xlsx'){
        Test-LocalImportPackage $file
        $import=$null;$importExcel=$null
        try{
            $importExcel=New-Object -ComObject Excel.Application
            $importExcel.Visible=$false;$importExcel.DisplayAlerts=$false
            $import=Open-ControlWorkbook $importExcel $file $true
            $used=$import.Worksheets.Item(1).UsedRange
            if($used.Rows.Count -gt 500 -or $used.Columns.Count -gt 30){throw 'A origem excede 500 linhas ou 30 colunas. Exporte apenas o relatório necessário.'}
            $builder=New-Object Text.StringBuilder
            for($r=1;$r -le $used.Rows.Count;$r++){
                $line=New-Object System.Collections.Generic.List[string]
                for($c=1;$c -le $used.Columns.Count;$c++){$line.Add([string]$used.Cells.Item($r,$c).Text)}
                $null=$builder.AppendLine(($line -join ' | '))
                if($builder.Length -gt 24000){throw 'A origem excede 24.000 caracteres. Divida o relatório.'}
            }
            $text=$builder.ToString()
        }finally{
            Close-ControlAttemptWorkbook $import
            Close-ControlAttemptExcel $importExcel
            if($import -and [Runtime.InteropServices.Marshal]::IsComObject($import)){try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($import)}catch{}}
            if($importExcel -and [Runtime.InteropServices.Marshal]::IsComObject($importExcel)){try{[void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($importExcel)}catch{}}
        }
    }else{throw 'Formato não suportado. Use TXT, CSV ou XLSX.'}
    if($text.Length -gt 24000){throw 'O texto excede 24.000 caracteres. Divida a origem; nenhum trecho foi cortado.'}
    Write-Central 'text' $text;Write-Central 'source' ([IO.Path]::GetFileName($file))
    Set-Status 'Arquivo importado. Confira o processo e o contexto; escolha Analisar texto para enviá-lo à Groq.'
}
function Install-Poppins {
    try{
        $fontDir=Join-Path $PSScriptRoot 'Fontes'
        if(-not(Test-Path -LiteralPath $fontDir)){return}
        $dest=Join-Path $env:LOCALAPPDATA 'Microsoft\Windows\Fonts';$null=New-Item -ItemType Directory -Force -Path $dest
        $reg='HKCU:\Software\Microsoft\Windows NT\CurrentVersion\Fonts';$null=New-Item -Path $reg -Force
        foreach($font in Get-ChildItem -LiteralPath $fontDir -Filter '*.ttf'){
            $target=Join-Path $dest $font.Name
            if(-not(Test-Path -LiteralPath $target)){Copy-Item -LiteralPath $font.FullName -Destination $target}
            $null=New-ItemProperty -Path $reg -Name ($font.BaseName+' (TrueType)') -Value $target -PropertyType String -Force
        }
        # A família privada permite Poppins nesta janela mesmo antes de reabrir o Excel.
        $script:PrivateFonts=New-Object Drawing.Text.PrivateFontCollection
        foreach($font in Get-ChildItem -LiteralPath $fontDir -Filter '*.ttf'){$script:PrivateFonts.AddFontFile($font.FullName)}
    }catch{Write-LocalLog 'FONTE' 'Poppins não pôde ser instalada automaticamente; use os arquivos da pasta Fontes.'}
}
function Write-ControlConnectionLog([string]$Message) {
    try{Write-LocalLog 'CONEXAO' $Message}catch{}
}
function Close-ControlAttemptWorkbook([object]$Workbook) {
    if(-not $Workbook){return}
    try{
        # Só descartar uma abertura limpa. Uma edição feita no Excel fica aberta.
        if($Workbook.Saved){$Workbook.Close($false)}
        else{$Workbook.Application.Visible=$true;Write-ControlConnectionLog 'A planilha aberta na tentativa permaneceu visível porque contém alterações.'}
    }catch{Write-ControlConnectionLog 'Não foi possível encerrar a planilha da tentativa; ela permaneceu disponível no Excel.'}
}
function Close-ControlAttemptExcel([object]$Excel) {
    if(-not $Excel){return}
    try{if($Excel.Workbooks.Count -eq 0){$Excel.Quit()}else{$Excel.Visible=$true}}
    catch{Write-ControlConnectionLog 'Não foi possível encerrar a instância criada durante a tentativa de conexão.'}
}
function Get-ControlExcelApplication {
    if($script:Excel){try{$count=$script:Excel.Workbooks.Count;if($null -ne $count -and [int]$count -ge 0){return [pscustomobject]@{Excel=$script:Excel;Created=$false}}}catch{}}
    try{$excel=[Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application');return [pscustomobject]@{Excel=$excel;Created=$false}}
    catch{
        try{$excel=New-Object -ComObject Excel.Application;return [pscustomobject]@{Excel=$excel;Created=$true}}
        catch{throw 'Não foi possível iniciar o Excel. Abra o Excel instalado no computador, conclua a ativação se for solicitada e tente conectar novamente.'}
    }
}
function Open-ControlWorkbook([object]$Excel,[string]$Path,[bool]$ReadOnly=$false) {
    $wb=$null;$openError=$null;$restoreError=$null;$security=$null;$securityRead=$false
    try{
        $security=$Excel.AutomationSecurity
        if($null -eq $security -or [int]$security -notin @(1,2,3)){throw 'Não foi possível consultar a configuração de segurança do Excel. Feche os avisos do Excel e tente novamente.'}
        $securityRead=$true
        $Excel.AutomationSecurity=3
        if($Excel.AutomationSecurity -ne 3){throw 'O Excel não confirmou a abertura segura da planilha. Feche os avisos do Excel e tente novamente.'}
        # UpdateLinks=0 controla esta abertura sem mudar as preferências globais de links.
        $wb=$Excel.Workbooks.Open($Path,0,$ReadOnly)
    }catch{$openError=$_}
    finally{
        if($securityRead){try{$Excel.AutomationSecurity=$security;if($Excel.AutomationSecurity -ne $security){throw 'A configuração anterior de segurança não foi restaurada.'}}catch{$restoreError=$_}}
    }
    if($openError){
        if($restoreError){Write-ControlConnectionLog 'O Excel também impediu a restauração da preferência de segurança; a falha original da abertura foi preservada.'}
        throw $openError
    }
    if($restoreError){
        Close-ControlAttemptWorkbook $wb
        throw 'O Excel abriu o arquivo, mas impediu a restauração da configuração de segurança. A conexão foi cancelada. Feche as janelas de aviso do Excel e tente novamente.'
    }
    return $wb
}
function Connect-Workbook([string]$Path) {
    if(-not(Test-Path -LiteralPath $Path -PathType Leaf)){throw 'Arquivo não encontrado. Clique em Escolher planilha e selecione o arquivo .xlsx do controle.'}
    $full=[IO.Path]::GetFullPath($Path)
    if([IO.Path]::GetExtension($full).ToLowerInvariant() -ne '.xlsx'){throw 'Escolha a planilha .xlsx deste controle.'}
    $names=@('Excel','Wb','Central','WorkbookDir','BackupDir','CacheDir','LedgerPath','Ledger','Pending','Preview','FinancialBefore','CaseFamilies','Context','BatchHash','DeskRows','DeskCurrentNumber','DeskConnectionUnavailable')
    $previous=@{}
    foreach($name in $names){$v=Get-Variable -Name $name -Scope Script -ErrorAction SilentlyContinue;$previous[$name]=@{Exists=($null -ne $v);Value=$(if($v){$v.Value}else{$null})}}
    $previousPath=[string]$script:PathBox.Text
    $excel=$null;$wb=$null;$opened=$false;$created=$false;$published=$false;$success=$false
    try{
        $application=Get-ControlExcelApplication;$excel=$application.Excel;$created=[bool]$application.Created
        if(-not $excel.Ready){throw 'O Excel está com uma edição ou uma janela aberta. Conclua essa ação e clique em Tentar novamente.'}
        foreach($candidate in $excel.Workbooks){
            if([string]::Equals([IO.Path]::GetFullPath($candidate.FullName),$full,[StringComparison]::OrdinalIgnoreCase)){$wb=$candidate;break}
        }
        if(-not $wb){$wb=Open-ControlWorkbook $excel $full $false;$opened=$true}
        if($wb.ReadOnly){throw 'A planilha está somente leitura ou aberta em outra instância do Excel. Feche a cópia bloqueada e conecte novamente.'}
        try{
            $central=$wb.Worksheets.Item('Central IA');$null=$wb.Worksheets.Item('Painel');$null=$wb.Worksheets.Item('Configurar')
            foreach($s in @('Processos','Valores','Andamentos')){
                $sheet=$wb.Worksheets.Item($s)
                if([string]$sheet.Range('A9').Value2 -cne 'Número do processo'){throw "A aba $s não pertence ao modelo esperado."}
            }
            if([string]$central.Range('I28').Value2 -notin @('ID','ID da proposta')){throw 'A aba Central IA não pertence ao modelo esperado.'}
            $caseSheet=$wb.Worksheets.Item('Processos')
            if([string]$caseSheet.Range('H9').Value2 -cne 'Parte autora' -or [string]$caseSheet.Range('I9').Value2 -cne 'Parte ré'){throw 'Os campos Parte autora e Parte ré não foram encontrados.'}
        }catch{throw ('Esta planilha não é compatível com o controle. Escolha o arquivo do pacote ou sua cópia atualizada do mesmo modelo. '+$_.Exception.Message)}
        $workbookDir=Join-Path $script:DataDir (Get-TextHash $full.ToLowerInvariant()).Substring(0,20)
        $backupDir=Join-Path $workbookDir 'Backups';$cacheDir=Join-Path $workbookDir 'AnalisesProtegidas'
        $null=New-Item -ItemType Directory -Path $backupDir -Force;$null=New-Item -ItemType Directory -Path $cacheDir -Force
        $published=$true
        $script:Excel=$excel;$script:Wb=$wb;$script:Central=$central
        $script:WorkbookDir=$workbookDir;$script:BackupDir=$backupDir;$script:CacheDir=$cacheDir;$script:LedgerPath=Join-Path $workbookDir 'aplicacoes.json'
        Load-Ledger
        $script:Pending=@();$script:Preview=@{};$script:FinancialBefore=@{};$script:CaseFamilies=@{};$script:Context=$null;$script:BatchHash=''
        $script:PathBox.Text=$full
        # O carregamento dos dados também faz parte da tentativa: se falhar, restaurar o vínculo anterior.
        Initialize-DeskAfterConnect
        Write-Central 'connection' 'Conectado';Write-Central 'command' '';Write-Central 'mode' 'Revisar tudo'
        $excel.Visible=$true
        $success=$true
        if($previous.Central.Value -and -not [object]::ReferenceEquals($previous.Wb.Value,$wb)){
            try{$previous.Central.Value.Range($script:Cells.connection).Value2='Desconectado'}catch{}
        }
        # Falhas de foco não invalidam uma planilha já carregada e utilizável.
        try{$wb.Activate();$wb.Worksheets.Item('Painel').Activate();$script:Form.Activate()}catch{Write-ControlConnectionLog 'Conexão concluída; o Excel não aceitou mudar a janela ativa.'}
        Set-Status 'Planilha conectada. Busque um processo ou clique em Novo processo.'
        Write-ControlConnectionLog 'Planilha validada e assistente ativo.'
    }catch{
        $failure=$_
        if($published -and -not $success){
            foreach($name in $names){if($previous[$name].Exists){Set-Variable -Name $name -Scope Script -Value $previous[$name].Value}else{Remove-Variable -Name $name -Scope Script -ErrorAction SilentlyContinue}}
            $script:PathBox.Text=$previousPath
            if(Get-Variable -Name DeskDetailLoadedNumber -Scope Script -ErrorAction SilentlyContinue){$script:DeskDetailLoadedNumber=''}
            if(Get-Command Update-DeskFilter -ErrorAction SilentlyContinue){try{Update-DeskFilter}catch{}}
        }
        if(-not $success){
            if($opened){Close-ControlAttemptWorkbook $wb}
            if($created){Close-ControlAttemptExcel $excel}
        }
        throw $failure
    }finally{
        if(Get-Command Update-DeskConnectionState -ErrorAction SilentlyContinue){try{Update-DeskConnectionState}catch{}}
    }
}

function Run-Command([string]$Command) {
    switch -Exact ($Command){
        'Analisar texto'{Analyze-Text}
        'Aplicar aprovados'{
            if(-not $script:Pending.Count){throw 'Não há prévia nesta sessão. Escolha Analisar texto primeiro.'}
            $selected=@($script:Pending | Where-Object {-not $_.Applied -and [string]$script:Central.Cells.Item($_.PreviewRow,1).Value2 -ceq 'Sim'})
            Apply-Proposals $selected $true
        }
        'Importar arquivo'{Import-Source}
        'Limpar entrada'{
            Write-Central 'text' '';Write-Central 'case' '';Write-Central 'source' '';Write-Central 'command' ''
            $script:Central.Range('A29:J228').ClearContents();$script:Pending=@();$script:Preview=@{}
            Set-Status 'Entrada e prévia limpas. O cadastro e o histórico foram preservados.'
        }
        'Trocar chave'{Show-KeyDialog;Set-Status 'Configuração de chave encerrada. Se salvou uma chave, a próxima análise usará a nova configuração.'}
        default{throw 'Comando desconhecido. Use a lista da célula C6.'}
    }
}

Install-Poppins
$regularFamily=$null
if($script:PrivateFonts){foreach($family in $script:PrivateFonts.Families){if($family.Name -eq 'Poppins' -and $family.IsStyleAvailable([Drawing.FontStyle]::Regular)){$regularFamily=$family;break}}}
$script:UiFont=if($regularFamily){New-Object Drawing.Font($regularFamily,10,[Drawing.FontStyle]::Regular)}else{New-Object Drawing.Font('Segoe UI',10,[Drawing.FontStyle]::Regular)}
. (Join-Path $PSScriptRoot 'Juridico.Cadastro.ps1')
. (Join-Path $PSScriptRoot 'Juridico.Interface.ps1')
$script:RequestedWorkbookPath=$WorkbookPath
Start-DeskInterface
