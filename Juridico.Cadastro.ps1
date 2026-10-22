# Formulários manuais. O Excel continua sendo a fonte dos dados.
# Dot-source depois das funções transacionais do assistente. Windows PowerShell 5.1.

function Get-DeskFieldColumns([string]$Area) {
    if($Area -eq 'Processos'){return @('A','B','C','D','E','F','G','H','I','J','K','L','M','N','O','P','Q','S','T','U','AW','AX')}
    if($Area -eq 'Valores'){return @('B','C','D','E','G','H','J','K','T','U','Y','Z','AA')}
    throw 'Área de cadastro desconhecida.'
}
function Get-DeskColumnNumber([string]$Column) {
    if($Column -cnotmatch '^[A-Z]{1,2}$'){throw 'Coluna inválida.'}
    $n=0;foreach($c in $Column.ToCharArray()){$n=$n*26+[int]$c-64};return $n
}
function Read-DeskBlock([string]$Area) {
    if(-not $script:Wb){throw 'Abra uma planilha para começar.'}
    $ranges=@{Processos='A10:BD509';Valores='A10:AB509';Andamentos='A10:F2009';Configurar='A10:D39'}
    if(-not $ranges.ContainsKey($Area)){throw 'Área de leitura desconhecida.'}
    return ,($script:Wb.Worksheets.Item($Area).Range($ranges[$Area]).Value2)
}
function Get-DeskBlockValue([object]$Block,[int]$Row,[int]$Column) {
    return $Block[($Block.GetLowerBound(0)+$Row-1),($Block.GetLowerBound(1)+$Column-1)]
}
function Get-DeskIndex([object]$Block,[string]$Area) {
    $index=@{}
    for($r=1;$r -le $Block.GetLength(0);$r++){
        $number=Get-DeskBlockValue $Block $r 1
        if(Is-Blank $number){continue}
        if($number -isnot [string]){throw "Há um número de processo armazenado como número em $Area, linha $($r+9). Confira esse registro no Excel antes de salvar."}
        $key=Get-CaseKey $number
        if($index.ContainsKey($key)){throw "O mesmo identificador aparece duas vezes em $Area. Diferencie os registros antes de continuar."}
        $index[$key]=[pscustomobject]@{Number=[string]$number;Row=$r+9;Offset=$r}
    }
    return $index
}
function Get-DeskLists {
    $block=Read-DeskBlock 'Configurar';$result=@{}
    foreach($spec in @(@('Office',1,20),@('Company',2,20),@('Consortium',3,30),@('Theme',4,10))){
        $items=New-Object System.Collections.Generic.List[string]
        for($r=1;$r -le [int]$spec[2];$r++){
            $v=[string](Get-DeskBlockValue $block $r ([int]$spec[1]))
            if(-not [string]::IsNullOrWhiteSpace($v)){$items.Add($v.Trim())}
        }
        $result[[string]$spec[0]]=@($items.ToArray() | Sort-Object -Unique)
    }
    $result.Nature=@('Judicial','Administrativo','Inquérito / Policial','Controle externo')
    $result.Status=@('Ativo','Suspenso','Encerrado','Arquivado')
    $result.Priority=@('Normal','Alta','Crítica')
    $result.YesNo=@('Sim','Não')
    return $result
}
function Format-DeskDate([object]$Value,[bool]$WithTime=$false) {
    if(Is-Blank $Value){return ''}
    if($Value -is [datetime]){$date=$Value}
    elseif($Value -is [ValueType]){
        try{$date=[datetime]::FromOADate([double]$Value)}catch{return [string]$Value}
    }else{return [string]$Value}
    if($WithTime){return $date.ToString('dd/MM/yyyy HH:mm')}
    return $date.ToString('dd/MM/yyyy')
}
function Format-DeskField([string]$Column,[object]$Value,[string]$Area) {
    if(Is-Blank $Value){return ''}
    if($Area -notin @('Processos','Valores')){throw 'Área de cadastro desconhecida.'}
    if(($Area -eq 'Processos' -and $Column -in @('L','S')) -or ($Area -eq 'Valores' -and $Column -eq 'J')){return Format-DeskDate $Value}
    $culture=[Globalization.CultureInfo]::GetCultureInfo('pt-BR')
    if($Area -eq 'Valores' -and $Value -is [ValueType]){
        if($Column -in @('B','C','D','E','G','U')){return ([decimal]$Value).ToString('N2',$culture)}
        if($Column -eq 'H'){return (([decimal]$Value)*100).ToString('0.########',$culture)+'%'}
    }
    return [string]$Value
}
function ConvertFrom-DeskNumber([string]$Text,[bool]$Percentage=$false) {
    $s=$Text.Trim() -replace '\u00a0',' '
    if($Percentage){$s=$s -replace '\s*%\s*$',''}else{$s=$s -replace '^R\$\s*',''}
    $s=$s.Trim()
    # O formulário usa o padrão brasileiro: 1.234,56. Nunca adivinhar a localidade.
    if($s -notmatch '^(?:\d+|\d{1,3}(?:\.\d{3})+)(?:,\d{1,8})?$'){throw 'Informe o valor no formato brasileiro, por exemplo 1.234,56. Não use valores negativos.'}
    $number=[decimal]::Parse($s,[Globalization.NumberStyles]::Number,[Globalization.CultureInfo]::GetCultureInfo('pt-BR'))
    if($Percentage){
        if($number -gt 100){throw 'A probabilidade de êxito deve estar entre 0% e 100%.'}
        return [double]($number/100)
    }
    if($s -match ',\d{3,}$'){throw 'Use no máximo duas casas decimais para valores em reais.'}
    if($number -gt 9999999999999.99){throw 'O valor é grande demais. Confira os separadores e as casas decimais.'}
    return [double]$number
}
function ConvertFrom-DeskDate([string]$Text) {
    $date=[datetime]::MinValue
    if(-not [datetime]::TryParseExact($Text.Trim(),'dd/MM/yyyy',[Globalization.CultureInfo]::GetCultureInfo('pt-BR'),[Globalization.DateTimeStyles]::None,[ref]$date)){throw 'Informe a data no formato dia/mês/ano, por exemplo 21/10/2026.'}
    if($date.Year -lt 1900){throw 'Confira a data: o ano deve ser 1900 ou posterior.'}
    return $date.ToOADate()
}
function ConvertFrom-DeskField([string]$Column,[object]$InputValue,[string]$Area,[hashtable]$Lists) {
    $allowed=Get-DeskFieldColumns $Area
    if($allowed -cnotcontains $Column){throw "O campo $Column não está disponível nesse formulário."}
    $text=if($null -eq $InputValue){''}else{[string]$InputValue}
    $text=$text.Trim()
    if($text.Length -eq 0){return $null}
    if($text.Length -gt 32767){throw 'O texto é grande demais para uma célula. Divida o conteúdo em mais de uma atualização.'}
    if($text -match '[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]'){throw 'O campo contém caracteres de controle inválidos.'}
    if($Area -eq 'Processos'){
        if($Column -eq 'A'){$null=Get-CaseKey $text;return $text}
        if($Column -in @('L','S')){return ConvertFrom-DeskDate $text}
        $listNames=@{B='Nature';C='Status';D='Office';E='Company';F='Consortium';G='Theme';N='Priority';O='YesNo';P='YesNo'}
        if($listNames.ContainsKey($Column)){
            $listName=$listNames[$Column]
            $match=@($Lists[$listName] | Where-Object {[string]::Equals([string]$_,$text,[StringComparison]::OrdinalIgnoreCase)})
            if($match.Count -ne 1){
                if($listName -in @('Office','Company','Consortium','Theme')){throw 'Escolha um nome da lista ou use Cadastros para acrescentá-lo primeiro.'}
                throw 'Escolha uma opção válida da lista.'
            }
            return [string]$match[0]
        }
        if($Column -eq 'T' -and $text.Length -gt 2048){throw 'O link dos documentos é grande demais.'}
    }else{
        if($Column -in @('B','C','D','E','G','U')){return ConvertFrom-DeskNumber $text}
        if($Column -eq 'H'){return ConvertFrom-DeskNumber $text $true}
        if($Column -eq 'J'){return ConvertFrom-DeskDate $text}
        if($Column -eq 'T'){
            $match=@(@('Sim','Não','Revisar') | Where-Object {[string]::Equals($_,$text,[StringComparison]::OrdinalIgnoreCase)})
            if($match.Count -ne 1){throw 'A inclusão nos totais deve ser Sim, Não ou Revisar.'}
            return $match[0]
        }
        # Y, Z e AA preservam as descrições originais. Perda não vira êxito.
    }
    return $text
}
function Get-DeskRows {
    $block=Read-DeskBlock 'Processos';$index=Get-DeskIndex $block 'Processos'
    $rows=New-Object System.Collections.Generic.List[object]
    foreach($entry in @($index.Values | Sort-Object Row)){
        $r=$entry.Offset;$status=[string](Get-DeskBlockValue $block $r 3);$priority=[string](Get-DeskBlockValue $block $r 14)
        $critical=($status -in @('Ativo','Suspenso')) -and ($priority -eq 'Crítica' -or (Get-DeskBlockValue $block $r 15) -eq 'Sim' -or (Get-DeskBlockValue $block $r 16) -eq 'Sim')
        $updated=Get-DeskBlockValue $block $r 26
        $rows.Add([pscustomobject]@{
            Number=$entry.Number;Row=$entry.Row;Nature=[string](Get-DeskBlockValue $block $r 2);Status=$status
            Office=[string](Get-DeskBlockValue $block $r 4);Company=[string](Get-DeskBlockValue $block $r 5)
            Consortium=[string](Get-DeskBlockValue $block $r 6);Theme=[string](Get-DeskBlockValue $block $r 7)
            Author=[string](Get-DeskBlockValue $block $r 8);Defendant=[string](Get-DeskBlockValue $block $r 9)
            Priority=$priority;Critical=[bool]$critical;NeedsUpdate=((Get-DeskBlockValue $block $r 30) -eq 1)
            LastUpdate=$(if(Is-Blank $updated){''}elseif($updated -eq 0){''}else{Format-DeskDate $updated})
            LastText=[string](Get-DeskBlockValue $block $r 29)
        })
    }
    return $rows.ToArray()
}
function Get-DeskRecord([string]$Number) {
    $key=Get-CaseKey $Number;$pb=Read-DeskBlock 'Processos';$vb=Read-DeskBlock 'Valores';$hb=Read-DeskBlock 'Andamentos'
    $pi=Get-DeskIndex $pb 'Processos';$vi=Get-DeskIndex $vb 'Valores'
    if(-not $pi.ContainsKey($key)){throw 'Esse processo não foi encontrado. Atualize a lista e selecione-o novamente.'}
    $entry=$pi[$key];$fields=@{};$money=@{}
    foreach($column in (Get-DeskFieldColumns 'Processos')){$fields[$column]=Get-DeskBlockValue $pb $entry.Offset (Get-DeskColumnNumber $column)}
    $moneyRow=0
    if($vi.ContainsKey($key)){
        $m=$vi[$key]
        if($m.Number -cne $entry.Number){throw 'O identificador está escrito de forma diferente em Cadastro e Valores. Padronize o registro no Excel antes de salvar.'}
        $moneyRow=$m.Row
        foreach($column in (Get-DeskFieldColumns 'Valores')){$money[$column]=Get-DeskBlockValue $vb $m.Offset (Get-DeskColumnNumber $column)}
    }else{foreach($column in (Get-DeskFieldColumns 'Valores')){$money[$column]=$null}}
    $history=New-Object System.Collections.Generic.List[object]
    for($r=1;$r -le $hb.GetLength(0);$r++){
        $n=Get-DeskBlockValue $hb $r 1
        if(Is-Blank $n){continue}
        if($n -isnot [string]){continue}
        if((Get-CaseKey $n) -cne $key){continue}
        $date=Get-DeskBlockValue $hb $r 2;$registered=Get-DeskBlockValue $hb $r 4
        $sort=if($registered -is [ValueType]){[double]$registered}else{0}
        $history.Add([pscustomobject]@{Date=(Format-DeskDate $date);Registered=(Format-DeskDate $registered $true);Text=[string](Get-DeskBlockValue $hb $r 3);Source=[string](Get-DeskBlockValue $hb $r 5);Notes=[string](Get-DeskBlockValue $hb $r 6);Row=$r+9;RegisteredValue=$sort})
    }
    $family=@($pi.Values | Where-Object {(Get-CaseBaseKey $_.Number) -ceq (Get-CaseBaseKey $entry.Number)} | ForEach-Object {$_.Number} | Sort-Object)
    $snapshot=@{Workbook=[string]$script:Wb.FullName;Number=$entry.Number;Key=$key;ProcessRow=$entry.Row;MoneyRow=$moneyRow;Fields=$fields.Clone();Money=$money.Clone();Family=$family}
    return [pscustomobject]@{Number=$entry.Number;Fields=$fields;Money=$money;History=@($history.ToArray() | Sort-Object -Property @{Expression='RegisteredValue';Descending=$true},@{Expression='Row';Descending=$true});Snapshot=$snapshot}
}
function Assert-DeskSnapshot([object]$Snapshot,[object]$Current,[string]$Area) {
    if($null -eq $Snapshot -or $Snapshot -isnot [hashtable]){throw 'Abra o processo novamente antes de salvar.'}
    foreach($field in @('Workbook','Number','Key','Fields','Money','Family','MoneyRow')){if(-not $Snapshot.ContainsKey($field)){throw 'O formulário precisa ser recarregado antes de salvar.'}}
    if(-not [string]::Equals([string]$Snapshot.Workbook,[string]$script:Wb.FullName,[StringComparison]::OrdinalIgnoreCase)){throw 'A planilha foi trocada. Abra o processo novamente antes de salvar.'}
    if($Snapshot.Number -cne $Current.Number -or $Snapshot.Key -cne $Current.Snapshot.Key){throw 'O identificador mudou. Abra o processo novamente antes de salvar.'}
    if((@($Snapshot.Family) -join "`n") -cne (@($Current.Snapshot.Family) -join "`n")){throw 'Os registros com esse número mudaram. Abra o processo novamente antes de salvar.'}
    $name=if($Area -eq 'Processos'){'Fields'}else{'Money'}
    foreach($column in (Get-DeskFieldColumns $Area)){
        if(-not $Snapshot[$name].ContainsKey($column) -or -not(Test-ValueEqual $Snapshot[$name][$column] $Current.Snapshot[$name][$column])){throw 'Esse cadastro foi alterado depois de abrir o formulário. Recarregue o processo para não substituir uma atualização mais recente.'}
    }
    if($Area -eq 'Valores' -and (($Snapshot.MoneyRow -eq 0) -ne ($Current.Snapshot.MoneyRow -eq 0))){throw 'Os valores foram cadastrados em outra edição. Recarregue o processo antes de salvar.'}
}
function Get-DeskFreeRow([string]$Area,[object]$Block) {
    # Valores!T já vem preenchida com Sim nas linhas disponíveis do modelo.
    # Outros conteúdos nessa coluna indicam uma linha reservada e são preservados.
    $inputColumns=if($Area -eq 'Processos'){@(1..21)+@(46..53)+@(56)}elseif($Area -eq 'Valores'){@(1..11)+@(21..28)}elseif($Area -eq 'Andamentos'){@(1..6)}else{throw 'Área de cadastro desconhecida.'}
    for($r=1;$r -le $Block.GetLength(0);$r++){
        if($Area -eq 'Valores'){$include=Get-DeskBlockValue $Block $r 20;if(-not(Is-Blank $include) -and $include -cne 'Sim'){continue}}
        $empty=$true
        foreach($c in $inputColumns){if(-not(Is-Blank (Get-DeskBlockValue $Block $r $c))){$empty=$false;break}}
        if($empty){return $r+9}
    }
    throw "A capacidade de $Area foi atingida. Nenhum dado foi salvo."
}
function Add-DeskWrite([System.Collections.Generic.List[object]]$Writes,[string]$Area,[int]$Row,[string]$Column,[object]$Value,[object]$Before) {
    $max=if($Area -eq 'Andamentos'){2009}else{509}
    if($Row -lt 10 -or $Row -gt $max){throw 'Linha fora da área de cadastro.'}
    $allowed=if($Area -eq 'Andamentos'){@('A','B','C','D','E','F')}elseif($Area -eq 'Valores'){@('A')+(Get-DeskFieldColumns 'Valores')}else{Get-DeskFieldColumns $Area}
    if($allowed -cnotcontains $Column){throw 'Esse campo não está disponível para gravação.'}
    $address=$Column+$Row;$cell=$script:Wb.Worksheets.Item($Area).Range($address)
    if($cell.HasFormula){throw "O campo $Area!$address contém uma fórmula. Nenhum dado foi salvo."}
    if(-not(Test-ValueEqual $cell.Value2 $Before)){throw 'O registro mudou durante a edição. Recarregue o formulário antes de salvar.'}
    foreach($write in $Writes){if($write.Sheet -eq $Area -and $write.Address -eq $address){$write.Value=$Value;return}}
    $Writes.Add([pscustomobject]@{Sheet=$Area;Address=$address;Value=$Value;Before=$Before;Format=$cell.NumberFormat})
}
function Reset-DeskPreview {
    $script:Pending=@();$script:Preview=@{}
    if($script:Central){try{$script:Central.Range('A29:J228').ClearContents();Write-Central 'command' ''}catch{}}
}
function Save-DeskProcess([string]$Number,[hashtable]$Fields,[object]$Snapshot) {
    if(-not $script:Wb){throw 'Abra uma planilha para começar.'}
    if($script:Wb.ReadOnly){throw 'A planilha está em modo somente leitura.'}
    $lists=Get-DeskLists;$new=[string]::IsNullOrWhiteSpace($Number);$values=@{};$before=@{}
    if($new){
        foreach($required in @('A','B','C','D','E')){if(-not $Fields.ContainsKey($required) -or [string]::IsNullOrWhiteSpace([string]$Fields[$required])){throw 'Preencha número, natureza, situação, escritório e empresa para cadastrar o processo.'}}
        $pb=Read-DeskBlock 'Processos';$pi=Get-DeskIndex $pb 'Processos'
        $Number=([string]$Fields.A).Trim();$key=Get-CaseKey $Number
        if($pi.ContainsKey($key)){throw 'Esse processo já está cadastrado. Localize-o na lista e escolha Editar cadastro.'}
        $family=@($pi.Values | Where-Object {(Get-CaseBaseKey $_.Number) -ceq (Get-CaseBaseKey $Number)})
        if($family.Count -and ((Get-CaseVariantId $Number) -eq '' -or @($family | Where-Object {(Get-CaseVariantId $_.Number) -eq ''}).Count)){throw 'Esse número já possui cadastro. Se forem registros distintos, use identificadores completos com [ID] para distingui-los.'}
        $row=Get-DeskFreeRow 'Processos' $pb
        foreach($column in (Get-DeskFieldColumns 'Processos')){$before[$column]=$null}
    }else{
        $current=Get-DeskRecord $Number;Assert-DeskSnapshot $Snapshot $current 'Processos'
        $Number=$current.Number;$row=$current.Snapshot.ProcessRow;$before=$current.Fields
        if($Fields.ContainsKey('A') -and [string]$Fields.A -cne $Number){throw 'O número identifica o processo e não pode ser alterado nesse formulário.'}
    }
    foreach($column in $Fields.Keys){
        if((Get-DeskFieldColumns 'Processos') -cnotcontains [string]$column){throw 'O formulário enviou um campo de cadastro desconhecido.'}
        $text=if($null -eq $Fields[$column]){''}else{[string]$Fields[$column]}
        if(-not $new -and $text -ceq (Format-DeskField $column $before[$column] 'Processos')){$values[$column]=$before[$column];continue}
        $values[$column]=ConvertFrom-DeskField $column $text 'Processos' $lists
    }
    $writes=New-Object System.Collections.Generic.List[object]
    foreach($column in $values.Keys){if(-not(Test-ValueEqual $before[$column] $values[$column])){Add-DeskWrite $writes 'Processos' $row $column $values[$column] $before[$column]}}
    if($writes.Count -eq 0){return [pscustomobject]@{Number=$Number;Changed=$false;Created=$false}}
    if(-not $new){Add-DeskWrite $writes 'Processos' $row 'A' $Number $Number}
    Commit-Writes $writes @();Reset-DeskPreview
    return [pscustomobject]@{Number=$Number;Changed=$true;Created=$new}
}
function Save-DeskValues([string]$Number,[hashtable]$Fields,[object]$Snapshot) {
    if(-not $script:Wb){throw 'Abra uma planilha para começar.'}
    if($script:Wb.ReadOnly){throw 'A planilha está em modo somente leitura.'}
    $current=Get-DeskRecord $Number;Assert-DeskSnapshot $Snapshot $current 'Valores';$Number=$current.Number
    $new=$current.Snapshot.MoneyRow -eq 0;$before=$current.Money;$values=@{};$lists=Get-DeskLists
    foreach($column in $Fields.Keys){
        if((Get-DeskFieldColumns 'Valores') -cnotcontains [string]$column){throw 'O formulário enviou um campo de valores desconhecido.'}
        $text=if($null -eq $Fields[$column]){''}else{[string]$Fields[$column]}
        if($text -ceq (Format-DeskField $column $before[$column] 'Valores')){$values[$column]=$before[$column];continue}
        $values[$column]=ConvertFrom-DeskField $column $text 'Valores' $lists
    }
    $changes=@($values.Keys | Where-Object {-not(Test-ValueEqual $before[$_] $values[$_])})
    if($changes.Count -eq 0){return [pscustomobject]@{Number=$Number;Changed=$false;Created=$false}}
    if($new){
        $block=Read-DeskBlock 'Valores';$row=Get-DeskFreeRow 'Valores' $block
        # Capturar o padrão real do modelo: a transação valida Before antes de alterar.
        $before.T=Get-DeskBlockValue $block ($row-9) 20
    }else{$row=$current.Snapshot.MoneyRow}
    $writes=New-Object System.Collections.Generic.List[object]
    Add-DeskWrite $writes 'Processos' $current.Snapshot.ProcessRow 'A' $Number $Number
    Add-DeskWrite $writes 'Valores' $row 'A' $Number $(if($new){$null}else{$Number})
    foreach($column in $changes){Add-DeskWrite $writes 'Valores' $row $column $values[$column] $before[$column]}
    if($new -and -not $Fields.ContainsKey('T')){
        $include=if(@($current.Snapshot.Family).Count -eq 1){'Sim'}else{'Revisar'}
        Add-DeskWrite $writes 'Valores' $row 'T' $include $before.T
    }
    Commit-Writes $writes @();Reset-DeskPreview
    return [pscustomobject]@{Number=$Number;Changed=$true;Created=$new}
}
function Add-DeskHistory([string]$Number,[Alias('DateBR')][string]$Date,[string]$Description,[string]$Source,[string]$Notes) {
    if(-not $script:Wb){throw 'Abra uma planilha para começar.'}
    if($script:Wb.ReadOnly){throw 'A planilha está em modo somente leitura.'}
    if([string]::IsNullOrWhiteSpace($Description)){throw 'Descreva o que aconteceu no processo.'}
    if([string]::IsNullOrWhiteSpace($Date)){throw 'Informe a data do andamento.'}
    if([string]::IsNullOrWhiteSpace($Source)){throw 'Informe de onde veio essa atualização.'}
    $typedDate=ConvertFrom-DeskDate $Date
    foreach($text in @($Description,$Source,$Notes)){if($text.Length -gt 32767 -or $text -match '[\x00-\x08\x0b\x0c\x0e-\x1f\x7f]'){throw 'Um texto da atualização é grande demais ou contém caracteres inválidos.'}}
    $current=Get-DeskRecord $Number;$Number=$current.Number;$block=Read-DeskBlock 'Andamentos'
    for($r=1;$r -le $block.GetLength(0);$r++){
        $n=Get-DeskBlockValue $block $r 1
        if($n -is [string] -and -not [string]::IsNullOrWhiteSpace($n) -and (Get-CaseKey $n) -ceq (Get-CaseKey $Number)){
            if((Test-ValueEqual (Get-DeskBlockValue $block $r 2) $typedDate) -and [string](Get-DeskBlockValue $block $r 3) -ceq $Description.Trim() -and [string](Get-DeskBlockValue $block $r 5) -ceq $Source.Trim() -and [string](Get-DeskBlockValue $block $r 6) -ceq $Notes.Trim()){throw 'Essa atualização já está registrada com a mesma data, texto e fonte.'}
        }
    }
    $row=Get-DeskFreeRow 'Andamentos' $block;$registered=[datetime]::Now.ToOADate()
    # Datas iguais de recebimento geram empate na planilha. Garantir um instante único.
    $max=0.0
    foreach($item in $current.History){if($item.RegisteredValue -gt $max){$max=[double]$item.RegisteredValue}}
    if($registered -le $max){$registered=$max+(1.0/86400)}
    $values=@($Number,$typedDate,$Description.Trim(),$registered,$Source.Trim(),$Notes.Trim())
    $writes=New-Object System.Collections.Generic.List[object]
    Add-DeskWrite $writes 'Processos' $current.Snapshot.ProcessRow 'A' $Number $Number
    for($c=0;$c -lt 6;$c++){Add-DeskWrite $writes 'Andamentos' $row ([string][char](65+$c)) $values[$c] $null}
    Commit-Writes $writes @();Reset-DeskPreview
    return [pscustomobject]@{Number=$Number;Changed=$true;Row=$row}
}
function Add-DeskListValue([string]$List,[string]$Value) {
    if(-not $script:Wb){throw 'Abra uma planilha para começar.'}
    if($script:Wb.ReadOnly){throw 'A planilha está em modo somente leitura.'}
    $specs=@{Office=@('A',1,20);Company=@('B',2,20);Consortium=@('C',3,30);Theme=@('D',4,10)}
    if(-not $specs.ContainsKey($List)){throw 'Tipo de cadastro desconhecido.'}
    $value=$Value.Trim()
    if($value.Length -eq 0 -or $value.Length -gt 200 -or $value -match '[\x00-\x1f\x7f]'){throw 'Informe um nome válido, com até 200 caracteres.'}
    $spec=$specs[$List];$block=Read-DeskBlock 'Configurar';$free=0
    for($r=1;$r -le [int]$spec[2];$r++){
        $current=Get-DeskBlockValue $block $r ([int]$spec[1])
        if([string]::Equals([string]$current,$value,[StringComparison]::OrdinalIgnoreCase)){return [pscustomobject]@{Value=[string]$current;Added=$false}}
        if($free -eq 0 -and (Is-Blank $current)){$free=$r+9}
    }
    if($free -eq 0){throw 'A lista está completa. Nenhum nome foi substituído.'}
    $address=[string]$spec[0]+$free;$cell=$script:Wb.Worksheets.Item('Configurar').Range($address)
    if($cell.HasFormula -or -not(Is-Blank $cell.Value2)){throw 'A lista mudou enquanto você preenchia. Atualize os cadastros e tente novamente.'}
    $writes=New-Object System.Collections.Generic.List[object]
    $writes.Add([pscustomobject]@{Sheet='Configurar';Address=$address;Value=$value;Before=$cell.Value2;Format=$cell.NumberFormat})
    Commit-Writes $writes @();Reset-DeskPreview
    return [pscustomobject]@{Value=$value;Added=$true}
}
