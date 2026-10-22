# Interface de uso diário. Carregada no mesmo escopo do assistente.
Set-StrictMode -Version 2.0
$script:DeskRows=@();$script:DeskLoading=$false;$script:DeskCurrentNumber=''
$script:DeskInk=[Drawing.Color]::FromArgb(28,47,73)
$script:DeskBlue=[Drawing.Color]::FromArgb(42,91,161)
$script:DeskBackground=[Drawing.Color]::FromArgb(242,245,249)
$script:DeskMuted=[Drawing.Color]::FromArgb(83,102,125)
$script:DeskBorder=[Drawing.Color]::FromArgb(220,228,237)
$script:DeskConnectionLabel=$null;$script:DeskOffline=$null;$script:DeskLastAttemptPath=''
$script:DeskConnectionUnavailable=$false
$script:DeskDetailLoadedNumber='';$script:DeskDetailRevision=-1;$script:DeskDataRevision=0;$script:DeskDetailWorkbookPath=''
$script:DeskWorkbookControls=New-Object System.Collections.Generic.List[object]
$script:DeskSelectionControls=New-Object System.Collections.Generic.List[object]
$script:DeskMetricLabels=@{};$script:DeskMetricCards=@{}

function New-DeskButton([string]$Text,[string]$Command,[int]$Width=115,[bool]$Primary=$false) {
    $b=New-Object Windows.Forms.Button
    $b.Text=$Text;$b.Tag=$Command;$b.Width=$Width;$b.Height=38
    $b.Margin=New-Object Windows.Forms.Padding(0,0,8,0)
    $b.FlatStyle=[Windows.Forms.FlatStyle]::Flat;$b.FlatAppearance.BorderSize=1
    $b.FlatAppearance.BorderColor=$script:DeskBorder
    $b.BackColor=if($Primary){$script:DeskBlue}else{[Drawing.Color]::White}
    $b.ForeColor=if($Primary){[Drawing.Color]::White}else{$script:DeskInk}
    $b.FlatAppearance.MouseOverBackColor=if($Primary){[Drawing.Color]::FromArgb(34,76,137)}else{[Drawing.Color]::FromArgb(235,242,251)}
    $b.Font=New-Object Drawing.Font($script:UiFont.FontFamily,9,[Drawing.FontStyle]::Regular)
    $b.Cursor=[Windows.Forms.Cursors]::Hand;$b.UseVisualStyleBackColor=$false
    $b.AccessibleName=$Text
    return $b
}
function New-DeskGrid {
    $g=New-Object Windows.Forms.DataGridView
    $g.Dock='Fill';$g.ReadOnly=$true;$g.AllowUserToAddRows=$false;$g.AllowUserToDeleteRows=$false
    $g.AllowUserToResizeRows=$false;$g.RowHeadersVisible=$false;$g.MultiSelect=$false
    $g.SelectionMode='FullRowSelect';$g.BackgroundColor=[Drawing.Color]::White
    $g.BorderStyle='None';$g.AutoGenerateColumns=$false;$g.EnableHeadersVisualStyles=$false
    $g.ColumnHeadersDefaultCellStyle.BackColor=[Drawing.Color]::FromArgb(234,240,247)
    $g.ColumnHeadersDefaultCellStyle.ForeColor=$script:DeskInk
    $g.ColumnHeadersDefaultCellStyle.Font=New-Object Drawing.Font($script:UiFont.FontFamily,9,[Drawing.FontStyle]::Bold)
    $g.ColumnHeadersHeight=38;$g.ColumnHeadersHeightSizeMode='DisableResizing'
    $g.ColumnHeadersBorderStyle='None';$g.CellBorderStyle='SingleHorizontal'
    $g.DefaultCellStyle.Font=New-Object Drawing.Font($script:UiFont.FontFamily,9,[Drawing.FontStyle]::Regular)
    $g.DefaultCellStyle.ForeColor=$script:DeskInk
    $g.DefaultCellStyle.SelectionBackColor=[Drawing.Color]::FromArgb(224,235,251)
    $g.DefaultCellStyle.SelectionForeColor=$script:DeskInk
    $g.DefaultCellStyle.Padding=New-Object Windows.Forms.Padding(5,4,5,4)
    $g.AlternatingRowsDefaultCellStyle.BackColor=[Drawing.Color]::FromArgb(248,250,253)
    $g.GridColor=[Drawing.Color]::FromArgb(235,240,246);$g.RowTemplate.Height=40
    return $g
}
function Add-DeskGridColumn($Grid,[string]$Name,[string]$Title,[int]$Width,[bool]$Fill=$false) {
    $c=New-Object Windows.Forms.DataGridViewTextBoxColumn
    $c.Name=$Name;$c.HeaderText=$Title;$c.Width=$Width;$c.SortMode='NotSortable'
    if($Fill){$c.AutoSizeMode='Fill';$c.MinimumWidth=$Width}
    [void]$Grid.Columns.Add($c)
}
function Fit-DeskWindow($Window) {
    $area=[Windows.Forms.Screen]::FromControl($Window).WorkingArea
    $maxWidth=[Math]::Max(400,$area.Width-16);$maxHeight=[Math]::Max(300,$area.Height-16)
    $Window.MinimumSize=New-Object Drawing.Size([Math]::Min($Window.MinimumSize.Width,$maxWidth),[Math]::Min($Window.MinimumSize.Height,$maxHeight))
    $Window.Size=New-Object Drawing.Size([Math]::Min($Window.Width,$maxWidth),[Math]::Min($Window.Height,$maxHeight))
    $Window.Left=[Math]::Max($area.Left,[Math]::Min($Window.Left,$area.Right-$Window.Width))
    $Window.Top=[Math]::Max($area.Top,[Math]::Min($Window.Top,$area.Bottom-$Window.Height))
}
function Get-DeskSelectedNumber {
    if($script:DeskGrid.SelectedRows.Count -eq 0){return ''}
    $r=$script:DeskGrid.SelectedRows[0]
    if(-not $r.Tag){return ''}
    return [string]$r.Tag.Number
}
function Test-DeskConnected {
    if(-not $script:Wb -or $script:DeskConnectionUnavailable){throw 'Conecte sua planilha em Escolher planilha para começar.'}
    try{$ready=$script:Excel.Ready;$fullName=[string]$script:Wb.FullName;if([string]::IsNullOrWhiteSpace($fullName)){throw 'A planilha não está mais aberta.'}}
    catch{Set-DeskConnectionUnavailable;throw 'A conexão com o Excel foi interrompida. Abra o Excel e clique em Tentar novamente.'}
    if(-not $ready){throw 'O Excel está ocupado. Conclua a edição ou feche a caixa de aviso aberta no Excel e tente novamente.'}
}
function Set-DeskConnectionUnavailable {
    $script:DeskConnectionUnavailable=$true
    if($script:PathBox.Text){$script:DeskLastAttemptPath=$script:PathBox.Text}
    Update-DeskConnectionState
}
function Set-DeskConnectionError([string]$Message) {
    $friendly=$Message
    if($Message -match '0x800AC472|0x80010001|RPC_E_CALL_REJECTED'){$friendly='O Excel está ocupado. Conclua a edição ou feche a caixa de aviso aberta no Excel e clique em Tentar novamente.'}
    Set-Status $friendly $true
    if($script:DeskOffline){$script:DeskOfflineMessage.Text=$friendly}
    Update-DeskConnectionState
}
function Invoke-DeskAction([scriptblock]$Action,[bool]$RequireWorkbook=$true) {
    if($script:Busy){return}
    $script:Busy=$true;$script:Form.UseWaitCursor=$true
    try{if($RequireWorkbook){Test-DeskConnected}; & $Action}
    catch{Set-DeskConnectionError $_.Exception.Message}
    finally{$script:Form.UseWaitCursor=$false;$script:Busy=$false;Update-DeskConnectionState}
}
function Get-DeskSearchText([object]$Row) {
    return (@($Row.Number,$Row.Office,$Row.Company,$Row.Consortium,$Row.Theme,$Row.Author,$Row.Defendant) -join ' ').ToLowerInvariant()
}
function Set-DeskQuickFilter([string]$Filter) {
    if(-not $script:Wb -or $script:Busy){return}
    $script:DeskLoading=$true
    try{$script:DeskSearch.Clear();$script:DeskFilter.SelectedItem=$Filter}finally{$script:DeskLoading=$false}
    Update-DeskFilter
}
function Update-DeskConnectionState {
    if(-not $script:DeskConnectionLabel){return}
    $connected=$null -ne $script:Wb -and -not $script:DeskConnectionUnavailable
    $selected=$connected -and -not [string]::IsNullOrWhiteSpace((Get-DeskSelectedNumber))
    foreach($c in $script:DeskWorkbookControls){$c.Enabled=$connected}
    foreach($c in $script:DeskSelectionControls){$c.Enabled=$selected}
    $script:DeskSearch.Enabled=$connected;$script:DeskFilter.Enabled=$connected
    if($connected){
        $name=[IO.Path]::GetFileName($script:PathBox.Text)
        $script:DeskConnectionLabel.Text='Planilha conectada  |  '+$name
        $script:DeskConnectionLabel.ForeColor=[Drawing.Color]::FromArgb(33,107,83)
        $script:DeskMetricLabels.Total.Text=[string]$script:DeskRows.Count
        $script:DeskMetricLabels.Active.Text=[string]@($script:DeskRows | Where-Object {$_.Status -in @('Ativo','Suspenso')}).Count
        $script:DeskMetricLabels.Critical.Text=[string]@($script:DeskRows | Where-Object {$_.Critical}).Count
        $script:DeskMetricLabels.Stale.Text=[string]@($script:DeskRows | Where-Object {$_.NeedsUpdate}).Count
    }else{
        $script:DeskConnectionLabel.Text='Planilha desconectada';$script:DeskConnectionLabel.ForeColor=$script:DeskMuted
        foreach($label in $script:DeskMetricLabels.Values){$label.Text='...'}
    }
    foreach($card in $script:DeskMetricCards.Values){
        $card.Enabled=$connected
        $card.BackColor=if($connected -and [string]$card.Tag -eq [string]$script:DeskFilter.SelectedItem){[Drawing.Color]::FromArgb(226,237,252)}else{[Drawing.Color]::White}
    }
    $script:DeskOffline.Visible=-not $connected
    if(-not $connected){$script:DeskOffline.BringToFront()}
    $script:DeskEmpty.Visible=$connected -and $script:DeskGrid.Rows.Count -eq 0
    if($script:DeskEmpty.Visible){
        if($script:DeskRows.Count -eq 0){$script:DeskEmptyTitle.Text='Cadastre o primeiro processo';$script:DeskEmptyText.Text='Os registros cadastrados aparecerão aqui para consulta e acompanhamento.';$script:DeskEmptyButton.Text='Novo processo';$script:DeskEmptyButton.Tag='new'}
        else{$script:DeskEmptyTitle.Text='Nenhum processo encontrado';$script:DeskEmptyText.Text='Tente outro número ou nome, ou limpe os filtros para ver todos os processos.';$script:DeskEmptyButton.Text='Limpar busca';$script:DeskEmptyButton.Tag='clear'}
        $script:DeskEmpty.BringToFront()
    }
}
function Update-DeskFilter([string]$SelectNumber='') {
    if($script:DeskLoading){return}
    $keep=if($SelectNumber){$SelectNumber}else{Get-DeskSelectedNumber}
    if(-not $keep){$keep=$script:DeskCurrentNumber}
    $script:DeskLoading=$true
    try{
        $q=$script:DeskSearch.Text.Trim().ToLowerInvariant();$filter=[string]$script:DeskFilter.SelectedItem
        $script:DeskGrid.Rows.Clear()
        foreach($r in $script:DeskRows){
            if($q -and -not (Get-DeskSearchText $r).Contains($q)){continue}
            if($filter -eq 'Em acompanhamento' -and $r.Status -notin @('Ativo','Suspenso')){continue}
            if($filter -eq 'Críticos' -and -not $r.Critical){continue}
            if($filter -eq 'Sem atualização' -and -not $r.NeedsUpdate){continue}
            if($filter -in @('Ativo','Suspenso','Encerrado','Arquivado') -and $r.Status -ne $filter){continue}
            $status=if($r.Critical){$r.Status+' / crítico'}else{$r.Status}
            $updated=if($r.LastUpdate){$r.LastUpdate}else{'Sem registro'}
            $i=$script:DeskGrid.Rows.Add([object[]]@($r.Number,$r.Office,$r.Company,$status,$updated))
            $row=$script:DeskGrid.Rows[$i];$row.Tag=$r
            if($r.Critical){$row.Cells['Status'].Style.ForeColor=[Drawing.Color]::FromArgb(160,53,42)}
            foreach($cell in $row.Cells){$cell.ToolTipText=[string]$cell.Value}
        }
        $script:DeskGrid.ClearSelection();$chosen=$null
        foreach($row in $script:DeskGrid.Rows){if([string]$row.Tag.Number -ceq $keep){$chosen=$row;break}}
        if(-not $chosen -and $script:DeskGrid.Rows.Count){$chosen=$script:DeskGrid.Rows[0]}
        if($chosen){$chosen.Selected=$true;$script:DeskGrid.CurrentCell=$chosen.Cells[0]}
        $script:DeskCount.Text=('{0} de {1} processos' -f $script:DeskGrid.Rows.Count,$script:DeskRows.Count)
    }finally{$script:DeskLoading=$false}
    Show-DeskSelection;Update-DeskConnectionState
}
function Refresh-DeskWorkspace([string]$SelectNumber='') {
    Test-DeskConnected;$script:Excel.Calculate()
    if($SelectNumber){
        $script:DeskCurrentNumber=$SelectNumber;$script:DeskLoading=$true
        try{$script:DeskSearch.Clear();$script:DeskFilter.SelectedIndex=0}finally{$script:DeskLoading=$false}
    }
    $script:DeskRows=@(Get-DeskRows);$script:DeskDataRevision++;Update-DeskFilter $SelectNumber
}
function Initialize-DeskAfterConnect {
    $script:DeskConnectionUnavailable=$false;$script:DeskDetailLoadedNumber='';$script:DeskCurrentNumber='';$script:DeskLoading=$true
    try{$script:DeskSearch.Clear();$script:DeskFilter.SelectedIndex=0}finally{$script:DeskLoading=$false}
    Refresh-DeskWorkspace
    Set-Status 'Pronto. Selecione um processo para consultar ou registrar uma atualização.'
}
function Get-DeskReadable([object]$Value) {
    if($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)){return 'Não informado'}
    return [string]$Value
}
function Format-DeskSummaryMoney([string]$Column,[object]$Value) {
    if($null -eq $Value -or [string]::IsNullOrWhiteSpace([string]$Value)){return 'Não informado'}
    return 'R$ '+(Format-DeskField $Column $Value 'Valores')
}
function Show-DeskSelection {
    if($script:DeskLoading){return}
    $number=Get-DeskSelectedNumber;$script:DeskCurrentNumber=$number
    if(-not $script:DeskConnectionUnavailable -and $number -and $number -ceq $script:DeskDetailLoadedNumber -and $script:DeskDetailRevision -eq $script:DeskDataRevision -and $script:DeskDetailWorkbookPath -ceq $script:PathBox.Text){Update-DeskConnectionState;return}
    $script:DeskDetailLoadedNumber=''
    $script:DeskHistory.Rows.Clear()
    if(-not $script:Wb -or $script:DeskConnectionUnavailable -or -not $number){
        $script:DeskDetailTitle.Text='Selecione um processo';$script:DeskDetailMeta.Text='Os detalhes aparecem aqui.'
        $script:DeskSummary.Text='Escolha um processo na lista para consultar os dados, informar valores ou registrar um andamento.'
        Update-DeskConnectionState;return
    }
    $script:DeskDetailTitle.Text=$number;$script:DeskDetailMeta.Text='Carregando os detalhes...';$script:DeskSummary.Clear()
    try{
        if(-not $script:Excel.Ready){$script:DeskDetailMeta.Text='O Excel está ocupado.';$script:DeskSummary.Text='Conclua a edição ou feche a caixa de aviso no Excel. Depois clique em Recarregar.';return}
        $r=Get-DeskRecord $number;$f=$r.Fields;$m=$r.Money
        $script:DeskDetailTitle.Text=$number
        $script:DeskDetailMeta.Text=(Get-DeskReadable $f.C)+'  |  Prioridade: '+(Get-DeskReadable $f.N)
        $lines=New-Object System.Collections.Generic.List[string]
        $lines.Add('Cadastro');$lines.Add('Escritório: '+(Get-DeskReadable $f.D));$lines.Add('Empresa: '+(Get-DeskReadable $f.E))
        $lines.Add('Consórcio: '+(Get-DeskReadable $f.F));$lines.Add('Tema: '+(Get-DeskReadable $f.G))
        $lines.Add('');$lines.Add('Partes e local');$lines.Add('Parte autora: '+(Get-DeskReadable $f.H));$lines.Add('Parte ré: '+(Get-DeskReadable $f.I))
        $lines.Add('Tribunal: '+(Get-DeskReadable $f.K));$lines.Add('Comarca / órgão: '+(Get-DeskReadable $f.J))
        $lines.Add('');$lines.Add('Fatos e motivos');$lines.Add((Get-DeskReadable $f.M))
        $lines.Add('');$lines.Add('Valores informados');$lines.Add('Causa: '+(Format-DeskSummaryMoney 'B' $m.B))
        $lines.Add('Depósitos: '+(Format-DeskSummaryMoney 'D' $m.D));$lines.Add('Bloqueios em dinheiro: '+(Format-DeskSummaryMoney 'E' $m.E))
        if($f.Q){$lines.Add('');$lines.Add('Motivo da criticidade');$lines.Add([string]$f.Q)}
        if($f.U){$lines.Add('');$lines.Add('Observações');$lines.Add([string]$f.U)}
        $script:DeskSummary.Text=$lines -join "`r`n"
        $script:DeskSummary.SelectAll();$script:DeskSummary.SelectionFont=New-Object Drawing.Font($script:UiFont.FontFamily,9,[Drawing.FontStyle]::Regular);$script:DeskSummary.SelectionColor=$script:DeskInk
        foreach($heading in @('Cadastro','Partes e local','Fatos e motivos','Valores informados','Motivo da criticidade','Observações')){
            $at=$script:DeskSummary.Find($heading,[Windows.Forms.RichTextBoxFinds]::MatchCase)
            if($at -ge 0){$script:DeskSummary.Select($at,$heading.Length);$script:DeskSummary.SelectionFont=New-Object Drawing.Font($script:UiFont.FontFamily,10,[Drawing.FontStyle]::Bold);$script:DeskSummary.SelectionColor=$script:DeskInk}
        }
        $script:DeskSummary.Select(0,0);$script:DeskSummary.ScrollToCaret()
        foreach($h in @($r.History)){$i=$script:DeskHistory.Rows.Add([object[]]@($h.Date,$h.Source,$h.Text,$h.Registered));$script:DeskHistory.Rows[$i].Tag=$h}
        $script:DeskHistory.ClearSelection()
        $script:DeskHistoryHint.Text=if(@($r.History).Count){'Dê dois cliques para ler o andamento completo.'}else{'Ainda não há andamentos. Use Registrar andamento para incluir o primeiro.'}
        $script:DeskDetailLoadedNumber=$number;$script:DeskDetailRevision=$script:DeskDataRevision;$script:DeskDetailWorkbookPath=$script:PathBox.Text
    }catch{$script:DeskDetailMeta.Text='Não foi possível carregar os detalhes.';$script:DeskSummary.Text='Clique em Recarregar para tentar novamente.';Set-Status $_.Exception.Message $true}
    finally{Update-DeskConnectionState}
}
function New-DeskDialog([string]$Title,[string]$Hint,[int]$Width=880,[int]$Height=680) {
    $script:Form.UseWaitCursor=$false
    $d=New-Object Windows.Forms.Form;$d.Text=$Title;$d.StartPosition='CenterParent'
    $area=[Windows.Forms.Screen]::FromControl($script:Form).WorkingArea
    $d.Size=New-Object Drawing.Size([Math]::Min($Width,$area.Width-24),[Math]::Min($Height,$area.Height-24));$d.MinimumSize=New-Object Drawing.Size([Math]::Min(760,$area.Width-24),[Math]::Min(560,$area.Height-24))
    $d.Font=$script:UiFont;$d.BackColor=$script:DeskBackground;$d.AutoScaleDimensions=New-Object Drawing.SizeF(96,96);$d.AutoScaleMode='Dpi'
    $d.Add_Shown({Fit-DeskWindow $this})
    $d.Tag=@{Fields=@{};Loading=$true;Dirty=$false;Saved=$false;Processing=$false;Kind='';Snapshot=$null;Number='';ResultNumber='';Status=$null;Body=$null;Bottom=$null}
    $titlePanel=New-Object Windows.Forms.Panel;$titlePanel.Dock='Top';$titlePanel.Height=78;$titlePanel.Padding=New-Object Windows.Forms.Padding(20,12,20,8)
    $label=New-Object Windows.Forms.Label;$label.Text=$Title;$label.Dock='Top';$label.Height=30;$label.Font=New-Object Drawing.Font($script:UiFont.FontFamily,14,[Drawing.FontStyle]::Bold);$label.ForeColor=$script:DeskInk
    $sub=New-Object Windows.Forms.Label;$sub.Text=$Hint;$sub.Dock='Fill';$sub.ForeColor=[Drawing.Color]::FromArgb(88,105,126)
    $titlePanel.Controls.Add($sub);$titlePanel.Controls.Add($label)
    $status=New-Object Windows.Forms.Label;$status.Dock='Bottom';$status.Height=58;$status.Padding=New-Object Windows.Forms.Padding(20,8,20,8);$status.ForeColor=$script:DeskInk
    $bottom=New-Object Windows.Forms.FlowLayoutPanel;$bottom.Dock='Bottom';$bottom.Height=54;$bottom.Padding=New-Object Windows.Forms.Padding(20,6,12,8)
    $body=New-Object Windows.Forms.Panel;$body.Dock='Fill';$body.Padding=New-Object Windows.Forms.Padding(16,0,16,0)
    $d.Controls.Add($body);$d.Controls.Add($status);$d.Controls.Add($bottom);$d.Controls.Add($titlePanel)
    $d.Tag.Status=$status;$d.Tag.Body=$body;$d.Tag.Bottom=$bottom
    $d.Add_FormClosing({
        param($sender,$e)
        $s=$sender.Tag
        if($s.Processing){$e.Cancel=$true;return}
        if($s.Dirty -and -not $s.Saved){
            $message=if($s.Kind -eq 'ai'){'As alterações já salvas serão mantidas. Deseja sair e descartar o texto e as propostas que ainda não salvou?'}else{'Há informações que ainda não foram salvas. Deseja sair e descartar essa edição?'}
            $answer=[Windows.Forms.MessageBox]::Show($sender,$message,'Sair sem salvar',[Windows.Forms.MessageBoxButtons]::YesNo,[Windows.Forms.MessageBoxIcon]::Question)
            if($answer -ne [Windows.Forms.DialogResult]::Yes){$e.Cancel=$true}
        }
    })
    return $d
}
function New-DeskFieldsPage($Tabs,[string]$Title) {
    $page=New-Object Windows.Forms.TabPage;$page.Text=$Title;$page.BackColor=[Drawing.Color]::White;$page.AutoScroll=$true
    $table=New-Object Windows.Forms.TableLayoutPanel;$table.Dock='Top';$table.AutoSize=$true;$table.ColumnCount=2;$table.RowCount=0
    $table.Padding=New-Object Windows.Forms.Padding(8)
    [void]$table.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle([Windows.Forms.SizeType]::Percent,50)))
    [void]$table.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle([Windows.Forms.SizeType]::Percent,50)))
    $table.Tag=@{Count=0};$page.Controls.Add($table);[void]$Tabs.TabPages.Add($page)
    return $table
}
function Add-DeskField($Dialog,$Table,[string]$Key,[string]$Label,[string]$Kind,[object[]]$Choices,[object]$Value,[bool]$Wide=$false) {
    $legacyDate=$false;$parsedDate=[datetime]::MinValue
    if($Kind -eq 'date' -and -not [string]::IsNullOrWhiteSpace([string]$Value)){
        if(-not [datetime]::TryParseExact([string]$Value,'dd/MM/yyyy',[Globalization.CultureInfo]::GetCultureInfo('pt-BR'),[Globalization.DateTimeStyles]::None,[ref]$parsedDate)){$legacyDate=$true;$Kind='text'}
    }
    $index=[int]$Table.Tag.Count
    if($Wide -and ($index % 2)){ $index++ }
    $row=[int][Math]::Floor($index/2);$column=$index%2
    while($Table.RowCount -le $row){$Table.RowCount++;[void]$Table.RowStyles.Add((New-Object Windows.Forms.RowStyle([Windows.Forms.SizeType]::AutoSize)))}
    $p=New-Object Windows.Forms.Panel;$p.Dock='Fill';$p.Height=if($Kind -eq 'multiline'){128}else{70};$p.Margin=New-Object Windows.Forms.Padding(6,4,6,4)
    $l=New-Object Windows.Forms.Label;$l.Text=$Label;$l.Dock='Top';$l.Height=26;$l.ForeColor=$script:DeskInk
    if($legacyDate){$l.Text=$Label+' (texto existente)';$l.Height=44;$p.Height=88}
    if($Kind -eq 'combo'){
        $c=New-Object Windows.Forms.ComboBox;$c.DropDownStyle='DropDown';$c.AutoCompleteMode='SuggestAppend';$c.AutoCompleteSource='ListItems'
        [void]$c.Items.Add('');foreach($choice in @($Choices)){if($null -ne $choice -and [string]$choice -ne ''){[void]$c.Items.Add([string]$choice)}}
    }elseif($Kind -eq 'date'){
        $c=New-Object Windows.Forms.DateTimePicker;$c.Format='Custom';$c.CustomFormat='dd/MM/yyyy';$c.ShowCheckBox=$true;$c.Checked=$false
    }else{
        $c=New-Object Windows.Forms.TextBox
        if($Kind -eq 'multiline'){$c.Multiline=$true;$c.ScrollBars='Vertical';$c.AcceptsReturn=$true;$c.MaxLength=24000}
        else{$c.MaxLength=6000}
    }
    $c.Dock='Fill';$c.BackColor=[Drawing.Color]::FromArgb(255,250,233)
    if($Kind -eq 'date'){
        if(-not [string]::IsNullOrWhiteSpace([string]$Value)){
            $parsed=[datetime]::MinValue
            if([datetime]::TryParseExact([string]$Value,'dd/MM/yyyy',[Globalization.CultureInfo]::GetCultureInfo('pt-BR'),[Globalization.DateTimeStyles]::None,[ref]$parsed)){$c.Value=$parsed;$c.Checked=$true}
        }
        $c.Add_ValueChanged({$f=$this.FindForm();if($f -and -not $f.Tag.Loading){$f.Tag.Dirty=$true}})
    }else{
        $c.Text=[string]$Value
        if($legacyDate){$c.Tag=@{LegacyDate=$true;OriginalText=[string]$Value}}
        $c.Add_TextChanged({$f=$this.FindForm();if($f -and -not $f.Tag.Loading){$f.Tag.Dirty=$true;if($f.Tag.ContainsKey('AnalysisReady')){$f.Tag.AnalysisReady=$false;$f.Tag.Apply.Enabled=$false}}})
    }
    $p.Controls.Add($c);$p.Controls.Add($l);$Table.Controls.Add($p,$column,$row)
    if($Wide){$Table.SetColumnSpan($p,2);$Table.Tag.Count=$index+2}else{$Table.Tag.Count=$index+1}
    $Dialog.Tag.Fields[$Key]=$c
    return $c
}
function Read-DeskFormFields($Form) {
    $fields=@{}
    foreach($key in $Form.Tag.Fields.Keys){
        $c=$Form.Tag.Fields[$key]
        if($c.Tag -is [hashtable] -and $c.Tag.ContainsKey('LegacyDate') -and [string]$c.Text -ceq [string]$c.Tag.OriginalText){continue}
        if($c -is [Windows.Forms.DateTimePicker]){$fields[$key]=if($c.Checked){$c.Value.ToString('dd/MM/yyyy')}else{''}}
        else{$fields[$key]=[string]$c.Text}
    }
    return $fields
}
function Add-DeskSaveButtons($Dialog) {
    $save=New-DeskButton 'Salvar' 'save' 130 $true;$save.Add_Click({Save-DeskForm ($this.FindForm())})
    $cancel=New-DeskButton 'Cancelar' 'cancel' 115;$cancel.Add_Click({$this.FindForm().Close()})
    $Dialog.Tag.Bottom.Controls.AddRange(@($save,$cancel));$Dialog.CancelButton=$cancel
}
function Save-DeskForm($Form) {
    $s=$Form.Tag;if($s.Processing){return}
    $s.Processing=$true;$Form.UseWaitCursor=$true
    try{
        Test-DeskConnected;$fields=Read-DeskFormFields $Form
        if($s.Kind -eq 'process'){
            $result=Save-DeskProcess $s.Number $fields $s.Snapshot;$s.ResultNumber=[string]$result.Number
        }elseif($s.Kind -eq 'money'){
            if($s.Snapshot.MoneyRow -eq 0 -and [string]::IsNullOrWhiteSpace($fields.T)){$fields.Remove('T')}
            Save-DeskValues $s.Number $fields $s.Snapshot;$s.ResultNumber=$s.Number
        }elseif($s.Kind -eq 'history'){
            Add-DeskHistory $s.Number $fields.Date $fields.Description $fields.Source $fields.Notes;$s.ResultNumber=$s.Number
        }else{throw 'Tipo de formulário inválido.'}
        $s.Saved=$true;$s.Dirty=$false;$Form.DialogResult=[Windows.Forms.DialogResult]::OK
    }catch{$s.Status.Text=$_.Exception.Message;$s.Status.ForeColor=[Drawing.Color]::Firebrick}
    finally{$s.Processing=$false;$Form.UseWaitCursor=$false}
    if($s.Saved){$Form.Close()}
}
function Show-DeskProcess([string]$Number='') {
    $record=if($Number){Get-DeskRecord $Number}else{$null};$lists=Get-DeskLists
    $title=if($Number){'Editar cadastro'}else{'Novo processo'}
    $d=New-DeskDialog $title 'Preencha o essencial na primeira aba. Os demais dados podem ser completados depois.'
    $d.Tag.Kind='process';$d.Tag.Number=$Number;if($record){$d.Tag.Snapshot=$record.Snapshot}
    $tabs=New-Object Windows.Forms.TabControl;$tabs.Dock='Fill';$d.Tag.Body.Controls.Add($tabs)
    $main=New-DeskFieldsPage $tabs 'Dados principais';$detail=New-DeskFieldsPage $tabs 'Informações complementares'
    $specs=@(
        @('A','Número completo *','text',$null,$main,$false),@('B','Natureza *','combo',$lists.Nature,$main,$false),
        @('C','Situação *','combo',$lists.Status,$main,$false),@('D','Escritório *','combo',$lists.Office,$main,$false),
        @('E','Empresa responsável *','combo',$lists.Company,$main,$false),@('F','Consórcio','combo',$lists.Consortium,$main,$false),
        @('G','Tema','combo',$lists.Theme,$main,$false),@('H','Parte autora','text',$null,$main,$false),
        @('I','Parte ré','text',$null,$main,$false),@('M','Fatos e motivos','multiline',$null,$main,$true),
        @('J','Comarca / órgão','text',$null,$detail,$false),@('K','Tribunal','text',$null,$detail,$false),
        @('L','Data de instauração, se conhecida','date',$null,$detail,$false),@('AW','Fase','text',$null,$detail,$false),
        @('AX','Contrato','text',$null,$detail,$false),@('N','Prioridade','combo',$lists.Priority,$detail,$false),
        @('O','Bloqueio vigente?','combo',$lists.YesNo,$detail,$false),@('P','Penhora vigente?','combo',$lists.YesNo,$detail,$false),
        @('Q','Motivo da criticidade','multiline',$null,$detail,$true),@('S','Previsão de encerramento, se informada','date',$null,$detail,$false),
        @('T','Link dos documentos','text',$null,$detail,$false),@('U','Observações','multiline',$null,$detail,$true)
    )
    foreach($x in $specs){$value=if($record){Format-DeskField $x[0] $record.Fields[$x[0]] 'Processos'}else{''};$null=Add-DeskField $d $x[4] $x[0] $x[1] $x[2] $x[3] $value $x[5]}
    if($Number){$d.Tag.Fields.A.ReadOnly=$true;$d.Tag.Fields.A.BackColor=$script:DeskBackground}
    $d.Tag.Status.Text='* Obrigatório no cadastro novo. Datas desmarcadas ficam em branco. Ao editar, apagar um campo remove seu conteúdo.'
    Add-DeskSaveButtons $d;$d.Tag.Loading=$false
    $null=$d.ShowDialog($script:Form);$saved=$d.Tag.Saved;$selected=$d.Tag.ResultNumber;$d.Dispose()
    if($saved){Refresh-DeskWorkspace $selected;Set-Status 'Cadastro salvo. O processo já está disponível para consulta e atualização.'}
}
function Show-DeskValues([string]$Number) {
    $r=Get-DeskRecord $Number
    $d=New-DeskDialog 'Valores do processo' $Number;$d.Tag.Kind='money';$d.Tag.Number=$Number;$d.Tag.Snapshot=$r.Snapshot
    $tabs=New-Object Windows.Forms.TabControl;$tabs.Dock='Fill';$d.Tag.Body.Controls.Add($tabs)
    $main=New-DeskFieldsPage $tabs 'Valores conhecidos';$detail=New-DeskFieldsPage $tabs 'Avaliação e fonte'
    $specs=@(
        @('B','Valor da causa (R$)','text',$null,$main,$false),@('U','Valor atualizado da causa (R$)','text',$null,$main,$false),
        @('C','Condenação / execução (R$)','text',$null,$main,$false),@('D','Saldo de depósitos judiciais (R$)','text',$null,$main,$false),
        @('E','Saldo bloqueado em dinheiro (R$)','text',$null,$main,$false),@('G','Benefício econômico estimado (R$)','text',$null,$main,$false),
        @('H','Probabilidade de êxito (%)','text',$null,$main,$false),@('T','Incluir nos totais?','combo',@('Sim','Revisar','Não'),$main,$false),
        @('Z','Probabilidade de perda informada','combo',@('Remoto','Possível','Provável','Não avaliado'),$detail,$false),@('AA','Duração estimada informada','text',$null,$detail,$false),
        @('J','Data da revisão','date',$null,$detail,$false),@('K','Fonte / observações financeiras','multiline',$null,$detail,$true),@('Y','Descrição do benefício econômico','multiline',$null,$detail,$true)
    )
    foreach($x in $specs){$null=Add-DeskField $d $x[4] $x[0] $x[1] $x[2] $x[3] (Format-DeskField $x[0] $r.Money[$x[0]] 'Valores') $x[5]}
    $d.Tag.Status.Text='Deixe em branco o que não sabe. Zero significa valor conhecido igual a zero. Somente Sim entra nos totais.'
    Add-DeskSaveButtons $d;$d.Tag.Loading=$false;$null=$d.ShowDialog($script:Form);$saved=$d.Tag.Saved;$d.Dispose()
    if($saved){Refresh-DeskWorkspace $Number;Set-Status 'Valores salvos. O painel e a ficha foram recalculados.'}
}
function Show-DeskHistoryForm([string]$Number) {
    $d=New-DeskDialog 'Registrar andamento' $Number 840 640;$d.Tag.Kind='history';$d.Tag.Number=$Number
    $tabs=New-Object Windows.Forms.TabControl;$tabs.Dock='Fill';$d.Tag.Body.Controls.Add($tabs);$p=New-DeskFieldsPage $tabs 'Nova atualização'
    $null=Add-DeskField $d $p 'Date' 'Data em que o fato aconteceu *' 'date' $null ''
    $null=Add-DeskField $d $p 'Source' 'Quem informou / fonte *' 'text' $null ''
    $null=Add-DeskField $d $p 'Description' 'O que aconteceu? *' 'multiline' $null '' $true
    $null=Add-DeskField $d $p 'Notes' 'Texto original ou observações' 'multiline' $null '' $true
    $d.Tag.Status.Text='O registro será acrescentado ao histórico. A data de recebimento é gravada automaticamente.'
    Add-DeskSaveButtons $d;$d.Tag.Loading=$false;$null=$d.ShowDialog($script:Form);$saved=$d.Tag.Saved;$d.Dispose()
    if($saved){Refresh-DeskWorkspace $Number;Set-Status 'Andamento registrado. O histórico anterior foi preservado.'}
}
function Show-DeskHistoryText {
    if($script:DeskHistory.SelectedRows.Count -eq 0){return}
    $h=$script:DeskHistory.SelectedRows[0].Tag
    $d=New-DeskDialog 'Andamento completo' ([string]$h.Date+'  '+[string]$h.Source) 820 620
    $t=New-Object Windows.Forms.TextBox;$t.Dock='Fill';$t.Multiline=$true;$t.ReadOnly=$true;$t.ScrollBars='Vertical';$t.BackColor=[Drawing.Color]::White
    $t.Text=('Registrado em: '+[string]$h.Registered+"`r`n`r`n"+[string]$h.Text+"`r`n`r`n"+[string]$h.Notes);$d.Tag.Body.Controls.Add($t)
    $close=New-DeskButton 'Fechar' 'close';$close.Add_Click({$this.FindForm().Close()});$d.Tag.Bottom.Controls.Add($close)
    $d.Tag.Loading=$false;$null=$d.ShowDialog($script:Form);$d.Dispose()
}
function Sync-DeskAIContext($Form) {
    $fields=Read-DeskFormFields $Form
    foreach($name in @('case','source','office','company','consortium','text')){Write-Central $name $fields[$name]}
    Write-Central 'mode' 'Revisar tudo';Write-Central 'command' ''
}
function Refresh-DeskAIGrid($Form) {
    $g=$Form.Tag.Grid;$g.Rows.Clear()
    foreach($p in $script:Pending){
        $proposed=if($p.Operation -eq 'CreateCase'){'Número: '+$p.Case+'; escritório: '+$p.Value.D+'; empresa: '+$p.Value.E+'; consórcio: '+$p.Value.F}
        elseif($p.Operation -eq 'AppendHistory'){(Get-SafeDisplay $p.Value[1] 'date')+'  '+$p.Value[2]}
        else{Get-SafeDisplay $p.Value $p.Type}
        $i=$g.Rows.Add([object[]]@($false,$p.Case,$p.Field,(Get-SafeDisplay $p.Before $p.Type),$proposed,$p.Evidence,$p.Warning, $(if($p.Applied){'Salvo'}else{'Aguardando'})))
        $row=$g.Rows[$i];$row.Tag=$p;$row.Cells[0].ReadOnly=[bool]$p.Applied
        if($p.Applied){$row.DefaultCellStyle.ForeColor=[Drawing.Color]::Gray}
    }
    $g.ClearSelection();$Form.Tag.Apply.Enabled=$Form.Tag.AnalysisReady -and @($script:Pending | Where-Object {-not $_.Applied}).Count -gt 0
}
function Invoke-DeskAIAction($Button) {
    $form=$Button.FindForm();$s=$form.Tag;if($s.Processing){return}
    $s.Processing=$true;$form.UseWaitCursor=$true
    try{
        Test-DeskConnected
        switch([string]$Button.Tag){
            'analyze'{
                $s.AnalysisReady=$false;$s.Apply.Enabled=$false
                Sync-DeskAIContext $form;$s.Status.ForeColor=$script:DeskInk;$s.Status.Text='Analisando o texto. Aguarde a resposta da IA...';$s.Status.Refresh()
                Analyze-Text;$s.AnalysisReady=$true;Refresh-DeskAIGrid $form;$s.Tabs.SelectedIndex=1
                $s.Status.Text=('{0} alteração(ões) proposta(s). Marque a coluna Salvar e clique em Salvar selecionados.' -f $script:Pending.Count)
                if($script:AnalysisWarnings.Count){$s.Status.Text+=' '+($script:AnalysisWarnings -join ' ')}
            }
            'import'{
                Sync-DeskAIContext $form;Import-Source;$s.Fields.text.Text=Read-Central 'text';$s.Fields.source.Text=Read-Central 'source'
                $s.AnalysisReady=$false;$s.Apply.Enabled=$false;$s.Status.Text='Texto importado. Confira o conteúdo e clique em Analisar texto.';$form.Activate()
            }
            'apply'{
                if(-not $s.AnalysisReady){throw 'O texto ou contexto mudou. Clique em Analisar texto novamente.'}
                $null=$s.Grid.EndEdit();$selected=New-Object System.Collections.Generic.List[object]
                foreach($row in $s.Grid.Rows){$p=$row.Tag;if(-not $p.Applied -and [bool]$row.Cells[0].Value){$selected.Add($p)}}
                if($selected.Count -eq 0){throw 'Marque as alterações que deseja salvar.'}
                Sync-DeskAIContext $form
                foreach($p in $script:Pending){if(-not $p.Applied){$script:Central.Cells.Item($p.PreviewRow,1).Value2='Não'}}
                foreach($p in $selected){$script:Central.Cells.Item($p.PreviewRow,1).Value2='Sim'}
                Apply-Proposals $selected.ToArray() $true;Refresh-DeskAIGrid $form
                $s.Status.ForeColor=$script:DeskInk;$s.Status.Text=$script:StatusLabel.Text
                if(@($script:Pending | Where-Object {-not $_.Applied}).Count -eq 0){$s.Dirty=$false}
            }
        }
    }catch{$s.Status.Text=$_.Exception.Message;$s.Status.ForeColor=[Drawing.Color]::Firebrick}
    finally{$s.Processing=$false;$form.UseWaitCursor=$false}
}
function Show-DeskAI([string]$Number='') {
    $lists=Get-DeskLists;$record=if($Number){Get-DeskRecord $Number}else{$null}
    $d=New-DeskDialog 'Atualizar com IA' 'Cole o relato, confira as propostas e salve apenas o que aprovar.' 1140 720
    $d.Tag.Kind='ai';$d.Tag.AnalysisReady=$false
    $tabs=New-Object Windows.Forms.TabControl;$tabs.Dock='Fill';$d.Tag.Body.Controls.Add($tabs);$d.Tag.Tabs=$tabs
    $entry=New-DeskFieldsPage $tabs '1. Texto e contexto'
    $null=Add-DeskField $d $entry 'case' 'Processo, quando o relato tratar de um só' 'combo' @($script:DeskRows | ForEach-Object {$_.Number}) $Number
    $null=Add-DeskField $d $entry 'source' 'Origem do texto' 'text' $null ''
    $null=Add-DeskField $d $entry 'office' 'Escritório, para cadastro novo' 'combo' $lists.Office $(if($record){$record.Fields.D}else{''})
    $null=Add-DeskField $d $entry 'company' 'Empresa, para cadastro novo' 'combo' $lists.Company $(if($record){$record.Fields.E}else{''})
    $null=Add-DeskField $d $entry 'consortium' 'Consórcio, se houver' 'combo' $lists.Consortium $(if($record){$record.Fields.F}else{''})
    $text=Add-DeskField $d $entry 'text' 'Texto recebido' 'multiline' $null '' $true;$text.Parent.Height=235
    $page=New-Object Windows.Forms.TabPage;$page.Text='2. Conferir alterações';$page.BackColor=[Drawing.Color]::White
    $grid=New-DeskGrid;$grid.ReadOnly=$false;$grid.MultiSelect=$false;$grid.DefaultCellStyle.WrapMode='True';$grid.AutoSizeRowsMode='AllCells';$grid.SelectionMode='CellSelect'
    $check=New-Object Windows.Forms.DataGridViewCheckBoxColumn;$check.Name='Choose';$check.HeaderText='Salvar';$check.Width=56;$check.SortMode='NotSortable';[void]$grid.Columns.Add($check)
    foreach($c in @(@('Case','Processo',190),@('Field','Campo',140),@('Before','Atual',140),@('After','Proposto',210),@('Evidence','Trecho da fonte',230),@('Warning','Observação',180),@('State','Estado',90))){Add-DeskGridColumn $grid $c[0] $c[1] $c[2];$grid.Columns[$c[0]].ReadOnly=$true}
    $grid.Add_CurrentCellDirtyStateChanged({if($this.IsCurrentCellDirty){$null=$this.CommitEdit([Windows.Forms.DataGridViewDataErrorContexts]::Commit)}})
    $page.Controls.Add($grid);$tabs.TabPages.Add($page);$d.Tag.Grid=$grid
    $analyze=New-DeskButton 'Analisar texto' 'analyze' 150 $true;$analyze.Add_Click({Invoke-DeskAIAction $this})
    $apply=New-DeskButton 'Salvar selecionados' 'apply' 180 $true;$apply.Enabled=$false;$apply.Add_Click({Invoke-DeskAIAction $this});$d.Tag.Apply=$apply
    $import=New-DeskButton 'Importar texto' 'import' 145;$import.Add_Click({Invoke-DeskAIAction $this})
    $close=New-DeskButton 'Fechar' 'close';$close.Add_Click({$this.FindForm().Close()})
    $d.Tag.Bottom.Controls.AddRange(@($analyze,$apply,$import,$close))
    $d.Tag.Status.Text='A IA recebe somente o texto e o contexto informados. Nenhuma proposta é salva automaticamente.'
    $d.Tag.Loading=$false;$null=$d.ShowDialog($script:Form);$d.Dispose();Refresh-DeskWorkspace $Number
}
function Show-DeskListForm {
    $d=New-DeskDialog 'Cadastros auxiliares' 'Inclua os nomes que você usa nos formulários.' 780 580
    $types=New-Object Windows.Forms.ComboBox;$types.DropDownStyle='DropDownList';$types.Dock='Top';$types.Height=32
    [void]$types.Items.AddRange([object[]]@('Escritórios','Empresas','Consórcios','Temas'))
    $list=New-Object Windows.Forms.ListBox;$list.Dock='Fill'
    $entry=New-Object Windows.Forms.TextBox;$entry.Dock='Bottom';$entry.Height=32;$entry.BackColor=[Drawing.Color]::FromArgb(255,250,233)
    $hint=New-Object Windows.Forms.Label;$hint.Text='Novo nome';$hint.Dock='Bottom';$hint.Height=25
    $d.Tag.Body.Controls.Add($list);$d.Tag.Body.Controls.Add($hint);$d.Tag.Body.Controls.Add($entry);$d.Tag.Body.Controls.Add($types)
    $d.Tag.Types=$types;$d.Tag.List=$list;$d.Tag.Entry=$entry
    $types.Add_SelectedIndexChanged({
        $f=$this.FindForm();$s=$f.Tag;$keys=@('Office','Company','Consortium','Theme');$names=Get-DeskLists
        $s.List.Items.Clear();foreach($name in @($names[$keys[$this.SelectedIndex]])){[void]$s.List.Items.Add([string]$name)}
    })
    $save=New-DeskButton 'Adicionar' 'add' 140 $true;$save.Add_Click({
        $f=$this.FindForm();$s=$f.Tag
        try{Test-DeskConnected;$keys=@('Office','Company','Consortium','Theme');Add-DeskListValue $keys[$s.Types.SelectedIndex] $s.Entry.Text;$s.Entry.Clear()
            $names=Get-DeskLists;$s.List.Items.Clear();foreach($name in @($names[$keys[$s.Types.SelectedIndex]])){[void]$s.List.Items.Add([string]$name)}
            $s.Status.ForeColor=$script:DeskInk;$s.Status.Text='Nome incluído. Ele já aparece nas listas dos formulários.'
        }catch{$s.Status.Text=$_.Exception.Message;$s.Status.ForeColor=[Drawing.Color]::Firebrick}
    })
    $close=New-DeskButton 'Fechar' 'close';$close.Add_Click({$this.FindForm().Close()});$d.Tag.Bottom.Controls.AddRange(@($save,$close))
    $d.Tag.Loading=$false;$types.SelectedIndex=0;$null=$d.ShowDialog($script:Form);$d.Dispose()
}
function Invoke-DeskMainCommand([string]$Command) {
    switch($Command){
        'new'{Invoke-DeskAction {Show-DeskProcess ''}}
        'edit'{Invoke-DeskAction {$n=Get-DeskSelectedNumber;if(-not $n){throw 'Selecione um processo na lista.'};Show-DeskProcess $n}}
        'history'{Invoke-DeskAction {$n=Get-DeskSelectedNumber;if(-not $n){throw 'Selecione o processo que recebeu a atualização.'};Show-DeskHistoryForm $n}}
        'money'{Invoke-DeskAction {$n=Get-DeskSelectedNumber;if(-not $n){throw 'Selecione o processo para informar os valores.'};Show-DeskValues $n}}
        'ai'{Invoke-DeskAction {Show-DeskAI (Get-DeskSelectedNumber)}}
        'reload'{Invoke-DeskAction {Refresh-DeskWorkspace;Set-Status 'Lista atualizada com os dados da planilha.'}}
        'clear'{Set-DeskQuickFilter 'Todos';$script:DeskSearch.Focus()}
        'choose'{Invoke-DeskAction {$pick=New-Object Windows.Forms.OpenFileDialog;try{$pick.Filter='Planilha Excel (*.xlsx)|*.xlsx';$pick.Title='Escolha o controle jurídico';if($pick.ShowDialog($script:Form) -eq [Windows.Forms.DialogResult]::OK){$script:DeskLastAttemptPath=$pick.FileName;Connect-Workbook $pick.FileName}}finally{$pick.Dispose()}} $false}
        'retry'{Invoke-DeskAction {
            if([string]::IsNullOrWhiteSpace($script:DeskLastAttemptPath)){throw 'Clique em Escolher planilha para selecionar seu arquivo.'}
            Set-Status 'Conectando à planilha. Aguarde...';$script:StatusLabel.Refresh();Connect-Workbook $script:DeskLastAttemptPath
        } $false}
        'excel'{Invoke-DeskAction {$script:Excel.Visible=$true;$script:Wb.Activate();$script:Wb.Worksheets.Item('Painel').Activate();Set-Status 'O painel está aberto no Excel.'}}
        'lists'{Invoke-DeskAction {Show-DeskListForm}}
        'key'{Invoke-DeskAction {Show-KeyDialog} $false}
        'backup'{Invoke-DeskAction {Start-Process explorer.exe -ArgumentList ('"'+$script:BackupDir+'"')}}
        'guide'{Invoke-DeskAction {Start-Process -FilePath (Join-Path $PSScriptRoot 'COMECE-AQUI.html')} $false}
    }
}
function New-DeskMetricCard([string]$Key,[string]$Title,[string]$Filter,[Drawing.Color]$Accent) {
    $card=New-Object Windows.Forms.Panel;$card.Dock='Fill';$card.BackColor=[Drawing.Color]::White
    $card.Margin=New-Object Windows.Forms.Padding(0,0,10,0);$card.Padding=New-Object Windows.Forms.Padding(15,7,12,6);$card.Tag=$Filter
    $stripe=New-Object Windows.Forms.Panel;$stripe.Dock='Left';$stripe.Width=3;$stripe.BackColor=$Accent
    $number=New-Object Windows.Forms.Label;$number.Text='...';$number.Dock='Top';$number.Height=35
    $number.Font=New-Object Drawing.Font($script:UiFont.FontFamily,22,[Drawing.FontStyle]::Bold);$number.ForeColor=$Accent;$number.Tag=$Filter
    $caption=New-Object Windows.Forms.Label;$caption.Text=$Title;$caption.Dock='Fill';$caption.ForeColor=$script:DeskMuted;$caption.Tag=$Filter
    $caption.Font=New-Object Drawing.Font($script:UiFont.FontFamily,9,[Drawing.FontStyle]::Regular)
    $card.Controls.Add($caption);$card.Controls.Add($number);$card.Controls.Add($stripe)
    foreach($control in @($card,$number,$caption)){$control.Cursor=[Windows.Forms.Cursors]::Hand;$control.Add_Click({Set-DeskQuickFilter ([string]$this.Tag)})}
    $script:DeskMetricLabels[$Key]=$number;$script:DeskMetricCards[$Key]=$card
    return $card
}
function Center-DeskMessage($HostPanel) {
    $content=$HostPanel.Tag
    if(-not $content){return}
    $content.Width=[Math]::Min(600,[Math]::Max(260,$HostPanel.ClientSize.Width-48))
    $content.Left=[Math]::Max(12,[int](($HostPanel.ClientSize.Width-$content.Width)/2))
    $content.Top=[Math]::Max(16,[int](($HostPanel.ClientSize.Height-$content.Height)/2))
}
function Start-DeskInterface {
    $script:Form=New-Object Windows.Forms.Form;$script:Form.Text='Controle jurídico';$script:Form.StartPosition='CenterScreen'
    $script:Form.Font=$script:UiFont;$script:Form.AutoScaleDimensions=New-Object Drawing.SizeF(96,96);$script:Form.AutoScaleMode='Dpi'
    $script:Form.Size=New-Object Drawing.Size(1240,800);$script:Form.MinimumSize=New-Object Drawing.Size(980,620)
    $script:Form.BackColor=$script:DeskBackground
    $script:DeskTips=New-Object Windows.Forms.ToolTip;$script:DeskTips.AutoPopDelay=15000

    # Cabeçalho compacto; apenas uma ação principal.
    $header=New-Object Windows.Forms.TableLayoutPanel;$header.Dock='Top';$header.Height=70;$header.Padding=New-Object Windows.Forms.Padding(22,9,14,8)
    $header.ColumnCount=2;$header.RowCount=1
    [void]$header.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle([Windows.Forms.SizeType]::Percent,100)))
    [void]$header.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle([Windows.Forms.SizeType]::Absolute,422)))
    $identity=New-Object Windows.Forms.Panel;$identity.Dock='Fill';$identity.Margin=New-Object Windows.Forms.Padding(0)
    $title=New-Object Windows.Forms.Label;$title.Text='Controle jurídico';$title.Dock='Top';$title.Height=32
    $title.ForeColor=$script:DeskInk;$title.Font=New-Object Drawing.Font($script:UiFont.FontFamily,18,[Drawing.FontStyle]::Bold)
    $script:DeskConnectionLabel=New-Object Windows.Forms.Label;$script:DeskConnectionLabel.Text='Conectando sua planilha...';$script:DeskConnectionLabel.Dock='Fill';$script:DeskConnectionLabel.AutoEllipsis=$true
    $script:DeskConnectionLabel.ForeColor=$script:DeskMuted;$script:DeskConnectionLabel.Font=New-Object Drawing.Font($script:UiFont.FontFamily,9,[Drawing.FontStyle]::Regular)
    $identity.Controls.Add($script:DeskConnectionLabel);$identity.Controls.Add($title)
    $script:PathBox=New-Object Windows.Forms.TextBox;$script:PathBox.ReadOnly=$true;$script:PathBox.Visible=$false
    $script:PathBox.Add_TextChanged({$script:DeskTips.SetToolTip($script:DeskConnectionLabel,$this.Text)})
    $identity.Controls.Add($script:PathBox)
    $primary=New-Object Windows.Forms.FlowLayoutPanel;$primary.Dock='Fill';$primary.Padding=New-Object Windows.Forms.Padding(0,5,0,0);$primary.Margin=New-Object Windows.Forms.Padding(0);$primary.WrapContents=$false
    foreach($spec in @(@('Novo processo','new',145,$true),@('Atualizar com IA','ai',155,$false))){
        $b=New-DeskButton $spec[0] $spec[1] $spec[2] $spec[3];$b.Add_Click({Invoke-DeskMainCommand ([string]$this.Tag)});$primary.Controls.Add($b);$script:DeskWorkbookControls.Add($b)
    }
    $menu=New-Object Windows.Forms.ContextMenuStrip
    foreach($spec in @(@('Escolher planilha','choose'),@('Tentar conexão novamente','retry'),@('Abrir painel no Excel','excel'),@('Cadastros auxiliares','lists'),@('Configurar IA','key'),@('Abrir backups','backup'),@('Como usar','guide'))){
        $item=New-Object Windows.Forms.ToolStripMenuItem;$item.Text=$spec[0];$item.Tag=$spec[1];$item.Add_Click({Invoke-DeskMainCommand ([string]$this.Tag)});[void]$menu.Items.Add($item)
        if($spec[1] -in @('excel','lists','backup')){$script:DeskWorkbookControls.Add($item)}
    }
    $options=New-DeskButton 'Opções' 'options' 90;$options.ContextMenuStrip=$menu;$options.Add_Click({$this.ContextMenuStrip.Show($this,0,$this.Height)});$primary.Controls.Add($options)
    $header.Controls.Add($identity,0,0);$header.Controls.Add($primary,1,0)

    # Os cartões são atalhos para os filtros, calculados com os registros atuais.
    $metrics=New-Object Windows.Forms.TableLayoutPanel;$metrics.Dock='Top';$metrics.Height=80;$metrics.Padding=New-Object Windows.Forms.Padding(22,0,12,12);$metrics.ColumnCount=4;$metrics.RowCount=1
    for($i=0;$i -lt 4;$i++){[void]$metrics.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle([Windows.Forms.SizeType]::Percent,25)))}
    $metricSpecs=@(@('Total','Total de processos','Todos',$script:DeskInk),@('Active','Em acompanhamento','Em acompanhamento',$script:DeskBlue),@('Critical','Críticos','Críticos',[Drawing.Color]::FromArgb(164,55,44)),@('Stale','Sem atualização','Sem atualização',[Drawing.Color]::FromArgb(138,91,21)))
    for($i=0;$i -lt $metricSpecs.Count;$i++){$m=$metricSpecs[$i];$metrics.Controls.Add((New-DeskMetricCard $m[0] $m[1] $m[2] $m[3]),$i,0)}

    $body=New-Object Windows.Forms.Panel;$body.Dock='Fill';$body.Padding=New-Object Windows.Forms.Padding(22,0,22,12)
    $split=New-Object Windows.Forms.SplitContainer;$split.Size=New-Object Drawing.Size(1190,500);$split.SplitterDistance=730;$split.SplitterWidth=12;$split.Dock='Fill'
    $split.Panel1MinSize=490;$split.Panel2MinSize=310;$script:DeskSplit=$split
    $left=New-Object Windows.Forms.Panel;$left.Dock='Fill';$left.BackColor=[Drawing.Color]::White;$left.Padding=New-Object Windows.Forms.Padding(12)
    $listHead=New-Object Windows.Forms.Panel;$listHead.Dock='Top';$listHead.Height=40
    $listTitle=New-Object Windows.Forms.Label;$listTitle.Text='Processos';$listTitle.Dock='Left';$listTitle.Width=150;$listTitle.TextAlign='MiddleLeft';$listTitle.ForeColor=$script:DeskInk;$listTitle.Font=New-Object Drawing.Font($script:UiFont.FontFamily,13,[Drawing.FontStyle]::Bold)
    $reload=New-DeskButton 'Recarregar' 'reload' 108;$reload.Dock='Right';$reload.Add_Click({Invoke-DeskMainCommand 'reload'});$script:DeskWorkbookControls.Add($reload)
    $listHead.Controls.Add($listTitle);$listHead.Controls.Add($reload)
    $filters=New-Object Windows.Forms.TableLayoutPanel;$filters.Dock='Top';$filters.Height=66;$filters.Padding=New-Object Windows.Forms.Padding(0,4,0,8);$filters.ColumnCount=2;$filters.RowCount=2
    [void]$filters.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle([Windows.Forms.SizeType]::Percent,62)))
    [void]$filters.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle([Windows.Forms.SizeType]::Percent,38)))
    [void]$filters.RowStyles.Add((New-Object Windows.Forms.RowStyle([Windows.Forms.SizeType]::Absolute,24)))
    [void]$filters.RowStyles.Add((New-Object Windows.Forms.RowStyle([Windows.Forms.SizeType]::Percent,100)))
    $searchLabel=New-Object Windows.Forms.Label;$searchLabel.Text='Buscar processo, parte ou empresa';$searchLabel.Dock='Fill';$searchLabel.ForeColor=$script:DeskMuted;$searchLabel.Font=New-Object Drawing.Font($script:UiFont.FontFamily,9,[Drawing.FontStyle]::Regular)
    $filterLabel=New-Object Windows.Forms.Label;$filterLabel.Text='Mostrar';$filterLabel.Dock='Fill';$filterLabel.ForeColor=$script:DeskMuted;$filterLabel.Font=$searchLabel.Font
    $script:DeskSearch=New-Object Windows.Forms.TextBox;$script:DeskSearch.Dock='Fill';$script:DeskSearch.Margin=New-Object Windows.Forms.Padding(0,0,10,0);$script:DeskSearch.AccessibleName='Buscar processo, parte, empresa ou escritório'
    $script:DeskFilter=New-Object Windows.Forms.ComboBox;$script:DeskFilter.Dock='Fill';$script:DeskFilter.DropDownStyle='DropDownList';$script:DeskFilter.Margin=New-Object Windows.Forms.Padding(0);$script:DeskFilter.AccessibleName='Filtrar processos'
    [void]$script:DeskFilter.Items.AddRange([object[]]@('Todos','Em acompanhamento','Críticos','Sem atualização','Ativo','Suspenso','Encerrado','Arquivado'));$script:DeskFilter.SelectedIndex=0
    $filters.Controls.Add($searchLabel,0,0);$filters.Controls.Add($filterLabel,1,0);$filters.Controls.Add($script:DeskSearch,0,1);$filters.Controls.Add($script:DeskFilter,1,1)
    $gridHost=New-Object Windows.Forms.Panel;$gridHost.Dock='Fill'
    $script:DeskGrid=New-DeskGrid;$script:DeskGrid.ScrollBars='Vertical';$script:DeskGrid.AccessibleName='Lista de processos'
    foreach($spec in @(@('Number','Processo',128,34),@('Office','Escritório',65,16),@('Company','Empresa',65,17),@('Status','Situação',76,17),@('LastUpdate','Atualização',80,16))){
        Add-DeskGridColumn $script:DeskGrid $spec[0] $spec[1] $spec[2] $true;$script:DeskGrid.Columns[$spec[0]].FillWeight=$spec[3]
    }
    $script:DeskGrid.Add_SelectionChanged({Show-DeskSelection});$script:DeskGrid.Add_CellDoubleClick({param($sender,$e)if($e.RowIndex -ge 0){Invoke-DeskMainCommand 'edit'}})
    $script:DeskCount=New-Object Windows.Forms.Label;$script:DeskCount.Text='Conecte a planilha para carregar os processos';$script:DeskCount.Dock='Bottom';$script:DeskCount.Height=26;$script:DeskCount.Padding=New-Object Windows.Forms.Padding(0,6,0,0);$script:DeskCount.ForeColor=$script:DeskMuted;$script:DeskCount.Font=New-Object Drawing.Font($script:UiFont.FontFamily,9,[Drawing.FontStyle]::Regular)
    $script:DeskEmpty=New-Object Windows.Forms.Panel;$script:DeskEmpty.Dock='Fill';$script:DeskEmpty.BackColor=[Drawing.Color]::White;$script:DeskEmpty.Visible=$false
    $emptyContent=New-Object Windows.Forms.Panel;$emptyContent.Height=172
    $script:DeskEmptyTitle=New-Object Windows.Forms.Label;$script:DeskEmptyTitle.Dock='Top';$script:DeskEmptyTitle.Height=40;$script:DeskEmptyTitle.TextAlign='MiddleCenter';$script:DeskEmptyTitle.Font=New-Object Drawing.Font($script:UiFont.FontFamily,13,[Drawing.FontStyle]::Bold);$script:DeskEmptyTitle.ForeColor=$script:DeskInk
    $script:DeskEmptyText=New-Object Windows.Forms.Label;$script:DeskEmptyText.Dock='Top';$script:DeskEmptyText.Height=64;$script:DeskEmptyText.TextAlign='TopCenter';$script:DeskEmptyText.ForeColor=$script:DeskMuted
    $emptyActions=New-Object Windows.Forms.FlowLayoutPanel;$emptyActions.Dock='Bottom';$emptyActions.Height=48;$emptyActions.FlowDirection='TopDown';$emptyActions.WrapContents=$false
    $script:DeskEmptyButton=New-DeskButton 'Limpar busca' 'clear' 150;$script:DeskEmptyButton.Add_Click({Invoke-DeskMainCommand ([string]$this.Tag)});$emptyActions.Controls.Add($script:DeskEmptyButton)
    $emptyActions.Add_Resize({if($this.Controls.Count){$this.Padding=New-Object Windows.Forms.Padding([Math]::Max(0,[int](($this.Width-$this.Controls[0].Width)/2)),0,0,0)}})
    $emptyContent.Controls.Add($script:DeskEmptyText);$emptyContent.Controls.Add($script:DeskEmptyTitle);$emptyContent.Controls.Add($emptyActions)
    $script:DeskEmpty.Controls.Add($emptyContent);$script:DeskEmpty.Tag=$emptyContent;$script:DeskEmpty.Add_Resize({Center-DeskMessage $this})
    $gridHost.Controls.Add($script:DeskGrid);$gridHost.Controls.Add($script:DeskEmpty)
    $left.Controls.Add($gridHost);$left.Controls.Add($script:DeskCount);$left.Controls.Add($filters);$left.Controls.Add($listHead);$split.Panel1.Controls.Add($left)

    # Ações de edição acompanham o processo selecionado, ao lado de seus dados.
    $right=New-Object Windows.Forms.Panel;$right.Dock='Fill';$right.BackColor=[Drawing.Color]::White;$right.Padding=New-Object Windows.Forms.Padding(13,12,13,10)
    $selectedHead=New-Object Windows.Forms.Panel;$selectedHead.Dock='Top';$selectedHead.Height=86
    $selectedCaption=New-Object Windows.Forms.Label;$selectedCaption.Text='PROCESSO SELECIONADO';$selectedCaption.Dock='Top';$selectedCaption.Height=22;$selectedCaption.ForeColor=$script:DeskMuted;$selectedCaption.Font=New-Object Drawing.Font($script:UiFont.FontFamily,8,[Drawing.FontStyle]::Regular)
    $script:DeskDetailTitle=New-Object Windows.Forms.Label;$script:DeskDetailTitle.Text='Selecione um processo';$script:DeskDetailTitle.Dock='Top';$script:DeskDetailTitle.Height=42;$script:DeskDetailTitle.ForeColor=$script:DeskInk;$script:DeskDetailTitle.Font=New-Object Drawing.Font($script:UiFont.FontFamily,11,[Drawing.FontStyle]::Bold)
    $script:DeskDetailMeta=New-Object Windows.Forms.Label;$script:DeskDetailMeta.Dock='Fill';$script:DeskDetailMeta.Text='Os detalhes aparecem aqui.';$script:DeskDetailMeta.ForeColor=$script:DeskMuted;$script:DeskDetailMeta.Font=New-Object Drawing.Font($script:UiFont.FontFamily,9,[Drawing.FontStyle]::Regular)
    $selectedHead.Controls.Add($script:DeskDetailMeta);$selectedHead.Controls.Add($script:DeskDetailTitle);$selectedHead.Controls.Add($selectedCaption)
    $caseActions=New-Object Windows.Forms.TableLayoutPanel;$caseActions.Dock='Top';$caseActions.Height=84;$caseActions.ColumnCount=2;$caseActions.RowCount=2;$caseActions.Padding=New-Object Windows.Forms.Padding(0,0,0,8)
    for($i=0;$i -lt 2;$i++){[void]$caseActions.ColumnStyles.Add((New-Object Windows.Forms.ColumnStyle([Windows.Forms.SizeType]::Percent,50)));[void]$caseActions.RowStyles.Add((New-Object Windows.Forms.RowStyle([Windows.Forms.SizeType]::Percent,50)))}
    foreach($spec in @(@('Editar cadastro','edit',0,0),@('Valores','money',1,0),@('Registrar andamento','history',0,1))){
        $b=New-DeskButton $spec[0] $spec[1] 120;$b.Dock='Fill';$b.Margin=New-Object Windows.Forms.Padding(0,0,6,6);$b.Add_Click({Invoke-DeskMainCommand ([string]$this.Tag)});$caseActions.Controls.Add($b,$spec[2],$spec[3]);$script:DeskSelectionControls.Add($b)
        if($spec[1] -eq 'history'){$caseActions.SetColumnSpan($b,2)}
    }
    $detail=New-Object Windows.Forms.TabControl;$detail.Dock='Fill';$detail.Padding=New-Object Drawing.Point(14,7)
    $summaryPage=New-Object Windows.Forms.TabPage;$summaryPage.Text='Resumo';$summaryPage.BackColor=[Drawing.Color]::White;$summaryPage.Padding=New-Object Windows.Forms.Padding(8,10,6,6)
    $historyPage=New-Object Windows.Forms.TabPage;$historyPage.Text='Histórico';$historyPage.BackColor=[Drawing.Color]::White;$historyPage.Padding=New-Object Windows.Forms.Padding(4,8,4,4)
    $script:DeskSummary=New-Object Windows.Forms.RichTextBox;$script:DeskSummary.Dock='Fill';$script:DeskSummary.ReadOnly=$true;$script:DeskSummary.BorderStyle='None';$script:DeskSummary.BackColor=[Drawing.Color]::White;$script:DeskSummary.ForeColor=$script:DeskInk
    $script:DeskSummary.Font=New-Object Drawing.Font($script:UiFont.FontFamily,9,[Drawing.FontStyle]::Regular);$script:DeskSummary.DetectUrls=$false;$script:DeskSummary.ScrollBars='Vertical';$script:DeskSummary.Text='Escolha um processo na lista para consultar os dados e registrar atualizações.'
    $summaryPage.Controls.Add($script:DeskSummary)
    $script:DeskHistory=New-DeskGrid;$script:DeskHistory.DefaultCellStyle.WrapMode='True';$script:DeskHistory.AutoSizeRowsMode='AllCells';$script:DeskHistory.ScrollBars='Vertical'
    Add-DeskGridColumn $script:DeskHistory 'Date' 'Data' 72 $true;Add-DeskGridColumn $script:DeskHistory 'Source' 'Fonte' 68 $true;Add-DeskGridColumn $script:DeskHistory 'Text' 'Andamento' 100 $true;Add-DeskGridColumn $script:DeskHistory 'Registered' 'Registrado em' 90
    $script:DeskHistory.Columns['Date'].FillWeight=25;$script:DeskHistory.Columns['Source'].FillWeight=25;$script:DeskHistory.Columns['Text'].FillWeight=50;$script:DeskHistory.Columns['Registered'].Visible=$false
    $script:DeskHistory.Add_CellDoubleClick({param($sender,$e)if($e.RowIndex -ge 0){Invoke-DeskAction {Show-DeskHistoryText}}})
    $script:DeskHistoryHint=New-Object Windows.Forms.Label;$script:DeskHistoryHint.Text='Dê dois cliques para ler o andamento completo.';$script:DeskHistoryHint.Dock='Bottom';$script:DeskHistoryHint.Height=48;$script:DeskHistoryHint.Padding=New-Object Windows.Forms.Padding(4,8,4,0);$script:DeskHistoryHint.ForeColor=$script:DeskMuted;$script:DeskHistoryHint.Font=New-Object Drawing.Font($script:UiFont.FontFamily,8,[Drawing.FontStyle]::Regular)
    $historyPage.Controls.Add($script:DeskHistory);$historyPage.Controls.Add($script:DeskHistoryHint);$detail.TabPages.AddRange(@($summaryPage,$historyPage))
    $right.Controls.Add($detail);$right.Controls.Add($caseActions);$right.Controls.Add($selectedHead);$split.Panel2.Controls.Add($right)

    # Uma conexão ausente tem explicação e saída, nunca uma tabela vazia sem contexto.
    $script:DeskOffline=New-Object Windows.Forms.Panel;$script:DeskOffline.Dock='Fill';$script:DeskOffline.BackColor=[Drawing.Color]::White
    $offlineContent=New-Object Windows.Forms.Panel;$offlineContent.Height=245
    $offlineTitle=New-Object Windows.Forms.Label;$offlineTitle.Text='Vamos conectar sua planilha';$offlineTitle.Dock='Top';$offlineTitle.Height=48;$offlineTitle.TextAlign='MiddleCenter';$offlineTitle.ForeColor=$script:DeskInk;$offlineTitle.Font=New-Object Drawing.Font($script:UiFont.FontFamily,19,[Drawing.FontStyle]::Bold)
    $offlineHint=New-Object Windows.Forms.Label;$offlineHint.Text='Escolha o arquivo do controle para carregar os processos e continuar de onde parou.';$offlineHint.Dock='Top';$offlineHint.Height=62;$offlineHint.TextAlign='TopCenter';$offlineHint.ForeColor=$script:DeskMuted
    $script:DeskOfflineMessage=New-Object Windows.Forms.Label;$script:DeskOfflineMessage.Text='A conexão será iniciada automaticamente.';$script:DeskOfflineMessage.Dock='Fill';$script:DeskOfflineMessage.TextAlign='TopCenter';$script:DeskOfflineMessage.ForeColor=[Drawing.Color]::FromArgb(151,68,38)
    $offlineActions=New-Object Windows.Forms.FlowLayoutPanel;$offlineActions.Dock='Bottom';$offlineActions.Height=50;$offlineActions.WrapContents=$false
    foreach($spec in @(@('Escolher planilha','choose',170,$true),@('Tentar novamente','retry',165,$false))){$b=New-DeskButton $spec[0] $spec[1] $spec[2] $spec[3];$b.Add_Click({Invoke-DeskMainCommand ([string]$this.Tag)});$offlineActions.Controls.Add($b)}
    $offlineActions.Add_Resize({$total=0;foreach($b in $this.Controls){$total+=$b.Width+$b.Margin.Horizontal};$this.Padding=New-Object Windows.Forms.Padding([Math]::Max(0,[int](($this.Width-$total)/2)),0,0,0)})
    $offlineContent.Controls.Add($script:DeskOfflineMessage);$offlineContent.Controls.Add($offlineHint);$offlineContent.Controls.Add($offlineTitle);$offlineContent.Controls.Add($offlineActions)
    $script:DeskOffline.Controls.Add($offlineContent);$script:DeskOffline.Tag=$offlineContent;$script:DeskOffline.Add_Resize({Center-DeskMessage $this})
    $body.Controls.Add($split);$body.Controls.Add($script:DeskOffline)
    $script:StatusLabel=New-Object Windows.Forms.Label;$script:StatusLabel.Dock='Bottom';$script:StatusLabel.Height=48;$script:StatusLabel.Padding=New-Object Windows.Forms.Padding(22,10,22,6);$script:StatusLabel.BackColor=[Drawing.Color]::White;$script:StatusLabel.Text='Abrindo a planilha...';$script:StatusLabel.AutoEllipsis=$true
    $script:StatusLabel.Font=New-Object Drawing.Font($script:UiFont.FontFamily,9,[Drawing.FontStyle]::Regular)
    $script:StatusLabel.Add_TextChanged({$script:DeskTips.SetToolTip($this,$this.Text)})
    $script:Form.Controls.Add($body);$script:Form.Controls.Add($script:StatusLabel);$script:Form.Controls.Add($metrics);$script:Form.Controls.Add($header)
    $script:DeskSearch.Add_TextChanged({Update-DeskFilter});$script:DeskFilter.Add_SelectedIndexChanged({Update-DeskFilter})
    $script:Form.KeyPreview=$true;$script:Form.Add_KeyDown({param($sender,$e)if($e.Control -and $e.KeyCode -eq [Windows.Forms.Keys]::F){$script:DeskSearch.Focus();$e.SuppressKeyPress=$true}elseif($e.KeyCode -eq [Windows.Forms.Keys]::F5){Invoke-DeskMainCommand 'reload';$e.SuppressKeyPress=$true}})

    $script:DeskTimer=New-Object Windows.Forms.Timer;$script:DeskTimer.Interval=1500
    $script:DeskTimer.Add_Tick({
        if($script:Busy -or -not $script:Wb -or $script:DeskConnectionUnavailable){return};$script:Busy=$true
        try{if(-not $script:Excel.Ready){return};$command=(Read-Central 'command').Trim();if($command){Write-Central 'command' '';Run-Command $command;Refresh-DeskWorkspace}}
        catch [Runtime.InteropServices.COMException]{Set-DeskConnectionUnavailable;Set-DeskConnectionError 'A conexão com o Excel foi interrompida. Abra o Excel e clique em Tentar novamente.'}
        catch{Set-Status $_.Exception.Message $true;Write-LocalLog 'ERRO' $_.Exception.Message}
        finally{$script:Busy=$false}
    })
    $script:Form.Add_Shown({
        Fit-DeskWindow $script:Form
        $min=$script:DeskSplit.Panel1MinSize;$max=$script:DeskSplit.Width-$script:DeskSplit.Panel2MinSize-$script:DeskSplit.SplitterWidth
        if($max -ge $min){$script:DeskSplit.SplitterDistance=[Math]::Min($max,[Math]::Max($min,[int]($script:DeskSplit.Width*0.62)))}
        Center-DeskMessage $script:DeskOffline;Center-DeskMessage $script:DeskEmpty;Update-DeskConnectionState
        $script:Busy=$true
        try{
            $candidate=if($script:RequestedWorkbookPath){$script:RequestedWorkbookPath}else{Join-Path $PSScriptRoot 'Controle_Juridico_IA.xlsx'}
            $script:DeskLastAttemptPath=$candidate
            if(Test-Path -LiteralPath $candidate){Connect-Workbook $candidate}else{Set-DeskConnectionError 'Escolha sua planilha do controle para carregar os processos.'}
        }catch{Set-DeskConnectionError $_.Exception.Message}finally{$script:Busy=$false;$script:DeskTimer.Start();Update-DeskConnectionState}
    })
    $script:Form.Add_FormClosing({
        param($sender,$e)
        if($script:Busy){$e.Cancel=$true;Set-Status 'Aguarde a operação terminar antes de fechar.';return}
        $script:DeskTimer.Stop()
        if($script:Central -and -not $script:DeskConnectionUnavailable){try{Write-Central 'connection' 'Assistente fechado';$script:Wb.Save()}catch{}}
        $script:Key=$null
    })
    $null=$script:Form.ShowDialog();$script:DeskTimer.Dispose();$script:DeskTips.Dispose();$menu.Dispose();$script:Form.Dispose()
}
