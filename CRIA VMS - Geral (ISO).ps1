#Requires -RunAsAdministrator
#Requires -Modules Hyper-V
<#
.SYNOPSIS
    Cria máquinas virtuais genéricas no Hyper-V a partir de uma imagem ISO.

.DESCRIPTION
    Diferente dos scripts SRV2025, aqui não existe disco pai: o disco do
    sistema é criado vazio e a instalação sai de uma ISO escolhida no
    formulário. Serve para qualquer sistema - Windows, Linux, pfSense,
    OPNsense, appliances em geral.

    A geração da VM é escolhida no formulário e o script ajusta sozinho o que
    depende dela:

      Geração 2 (UEFI) - Secure Boot configurável (inclusive desligado, que é
                         o que FreeBSD/pfSense e várias distribuições Linux
                         exigem), TPM opcional para Windows 11, boot pela ISO
                         definido via Set-VMFirmware.
      Geração 1 (BIOS) - sem Secure Boot e sem TPM, ordem de boot definida via
                         Set-VMBios. Use para sistemas de 32 bits ou
                         instaladores antigos.

    Nada é criado no host enquanto o formulário não for confirmado.
#>

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ---------------------------------------------------------------------------
# Padrões: ajuste aqui os valores iniciais do formulário
# ---------------------------------------------------------------------------
$padrao = @{
    Prefixo    = ''
    StartupGB  = 4
    MinimaGB   = 1
    MaximaGB   = 8
    Dinamica   = $false
    Cores      = 2
    PastaISO   = 'C:\HYPERV\ISO'
    PastaVHD   = 'C:\HYPERV\VHD'
    PastaVM    = 'C:\HYPERV\MAQUINAS'
    DiscoSoGB  = 60
    DiscoTipo  = 'Dinâmico'
    DiscoAddGB = 127
}

# limite prático da controladora SCSI 0 (64 slots, menos o disco do sistema)
$maxDiscos = 60

# ---------------------------------------------------------------------------
# Dados do host
# ---------------------------------------------------------------------------
$cs       = Get-CimInstance Win32_ComputerSystem
$maxCores = [int]$cs.NumberOfLogicalProcessors
$maxRamGB = [int][math]::Floor($cs.TotalPhysicalMemory / 1GB)
$switches = @(Get-VMSwitch | Select-Object -ExpandProperty Name | Sort-Object)

# lista de discos adicionais planejados
$discos = New-Object System.Collections.Generic.List[psobject]

# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------
function New-Rotulo {
    param([string]$Texto, [int]$X, [int]$Y, [int]$Largura = 140)
    $l = New-Object System.Windows.Forms.Label
    $l.Text     = $Texto
    $l.Location = New-Object System.Drawing.Point($X, ($Y + 3))
    $l.Size     = New-Object System.Drawing.Size($Largura, 20)
    return $l
}

function New-CampoGB {
    param([int]$X, [int]$Y, [decimal]$Valor, [int]$MaximoGB)
    $n = New-Object System.Windows.Forms.NumericUpDown
    $n.Location      = New-Object System.Drawing.Point($X, $Y)
    $n.Size          = New-Object System.Drawing.Size(65, 23)
    $n.DecimalPlaces = 1
    $n.Increment     = 0.5
    $n.Minimum       = 0.5
    $n.Maximum       = [math]::Max(0.5, $MaximoGB)
    $n.Value         = [math]::Min($Valor, $n.Maximum)
    return $n
}

function New-BotaoProcurar {
    param([int]$X, [int]$Y)
    $b = New-Object System.Windows.Forms.Button
    $b.Text     = '...'
    $b.Location = New-Object System.Drawing.Point($X, ($Y - 1))
    $b.Size     = New-Object System.Drawing.Size(40, 25)
    return $b
}

# Hyper-V trabalha com múltiplos de 2 MB
function ConvertTo-BytesMemoria {
    param([decimal]$GB)
    $bytes = [int64]($GB * 1GB)
    return [int64]([math]::Round($bytes / 2MB) * 2MB)
}

function Get-EspacoLivreGB {
    param([string]$Caminho)
    try {
        $raiz = [System.IO.Path]::GetPathRoot($Caminho)
        if ([string]::IsNullOrWhiteSpace($raiz)) { return $null }
        $id = $raiz.TrimEnd('\')
        if ($id -notmatch '^[A-Za-z]:$') { return $null }   # UNC nao entra aqui
        $d = Get-CimInstance Win32_LogicalDisk -Filter "DeviceID='$id'" -ErrorAction Stop
        if ($d) { return [math]::Round($d.FreeSpace / 1GB, 1) }
    } catch { }
    return $null
}

function Select-Pasta {
    param([System.Windows.Forms.TextBox]$Caixa)
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.SelectedPath = $Caixa.Text
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $Caixa.Text = $dlg.SelectedPath
    }
}

# ---------------------------------------------------------------------------
# Formulário
# ---------------------------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Criação de VM a partir de ISO - Hyper-V'
$form.Size            = New-Object System.Drawing.Size(600, 655)
$form.StartPosition   = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox     = $false
$form.MinimizeBox     = $false
$form.Font            = New-Object System.Drawing.Font('Segoe UI', 9)

$tabs          = New-Object System.Windows.Forms.TabControl
$tabs.Location = New-Object System.Drawing.Point(12, 12)
$tabs.Size     = New-Object System.Drawing.Size(560, 500)

$tabVM         = New-Object System.Windows.Forms.TabPage
$tabVM.Text    = 'Máquina virtual'
$tabDisco      = New-Object System.Windows.Forms.TabPage
$tabDisco.Text = 'Disco e mídia'
$tabOpc        = New-Object System.Windows.Forms.TabPage
$tabOpc.Text   = 'Opções'
$tabs.Controls.AddRange(@($tabVM, $tabDisco, $tabOpc))
$form.Controls.Add($tabs)

# ===========================================================================
# Aba 1 - Máquina virtual
# ===========================================================================
$y = 15
$txtNome          = New-Object System.Windows.Forms.TextBox
$txtNome.Location = New-Object System.Drawing.Point(155, $y)
$txtNome.Size     = New-Object System.Drawing.Size(385, 23)
$tabVM.Controls.AddRange(@((New-Rotulo 'Nome da VM:' 12 $y), $txtNome))

$y = 45
$txtPrefixo          = New-Object System.Windows.Forms.TextBox
$txtPrefixo.Location = New-Object System.Drawing.Point(155, $y)
$txtPrefixo.Size     = New-Object System.Drawing.Size(150, 23)
$txtPrefixo.Text     = $padrao.Prefixo
$tabVM.Controls.AddRange(@(
    (New-Rotulo 'Prefixo (opcional):' 12 $y),
    $txtPrefixo,
    (New-Rotulo 'ex.: LAB, SRV2025' 315 $y 200)))

$y = 75
$cboGeracao               = New-Object System.Windows.Forms.ComboBox
$cboGeracao.Location      = New-Object System.Drawing.Point(155, $y)
$cboGeracao.Size          = New-Object System.Drawing.Size(200, 23)
$cboGeracao.DropDownStyle = 'DropDownList'
[void]$cboGeracao.Items.AddRange(@('Geração 2 (UEFI)', 'Geração 1 (BIOS)'))
$cboGeracao.SelectedIndex = 0
$tabVM.Controls.AddRange(@((New-Rotulo 'Geração:' 12 $y), $cboGeracao))

$y = 105
$numCores          = New-Object System.Windows.Forms.NumericUpDown
$numCores.Location = New-Object System.Drawing.Point(155, $y)
$numCores.Size     = New-Object System.Drawing.Size(80, 23)
$numCores.Minimum  = 1
$numCores.Maximum  = $maxCores
$numCores.Value    = [math]::Min($padrao.Cores, $maxCores)
$tabVM.Controls.AddRange(@(
    (New-Rotulo 'Processadores:' 12 $y),
    $numCores,
    (New-Rotulo "(host: $maxCores lógicos)" 240 $y 200)))

$y = 135
$cboSwitch               = New-Object System.Windows.Forms.ComboBox
$cboSwitch.Location      = New-Object System.Drawing.Point(155, $y)
$cboSwitch.Size          = New-Object System.Drawing.Size(385, 23)
$cboSwitch.DropDownStyle = 'DropDownList'
[void]$cboSwitch.Items.Add('<sem conexão de rede>')
if ($switches.Count -gt 0) {
    [void]$cboSwitch.Items.AddRange($switches)
    $cboSwitch.SelectedIndex = 1
} else {
    $cboSwitch.SelectedIndex = 0
}
$tabVM.Controls.AddRange(@((New-Rotulo 'Switch virtual:' 12 $y), $cboSwitch))

$y = 165
$chkVlan          = New-Object System.Windows.Forms.CheckBox
$chkVlan.Text     = 'VLAN de acesso:'
$chkVlan.Location = New-Object System.Drawing.Point(155, $y)
$chkVlan.Size     = New-Object System.Drawing.Size(130, 22)

$numVlan          = New-Object System.Windows.Forms.NumericUpDown
$numVlan.Location = New-Object System.Drawing.Point(290, ($y - 2))
$numVlan.Size     = New-Object System.Drawing.Size(70, 23)
$numVlan.Minimum  = 1
$numVlan.Maximum  = 4094
$numVlan.Value    = 10
$numVlan.Enabled  = $false
$chkVlan.Add_CheckedChanged({ $numVlan.Enabled = $chkVlan.Checked })
$tabVM.Controls.AddRange(@((New-Rotulo 'Rede:' 12 $y), $chkVlan, $numVlan))

# --- Memória ---------------------------------------------------------------
$grpMem          = New-Object System.Windows.Forms.GroupBox
$grpMem.Text     = 'Memória'
$grpMem.Location = New-Object System.Drawing.Point(10, 200)
$grpMem.Size     = New-Object System.Drawing.Size(528, 105)

$chkDinamica          = New-Object System.Windows.Forms.CheckBox
$chkDinamica.Text     = 'Memória dinâmica'
$chkDinamica.Location = New-Object System.Drawing.Point(15, 22)
$chkDinamica.Size     = New-Object System.Drawing.Size(200, 22)
$chkDinamica.Checked  = $padrao.Dinamica

$numStartup = New-CampoGB 90  50 $padrao.StartupGB $maxRamGB
$numMin     = New-CampoGB 250 50 $padrao.MinimaGB  $maxRamGB
$numMax     = New-CampoGB 410 50 $padrao.MaximaGB  $maxRamGB

$lblMemInfo           = New-Object System.Windows.Forms.Label
$lblMemInfo.Location  = New-Object System.Drawing.Point(15, 78)
$lblMemInfo.Size      = New-Object System.Drawing.Size(500, 20)
$lblMemInfo.ForeColor = [System.Drawing.Color]::DimGray

$grpMem.Controls.AddRange(@(
    $chkDinamica,
    (New-Rotulo 'Inicial (GB):' 15  50 75), $numStartup,
    (New-Rotulo 'Mínima (GB):' 175 50 75), $numMin,
    (New-Rotulo 'Máxima (GB):' 335 50 75), $numMax,
    $lblMemInfo))
$tabVM.Controls.Add($grpMem)

# ===========================================================================
# Aba 2 - Disco e mídia
# ===========================================================================
$y = 15
$txtISO          = New-Object System.Windows.Forms.TextBox
$txtISO.Location = New-Object System.Drawing.Point(155, $y)
$txtISO.Size     = New-Object System.Drawing.Size(340, 23)
$btnISO          = New-BotaoProcurar 500 $y
$btnISO.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = 'Imagens ISO (*.iso)|*.iso|Todos os arquivos (*.*)|*.*'
    $inicial = if (Test-Path -LiteralPath $txtISO.Text) { Split-Path -Parent $txtISO.Text }
               else { $padrao.PastaISO }
    if (Test-Path -LiteralPath $inicial) { $dlg.InitialDirectory = $inicial }
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtISO.Text = $dlg.FileName
    }
})
$lblISOdica           = New-Object System.Windows.Forms.Label
$lblISOdica.Text      = 'Deixe vazio para criar a VM sem mídia e instalar depois.'
$lblISOdica.Location  = New-Object System.Drawing.Point(155, ($y + 26))
$lblISOdica.Size      = New-Object System.Drawing.Size(385, 18)
$lblISOdica.ForeColor = [System.Drawing.Color]::DimGray
$tabDisco.Controls.AddRange(@((New-Rotulo 'Imagem ISO:' 12 $y), $txtISO, $btnISO, $lblISOdica))

$y = 60
$txtPastaVM          = New-Object System.Windows.Forms.TextBox
$txtPastaVM.Location = New-Object System.Drawing.Point(155, $y)
$txtPastaVM.Size     = New-Object System.Drawing.Size(340, 23)
$txtPastaVM.Text     = $padrao.PastaVM
$btnPastaVM          = New-BotaoProcurar 500 $y
$btnPastaVM.Add_Click({ Select-Pasta $txtPastaVM })
$tabDisco.Controls.AddRange(@((New-Rotulo 'Pasta das VMs:' 12 $y), $txtPastaVM, $btnPastaVM))

$y = 90
$txtPastaVHD          = New-Object System.Windows.Forms.TextBox
$txtPastaVHD.Location = New-Object System.Drawing.Point(155, $y)
$txtPastaVHD.Size     = New-Object System.Drawing.Size(340, 23)
$txtPastaVHD.Text     = $padrao.PastaVHD
$btnPastaVHD          = New-BotaoProcurar 500 $y
$btnPastaVHD.Add_Click({ Select-Pasta $txtPastaVHD })
$tabDisco.Controls.AddRange(@((New-Rotulo 'Pasta dos VHDs:' 12 $y), $txtPastaVHD, $btnPastaVHD))

# --- Disco do sistema ------------------------------------------------------
$grpSo          = New-Object System.Windows.Forms.GroupBox
$grpSo.Text     = 'Disco do sistema'
$grpSo.Location = New-Object System.Drawing.Point(10, 125)
$grpSo.Size     = New-Object System.Drawing.Size(528, 60)

$numDiscoSo          = New-Object System.Windows.Forms.NumericUpDown
$numDiscoSo.Location = New-Object System.Drawing.Point(110, 22)
$numDiscoSo.Size     = New-Object System.Drawing.Size(70, 23)
$numDiscoSo.Minimum  = 8
$numDiscoSo.Maximum  = 65536
$numDiscoSo.Value    = $padrao.DiscoSoGB

$cboTipoSo               = New-Object System.Windows.Forms.ComboBox
$cboTipoSo.Location      = New-Object System.Drawing.Point(245, 22)
$cboTipoSo.Size          = New-Object System.Drawing.Size(110, 23)
$cboTipoSo.DropDownStyle = 'DropDownList'
[void]$cboTipoSo.Items.AddRange(@('Dinâmico', 'Fixo'))
$cboTipoSo.SelectedItem  = $padrao.DiscoTipo

$grpSo.Controls.AddRange(@(
    (New-Rotulo 'Tamanho (GB):' 15 25 90), $numDiscoSo,
    (New-Rotulo 'Tipo:' 200 25 40), $cboTipoSo))
$tabDisco.Controls.Add($grpSo)

# --- Discos adicionais -----------------------------------------------------
$grpAdd          = New-Object System.Windows.Forms.GroupBox
$grpAdd.Text     = 'Discos adicionais'
$grpAdd.Location = New-Object System.Drawing.Point(10, 193)
$grpAdd.Size     = New-Object System.Drawing.Size(528, 85)

$numQtd          = New-Object System.Windows.Forms.NumericUpDown
$numQtd.Location = New-Object System.Drawing.Point(95, 22)
$numQtd.Size     = New-Object System.Drawing.Size(55, 23)
$numQtd.Minimum  = 1
$numQtd.Maximum  = $maxDiscos
$numQtd.Value    = 1

$numTamanho          = New-Object System.Windows.Forms.NumericUpDown
$numTamanho.Location = New-Object System.Drawing.Point(258, 22)
$numTamanho.Size     = New-Object System.Drawing.Size(65, 23)
$numTamanho.Minimum  = 1
$numTamanho.Maximum  = 65536
$numTamanho.Value    = $padrao.DiscoAddGB

$cboTipo               = New-Object System.Windows.Forms.ComboBox
$cboTipo.Location      = New-Object System.Drawing.Point(378, 22)
$cboTipo.Size          = New-Object System.Drawing.Size(100, 23)
$cboTipo.DropDownStyle = 'DropDownList'
[void]$cboTipo.Items.AddRange(@('Dinâmico', 'Fixo'))
$cboTipo.SelectedItem  = $padrao.DiscoTipo

$lblDica           = New-Object System.Windows.Forms.Label
$lblDica.Text      = 'Clique em Adicionar quantas vezes precisar - dá para misturar tamanhos.'
$lblDica.Location  = New-Object System.Drawing.Point(15, 55)
$lblDica.Size      = New-Object System.Drawing.Size(350, 20)
$lblDica.ForeColor = [System.Drawing.Color]::DimGray

$btnAddDisco          = New-Object System.Windows.Forms.Button
$btnAddDisco.Text     = 'Adicionar'
$btnAddDisco.Location = New-Object System.Drawing.Point(378, 50)
$btnAddDisco.Size     = New-Object System.Drawing.Size(100, 25)

$grpAdd.Controls.AddRange(@(
    (New-Rotulo 'Quantidade:' 15 22 75),    $numQtd,
    (New-Rotulo 'Tamanho (GB):' 165 22 90), $numTamanho,
    (New-Rotulo 'Tipo:' 335 22 40),         $cboTipo,
    $lblDica, $btnAddDisco))
$tabDisco.Controls.Add($grpAdd)

$lstDiscos               = New-Object System.Windows.Forms.ListView
$lstDiscos.Location      = New-Object System.Drawing.Point(10, 285)
$lstDiscos.Size          = New-Object System.Drawing.Size(528, 115)
$lstDiscos.View          = 'Details'
$lstDiscos.FullRowSelect = $true
$lstDiscos.GridLines     = $true
[void]$lstDiscos.Columns.Add('#', 35)
[void]$lstDiscos.Columns.Add('Arquivo', 300)
[void]$lstDiscos.Columns.Add('Tamanho', 80)
[void]$lstDiscos.Columns.Add('Tipo', 90)
$tabDisco.Controls.Add($lstDiscos)

$btnRemover          = New-Object System.Windows.Forms.Button
$btnRemover.Text     = 'Remover selecionado'
$btnRemover.Location = New-Object System.Drawing.Point(10, 406)
$btnRemover.Size     = New-Object System.Drawing.Size(150, 25)

$btnLimpar          = New-Object System.Windows.Forms.Button
$btnLimpar.Text     = 'Limpar lista'
$btnLimpar.Location = New-Object System.Drawing.Point(168, 406)
$btnLimpar.Size     = New-Object System.Drawing.Size(110, 25)

$lblTotais          = New-Object System.Windows.Forms.Label
$lblTotais.Location = New-Object System.Drawing.Point(10, 436)
$lblTotais.Size     = New-Object System.Drawing.Size(528, 32)

$tabDisco.Controls.AddRange(@($btnRemover, $btnLimpar, $lblTotais))

# ===========================================================================
# Aba 3 - Opções
# ===========================================================================
$grpFw          = New-Object System.Windows.Forms.GroupBox
$grpFw.Text     = 'Firmware'
$grpFw.Location = New-Object System.Drawing.Point(10, 15)
$grpFw.Size     = New-Object System.Drawing.Size(528, 85)

$cboSecureBoot               = New-Object System.Windows.Forms.ComboBox
$cboSecureBoot.Location      = New-Object System.Drawing.Point(110, 22)
$cboSecureBoot.Size          = New-Object System.Drawing.Size(280, 23)
$cboSecureBoot.DropDownStyle = 'DropDownList'
[void]$cboSecureBoot.Items.AddRange(@(
    'Ligado - Microsoft Windows',
    'Ligado - Microsoft UEFI CA (Linux)',
    'Desligado'))
$cboSecureBoot.SelectedIndex = 0

$chkTPM          = New-Object System.Windows.Forms.CheckBox
$chkTPM.Text     = 'Habilitar TPM (necessário para Windows 11)'
$chkTPM.Location = New-Object System.Drawing.Point(15, 52)
$chkTPM.Size     = New-Object System.Drawing.Size(400, 22)

$grpFw.Controls.AddRange(@((New-Rotulo 'Secure Boot:' 15 25 90), $cboSecureBoot, $chkTPM))
$tabOpc.Controls.Add($grpFw)

$lblFwAviso           = New-Object System.Windows.Forms.Label
$lblFwAviso.Location  = New-Object System.Drawing.Point(12, 105)
$lblFwAviso.Size      = New-Object System.Drawing.Size(528, 34)
$lblFwAviso.ForeColor = [System.Drawing.Color]::DimGray
$tabOpc.Controls.Add($lblFwAviso)

$grp          = New-Object System.Windows.Forms.GroupBox
$grp.Text     = 'Máquina'
$grp.Location = New-Object System.Drawing.Point(10, 145)
$grp.Size     = New-Object System.Drawing.Size(528, 105)

$chkNested          = New-Object System.Windows.Forms.CheckBox
$chkNested.Text     = 'Virtualização aninhada'
$chkNested.Location = New-Object System.Drawing.Point(15, 25)
$chkNested.Size     = New-Object System.Drawing.Size(240, 22)

$chkMac          = New-Object System.Windows.Forms.CheckBox
$chkMac.Text     = 'MAC address spoofing'
$chkMac.Location = New-Object System.Drawing.Point(15, 50)
$chkMac.Size     = New-Object System.Drawing.Size(240, 22)

$chkGuest          = New-Object System.Windows.Forms.CheckBox
$chkGuest.Text     = 'Serviço de convidado'
$chkGuest.Location = New-Object System.Drawing.Point(15, 75)
$chkGuest.Size     = New-Object System.Drawing.Size(240, 22)
$chkGuest.Checked  = $true

$chkProducao          = New-Object System.Windows.Forms.CheckBox
$chkProducao.Text     = 'Checkpoint de produção'
$chkProducao.Location = New-Object System.Drawing.Point(290, 25)
$chkProducao.Size     = New-Object System.Drawing.Size(230, 22)
$chkProducao.Checked  = $true

$chkAutoChk          = New-Object System.Windows.Forms.CheckBox
$chkAutoChk.Text     = 'Desativar checkpoints automáticos'
$chkAutoChk.Location = New-Object System.Drawing.Point(290, 50)
$chkAutoChk.Size     = New-Object System.Drawing.Size(230, 22)
$chkAutoChk.Checked  = $true

$chkIniciar          = New-Object System.Windows.Forms.CheckBox
$chkIniciar.Text     = 'Iniciar a VM ao final'
$chkIniciar.Location = New-Object System.Drawing.Point(290, 75)
$chkIniciar.Size     = New-Object System.Drawing.Size(230, 22)

$grp.Controls.AddRange(@($chkNested, $chkMac, $chkGuest, $chkProducao, $chkAutoChk, $chkIniciar))
$tabOpc.Controls.Add($grp)

# ===========================================================================
# Nomes derivados e atualização da lista
# ===========================================================================
function Get-Geracao { if ($cboGeracao.SelectedIndex -eq 0) { 2 } else { 1 } }

function Get-NomeVM { "$($txtPrefixo.Text.Trim()) $($txtNome.Text.Trim())".Trim() }

function Get-BaseArquivo {
    $p = $txtPrefixo.Text.Trim()
    $n = $txtNome.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($p)) { return $n }
    return "$p-$n"
}

function Get-CaminhoVHD {
    Join-Path $txtPastaVHD.Text ('{0}-so.vhdx' -f (Get-BaseArquivo))
}

function Get-CaminhoDisco {
    param([int]$Indice)
    Join-Path $txtPastaVHD.Text ('{0}-DISCO{1:D2}.vhdx' -f (Get-BaseArquivo), $Indice)
}

$atualizarTotais = {
    $soGB   = [int]$numDiscoSo.Value
    $soFixo = ($cboTipoSo.SelectedItem -eq 'Fixo')

    $total = $soGB + (($discos | Measure-Object -Property TamanhoGB -Sum).Sum)
    $fixos = ($discos | Where-Object { $_.Tipo -eq 'Fixo' } | Measure-Object -Property TamanhoGB -Sum).Sum
    if (-not $fixos) { $fixos = 0 }
    if ($soFixo) { $fixos += $soGB }

    $texto = "Disco do sistema + $($discos.Count) adicional(is) = $total GB provisionados"
    if ($fixos -gt 0) { $texto += " ($fixos GB alocados de imediato, em discos fixos)" }

    $livre = Get-EspacoLivreGB $txtPastaVHD.Text
    if ($null -ne $livre) {
        $texto += ".`r`nLivre no destino: $livre GB."
        if ($fixos -gt $livre) {
            $lblTotais.ForeColor = [System.Drawing.Color]::Firebrick
            $texto += ' Espaço insuficiente para os discos fixos.'
        } else {
            $lblTotais.ForeColor = [System.Drawing.SystemColors]::ControlText
        }
    } else {
        $lblTotais.ForeColor = [System.Drawing.SystemColors]::ControlText
        $texto += '.'
    }
    $lblTotais.Text = $texto
}

$renderDiscos = {
    $lstDiscos.BeginUpdate()
    $lstDiscos.Items.Clear()
    for ($i = 0; $i -lt $discos.Count; $i++) {
        $d    = $discos[$i]
        $item = New-Object System.Windows.Forms.ListViewItem(('{0:D2}' -f ($i + 1)))
        [void]$item.SubItems.Add((Split-Path -Leaf (Get-CaminhoDisco ($i + 1))))
        [void]$item.SubItems.Add("$($d.TamanhoGB) GB")
        [void]$item.SubItems.Add($d.Tipo)
        [void]$lstDiscos.Items.Add($item)
    }
    $lstDiscos.EndUpdate()
    & $atualizarTotais
}

$btnAddDisco.Add_Click({
    $qtd = [int]$numQtd.Value
    if (($discos.Count + $qtd) -gt $maxDiscos) {
        [void][System.Windows.Forms.MessageBox]::Show(
            "A controladora SCSI comporta até $maxDiscos discos além do disco do sistema.",
            'Limite de discos',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }
    for ($i = 1; $i -le $qtd; $i++) {
        $discos.Add([pscustomobject]@{
            TamanhoGB = [int]$numTamanho.Value
            Tipo      = [string]$cboTipo.SelectedItem
        })
    }
    & $renderDiscos
})

$btnRemover.Add_Click({
    if ($lstDiscos.SelectedIndices.Count -eq 0) { return }
    foreach ($idx in (@($lstDiscos.SelectedIndices) | Sort-Object -Descending)) {
        $discos.RemoveAt($idx)
    }
    & $renderDiscos
})

$btnLimpar.Add_Click({
    $discos.Clear()
    & $renderDiscos
})

$numDiscoSo.Add_ValueChanged($atualizarTotais)
$cboTipoSo.Add_SelectedIndexChanged($atualizarTotais)

# --- Regras da geração -----------------------------------------------------
# Secure Boot e TPM só existem na geração 2; a geração 1 usa Set-VMBios para
# a ordem de boot e não tem firmware configurável.
$atualizarGeracao = {
    $gen2 = ((Get-Geracao) -eq 2)
    $cboSecureBoot.Enabled = $gen2
    $chkTPM.Enabled        = $gen2
    if (-not $gen2) { $chkTPM.Checked = $false }
    $lblFwAviso.Text = if ($gen2) {
        'Geração 2: para pfSense/FreeBSD e várias distribuições Linux, escolha Desligado ou Microsoft UEFI CA - o template Windows impede o boot.'
    } else {
        'Geração 1 (BIOS): sem Secure Boot e sem TPM. Use para sistemas de 32 bits ou instaladores antigos.'
    }
}
$cboGeracao.Add_SelectedIndexChanged($atualizarGeracao)
& $atualizarGeracao

# --- Regras entre memória e virtualização aninhada -------------------------
# A virtualização aninhada exige memória estática, então ela desliga (e trava)
# a memória dinâmica; os campos mínima/máxima só valem no modo dinâmico.
$atualizarMemoria = {
    if ($chkNested.Checked) {
        $chkDinamica.Checked = $false
        $chkDinamica.Enabled = $false
        $lblMemInfo.Text = "Host: $maxRamGB GB. A virtualização aninhada exige memória estática."
    } else {
        $chkDinamica.Enabled = $true
        $lblMemInfo.Text = "Host: $maxRamGB GB."
    }
    $numMin.Enabled = $chkDinamica.Checked
    $numMax.Enabled = $chkDinamica.Checked
}
$chkNested.Add_CheckedChanged($atualizarMemoria)
$chkDinamica.Add_CheckedChanged($atualizarMemoria)
& $atualizarMemoria

# --- Prévia ----------------------------------------------------------------
$lblPreview           = New-Object System.Windows.Forms.Label
$lblPreview.Location  = New-Object System.Drawing.Point(12, 522)
$lblPreview.Size      = New-Object System.Drawing.Size(560, 40)
$lblPreview.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblPreview)

$atualizarPreview = {
    if ([string]::IsNullOrWhiteSpace($txtNome.Text)) {
        $lblPreview.Text = ''
    } else {
        $lblPreview.Text = "VM:   $(Get-NomeVM)   (geração $(Get-Geracao))`r`nVHDX: $(Get-CaminhoVHD)"
    }
    & $renderDiscos   # os nomes dos discos derivam do nome da VM
}
$txtNome.Add_TextChanged($atualizarPreview)
$txtPrefixo.Add_TextChanged($atualizarPreview)
$txtPastaVHD.Add_TextChanged($atualizarPreview)
$cboGeracao.Add_SelectedIndexChanged($atualizarPreview)
& $renderDiscos

# --- Botões ----------------------------------------------------------------
$btnOk          = New-Object System.Windows.Forms.Button
$btnOk.Text     = 'Criar VM'
$btnOk.Location = New-Object System.Drawing.Point(372, 570)
$btnOk.Size     = New-Object System.Drawing.Size(95, 30)

$btnCancel              = New-Object System.Windows.Forms.Button
$btnCancel.Text         = 'Cancelar'
$btnCancel.Location     = New-Object System.Drawing.Point(477, 570)
$btnCancel.Size         = New-Object System.Drawing.Size(95, 30)
$btnCancel.DialogResult = [System.Windows.Forms.DialogResult]::Cancel

$btnOk.Add_Click({
    $erros = New-Object System.Collections.Generic.List[string]

    $nome = $txtNome.Text.Trim()
    if ([string]::IsNullOrWhiteSpace($nome)) {
        $erros.Add('Informe o nome da VM.')
    } elseif ($nome -match '[\\/:*?"<>|]') {
        $erros.Add('O nome da VM contém caracteres inválidos ( \ / : * ? " < > | ).')
    } elseif (Get-VM -Name (Get-NomeVM) -ErrorAction SilentlyContinue) {
        $erros.Add("Já existe uma VM chamada '$(Get-NomeVM)'.")
    }

    if (-not [string]::IsNullOrWhiteSpace($txtISO.Text)) {
        if (-not (Test-Path -LiteralPath $txtISO.Text -PathType Leaf)) {
            $erros.Add("Imagem ISO não encontrada: $($txtISO.Text)")
        } elseif ([System.IO.Path]::GetExtension($txtISO.Text) -ne '.iso') {
            $erros.Add('O arquivo de mídia informado não tem extensão .iso.')
        }
    }

    if ($chkDinamica.Checked) {
        if ($numMin.Value -gt $numStartup.Value) {
            $erros.Add('A memória mínima não pode ser maior que a inicial.')
        }
        if ($numMax.Value -lt $numStartup.Value) {
            $erros.Add('A memória máxima não pode ser menor que a inicial.')
        }
    }

    if ($nome -and (Test-Path -LiteralPath (Get-CaminhoVHD))) {
        $erros.Add("Já existe um VHDX em: $(Get-CaminhoVHD)")
    }

    if ($nome) {
        $existentes = @()
        for ($i = 1; $i -le $discos.Count; $i++) {
            $c = Get-CaminhoDisco $i
            if (Test-Path -LiteralPath $c) { $existentes += (Split-Path -Leaf $c) }
        }
        if ($existentes.Count -gt 0) {
            $erros.Add("Já existem discos com estes nomes no destino: $($existentes -join ', ')")
        }
    }

    # discos fixos precisam do espaço na hora da criação
    $fixos = ($discos | Where-Object { $_.Tipo -eq 'Fixo' } | Measure-Object -Property TamanhoGB -Sum).Sum
    if (-not $fixos) { $fixos = 0 }
    if ($cboTipoSo.SelectedItem -eq 'Fixo') { $fixos += [int]$numDiscoSo.Value }
    $livre = Get-EspacoLivreGB $txtPastaVHD.Text
    if ($fixos -gt 0 -and $livre -and ($fixos -gt $livre)) {
        $erros.Add("Os discos fixos somam $fixos GB e há apenas $livre GB livres no destino.")
    }

    if ($erros.Count -gt 0) {
        [void][System.Windows.Forms.MessageBox]::Show(
            ($erros -join "`r`n"), 'Corrija os campos',
            [System.Windows.Forms.MessageBoxButtons]::OK,
            [System.Windows.Forms.MessageBoxIcon]::Warning)
        return
    }

    $form.DialogResult = [System.Windows.Forms.DialogResult]::OK
    $form.Close()
})

$form.Controls.AddRange(@($btnOk, $btnCancel))
$form.AcceptButton = $btnOk
$form.CancelButton = $btnCancel

if ($form.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

# ---------------------------------------------------------------------------
# Valores confirmados
# ---------------------------------------------------------------------------
$vmName     = Get-NomeVM
$geracao    = Get-Geracao
$vhd        = Get-CaminhoVHD
$iso        = $txtISO.Text.Trim()
$pastaVM    = $txtPastaVM.Text
$pastaVHD   = $txtPastaVHD.Text
$dinamica   = $chkDinamica.Checked
[int]$cores = [int]$numCores.Value

$switch = if ($cboSwitch.SelectedIndex -le 0) { $null } else { [string]$cboSwitch.SelectedItem }

$memStartup = ConvertTo-BytesMemoria $numStartup.Value
$memMin     = ConvertTo-BytesMemoria $numMin.Value
$memMax     = ConvertTo-BytesMemoria $numMax.Value

$discoSoBytes = [int64]$numDiscoSo.Value * 1GB
$discoSoFixo  = ($cboTipoSo.SelectedItem -eq 'Fixo')

$planoDiscos = @()
for ($i = 1; $i -le $discos.Count; $i++) {
    $planoDiscos += [pscustomobject]@{
        Indice    = $i
        Caminho   = Get-CaminhoDisco $i
        TamanhoGB = $discos[$i - 1].TamanhoGB
        Tipo      = $discos[$i - 1].Tipo
    }
}

# ---------------------------------------------------------------------------
# Janela de progresso (discos fixos podem demorar bastante)
# ---------------------------------------------------------------------------
$passos = 6 + $planoDiscos.Count

$formProg                 = New-Object System.Windows.Forms.Form
$formProg.Text            = 'Criando a máquina virtual...'
$formProg.Size            = New-Object System.Drawing.Size(440, 135)
$formProg.StartPosition   = 'CenterScreen'
$formProg.FormBorderStyle = 'FixedDialog'
$formProg.ControlBox      = $false
$formProg.Font            = New-Object System.Drawing.Font('Segoe UI', 9)

$lblProg          = New-Object System.Windows.Forms.Label
$lblProg.Location = New-Object System.Drawing.Point(15, 18)
$lblProg.Size     = New-Object System.Drawing.Size(400, 20)

$barProg          = New-Object System.Windows.Forms.ProgressBar
$barProg.Location = New-Object System.Drawing.Point(15, 45)
$barProg.Size     = New-Object System.Drawing.Size(400, 20)
$barProg.Maximum  = $passos
$barProg.Value    = 0

$formProg.Controls.AddRange(@($lblProg, $barProg))

function Step-Progresso {
    param([string]$Texto)
    $lblProg.Text = $Texto
    if ($barProg.Value -lt $barProg.Maximum) { $barProg.Value++ }
    [System.Windows.Forms.Application]::DoEvents()
}

# ---------------------------------------------------------------------------
# Criação
# ---------------------------------------------------------------------------
$criados = New-Object System.Collections.Generic.List[string]
$vm      = $null

$formProg.Show()
$formProg.Refresh()

try {
    Step-Progresso 'Verificando as pastas de destino...'
    foreach ($p in @($pastaVHD, $pastaVM)) {
        if (-not (Test-Path -LiteralPath $p)) {
            New-Item -ItemType Directory -Path $p -Force | Out-Null
        }
    }

    Step-Progresso 'Criando o disco do sistema...'
    if ($discoSoFixo) {
        New-VHD -Path $vhd -SizeBytes $discoSoBytes -Fixed | Out-Null
    } else {
        New-VHD -Path $vhd -SizeBytes $discoSoBytes -Dynamic | Out-Null
    }
    $criados.Add($vhd)

    Step-Progresso "Criando a VM $vmName (geração $geracao)..."
    $parametros = @{
        Name               = $vmName
        MemoryStartupBytes = $memStartup
        Path               = $pastaVM
        Generation         = $geracao
        VHDPath            = $vhd
    }
    if ($switch) { $parametros['SwitchName'] = $switch }
    $vm = New-VM @parametros

    Step-Progresso 'Aplicando memória e processadores...'
    # memória antes do processador: a virtualização aninhada só é aceita
    # com memória dinâmica desligada
    if ($dinamica) {
        Set-VMMemory -VM $vm -DynamicMemoryEnabled $true `
                     -StartupBytes $memStartup -MinimumBytes $memMin -MaximumBytes $memMax
    } else {
        Set-VMMemory -VM $vm -DynamicMemoryEnabled $false -StartupBytes $memStartup
    }

    if ($chkNested.Checked) {
        Set-VMProcessor -VM $vm -Count $cores -ExposeVirtualizationExtensions $true
    } else {
        Set-VMProcessor -VM $vm -Count $cores
    }

    Step-Progresso 'Configurando firmware e mídia de instalação...'

    # firmware: só a geração 2 tem Secure Boot e TPM
    if ($geracao -eq 2) {
        switch ($cboSecureBoot.SelectedIndex) {
            0 { Set-VMFirmware -VM $vm -EnableSecureBoot On  -SecureBootTemplate 'MicrosoftWindows' }
            1 { Set-VMFirmware -VM $vm -EnableSecureBoot On  -SecureBootTemplate 'MicrosoftUEFICertificateAuthority' }
            2 { Set-VMFirmware -VM $vm -EnableSecureBoot Off }
        }
        if ($chkTPM.Checked) {
            Set-VMKeyProtector -VM $vm -NewLocalKeyProtector
            Enable-VMTPM -VM $vm
        }
    }

    # ISO: reaproveita o drive de DVD que a geração 1 já traz de fábrica
    if ($iso) {
        $dvd = Get-VMDvdDrive -VM $vm | Select-Object -First 1
        if ($dvd) {
            Set-VMDvdDrive -VMName $vmName `
                           -ControllerNumber $dvd.ControllerNumber `
                           -ControllerLocation $dvd.ControllerLocation -Path $iso
        } else {
            # -Passthru devolve o drive recem-criado, sem depender de reconsulta
            $dvd = Add-VMDvdDrive -VM $vm -Path $iso -Passthru
        }

        # bootar pela ISO: cada geração usa um cmdlet diferente
        if ($geracao -eq 2) {
            if ($dvd) { Set-VMFirmware -VM $vm -FirstBootDevice $dvd }
        } else {
            Set-VMBios -VM $vm -StartupOrder @('CD', 'IDE', 'LegacyNetworkAdapter', 'Floppy')
        }
    }

    if ($switch) {
        if ($chkMac.Checked) {
            Get-VMNetworkAdapter -VM $vm | Set-VMNetworkAdapter -MacAddressSpoofing On
        }
        if ($chkVlan.Checked) {
            Get-VMNetworkAdapter -VM $vm |
                Set-VMNetworkAdapterVlan -Access -VlanId ([int]$numVlan.Value)
        }
    }

    if ($chkProducao.Checked) { Set-VM -VM $vm -CheckpointType Production }
    if ($chkAutoChk.Checked)  { Set-VM -VM $vm -AutomaticCheckpointsEnabled $false }
    if ($chkGuest.Checked) {
        # ID fixo do "Serviço de Convidado" - o nome muda conforme o idioma do host
        Get-VMIntegrationService -VMName $vmName |
            Where-Object { $_.Id -like '*6C09BB55*' } |
            Enable-VMIntegrationService
    }

    if ($planoDiscos.Count -gt 0 -and -not (Get-VMScsiController -VM $vm)) {
        Add-VMScsiController -VM $vm
    }

    foreach ($d in $planoDiscos) {
        Step-Progresso ("Disco {0} de {1} ({2} GB, {3})..." -f `
            $d.Indice, $planoDiscos.Count, $d.TamanhoGB, $d.Tipo.ToLower())

        $bytes = [int64]$d.TamanhoGB * 1GB
        if ($d.Tipo -eq 'Fixo') {
            New-VHD -Path $d.Caminho -SizeBytes $bytes -Fixed | Out-Null
        } else {
            New-VHD -Path $d.Caminho -SizeBytes $bytes -Dynamic | Out-Null
        }
        $criados.Add($d.Caminho)

        Add-VMHardDiskDrive -VM $vm -Path $d.Caminho -ControllerType SCSI -ControllerNumber 0
    }

    Step-Progresso 'Finalizando...'
    if ($chkIniciar.Checked) { Start-VM -VM $vm }

    $descMemoria = if ($dinamica) {
        "dinâmica - inicial $($numStartup.Value) GB, mín. $($numMin.Value) GB, máx. $($numMax.Value) GB"
    } else {
        "estática - $($numStartup.Value) GB"
    }

    $descFirmware = if ($geracao -eq 2) {
        $sb = [string]$cboSecureBoot.SelectedItem
        if ($chkTPM.Checked) { "UEFI - Secure Boot $sb, TPM habilitado" } else { "UEFI - Secure Boot $sb" }
    } else {
        'BIOS (geração 1)'
    }

    $descDiscos = if ($planoDiscos.Count -eq 0) {
        'nenhum'
    } else {
        $somaGB = ($planoDiscos | Measure-Object -Property TamanhoGB -Sum).Sum
        "$($planoDiscos.Count) disco(s), $somaGB GB no total"
    }

    $resumo = @(
        "VM:            $vmName"
        "Geração:       $geracao"
        "Firmware:      $descFirmware"
        "Memória:       $descMemoria"
        "Processadores: $cores"
        "Rede:          $(if ($switch) { $switch } else { 'sem conexão' })"
        "ISO:           $(if ($iso) { Split-Path -Leaf $iso } else { 'nenhuma' })"
        "Disco:         $($numDiscoSo.Value) GB ($($cboTipoSo.SelectedItem.ToString().ToLower()))"
        "Discos extras: $descDiscos"
    ) -join "`r`n"

    $formProg.Close()

    [void][System.Windows.Forms.MessageBox]::Show(
        "A criação da máquina virtual foi finalizada!`r`n`r`n$resumo",
        'Processo concluído',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information)
}
catch {
    $mensagem = $_.Exception.Message
    $formProg.Close()

    # se a VM nao chegou a ser criada, nada esta anexado: limpa os VHDX orfaos.
    # se ela existe, os discos ja criados ficam anexados a ela e sao preservados.
    $vmExiste = [bool](Get-VM -Name $vmName -ErrorAction SilentlyContinue)
    if (-not $vmExiste) {
        foreach ($arquivo in $criados) {
            Remove-Item -LiteralPath $arquivo -Force -ErrorAction SilentlyContinue
        }
        $extra = 'Nenhuma alteração foi mantida no host.'
    } else {
        $extra = "A VM '$vmName' foi criada e os discos já anexados foram preservados. " +
                 'Revise no Gerenciador do Hyper-V antes de rodar o script de novo.'
    }

    [void][System.Windows.Forms.MessageBox]::Show(
        "Falha ao criar a VM:`r`n`r`n$mensagem`r`n`r`n$extra",
        'Erro',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error)
    throw
}
finally {
    $formProg.Dispose()
    $form.Dispose()
}
