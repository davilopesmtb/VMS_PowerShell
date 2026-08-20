#Requires -RunAsAdministrator
#Requires -Modules Hyper-V
<#
.SYNOPSIS
    Cria uma VM de laboratório no Hyper-V com discos de storage adicionais.

.DESCRIPTION
    Versão em formulário único do "VM com HD Storage.ps1" / "VM com 2 HDs.ps1".
    Em vez das variáveis fixas $vhd01..$vhd06, os discos adicionais são montados
    em uma lista dentro do formulário: informe quantidade, tamanho e tipo e
    clique em Adicionar quantas vezes quiser. Dá para misturar tamanhos
    (ex.: 2 discos de 127 GB + 1 de 500 GB) sem editar o script.

    Nada é criado no host enquanto o formulário não for confirmado.

    O disco do sistema é criado como disco de diferenciação a partir de um
    VHDX modelo (disco pai), que deve estar preparado (sysprep) e somente
    leitura. Os discos de storage são criados vazios e anexados na
    controladora SCSI 0.
#>

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ---------------------------------------------------------------------------
# Padrões: ajuste aqui os valores iniciais do formulário
# ---------------------------------------------------------------------------
$padrao = @{
    Prefixo      = 'SRV2025'
    StartupGB    = 4
    MinimaGB     = 1
    MaximaGB     = 8
    Dinamica     = $false
    Cores        = 2
    DiscoPai     = 'C:\HYPERV\VHD_Modelo\WS-2025-MODEL.vhdx'
    PastaVHD     = 'C:\HYPERV\VHD'
    PastaVM      = 'C:\HYPERV\MAQUINAS'
    PastaStorage = 'C:\HYPERV\STORAGE'
    DiscoGB      = 127
    DiscoTipo    = 'Dinâmico'
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

# lista de discos adicionais planejados (preenchida pelo formulário)
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
$form.Text            = 'Criação de VM com discos de storage - Hyper-V'
$form.Size            = New-Object System.Drawing.Size(600, 645)
$form.StartPosition   = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox     = $false
$form.MinimizeBox     = $false
$form.Font            = New-Object System.Drawing.Font('Segoe UI', 9)

$tabs          = New-Object System.Windows.Forms.TabControl
$tabs.Location = New-Object System.Drawing.Point(12, 12)
$tabs.Size     = New-Object System.Drawing.Size(560, 490)

$tabVM      = New-Object System.Windows.Forms.TabPage
$tabVM.Text = 'Máquina virtual'
$tabDisco      = New-Object System.Windows.Forms.TabPage
$tabDisco.Text = 'Discos adicionais'
$tabs.Controls.AddRange(@($tabVM, $tabDisco))
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
$tabVM.Controls.AddRange(@((New-Rotulo 'Prefixo:' 12 $y), $txtPrefixo))

$y = 75
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

$y = 105
$cboSwitch               = New-Object System.Windows.Forms.ComboBox
$cboSwitch.Location      = New-Object System.Drawing.Point(155, $y)
$cboSwitch.Size          = New-Object System.Drawing.Size(385, 23)
$cboSwitch.DropDownStyle = 'DropDownList'
if ($switches.Count -gt 0) {
    [void]$cboSwitch.Items.AddRange($switches)
    $cboSwitch.SelectedIndex = 0
}
$tabVM.Controls.AddRange(@((New-Rotulo 'Switch virtual:' 12 $y), $cboSwitch))

$y = 135
$txtDiscoPai          = New-Object System.Windows.Forms.TextBox
$txtDiscoPai.Location = New-Object System.Drawing.Point(155, $y)
$txtDiscoPai.Size     = New-Object System.Drawing.Size(340, 23)
$txtDiscoPai.Text     = $padrao.DiscoPai
$btnDiscoPai          = New-BotaoProcurar 500 $y
$btnDiscoPai.Add_Click({
    $dlg = New-Object System.Windows.Forms.OpenFileDialog
    $dlg.Filter = 'Discos virtuais (*.vhdx;*.vhd)|*.vhdx;*.vhd'
    if (Test-Path -LiteralPath $txtDiscoPai.Text) {
        $dlg.InitialDirectory = Split-Path -Parent $txtDiscoPai.Text
    }
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtDiscoPai.Text = $dlg.FileName
    }
})
$tabVM.Controls.AddRange(@((New-Rotulo 'Disco pai (modelo):' 12 $y), $txtDiscoPai, $btnDiscoPai))

$y = 165
$txtPastaVHD          = New-Object System.Windows.Forms.TextBox
$txtPastaVHD.Location = New-Object System.Drawing.Point(155, $y)
$txtPastaVHD.Size     = New-Object System.Drawing.Size(340, 23)
$txtPastaVHD.Text     = $padrao.PastaVHD
$btnPastaVHD          = New-BotaoProcurar 500 $y
$btnPastaVHD.Add_Click({ Select-Pasta $txtPastaVHD })
$tabVM.Controls.AddRange(@((New-Rotulo 'Pasta dos VHDs:' 12 $y), $txtPastaVHD, $btnPastaVHD))

$y = 195
$txtPastaVM          = New-Object System.Windows.Forms.TextBox
$txtPastaVM.Location = New-Object System.Drawing.Point(155, $y)
$txtPastaVM.Size     = New-Object System.Drawing.Size(340, 23)
$txtPastaVM.Text     = $padrao.PastaVM
$btnPastaVM          = New-BotaoProcurar 500 $y
$btnPastaVM.Add_Click({ Select-Pasta $txtPastaVM })
$tabVM.Controls.AddRange(@((New-Rotulo 'Pasta das VMs:' 12 $y), $txtPastaVM, $btnPastaVM))

# --- Memória ---------------------------------------------------------------
$grpMem          = New-Object System.Windows.Forms.GroupBox
$grpMem.Text     = 'Memória'
$grpMem.Location = New-Object System.Drawing.Point(10, 228)
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

# --- Opções ----------------------------------------------------------------
$grp          = New-Object System.Windows.Forms.GroupBox
$grp.Text     = 'Opções'
$grp.Location = New-Object System.Drawing.Point(10, 343)
$grp.Size     = New-Object System.Drawing.Size(528, 105)

$chkNested          = New-Object System.Windows.Forms.CheckBox
$chkNested.Text     = 'Virtualização aninhada'
$chkNested.Location = New-Object System.Drawing.Point(15, 25)
$chkNested.Size     = New-Object System.Drawing.Size(240, 22)
$chkNested.Checked  = $true

$chkMac          = New-Object System.Windows.Forms.CheckBox
$chkMac.Text     = 'MAC address spoofing (rede aninhada)'
$chkMac.Location = New-Object System.Drawing.Point(15, 50)
$chkMac.Size     = New-Object System.Drawing.Size(270, 22)
$chkMac.Checked  = $true

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
$tabVM.Controls.Add($grp)

# ===========================================================================
# Aba 2 - Discos adicionais
# ===========================================================================
$y = 15
$txtPastaStorage          = New-Object System.Windows.Forms.TextBox
$txtPastaStorage.Location = New-Object System.Drawing.Point(155, $y)
$txtPastaStorage.Size     = New-Object System.Drawing.Size(340, 23)
$txtPastaStorage.Text     = $padrao.PastaStorage
$btnPastaStorage          = New-BotaoProcurar 500 $y
$btnPastaStorage.Add_Click({ Select-Pasta $txtPastaStorage })
$tabDisco.Controls.AddRange(@((New-Rotulo 'Pasta dos discos:' 12 $y), $txtPastaStorage, $btnPastaStorage))

$grpAdd          = New-Object System.Windows.Forms.GroupBox
$grpAdd.Text     = 'Adicionar discos'
$grpAdd.Location = New-Object System.Drawing.Point(10, 45)
$grpAdd.Size     = New-Object System.Drawing.Size(528, 100)

$numQtd          = New-Object System.Windows.Forms.NumericUpDown
$numQtd.Location = New-Object System.Drawing.Point(95, 25)
$numQtd.Size     = New-Object System.Drawing.Size(60, 23)
$numQtd.Minimum  = 1
$numQtd.Maximum  = $maxDiscos
$numQtd.Value    = 1

$numTamanho          = New-Object System.Windows.Forms.NumericUpDown
$numTamanho.Location = New-Object System.Drawing.Point(270, 25)
$numTamanho.Size     = New-Object System.Drawing.Size(70, 23)
$numTamanho.Minimum  = 1
$numTamanho.Maximum  = 65536
$numTamanho.Value    = $padrao.DiscoGB

$cboTipo               = New-Object System.Windows.Forms.ComboBox
$cboTipo.Location      = New-Object System.Drawing.Point(395, 25)
$cboTipo.Size          = New-Object System.Drawing.Size(110, 23)
$cboTipo.DropDownStyle = 'DropDownList'
[void]$cboTipo.Items.AddRange(@('Dinâmico', 'Fixo'))
$cboTipo.SelectedItem  = $padrao.DiscoTipo

$lblDica           = New-Object System.Windows.Forms.Label
$lblDica.Text      = 'Clique em Adicionar quantas vezes precisar - dá para misturar tamanhos.'
$lblDica.Location  = New-Object System.Drawing.Point(15, 65)
$lblDica.Size      = New-Object System.Drawing.Size(370, 20)
$lblDica.ForeColor = [System.Drawing.Color]::DimGray

$btnAddDisco          = New-Object System.Windows.Forms.Button
$btnAddDisco.Text     = 'Adicionar'
$btnAddDisco.Location = New-Object System.Drawing.Point(395, 60)
$btnAddDisco.Size     = New-Object System.Drawing.Size(110, 27)

$grpAdd.Controls.AddRange(@(
    (New-Rotulo 'Quantidade:' 15 25 75),   $numQtd,
    (New-Rotulo 'Tamanho (GB):' 175 25 90), $numTamanho,
    (New-Rotulo 'Tipo:' 355 25 40),         $cboTipo,
    $lblDica, $btnAddDisco))
$tabDisco.Controls.Add($grpAdd)

$lstDiscos               = New-Object System.Windows.Forms.ListView
$lstDiscos.Location      = New-Object System.Drawing.Point(10, 155)
$lstDiscos.Size          = New-Object System.Drawing.Size(528, 230)
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
$btnRemover.Location = New-Object System.Drawing.Point(10, 392)
$btnRemover.Size     = New-Object System.Drawing.Size(150, 27)

$btnLimpar          = New-Object System.Windows.Forms.Button
$btnLimpar.Text     = 'Limpar lista'
$btnLimpar.Location = New-Object System.Drawing.Point(168, 392)
$btnLimpar.Size     = New-Object System.Drawing.Size(110, 27)

$lblTotais          = New-Object System.Windows.Forms.Label
$lblTotais.Location = New-Object System.Drawing.Point(10, 428)
$lblTotais.Size     = New-Object System.Drawing.Size(528, 32)

$tabDisco.Controls.AddRange(@($btnRemover, $btnLimpar, $lblTotais))

# ===========================================================================
# Nomes derivados e atualizacao da lista
# ===========================================================================
function Get-NomeVM { "$($txtPrefixo.Text.Trim()) $($txtNome.Text.Trim())".Trim() }

function Get-CaminhoVHD {
    Join-Path $txtPastaVHD.Text ('{0}-{1}-so.vhdx' -f $txtPrefixo.Text.Trim(), $txtNome.Text.Trim())
}

function Get-CaminhoDisco {
    param([int]$Indice)
    Join-Path $txtPastaStorage.Text `
        ('{0}-{1}-DISCO{2:D2}-st.vhdx' -f $txtPrefixo.Text.Trim(), $txtNome.Text.Trim(), $Indice)
}

$atualizarTotais = {
    if ($discos.Count -eq 0) {
        $lblTotais.Text      = 'Nenhum disco adicional. A VM sera criada apenas com o disco do sistema.'
        $lblTotais.ForeColor = [System.Drawing.Color]::DimGray
        return
    }

    $total = ($discos | Measure-Object -Property TamanhoGB -Sum).Sum
    $fixos = ($discos | Where-Object { $_.Tipo -eq 'Fixo' } | Measure-Object -Property TamanhoGB -Sum).Sum
    if (-not $fixos) { $fixos = 0 }

    $texto = "$($discos.Count) disco(s) - $total GB no total"
    if ($fixos -gt 0) { $texto += " ($fixos GB alocados de imediato, em discos fixos)" }

    $livre = Get-EspacoLivreGB $txtPastaStorage.Text
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
$lblPreview.Location  = New-Object System.Drawing.Point(12, 512)
$lblPreview.Size      = New-Object System.Drawing.Size(560, 40)
$lblPreview.ForeColor = [System.Drawing.Color]::DimGray
$form.Controls.Add($lblPreview)

$atualizarPreview = {
    if ([string]::IsNullOrWhiteSpace($txtNome.Text)) {
        $lblPreview.Text = ''
    } else {
        $lblPreview.Text = "VM:   $(Get-NomeVM)`r`nVHDX: $(Get-CaminhoVHD)"
    }
    & $renderDiscos   # os nomes dos discos derivam do nome da VM
}
$txtNome.Add_TextChanged($atualizarPreview)
$txtPrefixo.Add_TextChanged($atualizarPreview)
$txtPastaVHD.Add_TextChanged($atualizarPreview)
$txtPastaStorage.Add_TextChanged($renderDiscos)
& $renderDiscos

# --- Botões ----------------------------------------------------------------
$btnOk          = New-Object System.Windows.Forms.Button
$btnOk.Text     = 'Criar VM'
$btnOk.Location = New-Object System.Drawing.Point(372, 560)
$btnOk.Size     = New-Object System.Drawing.Size(95, 30)

$btnCancel              = New-Object System.Windows.Forms.Button
$btnCancel.Text         = 'Cancelar'
$btnCancel.Location     = New-Object System.Drawing.Point(477, 560)
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

    if (-not $cboSwitch.SelectedItem) { $erros.Add('Nenhum switch virtual selecionado.') }

    if ($chkDinamica.Checked) {
        if ($numMin.Value -gt $numStartup.Value) {
            $erros.Add('A memória mínima não pode ser maior que a inicial.')
        }
        if ($numMax.Value -lt $numStartup.Value) {
            $erros.Add('A memória máxima não pode ser menor que a inicial.')
        }
    }

    if (-not (Test-Path -LiteralPath $txtDiscoPai.Text)) {
        $erros.Add("Disco pai não encontrado: $($txtDiscoPai.Text)")
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
    $livre = Get-EspacoLivreGB $txtPastaStorage.Text
    if ($fixos -and $livre -and ($fixos -gt $livre)) {
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

if ($switches.Count -eq 0) {
    [void][System.Windows.Forms.MessageBox]::Show(
        'Nenhum switch virtual encontrado neste host. Crie um switch no Gerenciador do Hyper-V antes de continuar.',
        'Sem switch virtual',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error)
    return
}

if ($form.ShowDialog() -ne [System.Windows.Forms.DialogResult]::OK) { return }

# ---------------------------------------------------------------------------
# Valores confirmados
# ---------------------------------------------------------------------------
$vmName       = Get-NomeVM
$vhd          = Get-CaminhoVHD
$discoPai     = $txtDiscoPai.Text
$pastaVM      = $txtPastaVM.Text
$pastaVHD     = $txtPastaVHD.Text
$pastaStorage = $txtPastaStorage.Text
$switch       = [string]$cboSwitch.SelectedItem
$dinamica     = $chkDinamica.Checked
[int]$cores   = [int]$numCores.Value

$memStartup = ConvertTo-BytesMemoria $numStartup.Value
$memMin     = ConvertTo-BytesMemoria $numMin.Value
$memMax     = ConvertTo-BytesMemoria $numMax.Value

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
$passos = 5 + $planoDiscos.Count

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
    foreach ($p in @($pastaVHD, $pastaVM, $pastaStorage)) {
        if (-not (Test-Path -LiteralPath $p)) {
            New-Item -ItemType Directory -Path $p -Force | Out-Null
        }
    }

    Step-Progresso 'Criando o disco do sistema...'
    New-VHD -Path $vhd -ParentPath $discoPai -Differencing | Out-Null
    $criados.Add($vhd)

    Step-Progresso "Criando a VM $vmName..."
    $vm = New-VM -Name $vmName -MemoryStartupBytes $memStartup -Path $pastaVM `
                 -Generation 2 -VHDPath $vhd -SwitchName $switch

    Step-Progresso 'Aplicando memória, processadores e opções...'
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

    if ($chkMac.Checked) {
        Get-VMNetworkAdapter -VM $vm | Set-VMNetworkAdapter -MacAddressSpoofing On
    }
    if ($chkProducao.Checked) { Set-VM -VM $vm -CheckpointType Production }
    if ($chkAutoChk.Checked)  { Set-VM -VM $vm -AutomaticCheckpointsEnabled $false }
    if ($chkGuest.Checked) {
        # ID fixo do "Serviço de Convidado" - o nome muda conforme o idioma do host
        Get-VMIntegrationService -VMName $vmName |
            Where-Object { $_.Id -like '*6C09BB55*' } |
            Enable-VMIntegrationService
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

    $descDiscos = if ($planoDiscos.Count -eq 0) {
        'nenhum'
    } else {
        $somaGB = ($planoDiscos | Measure-Object -Property TamanhoGB -Sum).Sum
        "$($planoDiscos.Count) disco(s), $somaGB GB no total"
    }

    $resumo = @(
        "VM:            $vmName"
        "Memória:       $descMemoria"
        "Processadores: $cores"
        "Switch:        $switch"
        "VHDX:          $vhd"
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
