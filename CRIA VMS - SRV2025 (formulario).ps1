#Requires -RunAsAdministrator
#Requires -Modules Hyper-V
<#
.SYNOPSIS
    Cria uma VM de laboratório no Hyper-V a partir de um único formulário.

.DESCRIPTION
    Substitui a sequência de InputBox por uma única janela (Windows Forms) com
    todos os campos: nome, memória (inicial, mínima e máxima), processadores,
    switch, caminhos e opções.
    Nada é criado no host enquanto o formulário não for confirmado.

    O disco do sistema é criado como disco de diferenciação a partir de um
    VHDX modelo (disco pai), que deve estar preparado (sysprep) e somente
    leitura.
#>

$ErrorActionPreference = 'Stop'

Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
[System.Windows.Forms.Application]::EnableVisualStyles()

# ---------------------------------------------------------------------------
# Padrões: ajuste aqui os valores iniciais do formulário
# ---------------------------------------------------------------------------
$padrao = @{
    Prefixo   = 'SRV2025'
    StartupGB = 4
    MinimaGB  = 1
    MaximaGB  = 8
    Dinamica  = $false
    Cores     = 2
    DiscoPai  = 'C:\HYPERV\VHD_Modelo\WS-2025-MODEL.vhdx'
    PastaVHD  = 'C:\HYPERV\VHD'
    PastaVM   = 'C:\HYPERV\MAQUINAS'
}

# ---------------------------------------------------------------------------
# Dados do host (usados para limitar os campos numéricos)
# ---------------------------------------------------------------------------
$cs       = Get-CimInstance Win32_ComputerSystem
$maxCores = [int]$cs.NumberOfLogicalProcessors
$maxRamGB = [int][math]::Floor($cs.TotalPhysicalMemory / 1GB)
$switches = @(Get-VMSwitch | Select-Object -ExpandProperty Name | Sort-Object)

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
    param([int]$Y)
    $b = New-Object System.Windows.Forms.Button
    $b.Text     = '...'
    $b.Location = New-Object System.Drawing.Point(500, ($Y - 1))
    $b.Size     = New-Object System.Drawing.Size(40, 25)
    return $b
}

# Hyper-V trabalha com múltiplos de 2 MB
function ConvertTo-BytesMemoria {
    param([decimal]$GB)
    $bytes = [int64]($GB * 1GB)
    return [int64]([math]::Round($bytes / 2MB) * 2MB)
}

# ---------------------------------------------------------------------------
# Formulário
# ---------------------------------------------------------------------------
$form                 = New-Object System.Windows.Forms.Form
$form.Text            = 'Criação de VM - Hyper-V'
$form.Size            = New-Object System.Drawing.Size(570, 590)
$form.StartPosition   = 'CenterScreen'
$form.FormBorderStyle = 'FixedDialog'
$form.MaximizeBox     = $false
$form.MinimizeBox     = $false
$form.Font            = New-Object System.Drawing.Font('Segoe UI', 9)

# --- Nome ------------------------------------------------------------------
$y = 15
$txtNome          = New-Object System.Windows.Forms.TextBox
$txtNome.Location = New-Object System.Drawing.Point(155, $y)
$txtNome.Size     = New-Object System.Drawing.Size(385, 23)
$form.Controls.AddRange(@((New-Rotulo 'Nome da VM:' 12 $y), $txtNome))

# --- Prefixo ---------------------------------------------------------------
$y = 45
$txtPrefixo          = New-Object System.Windows.Forms.TextBox
$txtPrefixo.Location = New-Object System.Drawing.Point(155, $y)
$txtPrefixo.Size     = New-Object System.Drawing.Size(150, 23)
$txtPrefixo.Text     = $padrao.Prefixo
$form.Controls.AddRange(@((New-Rotulo 'Prefixo:' 12 $y), $txtPrefixo))

# --- Processadores ---------------------------------------------------------
$y = 75
$numCores          = New-Object System.Windows.Forms.NumericUpDown
$numCores.Location = New-Object System.Drawing.Point(155, $y)
$numCores.Size     = New-Object System.Drawing.Size(80, 23)
$numCores.Minimum  = 1
$numCores.Maximum  = $maxCores
$numCores.Value    = [math]::Min($padrao.Cores, $maxCores)
$form.Controls.AddRange(@(
    (New-Rotulo 'Processadores:' 12 $y),
    $numCores,
    (New-Rotulo "(host: $maxCores lógicos)" 240 $y 200)))

# --- Switch ----------------------------------------------------------------
$y = 105
$cboSwitch               = New-Object System.Windows.Forms.ComboBox
$cboSwitch.Location      = New-Object System.Drawing.Point(155, $y)
$cboSwitch.Size          = New-Object System.Drawing.Size(385, 23)
$cboSwitch.DropDownStyle = 'DropDownList'
if ($switches.Count -gt 0) {
    [void]$cboSwitch.Items.AddRange($switches)
    $cboSwitch.SelectedIndex = 0
}
$form.Controls.AddRange(@((New-Rotulo 'Switch virtual:' 12 $y), $cboSwitch))

# --- Disco pai -------------------------------------------------------------
$y = 135
$txtDiscoPai          = New-Object System.Windows.Forms.TextBox
$txtDiscoPai.Location = New-Object System.Drawing.Point(155, $y)
$txtDiscoPai.Size     = New-Object System.Drawing.Size(340, 23)
$txtDiscoPai.Text     = $padrao.DiscoPai
$btnDiscoPai          = New-BotaoProcurar $y
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
$form.Controls.AddRange(@((New-Rotulo 'Disco pai (modelo):' 12 $y), $txtDiscoPai, $btnDiscoPai))

# --- Pasta dos VHDs --------------------------------------------------------
$y = 165
$txtPastaVHD          = New-Object System.Windows.Forms.TextBox
$txtPastaVHD.Location = New-Object System.Drawing.Point(155, $y)
$txtPastaVHD.Size     = New-Object System.Drawing.Size(340, 23)
$txtPastaVHD.Text     = $padrao.PastaVHD
$btnPastaVHD          = New-BotaoProcurar $y
$btnPastaVHD.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.SelectedPath = $txtPastaVHD.Text
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtPastaVHD.Text = $dlg.SelectedPath
    }
})
$form.Controls.AddRange(@((New-Rotulo 'Pasta dos VHDs:' 12 $y), $txtPastaVHD, $btnPastaVHD))

# --- Pasta das VMs ---------------------------------------------------------
$y = 195
$txtPastaVM          = New-Object System.Windows.Forms.TextBox
$txtPastaVM.Location = New-Object System.Drawing.Point(155, $y)
$txtPastaVM.Size     = New-Object System.Drawing.Size(340, 23)
$txtPastaVM.Text     = $padrao.PastaVM
$btnPastaVM          = New-BotaoProcurar $y
$btnPastaVM.Add_Click({
    $dlg = New-Object System.Windows.Forms.FolderBrowserDialog
    $dlg.SelectedPath = $txtPastaVM.Text
    if ($dlg.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
        $txtPastaVM.Text = $dlg.SelectedPath
    }
})
$form.Controls.AddRange(@((New-Rotulo 'Pasta das VMs:' 12 $y), $txtPastaVM, $btnPastaVM))

# --- Memória ---------------------------------------------------------------
$grpMem          = New-Object System.Windows.Forms.GroupBox
$grpMem.Text     = 'Memória'
$grpMem.Location = New-Object System.Drawing.Point(12, 228)
$grpMem.Size     = New-Object System.Drawing.Size(528, 105)

$chkDinamica          = New-Object System.Windows.Forms.CheckBox
$chkDinamica.Text     = 'Memória dinâmica'
$chkDinamica.Location = New-Object System.Drawing.Point(15, 22)
$chkDinamica.Size     = New-Object System.Drawing.Size(200, 22)
$chkDinamica.Checked  = $padrao.Dinamica

$numStartup = New-CampoGB 90  50 $padrao.StartupGB $maxRamGB
$numMin     = New-CampoGB 250 50 $padrao.MinimaGB  $maxRamGB
$numMax     = New-CampoGB 410 50 $padrao.MaximaGB  $maxRamGB

$lblMemInfo          = New-Object System.Windows.Forms.Label
$lblMemInfo.Location = New-Object System.Drawing.Point(15, 78)
$lblMemInfo.Size     = New-Object System.Drawing.Size(500, 20)
$lblMemInfo.ForeColor = [System.Drawing.Color]::DimGray

$grpMem.Controls.AddRange(@(
    $chkDinamica,
    (New-Rotulo 'Inicial (GB):' 15  50 75), $numStartup,
    (New-Rotulo 'Mínima (GB):' 175 50 75), $numMin,
    (New-Rotulo 'Máxima (GB):' 335 50 75), $numMax,
    $lblMemInfo))
$form.Controls.Add($grpMem)

# --- Opções ----------------------------------------------------------------
$grp          = New-Object System.Windows.Forms.GroupBox
$grp.Text     = 'Opções'
$grp.Location = New-Object System.Drawing.Point(12, 343)
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
$form.Controls.Add($grp)

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
$lblPreview.Location  = New-Object System.Drawing.Point(12, 458)
$lblPreview.Size      = New-Object System.Drawing.Size(528, 40)
$lblPreview.ForeColor = [System.Drawing.Color]::DimGray

function Get-NomeVM  { "$($txtPrefixo.Text.Trim()) $($txtNome.Text.Trim())".Trim() }
function Get-CaminhoVHD {
    Join-Path $txtPastaVHD.Text ("{0}-{1}-so.vhdx" -f $txtPrefixo.Text.Trim(), $txtNome.Text.Trim())
}

$atualizarPreview = {
    if ([string]::IsNullOrWhiteSpace($txtNome.Text)) {
        $lblPreview.Text = ''
    } else {
        $lblPreview.Text = "VM:   $(Get-NomeVM)`r`nVHDX: $(Get-CaminhoVHD)"
    }
}
$txtNome.Add_TextChanged($atualizarPreview)
$txtPrefixo.Add_TextChanged($atualizarPreview)
$txtPastaVHD.Add_TextChanged($atualizarPreview)
$form.Controls.Add($lblPreview)

# --- Botões ----------------------------------------------------------------
$btnOk          = New-Object System.Windows.Forms.Button
$btnOk.Text     = 'Criar VM'
$btnOk.Location = New-Object System.Drawing.Point(340, 505)
$btnOk.Size     = New-Object System.Drawing.Size(95, 30)

$btnCancel              = New-Object System.Windows.Forms.Button
$btnCancel.Text         = 'Cancelar'
$btnCancel.Location     = New-Object System.Drawing.Point(445, 505)
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
$vmName   = Get-NomeVM
$vhd      = Get-CaminhoVHD
$discoPai = $txtDiscoPai.Text
$pastaVM  = $txtPastaVM.Text
$pastaVHD = $txtPastaVHD.Text
$switch   = [string]$cboSwitch.SelectedItem
$dinamica = $chkDinamica.Checked
[int]$cores = [int]$numCores.Value

$memStartup = ConvertTo-BytesMemoria $numStartup.Value
$memMin     = ConvertTo-BytesMemoria $numMin.Value
$memMax     = ConvertTo-BytesMemoria $numMax.Value

# ---------------------------------------------------------------------------
# Criação
# ---------------------------------------------------------------------------
$vhdCriado = $false
try {
    foreach ($p in @($pastaVHD, $pastaVM)) {
        if (-not (Test-Path -LiteralPath $p)) {
            New-Item -ItemType Directory -Path $p -Force | Out-Null
        }
    }

    New-VHD -Path $vhd -ParentPath $discoPai -Differencing | Out-Null
    $vhdCriado = $true

    $vm = New-VM -Name $vmName -MemoryStartupBytes $memStartup -Path $pastaVM `
                 -Generation 2 -VHDPath $vhd -SwitchName $switch

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
    if ($chkIniciar.Checked) { Start-VM -VM $vm }

    $descMemoria = if ($dinamica) {
        "dinâmica - inicial $($numStartup.Value) GB, mín. $($numMin.Value) GB, máx. $($numMax.Value) GB"
    } else {
        "estática - $($numStartup.Value) GB"
    }

    $resumo = @(
        "VM:            $vmName"
        "Memória:       $descMemoria"
        "Processadores: $cores"
        "Switch:        $switch"
        "VHDX:          $vhd"
    ) -join "`r`n"

    [void][System.Windows.Forms.MessageBox]::Show(
        "A criação da máquina virtual foi finalizada!`r`n`r`n$resumo",
        'Processo concluído',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Information)
}
catch {
    $mensagem = $_.Exception.Message

    # desfaz o VHDX órfão se a VM não chegou a ser criada
    if ($vhdCriado -and -not (Get-VM -Name $vmName -ErrorAction SilentlyContinue)) {
        Remove-Item -LiteralPath $vhd -Force -ErrorAction SilentlyContinue
    }

    [void][System.Windows.Forms.MessageBox]::Show(
        "Falha ao criar a VM:`r`n`r`n$mensagem",
        'Erro',
        [System.Windows.Forms.MessageBoxButtons]::OK,
        [System.Windows.Forms.MessageBoxIcon]::Error)
    throw
}
