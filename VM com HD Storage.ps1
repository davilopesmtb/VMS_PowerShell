﻿#prefixo geral$prefixo = "STORAGE01"#Defina na variável a baixo o nome da VM a ser criada.$vmname = "CLF - LAB1 - $prefixo"#Defina na variável a baixo o nome/prefixo do HD do Storage$nome_hd_storage = "$prefixo"#Defina na variável a baixo o nome e caminho do HD principal da VM$vhd00 = "D:\Hyper-V\VHD\CLF-LAB1-$prefixo.vhdx"#Defina na variável a baixo o nome e caminho do HD Diferencial/Modelo$DiscoPai = "D:\Hyper-V\VHD_PAI\Win2016Pai.vhdx"#Comando para criar HDs virtuais do laboratório 1 de cluster tendo disco diferencial como modelo.New-VHD -Path $vhd00 -ParentPath $DiscoPai -Differencing#Comando para criar máquinas virtuais do laboratório 1 de clusterNew-VM -Name $vmname -MemoryStartupBytes 4GB -Path "D:\Hyper-V\MAQUINAS" -Generation 2 -VHDPath $vhd00#Comando para criar HDs STORAGE virtuais do laboratório 1 de cluster$vhd01 = "D:\Hyper-V\STORAGE\CLF-LAB1-$nome_hd_storage-DISCO01-st.vhdx"$vhd02 = "D:\Hyper-V\STORAGE\CLF-LAB1-$nome_hd_storage-DISCO02-st.vhdx"$vhd03 = "D:\Hyper-V\STORAGE\CLF-LAB1-$nome_hd_storage-DISCO03-st.vhdx"$vhd04 = "D:\Hyper-V\STORAGE\CLF-LAB1-$nome_hd_storage-DISCO04-st.vhdx" $vhd05 = "D:\Hyper-V\STORAGE\CLF-LAB1-$nome_hd_storage-DISCO05-st.vhdx"$vhd06 = "D:\Hyper-V\STORAGE\CLF-LAB1-$nome_hd_storage-DISCO06-st.vhdx"New-VHD -Path $vhd01 -SizeBytes 127gb -Dynamic New-VHD -Path $vhd02 -SizeBytes 127gb -Dynamic New-VHD -Path $vhd03 -SizeBytes 127gb -Dynamic New-VHD -Path $vhd04 -SizeBytes 127gb -Dynamic New-VHD -Path $vhd05 -SizeBytes 127gb -Dynamic New-VHD -Path $vhd06 -SizeBytes 127gb -Dynamic #Comando para adicionar o HDs de STORSGE a VM criadaGet-VM $vmname | Add-VMHardDiskDrive -Path $vhd01 -ControllerType SCSI -ControllerNumber 0Get-VM $vmname | Add-VMHardDiskDrive -Path $vhd02 -ControllerType SCSI -ControllerNumber 0Get-VM $vmname | Add-VMHardDiskDrive -Path $vhd03 -ControllerType SCSI -ControllerNumber 0Get-VM $vmname | Add-VMHardDiskDrive -Path $vhd04 -ControllerType SCSI -ControllerNumber 0Get-VM $vmname | Add-VMHardDiskDrive -Path $vhd05 -ControllerType SCSI -ControllerNumber 0Get-VM $vmname | Add-VMHardDiskDrive -Path $vhd06 -ControllerType SCSI -ControllerNumber 0