﻿#Comando para criar máquinas virtuais do laboratório 1 de cluster#Defina na variável a baixo o nome da VM$nome = Read-Host$vhd = "D:\Hyper-V\VHD\WINDOWS10-$nome-so.vhdx"$DiscoPai = "D:\Hyper-V\VHD_PAI\WINDOWS10-PAI.vhdx"New-VHD -Path $vhd -ParentPath $DiscoPai -DifferencingNew-VM -Name "WINDOWS10 $nome" -MemoryStartupBytes 1GB -Path "D:\Hyper-V\MAQUINAS" -Generation 2 -VHDPath $vhd