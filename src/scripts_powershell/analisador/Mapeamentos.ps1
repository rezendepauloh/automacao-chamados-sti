[Diagnostics.CodeAnalysis.SuppressMessageAttribute('PSUseDeclaredVarsMoreThanAssignments', '', Justification = 'Variables are exported to the parent script via dot-sourcing.')]
param()

# -------------------------------------------------------------------------
# TABELAS DE MAPEAMENTO DE HARDWARE (Mapeamentos.ps1)
# -------------------------------------------------------------------------
# Este arquivo contém os mapeamentos de computadores (nomes amigáveis),
# monitores (via código EDID), fabricantes de memória RAM (códigos JEDEC)
# e tecnologias de saída de vídeo (VideoOutputTechnology).
# Salve este arquivo na mesma pasta do script principal.
# -------------------------------------------------------------------------

#region “Tabela de Mapeamento de Computadores e Portas de Vídeo”
$ModelMap = @{
    "3209N4P"                             = @{ FriendlyName = "Lenovo ThinkCentre M92p";          VideoPorts = "1x VGA / 1x DisplayPort" }
    "10A9000WBP"                          = @{ FriendlyName = "Lenovo ThinkCentre M93p";          VideoPorts = "1x VGA / 1x DisplayPort" }
    "11DUSD3R00"                          = @{ FriendlyName = "Lenovo ThinkCentre M70q";          VideoPorts = "1x VGA / 1x DisplayPort / 1x HDMI" }
    "12TES8R800"                          = @{ FriendlyName = "Lenovo ThinkCentre M70q v2";       VideoPorts = "1x VGA / 1x DisplayPort / 1x HDMI" }
    "HP EliteDesk 800 G5 Desktop Mini"    = @{ FriendlyName = "HP EliteDesk 800 G5 Desktop Mini"; VideoPorts = "1x VGA / 2x DisplayPort" }
    "HP EliteDesk 800 G3 DM 35W (Brazil)" = @{ FriendlyName = "HP EliteDesk 800 G3 DM 35W";       VideoPorts = "1x VGA / 2x DisplayPort" }
    "HP ProDesk 600 G1 SFF"               = @{ FriendlyName = "HP ProDesk 600 G1 SFF";            VideoPorts = "1x VGA / 2x DisplayPort" }
    "HP Z440 Workstation"                 = @{ FriendlyName = "HP Z440 Workstation";              VideoPorts = "1x VGA / Placa de vídeo com 4x DisplayPort" }
    "HP Pro SFF 400 G9 Desktop PC"        = @{ FriendlyName = "HP Pro SFF 400 G9 Desktop PC";     VideoPorts = "1x VGA / 1x DisplayPort / 1x HDMI" }
    "OptiPlex 7050"                       = @{ FriendlyName = "Dell OptiPlex 7050";               VideoPorts = "1x HDMI / 2x DisplayPort" }
    "OptiPlex 3020"                       = @{ FriendlyName = "Dell OptiPlex 3020";               VideoPorts = "1x VGA / 1x DisplayPort" }
    "20W1S6CB00"                          = @{ FriendlyName = "Lenovo T14 Gen 2";                 VideoPorts = "1x HDMI" }
    "13E0S00400"                          = @{ FriendlyName = "Lenovo E14";                       VideoPorts = "1x HDMI" }
}
#endregion

#region “Tabela de Mapeamento de Monitores via código EDID”
# https://devblogs.microsoft.com/scripting/use-powershell-to-discover-multi-monitor-information/
# https://raw.githubusercontent.com/linuxhw/EDID/master/DigitalDisplay.md
$MonitorMap = @{
    "DELA0D7"  = @{ Nome = "Dell P2217H";        Polegada = "21.7";  'Resolução' = "1920x1080"; Tamanho = "480x270mm"; VideoPorts = "1x VGA / 1x DisplayPort / 1x HDMI" }
    "DELA113"  = @{ Nome = "Dell P2217H";        Polegada = "21.7";  'Resolução' = "1920x1080"; Tamanho = "480x270mm"; VideoPorts = "1x VGA / 1x DisplayPort / 1x HDMI" }
    "DELA0D9"  = @{ Nome = "Dell P2217H";        Polegada = "21.7";  'Resolução' = "1920x1080"; Tamanho = "480x270mm"; VideoPorts = "1x VGA / 1x DisplayPort / 1x HDMI" }
    "DELA0D8"  = @{ Nome = "Dell P2217H";        Polegada = "21.7";  'Resolução' = "1920x1080"; Tamanho = "480x270mm"; VideoPorts = "1x VGA / 1x DisplayPort / 1x HDMI" }
    "DEL40F4"  = @{ Nome = "Dell P2317H";        Polegada = "23.1";  'Resolução' = "1920x1080"; Tamanho = "510x290mm"; VideoPorts = "1x VGA / 1x DisplayPort / 1x HDMI" }
    "DEL40F3"  = @{ Nome = "Dell P2317H";        Polegada = "23.1";  'Resolução' = "1920x1080"; Tamanho = "510x290mm"; VideoPorts = "1x VGA / 1x DisplayPort / 1x HDMI" }
    "DELA026"  = @{ Nome = "Dell P2317H";        Polegada = "24.2";  'Resolução' = "1920x1080"; Tamanho = "510x290mm"; VideoPorts = "1x VGA / 1x DisplayPort / 1x HDMI" }
    "DELD03A"  = @{ Nome = "Dell E1914H";        Polegada = "18.5";  'Resolução' = "1366x768";  Tamanho = "410x230mm"; VideoPorts = "1x VGA" }
    "DELA1D4"  = @{ Nome = "Dell C2423H";        Polegada = "24.0";  'Resolução' = "1920x1080"; Tamanho = "510x290mm"; VideoPorts = "1x HDMI / 2x DisplayPort" } # Novo monitor com câmera
    "DELA1D5"  = @{ Nome = "Dell C2423H";        Polegada = "24.0";  'Resolução' = "1920x1080"; Tamanho = "510x290mm"; VideoPorts = "1x HDMI / 2x DisplayPort" }
    "AOC2401"  = @{ Nome = "AOC 24P1U";          Polegada = "24";    'Resolução' = "1920x1080"; Tamanho = "530x300mm"; VideoPorts = "1x VGA / 1x DisplayPort / 1x HDMI" }
    "AOC2470"  = @{ Nome = "AOC M2470PWH";       Polegada = "23.4";  'Resolução' = "1920x1080"; Tamanho = "520x290mm"; VideoPorts = "1x VGA / 2x HDMI" }
    "GSM4ED3"  = @{ Nome = "LG E2011";           Polegada = "20.3";  'Resolução' = "1600x900";  Tamanho = "450x250mm"; VideoPorts = "1x VGA / 1x DVI" }
    "GSM4ED4"  = @{ Nome = "LG E2011";           Polegada = "20.3";  'Resolução' = "1600x900";  Tamanho = "450x250mm"; VideoPorts = "1x VGA / 1x DVI" }
    "GSM4C1E"  = @{ Nome = "LG E2011";           Polegada = "20.3";  'Resolução' = "1600x900";  Tamanho = "450x250mm"; VideoPorts = "1x VGA / 1x DVI" }
    "HWP3139"  = @{ Nome = "HP 20";              Polegada = "20.3";  'Resolução' = "1920x1080"; Tamanho = "450x250mm"; VideoPorts = "1x VGA / 1x DVI" }
    "HWP289A"  = @{ Nome = "HP L190hb";          Polegada = "19.1";  'Resolução' = "1440x900";  Tamanho = "410x260mm"; VideoPorts = "1x VGA / 1x DVI" }
    "HWP3279"  = @{ Nome = "HP H P E3";          Polegada = "23.1";  'Resolução' = "1920x1080"; Tamanho = "510x290mm"; VideoPorts = "1x VGA / 1x DisplayPort / 1x HDMI" }
    "HWP2990"  = @{ Nome = "HP W2072a";          Polegada = "19.9";  'Resolução' = "1600x900";  Tamanho = "440x250mm"; VideoPorts = "1x VGA / 1x DVI" }
    "GSM0001"  = @{ Nome = "LG TV";              Polegada = "41.9";  'Resolução' = "1920x1080"; Tamanho = "930x520mm"; VideoPorts = "" }  # geralmente HDMI
    "SAM0C39"  = @{ Nome = "Samsung";            Polegada = "31.5";  'Resolução' = "1920x1080"; Tamanho = "700x390mm"; VideoPorts = "" }  # geralmente HDMI
    "LEN40A1"  = @{ Nome = "Lenovo";             Polegada = "13.9";  'Resolução' = "1600x900";  Tamanho = "310x170mm"; VideoPorts = "" }  # VGA/HDMI comum
    "LEN61F9"  = @{ Nome = "Lenovo ThinkVision S24e"; Polegada = "24.0"; 'Resolução' = "1920x1080"; Tamanho = "450x250mm"; VideoPorts = "1x VGA / 1x HDMI" }
    "DELD0D8"  = @{ Nome = "Dell P2419H";        Polegada = "23.8";  'Resolução' = "1920x1080"; Tamanho = "530x300mm"; VideoPorts = "1x VGA / 1x DisplayPort / 1x HDMI" }
    "DELD0DA"  = @{ Nome = "Dell P2419H";        Polegada = "23.8";  'Resolução' = "1920x1080"; Tamanho = "530x300mm"; VideoPorts = "1x VGA / 1x DisplayPort / 1x HDMI" }
}
#endregion

#region “Tabela de Mapeamento de Fabricantes de Monitores”
$ManufacturerMap = @{
    "DEL" = 'Dell'
    "HWP" = 'HP'
    "GSM" = 'LG'
    "SAM" = 'Samsung'
    "AOC" = 'AOC'
    "LEN" = 'Lenovo'
    "ACR" = 'Acer'
    "ASU" = 'Asus'
    "BEN" = 'BenQ'
    "SNY" = 'Sony'
    "PHL" = 'Philips'
    "NEC" = 'NEC'
    "VSC" = 'ViewSonic'
}
#endregion

#region “Tabela de Mapeamento Geral de Fabricantes”
# Mapeamento para padronizar nomes de fabricantes (ex: LENOVO -> Lenovo)
$GeneralManufacturerMap = @{
    "LENOVO"          = "Lenovo"
    "DELL"            = "Dell"
    "DELL INC."       = "Dell"
    "HP"              = "HP"
    "HEWLETT-PACKARD" = "HP"
    "INTEL"           = "Intel"
    "REALTEK"         = "Realtek"
    "NVIDIA"          = "Nvidia"
    "MICROSOFT"       = "Microsoft"
}
#endregion

#region “Tabela de Mapeamento de Fabricantes de Memória RAM (Códigos JEDEC)”
# Lista de códigos hexadecimal retornados pelo WMI (Win32_PhysicalMemory) mapeados para fabricantes oficiais
$jedecMap = @{
    0x4B   = 'Kingston'
    0x73   = 'Samsung'
    0x0198 = 'Kingston'
    0x802C = 'Micron'
    0x80AD = 'Hynix'
    0x015B = 'Elpida'
    0x43   = 'Corsair'
    0x4D   = 'Micron'
    0xAD   = 'Hynix'
    0x83   = 'Infinion'
    0x85   = 'Nanya'
    0x02FE = 'Elpida'
    
    # Adições e Expansões de Fabricantes Comuns do MPMS
    0x0194 = 'Smart Modular'
    0x029E = 'ADATA'
    0x04CB = 'ADATA'
    0x014F = 'Transcend'
    0x02C4 = 'Crucial'
    0x059B = 'Crucial'
    0x070B = 'Crucial'
    0x04CD = 'G.Skill'
    0x0325 = 'Patriot'
    0xA    = 'Union Memory'
    0x01A4 = 'Union Memory'
}
#endregion

#region “Tabela de Mapeamento de Tecnologias de Saída de Vídeo”
# Mapeamento oficial do VideoOutputTechnology de acordo com a especificação D3DKMDT_VIDEO_OUTPUT_TECHNOLOGY da Microsoft
$videoTechMap = @{
    -1 = 'Outro'
    0  = 'VGA'
    1  = 'S-Video'
    2  = 'Video Composto'
    3  = 'Video Componente'
    4  = 'DVI'
    5  = 'HDMI'
    6  = 'LVDS (Tela Interna)'
    8  = 'D-JPN'
    9  = 'SDI'
    10 = 'DisplayPort'
    11 = 'DisplayPort (Embedded)'
    12 = 'UDI'
    13 = 'UDI (Embedded)'
    14 = 'SDTV'
    15 = 'Miracast'
    16 = 'Indirect Wired'
}
#endregion

#region “Tabela de Mapeamento de Tipos de Chassi”
# Mapeamento de códigos SMBIOS para descrição de tipo de chassi
$chassisTypeMap = @{
    1  = "Other"
    2  = "Unknown"
    3  = "Desktop"
    4  = "Low Profile Desktop"
    5  = "Pizza Box"
    6  = "Mini Tower"
    7  = "Tower"
    8  = "Portable"
    9  = "Laptop"
    10 = "Notebook"
    11 = "Hand Held"
    12 = "Docking Station"
    13 = "All in One"
    14 = "Sub Notebook"
    15 = "Space-Saving"
    16 = "Lunch Box"
    17 = "Main System Chassis"
    18 = "Expansion Chassis"
    19 = "SubChassis"
    20 = "Bus Expansion Chassis"
    21 = "Peripheral Chassis"
    22 = "RAID Chassis"
    23 = "Rack Mount Chassis"
    24 = "Sealed-Case PC"
    25 = "Multi-System Chassis"
    26 = "CompactPCI"
    27 = "AdvancedTCA"
    28 = "Blade"
    29 = "Blade Enclosure"
    30 = "Tablet"
    31 = "Convertible"
    32 = "Detachable"
    33 = "IoT Gateway"
    34 = "Embedded PC"
    35 = "Mini PC"
    36 = "Stick PC"
}
#endregion

#region “Tabela de Mapeamento de Unidades de Disco”
# Mapeamento de modelos brutos de SSD/HDD para Fabricante e Modelo Amigável
$DiskModelMap = @{
    "SAMSUNG MZVLB256HAHQ-000H1" = @{ Fabricante = "Samsung";        Modelo = "PM981a NVMe 256GB" }
    "SKHynix_HFS256GEJ9X113N"    = @{ Fabricante = "SK Hynix";       Modelo = "BC501 NVMe 256GB" }
    "KINGSTON SA400S37240G"      = @{ Fabricante = "Kingston";       Modelo = "A400 SATA 240GB" }
    "KINGSTON SA400S37480G"      = @{ Fabricante = "Kingston";       Modelo = "A400 SATA 480GB" }
    "KINGSTON SNVS250G"          = @{ Fabricante = "Kingston";       Modelo = "NV1 NVMe 250GB" }
    "KINGSTON SNVS500G"          = @{ Fabricante = "Kingston";       Modelo = "NV1 NVMe 500GB" }
    "CT250MX500SSD1"             = @{ Fabricante = "Crucial";        Modelo = "MX500 SATA 250GB" }
    "CT500MX500SSD1"             = @{ Fabricante = "Crucial";        Modelo = "MX500 SATA 500GB" }
    "WDC WD10EZEX-08WN4A0"       = @{ Fabricante = "Western Digital"; Modelo = "Blue HDD 1TB" }
    "WDC WDS250G2B0A"            = @{ Fabricante = "Western Digital"; Modelo = "Blue SATA 250GB" }
    "ST1000DM010-2EP102"         = @{ Fabricante = "Seagate";        Modelo = "Barracuda HDD 1TB" }
}
#endregion
