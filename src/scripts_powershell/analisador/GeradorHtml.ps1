function Get-DeviceReportHtml {
    param(
        [PSCustomObject]$systemData,
        [System.Collections.ArrayList]$usersData,
        [switch]$SkipMajorData
    )

    $extraSections = ""
    $extraLinks = ""

    if (-not $SkipMajorData) {
        $extraLinks = @'
        <a href="#sec-users-all">👥 Usuários & Perfis</a>
        <a href="#sec-users-chart" class="sub-link">📊 Gráfico de Perfis</a>
        <a href="#sec-programs">📦 Programas Instalados</a>
        <a href="#sec-programs-chart" class="sub-link">📊 Gráfico de Programas</a>
        <a href="#sec-services">⚙️ Serviços em Execução</a>
        <a href="#sec-services-chart" class="sub-link">📊 Gráfico de Serviços</a>
        <a href="#sec-processes">📊 Processos Ativos</a>
        <a href="#sec-processes-chart" class="sub-link">📊 Gráfico de Processos</a>
        <a href="#sec-drivers">🔌 Drivers do Sistema</a>
        <a href="#sec-drivers-chart" class="sub-link">📊 Gráfico de Drivers</a>
'@

        $extraSections = @"
        <div id="sec-users-all" class="section-card">
            <div class="card-header">
                <h2>👥 Usuários & Perfis no Disco</h2>
                <span class="accordion-icon">▼</span>
            </div>
            <div class="card-body">
                $(Convert-ToHtmlFragment -InputObject $usersData)
            </div>
        </div>

        <div id="sec-users-chart" class="section-card">
            <div class="card-header">
                <h2>📊 Gráfico: Consumo por Perfil (GB)</h2>
                <span class="accordion-icon">▼</span>
            </div>
            <div class="card-body">
                <div class="charts-row">
                    <div class="chart-card" style="grid-column: 1 / -1;">
                        <div id="userDiskChartContainer" style="position: relative; height: 260px; display: flex; flex-direction: column; justify-content: center;"></div>
                    </div>
                </div>
            </div>
        </div>

        <div id="sec-programs" class="section-card">
            <div class="card-header">
                <div class="section-header-row">
                    <h2>📦 Programas Instalados</h2>
                    <div class="search-box" onclick="event.stopPropagation()">
                        <span class="search-icon">🔍</span>
                        <input type="text" class="table-filter" data-target="sec-programs" placeholder="Pesquisar programa...">
                        <span class="clear-icon" title="Limpar pesquisa">✖</span>
                    </div>
                </div>
                <span class="accordion-icon">▼</span>
            </div>
            <div class="card-body">
                $(Convert-ToHtmlFragment -InputObject $systemData.InstalledPrograms)
            </div>
        </div>

        <div id="sec-programs-chart" class="section-card">
            <div class="card-header">
                <h2>📊 Gráfico: Programas por Ano de Instalação</h2>
                <span class="accordion-icon">▼</span>
            </div>
            <div class="card-body">
                <div class="charts-row">
                    <div class="chart-card" style="grid-column: 1 / -1;">
                        <div id="programsChartContainer" style="position: relative; height: 260px; display: flex; flex-direction: column; justify-content: center;"></div>
                    </div>
                </div>
            </div>
        </div>

        <div id="sec-services" class="section-card">
            <div class="card-header">
                <div class="section-header-row">
                    <h2>⚙️ Serviços em Execução</h2>
                    <div class="search-box" onclick="event.stopPropagation()">
                        <span class="search-icon">🔍</span>
                        <input type="text" class="table-filter" data-target="sec-services" placeholder="Pesquisar serviço...">
                        <span class="clear-icon" title="Limpar pesquisa">✖</span>
                    </div>
                </div>
                <span class="accordion-icon">▼</span>
            </div>
            <div class="card-body">
                $(Convert-ToHtmlFragment -InputObject $systemData.RunningServices)
            </div>
        </div>

        <div id="sec-services-chart" class="section-card">
            <div class="card-header">
                <h2>📊 Gráfico: Serviços por Tipo de Inicialização</h2>
                <span class="accordion-icon">▼</span>
            </div>
            <div class="card-body">
                <div class="charts-row">
                    <div class="chart-card" style="grid-column: 1 / -1;">
                        <div id="servicesChartContainer" style="position: relative; height: 220px; display: flex; flex-direction: column; justify-content: center;"></div>
                    </div>
                </div>
            </div>
        </div>

        <div id="sec-processes" class="section-card">
            <div class="card-header">
                <div class="section-header-row">
                    <h2>📊 Processos Ativos</h2>
                    <div class="search-box" onclick="event.stopPropagation()">
                        <span class="search-icon">🔍</span>
                        <input type="text" class="table-filter" data-target="sec-processes" placeholder="Pesquisar processo...">
                        <span class="clear-icon" title="Limpar pesquisa">✖</span>
                    </div>
                </div>
                <span class="accordion-icon">▼</span>
            </div>
            <div class="card-body">
                $(Convert-ToHtmlFragment -InputObject $systemData.ActiveProcesses)
            </div>
        </div>

        <div id="sec-processes-chart" class="section-card">
            <div class="card-header">
                <h2>📊 Gráfico: Métricas e Uso de RAM por Processo</h2>
                <span class="accordion-icon">▼</span>
            </div>
            <div class="card-body">
                <div class="charts-row">
                    <div class="chart-card" style="grid-column: 1 / -1;">
                        <div class="chart-stats" style="margin-bottom: 15px;">
                            <div class="stat-box"><span class="stat-val" id="stat-total-proc">0</span><span class="stat-lbl">Processos Totais</span></div>
                            <div class="stat-box"><span class="stat-val" id="stat-total-threads">0</span><span class="stat-lbl">Threads Totais</span></div>
                            <div class="stat-box"><span class="stat-val" id="stat-total-handles">0</span><span class="stat-lbl">Handles Totais</span></div>
                        </div>
                        <div id="processesChartContainer" style="position: relative; height: 260px; display: flex; flex-direction: column; justify-content: center;"></div>
                    </div>
                </div>
            </div>
        </div>

        <div id="sec-drivers" class="section-card">
            <div class="card-header">
                <div class="section-header-row">
                    <h2>🔌 Drivers do Sistema</h2>
                    <div class="search-box" onclick="event.stopPropagation()">
                        <span class="search-icon">🔍</span>
                        <input type="text" class="table-filter" data-target="sec-drivers" placeholder="Pesquisar driver...">
                        <span class="clear-icon" title="Limpar pesquisa">✖</span>
                    </div>
                </div>
                <span class="accordion-icon">▼</span>
            </div>
            <div class="card-body">
                $(Convert-ToHtmlFragment -InputObject $systemData.InstalledDrivers)
            </div>
        </div>

        <div id="sec-drivers-chart" class="section-card">
            <div class="card-header">
                <h2>📊 Gráfico: Origem dos Drivers (Microsoft vs Terceiros)</h2>
                <span class="accordion-icon">▼</span>
            </div>
            <div class="card-body">
                <div class="charts-row">
                    <div class="chart-card" style="grid-column: 1 / -1;">
                        <div id="driversChartContainer" style="position: relative; height: 220px; display: flex; flex-direction: column; justify-content: center;"></div>
                    </div>
                </div>
            </div>
        </div>
"@
    }

    $mappedPrintersHtml = ""
    if ($systemData.PrintersMapped) {
        $mappedPrintersHtml = '<h3 style="margin-top:15px; font-size:14px; color:#475569;">Impressoras Mapeadas de Usuários</h3>' + (Convert-ToHtmlFragment -InputObject $systemData.PrintersMapped)
    }

    $currentUserObj = [PSCustomObject]@{
        'Usuário Logado agora' = $systemData.CurrentUser
        'DisplayName AD'       = $systemData.CurrentUserDisplayName
        'Domínio'              = $systemData.CurrentUserDomain
        'Nome do Computador'   = $systemData.HardwareInfo.Computador
    }

    $nowStr = Get-Date -Format 'dd/MM/yyyy HH:mm:ss'
    $compName = $systemData.HardwareInfo.Computador
    $hostName = $systemData.OS.'Nome da Máquina'

    $htmlContent = @"
<!DOCTYPE html>
<html>
<head>
    <meta charset="utf-8" />
    <title>Relatório de $compName ($hostName)</title>
    <link href="https://fonts.googleapis.com/css2?family=Inter:wght@300;400;500;600;700&display=swap" rel="stylesheet">
    <style>
        /* Custom Scrollbars (Chrome, Edge, Safari) */
        ::-webkit-scrollbar {
            width: 6px;
            height: 6px;
        }
        ::-webkit-scrollbar-track {
            background: #f1f5f9;
            border-radius: 4px;
        }
        ::-webkit-scrollbar-thumb {
            background: #cbd5e1;
            border-radius: 4px;
        }
        ::-webkit-scrollbar-thumb:hover {
            background: #3b82f6;
        }

        /* Custom Scrollbars para Firefox (Padrão W3C CSS Scrollbars) */
        * {
            scrollbar-width: thin;
            scrollbar-color: #cbd5e1 #f1f5f9;
        }

        /* === Estilos gerais === */
        html {
            scroll-behavior: smooth;
        }
        body {
            font-family: 'Inter', system-ui, -apple-system, sans-serif;
            margin: 0;
            padding: 0;
            display: flex;
            background-color: #f8fafc;
            color: #1e293b;
            transition: margin-left 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        }

        /* === Quando body tiver a classe “sidebar-hidden”, recolhe a sidebar === */
        body.sidebar-hidden .sidebar {
            transform: translateX(-100%);
        }
        body.sidebar-hidden .main-content {
            margin-left: 20px;
            width: calc(100% - 20px);
        }

        /* === Botão “hamburger” para alternar sidebar === */
        #btn-toggle {
            position: fixed;
            top: 15px;
            left: 15px;
            width: 36px;
            height: 36px;
            background: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 8px;
            cursor: pointer;
            z-index: 1000;
            display: flex;
            flex-direction: column;
            justify-content: center;
            align-items: center;
            gap: 4px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.1);
            transition: all 0.2s ease;
        }
        #btn-toggle span {
            display: block;
            width: 18px;
            height: 2px;
            background-color: #475569;
            border-radius: 1px;
            transition: all 0.2s ease;
        }
        #btn-toggle:hover {
            background-color: #f1f5f9;
            border-color: #cbd5e1;
        }

        /* === Sidebar === */
        .sidebar {
            width: 250px;
            background: #0f172a;
            position: fixed;
            top: 0;
            bottom: 0;
            left: 0;
            padding: 70px 16px 20px 16px;
            box-sizing: border-box;
            overflow-y: auto;
            transition: transform 0.3s cubic-bezier(0.4, 0, 0.2, 1);
            z-index: 900;
            box-shadow: 2px 0 8px rgba(0,0,0,0.15);
        }
        .sidebar-title {
            font-size: 11px;
            font-weight: 700;
            text-transform: uppercase;
            letter-spacing: 1px;
            color: #38bdf8;
            margin: 0 0 16px 12px;
            border-bottom: 1px solid #1e293b;
            padding-bottom: 8px;
        }
        .sidebar a {
            display: flex;
            align-items: center;
            gap: 10px;
            color: #cbd5e1;
            text-decoration: none;
            padding: 8px 12px;
            margin-bottom: 2px;
            border-radius: 6px;
            font-size: 13px;
            font-weight: 500;
            transition: all 0.15s ease;
        }
        .sidebar a.sub-link {
            padding-left: 24px;
            font-size: 12px;
            color: #94a3b8;
        }
        .sidebar a:hover {
            color: #38bdf8;
            background-color: #1e293b;
            transform: translateX(2px);
        }
        .sidebar a.active {
            background-color: #1e293b;
            color: #38bdf8;
            font-weight: 600;
            border-left: 3px solid #38bdf8;
            padding-left: 9px;
        }
        .sidebar a.sub-link.active {
            padding-left: 21px;
        }

        /* === Conteúdo Principal === */
        .main-content {
            margin-left: 270px;
            padding: 30px 40px;
            width: calc(100% - 270px);
            box-sizing: border-box;
            transition: margin-left 0.3s cubic-bezier(0.4, 0, 0.2, 1), width 0.3s cubic-bezier(0.4, 0, 0.2, 1);
        }

        /* Header / Banner Principal */
        .report-header {
            background: linear-gradient(135deg, #1e293b 0%, #0f172a 100%);
            color: #ffffff;
            padding: 24px 30px;
            border-radius: 12px;
            margin-bottom: 24px;
            box-shadow: 0 4px 6px -1px rgba(0,0,0,0.1);
        }
        .report-header h1 {
            margin: 0 0 8px 0;
            font-size: 24px;
            font-weight: 700;
        }
        .report-header .subtitle {
            margin: 0;
            color: #94a3b8;
            font-size: 14px;
        }

        /* Seções Accordion */
        .section-card {
            background: #ffffff;
            border: 1px solid #e2e8f0;
            border-radius: 10px;
            margin-bottom: 20px;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
            overflow: hidden;
            transition: all 0.2s ease;
        }
        .card-header {
            padding: 16px 24px;
            cursor: pointer;
            display: flex;
            justify-content: space-between;
            align-items: center;
            user-select: none;
            background-color: #ffffff;
            transition: background-color 0.15s ease;
        }
        .card-header:hover {
            background-color: #f8fafc;
        }
        .card-header h2 {
            margin: 0;
            font-size: 16px;
            font-weight: 600;
            color: #0f172a;
            display: flex;
            align-items: center;
            gap: 8px;
            border-bottom: none;
            padding-bottom: 0;
            flex: 1;
        }
        .accordion-icon {
            font-size: 12px;
            color: #64748b;
            transition: transform 0.25s cubic-bezier(0.4, 0, 0.2, 1);
            margin-left: 12px;
        }
        .section-card.collapsed .accordion-icon {
            transform: rotate(-90deg);
        }
        .card-body {
            padding: 0 24px 20px 24px;
            border-top: 1px solid #f1f5f9;
        }
        .section-card.collapsed .card-body {
            display: none;
        }

        /* Header de Seção com Busca */
        .section-header-row {
            display: flex;
            justify-content: space-between;
            align-items: center;
            width: 100%;
        }
        .search-box {
            position: relative;
            display: flex;
            align-items: center;
            width: 260px;
        }
        .search-box .search-icon {
            position: absolute;
            left: 12px;
            color: #64748b;
            font-size: 0.9em;
            pointer-events: none;
        }
        .search-box .table-filter {
            width: 100%;
            padding: 6px 30px 6px 34px;
            font-size: 0.85em;
            font-family: 'Inter', system-ui, sans-serif;
            color: #1e293b;
            background-color: #f8fafc;
            border: 1px solid #cbd5e1;
            border-radius: 20px;
            outline: none;
            box-sizing: border-box;
            transition: all 0.2s ease;
        }
        .search-box .table-filter:focus {
            background-color: #ffffff;
            border-color: #3b82f6;
            box-shadow: 0 0 0 3px rgba(59, 130, 246, 0.15);
        }
        .search-box .clear-icon {
            position: absolute;
            right: 12px;
            color: #94a3b8;
            font-size: 0.9em;
            cursor: pointer;
            display: none;
            transition: color 0.15s ease;
        }
        .search-box .clear-icon:hover {
            color: #ef4444;
        }
        .no-results-row {
            text-align: center;
            padding: 16px !important;
            color: #64748b;
            font-style: italic;
        }

        /* Tabelas */
        .table-container {
            overflow-x: auto;
            margin-top: 15px;
        }
        table {
            width: 100%;
            border-collapse: collapse;
            font-size: 13px;
            text-align: left;
        }
        th {
            background-color: #f8fafc;
            color: #475569;
            font-weight: 600;
            padding: 10px 12px;
            border-bottom: 1px solid #e2e8f0;
            white-space: nowrap;
            user-select: none;
        }
        td {
            padding: 10px 12px;
            border-bottom: 1px solid #f1f5f9;
            color: #334155;
        }
        tr:nth-child(even) {
            background-color: #f8fafc;
        }
        tr:hover td {
            background-color: #f1f5f9;
        }

        /* Ordenação de Cabeçalho */
        th.sortable {
            cursor: pointer;
            position: relative;
            padding-right: 28px;
            transition: background-color 0.15s ease, color 0.15s ease;
        }
        th.sortable:hover {
            background-color: #e2e8f0;
            color: #0f172a;
        }
        .sort-arrow {
            position: absolute;
            right: 10px;
            top: 50%;
            transform: translateY(-50%);
            font-size: 0.85em;
            opacity: 0.4;
            transition: all 0.15s ease;
        }
        th.sortable:hover .sort-arrow {
            opacity: 1;
        }
        th.sort-asc .sort-arrow, th.sort-desc .sort-arrow {
            opacity: 1;
            color: #2563eb;
        }

        /* Gráficos */
        .charts-row {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(350px, 1fr));
            gap: 20px;
            margin-top: 15px;
        }
        .chart-card {
            background: #ffffff;
            border-radius: 8px;
            border: 1px solid #e2e8f0;
            box-shadow: 0 1px 3px rgba(0,0,0,0.05);
            padding: 16px;
            box-sizing: border-box;
            display: flex;
            flex-direction: column;
        }
        .chart-stats {
            display: grid;
            grid-template-columns: repeat(3, 1fr);
            gap: 12px;
        }
        .stat-box {
            background-color: #f8fafc;
            border: 1px solid #e2e8f0;
            border-radius: 6px;
            padding: 10px 8px;
            text-align: center;
            transition: all 0.2s ease;
        }
        .stat-box:hover {
            background-color: #f1f5f9;
            transform: translateY(-2px);
        }
        .stat-val {
            display: block;
            font-size: 1.4em;
            font-weight: 700;
            color: #2563eb;
            margin-bottom: 2px;
        }
        .stat-lbl {
            font-size: 0.7em;
            color: #64748b;
            text-transform: uppercase;
            font-weight: 600;
            letter-spacing: 0.05em;
        }
        .fallback-bars-list {
            display: flex;
            flex-direction: column;
            gap: 10px;
        }
        .bar-row {
            display: flex;
            align-items: center;
            gap: 10px;
            font-size: 0.85em;
            cursor: pointer;
            padding: 2px 0;
        }
        .bar-label {
            width: 140px;
            font-weight: 500;
            color: #334155;
            white-space: nowrap;
            overflow: hidden;
            text-overflow: ellipsis;
        }
        .bar-track {
            flex: 1;
            height: 10px;
            background-color: #f1f5f9;
            border-radius: 5px;
            overflow: hidden;
            position: relative;
        }
        .bar-fill {
            height: 100%;
            background: linear-gradient(90deg, #3b82f6 0%, #2563eb 100%);
            border-radius: 5px;
            transition: width 0.8s cubic-bezier(0.4, 0, 0.2, 1);
        }
        .bar-row:hover .bar-fill {
            background: linear-gradient(90deg, #60a5fa 0%, #3b82f6 100%);
        }
        .bar-value {
            width: 75px;
            text-align: right;
            font-weight: 600;
            color: #1e293b;
        }

        /* Rodapé */
        .report-footer {
            text-align: center;
            padding: 20px 0;
            color: #94a3b8;
            font-size: 12px;
        }

        /* Responsividade para Impressão */
        @media print {
            .sidebar, #btn-toggle, .search-box, .accordion-icon { display: none !important; }
            .main-content { margin-left: 0 !important; width: 100% !important; padding: 0 !important; }
            .section-card { page-break-inside: avoid; border: 1px solid #ccc; box-shadow: none; margin-bottom: 15px; }
            .section-card.collapsed .card-body { display: block !important; }
        }
    </style>
    <script src="https://cdn.jsdelivr.net/npm/chart.js" defer></script>
</head>
<body>
    <button id="btn-toggle" onclick="toggleSidebar()" title="Alternar Menu">
        <span></span><span></span><span></span>
    </button>

    <div class="sidebar">
        <div class="sidebar-title">Navegação Rápida</div>
        <a href="#sec-hw">💻 Hardware Principal</a>
        <a href="#sec-user">👤 Usuário Atual</a>
        <a href="#sec-os">🖥️ Sistema Operacional</a>
        <a href="#sec-bios">⚙️ BIOS</a>
        <a href="#sec-cpu">🧠 Processador</a>
        <a href="#sec-mb">🎛️ Placa-Mãe</a>
        <a href="#sec-chassis">📦 Gabinete / Chassi</a>
        <a href="#sec-ram">⚡ Memória RAM</a>
        <a href="#sec-ram-chart" class="sub-link">📊 Gráfico de RAM</a>
        <a href="#sec-disks">💾 Discos Físicos</a>
        <a href="#sec-ldisks">💽 Discos Lógicos</a>
        <a href="#sec-disk-chart" class="sub-link">📊 Gráfico de Discos</a>
        <a href="#sec-net">🌐 Rede e Conexões</a>
        <a href="#sec-monitors">🖥️ Monitores</a>
        <a href="#sec-gpu">🎮 Vídeo / GPU</a>
        <a href="#sec-audio">🔊 Dispositivos de Áudio</a>
        <a href="#sec-printers">🖨️ Impressoras</a>
        $extraLinks
    </div>

    <div class="main-content">
        <div class="report-header">
            <h1>Relatório de $compName ($hostName)</h1>
            <div class="subtitle">Dispositivo: <strong>$compName</strong> | Nome da Máquina: <strong>$hostName</strong> | Gerado em: $nowStr</div>
        </div>

        <div id="sec-hw" class="section-card">
            <div class="card-header">
                <h2>💻 Hardware Principal</h2>
                <span class="accordion-icon">▼</span>
            </div>
            <div class="card-body">
                $(Convert-ToHtmlFragment -InputObject $systemData.HardwareInfo)
            </div>
        </div>

        <div id="sec-user" class="section-card">
            <div class="card-header">
                <h2>👤 Usuário Atualmente Logado</h2>
                <span class="accordion-icon">▼</span>
            </div>
            <div class="card-body">
                $(Convert-ToHtmlFragment -InputObject $currentUserObj)
            </div>
        </div>

        <div id="sec-os" class="section-card">
            <div class="card-header">
                <h2>🖥️ Sistema Operacional</h2>
                <span class="accordion-icon">▼</span>
            </div>
            <div class="card-body">
                $(Convert-ToHtmlFragment -InputObject $systemData.OS)
            </div>
        </div>

        <div id="sec-bios" class="section-card">
            <div class="card-header">
                <h2>⚙️ BIOS</h2>
                <span class="accordion-icon">▼</span>
            </div>
            <div class="card-body">
                $(Convert-ToHtmlFragment -InputObject $systemData.BIOS)
            </div>
        </div>

        <div id="sec-cpu" class="section-card">
            <div class="card-header">
                <h2>🧠 Processador</h2>
                <span class="accordion-icon">▼</span>
            </div>
            <div class="card-body">
                $(Convert-ToHtmlFragment -InputObject $systemData.Processor)
            </div>
        </div>

        <div id="sec-mb" class="section-card">
            <div class="card-header">
                <h2>🎛️ Placa-Mãe</h2>
                <span class="accordion-icon">▼</span>
            </div>
            <div class="card-body">
                $(Convert-ToHtmlFragment -InputObject $systemData.Motherboard)
            </div>
        </div>

        <div id="sec-chassis" class="section-card">
            <div class="card-header">
                <h2>📦 Gabinete / Chassi</h2>
                <span class="accordion-icon">▼</span>
            </div>
            <div class="card-body">
                $(Convert-ToHtmlFragment -InputObject $systemData.Chassis)
            </div>
        </div>

        <div id="sec-ram" class="section-card">
            <div class="card-header">
                <h2>⚡ Memória RAM</h2>
                <span class="accordion-icon">▼</span>
            </div>
            <div class="card-body">
                $(Convert-ToHtmlFragment -InputObject $systemData.MemoryModules)
            </div>
        </div>

        <div id="sec-ram-chart" class="section-card">
            <div class="card-header">
                <h2>📊 Gráfico: Uso de Slots de Memória RAM</h2>
                <span class="accordion-icon">▼</span>
            </div>
            <div class="card-body">
                <div class="charts-row">
                    <div class="chart-card" style="grid-column: 1 / -1;">
                        <div id="memoryChartContainer" style="position: relative; height: 180px; display: flex; flex-direction: column; justify-content: center;"></div>
                    </div>
                </div>
            </div>
        </div>

        <div id="sec-disks" class="section-card">
            <div class="card-header">
                <h2>💾 Discos Físicos</h2>
                <span class="accordion-icon">▼</span>
            </div>
            <div class="card-body">
                $(Convert-ToHtmlFragment -InputObject $systemData.DiskDrives)
            </div>
        </div>

        <div id="sec-ldisks" class="section-card">
            <div class="card-header">
                <h2>💽 Discos Lógicos</h2>
                <span class="accordion-icon">▼</span>
            </div>
            <div class="card-body">
                $(Convert-ToHtmlFragment -InputObject $systemData.LogicalDisks)
            </div>
        </div>

        <div id="sec-disk-chart" class="section-card">
            <div class="card-header">
                <h2>📊 Gráfico: Armazenamento dos Discos</h2>
                <span class="accordion-icon">▼</span>
            </div>
            <div class="card-body">
                <div class="charts-row">
                    <div class="chart-card" style="grid-column: 1 / -1;">
                        <div id="logicalDiskChartsContainer" style="display: flex; flex-wrap: wrap; gap: 20px; justify-content: center; align-items: center; min-height: 240px;"></div>
                    </div>
                </div>
            </div>
        </div>

        <div id="sec-net" class="section-card">
            <div class="card-header">
                <h2>🌐 Rede e Conexões</h2>
                <span class="accordion-icon">▼</span>
            </div>
            <div class="card-body">
                $(Convert-ToHtmlFragment -InputObject $systemData.NetworkConfigs)
            </div>
        </div>

        <div id="sec-monitors" class="section-card">
            <div class="card-header">
                <h2>🖥️ Monitores Conectados</h2>
                <span class="accordion-icon">▼</span>
            </div>
            <div class="card-body">
                $(Convert-ToHtmlFragment -InputObject $systemData.Monitors)
            </div>
        </div>

        <div id="sec-gpu" class="section-card">
            <div class="card-header">
                <h2>🎮 Adaptador de Vídeo / GPU</h2>
                <span class="accordion-icon">▼</span>
            </div>
            <div class="card-body">
                $(Convert-ToHtmlFragment -InputObject $systemData.VideoController)
            </div>
        </div>

        <div id="sec-audio" class="section-card">
            <div class="card-header">
                <h2>🔊 Dispositivos de Áudio</h2>
                <span class="accordion-icon">▼</span>
            </div>
            <div class="card-body">
                $(Convert-ToHtmlFragment -InputObject $systemData.SoundDevices)
            </div>
        </div>

        <div id="sec-printers" class="section-card">
            <div class="card-header">
                <h2>🖨️ Impressoras Locais e Mapeadas</h2>
                <span class="accordion-icon">▼</span>
            </div>
            <div class="card-body">
                $(Convert-ToHtmlFragment -InputObject $systemData.PrintersLocal)
                $mappedPrintersHtml
            </div>
        </div>

        $extraSections

        <div class="report-footer">
            Relatório gerado automaticamente pela suíte de Automação de Suporte Técnico MPMS.
        </div>
    </div>

    <script>
        function toggleSidebar() {
            document.body.classList.toggle('sidebar-hidden');
        }

        document.addEventListener('DOMContentLoaded', function() {
            // 0) Funcionalidade de Accordions (Expandir/Recolher Seções)
            var cardHeaders = document.querySelectorAll('.section-card .card-header');
            cardHeaders.forEach(function(header) {
                header.addEventListener('click', function() {
                    var card = header.closest('.section-card');
                    if (card) {
                        card.classList.toggle('collapsed');
                    }
                });
            });

            // 1) Ordenação inteligente de todas as tabelas ao clicar no <th>
            var tables = document.querySelectorAll('.main-content table');
            tables.forEach(function(table) {
                var headers = table.querySelectorAll('th');
                if (headers.length === 0) return;

                headers.forEach(function(header, colIndex) {
                    header.classList.add('sortable');
                    header.title = 'Clique para ordenar esta coluna';

                    var arrow = document.createElement('span');
                    arrow.className = 'sort-arrow';
                    arrow.innerHTML = '↕';
                    header.appendChild(arrow);

                    header.addEventListener('click', function(e) {
                        e.stopPropagation();
                        var currentOrder = header.getAttribute('data-sort') || 'none';
                        var nextOrder = currentOrder === 'asc' ? 'desc' : 'asc';

                        headers.forEach(function(h) {
                            h.removeAttribute('data-sort');
                            h.classList.remove('sort-asc', 'sort-desc');
                            var a = h.querySelector('.sort-arrow');
                            if (a) a.innerHTML = '↕';
                        });

                        header.setAttribute('data-sort', nextOrder);
                        header.classList.add(nextOrder === 'asc' ? 'sort-asc' : 'sort-desc');
                        arrow.innerHTML = nextOrder === 'asc' ? '▲' : '▼';

                        var rows = Array.from(table.querySelectorAll('tr')).filter(function(row) {
                            return !row.querySelector('th') && !row.classList.contains('no-results-row');
                        });

                        if (rows.length <= 1) return;

                        var isAsc = nextOrder === 'asc';

                        rows.sort(function(rowA, rowB) {
                            var cellA = rowA.cells[colIndex] ? rowA.cells[colIndex].textContent.trim() : '';
                            var cellB = rowB.cells[colIndex] ? rowB.cells[colIndex].textContent.trim() : '';

                            if (cellA === cellB) return 0;

                            var cleanNum = function(str) {
                                if (str.indexOf('/') !== -1 && str.indexOf(' ') === -1) return NaN;
                                var s = str.replace(/[^\d.,-]/g, '').replace(',', '.');
                                return s !== '' ? parseFloat(s) : NaN;
                            };

                            var numA = cleanNum(cellA);
                            var numB = cleanNum(cellB);

                            if (!isNaN(numA) && !isNaN(numB)) {
                                return isAsc ? numA - numB : numB - numA;
                            }

                            var parseDate = function(str) {
                                var parts = str.split(' ');
                                var datePart = parts[0] ? parts[0].split('/') : [];
                                if (datePart.length === 3) {
                                    var timePart = parts[1] ? parts[1].split(':') : [0, 0, 0];
                                    var d = parseInt(datePart[0], 10);
                                    var m = parseInt(datePart[1], 10) - 1;
                                    var y = parseInt(datePart[2], 10);
                                    var hr = timePart[0] ? parseInt(timePart[0], 10) : 0;
                                    var min = timePart[1] ? parseInt(timePart[1], 10) : 0;
                                    var sec = timePart[2] ? parseInt(timePart[2], 10) : 0;
                                    return new Date(y, m, d, hr, min, sec);
                                }
                                return null;
                            };

                            var dateA = parseDate(cellA);
                            var dateB = parseDate(cellB);

                            var isDateA = dateA && !isNaN(dateA.getTime());
                            var isDateB = dateB && !isNaN(dateB.getTime());

                            if (isDateA || isDateB) {
                                if (isDateA && isDateB) {
                                    return isAsc ? dateA - dateB : dateB - dateA;
                                }
                                return isDateA ? (isAsc ? -1 : 1) : (isAsc ? 1 : -1);
                            }

                            return isAsc 
                                ? cellA.localeCompare(cellB, undefined, { numeric: true, sensitivity: 'base' })
                                : cellB.localeCompare(cellA, undefined, { numeric: true, sensitivity: 'base' });
                        });

                        var parent = rows[0].parentNode;
                        rows.forEach(function(row) {
                            parent.appendChild(row);
                        });
                    });
                });
            });

            // 2) Scroll Spy Inteligente (inclui links normais e sub-links dos gráficos)
            var sections = document.querySelectorAll('.main-content .section-card[id]');
            var navLinks = document.querySelectorAll('.sidebar a[href^="#"]');

            var updateActiveSection = function() {
                var scrollPosition = window.scrollY || document.documentElement.scrollTop;
                var windowHeight = window.innerHeight;
                var docHeight = document.documentElement.scrollHeight;
                var currentSectionId = '';

                if (scrollPosition + windowHeight >= docHeight - 15) {
                    if (sections.length > 0) {
                        currentSectionId = sections[sections.length - 1].getAttribute('id');
                    }
                } else {
                    for (var i = 0; i < sections.length; i++) {
                        var section = sections[i];
                        var sectionTop = section.offsetTop - 120;
                        if (scrollPosition >= sectionTop) {
                            currentSectionId = section.getAttribute('id');
                        } else {
                            break;
                        }
                    }
                }

                if (currentSectionId) {
                    navLinks.forEach(function(link) {
                        if (link.getAttribute('href') === '#' + currentSectionId) {
                            link.classList.add('active');
                            var sidebar = document.querySelector('.sidebar');
                            if (sidebar && (link.offsetTop < sidebar.scrollTop || link.offsetTop > sidebar.scrollTop + sidebar.clientHeight)) {
                                link.scrollIntoView({ behavior: 'smooth', block: 'nearest' });
                            }
                        } else {
                            link.classList.remove('active');
                        }
                    });
                }
            };

            window.addEventListener('scroll', updateActiveSection, { passive: true });
            updateActiveSection();

            // 3) Filtro Dinâmico de Pesquisa nas Tabelas (ao pesquisar, expande o card se estiver recolhido)
            var filterInputs = document.querySelectorAll('.table-filter');
            filterInputs.forEach(function(input) {
                var clearBtn = input.nextElementSibling;
                var targetId = input.getAttribute('data-target');
                var cardDiv = document.getElementById(targetId);
                var table = cardDiv ? cardDiv.querySelector('table') : null;
                
                if (!table) return;

                var tbody = table.querySelector('tbody') || table;
                var rows = tbody.querySelectorAll('tr');
                var colCount = table.querySelector('tr') ? table.querySelector('tr').cells.length : 1;

                var filterTable = function() {
                    var query = input.value.trim().toLowerCase();
                    if (clearBtn) {
                        clearBtn.style.display = query.length > 0 ? 'inline' : 'none';
                    }

                    if (query.length > 0 && cardDiv.classList.contains('collapsed')) {
                        cardDiv.classList.remove('collapsed');
                    }

                    var noRes = tbody.querySelector('.no-results-row');
                    if (noRes) noRes.remove();

                    var visibleCount = 0;
                    rows.forEach(function(row) {
                        if (row.querySelector('th')) return;
                        
                        var rowText = row.textContent.toLowerCase();
                        if (query === '' || rowText.indexOf(query) !== -1) {
                            row.style.display = '';
                            visibleCount++;
                        } else {
                            row.style.display = 'none';
                        }
                    });

                    if (visibleCount === 0 && query !== '') {
                        var trNoRes = document.createElement('tr');
                        trNoRes.className = 'no-results-row';
                        trNoRes.innerHTML = '<td colspan="' + colCount + '">Sem registros para "' + input.value.trim() + '"</td>';
                        tbody.appendChild(trNoRes);
                    }

                    if (targetId === 'sec-processes' && typeof updateProcessesChart === 'function') {
                        updateProcessesChart();
                    }
                };

                input.addEventListener('input', filterTable);
                if (clearBtn) {
                    clearBtn.addEventListener('click', function(e) {
                        e.stopPropagation();
                        input.value = '';
                        filterTable();
                        input.focus();
                    });
                }
            });

            // 4) Geração de Gráficos (Chart.js ou Fallback CSS)
            
            // a) Gráfico de Memória RAM
            var secRam = document.getElementById('sec-ram');
            var tableRam = secRam ? secRam.querySelector('table') : null;
            if (tableRam) {
                var memoryData = [];
                var modIdx = 1;
                tableRam.querySelectorAll('tr').forEach(function(row) {
                    if (row.querySelector('th')) return;
                    var cells = row.cells;
                    if (cells.length >= 2) {
                        var capText = cells[1].innerText.trim();
                        var capVal = parseFloat(capText.replace(/[^\d.,]/g, '').replace(',', '.')) || 0;
                        memoryData.push({ label: 'Módulo ' + modIdx++, capacity: capVal });
                    }
                });

                var memContainer = document.getElementById('memoryChartContainer');
                if (memContainer && memoryData.length > 0) {
                    if (typeof Chart !== 'undefined') {
                        memContainer.innerHTML = '<canvas id="memoryChart" style="width: 100%; height: 180px;"></canvas>';
                        var ctxM = document.getElementById('memoryChart').getContext('2d');
                        new Chart(ctxM, {
                            type: 'bar',
                            data: {
                                labels: memoryData.map(function(m) { return m.label; }),
                                datasets: [{
                                    label: 'Capacidade (GB)',
                                    data: memoryData.map(function(m) { return m.capacity; }),
                                    backgroundColor: 'rgba(59, 130, 246, 0.2)',
                                    borderColor: '#3b82f6',
                                    borderWidth: 2,
                                    borderRadius: 4
                                }]
                            },
                            options: {
                                responsive: true,
                                maintainAspectRatio: false,
                                plugins: { legend: { display: false } },
                                scales: {
                                    x: { grid: { display: false } },
                                    y: { beginAtZero: true, grace: '15%', ticks: { callback: function(v) { return v + ' GB'; } } }
                                }
                            }
                        });
                    } else {
                        var maxCap = Math.max.apply(null, memoryData.map(function(m) { return m.capacity; })) || 1;
                        var htmlM = '<div class="fallback-bars-list">';
                        memoryData.forEach(function(m) {
                            var pct = (m.capacity / maxCap) * 100;
                            htmlM += '<div class="bar-row"><div class="bar-label">' + m.label + '</div><div class="bar-track"><div class="bar-fill" style="width:' + pct + '%;"></div></div><div class="bar-value">' + m.capacity + ' GB</div></div>';
                        });
                        htmlM += '</div>';
                        memContainer.innerHTML = htmlM;
                    }
                }
            }

            // b) Gráfico de Armazenamento de Discos Lógicos
            var secLdisks = document.getElementById('sec-ldisks');
            var tableDisks = secLdisks ? secLdisks.querySelector('table') : null;
            if (tableDisks) {
                var diskData = [];
                tableDisks.querySelectorAll('tr').forEach(function(row) {
                    if (row.querySelector('th')) return;
                    var cells = row.cells;
                    if (cells.length >= 4) {
                        var devId = cells[0].innerText.trim();
                        var freeText = cells[1].innerText.trim();
                        var usedText = cells[2].innerText.trim();
                        var totalText = cells[3].innerText.trim();
                        var volName = cells[4] ? cells[4].innerText.trim() : '';

                        var freeVal = parseFloat(freeText.replace(/[^\d.,]/g, '').replace(',', '.')) || 0;
                        var usedVal = parseFloat(usedText.replace(/[^\d.,]/g, '').replace(',', '.')) || 0;
                        var totalVal = parseFloat(totalText.replace(/[^\d.,]/g, '').replace(',', '.')) || 0;

                        if (totalVal > 0) {
                            diskData.push({ devId: devId, volName: volName, free: freeVal, used: usedVal, total: totalVal });
                        }
                    }
                });

                var diskContainer = document.getElementById('logicalDiskChartsContainer');
                if (diskContainer && diskData.length > 0) {
                    diskContainer.innerHTML = '';
                    diskData.forEach(function(disk, idx) {
                        var cardId = 'diskPieChart_' + idx;
                        var wrapper = document.createElement('div');
                        wrapper.style.cssText = 'flex: 1 1 280px; max-width: 380px; background: #ffffff; border: 1px solid #e2e8f0; border-radius: 8px; padding: 15px; text-align: center;';

                        var title = document.createElement('h4');
                        title.style.cssText = 'margin: 0 0 12px 0; color: #1e293b; font-size: 15px; font-weight: 600;';
                        title.innerText = 'Disco ' + disk.devId + (disk.volName && disk.volName !== 'Windows' ? ' (' + disk.volName + ')' : '');
                        wrapper.appendChild(title);

                        if (typeof Chart !== 'undefined') {
                            var canvasWrapper = document.createElement('div');
                            canvasWrapper.style.cssText = 'position: relative; width: 100%; height: 200px;';
                            canvasWrapper.innerHTML = '<canvas id="' + cardId + '"></canvas>';
                            wrapper.appendChild(canvasWrapper);
                            diskContainer.appendChild(wrapper);

                            var ctxD = document.getElementById(cardId).getContext('2d');
                            new Chart(ctxD, {
                                type: 'pie',
                                data: {
                                    labels: ['Usado (GB)', 'Livre (GB)'],
                                    datasets: [{
                                        data: [disk.used, disk.free],
                                        backgroundColor: ['#0284c7', '#38bdf8'],
                                        borderWidth: 2,
                                        borderColor: '#ffffff'
                                    }]
                                },
                                options: {
                                    responsive: true,
                                    maintainAspectRatio: false,
                                    plugins: { legend: { position: 'bottom' } }
                                }
                            });
                        } else {
                            var pctUsed = (disk.used / disk.total) * 100;
                            var pctFree = 100 - pctUsed;
                            var fallbackHtml = '<div style="display: flex; flex-direction: column; align-items: center; justify-content: center; height: 180px;">' +
                                '<div style="width: 80px; height: 80px; border-radius: 50%; background: conic-gradient(#0284c7 0% ' + pctUsed + '%, #38bdf8 ' + pctUsed + '% 100%); margin-bottom: 15px;"></div>' +
                                '<div style="display: flex; justify-content: space-around; width: 100%; font-size: 12px;">' +
                                    '<div>Usado: ' + disk.used.toFixed(1) + ' GB (' + pctUsed.toFixed(0) + '%)</div>' +
                                    '<div>Livre: ' + disk.free.toFixed(1) + ' GB (' + pctFree.toFixed(0) + '%)</div>' +
                                '</div></div>';
                            wrapper.innerHTML += fallbackHtml;
                            diskContainer.appendChild(wrapper);
                        }
                    });
                }
            }

            // c) Gráfico de Espaço em Disco por Perfil de Usuário
            var secUsers = document.getElementById('sec-users-all');
            var tableUsers = secUsers ? secUsers.querySelector('table') : null;
            if (tableUsers) {
                var userData = [];
                tableUsers.querySelectorAll('tr').forEach(function(row) {
                    if (row.querySelector('th')) return;
                    var cells = row.cells;
                    if (cells.length >= 10) {
                        var name = cells[0].innerText.trim();
                        var user = cells[1].innerText.trim();
                        var sizeText = cells[9].innerText.trim();
                        var sizeVal = parseFloat(sizeText.replace(/[^\d.,-]/g, '').replace(',', '.')) || 0;
                        if (sizeVal > 0) {
                            userData.push({ name: name || user, user: user, size: sizeVal });
                        }
                    }
                });

                var userContainer = document.getElementById('userDiskChartContainer');
                if (userContainer && userData.length > 0) {
                    userData.sort(function(a, b) { return b.size - a.size; });
                    if (typeof Chart !== 'undefined') {
                        userContainer.innerHTML = '<canvas id="userDiskChart" style="width: 100%; height: 260px;"></canvas>';
                        var ctxU = document.getElementById('userDiskChart').getContext('2d');
                        new Chart(ctxU, {
                            type: 'bar',
                            data: {
                                labels: userData.map(function(u) { return u.name; }),
                                datasets: [{
                                    label: 'Espaço em Disco (GB)',
                                    data: userData.map(function(u) { return u.size; }),
                                    backgroundColor: 'rgba(14, 116, 144, 0.2)',
                                    borderColor: '#0e7490',
                                    borderWidth: 2,
                                    borderRadius: 4
                                }]
                            },
                            options: {
                                responsive: true,
                                maintainAspectRatio: false,
                                plugins: { legend: { display: false } },
                                scales: {
                                    x: { grid: { display: false } },
                                    y: { beginAtZero: true, grace: '15%', ticks: { callback: function(v) { return v + ' GB'; } } }
                                }
                            }
                        });
                    } else {
                        var maxSize = userData[0] ? userData[0].size : 1;
                        var htmlU = '<div class="fallback-bars-list">';
                        userData.forEach(function(u) {
                            var pct = (u.size / maxSize) * 100;
                            htmlU += '<div class="bar-row"><div class="bar-label" title="' + u.user + '">' + u.name + '</div><div class="bar-track"><div class="bar-fill" style="width:' + pct + '%; background-color:#0e7490;"></div></div><div class="bar-value">' + u.size.toFixed(2) + ' GB</div></div>';
                        });
                        htmlU += '</div>';
                        userContainer.innerHTML = htmlU;
                    }
                } else if (userContainer) {
                    userContainer.innerHTML = '<p style="text-align: center; color: #64748b; font-style: italic;">Nenhum dado de espaço em disco de usuário ativo disponível.</p>';
                }
            }

            // d) Gráfico de Programas por Ano
            var secProg = document.getElementById('sec-programs');
            var tableProg = secProg ? secProg.querySelector('table') : null;
            if (tableProg) {
                var yearCounts = {};
                tableProg.querySelectorAll('tr').forEach(function(row) {
                    if (row.querySelector('th')) return;
                    var cells = row.cells;
                    if (cells.length >= 4) {
                        var dateText = cells[3].innerText.trim();
                        var year = dateText.substring(dateText.length - 4);
                        if (/^\d{4}$/.test(year)) {
                            yearCounts[year] = (yearCounts[year] || 0) + 1;
                        }
                    }
                });

                var years = Object.keys(yearCounts).sort();
                var progContainer = document.getElementById('programsChartContainer');
                if (progContainer && years.length > 0) {
                    if (typeof Chart !== 'undefined') {
                        progContainer.innerHTML = '<canvas id="programsChart" style="width: 100%; height: 260px;"></canvas>';
                        var ctxP = document.getElementById('programsChart').getContext('2d');
                        new Chart(ctxP, {
                            type: 'bar',
                            data: {
                                labels: years,
                                datasets: [{
                                    label: 'Programas Instalados',
                                    data: years.map(function(y) { return yearCounts[y]; }),
                                    backgroundColor: 'rgba(16, 185, 129, 0.2)',
                                    borderColor: '#10b8a5',
                                    borderWidth: 2,
                                    borderRadius: 4
                                }]
                            },
                            options: {
                                responsive: true,
                                maintainAspectRatio: false,
                                plugins: { legend: { display: false } },
                                scales: {
                                    x: { grid: { display: false } },
                                    y: { beginAtZero: true, grace: '15%', ticks: { stepSize: 1 } }
                                }
                            }
                        });
                    } else {
                        var maxProg = Math.max.apply(null, years.map(function(y) { return yearCounts[y]; })) || 1;
                        var htmlP = '<div class="fallback-bars-list">';
                        years.forEach(function(y) {
                            var pct = (yearCounts[y] / maxProg) * 100;
                            htmlP += '<div class="bar-row"><div class="bar-label">Ano ' + y + '</div><div class="bar-track"><div class="bar-fill" style="width:' + pct + '%; background-color:#10b8a5;"></div></div><div class="bar-value">' + yearCounts[y] + ' progs</div></div>';
                        });
                        htmlP += '</div>';
                        progContainer.innerHTML = htmlP;
                    }
                }
            }

            // e) Gráfico de Serviços (Tipo de Inicialização)
            var secServ = document.getElementById('sec-services');
            var tableServ = secServ ? secServ.querySelector('table') : null;
            if (tableServ) {
                var modeCounts = {};
                tableServ.querySelectorAll('tr').forEach(function(row) {
                    if (row.querySelector('th')) return;
                    var cells = row.cells;
                    if (cells.length >= 4) {
                        var mode = cells[3].innerText.trim();
                        if (mode) modeCounts[mode] = (modeCounts[mode] || 0) + 1;
                    }
                });

                var modes = Object.keys(modeCounts);
                var servContainer = document.getElementById('servicesChartContainer');
                if (servContainer && modes.length > 0) {
                    if (typeof Chart !== 'undefined') {
                        servContainer.innerHTML = '<canvas id="servicesChart" style="width: 100%; height: 220px;"></canvas>';
                        var ctxS = document.getElementById('servicesChart').getContext('2d');
                        new Chart(ctxS, {
                            type: 'pie',
                            data: {
                                labels: modes,
                                datasets: [{
                                    data: modes.map(function(m) { return modeCounts[m]; }),
                                    backgroundColor: ['#f59e0b', '#10b981', '#3b82f6', '#ef4444', '#8b5cf6'],
                                    borderWidth: 2,
                                    borderColor: '#ffffff'
                                }]
                            },
                            options: {
                                responsive: true,
                                maintainAspectRatio: false,
                                plugins: { legend: { position: 'bottom' } }
                            }
                        });
                    } else {
                        var totalServ = modes.reduce(function(acc, m) { return acc + modeCounts[m]; }, 0) || 1;
                        var htmlS = '<div class="fallback-bars-list">';
                        modes.forEach(function(m) {
                            var pct = (modeCounts[m] / totalServ) * 100;
                            htmlS += '<div class="bar-row"><div class="bar-label">' + m + '</div><div class="bar-track"><div class="bar-fill" style="width:' + pct + '%; background-color:#3b82f6;"></div></div><div class="bar-value">' + modeCounts[m] + ' (' + pct.toFixed(0) + '%)</div></div>';
                        });
                        htmlS += '</div>';
                        servContainer.innerHTML = htmlS;
                    }
                }
            }

            // f) Gráfico de Processos (Top Consumidores de RAM & Métricas)
            window.updateProcessesChart = function() {
                var secProc = document.getElementById('sec-processes');
                var tableProc = secProc ? secProc.querySelector('table') : null;
                if (!tableProc) return;

                var procData = [];
                tableProc.querySelectorAll('tr').forEach(function(row) {
                    if (row.querySelector('th') || row.style.display === 'none' || row.classList.contains('no-results-row')) return;
                    var cells = row.cells;
                    if (cells.length >= 6) {
                        var name = cells[1].innerText.trim();
                        var user = cells[2].innerText.trim();
                        var memoryText = cells[3].innerText.trim();
                        var memoryVal = parseFloat(memoryText.replace(/[^\d.,-]/g, '').replace(',', '.')) || 0;
                        var threads = parseInt(cells[4].innerText.trim(), 10) || 0;
                        var handles = parseInt(cells[5].innerText.trim(), 10) || 0;
                        procData.push({ name: name, user: user, memory: memoryVal, threads: threads, handles: handles });
                    }
                });

                var totalProcEl = document.getElementById('stat-total-proc');
                var totalThreadsEl = document.getElementById('stat-total-threads');
                var totalHandlesEl = document.getElementById('stat-total-handles');
                var chartContainer = document.getElementById('processesChartContainer');

                if (procData.length > 0) {
                    var totalProc = procData.length;
                    var totalThreads = procData.reduce(function(acc, p) { return acc + p.threads; }, 0);
                    var totalHandles = procData.reduce(function(acc, p) { return acc + p.handles; }, 0);

                    if (totalProcEl) totalProcEl.innerText = totalProc;
                    if (totalThreadsEl) totalThreadsEl.innerText = totalThreads.toLocaleString('pt-BR');
                    if (totalHandlesEl) totalHandlesEl.innerText = totalHandles.toLocaleString('pt-BR');

                    var topMemory = procData.slice().sort(function(a, b) { return b.memory - a.memory; }).slice(0, 10);
                    if (chartContainer) {
                        if (typeof Chart !== 'undefined') {
                            if (window.procMemoryChartObj) {
                                window.procMemoryChartObj.destroy();
                            }
                            chartContainer.innerHTML = '<canvas id="procMemoryChart" style="width: 100%; height: 260px;"></canvas>';
                            var ctxProc = document.getElementById('procMemoryChart').getContext('2d');
                            window.procMemoryChartObj = new Chart(ctxProc, {
                                type: 'bar',
                                data: {
                                    labels: topMemory.map(function(p) { return p.name; }),
                                    datasets: [{
                                        label: 'Uso de Memória (MB)',
                                        data: topMemory.map(function(p) { return p.memory; }),
                                        backgroundColor: 'rgba(37, 99, 235, 0.15)',
                                        borderColor: '#2563eb',
                                        borderWidth: 2,
                                        borderRadius: 4
                                    }]
                                },
                                options: {
                                    responsive: true,
                                    maintainAspectRatio: false,
                                    plugins: { legend: { display: false } },
                                    scales: {
                                        x: { grid: { display: false } },
                                        y: { beginAtZero: true, grace: '15%', ticks: { callback: function(v) { return v + ' MB'; } } }
                                    }
                                }
                            });
                        } else {
                            var maxMem = topMemory[0] ? topMemory[0].memory : 1;
                            var htmlProc = '<div class="fallback-bars-list">';
                            topMemory.forEach(function(p) {
                                var pct = (p.memory / maxMem) * 100;
                                htmlProc += '<div class="bar-row"><div class="bar-label" title="' + p.user + '">' + p.name + '</div><div class="bar-track"><div class="bar-fill" style="width:' + pct + '%;"></div></div><div class="bar-value">' + p.memory.toFixed(1) + ' MB</div></div>';
                            });
                            htmlProc += '</div>';
                            chartContainer.innerHTML = htmlProc;
                        }
                    }
                } else {
                    if (totalProcEl) totalProcEl.innerText = '0';
                    if (totalThreadsEl) totalThreadsEl.innerText = '0';
                    if (totalHandlesEl) totalHandlesEl.innerText = '0';
                    if (chartContainer) {
                        if (window.procMemoryChartObj) {
                            window.procMemoryChartObj.destroy();
                            window.procMemoryChartObj = null;
                        }
                        chartContainer.innerHTML = '<p style="text-align: center; color: #64748b; font-style: italic; margin-top: 100px;">Nenhum processo correspondente à busca.</p>';
                    }
                }
            };
            updateProcessesChart();

            // g) Gráfico de Drivers (Origem)
            var secDrv = document.getElementById('sec-drivers');
            var tableDrv = secDrv ? secDrv.querySelector('table') : null;
            if (tableDrv) {
                var msCount = 0;
                var otherCount = 0;
                tableDrv.querySelectorAll('tr').forEach(function(row) {
                    if (row.querySelector('th')) return;
                    var cells = row.cells;
                    if (cells.length >= 2) {
                        var mfr = cells[1].innerText.trim().toLowerCase();
                        if (mfr.indexOf('microsoft') !== -1) {
                            msCount++;
                        } else {
                            otherCount++;
                        }
                    }
                });

                var drvContainer = document.getElementById('driversChartContainer');
                if (drvContainer && (msCount > 0 || otherCount > 0)) {
                    if (typeof Chart !== 'undefined') {
                        drvContainer.innerHTML = '<canvas id="driversChart" style="width: 100%; height: 220px;"></canvas>';
                        var ctxDr = document.getElementById('driversChart').getContext('2d');
                        new Chart(ctxDr, {
                            type: 'pie',
                            data: {
                                labels: ['Microsoft', 'Terceiros'],
                                datasets: [{
                                    data: [msCount, otherCount],
                                    backgroundColor: ['#0284c7', '#f43f5e'],
                                    borderWidth: 2,
                                    borderColor: '#ffffff'
                                }]
                            },
                            options: {
                                responsive: true,
                                maintainAspectRatio: false,
                                plugins: { legend: { position: 'bottom' } }
                            }
                        });
                    } else {
                        var totalDrv = msCount + otherCount;
                        var pctMs = (msCount / totalDrv) * 100;
                        var pctOther = (otherCount / totalDrv) * 100;
                        var htmlDr = '<div class="fallback-bars-list">' +
                            '<div class="bar-row"><div class="bar-label">Microsoft</div><div class="bar-track"><div class="bar-fill" style="width:' + pctMs + '%; background-color:#0284c7;"></div></div><div class="bar-value">' + msCount + ' (' + pctMs.toFixed(0) + '%)</div></div>' +
                            '<div class="bar-row"><div class="bar-label">Terceiros</div><div class="bar-track"><div class="bar-fill" style="width:' + pctOther + '%; background-color:#f43f5e;"></div></div><div class="bar-value">' + otherCount + ' (' + pctOther.toFixed(0) + '%)</div></div>' +
                            '</div>';
                        drvContainer.innerHTML = htmlDr;
                    }
                }
            }
        });
    </script>
</body>
</html>
"@

    return $htmlContent
}
