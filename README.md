<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>NetCAD - Suite de Redes Estructuradas (TIA/EIA-568)</title>
    <!-- Tailwind CSS CDN -->
    <script src="https://cdn.tailwindcss.com"></script>
    <!-- Lucide Icons -->
    <script src="https://unpkg.com/lucide@latest"></script>
    <!-- SheetJS para manejo de Excel -->
    <script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>
    <!-- Google GenAI SDK (para la IA integrada) -->
    <script type="importmap">
      {
        "imports": {
          "@google/genai": "https://esm.run/@google/genai"
        }
      }
    </script>
    <style>
        @keyframes pulse-slow {
            0%, 100% { opacity: 1; }
            50% { opacity: 0.4; }
        }
        .animate-pulse-slow {
            animation: pulse-slow 2s cubic-bezier(0.4, 0, 0.6, 1) infinite;
        }
    </style>
</head>
<body class="bg-slate-900 text-slate-100 min-h-screen flex flex-col font-sans">

    <!-- BARRA SUPERIOR -->
    <header class="bg-slate-800 border-b border-slate-700 px-6 py-4 flex flex-wrap items-center justify-between shadow-md">
        <div class="flex items-center space-x-3">
            <div class="bg-blue-600 p-2.5 rounded-xl text-white shadow-lg shadow-blue-500/30 flex items-center justify-center">
                <i data-lucide="network" class="w-7 h-7"></i>
            </div>
            <div>
                <h1 class="text-xl font-black tracking-wide text-white flex items-center gap-2">
                    NetCAD <span class="text-xs bg-blue-500/20 text-blue-400 px-2 py-0.5 rounded-full border border-blue-500/30 font-semibold">INACAP v3.1</span>
                </h1>
                <p class="text-xs text-slate-400">Diseño CAD, Trazado TIA/EIA-568, Cubicación y Presupuesto CLP</p>
            </div>
        </div>

        <div class="flex items-center space-x-2 mt-3 md:mt-0 flex-wrap gap-2">
            <button onclick="downloadExcelTemplate()" class="bg-slate-700 hover:bg-slate-600 text-slate-200 text-xs font-semibold px-3 py-2 rounded-lg transition flex items-center gap-1.5 border border-slate-600 shadow">
                <i data-lucide="file-spreadsheet" class="w-4 h-4 text-emerald-400"></i> Plantilla Excel
            </button>
            <label class="bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-semibold px-3 py-2 rounded-lg transition flex items-center gap-1.5 cursor-pointer shadow">
                <i data-lucide="upload" class="w-4 h-4"></i> Importar Excel
                <input type="file" id="excelFileInput" accept=".xlsx, .xls, .csv" class="hidden" onchange="handleExcelImport(event)">
            </label>
            <button onclick="exportProjectJSON()" class="bg-slate-700 hover:bg-slate-600 text-slate-200 text-xs font-semibold px-3 py-2 rounded-lg transition flex items-center gap-1.5 border border-slate-600">
                <i data-lucide="save" class="w-4 h-4 text-blue-400"></i> Guardar Proyecto
            </button>
            <label class="bg-slate-700 hover:bg-slate-600 text-slate-200 text-xs font-semibold px-3 py-2 rounded-lg transition flex items-center gap-1.5 border border-slate-600 cursor-pointer">
                <i data-lucide="folder-open" class="w-4 h-4 text-amber-400"></i> Cargar JSON
                <input type="file" id="jsonFileInput" accept=".json" class="hidden" onchange="importProjectJSON(event)">
            </label>
            <button onclick="generateReadmeModal()" class="bg-purple-600 hover:bg-purple-500 text-white text-xs font-semibold px-3 py-2 rounded-lg transition flex items-center gap-1.5 shadow">
                <i data-lucide="code" class="w-4 h-4"></i> Generar README.md
            </button>
        </div>
    </header>

    <!-- NAVEGACIÓN DE PESTAÑAS -->
    <nav class="bg-slate-800/60 border-b border-slate-700/80 px-6 flex space-x-1 overflow-x-auto">
        <button onclick="switchTab('cad')" id="btn-tab-cad" class="tab-btn px-4 py-3 text-sm font-semibold border-b-2 border-blue-500 text-blue-400 flex items-center gap-2 transition">
            <i data-lucide="map" class="w-4 h-4"></i> 1. Plano CAD & Trazado (TIA-568)
        </button>
        <button onclick="switchTab('topo')" id="btn-tab-topo" class="tab-btn px-4 py-3 text-sm font-semibold border-b-2 border-transparent text-slate-400 hover:text-slate-200 flex items-center gap-2 transition">
            <i data-lucide="git-commit" class="w-4 h-4"></i> 2. Topología en Estrella
        </button>
        <button onclick="switchTab('cub')" id="btn-tab-cub" class="tab-btn px-4 py-3 text-sm font-semibold border-b-2 border-transparent text-slate-400 hover:text-slate-200 flex items-center gap-2 transition">
            <i data-lucide="box" class="w-4 h-4"></i> 3. Cubicación de Materiales
        </button>
        <button onclick="switchTab('pre')" id="btn-tab-pre" class="tab-btn px-4 py-3 text-sm font-semibold border-b-2 border-transparent text-slate-400 hover:text-slate-200 flex items-center gap-2 transition">
            <i data-lucide="dollar-sign" class="w-4 h-4"></i> 4. Presupuesto CLP & IVA (19%)
        </button>
        <button onclick="switchTab('gan')" id="btn-tab-gan" class="tab-btn px-4 py-3 text-sm font-semibold border-b-2 border-transparent text-slate-400 hover:text-slate-200 flex items-center gap-2 transition">
            <i data-lucide="calendar" class="w-4 h-4"></i> 5. Carta Gantt & Riesgos
        </button>
    </nav>

    <!-- CONTENIDO PRINCIPAL -->
    <main class="flex-1 p-6 max-w-7xl w-full mx-auto">

        <!-- PESTAÑA 1: CAD & TRAZADO -->
        <section id="tab-cad" class="tab-content space-y-6">
            <div class="grid grid-cols-1 lg:grid-cols-4 gap-6">
                <!-- Panel de Herramientas CAD -->
                <div class="bg-slate-800 border border-slate-700 rounded-2xl p-5 space-y-5 shadow-xl">
                    <h2 class="text-sm font-bold uppercase tracking-wider text-slate-300 flex items-center gap-2">
                        <i data-lucide="sliders" class="w-4 h-4 text-blue-400"></i> Herramientas de Plano
                    </h2>

                    <div class="space-y-3">
                        <label class="block text-xs font-medium text-slate-400">Subir Plano Arquitectónico (PNG/JPG)</label>
                        <input type="file" id="floorplanFile" accept="image/*" onchange="loadFloorplan(event)" class="w-full text-xs text-slate-400 file:mr-3 file:py-2 file:px-3 file:rounded-lg file:border-0 file:text-xs file:font-semibold file:bg-blue-600 file:text-white hover:file:bg-blue-500 cursor-pointer">
                    </div>

                    <div class="border-t border-slate-700 pt-4 space-y-3">
                        <label class="block text-xs font-medium text-slate-400">Calibración Métrica (Escala)</label>
                        <div class="flex gap-2">
                            <input type="number" id="scaleMeters" value="10" placeholder="Metros reales" class="bg-slate-900 border border-slate-700 rounded-lg px-3 py-1.5 text-sm w-full focus:outline-none focus:border-blue-500">
                            <button onclick="setCalibrationMode()" id="btnCalibrate" class="bg-amber-600 hover:bg-amber-500 text-white px-3 py-1.5 rounded-lg text-xs font-semibold transition whitespace-nowrap shadow">Calibrar</button>
                        </div>
                        <p id="scaleStatus" class="text-xs text-emerald-400 font-medium">Escala actual: 1 metro = 25 píxeles</p>
                    </div>

                    <div class="border-t border-slate-700 pt-4 space-y-3">
                        <label class="block text-xs font-medium text-slate-400">Herramienta de Ubicación</label>
                        <div class="grid grid-cols-2 gap-2">
                            <button onclick="setTool('rack')" id="tool-rack" class="tool-btn bg-blue-600 text-white p-2.5 rounded-xl text-xs font-semibold flex items-center justify-center gap-1.5 shadow transition">
                                <i data-lucide="server" class="w-4 h-4"></i> Rack MDF
                            </button>
                            <button onclick="setTool('outlet')" id="tool-outlet" class="tool-btn bg-slate-700 hover:bg-slate-600 text-slate-300 p-2.5 rounded-xl text-xs font-semibold flex items-center justify-center gap-1.5 transition">
                                <i data-lucide="wifi" class="w-4 h-4"></i> Salida / AP
                            </button>
                        </div>
                    </div>

                    <div class="border-t border-slate-700 pt-4 space-y-2">
                        <button onclick="clearCanvasElements()" class="w-full bg-rose-600/20 hover:bg-rose-600/30 text-rose-400 border border-rose-500/30 text-xs font-semibold py-2 rounded-xl transition flex items-center justify-center gap-1.5">
                            <i data-lucide="trash-2" class="w-4 h-4"></i> Limpiar Elementos
                        </button>
                        <button onclick="exportCanvasPNG()" class="w-full bg-slate-700 hover:bg-slate-600 text-slate-200 text-xs font-semibold py-2 rounded-xl transition flex items-center justify-center gap-1.5 border border-slate-600">
                            <i data-lucide="download" class="w-4 h-4 text-emerald-400"></i> Descargar Imagen del Plano
                        </button>
                    </div>
                </div>

                <!-- Lienzo Interactivo CAD -->
                <div class="lg:col-span-3 bg-slate-800 border border-slate-700 rounded-2xl p-5 flex flex-col shadow-xl relative overflow-hidden">
                    <div class="flex items-center justify-between mb-3">
                        <h3 class="text-sm font-bold text-slate-200 flex items-center gap-2">
                            <i data-lucide="maximize" class="w-4 h-4 text-blue-400"></i> Lienzo CAD (Haga clic para ubicar nodos y trazar en estrella)
                        </h3>
                        <span id="canvasStats" class="text-xs bg-slate-700 px-3 py-1 rounded-full text-slate-300 border border-slate-600">0 Salidas conectadas | Distancia máx: 0m</span>
                    </div>

                    <div id="canvasContainer" class="relative w-full h-[520px] bg-slate-950 rounded-xl border border-slate-700 overflow-hidden cursor-crosshair flex items-center justify-center">
                        <canvas id="cadCanvas" class="absolute inset-0 w-full h-full"></canvas>
                        <div id="canvasPlaceholder" class="absolute inset-0 flex flex-col items-center justify-center text-center p-6 text-slate-500 pointer-events-none">
                            <i data-lucide="map-pin" class="w-12 h-12 mb-2 opacity-40"></i>
                            <p class="text-sm font-semibold text-slate-400">Plano listo para trazar</p>
                            <p class="text-xs">Sube un plano arquitectónico o usa el fondo cuadriculado por defecto. Coloca el Rack y las Salidas UTP.</p>
                        </div>
                    </div>
                </div>
            </div>

            <!-- Tabla de Enlaces y Verificación Normativa TIA-568 -->
            <div class="bg-slate-800 border border-slate-700 rounded-2xl p-6 shadow-xl space-y-4">
                <div class="flex items-center justify-between flex-wrap gap-2">
                    <div>
                        <h3 class="text-base font-bold text-white flex items-center gap-2">
                            <i data-lucide="check-circle-2" class="w-5 h-5 text-emerald-400"></i> Registro de Enlaces Permanentes (Norma TIA/EIA-568: Máximo 90m)
                        </h3>
                        <p class="text-xs text-slate-400">Cálculo automático de distancia horizontal + caída vertical al Rack principal.</p>
                    </div>
                </div>
                <div class="overflow-x-auto">
                    <table class="w-full text-left border-collapse text-xs">
                        <thead>
                            <tr class="bg-slate-900/80 text-slate-400 uppercase tracking-wider border-b border-slate-700">
                                <th class="p-3">ID Salida / AP</th>
                                <th class="p-3">Destino / Ubicación</th>
                                <th class="p-3">Distancia Horizontal</th>
                                <th class="p-3">Holgura / Vertical</th>
                                <th class="p-3">Distancia Total</th>
                                <th class="p-3">Estado Normativo (≤ 90m)</th>
                            </tr>
                        </thead>
                        <tbody id="linksTableBody" class="divide-y divide-slate-700/60 text-slate-200">
                            <tr><td colspan="6" class="p-4 text-center text-slate-500">Coloque salidas en el plano CAD para generar el reporte de enlaces.</td></tr>
                        </tbody>
                    </table>
                </div>
            </div>
        </section>

        <!-- PESTAÑA 2: TOPOLOGÍA EN ESTRELLA -->
        <section id="tab-topo" class="tab-content hidden space-y-6">
            <div class="bg-slate-800 border border-slate-700 rounded-2xl p-6 shadow-xl space-y-5">
                <div class="flex items-center justify-between flex-wrap gap-2">
                    <div>
                        <h3 class="text-lg font-bold text-white flex items-center gap-2">
                            <i data-lucide="git-commit" class="w-5 h-5 text-blue-400"></i> Diagrama Esquemático en Topología Estrella
                        </h3>
                        <p class="text-xs text-slate-400">Representación lógica centralizada desde el Patch Panel / Switch MDF hacia cada punto de red.</p>
                    </div>
                    <button onclick="renderTopologyDiagram()" class="bg-blue-600 hover:bg-blue-500 text-white text-xs font-semibold px-4 py-2 rounded-xl transition flex items-center gap-1.5 shadow">
                        <i data-lucide="refresh-cw" class="w-4 h-4"></i> Actualizar Diagrama
                    </button>
                </div>
                <div class="w-full h-[450px] bg-slate-950 rounded-xl border border-slate-700 flex items-center justify-center relative overflow-hidden p-4">
                    <canvas id="topoCanvas" class="w-full h-full"></canvas>
                </div>
            </div>
        </section>

        <!-- PESTAÑA 3: CUBICACIÓN -->
        <section id="tab-cub" class="tab-content hidden space-y-6">
            <div class="bg-slate-800 border border-slate-700 rounded-2xl p-6 shadow-xl space-y-6">
                <div class="flex items-center justify-between flex-wrap gap-4">
                    <div>
                        <h3 class="text-lg font-bold text-white flex items-center gap-2">
                            <i data-lucide="box" class="w-5 h-5 text-emerald-400"></i> Cubicación Automática de Materiales (Cat 6 / Cat 6A)
                        </h3>
                        <p class="text-xs text-slate-400">Calculado en base a los enlaces trazados en el plano más 10% de holgura y merma reglamentaria.</p>
                    </div>
                    <button onclick="calculateCubicacion()" class="bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-semibold px-4 py-2 rounded-xl transition flex items-center gap-1.5 shadow">
                        <i data-lucide="calculator" class="w-4 h-4"></i> Recalcular Cubicación
                    </button>
                </div>

                <div class="grid grid-cols-1 md:grid-cols-3 gap-4" id="cubicationCards">
                    <!-- Se llena por JS -->
                </div>

                <div class="overflow-x-auto pt-2">
                    <table class="w-full text-left border-collapse text-xs">
                        <thead>
                            <tr class="bg-slate-900/80 text-slate-400 uppercase tracking-wider border-b border-slate-700">
                                <th class="p-3">Ítem / Material</th>
                                <th class="p-3">Especificación técnica</th>
                                <th class="p-3">Cantidad Total</th>
                                <th class="p-3">Unidad de Medida</th>
                                <th class="p-3">Observación / Merma</th>
                            </tr>
                        </thead>
                        <tbody id="cubicationTableBody" class="divide-y divide-slate-700/60 text-slate-200">
                            <!-- Se llena por JS -->
                        </tbody>
                    </table>
                </div>
            </div>
        </section>

        <!-- PESTAÑA 4: PRESUPUESTO CLP & IVA -->
        <section id="tab-pre" class="tab-content hidden space-y-6">
            <div class="bg-slate-800 border border-slate-700 rounded-2xl p-6 shadow-xl space-y-6">
                <div class="flex items-center justify-between flex-wrap gap-4">
                    <div>
                        <h3 class="text-lg font-bold text-white flex items-center gap-2">
                            <i data-lucide="dollar-sign" class="w-5 h-5 text-amber-400"></i> Presupuesto del Proyecto en Pesos Chilenos (CLP) con IVA (19%)
                        </h3>
                        <p class="text-xs text-slate-400">Costos unitarios de mercado chileno (Sodimac, PC Factory, Mayoristas de Redes).</p>
                    </div>
                    <div class="flex gap-2">
                        <button onclick="addBudgetRow()" class="bg-blue-600 hover:bg-blue-500 text-white text-xs font-semibold px-3 py-2 rounded-xl transition flex items-center gap-1.5 shadow">
                            <i data-lucide="plus" class="w-4 h-4"></i> Agregar Ítem
                        </button>
                        <button onclick="exportToExcel()" class="bg-emerald-600 hover:bg-emerald-500 text-white text-xs font-semibold px-3 py-2 rounded-xl transition flex items-center gap-1.5 shadow">
                            <i data-lucide="file-spreadsheet" class="w-4 h-4"></i> Exportar a Excel
                        </button>
                    </div>
                </div>

                <div class="overflow-x-auto">
                    <table class="w-full text-left border-collapse text-xs">
                        <thead>
                            <tr class="bg-slate-900/80 text-slate-400 uppercase tracking-wider border-b border-slate-700">
                                <th class="p-3">Descripción / Insumo</th>
                                <th class="p-3">Categoría</th>
                                <th class="p-3">Cantidad</th>
                                <th class="p-3">P. Unitario (CLP)</th>
                                <th class="p-3">Total (CLP)</th>
                                <th class="p-3">Proveedor / Fuente</th>
                                <th class="p-3 text-center">Acción</th>
                            </tr>
                        </thead>
                        <tbody id="budgetTableBody" class="divide-y divide-slate-700/60 text-slate-200">
                            <!-- Se llena por JS -->
                        </tbody>
                    </table>
                </div>

                <!-- Resumen Financiero -->
                <div class="bg-slate-900/90 border border-slate-700 rounded-2xl p-5 flex flex-col md:flex-row justify-between items-center gap-4">
                    <div class="text-slate-300 text-xs space-y-1">
                        <p>🔹 **Mano de obra estimada**: Incluye certificación de puntos Fluke y etiquetado TIA-606-B.</p>
                        <p>🔹 **Tributación**: Impuesto al Valor Agregado (IVA) aplicado al 19% según legislación chilena.</p>
                    </div>
                    <div class="text-right space-y-1 min-w-[260px]">
                        <div class="flex justify-between text-xs text-slate-400">
                            <span>Neto Total:</span>
                            <span id="subtotalCLP" class="font-semibold text-slate-200">$0 CLP</span>
                        </div>
                        <div class="flex justify-between text-xs text-slate-400">
                            <span>IVA (19%):</span>
                            <span id="ivaCLP" class="font-semibold text-slate-200">$0 CLP</span>
                        </div>
                        <div class="flex justify-between text-base font-black text-emerald-400 border-t border-slate-700 pt-2">
                            <span>TOTAL PROYECTO:</span>
                            <span id="totalCLP">$0 CLP</span>
                        </div>
                    </div>
                </div>
            </div>
        </section>

        <!-- PESTAÑA 5: CARTA GANTT Y RIESGOS -->
        <section id="tab-gan" class="tab-content hidden space-y-6">
            <div class="grid grid-cols-1 lg:grid-cols-2 gap-6">
                <!-- Carta Gantt -->
                <div class="bg-slate-800 border border-slate-700 rounded-2xl p-6 shadow-xl space-y-4">
                    <h3 class="text-base font-bold text-white flex items-center gap-2">
                        <i data-lucide="calendar" class="w-5 h-5 text-blue-400"></i> Planificación y Carta Gantt (Instalación)
                    </h3>
                    <p class="text-xs text-slate-400">Cronograma de tareas para la ejecución del proyecto de red estructurada.</p>
                    
                    <div class="space-y-3 pt-2">
                        <div class="bg-slate-900 p-3 rounded-xl border border-slate-700 text-xs space-y-1">
                            <div class="flex justify-between font-bold text-slate-200"><span>1. Replanteo y Canalización</span> <span class="text-blue-400">Días 1-2</span></div>
                            <p class="text-slate-400">Responsable: Instalador Senior | TIA-569-D.</p>
                        </div>
                        <div class="bg-slate-900 p-3 rounded-xl border border-slate-700 text-xs space-y-1">
                            <div class="flex justify-between font-bold text-slate-200"><span>2. Tendido de Cable UTP Cat 6</span> <span class="text-blue-400">Días 3-4</span></div>
                            <p class="text-slate-400">Responsable: Ayudante Técnico | Sin superar radio de curvatura.</p>
                        </div>
                        <div class="bg-slate-900 p-3 rounded-xl border border-slate-700 text-xs space-y-1">
                            <div class="flex justify-between font-bold text-slate-200"><span>3. Terminación de Jacks y Patch Panel</span> <span class="text-blue-400">Días 5</span></div>
                            <p class="text-slate-400">Responsable: Técnico Certificador | TIA/EIA-568-B.</p>
                        </div>
                        <div class="bg-slate-900 p-3 rounded-xl border border-slate-700 text-xs space-y-1">
                            <div class="flex justify-between font-bold text-slate-200"><span>4. Certificación Fluke & Entrega</span> <span class="text-blue-400">Días 6</span></div>
                            <p class="text-slate-400">Responsable: Jefe de Proyecto | Documentación y entrega INACAP.</p>
                        </div>
                    </div>
                </div>

                <!-- Matriz de Riesgos -->
                <div class="bg-slate-800 border border-slate-700 rounded-2xl p-6 shadow-xl space-y-4">
                    <h3 class="text-base font-bold text-white flex items-center gap-2">
                        <i data-lucide="shield-alert" class="w-5 h-5 text-amber-400"></i> Matriz de Riesgos (Probabilidad × Impacto)
                    </h3>
                    <p class="text-xs text-slate-400">Evaluación de riesgos técnicos, de plazo y de costo.</p>

                    <div class="overflow-x-auto pt-2">
                        <table class="w-full text-left border-collapse text-xs">
                            <thead>
                                <tr class="bg-slate-900/80 text-slate-400 uppercase tracking-wider border-b border-slate-700">
                                    <th class="p-3">Riesgo Identificado</th>
                                    <th class="p-3">Prob (1-3)</th>
                                    <th class="p-3">Imp (1-3)</th>
                                    <th class="p-3">P × I</th>
                                    <th class="p-3">Medida Mitigación</th>
                                </tr>
                            </thead>
                            <tbody class="divide-y divide-slate-700/60 text-slate-200">
                                <tr>
                                    <td class="p-3 font-medium">Enlace > 90 metros</td>
                                    <td class="p-3">2</td>
                                    <td class="p-3">3</td>
                                    <td class="p-3 text-amber-400 font-bold">6 (Medio)</td>
                                    <td class="p-3 text-slate-400">Reubicar Rack IDF intermedio.</td>
                                </tr>
                                <tr>
                                    <td class="p-3 font-medium">Daño de cable por tracción</td>
                                    <td class="p-3">1</td>
                                    <td class="p-3">3</td>
                                    <td class="p-3 text-emerald-400 font-bold">3 (Bajo)</td>
                                    <td class="p-3 text-slate-400">Uso de gel lubricante y tensiómetro.</td>
                                </tr>
                                <tr>
                                    <td class="p-3 font-medium">Retraso en despacho de materiales</td>
                                    <td class="p-3">2</td>
                                    <td class="p-3">2</td>
                                    <td class="p-3 text-emerald-400 font-bold">4 (Bajo)</td>
                                    <td class="p-3 text-slate-400">Stock de seguridad del 10% en bodega.</td>
                                </tr>
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </section>

    </main>

    <!-- CHAT FLOTANTE DE IA (GEMINI) -->
    <div id="aiChatWidget" class="fixed bottom-6 right-6 z-50 flex flex-col items-end">
        <!-- Ventana de Chat -->
        <div id="chatWindow" class="hidden w-96 h-[500px] bg-slate-800 border border-slate-700 rounded-2xl shadow-2xl flex flex-col mb-3 overflow-hidden">
            <!-- Header del Chat -->
            <div class="bg-slate-900 px-4 py-3 border-b border-slate-700 flex items-center justify-between">
                <div class="flex items-center space-x-2">
                    <div class="bg-blue-600 p-1.5 rounded-lg text-white">
                        <i data-lucide="bot" class="w-4 h-4"></i>
                    </div>
                    <div>
                        <h3 class="text-xs font-bold text-white">NetCAD AI (Asistente INACAP)</h3>
                        <p class="text-[10px] text-emerald-400 flex items-center gap-1">● Conectado a Gemini</p>
                    </div>
                </div>
                <div class="flex items-center space-x-1">
                    <button onclick="toggleApiKeyModal()" class="text-slate-400 hover:text-white p-1 rounded-lg transition" title="Configurar API Key">
                        <i data-lucide="settings" class="w-4 h-4"></i>
                    </button>
                    <button onclick="toggleChat()" class="text-slate-400 hover:text-white p-1 rounded-lg transition">
                        <i data-lucide="x" class="w-4 h-4"></i>
                    </button>
                </div>
            </div>

            <!-- Botones de Ayuda Rápida e IA -->
            <div class="bg-slate-900/50 px-3 py-2 border-b border-slate-700/60 flex gap-1.5 overflow-x-auto text-[11px]">
                <button onclick="aiAnalyzePlan()" class="bg-blue-600/20 hover:bg-blue-600/30 text-blue-400 border border-blue-500/30 px-2 py-1 rounded-lg whitespace-nowrap transition font-medium">💡 Analizar mi plano</button>
                <button onclick="aiSimulateDefense()" class="bg-purple-600/20 hover:bg-purple-600/30 text-purple-400 border border-purple-500/30 px-2 py-1 rounded-lg whitespace-nowrap transition font-medium">🎓 Simular Defensa</button>
                <button onclick="aiAskNorma()" class="bg-emerald-600/20 hover:bg-emerald-600/30 text-emerald-400 border border-emerald-500/30 px-2 py-1 rounded-lg whitespace-nowrap transition font-medium">📏 Norma TIA-568</button>
            </div>

            <!-- Mensajes -->
            <div id="chatMessages" class="flex-1 p-4 overflow-y-auto space-y-3 text-xs text-slate-200">
                <div class="bg-slate-900/80 p-3 rounded-xl border border-slate-700 space-y-1">
                    <p class="font-bold text-blue-400">¡Hola, Nicolas!</p>
                    <p>Soy tu asistente experto en redes estructuradas para tu proyecto de INACAP. Haz clic en los botones de arriba o pregúntame lo que necesites sobre cubicación, la norma de los 90 metros o la defensa oral.</p>
                </div>
            </div>

            <!-- Input de Chat -->
            <div class="p-3 bg-slate-900 border-t border-slate-700 flex gap-2">
                <input type="text" id="chatInput" placeholder="Escribe tu consulta técnica..." onkeypress="handleChatKey(event)" class="bg-slate-800 border border-slate-700 rounded-xl px-3 py-2 text-xs flex-1 focus:outline-none focus:border-blue-500 text-slate-200">
                <button onclick="sendChatMessage()" class="bg-blue-600 hover:bg-blue-500 text-white p-2 rounded-xl transition shadow">
                    <i data-lucide="send" class="w-4 h-4"></i>
                </button>
            </div>
        </div>

        <!-- Botón Flotante -->
        <button onclick="toggleChat()" class="bg-blue-600 hover:bg-blue-500 text-white p-4 rounded-full shadow-2xl transition flex items-center justify-center gap-2 group">
            <i data-lucide="message-square" class="w-6 h-6"></i>
            <span class="max-w-0 overflow-hidden group-hover:max-w-xs transition-all duration-300 ease-in-out whitespace-nowrap text-xs font-semibold">Asistente IA</span>
        </button>
    </div>

    <!-- MODAL DE CONFIGURACIÓN DE API KEY -->
    <div id="apiKeyModal" class="fixed inset-0 bg-slate-950/80 z-50 hidden flex items-center justify-center p-4">
        <div class="bg-slate-800 border border-slate-700 rounded-2xl p-6 max-w-md w-full space-y-4 shadow-2xl">
            <h3 class="text-base font-bold text-white flex items-center gap-2">
                <i data-lucide="key" class="w-5 h-5 text-blue-400"></i> Configurar Clave API de Gemini
            </h3>
            <p class="text-xs text-slate-400">Para activar el asistente inteligente dentro de tu página, ingresa tu clave de API de Google AI Studio:</p>
            <input type="password" id="apiKeyInput" placeholder="AIzaSy..." class="w-full bg-slate-900 border border-slate-700 rounded-xl px-3 py-2 text-xs text-slate-200 focus:outline-none focus:border-blue-500">
            <div class="flex justify-end gap-2 pt-2">
                <button onclick="toggleApiKeyModal()" class="bg-slate-700 hover:bg-slate-600 text-slate-300 text-xs font-semibold px-4 py-2 rounded-xl transition">Cancelar</button>
                <button onclick="saveApiKey()" class="bg-blue-600 hover:bg-blue-500 text-white text-xs font-semibold px-4 py-2 rounded-xl transition shadow">Guardar Clave</button>
            </div>
        </div>
    </div>

    <!-- MODAL DE README.md -->
    <div id="readmeModal" class="fixed inset-0 bg-slate-950/80 z-50 hidden flex items-center justify-center p-4">
        <div class="bg-slate-800 border border-slate-700 rounded-2xl p-6 max-w-2xl w-full space-y-4 shadow-2xl">
            <h3 class="text-base font-bold text-white flex items-center gap-2">
                <i data-lucide="code" class="w-5 h-5 text-purple-400"></i> Código Markdown para tu Repositorio GitHub (`README.md`)
            </h3>
            <p class="text-xs text-slate-400">Copia este texto y pégalo en el archivo `README.md` de tu repositorio grupal para cumplir con la pauta de INACAP:</p>
            <textarea id="readmeTextarea" readonly class="w-full h-64 bg-slate-900 border border-slate-700 rounded-xl p-3 text-xs font-mono text-slate-300 focus:outline-none"></textarea>
            <div class="flex justify-end gap-2 pt-2">
                <button onclick="copyReadme()" class="bg-purple-600 hover:bg-purple-500 text-white text-xs font-semibold px-4 py-2 rounded-xl transition flex items-center gap-1.5 shadow">
                    <i data-lucide="copy" class="w-4 h-4"></i> Copiar al Portapapeles
                </button>
                <button onclick="closeReadmeModal()" class="bg-slate-700 hover:bg-slate-600 text-slate-300 text-xs font-semibold px-4 py-2 rounded-xl transition">Cerrar</button>
            </div>
        </div>
    </div>

    <!-- SCRIPTS DE LA APLICACIÓN -->
    <script>
        lucide.createIcons();

        let activeTab = 'cad';
        let currentTool = 'rack';
        let isCalibrating = false;
        let scalePixelsPerMeter = 25;
        let floorplanImg = null;
        let elements = [
            { id: 'MDF-1', type: 'rack', x: 100, y: 150, name: 'Rack Principal MDF' },
            { id: 'OUT-1', type: 'outlet', x: 300, y: 200, name: 'Oficina 01 (Doble)' },
            { id: 'OUT-2', type: 'outlet', x: 450, y: 320, name: 'Oficina 02 (Simple)' },
            { id: 'AP-1', type: 'outlet', x: 350, y: 100, name: 'Access Point Wi-Fi 6' }
        ];

        let budgetItems = [
            { desc: 'Cable UTP Cat 6 (Bobina 305m)', cat: 'Materiales', qty: 2, unitPrice: 125000, provider: 'Sodimac Constructor' },
            { desc: 'Conectores RJ45 Cat 6 (Caja 100 un)', cat: 'Materiales', qty: 1, unitPrice: 28900, provider: 'PC Factory' },
            { desc: 'Patch Panel 24 Puertos Cat 6', cat: 'Materiales', qty: 1, unitPrice: 45000, provider: 'Mercado Libre' },
            { desc: 'Faceplate Doble + Jacks', cat: 'Materiales', qty: 12, unitPrice: 4500, provider: 'Electrónica Local' },
            { desc: 'Canaleta PVC 40x25 mm (barra 2m)', cat: 'Canalización', qty: 30, unitPrice: 3200, provider: 'Sodimac' },
            { desc: 'Mano de Obra & Certificación Fluke', cat: 'Servicios', qty: 1, unitPrice: 180000, provider: 'Técnico Certificador' }
        ];

        let geminiApiKey = localStorage.getItem('gemini_api_key') || '';

        function switchTab(tabId) {
            activeTab = tabId;
            document.querySelectorAll('.tab-content').forEach(el => el.classList.add('hidden'));
            document.querySelectorAll('.tab-btn').forEach(el => {
                el.classList.remove('border-blue-500', 'text-blue-400');
                el.classList.add('border-transparent', 'text-slate-400');
            });

            document.getElementById(`tab-${tabId}`).classList.remove('hidden');
            const btn = document.getElementById(`btn-tab-${tabId}`);
            btn.classList.remove('border-transparent', 'text-slate-400');
            btn.classList.add('border-blue-500', 'text-blue-400');

            if (tabId === 'cad') { resizeCanvas(); drawCad(); }
            if (tabId === 'topo') { renderTopologyDiagram(); }
            if (tabId === 'cub') { calculateCubicacion(); }
            if (tabId === 'pre') { renderBudgetTable(); }
        }

        function setTool(tool) {
            currentTool = tool;
            document.querySelectorAll('.tool-btn').forEach(el => {
                el.classList.remove('bg-blue-600', 'text-white');
                el.classList.add('bg-slate-700', 'text-slate-300');
            });
            const activeBtn = event.currentTarget;
            activeBtn.classList.remove('bg-slate-700', 'text-slate-300');
            activeBtn.classList.add('bg-blue-600', 'text-white');
        }

        function setCalibrationMode() {
            isCalibrating = true;
            alert("Modo calibración activado.");
        }

        function loadFloorplan(event) {
            const file = event.target.files[0];
            if (file) {
                const reader = new FileReader();
                reader.onload = function(e) {
                    floorplanImg = new Image();
                    floorplanImg.onload = function() {
                        document.getElementById('canvasPlaceholder').style.display = 'none';
                        drawCad();
                    }
                    floorplanImg.src = e.target.result;
                }
                reader.readAsDataURL(file);
            }
        }

        const canvas = document.getElementById('cadCanvas');
        const ctx = canvas.getContext('2d');

        function resizeCanvas() {
            const container = canvas.parentElement;
            canvas.width = container.clientWidth;
            canvas.height = container.clientHeight;
            drawCad();
        }

        window.addEventListener('resize', () => { if(activeTab === 'cad') resizeCanvas(); });

        function drawCad() {
            if (!ctx) return;
            ctx.clearRect(0, 0, canvas.width, canvas.height);

            if (floorplanImg) {
                ctx.drawImage(floorplanImg, 0, 0, canvas.width, canvas.height);
            } else {
                ctx.strokeStyle = '#1e293b';
                ctx.lineWidth = 1;
                const gridSize = 30;
                for (let x = 0; x < canvas.width; x += gridSize) {
                    ctx.beginPath(); ctx.moveTo(x, 0); ctx.lineTo(x, canvas.height); ctx.stroke();
                }
                for (let y = 0; y < canvas.height; y += gridSize) {
                    ctx.beginPath(); ctx.moveTo(0, y); ctx.lineTo(canvas.width, y); ctx.stroke();
                }
            }

            const mdf = elements.find(el => el.type === 'rack') || { x: 80, y: 80 };
            let maxDist = 0;
            let connectedCount = 0;

            elements.forEach(el => {
                if (el.type === 'outlet') {
                    connectedCount++;
                    const dx = el.x - mdf.x;
                    const dy = el.y - mdf.y;
                    const pixelDist = Math.sqrt(dx*dx + dy*dy);
                    const realDist = (pixelDist / scalePixelsPerMeter) + 5;

                    if (realDist > maxDist) maxDist = realDist;

                    ctx.beginPath();
                    ctx.moveTo(mdf.x, mdf.y);
                    ctx.lineTo(el.x, el.y);
                    ctx.strokeStyle = realDist > 90 ? '#ef4444' : '#3b82f6';
                    ctx.lineWidth = realDist > 90 ? 3 : 2;
                    if (realDist > 90) ctx.setLineDash([4, 4]);
                    else ctx.setLineDash([]);
                    ctx.stroke();
                    ctx.setLineDash([]);

                    ctx.fillStyle = realDist > 90 ? '#f87171' : '#94a3b8';
                    ctx.font = '10px sans-serif';
                    ctx.fillText(`${realDist.toFixed(1)}m`, (mdf.x + el.x)/2 + 5, (mdf.y + el.y)/2 - 5);
                }
            });

            elements.forEach(el => {
                ctx.beginPath();
                ctx.arc(el.x, el.y, el.type === 'rack' ? 16 : 12, 0, Math.PI * 2);
                ctx.fillStyle = el.type === 'rack' ? '#2563eb' : '#10b981';
                ctx.fill();
                ctx.lineWidth = 3;
                ctx.strokeStyle = '#ffffff';
                ctx.stroke();

                ctx.fillStyle = '#ffffff';
                ctx.font = 'bold 11px sans-serif';
                ctx.fillText(el.id, el.x - 12, el.y - 20);
            });

            document.getElementById('canvasStats').innerText = `${connectedCount} Salidas conectadas | Distancia máx: ${maxDist.toFixed(1)}m`;
            updateLinksTable(mdf);
        }

        canvas.addEventListener('click', (e) => {
            const rect = canvas.getBoundingClientRect();
            const x = e.clientX - rect.left;
            const y = e.clientY - rect.top;

            const idNum = elements.length + 1;
            const newId = currentTool === 'rack' ? `MDF-${idNum}` : `OUT-${idNum}`;
            const newName = currentTool === 'rack' ? 'Rack Principal MDF' : `Puesto de Red ${idNum}`;

            elements.push({ id: newId, type: currentTool, x: x, y: y, name: newName });
            drawCad();
        });

        function clearCanvasElements() {
            elements = [{ id: 'MDF-1', type: 'rack', x: 100, y: 150, name: 'Rack Principal MDF' }];
            drawCad();
        }

        function exportCanvasPNG() {
            const link = document.createElement('a');
            link.download = 'plano-red-netcad.png';
            link.href = canvas.toDataURL();
            link.click();
        }

        function updateLinksTable(mdf) {
            const tbody = document.getElementById('linksTableBody');
            tbody.innerHTML = '';

            const outlets = elements.filter(el => el.type === 'outlet');
            if (outlets.length === 0) {
                tbody.innerHTML = `<tr><td colspan="6" class="p-4 text-center text-slate-500">No hay salidas de red ubicadas en el plano.</td></tr>`;
                return;
            }

            outlets.forEach(el => {
                const dx = el.x - mdf.x;
                const dy = el.y - mdf.y;
                const pixelDist = Math.sqrt(dx*dx + dy*dy);
                const horiz = pixelDist / scalePixelsPerMeter;
                const vertical = 5.0;
                const total = horiz + vertical;

                const isCompliant = total <= 90;

                const tr = document.createElement('tr');
                tr.className = "hover:bg-slate-900/40 transition";
                tr.innerHTML = `
                    <td class="p-3 font-bold text-blue-400">${el.id}</td>
                    <td class="p-3">${el.name}</td>
                    <td class="p-3">${horiz.toFixed(1)} m</td>
                    <td class="p-3">${vertical.toFixed(1)} m (Holgura)</td>
                    <td class="p-3 font-semibold">${total.toFixed(1)} m</td>
                    <td class="p-3">
                        <span class="px-2.5 py-1 rounded-full text-[10px] font-bold ${isCompliant ? 'bg-emerald-500/20 text-emerald-400 border border-emerald-500/30' : 'bg-rose-500/20 text-rose-400 border border-rose-500/30 animate-pulse'}">
                            ${isCompliant ? 'CUMPLE (≤ 90m)' : 'EXCEDE (> 90m)'}
                        </span>
                    </td>
                `;
                tbody.appendChild(tr);
            });
        }

        function renderTopologyDiagram() {
            const topoCanvas = document.getElementById('topoCanvas');
            const tCtx = topoCanvas.getContext('2d');
            topoCanvas.width = topoCanvas.parentElement.clientWidth;
            topoCanvas.height = topoCanvas.parentElement.clientHeight;

            tCtx.clearRect(0, 0, topoCanvas.width, topoCanvas.height);

            const centerX = topoCanvas.width / 2;
            const centerY = topoCanvas.height / 2;

            tCtx.beginPath();
            tCtx.arc(centerX, centerY, 35, 0, Math.PI * 2);
            tCtx.fillStyle = '#2563eb';
            tCtx.fill();
            tCtx.lineWidth = 4;
            tCtx.strokeStyle = '#ffffff';
            tCtx.stroke();

            tCtx.fillStyle = '#ffffff';
            tCtx.font = 'bold 12px sans-serif';
            tCtx.textAlign = 'center';
            tCtx.fillText('SWITCH / MDF', centerX, centerY - 5);

            const outlets = elements.filter(el => el.type === 'outlet');
            const count = outlets.length;
            if (count === 0) return;

            const radius = Math.min(centerX, centerY) - 80;

            outlets.forEach((el, index) => {
                const angle = (index / count) * 2 * Math.PI;
                const nodeX = centerX + radius * Math.cos(angle);
                const nodeY = centerY + radius * Math.sin(angle);

                tCtx.beginPath();
                tCtx.moveTo(centerX, centerY);
                tCtx.lineTo(nodeX, nodeY);
                tCtx.strokeStyle = '#3b82f6';
                tCtx.lineWidth = 2;
                tCtx.stroke();

                tCtx.beginPath();
                tCtx.arc(nodeX, nodeY, 20, 0, Math.PI * 2);
                tCtx.fillStyle = '#10b981';
                tCtx.fill();
                tCtx.lineWidth = 2;
                tCtx.strokeStyle = '#ffffff';
                tCtx.stroke();

                tCtx.fillStyle = '#ffffff';
                tCtx.font = 'bold 10px sans-serif';
                tCtx.fillText(el.id, nodeX, nodeY + 4);
            });
        }

        function calculateCubicacion() {
            const outletsCount = elements.filter(el => el.type === 'outlet').length || 4;
            const totalCableMeters = (outletsCount * 35 * 1.10).toFixed(0);
            const boxesOfCable = Math.ceil(totalCableMeters / 305);
            const jacksCount = outletsCount * 2;
            const patchCords = outletsCount * 2;

            const cardsContainer = document.getElementById('cubicationCards');
            cardsContainer.innerHTML = `
                <div class="bg-slate-900 border border-slate-700 p-4 rounded-xl space-y-1">
                    <p class="text-xs text-slate-400">Cable UTP Cat 6 Total</p>
                    <p class="text-xl font-black text-white">${totalCableMeters} metros</p>
                    <p class="text-[11px] text-emerald-400">Equivale a ${boxesOfCable} bobina(s) de 305m</p>
                </div>
                <div class="bg-slate-900 border border-slate-700 p-4 rounded-xl space-y-1">
                    <p class="text-xs text-slate-400">Salidas y Conectores</p>
                    <p class="text-xl font-black text-white">${jacksCount} Jacks RJ45</p>
                    <p class="text-[11px] text-blue-400">Incluye faceplates y tomas dobles</p>
                </div>
                <div class="bg-slate-900 border border-slate-700 p-4 rounded-xl space-y-1">
                    <p class="text-xs text-slate-400">Patch Cords de Conexión</p>
                    <p class="text-xl font-black text-white">${patchCords} unidades</p>
                    <p class="text-[11px] text-amber-400">Cat 6 certificados (2m y 3m)</p>
                </div>
            `;

            const tbody = document.getElementById('cubicationTableBody');
            tbody.innerHTML = `
                <tr><td class="p-3 font-semibold">Cable UTP Cat 6 (Holgura 10%)</td><td class="p-3">ANSI/TIA-568.2-D</td><td class="p-3">${totalCableMeters}</td><td class="p-3">Metros</td><td class="p-3 text-slate-400">Incluye merma de tendido y bajadas</td></tr>
                <tr><td class="p-3 font-semibold">Jacks RJ45 Cat 6</td><td class="p-3">Tipo Keystone Jack</td><td class="p-3">${jacksCount}</td><td class="p-3">Unidades</td><td class="p-3 text-slate-400">Montaje en gabinete y área de trabajo</td></tr>
                <tr><td class="p-3 font-semibold">Faceplate 2 Slots</td><td class="p-3">Placa mural 86x86mm</td><td class="p-3">${outletsCount}</td><td class="p-3">Unidades</td><td class="p-3 text-slate-400">Con ventanas antipolvo</td></tr>
                <tr><td class="p-3 font-semibold">Patch Panel 24 Puertos</td><td class="p-3">Rack 19" 1U Cat 6</td><td class="p-3">1</td><td class="p-3">Unidad</td><td class="p-3 text-slate-400">Incluye organizador de cables horizontal</td></tr>
                <tr><td class="p-3 font-semibold">Canaleta PVC 40x25mm</td><td class="p-3">Con adhesivo y divisiones</td><td class="p-3">30</td><td class="p-3">Barras (2m)</td><td class="p-3 text-slate-400">Canalización perimetral segura</td></tr>
            `;
        }

        function renderBudgetTable() {
            const tbody = document.getElementById('budgetTableBody');
            tbody.innerHTML = '';

            let subtotal = 0;

            budgetItems.forEach((item, index) => {
                const totalItem = item.qty * item.unitPrice;
                subtotal += totalItem;

                const tr = document.createElement('tr');
                tr.className = "hover:bg-slate-900/40 transition";
                tr.innerHTML = `
                    <td class="p-3 font-medium"><input type="text" value="${item.desc}" onchange="updateBudgetItem(${index}, 'desc', this.value)" class="bg-slate-900 border border-slate-700 rounded px-2 py-1 w-full text-xs text-slate-200"></td>
                    <td class="p-3"><span class="bg-slate-700 px-2 py-0.5 rounded text-[10px]">${item.cat}</span></td>
                    <td class="p-3"><input type="number" value="${item.qty}" onchange="updateBudgetItem(${index}, 'qty', this.value)" class="bg-slate-900 border border-slate-700 rounded px-2 py-1 w-16 text-xs text-slate-200"></td>
                    <td class="p-3"><input type="number" value="${item.unitPrice}" onchange="updateBudgetItem(${index}, 'unitPrice', this.value)" class="bg-slate-900 border border-slate-700 rounded px-2 py-1 w-28 text-xs text-slate-200"></td>
                    <td class="p-3 font-semibold text-emerald-400">$${totalItem.toLocaleString('es-CL')} CLP</td>
                    <td class="p-3 text-slate-400">${item.provider}</td>
                    <td class="p-3 text-center">
                        <button onclick="removeBudgetItem(${index})" class="text-rose-400 hover:text-rose-300 p-1"><i data-lucide="trash-2" class="w-4 h-4"></i></button>
                    </td>
                `;
                tbody.appendChild(tr);
            });

            lucide.createIcons();

            const iva = subtotal * 0.19;
            const total = subtotal + iva;

            document.getElementById('subtotalCLP').innerText = `$${subtotal.toLocaleString('es-CL')} CLP`;
            document.getElementById('ivaCLP').innerText = `$${iva.toLocaleString('es-CL')} CLP`;
            document.getElementById('totalCLP').innerText = `$${total.toLocaleString('es-CL')} CLP`;
        }

        function addBudgetRow() {
            budgetItems.push({ desc: 'Nuevo Ítem de Red', cat: 'Materiales', qty: 1, unitPrice: 10000, provider: 'Proveedor Local' });
            renderBudgetTable();
        }

        function updateBudgetItem(index, field, value) {
            budgetItems[index][field] = field === 'qty' || field === 'unitPrice' ? Number(value) : value;
            renderBudgetTable();
        }

        function removeBudgetItem(index) {
            budgetItems.splice(index, 1);
            renderBudgetTable();
        }

        function downloadExcelTemplate() {
            const wb = XLSX.utils.book_new();
            const wsData = [
                ["Descripción", "Categoría", "Cantidad", "Precio Unitario (CLP)", "Proveedor"],
                ["Cable UTP Cat 6 (Bobina 305m)", "Materiales", 2, 125000, "Sodimac"],
                ["Conectores RJ45 Cat 6", "Materiales", 1, 28900, "PC Factory"],
                ["Patch Panel 24 Puertos", "Materiales", 1, 45000, "Mercado Libre"]
            ];
            const ws = XLSX.utils.aoa_to_sheet(wsData);
            XLSX.utils.book_append_sheet(wb, ws, "Presupuesto");
            XLSX.writeFile(wb, "Plantilla_Presupuesto_NetCAD.xlsx");
        }

        function exportToExcel() {
            const wb = XLSX.utils.book_new();
            const wsData = [["Descripción", "Categoría", "Cantidad", "Precio Unitario (CLP)", "Total (CLP)", "Proveedor"]];
            budgetItems.forEach(item => {
                wsData.push([item.desc, item.cat, item.qty, item.unitPrice, item.qty * item.unitPrice, item.provider]);
            });
            const ws = XLSX.utils.aoa_to_sheet(wsData);
            XLSX.utils.book_append_sheet(wb, ws, "Presupuesto_NetCAD");
            XLSX.writeFile(wb, "Presupuesto_Redes_NetCAD.xlsx");
        }

        function handleExcelImport(event) {
            const file = event.target.files[0];
            if (!file) return;

            const reader = new FileReader();
            reader.onload = function(e) {
                const data = new Uint8Array(e.target.result);
                const workbook = XLSX.read(data, { type: 'array' });
                const firstSheetName = workbook.SheetNames[0];
                const worksheet = workbook.Sheets[firstSheetName];
                const json = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

                if (json.length > 1) {
                    budgetItems = [];
                    for (let i = 1; i < json.length; i++) {
                        const row = json[i];
                        if (row[0] && row[2]) {
                            budgetItems.push({
                                desc: row[0],
                                cat: row[1] || 'Materiales',
                                qty: Number(row[2]) || 1,
                                unitPrice: Number(row[3]) || 0,
                                provider: row[4] || 'Importado de Excel'
                            });
                        }
                    }
                    renderBudgetTable();
                    alert("¡Excel importado y datos traspasados con éxito!");
                }
            };
            reader.readAsArrayBuffer(file);
        }

        function exportProjectJSON() {
            const projectData = { elements, budgetItems, scalePixelsPerMeter };
            const dataStr = "data:text/json;charset=utf-8," + encodeURIComponent(JSON.stringify(projectData));
            const downloadAnchor = document.createElement('a');
            downloadAnchor.setAttribute("href", dataStr);
            downloadAnchor.setAttribute("download", "proyecto_netcad_inacap.json");
            document.body.appendChild(downloadAnchor);
            downloadAnchor.click();
            downloadAnchor.remove();
        }

        function importProjectJSON(event) {
            const file = event.target.files[0];
            if (!file) return;
            const reader = new FileReader();
            reader.onload = function(e) {
                try {
                    const data = JSON.parse(e.target.result);
                    if (data.elements) elements = data.elements;
                    if (data.budgetItems) budgetItems = data.budgetItems;
                    if (data.scalePixelsPerMeter) scalePixelsPerMeter = data.scalePixelsPerMeter;
                    drawCad();
                    renderBudgetTable();
                    alert("¡Proyecto cargado correctamente!");
                } catch(err) {
                    alert("Error al parsear el archivo JSON.");
                }
            };
            reader.readAsText(file);
        }

        function generateReadmeModal() {
            const outletsCount = elements.filter(el => el.type === 'outlet').length;
            const readmeContent = `# Proyecto de Redes Estructuradas - INACAP Sede La Serena

## Integrantes del Grupo
* **Nicolas Torres** (Estudiante Líder / Redes)
* Compañeros de Grupo (INACAP 2026)

## Descripción del Proyecto
Este repositorio contiene la planificación, trazado CAD, cubicación y presupuesto de la infraestructura de red de datos (UTP Cat 6 / Cat 6A) desarrollada con la herramienta **NetCAD**.

### Características Principales:
* **Topología en Estrella**: ${outletsCount} puestos de trabajo conectados al MDF principal sin superar los 90 metros (Norma TIA/EIA-568).
* **Cubicación Automática**: Cálculo de metros de cable con 10% de holgura, jacks, faceplates y patch panels.
* **Presupuesto CLP**: Valorización detallada en pesos chilenos considerando IVA (19%).
* **Despliegue**: Publicado en Cloudflare Pages.
`;
            document.getElementById('readmeTextarea').value = readmeContent;
            document.getElementById('readmeModal').classList.remove('hidden');
        }

        function closeReadmeModal() {
            document.getElementById('readmeModal').classList.add('hidden');
        }

        function copyReadme() {
            const textarea = document.getElementById('readmeTextarea');
            textarea.select();
            document.execCommand('copy');
            alert("¡Código Markdown copiado al portapapeles!");
        }

        function toggleChat() {
            const win = document.getElementById('chatWindow');
            win.classList.toggle('hidden');
        }

        function toggleApiKeyModal() {
            const modal = document.getElementById('apiKeyModal');
            modal.classList.toggle('hidden');
            if(!modal.classList.contains('hidden')) {
                document.getElementById('apiKeyInput').value = geminiApiKey;
            }
        }

        function saveApiKey() {
            geminiApiKey = document.getElementById('apiKeyInput').value.trim();
            localStorage.setItem('gemini_api_key', geminiApiKey);
            toggleApiKeyModal();
            appendChatMessage("Sistema", "¡Clave de API guardada con éxito!");
        }

        function handleChatKey(e) {
            if (e.key === 'Enter') sendChatMessage();
        }

        async function sendChatMessage() {
            const input = document.getElementById('chatInput');
            const text = input.value.trim();
            if (!text) return;

            appendChatMessage("Tú", text);
            input.value = '';

            if (!geminiApiKey) {
                appendChatMessage("NetCAD AI", "⚠️ Configura tu clave API haciendo clic en la rueda dentada (⚙️).");
                return;
            }

            try {
                const { GoogleGenAI } = await import("@google/genai");
                const ai = new GoogleGenAI({ apiKey: geminiApiKey });
                
                appendChatMessage("NetCAD AI", "Analizando con Gemini...", true);

                const response = await ai.models.generateContent({
                    model: 'gemini-2.5-flash',
                    contents: `Actúa como un profesor experto en redes estructuradas de INACAP. Ayuda al estudiante con respecto a: ${text}`
                });

                const chatMsgs = document.getElementById('chatMessages');
                chatMsgs.removeChild(chatMsgs.lastChild);

                appendChatMessage("NetCAD AI", response.text);
            } catch (error) {
                const chatMsgs = document.getElementById('chatMessages');
                if(chatMsgs.lastChild) chatMsgs.removeChild(chatMsgs.lastChild);
                appendChatMessage("NetCAD AI", "Error al conectar con Gemini: " + error.message);
            }
        }

        function appendChatMessage(sender, text, isTemporary = false) {
            const chatMsgs = document.getElementById('chatMessages');
            const div = document.createElement('div');
            div.className = `p-3 rounded-xl border text-xs space-y-1 ${sender === 'Tú' ? 'bg-blue-600/20 border-blue-500/30 ml-6' : 'bg-slate-900/80 border-slate-700 mr-6'}`;
            div.innerHTML = `<p class="font-bold ${sender === 'Tú' ? 'text-blue-400' : 'text-emerald-400'}">${sender}</p><p class="text-slate-200">${text}</p>`;
            chatMsgs.appendChild(div);
            chatMsgs.scrollTop = chatMsgs.scrollHeight;
        }

        async function aiAnalyzePlan() {
            const outlets = elements.filter(el => el.type === 'outlet').length;
            const prompt = `Analiza mi plano actual: tengo ${outlets} salidas conectadas al MDF principal bajo la norma TIA/EIA-568. Dame 3 recomendaciones técnicas clave.`;
            document.getElementById('chatInput').value = prompt;
            sendChatMessage();
        }

        async function aiSimulateDefense() {
            const prompt = `Hazme una pregunta de defensa oral sorpresa (como profesor de INACAP) sobre mi proyecto de redes estructuradas.`;
            document.getElementById('chatInput').value = prompt;
            sendChatMessage();
        }

        async function aiAskNorma() {
            const prompt = `Explícame brevemente qué dice la norma TIA/EIA-568 sobre la distancia máxima del enlace permanente.`;
            document.getElementById('chatInput').value = prompt;
            sendChatMessage();
        }
    </script>
</body>
</html>
