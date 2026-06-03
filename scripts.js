// Global State variables
let globalActivities = [];
let zoomScale = 1;
let translateX = 0;
let translateY = 0;
let isDragging = false;
let startX = 0;
let startY = 0;
let fileHandle = null;
let lastModified = 0;
let currentWorkbook = null;
let syncInterval = null;
let isCriticalPathFocused = false;
let activeTab = 'diagram';
let previousCalcMode = 'manual';

// Initialize Mermaid
mermaid.initialize({
    startOnLoad: false,
    securityLevel: 'loose',
    theme: 'base',
    themeVariables: {
        background: 'transparent',
        primaryColor: 'transparent',
        edgeColor: '#374151',
        lineColor: '#374151',
        transition: 'none'
    },
    flowchart: {
        htmlLabels: true,
        useMaxWidth: false,
        curve: 'basis'
    }
});

function clearInput() {
    document.getElementById('activityInput').value = '';
    showToast("Data Input Cleared");
}

/**
 * Links a local CSV or Excel file via File System Access API
 */
async function linkLocalFile() {
    try {
        [fileHandle] = await window.showOpenFilePicker({
            types: [
                {
                    description: 'Excel or CSV Files',
                    accept: {
                        'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet': ['.xlsx'],
                        'text/csv': ['.csv']
                    }
                },
            ],
            excludeAcceptAllOption: true,
            multiple: false
        });

        showToast("File Linked: " + fileHandle.name);
        document.getElementById('linkFileBtn').classList.add('active');
        document.getElementById('linkFileBtn').innerText = "Linked: " + fileHandle.name;
        
        // Initial Sync
        await syncWithLinkedFile();

        // Start polling for "Live" effect
        if (syncInterval) clearInterval(syncInterval);
        syncInterval = setInterval(syncWithLinkedFile, 2000);

    } catch (err) {
        if (err.name !== 'AbortError') {
            console.error(err);
            showToast("Failed to link file", true);
        }
    }
}

/**
 * Reads the linked file, parses it, and updates the PDM
 */
async function syncWithLinkedFile(forceUpdate = false) {
    if (!fileHandle) return;
    if (forceUpdate !== true) forceUpdate = false;

    try {
        const file = await fileHandle.getFile();
        const isModified = file.lastModified > lastModified;
        
        // Only re-read file from disk if it changed or we don't have it cached
        if (isModified || !currentWorkbook) {
            lastModified = file.lastModified;
            if (file.name.endsWith('.xlsx')) {
                const data = await file.arrayBuffer();
                currentWorkbook = XLSX.read(data);
            }
        }

        // Stop if nothing changed and no manual override
        if (!isModified && !forceUpdate && globalActivities.length > 0) return;

        const sheetNameInput = document.getElementById('excelSheetName').value.trim();

        let content = "";
        if (file.name.endsWith('.xlsx')) {
            const targetSheetName = sheetNameInput || currentWorkbook.SheetNames[0];
            const worksheet = currentWorkbook.Sheets[targetSheetName];
            if (!worksheet) {
                throw new Error(`Sheet "${targetSheetName}" not found.`);
            }
            content = XLSX.utils.sheet_to_csv(worksheet);
        } else {
            content = await file.text();
        }

        if (content) {
            // Remove all quotation marks from the data as requested
            document.getElementById('activityInput').value = content.replace(/"/g, '').trim();
            generatePDM();
            console.log("Auto-synced with " + fileHandle.name);
        }
    } catch (err) {
        console.error("Sync Error:", err);
        // If permission is lost (e.g. page refresh), stop sync
        if (err.name === 'NotAllowedError') {
            clearInterval(syncInterval);
            document.getElementById('linkFileBtn').classList.remove('active');
            document.getElementById('linkFileBtn').innerText = "Link Excel/CSV";
        }
    }
}

// Helper to generate Mermaid-safe IDs
function getSafeId(id) {
    return 'node_' + String(id).replace(/[^a-z0-9]/gi, '_');
}

function showHelp() {
    const calcMode = document.getElementById('calcModeSelect')?.value || 'manual';
    if (calcMode === 'auto') {
        alert("PDM Studio Auto CPM CSV Format:\n\nFormat: Item ID, [Node No (Optional)], Activity Name, Duration, [Predecessor IDs]\n\n- Node No: If omitted, sequence numbers (10, 20, 30...) are automatically assigned.\n- Predecessors: Semicolon ';' separated list of activity IDs (e.g., A1;B1). Use '-' or leave blank if no predecessors.\n\nExample:\nA1, 10, Setup, 5\nB1, Design Phase, 10, A1\nC1, Code Phase, 15, B1\nD1, Release, 5, B1;C1");
    } else {
        alert("PDM Studio Manual CSV Format:\n\nFormat: Item ID, [Node No (Optional)], Activity Name, ES, EF, LS, LF\n\n- Node No: If omitted, sequence numbers (10, 20, 30...) are automatically assigned.\n- ES/EF/LS/LF: Early Start, Early Finish, Late Start, Late Finish.\n\nExample:\nA1, 10, Design Phase, 0, 5, 0, 5\nB1, Coding Phase, 5, 15, 7, 17");
    }
}

// Handles switching between manual CSV input format and automatic CPM format
function changeCalcMode() {
    const calcModeSelect = document.getElementById('calcModeSelect');
    if (!calcModeSelect) return;
    const newMode = calcModeSelect.value;
    if (newMode === previousCalcMode) return;

    const textarea = document.getElementById('activityInput');
    const manualSample = `A1, 10, Setup, 0, 5, 0, 5\nB1, 20, Design, 5, 15, 5, 15\nC1, 30, Logic, 15, 30, 15, 30\nD1, 40, QA, 15, 23, 22, 30\nE1, 50, Release, 30, 35, 30, 35`;
    const autoSample = `A1, 10, Setup, 5\nB1, 20, Design, 10, A1\nC1, 30, Logic, 15, B1\nD1, 40, QA, 8, B1\nE1, 50, Release, 5, C1;D1`;

    // If the user hasn't edited the default text or it's empty, switch it to the other default
    const currentVal = textarea.value.trim();
    if (currentVal === "" || currentVal === manualSample || currentVal === autoSample) {
        textarea.value = (newMode === 'auto') ? autoSample : manualSample;
    }

    previousCalcMode = newMode;
    generatePDM();
}

// Change App Theme
function changeTheme() {
    const select = document.getElementById('themeSelect');
    const theme = select.value;
    
    document.body.className = '';
    if (theme !== 'midnight') {
        document.body.classList.add('theme-' + theme);
    }
    
    // Re-render chart to apply correct theme colors on elements
    generatePDM();
}

// Parse CSV activities
function parseInput(input, calcMode = 'manual') {
    const lines = input.split('\n')
        .map(line => line.trim())
        .filter(line => line.length > 0);

    const activities = [];
    const usedIds = new Set();

    let index = 0;
    for (const line of lines) {
        // Remove quotation marks and split by comma
        const parts = line.split(',').map(p => p.trim().replace(/"/g, ''));
        
        if (calcMode === 'auto') {
            if (parts.length < 3) continue;
            
            const id = parts[0];
            if (id.toLowerCase() === "item id" || id.toLowerCase() === "id") continue;

            if (usedIds.has(id)) {
                throw new Error(`Duplicate Item ID found: "${id}" on line ${index + 1}. IDs must be unique.`);
            }

            // Format: Item ID, [Node No], Name, Duration, Predecessors
            // If parts[1] is numeric and length >= 4, it is Node No
            const hasNodeNo = parts.length >= 4 && !isNaN(parseInt(parts[1]));
            
            const nodeNo = hasNodeNo ? parts[1] : String((index + 1) * 10);
            const name = hasNodeNo ? parts[2] : parts[1];
            const durationIndex = hasNodeNo ? 3 : 2;
            const duration = parseInt(parts[durationIndex]);
            
            const predsIndex = durationIndex + 1;
            const predsStr = parts[predsIndex] || "";
            const predecessorIds = predsStr.split(/[;|]/).map(p => p.trim()).filter(p => p.length > 0 && p !== "-");

            if (isNaN(duration)) {
                throw new Error(`Line ${index + 1} (${id}): Duration must be a valid number.`);
            }
            if (duration < 0) {
                throw new Error(`Line ${index + 1} (${id}): Duration cannot be negative.`);
            }

            activities.push({
                id,
                nodeNo,
                name,
                duration,
                predecessorIds,
                preds: []
            });
            usedIds.add(id);
        } else {
            if (parts.length < 6) continue; 

            const hasNodeNo = parts.length >= 7 && !isNaN(parseInt(parts[1]));
            
            const id = parts[0];
            if (id.toLowerCase() === "item id" || id.toLowerCase() === "id") continue;

            if (usedIds.has(id)) {
                throw new Error(`Duplicate Item ID found: "${id}" on line ${index + 1}. IDs must be unique.`);
            }

            const nodeNo = hasNodeNo ? parts[1] : String((index + 1) * 10);
            const name = hasNodeNo ? parts[2] : parts[1];
            
            const esIndex = hasNodeNo ? 3 : 2;
            const es = parseInt(parts[esIndex]);
            const ef = parseInt(parts[esIndex + 1]);
            const ls = parseInt(parts[esIndex + 2]);
            const lf = parseInt(parts[esIndex + 3]);

            // Validation: Check for numbers and logic consistency
            if ([es, ef, ls, lf].some(val => isNaN(val))) {
                throw new Error(`Line ${index + 1} (${id}): Timing values (ES, EF, LS, LF) must be valid numbers.`);
            }
            if (ef < es) {
                throw new Error(`Line ${index + 1} (${id}): Early Finish (${ef}) cannot be less than Early Start (${es}).`);
            }
            if (lf < ls) {
                throw new Error(`Line ${index + 1} (${id}): Late Finish (${lf}) cannot be less than Late Start (${ls}).`);
            }

            activities.push({
                id,
                nodeNo,
                name,
                es, ef, ls, lf,
                duration: ef - es,
                slack: lf - ef,
                preds: []
            });
            usedIds.add(id);
        }
        index++;
    }
    return activities;
}

// Precedence calculation (Precedence Diagram Method)
function calculatePDM(activities) {
    if (activities.length === 0) return [];

    // Calculate critical state
    activities.forEach(a => {
        a.isCritical = a.slack <= 0;
    });

    // Find project start node
    const minES = Math.min(...activities.map(a => a.es));
    
    activities.forEach(i => {
        if (i.es === minES) return; // Project start nodes don't search predecessors

        // Find candidate predecessors (EF <= ES)
        const candidates = activities.filter(j => j.id !== i.id && j.ef <= i.es);
        
        candidates.forEach(j => {
            // Filter out redundant transitive dependencies
            const isRedundant = candidates.some(k => 
                k.id !== j.id && k.id !== i.id && j.ef <= k.es && k.ef <= i.es
            );
            
            if (!isRedundant) {
                i.preds.push({
                    id: j.id,
                    lag: i.es - j.ef
                });
            }
        });
    });

    return activities;
}

// Automatic Critical Path Method calculation (Forward & Backward pass)
function calculateCPM(activities) {
    if (activities.length === 0) return [];

    // 1. Build graph structure
    const adj = {};
    const inDegree = {};
    const nodes = {};

    activities.forEach(a => {
        nodes[a.id] = a;
        adj[a.id] = [];
        inDegree[a.id] = 0;
    });

    // Populate dependencies and inDegrees
    activities.forEach(a => {
        a.predecessorIds.forEach(predId => {
            if (nodes[predId]) {
                adj[predId].push(a.id);
                inDegree[a.id]++;
            } else {
                throw new Error(`Task "${a.id}" references undefined predecessor "${predId}".`);
            }
        });
    });

    // 2. Topological Sort (Kahn's Algorithm) to check for circular dependencies
    const queue = [];
    activities.forEach(a => {
        if (inDegree[a.id] === 0) {
            queue.push(a.id);
        }
    });

    const topoOrder = [];
    while (queue.length > 0) {
        const u = queue.shift();
        topoOrder.push(u);
        adj[u].forEach(v => {
            inDegree[v]--;
            if (inDegree[v] === 0) {
                queue.push(v);
            }
        });
    }

    if (topoOrder.length < activities.length) {
        throw new Error("Circular dependency detected in project tasks!");
    }

    // 3. Forward Pass (Calculate Early Start and Early Finish)
    activities.forEach(a => {
        a.es = 0;
        a.ef = 0;
    });

    topoOrder.forEach(id => {
        const u = nodes[id];
        if (u.predecessorIds.length > 0) {
            let maxEF = 0;
            u.predecessorIds.forEach(predId => {
                const pred = nodes[predId];
                if (pred && pred.ef > maxEF) {
                    maxEF = pred.ef;
                }
            });
            u.es = maxEF;
        } else {
            u.es = 0;
        }
        u.ef = u.es + u.duration;
    });

    // 4. Backward Pass (Calculate Late Start and Late Finish)
    const projectDuration = Math.max(...activities.map(a => a.ef), 0);

    // Initialize late times with project duration
    activities.forEach(a => {
        a.lf = projectDuration;
        a.ls = projectDuration;
    });

    // Iterate backwards through topological order
    for (let i = topoOrder.length - 1; i >= 0; i--) {
        const id = topoOrder[i];
        const u = nodes[id];
        
        // Find all successors of u
        const successors = activities.filter(a => a.predecessorIds.includes(u.id));
        if (successors.length > 0) {
            let minLS = Infinity;
            successors.forEach(succ => {
                if (succ.ls < minLS) {
                    minLS = succ.ls;
                }
            });
            u.lf = minLS;
        } else {
            u.lf = projectDuration;
        }
        u.ls = u.lf - u.duration;
        u.slack = u.lf - u.ef;
        u.isCritical = u.slack <= 0;
    }

    // 5. Populate preds array for visual rendering compatibility
    activities.forEach(a => {
        a.preds = [];
        a.predecessorIds.forEach(predId => {
            const predNode = nodes[predId];
            if (predNode) {
                a.preds.push({
                    id: predId,
                    lag: a.es - predNode.ef
                });
            }
        });
    });

    return activities;
}

/**
 * Determines the sequence of the critical path by tracing nodes with zero slack
 */
function getCriticalPathSequence(activities) {
    const criticalNodes = activities.filter(a => a.isCritical);
    if (criticalNodes.length === 0) return "No critical path identified.";

    // Sort by Early Start to find the chronological sequence
    const sequence = criticalNodes
        .sort((a, b) => a.es - b.es)
        .map(a => a.id);

    return sequence.join(' → ');
}

/**
 * Toggles a persistent focus on the critical path by dimming non-critical elements
 */
function toggleCriticalPathFocus() {
    if (isCriticalPathFocused) {
        clearHighlightChain();
        isCriticalPathFocused = false;
        showToast("Critical Path Focus Disabled");
        return;
    }

    const criticalIds = new Set(globalActivities.filter(a => a.isCritical).map(a => a.id));
    if (criticalIds.size === 0) {
        showToast("No critical path to highlight", true);
        return;
    }

    // Apply dimming effect to non-critical nodes
    const allNodes = document.querySelectorAll('#diagram .node');
    globalActivities.forEach(a => {
        const safeId = getSafeId(a.id);
        allNodes.forEach(el => {
            if (el.id !== safeId && !el.id.endsWith('-' + safeId)) return;
            
            if (criticalIds.has(a.id)) {
                el.style.opacity = '1';
                el.style.filter = 'none';
            } else {
                el.style.opacity = '0.15';
                el.style.filter = 'grayscale(80%) blur(1px)';
            }
            el.style.transition = 'all 0.3s ease';
        });
    });

    // Apply dimming effect to non-critical edges
    const edges = document.querySelectorAll('#diagram .edgePath');
    edges.forEach(edge => {
        const classes = Array.from(edge.classList).join(' ');
        let activeCriticalNodes = 0;
        
        criticalIds.forEach(id => {
            if (classes.includes(getSafeId(id))) activeCriticalNodes++;
        });

        const isCriticalEdge = activeCriticalNodes >= 2;
        edge.style.opacity = isCriticalEdge ? '1' : '0.1';
        edge.style.transition = 'all 0.3s ease';

        // Also dim/hide markers (arrowheads) for non-critical edges
        const markerEnd = edge.querySelector('.path').getAttribute('marker-end');
        if (markerEnd) {
            const markerId = markerEnd.replace(/url\(["']?#|["']?\)/g, '');
            const marker = document.getElementById(markerId);
            if (marker) marker.style.opacity = isCriticalEdge ? '1' : '0.1';
        }
    });

    isCriticalPathFocused = true;
    showToast("Focus: " + getCriticalPathSequence(globalActivities));
}

// Generate Mermaid code string with embedded HTML tables
function generateMermaid(activities) {
    const flowDir = document.getElementById('directionSelect').value || 'LR';
    const theme = document.getElementById('themeSelect').value;
    
    // Mermaid's internal parser fails on CSS variables like var(--color-critical).
    // We map our theme variables to literal hex codes for the parser.
    const themePalette = {
        midnight: { critical: '#f43f5e', normal: '#374151' },
        cyberpunk: { critical: '#ff0055', normal: '#ec4899' },
        emerald: { critical: '#fb7185', normal: '#15803d' },
        light: { critical: '#e11d48', normal: '#cbd5e1' }
    };
    const palette = themePalette[theme] || themePalette.midnight;

    let code = `flowchart ${flowDir}\n`;

    activities.forEach(a => {
        const safeId = getSafeId(a.id);
        const cleanName = a.name.replace(/[<>"/]/g, "'");
        const cardClass = a.isCritical ? 'critical' : 'normal';

        // Elegant Table design matching the sketch
        const tableHtml = `
            <table class='pdm-node-table ${cardClass}' data-node-id='${a.id}'>
                <tr>
                    <td colspan='2' class='pdm-cell cell-item'>
                        <span class='cell-val'>${a.id}</span>
                        <span class='cell-lbl'>Item No.</span>
                    </td>
                    <td colspan='2' class='pdm-cell cell-node'>
                        <span class='cell-val'>${a.nodeNo}</span>
                        <span class='cell-lbl'>Node No.</span>
                    </td>
                    <td colspan='2' class='pdm-cell cell-duration'>
                        <span class='cell-val'>${a.duration}</span>
                        <span class='cell-lbl'>Duration</span>
                    </td>
                </tr>
                <tr>
                    <td colspan='6' class='pdm-cell cell-desc'>${cleanName}</td>
                </tr>
                <tr>
                    <td colspan='3' class='pdm-cell cell-es'>
                        <span class='cell-val'>${a.es}</span>
                        <span class='cell-lbl'>Early Start</span>
                    </td>
                    <td colspan='3' class='pdm-cell cell-ef'>
                        <span class='cell-val'>${a.ef}</span>
                        <span class='cell-lbl'>Early Finish</span>
                    </td>
                </tr>
                <tr>
                    <td colspan='3' class='pdm-cell cell-ls'>
                        <span class='cell-val'>${a.ls}</span>
                        <span class='cell-lbl'>Late Start</span>
                    </td>
                    <td colspan='3' class='pdm-cell cell-lf'>
                        <span class='cell-val'>${a.lf}</span>
                        <span class='cell-lbl'>Late Finish</span>
                    </td>
                </tr>
            </table>
        `.replace(/\s+/g, ' ').trim();

        code += `    ${safeId}["${tableHtml}"]\n`;
    });

    // Set link arrows and style links
    let linkIndex = 0;
    const linkStyles = [];

    activities.forEach(a => {
        const safeId = getSafeId(a.id);
        a.preds.forEach(pred => {
            const safePId = getSafeId(pred.id);
            const predNode = activities.find(act => act.id === pred.id);
            const isLinkCritical = a.isCritical && predNode && predNode.isCritical;

            let arrow = "-->";
            if (pred.lag > 0) {
                arrow = `-- "Lag: ${pred.lag}" -->`;
            } else if (pred.lag < 0) {
                arrow = `-- "Lead: ${Math.abs(pred.lag)}" -->`;
            }
            code += `    ${safePId} ${arrow} ${safeId}\n`;

            if (isLinkCritical) {
                linkStyles.push(`linkStyle ${linkIndex} stroke:${palette.critical},stroke-width:3px,stroke-dasharray:none;`);
            } else {
                linkStyles.push(`linkStyle ${linkIndex} stroke:${palette.normal},stroke-width:2px,stroke-dasharray:4;`);
            }
            linkIndex++;
        });
    });

    // Set node transparent style to let HTML table design take complete style control
    activities.forEach(a => {
        const safeId = getSafeId(a.id);
        code += `    style ${safeId} fill:none,stroke:none,stroke-width:0px;\n`;
    });

    code += '\n' + linkStyles.join('\n') + '\n';
    return code;
}

// Generate PDM main runner
async function generatePDM() {
    const input = document.getElementById('activityInput').value;
    const container = document.getElementById('diagram');
    const calcMode = document.getElementById('calcModeSelect')?.value || 'manual';
    
    try {
        let activities = parseInput(input, calcMode);
        if (activities.length === 0) {
            container.innerHTML = `<div style="padding: 40px; text-align: center; color: var(--text-secondary);">Please enter activity data in the sidebar CSV field.</div>`;
            document.getElementById('barchartContainer').innerHTML = `<div style="padding: 40px; text-align: center; color: var(--text-secondary);">Please enter activity data in the sidebar CSV field.</div>`;
            document.getElementById('analysisContainer').innerHTML = `<div style="padding: 40px; text-align: center; color: var(--text-secondary);">Please enter activity data in the sidebar CSV field.</div>`;
            updateStats(0, 0, 0);
            return;
        }

        if (calcMode === 'auto') {
            activities = calculateCPM(activities);
        } else {
            activities = calculatePDM(activities);
        }
        
        globalActivities = activities;
        const mermaidCode = generateMermaid(activities);
        
        container.innerHTML = mermaidCode;
        container.removeAttribute('data-processed');
        isCriticalPathFocused = false; // Reset focus state on new render
        
        await mermaid.run({
            nodes: [container],
        });

        // Update Stats Dashboard Widget
        const totalDuration = Math.max(...activities.map(a => a.ef), 0);
        const totalNodes = activities.length;
        const criticalNodes = activities.filter(a => a.isCritical).length;
        updateStats(totalDuration, totalNodes, criticalNodes);

        console.log("Critical Path Sequence:", getCriticalPathSequence(activities));

        // CSS styling modifications post-render
        colorArrowheadMarkers();
        setupInteractiveHighlighting();
        resetZoom();

        // Update active tab contents if we are in barchart or analysis
        if (activeTab === 'barchart') {
            renderBarChart(globalActivities);
        } else if (activeTab === 'analysis') {
            renderAnalysis(globalActivities);
        }
        
    } catch (err) {
        const errorHtml = `<div style="color: var(--color-critical); padding: 30px; background-color: rgba(244, 63, 94, 0.05); border: 1px solid var(--color-critical); border-radius: 8px;"><strong>Error Parsing Data:</strong><br/>${err.message}</div>`;
        container.innerHTML = errorHtml;
        document.getElementById('barchartContainer').innerHTML = errorHtml;
        document.getElementById('analysisContainer').innerHTML = errorHtml;
        updateStats(0, 0, 0);
        console.error("PDM Generator Error:", err);
    }
}

// Viewport tabs switching logic
function switchTab(tabId) {
    activeTab = tabId;
    
    // Update tab button active states
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.classList.remove('active');
    });
    const activeBtn = document.getElementById(`tabBtn_${tabId}`);
    if (activeBtn) activeBtn.classList.add('active');

    // Update tab content active states
    document.querySelectorAll('.tab-content').forEach(content => {
        content.classList.remove('active');
    });
    const activeContent = document.getElementById(`tabContent_${tabId}`);
    if (activeContent) activeContent.classList.add('active');

    // Show/Hide export buttons based on the active tab
    const exportSvg = document.getElementById('exportSvgBtn');
    const exportPng = document.getElementById('exportPngBtn');
    const copyMermaid = document.getElementById('copyMermaidBtn');

    if (tabId === 'diagram') {
        if (exportSvg) exportSvg.style.display = '';
        if (exportPng) exportPng.style.display = '';
        if (copyMermaid) copyMermaid.style.display = '';
    } else {
        if (exportSvg) exportSvg.style.display = 'none';
        if (exportPng) exportPng.style.display = 'none';
        if (copyMermaid) copyMermaid.style.display = 'none';
    }

    // Draw content
    if (tabId === 'barchart') {
        renderBarChart(globalActivities);
    } else if (tabId === 'analysis') {
        renderAnalysis(globalActivities);
    }
}

// Renders the interactive Gantt chart representation
function renderBarChart(activities) {
    const container = document.getElementById('barchartContainer');
    if (!container) return;

    if (activities.length === 0) {
        container.innerHTML = `<div style="padding: 40px; text-align: center; color: var(--text-secondary);">No data available. Enter activity data and render first.</div>`;
        return;
    }

    const projectDuration = Math.max(...activities.map(a => a.ef), 0);
    const colWidth = 50; // pixels per time unit
    const totalCols = Math.max(projectDuration + 2, 12); // padding columns

    // 1. Build Left Column Tasks List HTML
    let tasksHtml = `
        <div class="gantt-task-col">
            <div class="gantt-task-header">Task Activity</div>
            <div style="display: flex; flex-direction: column;">
    `;

    activities.forEach(a => {
        const isCritical = a.isCritical;
        const idClass = isCritical ? 'critical' : '';
        tasksHtml += `
                <div class="gantt-task-row">
                    <div class="gantt-task-name" title="${a.name}">${a.name}</div>
                    <div class="gantt-task-meta">
                        <span class="gantt-task-id ${idClass}">${a.id}</span>
                        <span>Dur: ${a.duration}</span>
                        <span>Slack: ${a.slack}</span>
                    </div>
                </div>
        `;
    });

    tasksHtml += `
            </div>
        </div>
    `;

    // 2. Build Right Column Timeline Header HTML
    let timelineHeaderHtml = `
        <div class="gantt-timeline-header" style="width: ${totalCols * colWidth}px;">
    `;
    for (let i = 0; i < totalCols; i++) {
        timelineHeaderHtml += `<div class="gantt-time-unit" style="width: ${colWidth}px; flex-shrink: 0;">${i}</div>`;
    }
    timelineHeaderHtml += `</div>`;

    // 3. Build Right Column Timeline Rows HTML
    let timelineRowsHtml = `
        <div style="display: flex; flex-direction: column; width: ${totalCols * colWidth}px;">
    `;

    activities.forEach(a => {
        const isCritical = a.isCritical;
        const barClass = isCritical ? 'critical' : 'normal';
        
        const barLeft = a.es * colWidth;
        const barWidth = a.duration * colWidth;
        const slackLeft = a.ef * colWidth;
        const slackWidth = a.slack * colWidth;

        timelineRowsHtml += `
            <div class="gantt-timeline-row" style="background-image: linear-gradient(90deg, var(--border-primary) 1px, transparent 1px); background-size: ${colWidth}px 100%;">
                <div class="gantt-bar-container">
                    <!-- Task Bar -->
                    <div class="gantt-bar ${barClass}" style="left: ${barLeft}px; width: ${barWidth}px;" title="${a.name} (${a.id}): ES=${a.es}, EF=${a.ef}, LS=${a.ls}, LF=${a.lf}, Slack=${a.slack}">
                        ${a.duration >= 1 ? `<span style="padding: 0 4px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap;">${a.id}</span>` : ''}
                    </div>
        `;

        if (a.slack > 0) {
            timelineRowsHtml += `
                    <!-- Slack Bar -->
                    <div class="gantt-slack-bar" style="left: ${slackLeft}px; width: ${slackWidth}px;" title="Slack/Float: ${a.slack}"></div>
            `;
        }

        timelineRowsHtml += `
                </div>
            </div>
        `;
    });

    timelineRowsHtml += `</div>`;

    // 4. Combine into final layout
    container.innerHTML = `
        <div class="gantt-chart">
            ${tasksHtml}
            <div class="gantt-timeline-col-wrapper">
                <div style="display: flex; flex-direction: column; width: fit-content;">
                    ${timelineHeaderHtml}
                    ${timelineRowsHtml}
                </div>
            </div>
        </div>
    `;
}

// Renders the Critical Path analysis breakdown and schedule metrics
function renderAnalysis(activities) {
    const container = document.getElementById('analysisContainer');
    if (!container) return;

    if (activities.length === 0) {
        container.innerHTML = `<div style="padding: 40px; text-align: center; color: var(--text-secondary);">No data available. Enter activity data and render first.</div>`;
        return;
    }

    const criticalNodes = activities.filter(a => a.isCritical).sort((a, b) => a.es - b.es);
    const totalDuration = Math.max(...activities.map(a => a.ef), 0);
    
    // 1. Build Critical Path Sequence Flow
    let sequenceHtml = "";
    if (criticalNodes.length === 0) {
        sequenceHtml = "<div style='color: var(--color-critical); font-weight: 600;'>No critical path identified. All tasks have float.</div>";
    } else {
        criticalNodes.forEach((node, idx) => {
            sequenceHtml += `
                <div class="seq-node">
                    <span style="font-size: 14px;">${node.id}</span>
                    <span class="seq-node-name" title="${node.name}">${node.name}</span>
                    <span style="font-size: 9px; opacity: 0.7;">Dur: ${node.duration}</span>
                </div>
            `;
            if (idx < criticalNodes.length - 1) {
                sequenceHtml += `<div class="seq-arrow">→</div>`;
            }
        });
    }

    // 2. Build Critical Tasks Table
    let tableRowsHtml = "";
    if (criticalNodes.length === 0) {
        tableRowsHtml = "<tr><td colspan='7' style='text-align: center; color: var(--text-secondary);'>No critical tasks.</td></tr>";
    } else {
        criticalNodes.forEach(node => {
            tableRowsHtml += `
                <tr>
                    <td style="font-weight: 700; color: var(--color-critical);">${node.id}</td>
                    <td>${node.name}</td>
                    <td style="font-weight: 600;">${node.duration}</td>
                    <td>${node.es}</td>
                    <td>${node.ef}</td>
                    <td>${node.ls}</td>
                    <td>${node.lf}</td>
                </tr>
            `;
        });
    }

    // 3. Schedule Health and Risks
    const nearCriticalNodes = activities.filter(a => !a.isCritical && a.slack > 0 && a.slack <= 3).sort((a, b) => a.slack - b.slack);
    const longestCritical = criticalNodes.reduce((max, node) => node.duration > max.duration ? node : max, { duration: 0, id: "None" });
    
    let risksHtml = "";
    
    // Risk 1: Longest Critical Task
    if (longestCritical.id !== "None" && longestCritical.duration > 0) {
        risksHtml += `
            <div class="risk-item high">
                <div class="risk-icon" style="color: var(--color-critical); flex-shrink: 0;">⚠️</div>
                <div class="risk-details">
                    <div class="risk-title">Bottleneck Risk: Task ${longestCritical.id}</div>
                    <div class="risk-desc">"${longestCritical.name}" is the longest task on the critical path (duration: ${longestCritical.duration}). Any delay here directly delays the project completion.</div>
                </div>
            </div>
        `;
    }

    // Risk 2: Near-critical tasks
    if (nearCriticalNodes.length > 0) {
        nearCriticalNodes.forEach(node => {
            risksHtml += `
                <div class="risk-item warn">
                    <div class="risk-icon" style="color: #fbbf24; flex-shrink: 0;">⚠️</div>
                    <div class="risk-details">
                        <div class="risk-title">Near-Critical: Task ${node.id}</div>
                        <div class="risk-desc">"${node.name}" has very low float/slack (${node.slack} units). A small delay of ${node.slack} units will make it critical and may delay the project.</div>
                    </div>
                </div>
            `;
        });
    }

    // Health Indicators
    const criticalRatio = activities.length > 0 ? (criticalNodes.length / activities.length) : 0;
    let healthSummary = "";
    let healthClass = "";
    if (criticalRatio > 0.7) {
        healthSummary = "Highly Constrained Schedule. Over 70% of your tasks are critical. There is almost zero flexibility; any delay will cascade.";
        healthClass = "color: var(--color-critical);";
    } else if (criticalRatio > 0.4) {
        healthSummary = "Moderately Constrained Schedule. Between 40% and 70% of tasks are critical. Manage resources carefully.";
        healthClass = "color: #fbbf24;";
    } else {
        healthSummary = "Healthy Schedule. Less than 40% of tasks are critical, leaving ample float/flexibility in most paths.";
        healthClass = "color: #34d399;";
    }

    const html = `
        <div class="analysis-grid">
            <div class="analysis-card full-width">
                <h3>
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" style="color: var(--color-critical); flex-shrink: 0;">
                        <path d="M22 11.08V12a10 10 0 1 1-5.93-9.14" />
                        <polyline points="22 4 12 14.01 9 11.01" />
                    </svg>
                    Critical Path Sequence (Longest Path)
                </h3>
                <div class="critical-sequence-flow">
                    ${sequenceHtml}
                </div>
            </div>
            <div class="analysis-card">
                <h3>
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" style="flex-shrink: 0;">
                        <line x1="8" y1="6" x2="21" y2="6" />
                        <line x1="8" y1="12" x2="21" y2="12" />
                        <line x1="8" y1="18" x2="21" y2="18" />
                        <line x1="3" y1="6" x2="3.01" y2="6" />
                        <line x1="3" y1="12" x2="3.01" y2="12" />
                        <line x1="3" y1="18" x2="3.01" y2="18" />
                    </svg>
                    Critical Tasks Breakdown
                </h3>
                <div style="overflow-x: auto;">
                    <table class="analysis-table">
                        <thead>
                            <tr>
                                <th>ID</th>
                                <th>Name</th>
                                <th>Duration</th>
                                <th>ES</th>
                                <th>EF</th>
                                <th>LS</th>
                                <th>LF</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${tableRowsHtml}
                        </tbody>
                    </table>
                </div>
            </div>
            <div class="analysis-card">
                <h3>
                    <svg width="16" height="16" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" style="flex-shrink: 0;">
                        <path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z" />
                        <line x1="12" y1="9" x2="12" y2="13" />
                        <line x1="12" y1="17" x2="12.01" y2="17" />
                    </svg>
                    Schedule Health & Risks
                </h3>
                <div style="margin-bottom: 20px; padding: 12px; border-radius: 8px; border: 1px solid var(--border-primary); background-color: rgba(255,255,255,0.01);">
                    <div style="font-weight: 700; font-size: 13px; margin-bottom: 4px; ${healthClass}">
                        ${Math.round(criticalRatio * 100)}% Critical Tasks
                    </div>
                    <div style="font-size: 11px; color: var(--text-secondary); line-height: 1.4;">
                        ${healthSummary}
                    </div>
                </div>
                <div class="risks-list">
                    ${risksHtml || "<div style='color: var(--text-secondary); font-size: 12px; text-align: center; padding: 10px 0;'>No active schedule risks identified. All non-critical tasks have healthy float margins.</div>"}
                </div>
            </div>
        </div>
    `;

    container.innerHTML = html;
}

// Color Arrowheads matching link color in SVG
function colorArrowheadMarkers() {
    const svg = document.querySelector('#diagram svg');
    if (!svg) return;
    
    const paths = svg.querySelectorAll('.edgePath .path');
    paths.forEach((path) => {
        const strokeColor = path.style.stroke || window.getComputedStyle(path).stroke;
        const markerEnd = path.getAttribute('marker-end');
        if (markerEnd) {
            const markerId = markerEnd.replace(/url\(["']?#|["']?\)/g, '');
            const marker = svg.getElementById(markerId);
            if (marker) {
                const markerPath = marker.querySelector('path');
                if (markerPath) {
                    markerPath.style.fill = strokeColor;
                    markerPath.style.stroke = 'none';
                }
            }
        }
    });
}

// Interactive chain hover highlighting logic
function setupInteractiveHighlighting() {
    document.querySelectorAll('.pdm-node-table').forEach(table => {
        table.addEventListener('mouseenter', () => {
            const id = table.getAttribute('data-node-id');
            highlightNodeChain(id);
        });
        table.addEventListener('mouseleave', () => {
            clearHighlightChain();
        });
    });
}

function highlightNodeChain(targetId) {
    const activeIds = new Set([targetId]);
    
    // Build an adjacency list for successors for faster lookups
    const successorMap = {};
    globalActivities.forEach(a => {
        a.preds.forEach(p => {
            if (!successorMap[p.id]) successorMap[p.id] = [];
            successorMap[p.id].push(a.id);
        });
    });
    
    // Find predecessors recursively
    function findPreds(id) {
        const node = globalActivities.find(a => a.id === id);
        if (!node) return;
        node.preds.forEach(p => {
            if (!activeIds.has(p.id)) {
                activeIds.add(p.id);
                findPreds(p.id);
            }
        });
    }
    
    // Find successors recursively
    function findSuccs(id) {
        const succs = successorMap[id] || [];
        succs.forEach(sId => {
            if (!activeIds.has(sId)) {
                activeIds.add(sId);
                findSuccs(sId);
            }
        });
    }
    
    findPreds(targetId);
    findSuccs(targetId);
    
    // Apply opacity styles to node SVGs
    const allNodes = document.querySelectorAll('#diagram .node');
    globalActivities.forEach(a => {
        const safeId = getSafeId(a.id);
        allNodes.forEach(el => {
            if (el.id !== safeId && !el.id.endsWith('-' + safeId)) return;
            if (activeIds.has(a.id)) {
                el.style.opacity = '1';
                el.style.filter = 'none';
            } else {
                el.style.opacity = '0.15';
                el.style.filter = 'grayscale(80%) blur(1px)';
            }
            el.style.transition = 'all 0.25s ease';
        });
    });
    
    // Dim/Highlight connection edge lines
    const edges = document.querySelectorAll('#diagram .edgePath');
    edges.forEach(edge => {
        const classes = Array.from(edge.classList).join(' ');
        let activeEdge = false;
        
        // If the edge connects nodes in the active set, keep it active
        // Mermaid classes include node IDs like 'ls-node_A1' and 'le-node_B1'
        let countActiveNodes = 0;
        activeIds.forEach(id => {
            const safeId = 'node_' + id.replace(/[^a-z0-9]/gi, '_');
            if (classes.includes(safeId)) {
                countActiveNodes++;
            }
        });
        
        if (countActiveNodes >= 2) {
            edge.style.opacity = '1';
        } else {
            edge.style.opacity = '0.1';
        }
        edge.style.transition = 'all 0.25s ease';
    });
}

function clearHighlightChain() {
    const elements = document.querySelectorAll('#diagram .node, #diagram .edgePath, #diagram marker');
    elements.forEach(el => {
        el.style.opacity = '1';
        el.style.filter = 'none';
    });
    isCriticalPathFocused = false;
}

// Update Statistics Panel Info
function updateStats(duration, total, critical) {
    document.getElementById('statDuration').innerText = duration;
    document.getElementById('statNodes').innerText = total;
    document.getElementById('statCritical').innerText = critical;
    
    const ratio = total > 0 ? Math.round((critical / total) * 100) : 0;
    document.getElementById('statEfficiency').innerText = ratio + '%';
}

// Interactive Zoom & Pan script logic
const wrapper = document.getElementById('diagramWrapper');
const diagram = document.getElementById('diagram');

wrapper.addEventListener('mousedown', (e) => {
    if (e.target.closest('.pdm-node-table')) return; // Avoid drag on button clicks/hover
    isDragging = true;
    startX = e.clientX - translateX;
    startY = e.clientY - translateY;
});

window.addEventListener('mouseup', () => {
    isDragging = false;
});

wrapper.addEventListener('mousemove', (e) => {
    if (!isDragging) return;
    translateX = e.clientX - startX;
    translateY = e.clientY - startY;
    updateTransform();
});

wrapper.addEventListener('wheel', (e) => {
    e.preventDefault();
    const zoomIntensity = 0.05;
    
    const rect = wrapper.getBoundingClientRect();
    const mouseX = e.clientX - rect.left;
    const mouseY = e.clientY - rect.top;
    
    const wheel = e.deltaY < 0 ? 1 : -1;
    const zoomFactor = Math.exp(wheel * zoomIntensity);
    
    const newScale = Math.min(Math.max(zoomScale * zoomFactor, 0.15), 3.5);
    
    translateX = mouseX - (mouseX - translateX) * (newScale / zoomScale);
    translateY = mouseY - (mouseY - translateY) * (newScale / zoomScale);
    zoomScale = newScale;
    
    updateTransform();
}, { passive: false });

function adjustZoom(factor) {
    const wrapperRect = wrapper.getBoundingClientRect();
    const centerX = wrapperRect.width / 2;
    const centerY = wrapperRect.height / 2;
    
    const newScale = Math.min(Math.max(zoomScale * factor, 0.15), 3.5);
    
    translateX = centerX - (centerX - translateX) * (newScale / zoomScale);
    translateY = centerY - (centerY - translateY) * (newScale / zoomScale);
    zoomScale = newScale;
    
    updateTransform();
}

function resetZoom() {
    // Find SVG dimensions if loaded and fit it center
    const svgEl = document.querySelector('#diagram svg');
    if (svgEl) {
        const wrapperRect = wrapper.getBoundingClientRect();
        const svgWidth = svgEl.viewBox.baseVal.width || svgEl.getBoundingClientRect().width;
        const svgHeight = svgEl.viewBox.baseVal.height || svgEl.getBoundingClientRect().height;
        
        // Fit scale
        const scaleX = (wrapperRect.width - 60) / svgWidth;
        const scaleY = (wrapperRect.height - 60) / svgHeight;
        zoomScale = Math.min(Math.min(scaleX, scaleY), 1.0); // Fit completely, max zoom 1.0
        
        translateX = (wrapperRect.width - svgWidth * zoomScale) / 2;
        translateY = (wrapperRect.height - svgHeight * zoomScale) / 2;
    } else {
        zoomScale = 1;
        translateX = 0;
        translateY = 0;
    }
    updateTransform();
}

function updateTransform() {
    diagram.style.transform = `translate(${translateX}px, ${translateY}px) scale(${zoomScale})`;
}

// Export Functionality
function copyMermaidCode() {
    const svgEl = document.querySelector('#diagram svg');
    if (!svgEl) {
        showToast("No active diagram to copy", true);
        return;
    }
    const mermaidMarkup = generateMermaid(globalActivities);
    navigator.clipboard.writeText(mermaidMarkup).then(() => {
        showToast("Mermaid markup copied to clipboard");
    }).catch(err => {
        showToast("Failed to copy text", true);
    });
}

function downloadSVG() {
    const svgEl = document.querySelector('#diagram svg');
    if (!svgEl) {
        showToast("No active diagram to export", true);
        return;
    }
    
    // Build a clean stylesheet to embed in the SVG so fonts and colors transfer correctly
    const styles = Array.from(document.styleSheets)
        .map(styleSheet => {
            try {
                return Array.from(styleSheet.cssRules)
                    .map(rule => rule.cssText)
                    .join('\n');
            } catch (e) {
                return '';
            }
        })
        .join('\n');
    
    const svgClone = svgEl.cloneNode(true);
    const styleElement = document.createElement('style');
    styleElement.textContent = styles;
    svgClone.insertBefore(styleElement, svgClone.firstChild);
    
    const svgString = new XMLSerializer().serializeToString(svgClone);
    const blob = new Blob([svgString], { type: 'image/svg+xml;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    
    const link = document.createElement('a');
    link.href = url;
    link.download = 'pdm_diagram.svg';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
    showToast("SVG Exported Successfully");
}

function downloadPNG() {
    const svgEl = document.querySelector('#diagram svg');
    if (!svgEl) {
        showToast("No active diagram to export", true);
        return;
    }

    // High-quality render: extract stylesheet and clone SVG
    const styles = Array.from(document.styleSheets)
        .map(styleSheet => {
            try {
                return Array.from(styleSheet.cssRules)
                    .map(rule => rule.cssText)
                    .join('\n');
            } catch (e) {
                return '';
            }
        })
        .join('\n');
    
    const svgClone = svgEl.cloneNode(true);
    const styleElement = document.createElement('style');
    styleElement.textContent = styles;
    svgClone.insertBefore(styleElement, svgClone.firstChild);
    
    const svgString = new XMLSerializer().serializeToString(svgClone);
    const svgSize = svgEl.getBoundingClientRect();
    
    const img = new Image();
    const svgBlob = new Blob([svgString], { type: 'image/svg+xml;charset=utf-8' });
    const url = URL.createObjectURL(svgBlob);
    
    img.onload = function() {
        const canvas = document.createElement('canvas');
        const scale = 2; // Increase scale for HD resolution
        canvas.width = svgSize.width * scale;
        canvas.height = svgSize.height * scale;
        
        const ctx = canvas.getContext('2d');
        ctx.scale(scale, scale);
        
        // Draw theme background
        const computedStyle = window.getComputedStyle(document.body);
        ctx.fillStyle = computedStyle.getPropertyValue('--bg-viewport').trim() || '#0b0f19';
        ctx.fillRect(0, 0, svgSize.width, svgSize.height);
        
        ctx.drawImage(img, 0, 0);
        
        const pngUrl = canvas.toDataURL('image/png');
        const link = document.createElement('a');
        link.href = pngUrl;
        link.download = 'pdm_diagram.png';
        document.body.appendChild(link);
        link.click();
        document.body.removeChild(link);
        
        URL.revokeObjectURL(url);
        showToast("PNG Exported Successfully");
    };
    
    img.onerror = function() {
        showToast("Error exporting PNG (foreignObject restrictions)", true);
    };
    
    img.src = url;
}

// Notification Toast triggers
function showToast(message, isError = false) {
    const toast = document.getElementById('toast');
    const toastMsg = document.getElementById('toastMessage');
    const toastIcon = document.getElementById('toastIcon');
    
    toastMsg.innerText = message;
    if (isError) {
        toast.classList.add('error');
        toastIcon.innerHTML = `<path d="M10.29 3.86L1.82 18a2 2 0 0 0 1.71 3h16.94a2 2 0 0 0 1.71-3L13.71 3.86a2 2 0 0 0-3.42 0z"/><line x1="12" y1="9" x2="12" y2="13"/><line x1="12" y1="17" x2="12.01" y2="17"/>`;
    } else {
        toast.classList.remove('error');
        toastIcon.innerHTML = `<circle cx="12" cy="12" r="10"/><path d="m9 12 2 2 4-4"/>`;
    }
    
    toast.classList.add('show');
    setTimeout(() => {
        toast.classList.remove('show');
    }, 3000);
}

// Initial run on window load
window.onload = function() {
    generatePDM();
};