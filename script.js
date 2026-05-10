// Регистрируем плагин для цифр внутри круга
Chart.register(ChartDataLabels);

let chartInstance = null;
let showLabels = true;
let rawDataRows = []; // Переменная для хранения всех строк загруженного файла

// 1. Расширенный массив на 30 уникальных цветов
const pieColors = [
    '#00ff88', '#2196F3', '#f44336', '#FFEB3B', '#E91E63', 
    '#9C27B0', '#00BCD4', '#FF9800', '#795548', '#607D8B',
    '#00d2ff', '#3a7bd5', '#ff512f', '#dd2476', '#ae10ff',
    '#00b09b', '#96c93d', '#f7971e', '#ffd200', '#1e3c72',
    '#ff4b2b', '#4568dc', '#b06ab3', '#0575e6', '#00f260',
    '#f12711', '#f5af19', '#8e2de2', '#4a00e0', '#f00000'
];

// Плагин для отображения чисел над столбцами (Bar/Line)
const labelsPlugin = {
    id: 'labelsPlugin',
    afterDatasetsDraw(chart) {
        if (!showLabels || chart.config.type === 'pie') return;
        const { ctx, data } = chart;
        ctx.save();
        ctx.font = 'bold 15px Segoe UI';
        ctx.fillStyle = '#ffffff';
        ctx.textAlign = 'center';

        chart.getDatasetMeta(0).data.forEach((datapoint, index) => {
            const value = data.datasets[0].data[index];
            ctx.fillText(value, datapoint.x, datapoint.y - 12);
        });
        ctx.restore();
    }
};

// СЛУШАТЕЛЬ ЗАГРУЗКИ
document.getElementById('upload').addEventListener('change', function(event) {
    const file = event.target.files[0];
    if (!file) return;
    
    document.getElementById('main-title').innerText = file.name.replace(/\.[^/.]+$/, "");
    
    const fileName = file.name.toLowerCase();
    const reader = new FileReader();

    if (fileName.endsWith('.csv')) {
        reader.onload = function(e) {
            // Читаем CSV с учетом кириллицы
            const text = e.target.result;
            const lines = text.trim().split('\n');
            rawDataRows = lines.map(line => line.split(','));
            
            fillSelectors(rawDataRows[0]);
            applyColumnChange();
        };
        reader.readAsText(file, 'windows-1251'); 
    } 
    else if (fileName.endsWith('.xlsx') || fileName.endsWith('.xls')) {
        reader.onload = function(e) {
            const dataArr = new Uint8Array(e.target.result);
            const workbook = XLSX.read(dataArr, {type: 'array'});
            const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
            
            // Сохраняем все данные в массив
            rawDataRows = XLSX.utils.sheet_to_json(firstSheet, {header: 1});
            
            if (rawDataRows.length > 0) {
                fillSelectors(rawDataRows[0]);
                applyColumnChange();
            }
        };
        reader.readAsArrayBuffer(file);
    }
});

// Заполнение выпадающих списков заголовками столбцов
function fillSelectors(headers) {
    const lSelect = document.getElementById('label-select');
    const vSelect = document.getElementById('value-select');
    lSelect.innerHTML = ''; vSelect.innerHTML = '';
    
    headers.forEach((h, i) => {
        let title = h ? h.toString().trim() : "Столбец " + (i + 1);
        let opt = `<option value="${i}">${title}</option>`;
        lSelect.innerHTML += opt;
        vSelect.innerHTML += opt;
    });
    
    lSelect.selectedIndex = 0;
    vSelect.selectedIndex = 1; // Предполагаем, что цифры во второй колонке
    document.getElementById('column-selectors').style.display = 'flex';
}

// Применение изменений при выборе нового столбца
function applyColumnChange() {
    const lIdx = parseInt(document.getElementById('label-select').value);
    const vIdx = parseInt(document.getElementById('value-select').value);
    
    const labels = [], values = [];
    let total = 0;

    // Пропускаем первую строку (заголовки)
    rawDataRows.slice(1).forEach(row => {
        if (row[vIdx] !== undefined) {
            const val = parseFloat(row[vIdx].toString().replace(/\s/g, '').replace(',', '.'));
            if (!isNaN(val)) {
                labels.push(row[lIdx] ? row[lIdx].toString().trim() : "");
                values.push(val);
                total += val;
            }
        }
    });

    const data = {
        labels, 
        values, 
        total, 
        max: values.length ? Math.max(...values) : 0, 
        min: values.length ? Math.min(...values) : 0
    };

    updateStats(data);
    const currentType = chartInstance ? chartInstance.config.type : 'bar';
    drawChart(data.labels, data.values, currentType);
}

function updateStats(data) {
    document.getElementById('stats-panel').style.display = 'grid';
    document.getElementById('total-val').innerText = data.total.toLocaleString();
    document.getElementById('max-val').innerText = data.max.toLocaleString();
    document.getElementById('min-val').innerText = data.min.toLocaleString();
}

function drawChart(labels, values, type) {
    const ctx = document.getElementById('myChart').getContext('2d');
    if (chartInstance) chartInstance.destroy();

    chartInstance = new Chart(ctx, {
        type: type,
        data: {
            labels: labels,
            datasets: [{
                data: values,
                borderColor: type === 'pie' ? '#1e1e1e' : '#00ff88',
                backgroundColor: type === 'pie' ? pieColors : 'rgba(0, 255, 136, 0.25)',
                borderWidth: 2,
                tension: 0.3,
                fill: true
            }]
        },
        options: {
            responsive: true,
            maintainAspectRatio: false,
            layout: { padding: { top: 40, bottom: 10 } },
            plugins: {
                datalabels: {
                    display: function() { return showLabels && type === 'pie'; },
                    color: '#000000',
                    font: { weight: 'bold', size: 14 },
                    formatter: (value) => value
                },
                legend: { 
                    display: type === 'pie', 
                    position: 'top',
                    labels: { color: '#fff', font: { size: 14, weight: 'bold' } } 
                },
                tooltip: { enabled: true }
            },
            scales: type === 'pie' ? {} : {
                y: { 
                    beginAtZero: true, grace: '15%', 
                    grid: { color: '#333' }, 
                    ticks: { color: '#aaa' } 
                },
                x: { 
                    grid: { color: '#333' }, 
                    ticks: { color: '#aaa' } 
                }
            }
        },
        plugins: [labelsPlugin]
    });
}

function toggleNumbers() {
    showLabels = !showLabels;
    document.getElementById('toggle-labels').innerText = showLabels ? "Цифры: ВКЛ" : "Цифры: ВЫКЛ";
    if (chartInstance) chartInstance.update();
}

function updateChartType(newType) {
    if (chartInstance) drawChart(chartInstance.data.labels, chartInstance.data.datasets[0].data, newType);
}

function downloadChart() {
    if (chartInstance) {
        const link = document.createElement('a');
        link.download = document.getElementById('main-title').innerText + '.png';
        link.href = document.getElementById('myChart').toDataURL('image/png');
        link.click();
    }
}

function resetChart() { location.reload(); }