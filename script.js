// متغيرات لتخزين البيانات وكائنات الشارتات
let originalData = []; 
let myCharts = {};    

// --- الإعدادات الجديدة للربط التلقائي ---
const API_KEY = 'AIzaSyANZIOWxCtxAcLVrFebSRF9xk93DvrUiCs';
const SPREADSHEET_ID = '1avAzf7ROjVAy43_yDTfppUAhg6JdM191_wGeLOfICWA';
const SHEET_NAME = 'Sheet1'; // تأكد أن هذا هو اسم الورقة داخل الملف
const RANGE = `${SHEET_NAME}!A:Z`; 

// 1. المحرك الرئيسي الجديد: جلب البيانات تلقائياً من Google Sheets عند تحميل الصفحة
window.addEventListener('DOMContentLoaded', function() {
    fetchLiveSheetData();
});

async function fetchLiveSheetData() {
    console.log("Checking for latest data from Google Sheets...");
    const url = `https://sheets.googleapis.com/v4/spreadsheets/${SPREADSHEET_ID}/values/${RANGE}?key=${API_KEY}`;

    try {
        const response = await fetch(url);
        const data = await response.json();

        if (data.values && data.values.length > 1) {
            const headers = data.values[0]; // الصف الأول هو العناوين
            const rows = data.values.slice(1); // باقي الصفوف هي البيانات

            // تحويل المصفوفة إلى تنسيق JSON (Object) بنفس أسماء الأعمدة القديمة
            originalData = rows.map(row => {
                let item = {};
                headers.forEach((header, index) => {
                    item[header] = row[index] || ""; 
                });
                return item;
            });

            console.log("✅ Data Loaded Successfully:", originalData.length, "rows");
            
            // تشغيل الداشبورد تلقائياً
            populateSlicers(originalData);
            updateDashboard(originalData);
            
        } else {
            console.error("لم يتم العثور على بيانات في الشيت.");
            alert("تنبيه: شيت البيانات فارغ حالياً.");
        }
    } catch (error) {
        console.error("❌ Error fetching data:", error);
        alert("فشل الاتصال التلقائي بقاعدة البيانات. تأكد من إعدادات الـ API.");
    }
}

// دالة معالجة التواريخ (تم تعديلها لتناسب نص التاريخ القادم من جوجل شيت)
function parseExcelDate(dateValue) {
    if (!dateValue) return null;
    const date = new Date(dateValue);
    return isNaN(date.getTime()) ? null : date;
}

// 2. دالة تحديث جميع العناصر
function updateDashboard(data) {
    renderStatusDonut(data);
    renderLineChart(data);
    renderFunnel('qualityFunnel', data, 'فى حاله كانت الشكوى جوده برجاء تحديد نوع الشكوى');
    renderFunnel('typeFunnel', data, 'نوع الشكوى');
    renderFunnel('productFunnel', data, 'الشكوى على اي منتج؟');
    
    renderLeaderboard(data); 
    renderRawDataTable(data); 
}

// 3. نظام الفلاتر الذكي (Slicers)
function populateSlicers(data) {
    const configs = [
        { id: 'branchSlicer', column: 'اختر الفرع', label: 'جميع الفروع' },
        { id: 'typeSlicer', column: 'نوع الشكوى', label: 'الكل' },
        { id: 'managerSlicer', column: 'مدير المنطقة المسؤول', label: 'الكل' }
    ];

    configs.forEach(conf => {
        const slicer = document.getElementById(conf.id);
        if (!slicer) return;

        const uniqueValues = [...new Set(data.map(item => item[conf.column]))]
            .filter(val => {
                if (!val) return false;
                if (conf.id === 'branchSlicer') return isNaN(val); 
                return true;
            });
        
        slicer.innerHTML = `<option value="all">${conf.label}</option>`;
        uniqueValues.sort().forEach(val => {
            const option = document.createElement('option');
            option.value = val;
            option.textContent = val;
            slicer.appendChild(option);
        });

        slicer.onchange = filterData;
    });

    flatpickr("#datePicker", {
        mode: "range",
        dateFormat: "Y-m-d",
        onClose: function(selectedDates) {
            if (selectedDates.length === 2) {
                filterData();
            }
        }
    });
}

// 4. منطق الفلترة التراكمي
function filterData() {
    const branchVal = document.getElementById('branchSlicer').value;
    const typeVal = document.getElementById('typeSlicer').value;
    const managerVal = document.getElementById('managerSlicer').value;
    
    const datePicker = document.getElementById('datePicker')._flatpickr;
    const selectedDates = datePicker ? datePicker.selectedDates : [];

    let filtered = originalData;

    if (branchVal !== 'all') filtered = filtered.filter(item => item['اختر الفرع'] === branchVal);
    if (typeVal !== 'all') filtered = filtered.filter(item => item['نوع الشكوى'] === typeVal);
    if (managerVal !== 'all') filtered = filtered.filter(item => item['مدير المنطقة المسؤول'] === managerVal);

    if (selectedDates.length === 2) {
        const start = new Date(selectedDates[0]);
        const end = new Date(selectedDates[1]);
        start.setHours(0, 0, 0, 0);
        end.setHours(23, 59, 59, 999);

        filtered = filtered.filter(item => {
            const itemDate = parseExcelDate(item['التاريخ']);
            if (!itemDate) return false;
            const compareDate = new Date(itemDate);
            compareDate.setHours(0, 0, 0, 0);
            return compareDate >= start && compareDate <= end;
        });
    }

    updateDashboard(filtered);
}

// إعادة الضبط (Reset Filters)
function resetFilters() {
    const ids = ['branchSlicer', 'typeSlicer', 'managerSlicer'];
    ids.forEach(id => {
        const el = document.getElementById(id);
        if (el) el.value = 'all';
    });

    const datePickerElem = document.getElementById('datePicker');
    if (datePickerElem && datePickerElem._flatpickr) {
        datePickerElem._flatpickr.clear();
    }

    updateDashboard(originalData);
}

// 5. رسم الـ Donut Chart
function renderStatusDonut(data) {
    const chartDom = document.getElementById('statusDonutChart');
    if (!chartDom) return;
    
    let myChart = echarts.getInstanceByDom(chartDom);
    if (myChart) { myChart.dispose(); }
    myChart = echarts.init(chartDom);

    const uniqueIds = [...new Set(data.map(item => item['INDEX']))];
    const totalBadge = document.getElementById('totalCasesNumber');
    if (totalBadge) totalBadge.textContent = uniqueIds.length.toLocaleString();

    const seenIds = new Set();
    const counts = {};
    let totalUniqueCount = 0;

    data.forEach(r => {
        if (!seenIds.has(r['INDEX'])) {
            const status = r['مدى الاجراء المتخذ'] || 'غير محدد';
            counts[status] = (counts[status] || 0) + 1;
            seenIds.add(r['INDEX']);
            totalUniqueCount++;
        }
    });

    const statusNames = Object.keys(counts);
    const baseColors = ['#8B0000', '#B22222', '#CD5C5C', '#E9967A', '#F08080', '#FFE4E1'];

    const chartData = statusNames.map((name, index) => ({
        name: name,
        value: counts[name],
        itemStyle: { color: baseColors[index] || '#4a4a4a' }
    }));

    const option = {
        tooltip: { 
            trigger: 'item', 
            confine: true,
            formatter: '{b}: <b>{c}</b> ({d}%)'
        },
        legend: { 
            orient: 'horizontal', 
            top: '5',             // مسافة بسيطة من أعلى الكارت
            left: 'center', 
            type: 'scroll',
            textStyle: { color: '#ccc', fontSize: 11 },
            pageIconColor: '#d32f2f',
            // إضافة النسبة المئوية بجانب كل اسم في الـ Legend
            formatter: function(name) {
                const item = chartData.find(d => d.name === name);
                const p = totalUniqueCount > 0 ? ((item.value / totalUniqueCount) * 100).toFixed(1) : 0;
                return `${name} (${p}%)`;
            }
        },
        series: [{
            name: 'حالة الإجراء',
            type: 'pie',
            radius: ['45%', '70%'], 
            center: ['50%', '60%'], // نزلنا الشارت شوية لتحت عشان الـ Legend اللي بقى واخد مساحة فوق
            avoidLabelOverlap: true,
            itemStyle: { borderRadius: 8, borderColor: '#242426', borderWidth: 2 },
            label: { show: false }, 
            data: chartData
        }]
    };
    
    myChart.setOption(option);
    myCharts['donut'] = myChart;
}

// 6. رسم الـ Line Chart (نسخة مصححة ومجربة)
function renderLineChart(data) {
    const chartDom = document.getElementById('lineChart');
    if (!chartDom) return;
    const myChart = echarts.init(chartDom);

    const timeData = {};
    const seenInDay = new Set();

    data.forEach(r => {
        let itemDate = parseExcelDate(r['التاريخ']);
        if (itemDate) {
            let dateStr = itemDate.toLocaleDateString('en-CA'); 
            let uniqueKey = `${dateStr}_${r['INDEX']}`;

            if (!seenInDay.has(uniqueKey)) {
                timeData[dateStr] = (timeData[dateStr] || 0) + 1;
                seenInDay.add(uniqueKey);
            }
        }
    });

    const sortedDates = Object.keys(timeData).sort();
    const values = sortedDates.map(d => timeData[d]);

    const option = {
        tooltip: { 
            trigger: 'axis', 
            backgroundColor: 'rgba(20, 20, 20, 0.9)', 
            textStyle: { color: '#eee' },
            confine: true 
        },
        grid: { left: '3%', right: '4%', bottom: '15%', containLabel: true },
        dataZoom: [
            {
                type: 'inside', 
                start: 70,      
                end: 100
            }
        ],
        xAxis: { 
            type: 'category', 
            data: sortedDates, 
            axisTick: { show: false },
            axisLabel: { 
                color: '#aaa', 
                fontSize: 10,
                hideOverlap: true,
                formatter: (value) => {
                    const date = new Date(value);
                    const monthNames = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"];
                    return `${monthNames[date.getMonth()]} ${date.getDate()}`;
                }
            } 
        },
        yAxis: { 
            type: 'value', 
            splitLine: { lineStyle: { color: '#222' } }, 
            axisLabel: { color: '#888', fontSize: 10 } 
        },
        series: [{
            name: 'عدد الشكاوى اليومية',
            type: 'line',
            smooth: true,
            symbol: 'circle',
            showSymbol: false,
            lineStyle: { width: 2, color: '#d32f2f' },
            areaStyle: {
                opacity: 0.2,
                color: new echarts.graphic.LinearGradient(0, 0, 0, 1, [
                    { offset: 0, color: '#d32f2f' },
                    { offset: 1, color: 'transparent' }
                ])
            },
            data: values
        }]
    };

    myChart.setOption(option, true);
    myCharts['line'] = myChart;
}
// 8. رسم الـ Funnel Charts
function renderFunnel(id, data, column) {
    const chartDom = document.getElementById(id);
    if (!chartDom) return;
    const myChart = echarts.init(chartDom);

    const counts = {};
    const excludedValues = ['لا علاقة لها بالجودة', 'لا علاقة لها بالمنتج', 'لا يوجد تفاصيل من العميل', '...', 'اخرى', 'أخرى', 'غير محدد'];

    data.forEach(r => {
        const val = r[column] || 'غير محدد';
        if (!excludedValues.includes(String(val).trim())) {
            counts[val] = (counts[val] || 0) + 1;
        }
    });

    const sortedEntries = Object.entries(counts).sort((a, b) => b[1] - a[1]).slice(0, 10); 
    const chartData = sortedEntries.map(([name, value]) => ({ name, value }));
    const totalItems = chartData.length;

    const colorPalette = chartData.map((_, index) => {
        const ratio = index / (totalItems > 1 ? totalItems - 1 : 1);
        const r = Math.round(139 + (ratio * (255 - 139))); 
        const g = Math.round(0 + (ratio * 228));        
        const b = Math.round(0 + (ratio * 225));        
        return `rgb(${r}, ${g}, ${b})`;
    });

    const option = {
        tooltip: { 
            trigger: 'item', 
            formatter: '{b}: <b>{c}</b>', 
            backgroundColor: 'rgba(20, 20, 20, 0.9)', 
            textStyle: { color: '#eee' } 
        },
        series: [{
            type: 'funnel', 
            left: '35%', 
            width: '55%', 
            top: 20, 
            bottom: 20,
            label: {
                show: true, 
                position: 'left', 
                formatter: (params) => `{name|${params.name}} {count|(${params.value.toLocaleString()})}`,
                rich: { 
                    name: { color: '#eee', fontSize: 11 }, 
                    count: { color: '#aaa', fontSize: 10, fontWeight: 'bold' } 
                }
            },
            data: chartData, 
            color: colorPalette 
        }]
    };
    myChart.setOption(option, true);
    myCharts[id] = myChart;
}

// 8. جدول المتصدرين
function renderLeaderboard(data) {
    const tbody = document.querySelector('#branchLeaderboard tbody');
    if (!tbody) return;
    tbody.innerHTML = '';
    const branchStats = {};
    const seenBranchId = new Set();

    data.forEach(item => {
        const key = `${item['INDEX']}_${item['اختر الفرع']}`;
        if (!seenBranchId.has(key)) {
            const branch = item['اختر الفرع'] || 'فرع غير معروف';
            branchStats[branch] = (branchStats[branch] || 0) + 1;
            seenBranchId.add(key);
        }
    });

    const topBranches = Object.entries(branchStats).sort((a, b) => b[1] - a[1]).slice(0, 15);
    topBranches.forEach((branch, index) => {
        const tr = document.createElement('tr');
        tr.innerHTML = `<td>${String(index + 1).padStart(2, '0')}</td><td>${branch[0]}</td><td style="text-align: center;">${branch[1].toLocaleString()}</td>`;
        tbody.appendChild(tr);
    });
}

// 9. جدول البيانات الخام
function renderRawDataTable(data) {
    const tbody = document.querySelector('#complainsRawTable tbody');
    const countBadge = document.getElementById('rawTableCount');
    if (!tbody) return;
    tbody.innerHTML = '';

    const grouped = data.reduce((acc, item) => {
        const id = item['INDEX'];
        if (!acc[id]) { acc[id] = { ...item, types: new Set([item['نوع الشكوى']]) }; } 
        else { acc[id].types.add(item['نوع الشكوى']); }
        return acc;
    }, {});

    const displayData = Object.values(grouped).sort((a, b) => (parseExcelDate(b['التاريخ']) || 0) - (parseExcelDate(a['التاريخ']) || 0)).slice(0, 100);

    if (countBadge) countBadge.textContent = `${Object.keys(grouped).length.toLocaleString()} شكوى فريدة`;

    displayData.forEach((item) => {
        const tr = document.createElement('tr');
        const dObj = parseExcelDate(item['التاريخ']);
        const formattedDate = dObj ? dObj.toLocaleDateString('ar-EG') : '-';
        tr.innerHTML = `<td>${item['INDEX']}</td><td>${item['اختر الفرع'] || '-'}</td><td>${item['مدير المنطقة المسؤول'] || '-'}</td><td>${Array.from(item.types).join(' + ')}</td><td>${item['الشكوى على اي منتج؟'] || '-'}</td><td>${item['محتوى شكوى العميل'] || '-'}</td><td>${formattedDate}</td>`;
        tbody.appendChild(tr);
    });
}

window.onresize = function() { Object.values(myCharts).forEach(chart => chart.resize()); };
