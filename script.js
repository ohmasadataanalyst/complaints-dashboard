// =============================================
// CONFIG
// =============================================
const API_KEY = 'AIzaSyANZIOWxCtxAcLVrFebSRF9xk93DvrUiCs';
const SPREADSHEET_ID = '1avAzf7ROjVAy43_yDTfppUAhg6JdM191_wGeLOfICWA';
const SHEET_NAME = 'Sheet1';
const RANGE = SHEET_NAME + '!A:Z';

let originalData = [];
let myCharts = {};

const multiSelectState = {
    branchSlicer: new Set(),
    typeSlicer: new Set(),
    productSlicer: new Set(),
    managerSlicer: new Set(),
    qualitySlicer: new Set(),
    foodSlicer: new Set()
};

// Comparison modal filter state (separate from main filters)
const compFilterState = {
    branch: new Set(),
    type: new Set(),
    product: new Set(),
    manager: new Set(),
    quality: new Set(),
    food: new Set()
};

const slicerConfigs = [
    { id: 'branchSlicer',   column: 'اختر الفرع',              label: 'جميع الفروع', labelText: 'الفرع',         containerId: 'slicer-branch',   filterBranch: true },
    { id: 'typeSlicer',     column: 'نوع الشكوى',              label: 'الكل',        labelText: 'نوع الشكوى',    containerId: 'slicer-type' },
    { id: 'productSlicer',  column: 'الشكوى على اي منتج؟',    label: 'الكل',        labelText: 'المنتج',         containerId: 'slicer-product' },
    { id: 'managerSlicer',  column: 'مدير المنطقة المسؤول',                                                                    label: 'الكل', labelText: 'مدير المنطقة',       containerId: 'slicer-manager' },
    { id: 'qualitySlicer', column: 'فى حاله كانت الشكوى جوده برجاء تحديد نوع الشكوى',                                              label: 'الكل', labelText: 'نوع شكوى الجودة',    containerId: 'slicer-quality', excludeValues: ['لا علاقة لها بالجودة'] },
    { id: 'foodSlicer',    column: 'في حال كانت الشكوى تخص التلوث الغذائي (السلامة الغذائية) يرجى اختيار الكود',                   label: 'الكل', labelText: 'كود السلامة الغذائية', containerId: 'slicer-food' }
];

// =============================================
// INIT
// =============================================
window.addEventListener('DOMContentLoaded', function() {
    initMultiSelectDropdowns();
    fetchLiveSheetData();
});

async function fetchLiveSheetData() {
    var url = 'https://sheets.googleapis.com/v4/spreadsheets/' + SPREADSHEET_ID + '/values/' + RANGE + '?key=' + API_KEY;
    try {
        var response = await fetch(url);
        var data = await response.json();
        if (data.values && data.values.length > 1) {
            var headers = data.values[0];
            var rows = data.values.slice(1);
            originalData = rows.map(function(row) {
                var item = {};
                headers.forEach(function(h, i) { item[h] = row[i] || ''; });
                return item;
            });
            console.log('Data loaded:', originalData.length, 'rows');
            populateSlicers(originalData);
            updateDashboard(originalData);
        } else {
            alert('تنبيه: شيت البيانات فارغ حالياً.');
        }
    } catch(e) {
        console.error(e);
        alert('تعذر الاتصال بـ Google Sheets. تحقق من إعدادات الـ API.');
    }
}

function parseExcelDate(v) {
    if (!v) return null;
    var d = new Date(v);
    return isNaN(d.getTime()) ? null : d;
}

function updateDashboard(data) {
    renderStatusDonut(data);
    renderLineChart(data);
    renderFunnel('qualityFunnel', data, 'فى حاله كانت الشكوى جوده برجاء تحديد نوع الشكوى');
    renderFunnel('typeFunnel', data, 'نوع الشكوى');
    renderFunnel('productFunnel', data, 'الشكوى على اي منتج؟');
    renderLeaderboard(data);
    renderRawDataTable(data);
    renderCompensation(data);
}

// =============================================
// MULTI-SELECT LOGIC
// =============================================
function initMultiSelectDropdowns() {
    document.addEventListener('click', function(e) {
        if (!e.target.closest('.multi-select-wrapper')) {
            document.querySelectorAll('.ms-dropdown').forEach(function(d) { d.classList.remove('open'); });
        }
    });
}

function buildMultiSelectHTML(id, label) {
    return '<div class="multi-select-wrapper" id="wrapper-' + id + '">' +
        '<div class="ms-trigger" onclick="toggleMsDropdown(\'' + id + '\')">' +
            '<span class="ms-label" id="label-' + id + '">' + label + '</span>' +
            '<span class="ms-arrow">▾</span>' +
        '</div>' +
        '<div class="ms-dropdown" id="dropdown-' + id + '">' +
            '<div class="ms-search-wrap"><input class="ms-search" type="text" placeholder="بحث..." oninput="filterMsDropdown(\'' + id + '\', this.value)"></div>' +
            '<div class="ms-options" id="options-' + id + '"></div>' +
            '<div class="ms-footer">' +
                '<button onclick="selectAllMs(\'' + id + '\')">تحديد الكل</button>' +
                '<button onclick="clearAllMs(\'' + id + '\')">إلغاء الكل</button>' +
            '</div>' +
        '</div>' +
    '</div>';
}

function toggleMsDropdown(id) {
    var d = document.getElementById('dropdown-' + id);
    var isOpen = d.classList.contains('open');
    document.querySelectorAll('.ms-dropdown').forEach(function(x) { x.classList.remove('open'); });
    if (!isOpen) d.classList.add('open');
}

function filterMsDropdown(id, q) {
    q = q.toLowerCase();
    document.querySelectorAll('#options-' + id + ' .ms-option').forEach(function(opt) {
        opt.style.display = opt.dataset.value.toLowerCase().includes(q) ? '' : 'none';
    });
}

function selectAllMs(id) {
    document.querySelectorAll('#options-' + id + ' .ms-option').forEach(function(opt) {
        if (opt.style.display !== 'none') { multiSelectState[id].add(opt.dataset.value); opt.classList.add('selected'); }
    });
    updateMsLabel(id); filterData();
}

function clearAllMs(id) {
    multiSelectState[id].clear();
    document.querySelectorAll('#options-' + id + ' .ms-option').forEach(function(o) { o.classList.remove('selected'); });
    updateMsLabel(id); filterData();
}

function toggleMsOption(id, value, el) {
    if (multiSelectState[id].has(value)) { multiSelectState[id].delete(value); el.classList.remove('selected'); }
    else { multiSelectState[id].add(value); el.classList.add('selected'); }
    updateMsLabel(id); filterData();
}

function updateMsLabel(id) {
    var el = document.getElementById('label-' + id);
    if (!el) return;
    var count = multiSelectState[id].size;
    var defaults = { branchSlicer:'جميع الفروع', typeSlicer:'الكل', productSlicer:'الكل', managerSlicer:'الكل', qualitySlicer:'الكل', foodSlicer:'الكل' };
    el.textContent = count === 0 ? defaults[id] : ('تم اختيار ' + count + ' عنصر');
    el.style.color = count > 0 ? '#d32f2f' : '';
}

function populateMsOptions(id, values) {
    var c = document.getElementById('options-' + id);
    if (!c) return;
    c.innerHTML = '';
    values.sort().forEach(function(val) {
        var div = document.createElement('div');
        div.className = 'ms-option';
        div.dataset.value = val;
        div.textContent = val;
        div.onclick = function() { toggleMsOption(id, val, div); };
        c.appendChild(div);
    });
}

function populateSlicers(data) {
    slicerConfigs.forEach(function(conf) {
        var container = document.getElementById(conf.containerId);
        if (!container) return;
        container.innerHTML = '<label>' + conf.labelText + '</label>' + buildMultiSelectHTML(conf.id, conf.label);
        var unique = [...new Set(data.map(function(item) { return item[conf.column]; }))].filter(function(val) {
            if (!val) return false;
            if (conf.filterBranch) return isNaN(val);
            if (conf.excludeValues && conf.excludeValues.includes(val)) return false;
            return true;
        });
        populateMsOptions(conf.id, unique);
    });

    flatpickr('#datePicker', {
        mode: 'range', dateFormat: 'Y-m-d',
        onClose: function(sel) { if (sel.length === 2) filterData(); }
    });
}

// =============================================
// FILTERING
// =============================================
function filterData() {
    var filtered = originalData;

    slicerConfigs.forEach(function(conf) {
        var sel = multiSelectState[conf.id];
        if (sel.size > 0) {
            filtered = filtered.filter(function(item) { return sel.has(item[conf.column]); });
        }
    });

    var dpEl = document.getElementById('datePicker');
    var dp = dpEl && dpEl._flatpickr;
    var selDates = dp ? dp.selectedDates : [];
    if (selDates.length === 2) {
        var start = new Date(selDates[0]); start.setHours(0,0,0,0);
        var end = new Date(selDates[1]); end.setHours(23,59,59,999);
        filtered = filtered.filter(function(item) {
            var d = parseExcelDate(item['التاريخ']);
            if (!d) return false;
            var c = new Date(d); c.setHours(0,0,0,0);
            return c >= start && c <= end;
        });
    }
    updateDashboard(filtered);
}

function resetFilters() {
    slicerConfigs.forEach(function(conf) {
        multiSelectState[conf.id].clear();
        document.querySelectorAll('#options-' + conf.id + ' .ms-option').forEach(function(o) { o.classList.remove('selected'); });
        updateMsLabel(conf.id);
    });
    var dpEl = document.getElementById('datePicker');
    if (dpEl && dpEl._flatpickr) dpEl._flatpickr.clear();
    updateDashboard(originalData);
}

// =============================================
// COMPENSATION CARD
// =============================================
function renderCompensation(data) {
    var valEl = document.getElementById('compValue');
    var subEl = document.getElementById('compSub');
    var badgeEl = document.getElementById('compBadge');
    if (!valEl) return;

    var colKey = 'قيمة التعويض';
    var total = 0; var count = 0; var hasData = false;
    var seenIds = new Set();

    data.forEach(function(r) {
        if (r['INDEX'] && !seenIds.has(r['INDEX'])) {
            seenIds.add(r['INDEX']);
            var raw = r[colKey];
            if (raw && raw !== '') {
                var num = parseFloat(String(raw).replace(/,/g,''));
                if (!isNaN(num)) { total += num; count++; hasData = true; }
            }
        }
    });

    if (!hasData) {
        valEl.textContent = 'لا توجد بيانات';
        subEl.textContent = 'عمود "قيمة التعويض" غير موجود في الشيت بعد';
        badgeEl.textContent = '—';
        valEl.style.color = '#666';
    } else {
        valEl.textContent = total.toLocaleString('ar-EG') + ' ر.س';
        valEl.style.color = '#d32f2f';
        subEl.textContent = 'بناءً على ' + count.toLocaleString() + ' حالة لديها قيمة تعويض مسجلة';
        var anyFilter = slicerConfigs.some(function(c) { return multiSelectState[c.id].size > 0; });
        var dpEl = document.getElementById('datePicker');
        var dp = dpEl && dpEl._flatpickr;
        var dateFiltered = dp && dp.selectedDates.length === 2;
        badgeEl.textContent = (anyFilter || dateFiltered) ? 'فلاتر مطبقة' : 'كل الفترات';
    }
}

// =============================================
// DONUT CHART
// =============================================
function renderStatusDonut(data) {
    var chartDom = document.getElementById('statusDonutChartV2');
    if (!chartDom) return;
    var myChart = echarts.getInstanceByDom(chartDom);
    if (myChart) myChart.dispose();
    myChart = echarts.init(chartDom);

    var seenIds = new Set(); var counts = {}; var total = 0;
    data.forEach(function(r) {
        if (r['INDEX'] && !seenIds.has(r['INDEX'])) {
            var s = r['مدى الاجراء المتخذ'] || 'غير محدد';
            counts[s] = (counts[s]||0)+1; seenIds.add(r['INDEX']); total++;
        }
    });

    var baseColors = ['#8B0000','#B22222','#CD5C5C','#E9967A','#F08080','#FFE4E1'];
    var chartData = Object.keys(counts).map(function(name,i) {
        return { name:name, value:counts[name], itemStyle:{ color: baseColors[i]||'#4a4a4a' } };
    });

    myChart.setOption({
        tooltip:{ trigger:'item', confine:true, backgroundColor:'rgba(30,30,30,0.9)', borderWidth:0, textStyle:{color:'#fff',fontSize:12},
            formatter:function(p){ return p.marker+' '+p.name+': <b>'+p.value.toLocaleString()+'</b> ('+p.percent+'%)'; }},
        legend:{ orient:'horizontal', top:'5%', left:'center', type:'scroll', textStyle:{color:'#ccc',fontSize:11} },
        graphic:[{ type:'group', left:'center', top:'48%', children:[
            { type:'text', left:'center', top:'middle', style:{fill:'#d32f2f', text:total.toLocaleString(), font:'bold 28px sans-serif'} },
            { type:'text', left:'center', top:35, style:{fill:'#aaa', text:'إجمالي الحالات', font:'13px sans-serif'} }
        ]}],
        series:[{ type:'pie', center:['50%','55%'], radius:['50%','75%'], avoidLabelOverlap:false,
            itemStyle:{borderRadius:8, borderColor:'#242426', borderWidth:2}, label:{show:false}, data:chartData }]
    }, true);
    myCharts['donut'] = myChart;
}

// =============================================
// LINE CHART
// =============================================
function renderLineChart(data) {
    var chartDom = document.getElementById('lineChart');
    if (!chartDom) return;
    var mc = echarts.getInstanceByDom(chartDom); if (mc) mc.dispose();
    var myChart = echarts.init(chartDom);

    var timeData = {}; var seen = new Set();
    data.forEach(function(r) {
        var d = parseExcelDate(r['التاريخ']);
        if (d) {
            var ds = d.toLocaleDateString('en-CA');
            var uk = ds+'_'+r['INDEX'];
            if (!seen.has(uk)) { timeData[ds]=(timeData[ds]||0)+1; seen.add(uk); }
        }
    });
    var sortedDates = Object.keys(timeData).sort();
    var values = sortedDates.map(function(d){ return timeData[d]; });
    var mn = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

    myChart.setOption({
        tooltip:{trigger:'axis', backgroundColor:'rgba(20,20,20,0.9)', textStyle:{color:'#eee'}, confine:true},
        grid:{left:'3%',right:'4%',bottom:'15%',containLabel:true},
        dataZoom:[{type:'inside',start:70,end:100}],
        xAxis:{type:'category', data:sortedDates, axisTick:{show:false},
            axisLabel:{color:'#aaa',fontSize:10,hideOverlap:true, formatter:function(v){ var d=new Date(v); return mn[d.getMonth()]+' '+d.getDate(); }}},
        yAxis:{type:'value', splitLine:{lineStyle:{color:'#222'}}, axisLabel:{color:'#888',fontSize:10}},
        series:[{name:'عدد الشكاوى اليومية', type:'line', smooth:true, symbol:'circle', showSymbol:false,
            lineStyle:{width:2,color:'#d32f2f'},
            areaStyle:{opacity:0.2, color:new echarts.graphic.LinearGradient(0,0,0,1,[{offset:0,color:'#d32f2f'},{offset:1,color:'transparent'}])},
            data:values}]
    }, true);
    myCharts['line'] = myChart;
}

// =============================================
// FUNNELS
// =============================================
function renderFunnel(id, data, column) {
    var chartDom = document.getElementById(id);
    if (!chartDom) return;
    var mc = echarts.getInstanceByDom(chartDom); if (mc) mc.dispose();
    var myChart = echarts.init(chartDom);

    var counts = {};
    var excluded = ['لا علاقة لها بالجودة','لا علاقة لها بالمنتج','لا يوجد تفاصيل من العميل','...','اخرى','أخرى','غير محدد'];
    data.forEach(function(r) {
        var val = r[column]||'غير محدد';
        if (!excluded.includes(String(val).trim())) counts[val]=(counts[val]||0)+1;
    });

    var sorted = Object.entries(counts).sort(function(a,b){return b[1]-a[1];}).slice(0,10);
    var chartData = sorted.map(function(e){ return {name:e[0],value:e[1]}; });
    var n = chartData.length;
    var colors = chartData.map(function(_,i){
        var ratio = i/(n>1?n-1:1);
        return 'rgb('+Math.round(139+(ratio*116))+','+Math.round(ratio*228)+','+Math.round(ratio*225)+')';
    });

    myChart.setOption({
        tooltip:{trigger:'item', formatter:'{b}: <b>{c}</b>', backgroundColor:'rgba(20,20,20,0.9)', textStyle:{color:'#eee'}},
        series:[{type:'funnel', left:'35%', width:'55%', top:20, bottom:20,
            label:{show:true, position:'left',
                formatter:function(p){ return '{name|'+p.name+'} {count|('+p.value.toLocaleString()+')}'  },
                rich:{name:{color:'#eee',fontSize:11}, count:{color:'#aaa',fontSize:10,fontWeight:'bold'}}},
            data:chartData, color:colors}]
    }, true);
    myCharts[id] = myChart;
}

// =============================================
// LEADERBOARD
// =============================================
function renderLeaderboard(data) {
    var tbody = document.querySelector('#branchLeaderboard tbody');
    if (!tbody) return;
    tbody.innerHTML = '';
    var stats = {}; var seen = new Set();
    data.forEach(function(item) {
        var k = item['INDEX']+'_'+item['اختر الفرع'];
        if (!seen.has(k)) {
            var b = item['اختر الفرع']||'فرع غير معروف';
            stats[b]=(stats[b]||0)+1; seen.add(k);
        }
    });
    Object.entries(stats).sort(function(a,b){return b[1]-a[1];}).slice(0,15).forEach(function(b,i){
        var tr = document.createElement('tr');
        tr.innerHTML = '<td>'+String(i+1).padStart(2,'0')+'</td><td>'+b[0]+'</td><td style="text-align:center;">'+b[1].toLocaleString()+'</td>';
        tbody.appendChild(tr);
    });
}

// =============================================
// RAW TABLE
// =============================================
function renderRawDataTable(data) {
    var tbody = document.querySelector('#complainsRawTable tbody');
    if (!tbody) return;
    tbody.innerHTML = '';

    var grouped = data.reduce(function(acc,item) {
        var id = item['INDEX'];
        if (!acc[id]) acc[id] = Object.assign({}, item, {types: new Set([item['نوع الشكوى']])});
        else acc[id].types.add(item['نوع الشكوى']);
        return acc;
    }, {});

    Object.values(grouped)
        .sort(function(a,b){ return (parseExcelDate(b['التاريخ'])||0)-(parseExcelDate(a['التاريخ'])||0); })
        .slice(0,100)
        .forEach(function(item) {
            var tr = document.createElement('tr');
            var d = parseExcelDate(item['التاريخ']);
            tr.innerHTML = '<td>'+item['INDEX']+'</td><td>'+(item['اختر الفرع']||'-')+'</td>' +
                '<td>'+Array.from(item.types).join(' + ')+'</td>' +
                '<td>'+(item['محتوى شكوى العميل']||'-')+'</td>' +
                '<td>'+(d ? d.toLocaleDateString('ar-EG') : '-')+'</td>';
            tbody.appendChild(tr);
        });
}

// =============================================
// PERIOD COMPARISON — FILTERS
// =============================================
var compDatePickers = {};
var compFilterSlicers = [
    { key: 'branch',   column: 'اختر الفرع',            label: 'جميع الفروع', labelText: 'الفرع',        filterBranch: true },
    { key: 'type',     column: 'نوع الشكوى',            label: 'الكل',        labelText: 'نوع الشكوى' },
    { key: 'product',  column: 'الشكوى على اي منتج؟',  label: 'الكل',        labelText: 'المنتج' },
    { key: 'manager',  column: 'مدير المنطقة المسؤول',                                                                    label: 'الكل', labelText: 'مدير المنطقة' },
    { key: 'quality',  column: 'فى حاله كانت الشكوى جوده برجاء تحديد نوع الشكوى',                                              label: 'الكل', labelText: 'نوع شكوى الجودة',    excludeValues: ['لا علاقة لها بالجودة'] },
    { key: 'food',     column: 'في حال كانت الشكوى تخص التلوث الغذائي (السلامة الغذائية) يرجى اختيار الكود',                   label: 'الكل', labelText: 'كود السلامة الغذائية' }
];

function buildCompMsHTML(key, label) {
    var id = 'comp-ms-' + key;
    return '<div class="multi-select-wrapper" id="wrapper-' + id + '">' +
        '<div class="ms-trigger" onclick="toggleMsDropdown(\'' + id + '\')">' +
            '<span class="ms-label" id="label-' + id + '">' + label + '</span>' +
            '<span class="ms-arrow">▾</span>' +
        '</div>' +
        '<div class="ms-dropdown" id="dropdown-' + id + '">' +
            '<div class="ms-search-wrap"><input class="ms-search" type="text" placeholder="بحث..." oninput="filterMsDropdown(\'' + id + '\', this.value)"></div>' +
            '<div class="ms-options" id="options-' + id + '"></div>' +
            '<div class="ms-footer">' +
                '<button onclick="compSelectAll(\'' + key + '\')">تحديد الكل</button>' +
                '<button onclick="compClearAll(\'' + key + '\')">إلغاء الكل</button>' +
            '</div>' +
        '</div>' +
    '</div>';
}

function compSelectAll(key) {
    var id = 'comp-ms-' + key;
    document.querySelectorAll('#options-' + id + ' .ms-option').forEach(function(opt) {
        if (opt.style.display !== 'none') { compFilterState[key].add(opt.dataset.value); opt.classList.add('selected'); }
    });
    updateCompMsLabel(key);
}

function compClearAll(key) {
    var id = 'comp-ms-' + key;
    compFilterState[key].clear();
    document.querySelectorAll('#options-' + id + ' .ms-option').forEach(function(o) { o.classList.remove('selected'); });
    updateCompMsLabel(key);
}

function toggleCompMsOption(key, value, el) {
    if (compFilterState[key].has(value)) { compFilterState[key].delete(value); el.classList.remove('selected'); }
    else { compFilterState[key].add(value); el.classList.add('selected'); }
    updateCompMsLabel(key);
}

function updateCompMsLabel(key) {
    var id = 'comp-ms-' + key;
    var el = document.getElementById('label-' + id);
    if (!el) return;
    var count = compFilterState[key].size;
    var conf = compFilterSlicers.find(function(c){ return c.key === key; });
    el.textContent = count === 0 ? (conf ? conf.label : 'الكل') : ('تم اختيار ' + count + ' عنصر');
    el.style.color = count > 0 ? '#42a5f5' : '';
}

function buildCompFiltersHTML() {
    var html = '<div class="comp-filters-bar">';
    compFilterSlicers.forEach(function(conf) {
        var id = 'comp-ms-' + conf.key;
        var unique = [...new Set(originalData.map(function(item){ return item[conf.column]; }))].filter(function(val){
            if (!val) return false;
            if (conf.filterBranch) return isNaN(val);
            if (conf.excludeValues && conf.excludeValues.includes(val)) return false;
            return true;
        }).sort();
        html += '<div class="comp-filter-slicer">';
        html += '<label>' + conf.labelText + '</label>';
        html += buildCompMsHTML(conf.key, conf.label);
        html += '</div>';
        // store unique values for later population
        conf._unique = unique;
    });
    html += '<button class="comp-reset-btn" onclick="resetCompFilters()">RESET</button>';
    html += '</div>';
    return html;
}

function populateCompFilterOptions() {
    compFilterSlicers.forEach(function(conf) {
        var id = 'comp-ms-' + conf.key;
        var c = document.getElementById('options-' + id);
        if (!c || !conf._unique) return;
        c.innerHTML = '';
        conf._unique.forEach(function(val) {
            var div = document.createElement('div');
            div.className = 'ms-option';
            div.dataset.value = val;
            div.textContent = val;
            var key = conf.key;
            div.onclick = function() { toggleCompMsOption(key, val, div); };
            c.appendChild(div);
        });
    });
}

function resetCompFilters() {
    compFilterSlicers.forEach(function(conf) {
        compFilterState[conf.key].clear();
        var id = 'comp-ms-' + conf.key;
        document.querySelectorAll('#options-' + id + ' .ms-option').forEach(function(o){ o.classList.remove('selected'); });
        updateCompMsLabel(conf.key);
    });
}

function applyCompFilters(data) {
    var filtered = data;
    compFilterSlicers.forEach(function(conf) {
        var sel = compFilterState[conf.key];
        if (sel.size > 0) {
            filtered = filtered.filter(function(item){ return sel.has(item[conf.column]); });
        }
    });
    return filtered;
}

// =============================================
// PERIOD COMPARISON — MAIN
// =============================================
function initComparisonDatePickers() {
    if (compDatePickers['aStart']) return;
    ['periodAStart','periodAEnd','periodBStart','periodBEnd'].forEach(function(id) {
        compDatePickers[id] = flatpickr('#'+id, { dateFormat:'Y-m-d', allowInput:false });
    });
}

function openComparisonModal() {
    document.getElementById('comparisonModal').classList.add('open');
    setTimeout(function() {
        initComparisonDatePickers();
        // Inject filter bar if not already there
        var filterContainer = document.getElementById('compFilterContainer');
        if (filterContainer && filterContainer.innerHTML.trim() === '') {
            filterContainer.innerHTML = buildCompFiltersHTML();
            populateCompFilterOptions();
        }
    }, 50);
}

function runComparison() {
    var aStart = document.getElementById('periodAStart').value;
    var aEnd   = document.getElementById('periodAEnd').value;
    var bStart = document.getElementById('periodBStart').value;
    var bEnd   = document.getElementById('periodBEnd').value;

    if (!aStart || !aEnd || !bStart || !bEnd) {
        alert('يرجى تحديد تاريخ البداية والنهاية للفترتين.');
        return;
    }

    // Apply comparison filters on top of date ranges
    var dataA = applyCompFilters(filterByDateRange(originalData, aStart, aEnd));
    var dataB = applyCompFilters(filterByDateRange(originalData, bStart, bEnd));

    var resultsEl = document.getElementById('comparisonResults');
    resultsEl.innerHTML = buildComparisonHTML(dataA, dataB, aStart, aEnd, bStart, bEnd);
    resultsEl.classList.add('visible');

    // Render the dual line chart
    setTimeout(function() {
        renderCompLineChart(dataA, dataB, aStart, aEnd, bStart, bEnd);
    }, 100);
}

// =============================================
// PERIOD COMPARISON — DUAL LINE CHART
// =============================================
function renderCompLineChart(dataA, dataB, aStart, aEnd, bStart, bEnd) {
    var chartDom = document.getElementById('compLineChart');
    if (!chartDom) return;
    var mc = echarts.getInstanceByDom(chartDom); if (mc) mc.dispose();
    var myChart = echarts.init(chartDom);

    function buildTimeSeries(data) {
        var timeData = {}; var seen = new Set();
        data.forEach(function(r) {
            var d = parseExcelDate(r['التاريخ']);
            if (d) {
                var ds = d.toLocaleDateString('en-CA');
                var uk = ds + '_' + r['INDEX'];
                if (!seen.has(uk)) { timeData[ds] = (timeData[ds]||0)+1; seen.add(uk); }
            }
        });
        return timeData;
    }

    var tdA = buildTimeSeries(dataA);
    var tdB = buildTimeSeries(dataB);

    // Normalize dates so both series align on day-offset (day 1, day 2, etc.)
    var datesA = Object.keys(tdA).sort();
    var datesB = Object.keys(tdB).sort();
    var maxLen = Math.max(datesA.length, datesB.length);
    var xAxis = [];
    for (var i = 1; i <= maxLen; i++) xAxis.push('يوم ' + i);

    var valuesA = datesA.map(function(d){ return tdA[d]; });
    var valuesB = datesB.map(function(d){ return tdB[d]; });

    var mn = ['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

    myChart.setOption({
        tooltip: {
            trigger: 'axis',
            backgroundColor: 'rgba(20,20,20,0.95)',
            textStyle: { color: '#eee' },
            confine: true,
            formatter: function(params) {
                var s = '<b>' + params[0].axisValue + '</b><br>';
                params.forEach(function(p) {
                    s += p.marker + ' ' + p.seriesName + ': <b>' + (p.value||0) + '</b><br>';
                });
                return s;
            }
        },
        legend: {
            data: [
                { name: '🔵 الفترة الأولى (' + aStart + ' → ' + aEnd + ')', textStyle: { color: '#42a5f5' } },
                { name: '🔴 الفترة الثانية (' + bStart + ' → ' + bEnd + ')', textStyle: { color: '#ef5350' } }
            ],
            top: 0,
            textStyle: { color: '#ccc', fontSize: 11 }
        },
        grid: { left: '3%', right: '4%', bottom: '10%', top: '15%', containLabel: true },
        dataZoom: [{ type: 'inside', start: 0, end: 100 }],
        xAxis: {
            type: 'category',
            data: xAxis,
            axisTick: { show: false },
            axisLabel: { color: '#aaa', fontSize: 10 }
        },
        yAxis: {
            type: 'value',
            splitLine: { lineStyle: { color: '#222' } },
            axisLabel: { color: '#888', fontSize: 10 }
        },
        series: [
            {
                name: '🔵 الفترة الأولى (' + aStart + ' → ' + aEnd + ')',
                type: 'line', smooth: true, showSymbol: false,
                lineStyle: { width: 2, color: '#42a5f5' },
                areaStyle: { opacity: 0.1, color: new echarts.graphic.LinearGradient(0,0,0,1,[{offset:0,color:'#42a5f5'},{offset:1,color:'transparent'}]) },
                data: valuesA
            },
            {
                name: '🔴 الفترة الثانية (' + bStart + ' → ' + bEnd + ')',
                type: 'line', smooth: true, showSymbol: false,
                lineStyle: { width: 2, color: '#ef5350' },
                areaStyle: { opacity: 0.1, color: new echarts.graphic.LinearGradient(0,0,0,1,[{offset:0,color:'#ef5350'},{offset:1,color:'transparent'}]) },
                data: valuesB
            }
        ]
    }, true);
    myCharts['compLine'] = myChart;
}

// =============================================
// PERIOD COMPARISON — BUILD HTML
// =============================================
function buildComparisonHTML(dataA, dataB, aStart, aEnd, bStart, bEnd) {
    var uA = getUniqueComplaints(dataA);
    var uB = getUniqueComplaints(dataB);

    var totalA = uA.length; var totalB = uB.length;
    var compA = sumColumn(uA, 'قيمة التعويض');
    var compB = sumColumn(uB, 'قيمة التعويض');
    var daysA = Math.max(1, daysDiff(aStart, aEnd));
    var daysB = Math.max(1, daysDiff(bStart, bEnd));
    var avgA = (totalA/daysA).toFixed(1);
    var avgB = (totalB/daysB).toFixed(1);

    var actionA = countBy(uA, 'مدى الاجراء المتخذ');
    var actionB = countBy(uB, 'مدى الاجراء المتخذ');
    var allActions = [...new Set([...Object.keys(actionA), ...Object.keys(actionB)])];

    var typeA = countBy(dataA, 'نوع الشكوى');
    var typeB = countBy(dataB, 'نوع الشكوى');
    var allTypes = [...new Set([...Object.keys(typeA), ...Object.keys(typeB)])].filter(function(v){ return v && v!=='غير محدد'; });

    var prodA = countBy(dataA, 'الشكوى على اي منتج؟');
    var prodB = countBy(dataB, 'الشكوى على اي منتج؟');
    var allProds = [...new Set([...Object.keys(prodA), ...Object.keys(prodB)])].filter(function(v){
        return v && !['لا علاقة لها بالمنتج','غير محدد','أخرى','اخرى'].includes(v);
    });

    // ALL branches — no slice
    var branchA = countBy(uA, 'اختر الفرع');
    var branchB = countBy(uB, 'اختر الفرع');
    var allBranches = [...new Set([...Object.keys(branchA), ...Object.keys(branchB)])].filter(function(v){ return v && isNaN(v); });

    var qualA = countBy(dataA, 'فى حاله كانت الشكوى جوده برجاء تحديد نوع الشكوى');
    var qualB = countBy(dataB, 'فى حاله كانت الشكوى جوده برجاء تحديد نوع الشكوى');
    var allQual = [...new Set([...Object.keys(qualA), ...Object.keys(qualB)])].filter(function(v){
        return v && v !== 'لا علاقة لها بالجودة';
    });

    var html = '';

    // KPI Cards
    html += '<div class="comp-section-title">📊 ملخص عام</div>';
    html += '<div class="comp-kpi-grid">';
    html += kpiCard('إجمالي الشكاوى', totalA, totalB);
    html += kpiCard('متوسط يومي', avgA, avgB);
    html += kpiCard('إجمالي التعويضات', formatComp(compA), formatComp(compB));
    html += '</div>';

    // Dual Line Chart
    html += '<div class="comp-section-title">📈 الخط الزمني للفترتين</div>';
    html += '<div id="compLineChart" style="width:100%;height:320px;"></div>';

    // By Action
    html += '<div class="comp-section-title">✅ حسب الإجراء المتخذ</div>';
    html += '<div class="comp-table-wrap"><table class="comp-table">';
    html += '<thead><tr><th>الإجراء</th><th style="color:#42a5f5;">الفترة الأولى</th><th style="color:#ef5350;">الفترة الثانية</th><th>التغيير</th></tr></thead><tbody>';
    allActions.sort().forEach(function(action) {
        var vA=actionA[action]||0; var vB=actionB[action]||0;
        html += '<tr><td style="text-align:right;padding-right:16px;">'+action+'</td><td class="period-a-color">'+vA+'</td><td class="period-b-color">'+vB+'</td><td>'+changeSpan(vA,vB)+'</td></tr>';
    });
    html += '</tbody></table></div>';

    // By Type
    html += '<div class="comp-section-title">🏷 حسب نوع الشكوى</div>';
    html += '<div class="comp-table-wrap"><table class="comp-table">';
    html += '<thead><tr><th>النوع</th><th style="color:#42a5f5;">الفترة الأولى</th><th style="color:#ef5350;">الفترة الثانية</th><th>التغيير</th></tr></thead><tbody>';
    allTypes.sort(function(a,b){ return (typeB[b]||0)-(typeA[b]||0)||(typeA[a]||0)-(typeB[a]||0); }).forEach(function(t) {
        var vA=typeA[t]||0; var vB=typeB[t]||0;
        html += '<tr><td style="text-align:right;padding-right:16px;">'+t+'</td><td class="period-a-color">'+vA+'</td><td class="period-b-color">'+vB+'</td><td>'+changeSpan(vA,vB)+'</td></tr>';
    });
    html += '</tbody></table></div>';

    // By Product
    html += '<div class="comp-section-title">📦 حسب المنتج</div>';
    html += '<div class="comp-table-wrap"><table class="comp-table">';
    html += '<thead><tr><th>المنتج</th><th style="color:#42a5f5;">الفترة الأولى</th><th style="color:#ef5350;">الفترة الثانية</th><th>التغيير</th></tr></thead><tbody>';
    allProds.sort().forEach(function(p) {
        var vA=prodA[p]||0; var vB=prodB[p]||0;
        html += '<tr><td style="text-align:right;padding-right:16px;">'+p+'</td><td class="period-a-color">'+vA+'</td><td class="period-b-color">'+vB+'</td><td>'+changeSpan(vA,vB)+'</td></tr>';
    });
    html += '</tbody></table></div>';

    // By Quality
    if (allQual.length > 0) {
        html += '<div class="comp-section-title">⭐ حسب نوع شكوى الجودة</div>';
        html += '<div class="comp-table-wrap"><table class="comp-table">';
        html += '<thead><tr><th>النوع</th><th style="color:#42a5f5;">الفترة الأولى</th><th style="color:#ef5350;">الفترة الثانية</th><th>التغيير</th></tr></thead><tbody>';
        allQual.sort().forEach(function(q) {
            var vA=qualA[q]||0; var vB=qualB[q]||0;
            html += '<tr><td style="text-align:right;padding-right:16px;">'+q+'</td><td class="period-a-color">'+vA+'</td><td class="period-b-color">'+vB+'</td><td>'+changeSpan(vA,vB)+'</td></tr>';
        });
        html += '</tbody></table></div>';
    }

    // ALL Branches — sorted by combined total, no limit
    html += '<div class="comp-section-title">🏪 جميع الفروع شكاوي</div>';
    html += '<div class="comp-table-wrap"><table class="comp-table">';
    html += '<thead><tr><th>#</th><th>الفرع</th><th style="color:#42a5f5;">الفترة الأولى</th><th style="color:#ef5350;">الفترة الثانية</th><th>التغيير</th></tr></thead><tbody>';
    allBranches
        .sort(function(a,b){ return ((branchB[b]||0)+(branchA[b]||0))-((branchA[a]||0)+(branchB[a]||0)); })
        .forEach(function(b, i) {
            var vA=branchA[b]||0; var vB=branchB[b]||0;
            html += '<tr><td>'+(i+1)+'</td><td style="text-align:right;padding-right:16px;">'+b+'</td><td class="period-a-color">'+vA+'</td><td class="period-b-color">'+vB+'</td><td>'+changeSpan(vA,vB)+'</td></tr>';
        });
    html += '</tbody></table></div>';

    return html;
}

// =============================================
// HELPERS
// =============================================
function kpiCard(label, vA, vB) {
    return '<div class="comp-kpi-card">' +
        '<div class="comp-kpi-label">'+label+'</div>' +
        '<div class="comp-kpi-values">' +
            '<div class="comp-kpi-val"><div class="period-tag">🔵 الأولى</div><div class="val-num period-a-color">'+vA+'</div></div>' +
            '<div class="comp-kpi-val"><div class="period-tag">🔴 الثانية</div><div class="val-num period-b-color">'+vB+'</div></div>' +
        '</div>' +
    '</div>';
}

function changeSpan(vA, vB) {
    var d = vB - vA;
    if (d === 0) return '<span class="comp-change neutral">—</span>';
    var pct = vA > 0 ? ' ('+((d/vA)*100).toFixed(1)+'%)' : '';
    var cls = d > 0 ? 'up' : 'down';
    var sym = d > 0 ? '▲' : '▼';
    return '<span class="comp-change '+cls+'">'+sym+' '+Math.abs(d)+pct+'</span>';
}

function sumColumn(data, col) {
    var total = 0;
    data.forEach(function(r) {
        var v = parseFloat(String(r[col]||'').replace(/,/g,''));
        if (!isNaN(v)) total += v;
    });
    return total;
}

function formatComp(v) {
    if (v === 0) return '—';
    return v.toLocaleString('ar-EG') + ' ر.س';
}

function daysDiff(start, end) {
    return Math.round((new Date(end) - new Date(start)) / 86400000) + 1;
}

function filterByDateRange(data, startStr, endStr) {
    var start = new Date(startStr); start.setHours(0,0,0,0);
    var end   = new Date(endStr);   end.setHours(23,59,59,999);
    return data.filter(function(item) {
        var d = parseExcelDate(item['التاريخ']);
        if (!d) return false;
        var c = new Date(d); c.setHours(0,0,0,0);
        return c >= start && c <= end;
    });
}

function getUniqueComplaints(data) {
    var seen = new Set();
    return data.filter(function(r) {
        if (!r['INDEX'] || seen.has(r['INDEX'])) return false;
        seen.add(r['INDEX']); return true;
    });
}

function countBy(data, col) {
    var counts = {};
    data.forEach(function(r) {
        var v = r[col]||'غير محدد';
        counts[v]=(counts[v]||0)+1;
    });
    return counts;
}

function topN(counts, n) {
    return Object.entries(counts).sort(function(a,b){return b[1]-a[1];}).slice(0,n);
}

window.onresize = function() { Object.values(myCharts).forEach(function(c){ c.resize(); }); };

window.addEventListener('load', function() {
    setTimeout(function() { Object.values(myCharts).forEach(function(c){ c.resize(); }); }, 300);
    setTimeout(function() { Object.values(myCharts).forEach(function(c){ c.resize(); }); }, 800);
});
