const dropZone = document.getElementById('dropZone');
const fileInput = document.getElementById('fileInput');
const selectedFilesDiv = document.getElementById('selectedFiles');
const filesCountSpan = document.getElementById('filesCount');
const filesListDiv = document.getElementById('filesList');
const processBtn = document.getElementById('processBtn');
const loader = document.getElementById('loader');
const resultsDiv = document.getElementById('results');
const tasksBody = document.getElementById('tasksBody');
const statsDiv = document.getElementById('stats');
const downloadBtn = document.getElementById('downloadBtn');
const exportSheets = document.getElementById('exportSheets');
const sheetsUrl = document.getElementById('sheetsUrl');
const exportCalendar = document.getElementById('exportCalendar');
const calendarId = document.getElementById('calendarId');
const filterPanel = document.getElementById('filterPanel');
const filterHeader = document.getElementById('filterHeader');
const filterContent = document.getElementById('filterContent');
const filterArrow = document.getElementById('filterArrow');
const clearFiltersBtn = document.getElementById('clearFiltersBtn');

let selectedFiles = [];
let excelData = null;
let allTasks = [];
let currentTasks = [];
let isFilterOpen = false;

function updateFilesList() {
    if (selectedFiles.length === 0) {
        selectedFilesDiv.style.display = 'none';
        processBtn.disabled = true;
        return;
    }
    selectedFilesDiv.style.display = 'block';
    filesCountSpan.textContent = selectedFiles.length;
    let filesHtml = '';
    for (let i = 0; i < selectedFiles.length; i++) {
        const file = selectedFiles[i];
        filesHtml += '<div class="file-item"><span>📄 ' + file.name + '</span><button onclick="removeFile(' + i + ')">✕</button></div>';
    }
    filesListDiv.innerHTML = filesHtml;
    processBtn.disabled = false;
}

window.removeFile = function(index) {
    selectedFiles.splice(index, 1);
    updateFilesList();
};

function addFiles(files) {
    const allowed = ['.pdf', '.docx', '.doc'];
    const valid = [];
    for (let i = 0; i < files.length; i++) {
        const file = files[i];
        const ext = '.' + file.name.split('.').pop().toLowerCase();
        if (allowed.includes(ext)) {
            valid.push(file);
        }
    }
    if (valid.length === 0) {
        alert('Поддерживаются только PDF, DOCX, DOC');
        return;
    }
    for (let i = 0; i < valid.length; i++) {
        const newFile = valid[i];
        let exists = false;
        for (let j = 0; j < selectedFiles.length; j++) {
            if (selectedFiles[j].name === newFile.name && selectedFiles[j].size === newFile.size) {
                exists = true;
                break;
            }
        }
        if (!exists) {
            selectedFiles.push(newFile);
        }
    }
    updateFilesList();
}

dropZone.addEventListener('click', function() {
    fileInput.click();
});

dropZone.addEventListener('dragover', function(e) {
    e.preventDefault();
});

dropZone.addEventListener('drop', function(e) {
    e.preventDefault();
    if (e.dataTransfer.files.length) {
        addFiles(e.dataTransfer.files);
    }
});

fileInput.addEventListener('change', function(e) {
    if (e.target.files.length) {
        addFiles(e.target.files);
    }
    fileInput.value = '';
});

exportSheets.addEventListener('change', function() {
    sheetsUrl.disabled = !exportSheets.checked;
});

exportCalendar.addEventListener('change', function() {
    calendarId.disabled = !exportCalendar.checked;
});

function displayFilters(responsibles, totalCount) {
    const filterOptions = document.getElementById('filterOptions');
    const filterBadge = document.getElementById('filterBadge');
    
    filterBadge.textContent = totalCount;
    
    let html = '<label class="filter-checkbox"><input type="checkbox" id="filterAll" checked> Все (' + totalCount + ' задач)</label>';
    
    for (let i = 0; i < responsibles.length; i++) {
        const resp = responsibles[i];
        html += '<label class="filter-checkbox"><input type="checkbox" class="respFilter" value="' + resp.name.replace(/"/g, '&quot;') + '" data-count="' + resp.count + '" checked> ' + resp.name + ' (' + resp.count + ' задач)</label>';
    }
    
    filterOptions.innerHTML = html;
    filterPanel.style.display = 'block';
    
    const filterAll = document.getElementById('filterAll');
    filterAll.addEventListener('change', function() {
        const allCheckboxes = document.querySelectorAll('.respFilter');
        for (let i = 0; i < allCheckboxes.length; i++) {
            allCheckboxes[i].checked = this.checked;
        }
        applyFilter();
    });
    
    const respFilters = document.querySelectorAll('.respFilter');
    for (let i = 0; i < respFilters.length; i++) {
        respFilters[i].addEventListener('change', function() {
            const allCheckboxes = document.querySelectorAll('.respFilter');
            let allChecked = true;
            for (let j = 0; j < allCheckboxes.length; j++) {
                if (!allCheckboxes[j].checked) {
                    allChecked = false;
                    break;
                }
            }
            filterAll.checked = allChecked;
            applyFilter();
        });
    }
}

function applyFilter() {
    const selectedResponsibles = [];
    const respFilters = document.querySelectorAll('.respFilter:checked');
    for (let i = 0; i < respFilters.length; i++) {
        selectedResponsibles.push(respFilters[i].value);
    }
    
    const filterAll = document.getElementById('filterAll');
    if (selectedResponsibles.length === 0 || filterAll.checked) {
        currentTasks = [...allTasks];
    } else {
        currentTasks = [];
        for (let i = 0; i < allTasks.length; i++) {
            const task = allTasks[i];
            const taskResp = task.responsible || "Не назначен";
            if (selectedResponsibles.includes(taskResp)) {
                currentTasks.push(task);
            }
        }
    }
    
    updateStatsDisplay();
    displayTasksTable();
    
    const filterBadge = document.getElementById('filterBadge');
    filterBadge.textContent = currentTasks.length;
}

function updateStatsDisplay() {
    const total = currentTasks.length;
    let withResponsible = 0;
    let withDate = 0;
    for (let i = 0; i < currentTasks.length; i++) {
        if (currentTasks[i].responsible) withResponsible++;
        if (currentTasks[i].due_date_str) withDate++;
    }
    
    statsDiv.innerHTML = '<div class="stat-card"><div class="stat-value">' + total + '</div><div class="stat-label">Всего задач</div></div>' +
        '<div class="stat-card"><div class="stat-value">' + withResponsible + '</div><div class="stat-label">С ответственным</div></div>' +
        '<div class="stat-card"><div class="stat-value">' + withDate + '</div><div class="stat-label">С датой</div></div>';
}

function displayTasksTable() {
    if (!currentTasks || currentTasks.length === 0) {
        tasksBody.innerHTML = '<tr><td colspan="5" style="text-align: center;">Задачи не найдены</td></tr>';
        return;
    }
    
    let tasksHtml = '';
    for (let i = 0; i < currentTasks.length; i++) {
        const task = currentTasks[i];
        const summary = task.summary || (task.full_description || '').substring(0, 80);
        const description = (task.full_description || '').substring(0, 100);
        tasksHtml += '<tr>';
        tasksHtml += '<td>' + task.number + '</td>';
        tasksHtml += '<td>' + summary + ((task.full_description || '').length > 80 ? '...' : '') + '</td>';
        tasksHtml += '<td>' + description + ((task.full_description || '').length > 100 ? '...' : '') + '</td>';
        tasksHtml += '<td>' + (task.responsible || '-') + '</td>';
        tasksHtml += '<td>' + (task.due_date_str || '-') + '</td>';
        tasksHtml += '</tr>';
    }
    tasksBody.innerHTML = tasksHtml;
}

function setupDownload(data) {
    if (!data) return;
    downloadBtn.style.display = 'inline-block';
    downloadBtn.onclick = function() {
        try {
            const binaryString = atob(data);
            const bytes = new Uint8Array(binaryString.length);
            for (let i = 0; i < binaryString.length; i++) {
                bytes[i] = binaryString.charCodeAt(i);
            }
            const blob = new Blob([bytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            const now = new Date();
            const year = now.getFullYear();
            const month = String(now.getMonth() + 1).padStart(2, '0');
            const day = String(now.getDate()).padStart(2, '0');
            const hours = String(now.getHours()).padStart(2, '0');
            const minutes = String(now.getMinutes()).padStart(2, '0');
            const seconds = String(now.getSeconds()).padStart(2, '0');
            const timestamp = year + '-' + month + '-' + day + 'T' + hours + '-' + minutes + '-' + seconds;
            a.download = 'tasks_' + timestamp + '.xlsx';
            a.click();
            URL.revokeObjectURL(url);
        } catch (err) {
            console.error('Ошибка скачивания:', err);
            alert('Ошибка при скачивании файла');
        }
    };
}

if (filterHeader) {
    filterHeader.addEventListener('click', (e) => {
        if (e.target.classList && e.target.classList.contains('clear-btn')) return;
        isFilterOpen = !isFilterOpen;
        if (isFilterOpen) {
            filterContent.classList.add('show');
            filterArrow.classList.add('open');
        } else {
            filterContent.classList.remove('show');
            filterArrow.classList.remove('open');
        }
    });
}

if (clearFiltersBtn) {
    clearFiltersBtn.addEventListener('click', () => {
        const filterAll = document.getElementById('filterAll');
        if (filterAll) {
            filterAll.checked = true;
        }
        const respFilters = document.querySelectorAll('.respFilter');
        for (let i = 0; i < respFilters.length; i++) {
            respFilters[i].checked = true;
        }
        applyFilter();
    });
}

processBtn.addEventListener('click', async function() {
    if (selectedFiles.length === 0) return;
    
    processBtn.disabled = true;
    loader.style.display = 'block';
    resultsDiv.style.display = 'none';
    downloadBtn.style.display = 'none';
    filterPanel.style.display = 'none';
    
    const formData = new FormData();
    for (let i = 0; i < selectedFiles.length; i++) {
        formData.append('files', selectedFiles[i]);
    }
    formData.append('export_to_sheets', exportSheets.checked);
    formData.append('export_to_calendar', exportCalendar.checked);
    formData.append('sheets_url', sheetsUrl.value);
    formData.append('calendar_id', calendarId.value);
    
    try {
        const response = await fetch('http://localhost:8000/parse-batch', {
            method: 'POST',
            body: formData
        });
        
        const data = await response.json();
        
        if (data.success) {
            allTasks = data.tasks;
            currentTasks = [...allTasks];
            
            if (data.excel_base64) {
                excelData = data.excel_base64;
                setupDownload(excelData);
            }
            
            updateStatsDisplay();
            displayTasksTable();
            resultsDiv.style.display = 'block';
            
            if (data.responsibles && data.responsibles.length > 0) {
                displayFilters(data.responsibles, allTasks.length);
            }
            
            if (data.calendar_export === 'success') {
                alert('Задачи добавлены в Google Calendar');
            } else if (data.calendar_export) {
                alert('Ошибка Calendar: ' + data.calendar_export);
            }
        } else {
            alert('Ошибка: ' + (data.error || 'Неизвестная ошибка'));
        }
        
    } catch (err) {
        console.error('Ошибка:', err);
        alert('Ошибка соединения: ' + err.message);
    } finally {
        processBtn.disabled = false;
        loader.style.display = 'none';
    }
});