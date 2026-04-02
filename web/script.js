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

let selectedFiles = [];
let excelData = null;

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

function displayResults(tasks, stats, filesInfo) {
    let statsHtml = '<div class="stat-card"><div class="stat-value">' + stats.total + '</div><div class="stat-label">Всего задач</div></div>';
    statsHtml += '<div class="stat-card"><div class="stat-value">' + stats.with_responsible + '</div><div class="stat-label">С ответственным</div></div>';
    statsHtml += '<div class="stat-card"><div class="stat-value">' + stats.with_date + '</div><div class="stat-label">С датой</div></div>';
    if (filesInfo && filesInfo.length) {
        statsHtml += '<div class="stat-card"><div class="stat-value">' + filesInfo.length + '</div><div class="stat-label">Обработано файлов</div></div>';
    }
    statsDiv.innerHTML = statsHtml;
    
    let tasksHtml = '';
    for (let i = 0; i < tasks.length; i++) {
        const task = tasks[i];
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
    resultsDiv.style.display = 'block';
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

processBtn.addEventListener('click', async function() {
    if (selectedFiles.length === 0) return;
    
    processBtn.disabled = true;
    loader.style.display = 'block';
    resultsDiv.style.display = 'none';
    downloadBtn.style.display = 'none';
    
    const formData = new FormData();
    
    for (let i = 0; i < selectedFiles.length; i++) {
        formData.append('files', selectedFiles[i]);
    }
    
    formData.append('export_to_sheets', exportSheets.checked);
    formData.append('export_to_calendar', exportCalendar.checked);
    formData.append('sheets_url', sheetsUrl.value);
    formData.append('calendar_id', calendarId.value);
    
    // ===== ПРОВЕРКА: что отправляем =====
    console.log('📤 Отправляю FormData:');
    for (let pair of formData.entries()) {
        console.log(pair[0], '=', pair[1]);
    }
    // ===================================
    
    console.log('📤 Отправляю:', {
        files: selectedFiles.length,
        export_to_calendar: exportCalendar.checked,
        calendar_id: calendarId.value
    });
    
    try {
        const response = await fetch('http://localhost:8000/parse-batch', {
            method: 'POST',
            body: formData
        });
        
        console.log('📥 Ответ получен, статус:', response.status);
        
        if (!response.ok) {
            throw new Error('HTTP ошибка: ' + response.status);
        }
        
        const data = await response.json();
        console.log('📦 Данные:', data);
        
        if (data.success) {
            if (data.excel_base64) {
                excelData = data.excel_base64;
                setupDownload(excelData);
            }
            displayResults(data.tasks, data.statistics, data.files);
            
            if (data.calendar_export === 'success') {
                alert('✅ Задачи добавлены в Google Calendar');
            } else if (data.calendar_export) {
                alert('❌ Ошибка Calendar: ' + data.calendar_export);
            }
        } else {
            alert('Ошибка: ' + (data.error || 'Неизвестная ошибка'));
        }
        
    } catch (err) {
        console.error('❌ Ошибка:', err);
        alert('Ошибка соединения: ' + err.message + '\n\nПроверь, что сервер запущен (python3 backend.py)');
    } finally {
        processBtn.disabled = false;
        loader.style.display = 'none';
    }
});