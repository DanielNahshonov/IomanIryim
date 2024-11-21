// Глобальные переменные для данных обоих файлов
let filteredFile1Data = [];
let file2DataObject = [];

// Обработчик загрузки первого файла
document.getElementById('file1').addEventListener('change', function(event) {
    const file = event.target.files[0];
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = e.target.result;
        const workbook = XLSX.read(data, { type: 'array' });
        const file1Data = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1 });

        // Фильтруем данные, оставляем только нужные столбцы
        filteredFile1Data = file1Data.slice(1).filter(row => row[2] && row[4] && row[5] && row[10]) // Убедимся, что все нужные столбцы не пустые
    .map(row => {
        return {
            eshkol: row[2],
            makat: `${row[4]}-${row[5]}`,
            time: formatExcelDate(row[10]),
        };
    });

        console.log("Filtered first file data:", filteredFile1Data);

        if (filteredFile1Data.length > 0) {
            filteredFile1Data.forEach(({ eshkol, makat, time }) => {
                addRowToTable(eshkol, makat, time);
            });
        } else {
            console.log("No data found in the filtered first file.");
        }
    };
    reader.readAsArrayBuffer(file);
});

// Обработчик загрузки второго файла
document.getElementById('file2').addEventListener('change', function(event) {
    const file = event.target.files[0];
    const reader = new FileReader();
    
    reader.onload = function(e) {
        const data = e.target.result;
        const workbook = XLSX.read(data, { type: 'array' });
        const file2Data = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1 });
        console.log("Second file loaded:", file2Data);

        if (file2Data.length > 1) {
            // Очищаем таблицу перед добавлением новых данных
            const tableBody = document.getElementById('resultTable2').querySelector('tbody');
            tableBody.innerHTML = ''; // Очищаем таблицу перед добавлением новых данных

            file2Data.forEach((row, index) => {
                if (index > 0) { // Пропускаем первую строку с заголовками
                    const startTime = row[5];
                    const endTime = row[6];
                    const makatList = row[8];
                    const eshkol = row[1]
                    if (makatList) {
                        const makatArray = makatList.split(',').map(m => m.trim());
                        makatArray.forEach(makat => {
                            // Добавляем данные в объект
                            file2DataObject.push({
                                eshkol: eshkol,
                                makat: makat,
                                startTime: startTime,
                                endTime: endTime
                            });

                            // Создаем строку таблицы для каждого маката
                            const row = document.createElement('tr');
                            row.innerHTML = `
                                <td>${eshkol}</td>
                                <td>${makat}</td>
                                <td>${startTime}</td>
                                <td>${endTime}</td>
                            `;
                            tableBody.appendChild(row);
                        });
                    }
                }
            });

            console.log("Data added to file2DataObject:", file2DataObject);
        } else {
            console.log("No data found in the second file.");
        }
    };

    reader.readAsArrayBuffer(file);
});

// Основная функция для обработки анализа
document.getElementById('analyzeButton').addEventListener('click', analyzeFiles);

const analysisResults = [];
function analyzeFiles() {
    filteredFile1Data.forEach((row1) => {
        const eshkol1 = row1.eshkol
        const makat1 = row1.makat;
        const time1 = row1.time; // Время в формате строки "DD/MM/YYYY HH:MM:SS"
        const time1Date = stringToDate(time1); // Преобразуем строку в объект Date
        let isMatchFound = false;

        file2DataObject.forEach((row2) => {
            const makat2 = row2.makat;
            const startTime = row2.startTime; // Время в формате строки "DD/MM/YYYY HH:MM:SS"
            const endTime = row2.endTime; // Время в формате строки "DD/MM/YYYY HH:MM:SS"
            const startTimeDate = stringToDate(startTime); // Преобразуем строку в объект Date
            const endTimeDate = stringToDate(endTime); // Преобразуем строку в объект Date

            // Сравниваем значения времени как объекты Date
            if (makat2 === makat1 && isTimeInRange(time1Date, startTimeDate, endTimeDate)) {
                isMatchFound = true;
            }
        });

        // Добавляем только те результаты, у которых "No Match"
        if (!isMatchFound) {
            analysisResults.push({ eshkol: eshkol1 ,makat: makat1, time: time1, status: 'פיספוס' });
        }
    });

    displayAnalysisResults(analysisResults);
}

// Функция для преобразования строки даты в объект Date
function stringToDate(dateString) {
    const [day, month, year, hour, minute, second] = dateString.split(/[/ :]/);
    return new Date(year, month - 1, day, hour, minute, second);
}

// Функция для проверки, попадает ли время в промежуток
function isTimeInRange(timeToCheck, startTime, endTime) {
    return timeToCheck >= startTime && timeToCheck <= endTime;
}

// Функция для вывода результатов анализа
// Функция для вывода результатов анализа
function displayAnalysisResults(results) {
    const analysisTableBody = document.getElementById('analysisTable').querySelector('tbody');
    analysisTableBody.innerHTML = '';

    if (results.length === 0) {
        // Если нет результатов, выводим сообщение
        const messageRow = document.createElement('tr');
        messageRow.innerHTML = `<td colspan="4" style="text-align: center;">כל הכבוד לאושרית!!!</td>`;
        analysisTableBody.appendChild(messageRow);
    } else {
        // Если есть результаты, добавляем их в таблицу
        results.forEach(result => {
            const row = document.createElement('tr');
            row.innerHTML = `
                <td>${result.eshkol}</td>
                <td>${result.makat}</td>
                <td>${result.time}</td>
                <td>${result.status}</td>
            `;
            analysisTableBody.appendChild(row);
        });
    }
}

// Утилитарные функции

function jsDateToExcelDate(jsDate) {
    // Получаем количество миллисекунд с 1 января 1970 года
    const msInDay = 86400 * 1000;
    const excelBaseDate = new Date(1900, 0, 1); // 1 января 1900 года
    const timeDiff = jsDate - excelBaseDate; // Разница в миллисекундах между датами
    
    // Возвращаем количество дней в формате Excel
    return timeDiff / msInDay + 25569; // 25569 — это количество дней от 1900 года до 1970 года
}

function excelDateToJSDate(excelDate) {
    const jsDate = new Date((excelDate - 25569) * 86400 * 1000);
    const utcDate = new Date(jsDate.toUTCString());

    // Применяем округление до ближайшей минуты
    const roundedDate = roundToMinute(utcDate);

    return roundedDate;
}

function roundToMinute(date) {
    const msInMinute = 60000;
    return new Date(Math.round(date.getTime() / msInMinute) * msInMinute);
}

function formatExcelDate(excelDate) {
    const jsDate = excelDateToJSDate(excelDate);
    const day = jsDate.getUTCDate().toString().padStart(2, '0');
    const month = (jsDate.getUTCMonth() + 1).toString().padStart(2, '0');
    const year = jsDate.getUTCFullYear();
    const hours = jsDate.getUTCHours().toString().padStart(2, '0');
    const minutes = jsDate.getUTCMinutes().toString().padStart(2, '0');
    const seconds = jsDate.getUTCSeconds().toString().padStart(2, '0');
    return `${day}/${month}/${year} ${hours}:${minutes}:${seconds}`;
}

function addRowToTable(eshkol,makat, timeValue) {
    const tableBody = document.getElementById('resultTable1').querySelector('tbody');
    const row = document.createElement('tr');
    row.innerHTML = `<td>${eshkol}</td><td>${makat}</td><td>${timeValue}</td>`;
    tableBody.appendChild(row);
}
document.getElementById('exportButton').addEventListener('click', exportToExcel);

function exportToExcel() {
    if (!analysisResults || analysisResults.length === 0) {
        alert("Нет данных для экспорта.");
        return;
    }

    // Создаем массив данных для экспорта
    const exportData = analysisResults.map(result => ({
        "אשכול": result.eshkol,
        "מק\"ט": result.makat,
        "זמן": result.time,
        "סטטוס": result.status || "No Match"
    }));

    // Преобразуем данные в формат Excel
    const worksheet = XLSX.utils.json_to_sheet(exportData);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, "Analysis Results");

    // Генерируем файл и загружаем его
    XLSX.writeFile(workbook, "AnalysisResults.xlsx");
}