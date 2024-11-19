// Обработчик загрузки первого файла
document.getElementById('file1').addEventListener('change', function(event) {
    const file = event.target.files[0];
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = e.target.result;
        const workbook = XLSX.read(data, { type: 'array' });
        file1Data = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1 });

        // Фильтруем данные, оставляем только нужные столбцы
        const filteredFile1Data = file1Data.slice(1).map(row => {
            return {
                makat: `${row[4]}-${row[5]}`, // Формируем makat из двух столбцов
                time: formatExcelDate(row[10]), // Форматируем время
            };
        });

        console.log("Filtered first file data:", filteredFile1Data);

        if (filteredFile1Data.length > 0) {
            filteredFile1Data.forEach(({ makat, time }) => {
                addRowToTable(makat, time);
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
    
    // Создаем объект для хранения данных
    let file2DataObject = [];

    reader.onload = function(e) {
        const data = e.target.result;
        const workbook = XLSX.read(data, { type: 'array' });
        file2Data = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1 });
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
                    if (makatList) {
                        const makatArray = makatList.split(',').map(m => m.trim());
                        makatArray.forEach(makat => {
                            // Добавляем данные в объект
                            file2DataObject.push({
                                makat: makat,
                                startTime: startTime,
                                endTime: endTime
                            });

                            // Создаем строку таблицы для каждого маката
                            const row = document.createElement('tr');
                            row.innerHTML = `
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

function analyzeFiles() {
    const analysisResults = [];
    file1Data.forEach((row1, index1) => {
        if (index1 > 0) {
            const makat1 = `${row1[4]}-${row1[5]}`;
            const time1 = row1[10];
            const time1Date = excelDateToJSDate(time1);
            let isMatchFound = false;

            file2Data.forEach((row2, index2) => {
                if (index2 > 0) {
                    const makat2s = row2[8].split(',');
                    const startTime = row2[5];
                    const endTime = row2[6];
                    const startTimeDate = excelDateToJSDate(startTime);
                    const endTimeDate = excelDateToJSDate(endTime);

                    // Проверка на совпадение маката и времени в промежутке
                    if (makat2s.includes(makat1) && isTimeInRange(time1Date, startTimeDate, endTimeDate)) {
                        isMatchFound = true;
                    }
                }
            });

            const matchStatus = isMatchFound ? 'Match' : 'No Match';
            analysisResults.push({ makat: makat1, time: formatExcelDate(time1), status: matchStatus });
        }
    });

    displayAnalysisResults(analysisResults);
}

// Функция для проверки, попадает ли время в промежуток
function isTimeInRange(timeToCheck, startTime, endTime) {
    return timeToCheck >= startTime && timeToCheck <= endTime;
}

// Функция для вывода результатов анализа
function displayAnalysisResults(results) {
    const analysisTableBody = document.getElementById('analysisTable').querySelector('tbody');
    analysisTableBody.innerHTML = '';
    results.forEach(result => {
        const row = document.createElement('tr');
        row.innerHTML = `
            <td>${result.makat}</td>
            <td>${result.time}</td>
            <td>${result.status}</td>
        `;
        analysisTableBody.appendChild(row);
    });
}

// Утилитарные функции

function excelDateToJSDate(excelDate) {
    return new Date((excelDate - 25569) * 86400 * 1000);
}

function formatExcelDate(excelDate) {
    const jsDate = excelDateToJSDate(excelDate);
    const day = jsDate.getDate().toString().padStart(2, '0');
    const month = (jsDate.getMonth() + 1).toString().padStart(2, '0');
    const year = jsDate.getFullYear();
    const hours = jsDate.getHours().toString().padStart(2, '0');
    const minutes = jsDate.getMinutes().toString().padStart(2, '0');
    const seconds = jsDate.getSeconds().toString().padStart(2, '0');
    return `${day}/${month}/${year} ${hours}:${minutes}:${seconds}`;
}

function addRowToTable(makat, timeValue) {
    const tableBody = document.getElementById('resultTable1').querySelector('tbody');
    const row = document.createElement('tr');
    row.innerHTML = `<td>${makat}</td><td>${timeValue}</td>`;
    tableBody.appendChild(row);
}

function addRowToTableForFile2(makat, startTime, endTime) {
    const tableBody = document.getElementById('resultTable2').querySelector('tbody');
    const row = document.createElement('tr');
    row.innerHTML = `<td>${makat}</td><td>${startTime}</td><td>${endTime}</td>`;
    tableBody.appendChild(row);
}