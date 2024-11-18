document.getElementById('analyzeButton').addEventListener('click', processData);

let file1Data = null;
let file2Data = null;

// Функция для преобразования даты Excel в стандартный объект Date
function excelDateToJSDate(excelDate) {
    return new Date((excelDate - 25569) * 86400 * 1000);
}

// Загрузка первого файла
document.getElementById('file1').addEventListener('change', function(event) {
    const file = event.target.files[0];
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = e.target.result;
        const workbook = XLSX.read(data, { type: 'array' });
        file1Data = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1 }); // Прочитать все строки как массив

        console.log("First file loaded:", file1Data); // Лог для загрузки первого файла

        // Проверяем, есть ли данные и выводим значения из E2, F2 и K2
        if (file1Data.length > 1) {
            const e2Value = file1Data[1][4]; // Столбец E, строка 2
            const f2Value = file1Data[1][5]; // Столбец F, строка 2
            const k2Value = file1Data[1][10]; // Столбец K, строка 2
            const currentDate = new Date().toLocaleString(); // Получаем текущую дату и время
            console.log(`E2-F2-K2: ${e2Value} - ${f2Value} - ${k2Value} (Date: ${currentDate})`);
        } else {
            console.log("No data found in the first file.");
        }
    };
    reader.readAsArrayBuffer(file);  // Чтение как массив
});

// Загрузка второго файла
document.getElementById('file2').addEventListener('change', function(event) {
    const file = event.target.files[0];
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = e.target.result;
        const workbook = XLSX.read(data, { type: 'array' });
        file2Data = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]], { header: 1 }); // Прочитать все строки как массив

        console.log("Second file loaded:", file2Data); // Лог для загрузки второго файла

        // Проверяем, есть ли данные и выводим значения из F2 и G2
        if (file2Data.length > 1) {
            const f2Value = file2Data[1][5]; // Столбец F, строка 2 (индекс 5)
            const g2Value = file2Data[1][6]; // Столбец G, строка 2 (индекс 6)

            console.log(`F2: ${f2Value}, G2: ${g2Value}`);
        } else {
            console.log("No data found in the second file.");
        }
    };
    reader.readAsArrayBuffer(file);  // Чтение как массив
});

// Основная функция обработки данных
function processData() {
    console.log("Processing data...");

    if (!file1Data || !file2Data) {
        alert("Please upload both files.");
        console.log("Files not uploaded properly.");
        return;
    }

    const notFound = [];
    console.log("File 1 data:", file1Data);
    console.log("File 2 data:", file2Data);

    file1Data.forEach(row1 => {
        const e2Value = row1[4]; // Столбец E, строка 2 (индекс 4)
        const f2Value = row1[5]; // Столбец F, строка 2 (индекс 5)
        const e2_d2 = `${e2Value}-${f2Value}`;

        // Преобразуем дату из Excel в стандартную дату
        const eventTime = excelDateToJSDate(row1[10]); // Столбец K, строка 2 (индекс 10)

        if (isNaN(eventTime)) {
            console.log(`Invalid event date in file 1 for row: ${JSON.stringify(row1)}`);
            return;
        }

        const location = row1[2]; // Столбец C, строка 2 (индекс 2)

        console.log(`Checking row: ${e2_d2}, event time: ${eventTime}, location: ${location}`);

        const matchB = file2Data.filter(row2 => row2[1] === row1[1]); // Сравниваем столбцы B (индекс 1)
        console.log("Matching rows from file2:", matchB);

        if (matchB.length > 0) {
            matchB.forEach(row2 => {
                const startTime = excelDateToJSDate(row2[5]); // Столбец F, строка 2 (индекс 5)
                const endTime = excelDateToJSDate(row2[6]); // Столбец G, строка 2 (индекс 6)

                if (isNaN(startTime) || isNaN(endTime)) {
                    console.log(`Invalid start or end time in file 2 for row: ${JSON.stringify(row2)}`);
                    return;
                }

                console.log(`Comparing event time with row2's start time: ${startTime}, end time: ${endTime}`);

                // Проверяем, попадает ли время события в диапазон
                if (startTime <= eventTime && eventTime <= endTime) {
                    const events = row2[8].split(','); // Столбец I, строка 2 (индекс 8)
                    console.log("Events from file2:", events);

                    if (!events.includes(e2_d2)) {
                        console.log(`Event not found in file2: ${e2_d2}`);
                        notFound.push(row1);
                    } else {
                        console.log(`Event found: ${e2_d2}`);
                    }
                }
            });
        } else {
            console.log(`No match found for B column value: ${row1[1]}`);
            notFound.push(row1);
        }
    });

    displayNotFound(notFound);
}

// Функция для отображения не найденных записей
function displayNotFound(notFound) {
    console.log("Displaying not found results:", notFound);

    const tableBody = document.querySelector('#resultTable tbody');
    tableBody.innerHTML = ''; // Очищаем таблицу перед добавлением новых данных

    const fragment = document.createDocumentFragment();

    notFound.forEach((row, index) => {
        const tr = document.createElement('tr');
        
        const td1 = document.createElement('td');
        const td2 = document.createElement('td');

        td1.textContent = index + 1;
        td2.textContent = JSON.stringify(row);

        tr.appendChild(td1);
        tr.appendChild(td2);
        fragment.appendChild(tr);
    });

    tableBody.appendChild(fragment); // Добавляем все строки сразу
}