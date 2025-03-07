document.addEventListener('DOMContentLoaded', () => {
    // Получаем элементы DOM для дальнейшей работы
    const pdfInput = document.getElementById('pdfInput');
    const pdfContainer = document.getElementById('pdfContainer');
    const smartBufferToggle = document.getElementById('smartBufferToggle');
    const labSelect = document.getElementById('labSelect');
    const panelSelect = document.getElementById('panelSelect');
    const elementsList = document.getElementById('elements-list');
    const genderSelect = document.getElementById('genderSelect');
    const ageInput = document.getElementById('ageInput');

    // Переменные для управления состоянием приложения
    let currentIndex = 0; // Текущий индекс для вставки значений из буфера обмена
    let lastClipboardValue = ''; // Последнее значение буфера обмена для предотвращения дублирования
    let isSmartBufferEnabled = localStorage.getItem('smartBufferEnabled') === 'true'; // Включён ли "умный буфер" (из localStorage)
    let labs = []; // Список лабораторий
    let panels = []; // Список панелей
    let parameters = []; // Список параметров
    let markers = []; // Список маркеров
    let categories = []; // Список категорий
    let valueBuffer = {}; // Буфер значений для сохранения введённых данных
    let currentLabId = null; // ID текущей выбранной лаборатории
    let clientGender = genderSelect.value; // Пол клиента genderSelect.value
    let clientAge = parseInt(ageInput.value); // Возраст клиента

    // Настройка библиотеки pdf.js для работы с PDF-файлами
    pdfjsLib.GlobalWorkerOptions.workerSrc = 'https://cdnjs.cloudflare.com/ajax/libs/pdf.js/2.10.377/pdf.worker.min.js';
    smartBufferToggle.checked = isSmartBufferEnabled; // Устанавливаем состояние переключателя "умного буфера"

    // Загружаем данные из файла data.xlsx
    fetch('data.xlsx')
        .then(response => response.arrayBuffer())
        .then(data => {
            const workbook = XLSX.read(data, { type: 'array' }); // Читаем Excel-файл

            // Преобразуем листы Excel в JSON-объекты
            labs = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[0]]); // Лаборатории
            panels = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[1]]); // Панели
            parameters = XLSX.utils.sheet_to_json(workbook.Sheets[workbook.SheetNames[2]]); // Параметры

            // Обрабатываем лист с маркерами
            const markerSheetName = workbook.SheetNames[3];
            const markerSheet = workbook.Sheets[markerSheetName];
            markers = XLSX.utils.sheet_to_json(markerSheet, { header: 1 }).slice(1).map(row => {
                const marker = { 
                    id: row[0], // ID маркера
                    name: row[1], // Название маркера
                    references: {}, // Референсные значения для лабораторий
                    units: row[6], // Единицы измерения
                    gender: row[7], // Пол
                    ageFrom: row[8], // Минимальный возраст
                    ageTo: row[9], // Максимальный возраст
                    grades: parseGrades(row[10]) // Градации (например, "низко", "оптимально")
                };
                // Парсим референсные диапазоны для лабораторий
                for (let colIndex = 2; colIndex < 6; colIndex += 2) {
                    const labIdFrom = XLSX.utils.sheet_to_json(markerSheet, { header: 1 })[0][colIndex];
                    const labIdTo = XLSX.utils.sheet_to_json(markerSheet, { header: 1 })[0][colIndex + 1];
                    const labId = labIdFrom.split('.')[0];
                    if (labs.some(lab => lab.id === labId)) {
                        marker.references[labId] = {
                            from: row[colIndex], // Нижняя граница диапазона
                            to: row[colIndex + 1] // Верхняя граница диапазона
                        };
                    }
                }
                return marker;
            });

            // Загружаем категории
            const categorySheetName = workbook.SheetNames[4];
            categories = XLSX.utils.sheet_to_json(workbook.Sheets[categorySheetName]);

            populateLabSelect(); // Заполняем выпадающий список лабораторий
        })
        .catch(err => console.error('Ошибка загрузки data.xlsx:', err));

    // Обработчик загрузки PDF-файла
    pdfInput.addEventListener('change', async (e) => {
        const file = e.target.files[0];
        if (file && file.type === 'application/pdf') {
            const fileReader = new FileReader();
            fileReader.onload = async () => {
                const typedArray = new Uint8Array(fileReader.result);
                const pdf = await pdfjsLib.getDocument(typedArray).promise;
                pdfContainer.innerHTML = ''; // Очищаем контейнер перед рендерингом
                // Рендерим каждую страницу PDF
                for (let pageNum = 1; pageNum <= pdf.numPages; pageNum++) {
                    const page = await pdf.getPage(pageNum);
                    const viewport = page.getViewport({ scale: 1.5 });
                    const pageDiv = document.createElement('div');
                    pageDiv.className = 'page';
                    pdfContainer.appendChild(pageDiv);
                    const canvas = document.createElement('canvas');
                    canvas.width = viewport.width;
                    canvas.height = viewport.height;
                    const context = canvas.getContext('2d');
                    pageDiv.appendChild(canvas);
                    const textLayerDiv = document.createElement('div');
                    textLayerDiv.className = 'textLayer';
                    textLayerDiv.style.width = `${viewport.width}px`;
                    textLayerDiv.style.height = `${viewport.height}px`;
                    pageDiv.appendChild(textLayerDiv);
                    await page.render({ canvasContext: context, viewport }).promise; // Рендерим изображение
                    const textContent = await page.getTextContent();
                    pdfjsLib.renderTextLayer({ // Рендерим текстовый слой
                        textContent,
                        container: textLayerDiv,
                        viewport,
                        pageNumber: pageNum,
                    });
                }
            };
            fileReader.readAsArrayBuffer(file); // Читаем файл как массив байтов
        }
    });

    // Заполняем выпадающий список лабораторий
    function populateLabSelect() {
        labSelect.innerHTML = '<option value="">Выберите лабораторию</option>';
        labs.forEach(lab => {
            const option = document.createElement('option');
            option.value = lab.id;
            option.textContent = lab.name;
            labSelect.appendChild(option);
        });
    }

    // Обработчик выбора лаборатории
    labSelect.addEventListener('change', () => {
        const labId = labSelect.value;
        panelSelect.disabled = !labId; // Активируем/деактивируем выбор панели
        panelSelect.innerHTML = '<option value="">Выберите панель</option>';
        elementsList.innerHTML = ''; // Очищаем список элементов
        if (labId !== currentLabId) {
            valueBuffer = {}; // Сбрасываем буфер значений при смене лаборатории
            currentLabId = labId;
        }
        if (labId) {
            const labPanels = panels.filter(panel => panel.labId === labId); // Фильтруем панели по лаборатории
            labPanels.forEach(panel => {
                const option = document.createElement('option');
                option.value = panel.id;
                option.textContent = panel.name;
                panelSelect.appendChild(option);
            });
        }
    });

    // Функция для парсинга paramIds, включая диапазоны
function parseParamIds(paramIdsString) {
    const result = [];
    const parts = paramIdsString.split(',').map(part => part.trim());

    parts.forEach(part => {
        if (part.includes('-')) {
            // Это диапазон, например, "P000001-P000020"
            const [start, end] = part.split('-').map(id => id.trim());
            const startNum = parseInt(start.replace('P', ''), 10); // Извлекаем число из P000001
            const endNum = parseInt(end.replace('P', ''), 10);     // Извлекаем число из P000020

            if (!isNaN(startNum) && !isNaN(endNum) && startNum <= endNum) {
                // Генерируем все ID в диапазоне
                for (let i = startNum; i <= endNum; i++) {
                    const generatedId = `P${String(i).padStart(6, '0')}`; // Форматируем как P000001
                    result.push(generatedId);
                }
            } else {
                console.error(`Некорректный диапазон: ${part}`);
            }
        } else {
            // Это одиночный ID, например, "P000001"
            result.push(part);
        }
    });

    return result;
}

    // Обработчик выбора панели
    panelSelect.addEventListener('change', () => {
        updateParameters(); // Обновляем параметры при выборе панели
    });

    // Обработчик изменения пола клиента
    genderSelect.addEventListener('change', () => {
        clientGender = genderSelect.value;
        if (panelSelect.value) updateParameters(); // Обновляем параметры, если выбрана панель
    });

    // Обработчик изменения возраста клиента
    ageInput.addEventListener('input', () => {
        clientAge = parseInt(ageInput.value) || 0;
        if (panelSelect.value) updateParameters(); // Обновляем параметры, если выбрана панель
    });

    // Функция обновления параметров на основе выбранной панели
    function updateParameters() {
        const panelId = panelSelect.value;
        elementsList.innerHTML = ''; // Очищаем список элементов
        currentIndex = 0; // Сбрасываем индекс для буфера
        if (panelId) {
            const selectedPanel = panels.find(panel => panel.id === panelId);
            if (selectedPanel) {
                const paramIds = parseParamIds(selectedPanel.paramIds); // Получаем массив ID с учётом диапазонов
    
                let currentCategoryContainer = null;
                let lastCategoryId = null;
    
                paramIds.forEach(paramId => {
                    const param = parameters.find(p => p.id === paramId);
                    if (param) {
                        const categoryId = param.categoryId;
                        const category = categories.find(cat => cat.id === categoryId);
    
                        // Если категория сменилась или это первый параметр, создаём новый контейнер
                        if (categoryId !== lastCategoryId || !currentCategoryContainer) {
                            currentCategoryContainer = document.createElement('div');
                            currentCategoryContainer.className = 'category-container';
                            currentCategoryContainer.dataset.categoryId = categoryId;
    
                            const categoryHeader = document.createElement('div');
                            categoryHeader.className = 'category-header';
                            categoryHeader.textContent = category ? category.name : 'Без категории';
                            currentCategoryContainer.appendChild(categoryHeader);
    
                            elementsList.appendChild(currentCategoryContainer);
                            lastCategoryId = categoryId; // Обновляем последнюю категорию
                        }
    
                        // Добавляем параметр в текущий контейнер
                        const relatedMarkers = param.relatedMarkerIds.split(',').map(id => id.trim())
                            .map(markerId => markers.find(m => m.id === markerId))
                            .filter(Boolean); // Получаем связанные маркеры
                        if (relatedMarkers.length) {
                            createParameterContainer(param.name, relatedMarkers, currentCategoryContainer, categoryId);
                        }
                    }
                });
            }
        }
    }

    // Настройка поведения стандартного ввода (для числовых значений)
    function setupStandardInputBehavior(input, updateCallback) {
        input.addEventListener('input', (e) => {
            let value = e.target.value.trim();
            value = value.replace(/[^0-9.,<>]/g, ''); // Убираем всё, кроме цифр, точки, запятой и знаков <, >
            if (value.includes(',')) value = value.replace(',', '.'); // Заменяем запятую на точку

            const hasLessThan = value.startsWith('<');
            const hasGreaterThan = value.startsWith('>');
            const cleanValue = value.replace(/[<>]/g, '');

            e.target.value = value;
            updateCallback(value); // Обновляем индикатор и метку

            clearTimeout(input.dataset.timeout);
            input.dataset.timeout = setTimeout(() => {
                autoCorrectStandardInput(input); // Автокоррекция через 5 секунд
                updateCallback(input.value);
            }, 5000);
        });

        input.addEventListener('paste', () => {
            setTimeout(() => {
                autoCorrectStandardInput(input); // Автокоррекция после вставки
                updateCallback(input.value);
            }, 0);
        });
    }

    // Автокоррекция стандартного ввода
    function autoCorrectStandardInput(input) {
        // Берём текущее значение из поля ввода и убираем лишние пробелы с краёв
        let value = input.value.trim();
        // Проверяем, есть ли операторы < или >
        const hasLessThan = value.startsWith('<');
        const hasGreaterThan = value.startsWith('>');
        // Убираем операторы и лишние пробелы, оставляем только число
        let cleanValue = value.replace(/[<>]/g, '').trim();
        // Если есть точка, убираем лишние нули в конце и саму точку, если после неё ничего нет
        if (cleanValue.includes('.')) {
            cleanValue = cleanValue.replace(/(\.\d*?)0+$/, '$1').replace(/\.$/, '');
        }
        // Формируем итоговое значение: добавляем < или > (без пробела) перед числом
        input.value = (hasLessThan ? '<' : hasGreaterThan ? '>' : '') + cleanValue;
    }

    // Настройка поведения ввода для генетических данных
    function setupGeneticsInputBehavior(input, updateCallback) {
        input.addEventListener('input', (e) => {
            let value = e.target.value.trim();
            value = value.replace(/[^a-zA-Z0-9/]/g, '').toUpperCase(); // Оставляем только буквы, цифры и слэш, приводим к верхнему регистру
            value = value.replace(/\/+/g, '/'); // Убираем лишние слэши
            e.target.value = value;
            updateCallback(value);

            clearTimeout(input.dataset.timeout);
            input.dataset.timeout = setTimeout(() => {
                autoCorrectGeneticsInput(input); // Автокоррекция через 5 секунд
                updateCallback(input.value);
            }, 5000);
        });

        input.addEventListener('paste', () => {
            setTimeout(() => {
                autoCorrectGeneticsInput(input); // Автокоррекция после вставки
                updateCallback(input.value);
            }, 0);
        });
    }

    // Автокоррекция генетического ввода
    function autoCorrectGeneticsInput(input) {
        let value = input.value.trim();
        value = value.replace(/\s+/g, ''); // Убираем пробелы
        value = value.toUpperCase(); // Приводим к верхнему регистру
        value = value.replace(/\/+/g, '/'); // Убираем лишние слэши
        input.value = value;
    }

    // Создание контейнера для параметра
    function createParameterContainer(name, markerVariants, container, categoryId) {
        const element = document.createElement('div');
        element.className = 'element';

        const label = document.createElement('label');
        label.textContent = name; // Устанавливаем начальное название
        element.appendChild(label);

        const input = document.createElement('input');
        input.type = 'text';
        input.step = '0.01'; // Шаг для числового ввода
        if (valueBuffer[name]) {
            input.value = valueBuffer[name]; // Восстанавливаем значение из буфера
        }
        element.appendChild(input);

        const refSelect = document.createElement('select');
        const labId = labSelect.value;
        markerVariants.forEach(marker => {
            Object.entries(marker.references).forEach(([refLabId, range]) => {
                if (refLabId === labId) {
                    const option = document.createElement('option');
                    option.value = JSON.stringify({ range, grades: marker.grades }); // Сохраняем диапазон и градации
                    const units = marker.units.split(',')[0].split('[')[0]; // Берем первую единицу измерения
                    option.textContent = `${range.from} - ${range.to} ${units}`;
                    option.dataset.gender = marker.gender;
                    option.dataset.ageFrom = marker.ageFrom;
                    option.dataset.ageTo = marker.ageTo;
                    option.dataset.markerName = marker.name;
                    refSelect.appendChild(option);
                }
            });
        });
        element.appendChild(refSelect);

        // Выбираем подходящий диапазон на основе пола и возраста клиента
        const matchingOption = Array.from(refSelect.options).find(option => {
            const gender = option.dataset.gender;
            const ageFrom = parseFloat(option.dataset.ageFrom);
            const ageTo = parseFloat(option.dataset.ageTo);
            return (gender === clientGender || gender === 'универсальный') && clientAge >= ageFrom && clientAge <= ageTo;
        });
        if (matchingOption) refSelect.value = matchingOption.value;


        const indicator = document.createElement('div');
        indicator.className = 'grade-indicator'; // Индикатор градации (цвет и текст)
        element.appendChild(indicator);

        // Настраиваем поведение ввода в зависимости от категории
        if (categoryId === 'C5') {
            setupGeneticsInputBehavior(input, updateLabelAndIndicator); // Генетические данные
        } else if (['C1', 'C2', 'C3', 'C4', 'C6', 'C7'].includes(categoryId)) {
            setupStandardInputBehavior(input, updateLabelAndIndicator); // Стандартные числовые данные
        }

        // Обновление метки и индикатора на основе введённого значения
function updateLabelAndIndicator(rawValue) {
    // Получаем выбранный вариант из выпадающего списка референсных значений
    const selectedOption = refSelect.selectedOptions[0];
    // Если есть название маркера в выбранном варианте, используем его для метки
    if (selectedOption && selectedOption.dataset.markerName) {
        label.textContent = selectedOption.dataset.markerName;
    } else {
        // Иначе используем исходное имя параметра
        label.textContent = name;
    }

    // Получаем введённое значение: либо переданное, либо текущее из поля ввода
    const value = rawValue || input.value;
    // Парсим данные выбранного референсного диапазона (или пустой объект, если данных нет)
    const selectedData = JSON.parse(refSelect.value || '{}');
    const range = selectedData.range || {}; // Референсный диапазон (from и to)
    // Градации: либо из выбранного диапазона, либо из первого маркера по умолчанию
    const grades = selectedData.grades || markerVariants[0].grades;
    // Начальные значения для текста и цвета индикатора
    let gradeText = 'Нет данных';
    let gradeColor = '#808080';

    // Проверяем, является ли категория генетической (C5)
    if (categoryId === 'C5') {
        // Для генетических данных ищем точное совпадение значения в градациях
        if (value && grades[value]) {
            gradeText = grades[value].text || value; // Текст градации или само значение
            gradeColor = grades[value].color || '#808080'; // Цвет градации или серый по умолчанию
        }
    } else {
        // Для числовых данных обрабатываем стандартные значения и операторы <, >
        const hasLessThan = value.startsWith('<'); // Проверяем наличие знака "<"
        const hasGreaterThan = value.startsWith('>'); // Проверяем наличие знака ">"
        // Преобразуем значение в число, убирая операторы
        const numericValue = parseFloat(value.replace(/[<>]/g, ''));

        // Проверяем, что значение корректное и есть референсный диапазон
        if (!isNaN(numericValue) && range.from !== undefined && range.to !== undefined) {
            // Сортируем градации для правильного порядка обработки
            const sortedGrades = Object.entries(grades).sort((a, b) => {
                // Определяем, является ли верхняя граница бесконечностью (Infinity или null)
                const aToInfinity = a[1].to === Infinity || a[1].to == null;
                const bToInfinity = b[1].to === Infinity || b[1].to == null;
                // Определяем, является ли нижняя граница бесконечностью
                const aFromInfinity = a[1].from === Infinity;
                const bFromInfinity = b[1].from === Infinity;
                // Диапазоны с Infinity в конце
                if (aToInfinity && !bToInfinity) return 1;
                if (!aToInfinity && bToInfinity) return -1;
                if (aFromInfinity && !bFromInfinity) return 1;
                if (!aFromInfinity && bFromInfinity) return -1;
                // Внутри групп сортируем по нижней границе
                return a[1].from - b[1].from;
            });

            // Переменная для хранения лучшего подходящего диапазона
            let bestMatch = null;

            // Проходим по отсортированным градациям
            for (const [gradeName, { from, to }] of sortedGrades) {
                const fromIsInfinity = from === Infinity; // Нижняя граница — бесконечность?
                const toIsInfinity = to === Infinity || to == null; // Верхняя граница — бесконечность?

                // Логируем проверку текущей градации для отладки
                console.log(`Checking ${gradeName}: ${from} - ${to}, Value: ${numericValue}`);

                // Обработка значений с оператором "<"
                if (hasLessThan) {
                    // Подходит, если значение меньше или равно верхней границе или верхняя граница бесконечна
                    if (numericValue <= to || toIsInfinity) {
                        // Если ещё нет кандидата, берём первый подходящий
                        if (!bestMatch) {
                            bestMatch = { gradeName, from, to, toIsInfinity };
                        } 
                        // Если текущая верхняя граница конечна и либо предыдущий кандидат был с Infinity,
                        // либо текущая верхняя граница меньше предыдущей — обновляем кандидата
                        else if (!toIsInfinity && (bestMatch.toIsInfinity || to < bestMatch.to)) {
                            bestMatch = { gradeName, from, to, toIsInfinity };
                        }
                    }
                } 
                // Обработка значений с оператором ">"
                else if (hasGreaterThan) {
                    // Подходит, если нижняя граница конечна, значение больше или равно нижней границе,
                    // и значение меньше или равно верхней границе (или верхняя граница бесконечна)
                    if (!fromIsInfinity && numericValue >= from && (toIsInfinity || numericValue <= to)) {
                        // Если нет кандидата или текущая нижняя граница больше предыдущей — обновляем кандидата
                        if (!bestMatch || from > bestMatch.from) {
                            bestMatch = { gradeName, from, to, toIsInfinity };
                        }
                    }
                } 
                // Обработка обычных значений (без операторов)
                else {
                    // Подходит, если значение находится в пределах диапазона (учитывая бесконечности)
                    if ((fromIsInfinity || numericValue >= from) && (toIsInfinity || numericValue <= to)) {
                        gradeText = gradeName; // Устанавливаем текст градации
                        gradeColor = getGradeColor(gradeName); // Устанавливаем цвет
                        break; // Прерываем цикл, так как нашли точное совпадение
                    }
                }
            }

            // Если был оператор "<" или ">", применяем лучший найденный диапазон
            if (hasLessThan || hasGreaterThan) {
                if (bestMatch) {
                    gradeText = bestMatch.gradeName;
                    gradeColor = getGradeColor(bestMatch.gradeName);
                }
            }
        }
    }

    // Логируем результат для отладки
    console.log(`Result: ${gradeText}, Color: ${gradeColor}`);
    // Обновляем текст и цвет индикатора на интерфейсе
    indicator.textContent = gradeText;
    indicator.style.backgroundColor = gradeColor;
    // Сохраняем введённое значение в буфер
    valueBuffer[name] = value || '';
}

        refSelect.addEventListener('change', () => updateLabelAndIndicator(input.value)); // Обновляем при смене диапазона
        updateLabelAndIndicator(input.value); // Инициализируем индикатор

        container.appendChild(element); // Добавляем элемент в контейнер
    }

    // Парсинг градаций из строки
    function parseGrades(gradesStr) {
        const grades = {};
        console.log("Parsing grades:", gradesStr);
        gradesStr.split(',').forEach(grade => {
            const [key, value] = grade.split('[');
            const cleanValue = value.slice(0, -1);
            if (cleanValue.includes('#')) {
                const [text, color] = cleanValue.split(',');
                grades[key.trim()] = { text, color };
            } else {
                const [from, to] = cleanValue.split('-').map(val => val ? val.trim() : val);
                const parsedFrom = from === '∞' ? Infinity : parseFloat(from);
                // Если to отсутствует или пустое, считаем его Infinity
                const parsedTo = (!to || to === '∞') ? Infinity : parseFloat(to);
                grades[key.trim()] = { from: parsedFrom, to: parsedTo };
                console.log(`Grade: ${key.trim()} -> from: ${parsedFrom}, to: ${parsedTo}`);
            }
        });
        return grades;
    }

    // Нормализация числового текста (замена запятой на точку)
    function normalizeNumber(text) {
        return text.trim().replace(',', '.');
    }

    // Проверка, является ли значение стандартным числовым
    function isStandardInputValue(value) {
        // Убираем символы < и >, затем удаляем лишние пробелы
        const cleanValue = value.replace(/[<>]/g, '').trim();
        // Проверяем, что оставшаяся часть состоит только из цифр и точки (валидное число)
        // Также проверяем, что исходное значение либо равно чистому числу, либо начинается с < или >
        return /^[\d.]+$/.test(cleanValue) && (value.trim() === cleanValue || value.trim().startsWith('<') || value.trim().startsWith('>'));
    }

    // Вставка значения из буфера обмена
    function insertValue(value) {
        // Получаем все поля ввода внутри #elements-list
        const inputs = document.querySelectorAll('#elements-list input');
        
        // Проверяем, что текущий индекс находится в пределах количества полей ввода
        if (currentIndex < inputs.length) {
            // Текущее поле ввода, в которое будет вставлено значение
            const currentInput = inputs[currentIndex];
            
            // Определяем ID категории контейнера, в котором находится инпут (по умолчанию 'C1')
            const categoryId = currentInput.closest('.category-container')?.dataset.categoryId || 'C1';
            
            // Нормализуем значение из буфера обмена, убирая лишние пробелы
            let normalizedValue = value.trim();
            
            // Проверяем, является ли значение "мусорным" (только запятая, точка или пробелы)
            const isInvalid = /^[,.\s]*$/.test(normalizedValue);
            
            // Если значение некорректно, очищаем поле и возвращаем false
            if (isInvalid) {
                currentInput.value = ''; // Очищаем поле ввода
                currentInput.dispatchEvent(new Event('input')); // Обновляем индикатор
                return false; // Не увеличиваем индекс, остаёмся на текущем поле
            }
            
            // Обрабатываем негенетические категории (кроме C5)
            if (categoryId !== 'C5') {
                // Проверяем, есть ли операторы < или >
                const hasLessThan = normalizedValue.startsWith('<');
                const hasGreaterThan = normalizedValue.startsWith('>');
                
                // Убираем операторы и лишние пробелы, оставляем только число
                let cleanValue = normalizedValue.replace(/[<>]/g, '').trim();
                
                // Если есть точка, убираем лишние нули в конце и саму точку, если после неё ничего нет
                if (cleanValue.includes('.')) {
                    cleanValue = cleanValue.replace(/(\.\d*?)0+$/, '$1').replace(/\.$/, '');
                }
                
                // Формируем итоговое значение с оператором (если был) и числом без пробелов
                normalizedValue = (hasLessThan ? '<' : hasGreaterThan ? '>' : '') + cleanValue;
            }
            
            // Присваиваем нормализованное значение текущему полю ввода
            currentInput.value = normalizedValue;
            
            // Запускаем событие 'input' для обновления индикатора и других обработчиков
            currentInput.dispatchEvent(new Event('input'));
            
            // Плавная прокрутка к текущему инпуту с запасом 50 пикселей снизу
            const formContainer = document.querySelector('.form-container'); // Прокручиваемый контейнер
            const inputRect = currentInput.getBoundingClientRect(); // Позиция инпута относительно окна
            const containerRect = formContainer.getBoundingClientRect(); // Позиция контейнера относительно окна
            const scrollOffset = formContainer.scrollTop; // Текущая прокрутка контейнера
            
            // Вычисляем позицию инпута относительно начала контейнера
            const inputTopRelativeToContainer = inputRect.top - containerRect.top + scrollOffset;
            
            // Вычисляем высоту контейнера
            const containerHeight = formContainer.clientHeight;
            
            // Целевая позиция прокрутки: верх инпута минус высота контейнера плюс высота инпута и запас 50px
            const targetScrollTop = inputTopRelativeToContainer - containerHeight + inputRect.height + 150;
            
            // Выполняем плавную прокрутку контейнера к вычисленной позиции
            formContainer.scrollTo({
                top: Math.max(0, targetScrollTop), // Не уходим за верхнюю границу (0)
                behavior: 'smooth' // Плавное перемещение
            });
            
            // Увеличиваем индекс для следующего поля ввода только при успешной вставке
            currentIndex++;
            
            // Возвращаем true, сигнализируя об успешной вставке
            return true;
        }
        // Если индекс вышел за пределы, возвращаем false (хотя это редкий случай)
        return false;
    }

    // Проверка буфера обмена
    async function checkClipboard() {
        if (!isSmartBufferEnabled) return;
        try {
            const clipboardText = await navigator.clipboard.readText();
            const normalizedText = normalizeNumber(clipboardText);
            if (clipboardText !== lastClipboardValue) {
                const isStandard = isStandardInputValue(normalizedText);
                const isGenetics = /^[\w/]+$/.test(normalizedText);
                // Добавляем проверку на "мусорные" значения вроде точки или запятой
                const isInvalid = /^[,.\s]*$/.test(normalizedText); // Только запятая, точка или пробелы
                if (isStandard || isGenetics) {
                    lastClipboardValue = clipboardText;
                    // Вставляем значение и проверяем результат
                    const insertedSuccessfully = insertValue(normalizedText);
                    if (!insertedSuccessfully) {
                        lastClipboardValue = ''; // Сбрасываем буфер, чтобы дать шанс вставить новое значение
                    }
                } else if (isInvalid) {
                    // Если скопировано некорректное значение (например, только точка), вызываем insertValue для очистки
                    insertValue(normalizedText);
                    lastClipboardValue = ''; // Сбрасываем буфер для повторной попытки
                }
            }
        } catch (err) {
            console.error('Ошибка доступа к буферу обмена:', err);
        }
    }

    setInterval(checkClipboard, 100); // Проверяем буфер обмена каждые 100 мс

    // Обработчик переключателя "умного буфера"
    smartBufferToggle.addEventListener('change', () => {
        isSmartBufferEnabled = smartBufferToggle.checked;
        localStorage.setItem('smartBufferEnabled', isSmartBufferEnabled); // Сохраняем состояние в localStorage
    });

    // Горячая клавиша для переключения "умного буфера" (Ctrl+Q)
    document.addEventListener('keydown', (e) => {
        if (e.ctrlKey && e.key === 'q') {
            e.preventDefault();
            isSmartBufferEnabled = !isSmartBufferEnabled;
            smartBufferToggle.checked = isSmartBufferEnabled;
            localStorage.setItem('smartBufferEnabled', isSmartBufferEnabled);
        }
    });

    // Получение цвета для градации
function getGradeColor(gradeName) {
    switch (gradeName) {
        case 'Критически низко':
            return '#6A5ACD'; // Яркий сине-фиолетовый (SlateBlue)
        case 'Низко':
            return '#4682B4'; // Яркий синий (SteelBlue)
        case 'Небольшое занижение':
            return '#20B2AA'; // Яркий бирюзовый (LightSeaGreen)
        case 'Оптималь':
            return '#32CD32'; // Яркий мягкий зелёный (LimeGreen)
        case 'Выше нормы':
            return '#FF8C00'; // Яркий оранжевый (DarkOrange)
        case 'Серьезное завышение':
            return '#FF4500'; // Яркий красный (OrangeRed)
        case 'Критически высоко':
            return '#8B0000'; // Тёмно-красный (DarkRed, без изменений)
        default:
            return '#808080'; // Серый (по умолчанию)
    }
}
});