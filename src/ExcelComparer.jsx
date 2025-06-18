import React, { useState } from 'react';
import * as XLSX from 'xlsx-js-style';

export default function ExcelComparer() {
  const [file1, setFile1] = useState(null);
  const [wb1, setWb1] = useState(null);
  const [sheet1, setSheet1] = useState('');

  const [file2, setFile2] = useState(null);
  const [wb2, setWb2] = useState(null);
  const [sheet2, setSheet2] = useState('');

  // Состояния для полей формы нового товара
  const [newItemName, setNewItemName] = useState('');
  const [newItemQuantity, setNewItemQuantity] = useState(''); // Количество вложений нового товара
  const [newItemPrice, setNewItemPrice] = useState('');
  const [newItemLink, setNewItemLink] = useState('');
  const [newItemWeight, setNewItemWeight] = useState(''); // Общий вес вложений нового товара

  // НОВЫЕ СОСТОЯНИЯ ДЛЯ ДОПОЛНИТЕЛЬНЫХ ПОЛЕЙ
  const [newIndexInputValue, setNewIndexInputValue] = useState('');
  const [newContactPersonInputValue, setNewContactPersonInputValue] = useState('');

  const handleFileUpload = async (file, setWb, setSheet) => {
    const data = await file.arrayBuffer();
    const workbook = XLSX.read(data);
    setWb(workbook);
    setSheet(workbook.SheetNames[0]);
  };

  const handleCompare = () => {
    if (!wb1 || !wb2 || !sheet1 || !sheet2) {
      return alert('Выберите оба файла и листы.');
    }

    const sheetData1 = XLSX.utils.sheet_to_json(wb1.Sheets[sheet1], { defval: '' });
    const sheetData2 = XLSX.utils.sheet_to_json(wb2.Sheets[sheet2], {
      header: 1,
      defval: ''
    });

    const normalize = str => str?.toString().trim().toLowerCase();

    const allValues2 = new Set(
      sheetData2.flat().map(val => normalize(val))
    );

    const targetColumnName = "Номер-посылки-(накладной)";
    const matched = [];
    const unmatched = [];

    for (const row of sheetData1) {
      const value = normalize(row[targetColumnName]);
      if (allValues2.has(value)) {
        matched.push(row);
      } else {
        unmatched.push(row);
      }
    }

    const wb = XLSX.utils.book_new();

    const borderStyle = {
      style: "thin",
      color: { rgb: "000000" }
    };

    // === ЛОГИКА УДАЛЕНИЯ ДУБЛИКАТОВ ВЕСА (Первичное прохождение до всех модификаций) ===
    const initialWeightColumnName = "包裹重量一个包裹一次备注就行общий вес посылки"; // Оригинальное название
    if (!Object.keys(sheetData1[0] || {}).includes(initialWeightColumnName)) {
        alert(`Колонка "${initialWeightColumnName}" не найдена в файле.`);
        return;
    }
    
    const preModificationOrderMap = new Map();
    for (let i = 0; i < matched.length; ++i) {
        const rowObj = matched[i];
        const orderId = rowObj?.[targetColumnName];
        if (orderId !== undefined && orderId !== null && orderId !== '') {
            if (!preModificationOrderMap.has(orderId)) {
                preModificationOrderMap.set(orderId, []);
            }
            preModificationOrderMap.get(orderId).push(i);
        }
    }

    for (const indices of preModificationOrderMap.values()) {
        if (indices.length <= 1) continue;

        const firstRowIndexInOrder = indices[0];
        const firstWeight = normalize(matched[firstRowIndexInOrder]?.[initialWeightColumnName]);

        for (let i = 1; i < indices.length; ++i) {
            const currentRowIndexInOrder = indices[i];
            const row = matched[currentRowIndexInOrder];
            if (normalize(row?.[initialWeightColumnName]) === firstWeight) {
                row[initialWeightColumnName] = '';
            }
        }
    }
    // === КОНЕЦ ЛОГИКИ УДАЛЕНИЯ ДУБЛИКАТОВ ВЕСА (Первичное прохождение) ===


    // === ЛОГИКА: УДАЛЕНИЕ СТОЛБЦОВ ПО СПИСКУ ===
    const columnsToRemove = [
      "收件人", "收件地址", "收件人城市", "收件人州省", "国家简码",
      "收件国家", "寄件国家", "寄件申报品名", "收件申报品名", "海关编码",
      "重量", 
      "内部单号", "订单号", "发货时间", "物流方式", "袋号", "分拣人",
      "长宽高", "体积重", "仓库名称", "是否带电", "产品信息",
      "产品信息sku1", "SKU", "ENGLISH", "Брэнд/изготовитель", "Материал", 
      "Группа товаров", "单品价格цена за 1 вложение(товар) в юанях","Цена поставщика", "__EMPTY", "链接ССЫЛКА"
    ];

    const filterColumns = (dataArray) => {
      return dataArray.map(row => {
        const newRow = {};
        for (const key in row) {
          if (!columnsToRemove.includes(key)) {
            newRow[key] = row[key];
          }
        }
        return newRow;
      });
    };

    const filteredMatched = filterColumns(matched);
    const filteredUnmatched = filterColumns(unmatched);
    // === КОНЕЦ ЛОГИКИ УДАЛЕНИЯ СТОЛБЦОВ ===


    // === ЛОГИКА: ОЧИСТКА ЗАГОЛОВКОВ СТОЛБЦОВ ОТ КИТАЙСКИХ СИМВОЛОВ ===
    const sanitizeHeaderName = (header) => {
      const sanitized = header.replace(/[^a-zA-Zа-яА-Я0-9\s_.-]/g, ' ');
      return sanitized.replace(/\s+/g, ' ').trim();
    };

    const cleanObjectKeys = (dataArray) => {
      return dataArray.map(row => {
        const newRow = {};
        for (const key in row) {
          const newKey = sanitizeHeaderName(key);
          newRow[newKey] = row[key];
        }
        return newRow;
      });
    };

    let finalMatchedData = cleanObjectKeys(filteredMatched);
    let finalUnmatchedData = cleanObjectKeys(filteredUnmatched);
    // === КОНЕЦ ЛОГИКИ ОЧИСТКИ ЗАГОЛОВКОВ ===

    // === ЛОГИКА: КОПИРОВАНИЕ СТОЛБЦОВ В КОНЕЦ ===
    const originalColumnToCopy1 = sanitizeHeaderName("Номер-посылки-(накладной)");
    const originalColumnToCopy2 = sanitizeHeaderName("наименование на русском");

    const copyColumnSuffix = " (копия)";

    const addColumnCopies = (dataArray) => {
        return dataArray.map(row => {
            const newRow = { ...row };
            if (row[originalColumnToCopy1] !== undefined) {
                newRow[originalColumnToCopy1 + copyColumnSuffix] = row[originalColumnToCopy1];
            }
            if (row[originalColumnToCopy2] !== undefined) {
                newRow[originalColumnToCopy2 + copyColumnSuffix] = row[originalColumnToCopy2];
            }
            return newRow;
        });
    };

    finalMatchedData = addColumnCopies(finalMatchedData);
    finalUnmatchedData = addColumnCopies(finalUnmatchedData);
    // === КОНЕЦ ЛОГИКИ КОПИРОВАНИЯ СТОЛБЦОВ ===


    // === НОВАЯ ЛОГИКА: ДОБАВЛЕНИЕ СТРОК ИЗ ФОРМЫ И ОБНОВЛЕНИЕ СУЩЕСТВУЮЩИХ ПОЛЕЙ ===
    const processedMatchedData = []; // Новый массив для сборки всех строк
    
    // Очищенные названия столбцов, которые будут использоваться
    const totalItemsColumn = sanitizeHeaderName("общ. Количество товаров в посылке (накладной)");
    const totalWeightColumn = sanitizeHeaderName("общий вес посылки"); // Это поле теперь будет перезаписано
    const itemNameColumn = sanitizeHeaderName("наименование на русском");
    const itemQuantityColumn = sanitizeHeaderName("количество вложений");
    const itemPriceColumn = sanitizeHeaderName("Цена товара");
    const itemLinkColumn = sanitizeHeaderName("Ссылка на товар");
    const itemWeightColumn = sanitizeHeaderName("общий вес вложений"); // Вес для отдельного вложения
    const targetOrderIdColumn = sanitizeHeaderName(targetColumnName);

    // Группируем finalMatchedData по OrderId для удобства обработки
    const currentOrderItemsMap = new Map(); // orderId -> [actual row objects from finalMatchedData]
    for (const row of finalMatchedData) {
        const orderId = row[targetOrderIdColumn];
        if (!currentOrderItemsMap.has(orderId)) {
            currentOrderItemsMap.set(orderId, []);
        }
        currentOrderItemsMap.get(orderId).push(row);
    }

    // Получаем количество и вес нового товара из формы (парсим в числа)
    const parsedNewItemQuantity = parseFloat(newItemQuantity) || 0;
    const parsedNewItemWeight = parseFloat(newItemWeight) || 0;

    for (const [orderId, currentOrderRows] of currentOrderItemsMap.entries()) {
        const baseRow = currentOrderRows[0] || {}; // Берем первую строку заказа как шаблон

        // Создаем новую строку для добавленного товара
        const newItem = { ...baseRow }; // Копируем общие данные из первой строки заказа

        // Копируем данные из заказа (поля, которые должны быть одинаковыми для всего заказа)
        newItem[targetOrderIdColumn] = baseRow[targetOrderIdColumn];
        newItem[sanitizeHeaderName("ФАМИЛИЯ")] = baseRow[sanitizeHeaderName("ФАМИЛИЯ")];
        newItem[sanitizeHeaderName("ИМЯ")] = baseRow[sanitizeHeaderName("ИМЯ")];
        newItem[sanitizeHeaderName("ОТЧЕСТВО")] = baseRow[sanitizeHeaderName("ОТЧЕСТВО")];
        newItem[sanitizeHeaderName("АДРЕС")] = baseRow[sanitizeHeaderName("АДРЕС")];
        newItem[sanitizeHeaderName("ГОРОД")] = baseRow[sanitizeHeaderName("ГОРОД")];
        newItem[sanitizeHeaderName("ОБЛАСТЬ")] = baseRow[sanitizeHeaderName("ОБЛАСТЬ")];
        newItem[sanitizeHeaderName("индекс")] = baseRow[sanitizeHeaderName("индекс")];
        newItem[sanitizeHeaderName("ТЕЛЕФОН")] = baseRow[sanitizeHeaderName("ТЕЛЕФОН")];
        newItem[sanitizeHeaderName("серия паспорта")] = baseRow[sanitizeHeaderName("серия паспорта")];
        newItem[sanitizeHeaderName("номер паспорта")] = baseRow[sanitizeHeaderName("номер паспорта")];
        newItem[sanitizeHeaderName("дата выдачи")] = baseRow[sanitizeHeaderName("дата выдачи")];
        newItem[sanitizeHeaderName("дата рождения")] = baseRow[sanitizeHeaderName("дата рождения")];
        newItem[sanitizeHeaderName("ИНН")] = baseRow[sanitizeHeaderName("ИНН")];
        
        // Заполняем поля нового товара данными из формы
        newItem[itemNameColumn] = newItemName;
        newItem[itemQuantityColumn] = parsedNewItemQuantity;
        newItem[itemPriceColumn] = parseFloat(newItemPrice) || 0;
        newItem[itemLinkColumn] = newItemLink;
        newItem[itemWeightColumn] = parsedNewItemWeight;

        // Вычисляем новые общие значения для заказа
        // Берем текущие общие значения из первого элемента заказа
        const originalTotalItemsForOrder = parseFloat(baseRow[totalItemsColumn]) || 0;
        const originalTotalWeightForOrder = parseFloat(baseRow[totalWeightColumn]) || 0;

        // Прибавляем количество/вес нового товара к существующим общим значениям
        const updatedTotalItemsForOrder = originalTotalItemsForOrder + parsedNewItemQuantity;
        const updatedTotalWeightForOrder = originalTotalWeightForOrder + parsedNewItemWeight;


        // Обновляем все строки текущего заказа (и существующие, и только что добавленную)
        // с новыми общими значениями
        const orderRowsToProcess = [...currentOrderRows, newItem]; // Все строки этого заказа, включая новую

        orderRowsToProcess.forEach(row => {
            row[totalItemsColumn] = updatedTotalItemsForOrder;
            row[totalWeightColumn] = updatedTotalWeightForOrder;
        });

        // Добавляем все эти обработанные строки в финальный массив
        processedMatchedData.push(...orderRowsToProcess);
    }
    
    finalMatchedData = processedMatchedData; // Обновляем finalMatchedData
    // === КОНЕЦ ЛОГИКИ ДОБАВЛЕНИЯ СТРОК ИЗ ФОРМЫ ===

    // === ПЕРЕСТРОЙКА orderMap ===
    // orderMap теперь должен отражать новые индексы и структуру finalMatchedData
    // Это важно для логики снижения цен и подсветки
    const newOrderMap = new Map();
    for (let i = 0; i < finalMatchedData.length; ++i) {
        const rowObj = finalMatchedData[i];
        const orderId = rowObj?.[targetOrderIdColumn];
        if (orderId !== undefined && orderId !== null && orderId !== '') {
            if (!newOrderMap.has(orderId)) {
                newOrderMap.set(orderId, []);
            }
            newOrderMap.get(orderId).push(i);
        }
    }
    const orderMapForProcessing = newOrderMap; // Используем newOrderMap для следующих шагов

    // === ПОВТОРНАЯ ЛОГИКА УДАЛЕНИЯ ДУБЛИКАТОВ ОБЩЕГО ВЕСА ПОСЫЛКИ ===
    // Это нужно, чтобы "общий вес посылки" отображался только один раз на заказ
    // Применяется после всех изменений и добавления новых строк.
    for (const indices of orderMapForProcessing.values()) { // Используем актуальный orderMap
        if (indices.length <= 1) continue; // Нет дубликатов, если только один элемент

        const firstRowIndexInOrder = indices[0];
        const firstRowOfOrder = finalMatchedData[firstRowIndexInOrder];
        // Берем значение общего веса посылки из первой строки заказа (оно уже обновлено)
        const valueToKeep = firstRowOfOrder?.[totalWeightColumn];

        // Очищаем значение общего веса посылки в последующих строках того же заказа
        for (let i = 1; i < indices.length; ++i) {
            const currentRowIndex = indices[i];
            const row = finalMatchedData[currentRowIndex];
            row[totalWeightColumn] = ''; // Очищаем значение
        }
    }
    // === КОНЕЦ ПОВТОРНОЙ ЛОГИКИ УДАЛЕНИЯ ДУБЛИКАТОВ ОБЩЕГО ВЕСА ПОСЫЛКИ ===

    // === ЛОГИКА: ОБНОВЛЕНИЕ ДОПОЛНИТЕЛЬНЫХ ПОЛЕЙ "индекс" и "Контактное лицо" ===
    const indexColumn = sanitizeHeaderName("индекс");
    const contactPersonColumn = sanitizeHeaderName("Контактное лицо (телефон), получатель в РОССИИ");

    // Применяем значения из новых инпутов ко ВСЕМ строкам в finalMatchedData
    if (newIndexInputValue !== '') {
        finalMatchedData.forEach(row => {
            row[indexColumn] = newIndexInputValue;
        });
    }
    if (newContactPersonInputValue !== '') {
        finalMatchedData.forEach(row => {
            row[contactPersonColumn] = newContactPersonInputValue;
        });
    }
    // === КОНЕЦ ЛОГИКИ ОБНОВЛЕНИЯ ДОПОЛНИТЕЛЬНЫХ ПОЛЕЙ ===

    // === НОВАЯ ЛОГИКА: СНИЖЕНИЕ ЦЕН И ХРАНЕНИЕ ОРИГИНАЛЬНЫХ СУММ ДЛЯ ПОДСВЕТКИ ===
    const targetOrderSumValue = 15595; // Целевая сумма, к которой нужно привести заказы
    const originalOrderSums = new Map(); // orderId -> originalSum (для подсветки)

    // Проходим по каждому заказу (группе строк с одинаковым ID)
    for (const indices of orderMapForProcessing.values()) { // Используем orderMapForProcessing
        const orderIdValue = finalMatchedData[indices[0]][targetOrderIdColumn]; // ID текущего заказа

        // 1. Рассчитываем оригинальную сумму для этого заказа (до изменений цен)
        let currentOriginalOrderSum = 0;
        for (const originalRowIndex of indices) {
            const rowData = finalMatchedData[originalRowIndex];
            const quantity = parseFloat(rowData[itemQuantityColumn]) || 0;
            const price = parseFloat(rowData[itemPriceColumn]) || 0;
            currentOriginalOrderSum += quantity * price;
        }

        // Сохраняем оригинальные значения для использования в логике подсветки
        originalOrderSums.set(orderIdValue, currentOriginalOrderSum);

        // 2. Корректируем цены, если оригинальная сумма превышает целевую
        if (currentOriginalOrderSum > targetOrderSumValue && currentOriginalOrderSum > 0) {
            const reductionFactor = targetOrderSumValue / currentOriginalOrderSum;

            for (const originalRowIndex of indices) {
                const rowData = finalMatchedData[originalRowIndex];
                const quantity = parseFloat(rowData[itemQuantityColumn]) || 0;
                let currentPrice = parseFloat(rowData[itemPriceColumn]) || 0;

                if (quantity > 0 && currentPrice > 0) {
                    let newPrice = currentPrice * reductionFactor;
                    
                    rowData[itemPriceColumn] = Math.max(1, Math.round(newPrice));
                } else if (quantity === 0 && currentPrice > 0) {
                    // Если количество 0, но цена есть, это не влияет на общую сумму "кол-во * цена"
                }
            }
        }
    }
    // === КОНЕЦ ЛОГИКИ СНИЖЕНИЯ ЦЕН ===


    // === ЛОГИКА: ИЗМЕНЕНИЕ ПОРЯДКА СТОЛБЦОВ ===
    const customColumnOrder = [
        sanitizeHeaderName("Номер-посылки-(накладной)"),
        "ФАМИЛИЯ",
        "ИМЯ",
        "ОТЧЕСТВО",
        "АДРЕС",
        "ГОРОД",
        "ОБЛАСТЬ",
        "индекс",
        "ТЕЛЕФОН",
        sanitizeHeaderName("общ. Количество товаров в посылке (накладной)"),
        "количество вложений",
        "наименование на русском",
        "Цена товара",
        "Ссылка на товар",
        "серия паспорта",
        "номер паспорта",
        "дата выдачи",
        "дата рождения",
        "ИНН",
        "общий вес вложений",
        "общий вес посылки",
        sanitizeHeaderName("Контактное лицо (телефон), получатель в РОССИИ"),
    ];

    const getAllUniqueHeaders = (dataArray) => {
      const headers = new Set();
      dataArray.forEach(row => {
        Object.keys(row).forEach(key => headers.add(key));
      });
      return Array.from(headers);
    };

    let finalHeadersMatched = [...customColumnOrder];
    const actualHeadersMatched = getAllUniqueHeaders(finalMatchedData);
    actualHeadersMatched.forEach(header => {
      if (!finalHeadersMatched.includes(header)) {
        finalHeadersMatched.push(header);
      }
    });

    let finalHeadersUnmatched = [...customColumnOrder];
    const actualHeadersUnmatched = getAllUniqueHeaders(finalUnmatchedData);
    actualHeadersUnmatched.forEach(header => {
      if (!finalHeadersUnmatched.includes(header)) {
        finalHeadersUnmatched.push(header);
      }
    });
    // === КОНЕЦ ЛОГИКИ ИЗМЕНЕНИЯ ПОРЯДКА ===


    // Создаем wsMatched из ОТФИЛЬТРОВАННОГО, ОЧИЩЕННОГО И СКОПИРОВАННОГО массива с УКАЗАННЫМ ПОРЯДКОМ ЗАГОЛОВКОВ
    const wsMatched = XLSX.utils.json_to_sheet(finalMatchedData, { cellStyles: true, header: finalHeadersMatched });

    const range = wsMatched['!ref'] ? XLSX.utils.decode_range(wsMatched['!ref']) : null;

    // === ЛОГИКА: ПОДСВЕТКА ЗАКАЗОВ ===
    const highlightByValueColor = { rgb: "FFCC00" }; // Оранжевый для суммы заказа > 16000
    const highlightByQuantityColor = { rgb: "FFFF99" }; // Светло-желтый для "количество вложений" > 5

    if (range) {
        // Сначала определим, какие заказы должны быть оранжевыми
        const ordersToHighlightOrange = new Set();
        for (const indices of orderMapForProcessing.values()) {
            const orderIdValue = finalMatchedData[indices[0]][targetOrderIdColumn];
            const currentOriginalOrderSum = originalOrderSums.get(orderIdValue) || 0;
            if (currentOriginalOrderSum > 16000) {
                ordersToHighlightOrange.add(orderIdValue);
            }
        }

        // В этом Map будем хранить цвета, которые уже были применены к каждой ячейке.
        // Это позволит избежать перезаписи высокоприоритетного цвета низкоприоритетным.
        const cellColors = new Map(); // Key: cellRef (e.g., "A1"), Value: color (e.g., highlightByValueColor)

        // *** ИЗМЕНЕННАЯ ЛОГИКА ПРИОРИТЕТА: Сначала применяем ЖЕЛТЫЙ, затем ОРАНЖЕВЫЙ, но только если ячейка еще не ЖЕЛТАЯ ***

        // 1. Проходим по ВСЕМ строкам и сначала применяем светло-желтый цвет
        for (let R = range.s.r + 1; R <= range.e.r; ++R) { 
            const rowIndexInData = R - 1; 
            const rowData = finalMatchedData[rowIndexInData];
            if (!rowData) continue;

            const itemQuantity = parseFloat(rowData[itemQuantityColumn]) || 0;

            if (itemQuantity >= 5) {
                for (let C = range.s.c; C <= range.e.c; ++C) {
                    const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
                    const cell = wsMatched[cellRef];
                    if (cell) {
                        if (!cell.s) cell.s = {};
                        cell.s.fill = {
                            fgColor: highlightByQuantityColor,
                            patternType: "solid"
                        };
                        cellColors.set(cellRef, highlightByQuantityColor); // Записываем, что эта ячейка желтая
                    }
                }
            }
        }

        // 2. Затем, проходим снова, чтобы применить ОРАНЖЕВЫЙ, но только если ячейка еще НЕ БЫЛА окрашена в светло-желтый
        for (let R = range.s.r + 1; R <= range.e.r; ++R) { 
            const rowIndexInData = R - 1; 
            const rowData = finalMatchedData[rowIndexInData];
            if (!rowData) continue;

            const orderIdValue = rowData[targetOrderIdColumn];

            // Применяем оранжевый, только если заказ должен быть оранжевым И ячейка еще не желтая
            if (ordersToHighlightOrange.has(orderIdValue)) {
                for (let C = range.s.c; C <= range.e.c; ++C) {
                    const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
                    // Проверяем, был ли уже применен светло-желтый цвет к этой ячейке
                    if (!cellColors.has(cellRef) || cellColors.get(cellRef).rgb !== highlightByQuantityColor.rgb) {
                        const cell = wsMatched[cellRef];
                        if (cell) {
                            if (!cell.s) cell.s = {};
                            cell.s.fill = {
                                fgColor: highlightByValueColor,
                                patternType: "solid"
                            };
                            cellColors.set(cellRef, highlightByValueColor); // Записываем, что эта ячейка оранжевая
                        }
                    }
                }
            }
        }

        // Логика границ (выполняется отдельно, чтобы не влиять на заливку)
        for (let R = range.s.r + 1; R <= range.e.r; ++R) {
            const rowIndexInData = R - 1;
            const rowData = finalMatchedData[rowIndexInData];
            if (!rowData) continue;

            const orderIdValue = rowData[targetOrderIdColumn];
            const indicesForCurrentOrder = orderMapForProcessing.get(orderIdValue);

            if (indicesForCurrentOrder) {
                const isFirstRowOfOrder = rowIndexInData === indicesForCurrentOrder[0];
                const isLastRowOfOrder = rowIndexInData === indicesForCurrentOrder[indicesForCurrentOrder.length - 1];

                for (let C = range.s.c; C <= range.e.c; ++C) {
                    const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
                    const cell = wsMatched[cellRef];

                    if (cell) {
                        if (!cell.s) cell.s = {};

                        if (isFirstRowOfOrder) {
                            cell.s.border = {
                                ...(cell.s.border || {}),
                                top: borderStyle
                            };
                        }
                        if (isLastRowOfOrder) {
                            cell.s.border = {
                                ...(cell.s.border || {}),
                                bottom: borderStyle
                            };
                        }
                    }
                }
            }
        }
    }
    // === КОНЕЦ ЛОГИКИ ПОДСВЕТКИ ===


    // Создаем wsUnmatched - для этого листа подсветка не требуется по заданию
    const wsUnmatched = XLSX.utils.json_to_sheet(finalUnmatchedData, { header: finalHeadersUnmatched });

    XLSX.utils.book_append_sheet(wb, wsMatched, 'Совпавшие');
    XLSX.utils.book_append_sheet(wb, wsUnmatched, 'Не совпавшие');
    XLSX.writeFile(wb, 'Результат_сравнения.xlsx');
  };

  return (
    <div style={{ padding: '2rem' }}>
      <h2>Сравнение Excel по выбранным листам</h2>
      {/* Форма для выбора файлов */}
      <div style={{ marginTop: '2rem', border: '1px solid #a3e635', padding: '1rem', borderRadius: '8px' }}>
          <div style={{ display: 'flex', flexDirection: 'column' }}>
            <div>
              <label>Файл 1 (реестр): </label>
              <input
                type="file"
                accept=".xlsx, .xls"
                onChange={e => {
                  const file = e.target.files[0];
                  setFile1(file);
                  handleFileUpload(file, setWb1, setSheet1);
                }}
              />
              {wb1 && (
                <select value={sheet1} onChange={e => setSheet1(e.target.value)}>
                  {wb1.SheetNames.map(name => (
                    <option key={name} value={name}>
                      {name}
                    </option>
                  ))}
                </select>
              )}
            </div>
            <div style={{marginTop: '2rem'}}>
              <label>Файл 2 (заказы): </label>
              <input
                type="file"
                accept=".xlsx, .xls"
                onChange={e => {
                  const file = e.target.files[0];
                  setFile2(file);
                  handleFileUpload(file, setWb2, setSheet2);
                }}
              />
              {wb2 && (
                <select value={sheet2} onChange={e => setSheet2(e.target.value)}>
                  {wb2.SheetNames.map(name => (
                    <option key={name} value={name}>
                      {name}
                    </option>
                  ))}
                </select>
              )}
            </div>
        </div>
      </div>
      
      {/* Форма для добавления нового товара */}
      <div style={{ marginTop: '2rem', border: '1px solid #a3e635', padding: '1rem', borderRadius: '8px' }}>
          <h3>Данные для нового товара (добавится к каждому заказу)</h3>
          <div style={{ display: 'flex', flexDirection: 'column' }}>
              <div style={{ marginBottom: '1rem'}}>
                  <label style={{ display: 'block', marginBottom: '5px' }}>Название товара (на русском):</label>
                  <input type="text" value={newItemName} onChange={e => setNewItemName(e.target.value)} 
                         style={{ width: '90%', padding: '8px', borderRadius: '4px', border: '1px solid #ddd' }} />
              </div>
              <div style={{ marginBottom: '1rem'}}>
                  <label style={{ display: 'block', marginBottom: '5px' }}>Количество вложений (нового товара):</label>
                  <input type="number" value={newItemQuantity} onChange={e => setNewItemQuantity(e.target.value)} 
                         style={{ width: '90%', padding: '8px', borderRadius: '4px', border: '1px solid #ddd' }} />
              </div>
              <div style={{ marginBottom: '1rem'}}>
                  <label style={{ display: 'block', marginBottom: '5px' }}>Цена товара (нового товара):</label>
                  <input type="number" step="1" value={newItemPrice} onChange={e => setNewItemPrice(e.target.value)} 
                         style={{ width: '90%', padding: '8px', borderRadius: '4px', border: '1px solid #ddd' }} />
              </div>
              <div style={{ marginBottom: '1rem'}}>
                  <label style={{ display: 'block', marginBottom: '5px' }}>Ссылка на товар (нового товара):</label>
                  <input type="text" value={newItemLink} onChange={e => setNewItemLink(e.target.value)} 
                         style={{ width: '90%', padding: '8px', borderRadius: '4px', border: '1px solid #ddd' }} />
              </div>
              <div style={{ marginBottom: '1rem'}}>
                  <label style={{ display: 'block', marginBottom: '5px' }}>Вес нового вложения (для "общий вес вложений"):</label>
                  <input type="number" step="0.01" value={newItemWeight} onChange={e => setNewItemWeight(e.target.value)} 
                         style={{ width: '90%', padding: '8px', borderRadius: '4px', border: '1px solid #ddd' }} />
              </div>
          </div>
      </div>

      {/* НОВЫЙ БЛОК ВВОДА ДАННЫХ ДЛЯ ВСЕХ СТРОК */}
      <div style={{ marginTop: '2rem', border: '1px solid #a3e635', padding: '1rem', borderRadius: '8px' }}>
          <h3>Дополнительные данные для всех строк</h3>
          <div style={{ display: 'flex', flexDirection:'column', gap: '15px', maxWidth: '700px' }}>
              <div>
                  <label style={{ display: 'block', marginBottom: '5px' }}>Индекс:</label>
                  <input type="text" value={newIndexInputValue} onChange={e => setNewIndexInputValue(e.target.value)} 
                         style={{ width: '90%', padding: '8px', borderRadius: '4px', border: '1px solid #a3e635' }} />
              </div>
              <div>
                  <label style={{ display: 'block', marginBottom: '5px' }}>Контактное лицо (телефон), получатель в РОССИИ:</label>
                  <input type="text" value={newContactPersonInputValue} onChange={e => setNewContactPersonInputValue(e.target.value)} 
                         style={{ width: '90%', padding: '8px', borderRadius: '4px', border: '1px solid #a3e635' }} />
              </div>
          </div>
      </div>

      <button
        style={{ marginTop: '2rem', padding: '10px 20px', fontSize: '16px', cursor: 'pointer', borderRadius: '5px', border: 'none', backgroundColor: '#4CAF50', color: 'white' }}
        onClick={handleCompare}
        disabled={!wb1 || !wb2}
      >
        Сделать дело!
      </button>
    </div>
  );
}