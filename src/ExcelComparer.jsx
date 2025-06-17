import React, { useState } from 'react';
import * as XLSX from 'xlsx-js-style';

export default function ExcelComparer() {
  const [file1, setFile1] = useState(null);
  const [wb1, setWb1] = useState(null);
  const [sheet1, setSheet1] = useState('');

  const [file2, setFile2] = useState(null);
  const [wb2, setWb2] = useState(null);
  const [sheet2, setSheet2] = useState('');

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

    // === ЛОГИКА УДАЛЕНИЯ ДУБЛИКАТОВ ВЕСА ===
    const columns = Object.keys(sheetData1[0] || {});
    const weightColumnName = "包裹重量一个包裹一次备注就行общий вес посылки";

    if (!columns.includes(weightColumnName)) {
      alert(`Колонка "${weightColumnName}" не найдена в файле.`);
      return;
    }
    
    // orderMap теперь будет содержать 0-based индексы строк matched
    const orderMap = new Map();
    for (let i = 0; i < matched.length; ++i) {
      const rowObj = matched[i];
      const orderId = rowObj?.[targetColumnName];

      if (orderId !== undefined && orderId !== null && orderId !== '') {
        if (!orderMap.has(orderId)) {
          orderMap.set(orderId, []);
        }
        orderMap.get(orderId).push(i); // Сохраняем индекс строки в массиве matched
      }
    }

    for (const indices of orderMap.values()) {
      if (indices.length <= 1) continue;

      const firstRowIndexInMatched = indices[0];
      const firstWeight = normalize(matched[firstRowIndexInMatched]?.[weightColumnName]);

      for (let i = 1; i < indices.length; ++i) {
        const currentRowIndexInMatched = indices[i];
        const row = matched[currentRowIndexInMatched];

        if (normalize(row?.[weightColumnName]) === firstWeight) {
          row[weightColumnName] = '';
        }
      }
    }
    // === КОНЕЦ ЛОГИКИ УДАЛЕНИЯ ДУБЛИКАТОВ ВЕСА ===


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
        sanitizeHeaderName("Контактное лицо (телефон), получатель  в  РОССИИ"),
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

    // === НОВАЯ ЛОГИКА: ПОДСВЕТКА ЗАКАЗОВ ===
    const highlightByCountColor = { rgb: "FFFF99" }; // Светло-желтый для количества товаров > 5
    const highlightByValueColor = { rgb: "FFCC00" }; // Оранжевый для суммы заказа > 16000

    // Очищенные названия столбцов для расчетов
    const quantityColumn = sanitizeHeaderName("количество вложений");
    const priceColumn = sanitizeHeaderName("Цена товара");


    if (range) {
        // Проходим по orderMap для применения стилей
        for (const indices of orderMap.values()) {
            const itemCount = indices.length; // Количество строк (товаров) для текущего заказа

            // Высчитываем общую сумму заказа
            let orderSum = 0;
            for (const originalRowIndex of indices) {
                // Получаем данные строки из finalMatchedData по ее индексу
                const rowData = finalMatchedData[originalRowIndex];
                
                // Парсим значения как числа, приводим к 0, если не число
                const quantity = parseFloat(rowData[quantityColumn]) || 0;
                const price = parseFloat(rowData[priceColumn]) || 0;
                
                orderSum += quantity * price;
            }

            // Определяем, какие условия подсветки выполнены
            const shouldHighlightByCount = itemCount > 5;
            const shouldHighlightByValue = orderSum > 16000;

            let currentHighlightColor = null;

            // Приоритет: если сумма заказа большая, используем оранжевый цвет
            if (shouldHighlightByValue) {
                currentHighlightColor = highlightByValueColor;
            } 
            // Иначе, если товаров много, используем светло-желтый
            else if (shouldHighlightByCount) {
                currentHighlightColor = highlightByCountColor;
            }
            
            // Индексы строк в Excel (1-based), принадлежащих текущему заказу
            const firstRowExcel = indices[0] + 1;
            const lastRowExcel = indices[indices.length - 1] + 1;

            // Применяем стили ко всем ячейкам в строках этого заказа
            for (let R = firstRowExcel; R <= lastRowExcel; ++R) {
                for (let C = range.s.c; C <= range.e.c; ++C) {
                    const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
                    const cell = wsMatched[cellRef];

                    if (cell) {
                        if (!cell.s) cell.s = {};

                        // Применяем фоновую заливку, если цвет определен
                        if (currentHighlightColor) {
                            cell.s.fill = {
                                fgColor: currentHighlightColor,
                                patternType: "solid"
                            };
                        }

                        // Сохраняем логику границ (верхняя для первой строки, нижняя для последней)
                        if (R === firstRowExcel) {
                            cell.s.border = {
                                ...(cell.s.border || {}),
                                top: borderStyle
                            };
                        }
                        if (R === lastRowExcel) {
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
    // === КОНЕЦ НОВОЙ ЛОГИКИ ===


    // Создаем wsUnmatched - для этого листа подсветка не требуется по заданию
    const wsUnmatched = XLSX.utils.json_to_sheet(finalUnmatchedData, { header: finalHeadersUnmatched });

    XLSX.utils.book_append_sheet(wb, wsMatched, 'Совпавшие');
    XLSX.utils.book_append_sheet(wb, wsUnmatched, 'Не совпавшие');
    XLSX.writeFile(wb, 'Результат_сравнения.xlsx');
  };

  return (
    <div style={{ padding: '2rem' }}>
      <h2>Сравнение Excel по выбранным листам</h2>

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

      <div style={{ marginTop: '1rem' }}>
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

      <button
        style={{ marginTop: '2rem' }}
        onClick={handleCompare}
        disabled={!wb1 || !wb2}
      >
        Сравнить и скачать
      </button>
    </div>
  );
}