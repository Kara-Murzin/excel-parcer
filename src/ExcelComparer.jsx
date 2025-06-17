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
    
    const orderMap = new Map();
    for (let i = 0; i < matched.length; ++i) {
      const rowObj = matched[i];
      const orderId = rowObj?.[targetColumnName];

      if (orderId !== undefined && orderId !== null && orderId !== '') {
        if (!orderMap.has(orderId)) {
          orderMap.set(orderId, []);
        }
        orderMap.get(orderId).push(i);
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


    // === НОВАЯ ЛОГИКА: УДАЛЕНИЕ СТОЛБЦОВ ПО СПИСКУ ===
    const columnsToRemove = [
      "收件人", "收件地址", "收件人城市", "收件人州省", "国家简码",
      "收件国家", "寄件国家", "寄件申报品名", "收件申报品名", "海关编码",
      "重量", // Обратите внимание: у вас есть "包裹重量..." и "重量". Если это одно и то же, оставьте только одно в списке.
      "内部单号", "订单号", "发货时间", "物流方式", "袋号", "分拣人",
      "长宽高", "体积重", "仓库名称", "是否带电", "产品信息",
      "产品信息sku1", "SKU", "ENGLISH", "Брэнд/изготовитель", "Материал",
      "Группа товаров", "单品价格цена за 1 вложение(товар) в юанях","Цена поставщика", "__EMPTY"
    ];

    // Функция для фильтрации объектов данных, удаляющая указанные столбцы
    const filterColumns = (dataArray) => {
      return dataArray.map(row => {
        const newRow = {};
        for (const key in row) {
          // Если ключ не содержится в списке columnsToRemove, добавляем его в новую строку
          if (!columnsToRemove.includes(key)) {
            newRow[key] = row[key];
          }
        }
        return newRow;
      });
    };

    const filteredMatched = filterColumns(matched);
    const filteredUnmatched = filterColumns(unmatched);
    // === КОНЕЦ НОВОЙ ЛОГИКИ ===


    // Создаем wsMatched из ОТФИЛЬТРОВАННОГО массива
    const wsMatched = XLSX.utils.json_to_sheet(filteredMatched, { cellStyles: true });

    // Теперь получаем диапазон из свежесозданного wsMatched
    const range = wsMatched['!ref'] ? XLSX.utils.decode_range(wsMatched['!ref']) : null;

    // Применяем стили (границы)
    if (range) {
        // Мы должны пересоздать orderMap, основываясь на filteredMatched,
        // или найти соответствия original_row_index -> filtered_row_index
        // Но так как мы удаляем столбцы, а не строки, старые индексы (из matched) все еще применимы,
        // если считать их как индекс в списке (после фильтрации столбцов, порядок строк не меняется)
        // Единственная проблема: если какая-то строка была удалена из `matched` полностью,
        // но в вашем случае мы только очищаем поля, не удаляем строки из `matched`.
        // Поэтому, `orderMap` (с индексами в `matched`) можно использовать напрямую.
        for (const indices of orderMap.values()) {
            // Конвертируем 0-based индекс из `matched` в 1-based Excel row.
            // Индексы в `orderMap` ссылаются на исходный `matched` массив,
            // который имеет ту же последовательность строк, что и `filteredMatched`.
            const firstRowExcel = indices[0] + 1;
            const lastRowExcel = indices[indices.length - 1] + 1;

            for (let C = range.s.c; C <= range.e.c; ++C) {
                const cellTop = XLSX.utils.encode_cell({ r: firstRowExcel, c: C });
                if (wsMatched[cellTop]) {
                    wsMatched[cellTop].s = {
                        ...wsMatched[cellTop].s,
                        border: {
                            ...(wsMatched[cellTop].s?.border || {}),
                            top: borderStyle
                        }
                    };
                }

                const cellBottom = XLSX.utils.encode_cell({ r: lastRowExcel, c: C });
                if (wsMatched[cellBottom]) {
                    wsMatched[cellBottom].s = {
                        ...wsMatched[cellBottom].s,
                        border: {
                            ...(wsMatched[cellBottom].s?.border || {}),
                            bottom: borderStyle
                        }
                    };
                }
            }
        }
    }


    const wsUnmatched = XLSX.utils.json_to_sheet(filteredUnmatched); // Используем отфильтрованные данные

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