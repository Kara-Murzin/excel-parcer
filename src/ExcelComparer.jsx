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

    // Создаем пустой рабочий документ для результатов
    const wb = XLSX.utils.book_new();

    const borderStyle = {
      style: "thin",
      color: { rgb: "000000" }
    };

    // Построим карту: ключ = номер посылки, значения = индексы строк (в Excel)
    // Важно: orderMap строится на основе 'matched' массива, который еще не модифицирован по весу.
    // Однако, индексы строк используются для прямого доступа к 'matched' массиву.
    const orderMap = new Map();
    // Мы должны временно создать wsMatched, чтобы получить диапазон,
    // или пересчитать его из `matched.length`.
    // Для более надежного получения диапазона после сортировки/фильтрации,
    // лучше использовать данные после всех преобразований.
    // Переместим создание wsMatched и range после обработки веса.

    // === ЛОГИКА УДАЛЕНИЯ ДУБЛИКАТОВ ВЕСА ===
    const columns = Object.keys(sheetData1[0] || {}); // Добавил || {} на случай пустого sheetData1
    const weightColumnName = "包裹重量一个包裹一次备注就行общий вес посылки";

    if (!columns.includes(weightColumnName)) { // Проверка на наличие колонки веса
      alert(`Колонка "${weightColumnName}" не найдена в файле.`);
      return;
    }
    
    // Сначала построим orderMap, чтобы знать, какие строки относятся к одному заказу
    // Индексы здесь относятся к массиву `matched` (0-based)
    for (let i = 0; i < matched.length; ++i) {
      const rowObj = matched[i];
      const orderId = rowObj?.[targetColumnName];

      if (orderId !== undefined && orderId !== null && orderId !== '') { // Убедимся, что orderId не пустой
        if (!orderMap.has(orderId)) {
          orderMap.set(orderId, []);
        }
        orderMap.get(orderId).push(i); // Сохраняем 0-based индекс в matched
      }
    }

    // Удаляем дубли веса
    for (const indices of orderMap.values()) {
      if (indices.length <= 1) continue; // Если только одна строка для заказа, нет дубликатов веса

      // Индекс первой строки для текущего заказа в массиве `matched`
      const firstRowIndexInMatched = indices[0];
      const firstWeight = normalize(matched[firstRowIndexInMatched]?.[weightColumnName]);

      // Проходим по остальным строкам того же заказа и очищаем вес, если он дублируется
      for (let i = 1; i < indices.length; ++i) {
        const currentRowIndexInMatched = indices[i];
        const row = matched[currentRowIndexInMatched];

        if (normalize(row?.[weightColumnName]) === firstWeight) {
          row[weightColumnName] = ''; // Очищаем вес
        }
      }
    }

    // === КОНЕЦ ЛОГИКИ УДАЛЕНИЯ ДУБЛИКАТОВ ВЕСА ===

    // !!! ВОТ ЗДЕСЬ ПЕРЕНОСИМ СОЗДАНИЕ wsMatched !!!
    // Создаем wsMatched из МОДИФИЦИРОВАННОГО массива matched
    const wsMatched = XLSX.utils.json_to_sheet(matched, { cellStyles: true });

    // Теперь получаем диапазон из свежесозданного wsMatched
    const range = wsMatched['!ref'] ? XLSX.utils.decode_range(wsMatched['!ref']) : null;

    // Применяем стили (границы)
    if (range) { // Убедимся, что диапазон существует
        for (const indices of orderMap.values()) {
            const firstRowExcel = indices[0] + 1; // Convert 0-based matched index to 1-based Excel row
            const lastRowExcel = indices[indices.length - 1] + 1; // Convert 0-based matched index to 1-based Excel row

            for (let C = range.s.c; C <= range.e.c; ++C) {
                // Стиль верхней границы для первой строки заказа
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

                // Стиль нижней границы для последней строки заказа
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


    const wsUnmatched = XLSX.utils.json_to_sheet(unmatched);

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