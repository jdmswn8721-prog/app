import React, { useState, useRef, useEffect, useCallback } from 'react';
import * as XLSX from 'xlsx';
import html2canvas from 'html2canvas';

const refreshCellIds = (container) => {
  if (!container) return;
  const table = container.querySelector('table');
  if (!table) return;

  const rows = Array.from(table.querySelectorAll('tr'));
  const occupancy = [];

  rows.forEach((row, rowIndex) => {
    occupancy[rowIndex] = occupancy[rowIndex] || [];
    const cells = Array.from(row.querySelectorAll('th, td'));
    let colPointer = 0;

    cells.forEach((cell) => {
      while (occupancy[rowIndex][colPointer]) {
        colPointer++;
      }

      const colspan = parseInt(cell.getAttribute('colspan')) || 1;
      const rowspan = parseInt(cell.getAttribute('rowspan')) || 1;
      const cellId = `cell-${rowIndex}-${colPointer}`;
      cell.setAttribute('data-cell-id', cellId);

      for (let rs = 0; rs < rowspan; rs++) {
        const targetRow = rowIndex + rs;
        occupancy[targetRow] = occupancy[targetRow] || [];
        for (let cs = 0; cs < colspan; cs++) {
          occupancy[targetRow][colPointer + cs] = true;
        }
      }

      colPointer += colspan;
    });
  });
};

const validateTableStructure = (container) => {
  if (!container) return true;
  const table = container.querySelector('table');
  if (!table) return true;

  const rows = Array.from(table.querySelectorAll('tr'));
  if (rows.length === 0) return true;

  const occupancy = [];
  let maxCols = 0;

  for (let rowIndex = 0; rowIndex < rows.length; rowIndex++) {
    const row = rows[rowIndex];
    occupancy[rowIndex] = occupancy[rowIndex] || [];
    const cells = Array.from(row.querySelectorAll('th, td'));
    let colPointer = 0;

    for (const cell of cells) {
      while (occupancy[rowIndex][colPointer]) {
        colPointer++;
      }

      const colspan = parseInt(cell.getAttribute('colspan')) || 1;
      const rowspan = parseInt(cell.getAttribute('rowspan')) || 1;

      for (let rs = 0; rs < rowspan; rs++) {
        const targetRow = rowIndex + rs;
        occupancy[targetRow] = occupancy[targetRow] || [];
        for (let cs = 0; cs < colspan; cs++) {
          const targetCol = colPointer + cs;
          if (occupancy[targetRow][targetCol]) {
            return false;
          }
          occupancy[targetRow][targetCol] = true;
        }
      }

      colPointer += colspan;
    }

    maxCols = Math.max(maxCols, colPointer);
  }

  for (let r = 0; r < occupancy.length; r++) {
    const row = occupancy[r] || [];
    for (let c = 0; c < maxCols; c++) {
      if (!row[c]) {
        return false;
      }
    }
  }

  return true;
};

const restoreTableHtml = (container, html) => {
  if (!container) return;
  container.innerHTML = html;
  refreshCellIds(container);
};

function App() {
  const [sheets, setSheets] = useState([]);
  const [activeSheetIndex, setActiveSheetIndex] = useState(0);
  const [viewMode, setViewMode] = useState('edit');
  const [error, setError] = useState('');
  const [fileName, setFileName] = useState('');
  const tableRefs = useRef([]);
  const editSessionRef = useRef({ active: false, snapshot: '', sheetIndex: null });

  const [selectedCells, setSelectedCells] = useState(new Set());
  const [isSelecting, setIsSelecting] = useState(false);
  const [selectionStart, setSelectionStart] = useState(null);

  const [isMergeMode, setIsMergeMode] = useState(false);

  const [undoStack, setUndoStack] = useState([]);
  const [redoStack, setRedoStack] = useState([]);

  useEffect(() => {
    if (viewMode !== 'edit') return;
    const container = tableRefs.current[activeSheetIndex];
    if (!container) return;
    refreshCellIds(container);
  }, [sheets, activeSheetIndex, viewMode]);

  const worksheetToHtmlTable = (worksheet) => {
    if (!worksheet || !worksheet['!ref']) {
      return '<table></table>';
    }

    try {
      const range = XLSX.utils.decode_range(worksheet['!ref']);
      const merges = worksheet['!merges'] || [];

      const mergeMap = new Map();
      merges.forEach((merge) => {
        const { s, e } = merge;
        const colspan = e.c - s.c + 1;
        const rowspan = e.r - s.r + 1;

        for (let r = s.r; r <= e.r; r++) {
          for (let c = s.c; c <= e.c; c++) {
            const key = `${r},${c}`;
            mergeMap.set(key, {
              startRow: s.r,
              startCol: s.c,
              endRow: e.r,
              endCol: e.c,
              colspan,
              rowspan,
              isStart: r === s.r && c === s.c,
            });
          }
        }
      });

      let html = '<table>';

      for (let R = range.s.r; R <= range.e.r; ++R) {
        html += '<tr>';

        for (let C = range.s.c; C <= range.e.c; ++C) {
          const cellKey = `${R},${C}`;
          const mergeInfo = mergeMap.get(cellKey);

          if (mergeInfo) {
            if (!mergeInfo.isStart) {
              continue;
            }

            const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
            const cell = worksheet[cellAddress];

            let cellValue = '';
            if (cell) {
              if (cell.v !== undefined && cell.v !== null) {
                if (cell.t === 'd') {
                  const date = new Date(cell.v);
                  cellValue = date.toLocaleString('ko-KR');
                } else {
                  cellValue = String(cell.v);
                }
              }
            }

            const tag = R === range.s.r ? 'th' : 'td';
            const colspanAttr = mergeInfo.colspan > 1 ? ` colspan="${mergeInfo.colspan}"` : '';
            const rowspanAttr = mergeInfo.rowspan > 1 ? ` rowspan="${mergeInfo.rowspan}"` : '';

            html += `<${tag}${colspanAttr}${rowspanAttr}>${escapeHtml(cellValue)}</${tag}>`;
          } else {
            const cellAddress = XLSX.utils.encode_cell({ r: R, c: C });
            const cell = worksheet[cellAddress];

            let cellValue = '';
            if (cell) {
              if (cell.v !== undefined && cell.v !== null) {
                if (cell.t === 'd') {
                  const date = new Date(cell.v);
                  cellValue = date.toLocaleString('ko-KR');
                } else {
                  cellValue = String(cell.v);
                }
              }
            }

            const tag = R === range.s.r ? 'th' : 'td';
            html += `<${tag}>${escapeHtml(cellValue)}</${tag}>`;
          }
        }

        html += '</tr>';
      }

      html += '</table>';
      return html;
    } catch (err) {
      console.error('테이블 생성 오류:', err);
      return '<table></table>';
    }
  };

  const escapeHtml = (text) => {
    if (!text) return '';
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
  };

  const removeEmptyRowsAndCols = (tableHtml) => {
    if (!tableHtml) return tableHtml;

    try {
      const parser = new DOMParser();
      const doc = parser.parseFromString(tableHtml, 'text/html');
      const table = doc.querySelector('table');

      if (!table) return tableHtml;

      const rows = Array.from(table.querySelectorAll('tr'));
      if (rows.length === 0) return tableHtml;

      const rowData = rows.map((row) => {
        const cells = Array.from(row.querySelectorAll('th, td'));
        return cells.map((cell) => {
          const text = cell.textContent || cell.innerText || '';
          return text.trim();
        });
      });

      const emptyRowIndices = new Set();
      rowData.forEach((cells, rowIndex) => {
        const hasContent = cells.some((cellText) => cellText.length > 0);
        if (!hasContent) {
          emptyRowIndices.add(rowIndex);
        }
      });

      const emptyColIndices = new Set();
      if (rowData.length > 0) {
        const colCount = rowData[0].length;
        for (let colIndex = 0; colIndex < colCount; colIndex++) {
          const hasContent = rowData.some((cells, rowIndex) => {
            if (emptyRowIndices.has(rowIndex)) return false;
            return cells[colIndex] && cells[colIndex].length > 0;
          });
          if (!hasContent) {
            emptyColIndices.add(colIndex);
          }
        }
      }

      emptyRowIndices.forEach((rowIndex) => {
        if (rows[rowIndex]) {
          rows[rowIndex].remove();
        }
      });

      const remainingRows = Array.from(table.querySelectorAll('tr'));
      remainingRows.forEach((row) => {
        const cells = Array.from(row.querySelectorAll('th, td'));
        emptyColIndices.forEach((colIndex) => {
          if (cells[colIndex]) {
            cells[colIndex].remove();
          }
        });
      });

      return table.outerHTML;
    } catch (err) {
      console.error('빈 행/열 제거 오류:', err);
      return tableHtml;
    }
  };

  const handleFileUpload = (event) => {
    const file = event.target.files[0];
    if (!file) return;

    const isValidType =
      file.name.endsWith('.xlsx') ||
      file.name.endsWith('.xls') ||
      file.name.endsWith('.xlsm') ||
      file.name.endsWith('.csv') ||
      file.type === 'text/csv' ||
      file.type === 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' ||
      file.type === 'application/vnd.ms-excel' ||
      file.type === 'application/vnd.ms-excel.sheet.macroEnabled.12';

    if (!isValidType) {
      setError('지원하는 파일 형식: 엑셀(.xlsx, .xls, .xlsm) 또는 CSV(.csv)');
      return;
    }

    setError('');
    const reader = new FileReader();
    const isCSV = file.name.endsWith('.csv');

    reader.onload = (e) => {
      try {
        let workbook;
        const baseFileName = file.name.replace(/\.(xlsx|xls|xlsm|csv)$/i, '');

        if (isCSV) {
          let text = e.target.result;
          if (text.charCodeAt(0) === 0xfeff) {
            text = text.slice(1);
          }
          workbook = XLSX.read(text, {
            type: 'string',
            codepage: 65001,
          });
        } else {
          const data = new Uint8Array(e.target.result);
          workbook = XLSX.read(data, { type: 'array' });
        }

        if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
          setError('파일에서 시트를 찾을 수 없습니다.');
          return;
        }

        const extractedSheets = workbook.SheetNames.map((sheetName, index) => {
          const worksheet = workbook.Sheets[sheetName];
          let htmlTable = worksheetToHtmlTable(worksheet);
          const finalSheetName = isCSV ? baseFileName : sheetName || baseFileName;

          return {
            name: finalSheetName,
            content: htmlTable,
            index,
          };
        });

        setSheets(extractedSheets);
        setActiveSheetIndex(0);
        setFileName(baseFileName);
        setSelectedCells(new Set());
        setUndoStack([]);
        setRedoStack([]);
      } catch (err) {
        setError('파일 읽기 중 오류가 발생했습니다: ' + err.message);
        console.error('파일 읽기 오류:', err);
      }
    };

    reader.onerror = () => {
      setError('파일 읽기 실패');
    };

    if (isCSV) {
      reader.readAsText(file, 'UTF-8');
    } else {
      reader.readAsArrayBuffer(file);
    }
  };

  const getCellId = (cell) => {
    if (!cell) return null;
    if (cell.hasAttribute('data-cell-id')) {
      return cell.getAttribute('data-cell-id');
    }

    const container = cell.closest('.table-container');
    if (!container) return null;

    refreshCellIds(container);
    return cell.getAttribute('data-cell-id');
  };

  const handleCellSelect = (e) => {
    if (viewMode !== 'edit') return;

    const cell = e.target;
    if (cell.tagName !== 'TD' && cell.tagName !== 'TH') return;
    if (cell.isContentEditable) return;

    if (e.detail === 2) {
      cell.contentEditable = true;
      cell.focus();
      return;
    }

    e.preventDefault();

    let cellId = cell.getAttribute('data-cell-id');
    if (!cellId) {
      cellId = getCellId(cell);
    }

    if (!cellId) return;

    if (isMergeMode) {
      setSelectedCells((prev) => {
        const newSet = new Set(prev);
        if (newSet.has(cellId)) {
          newSet.delete(cellId);
        } else {
          newSet.add(cellId);
        }
        return newSet;
      });
      return;
    }

    if (e.ctrlKey || e.metaKey) {
      setSelectedCells((prev) => {
        const newSet = new Set(prev);
        if (newSet.has(cellId)) {
          newSet.delete(cellId);
        } else {
          newSet.add(cellId);
        }
        return newSet;
      });
    } else {
      setSelectedCells(new Set([cellId]));
      setSelectionStart(cellId);
    }
  };

  const handleCellMouseDown = (e) => {
    if (viewMode !== 'edit') return;

    const cell = e.target;
    if (cell.tagName !== 'TD' && cell.tagName !== 'TH') return;
    if (cell.isContentEditable) return;

    if (e.detail === 2) return;

    if (isMergeMode) {
      e.preventDefault();
      return;
    }

    e.preventDefault();
    setIsSelecting(true);

    let cellId = cell.getAttribute('data-cell-id');
    if (!cellId) {
      cellId = getCellId(cell);
    }

    if (cellId) {
      setSelectionStart(cellId);
      setSelectedCells(new Set([cellId]));
    }
  };

  const handleCellMouseOver = (e) => {
    if (!isSelecting || viewMode !== 'edit') return;

    const cell = e.target;
    if (cell.tagName !== 'TD' && cell.tagName !== 'TH') return;
    if (cell.isContentEditable) return;

    let cellId = cell.getAttribute('data-cell-id');
    if (!cellId) {
      cellId = getCellId(cell);
    }

    if (!cellId || !selectionStart) return;

    const startId = selectionStart;
    const [startRow, startCol] = startId.replace('cell-', '').split('-').map(Number);
    const [endRow, endCol] = cellId.replace('cell-', '').split('-').map(Number);

    const rowDiff = Math.abs(endRow - startRow);
    const colDiff = Math.abs(endCol - startCol);
    const isVerticalDrag = e.shiftKey || (rowDiff > colDiff && rowDiff > 0);

    const minRow = Math.min(startRow, endRow);
    const maxRow = Math.max(startRow, endRow);
    const minCol = isVerticalDrag ? startCol : Math.min(startCol, endCol);
    const maxCol = isVerticalDrag ? startCol : Math.max(startCol, endCol);

    const table = cell.closest('table');
    if (!table) return;

    const rows = Array.from(table.querySelectorAll('tr'));
    const selectedSet = new Set();

    if (isVerticalDrag) {
      const selectedCellIds = new Set();

      for (let r = minRow; r <= maxRow && r < rows.length; r++) {
        const occupiedPositions = new Set();
        for (let prevRow = 0; prevRow < r; prevRow++) {
          if (prevRow >= rows.length) break;
          const prevRowCells = Array.from(rows[prevRow].querySelectorAll('th, td'));
          let prevCol = 0;

          prevRowCells.forEach((prevCell) => {
            const prevColspan = parseInt(prevCell.getAttribute('colspan')) || 1;
            const prevRowspan = parseInt(prevCell.getAttribute('rowspan')) || 1;
            const prevCellStartRow = prevRow;
            const prevCellEndRow = prevCellStartRow + prevRowspan - 1;

            if (r >= prevCellStartRow && r <= prevCellEndRow) {
              for (let col = prevCol; col < prevCol + prevColspan; col++) {
                occupiedPositions.add(`${r}-${col}`);
              }
            }

            prevCol += prevColspan;
          });
        }

        const row = rows[r];
        const cells = Array.from(row.querySelectorAll('th, td'));
        let currentCol = 0;

        for (let i = 0; i < cells.length; i++) {
          const c = cells[i];
          const colspan = parseInt(c.getAttribute('colspan')) || 1;
          const rowspan = parseInt(c.getAttribute('rowspan')) || 1;

          while (occupiedPositions.has(`${r}-${currentCol}`)) {
            currentCol++;
          }

          const cellStartCol = currentCol;
          const cellEndCol = currentCol + colspan - 1;
          const cellStartRow = r;

          if (minCol >= cellStartCol && minCol <= cellEndCol) {
            const cellIdForCell = getCellId(c);
            if (cellIdForCell && !selectedCellIds.has(cellIdForCell)) {
              selectedCellIds.add(cellIdForCell);
              selectedSet.add(cellIdForCell);
            }
            break;
          }

          currentCol += colspan;
        }
      }
    } else {
      for (let r = minRow; r <= maxRow && r < rows.length; r++) {
        const row = rows[r];
        const cells = Array.from(row.querySelectorAll('th, td'));
        let currentCol = 0;

        for (let i = 0; i < cells.length; i++) {
          const c = cells[i];
          const colspan = parseInt(c.getAttribute('colspan')) || 1;

          if (currentCol <= maxCol && currentCol + colspan - 1 >= minCol) {
            const cellIdForCell = getCellId(c);
            if (cellIdForCell) {
              selectedSet.add(cellIdForCell);
            }
          }

          currentCol += colspan;
        }
      }
    }

    setSelectedCells(selectedSet);
  };

  const handleCellMouseUp = () => {
    setIsSelecting(false);
    setSelectionStart(null);
  };

  const handleCellDoubleClick = (e) => {
    if (viewMode !== 'edit') return;
    const cell = e.target;
    if (cell.tagName === 'TD' || cell.tagName === 'TH') {
      const container = tableRefs.current[activeSheetIndex];
      if (
        container &&
        (!editSessionRef.current.active || editSessionRef.current.sheetIndex !== activeSheetIndex)
      ) {
        editSessionRef.current = {
          active: true,
          snapshot: container.innerHTML,
          sheetIndex: activeSheetIndex,
        };
      }
      cell.contentEditable = true;
      cell.focus();
    }
  };

  const handleCellBlur = (e) => {
    if (viewMode !== 'edit') return;

    if (e.target.tagName === 'TD' || e.target.tagName === 'TH') {
      e.target.contentEditable = false;

      const container = tableRefs.current[activeSheetIndex];
      if (!container) {
        editSessionRef.current = { active: false, snapshot: '', sheetIndex: null };
        return;
      }

      refreshCellIds(container);
      const containerHtml = container.innerHTML;

      setSheets((prev) =>
        prev.map((sheet, idx) =>
          idx === activeSheetIndex ? { ...sheet, content: containerHtml } : sheet
        )
      );

      if (
        editSessionRef.current.active &&
        editSessionRef.current.sheetIndex === activeSheetIndex
      ) {
        const snapshot = editSessionRef.current.snapshot || '';
        if (snapshot !== containerHtml) {
          setUndoStack((prev) => [...prev, snapshot]);
          setRedoStack([]);
        }
        editSessionRef.current = { active: false, snapshot: '', sheetIndex: null };
      }
    }
  };

  const handleToggleMergeMode = () => {
    setIsMergeMode(!isMergeMode);
    if (!isMergeMode) {
      setSelectedCells(new Set());
      setError('');
    }
  };

  const handleMergeCells = () => {
    if (selectedCells.size < 2 || !tableRefs.current[activeSheetIndex]) {
      setError('병합할 셀을 2개 이상 선택해주세요.');
      return;
    }

    const container = tableRefs.current[activeSheetIndex];
    const table = container.querySelector('table');
    if (!table) return;

    const previousHtml = container.innerHTML;

    try {
      const rows = Array.from(table.querySelectorAll('tr'));
      const cellPositions = [];

      rows.forEach((row, rowIndex) => {
        const cells = Array.from(row.querySelectorAll('th, td'));
        let colIndex = 0;
        cells.forEach((cell) => {
          const colspan = parseInt(cell.getAttribute('colspan')) || 1;
          const cellId = cell.getAttribute('data-cell-id') || getCellId(cell);

          if (cellId && selectedCells.has(cellId)) {
            cellPositions.push({ row: rowIndex, col: colIndex, cell, cellId });
          }

          colIndex += colspan;
        });
      });

      if (cellPositions.length < 2) {
        setError('병합할 셀을 찾을 수 없습니다.');
        return;
      }

      const rowIndices = cellPositions.map((p) => p.row);
      const colIndices = cellPositions.map((p) => p.col);
      const minRow = Math.min(...rowIndices);
      const maxRow = Math.max(...rowIndices);
      const minCol = Math.min(...colIndices);
      const maxCol = Math.max(...colIndices);

      const expectedCells = (maxRow - minRow + 1) * (maxCol - minCol + 1);
      if (cellPositions.length !== expectedCells) {
        setError('병합하려면 직사각형 범위의 모든 셀을 선택해야 합니다.');
        return;
      }

      const firstCell = cellPositions.find((p) => p.row === minRow && p.col === minCol);
      if (!firstCell) {
        setError('병합할 첫 번째 셀을 찾을 수 없습니다.');
        return;
      }

      const colspan = maxCol - minCol + 1;
      const rowspan = maxRow - minRow + 1;

      let mergedContent = firstCell.cell.innerHTML.trim();
      const otherCellsContent = cellPositions
        .filter((p) => p !== firstCell && p.cell.innerHTML.trim())
        .map((p) => p.cell.innerHTML.trim())
        .filter((content) => content.length > 0);

      if (otherCellsContent.length > 0) {
        mergedContent = [mergedContent, ...otherCellsContent]
          .filter(Boolean)
          .join(' ');
      }

      firstCell.cell.setAttribute('colspan', colspan);
      if (rowspan > 1) {
        firstCell.cell.setAttribute('rowspan', rowspan);
      } else {
        firstCell.cell.removeAttribute('rowspan');
      }
      firstCell.cell.innerHTML = mergedContent || '';

      cellPositions
        .filter((p) => p !== firstCell)
        .forEach(({ cell }) => {
          cell.remove();
        });

      refreshCellIds(container);

      if (!validateTableStructure(container)) {
        restoreTableHtml(container, previousHtml);
        setError('셀 병합 후 표 구조가 깨져 작업을 되돌렸습니다.');
        return;
      }

      const containerHtml = container.innerHTML;
      setSheets((prev) =>
        prev.map((sheet, idx) =>
          idx === activeSheetIndex ? { ...sheet, content: containerHtml } : sheet
        )
      );
      setUndoStack((prev) => [...prev, previousHtml]);
      setRedoStack([]);
      setSelectedCells(new Set());
      editSessionRef.current = { active: false, snapshot: '', sheetIndex: null };
      setError('');
    } catch (err) {
      restoreTableHtml(container, previousHtml);
      setError('셀 병합 중 오류가 발생했습니다: ' + err.message);
      console.error('셀 병합 오류:', err);
    }
  };

  const handleDeleteCells = useCallback(() => {
    if (selectedCells.size === 0) {
      setError('삭제할 셀을 선택해주세요.');
      return;
    }

    const container = tableRefs.current[activeSheetIndex];
    const currentHtml = container
      ? container.innerHTML
      : sheets[activeSheetIndex]?.content || '';

    if (!currentHtml) {
      setError('삭제할 테이블을 찾을 수 없습니다.');
      return;
    }

    try {
      const parser = new DOMParser();
      const doc = parser.parseFromString(currentHtml, 'text/html');
      const table = doc.querySelector('table');

      if (!table) {
        setError('삭제할 테이블을 찾을 수 없습니다.');
        return;
      }

      let mutated = false;
      selectedCells.forEach((cellId) => {
        const cell = table.querySelector(`[data-cell-id="${cellId}"]`);
        if (cell) {
          cell.innerHTML = '';
          cell.removeAttribute('style');
          mutated = true;
        }
      });

      if (!mutated) {
        setError('삭제할 셀을 찾을 수 없습니다.');
        return;
      }

      const updatedHtml = doc.body.innerHTML || '';

      if (container) {
        container.innerHTML = updatedHtml;
        refreshCellIds(container);

        if (!validateTableStructure(container)) {
          restoreTableHtml(container, currentHtml);
          setError('셀 삭제로 인해 표 구조가 깨져 작업을 되돌렸습니다.');
          return;
        }
      }

      setSheets((prev) =>
        prev.map((sheet, idx) =>
          idx === activeSheetIndex ? { ...sheet, content: updatedHtml } : sheet
        )
      );
      setUndoStack((prev) => [...prev, currentHtml]);
      setRedoStack([]);
      setSelectedCells(new Set());
      editSessionRef.current = { active: false, snapshot: '', sheetIndex: null };
      setError('');
    } catch (err) {
      setError('셀 삭제 중 오류가 발생했습니다: ' + err.message);
      console.error('셀 삭제 오류:', err);
    }
  }, [selectedCells, activeSheetIndex, sheets]);

  useEffect(() => {
    const handleKeydown = (event) => {
      if (viewMode !== 'edit') return;
      if (selectedCells.size === 0) return;

      const target = event.target;
      if (target instanceof HTMLElement) {
        if (target.isContentEditable) return;
        const tagName = target.tagName;
        if (tagName === 'INPUT' || tagName === 'TEXTAREA' || tagName === 'SELECT') return;
      }

      if (event.key === 'Delete' || event.key === 'Backspace') {
        event.preventDefault();
        handleDeleteCells();
      }
    };

    document.addEventListener('keydown', handleKeydown);
    return () => {
      document.removeEventListener('keydown', handleKeydown);
    };
  }, [viewMode, selectedCells, handleDeleteCells]);

  const handleUndo = () => {
    if (undoStack.length === 0 || !tableRefs.current[activeSheetIndex]) {
      setError('되돌릴 작업이 없습니다.');
      return;
    }

    try {
      const container = tableRefs.current[activeSheetIndex];
      const currentHtml = container.innerHTML;
      const previousHtml = undoStack[undoStack.length - 1];

      setRedoStack((prev) => [...prev, currentHtml]);
      setUndoStack((prev) => prev.slice(0, -1));

      restoreTableHtml(container, previousHtml);

      setSheets((prev) =>
        prev.map((sheet, idx) =>
          idx === activeSheetIndex ? { ...sheet, content: previousHtml } : sheet
        )
      );

      setSelectedCells(new Set());
      editSessionRef.current = { active: false, snapshot: '', sheetIndex: null };
      setError('');
    } catch (err) {
      setError('되돌리기 중 오류가 발생했습니다: ' + err.message);
      console.error('Undo 오류:', err);
    }
  };

  const handleRedo = () => {
    if (redoStack.length === 0 || !tableRefs.current[activeSheetIndex]) {
      setError('다시 실행할 작업이 없습니다.');
      return;
    }

    try {
      const container = tableRefs.current[activeSheetIndex];
      const currentHtml = container.innerHTML;
      const nextHtml = redoStack[redoStack.length - 1];

      setUndoStack((prev) => [...prev, currentHtml]);
      setRedoStack((prev) => prev.slice(0, -1));

      restoreTableHtml(container, nextHtml);

      setSheets((prev) =>
        prev.map((sheet, idx) =>
          idx === activeSheetIndex ? { ...sheet, content: nextHtml } : sheet
        )
      );

      setSelectedCells(new Set());
      editSessionRef.current = { active: false, snapshot: '', sheetIndex: null };
      setError('');
    } catch (err) {
      setError('다시 실행 중 오류가 발생했습니다: ' + err.message);
      console.error('Redo 오류:', err);
    }
  };

  useEffect(() => {
    if (viewMode !== 'edit' || !tableRefs.current[activeSheetIndex]) return;

    const container = tableRefs.current[activeSheetIndex];
    const table = container.querySelector('table');
    if (!table) return;

    const rows = Array.from(table.querySelectorAll('tr'));

    rows.forEach((row) => {
      const cells = Array.from(row.querySelectorAll('th, td'));
      cells.forEach((cell) => {
        let cellId = cell.getAttribute('data-cell-id');
        if (!cellId) {
          cellId = getCellId(cell);
        }

        if (cellId && selectedCells.has(cellId)) {
          cell.style.backgroundColor = '#bfdbfe';
          cell.style.border = '2px solid #3b82f6';
        } else {
          if (!cell.hasAttribute('data-original-style')) {
            cell.style.backgroundColor = '';
            cell.style.border = '';
          }
        }
      });
    });
  }, [selectedCells, activeSheetIndex, viewMode]);

  useEffect(() => {
    setSelectedCells(new Set());
  }, [activeSheetIndex]);

  const getPureTableHtml = (tableHtml) => {
    const parser = new DOMParser();
    const doc = parser.parseFromString(tableHtml, 'text/html');
    const table = doc.querySelector('table');

    if (!table) return tableHtml;

    const removeAttributes = (element) => {
      element.removeAttribute('style');
      element.removeAttribute('class');
      element.removeAttribute('id');
      element.removeAttribute('contenteditable');
      element.removeAttribute('data-cell-id');

      Array.from(element.children).forEach((child) => {
        removeAttributes(child);
      });
    };

    removeAttributes(table);

    return table.outerHTML;
  };

  const applyInlineStyles = (html) => {
    try {
      const parser = new DOMParser();
      const doc = parser.parseFromString(html, 'text/html');
      const table = doc.querySelector('table');

      if (!table) return html;

      const cells = Array.from(table.querySelectorAll('th, td'));
      cells.forEach((cell) => {
        cell.removeAttribute('style');
      });
      table.removeAttribute('style');

      return table.outerHTML;
    } catch (err) {
      console.error('스타일 적용 오류:', err);
      return html;
    }
  };

  const formatTableByRows = (html) => {
    try {
      const parser = new DOMParser();
      const doc = parser.parseFromString(html, 'text/html');
      const table = doc.querySelector('table');

      if (!table) return html;

      const rows = Array.from(table.querySelectorAll('tr'));
      if (rows.length === 0) return html;

      const formattedRows = rows.map((row) => {
        const rowClone = row.cloneNode(true);
        Array.from(rowClone.querySelectorAll('*')).forEach((el) => {
          el.removeAttribute('contenteditable');
        });

        let rowHtml = rowClone.outerHTML;
        rowHtml = rowHtml.replace(/>\s+</g, '><');
        rowHtml = rowHtml.replace(/></g, '>\n  <');

        return rowHtml;
      });

      return `<table>\n${formattedRows.join('\n\n')}\n</table>`;
    } catch (err) {
      console.error('행 단위 포맷팅 오류:', err);
      return html;
    }
  };

  const handleSaveHtml = () => {
    if (sheets.length === 0) {
      setError('저장할 테이블이 없습니다.');
      return;
    }

    try {
      const fileNameInput = prompt('저장할 HTML 파일 이름을 입력하세요:', fileName || 'table');
      if (!fileNameInput) return;

      const cleanFileName = fileNameInput.replace(/[^a-zA-Z0-9가-힣_-]/g, '_');

      const currentSheet = sheets[activeSheetIndex];
      let tableHtml = tableRefs.current[activeSheetIndex]
        ? tableRefs.current[activeSheetIndex].innerHTML
        : currentSheet.content;

      tableHtml = getPureTableHtml(tableHtml);
      tableHtml = removeEmptyRowsAndCols(tableHtml);
      tableHtml = applyInlineStyles(tableHtml);
      tableHtml = formatTableByRows(tableHtml);

      const htmlContent = `<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>${currentSheet.name}</title>
  <style>
    table { border-collapse: collapse; }
    th, td { border: 1px solid #000000; }
  </style>
</head>
<body>
${tableHtml}
</body>
</html>`;

      const blob = new Blob([htmlContent], { type: 'text/html;charset=utf-8' });
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `${cleanFileName}.html`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      window.URL.revokeObjectURL(url);

      setViewMode('html');
      setError('');
    } catch (err) {
      setError('파일 저장 중 오류가 발생했습니다: ' + err.message);
    }
  };

  const handleSaveAllHtml = () => {
    if (sheets.length === 0) {
      setError('저장할 테이블이 없습니다.');
      return;
    }

    try {
      const fileNameInput = prompt('저장할 HTML 파일 이름을 입력하세요:', fileName || 'tables');
      if (!fileNameInput) return;

      const cleanFileName = fileNameInput.replace(/[^a-zA-Z0-9가-힣_-]/g, '_');

      const allTables = sheets
        .map((sheet, index) => {
          let tableHtml = tableRefs.current[index]
            ? tableRefs.current[index].innerHTML
            : sheet.content;
          tableHtml = getPureTableHtml(tableHtml);
          tableHtml = removeEmptyRowsAndCols(tableHtml);
          tableHtml = applyInlineStyles(tableHtml);
          tableHtml = formatTableByRows(tableHtml);
          return `<h2 style="margin-top:30px;margin-bottom:10px">${sheet.name}</h2>\n${tableHtml}`;
        })
        .join('\n\n');

      const htmlContent = `<!DOCTYPE html>
<html lang="ko">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Tables</title>
  <style>
    table { border-collapse: collapse; }
    th, td { border: 1px solid #000000; }
    h2 { margin-top: 30px; margin-bottom: 10px; }
  </style>
</head>
<body>
${allTables}
</body>
</html>`;

      const blob = new Blob([htmlContent], { type: 'text/html;charset=utf-8' });
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `${cleanFileName}.html`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      window.URL.revokeObjectURL(url);

      setViewMode('html');
      setError('');
    } catch (err) {
      setError('파일 저장 중 오류가 발생했습니다: ' + err.message);
    }
  };

  const handleReset = () => {
    setSheets([]);
    setActiveSheetIndex(0);
    setViewMode('edit');
    setError('');
    setFileName('');
    tableRefs.current = [];
    editSessionRef.current = { active: false, snapshot: '', sheetIndex: null };
    const fileInput = document.querySelector('input[type="file"]');
    if (fileInput) fileInput.value = '';
  };

  const handleExportJSONL = async () => {
    if (sheets.length === 0) {
      setError('내보낼 테이블이 없습니다.');
      return;
    }

    try {
      setError('이미지 변환 중...');
      const container = tableRefs.current[activeSheetIndex];
      if (!container) {
        setError('현재 시트를 찾을 수 없습니다.');
        return;
      }

      let tableHtml = container.innerHTML;
      tableHtml = getPureTableHtml(tableHtml);
      tableHtml = removeEmptyRowsAndCols(tableHtml);

      let jsonEntry;

      try {
        const canvas = await html2canvas(container, {
          backgroundColor: '#ffffff',
          scale: 2,
          logging: false,
          useCORS: true,
        });

        const imageBase64 = canvas.toDataURL('image/png');

        jsonEntry = {
          image: imageBase64,
          html: tableHtml,
        };
      } catch (imgError) {
        console.error(`시트 ${activeSheetIndex} 이미지 변환 오류:`, imgError);
        jsonEntry = {
          image: '',
          html: tableHtml,
        };
      }

      const jsonlContent = JSON.stringify(jsonEntry);

      const fileNameInput = prompt(
        '저장할 JSONL 파일 이름을 입력하세요:',
        sheets[activeSheetIndex].name || 'data'
      );
      if (!fileNameInput) {
        setError('');
        return;
      }

      const cleanFileName = fileNameInput.replace(/[^a-zA-Z0-9가-힣_-]/g, '_');
      const blob = new Blob([jsonlContent], { type: 'application/jsonl' });
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `${cleanFileName}.jsonl`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      window.URL.revokeObjectURL(url);

      setError('');
      alert('✅ 현재 시트가 JSONL 형식으로 저장되었습니다.');
    } catch (err) {
      setError('JSONL 내보내기 중 오류가 발생했습니다: ' + err.message);
      console.error('JSONL 내보내기 오류:', err);
    }
  };

  const handleExportAllJSONL = async () => {
    if (sheets.length === 0) {
      setError('내보낼 테이블이 없습니다.');
      return;
    }

    try {
      setError('이미지 변환 중...');
      const jsonlData = [];

      for (let index = 0; index < sheets.length; index++) {
        const sheet = sheets[index];

        let tableHtml = tableRefs.current[index]
          ? tableRefs.current[index].innerHTML
          : sheet.content;
        tableHtml = getPureTableHtml(tableHtml);
        tableHtml = removeEmptyRowsAndCols(tableHtml);

        let tempElement = null;
        let shouldRemoveTemp = false;

        if (!tableRefs.current[index]) {
          tempElement = document.createElement('div');
          tempElement.style.position = 'absolute';
          tempElement.style.left = '-9999px';
          tempElement.style.top = '-9999px';
          tempElement.innerHTML = tableHtml;
          document.body.appendChild(tempElement);
          shouldRemoveTemp = true;
        } else {
          tempElement = tableRefs.current[index];
        }

        try {
          const canvas = await html2canvas(tempElement, {
            backgroundColor: '#ffffff',
            scale: 2,
            logging: false,
            useCORS: true,
          });

          const imageBase64 = canvas.toDataURL('image/png');

          const jsonEntry = {
            image: imageBase64,
            html: tableHtml,
          };

          jsonlData.push(jsonEntry);

          if (shouldRemoveTemp && tempElement && tempElement.parentNode) {
            document.body.removeChild(tempElement);
          }
        } catch (imgError) {
          console.error(`시트 ${index} (${sheet.name}) 이미지 변환 오류:`, imgError);
          const jsonEntry = {
            image: '',
            html: tableHtml,
          };
          jsonlData.push(jsonEntry);

          if (shouldRemoveTemp && tempElement && tempElement.parentNode) {
            document.body.removeChild(tempElement);
          }
        }
      }

      const jsonlContent = jsonlData.map((entry) => JSON.stringify(entry)).join('\n');

      const fileNameInput = prompt('저장할 JSONL 파일 이름을 입력하세요:', fileName || 'data');
      if (!fileNameInput) {
        setError('');
        return;
      }

      const cleanFileName = fileNameInput.replace(/[^a-zA-Z0-9가-힣_-]/g, '_');
      const blob = new Blob([jsonlContent], { type: 'application/jsonl' });
      const url = window.URL.createObjectURL(blob);
      const link = document.createElement('a');
      link.href = url;
      link.download = `${cleanFileName}.jsonl`;
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      window.URL.revokeObjectURL(url);

      setError('');
      alert(`✅ ${jsonlData.length}개의 테이블이 JSONL 형식으로 저장되었습니다.`);
    } catch (err) {
      setError('JSONL 내보내기 중 오류가 발생했습니다: ' + err.message);
      console.error('JSONL 내보내기 오류:', err);
    }
  };

  return (
    <div style={styles.container}>
      <div style={styles.card}>
        <h1 style={styles.title}>엑셀 → HTML 테이블 변환기</h1>

        {error && <div style={styles.error}>{error}</div>}

        <div style={styles.section}>
          <label style={styles.label}>파일 업로드 (엑셀: .xlsx, .xls, .xlsm 또는 CSV: .csv)</label>
          <input
            type="file"
            accept=".xlsx,.xls,.xlsm,.csv"
            onChange={handleFileUpload}
            style={styles.fileInput}
          />
        </div>

        {sheets.length > 0 && (
          <div style={styles.info}>📊 총 {sheets.length}개의 시트를 발견했습니다.</div>
        )}

        <div style={styles.buttonGroup}>
          <button onClick={handleReset} style={{ ...styles.button, ...styles.secondaryButton }}>
            초기화
          </button>
        </div>
      </div>

      {sheets.length > 0 && (
        <>
          <div style={styles.card}>
            <div style={styles.header}>
              <h2 style={styles.subtitle}>
                {viewMode === 'edit' ? '편집 가능한 테이블' : 'HTML 테이블 보기'}
              </h2>
              <div
                style={{
                  display: 'flex',
                  flexDirection: 'column',
                  gap: '10px',
                  alignItems: 'flex-end',
                }}
              >
                {viewMode === 'edit' && (
                  <>
                    <div style={styles.buttonGroup}>
                      <button
                        onClick={handleSaveHtml}
                        style={{ ...styles.button, ...styles.successButton }}
                      >
                        현재 html
                      </button>
                      <button
                        onClick={handleSaveAllHtml}
                        style={{ ...styles.button, ...styles.primaryButton }}
                      >
                        전체 html
                      </button>
                      <button
                        onClick={handleExportJSONL}
                        style={{
                          ...styles.button,
                          backgroundColor: '#af52de',
                          color: 'white',
                          boxShadow: '0 2px 8px rgba(175, 82, 222, 0.25)',
                        }}
                      >
                        현재 jsonl
                      </button>
                      <button
                        onClick={handleExportAllJSONL}
                        style={{
                          ...styles.button,
                          backgroundColor: '#af52de',
                          color: 'white',
                          boxShadow: '0 2px 8px rgba(175, 82, 222, 0.25)',
                        }}
                      >
                        전체 jsonl
                      </button>
                    </div>
                    <div style={styles.buttonGroup}>
                      <button
                        onClick={handleUndo}
                        disabled={undoStack.length === 0}
                        style={{
                          ...styles.button,
                          ...styles.secondaryButton,
                          padding: '11px',
                          minWidth: '44px',
                          opacity: undoStack.length === 0 ? 0.5 : 1,
                          cursor: undoStack.length === 0 ? 'not-allowed' : 'pointer',
                        }}
                        title="되돌리기"
                      >
                        ↩️
                      </button>
                      <button
                        onClick={handleRedo}
                        disabled={redoStack.length === 0}
                        style={{
                          ...styles.button,
                          ...styles.secondaryButton,
                          padding: '11px',
                          minWidth: '44px',
                          opacity: redoStack.length === 0 ? 0.5 : 1,
                          cursor: redoStack.length === 0 ? 'not-allowed' : 'pointer',
                        }}
                        title="다시 실행"
                      >
                        ↪️
                      </button>
                      <button
                        onClick={handleToggleMergeMode}
                        style={{
                          ...styles.button,
                          backgroundColor: isMergeMode ? '#34c759' : '#8e8e93',
                          color: 'white',
                          boxShadow: isMergeMode
                            ? '0 2px 8px rgba(52, 199, 89, 0.3)'
                            : '0 2px 8px rgba(0,0,0,0.1)',
                        }}
                      >
                        {isMergeMode ? '✅ 셀병합' : '🔗 셀병합'}
                      </button>
                      {isMergeMode && (
                        <button
                          onClick={handleMergeCells}
                          disabled={selectedCells.size < 2}
                          style={{
                            ...styles.button,
                            ...styles.successButton,
                            opacity: selectedCells.size < 2 ? 0.5 : 1,
                            cursor:
                              selectedCells.size < 2 ? 'not-allowed' : 'pointer',
                          }}
                        >
                          🔗 셀 병합
                        </button>
                      )}
                      <button
                        onClick={handleDeleteCells}
                        disabled={selectedCells.size === 0}
                        style={{
                          ...styles.button,
                          backgroundColor: '#ff3b30',
                          color: 'white',
                          boxShadow: '0 2px 8px rgba(255, 59, 48, 0.25)',
                          padding: '11px',
                          minWidth: '44px',
                          opacity: selectedCells.size === 0 ? 0.5 : 1,
                          cursor:
                            selectedCells.size === 0 ? 'not-allowed' : 'pointer',
                        }}
                        title="선택 셀 삭제"
                      >
                        🗑️
                      </button>
                    </div>
                  </>
                )}
                {viewMode === 'html' && (
                  <div style={styles.buttonGroup}>
                    <button
                      onClick={handleSaveHtml}
                      style={{ ...styles.button, ...styles.successButton }}
                    >
                      현재 html
                    </button>
                    <button
                      onClick={handleSaveAllHtml}
                      style={{ ...styles.button, ...styles.primaryButton }}
                    >
                      전체 html
                    </button>
                  </div>
                )}
              </div>
            </div>

            <div style={styles.tabContainer}>
              <button
                onClick={() => setViewMode('edit')}
                style={{
                  ...styles.tab,
                  ...(viewMode === 'edit' ? styles.activeTab : styles.inactiveTab),
                }}
              >
                📝 편집
              </button>
              <button
                onClick={() => setViewMode('html')}
                style={{
                  ...styles.tab,
                  ...(viewMode === 'html' ? styles.activeTab : styles.inactiveTab),
                }}
              >
                📄 HTML 보기
              </button>
            </div>

            <div style={styles.tabContainer}>
              {sheets.map((sheet, index) => (
                <button
                  key={index}
                  onClick={() => setActiveSheetIndex(index)}
                  style={{
                    ...styles.tab,
                    ...(activeSheetIndex === index
                      ? styles.activeTab
                      : styles.inactiveTab),
                  }}
                >
                  {sheet.name}
                </button>
              ))}
            </div>

            <div style={styles.tableWrapper}>
              {viewMode === 'edit' ? (
                <div
                  ref={(el) => {
                    tableRefs.current[activeSheetIndex] = el;
                  }}
                  onClick={handleCellSelect}
                  onMouseDown={handleCellMouseDown}
                  onMouseOver={handleCellMouseOver}
                  onMouseUp={handleCellMouseUp}
                  onMouseLeave={handleCellMouseUp}
                  onDoubleClick={handleCellDoubleClick}
                  onBlur={handleCellBlur}
                  dangerouslySetInnerHTML={{ __html: sheets[activeSheetIndex].content }}
                  className="table-container"
                />
              ) : (
                <pre style={styles.codeView}>
                  <code>
                    {(() => {
                      const currentSheet = sheets[activeSheetIndex];
                      let tableHtml = tableRefs.current[activeSheetIndex]
                        ? tableRefs.current[activeSheetIndex].innerHTML
                        : currentSheet.content;

                      tableHtml = getPureTableHtml(tableHtml);
                      tableHtml = removeEmptyRowsAndCols(tableHtml);
                      tableHtml = applyInlineStyles(tableHtml);
                      tableHtml = formatTableByRows(tableHtml);

                      return tableHtml;
                    })()}
                  </code>
                </pre>
              )}
            </div>
          </div>
        </>
      )}

      <style>{`
        .table-container table {
          border-collapse: separate;
          border-spacing: 0;
          width: 100%;
          background: white;
        }
        .table-container th,
        .table-container td {
          border: 1px solid #e5e5ea;
          padding: 12px 16px;
          text-align: left;
          font-size: 15px;
          transition: all 0.2s ease;
        }
        .table-container th {
          background-color: #f9f9fb;
          font-weight: 600;
          color: #1c1c1e;
          cursor: pointer;
          letter-spacing: -0.2px;
        }
        .table-container th:first-child {
          border-top-left-radius: 12px;
        }
        .table-container th:last-child {
          border-top-right-radius: 12px;
        }
        .table-container th:hover {
          background-color: #f5f5f7;
        }
        .table-container th[contenteditable="true"] {
          background-color: #e3f2fd;
          outline: none;
          border: 2px solid #007aff;
          box-shadow: 0 0 0 3px rgba(0, 122, 255, 0.15);
        }
        .table-container td {
          cursor: pointer;
          color: #3c3c43;
        }
        .table-container td:hover {
          background-color: #f5f5f7;
        }
        .table-container td[contenteditable="true"] {
          background-color: #e3f2fd;
          outline: none;
          border: 2px solid #007aff;
          box-shadow: 0 0 0 3px rgba(0, 122, 255, 0.15);
        }
        .table-container tr:last-child td:first-child {
          border-bottom-left-radius: 12px;
        }
        .table-container tr:last-child td:last-child {
          border-bottom-right-radius: 12px;
        }
        .table-container .cell-selected {
          background-color: #e3f2fd !important;
          border: 2px solid #007aff !important;
          box-shadow: 0 0 0 2px rgba(0, 122, 255, 0.1);
        }
        pre code {
          font-family: 'SF Mono', Menlo, Monaco, 'Courier New', monospace;
          font-size: 13px;
          line-height: 1.6;
          color: #1c1c1e;
        }
        button:hover:not(:disabled) {
          transform: translateY(-1px);
          box-shadow: 0 4px 12px rgba(0,0,0,0.15);
        }
        button:active:not(:disabled) {
          transform: translateY(0);
          box-shadow: 0 1px 4px rgba(0,0,0,0.1);
        }
        button:disabled {
          opacity: 0.5;
          cursor: not-allowed;
        }
        .viewModeGroup button:hover:not(:disabled) {
          transform: none;
        }
        .tab:hover:not(:disabled) {
          transform: translateY(-1px);
        }
      `}</style>
    </div>
  );
}

const styles = {
  container: {
    minHeight: '100vh',
    backgroundColor: '#f2f2f7',
    padding: '20px',
    fontFamily:
      '-apple-system, BlinkMacSystemFont, "SF Pro Display", "SF Pro Text", "Helvetica Neue", sans-serif',
  },
  card: {
    maxWidth: '1400px',
    margin: '0 auto 20px',
    backgroundColor: 'white',
    borderRadius: '16px',
    boxShadow: '0 4px 16px rgba(0,0,0,0.08)',
    padding: '28px',
    border: '0.5px solid rgba(0,0,0,0.04)',
  },
  title: {
    fontSize: '32px',
    fontWeight: '700',
    color: '#1c1c1e',
    marginBottom: '8px',
    marginTop: 0,
    letterSpacing: '-0.5px',
  },
  subtitle: {
    fontSize: '22px',
    fontWeight: '600',
    color: '#1c1c1e',
    margin: 0,
    letterSpacing: '-0.3px',
  },
  error: {
    backgroundColor: '#ffebee',
    border: 'none',
    color: '#c62828',
    padding: '14px 18px',
    borderRadius: '12px',
    marginBottom: '16px',
    fontSize: '15px',
    boxShadow: '0 2px 8px rgba(198, 40, 40, 0.12)',
  },
  info: {
    backgroundColor: '#e3f2fd',
    border: 'none',
    color: '#1565c0',
    padding: '14px 18px',
    borderRadius: '12px',
    marginBottom: '16px',
    fontSize: '15px',
    boxShadow: '0 2px 8px rgba(21, 101, 192, 0.12)',
  },
  section: {
    marginBottom: '20px',
  },
  label: {
    display: 'block',
    fontSize: '15px',
    fontWeight: '600',
    color: '#3c3c43',
    marginBottom: '10px',
    letterSpacing: '-0.2px',
  },
  fileInput: {
    display: 'block',
    width: '100%',
    fontSize: '15px',
    padding: '12px 16px',
    border: '1px solid #d1d1d6',
    borderRadius: '12px',
    backgroundColor: '#fafafa',
    transition: 'all 0.2s ease',
  },
  buttonGroup: {
    display: 'flex',
    gap: '10px',
    flexWrap: 'wrap',
  },
  button: {
    padding: '11px 20px',
    borderRadius: '12px',
    border: 'none',
    fontSize: '15px',
    fontWeight: '600',
    cursor: 'pointer',
    transition: 'all 0.25s ease',
    boxShadow: '0 2px 8px rgba(0,0,0,0.1)',
    letterSpacing: '-0.2px',
  },
  primaryButton: {
    backgroundColor: '#007aff',
    color: 'white',
  },
  secondaryButton: {
    backgroundColor: '#f2f2f7',
    color: '#3c3c43',
  },
  successButton: {
    backgroundColor: '#34c759',
    color: 'white',
  },
  header: {
    display: 'flex',
    justifyContent: 'space-between',
    alignItems: 'center',
    marginBottom: '20px',
    flexWrap: 'wrap',
    gap: '16px',
  },
  tabContainer: {
    display: 'flex',
    gap: '6px',
    marginBottom: '20px',
    overflowX: 'auto',
    paddingBottom: '4px',
    borderBottom: 'none',
  },
  tab: {
    padding: '10px 20px',
    border: 'none',
    borderRadius: '10px',
    fontSize: '15px',
    fontWeight: '600',
    cursor: 'pointer',
    transition: 'all 0.25s ease',
    whiteSpace: 'nowrap',
    letterSpacing: '-0.2px',
  },
  activeTab: {
    backgroundColor: '#007aff',
    color: 'white',
    boxShadow: '0 2px 8px rgba(0, 122, 255, 0.3)',
  },
  inactiveTab: {
    backgroundColor: '#f2f2f7',
    color: '#8e8e93',
  },
  tableWrapper: {
    overflowX: 'auto',
    border: 'none',
    borderRadius: '12px',
    backgroundColor: 'white',
    boxShadow: '0 2px 12px rgba(0,0,0,0.08)',
  },
  codeView: {
    margin: 0,
    padding: '20px',
    backgroundColor: '#fafafa',
    color: '#1c1c1e',
    borderRadius: '12px',
    overflow: 'auto',
    fontFamily: 'SF Mono, Menlo, Monaco, Courier New, monospace',
    fontSize: '13px',
    lineHeight: '1.6',
    whiteSpace: 'pre',
    maxHeight: '600px',
    border: '1px solid #e5e5ea',
    boxShadow: 'inset 0 1px 3px rgba(0,0,0,0.05)',
  },
};

export default App;
