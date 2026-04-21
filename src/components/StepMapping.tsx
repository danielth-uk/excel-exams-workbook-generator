import { useState, useMemo } from "react";
import { Button } from "./ui/button";
import { CardHeader, CardTitle, CardDescription, CardContent } from "./ui/card";
import { Table, TableBody, TableRow, TableCell } from "./ui/table";
import type { ExcelData, TemplateData, CellMapping } from "./ExcelWizard";
import { Trash2, Download, Loader2 } from "lucide-react";
import * as ExcelJS from "exceljs";

interface StepMappingProps {
  datesFile: ExcelData;
  templateFile: TemplateData;
  mappings: CellMapping[];
  selectedRows: Set<number>;
  sheetNameTemplate: string;
  onMappingsChange: (mappings: CellMapping[]) => void;
  onSheetNameTemplateChange: (template: string) => void;
  onBack: () => void;
  onReset: () => void;
}

export function StepMapping({
  datesFile,
  templateFile,
  mappings,
  selectedRows,
  sheetNameTemplate,
  onMappingsChange,
  onSheetNameTemplateChange,
  onBack,
  onReset,
}: StepMappingProps) {
  const [selectedCell, setSelectedCell] = useState<string | null>(null);
  const [selectedColumn, setSelectedColumn] = useState<number | null>(null);
  const [isGenerating, setIsGenerating] = useState(false);
  const [activePreviewTab, setActivePreviewTab] = useState(0);

  const isDateTimeColumn = (columnIndex: number): boolean => {
    // Check if column contains date/time values
    const sampleValues = datesFile.rows
      .slice(0, 10)
      .map((row) => row[columnIndex]);
    return sampleValues.some((val) => {
      if (!val) return false;
      // Check if it's a Date object or looks like date/time string
      if (val instanceof Date) return true;
      const str = String(val);
      // Simple check for time patterns like "09:00", "AM", "PM", dates
      return /\d{1,2}:\d{2}|AM|PM|\/\d{1,2}\/|^\d{4}-\d{2}-\d{2}/.test(str);
    });
  };

  const getCellAddress = (rowIdx: number, colIdx: number): string => {
    let colLetter = "";
    let col = colIdx + 1;
    while (col > 0) {
      const mod = (col - 1) % 26;
      colLetter = String.fromCharCode(65 + mod) + colLetter;
      col = Math.floor((col - mod) / 26);
    }
    return `${colLetter}${rowIdx + 1}`;
  };

  const getColumnLabel = (colIdx: number): string => {
    let label = "";
    let col = colIdx + 1;
    while (col > 0) {
      const mod = (col - 1) % 26;
      label = String.fromCharCode(65 + mod) + label;
      col = Math.floor((col - mod) / 26);
    }
    return label;
  };

  const handleCellClick = (rowIdx: number, colIdx: number) => {
    const cellAddress = getCellAddress(rowIdx, colIdx);
    setSelectedCell(cellAddress);
    setSelectedColumn(null);
  };

  const handleColumnSelect = (columnIndex: number) => {
    if (!selectedCell) return;

    // If it's a date/time column, show format options
    if (isDateTimeColumn(columnIndex)) {
      setSelectedColumn(columnIndex);
      return;
    }

    // Otherwise, add mapping directly
    const newMappings = mappings.filter((m) => m.cellAddress !== selectedCell);
    newMappings.push({
      cellAddress: selectedCell,
      columnIndex,
    });

    onMappingsChange(newMappings);
    setSelectedCell(null);
    setSelectedColumn(null);
  };

  const handleDateTimeFormatSelect = (
    format: "all" | "date" | "time" | "ampm" | "none",
  ) => {
    if (!selectedCell || selectedColumn === null) return;

    const newMappings = mappings.filter((m) => m.cellAddress !== selectedCell);
    newMappings.push({
      cellAddress: selectedCell,
      columnIndex: selectedColumn,
      dateTimeFormat: format,
    });

    onMappingsChange(newMappings);
    setSelectedCell(null);
    setSelectedColumn(null);
  };

  const handleRemoveMapping = (cellAddress: string) => {
    onMappingsChange(mappings.filter((m) => m.cellAddress !== cellAddress));
  };

  const getMappingForCell = (cellAddress: string) => {
    return mappings.find((m) => m.cellAddress === cellAddress);
  };

  const parseCellAddress = (address: string): { row: number; col: number } => {
    const match = address.match(/([A-Z]+)(\d+)/);
    if (!match) throw new Error(`Bad cell address: ${address}`);

    const colLetter = match[1];
    const rowNum = parseInt(match[2], 10);

    let colIndex = 0;
    for (let i = 0; i < colLetter.length; i++) {
      colIndex = colIndex * 26 + (colLetter.charCodeAt(i) - 64);
    }
    colIndex -= 1;

    return { row: rowNum - 1, col: colIndex };
  };

  const formatDateTime = (
    value: any,
    format: "all" | "date" | "time" | "ampm" | "none" | "clean",
  ): string => {
    if (!value) return "";
    if (format === "none") {
      // Strip double quotes from the value
      return String(value).replace(/^"(.*)"$/, '$1');
    }
    if (format === "clean") {
      // Strip double quotes from the value (same as none)
      return String(value).replace(/^"(.*)"$/, '$1');
    }

    let date: Date;
    let timeStr = String(value);

    if (value instanceof Date) {
      date = value;
    } else {
      const customMatch = timeStr.match(
        /(\d{2})\.(\d{2})\.(\d{2})[^\d]*(\d{1,2}):(\d{2})/,
      );
      if (customMatch) {
        const day = parseInt(customMatch[1], 10);
        const month = parseInt(customMatch[2], 10) - 1;
        const year = 2000 + parseInt(customMatch[3], 10);
        const hours = parseInt(customMatch[4], 10);
        const minutes = parseInt(customMatch[5], 10);
        const isPM = /PM/i.test(timeStr);
        const adjustedHours =
          isPM && hours !== 12 ? hours + 12 : !isPM && hours === 12 ? 0 : hours;
        date = new Date(year, month, day, adjustedHours, minutes);
      } else {
        const dateOnlyMatch = timeStr.match(
          /(\d{1,2})[\/\.](\d{1,2})[\/\.](\d{2,4})/,
        );
        if (dateOnlyMatch) {
          const day = parseInt(dateOnlyMatch[1], 10);
          const month = parseInt(dateOnlyMatch[2], 10) - 1;
          let year = parseInt(dateOnlyMatch[3], 10);
          if (year < 100) year += 2000;

          const timeMatch = timeStr.match(/(\d{1,2}):(\d{2})/);
          let hours = 0;
          let minutes = 0;

          if (timeMatch) {
            hours = parseInt(timeMatch[1], 10);
            minutes = parseInt(timeMatch[2], 10);
            const isPM = /PM/i.test(timeStr);
            const isAM = /AM/i.test(timeStr);
            if (isPM && hours !== 12) hours += 12;
            else if (isAM && hours === 12) hours = 0;
          }

          date = new Date(year, month, day, hours, minutes);
        } else {
          date = new Date(value);
          if (isNaN(date.getTime())) return String(value);
        }
      }
    }

    const days = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
    const day = days[date.getDay()];
    const dateNum = date.getDate().toString().padStart(2, "0");
    const monthNum = (date.getMonth() + 1).toString().padStart(2, "0");
    const yearShort = date.getFullYear().toString().slice(-2);
    const hours = date.getHours();
    const minutes = date.getMinutes().toString().padStart(2, "0");
    const hours12 = hours % 12 || 12;

    let ampmSmart: string;
    if (hours >= 10 && hours < 12) ampmSmart = "MID-AM";
    else if (hours === 12) ampmSmart = "Early PM";
    else if (hours >= 12) ampmSmart = "PM";
    else ampmSmart = "AM";

    switch (format) {
      case "date":
        return `${dateNum}.${monthNum}.${yearShort}`;
      case "time":
        return `${hours12}:${minutes}`;
      case "ampm":
        return ampmSmart;
      case "all":
      default:
        return `${day} ${dateNum}.${monthNum}.${yearShort} ${hours12}:${minutes} ${ampmSmart}`;
    }
  };

  const sanitizeSheetName = (name: string): string => {
    let sanitized = String(name).replace(/[\*\?\:\\\/\[\]]/g, "");
    if (sanitized.length > 31) sanitized = sanitized.substring(0, 31);
    if (!sanitized.trim()) sanitized = "Sheet";
    return sanitized.trim();
  };

  // Build sheet name from template
  const buildSheetName = (rowData: any[], rowIndex: number): string => {
    let name = sheetNameTemplate;
    
    // Replace {A:format}, {B:format}, etc. with formatted column values
    name = name.replace(/\{([A-Z]+):(date|time|all|ampm|none|clean)\}/g, (match, columnLetter, format) => {
      // Convert column letter to index
      let colIndex = 0;
      for (let i = 0; i < columnLetter.length; i++) {
        colIndex = colIndex * 26 + (columnLetter.charCodeAt(i) - 64);
      }
      colIndex -= 1; // Make 0-based
      
      const value = rowData[colIndex];
      if (value === null || value === undefined) return '';
      
      // Apply datetime formatting
      return formatDateTime(value, format as "date" | "time" | "all" | "ampm" | "none" | "clean");
    });
    
    // Replace {0:format}, {1:format}, etc. with formatted column values
    name = name.replace(/\{(\d+):(date|time|all|ampm|none|clean)\}/g, (match, index, format) => {
      const colIndex = parseInt(index, 10);
      const value = rowData[colIndex];
      if (value === null || value === undefined) return '';
      
      // Apply datetime formatting
      return formatDateTime(value, format as "date" | "time" | "all" | "ampm" | "none" | "clean");
    });
    
    // Replace {A}, {B}, {C}, etc. with column values (no formatting)
    name = name.replace(/\{([A-Z]+)\}/g, (match, columnLetter) => {
      // Convert column letter to index
      let colIndex = 0;
      for (let i = 0; i < columnLetter.length; i++) {
        colIndex = colIndex * 26 + (columnLetter.charCodeAt(i) - 64);
      }
      colIndex -= 1; // Make 0-based
      
      const value = rowData[colIndex];
      return value !== null && value !== undefined ? String(value) : '';
    });
    
    // Replace {0}, {1}, {2}, etc. with column indices (no formatting)
    name = name.replace(/\{(\d+)\}/g, (match, index) => {
      const colIndex = parseInt(index, 10);
      const value = rowData[colIndex];
      return value !== null && value !== undefined ? String(value) : '';
    });
    
    // If template results in empty string, use default
    if (!name.trim()) {
      name = `Sheet_${rowIndex + 1}`;
    }
    
    return sanitizeSheetName(name);
  };

  // Generate live preview for selected rows
  const livePreview = useMemo(() => {
    if (mappings.length === 0 || selectedRows.size === 0) return [];
    
    const selectedRowIndices = Array.from(selectedRows).sort((a, b) => a - b).slice(0, 3);
    
    return selectedRowIndices.map((rowIndex) => {
      const rowData = datesFile.rows[rowIndex];
      const sheetName = buildSheetName(rowData, rowIndex);
      const previewData: (string | number | null)[][] = [];
      const mappedCells = new Set<string>();

      templateFile.preview.forEach((templateRow, rowIdx) => {
        const previewRow: (string | number | null)[] = [];
        templateRow.forEach((cell, colIdx) => {
          const cellAddress = getCellAddress(rowIdx, colIdx);
          const mapping = mappings.find(m => m.cellAddress === cellAddress);
          
          if (mapping) {
            let dataValue = rowData[mapping.columnIndex];
            if (mapping.dateTimeFormat) {
              dataValue = formatDateTime(dataValue, mapping.dateTimeFormat);
            }
            previewRow.push(dataValue === null || dataValue === undefined ? null : String(dataValue));
            mappedCells.add(cellAddress);
          } else {
            previewRow.push(cell === null || cell === undefined ? null : String(cell));
          }
        });
        previewData.push(previewRow);
      });

      return { name: sheetName, data: previewData, mappedCells };
    });
  }, [mappings, selectedRows, datesFile.rows, templateFile.preview, sheetNameTemplate]);

  const handleDownload = async () => {
    setIsGenerating(true);

    try {
      const workbook = new ExcelJS.Workbook();
      const selectedRowIndices = Array.from(selectedRows).sort((a, b) => a - b);
      const usedSheetNames = new Set<string>();

      for (const rowIndex of selectedRowIndices) {
        const rowData = datesFile.rows[rowIndex];
        let sheetName = buildSheetName(rowData, rowIndex);

        let finalSheetName = sheetName;
        let duplicateCounter = 2;
        while (usedSheetNames.has(finalSheetName)) {
          const suffix = ` (${duplicateCounter})`;
          const maxBaseLength = 31 - suffix.length;
          const baseName = sheetName.substring(0, maxBaseLength);
          finalSheetName = `${baseName}${suffix}`;
          duplicateCounter++;
        }
        usedSheetNames.add(finalSheetName);

        const templateWorksheet = templateFile.workbook.worksheets[0];
        const newSheet = workbook.addWorksheet(finalSheetName);

        templateWorksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
          const newRow = newSheet.getRow(rowNumber);
          row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            const newCell = newRow.getCell(colNumber);
            newCell.value = cell.value;
            if (cell.style) newCell.style = cell.style;
          });
          newRow.commit();
        });

        for (const mapping of mappings) {
          const { row: rowNum, col: colNum } = parseCellAddress(mapping.cellAddress);
          const cell = newSheet.getRow(rowNum + 1).getCell(colNum + 1);
          let dataValue = rowData[mapping.columnIndex];

          if (mapping.dateTimeFormat) {
            dataValue = formatDateTime(dataValue, mapping.dateTimeFormat);
          }

          cell.value = dataValue;
        }

        templateWorksheet.columns.forEach((col, idx) => {
          if (col.width) {
            const newCol = newSheet.getColumn(idx + 1);
            newCol.width = col.width;
          }
        });
      }

      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = URL.createObjectURL(blob);

      const link = document.createElement('a');
      link.href = url;
      link.download = 'output.xlsx';
      document.body.appendChild(link);
      link.click();
      document.body.removeChild(link);
      
      setTimeout(() => URL.revokeObjectURL(url), 1000);
    } catch (error) {
      console.error("Error generating Excel:", error);
      alert("Ugh! Error making Excel! Check console!");
    } finally {
      setIsGenerating(false);
    }
  };

  return (
    <>
      <CardHeader>
        <CardTitle>Step 3: Map Columns & Preview</CardTitle>
        <CardDescription>
          Click template cell, pick column, see live preview! Download when ready!
        </CardDescription>
      </CardHeader>
      <CardContent className="space-y-6">
        {/* Sheet Name Template Input */}
        <div className="bg-slate-50 border border-slate-200 rounded-lg p-4">
          <label className="block mb-2">
            <span className="font-semibold text-slate-900 text-sm">Sheet Name Template</span>
            <p className="text-xs text-slate-600 mt-1 mb-2">
              Use <code className="bg-slate-200 px-1 rounded">{"{A}"}</code>, <code className="bg-slate-200 px-1 rounded">{"{B}"}</code> for columns. 
              Add format: <code className="bg-slate-200 px-1 rounded">{"{A:clean}"}</code>, <code className="bg-slate-200 px-1 rounded">{"{B:date}"}</code>, <code className="bg-slate-200 px-1 rounded">{"{C:time}"}</code>
            </p>
          </label>
          <div className="flex gap-2">
            <input
              type="text"
              value={sheetNameTemplate}
              onChange={(e) => onSheetNameTemplateChange(e.target.value)}
              placeholder="{A}"
              className="flex-1 px-3 py-2 text-sm border border-slate-300 rounded focus:outline-none focus:ring-2 focus:ring-slate-900"
            />
            <button
              onClick={() => onSheetNameTemplateChange("")}
              className="px-3 py-2 text-xs bg-red-100 hover:bg-red-200 text-red-700 rounded transition-colors"
            >
              Clear
            </button>
            <button
              onClick={() => onSheetNameTemplateChange("{A}")}
              className="px-3 py-2 text-xs bg-slate-200 hover:bg-slate-300 rounded transition-colors"
            >
              Reset
            </button>
          </div>
          
          {/* Quick add buttons for columns */}
          <div className="mt-2">
            <div className="text-xs font-semibold text-slate-700 mb-1">Quick Add Columns:</div>
            <div className="flex flex-wrap gap-1">
              {datesFile.headers.slice(0, 5).map((header, idx) => {
                const colLetter = String.fromCharCode(65 + idx);
                const isDateTime = isDateTimeColumn(idx);
                return (
                  <button
                    key={idx}
                    onClick={() => onSheetNameTemplateChange(sheetNameTemplate + `{${colLetter}}`)}
                    className="px-2 py-1 text-xs bg-white border border-slate-300 hover:bg-slate-100 rounded transition-colors"
                  >
                    + {colLetter} {isDateTime && '🕒'}
                  </button>
                );
              })}
            </div>
          </div>

          {/* Date/Time format quick add buttons */}
          <div className="mt-2">
            <div className="text-xs font-semibold text-slate-700 mb-1">Formats:</div>
            <div className="flex flex-wrap gap-1">
              <button
                onClick={() => onSheetNameTemplateChange(sheetNameTemplate + "{A:clean}")}
                className="px-2 py-1 text-xs bg-green-50 border border-green-200 hover:bg-green-100 rounded transition-colors"
              >
                + Clean (strip quotes)
              </button>
              <button
                onClick={() => onSheetNameTemplateChange(sheetNameTemplate + "{A:date}")}
                className="px-2 py-1 text-xs bg-blue-50 border border-blue-200 hover:bg-blue-100 rounded transition-colors"
              >
                + Date Only
              </button>
              <button
                onClick={() => onSheetNameTemplateChange(sheetNameTemplate + "{A:time}")}
                className="px-2 py-1 text-xs bg-blue-50 border border-blue-200 hover:bg-blue-100 rounded transition-colors"
              >
                + Time Only
              </button>
              <button
                onClick={() => onSheetNameTemplateChange(sheetNameTemplate + "{A:ampm}")}
                className="px-2 py-1 text-xs bg-blue-50 border border-blue-200 hover:bg-blue-100 rounded transition-colors"
              >
                + AM/PM
              </button>
              <button
                onClick={() => onSheetNameTemplateChange(sheetNameTemplate + "{A:all}")}
                className="px-2 py-1 text-xs bg-blue-50 border border-blue-200 hover:bg-blue-100 rounded transition-colors"
              >
                + Full DateTime
              </button>
            </div>
            <p className="text-xs text-slate-500 mt-1">
              💡 Replace "A" with your column letter (e.g., {"{B:clean}"}, {"{C:date}"}, {"{D:time}"})
            </p>
          </div>
        </div>

        <div className="grid grid-cols-3 gap-4">
          {/* Template preview */}
          <div>
            <h3 className="font-semibold text-slate-900 mb-2">
              Template
            </h3>
            <p className="text-sm text-slate-600 mb-4">
              Click cells to map
            </p>
            <div className="border rounded-lg overflow-auto max-h-[600px] bg-white">
              <table className="border-collapse w-full">
                <thead>
                  <tr>
                    <th className="sticky top-0 left-0 z-20 bg-slate-100 border border-slate-300 w-8 h-6 text-xs font-semibold text-slate-600"></th>
                    {templateFile.preview[0]?.map((_, colIdx) => (
                      <th
                        key={colIdx}
                        className="sticky top-0 z-10 bg-slate-100 border border-slate-300 min-w-[80px] h-6 text-xs font-semibold text-slate-600"
                      >
                        {getColumnLabel(colIdx)}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {templateFile.preview.map((row, rowIdx) => (
                    <tr key={rowIdx}>
                      <td className="sticky left-0 z-10 bg-slate-100 border border-slate-300 text-center text-xs font-semibold text-slate-600 w-8 h-6">
                        {rowIdx + 1}
                      </td>
                      {row.map((cell, colIdx) => {
                        const cellAddress = getCellAddress(rowIdx, colIdx);
                        const mapping = getMappingForCell(cellAddress);
                        const isSelected = selectedCell === cellAddress;

                        return (
                          <td
                            key={colIdx}
                            onClick={() => handleCellClick(rowIdx, colIdx)}
                            className={`border border-slate-300 px-2 py-1 text-xs cursor-pointer select-none transition-colors min-w-[80px] h-6 ${
                              isSelected
                                ? "bg-blue-100 border-blue-500 border-2"
                                : mapping
                                  ? "bg-green-100 border-green-500"
                                  : "hover:bg-slate-50"
                            }`}
                          >
                            <div className="flex items-center gap-1">
                              {mapping && (
                                <span className="text-xs text-green-700 font-medium">
                                  [{datesFile.headers[mapping.columnIndex]}]
                                </span>
                              )}
                              <span className={mapping ? "text-slate-500 text-xs" : "text-xs"}>
                                {String(cell || "")}
                              </span>
                            </div>
                          </td>
                        );
                      })}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>

          {/* Column selector */}
          <div className="max-h-[600px] overflow-y-auto">
            <h3 className="font-semibold text-slate-900 mb-2">
              {selectedCell
                ? `Select column for ${selectedCell}`
                : "Columns"}
            </h3>
            <p className="text-sm text-slate-600 mb-4">
              {selectedCell
                ? "Pick column"
                : "Click cell first"}
            </p>

            {/* Date/Time format selector */}
            {selectedColumn !== null && (
              <div className="mb-4 p-3 bg-blue-50 border border-blue-200 rounded-lg">
                <h4 className="font-semibold text-blue-900 mb-2 text-sm">
                  Date/Time Format
                </h4>
                <p className="text-xs text-blue-700 mb-2">
                  {datesFile.headers[selectedColumn]}
                </p>
                <div className="space-y-1">
                  <button
                    onClick={() => handleDateTimeFormatSelect("all")}
                    className="w-full text-left px-3 py-1.5 text-xs rounded border border-blue-300 bg-white hover:bg-blue-50 transition-colors"
                  >
                    <div className="font-medium">Full Date & Time</div>
                  </button>
                  <button
                    onClick={() => handleDateTimeFormatSelect("date")}
                    className="w-full text-left px-3 py-1.5 text-xs rounded border border-blue-300 bg-white hover:bg-blue-50 transition-colors"
                  >
                    <div className="font-medium">Date Only</div>
                  </button>
                  <button
                    onClick={() => handleDateTimeFormatSelect("time")}
                    className="w-full text-left px-3 py-1.5 text-xs rounded border border-blue-300 bg-white hover:bg-blue-50 transition-colors"
                  >
                    <div className="font-medium">Time Only</div>
                  </button>
                  <button
                    onClick={() => handleDateTimeFormatSelect("ampm")}
                    className="w-full text-left px-3 py-1.5 text-xs rounded border border-blue-300 bg-white hover:bg-blue-50 transition-colors"
                  >
                    <div className="font-medium">AM/PM (Smart)</div>
                  </button>
                  <button
                    onClick={() => handleDateTimeFormatSelect("none")}
                    className="w-full text-left px-3 py-1.5 text-xs rounded border border-blue-300 bg-white hover:bg-blue-50 transition-colors"
                  >
                    <div className="font-medium">No formatting</div>
                  </button>
                  <button
                    onClick={() => setSelectedColumn(null)}
                    className="w-full px-3 py-1.5 text-xs text-slate-600 hover:text-slate-900"
                  >
                    Cancel
                  </button>
                </div>
              </div>
            )}

            <div className="space-y-1">
              {datesFile.headers.map((header, idx) => {
                const isDateTime = isDateTimeColumn(idx);
                return (
                  <button
                    key={idx}
                    onClick={() => handleColumnSelect(idx)}
                    disabled={!selectedCell}
                    className={`w-full text-left px-3 py-2 text-xs rounded border transition-colors ${
                      selectedCell
                        ? "hover:bg-slate-100 border-slate-300 cursor-pointer"
                        : "border-slate-200 text-slate-400 cursor-not-allowed"
                    }`}
                  >
                    <div className="flex items-center gap-2">
                      <div className="font-medium">{header}</div>
                      {isDateTime && (
                        <span className="text-xs bg-blue-100 text-blue-700 px-1.5 py-0.5 rounded">
                          Date/Time
                        </span>
                      )}
                    </div>
                  </button>
                );
              })}
            </div>

            {/* Current mappings */}
            {mappings.length > 0 && (
              <div className="mt-4">
                <h4 className="font-semibold text-slate-900 mb-2 text-sm">
                  Mappings
                </h4>
                <div className="space-y-1">
                  {mappings.map((mapping) => (
                    <div
                      key={mapping.cellAddress}
                      className="flex items-center justify-between px-2 py-1.5 bg-green-50 rounded border border-green-200"
                    >
                      <span className="text-xs">
                        <span className="font-medium">
                          {mapping.cellAddress}
                        </span>
                        {" → "}
                        <span className="text-green-700">
                          {datesFile.headers[mapping.columnIndex]}
                        </span>
                        {mapping.dateTimeFormat && (
                          <span className="ml-1 text-xs bg-blue-100 text-blue-700 px-1 py-0.5 rounded">
                            {mapping.dateTimeFormat === "all" && "Full"}
                            {mapping.dateTimeFormat === "date" && "Date"}
                            {mapping.dateTimeFormat === "time" && "Time"}
                            {mapping.dateTimeFormat === "ampm" && "AM/PM"}
                            {mapping.dateTimeFormat === "none" && "Raw"}
                          </span>
                        )}
                      </span>
                      <button
                        onClick={() => handleRemoveMapping(mapping.cellAddress)}
                        className="text-red-500 hover:text-red-700"
                      >
                        <Trash2 className="h-3 w-3" />
                      </button>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>

          {/* Live Preview */}
          <div>
            <h3 className="font-semibold text-slate-900 mb-2">
              Live Preview
            </h3>
            <p className="text-sm text-slate-600 mb-4">
              {livePreview.length > 0 ? `Showing ${livePreview.length} of ${selectedRows.size} sheets` : "Map columns to see preview"}
            </p>
            
            {livePreview.length > 0 ? (
              <>
                {/* Preview tabs */}
                {livePreview.length > 1 && (
                  <div className="flex gap-1 mb-2">
                    {livePreview.map((sheet, index) => (
                      <button
                        key={index}
                        onClick={() => setActivePreviewTab(index)}
                        className={`px-2 py-1 text-xs rounded transition-colors ${
                          activePreviewTab === index
                            ? "bg-slate-900 text-white font-semibold"
                            : "bg-slate-200 text-slate-600 hover:bg-slate-300"
                        }`}
                      >
                        {sheet.name.substring(0, 10)}{sheet.name.length > 10 ? '...' : ''}
                      </button>
                    ))}
                  </div>
                )}
                
                <div className="border rounded-lg overflow-auto max-h-[600px] bg-white">
                  <table className="border-collapse w-full">
                    <thead>
                      <tr>
                        <th className="sticky top-0 left-0 z-20 bg-slate-100 border border-slate-300 w-8 h-6 text-xs font-semibold text-slate-600"></th>
                        {livePreview[activePreviewTab].data[0]?.map((_, colIndex) => (
                          <th
                            key={colIndex}
                            className="sticky top-0 z-10 bg-slate-100 border border-slate-300 min-w-[80px] h-6 text-xs font-semibold text-slate-600"
                          >
                            {String.fromCharCode(65 + (colIndex % 26))}
                          </th>
                        ))}
                      </tr>
                    </thead>
                    <tbody>
                      {livePreview[activePreviewTab].data.map((row, rowIndex) => (
                        <tr key={rowIndex}>
                          <td className="sticky left-0 z-10 bg-slate-100 border border-slate-300 text-center text-xs font-semibold text-slate-600 w-8 h-6">
                            {rowIndex + 1}
                          </td>
                          {row.map((cell, colIndex) => {
                            const cellAddress = `${String.fromCharCode(65 + colIndex)}${rowIndex + 1}`;
                            const isMapped = livePreview[activePreviewTab].mappedCells.has(cellAddress);
                            return (
                              <td
                                key={colIndex}
                                className={`border border-slate-300 px-2 py-1 text-xs min-w-[80px] h-6 ${
                                  isMapped ? "bg-blue-50 font-medium" : ""
                                }`}
                              >
                                {cell ?? ""}
                              </td>
                            );
                          })}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                </div>
              </>
            ) : (
              <div className="border rounded-lg p-8 text-center text-slate-400">
                <p className="text-sm">No mappings yet</p>
                <p className="text-xs mt-1">Click a template cell and select a column</p>
              </div>
            )}
          </div>
        </div>

        {/* Download button and navigation */}
        <div className="flex justify-between items-center pt-4 border-t">
          <Button variant="outline" onClick={onBack}>
            Back
          </Button>
          
          <div className="flex gap-3">
            <Button 
              onClick={handleDownload} 
              disabled={mappings.length === 0 || isGenerating}
              size="lg"
            >
              {isGenerating ? (
                <>
                  <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                  Generating...
                </>
              ) : (
                <>
                  <Download className="mr-2 h-4 w-4" />
                  Download Excel ({selectedRows.size} sheets)
                </>
              )}
            </Button>
            
            <Button variant="outline" onClick={onReset}>
              Start Over
            </Button>
          </div>
        </div>
      </CardContent>
    </>
  );
}
