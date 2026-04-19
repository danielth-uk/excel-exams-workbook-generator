import { useState } from 'react';
import { Button } from './ui/button';
import { CardHeader, CardTitle, CardDescription, CardContent } from './ui/card';
import { Table, TableBody, TableRow, TableCell } from './ui/table';
import type { ExcelData, TemplateData, CellMapping } from './ExcelWizard';
import { Trash2 } from 'lucide-react';

interface StepMappingProps {
  datesFile: ExcelData;
  templateFile: TemplateData;
  mappings: CellMapping[];
  onMappingsChange: (mappings: CellMapping[]) => void;
  onBack: () => void;
  onNext: () => void;
}

export function StepMapping({
  datesFile,
  templateFile,
  mappings,
  onMappingsChange,
  onBack,
  onNext,
}: StepMappingProps) {
  const [selectedCell, setSelectedCell] = useState<string | null>(null);
  const [selectedColumn, setSelectedColumn] = useState<number | null>(null);

  const isDateTimeColumn = (columnIndex: number): boolean => {
    // Check if column contains date/time values
    const sampleValues = datesFile.rows.slice(0, 10).map(row => row[columnIndex]);
    return sampleValues.some(val => {
      if (!val) return false;
      // Check if it's a Date object or looks like date/time string
      if (val instanceof Date) return true;
      const str = String(val);
      // Simple check for time patterns like "09:00", "AM", "PM", dates
      return /\d{1,2}:\d{2}|AM|PM|\/\d{1,2}\/|^\d{4}-\d{2}-\d{2}/.test(str);
    });
  };

  const getCellAddress = (rowIdx: number, colIdx: number): string => {
    let colLetter = '';
    let col = colIdx + 1;
    while (col > 0) {
      const mod = (col - 1) % 26;
      colLetter = String.fromCharCode(65 + mod) + colLetter;
      col = Math.floor((col - mod) / 26);
    }
    return `${colLetter}${rowIdx + 1}`;
  };

  const getColumnLabel = (colIdx: number): string => {
    let label = '';
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
    const newMappings = mappings.filter(m => m.cellAddress !== selectedCell);
    newMappings.push({
      cellAddress: selectedCell,
      columnIndex,
    });

    onMappingsChange(newMappings);
    setSelectedCell(null);
    setSelectedColumn(null);
  };

  const handleDateTimeFormatSelect = (format: 'all' | 'date' | 'time' | 'ampm') => {
    if (!selectedCell || selectedColumn === null) return;

    const newMappings = mappings.filter(m => m.cellAddress !== selectedCell);
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
    onMappingsChange(mappings.filter(m => m.cellAddress !== cellAddress));
  };

  const getMappingForCell = (cellAddress: string) => {
    return mappings.find(m => m.cellAddress === cellAddress);
  };

  return (
    <>
      <CardHeader>
        <CardTitle>Step 3: Map Columns to Template Cells</CardTitle>
        <CardDescription>
          Click template cell, then pick which column go there!
        </CardDescription>
      </CardHeader>
      <CardContent className="space-y-6">
        <div className="grid grid-cols-2 gap-6">
          {/* Template preview */}
          <div>
            <h3 className="font-semibold text-slate-900 mb-2">Template Preview</h3>
            <p className="text-sm text-slate-600 mb-4">
              Click on cells to map data
            </p>
            <div className="border rounded-lg overflow-auto max-h-[500px] bg-white">
              <table className="border-collapse w-full">
                <thead>
                  <tr>
                    <th className="sticky top-0 left-0 z-20 bg-slate-100 border border-slate-300 w-12 h-8 text-xs font-semibold text-slate-600"></th>
                    {templateFile.preview[0]?.map((_, colIdx) => (
                      <th
                        key={colIdx}
                        className="sticky top-0 z-10 bg-slate-100 border border-slate-300 min-w-[100px] h-8 text-xs font-semibold text-slate-600"
                      >
                        {getColumnLabel(colIdx)}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {templateFile.preview.map((row, rowIdx) => (
                    <tr key={rowIdx}>
                      <td className="sticky left-0 z-10 bg-slate-100 border border-slate-300 text-center text-xs font-semibold text-slate-600 w-12 h-8">
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
                            className={`border border-slate-300 px-2 py-1 text-sm cursor-pointer select-none transition-colors min-w-[100px] h-8 ${
                              isSelected
                                ? 'bg-blue-100 border-blue-500 border-2'
                                : mapping
                                ? 'bg-green-100 border-green-500'
                                : 'hover:bg-slate-50'
                            }`}
                          >
                            <div className="flex items-center gap-1">
                              {mapping && (
                                <span className="text-xs text-green-700 font-medium">
                                  [{datesFile.headers[mapping.columnIndex]}]
                                </span>
                              )}
                              <span className={mapping ? 'text-slate-500' : ''}>
                                {String(cell || '')}
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
          <div>
            <h3 className="font-semibold text-slate-900 mb-2">
              {selectedCell ? `Select column for ${selectedCell}` : 'Available Columns'}
            </h3>
            <p className="text-sm text-slate-600 mb-4">
              {selectedCell ? 'Pick which column data to use' : 'Click a template cell first'}
            </p>

            {/* Date/Time format selector */}
            {selectedColumn !== null && (
              <div className="mb-4 p-4 bg-blue-50 border border-blue-200 rounded-lg">
                <h4 className="font-semibold text-blue-900 mb-2">
                  Date/Time Column Detected!
                </h4>
                <p className="text-sm text-blue-700 mb-3">
                  Column: {datesFile.headers[selectedColumn]}
                </p>
                <div className="space-y-2">
                  <button
                    onClick={() => handleDateTimeFormatSelect('all')}
                    className="w-full text-left px-4 py-2 rounded-lg border border-blue-300 bg-white hover:bg-blue-50 transition-colors"
                  >
                    <div className="font-medium">All (Full Date & Time)</div>
                    <div className="text-xs text-slate-500">e.g., "Mon 02.04.26 09:00 AM"</div>
                  </button>
                  <button
                    onClick={() => handleDateTimeFormatSelect('date')}
                    className="w-full text-left px-4 py-2 rounded-lg border border-blue-300 bg-white hover:bg-blue-50 transition-colors"
                  >
                    <div className="font-medium">Date Only</div>
                    <div className="text-xs text-slate-500">e.g., "02.04.26" (DD.MM.YY)</div>
                  </button>
                  <button
                    onClick={() => handleDateTimeFormatSelect('time')}
                    className="w-full text-left px-4 py-2 rounded-lg border border-blue-300 bg-white hover:bg-blue-50 transition-colors"
                  >
                    <div className="font-medium">Time Only</div>
                    <div className="text-xs text-slate-500">e.g., "09:00"</div>
                  </button>
                  <button
                    onClick={() => handleDateTimeFormatSelect('ampm')}
                    className="w-full text-left px-4 py-2 rounded-lg border border-blue-300 bg-white hover:bg-blue-50 transition-colors"
                  >
                    <div className="font-medium">AM/PM (Smart)</div>
                    <div className="text-xs text-slate-500">Auto: "AM", "MID-AM" (10am+), "Early PM" (12pm), "PM"</div>
                  </button>
                  <button
                    onClick={() => setSelectedColumn(null)}
                    className="w-full px-4 py-2 text-sm text-slate-600 hover:text-slate-900"
                  >
                    Cancel
                  </button>
                </div>
              </div>
            )}
            
            <div className="space-y-2">
              {datesFile.headers.map((header, idx) => {
                const isDateTime = isDateTimeColumn(idx);
                return (
                  <button
                    key={idx}
                    onClick={() => handleColumnSelect(idx)}
                    disabled={!selectedCell}
                    className={`w-full text-left px-4 py-2 rounded-lg border transition-colors ${
                      selectedCell
                        ? 'hover:bg-slate-100 border-slate-300 cursor-pointer'
                        : 'border-slate-200 text-slate-400 cursor-not-allowed'
                    }`}
                  >
                    <div className="flex items-center gap-2">
                      <div className="font-medium">{header}</div>
                      {isDateTime && (
                        <span className="text-xs bg-blue-100 text-blue-700 px-2 py-0.5 rounded">
                          Date/Time
                        </span>
                      )}
                    </div>
                    <div className="text-xs text-slate-500">Column {String.fromCharCode(65 + idx)}</div>
                  </button>
                );
              })}
            </div>

            {/* Current mappings */}
            {mappings.length > 0 && (
              <div className="mt-6">
                <h4 className="font-semibold text-slate-900 mb-2">Current Mappings</h4>
                <div className="space-y-1">
                  {mappings.map((mapping) => (
                    <div
                      key={mapping.cellAddress}
                      className="flex items-center justify-between px-3 py-2 bg-green-50 rounded border border-green-200"
                    >
                      <span className="text-sm">
                        <span className="font-medium">{mapping.cellAddress}</span>
                        {' → '}
                        <span className="text-green-700">
                          {datesFile.headers[mapping.columnIndex]}
                        </span>
                        {mapping.dateTimeFormat && (
                          <span className="ml-2 text-xs bg-blue-100 text-blue-700 px-2 py-0.5 rounded">
                            {mapping.dateTimeFormat === 'all' && 'Full'}
                            {mapping.dateTimeFormat === 'date' && 'Date Only'}
                            {mapping.dateTimeFormat === 'time' && 'Time Only'}
                            {mapping.dateTimeFormat === 'ampm' && 'AM/PM (Smart)'}
                          </span>
                        )}
                      </span>
                      <button
                        onClick={() => handleRemoveMapping(mapping.cellAddress)}
                        className="text-red-500 hover:text-red-700"
                      >
                        <Trash2 className="h-4 w-4" />
                      </button>
                    </div>
                  ))}
                </div>
              </div>
            )}
          </div>
        </div>

        <div className="flex justify-between pt-4">
          <Button variant="outline" onClick={onBack}>
            Back
          </Button>
          <Button
            onClick={onNext}
            disabled={mappings.length === 0}
            size="lg"
          >
            Next: Generate ({mappings.length} mappings)
          </Button>
        </div>
      </CardContent>
    </>
  );
}
