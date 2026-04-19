import { useState } from 'react';
import { Button } from './ui/button';
import { CardHeader, CardTitle, CardDescription, CardContent } from './ui/card';
import { Table, TableHeader, TableBody, TableHead, TableRow, TableCell } from './ui/table';
import { Checkbox } from './ui/checkbox';
import type { ExcelData } from './ExcelWizard';

interface StepSelectRowsProps {
  datesFile: ExcelData;
  selectedRows: Set<number>;
  onSelectionChange: (selected: Set<number>) => void;
  onBack: () => void;
  onNext: () => void;
}

export function StepSelectRows({
  datesFile,
  selectedRows,
  onSelectionChange,
  onBack,
  onNext,
}: StepSelectRowsProps) {
  const [lastClickedIndex, setLastClickedIndex] = useState<number | null>(null);

  const handleRowClick = (rowIndex: number, e: React.MouseEvent) => {
    // Prevent text selection when shift-clicking
    if (e.shiftKey) {
      e.preventDefault();
    }

    const newSelected = new Set(selectedRows);

    if (e.shiftKey && lastClickedIndex !== null) {
      // Shift-click: select range
      const start = Math.min(lastClickedIndex, rowIndex);
      const end = Math.max(lastClickedIndex, rowIndex);
      for (let i = start; i <= end; i++) {
        newSelected.add(i);
      }
    } else if (e.ctrlKey || e.metaKey) {
      // Ctrl-click: toggle individual
      if (newSelected.has(rowIndex)) {
        newSelected.delete(rowIndex);
      } else {
        newSelected.add(rowIndex);
      }
    } else {
      // Normal click: toggle this row
      if (newSelected.has(rowIndex)) {
        newSelected.delete(rowIndex);
      } else {
        newSelected.add(rowIndex);
      }
    }

    setLastClickedIndex(rowIndex);
    onSelectionChange(newSelected);
  };

  const handleCheckboxChange = (rowIndex: number, checked: boolean) => {
    const newSelected = new Set(selectedRows);
    if (checked) {
      newSelected.add(rowIndex);
    } else {
      newSelected.delete(rowIndex);
    }
    onSelectionChange(newSelected);
    setLastClickedIndex(rowIndex);
  };

  const handleSelectAll = () => {
    const newSelected = new Set<number>();
    datesFile.rows.forEach((_, idx) => newSelected.add(idx));
    onSelectionChange(newSelected);
  };

  const handleDeselectAll = () => {
    onSelectionChange(new Set());
  };

  return (
    <>
      <CardHeader>
        <CardTitle>Step 2: Select Rows</CardTitle>
        <CardDescription>
          Check rows to include! Use Shift-click for range, Ctrl-click for many!
        </CardDescription>
      </CardHeader>
      <CardContent className="space-y-4">
        <div className="flex gap-2 mb-4">
          <Button variant="outline" size="sm" onClick={handleSelectAll}>
            Select All
          </Button>
          <Button variant="outline" size="sm" onClick={handleDeselectAll}>
            Deselect All
          </Button>
          <div className="ml-auto text-sm text-slate-600">
            {selectedRows.size} of {datesFile.rows.length} rows selected
          </div>
        </div>

        <div className="border rounded-lg overflow-auto max-h-[500px]">
          <Table>
            <TableHeader>
              <TableRow className="hover:bg-transparent">
                <TableHead className="w-16">
                  <Checkbox
                    checked={selectedRows.size === datesFile.rows.length}
                    onCheckedChange={(checked) => {
                      if (checked) {
                        handleSelectAll();
                      } else {
                        handleDeselectAll();
                      }
                    }}
                  />
                </TableHead>
                <TableHead className="w-16">#</TableHead>
                {datesFile.headers.map((header, idx) => (
                  <TableHead key={idx}>{header}</TableHead>
                ))}
              </TableRow>
            </TableHeader>
            <TableBody>
              {datesFile.rows.map((row, rowIndex) => {
                const isSelected = selectedRows.has(rowIndex);
                return (
                  <TableRow
                    key={rowIndex}
                    onClick={(e) => handleRowClick(rowIndex, e)}
                    className={`cursor-pointer select-none transition-colors ${
                      isSelected ? 'bg-slate-100' : 'hover:bg-slate-50'
                    }`}
                  >
                    <TableCell>
                      <Checkbox
                        checked={isSelected}
                        onCheckedChange={(checked) => handleCheckboxChange(rowIndex, checked)}
                      />
                    </TableCell>
                    <TableCell className="font-medium text-slate-500">
                      {rowIndex + 1}
                    </TableCell>
                    {row.map((cell, cellIdx) => (
                      <TableCell key={cellIdx}>
                        {String(cell || '')}
                      </TableCell>
                    ))}
                  </TableRow>
                );
              })}
            </TableBody>
          </Table>
        </div>

        <div className="flex justify-between pt-4">
          <Button variant="outline" onClick={onBack}>
            Back
          </Button>
          <Button
            onClick={onNext}
            disabled={selectedRows.size === 0}
            size="lg"
          >
            Next: Map Columns ({selectedRows.size} selected)
          </Button>
        </div>
      </CardContent>
    </>
  );
}
