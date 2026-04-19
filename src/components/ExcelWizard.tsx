import { useState } from 'react';
import type { Workbook } from 'exceljs';
import { StepUpload } from './StepUpload';
import { StepSelectRows } from './StepSelectRows';
import { StepMapping } from './StepMapping';
import { StepGenerate } from './StepGenerate';
import { Card } from './ui/card';

export type ExcelData = {
  headers: string[];
  rows: any[][];
};

export type TemplateData = {
  workbook: Workbook;
  sheetName: string;
  preview: any[][];
};

export type CellMapping = {
  cellAddress: string; // like "A1"
  columnIndex: number; // which column from dates file
  dateTimeFormat?: 'all' | 'date' | 'time' | 'ampm'; // for date/time columns
};

export function ExcelWizard() {
  const [step, setStep] = useState(1);
  const [datesFile, setDatesFile] = useState<ExcelData | null>(null);
  const [templateFile, setTemplateFile] = useState<TemplateData | null>(null);
  const [selectedRows, setSelectedRows] = useState<Set<number>>(new Set());
  const [mappings, setMappings] = useState<CellMapping[]>([]);

  return (
    <div className="min-h-screen bg-slate-50 p-8">
      <div className="max-w-7xl mx-auto">
        <h1 className="text-4xl font-bold text-slate-900 mb-2">Excel Sheet Maker</h1>
        <p className="text-slate-600 mb-8">
          Upload files, pick rows, map columns, get output! Easy like rock smash!
        </p>

        {/* Step indicator */}
        <div className="flex gap-2 mb-8">
          {[1, 2, 3, 4].map((s) => (
            <div
              key={s}
              className={`flex-1 h-2 rounded ${
                s <= step ? 'bg-slate-900' : 'bg-slate-200'
              }`}
            />
          ))}
        </div>

        <Card className="p-8">
          {step === 1 && (
            <StepUpload
              datesFile={datesFile}
              templateFile={templateFile}
              onDatesUpload={setDatesFile}
              onTemplateUpload={setTemplateFile}
              onNext={() => setStep(2)}
            />
          )}

          {step === 2 && datesFile && (
            <StepSelectRows
              datesFile={datesFile}
              selectedRows={selectedRows}
              onSelectionChange={setSelectedRows}
              onBack={() => setStep(1)}
              onNext={() => setStep(3)}
            />
          )}

          {step === 3 && datesFile && templateFile && (
            <StepMapping
              datesFile={datesFile}
              templateFile={templateFile}
              mappings={mappings}
              onMappingsChange={setMappings}
              onBack={() => setStep(2)}
              onNext={() => setStep(4)}
            />
          )}

          {step === 4 && datesFile && templateFile && (
            <StepGenerate
              datesFile={datesFile}
              templateFile={templateFile}
              selectedRows={selectedRows}
              mappings={mappings}
              onBack={() => setStep(3)}
              onReset={() => {
                setStep(1);
                setDatesFile(null);
                setTemplateFile(null);
                setSelectedRows(new Set());
                setMappings([]);
              }}
            />
          )}
        </Card>
      </div>
    </div>
  );
}
