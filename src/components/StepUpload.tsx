import { useRef, useState } from "react";
import { Upload } from "lucide-react";
import { Button } from "./ui/button";
import { CardHeader, CardTitle, CardDescription, CardContent } from "./ui/card";
import type { ExcelData, TemplateData } from "./ExcelWizard";
import * as ExcelJS from "exceljs";

interface StepUploadProps {
  datesFile: ExcelData | null;
  templateFile: TemplateData | null;
  onDatesUpload: (data: ExcelData) => void;
  onTemplateUpload: (data: TemplateData) => void;
  onNext: () => void;
}

export function StepUpload({
  datesFile,
  templateFile,
  onDatesUpload,
  onTemplateUpload,
  onNext,
}: StepUploadProps) {
  const datesInputRef = useRef<HTMLInputElement>(null);
  const templateInputRef = useRef<HTMLInputElement>(null);
  const [isDatesOver, setIsDatesOver] = useState(false);
  const [isTemplateOver, setIsTemplateOver] = useState(false);

  const processFile = async (file: File) => {
    const arrayBuffer = await file.arrayBuffer();
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.load(arrayBuffer);
    return { workbook, worksheet: workbook.worksheets[0] };
  };

  const getCellDisplayValue = (cellValue: any): any => {
    // If cell has formula, use the result value
    if (cellValue && typeof cellValue === "object" && "result" in cellValue) {
      return cellValue.result;
    }
    // If it's still an object, log it and try to extract useful info
    if (cellValue && typeof cellValue === "object") {
      // Check for common ExcelJS value types
      if ("text" in cellValue) return cellValue.text;
      if ("richText" in cellValue) {
        return cellValue.richText.map((rt: any) => rt.text).join("");
      }
      if ("hyperlink" in cellValue) return cellValue.hyperlink;
      if ("formula" in cellValue) return cellValue.formula;
      // If we can't extract anything useful, stringify it
      return JSON.stringify(cellValue);
    }
    return cellValue;
  };

  const processDatesFile = async (file: File) => {
    const { worksheet } = await processFile(file);
    const rows: any[][] = [];
    let headers: string[] = [];

    worksheet.eachRow((row, rowIndex) => {
      const rowData: any[] = [];
      row.eachCell({ includeEmpty: true }, (cell) => {
        rowData.push(getCellDisplayValue(cell.value));
      });

      if (rowIndex === 1) {
        headers = rowData.map((h) => String(h || ""));
      } else {
        rows.push(rowData);
      }
    });

    onDatesUpload({ headers, rows });
  };

  const processTemplateFile = async (file: File) => {
    const { workbook, worksheet } = await processFile(file);
    const preview: any[][] = [];

    // Find actual used range - look at all rows/cols with any data
    let maxRow = 0;
    let maxCol = 0;

    worksheet.eachRow({ includeEmpty: false }, (row, rowNumber) => {
      maxRow = Math.max(maxRow, rowNumber);
      row.eachCell({ includeEmpty: false }, (cell, colNumber) => {
        maxCol = Math.max(maxCol, colNumber);
      });
    });

    // Show at least 30 rows and 15 columns, or more if data exists
    const rowCount = Math.max(30, maxRow);
    const colCount = Math.max(15, maxCol);

    // Iterate through ALL rows by index to include blank rows
    for (let rowIdx = 1; rowIdx <= rowCount; rowIdx++) {
      const row = worksheet.getRow(rowIdx);
      const rowData: any[] = [];

      // Get all cells in the row, including empty ones
      for (let colIdx = 1; colIdx <= colCount; colIdx++) {
        const cell = row.getCell(colIdx);
        rowData.push(getCellDisplayValue(cell.value));
      }

      preview.push(rowData);
    }

    onTemplateUpload({
      workbook,
      sheetName: worksheet.name,
      preview,
    });
  };

  const handleDatesUpload = async (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;
    await processDatesFile(file);
  };

  const handleTemplateUpload = async (
    e: React.ChangeEvent<HTMLInputElement>,
  ) => {
    const file = e.target.files?.[0];
    if (!file) return;
    await processTemplateFile(file);
  };

  // Drag and drop handlers for dates file
  const handleDatesDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsDatesOver(true);
  };

  const handleDatesDragLeave = () => {
    setIsDatesOver(false);
  };

  const handleDatesDrop = async (e: React.DragEvent) => {
    e.preventDefault();
    setIsDatesOver(false);

    const file = e.dataTransfer.files?.[0];
    if (file && (file.name.endsWith(".xlsx") || file.name.endsWith(".xls"))) {
      await processDatesFile(file);
    }
  };

  // Drag and drop handlers for template file
  const handleTemplateDragOver = (e: React.DragEvent) => {
    e.preventDefault();
    setIsTemplateOver(true);
  };

  const handleTemplateDragLeave = () => {
    setIsTemplateOver(false);
  };

  const handleTemplateDrop = async (e: React.DragEvent) => {
    e.preventDefault();
    setIsTemplateOver(false);

    const file = e.dataTransfer.files?.[0];
    if (file && (file.name.endsWith(".xlsx") || file.name.endsWith(".xls"))) {
      await processTemplateFile(file);
    }
  };

  return (
    <>
      <CardHeader>
        <CardTitle>Step 1: Upload Files</CardTitle>
        <CardDescription>
          Drop dates file and template file here. Me read them!
        </CardDescription>
      </CardHeader>
      <CardContent className="space-y-6">
        {/* Dates file upload */}
        <div>
          <label className="block text-sm font-medium text-slate-900 mb-2">
            Dates File (with data rows)
          </label>
          <div
            onClick={() => datesInputRef.current?.click()}
            onDragOver={handleDatesDragOver}
            onDragLeave={handleDatesDragLeave}
            onDrop={handleDatesDrop}
            className={`border-2 border-dashed rounded-lg p-8 text-center cursor-pointer transition-colors ${
              isDatesOver
                ? "border-slate-900 bg-slate-50"
                : "border-slate-300 hover:border-slate-400"
            }`}
          >
            <Upload className="mx-auto h-12 w-12 text-slate-400 mb-2" />
            <p className="text-sm text-slate-600">
              {datesFile ? (
                <span className="text-slate-900 font-medium">
                  Loaded! {datesFile.rows.length} rows found
                </span>
              ) : (
                <>
                  <span className="font-medium">Click or drag file here</span>
                  <br />
                  <span className="text-xs">Excel file (.xlsx, .xls)</span>
                </>
              )}
            </p>
          </div>
          <input
            ref={datesInputRef}
            type="file"
            accept=".xlsx,.xls"
            className="hidden"
            onChange={handleDatesUpload}
          />
        </div>

        {/* Template file upload */}
        <div>
          <label className="block text-sm font-medium text-slate-900 mb-2">
            Template File (for output sheets)
          </label>
          <div
            onClick={() => templateInputRef.current?.click()}
            onDragOver={handleTemplateDragOver}
            onDragLeave={handleTemplateDragLeave}
            onDrop={handleTemplateDrop}
            className={`border-2 border-dashed rounded-lg p-8 text-center cursor-pointer transition-colors ${
              isTemplateOver
                ? "border-slate-900 bg-slate-50"
                : "border-slate-300 hover:border-slate-400"
            }`}
          >
            <Upload className="mx-auto h-12 w-12 text-slate-400 mb-2" />
            <p className="text-sm text-slate-600">
              {templateFile ? (
                <span className="text-slate-900 font-medium">
                  Loaded! Sheet: {templateFile.sheetName}
                </span>
              ) : (
                <>
                  <span className="font-medium">Click or drag file here</span>
                  <br />
                  <span className="text-xs">Excel file (.xlsx, .xls)</span>
                </>
              )}
            </p>
          </div>
          <input
            ref={templateInputRef}
            type="file"
            accept=".xlsx,.xls"
            className="hidden"
            onChange={handleTemplateUpload}
          />
        </div>

        <div className="flex justify-end pt-4">
          <Button
            onClick={onNext}
            disabled={!datesFile || !templateFile}
            size="lg"
          >
            Next: Select Rows
          </Button>
        </div>
      </CardContent>
    </>
  );
}
