import { useState } from "react";
import { Button } from "./ui/button";
import { CardHeader, CardTitle, CardDescription, CardContent } from "./ui/card";
import type { ExcelData, TemplateData, CellMapping } from "./ExcelWizard";
import * as ExcelJS from "exceljs";
import { Download, Loader2 } from "lucide-react";

interface StepGenerateProps {
  datesFile: ExcelData;
  templateFile: TemplateData;
  selectedRows: Set<number>;
  mappings: CellMapping[];
  onBack: () => void;
  onReset: () => void;
}

type SheetPreview = {
  name: string;
  data: (string | number | null)[][];
  mappedCells: Set<string>;
};

export function StepGenerate({
  datesFile,
  templateFile,
  selectedRows,
  mappings,
  onBack,
  onReset,
}: StepGenerateProps) {
  const [isGenerating, setIsGenerating] = useState(false);
  const [previewData, setPreviewData] = useState<SheetPreview[] | null>(null);
  const [activeTab, setActiveTab] = useState(0);
  const [downloadUrl, setDownloadUrl] = useState<string | null>(null);

  const parseCellAddress = (address: string): { row: number; col: number } => {
    const match = address.match(/([A-Z]+)(\d+)/);
    if (!match) throw new Error(`Bad cell address: ${address}`);

    const colLetter = match[1];
    const rowNum = parseInt(match[2], 10);

    // Convert column letter to index (A=0, B=1, etc)
    let colIndex = 0;
    for (let i = 0; i < colLetter.length; i++) {
      colIndex = colIndex * 26 + (colLetter.charCodeAt(i) - 64);
    }
    colIndex -= 1; // Make 0-based

    return { row: rowNum - 1, col: colIndex };
  };

  const sanitizeSheetName = (name: string): string => {
    // Remove invalid characters: * ? : \ / [ ]
    let sanitized = String(name).replace(/[\*\?\:\\\/\[\]]/g, "");

    // Excel sheet names max 31 chars
    if (sanitized.length > 31) {
      sanitized = sanitized.substring(0, 31);
    }

    // Can't be empty
    if (!sanitized.trim()) {
      sanitized = "Sheet";
    }

    return sanitized.trim();
  };

  const formatDateTime = (
    value: any,
    format: "all" | "date" | "time" | "ampm" | "none",
  ): string => {
    if (!value) return "";

    if (format === "none") {
      // Strip double quotes from the value
      return String(value).replace(/^"(.*)"$/, '$1');
    }

    let date: Date;
    let timeStr = String(value);

    if (value instanceof Date) {
      date = value;
    } else {
      // Check if it's already a formatted string like "THUR 23.04.26\nPM 12:00"
      const customMatch = timeStr.match(
        /(\d{2})\.(\d{2})\.(\d{2})[^\d]*(\d{1,2}):(\d{2})/,
      );
      if (customMatch) {
        // Parse DD.MM.YY HH:MM format
        const day = parseInt(customMatch[1], 10);
        const month = parseInt(customMatch[2], 10) - 1;
        const year = 2000 + parseInt(customMatch[3], 10);
        const hours = parseInt(customMatch[4], 10);
        const minutes = parseInt(customMatch[5], 10);

        // Check for PM in the string
        const isPM = /PM/i.test(timeStr);
        const adjustedHours =
          isPM && hours !== 12 ? hours + 12 : !isPM && hours === 12 ? 0 : hours;

        date = new Date(year, month, day, adjustedHours, minutes);
      } else {
        // Try parsing dd/mm/yyyy or dd.mm.yyyy formats
        const dateOnlyMatch = timeStr.match(
          /(\d{1,2})[\/\.](\d{1,2})[\/\.](\d{2,4})/,
        );
        if (dateOnlyMatch) {
          const day = parseInt(dateOnlyMatch[1], 10);
          const month = parseInt(dateOnlyMatch[2], 10) - 1;
          let year = parseInt(dateOnlyMatch[3], 10);
          // Convert 2-digit year to 4-digit
          if (year < 100) {
            year += 2000;
          }

          // Check if there's also time in the string
          const timeMatch = timeStr.match(/(\d{1,2}):(\d{2})/);
          let hours = 0;
          let minutes = 0;

          if (timeMatch) {
            hours = parseInt(timeMatch[1], 10);
            minutes = parseInt(timeMatch[2], 10);

            // Check for PM/AM
            const isPM = /PM/i.test(timeStr);
            const isAM = /AM/i.test(timeStr);
            if (isPM && hours !== 12) {
              hours += 12;
            } else if (isAM && hours === 12) {
              hours = 0;
            }
          }

          date = new Date(year, month, day, hours, minutes);
        } else {
          // Try to parse as standard date
          date = new Date(value);
          if (isNaN(date.getTime())) {
            // If can't parse, return original
            console.log("Cannot parse date:", value);
            return String(value);
          }
        }
      }
    }

    const days = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
    const months = [
      "Jan",
      "Feb",
      "Mar",
      "Apr",
      "May",
      "Jun",
      "Jul",
      "Aug",
      "Sep",
      "Oct",
      "Nov",
      "Dec",
    ];

    const day = days[date.getDay()];
    const dateNum = date.getDate().toString().padStart(2, "0");
    const monthNum = (date.getMonth() + 1).toString().padStart(2, "0");
    const yearShort = date.getFullYear().toString().slice(-2);
    const hours = date.getHours();
    const minutes = date.getMinutes().toString().padStart(2, "0");
    const hours12 = hours % 12 || 12;

    // Smart AM/PM logic
    let ampmSmart: string;
    if (hours >= 10 && hours < 12) {
      ampmSmart = "MID-AM";
    } else if (hours === 12) {
      ampmSmart = "Early PM";
    } else if (hours >= 12) {
      ampmSmart = "PM";
    } else {
      ampmSmart = "AM";
    }

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

  const generateExcel = async () => {
    setIsGenerating(true);

    try {
      const workbook = new ExcelJS.Workbook();
      const selectedRowIndices = Array.from(selectedRows).sort((a, b) => a - b);
      const usedSheetNames = new Set<string>();
      const previewSheets: SheetPreview[] = [];

      for (const rowIndex of selectedRowIndices) {
        const rowData = datesFile.rows[rowIndex];
        const rawSheetName = String(rowData[0] || `Sheet_${rowIndex + 1}`);
        let sheetName = sanitizeSheetName(rawSheetName);

        // Handle duplicate sheet names
        let finalSheetName = sheetName;
        let duplicateCounter = 2;
        while (usedSheetNames.has(finalSheetName)) {
          // Add counter suffix, making sure not to exceed 31 chars
          const suffix = ` (${duplicateCounter})`;
          const maxBaseLength = 31 - suffix.length;
          const baseName = sheetName.substring(0, maxBaseLength);
          finalSheetName = `${baseName}${suffix}`;
          duplicateCounter++;
        }
        usedSheetNames.add(finalSheetName);

        // Copy template worksheet
        const templateWorksheet = templateFile.workbook.worksheets[0];
        const newSheet = workbook.addWorksheet(finalSheetName);

        // Determine sheet dimensions
        let maxRow = 0;
        let maxCol = 0;
        templateWorksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
          maxRow = Math.max(maxRow, rowNumber);
          row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            maxCol = Math.max(maxCol, colNumber);
          });
        });

        // Initialize preview data array
        const previewDataArray: (string | number | null)[][] = [];
        const mappedCells = new Set<string>();

        // Copy all cells from template
        templateWorksheet.eachRow({ includeEmpty: true }, (row, rowNumber) => {
          const newRow = newSheet.getRow(rowNumber);
          const previewRow: (string | number | null)[] = [];

          row.eachCell({ includeEmpty: true }, (cell, colNumber) => {
            const newCell = newRow.getCell(colNumber);

            // Copy value
            newCell.value = cell.value;

            // Copy basic formatting
            if (cell.style) {
              newCell.style = cell.style;
            }

            // Add to preview (convert cell value to displayable format)
            const cellValue = cell.value;
            if (cellValue === null || cellValue === undefined) {
              previewRow.push(null);
            } else if (typeof cellValue === 'object' && 'richText' in cellValue) {
              previewRow.push(cellValue.richText.map((t: any) => t.text).join(''));
            } else if (typeof cellValue === 'object' && 'formula' in cellValue) {
              previewRow.push(`=${cellValue.formula}`);
            } else {
              previewRow.push(String(cellValue));
            }
          });

          // Fill remaining columns with null if needed
          while (previewRow.length < maxCol) {
            previewRow.push(null);
          }

          previewDataArray.push(previewRow);
          newRow.commit();
        });

        // Apply mappings - replace mapped cells with data
        for (const mapping of mappings) {
          const { row: rowNum, col: colNum } = parseCellAddress(
            mapping.cellAddress,
          );
          const cell = newSheet.getRow(rowNum + 1).getCell(colNum + 1);
          let dataValue = rowData[mapping.columnIndex];

          // Apply date/time formatting if specified
          if (mapping.dateTimeFormat) {
            dataValue = formatDateTime(dataValue, mapping.dateTimeFormat);
          }

          cell.value = dataValue;

          // Update preview data with mapped value
          if (previewDataArray[rowNum]) {
            previewDataArray[rowNum][colNum] = dataValue === null || dataValue === undefined ? null : String(dataValue);
          }

          // Track which cells were mapped
          mappedCells.add(mapping.cellAddress);
        }

        // Copy column widths
        templateWorksheet.columns.forEach((col, idx) => {
          if (col.width) {
            const newCol = newSheet.getColumn(idx + 1);
            newCol.width = col.width;
          }
        });

        // Add to preview data
        previewSheets.push({
          name: finalSheetName,
          data: previewDataArray,
          mappedCells,
        });
      }

      // Generate download blob
      const buffer = await workbook.xlsx.writeBuffer();
      const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      const url = URL.createObjectURL(blob);

      // Store URL for manual download later
      setDownloadUrl(url);

      // Set preview data to show preview UI
      setPreviewData(previewSheets);
      setActiveTab(0);
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
        <CardTitle>Step 4: Generate Output</CardTitle>
        <CardDescription>
          {previewData 
            ? "Review your Excel file below, then download when ready!"
            : "Generate preview first, then download! One sheet per row!"
          }
        </CardDescription>
      </CardHeader>
      <CardContent className="space-y-6">
        {!previewData && (
          <>
            <div className="bg-slate-50 rounded-lg p-6 space-y-3">
              <div className="flex justify-between">
                <span className="text-slate-600">Selected Rows:</span>
                <span className="font-semibold">{selectedRows.size}</span>
              </div>
              <div className="flex justify-between">
                <span className="text-slate-600">Column Mappings:</span>
                <span className="font-semibold">{mappings.length}</span>
              </div>
              <div className="flex justify-between">
                <span className="text-slate-600">Output Sheets:</span>
                <span className="font-semibold">{selectedRows.size}</span>
              </div>
            </div>

            <div className="border rounded-lg p-4">
              <h4 className="font-semibold mb-2">What happen:</h4>
              <ul className="text-sm text-slate-600 space-y-1 list-disc list-inside">
                <li>
                  Create {selectedRows.size} new sheets (one per selected row)
                </li>
                <li>Each sheet named from Column A</li>
                <li>Template copied to each sheet</li>
                <li>{mappings.length} cells replaced with row data</li>
              </ul>
            </div>

            <Button
              onClick={generateExcel}
              disabled={isGenerating}
              size="lg"
              className="w-full"
            >
              {isGenerating ? (
                <>
                  <Loader2 className="mr-2 h-4 w-4 animate-spin" />
                  Generating Preview... Please wait!
                </>
              ) : (
                "Generate Preview"
              )}
            </Button>

            <div className="flex justify-between pt-4">
              <Button variant="outline" onClick={onBack}>
                Back
              </Button>
            </div>
          </>
        )}

        {previewData && (
          <>
            <div className="bg-blue-50 border border-blue-200 rounded-lg p-4">
              <p className="text-blue-800 font-semibold text-center">
                Preview Generated Successfully!
              </p>
              <p className="text-blue-700 text-sm text-center mt-1">
                Review your {previewData.length} sheet{previewData.length !== 1 ? 's' : ''} below, then download when ready
              </p>
            </div>

            {/* Sheet tabs */}
            <div className="border rounded-lg overflow-hidden">
              <div className="flex gap-1 p-2 bg-slate-100 overflow-x-auto">
                {previewData.map((sheet, index) => (
                  <button
                    key={index}
                    onClick={() => setActiveTab(index)}
                    className={`px-4 py-2 rounded whitespace-nowrap transition-colors ${
                      activeTab === index
                        ? "bg-white shadow-sm font-semibold text-slate-900"
                        : "bg-slate-50 text-slate-600 hover:bg-slate-200"
                    }`}
                  >
                    {sheet.name}
                  </button>
                ))}
              </div>

              {/* Sheet preview content */}
              <div className="p-4 bg-white max-h-[600px] overflow-auto">
                {previewData[activeTab] && (
                  <div className="overflow-x-auto">
                    <table className="border-collapse border border-slate-300 text-sm">
                      <thead>
                        <tr className="bg-slate-100">
                          <th className="border border-slate-300 px-2 py-1 font-semibold text-slate-600 w-12"></th>
                          {previewData[activeTab].data[0]?.map((_, colIndex) => (
                            <th
                              key={colIndex}
                              className="border border-slate-300 px-2 py-1 font-semibold text-slate-600"
                            >
                              {String.fromCharCode(65 + (colIndex % 26))}
                            </th>
                          ))}
                        </tr>
                      </thead>
                      <tbody>
                        {previewData[activeTab].data.map((row, rowIndex) => (
                          <tr key={rowIndex}>
                            <td className="border border-slate-300 px-2 py-1 text-center text-slate-500 bg-slate-50 font-medium">
                              {rowIndex + 1}
                            </td>
                            {row.map((cell, colIndex) => {
                              const cellAddress = `${String.fromCharCode(65 + colIndex)}${rowIndex + 1}`;
                              const isMapped = previewData[activeTab].mappedCells.has(cellAddress);
                              return (
                                <td
                                  key={colIndex}
                                  className={`border border-slate-300 px-2 py-1 ${
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
                )}
              </div>

              {/* Legend */}
              <div className="p-3 bg-slate-50 border-t border-slate-200">
                <div className="flex items-center gap-4 text-xs text-slate-600">
                  <div className="flex items-center gap-2">
                    <div className="w-4 h-4 bg-blue-50 border border-slate-300"></div>
                    <span>Mapped cells (populated from data)</span>
                  </div>
                  <div className="flex items-center gap-2">
                    <div className="w-4 h-4 bg-white border border-slate-300"></div>
                    <span>Template cells</span>
                  </div>
                </div>
              </div>
            </div>

            {/* Download and action buttons */}
            <div className="space-y-3">
              {downloadUrl && (
                <a href={downloadUrl} download="output.xlsx" className="block">
                  <Button size="lg" className="w-full">
                    <Download className="mr-2 h-4 w-4" />
                    Download Excel File
                  </Button>
                </a>
              )}

              <div className="flex gap-3">
                <Button variant="outline" onClick={() => { setPreviewData(null); setDownloadUrl(null); }} className="flex-1">
                  Back to Edit Mappings
                </Button>
                <Button variant="outline" onClick={onReset} className="flex-1">
                  Start Over
                </Button>
              </div>
            </div>
          </>
        )}
      </CardContent>
    </>
  );
}
