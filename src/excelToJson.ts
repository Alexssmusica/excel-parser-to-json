import XLSX from 'multiverse-xlsx';
import extend from 'node.extend';

// Public interfaces (as requested)
export interface SheetConfig {
    header?: { rows: number } | undefined;
    range?: string | undefined;
    columnToKey?: { [key: string]: string } | undefined;
    includeEmptyLines?: boolean | undefined;
    sheetStubs?: boolean | undefined;
}

type ExcelToJsonConfig = (({ sourceFile: string } | { source: string | Buffer })
  & { sheets?: ReadonlyArray<string | (SheetConfig & { name: string })> | undefined }
  & SheetConfig);

// Internal helper types to preserve existing logic
type ExcelPrimitive = string | number | Date | null | undefined;
type RowData = Record<string, ExcelPrimitive>;
type ParsedSheetRows = Array<RowData | undefined>;
type ParsedData = { [key: string]: any[] };

interface InternalHeaderOptions {
    rows?: number;
    rowToKeys?: number | string;
}

interface InternalSheetConfig {
    name: string;
    header?: InternalHeaderOptions;
    range?: string;
    columnToKey?: { [key: string]: string };
    includeEmptyLines?: boolean;
    sheetStubs?: boolean;
    appendData?: Record<string, unknown>;
}

interface SheetsLimiter {
    numberOfSheetsToGet?: number;
}

type InternalSheetsConfig = Array<string | InternalSheetConfig> | SheetsLimiter;

interface InternalConvertConfig {
    sourceFile?: string;
    source?: Buffer | Uint8Array | string;
    header?: InternalHeaderOptions;
    range?: string;
    columnToKey?: { [key: string]: string };
    includeEmptyLines?: boolean;
    sheetStubs?: boolean;
    sheets?: InternalSheetsConfig;
}

interface WorkbookLike {
    Sheets: Record<string, any>;
}

type SheetCell = {
    t: string;
    v?: any;
    w?: any;
} | undefined;

const excelToJson = (function() {

    let _config: InternalConvertConfig = {} as InternalConvertConfig;

    const getCellRow = (cell: string): number => Number(cell.replace(/[A-z]/gi, ''));
    const getCellColumn = (cell: string): string => cell.replace(/[0-9]/g, '').toUpperCase();
    const getRangeBegin = (cell: string): string => (cell.match(/^[^:]*/)?.[0]) as string;
    const getRangeEnd = (cell: string): string => (cell.match(/[^:]*$/)?.[0]) as string;
    function getSheetCellValue(sheetCell: SheetCell): ExcelPrimitive {
        if (!sheetCell) {
            return undefined;
        }
        if (sheetCell.t === 'z' && _config.sheetStubs) {
            return null;
        }
        return (sheetCell.t === 'n' || sheetCell.t === 'd') ? sheetCell.v : (sheetCell.w && sheetCell.w.trim && sheetCell.w.trim()) || sheetCell.w;
    };

    const parseSheet = (sheetData: string | InternalSheetConfig, workbook: WorkbookLike): ParsedSheetRows => {
        const sheetName: string = ((sheetData as any).constructor == String) ? (sheetData as string) : (sheetData as InternalSheetConfig).name;
        const sheet: any = workbook.Sheets[sheetName];
        const columnToKey: { [key: string]: string } | undefined = (typeof sheetData !== 'string' && sheetData.columnToKey) || _config.columnToKey;
        const range: string | undefined = (typeof sheetData !== 'string' && sheetData.range) || _config.range;
        const headerRows: number | undefined = ((typeof sheetData !== 'string' && sheetData.header && sheetData.header.rows) || (_config.header && _config.header.rows)) as number | undefined;
        const headerRowToKeys: number | string | undefined = ((typeof sheetData !== 'string' && (sheetData.header as InternalHeaderOptions | undefined)?.rowToKeys) || (_config.header && (_config.header as InternalHeaderOptions).rowToKeys)) as number | string | undefined;

        let strictRangeColumns: { from: string; to: string } | undefined;
        let strictRangeRows: { from: number; to: number } | undefined;
        if (range) {
            strictRangeColumns = {
                from: getCellColumn(getRangeBegin(range)),
                to: getCellColumn(getRangeEnd(range))
            };

            strictRangeRows = {
                from: getCellRow(getRangeBegin(range)),
                to: getCellRow(getRangeEnd(range))
            };
        }

        let rows: ParsedSheetRows = [];
        for (let cell in sheet) {

            // !ref is not a data to be retrieved || this cell doesn't have a value
            if (cell == '!ref' || (sheet[cell].v === undefined && !(_config.sheetStubs && sheet[cell].t === 'z'))) {
                continue;
            }

            const row = getCellRow(cell);
            const column = getCellColumn(cell);

            // Is a Header row
            if (headerRows && row <= headerRows) {
                continue;
            }

            // This column is not _configured to be retrieved
            if (columnToKey && !(columnToKey[column] || columnToKey['*'])) {
                continue;
            }

            // This cell is out of the _configured range
            if ((strictRangeColumns && strictRangeRows) && (column < strictRangeColumns.from || column > strictRangeColumns.to || row < strictRangeRows.from || row > strictRangeRows.to)) {
                continue;
            }


            if (!rows[row]) rows[row] = {} as RowData;
            const rowData = rows[row] as RowData;
            let columnData: string = (columnToKey && (columnToKey[column] || columnToKey['*'])) ?
                (columnToKey[column] || columnToKey['*']) :
                (headerRowToKeys ? `{{${column}${headerRowToKeys}}}` : column);

            let dataVariables = columnData.match(/{{([^}}]+)}}/g);
            if (dataVariables) {
                dataVariables.forEach(dataVariable => {
                    let dataVariableRef = dataVariable.replace(/[\{\}]*/gi, '');
                    let variableValue: ExcelPrimitive;
                    switch (dataVariableRef) {
                        case 'columnHeader':
                            dataVariableRef = (headerRows) ? `${column}${headerRows}` : `${column + 1}`;
                        default:
                            variableValue = getSheetCellValue(sheet[dataVariableRef]);
                    }
                    columnData = columnData.replace(dataVariable, String(variableValue));
                });
            }

            if (columnData === '') {
                continue;
            }

            rowData[columnData] = getSheetCellValue(sheet[cell] as SheetCell);
            
            if (typeof sheetData !== 'string' && sheetData.appendData) {
                extend(true, rowData, sheetData.appendData);
            }
        }

        // removing first row i.e. 0th rows because first cell itself starts from A1
        rows.shift();

        // Cleaning empty if required
        if (!_config.includeEmptyLines) {
            rows = rows.filter(v => v !== null && v !== undefined);
        }

        return rows;
    };

    const convertExcelToJson = function(config: ExcelToJsonConfig | string, sourceFile?: string): ParsedData {
        _config = (config as any).constructor === String ? JSON.parse(config as string) : (config as InternalConvertConfig);
        _config.sourceFile = _config.sourceFile || sourceFile;
        
        // ignoring empty lines by default
        _config.includeEmptyLines = _config.includeEmptyLines || false;

        // at least sourceFile or source has to be defined and have a value
        if (!(_config.sourceFile || _config.source)) {
            throw new Error(':: \'sourceFile\' or \'source\' required for _config :: ');
        }

        let workbook: WorkbookLike;

        if (_config.source) {
            workbook = XLSX.read(_config.source as any, {
                sheetStubs: true,
                cellDates: true
            }) as unknown as WorkbookLike;
        } else {
            workbook = XLSX.readFile(_config.sourceFile as string, {
                sheetStubs: true,
                cellDates: true
            }) as unknown as WorkbookLike;
        }

        let sheetsToGet: Array<string | InternalSheetConfig> = ((_config.sheets as any) && ((_config.sheets as any).constructor === Array)) ?
            (_config.sheets as Array<string | InternalSheetConfig>) :
            Object.keys(workbook.Sheets).slice(0, (_config && (_config.sheets as SheetsLimiter | undefined) && (_config.sheets as SheetsLimiter).numberOfSheetsToGet) || undefined);

        let parsedData: ParsedData = {};
        sheetsToGet.forEach((sheetItem) => {
            const sheet: InternalSheetConfig = ((sheetItem as any).constructor == String) ? { name: sheetItem as string } as InternalSheetConfig : (sheetItem as InternalSheetConfig);

            parsedData[sheet.name] = parseSheet(sheet, workbook);
        });

        return parsedData;
    };

    return convertExcelToJson;
}());

export default excelToJson;
