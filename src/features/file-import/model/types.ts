export const EXCEL_FILE_EXTENSIONS = ['.xlsx', '.xls'] as const

export type ExcelFileExtension = (typeof EXCEL_FILE_EXTENSIONS)[number]

export interface SelectedExcelFile {
  file: File
  name: string
  size: number
}

