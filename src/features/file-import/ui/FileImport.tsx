import { ChangeEvent, useId, useState } from 'react'
import type { SelectedExcelFile } from '../model/types'
import './file-import.css'

function formatFileSize(bytes: number) {
  if (bytes < 1024 * 1024) {
    return `${Math.max(1, Math.round(bytes / 1024))} КБ`
  }

  return `${(bytes / 1024 / 1024).toFixed(1)} МБ`
}

export function FileImport() {
  const inputId = useId()
  const [selectedFile, setSelectedFile] = useState<SelectedExcelFile | null>(null)

  const handleFileChange = (event: ChangeEvent<HTMLInputElement>) => {
    const file = event.target.files?.[0]

    if (!file) return

    setSelectedFile({ file, name: file.name, size: file.size })
  }

  return (
    <section className="import-card" aria-labelledby="import-title">
      <div className="import-card__heading">
        <div>
          <h2 id="import-title">Загрузка файла</h2>
          <p>Поддерживаются форматы XLSX и XLS</p>
        </div>
        <span className="status-badge">Основа готова</span>
      </div>

      <div className="file-picker">
        <div className="file-picker__icon" aria-hidden="true">
          <svg viewBox="0 0 24 24" role="img">
            <path d="M12 16V4m0 0L7.5 8.5M12 4l4.5 4.5M5 14v4a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2v-4" />
          </svg>
        </div>

        <div className="file-picker__copy">
          <strong>{selectedFile ? selectedFile.name : 'Выберите Excel-файл'}</strong>
          <span>
            {selectedFile
              ? `${formatFileSize(selectedFile.size)} · готов к будущей обработке`
              : 'Позже здесь появятся проверка и импорт данных'}
          </span>
        </div>

        <label className="button" htmlFor={inputId}>
          {selectedFile ? 'Заменить файл' : 'Выбрать файл'}
        </label>
        <input
          id={inputId}
          className="visually-hidden"
          type="file"
          accept=".xlsx,.xls,application/vnd.ms-excel,application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
          onChange={handleFileChange}
        />
      </div>
    </section>
  )
}

