(function () {
  var REQUIRED_COLUMNS = [
    'Фото',
    'Артикул',
    'Номенклатура',
    'Продажи Розница, руб.',
    'Продажи Розница, шт',
    'Остатки В РОЗНИЦЕ, шт',
    'Остатки (Арвато основной), шт',
  ]
  var DATABASE_NAME = 'retail-analytic-local'
  var DATABASE_VERSION = 1
  var DATABASE_STORE = 'state'

  var input = document.getElementById('excel-file')
  var picker = document.getElementById('file-picker')
  var fileName = document.getElementById('file-name')
  var description = document.getElementById('file-description')
  var fileButton = document.getElementById('file-button')
  var comparisonInput = document.getElementById('comparison-excel-file')
  var comparisonPicker = document.getElementById('comparison-file-picker')
  var comparisonFileName = document.getElementById('comparison-file-name')
  var comparisonDescription = document.getElementById('comparison-file-description')
  var comparisonFileButton = document.getElementById('comparison-file-button')
  var confirmButton = document.getElementById('confirm-button')
  var feedback = document.getElementById('feedback')
  var debugPanel = document.getElementById('debug-panel')
  var debugList = document.getElementById('debug-list')
  var results = document.getElementById('results')
  var resultsTitle = document.getElementById('results-title')
  var resultsMeta = document.getElementById('results-meta')
  var exportPdfButton = document.getElementById('export-pdf-button')
  var exportExcelButton = document.getElementById('export-excel-button')
  var tableHead = document.getElementById('table-head')
  var tableBody = document.getElementById('table-body')
  var selectedFile = null
  var comparisonFile = null
  var activeImageUrls = []
  var tableDataset = null
  var selectedArticleKeys = new Set()
  var manualPreviousSales = new Map()
  var manualCurrentSales = new Map()
  var databasePromise = null

  function formatFileSize(bytes) {
    if (bytes < 1024 * 1024) {
      return Math.max(1, Math.round(bytes / 1024)) + ' КБ'
    }

    return (bytes / 1024 / 1024).toFixed(1) + ' МБ'
  }

  function getDateFromFileName(file) {
    if (!file || !file.name) return null

    var match = file.name.match(/(?:^|[^\d])(\d{2})[.\-_](\d{2})[.\-_](\d{4})(?=[^\d]|$)/)
    if (!match) return null

    var day = Number(match[1])
    var month = Number(match[2])
    var year = Number(match[3])
    var date = new Date(year, month - 1, day)

    if (
      date.getFullYear() !== year
      || date.getMonth() !== month - 1
      || date.getDate() !== day
    ) {
      return null
    }

    return match[1] + '.' + match[2] + '.' + match[3]
  }

  function updateResultsTitle() {
    var previousDate = getDateFromFileName(selectedFile)
    var currentDate = getDateFromFileName(comparisonFile)

    if (previousDate && currentDate) {
      resultsTitle.textContent = previousDate + ' → ' + currentDate
    } else if (currentDate) {
      resultsTitle.textContent = currentDate
    } else if (previousDate) {
      resultsTitle.textContent = previousDate
    } else {
      resultsTitle.textContent = 'Данные из файла'
    }
  }

  function openDatabase() {
    if (databasePromise) return databasePromise

    databasePromise = new Promise(function (resolve, reject) {
      if (typeof indexedDB === 'undefined') {
        reject(new Error('Локальное хранилище недоступно'))
        return
      }

      var request = indexedDB.open(DATABASE_NAME, DATABASE_VERSION)

      request.addEventListener('upgradeneeded', function () {
        var database = request.result
        if (!database.objectStoreNames.contains(DATABASE_STORE)) {
          database.createObjectStore(DATABASE_STORE)
        }
      })
      request.addEventListener('success', function () {
        resolve(request.result)
      })
      request.addEventListener('error', function () {
        reject(request.error || new Error('Не удалось открыть локальное хранилище'))
      })
    })

    return databasePromise
  }

  async function readLocalState(key) {
    var database = await openDatabase()

    return new Promise(function (resolve, reject) {
      var request = database.transaction(DATABASE_STORE, 'readonly').objectStore(DATABASE_STORE).get(key)
      request.addEventListener('success', function () {
        resolve(request.result)
      })
      request.addEventListener('error', function () {
        reject(request.error || new Error('Не удалось прочитать локальные данные'))
      })
    })
  }

  async function writeLocalState(key, value) {
    var database = await openDatabase()

    return new Promise(function (resolve, reject) {
      var transaction = database.transaction(DATABASE_STORE, 'readwrite')
      transaction.objectStore(DATABASE_STORE).put(value, key)
      transaction.addEventListener('complete', resolve)
      transaction.addEventListener('error', function () {
        reject(transaction.error || new Error('Не удалось сохранить локальные данные'))
      })
      transaction.addEventListener('abort', function () {
        reject(transaction.error || new Error('Сохранение локальных данных отменено'))
      })
    })
  }

  function showPersistenceWarning() {
    feedback.textContent = 'Данные загружены, но браузер не разрешил сохранить их локально.'
    feedback.classList.remove('feedback--error')
  }

  function persistSelectedArticles() {
    return writeLocalState('articles', Array.from(selectedArticleKeys))
  }

  function persistManualSales() {
    return writeLocalState('manualSales', {
      current: Array.from(manualCurrentSales.entries()),
      previous: Array.from(manualPreviousSales.entries()),
    })
  }

  function restoreManualSales(state) {
    manualPreviousSales = new Map(state && Array.isArray(state.previous) ? state.previous : [])
    manualCurrentSales = new Map(state && Array.isArray(state.current) ? state.current : [])
  }

  async function persistFile(key, arrayBuffer, file) {
    await writeLocalState(key, {
      buffer: arrayBuffer,
      lastModified: file.lastModified || Date.now(),
      name: file.name,
      type: file.type || 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
    })
  }

  function normalizeHeader(value) {
    return String(value == null ? '' : value)
      .replace(/\s+/g, ' ')
      .trim()
      .toLocaleLowerCase('ru-RU')
  }

  function resetResults() {
    activeImageUrls.forEach(function (url) {
      URL.revokeObjectURL(url)
    })
    activeImageUrls = []
    tableDataset = null
    selectedArticleKeys = new Set()
    manualPreviousSales = new Map()
    manualCurrentSales = new Map()
    results.hidden = true
    tableHead.replaceChildren()
    tableBody.replaceChildren()
    feedback.textContent = ''
    feedback.classList.remove('feedback--error')
    debugPanel.hidden = true
    debugList.replaceChildren()
  }

  function showFile(file) {
    resetResults()
    selectedFile = null
    confirmButton.hidden = true
    picker.classList.remove('file-picker--error', 'file-picker--selected')

    if (!file) return

    var extension = file.name.split('.').pop().toLowerCase()

    if (extension !== 'xlsx' && extension !== 'xls') {
      fileName.textContent = 'Неподдерживаемый формат'
      description.textContent = 'Выберите файл XLSX или XLS'
      picker.classList.add('file-picker--error')
      return
    }

    selectedFile = file
    fileName.textContent = file.name
    description.textContent = formatFileSize(file.size) + ' · файл готов к обработке'
    fileButton.textContent = 'Заменить файл'
    confirmButton.hidden = false
    picker.classList.add('file-picker--selected')
  }

  function showComparisonFile(file) {
    comparisonFile = null
    comparisonPicker.classList.remove('file-picker--error', 'file-picker--selected')

    if (!file) return

    var extension = file.name.split('.').pop().toLowerCase()

    if (extension !== 'xlsx' && extension !== 'xls') {
      comparisonFileName.textContent = 'Неподдерживаемый формат'
      comparisonDescription.textContent = 'Выберите файл XLSX или XLS'
      comparisonPicker.classList.add('file-picker--error')
      return
    }

    comparisonFile = file
    manualCurrentSales = new Map()
    comparisonFileName.textContent = file.name
    comparisonDescription.textContent = formatFileSize(file.size) + ' · данные готовы к выводу'
    comparisonFileButton.textContent = 'Заменить файл'
    comparisonPicker.classList.add('file-picker--selected')
  }

  function analyzeHeaders(rows) {
    var required = REQUIRED_COLUMNS.map(normalizeHeader)
    var bestMatch = {
      rowIndex: -1,
      foundCount: -1,
      missingColumns: REQUIRED_COLUMNS.slice(),
    }

    rows.slice(0, 100).forEach(function (row, rowIndex) {
      var normalizedRow = row.map(normalizeHeader)
      var missingColumns = REQUIRED_COLUMNS.filter(function (_, columnIndex) {
        return normalizedRow.indexOf(required[columnIndex]) === -1
      })
      var foundCount = REQUIRED_COLUMNS.length - missingColumns.length

      if (foundCount > bestMatch.foundCount) {
        bestMatch = {
          rowIndex: rowIndex,
          foundCount: foundCount,
          missingColumns: missingColumns,
        }
      }
    })

    return bestMatch
  }

  function showMissingColumns(columns) {
    var fragment = document.createDocumentFragment()

    columns.forEach(function (column) {
      var item = document.createElement('li')
      var label = document.createElement('code')
      label.textContent = column
      item.appendChild(label)
      fragment.appendChild(item)
    })

    debugList.replaceChildren(fragment)
    debugPanel.hidden = false
  }

  function resolveZipPath(basePath, targetPath) {
    var parts = basePath.split('/')
    parts.pop()

    targetPath.split('/').forEach(function (part) {
      if (!part || part === '.') return
      if (part === '..') {
        parts.pop()
      } else {
        parts.push(part)
      }
    })

    return parts.join('/')
  }

  function getRelationshipsPath(filePath) {
    var parts = filePath.split('/')
    var fileName = parts.pop()
    return parts.join('/') + '/_rels/' + fileName + '.rels'
  }

  async function readXml(zip, path) {
    var entry = zip.file(path)
    if (!entry) throw new Error('В Excel-файле отсутствует служебный файл ' + path + '.')

    var contents = await entry.async('string')
    return new DOMParser().parseFromString(contents, 'application/xml')
  }

  function findRelationship(xml, predicate) {
    return Array.from(xml.getElementsByTagNameNS('*', 'Relationship')).find(predicate)
  }

  function getRelationshipId(element) {
    return element.getAttributeNS(
      'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
      'id',
    ) || element.getAttribute('r:id')
  }

  function getMimeType(path) {
    var extension = path.split('.').pop().toLowerCase()
    if (extension === 'png') return 'image/png'
    if (extension === 'webp') return 'image/webp'
    if (extension === 'gif') return 'image/gif'
    return 'image/jpeg'
  }

  async function extractRetailImages(arrayBuffer, sourceFileName) {
    var imagesByRow = new Map()

    if (typeof JSZip === 'undefined' || !sourceFileName.toLowerCase().endsWith('.xlsx')) {
      return imagesByRow
    }

    var zip = await JSZip.loadAsync(arrayBuffer)
    var workbookPath = 'xl/workbook.xml'
    var workbookXml = await readXml(zip, workbookPath)
    var workbookRelsXml = await readXml(zip, getRelationshipsPath(workbookPath))
    var retailSheet = Array.from(workbookXml.getElementsByTagNameNS('*', 'sheet')).find(function (sheet) {
      return normalizeHeader(sheet.getAttribute('name')) === 'retail'
    })

    if (!retailSheet) return imagesByRow

    var sheetRelationshipId = getRelationshipId(retailSheet)
    var sheetRelationship = findRelationship(workbookRelsXml, function (relationship) {
      return relationship.getAttribute('Id') === sheetRelationshipId
    })

    if (!sheetRelationship) return imagesByRow

    var sheetPath = resolveZipPath(workbookPath, sheetRelationship.getAttribute('Target'))
    var sheetRelsEntry = zip.file(getRelationshipsPath(sheetPath))
    if (!sheetRelsEntry) return imagesByRow

    var sheetRelsXml = await readXml(zip, getRelationshipsPath(sheetPath))
    var drawingRelationship = findRelationship(sheetRelsXml, function (relationship) {
      return relationship.getAttribute('Type').endsWith('/drawing')
    })
    if (!drawingRelationship) return imagesByRow

    var drawingPath = resolveZipPath(sheetPath, drawingRelationship.getAttribute('Target'))
    var drawingXml = await readXml(zip, drawingPath)
    var drawingRelsXml = await readXml(zip, getRelationshipsPath(drawingPath))
    var imageRelationships = new Map()

    Array.from(drawingRelsXml.getElementsByTagNameNS('*', 'Relationship')).forEach(function (relationship) {
      imageRelationships.set(
        relationship.getAttribute('Id'),
        resolveZipPath(drawingPath, relationship.getAttribute('Target')),
      )
    })

    var anchors = Array.from(drawingXml.getElementsByTagNameNS('*', 'twoCellAnchor')).concat(
      Array.from(drawingXml.getElementsByTagNameNS('*', 'oneCellAnchor')),
    )

    var extractedImages = await Promise.all(anchors.map(async function (anchor) {
      var from = anchor.getElementsByTagNameNS('*', 'from')[0]
      var row = from && from.getElementsByTagNameNS('*', 'row')[0]
      var column = from && from.getElementsByTagNameNS('*', 'col')[0]
      var blip = anchor.getElementsByTagNameNS('*', 'blip')[0]
      var imageRelationshipId = blip && (
        blip.getAttributeNS(
          'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
          'embed',
        ) || blip.getAttribute('r:embed')
      )
      var imagePath = imageRelationships.get(imageRelationshipId)

      if (!row || !column || Number(column.textContent) !== 0 || !imagePath) return null

      var imageEntry = zip.file(imagePath)
      if (!imageEntry) return null

      var bytes = await imageEntry.async('uint8array')
      var url = URL.createObjectURL(new Blob([bytes], { type: getMimeType(imagePath) }))
      activeImageUrls.push(url)

      return { rowIndex: Number(row.textContent), url: url }
    }))

    extractedImages.forEach(function (image) {
      if (image) imagesByRow.set(image.rowIndex, image.url)
    })

    return imagesByRow
  }

  function normalizeArticle(value) {
    return String(value == null ? '' : value).replace(/\s+/g, '').toLocaleLowerCase('ru-RU')
  }

  function formatNumber(value, maximumFractionDigits) {
    return new Intl.NumberFormat('ru-RU', {
      maximumFractionDigits: maximumFractionDigits == null ? 2 : maximumFractionDigits,
    }).format(value)
  }

  function formatCellValue(value, columnIndex) {
    if (value == null || value === '') return '—'
    if (columnIndex >= 3 && typeof value === 'number') return formatNumber(value)
    return String(value)
  }

  function formatSignedNumber(value, maximumFractionDigits) {
    if (value === 0) return formatNumber(0, maximumFractionDigits)
    return (value > 0 ? '+' : '−') + formatNumber(Math.abs(value), maximumFractionDigits)
  }

  function parseManualSalesValue(value) {
    var normalized = String(value == null ? '' : value)
      .replace(/[\s\u00a0]/g, '')
      .replace(',', '.')
    var number = Number(normalized)

    return normalized !== '' && Number.isFinite(number) ? number : null
  }

  function createMissingSalesEditor(label, side, onSave) {
    var button = document.createElement('button')

    button.className = 'missing-sales-trigger'
    button.type = 'button'
    button.textContent = label
    button.setAttribute('aria-label', label + '. Ввести значение вручную')
    button.addEventListener('click', function () {
      var form = document.createElement('form')
      var input = document.createElement('input')
      var submitButton = document.createElement('button')
      var error = document.createElement('span')

      form.className = 'manual-sales-form'
      input.className = 'manual-sales-input'
      input.type = 'text'
      input.inputMode = 'decimal'
      input.autocomplete = 'off'
      input.placeholder = 'Сумма, руб.'
      input.setAttribute('aria-label', 'Продажи за ' + (side === 'previous' ? 'предыдущую' : 'текущую') + ' неделю')
      submitButton.className = 'manual-sales-submit'
      submitButton.type = 'submit'
      submitButton.textContent = 'Сохранить'
      error.className = 'manual-sales-error'
      error.setAttribute('aria-live', 'polite')
      form.append(input, submitButton, error)
      button.replaceWith(form)
      input.focus()

      form.addEventListener('submit', function (event) {
        event.preventDefault()
        var number = parseManualSalesValue(input.value)

        if (number == null) {
          error.textContent = 'Введите число'
          return
        }

        onSave(side, number)
      })
    })

    return button
  }

  function createSalesComparison(articleKey, currentValue, onManualValue) {
    if (!tableDataset.comparisonSalesByArticle) return null

    var comparison = document.createElement('span')
    comparison.className = 'sales-comparison'

    var baselineValue = manualPreviousSales.has(articleKey)
      ? manualPreviousSales.get(articleKey)
      : tableDataset.comparisonSalesByArticle.get(articleKey)

    if (typeof baselineValue !== 'number') {
      return createMissingSalesEditor('Нет в первом файле', 'previous', onManualValue)
    }

    if (typeof currentValue !== 'number') {
      return createMissingSalesEditor('Нет во втором файле', 'current', onManualValue)
    }

    var difference = currentValue - baselineValue
    var percentage = baselineValue === 0 ? null : difference / Math.abs(baselineValue) * 100
    var percentageLabel = percentage == null ? '—' : formatSignedNumber(percentage, 1) + '%'

    comparison.textContent = formatSignedNumber(difference, 2) + ' (' + percentageLabel + ')'
    comparison.classList.add(
      difference > 0
        ? 'sales-comparison--positive'
        : difference < 0
          ? 'sales-comparison--negative'
          : 'sales-comparison--neutral',
    )

    return comparison
  }

  function createProductRow(entry) {
    var tableRow = document.createElement('tr')
    var articleValue = String(entry.row[tableDataset.articleIndex] || '')
    var articleKey = normalizeArticle(articleValue)
    var actionCell = document.createElement('td')
    var removeButton = document.createElement('button')

    actionCell.className = 'row-action-cell'
    removeButton.className = 'remove-row'
    removeButton.type = 'button'
    removeButton.setAttribute('aria-label', 'Удалить артикул ' + articleValue)
    removeButton.innerHTML = '<span aria-hidden="true">−</span>'
    removeButton.addEventListener('click', function () {
      selectedArticleKeys.delete(articleKey)
      tableRow.remove()
      updateResultsMeta()
      persistSelectedArticles().catch(showPersistenceWarning)
    })
    actionCell.appendChild(removeButton)
    tableRow.appendChild(actionCell)

    tableDataset.columnIndexes.forEach(function (index, columnIndex) {
      var cell = document.createElement('td')
      var value = entry.row[index]

      if (columnIndex === 3 && manualCurrentSales.has(articleKey)) {
        value = manualCurrentSales.get(articleKey)
      }

      if (columnIndex === 0) {
        var imageUrl = tableDataset.imagesByRow.get(entry.rowIndex)
        cell.className = 'photo-cell'

        if (imageUrl) {
          var image = document.createElement('img')
          image.className = 'product-photo'
          image.src = imageUrl
          image.alt = 'Фото товара ' + String(entry.row[tableDataset.articleIndex] || '')
          image.loading = 'eager'
          cell.appendChild(image)
        } else {
          cell.textContent = '—'
        }
      } else {
        cell.textContent = formatCellValue(value, columnIndex)
      }

      tableRow.appendChild(cell)

      if (columnIndex === 3) {
        var changeCell = document.createElement('td')
        var comparison = createSalesComparison(articleKey, value, function (side, number) {
          if (side === 'previous') {
            manualPreviousSales.set(articleKey, number)
          } else {
            manualCurrentSales.set(articleKey, number)
          }

          tableRow.replaceWith(createProductRow(entry))
          persistManualSales().catch(showPersistenceWarning)
        })

        changeCell.className = 'sales-change-cell'
        if (comparison) {
          changeCell.appendChild(comparison)
        } else {
          changeCell.textContent = '—'
        }
        tableRow.appendChild(changeCell)
      }
    })

    return tableRow
  }

  function updateResultsMeta() {
    var count = selectedArticleKeys.size
    resultsMeta.textContent = count + ' ' + getRowWord(count) + ' · Retail'
    exportPdfButton.disabled = count === 0
    exportExcelButton.disabled = count === 0
  }

  function insertArticle(articleKey) {
    if (selectedArticleKeys.has(articleKey)) {
      return { ok: false, message: 'Этот артикул уже добавлен' }
    }

    var entry = tableDataset.rowsByArticle.get(articleKey)
    if (!entry) {
      return { ok: false, message: 'Артикул не найден на вкладке Retail' }
    }

    var addRow = tableBody.querySelector('.add-row')
    selectedArticleKeys.add(articleKey)
    tableBody.insertBefore(createProductRow(entry), addRow)
    updateResultsMeta()

    return { ok: true }
  }

  function showArticleEditor(cell, addRow) {
    var form = document.createElement('form')
    var controls = document.createElement('div')
    var input = document.createElement('input')
    var submitButton = document.createElement('button')
    var error = document.createElement('span')

    form.className = 'article-form'
    controls.className = 'article-form__controls'
    input.className = 'article-input'
    input.type = 'text'
    input.inputMode = 'text'
    input.autocomplete = 'off'
    input.placeholder = 'Введите артикул'
    input.setAttribute('aria-label', 'Артикул для добавления')
    submitButton.className = 'article-submit'
    submitButton.type = 'submit'
    submitButton.textContent = 'Добавить'
    error.className = 'article-error'
    error.setAttribute('aria-live', 'polite')

    controls.append(input, submitButton)
    form.append(controls, error)
    cell.replaceChildren(form)
    input.focus()

    form.addEventListener('submit', function (event) {
      event.preventDefault()
      var articleKey = normalizeArticle(input.value)

      if (!articleKey) {
        error.textContent = 'Введите артикул'
        return
      }

      var result = insertArticle(articleKey)
      if (!result.ok) {
        error.textContent = result.message
        return
      }

      addRow.replaceWith(createAddRow())
      persistSelectedArticles().catch(showPersistenceWarning)
    })
  }

  function createAddRow() {
    var addRow = document.createElement('tr')
    addRow.className = 'add-row'
    var actionCell = document.createElement('td')
    actionCell.className = 'row-action-cell'
    addRow.appendChild(actionCell)

    REQUIRED_COLUMNS.forEach(function (_, columnIndex) {
      var cell = document.createElement('td')

      if (columnIndex === 1) {
        var button = document.createElement('button')
        button.className = 'add-trigger'
        button.type = 'button'
        button.textContent = '+ Добавить'
        button.addEventListener('click', function () {
          showArticleEditor(cell, addRow)
        })
        cell.appendChild(button)
      }

      addRow.appendChild(cell)
      if (columnIndex === 3) addRow.appendChild(document.createElement('td'))
    })

    return addRow
  }

  function buildArticleData(rows, headerRowIndex) {
    var headerRow = rows[headerRowIndex]
    var columnIndexes = REQUIRED_COLUMNS.map(function (column) {
      return headerRow.map(normalizeHeader).indexOf(normalizeHeader(column))
    })
    var articleIndex = columnIndexes[1]
    var dataRows = rows.map(function (row, rowIndex) {
      return { row: row, rowIndex: rowIndex }
    }).slice(headerRowIndex + 1).filter(function (entry) {
      return normalizeArticle(entry.row[articleIndex]) !== ''
    })
    var rowsByArticle = new Map()

    dataRows.forEach(function (entry) {
      var articleKey = normalizeArticle(entry.row[articleIndex])
      if (!rowsByArticle.has(articleKey)) rowsByArticle.set(articleKey, entry)
    })

    return {
      articleIndex: articleIndex,
      columnIndexes: columnIndexes,
      rowsByArticle: rowsByArticle,
    }
  }

  function buildComparisonSales(rows, headerRowIndex) {
    var data = buildArticleData(rows, headerRowIndex)
    var salesIndex = data.columnIndexes[3]
    var salesByArticle = new Map()

    data.rowsByArticle.forEach(function (entry, articleKey) {
      salesByArticle.set(articleKey, entry.row[salesIndex])
    })

    return salesByArticle
  }

  function renderTable(rows, headerRowIndex, imagesByRow, comparisonSalesByArticle) {
    var primaryData = buildArticleData(rows, headerRowIndex)

    tableDataset = {
      articleIndex: primaryData.articleIndex,
      columnIndexes: primaryData.columnIndexes,
      comparisonSalesByArticle: comparisonSalesByArticle,
      imagesByRow: imagesByRow,
      rowsByArticle: primaryData.rowsByArticle,
    }
    selectedArticleKeys = new Set()

    var headingRow = document.createElement('tr')
    var actionHeading = document.createElement('th')
    actionHeading.className = 'row-action-heading'
    actionHeading.scope = 'col'
    actionHeading.setAttribute('aria-label', 'Действия со строкой')
    headingRow.appendChild(actionHeading)

    REQUIRED_COLUMNS.forEach(function (column) {
      var heading = document.createElement('th')
      heading.scope = 'col'
      heading.textContent = column
      headingRow.appendChild(heading)

      if (column === 'Продажи Розница, руб.') {
        var changeHeading = document.createElement('th')
        changeHeading.scope = 'col'
        changeHeading.textContent = 'Изменение продаж'
        headingRow.appendChild(changeHeading)
      }
    })
    tableHead.replaceChildren(headingRow)

    tableBody.replaceChildren(createAddRow())

    updateResultsMeta()
    results.hidden = false
    results.scrollIntoView({ behavior: 'smooth', block: 'start' })
  }

  function getRowWord(number) {
    var lastTwo = number % 100
    var last = number % 10

    if (lastTwo >= 11 && lastTwo <= 14) return 'строк'
    if (last === 1) return 'строка'
    if (last >= 2 && last <= 4) return 'строки'
    return 'строк'
  }

  function showError(message) {
    feedback.textContent = message
    feedback.classList.add('feedback--error')
    results.hidden = true
  }

  function waitForImages(container) {
    return Promise.all(Array.from(container.querySelectorAll('img')).map(function (image) {
      if (image.complete) return Promise.resolve()

      return new Promise(function (resolve) {
        image.addEventListener('load', resolve, { once: true })
        image.addEventListener('error', resolve, { once: true })
      })
    }))
  }

  function splitRowsIntoPdfPages(rows, headerHeight, maximumHeight) {
    var pages = []
    var currentPage = []
    var currentHeight = headerHeight

    rows.forEach(function (row) {
      var rowHeight = Math.max(48, row.getBoundingClientRect().height)

      if (currentPage.length > 0 && currentHeight + rowHeight > maximumHeight) {
        pages.push(currentPage)
        currentPage = []
        currentHeight = headerHeight
      }

      currentPage.push(row)
      currentHeight += rowHeight
    })

    if (currentPage.length > 0) pages.push(currentPage)
    return pages
  }

  function createPdfPageTable(sourceTable, rows, width) {
    var container = document.createElement('div')
    var table = sourceTable.cloneNode(false)
    var head = tableHead.cloneNode(true)
    var body = document.createElement('tbody')

    container.className = 'pdf-render-root'
    container.style.width = width + 'px'
    table.className = 'pdf-export-table'
    table.style.width = width + 'px'
    head.removeAttribute('id')
    table.appendChild(head)
    rows.forEach(function (row) {
      body.appendChild(row.cloneNode(true))
    })
    table.appendChild(body)
    container.appendChild(table)
    document.body.appendChild(container)

    return container
  }

  async function exportTableToPdf() {
    var JsPdf = window.jspdf && window.jspdf.jsPDF

    if (!JsPdf || typeof window.html2canvas !== 'function') {
      throw new Error('Не удалось загрузить модуль экспорта PDF.')
    }

    if (document.querySelector('.manual-sales-form')) {
      throw new Error('Сначала сохраните введённое значение продаж.')
    }

    var sourceTable = tableBody.closest('table')
    var rows = Array.from(tableBody.querySelectorAll('tr:not(.add-row)'))
    // Фиксированная ширина PDF: одинаковая с мобилы и десктопа, чтобы вся таблица помещалась.
    var tableWidth = 1560
    var pdf = new JsPdf({
      compress: true,
      format: 'a4',
      orientation: 'landscape',
      unit: 'mm',
    })
    var pageWidth = pdf.internal.pageSize.getWidth()
    var pageHeight = pdf.internal.pageSize.getHeight()
    var margin = 10
    var footerHeight = 8
    var contentWidth = pageWidth - margin * 2
    var contentHeight = pageHeight - margin * 2 - footerHeight
    var maximumTableHeight = contentHeight * tableWidth / contentWidth
    var headerHeight = tableHead.getBoundingClientRect().height
    var pages = splitRowsIntoPdfPages(rows, headerHeight, maximumTableHeight)

    for (var pageIndex = 0; pageIndex < pages.length; pageIndex += 1) {
      var renderContainer = createPdfPageTable(sourceTable, pages[pageIndex], tableWidth)

      try {
        await waitForImages(renderContainer)
        var canvas = await window.html2canvas(renderContainer, {
          backgroundColor: '#ffffff',
          logging: false,
          scale: 1.5,
          useCORS: true,
          width: tableWidth,
        })
        var renderedHeight = canvas.height * contentWidth / canvas.width
        var imageHeight = Math.min(contentHeight, renderedHeight)
        var imageWidth = canvas.width * imageHeight / canvas.height

        if (pageIndex > 0) pdf.addPage()
        pdf.addImage(
          canvas.toDataURL('image/png'),
          'PNG',
          margin + (contentWidth - imageWidth) / 2,
          margin,
          imageWidth,
          imageHeight,
          'table-page-' + pageIndex,
          'FAST',
        )
        pdf.setFont('helvetica', 'normal')
        pdf.setFontSize(9)
        pdf.setTextColor(102, 113, 127)
        pdf.text(String(pageIndex + 1), pageWidth / 2, pageHeight - 5, { align: 'center' })
      } finally {
        renderContainer.remove()
      }
    }

    pdf.save('retail-analytic.pdf')
  }

  // Колонки с числовыми данными — в Excel их выгружаем числами, а не текстом.
  var EXCEL_NUMERIC_HEADERS = {
    'Продажи Розница, руб.': true,
    'Продажи Розница, шт': true,
    'Остатки В РОЗНИЦЕ, шт': true,
    'Остатки (Арвато основной), шт': true,
  }

  function exportTableToExcel() {
    if (typeof XLSX === 'undefined') {
      throw new Error('Не удалось загрузить модуль экспорта Excel.')
    }

    if (document.querySelector('.manual-sales-form')) {
      throw new Error('Сначала сохраните введённое значение продаж.')
    }

    // Читаем уже собранную на экране таблицу, пропуская служебную колонку действий.
    var headers = Array.from(tableHead.querySelectorAll('th')).slice(1).map(function (cell) {
      return cell.textContent.trim()
    })

    var bodyRows = Array.from(tableBody.querySelectorAll('tr:not(.add-row)'))
    if (bodyRows.length === 0) {
      throw new Error('Добавьте хотя бы один артикул перед экспортом.')
    }

    var matrix = [headers]

    bodyRows.forEach(function (tableRow) {
      var record = Array.from(tableRow.children).slice(1).map(function (cell, columnIndex) {
        var text = cell.textContent.replace(/\s+/g, ' ').trim()
        if (text === '' || text === '—') return ''

        if (EXCEL_NUMERIC_HEADERS[headers[columnIndex]]) {
          var normalized = text.replace(/[\s ]/g, '').replace(',', '.')
          var number = Number(normalized)
          if (normalized !== '' && Number.isFinite(number)) return number
        }

        return text
      })

      matrix.push(record)
    })

    var worksheet = XLSX.utils.aoa_to_sheet(matrix)
    worksheet['!cols'] = headers.map(function (header) {
      return { wch: Math.min(40, Math.max(12, header.length + 2)) }
    })

    var workbook = XLSX.utils.book_new()
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Retail')
    XLSX.writeFile(workbook, 'retail-analytic.xlsx')
  }

  async function readRetailFile(file, label) {
    var buffer = await file.arrayBuffer()
    var workbook = XLSX.read(buffer, { type: 'array' })
    var retailSheetName = workbook.SheetNames.find(function (sheetName) {
      return normalizeHeader(sheetName) === 'retail'
    })
    var retailSheet = retailSheetName ? workbook.Sheets[retailSheetName] : null

    if (!retailSheet) {
      throw new Error('В ' + label + ' нет вкладки Retail.')
    }

    var rows = XLSX.utils.sheet_to_json(retailSheet, {
      header: 1,
      defval: '',
      raw: true,
      blankrows: true,
    })
    var headerAnalysis = analyzeHeaders(rows)

    if (headerAnalysis.missingColumns.length > 0) {
      showMissingColumns(headerAnalysis.missingColumns)
      throw new Error('На вкладке Retail в ' + label + ' не найдены все обязательные столбцы.')
    }

    return {
      buffer: buffer,
      headerRowIndex: headerAnalysis.rowIndex,
      rows: rows,
    }
  }

  async function processFile(options) {
    options = options || {}
    if (!selectedFile) return

    if (typeof XLSX === 'undefined') {
      showError('Не удалось загрузить модуль чтения Excel. Проверьте, что папка vendor находится рядом с index.html.')
      return
    }

    confirmButton.disabled = true
    confirmButton.textContent = 'Обрабатываем файл…'
    feedback.textContent = ''
    feedback.classList.remove('feedback--error')
    debugPanel.hidden = true
    debugList.replaceChildren()

    var articlesToRestore = Array.isArray(options.restoredArticles)
      ? options.restoredArticles
      : Array.from(selectedArticleKeys)

    if (options.restoredManualSales) {
      restoreManualSales(options.restoredManualSales)
    }

    try {
      var primaryData = await readRetailFile(selectedFile, 'основном файле')
      var comparisonData = comparisonFile
        ? await readRetailFile(comparisonFile, 'файле для сравнения')
        : null
      var comparisonSalesByArticle = comparisonData
        ? buildComparisonSales(primaryData.rows, primaryData.headerRowIndex)
        : null
      var displayedData = comparisonData || primaryData
      var displayedFile = comparisonData ? comparisonFile : selectedFile
      var imagesByRow = await extractRetailImages(displayedData.buffer, displayedFile.name)

      renderTable(
        displayedData.rows,
        displayedData.headerRowIndex,
        imagesByRow,
        comparisonSalesByArticle,
      )
      updateResultsTitle()

      articlesToRestore.forEach(function (articleKey) {
        insertArticle(normalizeArticle(articleKey))
      })

      if (options.persistWorkbook !== false) {
        try {
          await persistFile('workbook', primaryData.buffer, selectedFile)
          if (comparisonData) {
            await persistFile('comparisonWorkbook', comparisonData.buffer, comparisonFile)
          }
          await persistSelectedArticles()
          await persistManualSales()
        } catch (persistenceError) {
          showPersistenceWarning()
        }
      }
    } catch (error) {
      showError(error instanceof Error ? error.message : 'Не удалось прочитать Excel-файл.')
    } finally {
      confirmButton.disabled = false
      confirmButton.textContent = 'Подтвердить и показать таблицу'
    }
  }

  async function restoreSession() {
    try {
      var savedWorkbook = await readLocalState('workbook')
      if (!savedWorkbook || !savedWorkbook.buffer) return

      var savedState = await Promise.all([
        readLocalState('articles'),
        readLocalState('comparisonWorkbook'),
        readLocalState('manualSales'),
      ])
      var savedArticles = savedState[0]
      var savedComparisonWorkbook = savedState[1]
      var savedManualSales = savedState[2]
      var restoredFile = new File([savedWorkbook.buffer], savedWorkbook.name, {
        type: savedWorkbook.type,
        lastModified: savedWorkbook.lastModified,
      })

      showFile(restoredFile)
      description.textContent = formatFileSize(restoredFile.size) + ' · восстановлен локально'

      if (savedComparisonWorkbook && savedComparisonWorkbook.buffer) {
        var restoredComparisonFile = new File(
          [savedComparisonWorkbook.buffer],
          savedComparisonWorkbook.name,
          {
            type: savedComparisonWorkbook.type,
            lastModified: savedComparisonWorkbook.lastModified,
          },
        )
        showComparisonFile(restoredComparisonFile)
        comparisonDescription.textContent = formatFileSize(restoredComparisonFile.size) + ' · восстановлен локально'
      }

      processFile({
        persistWorkbook: false,
        restoredArticles: Array.isArray(savedArticles) ? savedArticles : [],
        restoredManualSales: savedManualSales,
      })
    } catch (error) {
      databasePromise = null
    }
  }

  input.addEventListener('change', function () {
    showFile(input.files && input.files[0])
  })

  comparisonInput.addEventListener('change', function () {
    showComparisonFile(comparisonInput.files && comparisonInput.files[0])
  })

  exportPdfButton.addEventListener('click', async function () {
    if (exportPdfButton.disabled) return

    var originalContents = exportPdfButton.innerHTML
    exportPdfButton.disabled = true
    exportPdfButton.textContent = 'Создаём PDF…'
    feedback.textContent = ''
    feedback.classList.remove('feedback--error')

    try {
      await exportTableToPdf()
    } catch (error) {
      feedback.textContent = error instanceof Error ? error.message : 'Не удалось создать PDF.'
      feedback.classList.add('feedback--error')
    } finally {
      exportPdfButton.innerHTML = originalContents
      exportPdfButton.disabled = selectedArticleKeys.size === 0
    }
  })

  exportExcelButton.addEventListener('click', function () {
    if (exportExcelButton.disabled) return

    var originalContents = exportExcelButton.innerHTML
    exportExcelButton.disabled = true
    exportExcelButton.textContent = 'Создаём Excel…'
    feedback.textContent = ''
    feedback.classList.remove('feedback--error')

    try {
      exportTableToExcel()
    } catch (error) {
      feedback.textContent = error instanceof Error ? error.message : 'Не удалось создать Excel.'
      feedback.classList.add('feedback--error')
    } finally {
      exportExcelButton.innerHTML = originalContents
      exportExcelButton.disabled = selectedArticleKeys.size === 0
    }
  })

  confirmButton.addEventListener('click', function () {
    processFile()
  })

  ;['dragenter', 'dragover'].forEach(function (eventName) {
    picker.addEventListener(eventName, function (event) {
      event.preventDefault()
      picker.classList.add('file-picker--dragging')
    })
  })

  ;['dragleave', 'drop'].forEach(function (eventName) {
    picker.addEventListener(eventName, function (event) {
      event.preventDefault()
      picker.classList.remove('file-picker--dragging')
    })
  })

  picker.addEventListener('drop', function (event) {
    showFile(event.dataTransfer && event.dataTransfer.files[0])
  })

  ;['dragenter', 'dragover'].forEach(function (eventName) {
    comparisonPicker.addEventListener(eventName, function (event) {
      event.preventDefault()
      comparisonPicker.classList.add('file-picker--dragging')
    })
  })

  ;['dragleave', 'drop'].forEach(function (eventName) {
    comparisonPicker.addEventListener(eventName, function (event) {
      event.preventDefault()
      comparisonPicker.classList.remove('file-picker--dragging')
    })
  })

  comparisonPicker.addEventListener('drop', function (event) {
    showComparisonFile(event.dataTransfer && event.dataTransfer.files[0])
  })

  restoreSession()
})()
