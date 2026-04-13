import { useDeferredValue, useEffect, useMemo, useRef, useState } from 'react'
import * as XLSX from 'xlsx'
import {
  Bar,
  BarChart,
  CartesianGrid,
  Cell,
  Pie,
  PieChart,
  ResponsiveContainer,
  Tooltip,
  XAxis,
  YAxis,
} from 'recharts'
import './App.css'

type DataRow = Record<string, string | number>
type SortDirection = 'asc' | 'desc'
type SortConfig = { column: string; direction: SortDirection } | null
type ParsedSheet = { detailRows: DataRow[]; summaryRows: DataRow[] }
type LowStockIssue = {
  productCode: string
  productName: string
  productLabel: string
  branch: string
  minStock: number
  currentStock: number
  deficit: number
}

const PIE_COLORS = ['#f97316', '#0ea5e9', '#22c55e', '#eab308', '#ef4444', '#8b5cf6']
const PAGE_SIZE = 200

const MAIN_COLUMNS = {
  code: 'รหัสสินค้า',
  name: 'ชื่อสินค้า',
  unit: 'หน่วย',
  minStock: 'จำนวนสต็อกสินค้าขั้นต่ำ',
  total: 'ยอดรวม',
}

const BRANCH_COLUMNS = [
  'Stock กลางเขาตาโล',
  'Stock กลางไร่วนาสินธุ์',
  'เขาตาโล',
  'ไร่วนาสินธุ์',
  'ชัยพรวิถี',
  'สถานที่ตัวอย่าง',
]

const HEADER_ALIAS: Record<string, string> = {
  รหัสสินค้า: MAIN_COLUMNS.code,
  ชื่อสินค้า: MAIN_COLUMNS.name,
  หน่วย: MAIN_COLUMNS.unit,
  จำนวนสต็อกสินค้าขั้นต่ำ: MAIN_COLUMNS.minStock,
  จำนวนสต็อกสินค้าขันต่ำ: MAIN_COLUMNS.minStock,
  ยอดรวม: MAIN_COLUMNS.total,
  stockกลางเขาตาโล: 'Stock กลางเขาตาโล',
  stockกลางไร่วนาสินธุ์: 'Stock กลางไร่วนาสินธุ์',
  เขาตาโล: 'เขาตาโล',
  ไร่วนาสินธุ์: 'ไร่วนาสินธุ์',
  ชัยพรวิถี: 'ชัยพรวิถี',
  สถานที่ตัวอย่าง: 'สถานที่ตัวอย่าง',
}

function normalizeText(value: unknown): string {
  return String(value ?? '')
    .trim()
    .replace(/\s+/g, ' ')
}

function normalizeHeaderKey(value: unknown): string {
  return normalizeText(value).replace(/\s+/g, '').toLowerCase()
}

function mapHeaderName(header: unknown, fallbackIndex: number): string {
  const raw = normalizeText(header)
  const normalizedKey = normalizeHeaderKey(raw)
  if (!raw) {
    return `Column ${fallbackIndex + 1}`
  }
  return HEADER_ALIAS[raw] ?? HEADER_ALIAS[normalizedKey] ?? raw
}

function parseFlexibleNumber(value: unknown): number | null {
  if (typeof value === 'number' && Number.isFinite(value)) {
    return value
  }

  if (typeof value !== 'string') {
    return null
  }

  const trimmed = value.trim()
  if (!trimmed) {
    return null
  }

  const cleaned = trimmed.replace(/,/g, '').replace(/%/g, '')
  const parsed = Number(cleaned)
  return Number.isFinite(parsed) ? parsed : null
}

function formatNumber(value: number): string {
  return new Intl.NumberFormat('th-TH', { maximumFractionDigits: 2 }).format(value)
}

function isAggregateItemRow(row: DataRow, headers: string[]): boolean {
  const fieldsToCheck = [
    row[MAIN_COLUMNS.code],
    row[MAIN_COLUMNS.name],
    row[MAIN_COLUMNS.total],
    row[headers[0]],
  ]
  const joined = fieldsToCheck.map((value) => normalizeText(value)).join(' ').toLowerCase()
  return /(ยอดรวม|รวมทั้งสิ้น|total|summary)/i.test(joined)
}

function downloadCsv(filename: string, rows: Record<string, string | number>[]) {
  if (rows.length === 0) {
    return
  }

  const headers = Object.keys(rows[0])
  const escapeCsv = (value: string | number) => `"${String(value).replace(/"/g, '""')}"`

  const csvContent = [
    headers.map((header) => escapeCsv(header)).join(','),
    ...rows.map((row) => headers.map((header) => escapeCsv(row[header] ?? '')).join(',')),
  ].join('\n')

  const blob = new Blob([`\ufeff${csvContent}`], { type: 'text/csv;charset=utf-8;' })
  const url = URL.createObjectURL(blob)
  const link = document.createElement('a')
  link.href = url
  link.download = filename
  link.click()
  URL.revokeObjectURL(url)
}

function downloadXlsx(filename: string, rows: Record<string, string | number>[], sheetName: string) {
  if (rows.length === 0) {
    return
  }
  const worksheet = XLSX.utils.json_to_sheet(rows)
  const workbook = XLSX.utils.book_new()
  XLSX.utils.book_append_sheet(workbook, worksheet, sheetName)
  XLSX.writeFile(workbook, filename)
}

function getHeadersFromRows(rows: DataRow[]): string[] {
  const headers = new Set<string>()
  for (const row of rows) {
    for (const key of Object.keys(row)) {
      if (key) {
        headers.add(key)
      }
    }
  }
  return [...headers]
}

function getSortableValue(row: DataRow, column: string): string | number {
  const raw = row[column]
  if (raw === null || typeof raw === 'undefined' || raw === '') {
    return ''
  }

  const asNumber = parseFlexibleNumber(raw)
  if (asNumber !== null) {
    return asNumber
  }

  return String(raw).toLowerCase()
}

function isLikelySummaryRow(row: DataRow, headers: string[]): boolean {
  const firstHeader = headers[0]
  const firstCell = normalizeText(row[firstHeader])
  const joined = headers.map((header) => normalizeText(row[header])).join(' ').toLowerCase()

  const summaryKeywords = /(ยอดรวม|รวมทั้งสิ้น|รวม|total|summary|วันที่|เวลา|พิมพ์|print)/i
  if (summaryKeywords.test(firstCell) || summaryKeywords.test(joined)) {
    return true
  }

  const numericCount = headers.reduce((count, header) => {
    return parseFlexibleNumber(row[header]) !== null ? count + 1 : count
  }, 0)

  const textCount = headers.reduce((count, header) => {
    const value = normalizeText(row[header])
    if (!value) {
      return count
    }
    return parseFlexibleNumber(value) === null ? count + 1 : count
  }, 0)

  return textCount <= 1 && numericCount >= Math.max(2, Math.floor(headers.length * 0.5))
}

function splitSummaryRows(rows: DataRow[], headers: string[]): ParsedSheet {
  const summaryRows: DataRow[] = []
  const detailRows = [...rows]

  while (detailRows.length > 0 && summaryRows.length < 2) {
    const lastRow = detailRows[detailRows.length - 1]
    if (!isLikelySummaryRow(lastRow, headers)) {
      break
    }
    summaryRows.unshift(lastRow)
    detailRows.pop()
  }

  return { detailRows, summaryRows }
}

function buildRowsFromSheet(worksheet: XLSX.WorkSheet): ParsedSheet {
  const matrix = XLSX.utils.sheet_to_json<unknown[]>(worksheet, {
    header: 1,
    raw: true,
    defval: '',
  })

  if (matrix.length === 0) {
    return { detailRows: [], summaryRows: [] }
  }

  const nonEmptyCounts: number[] = matrix.map((row) =>
    row.reduce<number>((count, cell) => {
      const value = normalizeText(cell)
      return value ? count + 1 : count
    }, 0),
  )

  let headerRowIndex = 0
  const scanLimit = Math.min(10, matrix.length)

  for (let index = 0; index < scanLimit; index += 1) {
    if (nonEmptyCounts[index] >= 2) {
      headerRowIndex = index
      break
    }
  }

  if (
    matrix.length > 1 &&
    nonEmptyCounts[0] <= 2 &&
    nonEmptyCounts[1] > nonEmptyCounts[0] &&
    nonEmptyCounts[1] >= 2
  ) {
    headerRowIndex = 1
  }

  const headerRow = matrix[headerRowIndex] ?? []
  const headers = headerRow.map((cell, index) => mapHeaderName(cell, index))

  const dataRows = matrix.slice(headerRowIndex + 1)
  const parsedRows: DataRow[] = []

  for (const row of dataRows) {
    const mapped: DataRow = {}
    let hasValue = false

    headers.forEach((header, index) => {
      const rawCell = row[index]
      const normalized = normalizeText(rawCell)
      mapped[header] = normalized
      if (normalized) {
        hasValue = true
      }
    })

    if (hasValue) {
      parsedRows.push(mapped)
    }
  }

  return splitSummaryRows(parsedRows, headers)
}

function getColumnStats(rows: DataRow[], headers: string[]) {
  const numericColumns = headers.filter((header) => {
    let valueCount = 0
    let numericCount = 0

    for (const row of rows) {
      const value = row[header]
      if (value === null || value === '' || typeof value === 'undefined') {
        continue
      }
      valueCount += 1
      if (parseFlexibleNumber(value) !== null) {
        numericCount += 1
      }
    }

    if (valueCount === 0) {
      return false
    }

    return numericCount / valueCount >= 0.7
  })

  const dimensionColumns = headers.filter((header) => !numericColumns.includes(header))
  return { numericColumns, dimensionColumns }
}

function App() {
  const columnPickerRef = useRef<HTMLDivElement | null>(null)
  const [isDarkMode, setIsDarkMode] = useState(() => {
    if (typeof window === 'undefined') {
      return false
    }

    const savedTheme = window.localStorage.getItem('info-cal-theme')
    if (savedTheme === 'dark') {
      return true
    }
    if (savedTheme === 'light') {
      return false
    }

    return window.matchMedia('(prefers-color-scheme: dark)').matches
  })
  const [sheetData, setSheetData] = useState<Record<string, ParsedSheet>>({})
  const [selectedSheet, setSelectedSheet] = useState('')
  const [selectedMetric, setSelectedMetric] = useState('')
  const [selectedDimension, setSelectedDimension] = useState('')
  const [searchText, setSearchText] = useState('')
  const [activeGroup, setActiveGroup] = useState('')
  const [sortConfig, setSortConfig] = useState<SortConfig>(null)
  const [fileInputKey, setFileInputKey] = useState(0)
  const [currentPage, setCurrentPage] = useState(1)
  const [showClearModal, setShowClearModal] = useState(false)
  const [selectedBranchForModal, setSelectedBranchForModal] = useState('')
  const [selectedChartGroupForModal, setSelectedChartGroupForModal] = useState('')
  const [chartModalSearchText, setChartModalSearchText] = useState('')
  const [chartModalSortConfig, setChartModalSortConfig] = useState<SortConfig>(null)
  const [isLowStockHighlightEnabled, setIsLowStockHighlightEnabled] = useState(false)
  const [visibleColumns, setVisibleColumns] = useState<string[]>([])
  const [showColumnPicker, setShowColumnPicker] = useState(false)
  const [modalSearchText, setModalSearchText] = useState('')
  const [modalSortConfig, setModalSortConfig] = useState<{ column: string; direction: SortDirection } | null>(
    { column: 'deficit', direction: 'desc' },
  )
  const [error, setError] = useState('')

  const deferredSearchText = useDeferredValue(searchText)
  const deferredModalSearchText = useDeferredValue(modalSearchText)
  const deferredChartModalSearchText = useDeferredValue(chartModalSearchText)

  const sheetNames = useMemo(() => Object.keys(sheetData), [sheetData])

  const rows = useMemo(() => {
    if (!selectedSheet || !sheetData[selectedSheet]) {
      return []
    }
    return sheetData[selectedSheet].detailRows
  }, [selectedSheet, sheetData])

  const summaryRows = useMemo(() => {
    if (!selectedSheet || !sheetData[selectedSheet]) {
      return []
    }
    return sheetData[selectedSheet].summaryRows
  }, [selectedSheet, sheetData])

  const headers = useMemo(() => getHeadersFromRows([...rows, ...summaryRows]), [rows, summaryRows])
  const displayedHeaders = useMemo(() => {
    const headerSet = new Set(visibleColumns)
    return headers.filter((header) => headerSet.has(header))
  }, [headers, visibleColumns])
  const allColumnsSelected = headers.length > 0 && displayedHeaders.length === headers.length
  const analysisRows = useMemo(
    () => rows.filter((row) => !isAggregateItemRow(row, headers)),
    [headers, rows],
  )

  const { numericColumns } = useMemo(() => getColumnStats(analysisRows, headers), [analysisRows, headers])

  const completionRate = useMemo(() => {
    if (rows.length === 0 || headers.length === 0) {
      return 0
    }

    const totalCells = rows.length * headers.length
    let filledCells = 0

    for (const row of rows) {
      for (const header of headers) {
        const value = row[header]
        if (value !== '' && value !== null && typeof value !== 'undefined') {
          filledCells += 1
        }
      }
    }

    return (filledCells / totalCells) * 100
  }, [analysisRows, headers])

  const chartData = useMemo(() => {
    if (!selectedMetric || !selectedDimension) {
      return []
    }

    const grouped = new Map<string, number>()

    for (const row of analysisRows) {
      const key = String(row[selectedDimension] ?? 'N/A').trim() || 'N/A'
      const amount = parseFlexibleNumber(row[selectedMetric])
      if (amount === null) {
        continue
      }
      grouped.set(key, (grouped.get(key) ?? 0) + amount)
    }

    return [...grouped.entries()]
      .map(([name, value]) => ({ name, value }))
      .sort((a, b) => b.value - a.value)
      .slice(0, 12)
  }, [rows, selectedDimension, selectedMetric])

  const groupFilteredRows = useMemo(() => {
    if (!activeGroup || !selectedDimension) {
      return analysisRows
    }

    return analysisRows.filter((row) => {
      const groupValue = String(row[selectedDimension] ?? 'N/A').trim() || 'N/A'
      return groupValue === activeGroup
    })
  }, [activeGroup, analysisRows, selectedDimension])

  const searchedRows = useMemo(() => {
    if (!deferredSearchText.trim()) {
      return groupFilteredRows
    }
    const keyword = deferredSearchText.toLowerCase()
    return groupFilteredRows.filter((row) =>
      headers.some((header) => String(row[header] ?? '').toLowerCase().includes(keyword)),
    )
  }, [deferredSearchText, groupFilteredRows, headers])

  const sortedRows = useMemo(() => {
    const nextRows = [...searchedRows]
    if (!sortConfig) {
      return nextRows
    }

    nextRows.sort((a, b) => {
      const aValue = getSortableValue(a, sortConfig.column)
      const bValue = getSortableValue(b, sortConfig.column)

      if (aValue === bValue) {
        return 0
      }

      const result = aValue > bValue ? 1 : -1
      return sortConfig.direction === 'asc' ? result : -result
    })

    return nextRows
  }, [searchedRows, sortConfig])

  const totalPages = useMemo(() => Math.max(1, Math.ceil(sortedRows.length / PAGE_SIZE)), [sortedRows])

  useEffect(() => {
    if (currentPage > totalPages) {
      setCurrentPage(totalPages)
    }
  }, [currentPage, totalPages])

  useEffect(() => {
    setCurrentPage(1)
  }, [selectedSheet, selectedMetric, selectedDimension, searchText, activeGroup, sortConfig])

  const pagedRows = useMemo(() => {
    const start = (currentPage - 1) * PAGE_SIZE
    return sortedRows.slice(start, start + PAGE_SIZE)
  }, [currentPage, sortedRows])

  const branchColumnsInFile = useMemo(
    () => BRANCH_COLUMNS.filter((branchName) => headers.includes(branchName)),
    [headers],
  )

  const lowStockIssues = useMemo(() => {
    const productNameCol = headers.includes(MAIN_COLUMNS.name) ? MAIN_COLUMNS.name : headers[0]
    const productCodeCol = headers.includes(MAIN_COLUMNS.code) ? MAIN_COLUMNS.code : ''

    const issues: LowStockIssue[] = []
    if (!headers.includes(MAIN_COLUMNS.minStock)) {
      return issues
    }

    for (const row of analysisRows) {
      const minStock = parseFlexibleNumber(row[MAIN_COLUMNS.minStock])
      if (minStock === null) {
        continue
      }

      const productName = normalizeText(row[productNameCol]) || 'ไม่ระบุสินค้า'
      const productCode = productCodeCol ? normalizeText(row[productCodeCol]) : ''

      for (const branchName of branchColumnsInFile) {
        const currentStock = parseFlexibleNumber(row[branchName])
        if (currentStock === null) {
          continue
        }
        if (currentStock < minStock) {
          const productLabel = productCode ? `${productCode} - ${productName}` : productName
          issues.push({
            productCode,
            productName,
            productLabel,
            branch: branchName,
            minStock,
            currentStock,
            deficit: minStock - currentStock,
          })
        }
      }
    }

    return issues
  }, [analysisRows, branchColumnsInFile, headers])

  const branchIssueCounts = useMemo(() => {
    const counts = new Map<string, number>()
    for (const item of lowStockIssues) {
      counts.set(item.branch, (counts.get(item.branch) ?? 0) + 1)
    }
    return counts
  }, [lowStockIssues])

  const lowStockByBranch = useMemo(() => {
    if (!headers.includes(MAIN_COLUMNS.minStock)) {
      return []
    }

    return branchColumnsInFile
      .map((branchName) => {
        return { branch: branchName, shortageItems: branchIssueCounts.get(branchName) ?? 0 }
      })
      .sort((a, b) => b.shortageItems - a.shortageItems)
  }, [branchColumnsInFile, branchIssueCounts, headers])

  const selectedBranchIssues = useMemo(() => {
    if (!selectedBranchForModal) {
      return []
    }
    return lowStockIssues
      .filter((item) => item.branch === selectedBranchForModal)
      .sort((a, b) => b.deficit - a.deficit)
  }, [lowStockIssues, selectedBranchForModal])

  const chartGroupRows = useMemo(() => {
    if (!selectedChartGroupForModal || !selectedDimension) {
      return []
    }

    return analysisRows.filter((row) => {
      const groupValue = String(row[selectedDimension] ?? 'N/A').trim() || 'N/A'
      return groupValue === selectedChartGroupForModal
    })
  }, [analysisRows, selectedChartGroupForModal, selectedDimension])

  const visibleChartGroupRows = useMemo(() => {
    let nextRows = chartGroupRows

    if (deferredChartModalSearchText.trim()) {
      const keyword = deferredChartModalSearchText.toLowerCase()
      nextRows = nextRows.filter((row) =>
        headers.some((header) => String(row[header] ?? '').toLowerCase().includes(keyword)),
      )
    }

    if (!chartModalSortConfig) {
      return nextRows
    }

    const sortable = [...nextRows]
    sortable.sort((a, b) => {
      const aValue = getSortableValue(a, chartModalSortConfig.column)
      const bValue = getSortableValue(b, chartModalSortConfig.column)

      if (aValue === bValue) {
        return 0
      }

      const result = aValue > bValue ? 1 : -1
      return chartModalSortConfig.direction === 'asc' ? result : -result
    })

    return sortable
  }, [chartGroupRows, deferredChartModalSearchText, chartModalSortConfig, headers])

  const visibleModalIssues = useMemo(() => {
    let nextRows = selectedBranchIssues

    if (deferredModalSearchText.trim()) {
      const keyword = deferredModalSearchText.toLowerCase()
      nextRows = nextRows.filter((item) => {
        return (
          item.productCode.toLowerCase().includes(keyword) ||
          item.productName.toLowerCase().includes(keyword) ||
          item.productLabel.toLowerCase().includes(keyword)
        )
      })
    }

    if (!modalSortConfig) {
      return nextRows
    }

    const sortable = [...nextRows]
    sortable.sort((a, b) => {
      const aValue = a[modalSortConfig.column as keyof LowStockIssue]
      const bValue = b[modalSortConfig.column as keyof LowStockIssue]

      if (aValue === bValue) {
        return 0
      }

      const result = aValue > bValue ? 1 : -1
      return modalSortConfig.direction === 'asc' ? result : -result
    })

    return sortable
  }, [deferredModalSearchText, modalSortConfig, selectedBranchIssues])

  useEffect(() => {
    if (!selectedBranchForModal) {
      return
    }
    setModalSearchText('')
    setModalSortConfig({ column: 'deficit', direction: 'desc' })
  }, [selectedBranchForModal])

  useEffect(() => {
    if (!selectedChartGroupForModal) {
      return
    }
    setChartModalSearchText('')
    setChartModalSortConfig(null)
  }, [selectedChartGroupForModal])

  useEffect(() => {
    setVisibleColumns(headers)
    setShowColumnPicker(false)
  }, [headers])

  useEffect(() => {
    document.documentElement.dataset.theme = isDarkMode ? 'dark' : 'light'
    window.localStorage.setItem('info-cal-theme', isDarkMode ? 'dark' : 'light')
  }, [isDarkMode])

  useEffect(() => {
    function handlePointerOutside(event: MouseEvent | TouchEvent) {
      if (!showColumnPicker || !columnPickerRef.current) {
        return
      }

      const target = event.target
      if (target instanceof Node && !columnPickerRef.current.contains(target)) {
        setShowColumnPicker(false)
      }
    }

    document.addEventListener('mousedown', handlePointerOutside)
    document.addEventListener('touchstart', handlePointerOutside)

    return () => {
      document.removeEventListener('mousedown', handlePointerOutside)
      document.removeEventListener('touchstart', handlePointerOutside)
    }
  }, [showColumnPicker])

  function exportBranchIssues(branch: string, fileType: 'csv' | 'xlsx') {
    const rowsForExport = lowStockIssues
      .filter((item) => item.branch === branch)
      .map((item) => ({
        สาขา: item.branch,
        รหัสสินค้า: item.productCode,
        ชื่อสินค้า: item.productName,
        ขั้นต่ำ: item.minStock,
        คงเหลือ: item.currentStock,
        ขาด: item.deficit,
      }))

    if (fileType === 'csv') {
      downloadCsv(`low-stock-${branch}.csv`, rowsForExport)
      return
    }
    downloadXlsx(`low-stock-${branch}.xlsx`, rowsForExport, 'LowStock')
  }

  function exportAllBranchIssues(fileType: 'csv' | 'xlsx') {
    const rowsForExport = lowStockIssues.map((item) => ({
      สาขา: item.branch,
      รหัสสินค้า: item.productCode,
      ชื่อสินค้า: item.productName,
      ขั้นต่ำ: item.minStock,
      คงเหลือ: item.currentStock,
      ขาด: item.deficit,
    }))

    if (fileType === 'csv') {
      downloadCsv('low-stock-all-branches.csv', rowsForExport)
      return
    }

    const workbook = XLSX.utils.book_new()

    for (const branchName of branchColumnsInFile) {
      const sheetRows = lowStockIssues
        .filter((item) => item.branch === branchName)
        .map((item) => ({
          รหัสสินค้า: item.productCode,
          ชื่อสินค้า: item.productName,
          ขั้นต่ำ: item.minStock,
          คงเหลือ: item.currentStock,
          ขาด: item.deficit,
        }))

      const rowsWithFallback =
        sheetRows.length > 0 ? sheetRows : [{ รหัสสินค้า: '-', ชื่อสินค้า: 'ไม่มีรายการต้องเบิก', ขั้นต่ำ: '-', คงเหลือ: '-', ขาด: '-' }]

      const safeSheetName = branchName.replace(/[\\/?*\[\]:]/g, '').slice(0, 31) || 'Branch'
      const worksheet = XLSX.utils.json_to_sheet(rowsWithFallback)
      XLSX.utils.book_append_sheet(workbook, worksheet, safeSheetName)
    }

    if (workbook.SheetNames.length === 0) {
      const worksheet = XLSX.utils.json_to_sheet([
        { หมายเหตุ: 'ไม่พบคอลัมน์สาขาในไฟล์ หรือไม่มีข้อมูลสำหรับส่งออก' },
      ])
      XLSX.utils.book_append_sheet(workbook, worksheet, 'Info')
    }

    XLSX.writeFile(workbook, 'low-stock-all-branches.xlsx')
  }

  function exportChartGroupRows(fileType: 'csv' | 'xlsx') {
    if (!selectedChartGroupForModal) {
      return
    }

    const rowsForExport = visibleChartGroupRows.map((row) => {
      const normalized: Record<string, string | number> = {}
      for (const header of headers) {
        normalized[header] = row[header] ?? ''
      }
      return normalized
    })

    if (fileType === 'csv') {
      downloadCsv(`group-${selectedChartGroupForModal}.csv`, rowsForExport)
      return
    }
    downloadXlsx(`group-${selectedChartGroupForModal}.xlsx`, rowsForExport, 'GroupData')
  }

  async function handleFileChange(event: React.ChangeEvent<HTMLInputElement>) {
    const file = event.target.files?.[0]
    if (!file) {
      return
    }

    try {
      setError('')
      const arrayBuffer = await file.arrayBuffer()
      const workbook = XLSX.read(arrayBuffer, { type: 'array' })

      const nextSheetRows: Record<string, ParsedSheet> = {}

      for (const sheetName of workbook.SheetNames) {
        const worksheet = workbook.Sheets[sheetName]
        nextSheetRows[sheetName] = buildRowsFromSheet(worksheet)
      }

      const firstSheet = workbook.SheetNames[0] ?? ''
      const firstDetailRows = nextSheetRows[firstSheet]?.detailRows ?? []
      const firstSummaryRows = nextSheetRows[firstSheet]?.summaryRows ?? []
      const firstHeaders = getHeadersFromRows([...firstDetailRows, ...firstSummaryRows])
      const firstCleanRows = firstDetailRows.filter((row) => !isAggregateItemRow(row, firstHeaders))
      const firstStats = getColumnStats(firstCleanRows, firstHeaders)

      setSheetData(nextSheetRows)
      setSelectedSheet(firstSheet)
      setSelectedMetric(
        firstHeaders.includes(MAIN_COLUMNS.total)
          ? MAIN_COLUMNS.total
          : (firstStats.numericColumns[0] ?? ''),
      )
      setSelectedDimension(
        firstHeaders.includes(MAIN_COLUMNS.code)
          ? MAIN_COLUMNS.code
          : (firstStats.dimensionColumns[0] ?? MAIN_COLUMNS.name),
      )
      setSearchText('')
      setActiveGroup('')
      setSortConfig(null)
      setCurrentPage(1)
      setSelectedBranchForModal('')
    } catch {
      setError('ไม่สามารถอ่านไฟล์ได้ กรุณาตรวจสอบว่าเป็นไฟล์ Excel ที่มี header แถวแรก')
      setSheetData({})
      setSelectedSheet('')
      setSelectedMetric('')
      setSelectedDimension('')
      setSearchText('')
      setActiveGroup('')
      setSortConfig(null)
      setCurrentPage(1)
      setSelectedBranchForModal('')
    }
  }

  function handleSheetChange(sheetName: string) {
    const nextRows = sheetData[sheetName]?.detailRows ?? []
    const nextSummaryRows = sheetData[sheetName]?.summaryRows ?? []
    const nextHeaders = getHeadersFromRows([...nextRows, ...nextSummaryRows])
    const nextCleanRows = nextRows.filter((row) => !isAggregateItemRow(row, nextHeaders))
    const nextStats = getColumnStats(nextCleanRows, nextHeaders)

    setSelectedSheet(sheetName)
    setSelectedMetric(
      nextHeaders.includes(MAIN_COLUMNS.total) ? MAIN_COLUMNS.total : (nextStats.numericColumns[0] ?? ''),
    )
    setSelectedDimension(
      nextHeaders.includes(MAIN_COLUMNS.code)
        ? MAIN_COLUMNS.code
        : (nextStats.dimensionColumns[0] ?? MAIN_COLUMNS.name),
    )
    setSearchText('')
    setActiveGroup('')
    setSortConfig(null)
    setCurrentPage(1)
    setSelectedBranchForModal('')
  }

  function clearAllDataConfirmed() {
    setSheetData({})
    setSelectedSheet('')
    setSelectedMetric('')
    setSelectedDimension('')
    setSearchText('')
    setActiveGroup('')
    setSortConfig(null)
    setCurrentPage(1)
    setError('')
    setShowClearModal(false)
    setSelectedBranchForModal('')
    setFileInputKey((current) => current + 1)
  }

  function handleSort(column: string) {
    setSortConfig((current) => {
      if (!current || current.column !== column) {
        return { column, direction: 'asc' }
      }

      if (current.direction === 'asc') {
        return { column, direction: 'desc' }
      }

      return null
    })
  }

  function handleModalSort(column: string) {
    setModalSortConfig((current) => {
      if (!current || current.column !== column) {
        return { column, direction: 'asc' }
      }

      if (current.direction === 'asc') {
        return { column, direction: 'desc' }
      }

      return null
    })
  }

  function handleChartModalSort(column: string) {
    setChartModalSortConfig((current) => {
      if (!current || current.column !== column) {
        return { column, direction: 'asc' }
      }

      if (current.direction === 'asc') {
        return { column, direction: 'desc' }
      }

      return null
    })
  }

  function openChartGroupModal(groupName: string) {
    if (!groupName) {
      return
    }
    setSelectedChartGroupForModal(groupName)
  }

  function toggleVisibleColumn(column: string) {
    setVisibleColumns((current) => {
      if (current.includes(column)) {
        return current.filter((item) => item !== column)
      }
      return [...current, column]
    })
  }

  function toggleAllVisibleColumns() {
    setVisibleColumns((current) => (current.length === headers.length ? [] : headers))
  }

  function isRowBelowMinimum(row: DataRow): boolean {
    const minStock = parseFlexibleNumber(row[MAIN_COLUMNS.minStock])
    if (minStock === null) {
      return false
    }

    for (const branchName of branchColumnsInFile) {
      const currentStock = parseFlexibleNumber(row[branchName])
      if (currentStock !== null && currentStock < minStock) {
        return true
      }
    }

    return false
  }

  return (
    <main className="dashboard-app">
      <nav className="top-navbar" aria-label="แถบนำทางหลัก">
        <div className="navbar-brand">Excel Insight</div>
        <div className="navbar-actions">
          <div className="navbar-page-title">แดชบอร์ด</div>
          <button
            type="button"
            className="theme-toggle"
            onClick={() => setIsDarkMode((current) => !current)}
            aria-label={isDarkMode ? 'สลับเป็นโหมดสว่าง' : 'สลับเป็นโหมดมืด'}
            title={isDarkMode ? 'สลับเป็นโหมดสว่าง' : 'สลับเป็นโหมดมืด'}
          >
            <span className="theme-toggle-icon" aria-hidden="true">
              {isDarkMode ? '☾' : '☀'}
            </span>
          </button>
        </div>
      </nav>

      <section className="card controls">
        <div className="control-row">
          <label className="field">
            <span>เลือกไฟล์ Excel</span>
            <input
              key={fileInputKey}
              type="file"
              accept=".xlsx,.xls,.csv"
              onChange={handleFileChange}
            />
          </label>

          <label className="field">
            <span>ชีต</span>
            <select
              value={selectedSheet}
              onChange={(event) => handleSheetChange(event.target.value)}
              disabled={sheetNames.length === 0}
            >
              {sheetNames.map((name) => (
                <option key={name} value={name}>
                  {name}
                </option>
              ))}
            </select>
          </label>

          <label className="field">
            <span>เรียงจาก (ใช้จัดกลุ่ม)</span>
            <select
              value={selectedDimension}
              onChange={(event) => {
                setSelectedDimension(event.target.value)
                setActiveGroup('')
              }}
              disabled={headers.length === 0}
            >
              {headers.map((header) => (
                <option key={header} value={header}>
                  {header}
                </option>
              ))}
            </select>
          </label>

          <label className="field">
            <span>ตัวชี้วัด</span>
            <select
              value={selectedMetric}
              onChange={(event) => {
                setSelectedMetric(event.target.value)
                setActiveGroup('')
              }}
              disabled={numericColumns.length === 0}
            >
              {numericColumns.map((header) => (
                <option key={header} value={header}>
                  {header}
                </option>
              ))}
            </select>
          </label>
        </div>

        <div className="action-row">
          <button type="button" className="danger-btn" onClick={() => setShowClearModal(true)}>
            ล้างข้อมูลทั้งหมด
          </button>
        </div>

        {error && <p className="error">{error}</p>}
      </section>

      <section className="kpi-grid">
        <article className="card kpi">
          <h2>จำนวนแถวข้อมูล</h2>
          <p>{formatNumber(rows.length)}</p>
        </article>
        <article className="card kpi">
          <h2>จำนวนแถวสรุป</h2>
          <p>{formatNumber(summaryRows.length)}</p>
        </article>
        <article className="card kpi">
          <h2>จำนวนสาขาที่ต้องเบิก</h2>
          <p>{formatNumber(lowStockByBranch.filter((item) => item.shortageItems > 0).length)}</p>
        </article>
        <article className="card kpi">
          <h2>ความสมบูรณ์ของข้อมูล</h2>
          <p>{completionRate.toFixed(1)}%</p>
        </article>
      </section>

      <section className="card low-stock-section">
        <div className="insight-head">
          <h2>สาขาที่ต้องเบิกสินค้า (Stock ต่ำกว่า Stock ขั้นต่ำ)</h2>
          <div className="export-actions">
            <button type="button" className="ghost-btn" onClick={() => exportAllBranchIssues('csv')}>
              Export CSV ทุกสาขา
            </button>
            <button type="button" className="ghost-btn" onClick={() => exportAllBranchIssues('xlsx')}>
              Export Excel ทุกสาขา
            </button>
          </div>
        </div>

        <div className="insight-grid">
          {lowStockByBranch.map((item) => (
            <button
              type="button"
              className={`insight-card ${item.shortageItems > 0 ? 'active' : ''}`}
              key={item.branch}
              onClick={() => setSelectedBranchForModal(item.branch)}
            >
              <span className="label">{item.branch}</span>
              <strong>{formatNumber(item.shortageItems)} รายการ</strong>
            </button>
          ))}
          {lowStockByBranch.length === 0 && (
            <p className="empty-inline">ไม่พบคอลัมน์สาขาในไฟล์หรือไม่มีข้อมูลขั้นต่ำ</p>
          )}
        </div>

      </section>

      <section className="chart-grid">
        <article className="card chart-card">
          <h2>กราฟแท่ง: ผลรวม {selectedMetric || 'ตัวชี้วัด'} แยกตาม {selectedDimension || 'มิติ'}</h2>
          <p className="chart-note">แสดงเฉพาะ 12 กลุ่มที่มีค่ามากที่สุด และไม่รวมแถวประเภท "ยอดรวม"</p>
          <p className="chart-click-hint">คลิกแท่งกราฟเพื่อดูรายละเอียดรายการในกลุ่มนั้น</p>
          {chartData.length > 0 ? (
            <ResponsiveContainer width="100%" height={320}>
              <BarChart data={chartData}>
                <CartesianGrid strokeDasharray="3 3" stroke="#d5dbdb" />
                <XAxis dataKey="name" tick={{ fontSize: 12 }} interval={0} angle={-20} height={70} />
                <YAxis tickFormatter={(value) => formatNumber(Number(value))} />
                <Tooltip formatter={(value) => formatNumber(Number(value ?? 0))} />
                <Bar
                  dataKey="value"
                  fill="#0ea5e9"
                  radius={[8, 8, 0, 0]}
                  cursor="pointer"
                  isAnimationActive={false}
                  onClick={(data: { name?: string }) => openChartGroupModal(String(data?.name ?? ''))}
                />
              </BarChart>
            </ResponsiveContainer>
          ) : (
            <p className="empty">อัปโหลดไฟล์และเลือกคอลัมน์เพื่อดูกราฟ</p>
          )}
        </article>

        <article className="card chart-card">
          <h2>กราฟวงกลม: สัดส่วน {selectedMetric || 'ตัวชี้วัด'} ของ 6 กลุ่มแรก</h2>
          <p className="chart-note">ใช้ข้อมูลชุดเดียวกับกราฟแท่ง เพื่อดูสัดส่วนภาพรวมอย่างเร็ว</p>
          <p className="chart-click-hint">คลิกชิ้นกราฟเพื่อดูรายละเอียดรายการในกลุ่มนั้น</p>
          {chartData.length > 0 ? (
            <ResponsiveContainer width="100%" height={320}>
              <PieChart>
                <Pie
                  data={chartData.slice(0, 6)}
                  dataKey="value"
                  nameKey="name"
                  outerRadius={120}
                  innerRadius={65}
                  cursor="pointer"
                  isAnimationActive={false}
                  onClick={(entry: { name?: string }) => openChartGroupModal(String(entry?.name ?? ''))}
                >
                  {chartData.slice(0, 6).map((entry, index) => (
                    <Cell key={entry.name} fill={PIE_COLORS[index % PIE_COLORS.length]} />
                  ))}
                </Pie>
                <Tooltip formatter={(value) => formatNumber(Number(value ?? 0))} />
              </PieChart>
            </ResponsiveContainer>
          ) : (
            <p className="empty">ยังไม่มีข้อมูลสำหรับสรุปสัดส่วน</p>
          )}
        </article>
      </section>

      <section className="card insight-section">
        <div className="insight-head">
          <h2>คลิกที่การ์ดเพื่อดูข้อมูลเฉพาะกลุ่มตามคอลัมน์มิติ ({selectedDimension || '-'})</h2>
          {activeGroup && (
            <button type="button" className="ghost-btn" onClick={() => setActiveGroup('')}>
              ล้างตัวกรองกลุ่ม
            </button>
          )}
        </div>

        <div className="insight-grid">
          {chartData.slice(0, 8).map((item) => (
            <button
              key={item.name}
              type="button"
              className={`insight-card ${activeGroup === item.name ? 'active' : ''}`}
              onClick={() => setActiveGroup(item.name)}
            >
              <span className="label">{item.name}</span>
              <strong>{formatNumber(item.value)}</strong>
            </button>
          ))}
          {chartData.length === 0 && <p className="empty-inline">ยังไม่มีข้อมูลกลุ่มให้เลือก</p>}
        </div>
      </section>

      <section className="card table-card">
        <div className="table-toolbar">
          <h2>
            ตัวอย่างข้อมูล ({PAGE_SIZE} แถว/หน้า)
            {activeGroup ? ` - กลุ่ม: ${activeGroup}` : ''}
          </h2>
          <div className="table-tools">
            <div className="column-picker-wrap" ref={columnPickerRef}>
              <button
                type="button"
                className="ghost-btn column-picker-trigger"
                onClick={() => setShowColumnPicker((current) => !current)}
              >
                เลือกคอลัมน์ ({displayedHeaders.length}/{headers.length})
              </button>
              {showColumnPicker && (
                <div className="column-picker-panel">
                  <div className="column-picker-actions">
                    <button type="button" className="ghost-btn compact-btn" onClick={toggleAllVisibleColumns}>
                      {allColumnsSelected ? 'ล้างทั้งหมด' : 'เลือกทั้งหมด'}
                    </button>
                  </div>
                  <div className="column-picker-list">
                    {headers.map((header) => (
                      <label key={header} className="column-picker-item">
                        <input
                          type="checkbox"
                          checked={displayedHeaders.includes(header)}
                          onChange={() => toggleVisibleColumn(header)}
                        />
                        <span>{header}</span>
                      </label>
                    ))}
                  </div>
                </div>
              )}
            </div>
            <label className="highlight-toggle">
              <input
                type="checkbox"
                checked={isLowStockHighlightEnabled}
                onChange={(event) => setIsLowStockHighlightEnabled(event.target.checked)}
              />
              <span>ไฮไลต์รายการที่ต้องเบิก</span>
            </label>
            <input
              type="search"
              placeholder="ค้นหาในตาราง"
              value={searchText}
              onChange={(event) => setSearchText(event.target.value)}
              disabled={rows.length === 0}
            />
          </div>
        </div>

        <div className="pager-row">
          <button
            type="button"
            className="ghost-btn"
            onClick={() => setCurrentPage((value) => Math.max(1, value - 1))}
            disabled={currentPage <= 1}
          >
            ก่อนหน้า
          </button>
          <span className="page-pill">
            หน้า {currentPage} / {totalPages}
          </span>
          <button
            type="button"
            className="ghost-btn"
            onClick={() => setCurrentPage((value) => Math.min(totalPages, value + 1))}
            disabled={currentPage >= totalPages}
          >
            ถัดไป
          </button>
        </div>

        <div className="table-scroll">
          {displayedHeaders.length === 0 ? (
            <div className="empty-table-state">ยังไม่ได้เลือกคอลัมน์ที่จะแสดง</div>
          ) : (
            <table style={{ minWidth: `${Math.max(760, displayedHeaders.length * 160)}px` }}>
              <thead>
                <tr>
                  {displayedHeaders.map((header) => (
                    <th key={header}>
                      <button type="button" className="sort-btn" onClick={() => handleSort(header)}>
                        {header}
                        {sortConfig?.column === header
                          ? sortConfig.direction === 'asc'
                            ? ' ▲'
                            : ' ▼'
                          : ''}
                      </button>
                    </th>
                  ))}
                </tr>
              </thead>
              <tbody>
                {pagedRows.map((row, index) => (
                  <tr
                    key={`row-${(currentPage - 1) * PAGE_SIZE + index}`}
                    className={
                      isLowStockHighlightEnabled && isRowBelowMinimum(row) ? 'row-low-stock' : undefined
                    }
                  >
                    {displayedHeaders.map((header) => (
                      <td key={`${index}-${header}`}>{String(row[header] ?? '')}</td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          )}
        </div>

        {summaryRows.length > 0 && displayedHeaders.length > 0 && (
          <div className="summary-panel">
            <h3>สรุปท้ายไฟล์ (แสดงทุกหน้าเสมอ)</h3>
            <div className="table-scroll">
              <table className="mini-table" style={{ minWidth: `${Math.max(760, displayedHeaders.length * 160)}px` }}>
                <tbody>
                  {summaryRows.map((row, index) => (
                    <tr key={`summary-${index}`}>
                      {displayedHeaders.map((header) => (
                        <td key={`summary-${index}-${header}`}>{String(row[header] ?? '')}</td>
                      ))}
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          </div>
        )}
      </section>

      {showClearModal && (
        <div className="modal-backdrop" role="presentation" onClick={() => setShowClearModal(false)}>
          <div className="modal-card" role="dialog" aria-modal="true" onClick={(event) => event.stopPropagation()}>
            <h3>ยืนยันการล้างข้อมูล</h3>
            <p>ต้องการล้างข้อมูลที่อัปโหลดทั้งหมดหรือไม่?</p>
            <div className="modal-actions">
              <button type="button" className="ghost-btn" onClick={() => setShowClearModal(false)}>
                ยกเลิก
              </button>
              <button type="button" className="danger-btn" onClick={clearAllDataConfirmed}>
                ยืนยันล้างข้อมูล
              </button>
            </div>
          </div>
        </div>
      )}

      {selectedBranchForModal && (
        <div className="modal-backdrop" role="presentation" onClick={() => setSelectedBranchForModal('')}>
          <div
            className="modal-card modal-wide"
            role="dialog"
            aria-modal="true"
            onClick={(event) => event.stopPropagation()}
          >
            <h3>รายการที่ต้องเบิก - {selectedBranchForModal}</h3>
            <p>แสดงสินค้าที่คงเหลือน้อยกว่าค่าขั้นต่ำในสาขานี้</p>
            <div className="modal-toolbar">
              <input
                type="search"
                className="modal-search-input"
                placeholder="ค้นหารหัสสินค้า/ชื่อสินค้า"
                value={modalSearchText}
                onChange={(event) => setModalSearchText(event.target.value)}
              />
            </div>
            <div className="modal-actions modal-actions-left">
              <button
                type="button"
                className="ghost-btn"
                onClick={() => exportBranchIssues(selectedBranchForModal, 'csv')}
              >
                Export CSV
              </button>
              <button
                type="button"
                className="ghost-btn"
                onClick={() => exportBranchIssues(selectedBranchForModal, 'xlsx')}
              >
                Export Excel
              </button>
            </div>
            <div className="table-scroll modal-table-scroll">
              <table className="mini-table">
                <thead>
                  <tr>
                    <th>
                      <button type="button" className="sort-btn" onClick={() => handleModalSort('productCode')}>
                        รหัสสินค้า
                        {modalSortConfig?.column === 'productCode'
                          ? modalSortConfig.direction === 'asc'
                            ? ' ▲'
                            : ' ▼'
                          : ''}
                      </button>
                    </th>
                    <th>
                      <button type="button" className="sort-btn" onClick={() => handleModalSort('productName')}>
                        ชื่อสินค้า
                        {modalSortConfig?.column === 'productName'
                          ? modalSortConfig.direction === 'asc'
                            ? ' ▲'
                            : ' ▼'
                          : ''}
                      </button>
                    </th>
                    <th>
                      <button type="button" className="sort-btn" onClick={() => handleModalSort('minStock')}>
                        ขั้นต่ำ
                        {modalSortConfig?.column === 'minStock'
                          ? modalSortConfig.direction === 'asc'
                            ? ' ▲'
                            : ' ▼'
                          : ''}
                      </button>
                    </th>
                    <th>
                      <button type="button" className="sort-btn" onClick={() => handleModalSort('currentStock')}>
                        คงเหลือ
                        {modalSortConfig?.column === 'currentStock'
                          ? modalSortConfig.direction === 'asc'
                            ? ' ▲'
                            : ' ▼'
                          : ''}
                      </button>
                    </th>
                    <th>
                      <button type="button" className="sort-btn" onClick={() => handleModalSort('deficit')}>
                        ขาด
                        {modalSortConfig?.column === 'deficit'
                          ? modalSortConfig.direction === 'asc'
                            ? ' ▲'
                            : ' ▼'
                          : ''}
                      </button>
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {visibleModalIssues.map((item, index) => (
                    <tr key={`${item.productLabel}-${index}`}>
                      <td>{item.productCode}</td>
                      <td>{item.productName}</td>
                      <td>{formatNumber(item.minStock)}</td>
                      <td>{formatNumber(item.currentStock)}</td>
                      <td>{formatNumber(item.deficit)}</td>
                    </tr>
                  ))}
                  {visibleModalIssues.length === 0 && (
                    <tr>
                      <td colSpan={5}>ไม่พบรายการที่ต้องเบิกสำหรับสาขานี้</td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
            <div className="modal-actions">
              <button type="button" className="ghost-btn" onClick={() => setSelectedBranchForModal('')}>
                ปิด
              </button>
            </div>
          </div>
        </div>
      )}

      {selectedChartGroupForModal && (
        <div className="modal-backdrop" role="presentation" onClick={() => setSelectedChartGroupForModal('')}>
          <div
            className="modal-card modal-wide"
            role="dialog"
            aria-modal="true"
            onClick={(event) => event.stopPropagation()}
          >
            <h3>รายละเอียดกลุ่ม - {selectedChartGroupForModal}</h3>
            <p>
              แสดงรายการข้อมูลทั้งหมดของกลุ่มนี้ โดยอ้างอิงจากมิติ "{selectedDimension || '-'}"
            </p>
            <div className="modal-toolbar">
              <input
                type="search"
                className="modal-search-input"
                placeholder="ค้นหาในรายการกลุ่ม"
                value={chartModalSearchText}
                onChange={(event) => setChartModalSearchText(event.target.value)}
              />
            </div>
            <div className="modal-actions modal-actions-left">
              <button type="button" className="ghost-btn" onClick={() => exportChartGroupRows('csv')}>
                Export CSV
              </button>
              <button type="button" className="ghost-btn" onClick={() => exportChartGroupRows('xlsx')}>
                Export Excel
              </button>
            </div>
            <div className="table-scroll modal-table-scroll">
              <table className="mini-table" style={{ minWidth: `${Math.max(760, headers.length * 160)}px` }}>
                <thead>
                  <tr>
                    {headers.map((header) => (
                      <th key={`chart-modal-${header}`}>
                        <button type="button" className="sort-btn" onClick={() => handleChartModalSort(header)}>
                          {header}
                          {chartModalSortConfig?.column === header
                            ? chartModalSortConfig.direction === 'asc'
                              ? ' ▲'
                              : ' ▼'
                            : ''}
                        </button>
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody>
                  {visibleChartGroupRows.map((row, rowIndex) => (
                    <tr key={`chart-modal-row-${rowIndex}`}>
                      {headers.map((header) => (
                        <td key={`chart-modal-cell-${rowIndex}-${header}`}>{String(row[header] ?? '')}</td>
                      ))}
                    </tr>
                  ))}
                  {visibleChartGroupRows.length === 0 && (
                    <tr>
                      <td colSpan={Math.max(1, headers.length)}>ไม่พบข้อมูลในกลุ่มนี้</td>
                    </tr>
                  )}
                </tbody>
              </table>
            </div>
            <div className="modal-actions">
              <button
                type="button"
                className="ghost-btn"
                onClick={() => setSelectedChartGroupForModal('')}
              >
                ปิด
              </button>
            </div>
          </div>
        </div>
      )}
    </main>
  )
}

export default App