import { useCallback, useEffect, useMemo, useState } from 'react'
import * as XLSX from 'xlsx-js-style'
import { saveAs } from 'file-saver'
import './App.css'

const DEVICE_EVENT_PAD_LENGTH = 6
const MILLISECONDS_PER_HOUR = 60 * 60 * 1000
const EVENT_CODES = {
  BOLUS: 2100,
}

const BOLUS_EXPORT_CODE = 262144

const EXPORT_EVENT_REMAP = {
  [EVENT_CODES.BOLUS]: {
    code: BOLUS_EXPORT_CODE,
    eventPort: 4,
    eventType: 0,
    eventLevel: 0,
  },
}

const EXPORT_FIELD_ORDER = [
  'id',
  'deviceSn',
  'datetime',
  'originalDatetime',
  'basal',
  'bolus',
  'event',
  'eventIndex',
  'eventPort',
  'eventType',
  'eventLevel',
  'eventValue',
  'remainingCapacity',
  'remainingInsulin',
  'autoMode',
  'delete',
]

const EXPORT_ALL_FIELD_ORDER = [
  'id',
  'deviceSn',
  'datetime',
  'basal',
  'bolus',
  'bolusTotal',
  'event',
  'deviceEventLabel',
  'eventIndex',
  'eventPort',
  'eventType',
  'eventLevel',
  'eventValue',
  'remainingCapacity',
  'remainingInsulin',
  'autoMode',
  'delete',
]
const PAGE_SIZE = 1000
const NARROW_COLUMNS = new Set([
  'basal',
  'bolus',
  'remainingCapacity',
  'remainingInsulin',
])

const parseDeviceEventValue = (rawValue) => {
  if (rawValue === undefined || rawValue === null) return null
  const trimmed = String(rawValue).trim()
  if (!trimmed) return null

  const numeric = /^0x/i.test(trimmed)
    ? Number.parseInt(trimmed, 16)
    : Number.parseInt(trimmed, 10)

  if (Number.isNaN(numeric) || numeric < 0) return NaN

  return numeric
}

const formatDeviceEventValue = (rawValue) => {
  const numeric = parseDeviceEventValue(rawValue)
  if (numeric === null) return ''
  if (Number.isNaN(numeric)) return String(rawValue)
  return `0x${numeric.toString(16).toUpperCase().padStart(DEVICE_EVENT_PAD_LENGTH, '0')}`
}

const splitDeviceEvent = (rawValue) => {
  const numeric = parseDeviceEventValue(rawValue)
  if (numeric === null || Number.isNaN(numeric)) {
    return { eventPort: '', eventType: '', eventLevel: '' }
  }

  return {
    eventPort: (numeric >> 16) & 0xff,
    eventType: (numeric >> 8) & 0xff,
    eventLevel: numeric & 0xff,
  }
}

const extractDeviceEventNumeric = (item) => {
  if (!item || typeof item !== 'object') return null
  const candidates = [
    item.event,
    item.deviceEvent,
    item.DEVICE_EVENT,
    item.DeviceEvent,
  ]

  for (const candidate of candidates) {
    const numeric = parseDeviceEventValue(candidate)
    if (numeric !== null && !Number.isNaN(numeric)) {
      return numeric
    }
  }

  return null
}

const combineDeviceEventParts = ({ eventPort, eventType, eventLevel }) => {
  const port = Number(eventPort)
  const type = Number(eventType)
  const level = Number(eventLevel)

  if (
    Number.isNaN(port) ||
    Number.isNaN(type) ||
    Number.isNaN(level) ||
    port < 0 ||
    type < 0 ||
    level < 0
  ) {
    return null
  }

  return ((port & 0xff) << 16) | ((type & 0xff) << 8) | (level & 0xff)
}

const DEVICE_EVENT_BASE_OPTIONS = [
  { code: 0x040000, label: '>>' },
  { code: 0x040700, label: '推杆定位' },
  { code: 0x040800, label: '推杆回退' },
  { code: 0x040500, label: '输注已暂停' },
  { code: 0x040c00, label: '输注已暂停' },
  { code: 0x050100, label: '上电复位' },
  { code: 0x040300, label: '监测到反转' },
  { code: 0x040a00, label: '临时基础率开始' },
  { code: 0x040b00, label: '临时基础率结束' },
  { code: 0x030000, label: '自动模式记录' },
  { code: 0x040900, label: '其他通知' },
  { code: 0x050400, label: '泵充电状态' },
  { code: 0x030100, label: '闭环基础率达到上限' },
  { code: 0x030200, label: '自动暂停基础率输注' },
  { code: 0x030500, label: '目标血糖改变' },
  { code: 0x030600, label: '自动基础率为0达到4小时' },
  { code: 0x030700, label: '闭环模式改变' },
  { code: 0x030800, label: '葡萄糖传感器故障,自动模式已退出。' },
  { code: 0x030900, label: '葡萄糖传感器过期，自动模式已退出' },
  { code: 0x040101, label: '药量低！' },
  { code: 0x040501, label: '输注已经暂停' },
  { code: 0x040601, label: '泵体即将关机！' },
  { code: 0x050001, label: '泵电量低！' },
  { code: 0x050201, label: '泵按键故障！' },
  { code: 0x050301, label: '进入强磁场，泵体可能失效！' },
  { code: 0x030101, label: '短期达到最大输注量，自动模式已退出!' },
  { code: 0x030201, label: '自动暂停基础率时间过长，自动模式已退出!' },
  { code: 0x030301, label: '持续高血糖！' },
  { code: 0x030401, label: '过长时间未接收到CGM的数据，自动模式已退出!' },
  { code: 0x050002, label: '泵体电量耗尽，输注已停止!!!' },
  { code: 0x040102, label: '药液耗尽，输注已停止!!!' },
  { code: 0x040202, label: '检测到阻塞，输注已停止!!!' },
  { code: 0x040302, label: '电机故障，输注已停止!!!' },
  { code: 0x050102, label: '非正常输注停止!!!' },
  { code: 0x040602, label: '长时间未操作泵体，输注已停止!!!' },
  { code: 2100, label: '大剂量' },
  { code: 2101, label: '开机上电' },
]

const DEVICE_EVENT_OPTIONS = DEVICE_EVENT_BASE_OPTIONS.map((option) => {
  const label =
    option.label && option.label.trim().length > 0
      ? option.label
      : formatDeviceEventValue(option.code)
  const { eventPort, eventType, eventLevel } = splitDeviceEvent(option.code)
  return {
    code: option.code,
    label,
    value: String(option.code),
    eventPort,
    eventType,
    eventLevel,
    hex: formatDeviceEventValue(option.code),
  }
}).sort((a, b) => a.code - b.code)

const DEVICE_EVENT_MAP = new Map(
  DEVICE_EVENT_OPTIONS.map((option) => [option.code, option])
)

const getDeviceEventOption = (code) => {
  if (code === '' || code === undefined || code === null) return undefined
  const numeric = parseDeviceEventValue(code)
  if (numeric === null || Number.isNaN(numeric)) return undefined
  return DEVICE_EVENT_MAP.get(numeric)
}

const getDeviceEventLabel = (code) => {
  const option = getDeviceEventOption(code)
  if (option) return option.label
  const numeric = parseDeviceEventValue(code)
  if (numeric === null || Number.isNaN(numeric)) return ''
  return formatDeviceEventValue(numeric)
}

const resolveDeviceEventCode = (rowCurrent) => {
  if (!rowCurrent) return null
  const candidate = parseDeviceEventValue(rowCurrent.deviceEvent)
  if (candidate !== null && !Number.isNaN(candidate)) {
    return candidate
  }

  const { eventPort, eventType, eventLevel } = rowCurrent
  const isEmpty =
    (eventPort === '' || eventPort === undefined || eventPort === null) &&
    (eventType === '' || eventType === undefined || eventType === null) &&
    (eventLevel === '' || eventLevel === undefined || eventLevel === null)
  if (isEmpty) return null

  const combined = combineDeviceEventParts({
    eventPort,
    eventType,
    eventLevel,
  })

  if (combined === null) return null
  return combined
}

const COLUMN_CONFIG = [
  {
    key: 'createTime',
    header: 'CREATE_TIME',
    label: '创建时间',
    hidden: true,
    type: 'text',
    map: (item) => item?.datetime ?? item?.time ?? '',
  },
  {
    key: 'eventIndex',
    header: 'EVENT_INDEX',
    label: '事件序号',
    type: 'number',
    readOnly: true,
    map: (item) => item?.eventIndex ?? '',
  },
  {
    key: 'deviceTime',
    header: 'DEVICE_TIME',
    label: '设备时间',
    type: 'text',
    map: (item) => item?.datetime ?? item?.time ?? '',
  },
  {
    key: 'deviceEvent',
    header: 'DEVICE_EVENT',
    label: '设备事件',
    type: 'number',
    editor: 'select',
    options: DEVICE_EVENT_OPTIONS.map(({ value, label }) => ({
      value,
      label,
    })),
    map: (item) => {
      const numeric = extractDeviceEventNumeric(item)
      if (numeric === null) return ''
      return numeric
    },
    parse: (value) => {
      if (value === '' || value === undefined || value === null) return ''
      const numeric = Number(value)
      return Number.isNaN(numeric) ? '' : numeric
    },
  },
  {
    key: 'basal',
    header: 'BASAL',
    label: '基础率',
    type: 'number',
    map: (item) => item?.basal ?? '',
  },
  {
    key: 'bolus',
    header: 'BOLUS',
    label: '大剂量',
    type: 'number',
    map: (item) => item?.bolus ?? '',
  },
  {
    key: 'remainingCapacity',
    header: 'REMAINING_CAPACITY',
    label: '剩余电量',
    type: 'number',
    map: (item) => item?.remainingCapacity ?? '',
  },
  {
    key: 'remainingInsulin',
    header: 'REMAINING_INSULIN',
    label: '剩余胰岛素',
    type: 'number',
    map: (item) => item?.remainingInsulin ?? '',
  },
  {
    key: 'eventPort',
    header: 'EVENT_PORT',
    label: '事件端口',
    hidden: true,
    type: 'number',
    map: (item) => {
      const numeric = extractDeviceEventNumeric(item)
      if (numeric !== null) {
        const { eventPort } = splitDeviceEvent(numeric)
        return eventPort
      }
      if (item?.eventPort !== undefined && item?.eventPort !== null) {
        return item.eventPort
      }
      return ''
    },
  },
  {
    key: 'eventType',
    header: 'EVENT_TYPE',
    label: '事件类型',
    hidden: true,
    type: 'number',
    map: (item) => {
      const numeric = extractDeviceEventNumeric(item)
      if (numeric !== null) {
        const { eventType } = splitDeviceEvent(numeric)
        return eventType
      }
      if (item?.eventType !== undefined && item?.eventType !== null) {
        return item.eventType
      }
      return ''
    },
  },
  {
    key: 'eventLevel',
    header: 'EVENT_LEVEL',
    label: '事件级别',
    hidden: true,
    type: 'number',
    map: (item) => {
      const numeric = extractDeviceEventNumeric(item)
      if (numeric !== null) {
        const { eventLevel } = splitDeviceEvent(numeric)
        return eventLevel
      }
      if (item?.eventLevel !== undefined && item?.eventLevel !== null) {
        return item.eventLevel
      }
      return ''
    },
  },
  {
    key: 'eventData',
    header: 'EVENT_Data',
    label: '事件数据',
    type: 'text',
    hidden: true,
    readOnly: true,
    map: (item) => {
      if (item?.eventValue !== undefined && item?.eventValue !== null) {
        return item.eventValue
      }
      if (item?.EVENT_Data !== undefined && item?.EVENT_Data !== null) {
        return item.EVENT_Data
      }
      return ''
    },
  },
  {
    key: 'autoMode',
    header: 'AUTO_MODE',
    label: '自动模式',
    editor: 'select',
    options: [
      { label: '开启', value: 'true' },
      { label: '关闭', value: 'false' },
    ],
    parse: (value) => {
      if (value === '' || value === undefined || value === null) return ''
      if (value === true || value === 'true' || value === 1 || value === '1')
        return true
      if (value === false || value === 'false' || value === 0 || value === '0')
        return false
      return Boolean(value)
    },
    map: (item) => item?.autoMode ?? '',
  },
  {
    key: 'basalDuration',
    header: '基础率时长',
    label: '基础率时长',
    hidden: true,
    type: 'number',
    map: (item) => item?.basalUnitPerHour ?? '',
  },
  {
    key: 'basalAmount',
    header: '基础注射量',
    label: '基础注射量',
    hidden: false,
    type: 'number',
    readOnly: true,
    map: (item) => item?.presetBasalUnitPerHour ?? '',
  },
  {
    key: 'bolusTotal',
    header: '大剂量总量',
    label: '大剂量总量',
    type: 'number',
    readOnly: true,
    map: (item) => item?.bolusSum ?? '',
  },
  {
    key: 'eventSummary',
    header: '事件',
    label: '事件',
    hidden: true,
    type: 'text',
    map: (item) => item?.event ?? '',
  },
  {
    key: 'recordId',
    header: 'ID',
    label: '记录ID',
    hidden: true,
    type: 'text',
    map: (item) => item?.id ?? '',
  },
  {
    key: 'deviceSn',
    header: 'SN',
    label: '设备SN',
    type: 'text',
    readOnly: true,
    map: (item) => item?.deviceSn ?? '',
  },
]

const EDITABLE_COLUMNS = COLUMN_CONFIG.filter((column) => !column.readOnly)
const VISIBLE_COLUMNS = COLUMN_CONFIG.filter((column) => !column.hidden)

const normaliseValue = (value) =>
  value === undefined || value === null ? '' : value

const toInputDateTime = (value) => {
  if (!value) return ''

  let dateInstance

  if (value instanceof Date) {
    dateInstance = value
  } else if (typeof value === 'number') {
    dateInstance = new Date(value)
  } else if (typeof value === 'string') {
    const trimmed = value.trim()
    if (!trimmed) return ''

    if (/^\d+$/.test(trimmed)) {
      dateInstance = new Date(Number(trimmed))
    } else if (trimmed.includes('T')) {
      dateInstance = new Date(trimmed)
    } else {
      dateInstance = new Date(trimmed.replace(' ', 'T'))
    }
  }

  if (!dateInstance || Number.isNaN(dateInstance.getTime())) return ''

  const local = new Date(dateInstance.getTime() - dateInstance.getTimezoneOffset() * 60000)
  return local.toISOString().slice(0, 19)
}

const fromInputDateTime = (value) => {
  if (!value) return ''
  const [datePart, timePart] = value.split('T')
  if (!datePart) return value
  const withSeconds =
    timePart && timePart.length === 5
      ? `${timePart}:00`
      : timePart && timePart.length === 8
        ? timePart
        : timePart ?? '00:00:00'
  return `${datePart} ${withSeconds}`
}

const formatDeviceTimeDisplay = (value) => {
  if (value === undefined || value === null) return ''
  if (typeof value === 'string') {
    const trimmed = value.trim()
    if (!trimmed) return ''
    if (trimmed.includes('T')) {
      return trimmed.replace('T', ' ')
    }
    return trimmed
  }
  if (value instanceof Date) {
    return value.toISOString().replace('T', ' ').slice(0, 19)
  }
  const parsed = parseDateTimeToMs(value)
  if (Number.isFinite(parsed)) {
    return new Date(parsed).toISOString().replace('T', ' ').slice(0, 19)
  }
  return String(value)
}

let _rowUidCounter = 0
const nextRowUid = () => `_row_${++_rowUidCounter}_${Date.now()}`

const buildRowModel = (item, index) => {
  const baseValues = {}

  COLUMN_CONFIG.forEach((column) => {
    if (column.key === 'status') return

    const mappedValue =
      typeof column.map === 'function'
        ? column.map(item, index)
        : item?.[column.key]

    if (column.key === 'deviceTime') {
      const inputValue = toInputDateTime(mappedValue)
      baseValues[column.key] = inputValue
    } else {
      let value = normaliseValue(mappedValue)
      if (typeof column.parse === 'function' && !column.readOnly) {
        value = column.parse(value)
      }
      baseValues[column.key] = value
    }
  })

  const deleteFlag = Number(item?.delete ?? item?.Delete ?? 0)

  return {
    id: nextRowUid(),
    sourceId: normaliseValue(item?.id) !== '' ? item.id : `${item?.deviceSn ?? 'row'}-${index}`,
    source: item,
    original: { ...baseValues },
    current: { ...baseValues },
    deleted: deleteFlag === 1,
    added: deleteFlag === 2,
  }
}

const parseDateTimeToMs = (value) => {
  if (!value && value !== 0) return Number.NaN
  if (value instanceof Date) return value.getTime()
  if (typeof value === 'number') return Number.isFinite(value) ? value : Number.NaN
  if (typeof value === 'string') {
    const trimmed = value.trim()
    if (!trimmed) return Number.NaN
    const normalised =
      trimmed.includes('T') || trimmed.endsWith('Z')
        ? trimmed
        : trimmed.replace(' ', 'T')
    const parsed = Date.parse(normalised)
    return Number.isNaN(parsed) ? Number.NaN : parsed
  }
  return Number.NaN
}

const formatDuration = (ms) => {
  if (!Number.isFinite(ms) || ms <= 0) return '00:00:00'
  const totalSeconds = Math.floor(ms / 1000)
  const hours = Math.floor(totalSeconds / 3600)
  const minutes = Math.floor((totalSeconds % 3600) / 60)
  const seconds = totalSeconds % 60
  const pad = (num) => String(num).padStart(2, '0')
  return `${pad(hours)}:${pad(minutes)}:${pad(seconds)}`
}

const applyBolusCalculations = (items) => {
  if (!Array.isArray(items) || items.length === 0) return items

  // 兼容性处理：如果 event 字段为空，通过 eventPort、eventType、eventLevel 生成
  const normalizedItems = items.map((item) => {
    const eventValue = item?.event ?? item?.deviceEvent ?? item?.DEVICE_EVENT
    const hasEvent = eventValue !== undefined && eventValue !== null && eventValue !== ''
    
    if (!hasEvent) {
      const eventPort = item?.eventPort
      const eventType = item?.eventType
      const eventLevel = item?.eventLevel
      
      // 如果三个字段都存在且有效，则生成 event
      if (
        eventPort !== undefined && eventPort !== null && eventPort !== '' &&
        eventType !== undefined && eventType !== null && eventType !== '' &&
        eventLevel !== undefined && eventLevel !== null && eventLevel !== ''
      ) {
        const combinedEvent = combineDeviceEventParts({
          eventPort,
          eventType,
          eventLevel,
        })
        
        if (combinedEvent !== null) {
          return {
            ...item,
            event: combinedEvent,
            deviceEvent: combinedEvent,
          }
        }
      }
    }
    
    return item
  })

  const decorated = normalizedItems
    .map((item, index) => {
      const timestamp = parseDateTimeToMs(
        item?.datetime ?? item?.time ?? item?.DEVICE_TIME ?? item?.createTime ?? item?.endtime
      )
      return { item, timestamp, index }
    })
    .sort((a, b) => {
      const { timestamp: ta, index: ia } = a
      const { timestamp: tb, index: ib } = b
      if (Number.isFinite(ta) && Number.isFinite(tb)) {
        if (ta === tb) return ia - ib
        return ta - tb
      }
      if (Number.isFinite(ta)) return -1
      if (Number.isFinite(tb)) return 1
      return ia - ib
    })

  const enriched = decorated.map(({ item }) => ({ ...item }))
  const timestamps = decorated.map(({ timestamp }) => timestamp)

  for (let index = 0; index < enriched.length - 1; index += 1) {
    const currentTime = timestamps[index]
    const nextTime = timestamps[index + 1]
    if (
      Number.isFinite(currentTime) &&
      Number.isFinite(nextTime) &&
      nextTime >= currentTime
    ) {
      const diff = nextTime - currentTime
      const diffHours = diff / MILLISECONDS_PER_HOUR
      const entity = enriched[index]
      const nextEntity = enriched[index + 1]
      
      // 计算基础率输注量 = （下一个事件时间 - 当前事件时间）* （basal / 160 ）U/hr
      const basalValue = Number(entity?.basal ?? 0)
      const basalAmount = Number.isFinite(basalValue) && Number.isFinite(diffHours) && diffHours > 0
        ? diffHours * (basalValue / 160)
        : 0

      // 将基础率输注量存储到当前事件
      enriched[index] = {
        ...entity,
        presetBasalUnitPerHour: Number.isFinite(basalAmount) ? Number(basalAmount.toFixed(3)) : basalAmount,
      }

      enriched[index + 1] = {
        ...nextEntity,
        baselSum: basalAmount,
        duction: formatDuration(diff),
        ductionTime: diff,
      }
    }
  }

  let activeEvent = null

  for (let index = 0; index < enriched.length; index += 1) {
    const entity = enriched[index]
    const currentTime = timestamps[index]
    const bolusValue = Number(entity?.bolus ?? 0)

    if (
      !activeEvent &&
      Number.isFinite(currentTime) &&
      Number.isFinite(bolusValue) &&
      bolusValue !== 0
    ) {
      activeEvent = {
        startIndex: index,
        startTime: currentTime,
        bolusValue,
      }

      const bolusCode = EVENT_CODES.BOLUS
      const option = DEVICE_EVENT_MAP.get(bolusCode)
      const components = option ?? splitDeviceEvent(bolusCode)

      enriched[index] = {
        ...entity,
        event: bolusCode,
        eventPort: components.eventPort ?? 0,
        eventType: components.eventType ?? 0,
        eventLevel: components.eventLevel ?? 0,
        deviceEvent: bolusCode,
      }
      continue
    }

    if (
      activeEvent &&
      Number.isFinite(currentTime) &&
      Number.isFinite(bolusValue) &&
      bolusValue === 0
    ) {
      const durationHours = Math.max(
        (currentTime - activeEvent.startTime) / MILLISECONDS_PER_HOUR,
        0
      )
      const totalBolus = (activeEvent.bolusValue / 160) * durationHours
      const startEntity = enriched[activeEvent.startIndex]

      startEntity.bolusSum = Number.isFinite(totalBolus)
        ? Number(totalBolus.toFixed(6))
        : totalBolus
      startEntity.bolusTotal = startEntity.bolusSum

      activeEvent = null
    }
  }

  if (activeEvent) {
    const startEntity = enriched[activeEvent.startIndex]
    if (startEntity.bolusSum === undefined) {
      startEntity.bolusSum = 0
      startEntity.bolusTotal = 0
    }
  }

  return enriched
}

function App() {
  const [rows, setRows] = useState([])
  const [sourceItems, setSourceItems] = useState([])
  const [error, setError] = useState('')
  const [startDate, setStartDate] = useState('')
  const [endDate, setEndDate] = useState('')
  const [isModalOpen, setIsModalOpen] = useState(false)
  const [selectedFileName, setSelectedFileName] = useState('')
  const [modalError, setModalError] = useState('')
  const [sourceFileBase, setSourceFileBase] = useState('')
  const [sourceFileExtension, setSourceFileExtension] = useState('json')
  const [currentPage, setCurrentPage] = useState(1)
  const [isLoading, setIsLoading] = useState(false)
  const [isExportModalOpen, setIsExportModalOpen] = useState(false)
  const [exportFileName, setExportFileName] = useState('')
  const [isExportAllModalOpen, setIsExportAllModalOpen] = useState(false)
  const [exportAllFileName, setExportAllFileName] = useState('')
  const [isHighPriorityModalOpen, setIsHighPriorityModalOpen] = useState(false)
  const [isFormatModalOpen, setIsFormatModalOpen] = useState(false)
  // expandedDeletedRows 记录被用户手动展开的已删除行
  const [expandedDeletedRows, setExpandedDeletedRows] = useState(new Set())

  const closeModal = useCallback(() => {
    setIsModalOpen(false)
    setModalError('')
  }, [])

  useEffect(() => {
    const maxPage = Math.max(1, Math.ceil(Math.max(rows.length, 1) / PAGE_SIZE))
    if (currentPage > maxPage) {
      setCurrentPage(maxPage)
    }
  }, [rows, currentPage])

  const paginatedRows = useMemo(() => {
    const start = (currentPage - 1) * PAGE_SIZE
    return rows.slice(start, start + PAGE_SIZE)
  }, [rows, currentPage])

  const totalPages = useMemo(
    () => Math.max(1, Math.ceil(Math.max(rows.length, 1) / PAGE_SIZE)),
    [rows.length]
  )

  const highPriorityRows = useMemo(
    () =>
      paginatedRows
        .map((row, index) => ({
          row,
          pageIndex: index,
        }))
        .filter(({ row }) =>
          Number(row.current?.eventLevel ?? row.source?.eventLevel ?? row.source?.EVENT_LEVEL) === 2
        ),
    [paginatedRows]
  )

  const currentHighPriorityCount = highPriorityRows.length

  const handlePageChange = useCallback(
    (page) => {
      const max = Math.max(1, Math.ceil(Math.max(rows.length, 1) / PAGE_SIZE))
      const next = Math.min(Math.max(1, page), max)
      if (next === currentPage) return
      setIsLoading(true)
      window.requestAnimationFrame(() => {
        setCurrentPage(next)
        setIsLoading(false)
      })
    },
    [rows.length, currentPage]
  )

  const buildUpdatedSource = useCallback(
    (row) => {
      const current = row?.current ?? {}
      const updated = { ...(row?.source ?? {}) }

      const rawDeviceEventCode = resolveDeviceEventCode(current)
      const remap =
        rawDeviceEventCode !== null && !Number.isNaN(rawDeviceEventCode)
          ? EXPORT_EVENT_REMAP[rawDeviceEventCode]
          : undefined
      const mappedDeviceEventCode =
        remap?.code ??
        (rawDeviceEventCode !== null && !Number.isNaN(rawDeviceEventCode)
          ? rawDeviceEventCode
          : rawDeviceEventCode)

      let deviceEventComponents = null
      if (remap || mappedDeviceEventCode === BOLUS_EXPORT_CODE) {
        deviceEventComponents = {
          eventPort: remap?.eventPort ?? 4,
          eventType: remap?.eventType ?? 0,
          eventLevel: remap?.eventLevel ?? 0,
        }
      } else if (mappedDeviceEventCode !== null && !Number.isNaN(mappedDeviceEventCode)) {
        deviceEventComponents = splitDeviceEvent(mappedDeviceEventCode)
      }

      if (mappedDeviceEventCode !== null && !Number.isNaN(mappedDeviceEventCode)) {
        updated.event = mappedDeviceEventCode
        updated.deviceEvent = mappedDeviceEventCode
      } else if (current.deviceEvent !== undefined && current.deviceEvent !== '') {
        updated.deviceEvent = current.deviceEvent
      }

      if (deviceEventComponents) {
        updated.eventPort = deviceEventComponents.eventPort
        updated.eventType = deviceEventComponents.eventType
        updated.eventLevel = deviceEventComponents.eventLevel
      } else {
        if (
          current.eventPort !== undefined &&
          current.eventPort !== ''
        ) {
          updated.eventPort = current.eventPort
        }
        if (
          current.eventType !== undefined &&
          current.eventType !== ''
        ) {
          updated.eventType = current.eventType
        }
        if (
          current.eventLevel !== undefined &&
          current.eventLevel !== ''
        ) {
          updated.eventLevel = current.eventLevel
        }
      }
      if (
        'eventValue' in updated &&
        current.eventData !== undefined &&
        current.eventData !== ''
      ) {
        updated.eventValue = current.eventData
      }
      if (
        'autoMode' in updated &&
        current.autoMode !== undefined &&
        current.autoMode !== ''
      ) {
        updated.autoMode = current.autoMode
      }
      if (
        'basalUnitPerHour' in updated &&
        current.basalDuration !== undefined &&
        current.basalDuration !== ''
      ) {
        updated.basalUnitPerHour = current.basalDuration
      }
      if (
        'presetBasalUnitPerHour' in updated &&
        current.basalAmount !== undefined &&
        current.basalAmount !== ''
      ) {
        updated.presetBasalUnitPerHour = current.basalAmount
      }
      if (
        'bolusSum' in updated &&
        current.bolusTotal !== undefined &&
        current.bolusTotal !== ''
      ) {
        updated.bolusSum = current.bolusTotal
      }
      if (
        'id' in updated &&
        current.recordId !== undefined &&
        current.recordId !== ''
      ) {
        updated.id = current.recordId
      }
      if (
        current.deviceTime !== undefined &&
        current.deviceTime !== null &&
        current.deviceTime !== ''
      ) {
        const formatted = fromInputDateTime(current.deviceTime)
        if ('datetime' in updated) updated.datetime = formatted
        if ('time' in updated) updated.time = formatted
      }

      const overrideKeys = new Set([
        'event',
        'deviceEvent',
        'eventPort',
        'eventType',
        'eventLevel',
      ])

      Object.keys(current).forEach((key) => {
        if (overrideKeys.has(key)) return
        if (key in updated && current[key] !== undefined && current[key] !== null) {
          updated[key] = current[key]
        }
      })

      return updated
    },
    []
  )

  const buildExportRecord = useCallback(
    (row) => {
      const updatedSource = buildUpdatedSource(row)
      const record = {}
      EXPORT_FIELD_ORDER.forEach((field) => {
        if (field === 'delete') {
          record.delete = row.added ? 2 : row.deleted ? 1 : 0
        } else if (field === 'originalDatetime') {
          // 保存原始时间字段，用于后端对比
          const originalTime =
            row.source?.datetime ?? row.source?.time ?? row.original?.deviceTime ?? ''
          record.originalDatetime = originalTime
        } else if (field in updatedSource) {
          record[field] = updatedSource[field]
        } else {
          record[field] = ''
        }
      })
      return record
    },
    [buildUpdatedSource]
  )

  const buildExportAllRecord = useCallback(
    (row) => {
      const updatedSource = buildUpdatedSource(row)
      const record = {}

      const eventCandidate =
        (updatedSource.event !== undefined && updatedSource.event !== '' && updatedSource.event !== null
          ? updatedSource.event
          : undefined) ??
        (updatedSource.deviceEvent !== undefined &&
        updatedSource.deviceEvent !== '' &&
        updatedSource.deviceEvent !== null
          ? updatedSource.deviceEvent
          : undefined) ??
        row.current.deviceEvent ??
        row.source?.event ??
        row.source?.deviceEvent ??
        ''

      const bolusTotalValue =
        row.current.bolusTotal ??
        row.current.bolusSum ??
        updatedSource.bolusTotal ??
        updatedSource.bolusSum ??
        row.source?.bolusTotal ??
        row.source?.bolusSum ??
        ''

      EXPORT_ALL_FIELD_ORDER.forEach((field) => {
        if (field === 'delete') {
          record.delete = row.added ? 2 : row.deleted ? 1 : 0
        } else if (field === 'deviceEventLabel') {
          record.deviceEventLabel = getDeviceEventLabel(eventCandidate)
        } else if (field === 'bolusTotal') {
          record.bolusTotal = bolusTotalValue
        } else if (field in updatedSource) {
          record[field] = updatedSource[field]
        } else if (field in row.current) {
          record[field] = row.current[field]
        } else if (field in (row.source ?? {})) {
          record[field] = row.source[field]
        } else {
          record[field] = ''
        }
      })

      return record
    },
    [buildUpdatedSource]
  )

  const isRowModified = useCallback(
    (row) =>
      EDITABLE_COLUMNS.some((column) => {
        const key = column.key
        return row.current[key] !== row.original[key]
      }),
    []
  )

  const summary = useMemo(() => {
    const totals = rows.reduce(
      (acc, row) => {
        if (row.added) {
          acc.added += 1
        } else if (row.deleted) {
          acc.deleted += 1
        } else if (isRowModified(row)) {
          acc.modified += 1
        }
        return acc
      },
      { modified: 0, deleted: 0, added: 0 }
    )

    return {
      total: rows.length,
      modified: totals.modified,
      deleted: totals.deleted,
      added: totals.added,
    }
  }, [rows, isRowModified])

  const changedRows = useMemo(
    () => rows.filter((row) => row.added || row.deleted || isRowModified(row)),
    [rows, isRowModified]
  )

  const exportStats = useMemo(() => {
    const deleted = changedRows.filter((row) => row.deleted).length
    const added = changedRows.filter((row) => row.added).length
    return {
      total: changedRows.length,
      deleted,
      added,
      modified: changedRows.length - deleted - added,
      modified: changedRows.length - deleted,
    }
  }, [changedRows])

  const setFromRawData = useCallback((raw, options = {}) => {
    const {
      sourceLabel = '',
      autoApply = false,
      suppressModal = false,
      fileBaseName = '',
      fileExtension = '',
    } = options

    if (!raw) {
      setRows([])
      setCurrentPage(1)
      setSourceItems([])
      setStartDate('')
      setEndDate('')
      setSelectedFileName('')
      setIsModalOpen(false)
      setSourceFileBase('')
      setSourceFileExtension('json')
      return
    }

    const normalisedArray = (() => {
      if (Array.isArray(raw)) return raw
      if (Array.isArray(raw?.records)) return raw.records
      if (Array.isArray(raw?.data)) return raw.data
      if (raw && typeof raw === 'object') return [raw]
      return []
    })()

    const filteredValues = normalisedArray.filter(
      (item) => item && typeof item === 'object'
    )

    setRows([])
    setCurrentPage(1)
    setError('')
    setSelectedFileName(sourceLabel)
    setModalError('')
    if (fileBaseName) {
      setSourceFileBase(fileBaseName)
    } else if (!sourceFileBase) {
      setSourceFileBase('pump-history')
    }
    if (fileExtension) {
      setSourceFileExtension(fileExtension.toLowerCase())
    } else if (!sourceFileExtension) {
      setSourceFileExtension('json')
    }

    if (!filteredValues.length) {
      setSourceItems([])
      setStartDate('')
      setEndDate('')
      setIsModalOpen(false)
      setModalError('')
      return
    }

    const timestamps = filteredValues
      .map((item) =>
        parseDateTimeToMs(
          item?.datetime ?? item?.time ?? item?.DEVICE_TIME ?? item?.createTime
        )
      )
      .filter((value) => Number.isFinite(value))

    if (timestamps.length) {
      const minTs = Math.min(...timestamps)
      const maxTs = Math.max(...timestamps)
      const pad = (num) => String(num).padStart(2, '0')
      const buildDateString = (ms) => {
        const date = new Date(ms)
        return `${date.getFullYear()}-${pad(date.getMonth() + 1)}-${pad(date.getDate())}`
      }

      setStartDate(buildDateString(minTs))
      setEndDate(buildDateString(maxTs))
    } else {
      setStartDate('')
      setEndDate('')
    }

    setSourceItems(filteredValues)

    if (autoApply) {
      const enrichedValues = applyBolusCalculations(filteredValues)
      setCurrentPage(1)
      setRows(enrichedValues.map(buildRowModel))
      setError('')
      closeModal()
      setIsLoading(false)
    } else {
      setIsModalOpen(!suppressModal)
      setModalError('')
    }
  }, [closeModal, sourceFileBase, sourceFileExtension])

  const applyDateFilter = useCallback(() => {
    if (!sourceItems.length) {
      setRows([])
      setModalError('')
      return
    }

    const parseDateBoundary = (value, isEnd) => {
      if (!value) return isEnd ? Number.POSITIVE_INFINITY : Number.NEGATIVE_INFINITY
      const base = new Date(`${value}T00:00:00`)
      if (Number.isNaN(base.getTime())) {
        return isEnd ? Number.POSITIVE_INFINITY : Number.NEGATIVE_INFINITY
      }
      if (isEnd) {
        return base.getTime() + MILLISECONDS_PER_HOUR * 24 - 1
      }
      return base.getTime()
    }

    const startMs = parseDateBoundary(startDate, false)
    const endMs = parseDateBoundary(endDate, true)

    if (startMs > endMs) {
      setModalError('开始日期不能晚于结束日期')
      return
    }

    setIsLoading(true)
    window.requestAnimationFrame(() => {
      const filtered = sourceItems.filter((item) => {
        const timestamp = parseDateTimeToMs(
          item?.datetime ?? item?.time ?? item?.DEVICE_TIME ?? item?.createTime
        )
        if (!Number.isFinite(timestamp)) return false
        return timestamp >= startMs && timestamp <= endMs
      })

      if (!filtered.length) {
        setRows([])
        setError('选定日期范围内没有数据，请调整日期后再试。')
        setModalError('选定日期范围内没有数据，请调整日期后再试。')
        setIsLoading(false)
        return
      }

      const enrichedValues = applyBolusCalculations(filtered)
      setCurrentPage(1)
      setRows(enrichedValues.map(buildRowModel))
      setError('')
      setModalError('')
      closeModal()
      setIsLoading(false)
    })
  }, [sourceItems, startDate, endDate, closeModal])

  const handleFileUpload = useCallback(
    async (event) => {
      const [file] = event.target.files ?? []
      event.target.value = ''
      if (!file) return

      try {
        setIsLoading(true)
        const fileText = await file.text()
        const parsed = JSON.parse(fileText)
        const rawName = file.name || 'pump-history.json'
        const lastDot = rawName.lastIndexOf('.')
        const baseName =
          lastDot > 0 ? rawName.slice(0, lastDot) : rawName || 'pump-history'
        const extension =
          lastDot > 0 ? rawName.slice(lastDot + 1) : 'json'

        setFromRawData(parsed, {
          sourceLabel: rawName,
          autoApply: false,
          suppressModal: false,
          fileBaseName: baseName,
          fileExtension: extension,
        })
        setError('')
      } catch (parseError) {
        console.error(parseError)
        setError('上传的 JSON 文件无法解析，请检查格式')
      } finally {
        setIsLoading(false)
      }
    },
    [setFromRawData]
  )

  const handleExcelUpload = useCallback(
    async (event) => {
      const [file] = event.target.files ?? []
      event.target.value = ''
      if (!file) return

      try {
        setIsLoading(true)
        const buffer = await file.arrayBuffer()
        const workbook = XLSX.read(buffer, { type: 'array' })
        const sheetName = workbook.SheetNames[0]
        if (!sheetName) {
          setError('Excel 文件中没有工作表')
          return
        }
        const sheet = workbook.Sheets[sheetName]
        const parsed = XLSX.utils.sheet_to_json(sheet)

        if (!parsed.length) {
          setError('Excel 文件中没有数据')
          return
        }

        const rawName = file.name || 'pump-history.xlsx'
        const lastDot = rawName.lastIndexOf('.')
        const baseName =
          lastDot > 0 ? rawName.slice(0, lastDot) : rawName || 'pump-history'
        const extension =
          lastDot > 0 ? rawName.slice(lastDot + 1) : 'xlsx'

        setFromRawData(parsed, {
          sourceLabel: rawName,
          autoApply: false,
          suppressModal: false,
          fileBaseName: baseName,
          fileExtension: extension,
        })
        setError('')
      } catch (parseError) {
        console.error(parseError)
        setError('上传的 Excel 文件无法解析，请检查格式')
      } finally {
        setIsLoading(false)
      }
    },
    [setFromRawData]
  )

  const handleCellChange = useCallback(
    (rowId, key, rawValue) => {
      setRows((currentRows) =>
        currentRows.map((row) => {
          if (row.id !== rowId) return row

          const column = EDITABLE_COLUMNS.find((item) => item.key === key)
          if (!column) return row

          if (key === 'deviceEvent') {
            const nextCurrent = { ...row.current }

            if (rawValue === '' || rawValue === undefined || rawValue === null) {
              nextCurrent.deviceEvent = ''
              nextCurrent.eventPort = ''
              nextCurrent.eventType = ''
              nextCurrent.eventLevel = ''
              if ('eventSummary' in nextCurrent) {
                nextCurrent.eventSummary = ''
              }
            } else {
              const numeric = Number(rawValue)
              if (Number.isNaN(numeric)) {
                nextCurrent.deviceEvent = ''
                nextCurrent.eventPort = ''
                nextCurrent.eventType = ''
                nextCurrent.eventLevel = ''
                if ('eventSummary' in nextCurrent) {
                  nextCurrent.eventSummary = ''
                }
              } else {
                const option = DEVICE_EVENT_MAP.get(numeric)
                const components = option ?? splitDeviceEvent(numeric)
                nextCurrent.deviceEvent = numeric
                nextCurrent.eventPort = components.eventPort
                nextCurrent.eventType = components.eventType
                nextCurrent.eventLevel = components.eventLevel
                if ('eventSummary' in nextCurrent) {
                  nextCurrent.eventSummary = numeric
                }
              }
            }

            return {
              ...row,
              current: nextCurrent,
            }
          }

          let value = rawValue
          if (key === 'deviceTime') {
            value = rawValue
          } else if (column.editor === 'select') {
            value = rawValue
          } else if (column.type === 'number') {
            value = rawValue === '' ? '' : Number(rawValue)
          }

          if (typeof column.parse === 'function') {
            value = column.parse(value)
          }

          return {
            ...row,
            current: {
              ...row.current,
              [key]: value,
            },
          }
        })
      )

      if (key === 'deviceEvent') {
        setError('')
      }

      // 时间、基础率、大剂量变化时重新计算衍生字段
      if (key === 'deviceTime' || key === 'basal' || key === 'bolus') {
        setRows((currentRows) => recalcDerived(currentRows))
      }
    },
    [setError]
  )

  const toggleDeleteRow = useCallback((rowId) => {
    setRows((currentRows) =>
      currentRows.map((row) =>
        row.id === rowId ? { ...row, deleted: !row.deleted } : row
      )
    )
    // 取消删除时，从展开集合中移除
    setExpandedDeletedRows((prev) => {
      if (!prev.has(rowId)) return prev
      const next = new Set(prev)
      next.delete(rowId)
      return next
    })
  }, [])

  const toggleCollapseRow = useCallback((rowId) => {
    setExpandedDeletedRows((prev) => {
      const next = new Set(prev)
      if (next.has(rowId)) {
        next.delete(rowId)
      } else {
        next.add(rowId)
      }
      return next
    })
  }, [])

  const duplicateRow = useCallback((rowId) => {
    setRows((currentRows) => {
      const index = currentRows.findIndex((row) => row.id === rowId)
      if (index === -1) return currentRows
      const sourceRow = currentRows[index]
      const cloned = {
        id: nextRowUid(),
        sourceId: sourceRow.sourceId,
        source: JSON.parse(JSON.stringify(sourceRow.source ?? {})),
        original: JSON.parse(JSON.stringify(sourceRow.current)),
        current: JSON.parse(JSON.stringify(sourceRow.current)),
        deleted: false,
        added: true,
      }
      const next = [...currentRows]
      next.splice(index + 1, 0, cloned)
      return recalcDerived(next)
    })
  }, [])

  const recalcDerived = (currentRows) => {
    // 将每行的当前编辑状态转换为 source-like 对象，用于重新计算
    const items = currentRows.map((row) => {
      const c = row.current
      const s = row.source ?? {}
      const deviceTime = c.deviceTime
        ? fromInputDateTime(c.deviceTime)
        : s.datetime ?? s.time ?? ''
      return {
        ...s,
        datetime: deviceTime,
        time: deviceTime,
        basal: c.basal !== '' && c.basal !== undefined ? Number(c.basal) : (s.basal ?? 0),
        bolus: c.bolus !== '' && c.bolus !== undefined ? Number(c.bolus) : (s.bolus ?? 0),
        event: s.event,
        deviceEvent: s.deviceEvent ?? s.event,
        eventPort: s.eventPort,
        eventType: s.eventType,
        eventLevel: s.eventLevel,
        deviceSn: s.deviceSn ?? c.deviceSn ?? '',
      }
    })

    const enriched = applyBolusCalculations(items)

    // 按时间排序后的 enriched 需要映射回原始行顺序
    // applyBolusCalculations 内部会排序，所以用 datetime 和 index 做匹配
    // 但我们需要保持原始顺序，所以直接按顺序传入并取回
    // 注意：applyBolusCalculations 会按时间排序，打乱顺序
    // 改为：我们自己按当前行顺序计算衍生字段

    const timestamps = items.map((item) =>
      parseDateTimeToMs(item.datetime ?? item.time ?? '')
    )

    // 重新计算基础注射量（与下一行的时间差 × basal/160）
    const derivedBasal = new Array(currentRows.length).fill(0)
    const derivedBolus = new Array(currentRows.length).fill('')

    // 按时间排序的索引
    const sortedIndices = timestamps
      .map((ts, i) => ({ ts, i }))
      .sort((a, b) => {
        if (Number.isFinite(a.ts) && Number.isFinite(b.ts)) {
          return a.ts === b.ts ? a.i - b.i : a.ts - b.ts
        }
        if (Number.isFinite(a.ts)) return -1
        if (Number.isFinite(b.ts)) return 1
        return a.i - b.i
      })
      .map(({ i }) => i)

    // 基础注射量
    for (let si = 0; si < sortedIndices.length - 1; si++) {
      const ci = sortedIndices[si]
      const ni = sortedIndices[si + 1]
      const ct = timestamps[ci]
      const nt = timestamps[ni]
      if (Number.isFinite(ct) && Number.isFinite(nt) && nt >= ct) {
        const diffHours = (nt - ct) / MILLISECONDS_PER_HOUR
        const basalValue = Number(items[ci].basal ?? 0)
        const basalAmount = Number.isFinite(basalValue) && diffHours > 0
          ? diffHours * (basalValue / 160)
          : 0
        derivedBasal[ci] = Number.isFinite(basalAmount) ? Number(basalAmount.toFixed(3)) : 0
      }
    }

    // 大剂量总量
    let activeEvent = null
    for (let si = 0; si < sortedIndices.length; si++) {
      const ci = sortedIndices[si]
      const ct = timestamps[ci]
      const bolusValue = Number(items[ci].bolus ?? 0)

      if (!activeEvent && Number.isFinite(ct) && Number.isFinite(bolusValue) && bolusValue !== 0) {
        activeEvent = { origIndex: ci, startTime: ct, bolusValue }
        continue
      }

      if (activeEvent && Number.isFinite(ct) && Number.isFinite(bolusValue) && bolusValue === 0) {
        const durationHours = Math.max((ct - activeEvent.startTime) / MILLISECONDS_PER_HOUR, 0)
        const totalBolus = (activeEvent.bolusValue / 160) * durationHours
        derivedBolus[activeEvent.origIndex] = Number.isFinite(totalBolus)
          ? Number(totalBolus.toFixed(6))
          : 0
        activeEvent = null
      }
    }
    if (activeEvent) {
      derivedBolus[activeEvent.origIndex] = 0
    }

    return currentRows.map((row, i) => ({
      ...row,
      source: {
        ...(row.source ?? {}),
        presetBasalUnitPerHour: derivedBasal[i],
        bolusSum: derivedBolus[i] !== '' ? derivedBolus[i] : (row.source?.bolusSum ?? ''),
      },
      current: {
        ...row.current,
        basalAmount: derivedBasal[i],
        bolusTotal: derivedBolus[i] !== '' ? derivedBolus[i] : (row.current.bolusTotal ?? ''),
      },
      original: {
        ...row.original,
        basalAmount: derivedBasal[i],
        bolusTotal: derivedBolus[i] !== '' ? derivedBolus[i] : row.original.bolusTotal,
      },
    }))
  }

  const resetRowChanges = useCallback((rowId) => {
    setRows((currentRows) => {
      const target = currentRows.find((row) => row.id === rowId)
      if (target?.added) {
        return currentRows.filter((row) => row.id !== rowId)
      }
      return currentRows.map((row) =>
        row.id === rowId
          ? {
              ...row,
              current: { ...row.original },
              deleted: false,
            }
          : row
      )
    })
  }, [])

  const resetAllChanges = useCallback(() => {
    setRows((currentRows) =>
      currentRows.map((row) => ({
        ...row,
        current: { ...row.original },
        deleted: false,
      }))
    )
    setError('')
  }, [setError])

  const exportToExcel = useCallback(
    (baseNameOverride) => {
      if (!changedRows.length) return false

      const baseInput = baseNameOverride || sourceFileBase || 'pump-history'
      const baseName = baseInput.replace(/\s+/g, '_')
    const timestamp = new Date()
      .toISOString()
      .replace(/[:T]/g, '-')
      .slice(0, 19)

      const sheetData = changedRows.map((row) => buildExportRecord(row))
      const headers = EXPORT_FIELD_ORDER

      const worksheet = XLSX.utils.json_to_sheet(sheetData, {
        header: headers,
        skipHeader: false,
      })
      const workbook = XLSX.utils.book_new()
      XLSX.utils.book_append_sheet(workbook, worksheet, 'PumpHistory')
      XLSX.writeFile(workbook, `${baseName}-${timestamp}.xlsx`)
      return true
    },
    [changedRows, sourceFileBase, buildExportRecord]
  )

  const exportToJson = useCallback(
    (baseNameOverride) => {
      if (!changedRows.length) return false

      const baseInput = baseNameOverride || sourceFileBase || 'pump-history'
      const baseName = baseInput.replace(/\s+/g, '_')
      const extension = (sourceFileExtension || 'json').replace(/^\.+/, '')
    const timestamp = new Date()
      .toISOString()
      .replace(/[:T]/g, '-')
      .slice(0, 19)

      const exportPayload = changedRows.map((row) => buildExportRecord(row))

      const blob = new Blob([JSON.stringify(exportPayload, null, 2)], {
        type: 'application/json',
      })

      saveAs(blob, `${baseName}-${timestamp}.${extension}`)
      return true
    },
    [changedRows, sourceFileBase, sourceFileExtension, buildExportRecord]
  )

  const exportAllToExcel = useCallback(
    (baseNameOverride) => {
      if (!rows.length) return false

      const baseInput = baseNameOverride || sourceFileBase || 'pump-history'
      const baseName = baseInput.replace(/\s+/g, '_')
      const timestamp = new Date()
        .toISOString()
        .replace(/[:T]/g, '-')
        .slice(0, 19)

      const sheetData = rows.map((row) => buildExportAllRecord(row))
      const headers = EXPORT_ALL_FIELD_ORDER

      // 1) 最终版 — 无颜色标记
      const wsFinal = XLSX.utils.json_to_sheet(sheetData, {
        header: headers,
        skipHeader: false,
      })
      const wbFinal = XLSX.utils.book_new()
      XLSX.utils.book_append_sheet(wbFinal, wsFinal, 'PumpHistory')
      XLSX.writeFile(wbFinal, `${baseName}-最终版-${timestamp}.xlsx`)

      // 2) 标记版 — 用颜色标记 新增/删除/修改
      const wsMarked = XLSX.utils.json_to_sheet(sheetData, {
        header: headers,
        skipHeader: false,
      })
      const colCount = headers.length
      for (let r = 0; r < rows.length; r++) {
        const row = rows[r]
        const modified = isRowModified(row)
        let bgColor = null
        if (row.added) bgColor = 'C6EFCE'
        else if (row.deleted) bgColor = 'FFC7CE'
        else if (modified) bgColor = 'FFEB9C'
        if (!bgColor) continue

        for (let c = 0; c < colCount; c++) {
          const cellRef = XLSX.utils.encode_cell({ r: r + 1, c })
          const cell = wsMarked[cellRef]
          if (cell) {
            cell.s = {
              fill: { fgColor: { rgb: bgColor }, patternType: 'solid' },
            }
          }
        }
      }
      const wbMarked = XLSX.utils.book_new()
      XLSX.utils.book_append_sheet(wbMarked, wsMarked, 'PumpHistory')
      XLSX.writeFile(wbMarked, `${baseName}-标记版-${timestamp}.xlsx`)

      return true
    },
    [rows, sourceFileBase, buildExportAllRecord, isRowModified]
  )

  const openExportModal = useCallback(() => {
    setExportFileName(sourceFileBase || 'pump-history')
    setIsExportModalOpen(true)
  }, [sourceFileBase])

  const closeExportModal = useCallback(() => {
    setIsExportModalOpen(false)
  }, [])

  const handleConfirmExport = useCallback(() => {
    const baseInput = exportFileName.trim() || sourceFileBase || 'pump-history'
    const hasExcel = exportToExcel(baseInput)
    const hasJson = exportToJson(baseInput)
    if (hasExcel || hasJson) {
      setIsExportModalOpen(false)
    }
  }, [exportFileName, sourceFileBase, exportToExcel, exportToJson])

  const openExportAllModal = useCallback(() => {
    setExportAllFileName(sourceFileBase || 'pump-history')
    setIsExportAllModalOpen(true)
  }, [sourceFileBase])

  const closeExportAllModal = useCallback(() => {
    setIsExportAllModalOpen(false)
  }, [])

  const handleConfirmExportAll = useCallback(() => {
    const baseInput = exportAllFileName.trim() || sourceFileBase || 'pump-history'
    const success = exportAllToExcel(baseInput)
    if (success) {
      setIsExportAllModalOpen(false)
    }
  }, [exportAllFileName, sourceFileBase, exportAllToExcel])

  const closeHighPriorityModal = useCallback(() => {
    setIsHighPriorityModalOpen(false)
  }, [])

  const closeFormatModal = useCallback(() => {
    setIsFormatModalOpen(false)
  }, [])

  return (
    <div className="app-shell">
      <header className="app-header">
      <div>
          <h1>泵体历史数据转换工具</h1>
          <p className="subtitle">
            导入 JSON / Excel 数据，在网页上编辑或标记删除，导出 Excel / JSON。
          </p>
      </div>
        <div className="summary">
          <span>总数：{summary.total}</span>
          <span className="summary-modified">已修改：{summary.modified}</span>
          <span className="summary-deleted">删除：{summary.deleted}</span>
          <span className="summary-added">新增：{summary.added}</span>
      </div>
      </header>

      <section className="controls">
        <label className="upload">
          导入 JSON
          <input
            type="file"
            accept="application/json,.json"
            onChange={handleFileUpload}
          />
        </label>
        <label className="upload">
          导入 Excel
          <input
            type="file"
            accept=".xlsx,.xls"
            onChange={handleExcelUpload}
          />
        </label>
        <button
          type="button"
          onClick={() => setIsFormatModalOpen(true)}
          title="查看 JSON 格式说明"
        >
          JSON 格式说明
        </button>
        <button
          type="button"
          onClick={() => {
            if (!sourceItems.length) return
            setModalError('')
            setIsModalOpen(true)
          }}
          disabled={!sourceItems.length}
        >
          {rows.length ? '调整筛选范围' : '选择加载范围'}
        </button>
     
        <button
          type="button"
          onClick={resetAllChanges}
          disabled={!rows.length || !changedRows.length}
        >
          一键还原
        </button>
        <button
          type="button"
          onClick={openExportAllModal}
          disabled={!rows.length}
        >
          导出全部 Excel
        </button>

        {selectedFileName && (
          <span className="dataset-tag" title={selectedFileName}>
            当前数据：{selectedFileName}
          </span>
        )}
        
        <div className="spacer" />
        <button
          type="button"
          onClick={() => openExportModal()}
          disabled={!rows.length}
        >
          导出 Excel + JSON
        </button>
      </section>

      {isModalOpen && (
        <div className="modal-overlay">
          <div className="modal-content">
            <div className="modal-header">
              <h2>选择加载范围</h2>
              <button
                type="button"
                className="modal-close"
                onClick={closeModal}
              >
                ×
              </button>
      </div>
            <p className="modal-path">
              数据来源：{selectedFileName || '未选择'}
            </p>
            <div className="date-range modal-range">
              <label>
                开始日期
                <input
                  type="date"
                  value={startDate}
                  onChange={(event) => setStartDate(event.target.value)}
                  disabled={!sourceItems.length}
                />
              </label>
              <label>
                结束日期
                <input
                  type="date"
                  value={endDate}
                  onChange={(event) => setEndDate(event.target.value)}
                  disabled={!sourceItems.length}
                />
              </label>
            </div>
            {modalError && <div className="modal-error">{modalError}</div>}
            <div className="modal-actions">
              <button
                type="button"
                className="modal-button secondary"
                onClick={closeModal}
              >
                取消
              </button>
              <button
                type="button"
                className="modal-button primary"
                onClick={applyDateFilter}
                disabled={!sourceItems.length}
              >
                加载数据
              </button>
            </div>
          </div>
        </div>
      )}

      {isExportAllModalOpen && (
        <div className="modal-overlay">
          <div className="modal-content">
            <div className="modal-header">
              <h2>导出全部 Excel</h2>
              <button
                type="button"
                className="modal-close"
                onClick={closeExportAllModal}
              >
                ×
              </button>
            </div>
            <div className="export-summary">
              <span>总记录：{rows.length} 条</span>
              <span>当前筛选：{startDate || endDate ? `${startDate || '最早'} ~ ${endDate || '最晚'}` : '全部数据'}</span>
            </div>
            <p className="export-hint">将同时导出 2 个文件：最终版（无标记）+ 标记版（颜色标记新增/删除/修改）</p>
            <label className="export-name">
              导出文件名（不含时间后缀）
              <input
                type="text"
                value={exportAllFileName}
                onChange={(event) => setExportAllFileName(event.target.value)}
              />
            </label>
            <div className="modal-actions">
              <button
                type="button"
                className="modal-button secondary"
                onClick={closeExportAllModal}
              >
                取消
              </button>
              <button
                type="button"
                className="modal-button primary"
                onClick={handleConfirmExportAll}
                disabled={!rows.length}
              >
                确认导出
              </button>
            </div>
          </div>
        </div>
      )}

      {isHighPriorityModalOpen && (
        <div className="modal-overlay">
          <div className="modal-content wide">
            <div className="modal-header">
              <h2>高优报警列表</h2>
              <button
                type="button"
                className="modal-close"
                onClick={closeHighPriorityModal}
              >
                ×
              </button>
            </div>
            <div className="export-summary">
              <span>
                当前页高优报警：{currentHighPriorityCount} 条（总第{' '}
                {(currentPage - 1) * PAGE_SIZE + 1} ~{' '}
                {Math.min(currentPage * PAGE_SIZE, rows.length)} 条）
              </span>
            </div>
            <div className="export-preview">
              {currentHighPriorityCount > 0 ? (
                <table className="preview-table">
                  <thead>
                    <tr>
                      <th>事件序号</th>
                      <th>设备时间</th>
                      <th>设备事件</th>
                    </tr>
                  </thead>
                  <tbody>
                    {highPriorityRows.map(({ row }) => {
                      const eventValue =
                        row.current.deviceEvent ??
                        row.source?.deviceEvent ??
                        row.source?.event ??
                        ''
                      const eventLabel = getDeviceEventLabel(eventValue)
                      const deviceTime = formatDeviceTimeDisplay(
                        row.current.deviceTime ??
                          row.source?.deviceTime ??
                          row.source?.datetime ??
                          row.source?.time ??
                          ''
                      )
                      const eventIndexValue =
                        row.current.eventIndex ?? row.source?.eventIndex ?? '-'

                      return (
                        <tr key={`hp-${row.id}`}>
                          <td>{eventIndexValue}</td>
                          <td>{deviceTime}</td>
                          <td>{eventLabel || formatDeviceEventValue(eventValue)}</td>
                      
                        </tr>
                      )
                    })}
                  </tbody>
                </table>
              ) : (
                <div className="empty-export">当前页没有高优报警。</div>
              )}
            </div>
            <div className="modal-actions">
              <button
                type="button"
                className="modal-button primary"
                onClick={closeHighPriorityModal}
              >
                知道了
              </button>
            </div>
          </div>
        </div>
      )}

      {isFormatModalOpen && (
        <div className="modal-overlay">
          <div className="modal-content wide">
            <div className="modal-header">
              <h2>JSON 格式说明</h2>
              <button
                type="button"
                className="modal-close"
                onClick={closeFormatModal}
              >
                ×
              </button>
            </div>
            <div className="format-info">
              <h3>示例格式：</h3>
              <pre className="format-example">
{`[
  {
    "id": 5847,
    "deviceSn": "D002AD",//泵体SN
    "datetime": "2025-11-08 09:06:11", //泵体
    "basal": 240, //单位时间步长  基础率速率 = basal / 160 U/hr
    "bolus": 0,   // 单位时间步长 大剂量速率 = bolus / 160 U/hr
    "event": 262144,//事件类型
    "eventIndex": 5,//事件序号
    "eventPort": 4,//事件端口
    "eventType": 0,//事件类型
    "eventLevel": 0,//事件级别
    "remainingCapacity": 38,//剩余电量
    "remainingInsulin": 130,//剩余胰岛素
    "autoMode": true //自动模式
  }
]`}
              </pre>
            </div>
            <div className="modal-actions">
              <button
                type="button"
                className="modal-button primary"
                onClick={closeFormatModal}
              >
                知道了
              </button>
            </div>
          </div>
        </div>
      )}

      {isExportModalOpen && (
        <div className="modal-overlay">
          <div className="modal-content wide">
            <div className="modal-header">
              <h2>导出确认</h2>
              <button
                type="button"
                className="modal-close"
                onClick={closeExportModal}
              >
                ×
              </button>
            </div>
            <div className="export-summary">
              <span>变更记录：{exportStats.total} 条</span>
              <span>新增：{exportStats.added} 条</span>
              <span>删除：{exportStats.deleted} 条</span>
              <span>修改：{exportStats.modified} 条</span>
            </div>
            <label className="export-name">
              导出文件名（不含时间后缀）
              <input
                type="text"
                value={exportFileName}
                onChange={(event) => setExportFileName(event.target.value)}
              />
            </label>
            <div className="export-preview">
              {changedRows.length ? (
                <table className="preview-table">
                  <thead>
                    <tr>
                      <th>事件序号</th>
                      <th>设备时间</th>
                      <th>设备事件</th>
                      <th>基础率</th>
                      <th>大剂量</th>
                      <th>剩余电量</th>
                      <th>剩余胰岛素</th>
                      <th>状态</th>
                      <th>操作</th>
                    </tr>
                  </thead>
                  <tbody>
                    {changedRows.map((row) => {
                      const statusLabel = row.added ? '新增' : row.deleted ? '删除' : '修改'
                      return (
                        <tr key={`export-${row.id}`}>
                          <td>{row.current.eventIndex ?? row.source?.eventIndex ?? '-'}</td>
                          <td>
                            {formatDeviceTimeDisplay(
                              row.current.deviceTime ??
                                row.source?.deviceTime ??
                                row.source?.time ??
                                row.source?.datetime
                            )}
                          </td>
                          <td>{getDeviceEventLabel(row.current.deviceEvent)}</td>
                          <td>{row.current.basal ?? row.source?.basal ?? '-'}</td>
                          <td>{row.current.bolus ?? row.source?.bolus ?? '-'}</td>
                          <td>
                            {row.current.remainingCapacity ??
                              row.source?.remainingCapacity ??
                              '-'}
                          </td>
                          <td>
                            {row.current.remainingInsulin ??
                              row.source?.remainingInsulin ??
                              '-'}
                          </td>
                          <td className={row.added ? 'status-add' : row.deleted ? 'status-delete' : 'status-modify'}>
                            {statusLabel}
                          </td>
                          <td>
                            <button
                              type="button"
                              className="modal-button secondary"
                              onClick={() => resetRowChanges(row.id)}
                            >
                              还原
                            </button>
                          </td>
                        </tr>
                      )
                    })}
                  </tbody>
                </table>
              ) : (
                <div className="empty-export">暂无修改或删除的数据。</div>
              )}
            </div>
            <div className="modal-actions">
              <button
                type="button"
                className="modal-button secondary"
                onClick={closeExportModal}
              >
                取消
              </button>
              <button
                type="button"
                className="modal-button primary"
                onClick={handleConfirmExport}
                disabled={!changedRows.length}
              >
                确认导出
              </button>
            </div>
          </div>
        </div>
      )}

      {error && <div className="error-banner">{error}</div>}

      {rows.length > 0 && (
        <div className="pagination">
          <div className="pagination-info">
            共 {rows.length} 条 · 每页 {PAGE_SIZE} 条 · 第 {currentPage} / {totalPages} 页 ·{' '}
            {currentHighPriorityCount > 0 ? (
              <button
                type="button"
                className="link-button"
                onClick={() => setIsHighPriorityModalOpen(true)}
              >
                当前页·高优报警 {currentHighPriorityCount} 条
              </button>
            ) : (
              <span>当前页高优报警 0 条</span>
            )}
          </div>
          <div className="pagination-actions">
            <button
              type="button"
              onClick={() => handlePageChange(1)}
              disabled={currentPage === 1}
            >
              首页
            </button>
            <button
              type="button"
              onClick={() => handlePageChange(currentPage - 1)}
              disabled={currentPage === 1}
            >
              上一页
            </button>
            <button
              type="button"
              onClick={() => handlePageChange(currentPage + 1)}
              disabled={currentPage === totalPages}
            >
              下一页
            </button>
            <button
              type="button"
              onClick={() => handlePageChange(totalPages)}
              disabled={currentPage === totalPages}
            >
              末页
            </button>
          </div>
        </div>
      )}

      <div className="table-wrapper">
        <table>
          <thead>
            <tr>
              {VISIBLE_COLUMNS.map((column) => {
                const columnClass =
                  column.key === 'deviceEvent'
                    ? 'column-device-event'
                    : column.key === 'deviceTime'
                      ? 'column-device-time'
                      : NARROW_COLUMNS.has(column.key)
                        ? 'column-narrow'
                        : ''
                return (
                  <th key={column.header} className={columnClass}>
                    {column.label ?? column.header}
                  </th>
                )
              })}
              <th>操作</th>
            </tr>
          </thead>
          <tbody>
            {rows.length === 0 ? (
              <tr>
                <td
                  colSpan={VISIBLE_COLUMNS.length + 1}
                  className="empty-state"
                >
                  {sourceItems.length
                    ? '请点击“选择加载范围”设置时间后加载数据。'
                    : '暂无数据，请导入 JSON / Excel 文件或加载示例数据。'}
                </td>
              </tr>
            ) : (
              paginatedRows.map((row) => {
                const modified = isRowModified(row)
                const eventLevelValue = Number(
                  row.current.eventLevel ??
                    row.source?.eventLevel ??
                    row.source?.EVENT_LEVEL ??
                    row.source?.event_level
                )
                const deviceEventValue = Number(
                  row.current.deviceEvent ??
                    row.source?.deviceEvent ??
                    row.source?.event ??
                    row.source?.DEVICE_EVENT
                )
                const isBolusEvent = deviceEventValue === EVENT_CODES.BOLUS

                const rowClass = [
                  'data-row',
                  row.deleted ? 'row-deleted' : '',
                  row.added ? 'row-added' : '',
                  !row.deleted && !row.added && modified ? 'row-modified' : '',
                ]
                  .filter(Boolean)
                  .join(' ')

                const isCollapsed = row.deleted && !expandedDeletedRows.has(row.id)

                if (isCollapsed) {
                  return (
                    <tr key={row.id} className="data-row row-deleted row-collapsed" onClick={() => toggleCollapseRow(row.id)} title="点击展开">
                      {VISIBLE_COLUMNS.map((column) => (
                        <td key={column.key} className={
                          column.key === 'deviceEvent'
                            ? 'column-device-event'
                            : column.key === 'deviceTime'
                              ? 'column-device-time'
                              : NARROW_COLUMNS.has(column.key)
                                ? 'column-narrow'
                                : undefined
                        } />
                      ))}
                      <td className="row-actions">
                        <button
                          type="button"
                          className="danger"
                          onClick={(e) => { e.stopPropagation(); toggleDeleteRow(row.id) }}
                        >
                          取消删除
                        </button>
                      </td>
                    </tr>
                  )
                }

                return (
                  <tr key={row.id} className={rowClass}>
                    {VISIBLE_COLUMNS.map((column) => {
                      const columnClass =
                        column.key === 'deviceEvent'
                          ? 'column-device-event'
                          : column.key === 'deviceTime'
                            ? 'column-device-time'
                            : NARROW_COLUMNS.has(column.key)
                              ? 'column-narrow'
                              : undefined

                      const cellValue = row.current[column.key]
                      const rawValue =
                        cellValue === undefined || cellValue === null
                          ? ''
                          : cellValue
                      const displayValue =
                        column.key === 'deviceEvent'
                          ? getDeviceEventLabel(cellValue)
                          : column.key === 'deviceTime'
                            ? formatDeviceTimeDisplay(rawValue)
                            : rawValue

                      const cellClass =
                        [
                          columnClass,
                          column.key === 'deviceEvent' && eventLevelValue === 2
                            ? 'device-event-critical'
                            : column.key === 'deviceEvent' && isBolusEvent
                              ? 'device-event-bolus'
                            : '',
                        ]
                          .filter(Boolean)
                          .join(' ') || undefined

                      if (column.readOnly && column.key !== 'deviceTime') {
                        return (
                          <td key={column.key} className={cellClass}>
                            {displayValue}
                          </td>
                        )
                      }

                      return (
                        <td
                          key={column.key}
                          className={cellClass}
                        >
                          {column.key === 'deviceTime' ? (
                            <input
                              type="datetime-local"
                              value={displayValue}
                              step="1"
                              onChange={(event) =>
                                handleCellChange(
                                  row.id,
                                  column.key,
                                  event.target.value
                                )
                              }
                              className="table-input wide"
                            />
                          ) : column.editor === 'select' ? (
                            <select
                              value={
                                rawValue === ''
                                  ? ''
                                  : String(rawValue)
                              }
                              onChange={(event) =>
                                handleCellChange(
                                  row.id,
                                  column.key,
                                  event.target.value
                                )
                              }
                              className="table-select"
                            >
                              <option value="">请选择</option>
                              {column.key === 'deviceEvent' &&
                                rawValue !== '' &&
                                !(column.options ?? []).some(
                                  (option) =>
                                    option.value === String(rawValue)
                                ) && (
                                  <option value={String(rawValue)}>
                                    {getDeviceEventLabel(rawValue)}
                                  </option>
                                )}
                              {(column.options ?? []).map((option) => (
                                <option key={option.value} value={option.value}>
                                  {option.label}
                                </option>
                              ))}
                            </select>
                          ) : (
                            <input
                              type={column.type === 'number' ? 'number' : 'text'}
                              value={
                                displayValue === ''
                                  ? ''
                                  : String(displayValue)
                              }
                              onChange={(event) =>
                                handleCellChange(
                                  row.id,
                                  column.key,
                                  event.target.value
                                )
                              }
                              className="table-input"
                            />
                          )}
                        </td>
                      )
                    })}
                    <td className="row-actions">
                      <button
                        type="button"
                        className="add"
                        onClick={() => duplicateRow(row.id)}
                      >
                        新增
                      </button>
                      <button
                        type="button"
                        className="danger"
                        onClick={() => toggleDeleteRow(row.id)}
                      >
                        {row.deleted ? '取消删除' : '删除'}
                      </button>
                      {row.deleted && (
                        <button
                          type="button"
                          onClick={() => toggleCollapseRow(row.id)}
                        >
                          折叠
                        </button>
                      )}
                      <button
                        type="button"
                        onClick={() => resetRowChanges(row.id)}
                        disabled={!row.deleted && !row.added && !modified}
                      >
                        还原
                      </button>
                    </td>
                  </tr>
                )
              })
            )}
          </tbody>
        </table>
      </div>

      {isLoading && (
        <div className="loading-overlay">
          <div className="loading-spinner" />
          <p>数据加载中…</p>
        </div>
      )}
    </div>
  )
}

export default App
