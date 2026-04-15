import { useMemo, useState } from 'react';
import {
  format,
  startOfMonth,
  endOfMonth,
  eachDayOfInterval,
  isToday,
  addMonths,
  subMonths,
  getDay,
  getDate,
  differenceInCalendarDays,
  parse,
  isSameDay,
} from 'date-fns';
import { ru } from 'date-fns/locale';
import { ChevronLeft, ChevronRight, User, Headphones, Filter, X, Check, Download, RotateCcw, Accessibility, FileText } from 'lucide-react';
import * as XLSX from 'xlsx';
import { DeliveryScheduleEntry, ScheduleEntry, VisitHistoryEntry, isScheduled, isDeliveryScheduled, getRouteColor, cn } from '../utils/schedule';

interface CalendarGridProps {
  entries: ScheduleEntry[];
  deliveryScheduleEntries: DeliveryScheduleEntry[];
  visitHistoryEntries: VisitHistoryEntry[];
}

type PointMeta = {
  ClientId: string;
  Branch: string;
  Name: string;
  Address: string;
  DeliveryZone?: string;
  hasSales: boolean;
  hasOperator: boolean;
};

type RouteOption = {
  key: string;
  routeCode: string;
  type: 'Торговый' | 'Оператор';
  label: string;
};

type ProximityFilterKey = 'same-day' | 'pm1' | 'pm2';
type FactOrderProximityKey = 'fo-same-day' | 'fo-pm1' | 'fo-pm2';

type ProximityOption = {
  key: ProximityFilterKey;
  label: string;
  description: string;
};

type FactOrderProximityOption = {
  key: FactOrderProximityKey;
  label: string;
  description: string;
};

type ContourSides = {
  top: boolean;
  right: boolean;
  bottom: boolean;
  left: boolean;
};

const getRouteSelectionKey = (routeCode: string, type: 'Торговый' | 'Оператор') => `${routeCode}::${type}`;
const getRouteTag = (type: 'Торговый' | 'Оператор') => (type === 'Торговый' ? 'ТП' : 'О');
const getRouteLabel = (routeCode: string, type: 'Торговый' | 'Оператор') => `${routeCode}_[${getRouteTag(type)}]`;

const PROXIMITY_OPTIONS: ProximityOption[] = [
  {
    key: 'same-day',
    label: 'день в день',
    description: 'Визит оператора совпадает с визитом торгового в тот же день',
  },
  {
    key: 'pm1',
    label: '+/-1 день',
    description: 'Между визитом оператора и торгового разница ровно 1 день',
  },
  {
    key: 'pm2',
    label: '+/-2 дня',
    description: 'Между визитом оператора и торгового разница ровно 2 дня',
  },
];

const FACT_ORDER_PROXIMITY_OPTIONS: FactOrderProximityOption[] = [
  {
    key: 'fo-same-day',
    label: 'день в день',
    description: 'Фактический заказ ТП совпадает с плановым визитом оператора в тот же день',
  },
  {
    key: 'fo-pm1',
    label: '+/-1 день',
    description: 'Фактический заказ ТП отстоит от планового визита оператора на 1 день',
  },
  {
    key: 'fo-pm2',
    label: '+/-2 дня',
    description: 'Фактический заказ ТП отстоит от планового визита оператора на 2 дня',
  },
];

const EXPORT_ROUTE_COLORS = [
  { bg: '#fee2e2', text: '#dc2626', border: '#fecaca' },
  { bg: '#dbeafe', text: '#2563eb', border: '#bfdbfe' },
  { bg: '#dcfce7', text: '#16a34a', border: '#bbf7d0' },
  { bg: '#fef3c7', text: '#d97706', border: '#fde68a' },
  { bg: '#f3e8ff', text: '#9333ea', border: '#e9d5ff' },
  { bg: '#fce7f3', text: '#db2777', border: '#fbcfe8' },
  { bg: '#e0e7ff', text: '#4f46e5', border: '#c7d2fe' },
  { bg: '#ffedd5', text: '#ea580c', border: '#fed7aa' },
];

const getExportRouteColor = (routeCode: string) => {
  let hash = 0;
  for (let i = 0; i < routeCode.length; i++) {
    hash = routeCode.charCodeAt(i) + ((hash << 5) - hash);
  }

  return EXPORT_ROUTE_COLORS[Math.abs(hash) % EXPORT_ROUTE_COLORS.length];
};

const escapeHtml = (value: unknown) =>
  String(value ?? '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');

const getFactOrderProximityKeyByDiff = (diff: number): FactOrderProximityKey | null => {
  if (diff === 0) return 'fo-same-day';
  if (diff === 1) return 'fo-pm1';
  if (diff === 2) return 'fo-pm2';
  return null;
};

const getProximityKeyByDiff = (diff: number): ProximityFilterKey | null => {
  if (diff === 0) return 'same-day';
  if (diff === 1) return 'pm1';
  if (diff === 2) return 'pm2';
  return null;
};

export function CalendarGrid({ entries, deliveryScheduleEntries, visitHistoryEntries }: CalendarGridProps) {
  const [currentDate, setCurrentDate] = useState(new Date());
  const [filterText, setFilterText] = useState('');
  const [isRouteFilterOpen, setIsRouteFilterOpen] = useState(false);
  const [routeFilterText, setRouteFilterText] = useState('');
  const [selectedRouteKeys, setSelectedRouteKeys] = useState<string[]>([]);
  const [isProximityFilterOpen, setIsProximityFilterOpen] = useState(false);
  const [selectedProximityKeys, setSelectedProximityKeys] = useState<ProximityFilterKey[]>([]);
  const [isFactOrderFilterOpen, setIsFactOrderFilterOpen] = useState(false);
  const [selectedFactOrderKeys, setSelectedFactOrderKeys] = useState<FactOrderProximityKey[]>([]);
  const [isJointCoverageOnly, setIsJointCoverageOnly] = useState(false);

  const monthStart = startOfMonth(currentDate);
  const monthEnd = endOfMonth(currentDate);
  const daysInMonth = eachDayOfInterval({ start: monthStart, end: monthEnd });

  const pointMetaMap = useMemo(() => {
    const map = new Map<string, PointMeta>();
    const addressPriorityMap = new Map<string, { address: string; hasPriority: boolean }>();
    const zonePriorityMap = new Map<string, { zone?: string; hasPriority: boolean }>();

    // First pass: collect addresses and delivery zones with priority (standard import has priority over EasyMerch)
    entries.forEach((entry) => {
      const hasPriority = !entry.EasyMerchDate;
      const existingAddress = addressPriorityMap.get(entry.ClientId);
      const existingZone = zonePriorityMap.get(entry.ClientId);

      if (!existingAddress || (hasPriority && !existingAddress.hasPriority)) {
        addressPriorityMap.set(entry.ClientId, {
          address: entry.Address,
          hasPriority,
        });
      }

      if (
        entry.DeliveryZone &&
        (!existingZone || (hasPriority && !existingZone.hasPriority) || !existingZone.zone)
      ) {
        zonePriorityMap.set(entry.ClientId, {
          zone: entry.DeliveryZone,
          hasPriority,
        });
      }
    });

    // Second pass: build point meta with prioritized addresses and delivery zones
    entries.forEach((entry) => {
      const existing = map.get(entry.ClientId);

      if (!existing) {
        const prioritizedAddress = addressPriorityMap.get(entry.ClientId)?.address || entry.Address;
        const prioritizedZone = zonePriorityMap.get(entry.ClientId)?.zone;
        map.set(entry.ClientId, {
          ClientId: entry.ClientId,
          Branch: entry.Branch,
          Name: entry.Name,
          Address: prioritizedAddress,
          DeliveryZone: prioritizedZone,
          hasSales: entry.Type === 'Торговый',
          hasOperator: entry.Type === 'Оператор',
        });
        return;
      }

      if (!existing.DeliveryZone) {
        existing.DeliveryZone = zonePriorityMap.get(entry.ClientId)?.zone;
      }
      if (entry.Type === 'Торговый') existing.hasSales = true;
      if (entry.Type === 'Оператор') existing.hasOperator = true;
    });

    return map;
  }, [entries]);

  const pointEntriesMap = useMemo(() => {
    const map = new Map<string, ScheduleEntry[]>();

    entries.forEach((entry) => {
      const existing = map.get(entry.ClientId);
      if (existing) {
        existing.push(entry);
      } else {
        map.set(entry.ClientId, [entry]);
      }
    });

    return map;
  }, [entries]);

  const deliveryScheduleMap = useMemo(() => {
    const map = new Map<string, DeliveryScheduleEntry[]>();

    deliveryScheduleEntries.forEach((entry) => {
      const zone = entry.ZoneNumber?.trim();
      if (!zone) return;
      const existing = map.get(zone);
      if (existing) {
        existing.push(entry);
      } else {
        map.set(zone, [entry]);
      }
    });

    return map;
  }, [deliveryScheduleEntries]);

  const visitHistoryMap = useMemo(() => {
    const map = new Map<string, VisitHistoryEntry[]>();

    visitHistoryEntries.forEach((entry) => {
      const normalizedClientId = String(entry.ClientId ?? '').trim();
      const normalizedDate = String(entry.Date ?? '').trim();
      const key = `${normalizedClientId}::${normalizedDate}`;
      const existing = map.get(key);
      if (existing) {
        existing.push(entry);
      } else {
        map.set(key, [entry]);
      }
    });

    return map;
  }, [visitHistoryEntries]);

  const formatDayKey = (day: Date) => format(day, 'dd.MM.yyyy');

  const formatOrderAmount = (amount: number) => new Intl.NumberFormat('ru-RU', {
    maximumFractionDigits: 0,
  }).format(amount);

  const formatDistanceKm = (meters: number) => {
    const km = meters / 1000;
    return `${km >= 10 ? km.toFixed(0) : km.toFixed(1)} км`;
  };

  const formatDistanceBadgeKm = (meters: number) => {
    const km = meters / 1000;
    return `${km >= 10 ? km.toFixed(0) : km.toFixed(1)}`;
  };

  const isPointDeliveryScheduled = (point: PointMeta, day: Date) => {
    const zone = point.DeliveryZone?.trim();
    if (!zone) return false;
    const schedules = deliveryScheduleMap.get(zone);
    if (!schedules || schedules.length === 0) return false;
    return schedules.some((entry) => isDeliveryScheduled(day, entry));
  };

  const getDeliveryScheduleSummary = (point: PointMeta) => {
    const zone = point.DeliveryZone?.trim();
    if (!zone) return '';

    const schedules = deliveryScheduleMap.get(zone);
    if (!schedules || schedules.length === 0) return zone;

    const formatDays = (entry: DeliveryScheduleEntry) => {
      const days: string[] = [];
      if (entry.Monday) days.push('ПН');
      if (entry.Tuesday) days.push('ВТ');
      if (entry.Wednesday) days.push('СР');
      if (entry.Thursday) days.push('ЧТ');
      if (entry.Friday) days.push('ПТ');
      if (entry.Saturday) days.push('СБ');
      if (entry.Sunday) days.push('ВС');
      return days.join(', ');
    };

    const summaries = schedules.map((entry) => {
      const frequency = String(entry.Frequency).padStart(2, '0');
      const days = formatDays(entry);
      return days ? `${zone} [${frequency}] ${days}` : `${zone} [${frequency}]`;
    });

    return summaries.join(' | ');
  };

  const routeOptions = useMemo<RouteOption[]>(() => {
    const seen = new Set<string>();
    const options: RouteOption[] = [];

    entries.forEach((entry) => {
      const key = getRouteSelectionKey(entry.RouteCode, entry.Type);
      if (seen.has(key)) return;
      seen.add(key);

      options.push({
        key,
        routeCode: entry.RouteCode,
        type: entry.Type,
        label: getRouteLabel(entry.RouteCode, entry.Type),
      });
    });

    return options.sort((a, b) => {
      const byRoute = a.routeCode.localeCompare(b.routeCode, 'ru');
      if (byRoute !== 0) return byRoute;
      if (a.type === b.type) return 0;
      return a.type === 'Торговый' ? -1 : 1;
    });
  }, [entries]);

  const filteredRouteOptions = useMemo(() => {
    if (!routeFilterText.trim()) return routeOptions;
    const query = routeFilterText.toLowerCase();

    return routeOptions.filter((option) => {
      const typeLabel = option.type === 'Торговый' ? 'торговый тп' : 'оператор о';
      return (
        option.label.toLowerCase().includes(query) ||
        option.routeCode.toLowerCase().includes(query) ||
        typeLabel.includes(query)
      );
    });
  }, [routeFilterText, routeOptions]);

  const routeFilteredPointIds = useMemo(() => {
    if (selectedRouteKeys.length === 0) return null;

    const selected = new Set(selectedRouteKeys);
    return new Set(
      entries
        .filter((entry) => selected.has(getRouteSelectionKey(entry.RouteCode, entry.Type)))
        .map((entry) => entry.ClientId),
    );
  }, [entries, selectedRouteKeys]);

  const visiblePoints = useMemo(() => {
    const allPoints = Array.from(pointMetaMap.values());
    const routeScopedPoints = routeFilteredPointIds
      ? allPoints.filter((point) => routeFilteredPointIds.has(point.ClientId))
      : allPoints;

    return routeScopedPoints.sort((a, b) => a.Name.localeCompare(b.Name, 'ru'));
  }, [pointMetaMap, routeFilteredPointIds]);

  const searchedPoints = useMemo(() => {
    if (!filterText.trim()) return visiblePoints;
    const lowerFilter = filterText.toLowerCase();

    return visiblePoints.filter((point) =>
      point.Name.toLowerCase().includes(lowerFilter) ||
      point.Address.toLowerCase().includes(lowerFilter) ||
      point.ClientId.toLowerCase().includes(lowerFilter) ||
      point.Branch.toLowerCase().includes(lowerFilter)
    );
  }, [visiblePoints, filterText]);

  const pointProximityMap = useMemo(() => {
    const map = new Map<string, Set<ProximityFilterKey>>();

    pointMetaMap.forEach((point, clientId) => {
      if (!point.hasSales || !point.hasOperator) {
        map.set(clientId, new Set());
        return;
      }

      const pointEntries = pointEntriesMap.get(clientId) ?? [];
      const salesDates = daysInMonth.filter((day) =>
        pointEntries.some((entry) => entry.Type === 'Торговый' && isScheduled(day, entry)),
      );
      const operatorDates = daysInMonth.filter((day) =>
        pointEntries.some((entry) => entry.Type === 'Оператор' && isScheduled(day, entry)),
      );

      const matches = new Set<ProximityFilterKey>();

      salesDates.forEach((salesDay) => {
        operatorDates.forEach((operatorDay) => {
          const diff = Math.abs(differenceInCalendarDays(operatorDay, salesDay));
          const matchKey = getProximityKeyByDiff(diff);
          if (matchKey) matches.add(matchKey);
        });
      });

      map.set(clientId, matches);
    });

    return map;
  }, [daysInMonth, pointEntriesMap, pointMetaMap]);

  const pointFactOrderProximityMap = useMemo(() => {
    const map = new Map<string, Set<FactOrderProximityKey>>();

    pointMetaMap.forEach((point, clientId) => {
      if (!point.hasOperator) {
        map.set(clientId, new Set());
        return;
      }

      const pointEntries = pointEntriesMap.get(clientId) ?? [];
      const operatorDates = daysInMonth.filter((day) =>
        pointEntries.some((entry) => entry.Type === 'Оператор' && isScheduled(day, entry)),
      );

      const factOrderDates = visitHistoryEntries
        .filter((entry) => String(entry.ClientId).trim() === clientId && (entry.OrderAmountRub ?? 0) > 0)
        .map((entry) => parse(String(entry.Date).trim(), 'dd.MM.yyyy', new Date()))
        .filter((date) => !Number.isNaN(date.getTime()));

      const matches = new Set<FactOrderProximityKey>();

      factOrderDates.forEach((factDate) => {
        operatorDates.forEach((operatorDay) => {
          const diff = Math.abs(differenceInCalendarDays(operatorDay, factDate));
          const matchKey = getFactOrderProximityKeyByDiff(diff);
          if (matchKey) matches.add(matchKey);
        });
      });

      map.set(clientId, matches);
    });

    return map;
  }, [daysInMonth, pointEntriesMap, pointMetaMap, visitHistoryEntries]);

  const selectedRouteSet = useMemo(() => new Set(selectedRouteKeys), [selectedRouteKeys]);
  const selectedProximitySet = useMemo(() => new Set(selectedProximityKeys), [selectedProximityKeys]);
  const selectedFactOrderSet = useMemo(() => new Set(selectedFactOrderKeys), [selectedFactOrderKeys]);

  const proximityContourMap = useMemo(() => {
    const map = new Map<string, ContourSides>();
    if (selectedProximitySet.size === 0) return map;

    const setContourSide = (clientId: string, day: Date, sides: Partial<ContourSides>) => {
      const key = `${clientId}::${formatDayKey(day)}`;
      const current = map.get(key) ?? { top: false, right: false, bottom: false, left: false };
      map.set(key, {
        top: current.top || !!sides.top,
        right: current.right || !!sides.right,
        bottom: current.bottom || !!sides.bottom,
        left: current.left || !!sides.left,
      });
    };

    pointMetaMap.forEach((point, clientId) => {
      if (!point.hasSales || !point.hasOperator) return;

      const pointEntries = pointEntriesMap.get(clientId) ?? [];
      const salesDates = daysInMonth.filter((day) =>
        pointEntries.some((entry) => entry.Type === 'Торговый' && isScheduled(day, entry)),
      );
      const operatorDates = daysInMonth.filter((day) =>
        pointEntries.some((entry) => entry.Type === 'Оператор' && isScheduled(day, entry)),
      );

      salesDates.forEach((salesDay) => {
        operatorDates.forEach((operatorDay) => {
          const diff = Math.abs(differenceInCalendarDays(operatorDay, salesDay));
          const matchKey = getProximityKeyByDiff(diff);
          if (!matchKey || !selectedProximitySet.has(matchKey)) return;

          const startDate = differenceInCalendarDays(operatorDay, salesDay) >= 0 ? salesDay : operatorDay;
          const endDate = differenceInCalendarDays(operatorDay, salesDay) >= 0 ? operatorDay : salesDay;
          const contourDays = daysInMonth.filter((day) =>
            differenceInCalendarDays(day, startDate) >= 0 && differenceInCalendarDays(endDate, day) >= 0,
          );

          contourDays.forEach((day) => {
            setContourSide(clientId, day, {
              top: true,
              bottom: true,
              left: isSameDay(day, startDate),
              right: isSameDay(day, endDate),
            });
          });
        });
      });
    });

    return map;
  }, [daysInMonth, pointEntriesMap, pointMetaMap, selectedProximitySet]);

  const factOrderContourMap = useMemo(() => {
    const map = new Map<string, ContourSides>();
    if (selectedFactOrderSet.size === 0) return map;

    const setContourSide = (clientId: string, day: Date, sides: Partial<ContourSides>) => {
      const key = `${clientId}::${formatDayKey(day)}`;
      const current = map.get(key) ?? { top: false, right: false, bottom: false, left: false };
      map.set(key, {
        top: current.top || !!sides.top,
        right: current.right || !!sides.right,
        bottom: current.bottom || !!sides.bottom,
        left: current.left || !!sides.left,
      });
    };

    pointMetaMap.forEach((point, clientId) => {
      if (!point.hasOperator) return;

      const pointEntries = pointEntriesMap.get(clientId) ?? [];
      const operatorDates = daysInMonth.filter((day) =>
        pointEntries.some((entry) => entry.Type === 'Оператор' && isScheduled(day, entry)),
      );

      if (operatorDates.length === 0) return;

      const factOrderDates = visitHistoryEntries
        .filter((entry) => String(entry.ClientId).trim() === clientId && (entry.OrderAmountRub ?? 0) > 0)
        .map((entry) => parse(String(entry.Date).trim(), 'dd.MM.yyyy', new Date()))
        .filter((date) => !Number.isNaN(date.getTime()));

      factOrderDates.forEach((factDate) => {
        operatorDates.forEach((operatorDay) => {
          const diff = Math.abs(differenceInCalendarDays(operatorDay, factDate));
          const matchKey = getFactOrderProximityKeyByDiff(diff);
          if (!matchKey || !selectedFactOrderSet.has(matchKey)) return;

          const startDate = differenceInCalendarDays(operatorDay, factDate) >= 0 ? factDate : operatorDay;
          const endDate = differenceInCalendarDays(operatorDay, factDate) >= 0 ? operatorDay : factDate;
          const contourDays = daysInMonth.filter((day) =>
            differenceInCalendarDays(day, startDate) >= 0 && differenceInCalendarDays(endDate, day) >= 0,
          );

          contourDays.forEach((day) => {
            setContourSide(clientId, day, {
              top: true,
              bottom: true,
              left: isSameDay(day, startDate),
              right: isSameDay(day, endDate),
            });
          });
        });
      });
    });

    return map;
  }, [daysInMonth, pointEntriesMap, pointMetaMap, selectedFactOrderSet, visitHistoryEntries]);

  const selectedRouteLabels = useMemo(
    () => routeOptions.filter((option) => selectedRouteSet.has(option.key)).map((option) => option.label),
    [routeOptions, selectedRouteSet],
  );

  const selectedProximityLabels = useMemo(
    () => PROXIMITY_OPTIONS.filter((option) => selectedProximitySet.has(option.key)).map((option) => option.label),
    [selectedProximitySet],
  );

  const selectedFactOrderLabels = useMemo(
    () => FACT_ORDER_PROXIMITY_OPTIONS.filter((option) => selectedFactOrderSet.has(option.key)).map((option) => option.label),
    [selectedFactOrderSet],
  );

  const hasSelectionFilters = selectedRouteLabels.length > 0 || selectedProximityLabels.length > 0 || selectedFactOrderLabels.length > 0 || isJointCoverageOnly;

  const filteredPoints = useMemo(() => {
    const proximityFiltered = selectedProximitySet.size === 0
      ? searchedPoints
      : searchedPoints.filter((point) => {
          const proximityMatches = pointProximityMap.get(point.ClientId);
          if (!proximityMatches || proximityMatches.size === 0) return false;

          for (const key of selectedProximitySet) {
            if (proximityMatches.has(key)) return true;
          }

          return false;
        });

    const factOrderFiltered = selectedFactOrderSet.size === 0
      ? proximityFiltered
      : proximityFiltered.filter((point) => {
          const proximityMatches = pointFactOrderProximityMap.get(point.ClientId);
          if (!proximityMatches || proximityMatches.size === 0) return false;

          for (const key of selectedFactOrderSet) {
            if (proximityMatches.has(key)) return true;
          }

          return false;
        });

    if (!isJointCoverageOnly) {
      return factOrderFiltered;
    }

    return factOrderFiltered.filter((point) => point.hasSales && point.hasOperator);
  }, [searchedPoints, selectedProximitySet, pointProximityMap, selectedFactOrderSet, pointFactOrderProximityMap, isJointCoverageOnly]);

  const getProximityLabelsForPoint = (clientId: string) => {
    const matches = pointProximityMap.get(clientId);
    if (!matches || matches.size === 0) return '';

    return PROXIMITY_OPTIONS.filter((option) => matches.has(option.key))
      .map((option) => option.label)
      .join(', ');
  };

  const getProximityContourSides = (clientId: string, day: Date) =>
    proximityContourMap.get(`${String(clientId).trim()}::${formatDayKey(day)}`);

  const getFactOrderContourSides = (clientId: string, day: Date) =>
    factOrderContourMap.get(`${String(clientId).trim()}::${formatDayKey(day)}`);

  const getContourClassName = (baseClassName: string, sides?: ContourSides) => {
    if (!sides) return '';

    return cn(
      baseClassName,
      sides.top && `${baseClassName}-top`,
      sides.right && `${baseClassName}-right`,
      sides.bottom && `${baseClassName}-bottom`,
      sides.left && `${baseClassName}-left`,
    );
  };

  const exportRows = useMemo(() => {
    return filteredPoints.map((point) => {
      const pointEntries = pointEntriesMap.get(point.ClientId) ?? [];
      const salesRoutes = Array.from(
        new Set(
          pointEntries
            .filter((entry) => entry.Type === 'Торговый')
            .map((entry) => getRouteLabel(entry.RouteCode, entry.Type)),
        ),
      ).join(', ');
      const operatorRoutes = Array.from(
        new Set(
          pointEntries
            .filter((entry) => entry.Type === 'Оператор')
            .map((entry) => getRouteLabel(entry.RouteCode, entry.Type)),
        ),
      ).join(', ');

      return {
        'код точки': point.ClientId,
        Наименование: point.Name,
        Адрес: point.Address,
        'номер маршрута Торгового': salesRoutes,
        'номер маршрута Оператора': operatorRoutes,
        'Близость визитов': getProximityLabelsForPoint(point.ClientId),
      };
    });
  }, [filteredPoints, pointEntriesMap, pointProximityMap]);

  const exportFilteredPointsToExcel = () => {
    const worksheet = XLSX.utils.json_to_sheet(exportRows);
    const workbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(workbook, worksheet, 'Точки');

    worksheet['!cols'] = [
      { wch: 16 },
      { wch: 40 },
      { wch: 60 },
      { wch: 28 },
      { wch: 28 },
      { wch: 24 },
    ];

    const fileMonth = format(currentDate, 'yyyy-MM');
    XLSX.writeFile(workbook, `точки_${fileMonth}.xlsx`);
  };

  const buildShareCellHtml = (pointCode: string, day: Date) => {
    const normalizedPointCode = String(pointCode).trim();
    const scheduledEvents = entries.filter(
      (entry) => String(entry.ClientId).trim() === normalizedPointCode && isScheduled(day, entry),
    );
    const visitFacts = visitHistoryMap.get(`${normalizedPointCode}::${formatDayKey(day)}`) ?? [];

    if (scheduledEvents.length === 0 && visitFacts.length === 0) return '';

    const salesReps = scheduledEvents.filter((entry) => entry.Type === 'Торговый');
    const operators = scheduledEvents.filter((entry) => entry.Type === 'Оператор');
    const factRoutes = Array.from(new Set(visitFacts.map((fact) => fact.RouteCode).filter(Boolean)));
    const hasOrder = visitFacts.some((fact) => (fact.OrderAmountRub ?? 0) > 0);
    const totalOrderAmount = Math.round(visitFacts.reduce((sum, fact) => sum + (fact.OrderAmountRub ?? 0), 0));
    const deviations = visitFacts
      .map((fact) => fact.CoordinateDeviationMeters)
      .filter((meters): meters is number => typeof meters === 'number' && Number.isFinite(meters));
    const hasDeviationWithin300 = deviations.some((meters) => meters <= 300);
    const deviationsOver300 = deviations.filter((meters) => meters > 300);
    const maxDeviationOver300 = !hasDeviationWithin300 && deviationsOver300.length > 0
      ? Math.max(...deviationsOver300)
      : null;
    const minDeviation = deviations.length > 0 ? Math.min(...deviations) : null;

    const salesBadges = salesReps.map((rep) => {
      const color = getExportRouteColor(rep.RouteCode);
      return `<span class="marker sales-marker" style="background:${color.bg};border-color:${color.border};color:${color.text}" title="${escapeHtml(`Торговый: ${getRouteLabel(rep.RouteCode, rep.Type)}`)}">ТП</span>`;
    });

    const operatorBadges = operators.map((op) =>
      `<span class="marker operator-marker" title="${escapeHtml(`Оператор: ${getRouteLabel(op.RouteCode, op.Type)}`)}">О</span>`,
    );

    const factBadges = visitFacts.length > 0
      ? [
          `<span class="marker fact-marker ${hasOrder ? 'fact-order' : ''}" title="${escapeHtml([
            `Факт: ${factRoutes.join(', ') || 'маршрут не указан'}`,
            hasOrder ? `Сумма отгрузки: ${formatOrderAmount(totalOrderAmount)} руб.` : 'Сумма отгрузки: нет',
            minDeviation !== null ? `Отклонение: ${formatDistanceKm(minDeviation)}` : 'Отклонение: нет данных',
          ].join(' | '))}">${hasOrder ? '₽' : 'Ф'}</span>`,
        ]
      : [];

    const deviationBadge = maxDeviationOver300
      ? `<span class="marker deviation-marker" title="${escapeHtml(`Отклонение: ${formatDistanceKm(maxDeviationOver300)}`)}">${escapeHtml(formatDistanceBadgeKm(maxDeviationOver300))}</span>`
      : '';

    return `<div class="cell-events">${[...salesBadges, ...operatorBadges, ...factBadges, deviationBadge].filter(Boolean).join('')}</div>`;
  };

  const exportCurrentViewToHtml = () => {
    const monthLabel = format(currentDate, 'LLLL yyyy', { locale: ru });
    const exportedAt = format(new Date(), 'dd.MM.yyyy HH:mm');
    const filterItems = [
      filterText.trim() ? `Поиск: ${filterText.trim()}` : '',
      selectedRouteLabels.length > 0 ? `Маршруты: ${selectedRouteLabels.join(', ')}` : '',
      selectedProximityLabels.length > 0 ? `Близость ПВ: ${selectedProximityLabels.join(', ')}` : '',
      selectedFactOrderLabels.length > 0 ? `Близость ФЗ: ${selectedFactOrderLabels.join(', ')}` : '',
      isJointCoverageOnly ? 'Покрытие: общее' : '',
    ].filter(Boolean);

    const headCells = daysInMonth.map((day) => {
      const isWeekend = [0, 6].includes(getDay(day));
      return `
        <th class="${isWeekend ? 'weekend' : ''}">
          <div class="weekday">${escapeHtml(format(day, 'EEEEEE', { locale: ru }))}</div>
          <div class="daynum">${getDate(day)}</div>
        </th>
      `;
    }).join('');

    const bodyRows = filteredPoints.length === 0
      ? `<tr><td class="empty" colspan="${daysInMonth.length + 1}">Нет данных для отображения.</td></tr>`
      : filteredPoints.map((point) => {
          const rows = daysInMonth.map((day) => {
            const hasDelivery = isPointDeliveryScheduled(point, day);
            const isWeekend = [0, 6].includes(getDay(day));
            const proximityContourSides = getProximityContourSides(point.ClientId, day);
            const factOrderContourSides = getFactOrderContourSides(point.ClientId, day);
            const cellClass = [
              hasDelivery ? 'delivery' : '',
              isWeekend ? 'weekend-cell' : '',
              getContourClassName('proximity-contour', proximityContourSides),
              getContourClassName('fact-order-contour', factOrderContourSides),
            ].filter(Boolean).join(' ');

            return `<td class="${cellClass}">${buildShareCellHtml(point.ClientId, day)}</td>`;
          }).join('');

          return `
            <tr>
              <td class="point">
                <div class="branch">${escapeHtml(point.Branch)}</div>
                <div class="name">${escapeHtml(point.Name)}</div>
                <div class="address">${escapeHtml(point.Address)}</div>
                <div class="meta">
                  <span>${escapeHtml(point.ClientId)}</span>
                  ${point.hasSales ? '<span class="mini sales">ТП</span>' : ''}
                  ${point.hasOperator ? '<span class="mini operator">О</span>' : ''}
                  ${point.DeliveryZone ? `<span class="zone" title="${escapeHtml(getDeliveryScheduleSummary(point))}">${escapeHtml(getDeliveryScheduleSummary(point))}</span>` : ''}
                </div>
              </td>
              ${rows}
            </tr>
          `;
        }).join('');

    const html = `<!doctype html>
<html lang="ru">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>Календарь маршрутов - ${escapeHtml(monthLabel)}</title>
  <style>
    :root { color-scheme: light; font-family: Inter, Arial, sans-serif; color: #111827; background: #edf2f7; }
    * { box-sizing: border-box; }
    html, body { height: 100%; }
    body { margin: 0; overflow: hidden; background: #edf2f7; }
    .screen { display: grid; grid-template-rows: auto 1fr; height: 100vh; padding: 12px; gap: 10px; }
    header { min-height: 0; background: #ffffff; border: 1px solid #d7dee8; border-radius: 8px; padding: 12px 14px; box-shadow: 0 6px 18px rgba(15, 23, 42, 0.08); }
    .title-row { display: flex; align-items: flex-start; justify-content: space-between; gap: 16px; }
    h1 { margin: 0; font-size: 20px; line-height: 1.15; letter-spacing: 0; }
    .exported { flex-shrink: 0; color: #64748b; font-size: 12px; text-align: right; }
    .summary { display: flex; flex-wrap: wrap; gap: 6px; align-items: center; margin-top: 8px; color: #475569; font-size: 12px; }
    .chip { display: inline-flex; align-items: center; max-width: 360px; overflow: hidden; text-overflow: ellipsis; white-space: nowrap; border: 1px solid #bfdbfe; background: #eff6ff; color: #1d4ed8; border-radius: 999px; padding: 3px 8px; font-weight: 700; }
    .legend { display: flex; flex-wrap: wrap; gap: 6px 12px; margin-top: 8px; font-size: 11px; color: #64748b; }
    .legend-mark { display: inline-flex; height: 18px; min-width: 18px; align-items: center; justify-content: center; border-radius: 5px; padding: 0 5px; font-size: 10px; font-weight: 900; }
    .wrap { min-height: 0; }
    .table-wrap { height: 100%; overflow: auto; border: 1px solid #d7dee8; border-radius: 8px; background: #ffffff; box-shadow: 0 10px 26px rgba(15, 23, 42, 0.1); }
    table { width: 100%; min-width: 1180px; border-collapse: separate; border-spacing: 0; table-layout: fixed; }
    th, td { border-right: 1px solid #e2e8f0; border-bottom: 1px solid #e2e8f0; }
    th { position: sticky; top: 0; z-index: 3; width: 36px; min-width: 36px; height: 34px; padding: 2px; background: #ffffff; text-align: center; }
    th.first { left: 0; z-index: 4; width: 230px; min-width: 230px; text-align: left; background: #f8fafc; color: #475569; font-size: 11px; text-transform: uppercase; }
    th.weekend { background: #fff1f2; }
    .weekday { color: #94a3b8; font-size: 8px; text-transform: capitalize; line-height: 1; }
    .daynum { color: #0f172a; font-size: 12px; font-weight: 900; line-height: 1.15; }
    td { width: 36px; min-width: 36px; height: 44px; padding: 2px; vertical-align: middle; background: #ffffff; }
    td.proximity-contour { position: relative; z-index: 1; }
    td.proximity-contour-top { border-top: 3px solid #f59e0b; }
    td.proximity-contour-right { border-right: 3px solid #f59e0b; }
    td.proximity-contour-bottom { border-bottom: 3px solid #f59e0b; }
    td.proximity-contour-left { border-left: 3px solid #f59e0b; }
    td.fact-order-contour { position: relative; z-index: 1; }
    td.fact-order-contour-top { border-top: 3px solid #ef4444; }
    td.fact-order-contour-right { border-right: 3px solid #ef4444; }
    td.fact-order-contour-bottom { border-bottom: 3px solid #ef4444; }
    td.fact-order-contour-left { border-left: 3px solid #ef4444; }
    td.weekend-cell { background: #fff7f7; }
    td.delivery { background: #eefbf3; }
    td.point { position: sticky; left: 0; z-index: 2; width: 230px; min-width: 230px; background: #ffffff; padding: 5px 7px; }
    .branch { color: #64748b; font-size: 8px; font-weight: 800; text-transform: uppercase; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .name { margin-top: 1px; color: #0f172a; font-size: 11px; font-weight: 900; line-height: 1.15; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .address { margin-top: 1px; color: #64748b; font-size: 8px; line-height: 1.15; white-space: nowrap; overflow: hidden; text-overflow: ellipsis; }
    .meta { display: flex; flex-wrap: nowrap; gap: 3px; align-items: center; margin-top: 3px; color: #64748b; font-size: 8px; font-family: ui-monospace, SFMono-Regular, Menlo, monospace; overflow: hidden; }
    .zone { color: #15803d; font-family: Inter, Arial, sans-serif; font-weight: 700; }
    .mini { border-radius: 4px; padding: 1px 4px; font-family: Inter, Arial, sans-serif; font-weight: 800; }
    .mini.sales { background: #fee2e2; color: #dc2626; }
    .mini.operator { background: #dcfce7; color: #16a34a; }
    .cell-events { display: flex; width: 100%; height: 100%; align-items: center; justify-content: center; gap: 2px; flex-wrap: wrap; overflow: hidden; }
    .marker { display: inline-flex; width: 18px; height: 18px; align-items: center; justify-content: center; flex: 0 0 auto; border: 1px solid #e5e7eb; border-radius: 5px; font-size: 8px; line-height: 1; font-weight: 900; letter-spacing: 0; box-shadow: 0 1px 2px rgba(15, 23, 42, 0.08); }
    .operator-marker { background: #dcfce7; border-color: #86efac; color: #16a34a; }
    .fact-marker { background: #f8fafc; border-color: #94a3b8; color: #334155; }
    .fact-order { background: #fef3c7; border-color: #f59e0b; color: #92400e; font-size: 11px; }
    .deviation-marker { min-width: 22px; width: auto; padding: 0 3px; background: #ffe4e6; border-color: #fb7185; color: #9f1239; }
    .empty { width: auto; height: 120px; text-align: center; color: #6b7280; font-size: 14px; }
    @media (max-width: 1300px) {
      .screen { padding: 8px; gap: 8px; }
      header { padding: 10px 12px; }
      h1 { font-size: 18px; }
      table { min-width: 1040px; }
      th, td { width: 32px; min-width: 32px; }
      td { height: 40px; }
      .marker { width: 16px; height: 16px; font-size: 7px; }
      .deviation-marker { min-width: 20px; }
    }
    @media print {
      body { overflow: visible; background: #ffffff; }
      .screen { display: block; height: auto; padding: 0; }
      header, th, td.point { position: static; }
      .wrap { padding: 0; height: auto; }
      .table-wrap { box-shadow: none; border: 0; }
    }
  </style>
</head>
<body>
  <div class="screen">
    <header>
      <div class="title-row">
        <h1>Календарь маршрутов: ${escapeHtml(monthLabel)}</h1>
        <div class="exported">Экспортировано<br>${escapeHtml(exportedAt)}</div>
      </div>
      <div class="summary">
        <span>Видимых точек: <strong>${filteredPoints.length}</strong></span>
        <span>Маршрутных записей: <strong>${entries.length}</strong></span>
        <span>Зон доставки: <strong>${deliveryScheduleEntries.length}</strong></span>
        <span>Фактов визитов: <strong>${visitHistoryEntries.length}</strong></span>
        ${filterItems.length > 0 ? filterItems.map((item) => `<span class="chip">${escapeHtml(item)}</span>`).join('') : '<span class="chip">Фильтры не применены</span>'}
      </div>
      <div class="legend">
        <span><span class="legend-mark" style="background:#fee2e2;color:#dc2626">ТП</span> план торгового</span>
        <span><span class="legend-mark" style="background:#dcfce7;color:#16a34a">О</span> план оператора</span>
        <span><span class="legend-mark" style="background:#f8fafc;color:#334155;border:1px solid #94a3b8">Ф</span> факт</span>
        <span><span class="legend-mark" style="background:#fef3c7;color:#92400e">₽</span> факт с заказом</span>
        <span><span class="legend-mark" style="background:#eefbf3;color:#15803d">Д</span> день доставки</span>
      </div>
    </header>
    <main class="wrap">
      <div class="table-wrap">
        <table>
          <thead>
            <tr>
              <th class="first">Точка доставки</th>
              ${headCells}
            </tr>
          </thead>
          <tbody>
            ${bodyRows}
          </tbody>
        </table>
      </div>
    </main>
  </div>
</body>
</html>`;

    const blob = new Blob([html], { type: 'text/html;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = `календарь_маршрутов_${format(currentDate, 'yyyy-MM')}.html`;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(url);
  };

  const prevMonth = () => setCurrentDate(subMonths(currentDate, 1));
  const nextMonth = () => setCurrentDate(addMonths(currentDate, 1));

  const toggleRouteSelection = (routeKey: string) => {
    setSelectedRouteKeys((current) =>
      current.includes(routeKey)
        ? current.filter((key) => key !== routeKey)
        : [...current, routeKey],
    );
  };

  const clearRouteSelection = () => {
    setSelectedRouteKeys([]);
    setRouteFilterText('');
  };

  const selectAllFilteredRoutes = () => {
    setSelectedRouteKeys((current) => {
      const merged = new Set(current);
      filteredRouteOptions.forEach((option) => merged.add(option.key));
      return Array.from(merged);
    });
  };

  const toggleProximitySelection = (key: ProximityFilterKey) => {
    setSelectedProximityKeys((current) =>
      current.includes(key)
        ? current.filter((item) => item !== key)
        : [...current, key],
    );
  };

  const clearProximitySelection = () => {
    setSelectedProximityKeys([]);
  };

  const toggleFactOrderSelection = (key: FactOrderProximityKey) => {
    setSelectedFactOrderKeys((current) =>
      current.includes(key)
        ? current.filter((item) => item !== key)
        : [...current, key],
    );
  };

  const clearFactOrderSelection = () => {
    setSelectedFactOrderKeys([]);
  };

  const clearSelectionFilters = () => {
    setSelectedRouteKeys([]);
    setRouteFilterText('');
    setSelectedProximityKeys([]);
    setSelectedFactOrderKeys([]);
    setIsJointCoverageOnly(false);
    setIsRouteFilterOpen(false);
    setIsProximityFilterOpen(false);
    setIsFactOrderFilterOpen(false);
  };

  const getCellContent = (pointCode: string, day: Date) => {
    const normalizedPointCode = String(pointCode).trim();
    const scheduledEvents = entries.filter(
      (entry) => String(entry.ClientId).trim() === normalizedPointCode && isScheduled(day, entry),
    );

    const visitFacts = visitHistoryMap.get(`${normalizedPointCode}::${formatDayKey(day)}`) ?? [];

    if (scheduledEvents.length === 0 && visitFacts.length === 0) return null;

    const salesReps = scheduledEvents.filter((entry) => entry.Type === 'Торговый');
    const operators = scheduledEvents.filter((entry) => entry.Type === 'Оператор');
    const factRoutes = Array.from(new Set(visitFacts.map((fact) => fact.RouteCode).filter(Boolean)));
    const primaryFactRoute = factRoutes[0] ?? '';
    const hasAnyPlan = salesReps.length > 0 || operators.length > 0;
    const hasFact = visitFacts.length > 0;
    const hasOrder = visitFacts.some((fact) => (fact.OrderAmountRub ?? 0) > 0);
    const totalOrderAmount = Math.round(visitFacts.reduce((sum, fact) => sum + (fact.OrderAmountRub ?? 0), 0));
    const deviations = visitFacts
      .map((fact) => fact.CoordinateDeviationMeters)
      .filter((meters): meters is number => typeof meters === 'number' && Number.isFinite(meters));
    const hasDeviationWithin300 = deviations.some((meters) => meters <= 300);
    const deviationsOver300 = deviations.filter((meters) => meters > 300);
    const maxDeviationOver300 = !hasDeviationWithin300 && deviationsOver300.length > 0
      ? Math.max(...deviationsOver300)
      : null;
    const minDeviation = deviations.length > 0 ? Math.min(...deviations) : null;

    return (
      <div className="relative flex items-center justify-center w-full h-full min-h-[52px] overflow-hidden px-0.5 py-0.5">
        {salesReps.map((rep, idx) => {
          const colorClass = getRouteColor(rep.RouteCode);
          const hasOverlap = salesReps.length > 1 || operators.length > 0 || hasFact;
          const hasTopBadges = !!maxDeviationOver300;

          return (
              <div
                key={`rep-${idx}`}
                className={cn(
                  'absolute z-[1] flex h-[24px] w-[24px] flex-col items-center justify-center rounded-md border p-[1px] shadow-sm transform transition-transform hover:scale-110 hover:z-[2] cursor-help',
                  colorClass,
                  hasOverlap ? (hasTopBadges ? 'top-4 left-0' : 'top-0 left-0') : hasTopBadges ? 'top-4' : '',
                )}
                title={`Торговый: ${getRouteLabel(rep.RouteCode, rep.Type)}`}
              >
                <User size={10} className="mx-auto" />
                <span className="text-[6px] block text-center leading-none mt-0.5 font-bold">план</span>
              </div>
          );
        })}

        {operators.map((op, idx) => {
          const hasOverlap = salesReps.length > 0 || operators.length > 1 || hasFact;
          const hasTopBadges = !!maxDeviationOver300;

          return (
            <div
              key={`op-${idx}`}
              className={cn(
                'absolute z-[1] flex h-[24px] w-[24px] flex-col items-center justify-center rounded-md border p-[1px] shadow-sm transform transition-transform hover:scale-110 hover:z-[2] cursor-help',
                hasOverlap ? 'bottom-0 right-0' : hasTopBadges ? 'top-4 right-0' : '',
              )}
              style={{
                backgroundColor: '#dcfce7',
                borderColor: '#86efac',
                color: '#16a34a',
              }}
              title={`Оператор: ${getRouteLabel(op.RouteCode, op.Type)}`}
            >
              <Headphones size={10} className="mx-auto" />
              <span className="text-[6px] block text-center leading-none mt-0.5 font-bold">план</span>
            </div>
          );
        })}

        {hasFact && (
          <div
            className={cn(
              'absolute z-[2] flex h-[24px] w-[24px] flex-col items-center justify-center rounded-md border-2 p-[1px] shadow-md cursor-help ring-1 ring-white/90',
              primaryFactRoute ? getRouteColor(primaryFactRoute) : 'border-slate-400 bg-slate-50 text-slate-700',
                hasAnyPlan
                  ? 'bottom-0 left-0'
                  : maxDeviationOver300
                    ? 'top-[30px] left-1/2 -translate-x-1/2'
                    : 'top-1/2 left-1/2 -translate-x-1/2 -translate-y-1/2',
            )}
            title={[
              `Факт: ${factRoutes.join(', ') || 'маршрут не указан'}`,
              hasOrder ? `Сумма отгрузки: ${formatOrderAmount(totalOrderAmount)} ₽` : 'Сумма отгрузки: нет',
              minDeviation !== null ? `Отклонение: ${formatDistanceKm(minDeviation)}` : 'Отклонение: нет данных',
            ].join(' • ')}
          >
            {hasDeviationWithin300 ? (
              <User size={10} className="mx-auto" strokeWidth={2.4} />
            ) : (
              <Accessibility size={10} className="mx-auto" strokeWidth={2.4} />
            )}
            <span
              className={cn(
                'block text-center leading-none mt-0.5 font-extrabold tracking-[0.02em]',
                hasOrder ? 'text-[7px]' : 'text-[6px] uppercase',
              )}
            >
              {hasOrder ? '₽' : 'факт'}
            </span>
          </div>
        )}

        {maxDeviationOver300 && (
          <div className="absolute inset-x-0 top-0.5 z-[3] flex flex-col items-center gap-0.5 px-0.5">
            <span
              className="max-w-full whitespace-nowrap rounded-md bg-rose-200 px-1.5 py-[2px] text-[9px] font-black leading-none text-rose-900 shadow-sm"
              title={`Отклонение: ${formatDistanceKm(maxDeviationOver300)}`}
            >
              {formatDistanceBadgeKm(maxDeviationOver300)}
            </span>
          </div>
        )}
      </div>
    );
  };

  return (
    <div className="flex min-h-0 flex-col h-full overflow-hidden bg-white shadow-lg rounded-lg border border-gray-200">
      <div className="sticky top-0 relative z-[120] border-b bg-gray-50 p-4 flex-shrink-0">
        <div className="flex items-center justify-between gap-4 whitespace-nowrap">
          <div className="flex items-center gap-4 shrink-0">
            <div className="relative">
              <input
                type="text"
                placeholder="Поиск точки..."
                className="pl-8 pr-4 py-2 border rounded-md text-sm focus:outline-none focus:ring-2 focus:ring-blue-500 w-64"
                value={filterText}
                onChange={(e) => setFilterText(e.target.value)}
              />
              <Filter className="absolute left-2 top-2.5 text-gray-400" size={16} />
              {filterText && (
                <button
                  onClick={() => setFilterText('')}
                  className="absolute right-2 top-2.5 text-gray-400 hover:text-gray-600"
                >
                  <X size={16} />
                </button>
              )}
            </div>

            <div className="relative">
              <button
                type="button"
                onClick={() => {
                  setIsRouteFilterOpen((prev) => !prev);
                  setIsProximityFilterOpen(false);
                  setIsFactOrderFilterOpen(false);
                }}
                className={cn(
                  'flex items-center gap-2 rounded-md border bg-white px-3 py-2 text-sm text-gray-700 shadow-sm transition-colors',
                  isRouteFilterOpen || selectedRouteKeys.length > 0
                    ? 'border-blue-300 ring-2 ring-blue-100'
                    : 'border-gray-300 hover:border-gray-400',
                )}
              >
                <Filter size={16} className="text-gray-500" />
                <span className="font-medium">Маршруты</span>
                <span className="text-xs text-gray-500">
                  {selectedRouteKeys.length > 0 ? `выбрано: ${selectedRouteKeys.length}` : 'все'}
                </span>
              </button>

              {isRouteFilterOpen && (
                <div className="absolute left-0 top-full z-[9999] mt-2 w-[360px] rounded-xl border border-gray-200 bg-white shadow-2xl">
                  <div className="border-b border-gray-100 p-3">
                    <div className="mb-3 flex items-center justify-between gap-2">
                      <div>
                        <div className="text-sm font-semibold text-gray-900">Фильтр по маршрутам</div>
                        <div className="text-xs text-gray-500">Можно выбрать один или несколько маршрутов</div>
                      </div>
                      <button
                        type="button"
                        onClick={() => setIsRouteFilterOpen(false)}
                        className="rounded p-1 text-gray-400 hover:bg-gray-100 hover:text-gray-600"
                        title="Закрыть"
                      >
                        <X size={16} />
                      </button>
                    </div>

                    <div className="relative mb-3">
                      <input
                        type="text"
                        placeholder="Поиск маршрута..."
                        value={routeFilterText}
                        onChange={(e) => setRouteFilterText(e.target.value)}
                        className="w-full rounded-md border border-gray-300 py-2 pl-8 pr-8 text-sm focus:outline-none focus:ring-2 focus:ring-blue-500"
                      />
                      <Filter className="absolute left-2 top-2.5 text-gray-400" size={16} />
                      {routeFilterText && (
                        <button
                          type="button"
                          onClick={() => setRouteFilterText('')}
                          className="absolute right-2 top-2.5 text-gray-400 hover:text-gray-600"
                        >
                          <X size={16} />
                        </button>
                      )}
                    </div>

                    <div className="flex items-center justify-between gap-2 text-xs">
                      <button
                        type="button"
                        onClick={selectAllFilteredRoutes}
                        className="font-medium text-blue-600 hover:text-blue-700"
                      >
                        Выбрать найденные
                      </button>
                      <button
                        type="button"
                        onClick={clearRouteSelection}
                        className="font-medium text-gray-500 hover:text-gray-700"
                      >
                        Сбросить
                      </button>
                    </div>
                  </div>

                  <div className="max-h-80 overflow-auto p-2">
                    {filteredRouteOptions.length === 0 ? (
                      <div className="px-2 py-6 text-center text-sm text-gray-500">Маршруты не найдены</div>
                    ) : (
                      filteredRouteOptions.map((option) => {
                        const isSelected = selectedRouteSet.has(option.key);
                        const isSales = option.type === 'Торговый';

                        return (
                          <button
                            key={option.key}
                            type="button"
                            onClick={() => toggleRouteSelection(option.key)}
                            className="flex w-full items-center gap-3 rounded-lg px-2 py-2 text-left hover:bg-gray-50"
                            aria-pressed={isSelected}
                            aria-label={`Выбрать маршрут ${option.label}`}
                          >
                            <span
                              className={cn(
                                'flex h-5 w-5 shrink-0 items-center justify-center rounded border transition-colors',
                                isSelected
                                  ? 'border-blue-600 bg-blue-600 text-white'
                                  : 'border-gray-300 bg-white text-transparent',
                              )}
                            >
                              <Check size={12} />
                            </span>

                            <span
                              className={cn(
                                'inline-flex h-5 w-5 shrink-0 items-center justify-center rounded border',
                                isSales
                                  ? getRouteColor(option.routeCode)
                                  : 'border-green-200 bg-green-50 text-green-600',
                              )}
                            >
                              {isSales ? <User size={12} /> : <Headphones size={12} />}
                            </span>

                            <span className="min-w-0 flex-1 truncate text-sm font-medium text-gray-700" title={option.label}>
                              {option.label}
                            </span>
                          </button>
                        );
                      })
                    )}
                  </div>
                </div>
              )}
            </div>

            <div className="relative">
              <button
                type="button"
                onClick={() => {
                  setIsProximityFilterOpen((prev) => !prev);
                  setIsRouteFilterOpen(false);
                  setIsFactOrderFilterOpen(false);
                }}
                title="Близость плановых визитов"
                className={cn(
                  'flex items-center gap-2 rounded-md border bg-white px-3 py-2 text-sm text-gray-700 shadow-sm transition-colors',
                  isProximityFilterOpen || selectedProximityKeys.length > 0
                    ? 'border-blue-300 ring-2 ring-blue-100'
                    : 'border-gray-300 hover:border-gray-400',
                )}
              >
                <Filter size={16} className="text-gray-500" />
                <span className="font-medium" title="Близость плановых визитов ТП и Оператора">Близость ПВ</span>
                <span className="text-xs text-gray-500">
                  {selectedProximityKeys.length > 0 ? `выбрано: ${selectedProximityKeys.length}` : 'все'}
                </span>
              </button>

              {isProximityFilterOpen && (
                <div className="absolute left-0 top-full z-[9999] mt-2 w-[380px] rounded-xl border border-gray-200 bg-white shadow-2xl">
                  <div className="border-b border-gray-100 p-3">
                    <div className="mb-3 flex items-center justify-between gap-2">
                      <div>
                        <div className="text-sm font-semibold text-gray-900">Фильтр близости плановых визитов</div>
                        <div className="text-xs text-gray-500">Можно выбрать один или несколько параметров</div>
                      </div>
                      <button
                        type="button"
                        onClick={() => setIsProximityFilterOpen(false)}
                        className="rounded p-1 text-gray-400 hover:bg-gray-100 hover:text-gray-600"
                        title="Закрыть"
                      >
                        <X size={16} />
                      </button>
                    </div>

                    <div className="flex items-center justify-end gap-2 text-xs">
                      <button
                        type="button"
                        onClick={clearProximitySelection}
                        className="font-medium text-gray-500 hover:text-gray-700"
                      >
                        Сбросить
                      </button>
                    </div>
                  </div>

                  <div className="p-2">
                    {PROXIMITY_OPTIONS.map((option) => {
                      const isSelected = selectedProximitySet.has(option.key);

                      return (
                        <button
                          key={option.key}
                          type="button"
                          onClick={() => toggleProximitySelection(option.key)}
                          className="flex w-full items-start gap-3 rounded-lg px-2 py-2 text-left hover:bg-gray-50"
                          aria-pressed={isSelected}
                          aria-label={`Выбрать фильтр ${option.label}`}
                        >
                          <span
                            className={cn(
                              'mt-0.5 flex h-5 w-5 shrink-0 items-center justify-center rounded border transition-colors',
                              isSelected
                                ? 'border-blue-600 bg-blue-600 text-white'
                                : 'border-gray-300 bg-white text-transparent',
                            )}
                          >
                            <Check size={12} />
                          </span>

                          <div className="min-w-0 flex-1">
                            <div className="text-sm font-medium text-gray-800">{option.label}</div>
                            <div className="text-xs leading-snug text-gray-500 break-words whitespace-normal">{option.description}</div>
                          </div>
                        </button>
                      );
                    })}
                  </div>
                </div>
              )}
            </div>

            <div className="relative">
              <button
                type="button"
                onClick={() => {
                  setIsFactOrderFilterOpen((prev) => !prev);
                  setIsRouteFilterOpen(false);
                  setIsProximityFilterOpen(false);
                }}
                title="Близость факт заказ ТП от визита Оператора"
                className={cn(
                  'flex items-center gap-2 rounded-md border bg-white px-3 py-2 text-sm text-gray-700 shadow-sm transition-colors',
                  isFactOrderFilterOpen || selectedFactOrderKeys.length > 0
                    ? 'border-blue-300 ring-2 ring-blue-100'
                    : 'border-gray-300 hover:border-gray-400',
                )}
              >
                <Filter size={16} className="text-gray-500" />
                <span className="font-medium" title="Близость фактического заказа ТП к плановому визиту оператора.">Близость ФЗ</span>
                <span className="text-xs text-gray-500">
                  {selectedFactOrderKeys.length > 0 ? `выбрано: ${selectedFactOrderKeys.length}` : 'все'}
                </span>
              </button>

              {isFactOrderFilterOpen && (
                <div className="absolute left-0 top-full z-[9999] mt-2 w-[380px] rounded-xl border border-gray-200 bg-white shadow-2xl">
                  <div className="border-b border-gray-100 p-3">
                    <div className="mb-3 flex items-center justify-between gap-2">
                      <div>
                        <div className="text-sm font-semibold text-gray-900">Фильтр близости факт-заказов</div>
                        <div className="text-xs text-gray-500">Можно выбрать один или несколько параметров</div>
                      </div>
                      <button
                        type="button"
                        onClick={() => setIsFactOrderFilterOpen(false)}
                        className="rounded p-1 text-gray-400 hover:bg-gray-100 hover:text-gray-600"
                        title="Закрыть"
                      >
                        <X size={16} />
                      </button>
                    </div>

                    <div className="flex items-center justify-end gap-2 text-xs">
                      <button
                        type="button"
                        onClick={clearFactOrderSelection}
                        className="font-medium text-gray-500 hover:text-gray-700"
                      >
                        Сбросить
                      </button>
                    </div>
                  </div>

                  <div className="p-2">
                    {FACT_ORDER_PROXIMITY_OPTIONS.map((option) => {
                      const isSelected = selectedFactOrderSet.has(option.key);

                      return (
                        <button
                          key={option.key}
                          type="button"
                          onClick={() => toggleFactOrderSelection(option.key)}
                          className="flex w-full items-start gap-3 rounded-lg px-2 py-2 text-left hover:bg-gray-50"
                          aria-pressed={isSelected}
                          aria-label={`Выбрать фильтр ${option.label}`}
                        >
                          <span
                            className={cn(
                              'mt-0.5 flex h-5 w-5 shrink-0 items-center justify-center rounded border transition-colors',
                              isSelected
                                ? 'border-blue-600 bg-blue-600 text-white'
                                : 'border-gray-300 bg-white text-transparent',
                            )}
                          >
                            <Check size={12} />
                          </span>

                          <div className="min-w-0 flex-1">
                            <div className="text-sm font-medium text-gray-800">{option.label}</div>
                            <div className="text-xs leading-snug text-gray-500 break-words whitespace-normal">{option.description}</div>
                          </div>
                        </button>
                      );
                    })}
                  </div>
                </div>
              )}
            </div>

            <button
              type="button"
              onClick={() => setIsJointCoverageOnly((prev) => !prev)}
              className={cn(
                'inline-flex items-center gap-2 rounded-md border bg-white px-3 py-2 text-sm text-gray-700 shadow-sm transition-colors shrink-0',
                isJointCoverageOnly
                  ? 'border-blue-300 ring-2 ring-blue-100 text-blue-700'
                  : 'border-gray-300 hover:border-gray-400',
              )}
              title="Показать только точки, которые одновременно посещаются торговым и оператором"
            >
              <Check size={16} className={isJointCoverageOnly ? 'text-blue-600' : 'text-gray-400'} />
              <span className="font-medium">Общее</span>
            </button>

            <div className="text-sm text-gray-500 shrink-0">Всего точек: {filteredPoints.length}</div>
          </div>

          <div className="flex items-center gap-1 bg-white px-2 py-1 rounded-md border shadow-sm shrink-0">
            <button onClick={prevMonth} className="p-1 hover:bg-gray-100 rounded-full">
              <ChevronLeft size={18} />
            </button>
            <span className="text-base font-semibold w-32 text-center capitalize text-gray-800">
              {format(currentDate, 'LLLL yyyy', { locale: ru })}
            </span>
            <button onClick={nextMonth} className="p-1 hover:bg-gray-100 rounded-full">
              <ChevronRight size={18} />
            </button>
          </div>

          <div className="flex items-center gap-2 shrink-0">
            <button
              type="button"
              onClick={exportCurrentViewToHtml}
              disabled={filteredPoints.length === 0}
              className={cn(
                'inline-flex h-10 w-10 items-center justify-center rounded-md border bg-white text-sm font-medium shadow-sm transition-colors',
                filteredPoints.length === 0
                  ? 'cursor-not-allowed border-gray-200 text-gray-400'
                  : 'border-gray-300 text-gray-700 hover:border-gray-400 hover:bg-gray-50',
              )}
              title="Сохранить текущий экран с фильтрами и календарем в HTML для просмотра без приложения"
              aria-label="Сохранить текущий экран в HTML"
            >
              <FileText size={16} className={filteredPoints.length === 0 ? 'text-gray-300' : 'text-gray-500'} />
            </button>

            <button
              type="button"
              onClick={exportFilteredPointsToExcel}
              disabled={exportRows.length === 0}
              className={cn(
                'inline-flex h-10 w-10 items-center justify-center rounded-md border bg-white text-sm font-medium shadow-sm transition-colors',
                exportRows.length === 0
                  ? 'cursor-not-allowed border-gray-200 text-gray-400'
                  : 'border-gray-300 text-gray-700 hover:border-gray-400 hover:bg-gray-50',
              )}
              title="Выгрузить текущий отфильтрованный список точек в Excel"
              aria-label="Экспортировать текущий список точек в Excel"
            >
              <Download size={16} className={exportRows.length === 0 ? 'text-gray-300' : 'text-gray-500'} />
            </button>
          </div>
        </div>
      </div>

      {hasSelectionFilters && (
        <div className="relative z-[110] border-b bg-white px-4 py-2">
          <div className="flex items-start justify-between gap-4">
            <div className="flex min-w-0 flex-1 flex-col gap-2">
              {selectedRouteLabels.length > 0 && (
                <div className="flex items-center gap-2 flex-wrap">
                  <span className="text-xs font-medium text-gray-500">Выбранные маршруты:</span>
                  {selectedRouteLabels.map((label) => (
                    <span
                      key={label}
                      className="inline-flex items-center rounded-full border border-blue-200 bg-blue-50 px-2 py-1 text-xs font-medium text-blue-700"
                      title={label}
                    >
                      {label}
                    </span>
                  ))}
                </div>
              )}

              {selectedProximityLabels.length > 0 && (
                <div className="flex items-center gap-2 flex-wrap">
                  <span className="text-xs font-medium text-gray-500">Близость ПВ:</span>
                  {selectedProximityLabels.map((label) => (
                    <span
                      key={label}
                      className="inline-flex items-center rounded-full border border-violet-200 bg-violet-50 px-2 py-1 text-xs font-medium text-violet-700"
                      title={label}
                    >
                      {label}
                    </span>
                  ))}
                </div>
              )}

              {selectedFactOrderLabels.length > 0 && (
                <div className="flex items-center gap-2 flex-wrap">
                  <span className="text-xs font-medium text-gray-500">Близость ФЗ:</span>
                  {selectedFactOrderLabels.map((label) => (
                    <span
                      key={label}
                      className="inline-flex items-center rounded-full border border-amber-200 bg-amber-50 px-2 py-1 text-xs font-medium text-amber-700"
                      title={label}
                    >
                      {label}
                    </span>
                  ))}
                </div>
              )}

              {isJointCoverageOnly && (
                <div className="flex items-center gap-2 flex-wrap">
                  <span className="text-xs font-medium text-gray-500">Покрытие:</span>
                  <span
                    className="inline-flex items-center rounded-full border border-emerald-200 bg-emerald-50 px-2 py-1 text-xs font-medium text-emerald-700"
                    title="Показаны только точки с одновременным покрытием торгового и оператора"
                  >
                    Общее
                  </span>
                </div>
              )}
            </div>

            <button
              type="button"
              onClick={clearSelectionFilters}
              className="inline-flex shrink-0 items-center gap-2 rounded-md border border-gray-300 bg-white px-3 py-2 text-sm font-medium text-gray-700 shadow-sm transition-colors hover:border-gray-400 hover:bg-gray-50"
              title="Очистить выбранные маршруты, близость ПВ, близость ФЗ и общее покрытие"
            >
              <RotateCcw size={16} className="text-gray-500" />
              <span>Очистить все фильтры</span>
            </button>
          </div>
        </div>
      )}

      <div className="relative z-0 min-h-0 flex-1 overflow-auto">
        <table className="table-fixed w-full border-separate border-spacing-0">
          <thead>
            <tr>
              <th className="sticky top-0 left-0 z-[80] bg-gray-50 border-b border-r p-2 text-left w-[250px] min-w-[200px] text-xs font-semibold text-gray-600 uppercase tracking-wider h-10 shadow-sm">
                Точка Доставки
              </th>
              {daysInMonth.map((day) => {
                const isWeekend = [0, 6].includes(getDay(day));

                return (
                  <th
                    key={day.toISOString()}
                    className={cn(
                      'sticky top-0 z-[70] border-b border-r border-gray-200 p-0.5 text-center h-10 shadow-sm',
                      isWeekend
                        ? isToday(day)
                          ? 'bg-red-100'
                          : 'bg-red-50'
                        : isToday(day)
                          ? 'bg-blue-50'
                          : 'bg-white',
                    )}
                  >
                    <div
                      className={cn(
                        'text-[9px] font-medium capitalize leading-none',
                        isWeekend ? 'text-red-400' : 'text-gray-400',
                      )}
                    >
                      {format(day, 'EEEEEE', { locale: ru })}
                    </div>
                    <div
                      className={cn(
                        'text-xs font-bold leading-tight',
                        isWeekend ? 'text-red-600' : 'text-gray-700',
                      )}
                    >
                      {getDate(day)}
                    </div>
                  </th>
                );
              })}
            </tr>
          </thead>
          <tbody className="bg-white divide-y divide-gray-100">
            {filteredPoints.length === 0 ? (
              <tr>
                <td colSpan={daysInMonth.length + 1} className="p-12 text-center text-gray-500">
                  Нет данных для отображения.
                </td>
              </tr>
            ) : (
              filteredPoints.map((point) => (
                <tr key={point.ClientId} className="hover:bg-gray-50 group">
                  <td className="sticky left-0 z-10 bg-white border-r px-2 py-1 border-b group-hover:bg-gray-50">
                    <div className="flex flex-col space-y-0.5 w-[230px]">
                      <div className="flex items-center justify-between gap-2 min-w-0">
                        <span className="text-[9px] text-gray-500 font-medium uppercase tracking-wide truncate">{point.Branch}</span>
                        <div className="flex items-center gap-1 shrink-0">
                          {point.hasSales && (
                            <span
                              className="inline-flex h-4 w-4 items-center justify-center rounded border border-red-200 bg-red-50 text-red-500"
                              title="Точка посещается торговым"
                            >
                              <User size={10} />
                            </span>
                          )}
                          {point.hasOperator && (
                            <span
                              className="inline-flex h-4 w-4 items-center justify-center rounded border border-green-200 bg-green-50 text-green-600"
                              title="Точка обслуживается оператором"
                            >
                              <Headphones size={10} />
                            </span>
                          )}
                        </div>
                      </div>
                      <span className="text-xs font-bold text-gray-900 leading-tight truncate" title={point.Name}>{point.Name}</span>
                      <span className="text-[9px] text-gray-500 truncate" title={point.Address}>{point.Address}</span>
                      <div className="flex items-center justify-between gap-2">
                        <span className="text-[9px] text-gray-400 font-mono">{point.ClientId}</span>
                        {point.DeliveryZone && (
                          <span
                            className="text-[9px] text-green-700 font-medium truncate"
                            title={getDeliveryScheduleSummary(point)}
                          >
                            {getDeliveryScheduleSummary(point)}
                          </span>
                        )}
                      </div>
                    </div>
                  </td>
                  {daysInMonth.map((day) => {
                    const hasDelivery = isPointDeliveryScheduled(point, day);
                    const isWeekend = [0, 6].includes(getDay(day));
                    const proximityContourSides = getProximityContourSides(point.ClientId, day);
                    const factOrderContourSides = getFactOrderContourSides(point.ClientId, day);

                    return (
                      <td
                        key={day.toISOString()}
                        className={cn(
                          'border-r border-b border-gray-200 p-0 relative h-[76px] align-top overflow-hidden',
                          hasDelivery
                            ? 'bg-green-50'
                            : isWeekend
                              ? 'bg-red-50/60'
                              : 'bg-white',
                        )}
                      >
                        {proximityContourSides && (
                          <div
                            className={cn(
                              'pointer-events-none absolute inset-[-1px] z-[3]',
                              proximityContourSides.top && 'border-t-[3px] border-t-amber-500',
                              proximityContourSides.right && 'border-r-[3px] border-r-amber-500',
                              proximityContourSides.bottom && 'border-b-[3px] border-b-amber-500',
                              proximityContourSides.left && 'border-l-[3px] border-l-amber-500',
                            )}
                          />
                        )}
                        {factOrderContourSides && (
                          <div
                            className={cn(
                              'pointer-events-none absolute inset-[-1px] z-[4]',
                              factOrderContourSides.top && 'border-t-[3px] border-t-red-600',
                              factOrderContourSides.right && 'border-r-[3px] border-r-red-600',
                              factOrderContourSides.bottom && 'border-b-[3px] border-b-red-600',
                              factOrderContourSides.left && 'border-l-[3px] border-l-red-600',
                            )}
                          />
                        )}
                        {getCellContent(point.ClientId, day)}
                      </td>
                    );
                  })}
                </tr>
              ))
            )}
          </tbody>
        </table>
      </div>
    </div>
  );
}
