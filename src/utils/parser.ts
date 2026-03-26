import { DeliveryScheduleEntry, ScheduleEntry, VisitHistoryEntry } from './schedule';
import { parse, isValid } from 'date-fns';

export const parseScheduleRow = (row: any): ScheduleEntry | null => {
  // Extract key fields, handling various possible header names (case-insensitive or partial match)
  const getVal = (keys: string[]) => {
    for (const key of keys) {
      if (row[key] !== undefined) return row[key];
      // Try lowercase keys in row
      const lowerKey = key.toLowerCase();
      const foundKey = Object.keys(row).find(k => k.toLowerCase() === lowerKey);
      if (foundKey) return row[foundKey];
    }
    return undefined;
  };

  const branch = String(getVal(['Филиал', 'Branch']) || '');
  const clientId = String(getVal(['ИД клиента', 'ClientId', 'ТочкаДоставкиКод']) || '');
  const name = String(getVal(['Название', 'Name']) || '');
  const address = String(getVal(['Адрес', 'Address']) || '');
  const deliveryZone = String(getVal(['Зона доставки', 'DeliveryZone']) || '');
  
  const frequencyRaw = getVal(['Частота посещений за 4 недели', 'Frequency']);
  const frequency = frequencyRaw !== undefined ? Number(frequencyRaw) : 0;
  
  const routeCode = String(getVal(['Маршрут', 'RouteCode', 'МаршрутКод']) || '');
  const typeRaw = String(getVal(['Тип покрытия', 'Type']) || 'Торговый');
  
  const dayOfWeekRaw = getVal(['День посещения', 'DayOfWeek', 'ДеньНедели']);
  const dayOfWeek = dayOfWeekRaw !== undefined ? Number(dayOfWeekRaw) : 0;

  if (!clientId || !dayOfWeek) return null;

  const type = typeRaw.toLowerCase().includes('оператор') ? 'Оператор' : 'Торговый';

  // Helper to check if any value in a week block is non-zero
  const checkWeek = (weekNum: number): boolean => {
    const days = ['пн', 'вт', 'ср', 'чт', 'пт', 'сб', 'вс'];
    const suffix = String(weekNum);

    for (const d of days) {
        const key = `${d}${suffix}`;
        // Find property case-insensitive
        const entry = Object.entries(row).find(([k]) => k.toLowerCase() === key);
        if (entry) {
             const val = Number(entry[1]);
             if (!isNaN(val) && val !== 0) return true;
        }
    }
    return false;
  };

  let cycle = 40; // Default to every week

  if (frequency === 4) {
    cycle = 40;
  } else if (frequency === 0) {
    // Only for Sat/Sun, treat as every week for those days
    cycle = 40;
  } else if (frequency === 2) {
    // Frequency 2: 2 times in 4 weeks
    const hasEven = checkWeek(2) || checkWeek(4);
    const hasOdd = checkWeek(1) || checkWeek(3);
    
    if (hasEven && !hasOdd) {
        cycle = 22; // Even weeks
    } else if (hasOdd && !hasEven) {
        cycle = 21; // Odd weeks
    } else {
        // If ambiguous, check even first as per "even - if..." logic priority?
        // Actually, let's assume if it has Even flags, it's Even.
        if (hasEven) cycle = 22;
        else cycle = 21;
    }
  } else if (frequency === 1) {
    // Frequency 1: 1 time in 4 weeks
    if (checkWeek(1)) cycle = 11;
    else if (checkWeek(2)) cycle = 12;
    else if (checkWeek(3)) cycle = 13;
    else if (checkWeek(4)) cycle = 14;
    else cycle = 11; // Default
  }

  return {
    Branch: branch,
    ClientId: clientId,
    Name: name,
    Address: address,
    DeliveryZone: deliveryZone,
    Type: type,
    RouteCode: routeCode,
    DayOfWeek: dayOfWeek,
    Cycle: cycle,
  };
};

// Parse EasyMerch format with date columns
export const parseEasyMerchRow = (row: any, headers: string[]): ScheduleEntry[] => {
  const normalizeKey = (value: string) => value.replace(/\uFEFF/g, '').trim().toLowerCase();

  const getVal = (keys: string[]) => {
    const rowKeys = Object.keys(row);

    for (const key of keys) {
      if (row[key] !== undefined) return row[key];
      const normalizedTarget = normalizeKey(key);
      const foundKey = rowKeys.find((k) => normalizeKey(k) === normalizedTarget);
      if (foundKey) return row[foundKey];
    }
    return undefined;
  };

  const getRowValueByHeader = (header: string) => {
    if (row[header] !== undefined) return row[header];
    const normalizedHeader = normalizeKey(header);
    const foundKey = Object.keys(row).find((k) => normalizeKey(k) === normalizedHeader);
    return foundKey ? row[foundKey] : undefined;
  };

  const clientId = String(getVal(['ИД клиента', 'ID клиента', 'ClientId']) || '').trim();
  const name = String(getVal(['Название', 'Name']) || '').trim();
  const address = String(getVal(['Адрес', 'Address']) || '').trim();
  const deliveryZone = String(getVal(['Зона доставки', 'DeliveryZone']) || '').trim();
  const routeCode = String(getVal(['Маршрут', 'Route']) || '').trim();
  const typeRaw = String(getVal(['Тип покрытия', 'Type']) || 'Оператор').trim();

  if (!clientId) return [];

  const type = typeRaw.toLowerCase().includes('оператор') ? 'Оператор' : 'Торговый';

  const entries: ScheduleEntry[] = [];

  // Find date columns (headers that look like dates DD.MM.YYYY)
  const dateColumns = headers
    .map((h) => String(h || '').trim())
    .filter((h) => /^\d{2}\.\d{2}\.\d{4}$/.test(h));

  for (const dateStr of dateColumns) {
    const cellValue = getRowValueByHeader(dateStr);
    if (cellValue === undefined || cellValue === null || String(cellValue).trim() === '') continue;

    // Parse the date
    const parsedDate = parse(dateStr, 'dd.MM.yyyy', new Date());
    if (!isValid(parsedDate)) continue;

    // Get day of week (1-7, where 1 is Monday)
    const jsDay = parsedDate.getDay();
    const dayOfWeek = jsDay === 0 ? 7 : jsDay;

    entries.push({
      Branch: '',
      ClientId: clientId,
      Name: name,
      Address: address,
      DeliveryZone: deliveryZone,
      Type: type,
      RouteCode: routeCode,
      DayOfWeek: dayOfWeek,
      Cycle: 40,
      EasyMerchDate: dateStr,
    });
  }

  return entries;
};

export const parseDeliveryScheduleRow = (row: any): DeliveryScheduleEntry | null => {
  const getVal = (keys: string[]) => {
    for (const key of keys) {
      if (row[key] !== undefined) return row[key];
      const lowerKey = key.toLowerCase();
      const foundKey = Object.keys(row).find((k) => k.toLowerCase() === lowerKey);
      if (foundKey) return row[foundKey];
    }
    return undefined;
  };

  const zoneNumber = String(getVal(['Номер зоны', 'ZoneNumber']) || '').trim();
  const frequencyRaw = String(getVal(['Частота (по неделям)', 'Frequency']) || '00').trim();
  const requestDate = String(getVal(['Дата запроса', 'RequestDate']) || '').trim();

  const yesNo = (value: unknown) => String(value || '').trim().toUpperCase() === 'Y';

  if (!zoneNumber) return null;

  return {
    ZoneNumber: zoneNumber,
    Frequency: Number(frequencyRaw),
    RequestDate: requestDate,
    Monday: yesNo(getVal(['Понедельник', 'Monday'])),
    Tuesday: yesNo(getVal(['Вторник', 'Tuesday'])),
    Wednesday: yesNo(getVal(['Среда', 'Wednesday'])),
    Thursday: yesNo(getVal(['Четверг', 'Thursday'])),
    Friday: yesNo(getVal(['Пятница', 'Friday'])),
    Saturday: yesNo(getVal(['Суббота', 'Saturday'])),
    Sunday: yesNo(getVal(['Воскресенье', 'Sunday'])),
  };
};

export const parseVisitHistoryRow = (row: any): VisitHistoryEntry | null => {
  const normalizeKey = (value: string) => value
    .replace(/\uFEFF/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();

  const getVal = (keys: string[]) => {
    const rowKeys = Object.keys(row);

    for (const key of keys) {
      if (row[key] !== undefined) return row[key];

      const normalizedTarget = normalizeKey(key);
      const foundKey = rowKeys.find((k) => normalizeKey(k) === normalizedTarget);
      if (foundKey) return row[foundKey];
    }
    return undefined;
  };

  const parseLocalizedNumber = (value: unknown): number | undefined => {
    const text = String(value ?? '').replace(/\u00A0/g, ' ').trim();
    if (!text) return undefined;
    const normalized = text.replace(/\s+/g, '').replace(',', '.');
    const parsed = Number(normalized);
    return Number.isFinite(parsed) ? parsed : undefined;
  };

  const date = String(getVal(['Дата', 'Date']) || '').trim();
  const routeCode = String(getVal(['Маршрут', 'Route', 'RouteCode']) || '').trim();
  const clientId = String(getVal(['ИД клиента', 'ClientId', 'ID клиента']) || '').trim();
  const name = String(getVal(['Название', 'Name']) || '').trim();
  const address = String(getVal(['Адрес', 'Address']) || '').trim();
  const coordinateDeviationMeters = parseLocalizedNumber(
    getVal([
      'Отклонение координат ТТ и визита м',
      'Отклонение координат ТТ и визита, м',
      'Отклонение координат тт и визита м',
      'CoordinateDeviationMeters',
    ]),
  );
  const orderAmountRub = parseLocalizedNumber(
    getVal([
      'Сумма заказа руб',
      'Сумма заказа, руб',
      'Сумма заказа',
      'OrderAmountRub',
    ]),
  );

  if (!date || !clientId) return null;

  return {
    Date: date,
    RouteCode: routeCode,
    ClientId: clientId,
    Name: name,
    Address: address,
    CoordinateDeviationMeters: coordinateDeviationMeters,
    OrderAmountRub: orderAmountRub,
  };
};
