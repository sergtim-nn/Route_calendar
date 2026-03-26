import { getISOWeek, getDay } from 'date-fns';
import { clsx, type ClassValue } from 'clsx';
import { twMerge } from 'tailwind-merge';

export function cn(...inputs: ClassValue[]) {
  return twMerge(clsx(inputs));
}

export type ScheduleEntry = {
  Branch: string;
  ClientId: string;
  Name: string;
  Address: string;
  DeliveryZone?: string;
  Type: 'Торговый' | 'Оператор';
  RouteCode: string;
  DayOfWeek: number; // 1 (Mon) - 7 (Sun)
  Cycle: number; // 40, 21, 22, 11, 12, 13, 14
  EasyMerchDate?: string; // Actual date for EasyMerch format (DD.MM.YYYY)
};

export type DeliveryScheduleEntry = {
  ZoneNumber: string;
  Frequency: number; // 00, 10, 20, 41, 42, 43, 44
  RequestDate?: string;
  Monday: boolean;
  Tuesday: boolean;
  Wednesday: boolean;
  Thursday: boolean;
  Friday: boolean;
  Saturday: boolean;
  Sunday: boolean;
};

export type VisitHistoryEntry = {
  Date: string;
  RouteCode: string;
  ClientId: string;
  Name: string;
  Address: string;
  CoordinateDeviationMeters?: number;
  OrderAmountRub?: number;
};

// Colors for different routes to distinguish them visually
const ROUTE_COLORS = [
  'bg-red-100 text-red-600 border-red-200',
  'bg-blue-100 text-blue-600 border-blue-200',
  'bg-green-100 text-green-600 border-green-200',
  'bg-yellow-100 text-yellow-600 border-yellow-200',
  'bg-purple-100 text-purple-600 border-purple-200',
  'bg-pink-100 text-pink-600 border-pink-200',
  'bg-indigo-100 text-indigo-600 border-indigo-200',
  'bg-orange-100 text-orange-600 border-orange-200',
];

export const getRouteColor = (routeCode: string) => {
  let hash = 0;
  for (let i = 0; i < routeCode.length; i++) {
    hash = routeCode.charCodeAt(i) + ((hash << 5) - hash);
  }
  const index = Math.abs(hash) % ROUTE_COLORS.length;
  return ROUTE_COLORS[index];
};

export const isScheduled = (date: Date, entry: ScheduleEntry): boolean => {
  // If this is an EasyMerch entry with a specific date, check that date
  if (entry.EasyMerchDate) {
    const [day, month, year] = entry.EasyMerchDate.split('.').map(Number);
    const entryDate = new Date(year, month - 1, day);
    return entryDate.toDateString() === date.toDateString();
  }
  
  // date-fns getDay returns 0 for Sunday, 1 for Monday.
  // We need to map Sunday (0) to 7 to match the input format (1-7).
  const day = getDay(date);
  const adjustedDay = day === 0 ? 7 : day;

  if (adjustedDay !== entry.DayOfWeek) return false;

  const isoWeek = getISOWeek(date);

  switch (entry.Cycle) {
    case 40: // Every week
      return true;
    case 21: // Odd weeks
      return isoWeek % 2 !== 0;
    case 22: // Even weeks
      return isoWeek % 2 === 0;
    case 11: // 1st week of 4-week cycle
      return (isoWeek - 1) % 4 === 0;
    case 12: // 2nd week of 4-week cycle
      return (isoWeek - 1) % 4 === 1;
    case 13: // 3rd week of 4-week cycle
      return (isoWeek - 1) % 4 === 2;
    case 14: // 4th week of 4-week cycle
      return (isoWeek - 1) % 4 === 3;
    default:
      return false;
  }
};

export const isDeliveryScheduled = (date: Date, entry: DeliveryScheduleEntry): boolean => {
  const day = getDay(date);
  const isoWeek = getISOWeek(date);

  const dayMatches = (
    (day === 1 && entry.Monday) ||
    (day === 2 && entry.Tuesday) ||
    (day === 3 && entry.Wednesday) ||
    (day === 4 && entry.Thursday) ||
    (day === 5 && entry.Friday) ||
    (day === 6 && entry.Saturday) ||
    (day === 0 && entry.Sunday)
  );

  if (!dayMatches) return false;

  switch (entry.Frequency) {
    case 0:
      return true;
    case 10:
      return isoWeek % 2 !== 0;
    case 20:
      return isoWeek % 2 === 0;
    case 41:
      return (isoWeek - 1) % 4 === 0;
    case 42:
      return (isoWeek - 1) % 4 === 1;
    case 43:
      return (isoWeek - 1) % 4 === 2;
    case 44:
      return (isoWeek - 1) % 4 === 3;
    default:
      return false;
  }
};
