import { useMemo, useState } from 'react';
import * as XLSX from 'xlsx';
import { Upload, FileSpreadsheet, Loader2, Download, ArchiveRestore, Save } from 'lucide-react';
import type { DeliveryScheduleEntry, ScheduleEntry, VisitHistoryEntry } from '../utils/schedule';
import { parseVisitHistoryRow } from '../utils/parser';

interface FileImportProps {
  onDataLoaded: (data: any[]) => void;
}

interface FileImportEasyMerchProps {
  onEasyMerchLoaded: (data: any[], headers: string[]) => void;
}

interface FileImportDeliveryScheduleProps {
  onDeliveryScheduleLoaded: (data: any[]) => void;
}

interface FileImportVisitHistoryProps {
  onVisitHistoryLoaded: (data: VisitHistoryEntry[]) => void;
}

interface BackupPayload {
  version: 1;
  exportedAt: string;
  entries?: ScheduleEntry[];
  deliveryScheduleEntries?: DeliveryScheduleEntry[];
  visitHistoryEntries?: VisitHistoryEntry[];
}

interface FileImportBackupProps {
  entries: ScheduleEntry[];
  deliveryScheduleEntries: DeliveryScheduleEntry[];
  visitHistoryEntries: VisitHistoryEntry[];
  onBackupRestore: (payload: BackupPayload) => void;
}

const normalizeEasyMerchHeader = (value: unknown): string => {
  if (value === undefined || value === null) return '';

  const text = String(value).replace(/\uFEFF/g, '').trim();
  const match = text.match(/^(\d{1,2})[./-](\d{1,2})[./-](\d{4})$/);

  if (match) {
    const [, day, month, year] = match;
    return `${day.padStart(2, '0')}.${month.padStart(2, '0')}.${year}`;
  }

  return text;
};

const parseLocalizedVisitHistoryNumber = (value: unknown): number | undefined => {
  if (value === undefined || value === null || value === '') return undefined;
  if (typeof value === 'number' && Number.isFinite(value)) return value;

  const text = String(value ?? '')
    .replace(/\u00A0/g, ' ')
    .replace(/\u202F/g, ' ')
    .trim();

  if (!text) return undefined;

  const normalized = text.replace(/\s+/g, '').replace(',', '.');
  const parsed = Number(normalized);
  return Number.isFinite(parsed) ? parsed : undefined;
};

const formatVisitHistoryDateValue = (value: unknown): string => {
  if (value === undefined || value === null || value === '') return '';

  if (value instanceof Date && !Number.isNaN(value.getTime())) {
    const day = String(value.getDate()).padStart(2, '0');
    const month = String(value.getMonth() + 1).padStart(2, '0');
    const year = value.getFullYear();
    return `${day}.${month}.${year}`;
  }

  if (typeof value === 'number' && Number.isFinite(value)) {
    const parsed = XLSX.SSF.parse_date_code(value);
    if (parsed) {
      const day = String(parsed.d).padStart(2, '0');
      const month = String(parsed.m).padStart(2, '0');
      const year = parsed.y;
      return `${day}.${month}.${year}`;
    }
  }

  return normalizeEasyMerchHeader(value);
};

const normalizeVisitHistoryHeader = (value: unknown): string => String(value ?? '')
  .replace(/\uFEFF/g, '')
  .replace(/\u00A0/g, ' ')
  .replace(/\u202F/g, ' ')
  .replace(/\s+/g, ' ')
  .trim()
  .toLowerCase();

const compactVisitHistoryHeader = (value: unknown): string => normalizeVisitHistoryHeader(value)
  .replace(/[^a-zа-яё0-9]/gi, '');

const isVisitHistoryDateHeader = (header: unknown): boolean => {
  const normalized = normalizeVisitHistoryHeader(header);
  return normalized === 'дата' || normalized === 'date' || normalized.includes('дата визита');
};

const getVisitHistoryHeaderScore = (row: Array<string | number | Date | null>): number => {
  const compactHeaders = row.map(compactVisitHistoryHeader);
  const hasAny = (values: string[]) => compactHeaders.some((header) => values.includes(header));
  const hasPartial = (values: string[]) => compactHeaders.some((header) => values.some((value) => header.includes(value)));

  let score = 0;
  if (hasAny(['дата', 'date']) || hasPartial(['датавизита'])) score += 3;
  if (hasAny(['маршрут', 'route', 'routecode', 'кодмаршрута'])) score += 2;
  if (hasAny(['идклиента', 'idклиента', 'clientid', 'кодклиента', 'клиенткод', 'кодтт', 'idтт', 'точкадоставкикод'])) score += 3;
  if (hasAny(['название', 'name', 'клиент', 'названиеклиента', 'наименованиеклиента', 'названиетт', 'наименованиетт'])) score += 1;
  if (hasAny(['адрес', 'address', 'адресклиента', 'адрестт', 'фактическийадрес'])) score += 1;
  if (hasPartial(['отклонениекоординат', 'суммазаказа'])) score += 1;

  return score;
};

const findVisitHistoryHeaderIndex = (rows: Array<Array<string | number | Date | null>>): number => {
  let bestIndex = 0;
  let bestScore = 0;

  rows.slice(0, 30).forEach((row, index) => {
    const score = getVisitHistoryHeaderScore(row);
    if (score > bestScore) {
      bestScore = score;
      bestIndex = index;
    }
  });

  return bestScore >= 5 ? bestIndex : 0;
};

const buildVisitHistoryRowObject = (
  headers: Array<string | number | Date | null>,
  row: Array<string | number | Date | null>,
): Record<string, unknown> => {
  const result: Record<string, unknown> = {};

  headers.forEach((header, index) => {
    const headerName = String(header ?? '').trim();
    if (!headerName) return;

    result[headerName] = isVisitHistoryDateHeader(headerName)
      ? formatVisitHistoryDateValue(row[index])
      : row[index];
  });

  return result;
};

const routeTemplateHeaders = [
  'Филиал',
  'ИД клиента',
  'Название',
  'Адрес',
  'Зона доставки',
  'Частота посещений за 4 недели',
  'Маршрут',
  'Тип покрытия',
  'День посещения',
  'пн1', 'вт1', 'ср1', 'чт1', 'пт1', 'сб1', 'вс1',
  'пн2', 'вт2', 'ср2', 'чт2', 'пт2', 'сб2', 'вс2',
  'пн3', 'вт3', 'ср3', 'чт3', 'пт3', 'сб3', 'вс3',
  'пн4', 'вт4', 'ср4', 'чт4', 'пт4', 'сб4', 'вс4',
];

const routeTemplateRows = [
  ['(828) Филиал Сочи', '10043616', 'Жвитиашвили Натела Гочавна ИП', '354340, Россия, Краснодарский край, Сочи г, Гастелло ул, д. 40 корп А', 'Юг-1', 2, 'N33_828NFD', 'Торговый', 5, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0],
  ['(828) Филиал Сочи', '10043616', 'Жвитиашвили Натела Гочавна ИП', '354340, Россия, Краснодарский край, Сочи г, Гастелло ул, д. 40 корп А', 'Юг-1', 2, 'Z52_01CON', 'Оператор', 5, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0],
  ['(828) Филиал Сочи', '10085444', 'Олимпик Сити ООО', '354000, Россия, Сочи г, Пластунская ул, д. 52 корп. 3', 'Центр', 2, 'N33_828NFD', 'Торговый', 2, 0, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 2, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0],
  ['(828) Филиал Сочи', '10085444', 'Олимпик Сити ООО', '354000, Россия, Сочи г, Пластунская ул, д. 52 корп. 3', 'Центр', 2, 'Z52_01CON', 'Оператор', 2, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 1, 0, 0, 0, 0, 0],
  ['(828) Филиал Сочи', '11115693', 'Фирма "Каньон" ООО', '354207, Россия, Краснодарский край, Сочи г, Батумское шоссе ул, д. 69 А', '', 0, 'N33_828NFD', 'Торговый', 7, 0, 0, 0, 0, 0, 0, 15, 0, 0, 0, 0, 0, 0, 15, 0, 0, 0, 0, 0, 0, 15, 0, 0, 0, 0, 0, 0, 15],
];

const easyMerchTemplateHeaders = [
  'ИД клиента',
  'Название',
  'Адрес',
  'Зона доставки',
  'Маршрут',
  'Тип покрытия',
  '01.03.2026', '02.03.2026', '03.03.2026', '04.03.2026', '05.03.2026', '06.03.2026', '07.03.2026', '08.03.2026', '09.03.2026', '10.03.2026', '11.03.2026', '12.03.2026', '13.03.2026', '14.03.2026', '15.03.2026', '16.03.2026', '17.03.2026', '18.03.2026', '19.03.2026', '20.03.2026', '21.03.2026', '22.03.2026', '23.03.2026', '24.03.2026', '25.03.2026', '26.03.2026', '27.03.2026', '28.03.2026', '29.03.2026', '30.03.2026', '31.03.2026',
];

const easyMerchTemplateRows = [
  ['14044406', 'Ладыгина Екатерина Александровна ИП', 'Астрахань г, Адмирала Нахимова ул, д. 149', 'Астрахань-1', '01TLM_01M', 'Оператор', '', '2x25', '', '', '', '', '', '', '', '2x1', '', '', '', '', '', '', '2x1', '', '', '', '', '', '', '2x1', '', '', '', '', '', '', '2x1'],
  ['8103477', 'Устьян Артур Борисович ИП', 'Верхнеармянская Хобза с, Разданская ул, д.25корп А', 'Сочи-Север', '01TLM_01M', 'Оператор', '', '2x1', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '2x1', '', '', '', '', '', '', '', '', '', '', '', '', '', '2x1'],
  ['8376262', 'Буюклян Алина Ашотовна, ИП', 'Горное Лоо с, Обзорная ул, д. 1', 'Сочи-Север', '01TLM_01M', 'Оператор', '', '', '2x1', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '2x12', '', '', '', '', '', '', '', '', '', '', '', '2x12'],
  ['8097846', 'Давыдов Денис Владимирович ИП', 'Астрахань г, Боевая ул, д. 75 корп Б', '', '01TLM_01M', 'Оператор', '', '', '', '2x1', '', '', '', '', '', '', '', '2x40', '', '', '', '', '', '', '2x1', '', '', '', '', '', '', '2x2', '', '', '', '', '', ''],
  ['8096669', 'Фармсервис ООО', 'Астрахань г, Адмиралтейская ул, д. 28', 'Центр', '01TLM_01M', 'Оператор', '', '', '', '', '2x13', '', '', '', '', '', '', '', '2x1', '', '', '', '', '', '', '2x1', '', '', '', '', '', '', '2x1', '', '', '', ''],
  ['13743743', 'Галустян Анжелика Гамаяковна ИП', 'Сочи г, Виноградная ул, д. 122', '', '01TLM_01M', 'Оператор', '', '', '', '', '2x1', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '', '2x37', '', '', '', '', '', '', ''],
];

const deliveryScheduleTemplateHeaders = [
  'Номер зоны',
  'Частота (по неделям)',
  'Дата запроса',
  'Понедельник',
  'Вторник',
  'Среда',
  'Четверг',
  'Пятница',
  'Суббота',
  'Воскресенье',
];

const deliveryScheduleTemplateRows = [
  ['Ч0Г', '00', '', 'N', 'N', 'Y', 'N', 'Y', 'N', 'N'],
  ['Ч6А', '00', '', 'Y', 'N', 'Y', 'N', 'Y', 'N', 'N'],
  ['Ч7В', '00', '', 'N', 'Y', 'N', 'Y', 'Y', 'N', 'N'],
  ['Ч8Б', '00', '', 'N', 'Y', 'N', 'Y', 'N', 'N', 'N'],
  ['Ю07', '00', '', 'N', 'N', 'Y', 'N', 'N', 'N', 'N'],
  ['Э18', '10', '', 'N', 'Y', 'N', 'N', 'N', 'N', 'N'],
  ['Э25', '20', '', 'N', 'Y', 'N', 'N', 'N', 'N', 'N'],
  ['Э27', '10', '', 'N', 'N', 'N', 'Y', 'N', 'N', 'N'],
  ['Э28', '20', '', 'N', 'N', 'N', 'Y', 'N', 'N', 'N'],
  ['H02', '00', '', 'Y', 'N', 'Y', 'Y', 'Y', 'N', 'N'],
  ['H09', '20', '', 'N', 'N', 'Y', 'N', 'N', 'N', 'N'],
  ['H10', '20', '', 'N', 'N', 'N', 'Y', 'N', 'N', 'N'],
  ['H11', '20', '', 'N', 'N', 'N', 'Y', 'N', 'N', 'N'],
];

const visitHistoryTemplateHeaders = [
  'Дата',
  'Маршрут',
  'ИД клиента',
  'Название',
  'Адрес',
  'Отклонение координат ТТ и визита м',
  'Сумма заказа руб',
];

const visitHistoryTemplateRows = [
  ['02.03.2026', '560_42CON', '10044444', 'КСК-Воронеж ООО', '394065, Воронежская обл, Воронеж г, Патриотов пр-кт, 11, Б', '16 237,72', ''],
  ['02.03.2026', '560_42CON', '9610713', 'Эталон Торг ООО', '394065, Воронежская обл, Воронеж г, Патриотов пр-кт, д.23 В', '299,31', ''],
  ['02.03.2026', '560_42CON', '9610585', 'Авалон ООО', '394065, Воронежская обл, Воронеж г, Патриотов пр-кт, д. 28 А', '1 410,24', ''],
  ['02.03.2026', '560_42CON', '10044347', 'КСК-Воронеж ООО', '394065, Воронежская обл, Воронеж г, Патриотов пр-кт, 116, А', '848,17', ''],
  ['02.03.2026', '560_42CON', '9611520', 'Визит ООО', '394048, Воронежская обл, Воронеж г, Междуреченская ул, д.1Б', '86,43', '5 170,18'],
  ['02.03.2026', '560_42CON', '9611519', 'Мнацаканян Нели Торгомовна ИП', '394048, Шилово рп, Междуреченская ул, д.10', '9 332,33', ''],
  ['02.03.2026', '560_42CON', '10044317', 'КСК-Воронеж ООО', '394048, Воронежская обл, Воронеж г, Междуреченская ул, 4', '1 360 184,80', ''],
];

const downloadTemplateWorkbook = (fileName: string, headers: string[], rows: Array<Array<string | number>>) => {
  const workbook = XLSX.utils.book_new();
  const worksheet = XLSX.utils.aoa_to_sheet([headers, ...rows]);

  const columnWidths = headers.map((header) => ({
    wch: Math.min(Math.max(String(header).length + 2, 12), 28),
  }));

  worksheet['!cols'] = columnWidths;
  XLSX.utils.book_append_sheet(workbook, worksheet, 'Шаблон');
  XLSX.writeFile(workbook, fileName);
};

export function FileImport({ onDataLoaded }: FileImportProps) {
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState(false);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setLoading(true);
    setError(null);
    setSuccess(false);

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const workbook = XLSX.read(bstr, { type: 'binary' });
        const wsname = workbook.SheetNames[0];
        const ws = workbook.Sheets[wsname];

        const data = XLSX.utils.sheet_to_json(ws, {
          raw: false,
          defval: '',
        });

        onDataLoaded(data);
        setSuccess(true);
        setTimeout(() => setSuccess(false), 3000);
      } catch (err) {
        console.error('Error parsing Excel:', err);
        setError('Не удалось прочитать файл. Убедитесь, что это корректный Excel файл.');
      } finally {
        setLoading(false);
      }
    };

    reader.readAsBinaryString(file);
  };

  return (
    <div className="p-6 bg-white rounded-lg shadow-sm border border-gray-200 text-center">
      <div className="mb-4 flex justify-center text-blue-500">
        <FileSpreadsheet size={48} />
      </div>
      <h3 className="text-lg font-medium text-gray-900 mb-2">Загрузка маршрутов</h3>
      <p className="text-sm text-gray-500 mb-6">
        Загрузите Excel файл с маршрутами (формат: .xlsx, .xls)
      </p>

      <div className="flex flex-col items-center gap-3">
        <div className="relative inline-block">
          <input
            type="file"
            accept=".xlsx, .xls"
            onChange={handleFileUpload}
            className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
            disabled={loading}
          />
          <button
            className="flex items-center space-x-2 px-4 py-2 bg-blue-600 text-white rounded-md hover:bg-blue-700 transition-colors disabled:opacity-50"
            disabled={loading}
          >
            {loading ? <Loader2 className="animate-spin" size={20} /> : <Upload size={20} />}
            <span>{loading ? 'Обработка...' : 'Выберите файл'}</span>
          </button>
        </div>

        <button
          type="button"
          onClick={() => downloadTemplateWorkbook('template-routes.xlsx', routeTemplateHeaders, routeTemplateRows)}
          className="inline-flex items-center space-x-2 px-4 py-2 text-sm font-medium text-blue-700 bg-blue-50 border border-blue-200 rounded-md hover:bg-blue-100 transition-colors"
        >
          <Download size={16} />
          <span>Скачать шаблон с примером</span>
        </button>
      </div>

      {success && (
        <div className="mt-4 text-sm text-green-600 bg-green-50 p-2 rounded">
          ✓ Файл успешно загружен
        </div>
      )}

      {error && (
        <div className="mt-4 text-sm text-red-600 bg-red-50 p-2 rounded">
          {error}
        </div>
      )}

      <div className="mt-6 text-left text-xs text-gray-400">
        <p className="font-semibold mb-1">Ожидаемый формат колонок:</p>
        <code className="bg-gray-100 p-1 rounded block overflow-x-auto text-xs">
          Филиал | ИД клиента | Название | Адрес | Зона доставки | Частота посещений за 4 недели | Маршрут | Тип покрытия | День посещения | пн1 ... вс4
        </code>
      </div>
    </div>
  );
}

export function FileImportEasyMerch({ onEasyMerchLoaded }: FileImportEasyMerchProps) {
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState(false);
  const [loadedRowsCount, setLoadedRowsCount] = useState(0);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setLoading(true);
    setError(null);
    setSuccess(false);
    setLoadedRowsCount(0);

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const workbook = XLSX.read(bstr, { type: 'binary', cellDates: true });
        const wsname = workbook.SheetNames[0];
        const ws = workbook.Sheets[wsname];

        const rows = XLSX.utils.sheet_to_json<(string | number | Date | null)[]>(ws, {
          header: 1,
          raw: false,
          defval: '',
          blankrows: false,
          dateNF: 'dd.mm.yyyy',
        });

        if (!rows.length) {
          throw new Error('EasyMerch sheet is empty');
        }

        const headers = (rows[0] ?? []).map((cell) => normalizeEasyMerchHeader(cell));

        const data = rows
          .slice(1)
          .filter((row) => row.some((cell) => String(cell ?? '').trim() !== ''))
          .map((row) => {
            const record: Record<string, string> = {};
            headers.forEach((header, index) => {
              if (!header) return;
              record[header] = String(row[index] ?? '').trim();
            });
            return record;
          });

        onEasyMerchLoaded(data, headers);
        setLoadedRowsCount(data.length);
        setSuccess(true);
        setTimeout(() => setSuccess(false), 3000);
      } catch (err) {
        console.error('Error parsing EasyMerch Excel:', err);
        setError('Не удалось прочитать файл. Убедитесь, что это корректный Excel файл EasyMerch.');
      } finally {
        setLoading(false);
      }
    };

    reader.readAsBinaryString(file);
  };

  return (
    <div className="p-6 bg-white rounded-lg shadow-sm border border-gray-200 text-center">
      <div className="mb-4 flex justify-center text-green-500">
        <FileSpreadsheet size={48} />
      </div>
      <h3 className="text-lg font-medium text-gray-900 mb-2">Загрузка маршрутов EasyMerch</h3>
      <p className="text-sm text-gray-500 mb-6">
        Загрузите Excel файл с маршрутами EasyMerch (формат с датами в заголовках)
      </p>

      <div className="flex flex-col items-center gap-3">
        <div className="relative inline-block">
          <input
            type="file"
            accept=".xlsx, .xls"
            onChange={handleFileUpload}
            className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
            disabled={loading}
          />
          <button
            className="flex items-center space-x-2 px-4 py-2 bg-green-600 text-white rounded-md hover:bg-green-700 transition-colors disabled:opacity-50"
            disabled={loading}
          >
            {loading ? <Loader2 className="animate-spin" size={20} /> : <Upload size={20} />}
            <span>{loading ? 'Обработка...' : 'Выберите файл'}</span>
          </button>
        </div>

        <button
          type="button"
          onClick={() => downloadTemplateWorkbook('template-easymerch.xlsx', easyMerchTemplateHeaders, easyMerchTemplateRows)}
          className="inline-flex items-center space-x-2 px-4 py-2 text-sm font-medium text-green-700 bg-green-50 border border-green-200 rounded-md hover:bg-green-100 transition-colors"
        >
          <Download size={16} />
          <span>Скачать шаблон с примером</span>
        </button>
      </div>

      {success && (
        <div className="mt-4 text-sm text-green-600 bg-green-50 p-2 rounded space-y-1">
          <div>✓ Файл успешно загружен</div>
          <div>
            Загружено строк EasyMerch: <span className="font-semibold">{loadedRowsCount}</span>
          </div>
        </div>
      )}

      {error && (
        <div className="mt-4 text-sm text-red-600 bg-red-50 p-2 rounded">
          {error}
        </div>
      )}

      <div className="mt-6 text-left text-xs text-gray-400">
        <p className="font-semibold mb-1">Ожидаемый формат:</p>
        <code className="bg-gray-100 p-1 rounded block overflow-x-auto text-xs">
          ИД клиента | Название | Адрес | Зона доставки | Маршрут | Тип покрытия | 01.03.2026 | 02.03.2026 | ...
        </code>
      </div>
    </div>
  );
}

export function FileImportDeliverySchedule({ onDeliveryScheduleLoaded }: FileImportDeliveryScheduleProps) {
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState(false);
  const [loadedRowsCount, setLoadedRowsCount] = useState(0);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setLoading(true);
    setError(null);
    setSuccess(false);
    setLoadedRowsCount(0);

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const workbook = XLSX.read(bstr, { type: 'binary' });
        const wsname = workbook.SheetNames[0];
        const ws = workbook.Sheets[wsname];

        const data = XLSX.utils.sheet_to_json(ws, {
          raw: false,
          defval: '',
        });

        onDeliveryScheduleLoaded(data);
        setLoadedRowsCount(data.length);
        setSuccess(true);
        setTimeout(() => setSuccess(false), 3000);
      } catch (err) {
        console.error('Error parsing delivery schedule Excel:', err);
        setError('Не удалось прочитать файл. Убедитесь, что это корректный Excel файл графика доставки.');
      } finally {
        setLoading(false);
      }
    };

    reader.readAsBinaryString(file);
  };

  return (
    <div className="p-6 bg-white rounded-lg shadow-sm border border-gray-200 text-center">
      <div className="mb-4 flex justify-center text-amber-500">
        <FileSpreadsheet size={48} />
      </div>
      <h3 className="text-lg font-medium text-gray-900 mb-2">График доставки</h3>
      <p className="text-sm text-gray-500 mb-6">
        Загрузите Excel файл с графиком доставки по зонам
      </p>

      <div className="flex flex-col items-center gap-3">
        <div className="relative inline-block">
          <input
            type="file"
            accept=".xlsx, .xls"
            onChange={handleFileUpload}
            className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
            disabled={loading}
          />
          <button
            className="flex items-center space-x-2 px-4 py-2 bg-amber-600 text-white rounded-md hover:bg-amber-700 transition-colors disabled:opacity-50"
            disabled={loading}
          >
            {loading ? <Loader2 className="animate-spin" size={20} /> : <Upload size={20} />}
            <span>{loading ? 'Обработка...' : 'Выберите файл'}</span>
          </button>
        </div>

        <button
          type="button"
          onClick={() => downloadTemplateWorkbook('template-delivery-schedule.xlsx', deliveryScheduleTemplateHeaders, deliveryScheduleTemplateRows)}
          className="inline-flex items-center space-x-2 px-4 py-2 text-sm font-medium text-amber-700 bg-amber-50 border border-amber-200 rounded-md hover:bg-amber-100 transition-colors"
        >
          <Download size={16} />
          <span>Скачать шаблон с примером</span>
        </button>
      </div>

      {success && (
        <div className="mt-4 text-sm text-green-600 bg-green-50 p-2 rounded space-y-1">
          <div>✓ Файл успешно загружен</div>
          <div>
            Загружено зон доставки: <span className="font-semibold">{loadedRowsCount}</span>
          </div>
        </div>
      )}

      {error && (
        <div className="mt-4 text-sm text-red-600 bg-red-50 p-2 rounded">
          {error}
        </div>
      )}

      <div className="mt-6 text-left text-xs text-gray-400">
        <p className="font-semibold mb-1">Ожидаемый формат:</p>
        <code className="bg-gray-100 p-1 rounded block overflow-x-auto text-xs">
          Номер зоны | Частота (по неделям) | Дата запроса | Понедельник | Вторник | Среда | Четверг | Пятница | Суббота | Воскресенье
        </code>
      </div>
    </div>
  );
}

export function FileImportVisitHistory({ onVisitHistoryLoaded }: FileImportVisitHistoryProps) {
  const [loading, setLoading] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState(false);
  const [loadedRowsCount, setLoadedRowsCount] = useState(0);

  const handleFileUpload = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setLoading(true);
    setError(null);
    setSuccess(false);
    setLoadedRowsCount(0);

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const bstr = evt.target?.result;
        const workbook = XLSX.read(bstr, { type: 'binary', cellDates: true });
        const wsname = workbook.SheetNames[0];
        const ws = workbook.Sheets[wsname];

        const rows = XLSX.utils.sheet_to_json<(string | number | Date | null)[]>(ws, {
          header: 1,
          raw: true,
          defval: '',
          blankrows: false,
          dateNF: 'dd.mm.yyyy',
        });

        if (!rows.length) {
          throw new Error('Visit history sheet is empty');
        }

        const headerIndex = findVisitHistoryHeaderIndex(rows);
        const headers = rows[headerIndex];

        const parsedRows: Array<VisitHistoryEntry | null> = rows
          .slice(headerIndex + 1)
          .filter((row) => row.some((cell) => String(cell ?? '').trim() !== ''))
          .map((row) => {
            const rowObject = buildVisitHistoryRowObject(headers, row);
            const parsedByHeaders = parseVisitHistoryRow(rowObject);

            if (parsedByHeaders) return parsedByHeaders;

            const date = formatVisitHistoryDateValue(row[0]);
            const routeCode = String(row[1] ?? '').trim();
            const clientId = String(row[2] ?? '').trim();
            const name = String(row[3] ?? '').trim();
            const address = String(row[4] ?? '').trim();
            const coordinateDeviationMeters = parseLocalizedVisitHistoryNumber(row[5]);
            const orderAmountRub = parseLocalizedVisitHistoryNumber(row[6]);

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
          });

        const data = parsedRows.filter((entry): entry is VisitHistoryEntry => entry !== null);

        onVisitHistoryLoaded(data);
        setLoadedRowsCount(data.length);
        setSuccess(true);
        setTimeout(() => setSuccess(false), 3000);
      } catch (err) {
        console.error('Error parsing visit history Excel:', err);
        setError('Не удалось прочитать файл. Убедитесь, что это корректный Excel файл истории визитов.');
      } finally {
        setLoading(false);
      }
    };

    reader.readAsBinaryString(file);
  };

  return (
    <div className="p-6 bg-white rounded-lg shadow-sm border border-gray-200 text-center">
      <div className="mb-4 flex justify-center text-violet-500">
        <FileSpreadsheet size={48} />
      </div>
      <h3 className="text-lg font-medium text-gray-900 mb-2">История визитов</h3>
      <p className="text-sm text-gray-500 mb-6">
        Загрузите Excel файл с историей фактических визитов
      </p>

      <div className="flex flex-col items-center gap-3">
        <div className="relative inline-block">
          <input
            type="file"
            accept=".xlsx, .xls"
            onChange={handleFileUpload}
            className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
            disabled={loading}
          />
          <button
            className="flex items-center space-x-2 px-4 py-2 bg-violet-600 text-white rounded-md hover:bg-violet-700 transition-colors disabled:opacity-50"
            disabled={loading}
          >
            {loading ? <Loader2 className="animate-spin" size={20} /> : <Upload size={20} />}
            <span>{loading ? 'Обработка...' : 'Выберите файл'}</span>
          </button>
        </div>

        <button
          type="button"
          onClick={() => downloadTemplateWorkbook('template-visit-history.xlsx', visitHistoryTemplateHeaders, visitHistoryTemplateRows)}
          className="inline-flex items-center space-x-2 px-4 py-2 text-sm font-medium text-violet-700 bg-violet-50 border border-violet-200 rounded-md hover:bg-violet-100 transition-colors"
        >
          <Download size={16} />
          <span>Скачать шаблон с примером</span>
        </button>
      </div>

      {success && (
        <div className="mt-4 text-sm text-green-600 bg-green-50 p-2 rounded space-y-1">
          <div>✓ Файл успешно загружен</div>
          <div>
            Загружено визитов: <span className="font-semibold">{loadedRowsCount}</span>
          </div>
        </div>
      )}

      {error && (
        <div className="mt-4 text-sm text-red-600 bg-red-50 p-2 rounded">
          {error}
        </div>
      )}

      <div className="mt-6 text-left text-xs text-gray-400">
        <p className="font-semibold mb-1">Ожидаемый формат:</p>
        <code className="bg-gray-100 p-1 rounded block overflow-x-auto text-xs">
          Дата | Маршрут | ИД клиента | Название | Адрес | Отклонение координат ТТ и визита м | Сумма заказа руб
        </code>
      </div>
    </div>
  );
}

export function FileImportBackup({
  entries,
  deliveryScheduleEntries,
  visitHistoryEntries,
  onBackupRestore,
}: FileImportBackupProps) {
  const [includeRoutes, setIncludeRoutes] = useState(true);
  const [includeDelivery, setIncludeDelivery] = useState(true);
  const [includeVisits, setIncludeVisits] = useState(true);
  const [error, setError] = useState<string | null>(null);
  const [success, setSuccess] = useState<string | null>(null);

  const backupPreview = useMemo(() => ({
    entries: includeRoutes ? entries.length : 0,
    delivery: includeDelivery ? deliveryScheduleEntries.length : 0,
    visits: includeVisits ? visitHistoryEntries.length : 0,
  }), [includeRoutes, includeDelivery, includeVisits, entries.length, deliveryScheduleEntries.length, visitHistoryEntries.length]);

  const hasAnythingSelected = includeRoutes || includeDelivery || includeVisits;
  const hasAnythingToBackup = backupPreview.entries > 0 || backupPreview.delivery > 0 || backupPreview.visits > 0;

  const handleDownloadBackup = () => {
    setError(null);
    setSuccess(null);

    if (!hasAnythingSelected) {
      setError('Выберите хотя бы один тип данных для бэкапа.');
      return;
    }

    if (!hasAnythingToBackup) {
      setError('Нет данных для бэкапа по выбранным типам.');
      return;
    }

    const payload: BackupPayload = {
      version: 1,
      exportedAt: new Date().toISOString(),
      ...(includeRoutes ? { entries } : {}),
      ...(includeDelivery ? { deliveryScheduleEntries } : {}),
      ...(includeVisits ? { visitHistoryEntries } : {}),
    };

    const blob = new Blob([JSON.stringify(payload, null, 2)], { type: 'application/json;charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    const timestamp = new Date().toISOString().slice(0, 19).replace(/[T:]/g, '-');
    link.href = url;
    link.download = `calendar-routes-backup-${timestamp}.json`;
    link.click();
    URL.revokeObjectURL(url);

    setSuccess('Бэкап успешно сохранен');
    setTimeout(() => setSuccess(null), 3000);
  };

  const handleRestoreBackup = (e: React.ChangeEvent<HTMLInputElement>) => {
    const file = e.target.files?.[0];
    if (!file) return;

    setError(null);
    setSuccess(null);

    const reader = new FileReader();
    reader.onload = (evt) => {
      try {
        const text = String(evt.target?.result ?? '');
        const parsed = JSON.parse(text) as BackupPayload;

        if (!parsed || typeof parsed !== 'object' || parsed.version !== 1) {
          throw new Error('Unsupported backup format');
        }

        onBackupRestore(parsed);
        setSuccess('Бэкап успешно восстановлен');
        setTimeout(() => setSuccess(null), 3000);
      } catch (err) {
        console.error('Error restoring backup:', err);
        setError('Не удалось восстановить бэкап. Убедитесь, что это корректный JSON-файл бэкапа.');
      }
    };

    reader.readAsText(file, 'utf-8');
    e.target.value = '';
  };

  return (
    <div className="p-6 bg-white rounded-lg shadow-sm border border-gray-200 text-center">
      <div className="mb-4 flex justify-center text-slate-600">
        <ArchiveRestore size={48} />
      </div>
      <h3 className="text-lg font-medium text-gray-900 mb-2">Бэкап</h3>
      <p className="text-sm text-gray-500 mb-4">
        Сохранение и восстановление текущих загруженных данных
      </p>

      <div className="space-y-2 text-left text-sm bg-gray-50 border border-gray-200 rounded-lg p-3 mb-4">
        <label className="flex items-center gap-2">
          <input type="checkbox" checked={includeRoutes} onChange={(e) => setIncludeRoutes(e.target.checked)} />
          <span>Маршруты ({entries.length})</span>
        </label>
        <label className="flex items-center gap-2">
          <input type="checkbox" checked={includeDelivery} onChange={(e) => setIncludeDelivery(e.target.checked)} />
          <span>График доставки ({deliveryScheduleEntries.length})</span>
        </label>
        <label className="flex items-center gap-2">
          <input type="checkbox" checked={includeVisits} onChange={(e) => setIncludeVisits(e.target.checked)} />
          <span>История визитов ({visitHistoryEntries.length})</span>
        </label>
      </div>

      <div className="flex flex-col items-center gap-3">
        <button
          type="button"
          onClick={handleDownloadBackup}
          className="inline-flex items-center space-x-2 px-4 py-2 text-sm font-medium text-slate-700 bg-slate-100 border border-slate-300 rounded-md hover:bg-slate-200 transition-colors"
        >
          <Save size={16} />
          <span>Скачать бэкап</span>
        </button>

        <div className="relative inline-block">
          <input
            type="file"
            accept=".json"
            onChange={handleRestoreBackup}
            className="absolute inset-0 w-full h-full opacity-0 cursor-pointer"
          />
          <button className="inline-flex items-center space-x-2 px-4 py-2 text-sm font-medium text-slate-700 bg-white border border-slate-300 rounded-md hover:bg-slate-50 transition-colors">
            <ArchiveRestore size={16} />
            <span>Восстановить бэкап</span>
          </button>
        </div>
      </div>

      {success && (
        <div className="mt-4 text-sm text-green-600 bg-green-50 p-2 rounded">
          {success}
        </div>
      )}

      {error && (
        <div className="mt-4 text-sm text-red-600 bg-red-50 p-2 rounded">
          {error}
        </div>
      )}

      <div className="mt-6 text-left text-xs text-gray-400">
        <p className="font-semibold mb-1">Что входит в бэкап:</p>
        <code className="bg-gray-100 p-1 rounded block overflow-x-auto text-xs">
          Маршруты | График доставки | История визитов
        </code>
      </div>
    </div>
  );
}
