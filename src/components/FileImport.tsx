import { useState } from 'react';
import * as XLSX from 'xlsx';
import { Upload, FileSpreadsheet, Loader2, Download } from 'lucide-react';

interface FileImportProps {
  onDataLoaded: (data: any[]) => void;
}

interface FileImportEasyMerchProps {
  onEasyMerchLoaded: (data: any[], headers: string[]) => void;
}

interface FileImportDeliveryScheduleProps {
  onDeliveryScheduleLoaded: (data: any[]) => void;
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
