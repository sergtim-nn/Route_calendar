import { useState } from 'react';
import { CalendarGrid } from './components/CalendarGrid';
import { FileImport, FileImportDeliverySchedule, FileImportEasyMerch } from './components/FileImport';
import { DeliveryScheduleEntry, ScheduleEntry } from './utils/schedule';
import { parseDeliveryScheduleRow, parseEasyMerchRow, parseScheduleRow } from './utils/parser';
import { Calendar, RefreshCw, ArrowRight } from 'lucide-react';

function App() {
  const [entries, setEntries] = useState<ScheduleEntry[]>([]);
  const [deliveryScheduleEntries, setDeliveryScheduleEntries] = useState<DeliveryScheduleEntry[]>([]);
  const [showCalendar, setShowCalendar] = useState(false);

  const handleDataLoaded = (data: any[]) => {
    const normalizedData: ScheduleEntry[] = data
      .map(parseScheduleRow)
      .filter((entry): entry is ScheduleEntry => entry !== null);

    setEntries(prev => [...prev, ...normalizedData]);
  };

  const handleEasyMerchLoaded = (data: any[], headers: string[]) => {
    const normalizedData: ScheduleEntry[] = data
      .flatMap((row) => parseEasyMerchRow(row, headers))
      .filter((entry): entry is ScheduleEntry => entry !== null);

    if (normalizedData.length === 0) {
      console.warn('EasyMerch file loaded but produced 0 schedule entries', { dataPreview: data.slice(0, 3), headers });
    }

    setEntries((prev) => [...prev, ...normalizedData]);
  };

  const handleDeliveryScheduleLoaded = (data: any[]) => {
    const normalizedData: DeliveryScheduleEntry[] = data
      .map(parseDeliveryScheduleRow)
      .filter((entry): entry is DeliveryScheduleEntry => entry !== null);

    setDeliveryScheduleEntries((prev) => [...prev, ...normalizedData]);
  };

  const handleReset = () => {
    setEntries([]);
    setDeliveryScheduleEntries([]);
    setShowCalendar(false);
  };

  const handleGoToCalendar = () => {
    if (entries.length > 0 || deliveryScheduleEntries.length > 0) {
      setShowCalendar(true);
      window.scrollTo({ top: 0, behavior: 'smooth' });
    }
  };

  const shouldShowUploadScreen = !showCalendar;

  return (
    <div className="min-h-screen bg-gray-100 font-sans text-gray-900">
      <header className="sticky top-0 z-[300] bg-white shadow-sm border-b border-gray-200">
        <div className="w-full px-4 sm:px-6 lg:px-8 h-16 flex items-center justify-between">
          <div className="flex items-center space-x-3">
            <div className="bg-blue-600 p-2 rounded-lg text-white">
              <Calendar size={24} />
            </div>
            <h1 className="text-xl font-bold text-gray-900">Календарь маршрутов</h1>
          </div>

          {(entries.length > 0 || deliveryScheduleEntries.length > 0) && (
            <button
              onClick={handleReset}
              className="flex items-center space-x-2 text-sm text-gray-500 hover:text-red-600 transition-colors"
            >
              <RefreshCw size={16} />
              <span>Сбросить данные</span>
            </button>
          )}
        </div>
      </header>

      <main className="w-full px-4 sm:px-6 lg:px-8 py-4 h-[calc(100vh-64px)] overflow-hidden">
        {shouldShowUploadScreen ? (
          <div className="h-full flex flex-col items-center justify-center">
            <div className="max-w-[1600px] w-full space-y-6">
              <div className="grid grid-cols-1 xl:grid-cols-3 gap-4">
                <FileImport onDataLoaded={handleDataLoaded} />
                <FileImportEasyMerch onEasyMerchLoaded={handleEasyMerchLoaded} />
                <FileImportDeliverySchedule onDeliveryScheduleLoaded={handleDeliveryScheduleLoaded} />
              </div>

              <div className="flex flex-col items-center gap-3">
                {(entries.length > 0 || deliveryScheduleEntries.length > 0) && (
                  <div className="text-sm text-gray-600 bg-white border border-gray-200 rounded-lg px-4 py-2 shadow-sm text-center">
                    <div>
                      Загружено маршрутных записей: <span className="font-semibold text-gray-900">{entries.length}</span>
                    </div>
                    <div>
                      Загружено зон графика доставки: <span className="font-semibold text-gray-900">{deliveryScheduleEntries.length}</span>
                    </div>
                  </div>
                )}

                <button
                  onClick={handleGoToCalendar}
                  disabled={entries.length === 0 && deliveryScheduleEntries.length === 0}
                  className="flex items-center space-x-2 px-6 py-3 bg-blue-600 text-white rounded-lg hover:bg-blue-700 disabled:bg-gray-300 disabled:cursor-not-allowed transition-colors font-medium"
                >
                  <span>Перейти в календарь</span>
                  <ArrowRight size={20} />
                </button>
              </div>

              <div className="mt-8 bg-blue-50 p-4 rounded-lg border border-blue-100 text-sm text-blue-800">
                <p className="font-semibold mb-2">Как это работает:</p>
                <ul className="list-disc list-inside space-y-1">
                  <li>Можно загрузить «Загрузка маршрутов», «EasyMerch» и «График доставки» в любом порядке, не переходя в календарь.</li>
                  <li>Все три карточки загрузки расположены рядом по горизонтали на широком экране.</li>
                  <li>Данные из разных файлов сохраняются в рамках одной сессии.</li>
                  <li>Переход в календарь выполняется только по кнопке «Перейти в календарь».</li>
                </ul>
              </div>
            </div>
          </div>
        ) : (
          <div className="h-full min-h-0 flex flex-col overflow-hidden">
            <CalendarGrid entries={entries} deliveryScheduleEntries={deliveryScheduleEntries} />
          </div>
        )}
      </main>
    </div>
  );
}

export default App;
