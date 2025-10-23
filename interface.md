# Документация интерфейсов для работы с табличным редактором (TypeScript)

## Основные типы

```typescript
type ExportFormat = 'xlsx' | 'csv' | 'pdf' | 'html';
type ChartType = 'line' | 'bar' | 'column' | 'pie' | 'area' | 'scatter';
type BorderStyle = 'thin' | 'medium' | 'thick' | 'dashed' | 'dotted';

interface ChartPosition {
  row: number;
  column: number;
  width?: number;
  height?: number;
}

interface CellAddress {
  row: number;
  column: number;
}
```

## Интерфейс `IWorkbook` (Рабочая книга)

```typescript
interface IWorkbook {
  /**
   * Возвращает список всех листов в книге
   */
  getSheets(): ISheet[];

  /**
   * Добавляет новый лист в книгу
   * @param name - название листа
   * @param position - позиция для вставки (опционально)
   */
  addSheet(name: string, position?: number): ISheet;

  /**
   * Удаляет лист из книги
   * @param sheet - объект ISheet или название листа
   */
  removeSheet(sheet: ISheet | string): boolean;

  /**
   * Переименовывает лист
   * @param sheet - объект ISheet или название листа
   * @param newName - новое название
   */
  renameSheet(sheet: ISheet | string, newName: string): boolean;

  /**
   * Возвращает активный лист
   */
  getActiveSheet(): ISheet | null;

  /**
   * Устанавливает активный лист
   * @param sheet - объект ISheet или название листа
   */
  setActiveSheet(sheet: ISheet | string): boolean;

  /**
   * Сохраняет рабочую книгу
   * @param filename - путь для сохранения (опционально)
   */
  save(filename?: string): boolean;

  /**
   * Загружает рабочую книгу из файла
   * @param filename - путь к файлу
   */
  load(filename: string): boolean;

  /**
   * Экспортирует книгу в указанном формате
   * @param format - формат экспорта
   * @param filename - путь для экспорта
   */
  exportAs(format: ExportFormat, filename: string): boolean;
}
```

## Интерфейс `ISheet` (Лист)

```typescript
interface ISheet {
  readonly name: string;
  readonly workbook: IWorkbook;

  /**
   * Возвращает ячейку по координатам
   * @param row - номер строки (начиная с 1)
   * @param column - номер колонки (начиная с 1)
   */
  getCell(row: number, column: number): ICell;

  /**
   * Возвращает диапазон ячеек
   * @param startRow - начальная строка
   * @param startCol - начальная колонка
   * @param endRow - конечная строка
   * @param endCol - конечная колонка
   */
  getRange(startRow: number, startCol: number, endRow: number, endCol: number): IRange;

  /**
   * Возвращает используемый диапазон (область с данными)
   */
  getUsedRange(): IRange | null;

  /**
   * Вставляет строки
   * @param row - номер строки для вставки
   * @param count - количество строк (по умолчанию 1)
   */
  insertRow(row: number, count?: number): boolean;

  /**
   * Удаляет строки
   * @param row - номер строки для удаления
   * @param count - количество строк (по умолчанию 1)
   */
  deleteRow(row: number, count?: number): boolean;

  /**
   * Вставляет колонки
   * @param column - номер колонки для вставки
   * @param count - количество колонок (по умолчанию 1)
   */
  insertColumn(column: number, count?: number): boolean;

  /**
   * Удаляет колонки
   * @param column - номер колонки для удаления
   * @param count - количество колонок (по умолчанию 1)
   */
  deleteColumn(column: number, count?: number): boolean;

  /**
   * Добавляет график на лист
   * @param dataRange - диапазон данных для графика
   * @param chartType - тип графика
   * @param position - позиция графика на листе
   */
  addChart(dataRange: IRange, chartType: ChartType, position: ChartPosition): IChart;

  /**
   * Возвращает все графики на листе
   */
  getCharts(): IChart[];
}
```

## Интерфейс `ICell` (Ячейка)

```typescript
interface ICell {
  value: any;
  readonly row: number;
  readonly column: number;
  readonly sheet: ISheet;

  /**
   * Возвращает отформатированное значение ячейки
   */
  getFormattedValue(): string;

  /**
   * Очищает ячейку (значение и форматирование)
   */
  clear(): void;

  /**
   * Устанавливает цвет заливки ячейки
   * @param color - цвет в формате HEX (#RRGGBB)
   */
  setBackgroundColor(color: string): void;

  /**
   * Устанавливает размер шрифта ячейки
   * @param size - размер шрифта
   */
  setFontSize(size: number): void;

  /**
   * Устанавливает цвет шрифта ячейки
   * @param color - цвет в формате HEX (#RRGGBB)
   */
  setFontColor(color: string): void;

  /**
   * Устанавливает жирное начертание
   * @param isBold - флаг жирного начертания
   */
  setBold(isBold?: boolean): void;

  /**
   * Устанавливает курсивное начертание
   * @param isItalic - флаг курсивного начертания
   */
  setItalic(isItalic?: boolean): void;
}
```

## Интерфейс `IRange` (Диапазон)

```typescript
interface IRange {
  readonly startRow: number;
  readonly startCol: number;
  readonly endRow: number;
  readonly endCol: number;
  readonly sheet: ISheet;

  /**
   * Возвращает значения диапазона как двумерный массив
   */
  getValues(): any[][];

  /**
   * Устанавливает значения для диапазона
   * @param values - двумерный массив значений
   */
  setValues(values: any[][]): boolean;

  /**
   * Устанавливает цвет заливки для всего диапазона
   * @param color - цвет в формате HEX (#RRGGBB)
   */
  setBackgroundColor(color: string): void;

  /**
   * Устанавливает размер шрифта для всего диапазона
   * @param size - размер шрифта
   */
  setFontSize(size: number): void;

  /**
   * Устанавливает цвет шрифта для всего диапазона
   * @param color - цвет в формате HEX (#RRGGBB)
   */
  setFontColor(color: string): void;

  /**
   * Устанавливает жирное начертание
   * @param isBold - флаг жирного начертания
   */
  setBold(isBold?: boolean): void;

  /**
   * Устанавливает курсивное начертание
   * @param isItalic - флаг курсивного начертания
   */
  setItalic(isItalic?: boolean): void;

  /**
   * Устанавливает границы для диапазона
   * @param style - стиль границы
   * @param color - цвет границы (опционально)
   */
  setBorder(style: BorderStyle, color?: string): void;

  /**
   * Объединяет ячейки диапазона
   */
  merge(): boolean;

  /**
   * Разъединяет ячейки диапазона
   */
  unmerge(): boolean;

  /**
   * Очищает значения и форматирование диапазона
   */
  clear(): void;

  /**
   * Возвращает отдельную ячейку из диапазона
   * @param relativeRow - относительная строка (от начала диапазона)
   * @param relativeCol - относительная колонка (от начала диапазона)
   */
  getCell(relativeRow: number, relativeCol: number): ICell;
}
```

## Интерфейс `IChart` (График)

```typescript
interface IChart {
  readonly id: string;
  readonly sheet: ISheet;

  /**
   * Устанавливает заголовок графика
   * @param title - заголовок
   */
  setTitle(title: string): void;

  /**
   * Устанавливает позицию графика на листе
   * @param position - позиция графика
   */
  setPosition(position: ChartPosition): void;

  /**
   * Устанавливает диапазон данных для графика
   * @param range - диапазон данных
   */
  setDataRange(range: IRange): boolean;

  /**
   * Устанавливает тип графика
   * @param chartType - тип графика
   */
  setChartType(chartType: ChartType): void;

  /**
   * Устанавливает отображение легенды
   * @param visible - видимость легенды
   */
  setLegend(visible: boolean): void;

  /**
   * Удаляет график с листа
   */
  remove(): boolean;
}
```

## Примеры использования интерфейсов

### Создание и работа с книгой

```typescript
// Создание реализации рабочей книги
class Workbook implements IWorkbook {
  // Реализация методов интерфейса...
  getSheets(): ISheet[] { /* ... */ }
  addSheet(name: string, position?: number): ISheet { /* ... */ }
  // ... остальные методы
}

// Использование
const workbook: IWorkbook = new Workbook();
const sheet: ISheet = workbook.addSheet('Данные');

// Работа с ячейками
const cell: ICell = sheet.getCell(1, 1);
cell.value = 'Заголовок';
cell.setBold(true);
cell.setBackgroundColor('#FF0000');

// Работа с диапазонами
const range: IRange = sheet.getRange(1, 1, 5, 3);
range.setValues([
  ['Товар', 'Цена', 'Количество'],
  ['Товар A', 100, 50],
  ['Товар B', 200, 30]
]);
range.setBorder('thin', '#000000');

// Создание графика
const chart: IChart = sheet.addChart(
  range,
  'column',
  { row: 7, column: 1, width: 600, height: 400 }
);
chart.setTitle('Статистика продаж');

// Сохранение
workbook.save('отчет.xlsx');
```

### Работа с несколькими листами

```typescript
function createReport(workbook: IWorkbook): void {
  // Создание листов
  const dataSheet: ISheet = workbook.addSheet('Данные');
  const summarySheet: ISheet = workbook.addSheet('Итоги');
  
  // Заполнение данными
  const dataRange: IRange = dataSheet.getRange(1, 1, 10, 4);
  // ... заполнение данных
  
  // Установка активного листа
  workbook.setActiveSheet(summarySheet);
  
  // Переименование
  workbook.renameSheet('Данные', 'Исходные данные');
}
```