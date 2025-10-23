# Документация по классам для работы с табличным редактором (TypeScript)

## Класс `Workbook` (Рабочая книга)

Класс для управления рабочей книгой табличного редактора.

### Конструктор

```typescript
constructor(filename?: string)
```
Создает новую рабочую книгу или загружает существующую.

**Параметры:**
- `filename` - путь к файлу для загрузки (опционально)

### Методы

#### Управление листами

```typescript
getSheets(): Sheet[]
```
Возвращает список всех листов в книге.

**Возвращает:**
- Массив объектов `Sheet`

---

```typescript
addSheet(name: string, position?: number): Sheet
```
Добавляет новый лист в книгу.

**Параметры:**
- `name` - название листа
- `position` - позиция для вставки (опционально)

**Возвращает:**
- Созданный объект `Sheet`

---

```typescript
removeSheet(sheet: Sheet | string): boolean
```
Удаляет лист из книги.

**Параметры:**
- `sheet` - объект Sheet или название листа

**Возвращает:**
- `true` если удаление успешно, `false` в противном случае

---

```typescript
renameSheet(sheet: Sheet | string, newName: string): boolean
```
Переименовывает лист.

**Параметры:**
- `sheet` - объект Sheet или название листа
- `newName` - новое название

**Возвращает:**
- `true` если переименование успешно, `false` в противном случае

---

```typescript
getActiveSheet(): Sheet | null
```
Возвращает активный лист.

**Возвращает:**
- Активный объект `Sheet` или `null`

---

```typescript
setActiveSheet(sheet: Sheet | string): boolean
```
Устанавливает активный лист.

**Параметры:**
- `sheet` - объект Sheet или название листа

**Возвращает:**
- `true` если установка успешна, `false` в противном случае

#### Операции с файлом

```typescript
save(filename?: string): boolean
```
Сохраняет рабочую книгу.

**Параметры:**
- `filename` - путь для сохранения (опционально)

**Возвращает:**
- `true` если сохранение успешно, `false` в противном случае

---

```typescript
load(filename: string): boolean
```
Загружает рабочую книгу из файла.

**Параметры:**
- `filename` - путь к файлу

**Возвращает:**
- `true` если загрузка успешна, `false` в противном случае

---

```typescript
exportAs(format: ExportFormat, filename: string): boolean
```
Экспортирует книгу в указанном формате.

**Параметры:**
- `format` - формат экспорта
- `filename` - путь для экспорта

**Возвращает:**
- `true` если экспорт успешен, `false` в противном случае

## Класс `Sheet` (Лист)

Класс для работы с отдельным листом в рабочей книге.

### Конструктор

```typescript
constructor(workbook: Workbook, name: string)
```
Создает новый лист.

**Параметры:**
- `workbook` - родительская рабочая книга
- `name` - название листа

### Свойства

```typescript
readonly name: string
```
Название листа (только для чтения).

---

```typescript
readonly workbook: Workbook
```
Родительская рабочая книга (только для чтения).

### Методы

#### Работа с ячейками

```typescript
getCell(row: number, column: number): Cell
```
Возвращает ячейку по координатам.

**Параметры:**
- `row` - номер строки (начиная с 1)
- `column` - номер колонки (начиная с 1)

**Возвращает:**
- Объект `Cell`

---

```typescript
getRange(startRow: number, startCol: number, endRow: number, endCol: number): Range
```
Возвращает диапазон ячеек.

**Параметры:**
- `startRow` - начальная строка
- `startCol` - начальная колонка
- `endRow` - конечная строка
- `endCol` - конечная колонка

**Возвращает:**
- Объект `Range`

---

```typescript
getUsedRange(): Range | null
```
Возвращает используемый диапазон (область с данными).

**Возвращает:**
- Объект `Range` или `null`

#### Управление строками и колонками

```typescript
insertRow(row: number, count: number = 1): boolean
```
Вставляет строки.

**Параметры:**
- `row` - номер строки для вставки
- `count` - количество строк (по умолчанию 1)

**Возвращает:**
- `true` если операция успешна

---

```typescript
deleteRow(row: number, count: number = 1): boolean
```
Удаляет строки.

**Параметры:**
- `row` - номер строки для удаления
- `count` - количество строк (по умолчанию 1)

**Возвращает:**
- `true` если операция успешна

---

```typescript
insertColumn(column: number, count: number = 1): boolean
```
Вставляет колонки.

**Параметры:**
- `column` - номер колонки для вставки
- `count` - количество колонок (по умолчанию 1)

**Возвращает:**
- `true` если операция успешна

---

```typescript
deleteColumn(column: number, count: number = 1): boolean
```
Удаляет колонки.

**Параметры:**
- `column` - номер колонки для удаления
- `count` - количество колонок (по умолчанию 1)

**Возвращает:**
- `true` если операция успешна

#### Графики

```typescript
addChart(dataRange: Range, chartType: ChartType, position: ChartPosition): Chart
```
Добавляет график на лист.

**Параметры:**
- `dataRange` - диапазон данных для графика
- `chartType` - тип графика
- `position` - позиция графика на листе

**Возвращает:**
- Объект `Chart`

---

```typescript
getCharts(): Chart[]
```
Возвращает все графики на листе.

**Возвращает:**
- Массив объектов `Chart`

## Класс `Cell` (Ячейка)

Класс для работы с отдельной ячейкой.

### Свойства

```typescript
value: any
```
Значение ячейки.

---

```typescript
readonly row: number
```
Номер строки (только для чтения).

---

```typescript
readonly column: number
```
Номер колонки (только для чтения).

---

```typescript
readonly sheet: Sheet
```
Родительский лист (только для чтения).

### Методы

```typescript
getFormattedValue(): string
```
Возвращает отформатированное значение ячейки.

**Возвращает:**
- Отформатированное значение как строка

---

```typescript
clear(): void
```
Очищает ячейку (значение и форматирование).

## Класс `Range` (Диапазон)

Класс для работы с диапазоном ячеек.

### Свойства

```typescript
readonly startRow: number
```
Начальная строка (только для чтения).

---

```typescript
readonly startCol: number
```
Начальная колонка (только для чтения).

---

```typescript
readonly endRow: number
```
Конечная строка (только для чтения).

---

```typescript
readonly endCol: number
```
Конечная колонка (только для чтения).

---

```typescript
readonly sheet: Sheet
```
Родительский лист (только для чтения).

### Методы

#### Получение данных

```typescript
getValues(): any[][]
```
Возвращает значения диапазона как двумерный массив.

**Возвращает:**
- Двумерный массив значений

---

```typescript
setValues(values: any[][]): boolean
```
Устанавливает значения для диапазона.

**Параметры:**
- `values` - двумерный массив значений

**Возвращает:**
- `true` если операция успешна

#### Форматирование

```typescript
setBackgroundColor(color: string): void
```
Устанавливает цвет заливки для всего диапазона.

**Параметры:**
- `color` - цвет в формате HEX (#RRGGBB)

---

```typescript
setFontSize(size: number): void
```
Устанавливает размер шрифта для всего диапазона.

**Параметры:**
- `size` - размер шрифта

---

```typescript
setFontColor(color: string): void
```
Устанавливает цвет шрифта для всего диапазона.

**Параметры:**
- `color` - цвет в формате HEX (#RRGGBB)

---

```typescript
setBold(isBold: boolean = true): void
```
Устанавливает жирное начертание.

**Параметры:**
- `isBold` - флаг жирного начертания

---

```typescript
setItalic(isItalic: boolean = true): void
```
Устанавливает курсивное начертание.

**Параметры:**
- `isItalic` - флаг курсивного начертания

---

```typescript
setBorder(style: BorderStyle, color?: string): void
```
Устанавливает границы для диапазона.

**Параметры:**
- `style` - стиль границы
- `color` - цвет границы (опционально)

#### Операции

```typescript
merge(): boolean
```
Объединяет ячейки диапазона.

**Возвращает:**
- `true` если операция успешна

---

```typescript
unmerge(): boolean
```
Разъединяет ячейки диапазона.

**Возвращает:**
- `true` если операция успешна

---

```typescript
clear(): void
```
Очищает значения и форматирование диапазона.

## Класс `Chart` (График)

Класс для работы с графиками.

### Свойства

```typescript
readonly id: string
```
Уникальный идентификатор графика (только для чтения).

---

```typescript
readonly sheet: Sheet
```
Родительский лист (только для чтения).

### Методы

```typescript
setTitle(title: string): void
```
Устанавливает заголовок графика.

**Параметры:**
- `title` - заголовок

---

```typescript
setPosition(position: ChartPosition): void
```
Устанавливает позицию графика на листе.

**Параметры:**
- `position` - позиция графика

---

```typescript
setDataRange(range: Range): boolean
```
Устанавливает диапазон данных для графика.

**Параметры:**
- `range` - диапазон данных

**Возвращает:**
- `true` если операция успешна

---

```typescript
remove(): boolean
```
Удаляет график с листа.

**Возвращает:**
- `true` если удаление успешно

## Типы

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
```

## Примеры использования

### Создание книги и добавление данных

```typescript
// Создание новой рабочей книги
const workbook = new Workbook();

// Добавление листа
const sheet = workbook.addSheet('Данные');

// Запись данных в ячейки
sheet.getCell(1, 1).value = 'Название';
sheet.getCell(1, 2).value = 'Количество';
sheet.getCell(2, 1).value = 'Товар A';
sheet.getCell(2, 2).value = 100;
sheet.getCell(3, 1).value = 'Товар B';
sheet.getCell(3, 2).value = 150;

// Форматирование заголовков
const headerRange = sheet.getRange(1, 1, 1, 2);
headerRange.setBackgroundColor('#4CAF50');
headerRange.setFontColor('#FFFFFF');
headerRange.setBold(true);

// Сохранение книги
workbook.save('отчет.xlsx');
```

### Работа с диапазонами и графиками

```typescript
// Получение диапазона данных
const dataRange = sheet.getRange(1, 1, 3, 2);

// Добавление графика
const chart = sheet.addChart(
  dataRange,
  'column',
  { row: 5, column: 1, width: 600, height: 400 }
);

chart.setTitle('Продажи по товарам');

// Экспорт в PDF
workbook.exportAs('pdf', 'отчет.pdf');
```

### Управление листами

```typescript
// Переименование листа
workbook.renameSheet('Данные', 'Статистика');

// Добавление нового листа
const summarySheet = workbook.addSheet('Итоги', 0); // В начало

// Установка активного листа
workbook.setActiveSheet('Итоги');

// Удаление листа
workbook.removeSheet('Статистика');
```