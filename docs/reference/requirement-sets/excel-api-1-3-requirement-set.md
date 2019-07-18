---
title: Набор обязательных элементов API JavaScript для Excel 1,3
description: Сведения о наборе требований ExcelApi 1,3
ms.date: 07/11/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 4698b0fad3122c8ecf52117c35d4928305d812fc
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771997"
---
# <a name="whats-new-in-excel-javascript-api-13"></a>Новые возможности API JavaScript для Excel 1.3

ExcelApi 1,3 добавлена поддержка привязки данных и базового доступа к сводным таблицам.

## <a name="api-list"></a>Список API

| Класс | Поля | Описание |
|:---|:---|:---|
|[Binding](/javascript/api/excel/excel.binding)|[delete()](/javascript/api/excel/excel.binding#delete--)|Удаляет привязку.|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[Add (Range: Range \| String, bindingType: "Range" \| "Table" \| "Text", ID: String)](/javascript/api/excel/excel.bindingcollection#add-range--bindingtype--id-)|Добавляет привязку к определенному объекту Range.|
||[Add (Range: строка \| Range, BindingType: Excel. bindingType, ID: строка)](/javascript/api/excel/excel.bindingcollection#add-range--bindingtype--id-)|Добавляет привязку к определенному объекту Range.|
||[Аддфромнамедитем (Name: строка, bindingType: "Range" \| "Table" \| "Text", ID: String)](/javascript/api/excel/excel.bindingcollection#addfromnameditem-name--bindingtype--id-)|Добавляет новую привязку с учетом именованного элемента в книге.|
||[Аддфромнамедитем (имя: строка, bindingType: Excel. BindingType, ID: строка)](/javascript/api/excel/excel.bindingcollection#addfromnameditem-name--bindingtype--id-)|Добавляет новую привязку с учетом именованного элемента в книге.|
||[Аддфромселектион (bindingType: "Range" \| "Table" \| "Text", ID: строка)](/javascript/api/excel/excel.bindingcollection#addfromselection-bindingtype--id-)|Добавляет новую привязку с учетом выделенного в настоящий момент фрагмента.|
||[Аддфромселектион (bindingType: Excel. BindingType, ID: String)](/javascript/api/excel/excel.bindingcollection#addfromselection-bindingtype--id-)|Добавляет новую привязку с учетом выделенного в настоящий момент фрагмента.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[name](/javascript/api/excel/excel.pivottable#name)|Имя сводной таблицы.|
||[worksheet](/javascript/api/excel/excel.pivottable#worksheet)|Лист, содержащий текущую сводную таблицу.|
||[refresh()](/javascript/api/excel/excel.pivottable#refresh--)|Обновляет сводную таблицу.|
||[Set (Properties: Excel. PivotTable)](/javascript/api/excel/excel.pivottable#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Пивоттаблеупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.pivottable#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getItem(name: string)](/javascript/api/excel/excel.pivottablecollection#getitem-name-)|Получает сводную таблицу по имени.|
||[items](/javascript/api/excel/excel.pivottablecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[refreshAll ()](/javascript/api/excel/excel.pivottablecollection#refreshall--)|Обновляет все сводные таблицы в коллекции.|
|[Пивоттаблеколлектионлоадоптионс](/javascript/api/excel/excel.pivottablecollectionloadoptions)|[$all](/javascript/api/excel/excel.pivottablecollectionloadoptions#$all)||
||[name](/javascript/api/excel/excel.pivottablecollectionloadoptions#name)|Для каждого элемента в коллекции: имя сводной таблицы.|
||[worksheet](/javascript/api/excel/excel.pivottablecollectionloadoptions#worksheet)|Для каждого элемента в коллекции: лист, содержащий текущую сводную таблицу.|
|[Пивоттабледата](/javascript/api/excel/excel.pivottabledata)|[name](/javascript/api/excel/excel.pivottabledata#name)|Имя сводной таблицы.|
|[Пивоттаблелоадоптионс](/javascript/api/excel/excel.pivottableloadoptions)|[$all](/javascript/api/excel/excel.pivottableloadoptions#$all)||
||[name](/javascript/api/excel/excel.pivottableloadoptions#name)|Имя сводной таблицы.|
||[worksheet](/javascript/api/excel/excel.pivottableloadoptions#worksheet)|Лист, содержащий текущую сводную таблицу.|
|[Пивоттаблеупдатедата](/javascript/api/excel/excel.pivottableupdatedata)|[name](/javascript/api/excel/excel.pivottableupdatedata#name)|Имя сводной таблицы.|
|[Range](/javascript/api/excel/excel.range)|[Жетвисиблевиев ()](/javascript/api/excel/excel.range#getvisibleview--)|Представляет видимые строки текущего диапазона.|
|[RangeView](/javascript/api/excel/excel.rangeview)|[formulas](/javascript/api/excel/excel.rangeview#formulas)|Представляет формулу в формате A1.|
||[formulasLocal](/javascript/api/excel/excel.rangeview#formulaslocal)|Представляет формулу в нотации стиля A1 на языке пользователя и в соответствии с его языковым стандартом. Например, английская формула =SUM(A1, 1.5) превратится в "=СУММ(A1; 1,5)" на русском языке.|
||[formulasR1C1](/javascript/api/excel/excel.rangeview#formulasr1c1)|Представляет формулу в формате R1C1.|
||[getRange()](/javascript/api/excel/excel.rangeview#getrange--)|Получает родительский диапазон, сопоставленный с текущим объектом RangeView.|
||[numberFormat](/javascript/api/excel/excel.rangeview#numberformat)|Представляет код в числовом формате Excel для данной ячейки.|
||[cellAddresses](/javascript/api/excel/excel.rangeview#celladdresses)|Представляет адреса ячеек RangeView. Только для чтения.|
||[Число](/javascript/api/excel/excel.rangeview#columncount)|Возвращает количество видимых столбцов. Только для чтения.|
||[индекс](/javascript/api/excel/excel.rangeview#index)|Возвращает значение, представляющее индекс RangeView. Только для чтения.|
||[Стро](/javascript/api/excel/excel.rangeview#rowcount)|Возвращает количество видимых строк. Только для чтения.|
||[строки](/javascript/api/excel/excel.rangeview#rows)|Представляет коллекцию видимых ячеек в диапазоне, сопоставленных с указанным диапазоном. Только для чтения.|
||[text](/javascript/api/excel/excel.rangeview#text)|Текстовые значения указанного диапазона. Текстовое значение не зависит от ширины ячейки. Замена знака #, которая происходит в пользовательском интерфейсе Excel, не повлияет на текстовое значение, возвращаемое API. Только для чтения.|
||[valueTypes](/javascript/api/excel/excel.rangeview#valuetypes)|Представляет тип данных каждой ячейки. Только для чтения.|
||[Set (Properties: Excel. RangeView)](/javascript/api/excel/excel.rangeview#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Ранжевиевупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.rangeview#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[values](/javascript/api/excel/excel.rangeview#values)|Представляет необработанные значения указанного объекта rangeView. Могут возвращаться строковые и числовые данные, а также логические значения. Ячейки, содержащие ошибку, вернут строку ошибки.|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getItemAt(index: number)](/javascript/api/excel/excel.rangeviewcollection#getitemat-index-)|Получает строку RangeView с помощью индекса. Используется нулевой индекс.|
||[items](/javascript/api/excel/excel.rangeviewcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Ранжевиевколлектионлоадоптионс](/javascript/api/excel/excel.rangeviewcollectionloadoptions)|[$all](/javascript/api/excel/excel.rangeviewcollectionloadoptions#$all)||
||[cellAddresses](/javascript/api/excel/excel.rangeviewcollectionloadoptions#celladdresses)|Для каждого элемента в коллекции: представляет адреса ячеек RangeView. Только для чтения.|
||[Число](/javascript/api/excel/excel.rangeviewcollectionloadoptions#columncount)|Для каждого элемента в коллекции: Возвращает число видимых столбцов. Только для чтения.|
||[formulas](/javascript/api/excel/excel.rangeviewcollectionloadoptions#formulas)|Для каждого элемента в коллекции: представляет формулу в нотации стиля a1.|
||[formulasLocal](/javascript/api/excel/excel.rangeviewcollectionloadoptions#formulaslocal)|Для каждого элемента в коллекции: представляет формулу в нотации стиля a1 в языке пользователя и в языковом стандартном форматировании.  Например, английская формула =SUM(A1, 1.5) превратится в "=СУММ(A1; 1,5)" на русском языке.|
||[formulasR1C1](/javascript/api/excel/excel.rangeviewcollectionloadoptions#formulasr1c1)|Для каждого элемента в коллекции: представляет формулу в нотации стиля R1C1.|
||[индекс](/javascript/api/excel/excel.rangeviewcollectionloadoptions#index)|Для каждого элемента в коллекции: Возвращает значение, представляющее индекс объекта RangeView. Только для чтения.|
||[numberFormat](/javascript/api/excel/excel.rangeviewcollectionloadoptions#numberformat)|Для каждого элемента в коллекции: представляет код числового формата Excel для данной ячейки.|
||[Стро](/javascript/api/excel/excel.rangeviewcollectionloadoptions#rowcount)|Для каждого элемента в коллекции: Возвращает количество видимых строк. Только для чтения.|
||[text](/javascript/api/excel/excel.rangeviewcollectionloadoptions#text)|Для каждого элемента в коллекции: текстовые значения указанного диапазона. Текстовое значение не зависит от ширины ячейки. Замена знака #, которая происходит в пользовательском интерфейсе Excel, не повлияет на текстовое значение, возвращаемое API. Только для чтения.|
||[valueTypes](/javascript/api/excel/excel.rangeviewcollectionloadoptions#valuetypes)|Для каждого элемента в коллекции: представляет тип данных каждой ячейки. Только для чтения.|
||[values](/javascript/api/excel/excel.rangeviewcollectionloadoptions#values)|Для каждого элемента в коллекции: представляет необработанные значения указанного представления диапазона. Могут возвращаться строковые и числовые данные, а также логические значения. Ячейки, содержащие ошибку, вернут строку ошибки.|
|[Ранжевиевдата](/javascript/api/excel/excel.rangeviewdata)|[cellAddresses](/javascript/api/excel/excel.rangeviewdata#celladdresses)|Представляет адреса ячеек RangeView. Только для чтения.|
||[Число](/javascript/api/excel/excel.rangeviewdata#columncount)|Возвращает количество видимых столбцов. Только для чтения.|
||[formulas](/javascript/api/excel/excel.rangeviewdata#formulas)|Представляет формулу в формате A1.|
||[formulasLocal](/javascript/api/excel/excel.rangeviewdata#formulaslocal)|Представляет формулу в нотации стиля A1 на языке пользователя и в соответствии с его языковым стандартом. Например, английская формула =SUM(A1, 1.5) превратится в "=СУММ(A1; 1,5)" на русском языке.|
||[formulasR1C1](/javascript/api/excel/excel.rangeviewdata#formulasr1c1)|Представляет формулу в формате R1C1.|
||[индекс](/javascript/api/excel/excel.rangeviewdata#index)|Возвращает значение, представляющее индекс RangeView. Только для чтения.|
||[numberFormat](/javascript/api/excel/excel.rangeviewdata#numberformat)|Представляет код в числовом формате Excel для данной ячейки.|
||[rowCount](/javascript/api/excel/excel.rangeviewdata#rowcount)|Возвращает количество видимых строк. Только для чтения.|
||[строки](/javascript/api/excel/excel.rangeviewdata#rows)|Представляет коллекцию видимых ячеек в диапазоне, сопоставленных с указанным диапазоном. Только для чтения.|
||[text](/javascript/api/excel/excel.rangeviewdata#text)|Текстовые значения указанного диапазона. Текстовое значение не зависит от ширины ячейки. Замена знака #, которая происходит в пользовательском интерфейсе Excel, не повлияет на текстовое значение, возвращаемое API. Только для чтения.|
||[valueTypes](/javascript/api/excel/excel.rangeviewdata#valuetypes)|Представляет тип данных каждой ячейки. Только для чтения.|
||[values](/javascript/api/excel/excel.rangeviewdata#values)|Представляет необработанные значения указанного объекта rangeView. Могут возвращаться строковые и числовые данные, а также логические значения. Ячейки, содержащие ошибку, вернут строку ошибки.|
|[Ранжевиевлоадоптионс](/javascript/api/excel/excel.rangeviewloadoptions)|[$all](/javascript/api/excel/excel.rangeviewloadoptions#$all)||
||[cellAddresses](/javascript/api/excel/excel.rangeviewloadoptions#celladdresses)|Представляет адреса ячеек RangeView. Только для чтения.|
||[Число](/javascript/api/excel/excel.rangeviewloadoptions#columncount)|Возвращает количество видимых столбцов. Только для чтения.|
||[formulas](/javascript/api/excel/excel.rangeviewloadoptions#formulas)|Представляет формулу в формате A1.|
||[formulasLocal](/javascript/api/excel/excel.rangeviewloadoptions#formulaslocal)|Представляет формулу в нотации стиля A1 на языке пользователя и в соответствии с его языковым стандартом. Например, английская формула =SUM(A1, 1.5) превратится в "=СУММ(A1; 1,5)" на русском языке.|
||[formulasR1C1](/javascript/api/excel/excel.rangeviewloadoptions#formulasr1c1)|Представляет формулу в формате R1C1.|
||[индекс](/javascript/api/excel/excel.rangeviewloadoptions#index)|Возвращает значение, представляющее индекс RangeView. Только для чтения.|
||[numberFormat](/javascript/api/excel/excel.rangeviewloadoptions#numberformat)|Представляет код в числовом формате Excel для данной ячейки.|
||[rowCount](/javascript/api/excel/excel.rangeviewloadoptions#rowcount)|Возвращает количество видимых строк. Только для чтения.|
||[text](/javascript/api/excel/excel.rangeviewloadoptions#text)|Текстовые значения указанного диапазона. Текстовое значение не зависит от ширины ячейки. Замена знака #, которая происходит в пользовательском интерфейсе Excel, не повлияет на текстовое значение, возвращаемое API. Только для чтения.|
||[valueTypes](/javascript/api/excel/excel.rangeviewloadoptions#valuetypes)|Представляет тип данных каждой ячейки. Только для чтения.|
||[values](/javascript/api/excel/excel.rangeviewloadoptions#values)|Представляет необработанные значения указанного объекта rangeView. Могут возвращаться строковые и числовые данные, а также логические значения. Ячейки, содержащие ошибку, вернут строку ошибки.|
|[Ранжевиевупдатедата](/javascript/api/excel/excel.rangeviewupdatedata)|[formulas](/javascript/api/excel/excel.rangeviewupdatedata#formulas)|Представляет формулу в формате A1.|
||[formulasLocal](/javascript/api/excel/excel.rangeviewupdatedata#formulaslocal)|Представляет формулу в нотации стиля A1 на языке пользователя и в соответствии с его языковым стандартом. Например, английская формула =SUM(A1, 1.5) превратится в "=СУММ(A1; 1,5)" на русском языке.|
||[formulasR1C1](/javascript/api/excel/excel.rangeviewupdatedata#formulasr1c1)|Представляет формулу в формате R1C1.|
||[numberFormat](/javascript/api/excel/excel.rangeviewupdatedata#numberformat)|Представляет код в числовом формате Excel для данной ячейки.|
||[values](/javascript/api/excel/excel.rangeviewupdatedata#values)|Представляет необработанные значения указанного объекта rangeView. Могут возвращаться строковые и числовые данные, а также логические значения. Ячейки, содержащие ошибку, вернут строку ошибки.|
|[Table](/javascript/api/excel/excel.table)|[highlightFirstColumn](/javascript/api/excel/excel.table#highlightfirstcolumn)|Указывает, содержит ли первый столбец специальное форматирование.|
||[highlightLastColumn](/javascript/api/excel/excel.table#highlightlastcolumn)|Указывает, содержит ли последний столбец специальное форматирование.|
||[showBandedColumns](/javascript/api/excel/excel.table#showbandedcolumns)|Указывает, чередуется ли форматирование четных и нечетных столбцов для более удобного просмотра таблицы.|
||[showBandedRows](/javascript/api/excel/excel.table#showbandedrows)|Указывает, чередуется ли форматирование четных и нечетных строк для более удобного просмотра таблицы.|
||[showFilterButton](/javascript/api/excel/excel.table#showfilterbutton)|Указывает, видны ли кнопки фильтрации в верхней части заголовков столбцов. Это свойство можно использовать, только если таблица содержит строку заголовков.|
|[Таблеколлектионлоадоптионс](/javascript/api/excel/excel.tablecollectionloadoptions)|[highlightFirstColumn](/javascript/api/excel/excel.tablecollectionloadoptions#highlightfirstcolumn)|Для каждого элемента в коллекции: указывает, содержит ли первый столбец специальное форматирование.|
||[highlightLastColumn](/javascript/api/excel/excel.tablecollectionloadoptions#highlightlastcolumn)|Для каждого элемента в коллекции: указывает, содержит ли последний столбец специальное форматирование.|
||[showBandedColumns](/javascript/api/excel/excel.tablecollectionloadoptions#showbandedcolumns)|Для каждого элемента в коллекции: указывает, отображаются ли в столбцах полоснее форматирование, в результате которой нечетные столбцы выделяются не так, как даже для упрощения чтения таблицы.|
||[showBandedRows](/javascript/api/excel/excel.tablecollectionloadoptions#showbandedrows)|Для каждого элемента в коллекции: указывает, отображаются ли в строках форматирование с чередованием, в результате чего нечетные строки выделяются иначе, чтобы упростить чтение таблицы.|
||[showFilterButton](/javascript/api/excel/excel.tablecollectionloadoptions#showfilterbutton)|Для каждого элемента в коллекции: указывает, отображаются ли кнопки фильтра в верхней части каждого заголовка столбца. Это свойство можно использовать, только если таблица содержит строку заголовков.|
|[TableData](/javascript/api/excel/excel.tabledata)|[highlightFirstColumn](/javascript/api/excel/excel.tabledata#highlightfirstcolumn)|Указывает, содержит ли первый столбец специальное форматирование.|
||[highlightLastColumn](/javascript/api/excel/excel.tabledata#highlightlastcolumn)|Указывает, содержит ли последний столбец специальное форматирование.|
||[showBandedColumns](/javascript/api/excel/excel.tabledata#showbandedcolumns)|Указывает, чередуется ли форматирование четных и нечетных столбцов для более удобного просмотра таблицы.|
||[showBandedRows](/javascript/api/excel/excel.tabledata#showbandedrows)|Указывает, чередуется ли форматирование четных и нечетных строк для более удобного просмотра таблицы.|
||[showFilterButton](/javascript/api/excel/excel.tabledata#showfilterbutton)|Указывает, видны ли кнопки фильтрации в верхней части заголовков столбцов. Это свойство можно использовать, только если таблица содержит строку заголовков.|
|[Таблелоадоптионс](/javascript/api/excel/excel.tableloadoptions)|[highlightFirstColumn](/javascript/api/excel/excel.tableloadoptions#highlightfirstcolumn)|Указывает, содержит ли первый столбец специальное форматирование.|
||[highlightLastColumn](/javascript/api/excel/excel.tableloadoptions#highlightlastcolumn)|Указывает, содержит ли последний столбец специальное форматирование.|
||[showBandedColumns](/javascript/api/excel/excel.tableloadoptions#showbandedcolumns)|Указывает, чередуется ли форматирование четных и нечетных столбцов для более удобного просмотра таблицы.|
||[showBandedRows](/javascript/api/excel/excel.tableloadoptions#showbandedrows)|Указывает, чередуется ли форматирование четных и нечетных строк для более удобного просмотра таблицы.|
||[showFilterButton](/javascript/api/excel/excel.tableloadoptions#showfilterbutton)|Указывает, видны ли кнопки фильтрации в верхней части заголовков столбцов. Это свойство можно использовать, только если таблица содержит строку заголовков.|
|[Таблеупдатедата](/javascript/api/excel/excel.tableupdatedata)|[highlightFirstColumn](/javascript/api/excel/excel.tableupdatedata#highlightfirstcolumn)|Указывает, содержит ли первый столбец специальное форматирование.|
||[highlightLastColumn](/javascript/api/excel/excel.tableupdatedata#highlightlastcolumn)|Указывает, содержит ли последний столбец специальное форматирование.|
||[showBandedColumns](/javascript/api/excel/excel.tableupdatedata#showbandedcolumns)|Указывает, чередуется ли форматирование четных и нечетных столбцов для более удобного просмотра таблицы.|
||[showBandedRows](/javascript/api/excel/excel.tableupdatedata#showbandedrows)|Указывает, чередуется ли форматирование четных и нечетных строк для более удобного просмотра таблицы.|
||[showFilterButton](/javascript/api/excel/excel.tableupdatedata#showfilterbutton)|Указывает, видны ли кнопки фильтрации в верхней части заголовков столбцов. Это свойство можно использовать, только если таблица содержит строку заголовков.|
|[Workbook](/javascript/api/excel/excel.workbook)|[Сводные таблицы](/javascript/api/excel/excel.workbook#pivottables)|Представляет коллекцию сводных таблиц, сопоставленных с книгой. Только для чтения.|
|[Воркбукдата](/javascript/api/excel/excel.workbookdata)|[Сводные таблицы](/javascript/api/excel/excel.workbookdata#pivottables)|Представляет коллекцию сводных таблиц, сопоставленных с книгой. Только для чтения.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[Сводные таблицы](/javascript/api/excel/excel.worksheet#pivottables)|Коллекция сводных таблиц на листе. Только для чтения.|
|[Воркшитдата](/javascript/api/excel/excel.worksheetdata)|[Сводные таблицы](/javascript/api/excel/excel.worksheetdata#pivottables)|Коллекция сводных таблиц на листе. Только для чтения.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel)
- [Наборы обязательных элементов API JavaScript для Excel](./excel-api-requirement-sets.md)
