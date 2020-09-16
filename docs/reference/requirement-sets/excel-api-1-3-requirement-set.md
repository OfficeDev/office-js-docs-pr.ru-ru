---
title: Набор обязательных элементов API JavaScript для Excel 1,3
description: Сведения о наборе требований ExcelApi 1,3.
ms.date: 07/26/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: ee72e8bde7b768b2edb3dcff5217325e2336e3ab
ms.sourcegitcommit: ed2a98b6fb5b432fa99c6cefa5ce52965dc25759
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/16/2020
ms.locfileid: "47819821"
---
# <a name="whats-new-in-excel-javascript-api-13"></a>Новые возможности API JavaScript для Excel 1.3

ExcelApi 1,3 добавлена поддержка привязки данных и базового доступа к сводным таблицам.

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в наборе обязательных элементов API JavaScript для Excel 1,3. Чтобы просмотреть справочную документацию по API для всех API, поддерживаемых набором обязательных элементов API JavaScript для Excel 1,3 или более ранней версии, обратитесь к разделам [API Excel в наборе требований 1,3](/javascript/api/excel?view=excel-js-1.3&preserve-view=true)

| Класс | Поля | Описание |
|:---|:---|:---|
|[Binding](/javascript/api/excel/excel.binding)|[delete()](/javascript/api/excel/excel.binding#delete--)|Удаляет привязку.|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[Add (Range: \| строка Range, bindingType: Excel. bindingType, ID: строка)](/javascript/api/excel/excel.bindingcollection#add-range--bindingtype--id-)|Добавляет привязку к определенному объекту Range.|
||[Аддфромнамедитем (имя: строка, bindingType: Excel. BindingType, ID: строка)](/javascript/api/excel/excel.bindingcollection#addfromnameditem-name--bindingtype--id-)|Добавляет новую привязку с учетом именованного элемента в книге.|
||[Аддфромселектион (bindingType: Excel. BindingType, ID: String)](/javascript/api/excel/excel.bindingcollection#addfromselection-bindingtype--id-)|Добавляет новую привязку с учетом выделенного в настоящий момент фрагмента.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[name](/javascript/api/excel/excel.pivottable#name)|Имя сводной таблицы.|
||[worksheet](/javascript/api/excel/excel.pivottable#worksheet)|Лист, содержащий текущую сводную таблицу.|
||[refresh()](/javascript/api/excel/excel.pivottable#refresh--)|Обновляет сводную таблицу.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getItem(name: string)](/javascript/api/excel/excel.pivottablecollection#getitem-name-)|Получает сводную таблицу по имени.|
||[items](/javascript/api/excel/excel.pivottablecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[refreshAll ()](/javascript/api/excel/excel.pivottablecollection#refreshall--)|Обновляет все сводные таблицы в коллекции.|
|[Range](/javascript/api/excel/excel.range)|[Жетвисиблевиев ()](/javascript/api/excel/excel.range#getvisibleview--)|Представляет видимые строки текущего диапазона.|
|[RangeView](/javascript/api/excel/excel.rangeview)|[formulas](/javascript/api/excel/excel.rangeview#formulas)|Представляет формулу в формате A1.|
||[formulasLocal](/javascript/api/excel/excel.rangeview#formulaslocal)|Представляет формулу в нотации стиля A1 на языке пользователя и в соответствии с его языковым стандартом. Например, английская формула =SUM(A1, 1.5) превратится в "=СУММ(A1; 1,5)" на русском языке.|
||[formulasR1C1](/javascript/api/excel/excel.rangeview#formulasr1c1)|Представляет формулу в формате R1C1.|
||[getRange()](/javascript/api/excel/excel.rangeview#getrange--)|Получает родительский диапазон, сопоставленный с текущим объектом RangeView.|
||[numberFormat](/javascript/api/excel/excel.rangeview#numberformat)|Представляет код в числовом формате Excel для данной ячейки.|
||[cellAddresses](/javascript/api/excel/excel.rangeview#celladdresses)|Представляет адреса ячеек RangeView. Только для чтения.|
||[Число](/javascript/api/excel/excel.rangeview#columncount)|Возвращает количество видимых столбцов. Только для чтения.|
||[index](/javascript/api/excel/excel.rangeview#index)|Возвращает значение, представляющее индекс RangeView. Только для чтения.|
||[Стро](/javascript/api/excel/excel.rangeview#rowcount)|Возвращает количество видимых строк. Только для чтения.|
||[строки](/javascript/api/excel/excel.rangeview#rows)|Представляет коллекцию видимых ячеек в диапазоне, сопоставленных с указанным диапазоном. Только для чтения.|
||[text](/javascript/api/excel/excel.rangeview#text)|Текстовые значения указанного диапазона. Текстовое значение не зависит от ширины ячейки. Замена знака #, которая происходит в пользовательском интерфейсе Excel, не повлияет на текстовое значение, возвращаемое API. Только для чтения.|
||[valueTypes](/javascript/api/excel/excel.rangeview#valuetypes)|Представляет тип данных каждой ячейки. Только для чтения.|
||[values](/javascript/api/excel/excel.rangeview#values)|Представляет необработанные значения указанного объекта rangeView. Могут возвращаться строковые и числовые данные, а также логические значения. Ячейки, содержащие ошибку, вернут строку ошибки.|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getItemAt(index: number)](/javascript/api/excel/excel.rangeviewcollection#getitemat-index-)|Получает строку RangeView с помощью индекса. Используется нулевой индекс.|
||[items](/javascript/api/excel/excel.rangeviewcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Table](/javascript/api/excel/excel.table)|[highlightFirstColumn](/javascript/api/excel/excel.table#highlightfirstcolumn)|Указывает, содержит ли первый столбец специальное форматирование.|
||[highlightLastColumn](/javascript/api/excel/excel.table#highlightlastcolumn)|Указывает, содержит ли последний столбец специальное форматирование.|
||[showBandedColumns](/javascript/api/excel/excel.table#showbandedcolumns)|Указывает, чередуется ли форматирование четных и нечетных столбцов для более удобного просмотра таблицы.|
||[showBandedRows](/javascript/api/excel/excel.table#showbandedrows)|Указывает, чередуется ли форматирование четных и нечетных строк для более удобного просмотра таблицы.|
||[showFilterButton](/javascript/api/excel/excel.table#showfilterbutton)|Указывает, видны ли кнопки фильтрации в верхней части заголовков столбцов. Это свойство можно использовать, только если таблица содержит строку заголовков.|
|[Workbook](/javascript/api/excel/excel.workbook)|[Сводные таблицы](/javascript/api/excel/excel.workbook#pivottables)|Представляет коллекцию сводных таблиц, сопоставленных с книгой. Только для чтения.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[Сводные таблицы](/javascript/api/excel/excel.worksheet#pivottables)|Коллекция сводных таблиц на листе. Только для чтения.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.3&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)
