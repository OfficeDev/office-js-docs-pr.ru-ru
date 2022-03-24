---
title: Excel API JavaScript установлено 1.3
description: Сведения о наборе требований ExcelApi 1.3.
ms.date: 11/09/2020
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 1bf8bc604c2c770f517878193994c1ed32640da1
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745341"
---
# <a name="whats-new-in-excel-javascript-api-13"></a>Новые возможности API JavaScript для Excel 1.3

В ExcelApi 1.3 добавлена поддержка привязки к данным и базового доступа к pivotTable.

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в Excel API JavaScript, за набором 1.3. Чтобы просмотреть справочную документацию API для всех API, поддерживаемых Excel API JavaScript, установленного 1.3 или ранее, см. в Excel API в наборе требований [1.3 или ранее](/javascript/api/excel?view=excel-js-1.3&preserve-view=true).

| Класс | Поля | Описание |
|:---|:---|:---|
|[Binding](/javascript/api/excel/excel.binding)|[delete()](/javascript/api/excel/excel.binding#excel-excel-binding-delete-member(1))|Удаляет привязку.|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[add(range: Range \| string, bindingType: Excel. BindingType, id: string)](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-add-member(1))|Добавляет привязку к определенному объекту Range.|
||[addFromNamedItem (имя: строка, bindingType: Excel. BindingType, id: string)](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-addfromnameditem-member(1))|Добавляет новую привязку с учетом именованного элемента в книге.|
||[addFromSelection (bindingType: Excel. BindingType, id: string)](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-addfromselection-member(1))|Добавляет новую привязку с учетом выделенного в настоящий момент фрагмента.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[name](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-name-member)|Имя сводной таблицы.|
||[refresh()](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-refresh-member(1))|Обновляет сводную таблицу.|
||[worksheet](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-worksheet-member)|Лист, содержащий текущую сводную таблицу.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getItem(name: string)](/javascript/api/excel/excel.pivottablecollection#excel-excel-pivottablecollection-getitem-member(1))|Получает сводную таблицу по имени.|
||[items](/javascript/api/excel/excel.pivottablecollection#excel-excel-pivottablecollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
||[refreshAll()](/javascript/api/excel/excel.pivottablecollection#excel-excel-pivottablecollection-refreshall-member(1))|Обновляет все сводные таблицы в коллекции.|
|[Range](/javascript/api/excel/excel.range)|[getVisibleView()](/javascript/api/excel/excel.range#excel-excel-range-getvisibleview-member(1))|Представляет видимые строки текущего диапазона.|
|[RangeView](/javascript/api/excel/excel.rangeview)|[cellAddresses](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-celladdresses-member)|Представляет адреса ячейки `RangeView`.|
||[columnCount](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-columncount-member)|Количество видимых столбцов.|
||[formulas](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-formulas-member)|Представляет формулу в формате A1.|
||[formulasLocal](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-formulaslocal-member)|Представляет формулу в нотации стиля A1 на языке пользователя и в соответствии с его языковым стандартом.|
||[formulasR1C1](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-formulasr1c1-member)|Представляет формулу в формате R1C1.|
||[getRange()](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-getrange-member(1))|Получает родительский диапазон, связанный с текущим `RangeView`.|
||[индекс](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-index-member)|Возвращает значение, которое представляет индекс индекса `RangeView`.|
||[numberFormat](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-numberformat-member)|Представляет код в числовом формате Excel для данной ячейки.|
||[rowCount](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-rowcount-member)|Количество видимых строк.|
||[строки](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-rows-member)|Представляет коллекцию видимых ячеек в диапазоне, сопоставленных с указанным диапазоном.|
||[text](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-text-member)|Текстовые значения указанного диапазона.|
||[valueTypes](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-valuetypes-member)|Представляет тип данных каждой ячейки.|
||[values](/javascript/api/excel/excel.rangeview#excel-excel-rangeview-values-member)|Представляет необработанные значения указанного объекта rangeView.|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getItemAt(index: number)](/javascript/api/excel/excel.rangeviewcollection#excel-excel-rangeviewcollection-getitemat-member(1))|Получает строку `RangeView` с помощью индекса.|
||[items](/javascript/api/excel/excel.rangeviewcollection#excel-excel-rangeviewcollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[Table](/javascript/api/excel/excel.table)|[highlightFirstColumn](/javascript/api/excel/excel.table#excel-excel-table-highlightfirstcolumn-member)|Указывает, содержит ли первый столбец специальный форматирование.|
||[highlightLastColumn](/javascript/api/excel/excel.table#excel-excel-table-highlightlastcolumn-member)|Указывает, содержит ли последний столбец специальный форматирование.|
||[showBandedColumns](/javascript/api/excel/excel.table#excel-excel-table-showbandedcolumns-member)|Указывает, показывают ли столбцы полосатую форматирование, в котором нечетные столбцы выделены иначе, чем даже, чтобы облегчить чтение таблицы.|
||[showBandedRows](/javascript/api/excel/excel.table#excel-excel-table-showbandedrows-member)|Указывает, показывают ли строки полосы форматирования, в которых нечетные строки выделяются иначе, чем четные, чтобы облегчить чтение таблицы.|
||[showFilterButton](/javascript/api/excel/excel.table#excel-excel-table-showfilterbutton-member)|Указывает, видны ли кнопки фильтра в верхней части каждого загона столбца.|
|[Workbook](/javascript/api/excel/excel.workbook)|[pivotTables](/javascript/api/excel/excel.workbook#excel-excel-workbook-pivottables-member)|Представляет коллекцию сводных таблиц, сопоставленных с книгой.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[pivotTables](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-pivottables-member)|Коллекция сводных таблиц на листе.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.3&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)
