---
title: Excel Набор API JavaScript 1.3
description: Сведения о наборе требований ExcelApi 1.3.
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: d3606b74e8a1099cd58631cc047a783f27a09a19
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671284"
---
# <a name="whats-new-in-excel-javascript-api-13"></a>Новые возможности API JavaScript для Excel 1.3

В ExcelApi 1.3 добавлена поддержка привязки к данным и базового доступа к pivotTable.

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в Excel API JavaScript, установленный 1.3. Чтобы просмотреть справочную документацию API для всех API, поддерживаемых Excel API JavaScript, за набором 1.3 или более ранних, см. в Excel API в наборе требований [1.3](/javascript/api/excel?view=excel-js-1.3&preserve-view=true)или ранее .

| Класс | Поля | Описание |
|:---|:---|:---|
|[Binding](/javascript/api/excel/excel.binding)|[delete()](/javascript/api/excel/excel.binding#delete__)|Удаляет привязку.|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[add(range: Range \| string, bindingType: Excel. BindingType, id: string)](/javascript/api/excel/excel.bindingcollection#add_range__bindingType__id_)|Добавляет привязку к определенному объекту Range.|
||[addFromNamedItem (имя: строка, bindingType: Excel. BindingType, id: string)](/javascript/api/excel/excel.bindingcollection#addFromNamedItem_name__bindingType__id_)|Добавляет новую привязку с учетом именованного элемента в книге.|
||[addFromSelection(bindingType: Excel. BindingType, id: string)](/javascript/api/excel/excel.bindingcollection#addFromSelection_bindingType__id_)|Добавляет новую привязку с учетом выделенного в настоящий момент фрагмента.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[name](/javascript/api/excel/excel.pivottable#name)|Имя сводной таблицы.|
||[worksheet](/javascript/api/excel/excel.pivottable#worksheet)|Лист, содержащий текущую сводную таблицу.|
||[refresh()](/javascript/api/excel/excel.pivottable#refresh__)|Обновляет сводную таблицу.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getItem(name: string)](/javascript/api/excel/excel.pivottablecollection#getItem_name_)|Получает сводную таблицу по имени.|
||[items](/javascript/api/excel/excel.pivottablecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[refreshAll()](/javascript/api/excel/excel.pivottablecollection#refreshAll__)|Обновляет все сводные таблицы в коллекции.|
|[Range](/javascript/api/excel/excel.range)|[getVisibleView()](/javascript/api/excel/excel.range#getVisibleView__)|Представляет видимые строки текущего диапазона.|
|[RangeView](/javascript/api/excel/excel.rangeview)|[formulas](/javascript/api/excel/excel.rangeview#formulas)|Представляет формулу в формате A1.|
||[formulasLocal](/javascript/api/excel/excel.rangeview#formulasLocal)|Представляет формулу в нотации стиля A1 на языке пользователя и в соответствии с его языковым стандартом.|
||[formulasR1C1](/javascript/api/excel/excel.rangeview#formulasR1C1)|Представляет формулу в формате R1C1.|
||[getRange()](/javascript/api/excel/excel.rangeview#getRange__)|Получает родительский диапазон, связанный с текущим `RangeView` .|
||[numberFormat](/javascript/api/excel/excel.rangeview#numberFormat)|Представляет код в числовом формате Excel для данной ячейки.|
||[cellAddresses](/javascript/api/excel/excel.rangeview#cellAddresses)|Представляет адреса ячейки `RangeView` .|
||[columnCount](/javascript/api/excel/excel.rangeview#columnCount)|Количество видимых столбцов.|
||[index](/javascript/api/excel/excel.rangeview#index)|Возвращает значение, которое представляет индекс `RangeView` индекса .|
||[rowCount](/javascript/api/excel/excel.rangeview#rowCount)|Количество видимых строк.|
||[строки](/javascript/api/excel/excel.rangeview#rows)|Представляет коллекцию видимых ячеек в диапазоне, сопоставленных с указанным диапазоном.|
||[text](/javascript/api/excel/excel.rangeview#text)|Текстовые значения указанного диапазона.|
||[valueTypes](/javascript/api/excel/excel.rangeview#valueTypes)|Представляет тип данных каждой ячейки.|
||[values](/javascript/api/excel/excel.rangeview#values)|Представляет необработанные значения указанного объекта rangeView.|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getItemAt(index: number)](/javascript/api/excel/excel.rangeviewcollection#getItemAt_index_)|Получает `RangeView` строку с помощью индекса.|
||[items](/javascript/api/excel/excel.rangeviewcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Table](/javascript/api/excel/excel.table)|[highlightFirstColumn](/javascript/api/excel/excel.table#highlightFirstColumn)|Указывает, содержит ли первый столбец специальный форматирование.|
||[highlightLastColumn](/javascript/api/excel/excel.table#highlightLastColumn)|Указывает, содержит ли последний столбец специальный форматирование.|
||[showBandedColumns](/javascript/api/excel/excel.table#showBandedColumns)|Указывает, показывают ли столбцы полосатую форматирование, в котором нечетные столбцы выделены иначе, чем даже, чтобы облегчить чтение таблицы.|
||[showBandedRows](/javascript/api/excel/excel.table#showBandedRows)|Указывает, показывают ли строки полосы форматирования, в которых нечетные строки выделяются иначе, чем четные, чтобы облегчить чтение таблицы.|
||[showFilterButton](/javascript/api/excel/excel.table#showFilterButton)|Указывает, видны ли кнопки фильтра в верхней части каждого загона столбца.|
|[Workbook](/javascript/api/excel/excel.workbook)|[pivotTables](/javascript/api/excel/excel.workbook#pivotTables)|Представляет коллекцию сводных таблиц, сопоставленных с книгой.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[pivotTables](/javascript/api/excel/excel.worksheet#pivotTables)|Коллекция сводных таблиц на листе.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.3&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)
