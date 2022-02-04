---
title: Excel API JavaScript установлено 1.13
description: Сведения о наборе требований ExcelApi 1.13.
ms.date: 07/09/2021
ms.prod: excel
ms.localizationpriority: medium
---

# <a name="whats-new-in-excel-javascript-api-113"></a>Новые возможности в Excel API JavaScript 1.13

В ExcelApi 1.13 добавлен метод вставки листа в книгу из строки с кодированной базой 64 и событие для обнаружения активации книги. Это также увеличило поддержку формул в диапазонах, добавив API для отслеживания изменений формул и поиска прямых зависимых ячеек формулы. Кроме того, она расширила поддержку PivotTable, добавив API PivotLayout для ALT-текста, стиля и управления пустыми ячейками.

| Функциональная область | Описание | Соответствующие объекты |
|:--- |:--- |:--- |
| [События с измененной формулой](../../excel/excel-add-ins-worksheets.md#detect-formula-changes) | Отслеживание изменений формул, в том числе источника и типа события, которое вызвало изменение. | [Таблица.onFormulaChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onformulachanged-member)|
| [Зависимые формулы](../../excel/excel-add-ins-ranges-precedents-dependents.md#get-the-direct-dependents-of-a-formula) | Найдите прямые зависимые ячейки формулы. | [Range.getDirectDependents](/javascript/api/excel/excel.range#excel-excel-range-getdirectdependents-member(1)) |
| [Вставка таблиц](../../excel/excel-add-ins-workbooks.md#insert-a-copy-of-an-existing-workbook-into-the-current-one) | Вставьте таблицы из другой книги в текущую книгу в качестве строки с кодом Base64. | [Workbook.insertWorksheetsFromBase64](/javascript/api/excel/excel.workbook#excel-excel-workbook-insertworksheetsfrombase64-member(1)) |
| [PivotTable PivotLayout](../../excel/excel-add-ins-pivottables.md#other-pivotlayout-functions) | Расширение класса PivotLayout, включая новую поддержку текста alt и управление пустыми ячейками. | [PivotLayout](/javascript/api/excel/excel.pivotlayout) |

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в Excel API JavaScript, за набором 1.13. Чтобы просмотреть справочную документацию API для всех API, поддерживаемых Excel API JavaScript, за набором 1.13 или ранее, см. в Excel API в наборе требований [1.13 или ранее](/javascript/api/excel?view=excel-js-1.13&preserve-view=true).

| Класс | Поля | Описание |
|:---|:---|:---|
|[FormulaChangedEventDetail](/javascript/api/excel/excel.formulachangedeventdetail)|[cellAddress](/javascript/api/excel/excel.formulachangedeventdetail#excel-excel-formulachangedeventdetail-celladdress-member)|Адрес ячейки, содержаной измененную формулу.|
||[previousFormula](/javascript/api/excel/excel.formulachangedeventdetail#excel-excel-formulachangedeventdetail-previousformula-member)|Представляет предыдущую формулу, прежде чем она была изменена.|
|[InsertWorksheetOptions](/javascript/api/excel/excel.insertworksheetoptions)|[positionType](/javascript/api/excel/excel.insertworksheetoptions#excel-excel-insertworksheetoptions-positiontype-member)|Положение вставки в текущей книге новых таблиц.|
||[relativeTo](/javascript/api/excel/excel.insertworksheetoptions#excel-excel-insertworksheetoptions-relativeto-member)|Таблица в текущей книге, которая ссылается на параметр `WorksheetPositionType` .|
||[sheetNamesToInsert](/javascript/api/excel/excel.insertworksheetoptions#excel-excel-insertworksheetoptions-sheetnamestoinsert-member)|Имена отдельных таблиц, которые необходимо вставить.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[altTextDescription](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-alttextdescription-member)|The alt text description of the PivotTable.|
||[altTextTitle](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-alttexttitle-member)|The alt text title of the PivotTable.|
||[displayBlankLineAfterEachItem(display: boolean)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-displayblanklineaftereachitem-member(1))|Задает, следует ли отображать пустую строку после каждого элемента.|
||[emptyCellText](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-emptycelltext-member)|Текст, который автоматически заполняется в любую пустую ячейку в PivotTable если `fillEmptyCells == true`.|
||[fillEmptyCells](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-fillemptycells-member)|Указывает, должны ли пустые ячейки в PivotTable заполняться с помощью `emptyCellText`.|
||[repeatAllItemLabels (repeatLabels: boolean)](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-repeatallitemlabels-member(1))|Задает параметр "Повторите все метки элементов" во всех полях в PivotTable.|
||[showFieldHeaders](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-showfieldheaders-member)|Указывает, отображаются ли в pivotTable полевые заголовок (подписи полей и отфильтровываемые выпадения).|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[refreshOnOpen](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-refreshonopen-member)|Указывает, обновляется ли pivotTable при открываемой книге.|
|[Range](/javascript/api/excel/excel.range)|[getDirectDependents()](/javascript/api/excel/excel.range#excel-excel-range-getdirectdependents-member(1))|Возвращает объект, `WorkbookRangeAreas` представляющего диапазон, содержащий все прямые иждивенцы ячейки в одной и той же таблице или в нескольких таблицах.|
||[getExtendedRange (направление: Excel. KeyboardDirection, activeCell?: Range \| string)](/javascript/api/excel/excel.range#excel-excel-range-getextendedrange-member(1))|Возвращает объект диапазона, который включает текущий диапазон и до края диапазона, в зависимости от предоставленного направления.|
||[getMergedAreasOrNullObject()](/javascript/api/excel/excel.range#excel-excel-range-getmergedareasornullobject-member(1))|Возвращает объект RangeAreas, который представляет объединенные области в этом диапазоне.|
||[getRangeEdge (направление: Excel. KeyboardDirection, activeCell?: Range \| string)](/javascript/api/excel/excel.range#excel-excel-range-getrangeedge-member(1))|Возвращает объект диапазона, который является краеугольным элементом области данных, соответствующей предоставленной направлению.|
|[Table](/javascript/api/excel/excel.table)|[resize (newRange: Range \| string)](/javascript/api/excel/excel.table#excel-excel-table-resize-member(1))|Resize the table to the new range.|
|[Workbook](/javascript/api/excel/excel.workbook)|[insertWorksheetsFromBase64(base64File: string, options?: Excel. InsertWorksheetOptions)](/javascript/api/excel/excel.workbook#excel-excel-workbook-insertworksheetsfrombase64-member(1))|Вставляет указанные таблицы из источника книги в текущую книгу.|
||[onActivated](/javascript/api/excel/excel.workbook#excel-excel-workbook-onactivated-member)|Возникает при активации книги.|
|[WorkbookActivatedEventArgs](/javascript/api/excel/excel.workbookactivatedeventargs)|[type](/javascript/api/excel/excel.workbookactivatedeventargs#excel-excel-workbookactivatedeventargs-type-member)|Получает тип события.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onFormulaChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onformulachanged-member)|Возникает, когда в этом таблице изменена одна или несколько формул.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onFormulaChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onformulachanged-member)|Возникает, когда одна или несколько формул меняются в любом таблице этой коллекции.|
|[WorksheetFormulaChangedEventArgs](/javascript/api/excel/excel.worksheetformulachangedeventargs)|[formulaDetails](/javascript/api/excel/excel.worksheetformulachangedeventargs#excel-excel-worksheetformulachangedeventargs-formuladetails-member)|Получает массив объектов `FormulaChangedEventDetail` , содержащих сведения обо всех измененных формулах.|
||[источник](/javascript/api/excel/excel.worksheetformulachangedeventargs#excel-excel-worksheetformulachangedeventargs-source-member)|Источник события.|
||[type](/javascript/api/excel/excel.worksheetformulachangedeventargs#excel-excel-worksheetformulachangedeventargs-type-member)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetformulachangedeventargs#excel-excel-worksheetformulachangedeventargs-worksheetid-member)|Получает ID таблицы, в которой изменена формула.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.13&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)
