---
title: Excel Набор API JavaScript 1.13
description: Сведения о наборе требований ExcelApi 1.13.
ms.date: 07/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: bfd9c23beda64565b44f16845e046fa1a2358d41
ms.sourcegitcommit: aa73ec6367eaf74399fbf8d6b7776d77895e9982
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/03/2021
ms.locfileid: "53290815"
---
# <a name="whats-new-in-excel-javascript-api-113"></a>Новые возможности в Excel API JavaScript 1.13

В ExcelApi 1.13 добавлен метод вставки листа в книгу из строки с кодированной базой 64 и событие для обнаружения активации книги. Это также увеличило поддержку формул в диапазонах, добавив API для отслеживания изменений формул и поиска прямых зависимых ячеек формулы. Кроме того, она расширила поддержку PivotTable, добавив API PivotLayout для ALT-текста, стиля и управления пустыми ячейками.

| Функциональная область | Описание | Соответствующие объекты |
|:--- |:--- |:--- |
| События с измененной формулой | Отслеживание изменений формул, в том числе источника и типа события, которое вызвало изменение. | [Таблица.onFormulaChanged](/javascript/api/excel/excel.worksheet#onFormulaChanged)|
| Зависимые формулы | Найдите прямые зависимые ячейки формулы. | [Range.getDirectDependents](/javascript/api/excel/excel.range#getDirectDependents__) |
| Вставка таблиц | Вставьте таблицы из другой книги в текущую книгу в качестве строки с кодом Base64. | [Workbook.insertWorksheetsFromBase64](/javascript/api/excel/excel.workbook#insertWorksheetsFromBase64_base64File__options_) |
| PivotTable PivotLayout | Расширение класса PivotLayout, включая новую поддержку текста alt и управление пустыми ячейками. | [PivotLayout](/javascript/api/excel/excel.pivotlayout) |

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в Excel API JavaScript, за набором 1.13. Чтобы просмотреть справочную документацию API для всех API, поддерживаемых Excel API JavaScript, установленного 1.13 или ранее, см. в Excel API в наборе требований [1.13](/javascript/api/excel?view=excel-js-1.13&preserve-view=true)или ранее .

| Класс | Поля | Описание |
|:---|:---|:---|
|[FormulaChangedEventDetail](/javascript/api/excel/excel.formulachangedeventdetail)|[cellAddress](/javascript/api/excel/excel.formulachangedeventdetail#celladdress)|Адрес ячейки, содержаной измененную формулу.|
||[previousFormula](/javascript/api/excel/excel.formulachangedeventdetail#previousformula)|Представляет предыдущую формулу, прежде чем она была изменена.|
|[InsertWorksheetOptions](/javascript/api/excel/excel.insertworksheetoptions)|[positionType](/javascript/api/excel/excel.insertworksheetoptions#positiontype)|Положение вставки в текущей книге новых таблиц.|
||[relativeTo](/javascript/api/excel/excel.insertworksheetoptions#relativeto)|Таблица в текущей книге, которая ссылается на `WorksheetPositionType` параметр.|
||[sheetNamesToInsert](/javascript/api/excel/excel.insertworksheetoptions#sheetnamestoinsert)|Имена отдельных таблиц, которые необходимо вставить.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[altTextDescription](/javascript/api/excel/excel.pivotlayout#alttextdescription)|The alt text description of the PivotTable.|
||[altTextTitle](/javascript/api/excel/excel.pivotlayout#alttexttitle)|The alt text title of the PivotTable.|
||[displayBlankLineAfterEachItem(display: boolean)](/javascript/api/excel/excel.pivotlayout#displayblanklineaftereachitem-display-)|Задает, следует ли отображать пустую строку после каждого элемента.|
||[emptyCellText](/javascript/api/excel/excel.pivotlayout#emptycelltext)|Текст, который автоматически заполняется в любую пустую ячейку в PivotTable если `fillEmptyCells == true` .|
||[fillEmptyCells](/javascript/api/excel/excel.pivotlayout#fillemptycells)|Указывает, должны ли пустые ячейки в PivotTable заполняться с `emptyCellText` помощью .|
||[repeatAllItemLabels (repeatLabels: boolean)](/javascript/api/excel/excel.pivotlayout#repeatallitemlabels-repeatlabels-)|Задает параметр "Повторите все метки элементов" во всех полях в PivotTable.|
||[showFieldHeaders](/javascript/api/excel/excel.pivotlayout#showfieldheaders)|Указывает, отображаются ли в pivotTable полевые заголовок (подписи полей и отфильтровываемые выпадения).|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[refreshOnOpen](/javascript/api/excel/excel.pivottable#refreshonopen)|Указывает, обновляется ли pivotTable при открываемой книге.|
|[Range](/javascript/api/excel/excel.range)|[getDirectDependents()](/javascript/api/excel/excel.range#getdirectdependents--)|Возвращает объект, представляющего диапазон, содержащий все прямые иждивенцы ячейки в одной и той же таблице или в нескольких `WorkbookRangeAreas` таблицах.|
||[getExtendedRange (направление: Excel. KeyboardDirection, activeCell?: Range \| string)](/javascript/api/excel/excel.range#getextendedrange-direction--activecell-)|Возвращает объект диапазона, который включает текущий диапазон и до края диапазона, в зависимости от предоставленного направления.|
||[getMergedAreasOrNullObject()](/javascript/api/excel/excel.range#getmergedareasornullobject--)|Возвращает объект RangeAreas, который представляет объединенные области в этом диапазоне.|
||[getRangeEdge (направление: Excel. KeyboardDirection, activeCell?: Range \| string)](/javascript/api/excel/excel.range#getrangeedge-direction--activecell-)|Возвращает объект диапазона, который является краеугольным элементом области данных, соответствующей предоставленной направлению.|
|[Table](/javascript/api/excel/excel.table)|[resize (newRange: Range \| string)](/javascript/api/excel/excel.table#resize-newrange-)|Resize the table to the new range.|
|[Workbook](/javascript/api/excel/excel.workbook)|[insertWorksheetsFromBase64(base64File: string, options?: Excel. InsertWorksheetOptions)](/javascript/api/excel/excel.workbook#insertworksheetsfrombase64-base64file--options-)|Вставляет указанные таблицы из источника книги в текущую книгу.|
||[onActivated](/javascript/api/excel/excel.workbook#onactivated)|Возникает при активации книги.|
|[WorkbookActivatedEventArgs](/javascript/api/excel/excel.workbookactivatedeventargs)|[type](/javascript/api/excel/excel.workbookactivatedeventargs#type)|Получает тип события.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onFormulaChanged](/javascript/api/excel/excel.worksheet#onformulachanged)|Возникает, когда в этом таблице изменена одна или несколько формул.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onFormulaChanged](/javascript/api/excel/excel.worksheetcollection#onformulachanged)|Возникает, когда одна или несколько формул меняются в любом таблице этой коллекции.|
|[WorksheetFormulaChangedEventArgs](/javascript/api/excel/excel.worksheetformulachangedeventargs)|[formulaDetails](/javascript/api/excel/excel.worksheetformulachangedeventargs#formuladetails)|Получает массив объектов, содержащих сведения обо всех `FormulaChangedEventDetail` измененных формулах.|
||[source](/javascript/api/excel/excel.worksheetformulachangedeventargs#source)|Источник события.|
||[type](/javascript/api/excel/excel.worksheetformulachangedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetformulachangedeventargs#worksheetid)|Получает ID таблицы, в которой изменена формула.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.13&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)
