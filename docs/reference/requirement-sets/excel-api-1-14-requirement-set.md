---
title: Excel API JavaScript установлено 1.14
description: Сведения о наборе требований ExcelApi 1.14.
ms.date: 12/08/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 93b1690a3c03e51dadb2110ec6382ca6ee86cfe1
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/23/2022
ms.locfileid: "63747018"
---
# <a name="whats-new-in-excel-javascript-api-114"></a>Новые возможности в Excel API JavaScript 1.14

В ExcelApi 1.14 добавлены объекты для управления функцией таблицы данных диаграммы, методом обнаружения всех ячеек-прецедентов формулы и событиями защиты листа для отслеживания изменений состояния защиты листа. Он также добавил несколько методов [`getItemOrNullObject`](../../develop/application-specific-api-model.md#ornullobject-methods-and-properties) для таких объектов`CommentCollection`, как , и `ShapeCollection``StyleCollection` для улучшения обработки ошибок.

| Функциональная область | Описание | Соответствующие объекты |
|:--- |:--- |:--- |
| [Таблицы данных диаграммы](../../excel/excel-add-ins-charts.md#add-and-format-a-chart-data-table) | Управление внешним видом, форматированием и видимостью таблиц данных на диаграммах. | [Chart](/javascript/api/excel/excel.chart), [ChartDataTable](/javascript/api/excel/excel.chartdatatable), [ChartDataTableFormat](/javascript/api/excel/excel.chartdatatableformat) |
| [Прецеденты формул](../../excel/excel-add-ins-ranges-precedents-dependents.md#get-the-precedents-of-a-formula) | Возвращаем все ячейки-прецеденты формулы. | [Range](/javascript/api/excel/excel.range) |
| Запросы | Извлечение атрибутов Power Query, таких как имя, дата обновления и количество запросов. | [Запрос](/javascript/api/excel/excel.query), [QueryCollection](/javascript/api/excel/excel.querycollection)|
| [События защиты таблиц](../../excel/excel-add-ins-worksheets.md#detect-changes-to-the-worksheet-protection-state) | Отслеживание изменений состояния защиты таблицы и источника этих изменений. | [WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs), [WorksheetCollection](/javascript/api/excel/excel.worksheet)[](/javascript/api/excel/excel.worksheetcollection) |

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в Excel API JavaScript, за набором 1.14. Чтобы просмотреть справочную документацию API для всех API, поддерживаемых Excel API JavaScript, установленного 1.14 или ранее, см. в Excel API в наборе требований [1.14 или ранее](/javascript/api/excel?view=excel-js-1.14&preserve-view=true).

| Класс | Поля | Описание |
|:---|:---|:---|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[clearColumnCriteria(columnIndex: number)](/javascript/api/excel/excel.autofilter#excel-excel-autofilter-clearcolumncriteria-member(1))|Очищает критерии фильтрации столбцов автофайлов.|
|[ChangeDirectionState](/javascript/api/excel/excel.changedirectionstate)|[deleteShiftDirection](/javascript/api/excel/excel.changedirectionstate#excel-excel-changedirectionstate-deleteshiftdirection-member)|Представляет направление (например, вверх или влево), которое остальные ячейки будут смещаться при удалении ячейки или ячейки.|
||[insertShiftDirection](/javascript/api/excel/excel.changedirectionstate#excel-excel-changedirectionstate-insertshiftdirection-member)|Представляет направление (например, вниз или вправо), в которое будут перенесены существующие ячейки при вставке новой ячейки или ячеек.|
|[Chart](/javascript/api/excel/excel.chart)|[getDataTable()](/javascript/api/excel/excel.chart#excel-excel-chart-getdatatable-member(1))|Получает таблицу данных на диаграмме.|
||[getDataTableOrNullObject()](/javascript/api/excel/excel.chart#excel-excel-chart-getdatatableornullobject-member(1))|Получает таблицу данных на диаграмме.|
|[ChartDataTable](/javascript/api/excel/excel.chartdatatable)|[format](/javascript/api/excel/excel.chartdatatable#excel-excel-chartdatatable-format-member)|Представляет формат таблицы данных диаграммы, которая включает заполняемую таблицу, шрифт и пограничный формат.|
||[showHorizontalBorder](/javascript/api/excel/excel.chartdatatable#excel-excel-chartdatatable-showhorizontalborder-member)|Указывает, следует ли отображать горизонтальную границу таблицы данных.|
||[showLegendKey](/javascript/api/excel/excel.chartdatatable#excel-excel-chartdatatable-showlegendkey-member)|Указывает, следует ли показывать ключ-легенду таблицы данных.|
||[showOutlineBorder](/javascript/api/excel/excel.chartdatatable#excel-excel-chartdatatable-showoutlineborder-member)|Указывает, следует ли отображать границу контура таблицы данных.|
||[showVerticalBorder](/javascript/api/excel/excel.chartdatatable#excel-excel-chartdatatable-showverticalborder-member)|Указывает, следует ли отображать вертикальную границу таблицы данных.|
||[visible](/javascript/api/excel/excel.chartdatatable#excel-excel-chartdatatable-visible-member)|Указывает, следует ли показывать таблицу данных диаграммы.|
|[ChartDataTableFormat](/javascript/api/excel/excel.chartdatatableformat)|[граница](/javascript/api/excel/excel.chartdatatableformat#excel-excel-chartdatatableformat-border-member)|Представляет пограничный формат таблицы данных диаграммы, которая включает цвет, стиль строки и вес.|
||[fill](/javascript/api/excel/excel.chartdatatableformat#excel-excel-chartdatatableformat-fill-member)|Представляет формат заливки объекта, включая сведения о форматировании фона.|
||[font](/javascript/api/excel/excel.chartdatatableformat#excel-excel-chartdatatableformat-font-member)|Представляет атрибуты шрифта (например, имя шрифта, размер шрифта и цвет) для текущего объекта.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[getItemOrNullObject(commentId: string)](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-getitemornullobject-member(1))|Получает примечание из коллекции на основе его идентификатора.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[getItemOrNullObject(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-getitemornullobject-member(1))|Возвращает ответ на примечание, определенное по идентификатору.|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[getItemOrNullObject(id: строка)](/javascript/api/excel/excel.conditionalformatcollection#excel-excel-conditionalformatcollection-getitemornullobject-member(1))|Возвращает условный формат, идентифицированный его ID.|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.groupshapecollection#excel-excel-groupshapecollection-getitemornullobject-member(1))|Получает фигуру с ее именем или ИД.|
|[Query](/javascript/api/excel/excel.query)|[error](/javascript/api/excel/excel.query#excel-excel-query-error-member)|Получает сообщение об ошибке запроса с последнего обновления запроса.|
||[loadedTo](/javascript/api/excel/excel.query#excel-excel-query-loadedto-member)|Загружает запрос на тип объекта.|
||[loadedToDataModel](/javascript/api/excel/excel.query#excel-excel-query-loadedtodatamodel-member)|Указывает, загружен ли запрос в модель данных.|
||[name](/javascript/api/excel/excel.query#excel-excel-query-name-member)|Получает имя запроса.|
||[refreshDate](/javascript/api/excel/excel.query#excel-excel-query-refreshdate-member)|Получает дату и время последнего обновления запроса.|
||[rowsLoadedCount](/javascript/api/excel/excel.query#excel-excel-query-rowsloadedcount-member)|Получает количество строк, загруженных при последнем обновлении запроса.|
|[QueryCollection](/javascript/api/excel/excel.querycollection)|[getCount()](/javascript/api/excel/excel.querycollection#excel-excel-querycollection-getcount-member(1))|Получает количество запросов в книге.|
||[getItem(key: string)](/javascript/api/excel/excel.querycollection#excel-excel-querycollection-getitem-member(1))|Получает запрос из коллекции на основе его имени.|
||[items](/javascript/api/excel/excel.querycollection#excel-excel-querycollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[Range](/javascript/api/excel/excel.range)|[getPrecedents()](/javascript/api/excel/excel.range#excel-excel-range-getprecedents-member(1))|Возвращает объект, `WorkbookRangeAreas` представляющего диапазон, содержащий все прецеденты ячейки в одной и той же таблице или в нескольких таблицах.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.shapecollection#excel-excel-shapecollection-getitemornullobject-member(1))|Получает фигуру с ее именем или ИД.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.stylecollection#excel-excel-stylecollection-getitemornullobject-member(1))|Получает стиль по имени.|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablescopedcollection#excel-excel-tablescopedcollection-getitemornullobject-member(1))|Получает таблицу по имени или ИД.|
|[Workbook](/javascript/api/excel/excel.workbook)|[запросы](/javascript/api/excel/excel.workbook#excel-excel-workbook-queries-member)|Возвращает коллекцию запросов Power Query, которые являются частью книги.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onProtectionChanged](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onprotectionchanged-member)|Возникает при смене состояния защиты таблицы.|
||[tabId](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-tabid-member)|Возвращает значение, представляющее этот таблицу, которую можно прочитать в Open Office XML.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[changeDirectionState](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-changedirectionstate-member)|Представляет изменение в направлении, в которое будут сдвигаться ячейки в таблице при удалении или вставке ячейки.|
||[triggerSource](/javascript/api/excel/excel.worksheetchangedeventargs#excel-excel-worksheetchangedeventargs-triggersource-member)|Представляет источник триггера события.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onProtectionChanged](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onprotectionchanged-member)|Возникает при смене состояния защиты таблицы.|
|[WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs)|[isProtected](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#excel-excel-worksheetprotectionchangedeventargs-isprotected-member)|Получает текущее состояние защиты таблицы.|
||[источник](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#excel-excel-worksheetprotectionchangedeventargs-source-member)|Источник события.|
||[type](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#excel-excel-worksheetprotectionchangedeventargs-type-member)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#excel-excel-worksheetprotectionchangedeventargs-worksheetid-member)|Получает ID таблицы, в которой изменен статус защиты.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.14&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)
