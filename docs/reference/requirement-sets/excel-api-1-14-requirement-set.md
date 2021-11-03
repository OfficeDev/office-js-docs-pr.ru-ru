---
title: Excel Набор API JavaScript 1.14
description: Сведения о наборе требований ExcelApi 1.14.
ms.date: 10/29/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 9cdf22d35125607237b724c88da2083ae78a9940
ms.sourcegitcommit: 23ce57b2702aca19054e31fcb2d2f015b4183ba1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/02/2021
ms.locfileid: "60681509"
---
# <a name="whats-new-in-excel-javascript-api-114"></a>Новые возможности в Excel API JavaScript 1.14

В ExcelApi 1.14 добавлены объекты для управления функцией таблицы данных диаграммы, методом обнаружения всех ячеек-прецедентов формулы и событиями защиты листа для отслеживания изменений состояния защиты листа. Он также добавил несколько [`getItemOrNullObject`](../../develop/application-specific-api-model.md#ornullobject-methods-and-properties) методов для таких объектов, как , и для `CommentCollection` улучшения `ShapeCollection` `StyleCollection` обработки ошибок.

| Функциональная область | Описание | Соответствующие объекты |
|:--- |:--- |:--- |
| Таблицы данных диаграммы | Управление внешним видом, форматированием и видимостью таблиц данных на диаграммах. | [Диаграмма](/javascript/api/excel/excel.chart), [ChartDataTable](/javascript/api/excel/excel.chartdatatable), [ChartDataTableFormat](/javascript/api/excel/excel.chartdatatableformat) |
| Прецеденты формул | Возвращаем все ячейки-прецеденты формулы. | [Range](/javascript/api/excel/excel.range) |
| Запросы | Извлечение атрибутов Power Query, таких как имя, дата обновления и количество запросов. | [Запрос](/javascript/api/excel/excel.query), [QueryCollection](/javascript/api/excel/excel.querycollection)|
| События защиты таблиц | Отслеживание изменений состояния защиты таблицы и источника этих изменений. | [WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs), [Таблица](/javascript/api/excel/excel.worksheet), [WorksheetCollection](/javascript/api/excel/excel.worksheetcollection) |

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в Excel API JavaScript, за набором 1.14. Чтобы просмотреть справочную документацию API для всех API, поддерживаемых Excel API JavaScript, установленного 1.14 или ранее, см. в Excel API в наборе требований [1.14](/javascript/api/excel?view=excel-js-1.14&preserve-view=true)или ранее .

| Класс | Поля | Описание |
|:---|:---|:---|
|[AutoFilter](/javascript/api/excel/excel.autofilter)|[clearColumnCriteria(columnIndex: number)](/javascript/api/excel/excel.autofilter#clearColumnCriteria_columnIndex_)|Очищает критерии фильтрации столбцов автофайлов.|
|[ChangeDirectionState](/javascript/api/excel/excel.changedirectionstate)|[deleteShiftDirection](/javascript/api/excel/excel.changedirectionstate#deleteShiftDirection)|Представляет направление (например, вверх или влево), которое остальные ячейки будут смещаться при удалении ячейки или ячейки.|
||[insertShiftDirection](/javascript/api/excel/excel.changedirectionstate#insertShiftDirection)|Представляет направление (например, вниз или вправо), в которое будут перенесены существующие ячейки при вставке новой ячейки или ячеек.|
|[Chart](/javascript/api/excel/excel.chart)|[getDataTable()](/javascript/api/excel/excel.chart#getDataTable__)|Получает таблицу данных на диаграмме.|
||[getDataTableOrNullObject()](/javascript/api/excel/excel.chart#getDataTableOrNullObject__)|Получает таблицу данных на диаграмме.|
|[ChartDataTable](/javascript/api/excel/excel.chartdatatable)|[format](/javascript/api/excel/excel.chartdatatable#format)|Представляет формат таблицы данных диаграммы, которая включает заполняемую таблицу, шрифт и пограничный формат.|
||[showHorizontalBorder](/javascript/api/excel/excel.chartdatatable#showHorizontalBorder)|Указывает, следует ли отображать горизонтальную границу таблицы данных.|
||[showLegendKey](/javascript/api/excel/excel.chartdatatable#showLegendKey)|Указывает, следует ли показывать ключ-легенду таблицы данных.|
||[showOutlineBorder](/javascript/api/excel/excel.chartdatatable#showOutlineBorder)|Указывает, следует ли отображать границу контура таблицы данных.|
||[showVerticalBorder](/javascript/api/excel/excel.chartdatatable#showVerticalBorder)|Указывает, следует ли отображать вертикальную границу таблицы данных.|
||[visible](/javascript/api/excel/excel.chartdatatable#visible)|Указывает, следует ли показывать таблицу данных диаграммы.|
|[ChartDataTableFormat](/javascript/api/excel/excel.chartdatatableformat)|[граница](/javascript/api/excel/excel.chartdatatableformat#border)|Представляет пограничный формат таблицы данных диаграммы, которая включает цвет, стиль строки и вес.|
||[fill](/javascript/api/excel/excel.chartdatatableformat#fill)|Представляет формат заливки объекта, включая сведения о форматировании фона.|
||[font](/javascript/api/excel/excel.chartdatatableformat#font)|Представляет атрибуты шрифта (например, имя шрифта, размер шрифта и цвет) для текущего объекта.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[getItemOrNullObject(commentId: string)](/javascript/api/excel/excel.commentcollection#getItemOrNullObject_commentId_)|Получает примечание из коллекции на основе его идентификатора.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[getItemOrNullObject(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getItemOrNullObject_commentReplyId_)|Возвращает ответ на примечание, определенное по идентификатору.|
|[ConditionalFormatCollection](/javascript/api/excel/excel.conditionalformatcollection)|[getItemOrNullObject(id: строка)](/javascript/api/excel/excel.conditionalformatcollection#getItemOrNullObject_id_)|Возвращает условный формат, идентифицированный его ID.|
|[GroupShapeCollection](/javascript/api/excel/excel.groupshapecollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.groupshapecollection#getItemOrNullObject_key_)|Получает фигуру с ее именем или ИД.|
|[Query](/javascript/api/excel/excel.query)|[error](/javascript/api/excel/excel.query#error)|Получает сообщение об ошибке запроса с последнего обновления запроса.|
||[loadedTo](/javascript/api/excel/excel.query#loadedTo)|Загружает запрос на тип объекта.|
||[loadedToDataModel](/javascript/api/excel/excel.query#loadedToDataModel)|Указывает, загружен ли запрос в модель данных.|
||[name](/javascript/api/excel/excel.query#name)|Получает имя запроса.|
||[refreshDate](/javascript/api/excel/excel.query#refreshDate)|Получает дату и время последнего обновления запроса.|
||[rowsLoadedCount](/javascript/api/excel/excel.query#rowsLoadedCount)|Получает количество строк, загруженных при последнем обновлении запроса.|
|[QueryCollection](/javascript/api/excel/excel.querycollection)|[getCount()](/javascript/api/excel/excel.querycollection#getCount__)|Получает количество запросов в книге.|
||[getItem(key: string)](/javascript/api/excel/excel.querycollection#getItem_key_)|Получает запрос из коллекции на основе его имени.|
||[items](/javascript/api/excel/excel.querycollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Range](/javascript/api/excel/excel.range)|[getPrecedents()](/javascript/api/excel/excel.range#getPrecedents__)|Возвращает объект, представляющего диапазон, содержащий все прецеденты ячейки в одной и той же таблице или `WorkbookRangeAreas` в нескольких таблицах.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.shapecollection#getItemOrNullObject_key_)|Получает фигуру с ее именем или ИД.|
|[StyleCollection](/javascript/api/excel/excel.stylecollection)|[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.stylecollection#getItemOrNullObject_name_)|Получает стиль по имени.|
|[TableScopedCollection](/javascript/api/excel/excel.tablescopedcollection)|[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablescopedcollection#getItemOrNullObject_key_)|Получает таблицу по имени или ИД.|
|[Workbook](/javascript/api/excel/excel.workbook)|[запросы](/javascript/api/excel/excel.workbook#queries)|Возвращает коллекцию запросов Power Query, которые являются частью книги.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[onProtectionChanged](/javascript/api/excel/excel.worksheet#onProtectionChanged)|Возникает при смене состояния защиты таблицы.|
||[tabId](/javascript/api/excel/excel.worksheet#tabId)|Возвращает значение, представляющее этот таблицу, которую можно прочитать в Open Office XML.|
|[WorksheetChangedEventArgs](/javascript/api/excel/excel.worksheetchangedeventargs)|[changeDirectionState](/javascript/api/excel/excel.worksheetchangedeventargs#changeDirectionState)|Представляет изменение в направлении, в которое будут сдвигаться ячейки в таблице при удалении или вставке ячейки.|
||[triggerSource](/javascript/api/excel/excel.worksheetchangedeventargs#triggerSource)|Представляет источник триггера события.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onProtectionChanged](/javascript/api/excel/excel.worksheetcollection#onProtectionChanged)|Возникает при смене состояния защиты таблицы.|
|[WorksheetProtectionChangedEventArgs](/javascript/api/excel/excel.worksheetprotectionchangedeventargs)|[isProtected](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#isProtected)|Получает текущее состояние защиты таблицы.|
||[источник](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#source)|Источник события.|
||[type](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetprotectionchangedeventargs#worksheetId)|Получает ID таблицы, в которой изменен статус защиты.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.14&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)
