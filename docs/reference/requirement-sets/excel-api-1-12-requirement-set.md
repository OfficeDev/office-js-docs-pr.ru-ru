---
title: Набор API JavaScript Excel 1.12
description: Сведения о наборе требований ExcelApi 1.12.
ms.date: 04/01/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: d66f5797d41c8c07f66fcc8069cd4687cd8d8118
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652218"
---
# <a name="whats-new-in-excel-javascript-api-112"></a>Что нового в API JavaScript Excel 1.12

ExcelApi 1.12 увеличил поддержку формул в диапазонах, добавив API для отслеживания динамических массивов и поиска прямых прецедентов формулы. Кроме того, добавлен контроль API фильтров PivotTable. Улучшения также были сделаны в областях комментариев, параметров культуры и пользовательских свойств.

| Функциональная область | Описание | Соответствующие объекты |
|:--- |:--- |:--- |
| [События комментариев](../../excel/excel-add-ins-comments.md#comment-events) | Добавляет события для добавления, изменения и удаления в коллекцию комментариев.| [CommentCollection](/javascript/api/excel/excel.commentcollection) |
| Параметры культуры [даты и времени](../../excel/excel-add-ins-workbooks.md#access-application-culture-settings) | Предоставляет доступ к дополнительным культурным настройкам даты и времени форматирования. | [CultureInfo](/javascript/api/excel/excel.cultureinfo), [Приложение NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [](/javascript/api/excel/excel.application) |
| [Прямые прецеденты](../../excel/excel-add-ins-ranges-precedents.md) | Возвращает диапазоны, используемые для оценки формулы ячейки.| [Range](/javascript/api/excel/excel.range#getdirectprecedents--) |
| Фильтры поворота | Применяет фильтры, управляемые значением, к полям PivotTable. | [PivotField](/javascript/api/excel/excel.pivotfield#applyfilter-filter-), [PivotFilters](/javascript/api/excel/excel.pivotFilters) |
| [Разлиение диапазона](../../excel/excel-add-ins-ranges-dynamic-arrays.md) | Позволяет надстройки находить диапазоны, связанные с [динамическими результатами массива.](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531) | [Range](/javascript/api/excel/excel.range) |
| [Настраиваемые свойства на уровне таблицы](../../excel/excel-add-ins-workbooks.md#worksheet-level-custom-properties) | Позволяет настраивать свойства на уровне таблицы, в дополнение к области на уровне книг. | [WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty), [WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в API Excel JavaScript, за набором 1.12. Чтобы просмотреть справочную документацию API для всех API, поддерживаемых требованием API Excel JavaScript, установленным 1.12 или ранее, см. в справке к API Excel в наборе требований [1.12 или более ранних](/javascript/api/excel?view=excel-js-1.12&preserve-view=true).

| Класс | Поля | Описание |
|:---|:---|:---|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[textOrientation](/javascript/api/excel/excel.chartaxistitle#textorientation)|Указывает угол, на который ориентирован текст для заголовка оси диаграммы.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[getDimensionValues(dimension: Excel.ChartSeriesDimension)](/javascript/api/excel/excel.chartseries#getdimensionvalues-dimension-)|Получает значения из одного измерения серии диаграмм.|
|[Comment](/javascript/api/excel/excel.comment)|[contentType](/javascript/api/excel/excel.comment#contenttype)|Получает тип контента комментария.|
|[CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs)|[commentDetails](/javascript/api/excel/excel.commentaddedeventargs#commentdetails)|Получите массив CommentDetail, содержащий id и Id-адреса комментариев соответствующих ответов.|
||[source](/javascript/api/excel/excel.commentaddedeventargs#source)|Указывает источник события.|
||[type](/javascript/api/excel/excel.commentaddedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.commentaddedeventargs#worksheetid)|Получает Id таблицы, в которой произошло событие.|
|[CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventargs)|[changeType](/javascript/api/excel/excel.commentchangedeventargs#changetype)|Получает тип изменений, который представляет, как запускается измененное событие.|
||[commentDetails](/javascript/api/excel/excel.commentchangedeventargs#commentdetails)|Получите массив CommentDetail, содержащий id и Id-адреса комментариев соответствующих ответов.|
||[source](/javascript/api/excel/excel.commentchangedeventargs#source)|Указывает источник события.|
||[type](/javascript/api/excel/excel.commentchangedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.commentchangedeventargs#worksheetid)|Получает Id таблицы, в которой произошло событие.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[onAdded](/javascript/api/excel/excel.commentcollection#onadded)|Возникает при добавлении комментариев.|
||[onChanged](/javascript/api/excel/excel.commentcollection#onchanged)|Возникает при смене комментариев или ответов в коллекции комментариев, в том числе при удалении ответов.|
||[onDeleted](/javascript/api/excel/excel.commentcollection#ondeleted)|Происходит, когда комментарии удаляются в коллекции комментариев.|
|[CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs)|[commentDetails](/javascript/api/excel/excel.commentdeletedeventargs#commentdetails)|Получите массив CommentDetail, содержащий id и Id-адреса комментариев соответствующих ответов.|
||[source](/javascript/api/excel/excel.commentdeletedeventargs#source)|Указывает источник события.|
||[type](/javascript/api/excel/excel.commentdeletedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.commentdeletedeventargs#worksheetid)|Получает Id таблицы, в которой произошло событие.|
|[CommentDetail](/javascript/api/excel/excel.commentdetail)|[commentId](/javascript/api/excel/excel.commentdetail#commentid)|Представляет id комментария.|
||[replyIds](/javascript/api/excel/excel.commentdetail#replyids)|Представляет ids соответствующих ответов принадлежат к комментарию.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[contentType](/javascript/api/excel/excel.commentreply#contenttype)|Тип контента ответа.|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[datetimeFormat](/javascript/api/excel/excel.cultureinfo#datetimeformat)|Определяет культурный формат отображения даты и времени.|
|[DatetimeFormatInfo](/javascript/api/excel/excel.datetimeformatinfo)|[dateSeparator](/javascript/api/excel/excel.datetimeformatinfo#dateseparator)|Получает строку, используемую в качестве сепаратора дат.|
||[longDatePattern](/javascript/api/excel/excel.datetimeformatinfo#longdatepattern)|Получает строку формата для длинного значения даты.|
||[longTimePattern](/javascript/api/excel/excel.datetimeformatinfo#longtimepattern)|Получает строку формата в течение длительного времени.|
||[shortDatePattern](/javascript/api/excel/excel.datetimeformatinfo#shortdatepattern)|Получает строку формата для краткого значения даты.|
||[timeSeparator](/javascript/api/excel/excel.datetimeformatinfo#timeseparator)|Получает строку, используемую в качестве сепаратора времени.|
|[PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter)|[компаратор](/javascript/api/excel/excel.pivotdatefilter#comparator)|Компаратор — это статическое значение, с которым сравниваются другие значения.|
||[условие](/javascript/api/excel/excel.pivotdatefilter#condition)|Указывает условие фильтра, которое определяет необходимые критерии фильтрации.|
||[эксклюзив](/javascript/api/excel/excel.pivotdatefilter#exclusive)|Если это так, фильтр *исключает элементы,* которые соответствуют критериям.|
||[lowerBound](/javascript/api/excel/excel.pivotdatefilter#lowerbound)|Нижний предел диапазона для состояния `Between` фильтра.|
||[upperBound](/javascript/api/excel/excel.pivotdatefilter#upperbound)|Верхний предел диапазона для состояния `Between` фильтра.|
||[wholeDays](/javascript/api/excel/excel.pivotdatefilter#wholedays)|Для `Equals` , , и фильтр `Before` `After` `Between` условия, указывает, если сравнения должны быть сделаны в течение целых дней.|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[applyFilter(filter: Excel.PivotFilters)](/javascript/api/excel/excel.pivotfield#applyfilter-filter-)|Задает один или несколько текущих pivotFilters поля и применяет их к полю.|
||[clearAllFilters()](/javascript/api/excel/excel.pivotfield#clearallfilters--)|Очищает все критерии от всех фильтров поля.|
||[clearFilter(filterType: Excel.PivotFilterType)](/javascript/api/excel/excel.pivotfield#clearfilter-filtertype-)|Очищает все существующие критерии от фильтра поля данного типа (если он применяется в настоящее время).|
||[getFilters()](/javascript/api/excel/excel.pivotfield#getfilters--)|Получает все фильтры, применяемые в настоящее время на поле.|
||[isFiltered(filterType?: Excel.PivotFilterType)](/javascript/api/excel/excel.pivotfield#isfiltered-filtertype-)|Проверяет, есть ли на поле примененные фильтры.|
|[PivotFilters](/javascript/api/excel/excel.pivotfilters)|[dateFilter](/javascript/api/excel/excel.pivotfilters#datefilter)|В настоящее время применяется фильтр дат PivotField.|
||[labelFilter](/javascript/api/excel/excel.pivotfilters#labelfilter)|Фильтр меток PivotField в настоящее время применяется.|
||[manualFilter](/javascript/api/excel/excel.pivotfilters#manualfilter)|В настоящее время применяется ручной фильтр PivotField.|
||[valueFilter](/javascript/api/excel/excel.pivotfilters#valuefilter)|В настоящее время применяется фильтр значений PivotField.|
|[PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter)|[компаратор](/javascript/api/excel/excel.pivotlabelfilter#comparator)|Компаратор — это статическое значение, с которым сравниваются другие значения.|
||[условие](/javascript/api/excel/excel.pivotlabelfilter#condition)|Указывает условие фильтра, которое определяет необходимые критерии фильтрации.|
||[эксклюзив](/javascript/api/excel/excel.pivotlabelfilter#exclusive)|Если это так, фильтр *исключает элементы,* которые соответствуют критериям.|
||[lowerBound](/javascript/api/excel/excel.pivotlabelfilter#lowerbound)|Нижний предел диапазона для условия Между фильтрами.|
||[substring](/javascript/api/excel/excel.pivotlabelfilter#substring)|Подстройка, используемая для `BeginsWith` `EndsWith` и `Contains` фильтрации условий.|
||[upperBound](/javascript/api/excel/excel.pivotlabelfilter#upperbound)|Верхний предел диапазона для условия Между фильтрами.|
|[PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter)|[selectedItems](/javascript/api/excel/excel.pivotmanualfilter#selecteditems)|Список выбранных элементов, которые необходимо фильтровать вручную.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[allowMultipleFiltersPerField](/javascript/api/excel/excel.pivottable#allowmultiplefiltersperfield)|Указывает, разрешает ли pivotTable применение нескольких pivotFilters в заданной pivotField в таблице.|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getCount()](/javascript/api/excel/excel.pivottablescopedcollection#getcount--)|Получает число pivotTables в коллекции.|
||[getFirst()](/javascript/api/excel/excel.pivottablescopedcollection#getfirst--)|Получает первый pivotTable в коллекции.|
||[getItem(key: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitem-key-)|Получает сводную таблицу по имени.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.pivottablescopedcollection#getitemornullobject-name-)|Получает сводную таблицу по имени.|
||[items](/javascript/api/excel/excel.pivottablescopedcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter)|[компаратор](/javascript/api/excel/excel.pivotvaluefilter#comparator)|Компаратор — это статическое значение, с которым сравниваются другие значения.|
||[условие](/javascript/api/excel/excel.pivotvaluefilter#condition)|Указывает условие фильтра, которое определяет необходимые критерии фильтрации.|
||[эксклюзив](/javascript/api/excel/excel.pivotvaluefilter#exclusive)|Если это так, фильтр *исключает элементы,* которые соответствуют критериям.|
||[lowerBound](/javascript/api/excel/excel.pivotvaluefilter#lowerbound)|Нижний предел диапазона для состояния `Between` фильтра.|
||[selectionType](/javascript/api/excel/excel.pivotvaluefilter#selectiontype)|Указывает, является ли фильтр для элементов верхнего и нижнего N, верхнего и нижнего N-процентов или суммы N верхнего или нижнего.|
||[пороговое значение](/javascript/api/excel/excel.pivotvaluefilter#threshold)|Пороговое число элементов , процентов или сумм, которые необходимо отфильтровать для состояния фильтра Верхнего или Нижнего уровня.|
||[upperBound](/javascript/api/excel/excel.pivotvaluefilter#upperbound)|Верхний предел диапазона для состояния `Between` фильтра.|
||[value](/javascript/api/excel/excel.pivotvaluefilter#value)|Имя выбранного "значения" в поле для фильтрации.|
|[Range](/javascript/api/excel/excel.range)|[getDirectPrecedents()](/javascript/api/excel/excel.range#getdirectprecedents--)|Возвращает объект WorkbookRangeAreas, представляющего диапазон, содержащий все прямые прецеденты ячейки в одной и той же таблице или в нескольких таблицах.|
||[getPivotTables (fullyContained?: boolean)](/javascript/api/excel/excel.range#getpivottables-fullycontained-)|Получает объемную коллекцию pivotTables, которые пересекаются с диапазоном.|
||[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|Получает объект диапазона, содержащий базовую ячейку для переносимой ячейки.|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|Получает объект диапазона, содержащий базовую ячейку для переносимой ячейки.|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|Получает объект range, содержащий диапазон переноса при вызове для базовой ячейки.|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|Получает объект range, содержащий диапазон переноса при вызове для базовой ячейки.|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|Указывает, есть ли во всех ячейках граница переноса.|
||[numberFormatCategories](/javascript/api/excel/excel.range#numberformatcategories)|Представляет категорию формата номеров каждой ячейки.|
||[savedAsArray](/javascript/api/excel/excel.range#savedasarray)|Представляет, будут ли все ячейки сохранены в качестве формулы массива.|
|[RangeAreasCollection](/javascript/api/excel/excel.rangeareascollection)|[getCount()](/javascript/api/excel/excel.rangeareascollection#getcount--)|Получает количество объектов RangeAreas в этой коллекции.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangeareascollection#getitemat-index-)|Возвращает объект RangeAreas в зависимости от положения в коллекции.|
||[items](/javascript/api/excel/excel.rangeareascollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[WorkbookRangeAreas](/javascript/api/excel/excel.workbookrangeareas)|[getRangeAreasBySheet (клавиша: строка)](/javascript/api/excel/excel.workbookrangeareas#getrangeareasbysheet-key-)|Возвращает объект на основе id или имени таблицы `RangeAreas` в коллекции.|
||[getRangeAreasOrNullObjectBySheet (key: string)](/javascript/api/excel/excel.workbookrangeareas#getrangeareasornullobjectbysheet-key-)|Возвращает объект на основе имени или ид таблицы `RangeAreas` в коллекции.|
||[addresses](/javascript/api/excel/excel.workbookrangeareas#addresses)|Возвращает массив адресов в стиле A1.|
||[areas](/javascript/api/excel/excel.workbookrangeareas#areas)|Возвращает `RangeAreasCollection` объект.|
||[диапазоны](/javascript/api/excel/excel.workbookrangeareas#ranges)|Возвращает диапазоны, составляющие этот объект в `RangeCollection` объекте.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[customProperties](/javascript/api/excel/excel.worksheet#customproperties)|Получает коллекцию пользовательских свойств на уровне таблицы.|
|[WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty)|[delete()](/javascript/api/excel/excel.worksheetcustomproperty#delete--)|Удаляет настраиваемое свойство.|
||[key](/javascript/api/excel/excel.worksheetcustomproperty#key)|Возвращает ключ настраиваемого свойства.|
||[value](/javascript/api/excel/excel.worksheetcustomproperty#value)|Получает или задает значение настраиваемого свойства.|
|[WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|[add(key: string, value: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#add-key--value-)|Добавляет новое настраиваемую свойство, которое сопополяет с предоставленным ключом.|
||[getCount()](/javascript/api/excel/excel.worksheetcustompropertycollection#getcount--)|Получает количество настраиваемого свойства на этом таблице.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitem-key-)|Возвращает объект настраиваемого свойства по ключу, указываемому без учета регистра.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitemornullobject-key-)|Возвращает объект настраиваемого свойства по ключу, указываемому без учета регистра.|
||[items](/javascript/api/excel/excel.worksheetcustompropertycollection#items)|Получает загруженные дочерние элементы в этой коллекции.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.12&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)
