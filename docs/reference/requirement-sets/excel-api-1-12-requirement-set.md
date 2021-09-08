---
title: Excel Требования к API JavaScript 1.12
description: Сведения о наборе требований ExcelApi 1.12.
ms.date: 04/01/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 10587b84ba476b91cdd56d8472e551348b3a718b
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2021
ms.locfileid: "58936946"
---
# <a name="whats-new-in-excel-javascript-api-112"></a>Новые возможности в Excel API JavaScript 1.12

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

В следующей таблице перечислены API в Excel API JavaScript, за набором 1.12. Чтобы просмотреть справочную документацию API для всех API, поддерживаемых Excel API JavaScript, установленного 1.12 или ранее, см. в Excel API в наборе требований [1.12](/javascript/api/excel?view=excel-js-1.12&preserve-view=true)или ранее .

| Класс | Поля | Описание |
|:---|:---|:---|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[textOrientation](/javascript/api/excel/excel.chartaxistitle#textOrientation)|Указывает угол, на который ориентирован текст для заголовка оси диаграммы.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[getDimensionValues (размер: Excel. ChartSeriesDimension)](/javascript/api/excel/excel.chartseries#getDimensionValues_dimension_)|Получает значения из одного измерения серии диаграмм.|
|[Comment](/javascript/api/excel/excel.comment)|[contentType](/javascript/api/excel/excel.comment#contentType)|Получает тип контента комментария.|
|[CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs)|[commentDetails](/javascript/api/excel/excel.commentaddedeventargs#commentDetails)|Получает массив, содержащий ID и ID комментариев связанных с ним `CommentDetail` ответов.|
||[source](/javascript/api/excel/excel.commentaddedeventargs#source)|Указывает источник события.|
||[type](/javascript/api/excel/excel.commentaddedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.commentaddedeventargs#worksheetId)|Получает ID таблицы, в которой произошло событие.|
|[CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventargs)|[changeType](/javascript/api/excel/excel.commentchangedeventargs#changeType)|Получает тип изменений, который представляет, как запускается измененное событие.|
||[commentDetails](/javascript/api/excel/excel.commentchangedeventargs#commentDetails)|Получите массив, содержащий ID и ID комментариев связанных с ним `CommentDetail` ответов.|
||[source](/javascript/api/excel/excel.commentchangedeventargs#source)|Указывает источник события.|
||[type](/javascript/api/excel/excel.commentchangedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.commentchangedeventargs#worksheetId)|Получает ID таблицы, в которой произошло событие.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[onAdded](/javascript/api/excel/excel.commentcollection#onAdded)|Возникает при добавлении комментариев.|
||[onChanged](/javascript/api/excel/excel.commentcollection#onChanged)|Возникает при смене комментариев или ответов в коллекции комментариев, в том числе при удалении ответов.|
||[onDeleted](/javascript/api/excel/excel.commentcollection#onDeleted)|Происходит, когда комментарии удаляются в коллекции комментариев.|
|[CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs)|[commentDetails](/javascript/api/excel/excel.commentdeletedeventargs#commentDetails)|Получает массив, содержащий ID и ID комментариев связанных с ним `CommentDetail` ответов.|
||[source](/javascript/api/excel/excel.commentdeletedeventargs#source)|Указывает источник события.|
||[type](/javascript/api/excel/excel.commentdeletedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.commentdeletedeventargs#worksheetId)|Получает ID таблицы, в которой произошло событие.|
|[CommentDetail](/javascript/api/excel/excel.commentdetail)|[commentId](/javascript/api/excel/excel.commentdetail#commentId)|Представляет ID комментария.|
||[replyIds](/javascript/api/excel/excel.commentdetail#replyIds)|Представляет ID-адреса соответствующих ответов, которые принадлежат комментарию.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[contentType](/javascript/api/excel/excel.commentreply#contentType)|Тип контента ответа.|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[datetimeFormat](/javascript/api/excel/excel.cultureinfo#datetimeFormat)|Определяет культурный формат отображения даты и времени.|
|[DatetimeFormatInfo](/javascript/api/excel/excel.datetimeformatinfo)|[dateSeparator](/javascript/api/excel/excel.datetimeformatinfo#dateSeparator)|Получает строку, используемую в качестве сепаратора дат.|
||[longDatePattern](/javascript/api/excel/excel.datetimeformatinfo#longDatePattern)|Получает строку формата для длинного значения даты.|
||[longTimePattern](/javascript/api/excel/excel.datetimeformatinfo#longTimePattern)|Получает строку формата в течение длительного времени.|
||[shortDatePattern](/javascript/api/excel/excel.datetimeformatinfo#shortDatePattern)|Получает строку формата для краткого значения даты.|
||[timeSeparator](/javascript/api/excel/excel.datetimeformatinfo#timeSeparator)|Получает строку, используемую в качестве сепаратора времени.|
|[PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter)|[компаратор](/javascript/api/excel/excel.pivotdatefilter#comparator)|Компаратор — это статическое значение, с которым сравниваются другие значения.|
||[условие](/javascript/api/excel/excel.pivotdatefilter#condition)|Указывает условие фильтра, которое определяет необходимые критерии фильтрации.|
||[эксклюзив](/javascript/api/excel/excel.pivotdatefilter#exclusive)|Если `true` фильтр исключает *элементы,* которые соответствуют критериям.|
||[lowerBound](/javascript/api/excel/excel.pivotdatefilter#lowerBound)|Нижний предел диапазона для состояния `between` фильтра.|
||[upperBound](/javascript/api/excel/excel.pivotdatefilter#upperBound)|Верхний предел диапазона для состояния `between` фильтра.|
||[wholeDays](/javascript/api/excel/excel.pivotdatefilter#wholeDays)|Для `equals` , , и фильтр `before` `after` `between` условия, указывает, если сравнения должны быть сделаны в течение целых дней.|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[applyFilter(filter: Excel. PivotFilters)](/javascript/api/excel/excel.pivotfield#applyFilter_filter_)|Задает один или несколько текущих pivotFilters поля и применяет их к полю.|
||[clearAllFilters()](/javascript/api/excel/excel.pivotfield#clearAllFilters__)|Очищает все критерии от всех фильтров поля.|
||[clearFilter(filterType: Excel. PivotFilterType)](/javascript/api/excel/excel.pivotfield#clearFilter_filterType_)|Очищает все существующие критерии от фильтра поля данного типа (если он применяется в настоящее время).|
||[getFilters()](/javascript/api/excel/excel.pivotfield#getFilters__)|Получает все фильтры, применяемые в настоящее время на поле.|
||[isFiltered (filterType?: Excel. PivotFilterType)](/javascript/api/excel/excel.pivotfield#isFiltered_filterType_)|Проверяет, есть ли на поле примененные фильтры.|
|[PivotFilters](/javascript/api/excel/excel.pivotfilters)|[dateFilter](/javascript/api/excel/excel.pivotfilters#dateFilter)|В настоящее время применяется фильтр дат PivotField.|
||[labelFilter](/javascript/api/excel/excel.pivotfilters#labelFilter)|Фильтр меток PivotField в настоящее время применяется.|
||[manualFilter](/javascript/api/excel/excel.pivotfilters#manualFilter)|В настоящее время применяется ручной фильтр PivotField.|
||[valueFilter](/javascript/api/excel/excel.pivotfilters#valueFilter)|В настоящее время применяется фильтр значений PivotField.|
|[PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter)|[компаратор](/javascript/api/excel/excel.pivotlabelfilter#comparator)|Компаратор — это статическое значение, с которым сравниваются другие значения.|
||[условие](/javascript/api/excel/excel.pivotlabelfilter#condition)|Указывает условие фильтра, которое определяет необходимые критерии фильтрации.|
||[эксклюзив](/javascript/api/excel/excel.pivotlabelfilter#exclusive)|Если `true` фильтр исключает *элементы,* которые соответствуют критериям.|
||[lowerBound](/javascript/api/excel/excel.pivotlabelfilter#lowerBound)|Нижний предел диапазона для состояния `between` фильтра.|
||[подстройка](/javascript/api/excel/excel.pivotlabelfilter#substring)|Подстройка, используемая для `beginsWith` `endsWith` и `contains` фильтрации условий.|
||[upperBound](/javascript/api/excel/excel.pivotlabelfilter#upperBound)|Верхний предел диапазона для состояния `between` фильтра.|
|[PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter)|[selectedItems](/javascript/api/excel/excel.pivotmanualfilter#selectedItems)|Список выбранных элементов, которые необходимо фильтровать вручную.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[allowMultipleFiltersPerField](/javascript/api/excel/excel.pivottable#allowMultipleFiltersPerField)|Указывает, разрешает ли pivotTable применение нескольких pivotFilters в заданной pivotField в таблице.|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getCount()](/javascript/api/excel/excel.pivottablescopedcollection#getCount__)|Получает число pivotTables в коллекции.|
||[getFirst()](/javascript/api/excel/excel.pivottablescopedcollection#getFirst__)|Получает первый pivotTable в коллекции.|
||[getItem(key: string)](/javascript/api/excel/excel.pivottablescopedcollection#getItem_key_)|Получает сводную таблицу по имени.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.pivottablescopedcollection#getItemOrNullObject_name_)|Получает сводную таблицу по имени.|
||[items](/javascript/api/excel/excel.pivottablescopedcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter)|[компаратор](/javascript/api/excel/excel.pivotvaluefilter#comparator)|Компаратор — это статическое значение, с которым сравниваются другие значения.|
||[условие](/javascript/api/excel/excel.pivotvaluefilter#condition)|Указывает условие фильтра, которое определяет необходимые критерии фильтрации.|
||[эксклюзив](/javascript/api/excel/excel.pivotvaluefilter#exclusive)|Если `true` фильтр исключает *элементы,* которые соответствуют критериям.|
||[lowerBound](/javascript/api/excel/excel.pivotvaluefilter#lowerBound)|Нижний предел диапазона для состояния `between` фильтра.|
||[selectionType](/javascript/api/excel/excel.pivotvaluefilter#selectionType)|Указывает, является ли фильтр для элементов верхнего и нижнего N, верхнего и нижнего N-процентов или суммы N верхнего или нижнего.|
||[пороговое значение](/javascript/api/excel/excel.pivotvaluefilter#threshold)|Пороговое число элементов , процентов или сумм, которые необходимо отфильтровать для состояния верхнего или нижнего фильтра.|
||[upperBound](/javascript/api/excel/excel.pivotvaluefilter#upperBound)|Верхний предел диапазона для состояния `between` фильтра.|
||[value](/javascript/api/excel/excel.pivotvaluefilter#value)|Имя выбранного "значения" в поле для фильтрации.|
|[Range](/javascript/api/excel/excel.range)|[getDirectPrecedents()](/javascript/api/excel/excel.range#getDirectPrecedents__)|Возвращает объект, представляющего диапазон, содержащий все прямые прецеденты ячейки в одной и той же таблице или `WorkbookRangeAreas` в нескольких таблицах.|
||[getPivotTables (fullyContained?: boolean)](/javascript/api/excel/excel.range#getPivotTables_fullyContained_)|Получает объемную коллекцию pivotTables, которые пересекаются с диапазоном.|
||[getSpillParent()](/javascript/api/excel/excel.range#getSpillParent__)|Получает объект диапазона, содержащий базовую ячейку для переносимой ячейки.|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getSpillParentOrNullObject__)|Получает объект диапазона, содержащий якорную ячейку для пролитой ячейки.|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getSpillingToRange__)|Получает объект range, содержащий диапазон переноса при вызове для базовой ячейки.|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getSpillingToRangeOrNullObject__)|Получает объект range, содержащий диапазон переноса при вызове для базовой ячейки.|
||[hasSpill](/javascript/api/excel/excel.range#hasSpill)|Указывает, есть ли во всех ячейках граница переноса.|
||[numberFormatCategories](/javascript/api/excel/excel.range#numberFormatCategories)|Представляет категорию формата номеров каждой ячейки.|
||[savedAsArray](/javascript/api/excel/excel.range#savedAsArray)|Представляет, будут ли сохранены все ячейки в качестве формулы массива.|
|[RangeAreasCollection](/javascript/api/excel/excel.rangeareascollection)|[getCount()](/javascript/api/excel/excel.rangeareascollection#getCount__)|Получает количество объектов `RangeAreas` в этой коллекции.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangeareascollection#getItemAt_index_)|Возвращает объект `RangeAreas` в зависимости от позиции в коллекции.|
||[items](/javascript/api/excel/excel.rangeareascollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[WorkbookRangeAreas](/javascript/api/excel/excel.workbookrangeareas)|[getRangeAreasBySheet (клавиша: строка)](/javascript/api/excel/excel.workbookrangeareas#getRangeAreasBySheet_key_)|Возвращает объект на основе ИД или имени таблицы `RangeAreas` в коллекции.|
||[getRangeAreasOrNullObjectBySheet (key: string)](/javascript/api/excel/excel.workbookrangeareas#getRangeAreasOrNullObjectBySheet_key_)|Возвращает объект на основе имени или ИД таблицы `RangeAreas` в коллекции.|
||[addresses](/javascript/api/excel/excel.workbookrangeareas#addresses)|Возвращает массив адресов в стиле A1.|
||[areas](/javascript/api/excel/excel.workbookrangeareas#areas)|Возвращает `RangeAreasCollection` объект.|
||[диапазоны](/javascript/api/excel/excel.workbookrangeareas#ranges)|Возвращает диапазоны, составляющие этот объект в `RangeCollection` объекте.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[customProperties](/javascript/api/excel/excel.worksheet#customProperties)|Получает коллекцию пользовательских свойств на уровне таблицы.|
|[WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty)|[delete()](/javascript/api/excel/excel.worksheetcustomproperty#delete__)|Удаляет настраиваемое свойство.|
||[key](/javascript/api/excel/excel.worksheetcustomproperty#key)|Возвращает ключ настраиваемого свойства.|
||[value](/javascript/api/excel/excel.worksheetcustomproperty#value)|Получает или задает значение настраиваемого свойства.|
|[WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|[add(key: string, value: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#add_key__value_)|Добавляет новое настраиваемую свойство, которое сопополяет с предоставленным ключом.|
||[getCount()](/javascript/api/excel/excel.worksheetcustompropertycollection#getCount__)|Получает количество настраиваемого свойства на этом таблице.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getItem_key_)|Возвращает объект настраиваемого свойства по ключу, указываемому без учета регистра.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getItemOrNullObject_key_)|Возвращает объект настраиваемого свойства по ключу, указываемому без учета регистра.|
||[items](/javascript/api/excel/excel.worksheetcustompropertycollection#items)|Получает загруженные дочерние элементы в этой коллекции.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.12&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)
