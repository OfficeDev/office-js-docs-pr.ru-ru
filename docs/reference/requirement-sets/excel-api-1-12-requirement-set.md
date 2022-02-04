---
title: Excel API JavaScript установлено 1.12
description: Сведения о наборе требований ExcelApi 1.12.
ms.date: 04/01/2021
ms.prod: excel
ms.localizationpriority: medium
---

# <a name="whats-new-in-excel-javascript-api-112"></a>Новые возможности в Excel API JavaScript 1.12

ExcelApi 1.12 увеличил поддержку формул в диапазонах, добавив API для отслеживания динамических массивов и поиска прямых прецедентов формулы. Кроме того, добавлен контроль API фильтров PivotTable. Улучшения также были сделаны в областях комментариев, параметров культуры и пользовательских свойств.

| Функциональная область | Описание | Соответствующие объекты |
|:--- |:--- |:--- |
| [События комментариев](../../excel/excel-add-ins-comments.md#comment-events) | Добавляет события для добавления, изменения и удаления в коллекцию комментариев.| [CommentCollection](/javascript/api/excel/excel.commentcollection) |
| Параметры культуры [даты и времени](../../excel/excel-add-ins-workbooks.md#access-application-culture-settings) | Предоставляет доступ к дополнительным культурным настройкам даты и времени форматирования. | [CultureInfo](/javascript/api/excel/excel.cultureinfo), [приложение NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [](/javascript/api/excel/excel.application) |
| [Прямые прецеденты](../../excel/excel-add-ins-ranges-precedents.md) | Возвращает диапазоны, используемые для оценки формулы ячейки.| [Range](/javascript/api/excel/excel.range#getdirectprecedents--) |
| Фильтры поворота | Применяет фильтры, управляемые значением, к полям PivotTable. | [PivotField](/javascript/api/excel/excel.pivotfield#applyfilter-filter-), [PivotFilters](/javascript/api/excel/excel.pivotfilters) |
| [Разлиение диапазона](../../excel/excel-add-ins-ranges-dynamic-arrays.md) | Позволяет надстройки находить диапазоны, связанные с [динамическими результатами массива](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531) . | [Range](/javascript/api/excel/excel.range) |
| [Настраиваемые свойства на уровне таблицы](../../excel/excel-add-ins-workbooks.md#worksheet-level-custom-properties) | Позволяет настраивать свойства на уровне таблицы, в дополнение к области на уровне книг. | [WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty), [WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в Excel API JavaScript, за набором 1.12. Чтобы просмотреть справочную документацию API для всех API, поддерживаемых Excel API JavaScript, установленного 1.12 или ранее, см. в Excel API в наборе требований [1.12 или ранее](/javascript/api/excel?view=excel-js-1.12&preserve-view=true).

| Класс | Поля | Описание |
|:---|:---|:---|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[textOrientation](/javascript/api/excel/excel.chartaxistitle#excel-excel-chartaxistitle-textorientation-member)|Указывает угол, на который ориентирован текст для заголовка оси диаграммы.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[getDimensionValues (размер: Excel. ChartSeriesDimension)](/javascript/api/excel/excel.chartseries#excel-excel-chartseries-getdimensionvalues-member(1))|Получает значения из одного измерения серии диаграмм.|
|[Comment](/javascript/api/excel/excel.comment)|[contentType](/javascript/api/excel/excel.comment#excel-excel-comment-contenttype-member)|Получает тип контента комментария.|
|[CommentAddedEventArgs](/javascript/api/excel/excel.commentaddedeventargs)|[commentDetails](/javascript/api/excel/excel.commentaddedeventargs#excel-excel-commentaddedeventargs-commentdetails-member)|Получает массив `CommentDetail` , содержащий ID и ID комментариев связанных с ним ответов.|
||[source](/javascript/api/excel/excel.commentaddedeventargs#excel-excel-commentaddedeventargs-source-member)|Указывает источник события.|
||[type](/javascript/api/excel/excel.commentaddedeventargs#excel-excel-commentaddedeventargs-type-member)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.commentaddedeventargs#excel-excel-commentaddedeventargs-worksheetid-member)|Получает ID таблицы, в которой произошло событие.|
|[CommentChangedEventArgs](/javascript/api/excel/excel.commentchangedeventargs)|[changeType](/javascript/api/excel/excel.commentchangedeventargs#excel-excel-commentchangedeventargs-changetype-member)|Получает тип изменений, который представляет, как запускается измененное событие.|
||[commentDetails](/javascript/api/excel/excel.commentchangedeventargs#excel-excel-commentchangedeventargs-commentdetails-member)|Получите массив `CommentDetail` , содержащий ID и ID комментариев связанных с ним ответов.|
||[source](/javascript/api/excel/excel.commentchangedeventargs#excel-excel-commentchangedeventargs-source-member)|Указывает источник события.|
||[type](/javascript/api/excel/excel.commentchangedeventargs#excel-excel-commentchangedeventargs-type-member)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.commentchangedeventargs#excel-excel-commentchangedeventargs-worksheetid-member)|Получает ID таблицы, в которой произошло событие.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[onAdded](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-onadded-member)|Возникает при добавлении комментариев.|
||[onChanged](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-onchanged-member)|Возникает при смене комментариев или ответов в коллекции комментариев, в том числе при удалении ответов.|
||[onDeleted](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-ondeleted-member)|Происходит, когда комментарии удаляются в коллекции комментариев.|
|[CommentDeletedEventArgs](/javascript/api/excel/excel.commentdeletedeventargs)|[commentDetails](/javascript/api/excel/excel.commentdeletedeventargs#excel-excel-commentdeletedeventargs-commentdetails-member)|Получает массив `CommentDetail` , содержащий ID и ID комментариев связанных с ним ответов.|
||[source](/javascript/api/excel/excel.commentdeletedeventargs#excel-excel-commentdeletedeventargs-source-member)|Указывает источник события.|
||[type](/javascript/api/excel/excel.commentdeletedeventargs#excel-excel-commentdeletedeventargs-type-member)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.commentdeletedeventargs#excel-excel-commentdeletedeventargs-worksheetid-member)|Получает ID таблицы, в которой произошло событие.|
|[CommentDetail](/javascript/api/excel/excel.commentdetail)|[commentId](/javascript/api/excel/excel.commentdetail#excel-excel-commentdetail-commentid-member)|Представляет ID комментария.|
||[replyIds](/javascript/api/excel/excel.commentdetail#excel-excel-commentdetail-replyids-member)|Представляет ID-адреса соответствующих ответов, которые принадлежат комментарию.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[contentType](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-contenttype-member)|Тип контента ответа.|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[datetimeFormat](/javascript/api/excel/excel.cultureinfo#excel-excel-cultureinfo-datetimeformat-member)|Определяет культурный формат отображения даты и времени.|
|[DatetimeFormatInfo](/javascript/api/excel/excel.datetimeformatinfo)|[dateSeparator](/javascript/api/excel/excel.datetimeformatinfo#excel-excel-datetimeformatinfo-dateseparator-member)|Получает строку, используемую в качестве сепаратора дат.|
||[longDatePattern](/javascript/api/excel/excel.datetimeformatinfo#excel-excel-datetimeformatinfo-longdatepattern-member)|Получает строку формата для длинного значения даты.|
||[longTimePattern](/javascript/api/excel/excel.datetimeformatinfo#excel-excel-datetimeformatinfo-longtimepattern-member)|Получает строку формата в течение длительного времени.|
||[shortDatePattern](/javascript/api/excel/excel.datetimeformatinfo#excel-excel-datetimeformatinfo-shortdatepattern-member)|Получает строку формата для краткого значения даты.|
||[timeSeparator](/javascript/api/excel/excel.datetimeformatinfo#excel-excel-datetimeformatinfo-timeseparator-member)|Получает строку, используемую в качестве сепаратора времени.|
|[PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter)|[компаратор](/javascript/api/excel/excel.pivotdatefilter#excel-excel-pivotdatefilter-comparator-member)|Компаратор — это статическое значение, с которым сравниваются другие значения.|
||[условие](/javascript/api/excel/excel.pivotdatefilter#excel-excel-pivotdatefilter-condition-member)|Указывает условие фильтра, которое определяет необходимые критерии фильтрации.|
||[эксклюзив](/javascript/api/excel/excel.pivotdatefilter#excel-excel-pivotdatefilter-exclusive-member)|Если `true`фильтр исключает *элементы* , которые соответствуют критериям.|
||[lowerBound](/javascript/api/excel/excel.pivotdatefilter#excel-excel-pivotdatefilter-lowerbound-member)|Нижний предел диапазона для состояния `between` фильтра.|
||[upperBound](/javascript/api/excel/excel.pivotdatefilter#excel-excel-pivotdatefilter-upperbound-member)|Верхний предел диапазона для состояния `between` фильтра.|
||[wholeDays](/javascript/api/excel/excel.pivotdatefilter#excel-excel-pivotdatefilter-wholedays-member)|Для `equals`, `before`, и `after`фильтр `between` условия, указывает, если сравнения должны быть сделаны в течение целых дней.|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[applyFilter(filter: Excel. PivotFilters)](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-applyfilter-member(1))|Задает один или несколько текущих pivotFilters поля и применяет их к полю.|
||[clearAllFilters()](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-clearallfilters-member(1))|Очищает все критерии от всех фильтров поля.|
||[clearFilter(filterType: Excel. PivotFilterType)](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-clearfilter-member(1))|Очищает все существующие критерии от фильтра поля данного типа (если он применяется в настоящее время).|
||[getFilters()](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-getfilters-member(1))|Получает все фильтры, применяемые в настоящее время на поле.|
||[isFiltered (filterType?: Excel. PivotFilterType)](/javascript/api/excel/excel.pivotfield#excel-excel-pivotfield-isfiltered-member(1))|Проверяет, есть ли на поле примененные фильтры.|
|[PivotFilters](/javascript/api/excel/excel.pivotfilters)|[dateFilter](/javascript/api/excel/excel.pivotfilters#excel-excel-pivotfilters-datefilter-member)|В настоящее время применяется фильтр дат PivotField.|
||[labelFilter](/javascript/api/excel/excel.pivotfilters#excel-excel-pivotfilters-labelfilter-member)|Фильтр меток PivotField в настоящее время применяется.|
||[manualFilter](/javascript/api/excel/excel.pivotfilters#excel-excel-pivotfilters-manualfilter-member)|В настоящее время применяется ручной фильтр PivotField.|
||[valueFilter](/javascript/api/excel/excel.pivotfilters#excel-excel-pivotfilters-valuefilter-member)|В настоящее время применяется фильтр значений PivotField.|
|[PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter)|[компаратор](/javascript/api/excel/excel.pivotlabelfilter#excel-excel-pivotlabelfilter-comparator-member)|Компаратор — это статическое значение, с которым сравниваются другие значения.|
||[условие](/javascript/api/excel/excel.pivotlabelfilter#excel-excel-pivotlabelfilter-condition-member)|Указывает условие фильтра, которое определяет необходимые критерии фильтрации.|
||[эксклюзив](/javascript/api/excel/excel.pivotlabelfilter#excel-excel-pivotlabelfilter-exclusive-member)|Если `true`фильтр исключает *элементы* , которые соответствуют критериям.|
||[lowerBound](/javascript/api/excel/excel.pivotlabelfilter#excel-excel-pivotlabelfilter-lowerbound-member)|Нижний предел диапазона для состояния `between` фильтра.|
||[substring](/javascript/api/excel/excel.pivotlabelfilter#excel-excel-pivotlabelfilter-substring-member)|Подстройка, используемая для `beginsWith`и `endsWith`фильтрации `contains` условий.|
||[upperBound](/javascript/api/excel/excel.pivotlabelfilter#excel-excel-pivotlabelfilter-upperbound-member)|Верхний предел диапазона для состояния `between` фильтра.|
|[PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter)|[selectedItems](/javascript/api/excel/excel.pivotmanualfilter#excel-excel-pivotmanualfilter-selecteditems-member)|Список выбранных элементов, которые необходимо фильтровать вручную.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[allowMultipleFiltersPerField](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-allowmultiplefiltersperfield-member)|Указывает, разрешает ли pivotTable применение нескольких pivotFilters в заданной pivotField в таблице.|
|[PivotTableScopedCollection](/javascript/api/excel/excel.pivottablescopedcollection)|[getCount()](/javascript/api/excel/excel.pivottablescopedcollection#excel-excel-pivottablescopedcollection-getcount-member(1))|Получает число pivotTables в коллекции.|
||[getFirst()](/javascript/api/excel/excel.pivottablescopedcollection#excel-excel-pivottablescopedcollection-getfirst-member(1))|Получает первый pivotTable в коллекции.|
||[getItem(key: string)](/javascript/api/excel/excel.pivottablescopedcollection#excel-excel-pivottablescopedcollection-getitem-member(1))|Получает сводную таблицу по имени.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.pivottablescopedcollection#excel-excel-pivottablescopedcollection-getitemornullobject-member(1))|Получает сводную таблицу по имени.|
||[items](/javascript/api/excel/excel.pivottablescopedcollection#excel-excel-pivottablescopedcollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter)|[компаратор](/javascript/api/excel/excel.pivotvaluefilter#excel-excel-pivotvaluefilter-comparator-member)|Компаратор — это статическое значение, с которым сравниваются другие значения.|
||[условие](/javascript/api/excel/excel.pivotvaluefilter#excel-excel-pivotvaluefilter-condition-member)|Указывает условие фильтра, которое определяет необходимые критерии фильтрации.|
||[эксклюзив](/javascript/api/excel/excel.pivotvaluefilter#excel-excel-pivotvaluefilter-exclusive-member)|Если `true`фильтр исключает *элементы* , которые соответствуют критериям.|
||[lowerBound](/javascript/api/excel/excel.pivotvaluefilter#excel-excel-pivotvaluefilter-lowerbound-member)|Нижний предел диапазона для состояния `between` фильтра.|
||[selectionType](/javascript/api/excel/excel.pivotvaluefilter#excel-excel-pivotvaluefilter-selectiontype-member)|Указывает, является ли фильтр для элементов верхнего и нижнего N, верхнего и нижнего N-процентов или суммы N верхнего или нижнего.|
||[пороговое значение](/javascript/api/excel/excel.pivotvaluefilter#excel-excel-pivotvaluefilter-threshold-member)|Пороговое число элементов , процентов или сумм, которые необходимо отфильтровать для состояния верхнего или нижнего фильтра.|
||[upperBound](/javascript/api/excel/excel.pivotvaluefilter#excel-excel-pivotvaluefilter-upperbound-member)|Верхний предел диапазона для состояния `between` фильтра.|
||[value](/javascript/api/excel/excel.pivotvaluefilter#excel-excel-pivotvaluefilter-value-member)|Имя выбранного "значения" в поле для фильтрации.|
|[Range](/javascript/api/excel/excel.range)|[getDirectPrecedents()](/javascript/api/excel/excel.range#excel-excel-range-getdirectprecedents-member(1))|Возвращает объект, `WorkbookRangeAreas` представляющего диапазон, содержащий все прямые прецеденты ячейки в одной и той же таблице или в нескольких таблицах.|
||[getPivotTables (fullyContained?: boolean)](/javascript/api/excel/excel.range#excel-excel-range-getpivottables-member(1))|Получает объемную коллекцию pivotTables, которые пересекаются с диапазоном.|
||[getSpillParent()](/javascript/api/excel/excel.range#excel-excel-range-getspillparent-member(1))|Получает объект диапазона, содержащий базовую ячейку для переносимой ячейки.|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#excel-excel-range-getspillparentornullobject-member(1))|Получает объект диапазона, содержащий якорную ячейку для пролитой ячейки.|
||[getSpillingToRange()](/javascript/api/excel/excel.range#excel-excel-range-getspillingtorange-member(1))|Получает объект range, содержащий диапазон переноса при вызове для базовой ячейки.|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#excel-excel-range-getspillingtorangeornullobject-member(1))|Получает объект range, содержащий диапазон переноса при вызове для базовой ячейки.|
||[hasSpill](/javascript/api/excel/excel.range#excel-excel-range-hasspill-member)|Указывает, есть ли во всех ячейках граница переноса.|
||[numberFormatCategories](/javascript/api/excel/excel.range#excel-excel-range-numberformatcategories-member)|Представляет категорию формата номеров каждой ячейки.|
||[savedAsArray](/javascript/api/excel/excel.range#excel-excel-range-savedasarray-member)|Представляет, будут ли сохранены все ячейки в качестве формулы массива.|
|[RangeAreasCollection](/javascript/api/excel/excel.rangeareascollection)|[getCount()](/javascript/api/excel/excel.rangeareascollection#excel-excel-rangeareascollection-getcount-member(1))|Получает количество объектов `RangeAreas` в этой коллекции.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangeareascollection#excel-excel-rangeareascollection-getitemat-member(1))|Возвращает объект `RangeAreas` в зависимости от позиции в коллекции.|
||[items](/javascript/api/excel/excel.rangeareascollection#excel-excel-rangeareascollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[WorkbookRangeAreas](/javascript/api/excel/excel.workbookrangeareas)|[addresses](/javascript/api/excel/excel.workbookrangeareas#excel-excel-workbookrangeareas-addresses-member)|Возвращает массив адресов в стиле A1.|
||[areas](/javascript/api/excel/excel.workbookrangeareas#excel-excel-workbookrangeareas-areas-member)|Возвращает объект `RangeAreasCollection` .|
||[getRangeAreasBySheet (клавиша: строка)](/javascript/api/excel/excel.workbookrangeareas#excel-excel-workbookrangeareas-getrangeareasbysheet-member(1))|Возвращает объект на `RangeAreas` основе ИД или имени таблицы в коллекции.|
||[getRangeAreasOrNullObjectBySheet (key: string)](/javascript/api/excel/excel.workbookrangeareas#excel-excel-workbookrangeareas-getrangeareasornullobjectbysheet-member(1))|Возвращает объект на `RangeAreas` основе имени или ИД таблицы в коллекции.|
||[диапазоны](/javascript/api/excel/excel.workbookrangeareas#excel-excel-workbookrangeareas-ranges-member)|Возвращает диапазоны, составляющие этот объект в объекте `RangeCollection` .|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[customProperties](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-customproperties-member)|Получает коллекцию пользовательских свойств на уровне таблицы.|
|[WorksheetCustomProperty](/javascript/api/excel/excel.worksheetcustomproperty)|[delete()](/javascript/api/excel/excel.worksheetcustomproperty#excel-excel-worksheetcustomproperty-delete-member(1))|Удаляет настраиваемое свойство.|
||[key](/javascript/api/excel/excel.worksheetcustomproperty#excel-excel-worksheetcustomproperty-key-member)|Возвращает ключ настраиваемого свойства.|
||[value](/javascript/api/excel/excel.worksheetcustomproperty#excel-excel-worksheetcustomproperty-value-member)|Получает или задает значение настраиваемого свойства.|
|[WorksheetCustomPropertyCollection](/javascript/api/excel/excel.worksheetcustompropertycollection)|[add(key: string, value: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#excel-excel-worksheetcustompropertycollection-add-member(1))|Добавляет новое настраиваемую свойство, которое сопополяет с предоставленным ключом.|
||[getCount()](/javascript/api/excel/excel.worksheetcustompropertycollection#excel-excel-worksheetcustompropertycollection-getcount-member(1))|Получает количество настраиваемого свойства на этом таблице.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#excel-excel-worksheetcustompropertycollection-getitem-member(1))|Возвращает объект настраиваемого свойства по ключу, указываемому без учета регистра.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#excel-excel-worksheetcustompropertycollection-getitemornullobject-member(1))|Возвращает объект настраиваемого свойства по ключу, указываемому без учета регистра.|
||[items](/javascript/api/excel/excel.worksheetcustompropertycollection#excel-excel-worksheetcustompropertycollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.12&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)
