---
title: Набор обязательных элементов API JavaScript для Excel 1,12
description: Сведения о наборе требований ExcelApi 1,12.
ms.date: 09/16/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 2f2fb04c914e26aacbd8815a1d173c8af9c09342
ms.sourcegitcommit: 0844ca7589ad3a6b0432fe126ca4e0ac9dbb80ce
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/18/2020
ms.locfileid: "47963824"
---
# <a name="whats-new-in-excel-javascript-api-112"></a>Новые возможности API JavaScript для Excel 1,12

В ExcelApi 1,12 расширена поддержка формул в диапазонах путем добавления API для отслеживания динамических массивов и поиска прямых и косвенных приоритетов формулы. Он также добавил API фильтров сводной таблицы. Кроме того, в области "Комментарии", "параметры культуры" и "настраиваемые свойства" были внесены улучшения.

| Функциональная область | Описание | Соответствующие объекты |
|:--- |:--- |:--- |
| [События комментариев](../../excel/excel-add-ins-events.md) | Добавляет события для добавления, изменения и удаления в коллекцию комментариев.| [CommentCollection](/javascript/api/excel/excel.commentcollection) |
| [Параметры культуры](../../excel/excel-add-ins-workbooks.md#access-application-culture-settings) даты и времени | Предоставляет доступ к дополнительным параметрам культуры в отношении форматирования даты и времени. | [CultureInfo](/javascript/api/excel/excel.cultureinfo), [NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [Application](/javascript/api/excel/excel.application) |
| Прямые и влияющие | Возвращает диапазоны, используемые для вычисления формулы ячейки.| [Range](/javascript/api/excel/excel.range#getdirectprecedents--) |
| Фильтры сводной таблицы | Применяет управляемые по значению фильтры к полям сводной таблицы. | [PivotField](/javascript/api/excel/excel.pivotfield#applyfilter-filter-), [PivotFilters](/javascript/api/excel/excel.pivotFilters) |
| [Сброс диапазона](../../excel/excel-add-ins-ranges-advanced.md#handle-dynamic-arrays-and-spilling) | Позволяет надстройкам находить диапазоны, связанные с результатами [динамических массивов](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531) . | [Range](/javascript/api/excel/excel.range) |
| [Настраиваемые свойства на уровне листа](../../excel/excel-add-ins-workbooks.md#worksheet-level-custom-properties) | Позволяет задать для настраиваемых свойств область на уровне листа, в дополнение к ограничению области действия на уровне книги. | [Воркшиткустомпроперти](/javascript/api/excel/excel.worksheetcustomproperty), [воркшиткустомпропертиколлектион](/javascript/api/excel/excel.worksheetcustompropertycollection)|

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в наборе обязательных элементов API JavaScript для Excel 1,12. Чтобы просмотреть справочную документацию по API для всех API, поддерживаемых набором обязательных элементов API JavaScript для Excel 1,12 или более ранней версии, обратитесь к разделам [API Excel в наборе требований 1,12](/javascript/api/excel?view=excel-js-1.12&preserve-view=true)

| Класс | Поля | Описание |
|:---|:---|:---|
|[ChartAxisTitle](/javascript/api/excel/excel.chartaxistitle)|[textOrientation](/javascript/api/excel/excel.chartaxistitle#textorientation)|Задает угол, по которому текст будет ориентирован на название оси диаграммы. Значение должно быть целым числом от – 90 до 90 или целым числом 180 для вертикально ориентированного текста.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[Жетдименсионвалуес (Dimension: Excel. Чартсериесдименсион)](/javascript/api/excel/excel.chartseries#getdimensionvalues-dimension-)|Получает значения из одного измерения ряда диаграммы. Это могут быть значения категории или значения данных, в зависимости от указанного измерения и способа сопоставления данных для ряда диаграммы.|
|[Comment](/javascript/api/excel/excel.comment)|[contentType](/javascript/api/excel/excel.comment#contenttype)|Получает тип контента комментария.|
|[комментаддедевентаргс](/javascript/api/excel/excel.commentaddedeventargs)|[комментдетаилс](/javascript/api/excel/excel.commentaddedeventargs#commentdetails)|Получение массива Комментдетаил, содержащего идентификатор комментария и идентификаторы связанных ответов.|
||[source](/javascript/api/excel/excel.commentaddedeventargs#source)|Указывает источник события. Дополнительные сведения см. в статье Excel.EventSource.|
||[type](/javascript/api/excel/excel.commentaddedeventargs#type)|Получает тип события. Дополнительные сведения см. в статье Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.commentaddedeventargs#worksheetid)|Получает идентификатор листа, в котором произошло событие.|
|[комментчанжедевентаргс](/javascript/api/excel/excel.commentchangedeventargs)|[changeType](/javascript/api/excel/excel.commentchangedeventargs#changetype)|Получает тип изменения, представляющий способ запуска события Changed.|
||[комментдетаилс](/javascript/api/excel/excel.commentchangedeventargs#commentdetails)|Получение массива Комментдетаил, содержащего идентификатор комментария и идентификаторы связанных ответов.|
||[source](/javascript/api/excel/excel.commentchangedeventargs#source)|Указывает источник события. Дополнительные сведения см. в статье Excel.EventSource.|
||[type](/javascript/api/excel/excel.commentchangedeventargs#type)|Получает тип события. Дополнительные сведения см. в статье Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.commentchangedeventargs#worksheetid)|Получает идентификатор листа, в котором произошло событие.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[onAdded](/javascript/api/excel/excel.commentcollection#onadded)|Возникает при добавлении комментариев.|
||[onChanged](/javascript/api/excel/excel.commentcollection#onchanged)|Происходит при изменении комментариев или ответов в коллекции комментариев, в том числе при удалении ответов.|
||[onDeleted](/javascript/api/excel/excel.commentcollection#ondeleted)|Возникает при удалении комментариев в коллекции комментариев.|
|[комментделетедевентаргс](/javascript/api/excel/excel.commentdeletedeventargs)|[комментдетаилс](/javascript/api/excel/excel.commentdeletedeventargs#commentdetails)|Получение массива Комментдетаил, содержащего идентификатор комментария и идентификаторы связанных ответов.|
||[source](/javascript/api/excel/excel.commentdeletedeventargs#source)|Указывает источник события. Дополнительные сведения см. в статье Excel.EventSource.|
||[type](/javascript/api/excel/excel.commentdeletedeventargs#type)|Получает тип события. Дополнительные сведения см. в статье Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.commentdeletedeventargs#worksheetid)|Получает идентификатор листа, в котором произошло событие.|
|[комментдетаил](/javascript/api/excel/excel.commentdetail)|[комментид](/javascript/api/excel/excel.commentdetail#commentid)|Представляет идентификатор комментария.|
||[реплидс](/javascript/api/excel/excel.commentdetail#replyids)|Представляет идентификаторы связанных ответов, относящихся к комментарию.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[contentType](/javascript/api/excel/excel.commentreply#contenttype)|Тип контента для ответа.|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[датетимеформат](/javascript/api/excel/excel.cultureinfo#datetimeformat)|Определяет формат отображения даты и времени, соответствующий культуре. Это основано на текущих параметрах языковых параметров системы.|
|[датетимеформатинфо](/javascript/api/excel/excel.datetimeformatinfo)|[датесепаратор](/javascript/api/excel/excel.datetimeformatinfo#dateseparator)|Получает строку, используемую в качестве разделителя даты. Это основано на текущих параметрах системы.|
||[лонгдатепаттерн](/javascript/api/excel/excel.datetimeformatinfo#longdatepattern)|Получает строку формата для длинного значения даты. Это основано на текущих параметрах системы.|
||[лонгтимепаттерн](/javascript/api/excel/excel.datetimeformatinfo#longtimepattern)|Получает строку формата для длинного значения времени. Это основано на текущих параметрах системы.|
||[шортдатепаттерн](/javascript/api/excel/excel.datetimeformatinfo#shortdatepattern)|Получает строку формата для краткого значения даты. Это основано на текущих параметрах системы.|
||[тимесепаратор](/javascript/api/excel/excel.datetimeformatinfo#timeseparator)|Получает строку, используемую в качестве разделителя времени. Это основано на текущих параметрах системы.|
|[пивотдатефилтер](/javascript/api/excel/excel.pivotdatefilter)|[блок](/javascript/api/excel/excel.pivotdatefilter#comparator)|Оператор сравнения — это статическое значение, с которым сравниваются другие значения. Тип сравнения определяется условием.|
||[установлен](/javascript/api/excel/excel.pivotdatefilter#condition)|Задает условие фильтра, которое определяет необходимые критерии фильтрации.|
||[применим](/javascript/api/excel/excel.pivotdatefilter#exclusive)|Если задано значение true, фильтр *исключает* элементы, соответствующие условиям. По умолчанию используется значение false (Filter для включения элементов, соответствующих условиям).|
||[ловербаунд](/javascript/api/excel/excel.pivotdatefilter#lowerbound)|Нижняя граница диапазона `Between` условия фильтра.|
||[уппербаунд](/javascript/api/excel/excel.pivotdatefilter#upperbound)|Верхняя граница диапазона `Between` условия фильтра.|
||[вхоледайс](/javascript/api/excel/excel.pivotdatefilter#wholedays)|`Equals`Условия для, `Before` , `After` , и `Between` условия фильтра указывает, следует ли производить сравнение в течение целых дней.|
|[PivotField](/javascript/api/excel/excel.pivotfield)|[applyFilter (Filter: Excel. PivotFilters)](/javascript/api/excel/excel.pivotfield#applyfilter-filter-)|Задает одно или несколько текущих PivotFilters поля и применяет их к полю.|
||[Клеараллфилтерс ()](/javascript/api/excel/excel.pivotfield#clearallfilters--)|Удаляет все условия из всех фильтров полей. При этом будет удалена любая активная фильтрация для поля.|
||[clearFilter (filterType: Excel. Пивотфилтертипе)](/javascript/api/excel/excel.pivotfield#clearfilter-filtertype-)|Удаляет все существующие критерии из фильтра поля данного типа (если он в настоящее время применяется).|
||[Фильтры ()](/javascript/api/excel/excel.pivotfield#getfilters--)|Получает все фильтры, применяемые в данный момент для поля.|
||[Фильтр (filterType?: Excel. Пивотфилтертипе)](/javascript/api/excel/excel.pivotfield#isfiltered-filtertype-)|Проверяет, применены ли фильтры к полю.|
|[PivotFilters](/javascript/api/excel/excel.pivotfilters)|[датефилтер](/javascript/api/excel/excel.pivotfilters#datefilter)|Применяемый в данный момент фильтр даты PivotField. Значение null, если значение не применяется.|
||[лабелфилтер](/javascript/api/excel/excel.pivotfilters#labelfilter)|Применяемый в данный момент фильтр меток PivotField. Значение null, если значение не применяется.|
||[мануалфилтер](/javascript/api/excel/excel.pivotfilters#manualfilter)|Применяемый в данный момент фильтр, выполняемый в PivotField. Значение null, если значение не применяется.|
||[валуефилтер](/javascript/api/excel/excel.pivotfilters#valuefilter)|Примененный в текущий момент фильтр значений PivotField. Значение null, если значение не применяется.|
|[пивотлабелфилтер](/javascript/api/excel/excel.pivotlabelfilter)|[блок](/javascript/api/excel/excel.pivotlabelfilter#comparator)|Оператор сравнения — это статическое значение, с которым сравниваются другие значения. Тип сравнения определяется условием.|
||[установлен](/javascript/api/excel/excel.pivotlabelfilter#condition)|Задает условие фильтра, которое определяет необходимые критерии фильтрации.|
||[применим](/javascript/api/excel/excel.pivotlabelfilter#exclusive)|Если задано значение true, фильтр *исключает* элементы, соответствующие условиям. По умолчанию используется значение false (Filter для включения элементов, соответствующих условиям).|
||[ловербаунд](/javascript/api/excel/excel.pivotlabelfilter#lowerbound)|Нижняя граница диапазона между условиями фильтра.|
||[substring](/javascript/api/excel/excel.pivotlabelfilter#substring)|Подстрока, используемая для `BeginsWith` `EndsWith` `Contains` условий фильтра и.|
||[уппербаунд](/javascript/api/excel/excel.pivotlabelfilter#upperbound)|Верхняя граница диапазона между условиями фильтра.|
|[пивотмануалфилтер](/javascript/api/excel/excel.pivotmanualfilter)|[selectedItems](/javascript/api/excel/excel.pivotmanualfilter#selecteditems)|Список выбранных элементов, которые необходимо фильтровать вручную. В выбранном поле должны быть существующие и допустимые элементы.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[алловмултиплефилтерсперфиелд](/javascript/api/excel/excel.pivottable#allowmultiplefiltersperfield)|Указывает, разрешена ли в сводной таблице возможность применения нескольких PivotFilters к заданному PivotField в таблице.|
|[пивоттаблескопедколлектион](/javascript/api/excel/excel.pivottablescopedcollection)|[getCount()](/javascript/api/excel/excel.pivottablescopedcollection#getcount--)|Получает количество сводных таблиц в коллекции.|
||[getFirst()](/javascript/api/excel/excel.pivottablescopedcollection#getfirst--)|Получает первую сводную таблицу в коллекции. Сводные таблицы в коллекции сортируются сверху вниз и слева направо, так как первая сводная таблица в коллекции является верхней левой.|
||[getItem(key: string)](/javascript/api/excel/excel.pivottablescopedcollection#getitem-key-)|Получает сводную таблицу по имени.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.pivottablescopedcollection#getitemornullobject-name-)|Получает сводную таблицу по имени. Если сводная таблица не существует, возвращает пустой объект.|
||[items](/javascript/api/excel/excel.pivottablescopedcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[пивотвалуефилтер](/javascript/api/excel/excel.pivotvaluefilter)|[блок](/javascript/api/excel/excel.pivotvaluefilter#comparator)|Оператор сравнения — это статическое значение, с которым сравниваются другие значения. Тип сравнения определяется условием.|
||[установлен](/javascript/api/excel/excel.pivotvaluefilter#condition)|Задает условие фильтра, которое определяет необходимые критерии фильтрации.|
||[применим](/javascript/api/excel/excel.pivotvaluefilter#exclusive)|Если задано значение true, фильтр *исключает* элементы, соответствующие условиям. По умолчанию используется значение false (Filter для включения элементов, соответствующих условиям).|
||[ловербаунд](/javascript/api/excel/excel.pivotvaluefilter#lowerbound)|Нижняя граница диапазона `Between` условия фильтра.|
||[селектионтипе](/javascript/api/excel/excel.pivotvaluefilter#selectiontype)|Указывает, используется ли фильтр для верхних и нижних N элементов, а также для первых и последних N процентов, а также для верхней и нижней N сумм.|
||[threshold](/javascript/api/excel/excel.pivotvaluefilter#threshold)|Пороговое значение "N" элементов, процентов или SUM, фильтруемое для условия фильтра Top/Bottom.|
||[уппербаунд](/javascript/api/excel/excel.pivotvaluefilter#upperbound)|Верхняя граница диапазона `Between` условия фильтра.|
||[value](/javascript/api/excel/excel.pivotvaluefilter#value)|Имя выбранного "значения" в поле, по которому будет осуществляться фильтрация.|
|[Range](/javascript/api/excel/excel.range)|[Жетдиректпрецедентс ()](/javascript/api/excel/excel.range#getdirectprecedents--)|Возвращает объект Воркбукранжеареас, который представляет диапазон, содержащий все прямые и непосредственные ячейки в ячейке на одном листе или на нескольких листах.|
||[PivotTable (Фулликонтаинед?: Boolean)](/javascript/api/excel/excel.range#getpivottables-fullycontained-)|Возвращает ограниченную коллекцию сводных таблиц, которые перекрывают диапазон.|
||[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|Получает объект диапазона, содержащий базовую ячейку для переносимой ячейки. Возвращает ошибку, если применяется к диапазону с несколькими ячейками.|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|Получает объект диапазона, содержащий базовую ячейку для переносимой ячейки.|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|Получает объект range, содержащий диапазон переноса при вызове для базовой ячейки. Возвращает ошибку, если применяется к диапазону с несколькими ячейками.|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|Получает объект range, содержащий диапазон переноса при вызове для базовой ячейки.|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|Указывает, есть ли во всех ячейках граница переноса.|
||[нумберформаткатегориес](/javascript/api/excel/excel.range#numberformatcategories)|Представляет категорию числового формата для каждой ячейки.|
||[саведасаррай](/javascript/api/excel/excel.range#savedasarray)|Указывает, следует ли сохранять все ячейки в виде формулы массива.|
|[ранжеареасколлектион](/javascript/api/excel/excel.rangeareascollection)|[getCount()](/javascript/api/excel/excel.rangeareascollection#getcount--)|Получает число объектов RangeAreas в коллекции.|
||[getItemAt(index: number)](/javascript/api/excel/excel.rangeareascollection#getitemat-index-)|Возвращает объект RangeAreas на основе позиции в коллекции.|
||[items](/javascript/api/excel/excel.rangeareascollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[воркбукранжеареас](/javascript/api/excel/excel.workbookrangeareas)|[Жетранжеареасбишит (Key: строка)](/javascript/api/excel/excel.workbookrangeareas#getrangeareasbysheet-key-)|Возвращает `RangeAreas` объект, основанный на идентификаторе или имени листа в коллекции.|
||[Жетранжеареасорнуллобжектбишит (Key: строка)](/javascript/api/excel/excel.workbookrangeareas#getrangeareasornullobjectbysheet-key-)|Возвращает `RangeAreas` объект, основанный на имени листа или идентификаторе в коллекции. Если лист не существует, возвращает пустой объект.|
||[addresses](/javascript/api/excel/excel.workbookrangeareas#addresses)|Возвращает массив адресов в стиле a1. Значение Address будет содержать имя листа для каждого прямоугольного блока ячеек (например, "Лист1! A1: B4, Лист1! D1: D4 "). Только для чтения.|
||[areas](/javascript/api/excel/excel.workbookrangeareas#areas)|Возвращает `RangeAreasCollection` объект. Каждый `RangeAreas` объект в коллекции представляет один или несколько диапазонов прямоугольников в одном листе.|
||[ячеек](/javascript/api/excel/excel.workbookrangeareas#ranges)|Возвращает диапазоны, составляющие этот объект в `RangeCollection` объекте.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[customProperties](/javascript/api/excel/excel.worksheet#customproperties)|Возвращает коллекцию настраиваемых свойств на уровне листа.|
|[воркшиткустомпроперти](/javascript/api/excel/excel.worksheetcustomproperty)|[delete()](/javascript/api/excel/excel.worksheetcustomproperty#delete--)|Удаляет настраиваемое свойство.|
||[key](/javascript/api/excel/excel.worksheetcustomproperty#key)|Возвращает ключ настраиваемого свойства. В настраиваемых ключах свойств не учитывается регистр. Ключ имеет ограничение в 255 символов (большие значения приведут к возникновению ошибки "InvalidArgument").|
||[value](/javascript/api/excel/excel.worksheetcustomproperty#value)|Получает или задает значение настраиваемого свойства.|
|[воркшиткустомпропертиколлектион](/javascript/api/excel/excel.worksheetcustompropertycollection)|[Add (Key: строка, Value: строка)](/javascript/api/excel/excel.worksheetcustompropertycollection#add-key--value-)|Добавляет новое настраиваемое свойство, которое сопоставляется с предоставленным ключом. При этом существующие настраиваемые свойства перезаписываются с помощью этого раздела.|
||[getCount()](/javascript/api/excel/excel.worksheetcustompropertycollection#getcount--)|Получает количество настраиваемых свойств на этом листе.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitem-key-)|Возвращает объект настраиваемого свойства по ключу, указываемому без учета регистра. Вызывается, если настраиваемое свойство не существует.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitemornullobject-key-)|Возвращает объект настраиваемого свойства по ключу, указываемому без учета регистра. Возвращает нулевой объект, если настраиваемое свойство не существует.|
||[items](/javascript/api/excel/excel.worksheetcustompropertycollection#items)|Получает загруженные дочерние элементы в этой коллекции.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.12&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)
