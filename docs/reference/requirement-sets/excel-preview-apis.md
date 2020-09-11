---
title: Предварительные версии API JavaScript для Excel
description: Сведения о предстоящих API JavaScript для Excel
ms.date: 06/29/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: d1701ad393b96e33f0007bfcb5609c93c13608a2
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/11/2020
ms.locfileid: "47430767"
---
# <a name="excel-javascript-preview-apis"></a>Предварительные версии API JavaScript для Excel

Новые API JavaScript для Excel сначала выпускаются в "предварительной версии", а затем становятся частью определенного нумерованного набора обязательных элементов после выполнения достаточного тестирования и получения отзывов пользователей.

В первой таблице представлен краткий обзор API, а в последующей таблице приведен подробный список.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| Функциональная область | Описание | Соответствующие объекты |
|:--- |:--- |:--- |
| [Параметры культуры](../../excel/excel-add-ins-workbooks.md#access-application-culture-settings) даты и времени | Предоставляет доступ к дополнительным параметрам культуры в отношении форматирования даты и времени. | [CultureInfo](/javascript/api/excel/excel.cultureinfo), [NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [Application](/javascript/api/excel/excel.application) |
| [Вставка книги](../../excel/excel-add-ins-workbooks.md#insert-a-copy-of-an-existing-workbook-into-the-current-one-preview) | Вставка одной книги в другую.  | [Workbook](/javascript/api/excel/excel.worksheetcollection) |
| Фильтры сводной таблицы | Применяет управляемые по значению фильтры к полям сводной таблицы. | [PivotField](/javascript/api/excel/excel.pivotfield#applyfilter-filter-), [PivotFilters](/javascript/api/excel/excel.pivotFilters) |
|Сброс диапазона | Позволяет надстройкам находить диапазоны, связанные с результатами [динамических массивов](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531) . | [Range](/javascript/api/excel/excel.range) |

## <a name="api-list"></a>Список API

В следующей таблице перечислены API JavaScript для Excel, находящиеся в предварительной версии. Чтобы просмотреть полный список всех интерфейсов API JavaScript для Excel (включая предварительные API и ранее выпущенные API), ознакомьтесь со статьями [все API JavaScript для Excel](/javascript/api/excel?view=excel-js-preview&preserve-view=true).

| Класс | Поля | Описание |
|:---|:---|:---|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[Жетдименсионвалуес (Dimension: Excel. Чартсериесдименсион)](/javascript/api/excel/excel.chartseries#getdimensionvalues-dimension-)|Получает значения из одного измерения ряда диаграммы. Это могут быть значения категории или значения данных, в зависимости от указанного измерения и способа сопоставления данных для ряда диаграммы.|
|[Comment](/javascript/api/excel/excel.comment)|[contentType](/javascript/api/excel/excel.comment#contenttype)|Получает тип контента комментария.|
|[комментаддедевентаргс](/javascript/api/excel/excel.commentaddedeventargs)|[комментдетаилс](/javascript/api/excel/excel.commentaddedeventargs#commentdetails)|Получает `CommentDetail` массив, СОДЕРЖАЩИЙ идентификатор комментария и идентификаторы связанных с ним ответов.|
||[source](/javascript/api/excel/excel.commentaddedeventargs#source)|Указывает источник события. `Excel.EventSource`Для получения дополнительных сведений см.|
||[type](/javascript/api/excel/excel.commentaddedeventargs#type)|Получает тип события. `Excel.EventType`Для получения дополнительных сведений см.|
||[worksheetId](/javascript/api/excel/excel.commentaddedeventargs#worksheetid)|Получает идентификатор листа, в котором произошло событие.|
|[комментчанжедевентаргс](/javascript/api/excel/excel.commentchangedeventargs)|[changeType](/javascript/api/excel/excel.commentchangedeventargs#changetype)|Получает тип изменения, представляющий способ запуска события Changed.|
||[комментдетаилс](/javascript/api/excel/excel.commentchangedeventargs#commentdetails)|Получает `CommentDetail` массив, СОДЕРЖАЩИЙ идентификатор комментария и идентификаторы связанных с ним ответов.|
||[source](/javascript/api/excel/excel.commentchangedeventargs#source)|Указывает источник события. `Excel.EventSource`Для получения дополнительных сведений см.|
||[type](/javascript/api/excel/excel.commentchangedeventargs#type)|Получает тип события. `Excel.EventType`Для получения дополнительных сведений см.|
||[worksheetId](/javascript/api/excel/excel.commentchangedeventargs#worksheetid)|Получает идентификатор листа, в котором произошло событие.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[onAdded](/javascript/api/excel/excel.commentcollection#onadded)|Возникает при добавлении комментариев.|
||[onChanged](/javascript/api/excel/excel.commentcollection#onchanged)|Происходит при изменении комментариев или ответов в коллекции комментариев, в том числе при удалении ответов.|
||[onDeleted](/javascript/api/excel/excel.commentcollection#ondeleted)|Возникает при удалении комментариев в коллекции комментариев.|
|[комментделетедевентаргс](/javascript/api/excel/excel.commentdeletedeventargs)|[комментдетаилс](/javascript/api/excel/excel.commentdeletedeventargs#commentdetails)|Получает `CommentDetail` массив, СОДЕРЖАЩИЙ идентификатор комментария и идентификаторы связанных с ним ответов.|
||[source](/javascript/api/excel/excel.commentdeletedeventargs#source)|Указывает источник события. `Excel.EventSource`Для получения дополнительных сведений см.|
||[type](/javascript/api/excel/excel.commentdeletedeventargs#type)|Получает тип события. `Excel.EventType`Для получения дополнительных сведений см.|
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
|[намедшитвиев](/javascript/api/excel/excel.namedsheetview)|[activate()](/javascript/api/excel/excel.namedsheetview#activate--)|Активирует это представление листа. Это эквивалентно использованию команды "переключиться" в пользовательском интерфейсе Excel.|
||[delete()](/javascript/api/excel/excel.namedsheetview#delete--)|Удаляет представление листа из листа.|
||[дублировать (имя?: строка)](/javascript/api/excel/excel.namedsheetview#duplicate-name-)|Создает копию этого представления листа.|
||[name](/javascript/api/excel/excel.namedsheetview#name)|Получает или задает имя представления листа.|
|[намедшитвиевколлектион](/javascript/api/excel/excel.namedsheetviewcollection)|[add(name: string)](/javascript/api/excel/excel.namedsheetviewcollection#add-name-)|Создает новое представление листа с заданным именем.|
||[Ентертемпорари ()](/javascript/api/excel/excel.namedsheetviewcollection#entertemporary--)|Создает и активирует новое временное представление листа.|
||[Exit ()](/javascript/api/excel/excel.namedsheetviewcollection#exit--)|Выполняет выход из текущего активного представления листа.|
||[onactive ()](/javascript/api/excel/excel.namedsheetviewcollection#getactive--)|Получает текущее активное представление листа.|
||[getCount()](/javascript/api/excel/excel.namedsheetviewcollection#getcount--)|Получает количество просмотров листа на этом листе.|
||[getItem(key: string)](/javascript/api/excel/excel.namedsheetviewcollection#getitem-key-)|Возвращает представление листа с использованием его имени.|
||[getItemAt(index: number)](/javascript/api/excel/excel.namedsheetviewcollection#getitemat-index-)|Получает представление листа по его индексу в коллекции.|
||[items](/javascript/api/excel/excel.namedsheetviewcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
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
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|Получает уникальную ячейку в сводной таблице на основе иерархии данных и элементов строк и столбцов соответствующих иерархий. Возвращаемая ячейка находится на пересечении указанной строки и столбца, содержащего данные из заданной иерархии. Этот метод является обратным вызову методов getPivotItems и getDataHierarchy для конкретной ячейки.|
||[пивотстиле](/javascript/api/excel/excel.pivotlayout#pivotstyle)|Стиль, примененный к сводной таблице.|
||[Сетстиле (Style: string \| пивоттаблестиле \| буилтинпивоттаблестиле)](/javascript/api/excel/excel.pivotlayout#setstyle-style-)|Задает стиль, применяемый к сводной таблице.|
|[пивотмануалфилтер](/javascript/api/excel/excel.pivotmanualfilter)|[selectedItems](/javascript/api/excel/excel.pivotmanualfilter#selecteditems)|Список выбранных элементов, которые необходимо фильтровать вручную. В выбранном поле должны быть существующие и допустимые элементы.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[алловмултиплефилтерсперфиелд](/javascript/api/excel/excel.pivottable#allowmultiplefiltersperfield)|Указывает, разрешена ли в сводной таблице возможность применения нескольких PivotFilters к заданному PivotField в таблице.|
|[пивотвалуефилтер](/javascript/api/excel/excel.pivotvaluefilter)|[блок](/javascript/api/excel/excel.pivotvaluefilter#comparator)|Оператор сравнения — это статическое значение, с которым сравниваются другие значения. Тип сравнения определяется условием.|
||[установлен](/javascript/api/excel/excel.pivotvaluefilter#condition)|Задает условие фильтра, которое определяет необходимые критерии фильтрации.|
||[применим](/javascript/api/excel/excel.pivotvaluefilter#exclusive)|Если задано значение true, фильтр *исключает* элементы, соответствующие условиям. По умолчанию используется значение false (Filter для включения элементов, соответствующих условиям).|
||[ловербаунд](/javascript/api/excel/excel.pivotvaluefilter#lowerbound)|Нижняя граница диапазона `Between` условия фильтра.|
||[селектионтипе](/javascript/api/excel/excel.pivotvaluefilter#selectiontype)|Указывает, используется ли фильтр для верхних и нижних N элементов, а также для первых и последних N процентов, а также для верхней и нижней N сумм.|
||[threshold](/javascript/api/excel/excel.pivotvaluefilter#threshold)|Пороговое значение "N" элементов, процентов или SUM, фильтруемое для условия фильтра Top/Bottom.|
||[уппербаунд](/javascript/api/excel/excel.pivotvaluefilter#upperbound)|Верхняя граница диапазона `Between` условия фильтра.|
||[value](/javascript/api/excel/excel.pivotvaluefilter#value)|Имя выбранного "значения" в поле, по которому будет осуществляться фильтрация.|
|[Range](/javascript/api/excel/excel.range)|[Жетдиректпрецедентс ()](/javascript/api/excel/excel.range#getdirectprecedents--)|Возвращает `WorkbookRangeAreas` объект, представляющий диапазон, который содержит все прямые и непосредственные ячейки в ячейке на одном листе или на нескольких листах.|
||[Жетмержедареас ()](/javascript/api/excel/excel.range#getmergedareas--)|Возвращает объект RangeAreas, представляющий Объединенные области в этом диапазоне. Обратите внимание, что если число Объединенных областей в этом диапазоне превышает 512, API не будет возвращать результат.|
||[Влияющие ()](/javascript/api/excel/excel.range#getprecedents--)|Возвращает `WorkbookRangeAreas` объект, представляющий диапазон, содержащий все влияющие ячейки на одном листе или на нескольких листах.|
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
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|Создает изображение SVG (масштабируемая векторная графика) из строки XML и добавляет его на лист. Возвращает объект Shape, представляющий новое изображение.|
|[Slicer](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|Представляет имя среза, используемое в формуле.|
||[слицерстиле](/javascript/api/excel/excel.slicer#slicerstyle)|Стиль, примененный к срезу.|
||[Сетстиле (Style: string \| пивоттаблестиле \| буилтинслицерстиле)](/javascript/api/excel/excel.slicer#setstyle-style-)|Задает стиль, применяемый к срезу.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|Изменяет таблицу для использования стиля таблицы по умолчанию.|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|Возникает, если применен фильтр к указанной таблице.|
||[tableStyle](/javascript/api/excel/excel.table#tablestyle)|Стиль, примененный к таблице.|
||[Сетстиле (Style: string \| пивоттаблестиле \| буилтинтаблестиле)](/javascript/api/excel/excel.table#setstyle-style-)|Задает стиль, применяемый к срезу.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|Возникает, если применен фильтр к любой таблице в книге или листе.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|Получает идентификатор таблицы, в которой применяется фильтр.|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|Получает тип события. Дополнительные сведения см. в статье Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|Получает идентификатор листа, содержащего таблицу.|
|[Workbook](/javascript/api/excel/excel.workbook)|[шовпивотфиелдлист](/javascript/api/excel/excel.workbook#showpivotfieldlist)|Указывает, отображается ли область списка полей сводной таблицы на уровне книги.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|Значение true, если в книге используется система дат 1904.|
|[воркбукранжеареас](/javascript/api/excel/excel.workbookrangeareas)|[Жетранжеареасбишит (Key: строка)](/javascript/api/excel/excel.workbookrangeareas#getrangeareasbysheet-key-)|Возвращает `RangeAreas` объект, основанный на идентификаторе или имени листа в коллекции.|
||[Жетранжеареасорнуллобжектбишит (Key: строка)](/javascript/api/excel/excel.workbookrangeareas#getrangeareasornullobjectbysheet-key-)|Возвращает `RangeAreas` объект, основанный на имени листа или идентификаторе в коллекции. Если лист не существует, возвращает пустой объект.|
||[addresses](/javascript/api/excel/excel.workbookrangeareas#addresses)|Возвращает массив адресов в стиле a1. Значение Address будет содержать имя листа для каждого прямоугольного блока ячеек (например, "Лист1! A1: B4, Лист1! D1: D4 "). Только для чтения.|
||[areas](/javascript/api/excel/excel.workbookrangeareas#areas)|Возвращает объект Ранжеареасколлектион, каждый RangeAreas в коллекции представляет один или несколько диапазонов прямоугольников в одном листе.|
||[ячеек](/javascript/api/excel/excel.workbookrangeareas#ranges)|Возвращает коллекцию диапазонов, состоящих из этого объекта.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[customProperties](/javascript/api/excel/excel.worksheet#customproperties)|Возвращает коллекцию настраиваемых свойств на уровне листа.|
||[намедшитвиевс](/javascript/api/excel/excel.worksheet#namedsheetviews)|Возвращает коллекцию представлений листа, присутствующих на листе.|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|Возникает, если применен фильтр к указанному листу.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Вставляет указанные листы книги в текущую книгу.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|Возникает при применении любого фильтра листа в книге.|
|[воркшиткустомпроперти](/javascript/api/excel/excel.worksheetcustomproperty)|[delete()](/javascript/api/excel/excel.worksheetcustomproperty#delete--)|Удаляет настраиваемое свойство.|
||[key](/javascript/api/excel/excel.worksheetcustomproperty#key)|Возвращает ключ настраиваемого свойства. В настраиваемых ключах свойств не учитывается регистр. Ключ имеет ограничение в 255 символов (большие значения приведут к возникновению ошибки "InvalidArgument").|
||[value](/javascript/api/excel/excel.worksheetcustomproperty#value)|Получает или задает значение настраиваемого свойства.|
|[воркшиткустомпропертиколлектион](/javascript/api/excel/excel.worksheetcustompropertycollection)|[Add (Key: строка, Value: строка)](/javascript/api/excel/excel.worksheetcustompropertycollection#add-key--value-)|Добавляет новое настраиваемое свойство, которое сопоставляется с предоставленным ключом. При этом существующие настраиваемые свойства перезаписываются с помощью этого раздела.|
||[getCount()](/javascript/api/excel/excel.worksheetcustompropertycollection#getcount--)|Получает количество настраиваемых свойств на этом листе.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitem-key-)|Возвращает объект настраиваемого свойства по ключу, указываемому без учета регистра. Вызывается, если настраиваемое свойство не существует.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitemornullobject-key-)|Возвращает объект настраиваемого свойства по ключу, указываемому без учета регистра. Возвращает нулевой объект, если настраиваемое свойство не существует.|
||[items](/javascript/api/excel/excel.worksheetcustompropertycollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Получает тип события. Дополнительные сведения см. в статье Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|Получает идентификатор листа, в котором применяется фильтр.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-preview&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](./excel-api-requirement-sets.md)
