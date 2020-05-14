---
title: Предварительные версии API JavaScript для Excel
description: Сведения о предстоящих API JavaScript для Excel
ms.date: 05/11/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: a867a0c80158e9c1effa40ba5a0aa83608ba57de
ms.sourcegitcommit: 682d18c9149b1153f9c38d28e2a90384e6a261dc
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/13/2020
ms.locfileid: "44218007"
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

В следующей таблице перечислены API JavaScript для Excel, находящиеся в предварительной версии. Чтобы просмотреть полный список всех интерфейсов API JavaScript для Excel (включая предварительные API и ранее выпущенные API), ознакомьтесь со статьями [все API JavaScript для Excel](/javascript/api/excel?view=excel-js-preview).

| Класс | Поля | Описание |
|:---|:---|:---|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[Жетдименсионвалуес (Dimension: Excel. Чартсериесдименсион)](/javascript/api/excel/excel.chartseries#getdimensionvalues-dimension-)|Получает значения из одного измерения ряда диаграммы. Это могут быть значения категории или значения данных, в зависимости от указанного измерения и способа сопоставления данных для ряда диаграммы.|
|[Comment](/javascript/api/excel/excel.comment)|[contentType](/javascript/api/excel/excel.comment#contenttype)|Получает тип контента комментария.|
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
|[Range](/javascript/api/excel/excel.range)|[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|Получает объект диапазона, содержащий базовую ячейку для переносимой ячейки. Возвращает ошибку, если применяется к диапазону с несколькими ячейками.|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|Получает объект диапазона, содержащий базовую ячейку для переносимой ячейки.|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|Получает объект range, содержащий диапазон переноса при вызове для базовой ячейки. Возвращает ошибку, если применяется к диапазону с несколькими ячейками.|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|Получает объект range, содержащий диапазон переноса при вызове для базовой ячейки.|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|Указывает, есть ли во всех ячейках граница переноса.|
||[нумберформаткатегориес](/javascript/api/excel/excel.range#numberformatcategories)|Представляет категорию числового формата для каждой ячейки.|
||[саведасаррай](/javascript/api/excel/excel.range#savedasarray)|Указывает, следует ли сохранять все ячейки в виде формулы массива.|
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
|[Worksheet](/javascript/api/excel/excel.worksheet)|[customProperties](/javascript/api/excel/excel.worksheet#customproperties)|Возвращает коллекцию настраиваемых свойств на уровне листа.|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|Возникает, если применен фильтр к указанному листу.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Вставляет указанные листы книги в текущую книгу.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|Возникает при применении любого фильтра листа в книге.|
|[воркшиткустомпроперти](/javascript/api/excel/excel.worksheetcustomproperty)|[delete()](/javascript/api/excel/excel.worksheetcustomproperty#delete--)|Удаляет настраиваемое свойство.|
||[key](/javascript/api/excel/excel.worksheetcustomproperty#key)|Возвращает ключ настраиваемого свойства. В настраиваемых ключах свойств не учитывается регистр.|
||[value](/javascript/api/excel/excel.worksheetcustomproperty#value)|Получает или задает значение настраиваемого свойства.|
|[воркшиткустомпропертиколлектион](/javascript/api/excel/excel.worksheetcustompropertycollection)|[Add (Key: строка, Value: строка)](/javascript/api/excel/excel.worksheetcustompropertycollection#add-key--value-)|Добавляет новое настраиваемое свойство, которое сопоставляется с предоставленным ключом. При этом существующие настраиваемые свойства перезаписываются с помощью этого раздела.|
||[getCount()](/javascript/api/excel/excel.worksheetcustompropertycollection#getcount--)|Получает количество настраиваемых свойств на этом листе.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitem-key-)|Возвращает объект настраиваемого свойства по ключу, указываемому без учета регистра. Вызывается, если настраиваемое свойство не существует.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitemornullobject-key-)|Возвращает объект настраиваемого свойства по ключу, указываемому без учета регистра. Возвращает нулевой объект, если настраиваемое свойство не существует.|
||[items](/javascript/api/excel/excel.worksheetcustompropertycollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Получает тип события. Дополнительные сведения см. в статье Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|Получает идентификатор листа, в котором применяется фильтр.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-preview)
- [Наборы обязательных элементов API JavaScript для Excel](./excel-api-requirement-sets.md)
