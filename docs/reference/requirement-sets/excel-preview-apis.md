---
title: Предварительные версии API JavaScript для Excel
description: Сведения о предстоящих API JavaScript для Excel
ms.date: 01/02/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 7b751cfa550813fe3cb57f0d90d40e5632baf953
ms.sourcegitcommit: abe8188684b55710261c69e206de83d3a6bd2ed3
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/08/2020
ms.locfileid: "40969547"
---
# <a name="excel-javascript-preview-apis"></a>Предварительные версии API JavaScript для Excel

Новые API JavaScript для Excel сначала выпускаются в "предварительной версии", а затем становятся частью определенного нумерованного набора обязательных элементов после выполнения достаточного тестирования и получения отзывов пользователей.

В первой таблице представлен краткий обзор API, а в последующей таблице приведен подробный список.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| Функциональная область | Описание | Соответствующие объекты |
|:--- |:--- |:--- |
| Параметры культуры | Получает региональные параметры системы для книги, например форматирование чисел. | [CultureInfo](/javascript/api/excel/excel.cultureinfo), [NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo) [Application](/javascript/api/excel/excel.application) |
| [Вставка книги](../../excel/excel-add-ins-workbooks.md#insert-a-copy-of-an-existing-workbook-into-the-current-one-preview) | Вставка одной книги в другую.  | [Workbook](/javascript/api/excel/excel.worksheetcollection) |
| [Сохранение](../../excel/excel-add-ins-workbooks.md#save-the-workbook-preview) и [закрытие](../../excel/excel-add-ins-workbooks.md#close-the-workbook-preview) рабочей книги | Сохранение и закрытие книг.  | [Workbook](/javascript/api/excel/excel.workbook) |

## <a name="api-list"></a>Список API

В следующей таблице перечислены API JavaScript для Excel, находящиеся в предварительной версии. Чтобы просмотреть полный список всех интерфейсов API JavaScript для Excel (включая предварительные API и ранее выпущенные API), ознакомьтесь со статьями [все API JavaScript для Excel](/javascript/api/excel?view=excel-js-preview).

| Класс | Поля | Описание |
|:---|:---|:---|
|[Application](/javascript/api/excel/excel.application)|[cultureInfo](/javascript/api/excel/excel.application#cultureinfo)|Предоставляет сведения, основанные на текущих параметрах языковых параметров системы. Сюда входят имена культур, форматирование чисел и другие параметры, зависящие от культуры.|
||[деЦималсепаратор](/javascript/api/excel/excel.application#decimalseparator)|Получает строку, используемую в качестве десятичного разделителя для числовых значений. Это основано на локальных параметрах Excel.|
||[саусандссепаратор](/javascript/api/excel/excel.application#thousandsseparator)|Получает строку, используемую для разделения групп цифр слева от десятичного разделителя для числовых значений. Это основано на локальных параметрах Excel.|
||[усесистемсепараторс](/javascript/api/excel/excel.application#usesystemseparators)|Указывает, включены ли системные разделители Microsoft Excel.|
|[ChartSeries](/javascript/api/excel/excel.chartseries)|[Жетдименсионвалуес (Dimension: Excel. Чартсериесдименсион)](/javascript/api/excel/excel.chartseries#getdimensionvalues-dimension-)|Получает значения из одного измерения ряда диаграммы. Это могут быть значения категории или значения данных, в зависимости от указанного измерения и способа сопоставления данных для ряда диаграммы.|
|[Comment](/javascript/api/excel/excel.comment)|[определяем](/javascript/api/excel/excel.comment#resolved)|Получает или задает состояние потока комментариев. Значение "true" означает, что поток комментариев находится в состоянии "разрешено".|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[определяем](/javascript/api/excel/excel.commentreply#resolved)|Получает или задает состояние ответа на комментарий. Значение "true" означает, что ответ комментария находится в состоянии "разрешено".|
|[CultureInfo](/javascript/api/excel/excel.cultureinfo)|[name](/javascript/api/excel/excel.cultureinfo#name)|Получает имя языка и региональных параметров в формате languagecode2-Country/regioncode2 (например, "zh-CN" или "en-US"). Это основано на текущих параметрах системы.|
||[numberFormatInfo](/javascript/api/excel/excel.cultureinfo#numberformatinfo)|Определяет формат отображения чисел, соответствующий культуре. Это основано на текущих параметрах языковых параметров системы.|
|[NumberFormatInfo](/javascript/api/excel/excel.numberformatinfo)|[нумбердеЦималсепаратор](/javascript/api/excel/excel.numberformatinfo#numberdecimalseparator)|Получает строку, используемую в качестве десятичного разделителя для числовых значений. Это основано на текущих параметрах системы.|
||[нумберграупсепаратор](/javascript/api/excel/excel.numberformatinfo#numbergroupseparator)|Получает строку, используемую для разделения групп цифр слева от десятичного разделителя для числовых значений. Это основано на текущих параметрах системы.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|Получает уникальную ячейку в сводной таблице на основе иерархии данных и элементов строк и столбцов соответствующих иерархий. Возвращаемая ячейка находится на пересечении указанной строки и столбца, содержащего данные из заданной иерархии. Этот метод является обратным вызову методов getPivotItems и getDataHierarchy для конкретной ячейки.|
|[Range](/javascript/api/excel/excel.range)|[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|Получает объект диапазона, содержащий базовую ячейку для переносимой ячейки. Возвращает ошибку, если применяется к диапазону с несколькими ячейками. Только для чтения.|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|Получает объект диапазона, содержащий базовую ячейку для переносимой ячейки. Только для чтения.|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|Получает объект range, содержащий диапазон переноса при вызове для базовой ячейки. Возвращает ошибку, если применяется к диапазону с несколькими ячейками. Только для чтения.|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|Получает объект range, содержащий диапазон переноса при вызове для базовой ячейки. Только для чтения.|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|Указывает, есть ли во всех ячейках граница переноса.|
||[саведасаррай](/javascript/api/excel/excel.range#savedasarray)|Указывает, следует ли сохранять все ячейки в виде формулы массива.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|Создает изображение SVG (масштабируемая векторная графика) из строки XML и добавляет его на лист. Возвращает объект Shape, представляющий новое изображение.|
|[Slicer](/javascript/api/excel/excel.slicer)|[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|Представляет имя среза, используемое в формуле.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|Изменяет таблицу для использования стиля таблицы по умолчанию.|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|Возникает, если применен фильтр к указанной таблице.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|Возникает, если применен фильтр к любой таблице в книге или листе.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|Представляет идентификатор таблицы, в которой применяется фильтр.|
||[тип](/javascript/api/excel/excel.tablefilteredeventargs#type)|Представляет тип события. Дополнительные сведения см. в статье Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|Представляет идентификатор листа, содержащего таблицу.|
|[Workbook](/javascript/api/excel/excel.workbook)|[close(closeBehavior?: Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close-closebehavior-)|Закрывает текущую книгу.|
||[save(saveBehavior?: Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save-savebehavior-)|Сохраняет текущую книгу.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|Значение true, если в книге используется система дат 1904.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[customProperties](/javascript/api/excel/excel.worksheet#customproperties)|Возвращает коллекцию настраиваемых свойств на уровне листа.|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|Возникает, если применен фильтр к указанному листу.|
||[онровхидденчанжед](/javascript/api/excel/excel.worksheet#onrowhiddenchanged)|Происходит при изменении скрытого состояния одной или нескольких строк на определенном листе.|
|[воркшиткалкулатедевентаргс](/javascript/api/excel/excel.worksheetcalculatedeventargs)|[address](/javascript/api/excel/excel.worksheetcalculatedeventargs#address)|Адрес диапазона, который выполнил вычисление.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Вставляет указанные листы книги в текущую книгу.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|Возникает при применении любого фильтра листа в книге.|
||[онровхидденчанжед](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)|Происходит при изменении скрытого состояния одной или нескольких строк на определенном листе.|
|[воркшиткустомпроперти](/javascript/api/excel/excel.worksheetcustomproperty)|[key](/javascript/api/excel/excel.worksheetcustomproperty#key)|Возвращает ключ настраиваемого свойства. Только для чтения.|
||[value](/javascript/api/excel/excel.worksheetcustomproperty#value)|Получает значение настраиваемого свойства. Только для чтения.|
|[воркшиткустомпропертиколлектион](/javascript/api/excel/excel.worksheetcustompropertycollection)|[getCount()](/javascript/api/excel/excel.worksheetcustompropertycollection#getcount--)|Получает количество настраиваемых свойств на этом листе.|
||[getItem(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitem-key-)|Возвращает объект настраиваемого свойства по ключу, указываемому без учета регистра. Вызывается, если настраиваемое свойство не существует.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcustompropertycollection#getitemornullobject-key-)|Возвращает объект настраиваемого свойства по ключу, указываемому без учета регистра. Возвращает нулевой объект, если настраиваемое свойство не существует.|
||[items](/javascript/api/excel/excel.worksheetcustompropertycollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[тип](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Представляет тип события. Дополнительные сведения см. в статье Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|Представляет идентификатор листа, в котором применен фильтр.|
|[воркшитровхидденчанжедевентаргс](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[address](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#address)|Получает адрес диапазона, представляющий измененную область конкретного листа.|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#changetype)|Получает тип изменения, которое представляет способ запуска события. Для `Excel.RowHiddenChangeType` получения дополнительных сведений см.|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#source)|Получает источник события. Дополнительные сведения см. в статье Excel.EventSource.|
||[type](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#type)|Получает тип события. Дополнительные сведения см. в статье Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#worksheetid)|Получает идентификатор листа, в котором изменены данные.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-preview)
- [Наборы обязательных элементов API JavaScript для Excel](./excel-api-requirement-sets.md)
