---
title: Набор обязательных элементов API JavaScript для Excel 1,10
description: Сведения о наборе требований ExcelApi 1,10.
ms.date: 10/22/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 8c6ecf04312e37e710111757b82240d3b4cfa751
ms.sourcegitcommit: ed2a98b6fb5b432fa99c6cefa5ce52965dc25759
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/16/2020
ms.locfileid: "47819786"
---
# <a name="whats-new-in-excel-javascript-api-110"></a>Новые возможности API JavaScript для Excel 1,10

В ExcelApi 1,10 представлены основные функции, такие как комментирование, структуры и срезы. Кроме того, добавлена поддержка событий для нажатия и сортировки на уровне листа.

| Функциональная область | Описание | Соответствующие объекты |
|:--- |:--- |:--- |
| [Примечания](../../excel/excel-add-ins-comments.md) | Добавление, редактирование и удаление примечаний. | [Comment](/javascript/api/excel/excel.comment), [CommentCollection](/javascript/api/excel/excel.commentcollection) |
| [Описываются](../../excel/excel-add-ins-ranges-advanced.md#group-data-for-an-outline) | Сгруппируйте строки и столбцы, чтобы сформировать свертываемые структуры. | [Диапазон](/javascript/api/excel/excel.range), [лист](/javascript/api/excel/excel.worksheet) |
| [Slicers](../../excel/excel-add-ins-pivottables.md#slicers) | Вставка и настройка срезов для таблиц и сводных таблиц. | [Slicer](/javascript/api/excel/excel.slicer) |
| [Дополнительные события листа](../../excel/excel-add-ins-events.md) | Прослушивать события Click и Sort на листе. | [Лист (события)](/javascript/api/excel/excel.worksheet#events) |

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в наборе обязательных элементов API JavaScript для Excel 1,10. Чтобы просмотреть справочную документацию по API для всех API, поддерживаемых набором обязательных элементов API JavaScript для Excel 1,10 или более ранней версии, обратитесь к разделам [API Excel в наборе требований 1,10](/javascript/api/excel?view=excel-js-1.10&preserve-view=true)

| Класс | Поля | Описание |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[content](/javascript/api/excel/excel.comment#content)|Получает или задает содержимое примечания. Строка является обычным текстом.|
||[delete()](/javascript/api/excel/excel.comment#delete--)|Удаляет комментарий и все подключенные ответы.|
||[getLocation()](/javascript/api/excel/excel.comment#getlocation--)|Получает ячейку, в которой находится этот комментарий.|
||[authorEmail](/javascript/api/excel/excel.comment#authoremail)|Получает электронную почту автора примечания.|
||[authorName](/javascript/api/excel/excel.comment#authorname)|Получает имя автора примечания.|
||[creationDate](/javascript/api/excel/excel.comment#creationdate)|Получает время создания примечания. Возвращает значение null, если примечание было преобразовано из заметки, так как у примечания нет даты создания.|
||[id](/javascript/api/excel/excel.comment#id)|Представляет идентификатор примечания. Только для чтения.|
||[replies](/javascript/api/excel/excel.comment#replies)|Представляет коллекцию объектов ответов, связанных с примечанием. Только для чтения.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[Add (Целладдресс: \| строка Range, Content: комментричконтент \| String, ContentType?: Excel. ContentType)](/javascript/api/excel/excel.commentcollection#add-celladdress--content--contenttype-)|Создает новое примечание с указанным содержимым в определенной ячейке. `InvalidArgument`Если указанный диапазон превышает одну ячейку, возникает ошибка.|
||[getCount()](/javascript/api/excel/excel.commentcollection#getcount--)|Получает количество примечаний в коллекции.|
||[getItem(commentId: string)](/javascript/api/excel/excel.commentcollection#getitem-commentid-)|Получает примечание из коллекции на основе его идентификатора. Только для чтения.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentcollection#getitemat-index-)|Получает примечание из коллекции на основе его позиции.|
||[getItemByCell(cellAddress: Range \| string)](/javascript/api/excel/excel.commentcollection#getitembycell-celladdress-)|Получает примечание из указанной ячейки.|
||[getItemByReplyId(replyId: string)](/javascript/api/excel/excel.commentcollection#getitembyreplyid-replyid-)|Получает комментарий, к которому подключен данный ответ.|
||[items](/javascript/api/excel/excel.commentcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[content](/javascript/api/excel/excel.commentreply#content)|Получает или задает содержимое ответа на примечание. Строка является обычным текстом.|
||[delete()](/javascript/api/excel/excel.commentreply#delete--)|Удаляет ответ на примечание.|
||[getLocation()](/javascript/api/excel/excel.commentreply#getlocation--)|Получает ячейку, в которой находится этот ответ на комментарий.|
||[getParentComment()](/javascript/api/excel/excel.commentreply#getparentcomment--)|Получает родительский комментарий для этого ответа.|
||[authorEmail](/javascript/api/excel/excel.commentreply#authoremail)|Получает электронную почту автора ответа на примечание.|
||[authorName](/javascript/api/excel/excel.commentreply#authorname)|Получает имя автора ответа на примечание.|
||[creationDate](/javascript/api/excel/excel.commentreply#creationdate)|Получает время создания ответа на примечание.|
||[id](/javascript/api/excel/excel.commentreply#id)|Представляет идентификатор ответа на примечание. Только для чтения.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[Добавить (контент: \| строка комментричконтент, ContentType?: Excel. ContentType)](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|Создает ответ на примечание.|
||[getCount()](/javascript/api/excel/excel.commentreplycollection#getcount--)|Получает количество ответов на примечания в коллекции.|
||[getItem(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getitem-commentreplyid-)|Возвращает ответ на примечание, определенное по идентификатору. Только для чтения.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentreplycollection#getitemat-index-)|Возвращает ответ на примечание на основе его позиции в коллекции.|
||[items](/javascript/api/excel/excel.commentreplycollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[комментричконтент](/javascript/api/excel/excel.commentrichcontent)||[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[enableFieldList](/javascript/api/excel/excel.pivotlayout#enablefieldlist)|Указывает, может ли список полей отображаться в пользовательском интерфейсе.|
|[PivotTableStyle](/javascript/api/excel/excel.pivottablestyle)|[delete()](/javascript/api/excel/excel.pivottablestyle#delete--)|Удаляет объект PivotTableStyle.|
||[duplicate()](/javascript/api/excel/excel.pivottablestyle#duplicate--)|Создает дубликат объекта PivotTableStyle с копиями всех элементов стиля.|
||[name](/javascript/api/excel/excel.pivottablestyle#name)|Получает имя объекта PivotTableStyle.|
||[readOnly](/javascript/api/excel/excel.pivottablestyle#readonly)|Указывает, доступен ли объект Пивоттаблестиле только для чтения. Только для чтения.|
|[PivotTableStyleCollection](/javascript/api/excel/excel.pivottablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.pivottablestylecollection#add-name--makeuniquename-)|Создает пустой объект PivotTableStyle с указанным именем.|
||[getCount()](/javascript/api/excel/excel.pivottablestylecollection#getcount--)|Получает количество стилей сводных таблиц в коллекции.|
||[getDefault()](/javascript/api/excel/excel.pivottablestylecollection#getdefault--)|Получает используемый по умолчанию объект PivotTableStyle для области родительского объекта.|
||[getItem(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getitem-name-)|Получает объект PivotTableStyle по имени.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.pivottablestylecollection#getitemornullobject-name-)|Получает объект PivotTableStyle по имени. Если объект PivotTableStyle не существует, возвращает пустой объект.|
||[items](/javascript/api/excel/excel.pivottablestylecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[setDefault(newDefaultStyle: PivotTableStyle \| string)](/javascript/api/excel/excel.pivottablestylecollection#setdefault-newdefaultstyle-)|Задает объект PivotTableStyle, используемый по умолчанию в области родительского объекта.|
|[Range](/javascript/api/excel/excel.range)|[Group (Граупоптион: Excel. Граупоптион)](/javascript/api/excel/excel.range#group-groupoption-)|Группирует столбцы и строки для структуры.|
||[Хидеграупдетаилс (Граупоптион: Excel. Граупоптион)](/javascript/api/excel/excel.range#hidegroupdetails-groupoption-)|Скрытие сведений о группе строк или столбцов.|
||[height](/javascript/api/excel/excel.range#height)|Возвращает расстояние в пунктах (для масштаба 100 %) от верхнего до нижнего края диапазона. Только для чтения.|
||[left](/javascript/api/excel/excel.range#left)|Возвращает расстояние в пунктах (для масштаба 100 %) от левого края листа до левого края диапазона. Только для чтения.|
||[top](/javascript/api/excel/excel.range#top)|Возвращает расстояние в пунктах для масштаба 100 % от верхнего края листа до верхнего края диапазона. Только для чтения.|
||[width](/javascript/api/excel/excel.range#width)|Возвращает расстояние в пунктах (для масштаба 100 %) от левого до правого края диапазона. Только для чтения.|
||[Шовграупдетаилс (Граупоптион: Excel. Граупоптион)](/javascript/api/excel/excel.range#showgroupdetails-groupoption-)|Отображение сведений о группе строк или столбцов.|
||[Разгруппировать (Граупоптион: Excel. Граупоптион)](/javascript/api/excel/excel.range#ungroup-groupoption-)|Разгруппирование столбцов и строк для структуры.|
|[Shape](/javascript/api/excel/excel.shape)|[copyTo(destinationSheet?: Worksheet \| string)](/javascript/api/excel/excel.shape#copyto-destinationsheet-)|Копирует и вставляет объект Shape.|
||[placement](/javascript/api/excel/excel.shape#placement)|Представляет способ прикрепления объекта к ячейкам под ним.|
|[Slicer](/javascript/api/excel/excel.slicer)|[caption](/javascript/api/excel/excel.slicer#caption)|Представляет подпись среза.|
||[clearFilters()](/javascript/api/excel/excel.slicer#clearfilters--)|Удаляет все фильтры, примененные к срезу.|
||[delete()](/javascript/api/excel/excel.slicer#delete--)|Удаляет срез.|
||[getSelectedItems()](/javascript/api/excel/excel.slicer#getselecteditems--)|Возвращает массив имен выбранных ключей элементов. Только для чтения.|
||[height](/javascript/api/excel/excel.slicer#height)|Представляет высоту среза (в пунктах).|
||[left](/javascript/api/excel/excel.slicer#left)|Представляет расстояние в пунктах от левого края среза до левого края листа.|
||[name](/javascript/api/excel/excel.slicer#name)|Представляет имя среза.|
||[id](/javascript/api/excel/excel.slicer#id)|Представляет уникальный идентификатор среза. Только для чтения.|
||[isFilterCleared](/javascript/api/excel/excel.slicer#isfiltercleared)|Значение true, если удалены все фильтры, примененные к срезу.|
||[slicerItems](/javascript/api/excel/excel.slicer#sliceritems)|Представляет коллекцию объектов SlicerItem, которые являются частью среза. Только для чтения.|
||[worksheet](/javascript/api/excel/excel.slicer#worksheet)|Представляет лист, содержащий срез. Только для чтения.|
||[selectItems(items?: string[])](/javascript/api/excel/excel.slicer#selectitems-items-)|Выбирает элементы срезов на основе их ключей. Предыдущие выбранные элементы очищаются.|
||[sortBy](/javascript/api/excel/excel.slicer#sortby)|Представляет порядок сортировки элементов в срезе. Возможные значения: "Датасаурцеордер", "Ascending", "Descending".|
||[style](/javascript/api/excel/excel.slicer#style)|Постоянное значение, представляющее стиль среза. Возможные значения: "SlicerStyleLight1", "SlicerStyleLight6", "TableStyleOther1", "TableStyleOther2", "SlicerStyleDark1" и "SlicerStyleDark6". Также можно указать настраиваемый пользовательский стиль, имеющийся в книге.|
||[top](/javascript/api/excel/excel.slicer#top)|Представляет расстояние в пунктах от верхнего края среза до верхнего края листа.|
||[width](/javascript/api/excel/excel.slicer#width)|Представляет ширину среза (в пунктах).|
|[SlicerCollection](/javascript/api/excel/excel.slicercollection)|[add(slicerSource: string \| PivotTable \| Table, sourceField: string \| PivotField \| number \| TableColumn, slicerDestination?: string \| Worksheet)](/javascript/api/excel/excel.slicercollection#add-slicersource--sourcefield--slicerdestination-)|Добавляет новый срез в книгу.|
||[getCount()](/javascript/api/excel/excel.slicercollection#getcount--)|Возвращает количество срезов в коллекции.|
||[getItem(key: string)](/javascript/api/excel/excel.slicercollection#getitem-key-)|Получает объект slicer по его имени или ИД.|
||[getItemAt(index: number)](/javascript/api/excel/excel.slicercollection#getitemat-index-)|Получает срез на основе его позиции в коллекции.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.slicercollection#getitemornullobject-key-)|Получает срез по его имени или ИД. Если срез не существует, возвращает пустой объект.|
||[items](/javascript/api/excel/excel.slicercollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[SlicerItem](/javascript/api/excel/excel.sliceritem)|[isSelected](/javascript/api/excel/excel.sliceritem#isselected)|Значение true, если выбран элемент среза.|
||[hasData](/javascript/api/excel/excel.sliceritem#hasdata)|Значение true, если элемент среза содержит данные. |
||[key](/javascript/api/excel/excel.sliceritem#key)|Представляет уникальное значение, соответствующее элементу среза.|
||[name](/javascript/api/excel/excel.sliceritem#name)|Представляет заголовок, отображаемый в пользовательском интерфейсе.|
|[SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection)|[getCount()](/javascript/api/excel/excel.sliceritemcollection#getcount--)|Возвращает количество элементов в срезе.|
||[getItem(key: string)](/javascript/api/excel/excel.sliceritemcollection#getitem-key-)|Получает объект элемента среза по ключу или имени.|
||[getItemAt(index: number)](/javascript/api/excel/excel.sliceritemcollection#getitemat-index-)|Получает элемент среза на основе его позиции в коллекции.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.sliceritemcollection#getitemornullobject-key-)|Получает элемент среза по ключу или имени. Если элемент среза не существует, возвращает пустой объект.|
||[items](/javascript/api/excel/excel.sliceritemcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[SlicerStyle](/javascript/api/excel/excel.slicerstyle)|[delete()](/javascript/api/excel/excel.slicerstyle#delete--)|Удаляет объект SlicerStyle.|
||[duplicate()](/javascript/api/excel/excel.slicerstyle#duplicate--)|Создает дубликат объекта SlicerStyle с копиями всех элементов стиля.|
||[name](/javascript/api/excel/excel.slicerstyle#name)|Получает имя объекта SlicerStyle.|
||[readOnly](/javascript/api/excel/excel.slicerstyle#readonly)|Указывает, доступен ли объект Слицерстиле только для чтения. Только для чтения.|
|[SlicerStyleCollection](/javascript/api/excel/excel.slicerstylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.slicerstylecollection#add-name--makeuniquename-)|Создает пустой объект SlicerStyle с указанным именем.|
||[getCount()](/javascript/api/excel/excel.slicerstylecollection#getcount--)|Получает количество стилей срезов в коллекции.|
||[getDefault()](/javascript/api/excel/excel.slicerstylecollection#getdefault--)|Получает используемый по умолчанию объект SlicerStyle для области родительского объекта.|
||[getItem(name: string)](/javascript/api/excel/excel.slicerstylecollection#getitem-name-)|Получает объект SlicerStyle по имени.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.slicerstylecollection#getitemornullobject-name-)|Получает объект SlicerStyle по имени. Если объект SlicerStyle не существует, возвращает пустой объект.|
||[items](/javascript/api/excel/excel.slicerstylecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[setDefault(newDefaultStyle: SlicerStyle \| string)](/javascript/api/excel/excel.slicerstylecollection#setdefault-newdefaultstyle-)|Задает объект SlicerStyle, используемый по умолчанию в области родительского объекта.|
|[TableStyle](/javascript/api/excel/excel.tablestyle)|[delete()](/javascript/api/excel/excel.tablestyle#delete--)|Удаляет объект TableStyle.|
||[duplicate()](/javascript/api/excel/excel.tablestyle#duplicate--)|Создает дубликат объекта TableStyle с копиями всех элементов стиля.|
||[name](/javascript/api/excel/excel.tablestyle#name)|Получает имя объекта TableStyle.|
||[readOnly](/javascript/api/excel/excel.tablestyle#readonly)|Указывает, доступен ли объект TableStyle только для чтения. Только для чтения.|
|[TableStyleCollection](/javascript/api/excel/excel.tablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.tablestylecollection#add-name--makeuniquename-)|Создает пустой объект TableStyle с указанным именем.|
||[getCount()](/javascript/api/excel/excel.tablestylecollection#getcount--)|Получает количество стилей таблиц в коллекции.|
||[getDefault()](/javascript/api/excel/excel.tablestylecollection#getdefault--)|Получает используемый по умолчанию объект TableStyle для области родительского объекта.|
||[getItem(name: string)](/javascript/api/excel/excel.tablestylecollection#getitem-name-)|Получает объект TableStyle по имени.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.tablestylecollection#getitemornullobject-name-)|Получает объект TableStyle по имени. Если объект TableStyle не существует, возвращает пустой объект.|
||[items](/javascript/api/excel/excel.tablestylecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[setDefault(newDefaultStyle: TableStyle \| string)](/javascript/api/excel/excel.tablestylecollection#setdefault-newdefaultstyle-)|Задает объект TableStyle, используемый по умолчанию в области родительского объекта.|
|[TimelineStyle](/javascript/api/excel/excel.timelinestyle)|[delete()](/javascript/api/excel/excel.timelinestyle#delete--)|Удаляет объект TableStyle.|
||[duplicate()](/javascript/api/excel/excel.timelinestyle#duplicate--)|Создает дубликат объекта TimelineStyle с копиями всех элементов стиля.|
||[name](/javascript/api/excel/excel.timelinestyle#name)|Получает имя объекта TimelineStyle.|
||[readOnly](/javascript/api/excel/excel.timelinestyle#readonly)|Указывает, доступен ли объект Тимелинестиле только для чтения. Только для чтения.|
|[TimelineStyleCollection](/javascript/api/excel/excel.timelinestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.timelinestylecollection#add-name--makeuniquename-)|Создает пустой объект TimelineStyle с указанным именем.|
||[getCount()](/javascript/api/excel/excel.timelinestylecollection#getcount--)|Получает количество стилей временной шкалы в коллекции.|
||[getDefault()](/javascript/api/excel/excel.timelinestylecollection#getdefault--)|Получает используемый по умолчанию объект TimelineStyle для области родительского объекта.|
||[getItem(name: string)](/javascript/api/excel/excel.timelinestylecollection#getitem-name-)|Получает объект TimelineStyle по имени.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.timelinestylecollection#getitemornullobject-name-)|Получает объект TimelineStyle по имени. Если объект TimelineStyle не существует, возвращает пустой объект.|
||[items](/javascript/api/excel/excel.timelinestylecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[setDefault(newDefaultStyle: TimelineStyle \| string)](/javascript/api/excel/excel.timelinestylecollection#setdefault-newdefaultstyle-)|Задает объект TimelineStyle, используемый по умолчанию в области родительского объекта.|
|[Workbook](/javascript/api/excel/excel.workbook)|[getActiveSlicer()](/javascript/api/excel/excel.workbook#getactiveslicer--)|Получает текущий активный срез в книге. Если активного среза нет, `ItemNotFound` создается исключение.|
||[getActiveSlicerOrNullObject()](/javascript/api/excel/excel.workbook#getactiveslicerornullobject--)|Получает текущий активный срез в книге. Если активный срез отсутствует, возвращается пустой объект.|
||[comments](/javascript/api/excel/excel.workbook#comments)|Представляет коллекцию примечаний, связанных с книгой. Только для чтения.|
||[pivotTableStyles](/javascript/api/excel/excel.workbook#pivottablestyles)|Представляет коллекцию объектов PivotTableStyles, связанных с книгой. Только для чтения.|
||[slicerStyles](/javascript/api/excel/excel.workbook#slicerstyles)|Представляет коллекцию объектов SlicerStyles, связанных с книгой. Только для чтения.|
||[slicers](/javascript/api/excel/excel.workbook#slicers)|Представляет коллекцию срезов, связанных с книгой. Только для чтения.|
||[tableStyles](/javascript/api/excel/excel.workbook#tablestyles)|Представляет коллекцию объектов TableStyles, связанных с книгой. Только для чтения.|
||[timelineStyles](/javascript/api/excel/excel.workbook#timelinestyles)|Представляет коллекцию объектов TimelineStyles, связанных с книгой. Только для чтения.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[comments](/javascript/api/excel/excel.worksheet#comments)|Возвращает коллекцию всех объектов Comments на листе. Только для чтения.|
||[onColumnSorted](/javascript/api/excel/excel.worksheet#oncolumnsorted)|Возникает при сортировке одного или нескольких столбцов. Происходит в результате операции сортировки слева направо.|
||[onRowSorted](/javascript/api/excel/excel.worksheet#onrowsorted)|Возникает при сортировке одной или нескольких строк. Происходит в результате операции сортировки сверху вниз.|
||[onSingleClicked](/javascript/api/excel/excel.worksheet#onsingleclicked)|Происходит, когда на листе происходит действие, нажатие которого выполняется влево. Это событие не будет инициировано при нажатии следующих условий:|
||[slicers](/javascript/api/excel/excel.worksheet#slicers)|Возвращает коллекцию срезов, которые входят в состав рабочего листа. Только для чтения.|
||[Шоваутлинелевелс (Ровлевелс: число, Колумнлевелс: число)](/javascript/api/excel/excel.worksheet#showoutlinelevels-rowlevels--columnlevels-)|Отображает группы строк или столбцов по их уровням структуры.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onColumnSorted](/javascript/api/excel/excel.worksheetcollection#oncolumnsorted)|Возникает при сортировке одного или нескольких столбцов. Происходит в результате операции сортировки слева направо.|
||[onRowSorted](/javascript/api/excel/excel.worksheetcollection#onrowsorted)|Возникает при сортировке одной или нескольких строк. Происходит в результате операции сортировки сверху вниз.|
||[onSingleClicked](/javascript/api/excel/excel.worksheetcollection#onsingleclicked)|Возникает, когда в коллекции листа происходит операция с нажатием и нажатием левой кнопкой мыши. Это событие не будет инициировано при нажатии следующих условий:|
|[WorksheetColumnSortedEventArgs](/javascript/api/excel/excel.worksheetcolumnsortedeventargs)|[address](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#address)|Получает адрес диапазона, представляющий отсортированные области конкретного листа. Возвращаются только столбцы, измененные в результате операции сортировки.|
||[source](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#source)|Получает источник события. Дополнительные сведения см. в статье Excel.EventSource.|
||[type](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#type)|Получает тип события. Дополнительные сведения см. в статье Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#worksheetid)|Получает идентификатор листа, в котором выполнена сортировка.|
|[WorksheetRowSortedEventArgs](/javascript/api/excel/excel.worksheetrowsortedeventargs)|[address](/javascript/api/excel/excel.worksheetrowsortedeventargs#address)|Получает адрес диапазона, представляющий отсортированные области конкретного листа. Возвращаются только строки, измененные в результате операции сортировки.|
||[source](/javascript/api/excel/excel.worksheetrowsortedeventargs#source)|Получает источник события. Дополнительные сведения см. в статье Excel.EventSource.|
||[type](/javascript/api/excel/excel.worksheetrowsortedeventargs#type)|Получает тип события. Дополнительные сведения см. в статье Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowsortedeventargs#worksheetid)|Получает идентификатор листа, в котором выполнена сортировка.|
|[WorksheetSingleClickedEventArgs](/javascript/api/excel/excel.worksheetsingleclickedeventargs)|[address](/javascript/api/excel/excel.worksheetsingleclickedeventargs#address)|Получает адрес, представляющий ячейку, по которой выполнен щелчок левой кнопкой мыши или нажатие, для определенного листа.|
||[offsetX](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsetx)|Расстояние (в пунктах) от левой щелчка/касания до левого (или правого для языков с письмом справа налево) границы линии сетки для ячейки с левой щелчком мыши.|
||[offsetY](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsety)|Расстояние в пунктах от точки щелчка левой кнопкой мыши или нажатия до верхнего края сетки ячейки, по которой выполнен щелчок левой кнопкой мыши или нажатие.|
||[type](/javascript/api/excel/excel.worksheetsingleclickedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetsingleclickedeventargs#worksheetid)|Получает идентификатор листа, в котором по ячейке выполнен щелчок левой кнопкой мыши или нажатие.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.10&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)