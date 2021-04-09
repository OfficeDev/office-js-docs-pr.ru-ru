---
title: Набор требований к API JavaScript Excel 1.10
description: Сведения о наборе требований ExcelApi 1.10.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 1bafdd2064166019c5c3f22aa4da1a2d0ec73f08
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/09/2021
ms.locfileid: "51650823"
---
# <a name="whats-new-in-excel-javascript-api-110"></a>Что нового в API JavaScript Excel 1.10

В ExcelApi 1.10 представлены ключевые функции, такие как комментарии, контуры и срезы. Кроме того, добавлена поддержка событий для нажатия и сортировки на уровне таблицы.

| Функциональная область | Описание | Соответствующие объекты |
|:--- |:--- |:--- |
| [Примечания](../../excel/excel-add-ins-comments.md) | Добавление, редактирование и удаление примечаний. | [Comment](/javascript/api/excel/excel.comment), [CommentCollection](/javascript/api/excel/excel.commentcollection) |
| [Контуры](../../excel/excel-add-ins-ranges-group.md) | Групповые строки и столбцы для формирования общих контуров. | [Диапазон](/javascript/api/excel/excel.range), [таблица](/javascript/api/excel/excel.worksheet) |
| [Slicers](../../excel/excel-add-ins-pivottables.md#filter-with-slicers) | Вставка и настройка срезов для таблиц и сводных таблиц. | [Slicer](/javascript/api/excel/excel.slicer) |
| [Дополнительные события таблицы](../../excel/excel-add-ins-events.md) | Щелкните кнопку мыши и отсортировать события в таблице. | [Таблица (События)](/javascript/api/excel/excel.worksheet#events) |

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в API Excel JavaScript, за набором 1.10. Чтобы просмотреть справочную документацию API для всех API, поддерживаемых требованиями API Excel JavaScript, установленных 1.10 или ранее, см. в справке к API Excel в наборе требований [1.10](/javascript/api/excel?view=excel-js-1.10&preserve-view=true)или более ранних .

| Класс | Поля | Описание |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[content](/javascript/api/excel/excel.comment#content)|Содержимое комментария.|
||[delete()](/javascript/api/excel/excel.comment#delete--)|Удаляет комментарий и все подключенные ответы.|
||[getLocation()](/javascript/api/excel/excel.comment#getlocation--)|Получает ячейку, в которой расположен этот комментарий.|
||[authorEmail](/javascript/api/excel/excel.comment#authoremail)|Получает электронную почту автора примечания.|
||[authorName](/javascript/api/excel/excel.comment#authorname)|Получает имя автора примечания.|
||[creationDate](/javascript/api/excel/excel.comment#creationdate)|Получает время создания примечания.|
||[id](/javascript/api/excel/excel.comment#id)|Указывает идентификатор комментария.|
||[replies](/javascript/api/excel/excel.comment#replies)|Представляет коллекцию объектов ответов, связанных с примечанием.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(cellAddress: Range \| string, content: string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentcollection#add-celladdress--content--contenttype-)|Создает новое примечание с указанным содержимым в определенной ячейке.|
||[getCount()](/javascript/api/excel/excel.commentcollection#getcount--)|Получает количество примечаний в коллекции.|
||[getItem(commentId: string)](/javascript/api/excel/excel.commentcollection#getitem-commentid-)|Получает примечание из коллекции на основе его идентификатора.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentcollection#getitemat-index-)|Получает примечание из коллекции на основе его позиции.|
||[getItemByCell(cellAddress: Range \| string)](/javascript/api/excel/excel.commentcollection#getitembycell-celladdress-)|Получает примечание из указанной ячейки.|
||[getItemByReplyId(replyId: string)](/javascript/api/excel/excel.commentcollection#getitembyreplyid-replyid-)|Получает комментарий, к которому подключен данный ответ.|
||[items](/javascript/api/excel/excel.commentcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[content](/javascript/api/excel/excel.commentreply#content)|Содержимое ответа на комментарий.|
||[delete()](/javascript/api/excel/excel.commentreply#delete--)|Удаляет ответ на примечание.|
||[getLocation()](/javascript/api/excel/excel.commentreply#getlocation--)|Получает ячейку, в которой находится ответ на этот комментарий.|
||[getParentComment()](/javascript/api/excel/excel.commentreply#getparentcomment--)|Получает родительский комментарий этого ответа.|
||[authorEmail](/javascript/api/excel/excel.commentreply#authoremail)|Получает электронную почту автора ответа на примечание.|
||[authorName](/javascript/api/excel/excel.commentreply#authorname)|Получает имя автора ответа на примечание.|
||[creationDate](/javascript/api/excel/excel.commentreply#creationdate)|Получает время создания ответа на примечание.|
||[id](/javascript/api/excel/excel.commentreply#id)|Указывает идентификатор ответа на комментарии.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|Создает ответ на примечание.|
||[getCount()](/javascript/api/excel/excel.commentreplycollection#getcount--)|Получает количество ответов на примечания в коллекции.|
||[getItem(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getitem-commentreplyid-)|Возвращает ответ на примечание, определенное по идентификатору.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentreplycollection#getitemat-index-)|Возвращает ответ на примечание на основе его позиции в коллекции.|
||[items](/javascript/api/excel/excel.commentreplycollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[enableFieldList](/javascript/api/excel/excel.pivotlayout#enablefieldlist)|Указывает, можно ли показывать список полей в пользовательском интерфейсе.|
|[PivotTableStyle](/javascript/api/excel/excel.pivottablestyle)|[delete()](/javascript/api/excel/excel.pivottablestyle#delete--)|Удаляет объект PivotTableStyle.|
||[duplicate()](/javascript/api/excel/excel.pivottablestyle#duplicate--)|Создает дубликат объекта PivotTableStyle с копиями всех элементов стиля.|
||[name](/javascript/api/excel/excel.pivottablestyle#name)|Получает имя объекта PivotTableStyle.|
||[readOnly](/javascript/api/excel/excel.pivottablestyle#readonly)|Указывает, предназначен ли объект PivotTableStyle только для чтения.|
|[PivotTableStyleCollection](/javascript/api/excel/excel.pivottablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.pivottablestylecollection#add-name--makeuniquename-)|Создает пустой объект PivotTableStyle с указанным именем.|
||[getCount()](/javascript/api/excel/excel.pivottablestylecollection#getcount--)|Получает количество стилей сводных таблиц в коллекции.|
||[getDefault()](/javascript/api/excel/excel.pivottablestylecollection#getdefault--)|Получает используемый по умолчанию объект PivotTableStyle для области родительского объекта.|
||[getItem(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getitem-name-)|Получает объект PivotTableStyle по имени.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.pivottablestylecollection#getitemornullobject-name-)|Получает объект PivotTableStyle по имени.|
||[items](/javascript/api/excel/excel.pivottablestylecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[setDefault(newDefaultStyle: PivotTableStyle \| string)](/javascript/api/excel/excel.pivottablestylecollection#setdefault-newdefaultstyle-)|Задает объект PivotTableStyle, используемый по умолчанию в области родительского объекта.|
|[Range](/javascript/api/excel/excel.range)|[group(groupOption: Excel.GroupOption)](/javascript/api/excel/excel.range#group-groupoption-)|Группы столбцов и строк для контура.|
||[hideGroupDetails(groupOption: Excel.GroupOption)](/javascript/api/excel/excel.range#hidegroupdetails-groupoption-)|Скрыть сведения о группе строки или столбца.|
||[height](/javascript/api/excel/excel.range#height)|Возвращает расстояние в пунктах (для масштаба 100 %) от верхнего до нижнего края диапазона.|
||[left](/javascript/api/excel/excel.range#left)|Возвращает расстояние в пунктах (для масштаба 100 %) от левого края листа до левого края диапазона.|
||[top](/javascript/api/excel/excel.range#top)|Возвращает расстояние в пунктах для масштаба 100 % от верхнего края листа до верхнего края диапазона.|
||[width](/javascript/api/excel/excel.range#width)|Возвращает расстояние в пунктах (для масштаба 100 %) от левого до правого края диапазона.|
||[showGroupDetails(groupOption: Excel.GroupOption)](/javascript/api/excel/excel.range#showgroupdetails-groupoption-)|Показать сведения о группе строки или столбца.|
||[ungroup(groupOption: Excel.GroupOption)](/javascript/api/excel/excel.range#ungroup-groupoption-)|Разгруппировка столбцов и строк для контура.|
|[Shape](/javascript/api/excel/excel.shape)|[copyTo(destinationSheet?: Worksheet \| string)](/javascript/api/excel/excel.shape#copyto-destinationsheet-)|Копирует и вставляет объект Shape.|
||[placement](/javascript/api/excel/excel.shape#placement)|Представляет способ прикрепления объекта к ячейкам под ним.|
|[Slicer](/javascript/api/excel/excel.slicer)|[caption](/javascript/api/excel/excel.slicer#caption)|Представляет подпись среза.|
||[clearFilters()](/javascript/api/excel/excel.slicer#clearfilters--)|Удаляет все фильтры, примененные к срезу.|
||[delete()](/javascript/api/excel/excel.slicer#delete--)|Удаляет срез.|
||[getSelectedItems()](/javascript/api/excel/excel.slicer#getselecteditems--)|Возвращает массив имен выбранных ключей элементов.|
||[height](/javascript/api/excel/excel.slicer#height)|Представляет высоту среза (в пунктах).|
||[left](/javascript/api/excel/excel.slicer#left)|Представляет расстояние в пунктах от левого края среза до левого края листа.|
||[name](/javascript/api/excel/excel.slicer#name)|Представляет имя среза.|
||[id](/javascript/api/excel/excel.slicer#id)|Представляет уникальный идентификатор среза.|
||[isFilterCleared](/javascript/api/excel/excel.slicer#isfiltercleared)|Значение true, если удалены все фильтры, примененные к срезу.|
||[slicerItems](/javascript/api/excel/excel.slicer#sliceritems)|Представляет коллекцию объектов SlicerItem, которые являются частью среза.|
||[worksheet](/javascript/api/excel/excel.slicer#worksheet)|Представляет лист, содержащий срез.|
||[selectItems(items?: string[])](/javascript/api/excel/excel.slicer#selectitems-items-)|Выбирает элементы среза на основе ключей.|
||[sortBy](/javascript/api/excel/excel.slicer#sortby)|Представляет порядок сортировки элементов в срезе.|
||[style](/javascript/api/excel/excel.slicer#style)|Постоянное значение, представляющее стиль среза.|
||[top](/javascript/api/excel/excel.slicer#top)|Представляет расстояние в пунктах от верхнего края среза до верхнего края листа.|
||[width](/javascript/api/excel/excel.slicer#width)|Представляет ширину среза (в пунктах).|
|[SlicerCollection](/javascript/api/excel/excel.slicercollection)|[add(slicerSource: string \| PivotTable \| Table, sourceField: string \| PivotField \| number \| TableColumn, slicerDestination?: string \| Worksheet)](/javascript/api/excel/excel.slicercollection#add-slicersource--sourcefield--slicerdestination-)|Добавляет новый срез в книгу.|
||[getCount()](/javascript/api/excel/excel.slicercollection#getcount--)|Возвращает количество срезов в коллекции.|
||[getItem(key: string)](/javascript/api/excel/excel.slicercollection#getitem-key-)|Получает объект slicer по его имени или ИД.|
||[getItemAt(index: number)](/javascript/api/excel/excel.slicercollection#getitemat-index-)|Получает срез на основе его позиции в коллекции.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.slicercollection#getitemornullobject-key-)|Получает срез с его именем или id.|
||[items](/javascript/api/excel/excel.slicercollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[SlicerItem](/javascript/api/excel/excel.sliceritem)|[isSelected](/javascript/api/excel/excel.sliceritem#isselected)|Значение true, если выбран элемент среза.|
||[hasData](/javascript/api/excel/excel.sliceritem#hasdata)|Значение true, если элемент среза содержит данные. |
||[key](/javascript/api/excel/excel.sliceritem#key)|Представляет уникальное значение, соответствующее элементу среза.|
||[name](/javascript/api/excel/excel.sliceritem#name)|Представляет название, отображаемую в пользовательском интерфейсе.|
|[SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection)|[getCount()](/javascript/api/excel/excel.sliceritemcollection#getcount--)|Возвращает количество элементов в срезе.|
||[getItem(key: string)](/javascript/api/excel/excel.sliceritemcollection#getitem-key-)|Получает объект элемента среза по ключу или имени.|
||[getItemAt(index: number)](/javascript/api/excel/excel.sliceritemcollection#getitemat-index-)|Получает элемент среза на основе его позиции в коллекции.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.sliceritemcollection#getitemornullobject-key-)|Получает элемент среза по ключу или имени.|
||[items](/javascript/api/excel/excel.sliceritemcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[SlicerStyle](/javascript/api/excel/excel.slicerstyle)|[delete()](/javascript/api/excel/excel.slicerstyle#delete--)|Удаляет объект SlicerStyle.|
||[duplicate()](/javascript/api/excel/excel.slicerstyle#duplicate--)|Создает дубликат объекта SlicerStyle с копиями всех элементов стиля.|
||[name](/javascript/api/excel/excel.slicerstyle#name)|Получает имя объекта SlicerStyle.|
||[readOnly](/javascript/api/excel/excel.slicerstyle#readonly)|Указывает, предназначен ли объект SlicerStyle только для чтения.|
|[SlicerStyleCollection](/javascript/api/excel/excel.slicerstylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.slicerstylecollection#add-name--makeuniquename-)|Создает пустой объект SlicerStyle с указанным именем.|
||[getCount()](/javascript/api/excel/excel.slicerstylecollection#getcount--)|Получает количество стилей срезов в коллекции.|
||[getDefault()](/javascript/api/excel/excel.slicerstylecollection#getdefault--)|Получает используемый по умолчанию объект SlicerStyle для области родительского объекта.|
||[getItem(name: string)](/javascript/api/excel/excel.slicerstylecollection#getitem-name-)|Получает объект SlicerStyle по имени.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.slicerstylecollection#getitemornullobject-name-)|Получает объект SlicerStyle по имени.|
||[items](/javascript/api/excel/excel.slicerstylecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[setDefault(newDefaultStyle: SlicerStyle \| string)](/javascript/api/excel/excel.slicerstylecollection#setdefault-newdefaultstyle-)|Задает объект SlicerStyle, используемый по умолчанию в области родительского объекта.|
|[TableStyle](/javascript/api/excel/excel.tablestyle)|[delete()](/javascript/api/excel/excel.tablestyle#delete--)|Удаляет объект TableStyle.|
||[duplicate()](/javascript/api/excel/excel.tablestyle#duplicate--)|Создает дубликат объекта TableStyle с копиями всех элементов стиля.|
||[name](/javascript/api/excel/excel.tablestyle#name)|Получает имя объекта TableStyle.|
||[readOnly](/javascript/api/excel/excel.tablestyle#readonly)|Указывает, предназначен ли объект TableStyle только для чтения.|
|[TableStyleCollection](/javascript/api/excel/excel.tablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.tablestylecollection#add-name--makeuniquename-)|Создает пустой объект TableStyle с указанным именем.|
||[getCount()](/javascript/api/excel/excel.tablestylecollection#getcount--)|Получает количество стилей таблиц в коллекции.|
||[getDefault()](/javascript/api/excel/excel.tablestylecollection#getdefault--)|Получает используемый по умолчанию объект TableStyle для области родительского объекта.|
||[getItem(name: string)](/javascript/api/excel/excel.tablestylecollection#getitem-name-)|Получает объект TableStyle по имени.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.tablestylecollection#getitemornullobject-name-)|Получает объект TableStyle по имени.|
||[items](/javascript/api/excel/excel.tablestylecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[setDefault(newDefaultStyle: TableStyle \| string)](/javascript/api/excel/excel.tablestylecollection#setdefault-newdefaultstyle-)|Задает объект TableStyle, используемый по умолчанию в области родительского объекта.|
|[TimelineStyle](/javascript/api/excel/excel.timelinestyle)|[delete()](/javascript/api/excel/excel.timelinestyle#delete--)|Удаляет объект TableStyle.|
||[duplicate()](/javascript/api/excel/excel.timelinestyle#duplicate--)|Создает дубликат объекта TimelineStyle с копиями всех элементов стиля.|
||[name](/javascript/api/excel/excel.timelinestyle#name)|Получает имя объекта TimelineStyle.|
||[readOnly](/javascript/api/excel/excel.timelinestyle#readonly)|Указывает, предназначен ли объект TimelineStyle только для чтения.|
|[TimelineStyleCollection](/javascript/api/excel/excel.timelinestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.timelinestylecollection#add-name--makeuniquename-)|Создает пустой объект TimelineStyle с указанным именем.|
||[getCount()](/javascript/api/excel/excel.timelinestylecollection#getcount--)|Получает количество стилей временной шкалы в коллекции.|
||[getDefault()](/javascript/api/excel/excel.timelinestylecollection#getdefault--)|Получает используемый по умолчанию объект TimelineStyle для области родительского объекта.|
||[getItem(name: string)](/javascript/api/excel/excel.timelinestylecollection#getitem-name-)|Получает объект TimelineStyle по имени.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.timelinestylecollection#getitemornullobject-name-)|Получает объект TimelineStyle по имени.|
||[items](/javascript/api/excel/excel.timelinestylecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[setDefault(newDefaultStyle: TimelineStyle \| string)](/javascript/api/excel/excel.timelinestylecollection#setdefault-newdefaultstyle-)|Задает объект TimelineStyle, используемый по умолчанию в области родительского объекта.|
|[Workbook](/javascript/api/excel/excel.workbook)|[getActiveSlicer()](/javascript/api/excel/excel.workbook#getactiveslicer--)|Получает текущий активный срез в книге.|
||[getActiveSlicerOrNullObject()](/javascript/api/excel/excel.workbook#getactiveslicerornullobject--)|Получает текущий активный срез в книге.|
||[comments](/javascript/api/excel/excel.workbook#comments)|Представляет коллекцию примечаний, связанных с книгой.|
||[pivotTableStyles](/javascript/api/excel/excel.workbook#pivottablestyles)|Представляет коллекцию объектов PivotTableStyles, связанных с книгой.|
||[slicerStyles](/javascript/api/excel/excel.workbook#slicerstyles)|Представляет коллекцию объектов SlicerStyles, связанных с книгой.|
||[slicers](/javascript/api/excel/excel.workbook#slicers)|Представляет коллекцию срезов, связанных с книгой.|
||[tableStyles](/javascript/api/excel/excel.workbook#tablestyles)|Представляет коллекцию объектов TableStyles, связанных с книгой.|
||[timelineStyles](/javascript/api/excel/excel.workbook#timelinestyles)|Представляет коллекцию объектов TimelineStyles, связанных с книгой.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[comments](/javascript/api/excel/excel.worksheet#comments)|Возвращает коллекцию всех объектов Comments на листе.|
||[onColumnSorted](/javascript/api/excel/excel.worksheet#oncolumnsorted)|Возникает при сортировке одного или нескольких столбцов.|
||[onRowSorted](/javascript/api/excel/excel.worksheet#onrowsorted)|Возникает при сортировке одной или нескольких строк.|
||[onSingleClicked](/javascript/api/excel/excel.worksheet#onsingleclicked)|Происходит, когда в таблице происходит действие левого щелчка или прослушиваемого действия.|
||[slicers](/javascript/api/excel/excel.worksheet#slicers)|Возвращает коллекцию срезов, которые являются частью таблицы.|
||[showOutlineLevels (rowLevels: number, columnLevels: number)](/javascript/api/excel/excel.worksheet#showoutlinelevels-rowlevels--columnlevels-)|Показывает группы строк или столбцов по уровням контура.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onColumnSorted](/javascript/api/excel/excel.worksheetcollection#oncolumnsorted)|Возникает при сортировке одного или нескольких столбцов.|
||[onRowSorted](/javascript/api/excel/excel.worksheetcollection#onrowsorted)|Возникает при сортировке одной или нескольких строк.|
||[onSingleClicked](/javascript/api/excel/excel.worksheetcollection#onsingleclicked)|Происходит, когда в коллекции таблицы происходит операция нажатием левой кнопкой мыши или нажатием на нее.|
|[WorksheetColumnSortedEventArgs](/javascript/api/excel/excel.worksheetcolumnsortedeventargs)|[address](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#address)|Получает адрес диапазона, представляющий отсортированные области конкретного листа.|
||[источник](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#source)|Получает источник события.|
||[type](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#worksheetid)|Получает идентификатор листа, в котором выполнена сортировка.|
|[WorksheetRowSortedEventArgs](/javascript/api/excel/excel.worksheetrowsortedeventargs)|[address](/javascript/api/excel/excel.worksheetrowsortedeventargs#address)|Получает адрес диапазона, представляющий отсортированные области конкретного листа.|
||[источник](/javascript/api/excel/excel.worksheetrowsortedeventargs#source)|Получает источник события.|
||[type](/javascript/api/excel/excel.worksheetrowsortedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowsortedeventargs#worksheetid)|Получает идентификатор листа, в котором выполнена сортировка.|
|[WorksheetSingleClickedEventArgs](/javascript/api/excel/excel.worksheetsingleclickedeventargs)|[address](/javascript/api/excel/excel.worksheetsingleclickedeventargs#address)|Получает адрес, представляющий ячейку, по которой выполнен щелчок левой кнопкой мыши или нажатие, для определенного листа.|
||[offsetX](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsetx)|Расстояние в точках от слева нажатой или прослушиваемой точки до левого (или правого для языков справа налево) границы сетки ячейки слева нажатой или прослушиваемой.|
||[offsetY](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsety)|Расстояние в пунктах от точки щелчка левой кнопкой мыши или нажатия до верхнего края сетки ячейки, по которой выполнен щелчок левой кнопкой мыши или нажатие.|
||[type](/javascript/api/excel/excel.worksheetsingleclickedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetsingleclickedeventargs#worksheetid)|Получает идентификатор листа, в котором по ячейке выполнен щелчок левой кнопкой мыши или нажатие.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.10&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)