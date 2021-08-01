---
title: Excel Набор API JavaScript 1.10
description: Сведения о наборе требований ExcelApi 1.10.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 7b620bb76f758bc2574e8bd99d2c45d3d4bfae39
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671226"
---
# <a name="whats-new-in-excel-javascript-api-110"></a>Новые возможности в Excel API JavaScript 1.10

В ExcelApi 1.10 представлены ключевые функции, такие как комментарии, контуры и срезы. Кроме того, добавлена поддержка событий для нажатия и сортировки на уровне таблицы.

| Функциональная область | Описание | Соответствующие объекты |
|:--- |:--- |:--- |
| [Примечания](../../excel/excel-add-ins-comments.md) | Добавление, редактирование и удаление примечаний. | [Comment](/javascript/api/excel/excel.comment), [CommentCollection](/javascript/api/excel/excel.commentcollection) |
| [Контуры](../../excel/excel-add-ins-ranges-group.md) | Групповые строки и столбцы для формирования общих контуров. | [Диапазон](/javascript/api/excel/excel.range), [таблица](/javascript/api/excel/excel.worksheet) |
| [Slicers](../../excel/excel-add-ins-pivottables.md#filter-with-slicers) | Вставка и настройка срезов для таблиц и сводных таблиц. | [Slicer](/javascript/api/excel/excel.slicer) |
| [Дополнительные события таблицы](../../excel/excel-add-ins-events.md) | Щелкните кнопку мыши и отсортировать события в таблице. | [Таблица (События)](/javascript/api/excel/excel.worksheet#events) |

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в Excel API JavaScript, за набором 1.10. Чтобы просмотреть справочную документацию API для всех API, поддерживаемых Excel API JavaScript, установленного 1.10 или ранее, см. в Excel API в наборе требований [1.10](/javascript/api/excel?view=excel-js-1.10&preserve-view=true)или ранее .

| Класс | Поля | Описание |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[content](/javascript/api/excel/excel.comment#content)|Содержимое комментария.|
||[delete()](/javascript/api/excel/excel.comment#delete__)|Удаляет комментарий и все подключенные ответы.|
||[getLocation()](/javascript/api/excel/excel.comment#getLocation__)|Получает ячейку, в которой расположен этот комментарий.|
||[authorEmail](/javascript/api/excel/excel.comment#authorEmail)|Получает электронную почту автора примечания.|
||[authorName](/javascript/api/excel/excel.comment#authorName)|Получает имя автора примечания.|
||[creationDate](/javascript/api/excel/excel.comment#creationDate)|Получает время создания примечания.|
||[id](/javascript/api/excel/excel.comment#id)|Указывает идентификатор комментария.|
||[replies](/javascript/api/excel/excel.comment#replies)|Представляет коллекцию объектов ответов, связанных с примечанием.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(cellAddress: Range \| string, content: string, contentType?: Excel. ContentType)](/javascript/api/excel/excel.commentcollection#add_cellAddress__content__contentType_)|Создает новое примечание с указанным содержимым в определенной ячейке.|
||[getCount()](/javascript/api/excel/excel.commentcollection#getCount__)|Получает количество примечаний в коллекции.|
||[getItem(commentId: string)](/javascript/api/excel/excel.commentcollection#getItem_commentId_)|Получает примечание из коллекции на основе его идентификатора.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentcollection#getItemAt_index_)|Получает примечание из коллекции на основе его позиции.|
||[getItemByCell(cellAddress: Range \| string)](/javascript/api/excel/excel.commentcollection#getItemByCell_cellAddress_)|Получает примечание из указанной ячейки.|
||[getItemByReplyId(replyId: string)](/javascript/api/excel/excel.commentcollection#getItemByReplyId_replyId_)|Получает комментарий, к которому подключен данный ответ.|
||[items](/javascript/api/excel/excel.commentcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[content](/javascript/api/excel/excel.commentreply#content)|Содержимое ответа на комментарий.|
||[delete()](/javascript/api/excel/excel.commentreply#delete__)|Удаляет ответ на примечание.|
||[getLocation()](/javascript/api/excel/excel.commentreply#getLocation__)|Получает ячейку, в которой находится ответ на этот комментарий.|
||[getParentComment()](/javascript/api/excel/excel.commentreply#getParentComment__)|Получает родительский комментарий этого ответа.|
||[authorEmail](/javascript/api/excel/excel.commentreply#authorEmail)|Получает электронную почту автора ответа на примечание.|
||[authorName](/javascript/api/excel/excel.commentreply#authorName)|Получает имя автора ответа на примечание.|
||[creationDate](/javascript/api/excel/excel.commentreply#creationDate)|Получает время создания ответа на примечание.|
||[id](/javascript/api/excel/excel.commentreply#id)|Указывает идентификатор ответа на комментарии.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentreplycollection#add_content__contentType_)|Создает ответ на комментарий для комментария.|
||[getCount()](/javascript/api/excel/excel.commentreplycollection#getCount__)|Получает количество ответов на примечания в коллекции.|
||[getItem(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getItem_commentReplyId_)|Возвращает ответ на примечание, определенное по идентификатору.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentreplycollection#getItemAt_index_)|Возвращает ответ на примечание на основе его позиции в коллекции.|
||[items](/javascript/api/excel/excel.commentreplycollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[enableFieldList](/javascript/api/excel/excel.pivotlayout#enableFieldList)|Указывает, можно ли показывать список полей в пользовательском интерфейсе.|
|[PivotTableStyle](/javascript/api/excel/excel.pivottablestyle)|[delete()](/javascript/api/excel/excel.pivottablestyle#delete__)|Удаляет стиль PivotTable.|
||[duplicate()](/javascript/api/excel/excel.pivottablestyle#duplicate__)|Создает дубликат этого стиля PivotTable с копиями всех элементов стиля.|
||[name](/javascript/api/excel/excel.pivottablestyle#name)|Получает имя стиля PivotTable.|
||[readOnly](/javascript/api/excel/excel.pivottablestyle#readOnly)|Указывает, является ли `PivotTableStyle` этот объект только для чтения.|
|[PivotTableStyleCollection](/javascript/api/excel/excel.pivottablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.pivottablestylecollection#add_name__makeUniqueName_)|Создает пробел `PivotTableStyle` с указанным именем.|
||[getCount()](/javascript/api/excel/excel.pivottablestylecollection#getCount__)|Получает количество стилей сводных таблиц в коллекции.|
||[getDefault()](/javascript/api/excel/excel.pivottablestylecollection#getDefault__)|Получает стиль PivotTable по умолчанию для области родительского объекта.|
||[getItem(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getItem_name_)|Получает `PivotTableStyle` имя.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.pivottablestylecollection#getItemOrNullObject_name_)|Получает `PivotTableStyle` имя.|
||[items](/javascript/api/excel/excel.pivottablestylecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[setDefault(newDefaultStyle: PivotTableStyle \| string)](/javascript/api/excel/excel.pivottablestylecollection#setDefault_newDefaultStyle_)|Задает стиль PivotTable по умолчанию для использования в области родительского объекта.|
|[Range](/javascript/api/excel/excel.range)|[group(groupOption: Excel. GroupOption)](/javascript/api/excel/excel.range#group_groupOption_)|Группы столбцов и строк для контура.|
||[hideGroupDetails(groupOption: Excel. GroupOption)](/javascript/api/excel/excel.range#hideGroupDetails_groupOption_)|Скрывает сведения о группе строки или столбца.|
||[height](/javascript/api/excel/excel.range#height)|Возвращает расстояние в точках для 100% масштабирования от верхнего края диапазона до нижнего края диапазона.|
||[left](/javascript/api/excel/excel.range#left)|Возвращает расстояние в точках для 100% масштабирования от левого края таблицы до левого края диапазона.|
||[top](/javascript/api/excel/excel.range#top)|Возвращает расстояние в точках для 100% масштабирования от верхнего края таблицы до верхнего края диапазона.|
||[width](/javascript/api/excel/excel.range#width)|Возвращает расстояние в точках для 100% масштабирования от левого края диапазона до правого края диапазона.|
||[showGroupDetails(groupOption: Excel. GroupOption)](/javascript/api/excel/excel.range#showGroupDetails_groupOption_)|Отображает сведения о группе строки или столбца.|
||[ungroup(groupOption: Excel. GroupOption)](/javascript/api/excel/excel.range#ungroup_groupOption_)|Разгруппировка столбцов и строк для контура.|
|[Shape](/javascript/api/excel/excel.shape)|[copyTo(destinationSheet?: Worksheet \| string)](/javascript/api/excel/excel.shape#copyTo_destinationSheet_)|Копирует и вклеит `Shape` объект.|
||[placement](/javascript/api/excel/excel.shape#placement)|Представляет способ прикрепления объекта к ячейкам под ним.|
|[Slicer](/javascript/api/excel/excel.slicer)|[caption](/javascript/api/excel/excel.slicer#caption)|Представляет подпись среза.|
||[clearFilters()](/javascript/api/excel/excel.slicer#clearFilters__)|Удаляет все фильтры, примененные к срезу.|
||[delete()](/javascript/api/excel/excel.slicer#delete__)|Удаляет срез.|
||[getSelectedItems()](/javascript/api/excel/excel.slicer#getSelectedItems__)|Возвращает массив имен выбранных ключей элементов.|
||[height](/javascript/api/excel/excel.slicer#height)|Представляет высоту среза (в пунктах).|
||[left](/javascript/api/excel/excel.slicer#left)|Представляет расстояние в пунктах от левого края среза до левого края листа.|
||[name](/javascript/api/excel/excel.slicer#name)|Представляет имя среза.|
||[id](/javascript/api/excel/excel.slicer#id)|Представляет уникальный ID среза.|
||[isFilterCleared](/javascript/api/excel/excel.slicer#isFilterCleared)|Значение, `true` если все фильтры, применяемые в настоящее время на срезе, будут очищены.|
||[slicerItems](/javascript/api/excel/excel.slicer#slicerItems)|Представляет коллекцию элементов slicer, которые являются частью среза.|
||[worksheet](/javascript/api/excel/excel.slicer#worksheet)|Представляет лист, содержащий срез.|
||[selectItems(items?: string[])](/javascript/api/excel/excel.slicer#selectItems_items_)|Выбирает элементы среза на основе ключей.|
||[sortBy](/javascript/api/excel/excel.slicer#sortBy)|Представляет порядок сортировки элементов в срезе.|
||[style](/javascript/api/excel/excel.slicer#style)|Постоянное значение, представляю которое представляет стиль среза.|
||[top](/javascript/api/excel/excel.slicer#top)|Представляет расстояние в пунктах от верхнего края среза до верхнего края листа.|
||[width](/javascript/api/excel/excel.slicer#width)|Представляет ширину среза (в пунктах).|
|[SlicerCollection](/javascript/api/excel/excel.slicercollection)|[add(slicerSource: string \| PivotTable \| Table, sourceField: string \| PivotField \| number \| TableColumn, slicerDestination?: string \| Worksheet)](/javascript/api/excel/excel.slicercollection#add_slicerSource__sourceField__slicerDestination_)|Добавляет новый срез в книгу.|
||[getCount()](/javascript/api/excel/excel.slicercollection#getCount__)|Возвращает количество срезов в коллекции.|
||[getItem(key: string)](/javascript/api/excel/excel.slicercollection#getItem_key_)|Получает объект slicer с его именем или ИД.|
||[getItemAt(index: number)](/javascript/api/excel/excel.slicercollection#getItemAt_index_)|Получает срез на основе его позиции в коллекции.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.slicercollection#getItemOrNullObject_key_)|Получает срез с его именем или ИД.|
||[items](/javascript/api/excel/excel.slicercollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[SlicerItem](/javascript/api/excel/excel.sliceritem)|[isSelected](/javascript/api/excel/excel.sliceritem#isSelected)|Значение, `true` если выбран элемент slicer.|
||[hasData](/javascript/api/excel/excel.sliceritem#hasData)|Значение, `true` если элемент slicer имеет данные.|
||[key](/javascript/api/excel/excel.sliceritem#key)|Представляет уникальное значение, соответствующее элементу среза.|
||[name](/javascript/api/excel/excel.sliceritem#name)|Представляет название, отображаемую в пользовательском Excel интерфейсе.|
|[SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection)|[getCount()](/javascript/api/excel/excel.sliceritemcollection#getCount__)|Возвращает количество элементов в срезе.|
||[getItem(key: string)](/javascript/api/excel/excel.sliceritemcollection#getItem_key_)|Получает объект элемента среза по ключу или имени.|
||[getItemAt(index: number)](/javascript/api/excel/excel.sliceritemcollection#getItemAt_index_)|Получает элемент среза на основе его позиции в коллекции.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.sliceritemcollection#getItemOrNullObject_key_)|Получает элемент среза по ключу или имени.|
||[items](/javascript/api/excel/excel.sliceritemcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[SlicerStyle](/javascript/api/excel/excel.slicerstyle)|[delete()](/javascript/api/excel/excel.slicerstyle#delete__)|Удаляет стиль среза.|
||[duplicate()](/javascript/api/excel/excel.slicerstyle#duplicate__)|Создает дубликат этого стиля среза с копиями всех элементов стиля.|
||[name](/javascript/api/excel/excel.slicerstyle#name)|Получает имя стиля slicer.|
||[readOnly](/javascript/api/excel/excel.slicerstyle#readOnly)|Указывает, является ли `SlicerStyle` этот объект только для чтения.|
|[SlicerStyleCollection](/javascript/api/excel/excel.slicerstylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.slicerstylecollection#add_name__makeUniqueName_)|Создает пустой стиль среза с указанным именем.|
||[getCount()](/javascript/api/excel/excel.slicerstylecollection#getCount__)|Получает количество стилей срезов в коллекции.|
||[getDefault()](/javascript/api/excel/excel.slicerstylecollection#getDefault__)|Получает по `SlicerStyle` умолчанию область родительского объекта.|
||[getItem(name: string)](/javascript/api/excel/excel.slicerstylecollection#getItem_name_)|Получает `SlicerStyle` имя.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.slicerstylecollection#getItemOrNullObject_name_)|Получает `SlicerStyle` имя.|
||[items](/javascript/api/excel/excel.slicerstylecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[setDefault(newDefaultStyle: SlicerStyle \| string)](/javascript/api/excel/excel.slicerstylecollection#setDefault_newDefaultStyle_)|Задает стиль среза по умолчанию для использования в области родительского объекта.|
|[TableStyle](/javascript/api/excel/excel.tablestyle)|[delete()](/javascript/api/excel/excel.tablestyle#delete__)|Удаляет стиль таблицы.|
||[duplicate()](/javascript/api/excel/excel.tablestyle#duplicate__)|Создает дубликат этого стиля таблицы с копиями всех элементов стиля.|
||[name](/javascript/api/excel/excel.tablestyle#name)|Получает имя стиля таблицы.|
||[readOnly](/javascript/api/excel/excel.tablestyle#readOnly)|Указывает, является ли `TableStyle` этот объект только для чтения.|
|[TableStyleCollection](/javascript/api/excel/excel.tablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.tablestylecollection#add_name__makeUniqueName_)|Создает пробел `TableStyle` с указанным именем.|
||[getCount()](/javascript/api/excel/excel.tablestylecollection#getCount__)|Получает количество стилей таблиц в коллекции.|
||[getDefault()](/javascript/api/excel/excel.tablestylecollection#getDefault__)|Получает стиль таблицы по умолчанию для области родительского объекта.|
||[getItem(name: string)](/javascript/api/excel/excel.tablestylecollection#getItem_name_)|Получает `TableStyle` имя.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.tablestylecollection#getItemOrNullObject_name_)|Получает `TableStyle` имя.|
||[items](/javascript/api/excel/excel.tablestylecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[setDefault(newDefaultStyle: TableStyle \| string)](/javascript/api/excel/excel.tablestylecollection#setDefault_newDefaultStyle_)|Задает стиль таблицы по умолчанию для использования в области родительского объекта.|
|[TimelineStyle](/javascript/api/excel/excel.timelinestyle)|[delete()](/javascript/api/excel/excel.timelinestyle#delete__)|Удаляет стиль таблицы.|
||[duplicate()](/javascript/api/excel/excel.timelinestyle#duplicate__)|Создает дубликат этого стиля временной шкалы с копиями всех элементов стиля.|
||[name](/javascript/api/excel/excel.timelinestyle#name)|Получает имя стиля timeline.|
||[readOnly](/javascript/api/excel/excel.timelinestyle#readOnly)|Указывает, является ли `TimelineStyle` этот объект только для чтения.|
|[TimelineStyleCollection](/javascript/api/excel/excel.timelinestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.timelinestylecollection#add_name__makeUniqueName_)|Создает пробел `TimelineStyle` с указанным именем.|
||[getCount()](/javascript/api/excel/excel.timelinestylecollection#getCount__)|Получает количество стилей временной шкалы в коллекции.|
||[getDefault()](/javascript/api/excel/excel.timelinestylecollection#getDefault__)|Получает стиль временной шкалы по умолчанию для области родительского объекта.|
||[getItem(name: string)](/javascript/api/excel/excel.timelinestylecollection#getItem_name_)|Получает `TimelineStyle` имя.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.timelinestylecollection#getItemOrNullObject_name_)|Получает `TimelineStyle` имя.|
||[items](/javascript/api/excel/excel.timelinestylecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[setDefault(newDefaultStyle: TimelineStyle \| string)](/javascript/api/excel/excel.timelinestylecollection#setDefault_newDefaultStyle_)|Задает стиль временной шкалы по умолчанию для использования в области родительского объекта.|
|[Workbook](/javascript/api/excel/excel.workbook)|[getActiveSlicer()](/javascript/api/excel/excel.workbook#getActiveSlicer__)|Получает текущий активный срез в книге.|
||[getActiveSlicerOrNullObject()](/javascript/api/excel/excel.workbook#getActiveSlicerOrNullObject__)|Получает текущий активный срез в книге.|
||[comments](/javascript/api/excel/excel.workbook#comments)|Представляет коллекцию комментариев, связанных с книгой.|
||[pivotTableStyles](/javascript/api/excel/excel.workbook#pivotTableStyles)|Представляет коллекцию объектов PivotTableStyles, связанных с книгой.|
||[slicerStyles](/javascript/api/excel/excel.workbook#slicerStyles)|Представляет коллекцию объектов SlicerStyles, связанных с книгой.|
||[slicers](/javascript/api/excel/excel.workbook#slicers)|Представляет коллекцию срезов, связанных с книгой.|
||[tableStyles](/javascript/api/excel/excel.workbook#tableStyles)|Представляет коллекцию объектов TableStyles, связанных с книгой.|
||[timelineStyles](/javascript/api/excel/excel.workbook#timelineStyles)|Представляет коллекцию объектов TimelineStyles, связанных с книгой.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[comments](/javascript/api/excel/excel.worksheet#comments)|Возвращает коллекцию всех объектов Comments на листе.|
||[onColumnSorted](/javascript/api/excel/excel.worksheet#onColumnSorted)|Возникает при сортировке одного или нескольких столбцов.|
||[onRowSorted](/javascript/api/excel/excel.worksheet#onRowSorted)|Возникает при сортировке одной или нескольких строк.|
||[onSingleClicked](/javascript/api/excel/excel.worksheet#onSingleClicked)|Происходит, когда в таблице происходит действие левого щелчка или прослушиваемого действия.|
||[slicers](/javascript/api/excel/excel.worksheet#slicers)|Возвращает коллекцию срезов, которые являются частью таблицы.|
||[showOutlineLevels (rowLevels: number, columnLevels: number)](/javascript/api/excel/excel.worksheet#showOutlineLevels_rowLevels__columnLevels_)|Показывает группы строк или столбцов по уровням контура.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onColumnSorted](/javascript/api/excel/excel.worksheetcollection#onColumnSorted)|Возникает при сортировке одного или нескольких столбцов.|
||[onRowSorted](/javascript/api/excel/excel.worksheetcollection#onRowSorted)|Возникает при сортировке одной или нескольких строк.|
||[onSingleClicked](/javascript/api/excel/excel.worksheetcollection#onSingleClicked)|Происходит, когда в коллекции таблицы происходит операция нажатием левой кнопкой мыши или нажатием на нее.|
|[WorksheetColumnSortedEventArgs](/javascript/api/excel/excel.worksheetcolumnsortedeventargs)|[address](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#address)|Получает адрес диапазона, представляющий отсортированные области конкретного листа.|
||[источник](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#source)|Получает источник события.|
||[type](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#worksheetId)|Получает ID таблицы, в которой произошла сортировка.|
|[WorksheetRowSortedEventArgs](/javascript/api/excel/excel.worksheetrowsortedeventargs)|[address](/javascript/api/excel/excel.worksheetrowsortedeventargs#address)|Получает адрес диапазона, представляющий отсортированные области конкретного листа.|
||[источник](/javascript/api/excel/excel.worksheetrowsortedeventargs#source)|Получает источник события.|
||[type](/javascript/api/excel/excel.worksheetrowsortedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowsortedeventargs#worksheetId)|Получает ID таблицы, в которой произошла сортировка.|
|[WorksheetSingleClickedEventArgs](/javascript/api/excel/excel.worksheetsingleclickedeventargs)|[address](/javascript/api/excel/excel.worksheetsingleclickedeventargs#address)|Получает адрес, представляющий ячейку, по которой выполнен щелчок левой кнопкой мыши или нажатие, для определенного листа.|
||[offsetX](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsetX)|Расстояние в точках от слева нажатой или прослушиваемой точки до левого (или правого для языков справа налево) границы сетки ячейки слева нажатой или прослушиваемой.|
||[offsetY](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsetY)|Расстояние в пунктах от точки щелчка левой кнопкой мыши или нажатия до верхнего края сетки ячейки, по которой выполнен щелчок левой кнопкой мыши или нажатие.|
||[type](/javascript/api/excel/excel.worksheetsingleclickedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetsingleclickedeventargs#worksheetId)|Получает ID таблицы, в которой ячейка была нажата левой кнопкой мыши или прослушивается.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.10&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)