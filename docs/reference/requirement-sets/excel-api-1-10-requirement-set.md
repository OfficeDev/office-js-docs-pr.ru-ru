---
title: Excel API JavaScript установлено 1.10
description: Сведения о наборе требований ExcelApi 1.10.
ms.date: 04/02/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 53cf0ec55a26f02a615a3c5eee0b718b818790d0
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746342"
---
# <a name="whats-new-in-excel-javascript-api-110"></a>Новые возможности в Excel API JavaScript 1.10

В ExcelApi 1.10 представлены ключевые функции, такие как комментарии, контуры и срезы. Кроме того, добавлена поддержка событий для нажатия и сортировки на уровне таблицы.

| Функциональная область | Описание | Соответствующие объекты |
|:--- |:--- |:--- |
| [Примечания](../../excel/excel-add-ins-comments.md) | Добавление, редактирование и удаление примечаний. | [Comment](/javascript/api/excel/excel.comment), [CommentCollection](/javascript/api/excel/excel.commentcollection) |
| [Контуры](../../excel/excel-add-ins-ranges-group.md) | Групповые строки и столбцы для формирования общих контуров. | [Диапазон](/javascript/api/excel/excel.range), [таблица](/javascript/api/excel/excel.worksheet) |
| [Slicers](../../excel/excel-add-ins-pivottables.md#filter-with-slicers) | Вставка и настройка срезов для таблиц и сводных таблиц. | [Slicer](/javascript/api/excel/excel.slicer) |
| [Дополнительные события таблицы](../../excel/excel-add-ins-events.md) | Щелкните кнопку мыши и отсортировать события в таблице. | [Таблица (События)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-events-member) |

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в Excel API JavaScript, за набором 1.10. Чтобы просмотреть справочную документацию API для всех API, поддерживаемых Excel API JavaScript, установленного 1.10 или ранее, см. Excel API в наборе требований [1.10 или ранее](/javascript/api/excel?view=excel-js-1.10&preserve-view=true).

| Класс | Поля | Описание |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[authorEmail](/javascript/api/excel/excel.comment#excel-excel-comment-authoremail-member)|Получает электронную почту автора примечания.|
||[authorName](/javascript/api/excel/excel.comment#excel-excel-comment-authorname-member)|Получает имя автора примечания.|
||[content](/javascript/api/excel/excel.comment#excel-excel-comment-content-member)|Содержимое комментария.|
||[creationDate](/javascript/api/excel/excel.comment#excel-excel-comment-creationdate-member)|Получает время создания примечания.|
||[delete()](/javascript/api/excel/excel.comment#excel-excel-comment-delete-member(1))|Удаляет комментарий и все подключенные ответы.|
||[getLocation()](/javascript/api/excel/excel.comment#excel-excel-comment-getlocation-member(1))|Получает ячейку, в которой расположен этот комментарий.|
||[id](/javascript/api/excel/excel.comment#excel-excel-comment-id-member)|Указывает идентификатор комментария.|
||[replies](/javascript/api/excel/excel.comment#excel-excel-comment-replies-member)|Представляет коллекцию объектов ответов, связанных с примечанием.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(cellAddress: Range \| string, content: string, contentType?: Excel. ContentType)](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-add-member(1))|Создает новое примечание с указанным содержимым в определенной ячейке.|
||[getCount()](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-getcount-member(1))|Получает количество примечаний в коллекции.|
||[getItem(commentId: string)](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-getitem-member(1))|Получает примечание из коллекции на основе его идентификатора.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-getitemat-member(1))|Получает примечание из коллекции на основе его позиции.|
||[getItemByCell(cellAddress: Range \| string)](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-getitembycell-member(1))|Получает примечание из указанной ячейки.|
||[getItemByReplyId(replyId: string)](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-getitembyreplyid-member(1))|Получает комментарий, к которому подключен данный ответ.|
||[items](/javascript/api/excel/excel.commentcollection#excel-excel-commentcollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[authorEmail](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-authoremail-member)|Получает электронную почту автора ответа на примечание.|
||[authorName](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-authorname-member)|Получает имя автора ответа на примечание.|
||[content](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-content-member)|Содержимое ответа на комментарий.|
||[creationDate](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-creationdate-member)|Получает время создания ответа на примечание.|
||[delete()](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-delete-member(1))|Удаляет ответ на примечание.|
||[getLocation()](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-getlocation-member(1))|Получает ячейку, в которой находится ответ на этот комментарий.|
||[getParentComment()](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-getparentcomment-member(1))|Получает родительский комментарий этого ответа.|
||[id](/javascript/api/excel/excel.commentreply#excel-excel-commentreply-id-member)|Указывает идентификатор ответа на комментарии.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-add-member(1))|Создает ответ на комментарий для комментария.|
||[getCount()](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-getcount-member(1))|Получает количество ответов на примечания в коллекции.|
||[getItem(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-getitem-member(1))|Возвращает ответ на примечание, определенное по идентификатору.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-getitemat-member(1))|Возвращает ответ на примечание на основе его позиции в коллекции.|
||[items](/javascript/api/excel/excel.commentreplycollection#excel-excel-commentreplycollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[enableFieldList](/javascript/api/excel/excel.pivotlayout#excel-excel-pivotlayout-enablefieldlist-member)|Указывает, можно ли показывать список полей в пользовательском интерфейсе.|
|[PivotTableStyle](/javascript/api/excel/excel.pivottablestyle)|[delete()](/javascript/api/excel/excel.pivottablestyle#excel-excel-pivottablestyle-delete-member(1))|Удаляет стиль PivotTable.|
||[duplicate()](/javascript/api/excel/excel.pivottablestyle#excel-excel-pivottablestyle-duplicate-member(1))|Создает дубликат этого стиля PivotTable с копиями всех элементов стиля.|
||[name](/javascript/api/excel/excel.pivottablestyle#excel-excel-pivottablestyle-name-member)|Получает имя стиля PivotTable.|
||[readOnly](/javascript/api/excel/excel.pivottablestyle#excel-excel-pivottablestyle-readonly-member)|Указывает, является ли `PivotTableStyle` этот объект только для чтения.|
|[PivotTableStyleCollection](/javascript/api/excel/excel.pivottablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-add-member(1))|Создает пробел с `PivotTableStyle` указанным именем.|
||[getCount()](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-getcount-member(1))|Получает количество стилей сводных таблиц в коллекции.|
||[getDefault()](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-getdefault-member(1))|Получает стиль PivotTable по умолчанию для области родительского объекта.|
||[getItem(name: string)](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-getitem-member(1))|Получает имя `PivotTableStyle` .|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-getitemornullobject-member(1))|Получает имя `PivotTableStyle` .|
||[items](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
||[setDefault(newDefaultStyle: PivotTableStyle \| string)](/javascript/api/excel/excel.pivottablestylecollection#excel-excel-pivottablestylecollection-setdefault-member(1))|Задает стиль PivotTable по умолчанию для использования в области родительского объекта.|
|[Range](/javascript/api/excel/excel.range)|[group(groupOption: Excel. GroupOption)](/javascript/api/excel/excel.range#excel-excel-range-group-member(1))|Группы столбцов и строк для контура.|
||[height](/javascript/api/excel/excel.range#excel-excel-range-height-member)|Возвращает расстояние в точках для 100% масштабирования от верхнего края диапазона до нижнего края диапазона.|
||[hideGroupDetails(groupOption: Excel. GroupOption)](/javascript/api/excel/excel.range#excel-excel-range-hidegroupdetails-member(1))|Скрывает сведения о группе строки или столбца.|
||[left](/javascript/api/excel/excel.range#excel-excel-range-left-member)|Возвращает расстояние в точках для 100% масштабирования от левого края таблицы до левого края диапазона.|
||[showGroupDetails(groupOption: Excel. GroupOption)](/javascript/api/excel/excel.range#excel-excel-range-showgroupdetails-member(1))|Отображает сведения о группе строки или столбца.|
||[top](/javascript/api/excel/excel.range#excel-excel-range-top-member)|Возвращает расстояние в точках для 100% масштабирования от верхнего края таблицы до верхнего края диапазона.|
||[ungroup(groupOption: Excel. GroupOption)](/javascript/api/excel/excel.range#excel-excel-range-ungroup-member(1))|Разгруппировка столбцов и строк для контура.|
||[width](/javascript/api/excel/excel.range#excel-excel-range-width-member)|Возвращает расстояние в точках для 100% масштабирования от левого края диапазона до правого края диапазона.|
|[Shape](/javascript/api/excel/excel.shape)|[copyTo(destinationSheet?: Worksheet \| string)](/javascript/api/excel/excel.shape#excel-excel-shape-copyto-member(1))|Копирует и вклеит `Shape` объект.|
||[placement](/javascript/api/excel/excel.shape#excel-excel-shape-placement-member)|Представляет способ прикрепления объекта к ячейкам под ним.|
|[Slicer](/javascript/api/excel/excel.slicer)|[caption](/javascript/api/excel/excel.slicer#excel-excel-slicer-caption-member)|Представляет подпись среза.|
||[clearFilters()](/javascript/api/excel/excel.slicer#excel-excel-slicer-clearfilters-member(1))|Удаляет все фильтры, примененные к срезу.|
||[delete()](/javascript/api/excel/excel.slicer#excel-excel-slicer-delete-member(1))|Удаляет срез.|
||[getSelectedItems()](/javascript/api/excel/excel.slicer#excel-excel-slicer-getselecteditems-member(1))|Возвращает массив имен выбранных ключей элементов.|
||[height](/javascript/api/excel/excel.slicer#excel-excel-slicer-height-member)|Представляет высоту среза (в пунктах).|
||[id](/javascript/api/excel/excel.slicer#excel-excel-slicer-id-member)|Представляет уникальный ID среза.|
||[isFilterCleared](/javascript/api/excel/excel.slicer#excel-excel-slicer-isfiltercleared-member)|Значение, `true` если все фильтры, применяемые в настоящее время на срезе, будут очищены.|
||[left](/javascript/api/excel/excel.slicer#excel-excel-slicer-left-member)|Представляет расстояние в пунктах от левого края среза до левого края листа.|
||[name](/javascript/api/excel/excel.slicer#excel-excel-slicer-name-member)|Представляет имя среза.|
||[selectItems(items?: string[])](/javascript/api/excel/excel.slicer#excel-excel-slicer-selectitems-member(1))|Выбирает элементы среза на основе ключей.|
||[slicerItems](/javascript/api/excel/excel.slicer#excel-excel-slicer-sliceritems-member)|Представляет коллекцию элементов slicer, которые являются частью среза.|
||[sortBy](/javascript/api/excel/excel.slicer#excel-excel-slicer-sortby-member)|Представляет порядок сортировки элементов в срезе.|
||[style](/javascript/api/excel/excel.slicer#excel-excel-slicer-style-member)|Постоянное значение, представляю которое представляет стиль среза.|
||[top](/javascript/api/excel/excel.slicer#excel-excel-slicer-top-member)|Представляет расстояние в пунктах от верхнего края среза до верхнего края листа.|
||[width](/javascript/api/excel/excel.slicer#excel-excel-slicer-width-member)|Представляет ширину среза (в пунктах).|
||[worksheet](/javascript/api/excel/excel.slicer#excel-excel-slicer-worksheet-member)|Представляет лист, содержащий срез.|
|[SlicerCollection](/javascript/api/excel/excel.slicercollection)|[add(slicerSource: string \| PivotTable \| Table, sourceField: string \| PivotField \| number \| TableColumn, slicerDestination?: string \| Worksheet)](/javascript/api/excel/excel.slicercollection#excel-excel-slicercollection-add-member(1))|Добавляет новый срез в книгу.|
||[getCount()](/javascript/api/excel/excel.slicercollection#excel-excel-slicercollection-getcount-member(1))|Возвращает количество срезов в коллекции.|
||[getItem(key: string)](/javascript/api/excel/excel.slicercollection#excel-excel-slicercollection-getitem-member(1))|Получает объект slicer с его именем или ИД.|
||[getItemAt(index: number)](/javascript/api/excel/excel.slicercollection#excel-excel-slicercollection-getitemat-member(1))|Получает срез на основе его позиции в коллекции.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.slicercollection#excel-excel-slicercollection-getitemornullobject-member(1))|Получает срез с его именем или ИД.|
||[items](/javascript/api/excel/excel.slicercollection#excel-excel-slicercollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[SlicerItem](/javascript/api/excel/excel.sliceritem)|[hasData](/javascript/api/excel/excel.sliceritem#excel-excel-sliceritem-hasdata-member)|Значение, если `true` элемент slicer имеет данные.|
||[isSelected](/javascript/api/excel/excel.sliceritem#excel-excel-sliceritem-isselected-member)|Значение, `true` если выбран элемент slicer.|
||[key](/javascript/api/excel/excel.sliceritem#excel-excel-sliceritem-key-member)|Представляет уникальное значение, соответствующее элементу среза.|
||[name](/javascript/api/excel/excel.sliceritem#excel-excel-sliceritem-name-member)|Представляет название, отображаемую в пользовательском Excel интерфейсе.|
|[SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection)|[getCount()](/javascript/api/excel/excel.sliceritemcollection#excel-excel-sliceritemcollection-getcount-member(1))|Возвращает количество элементов в срезе.|
||[getItem(key: string)](/javascript/api/excel/excel.sliceritemcollection#excel-excel-sliceritemcollection-getitem-member(1))|Получает объект элемента среза по ключу или имени.|
||[getItemAt(index: number)](/javascript/api/excel/excel.sliceritemcollection#excel-excel-sliceritemcollection-getitemat-member(1))|Получает элемент среза на основе его позиции в коллекции.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.sliceritemcollection#excel-excel-sliceritemcollection-getitemornullobject-member(1))|Получает элемент среза по ключу или имени.|
||[items](/javascript/api/excel/excel.sliceritemcollection#excel-excel-sliceritemcollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[SlicerStyle](/javascript/api/excel/excel.slicerstyle)|[delete()](/javascript/api/excel/excel.slicerstyle#excel-excel-slicerstyle-delete-member(1))|Удаляет стиль среза.|
||[duplicate()](/javascript/api/excel/excel.slicerstyle#excel-excel-slicerstyle-duplicate-member(1))|Создает дубликат этого стиля среза с копиями всех элементов стиля.|
||[name](/javascript/api/excel/excel.slicerstyle#excel-excel-slicerstyle-name-member)|Получает имя стиля slicer.|
||[readOnly](/javascript/api/excel/excel.slicerstyle#excel-excel-slicerstyle-readonly-member)|Указывает, является ли `SlicerStyle` этот объект только для чтения.|
|[SlicerStyleCollection](/javascript/api/excel/excel.slicerstylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-add-member(1))|Создает пустой стиль среза с указанным именем.|
||[getCount()](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-getcount-member(1))|Получает количество стилей срезов в коллекции.|
||[getDefault()](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-getdefault-member(1))|Получает по умолчанию `SlicerStyle` область родительского объекта.|
||[getItem(name: string)](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-getitem-member(1))|Получает имя `SlicerStyle` .|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-getitemornullobject-member(1))|Получает имя `SlicerStyle` .|
||[items](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
||[setDefault(newDefaultStyle: SlicerStyle \| string)](/javascript/api/excel/excel.slicerstylecollection#excel-excel-slicerstylecollection-setdefault-member(1))|Задает стиль среза по умолчанию для использования в области родительского объекта.|
|[TableStyle](/javascript/api/excel/excel.tablestyle)|[delete()](/javascript/api/excel/excel.tablestyle#excel-excel-tablestyle-delete-member(1))|Удаляет стиль таблицы.|
||[duplicate()](/javascript/api/excel/excel.tablestyle#excel-excel-tablestyle-duplicate-member(1))|Создает дубликат этого стиля таблицы с копиями всех элементов стиля.|
||[name](/javascript/api/excel/excel.tablestyle#excel-excel-tablestyle-name-member)|Получает имя стиля таблицы.|
||[readOnly](/javascript/api/excel/excel.tablestyle#excel-excel-tablestyle-readonly-member)|Указывает, является ли `TableStyle` этот объект только для чтения.|
|[TableStyleCollection](/javascript/api/excel/excel.tablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-add-member(1))|Создает пробел с `TableStyle` указанным именем.|
||[getCount()](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-getcount-member(1))|Получает количество стилей таблиц в коллекции.|
||[getDefault()](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-getdefault-member(1))|Получает стиль таблицы по умолчанию для области родительского объекта.|
||[getItem(name: string)](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-getitem-member(1))|Получает имя `TableStyle` .|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-getitemornullobject-member(1))|Получает имя `TableStyle` .|
||[items](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
||[setDefault(newDefaultStyle: TableStyle \| string)](/javascript/api/excel/excel.tablestylecollection#excel-excel-tablestylecollection-setdefault-member(1))|Задает стиль таблицы по умолчанию для использования в области родительского объекта.|
|[TimelineStyle](/javascript/api/excel/excel.timelinestyle)|[delete()](/javascript/api/excel/excel.timelinestyle#excel-excel-timelinestyle-delete-member(1))|Удаляет стиль таблицы.|
||[duplicate()](/javascript/api/excel/excel.timelinestyle#excel-excel-timelinestyle-duplicate-member(1))|Создает дубликат этого стиля временной шкалы с копиями всех элементов стиля.|
||[name](/javascript/api/excel/excel.timelinestyle#excel-excel-timelinestyle-name-member)|Получает имя стиля timeline.|
||[readOnly](/javascript/api/excel/excel.timelinestyle#excel-excel-timelinestyle-readonly-member)|Указывает, является ли `TimelineStyle` этот объект только для чтения.|
|[TimelineStyleCollection](/javascript/api/excel/excel.timelinestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-add-member(1))|Создает пробел с `TimelineStyle` указанным именем.|
||[getCount()](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-getcount-member(1))|Получает количество стилей временной шкалы в коллекции.|
||[getDefault()](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-getdefault-member(1))|Получает стиль временной шкалы по умолчанию для области родительского объекта.|
||[getItem(name: string)](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-getitem-member(1))|Получает имя `TimelineStyle` .|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-getitemornullobject-member(1))|Получает имя `TimelineStyle` .|
||[items](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
||[setDefault(newDefaultStyle: TimelineStyle \| string)](/javascript/api/excel/excel.timelinestylecollection#excel-excel-timelinestylecollection-setdefault-member(1))|Задает стиль временной шкалы по умолчанию для использования в области родительского объекта.|
|[Workbook](/javascript/api/excel/excel.workbook)|[comments](/javascript/api/excel/excel.workbook#excel-excel-workbook-comments-member)|Представляет коллекцию комментариев, связанных с книгой.|
||[getActiveSlicer()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getactiveslicer-member(1))|Получает текущий активный срез в книге.|
||[getActiveSlicerOrNullObject()](/javascript/api/excel/excel.workbook#excel-excel-workbook-getactiveslicerornullobject-member(1))|Получает текущий активный срез в книге.|
||[pivotTableStyles](/javascript/api/excel/excel.workbook#excel-excel-workbook-pivottablestyles-member)|Представляет коллекцию объектов PivotTableStyles, связанных с книгой.|
||[slicerStyles](/javascript/api/excel/excel.workbook#excel-excel-workbook-slicerstyles-member)|Представляет коллекцию объектов SlicerStyles, связанных с книгой.|
||[slicers](/javascript/api/excel/excel.workbook#excel-excel-workbook-slicers-member)|Представляет коллекцию срезов, связанных с книгой.|
||[tableStyles](/javascript/api/excel/excel.workbook#excel-excel-workbook-tablestyles-member)|Представляет коллекцию объектов TableStyles, связанных с книгой.|
||[timelineStyles](/javascript/api/excel/excel.workbook#excel-excel-workbook-timelinestyles-member)|Представляет коллекцию объектов TimelineStyles, связанных с книгой.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[comments](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-comments-member)|Возвращает коллекцию всех объектов Comments на листе.|
||[onColumnSorted](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-oncolumnsorted-member)|Возникает при сортировке одного или нескольких столбцов.|
||[onRowSorted](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onrowsorted-member)|Возникает при сортировке одной или нескольких строк.|
||[onSingleClicked](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-onsingleclicked-member)|Происходит, когда в таблице происходит действие левого щелчка или прослушиваемого действия.|
||[showOutlineLevels (rowLevels: number, columnLevels: number)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-showoutlinelevels-member(1))|Показывает группы строк или столбцов по уровням контура.|
||[slicers](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-slicers-member)|Возвращает коллекцию срезов, которые являются частью таблицы.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[onColumnSorted](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-oncolumnsorted-member)|Возникает при сортировке одного или нескольких столбцов.|
||[onRowSorted](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onrowsorted-member)|Возникает при сортировке одной или нескольких строк.|
||[onSingleClicked](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-onsingleclicked-member)|Происходит, когда в коллекции таблицы происходит операция нажатием левой кнопкой мыши или нажатием на нее.|
|[WorksheetColumnSortedEventArgs](/javascript/api/excel/excel.worksheetcolumnsortedeventargs)|[address](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#excel-excel-worksheetcolumnsortedeventargs-address-member)|Получает адрес диапазона, представляющий отсортированные области конкретного листа.|
||[источник](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#excel-excel-worksheetcolumnsortedeventargs-source-member)|Получает источник события.|
||[type](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#excel-excel-worksheetcolumnsortedeventargs-type-member)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#excel-excel-worksheetcolumnsortedeventargs-worksheetid-member)|Получает ID таблицы, в которой произошла сортировка.|
|[WorksheetRowSortedEventArgs](/javascript/api/excel/excel.worksheetrowsortedeventargs)|[address](/javascript/api/excel/excel.worksheetrowsortedeventargs#excel-excel-worksheetrowsortedeventargs-address-member)|Получает адрес диапазона, представляющий отсортированные области конкретного листа.|
||[источник](/javascript/api/excel/excel.worksheetrowsortedeventargs#excel-excel-worksheetrowsortedeventargs-source-member)|Получает источник события.|
||[type](/javascript/api/excel/excel.worksheetrowsortedeventargs#excel-excel-worksheetrowsortedeventargs-type-member)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowsortedeventargs#excel-excel-worksheetrowsortedeventargs-worksheetid-member)|Получает ID таблицы, в которой произошла сортировка.|
|[WorksheetSingleClickedEventArgs](/javascript/api/excel/excel.worksheetsingleclickedeventargs)|[address](/javascript/api/excel/excel.worksheetsingleclickedeventargs#excel-excel-worksheetsingleclickedeventargs-address-member)|Получает адрес, представляющий ячейку, по которой выполнен щелчок левой кнопкой мыши или нажатие, для определенного листа.|
||[offsetX](/javascript/api/excel/excel.worksheetsingleclickedeventargs#excel-excel-worksheetsingleclickedeventargs-offsetx-member)|Расстояние в точках от слева нажатой или прослушиваемой точки до левого (или правого для языков справа налево) границы сетки ячейки слева нажатой или прослушиваемой.|
||[offsetY](/javascript/api/excel/excel.worksheetsingleclickedeventargs#excel-excel-worksheetsingleclickedeventargs-offsety-member)|Расстояние в пунктах от точки щелчка левой кнопкой мыши или нажатия до верхнего края сетки ячейки, по которой выполнен щелчок левой кнопкой мыши или нажатие.|
||[type](/javascript/api/excel/excel.worksheetsingleclickedeventargs#excel-excel-worksheetsingleclickedeventargs-type-member)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetsingleclickedeventargs#excel-excel-worksheetsingleclickedeventargs-worksheetid-member)|Получает ID таблицы, в которой ячейка была нажата левой кнопкой мыши или прослушивается.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.10&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)