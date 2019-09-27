---
title: Предварительные версии API JavaScript для Excel
description: Сведения о предстоящих API JavaScript для Excel
ms.date: 09/24/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 95cecd1c6a0b26f945e96834fe39757490788a97
ms.sourcegitcommit: c8914ce0f48a0c19bbfc3276a80d090bb7ce68e1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/26/2019
ms.locfileid: "37235318"
---
# <a name="excel-javascript-preview-apis"></a>Предварительные версии API JavaScript для Excel

Новые API JavaScript для Excel сначала выпускаются в "предварительной версии", а затем становятся частью определенного нумерованного набора обязательных элементов после выполнения достаточного тестирования и получения отзывов пользователей.

В первой таблице представлен краткий обзор API, а в последующей таблице приведен подробный список.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

| Функциональная область | Описание | Соответствующие объекты |
|:--- |:--- |:--- |
| [Срез](../../excel/excel-add-ins-pivottables.md#slicers-preview) | Вставка и настройка срезов для таблиц и сводных таблиц. | [Slicer](/javascript/api/excel/excel.slicer) |
| [Примечания](../../excel/excel-add-ins-workbooks.md#comments-preview) | Добавление, редактирование и удаление примечаний. | [Comment](/javascript/api/excel/excel.comment), [CommentCollection](/javascript/api/excel/excel.commentcollection) |
| [Сохранение](../../excel/excel-add-ins-workbooks.md#save-the-workbook-preview) и [закрытие](../../excel/excel-add-ins-workbooks.md#close-the-workbook-preview) рабочей книги | Сохранение и закрытие книг.  | [Workbook](/javascript/api/excel/excel.workbook) |
| [Вставка книги](../../excel/excel-add-ins-workbooks.md#insert-a-copy-of-an-existing-workbook-into-the-current-one-preview) | Вставка одной книги в другую.  | [Workbook](/javascript/api/excel/excel.worksheetcollection) |

## <a name="api-list"></a>Список API

В следующей таблице перечислены API JavaScript для Excel, находящиеся в предварительной версии. Чтобы просмотреть полный список всех интерфейсов API JavaScript для Excel (включая предварительные API и ранее выпущенные API), ознакомьтесь со статьями [все API JavaScript для Excel](/javascript/api/excel?view=excel-js-preview).

| Класс | Поля | Описание |
|:---|:---|:---|
|[Comment](/javascript/api/excel/excel.comment)|[content](/javascript/api/excel/excel.comment#content)|Получает или задает содержимое примечания. Строка является обычным текстом.|
||[delete()](/javascript/api/excel/excel.comment#delete--)|Удаляет комментарий и все подключенные ответы.|
||[getLocation()](/javascript/api/excel/excel.comment#getlocation--)|Получает ячейку, в которой находится этот комментарий.|
||[authorEmail](/javascript/api/excel/excel.comment#authoremail)|Получает электронную почту автора примечания.|
||[authorName](/javascript/api/excel/excel.comment#authorname)|Получает имя автора примечания.|
||[creationDate](/javascript/api/excel/excel.comment#creationdate)|Получает время создания примечания. Возвращает значение null, если примечание было преобразовано из заметки, так как у примечания нет даты создания.|
||[id](/javascript/api/excel/excel.comment#id)|Представляет идентификатор примечания. Только для чтения.|
||[mentions](/javascript/api/excel/excel.comment#mentions)|Получает объекты (например, людей), которые упоминаются в комментариях.|
||[replies](/javascript/api/excel/excel.comment#replies)|Представляет коллекцию объектов ответов, связанных с примечанием. Только для чтения.|
||[ричконтент](/javascript/api/excel/excel.comment#richcontent)|Получает форматированный текст комментария (например, упоминания в комментариях). Эта строка не предназначена для отображения конечным пользователям. Надстройка должна использовать эту надстройку только для анализа форматированного содержимого комментариев.|
||[определяем](/javascript/api/excel/excel.comment#resolved)|Получает или задает состояние потока комментариев. Значение "true" означает, что поток комментариев находится в состоянии "разрешено".|
||[Упдатементионс (Контентвисментионс: Excel. Комментричконтент)](/javascript/api/excel/excel.comment#updatementions-contentwithmentions-)|Обновляет содержимое комментария с помощью специально отформатированной строки и списка упоминаний.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[Add (Целладдресс: строка \| Range, Content: комментричконтент \| String, ContentType?: Excel. ContentType)](/javascript/api/excel/excel.commentcollection#add-celladdress--content--contenttype-)|Создает новое примечание с указанным содержимым в определенной ячейке. Если `InvalidArgument` указанный диапазон превышает одну ячейку, возникает ошибка.|
||[getCount()](/javascript/api/excel/excel.commentcollection#getcount--)|Получает количество примечаний в коллекции.|
||[getItem(commentId: string)](/javascript/api/excel/excel.commentcollection#getitem-commentid-)|Получает примечание из коллекции на основе его идентификатора. Только для чтения.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentcollection#getitemat-index-)|Получает примечание из коллекции на основе его позиции.|
||[getItemByCell(cellAddress: Range \| string)](/javascript/api/excel/excel.commentcollection#getitembycell-celladdress-)|Получает примечание из указанной ячейки.|
||[getItemByReplyId(replyId: string)](/javascript/api/excel/excel.commentcollection#getitembyreplyid-replyid-)|Получает комментарий, к которому подключен данный ответ.|
||[items](/javascript/api/excel/excel.commentcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[комментментион](/javascript/api/excel/excel.commentmention)|[email](/javascript/api/excel/excel.commentmention#email)|Получает или задает адрес электронной почты объекта, который упоминается в примечании.|
||[id](/javascript/api/excel/excel.commentmention#id)|Получает или задает идентификатор объекта. Эта информация выравнивается по идентификатору в `CommentRichContent.richContent`файле.|
||[name](/javascript/api/excel/excel.commentmention#name)|Получает или задает имя объекта, упоминаемого в примечании.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[content](/javascript/api/excel/excel.commentreply#content)|Получает или задает содержимое ответа на примечание. Строка является обычным текстом.|
||[delete()](/javascript/api/excel/excel.commentreply#delete--)|Удаляет ответ на примечание.|
||[getLocation()](/javascript/api/excel/excel.commentreply#getlocation--)|Получает ячейку, в которой находится этот ответ на комментарий.|
||[getParentComment()](/javascript/api/excel/excel.commentreply#getparentcomment--)|Получает родительский комментарий для этого ответа.|
||[authorEmail](/javascript/api/excel/excel.commentreply#authoremail)|Получает электронную почту автора ответа на примечание.|
||[authorName](/javascript/api/excel/excel.commentreply#authorname)|Получает имя автора ответа на примечание.|
||[creationDate](/javascript/api/excel/excel.commentreply#creationdate)|Получает время создания ответа на примечание.|
||[id](/javascript/api/excel/excel.commentreply#id)|Представляет идентификатор ответа на примечание. Только для чтения.|
||[mentions](/javascript/api/excel/excel.commentreply#mentions)|Получает объекты (например, людей), которые упоминаются в комментариях.|
||[определяем](/javascript/api/excel/excel.commentreply#resolved)|Получает или задает состояние ответа на комментарий. Значение "true" означает, что ответ комментария находится в состоянии "разрешено".|
||[ричконтент](/javascript/api/excel/excel.commentreply#richcontent)|Получает форматированный текст комментария (например, упоминания в комментариях). Эта строка не предназначена для отображения конечным пользователям. Надстройка должна использовать эту надстройку только для анализа форматированного содержимого комментариев.|
||[Упдатементионс (Контентвисментионс: Excel. Комментричконтент)](/javascript/api/excel/excel.commentreply#updatementions-contentwithmentions-)|Обновляет содержимое комментария с помощью специально отформатированной строки и списка упоминаний.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[Добавить (контент: строка \| Комментричконтент, ContentType?: Excel. ContentType)](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|Создает ответ на примечание.|
||[getCount()](/javascript/api/excel/excel.commentreplycollection#getcount--)|Получает количество ответов на примечания в коллекции.|
||[getItem(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getitem-commentreplyid-)|Возвращает ответ на примечание, определенное по идентификатору. Только для чтения.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentreplycollection#getitemat-index-)|Возвращает ответ на примечание на основе его позиции в коллекции.|
||[items](/javascript/api/excel/excel.commentreplycollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[комментричконтент](/javascript/api/excel/excel.commentrichcontent)|[mentions](/javascript/api/excel/excel.commentrichcontent#mentions)|Массив, содержащий все сущности (например, люди), упомянутые в комментарии.|
||[ричконтент](/javascript/api/excel/excel.commentrichcontent#richcontent)||
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[enableFieldList](/javascript/api/excel/excel.pivotlayout#enablefieldlist)|Указывает, может ли список полей отображаться в пользовательском интерфейсе.|
||[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|Получает уникальную ячейку в сводной таблице на основе иерархии данных и элементов строк и столбцов соответствующих иерархий. Возвращаемая ячейка находится на пересечении указанной строки и столбца, содержащего данные из заданной иерархии. Этот метод является обратным вызову методов getPivotItems и getDataHierarchy для конкретной ячейки.|
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
|[Range](/javascript/api/excel/excel.range)|[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|Получает объект диапазона, содержащий базовую ячейку для переносимой ячейки. Возвращает ошибку, если применяется к диапазону с несколькими ячейками. Только для чтения.|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|Получает объект диапазона, содержащий базовую ячейку для переносимой ячейки. Только для чтения.|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|Получает объект range, содержащий диапазон переноса при вызове для базовой ячейки. Возвращает ошибку, если применяется к диапазону с несколькими ячейками. Только для чтения.|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|Получает объект range, содержащий диапазон переноса при вызове для базовой ячейки. Только для чтения.|
||[Group (Граупоптион: Excel. Граупоптион)](/javascript/api/excel/excel.range#group-groupoption-)|Группирует столбцы и строки для структуры.|
||[Хидеграупдетаилс (Граупоптион: Excel. Граупоптион)](/javascript/api/excel/excel.range#hidegroupdetails-groupoption-)|Скрытие сведений о группе строк или столбцов.|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|Указывает, есть ли во всех ячейках граница переноса.|
||[height](/javascript/api/excel/excel.range#height)|Возвращает расстояние в пунктах (для масштаба 100 %) от верхнего до нижнего края диапазона. Только для чтения.|
||[left](/javascript/api/excel/excel.range#left)|Возвращает расстояние в пунктах (для масштаба 100 %) от левого края листа до левого края диапазона. Только для чтения.|
||[саведасаррай](/javascript/api/excel/excel.range#savedasarray)|Указывает, следует ли сохранять все ячейки в виде формулы массива.|
||[top](/javascript/api/excel/excel.range#top)|Возвращает расстояние в пунктах для масштаба 100 % от верхнего края листа до верхнего края диапазона. Только для чтения.|
||[width](/javascript/api/excel/excel.range#width)|Возвращает расстояние в пунктах (для масштаба 100 %) от левого до правого края диапазона. Только для чтения.|
||[Шовграупдетаилс (Граупоптион: Excel. Граупоптион)](/javascript/api/excel/excel.range#showgroupdetails-groupoption-)|Отображение сведений о группе строк или столбцов.|
||[Разгруппировать (Граупоптион: Excel. Граупоптион)](/javascript/api/excel/excel.range#ungroup-groupoption-)|Разгруппирование столбцов и строк для структуры.|
|[RangeFormat](/javascript/api/excel/excel.rangeformat)|[Аджустиндент (Amount: число)](/javascript/api/excel/excel.rangeformat#adjustindent-amount-)|Настраивает отступ для форматирования диапазона. Значение отступа лежит в диапазоне от 0 до 250.|
|[Shape](/javascript/api/excel/excel.shape)|[copyTo(destinationSheet?: Worksheet \| string)](/javascript/api/excel/excel.shape#copyto-destinationsheet-)|Копирует и вставляет объект Shape.|
||[placement](/javascript/api/excel/excel.shape#placement)|Представляет способ прикрепления объекта к ячейкам под ним.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|Создает изображение SVG (масштабируемая векторная графика) из строки XML и добавляет его на лист. Возвращает объект Shape, представляющий новое изображение.|
|[Slicer](/javascript/api/excel/excel.slicer)|[caption](/javascript/api/excel/excel.slicer#caption)|Представляет подпись среза.|
||[clearFilters()](/javascript/api/excel/excel.slicer#clearfilters--)|Удаляет все фильтры, примененные к срезу.|
||[delete()](/javascript/api/excel/excel.slicer#delete--)|Удаляет срез.|
||[getSelectedItems()](/javascript/api/excel/excel.slicer#getselecteditems--)|Возвращает массив имен выбранных ключей элементов. Только для чтения.|
||[height](/javascript/api/excel/excel.slicer#height)|Представляет высоту среза (в пунктах).|
||[left](/javascript/api/excel/excel.slicer#left)|Представляет расстояние в пунктах от левого края среза до левого края листа.|
||[name](/javascript/api/excel/excel.slicer#name)|Представляет имя среза.|
||[nameInFormula](/javascript/api/excel/excel.slicer#nameinformula)|Представляет имя среза, используемое в формуле.|
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
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|Изменяет таблицу для использования стиля таблицы по умолчанию.|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|Возникает, если применен фильтр к указанной таблице.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|Возникает, если применен фильтр к любой таблице в книге или листе.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|Представляет идентификатор таблицы, в которой применяется фильтр.|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|Представляет тип события. Дополнительные сведения см. в статье Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|Представляет идентификатор листа, содержащего таблицу.|
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
|[Workbook](/javascript/api/excel/excel.workbook)|[close(closeBehavior?: Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close-closebehavior-)|Закрывает текущую книгу.|
||[getActiveSlicer()](/javascript/api/excel/excel.workbook#getactiveslicer--)|Получает текущий активный срез в книге. Если активного среза нет, создается `ItemNotFound` исключение.|
||[getActiveSlicerOrNullObject()](/javascript/api/excel/excel.workbook#getactiveslicerornullobject--)|Получает текущий активный срез в книге. Если активный срез отсутствует, возвращается пустой объект.|
||[comments](/javascript/api/excel/excel.workbook#comments)|Представляет коллекцию примечаний, связанных с книгой. Только для чтения.|
||[pivotTableStyles](/javascript/api/excel/excel.workbook#pivottablestyles)|Представляет коллекцию объектов PivotTableStyles, связанных с книгой. Только для чтения.|
||[slicerStyles](/javascript/api/excel/excel.workbook#slicerstyles)|Представляет коллекцию объектов SlicerStyles, связанных с книгой. Только для чтения.|
||[slicers](/javascript/api/excel/excel.workbook#slicers)|Представляет коллекцию срезов, связанных с книгой. Только для чтения.|
||[tableStyles](/javascript/api/excel/excel.workbook#tablestyles)|Представляет коллекцию объектов TableStyles, связанных с книгой. Только для чтения.|
||[timelineStyles](/javascript/api/excel/excel.workbook#timelinestyles)|Представляет коллекцию объектов TimelineStyles, связанных с книгой. Только для чтения.|
||[save(saveBehavior?: Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save-savebehavior-)|Сохраняет текущую книгу.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|Значение true, если в книге используется система дат 1904.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[comments](/javascript/api/excel/excel.worksheet#comments)|Возвращает коллекцию всех объектов Comments на листе. Только для чтения.|
||[onColumnSorted](/javascript/api/excel/excel.worksheet#oncolumnsorted)|Возникает при сортировке одного или нескольких столбцов. Это происходит в результате операции сортировки слева направо.|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|Возникает, если применен фильтр к указанному листу.|
||[онровхидденчанжед](/javascript/api/excel/excel.worksheet#onrowhiddenchanged)|Происходит при изменении скрытого состояния одной или нескольких строк на определенном листе.|
||[onRowSorted](/javascript/api/excel/excel.worksheet#onrowsorted)|Возникает при сортировке одной или нескольких строк. Это происходит в результате операции сортировки "начало сверху вниз".|
||[onSingleClicked](/javascript/api/excel/excel.worksheet#onsingleclicked)|Возникает, когда происходит щелчок левой кнопкой мыши или нажатие на листе. Это событие не будет инициировано при нажатии следующих условий:|
||[slicers](/javascript/api/excel/excel.worksheet#slicers)|Возвращает коллекцию срезов, имеющихся на листе. Только для чтения.|
||[Шоваутлинелевелс (Ровлевелс: число, Колумнлевелс: число)](/javascript/api/excel/excel.worksheet#showoutlinelevels-rowlevels--columnlevels-)|Отображает группы строк или столбцов по их уровням структуры.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Вставляет указанные листы книги в текущую книгу.|
||[onColumnSorted](/javascript/api/excel/excel.worksheetcollection#oncolumnsorted)|Возникает при сортировке одного или нескольких столбцов. Это происходит в результате операции сортировки слева направо.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|Возникает при применении любого фильтра листа в книге.|
||[онровхидденчанжед](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)|Происходит при изменении скрытого состояния одной или нескольких строк на определенном листе.|
||[onRowSorted](/javascript/api/excel/excel.worksheetcollection#onrowsorted)|Возникает при сортировке одной или нескольких строк. Это происходит в результате операции сортировки "начало сверху вниз".|
||[onSingleClicked](/javascript/api/excel/excel.worksheetcollection#onsingleclicked)|Возникает, когда в коллекции листа происходит операция с нажатием и нажатием левой кнопкой мыши. Это событие не будет инициировано при нажатии следующих условий:|
|[WorksheetColumnSortedEventArgs](/javascript/api/excel/excel.worksheetcolumnsortedeventargs)|[address](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#address)|Получает адрес диапазона, представляющий отсортированные области конкретного листа. Возвращаются только столбцы, измененные в результате операции сортировки.|
||[source](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#source)|Получает источник события. Дополнительные сведения см. в статье Excel.EventSource.|
||[type](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#type)|Получает тип события. Дополнительные сведения см. в статье Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#worksheetid)|Получает идентификатор листа, в котором выполнена сортировка.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Представляет тип события. Дополнительные сведения см. в статье Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|Представляет идентификатор листа, в котором применен фильтр.|
|[воркшитровхидденчанжедевентаргс](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[address](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#address)|Получает адрес диапазона, представляющий измененную область конкретного листа.|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#changetype)|Получает тип изменения, которое представляет способ запуска события. Для `Excel.RowHiddenChangeType` получения дополнительных сведений см.|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#source)|Получает источник события. Дополнительные сведения см. в статье Excel.EventSource.|
||[type](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#type)|Получает тип события. Дополнительные сведения см. в статье Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#worksheetid)|Получает идентификатор листа, в котором изменены данные.|
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

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-preview)
- [Наборы обязательных элементов API JavaScript для Excel](./excel-api-requirement-sets.md)
