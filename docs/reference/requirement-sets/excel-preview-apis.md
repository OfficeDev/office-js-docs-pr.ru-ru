---
title: Предварительные версии API JavaScript для Excel
description: Сведения о предстоящих API JavaScript для Excel
ms.date: 07/25/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 200b187b059c1b03ae3713b5afa11b2152aba0da
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064853"
---
# <a name="excel-javascript-preview-apis"></a>Предварительные версии API JavaScript для Excel

Новые API JavaScript для Excel сначала выпускаются в "предварительной версии", а затем становятся частью определенного нумерованного набора обязательных элементов после выполнения достаточного тестирования и получения отзывов пользователей.

В первой таблице представлен краткий обзор API, а в последующей таблице приведен подробный список.

> [!NOTE]
> API предварительной версии могут быть изменены и не предназначены для использования в рабочей среде. Рекомендуется использовать их только в тестовой среде и среде разработки. Не используйте API предварительной версии в рабочей среде или в важных деловых документах.
>
> Чтобы использовать API предварительной версии, нужно сослаться на **бета-версию** библиотеки в сети CDN (https://appsforoffice.microsoft.com/lib/beta/hosted/office.js), и также может потребоваться присоединение к программе предварительной оценки Office для получения последней сборки Office.

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
||[delete()](/javascript/api/excel/excel.comment#delete--)|Удаляет цепочку примечаний.|
||[getLocation()](/javascript/api/excel/excel.comment#getlocation--)|Получает ячейку, в которой находится этот комментарий.|
||[authorEmail](/javascript/api/excel/excel.comment#authoremail)|Получает электронную почту автора примечания.|
||[authorName](/javascript/api/excel/excel.comment#authorname)|Получает имя автора примечания.|
||[creationDate](/javascript/api/excel/excel.comment#creationdate)|Получает время создания примечания. Возвращает значение null, если примечание было преобразовано из заметки, так как у примечания нет даты создания.|
||[id](/javascript/api/excel/excel.comment#id)|Представляет идентификатор примечания. Только для чтения.|
||[replies](/javascript/api/excel/excel.comment#replies)|Представляет коллекцию объектов ответов, связанных с примечанием. Только для чтения.|
||[Set (Properties: Excel. Comment)](/javascript/api/excel/excel.comment#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Комментупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.comment#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[CommentCollection](/javascript/api/excel/excel.commentcollection)|[add(content: string, cellAddress: Range \| string, contentType?: "Plain")](/javascript/api/excel/excel.commentcollection#add-content--celladdress--contenttype-)|Создает новый комментарий (поток комментариев) с заданным содержимым в заданной ячейке. Если `InvalidArgument` указанный диапазон превышает одну ячейку, возникает ошибка.|
||[add(content: string, cellAddress: Range \| string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentcollection#add-content--celladdress--contenttype-)|Создает новый комментарий (поток комментариев) с заданным содержимым в заданной ячейке. Если `InvalidArgument` указанный диапазон превышает одну ячейку, возникает ошибка.|
||[getCount()](/javascript/api/excel/excel.commentcollection#getcount--)|Получает количество примечаний в коллекции.|
||[getItem(commentId: string)](/javascript/api/excel/excel.commentcollection#getitem-commentid-)|Получает примечание из коллекции на основе его идентификатора. Только для чтения.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentcollection#getitemat-index-)|Получает примечание из коллекции на основе его позиции.|
||[getItemByCell(cellAddress: Range \| string)](/javascript/api/excel/excel.commentcollection#getitembycell-celladdress-)|Получает примечание из указанной ячейки.|
||[getItemByReplyId(replyId: string)](/javascript/api/excel/excel.commentcollection#getitembyreplyid-replyid-)|Возвращает примечание, связанное с идентификатором ответа в коллекции.|
||[items](/javascript/api/excel/excel.commentcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Комментколлектиондата](/javascript/api/excel/excel.commentcollectiondata)|[items](/javascript/api/excel/excel.commentcollectiondata#items)||
|[Комментколлектионлоадоптионс](/javascript/api/excel/excel.commentcollectionloadoptions)|[$all](/javascript/api/excel/excel.commentcollectionloadoptions#$all)||
||[authorEmail](/javascript/api/excel/excel.commentcollectionloadoptions#authoremail)|Для каждого элемента в коллекции: получает электронную почту автора комментария.|
||[authorName](/javascript/api/excel/excel.commentcollectionloadoptions#authorname)|Для каждого элемента в коллекции: получает имя автора комментария.|
||[content](/javascript/api/excel/excel.commentcollectionloadoptions#content)|Для каждого элемента в коллекции: Получает или задает содержимое комментария. Строка является обычным текстом.|
||[creationDate](/javascript/api/excel/excel.commentcollectionloadoptions#creationdate)|Для каждого элемента в коллекции: получает время создания комментария. Возвращает значение null, если примечание было преобразовано из заметки, так как у примечания нет даты создания.|
||[id](/javascript/api/excel/excel.commentcollectionloadoptions#id)|Для каждого элемента в коллекции: представляет идентификатор комментария. Только для чтения.|
|[Комментколлектионупдатедата](/javascript/api/excel/excel.commentcollectionupdatedata)|[items](/javascript/api/excel/excel.commentcollectionupdatedata#items)||
|[Комментдата](/javascript/api/excel/excel.commentdata)|[authorEmail](/javascript/api/excel/excel.commentdata#authoremail)|Получает электронную почту автора примечания.|
||[authorName](/javascript/api/excel/excel.commentdata#authorname)|Получает имя автора примечания.|
||[content](/javascript/api/excel/excel.commentdata#content)|Получает или задает содержимое примечания. Строка является обычным текстом.|
||[creationDate](/javascript/api/excel/excel.commentdata#creationdate)|Получает время создания примечания. Возвращает значение null, если примечание было преобразовано из заметки, так как у примечания нет даты создания.|
||[id](/javascript/api/excel/excel.commentdata#id)|Представляет идентификатор примечания. Только для чтения.|
||[replies](/javascript/api/excel/excel.commentdata#replies)|Представляет коллекцию объектов ответов, связанных с примечанием. Только для чтения.|
|[Комментлоадоптионс](/javascript/api/excel/excel.commentloadoptions)|[$all](/javascript/api/excel/excel.commentloadoptions#$all)||
||[authorEmail](/javascript/api/excel/excel.commentloadoptions#authoremail)|Получает электронную почту автора примечания.|
||[authorName](/javascript/api/excel/excel.commentloadoptions#authorname)|Получает имя автора примечания.|
||[content](/javascript/api/excel/excel.commentloadoptions#content)|Получает или задает содержимое примечания. Строка является обычным текстом.|
||[creationDate](/javascript/api/excel/excel.commentloadoptions#creationdate)|Получает время создания примечания. Возвращает значение null, если примечание было преобразовано из заметки, так как у примечания нет даты создания.|
||[id](/javascript/api/excel/excel.commentloadoptions#id)|Представляет идентификатор примечания. Только для чтения.|
|[CommentReply](/javascript/api/excel/excel.commentreply)|[content](/javascript/api/excel/excel.commentreply#content)|Получает или задает содержимое ответа на примечание. Строка является обычным текстом.|
||[delete()](/javascript/api/excel/excel.commentreply#delete--)|Удаляет ответ на примечание.|
||[getLocation()](/javascript/api/excel/excel.commentreply#getlocation--)|Получает ячейку, в которой находится этот ответ на комментарий.|
||[getParentComment()](/javascript/api/excel/excel.commentreply#getparentcomment--)|Получает родительский комментарий для этого ответа.|
||[authorEmail](/javascript/api/excel/excel.commentreply#authoremail)|Получает электронную почту автора ответа на примечание.|
||[authorName](/javascript/api/excel/excel.commentreply#authorname)|Получает имя автора ответа на примечание.|
||[creationDate](/javascript/api/excel/excel.commentreply#creationdate)|Получает время создания ответа на примечание.|
||[id](/javascript/api/excel/excel.commentreply#id)|Представляет идентификатор ответа на примечание. Только для чтения.|
||[Set (Properties: Excel. Комментрепли)](/javascript/api/excel/excel.commentreply#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Комментреплюпдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.commentreply#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[CommentReplyCollection](/javascript/api/excel/excel.commentreplycollection)|[add(content: string, contentType?: "Plain")](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|Создает ответ на примечание.|
||[add(content: string, contentType?: Excel.ContentType)](/javascript/api/excel/excel.commentreplycollection#add-content--contenttype-)|Создает ответ на примечание.|
||[getCount()](/javascript/api/excel/excel.commentreplycollection#getcount--)|Получает количество ответов на примечания в коллекции.|
||[getItem(commentReplyId: string)](/javascript/api/excel/excel.commentreplycollection#getitem-commentreplyid-)|Возвращает ответ на примечание, определенное по идентификатору. Только для чтения.|
||[getItemAt(index: number)](/javascript/api/excel/excel.commentreplycollection#getitemat-index-)|Возвращает ответ на примечание на основе его позиции в коллекции.|
||[items](/javascript/api/excel/excel.commentreplycollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Комментрепликоллектиондата](/javascript/api/excel/excel.commentreplycollectiondata)|[items](/javascript/api/excel/excel.commentreplycollectiondata#items)||
|[Комментрепликоллектионлоадоптионс](/javascript/api/excel/excel.commentreplycollectionloadoptions)|[$all](/javascript/api/excel/excel.commentreplycollectionloadoptions#$all)||
||[authorEmail](/javascript/api/excel/excel.commentreplycollectionloadoptions#authoremail)|Для каждого элемента в коллекции: получает сообщение электронной почты автора комментария.|
||[authorName](/javascript/api/excel/excel.commentreplycollectionloadoptions#authorname)|Для каждого элемента в коллекции: получает имя автора ответа на комментарий.|
||[content](/javascript/api/excel/excel.commentreplycollectionloadoptions#content)|Для каждого элемента в коллекции: Получает или задает содержимое ответа на комментарий. Строка является обычным текстом.|
||[creationDate](/javascript/api/excel/excel.commentreplycollectionloadoptions#creationdate)|Для каждого элемента в коллекции: получает время создания ответа на комментарий.|
||[id](/javascript/api/excel/excel.commentreplycollectionloadoptions#id)|Для каждого элемента в коллекции — представляет идентификатор ответа комментария. Только для чтения.|
|[Комментрепликоллектионупдатедата](/javascript/api/excel/excel.commentreplycollectionupdatedata)|[items](/javascript/api/excel/excel.commentreplycollectionupdatedata#items)||
|[Комментреплидата](/javascript/api/excel/excel.commentreplydata)|[authorEmail](/javascript/api/excel/excel.commentreplydata#authoremail)|Получает электронную почту автора ответа на примечание.|
||[authorName](/javascript/api/excel/excel.commentreplydata#authorname)|Получает имя автора ответа на примечание.|
||[content](/javascript/api/excel/excel.commentreplydata#content)|Получает или задает содержимое ответа на примечание. Строка является обычным текстом.|
||[creationDate](/javascript/api/excel/excel.commentreplydata#creationdate)|Получает время создания ответа на примечание.|
||[id](/javascript/api/excel/excel.commentreplydata#id)|Представляет идентификатор ответа на примечание. Только для чтения.|
|[Комментреплилоадоптионс](/javascript/api/excel/excel.commentreplyloadoptions)|[$all](/javascript/api/excel/excel.commentreplyloadoptions#$all)||
||[authorEmail](/javascript/api/excel/excel.commentreplyloadoptions#authoremail)|Получает электронную почту автора ответа на примечание.|
||[authorName](/javascript/api/excel/excel.commentreplyloadoptions#authorname)|Получает имя автора ответа на примечание.|
||[content](/javascript/api/excel/excel.commentreplyloadoptions#content)|Получает или задает содержимое ответа на примечание. Строка является обычным текстом.|
||[creationDate](/javascript/api/excel/excel.commentreplyloadoptions#creationdate)|Получает время создания ответа на примечание.|
||[id](/javascript/api/excel/excel.commentreplyloadoptions#id)|Представляет идентификатор ответа на примечание. Только для чтения.|
|[Комментреплюпдатедата](/javascript/api/excel/excel.commentreplyupdatedata)|[content](/javascript/api/excel/excel.commentreplyupdatedata#content)|Получает или задает содержимое ответа на примечание. Строка является обычным текстом.|
|[Комментупдатедата](/javascript/api/excel/excel.commentupdatedata)|[content](/javascript/api/excel/excel.commentupdatedata#content)|Получает или задает содержимое примечания. Строка является обычным текстом.|
|[Граупшапеколлектионлоадоптионс](/javascript/api/excel/excel.groupshapecollectionloadoptions)|[placement](/javascript/api/excel/excel.groupshapecollectionloadoptions#placement)|Для каждого элемента в коллекции: указывает, как объект присоединен к ячейкам, расположенным под ним.|
|[PivotLayout](/javascript/api/excel/excel.pivotlayout)|[enableFieldList](/javascript/api/excel/excel.pivotlayout#enablefieldlist)|Указывает, может ли список полей отображаться в пользовательском интерфейсе.|
||[getCell(dataHierarchy: DataPivotHierarchy \| string, rowItems: Array<PivotItem \| string>, columnItems: Array<PivotItem \| string>)](/javascript/api/excel/excel.pivotlayout#getcell-datahierarchy--rowitems--columnitems-)|Получает уникальную ячейку в сводной таблице на основе иерархии данных и элементов строк и столбцов соответствующих иерархий. Возвращаемая ячейка находится на пересечении указанной строки и столбца, содержащего данные из заданной иерархии. Этот метод является обратным вызову методов getPivotItems и getDataHierarchy для конкретной ячейки.|
|[Пивотлайаутдата](/javascript/api/excel/excel.pivotlayoutdata)|[enableFieldList](/javascript/api/excel/excel.pivotlayoutdata#enablefieldlist)|Указывает, может ли список полей отображаться в пользовательском интерфейсе.|
|[Пивотлайаутлоадоптионс](/javascript/api/excel/excel.pivotlayoutloadoptions)|[enableFieldList](/javascript/api/excel/excel.pivotlayoutloadoptions#enablefieldlist)|Указывает, может ли список полей отображаться в пользовательском интерфейсе.|
|[Пивотлайаутупдатедата](/javascript/api/excel/excel.pivotlayoutupdatedata)|[enableFieldList](/javascript/api/excel/excel.pivotlayoutupdatedata#enablefieldlist)|Указывает, может ли список полей отображаться в пользовательском интерфейсе.|
|[PivotTableStyle](/javascript/api/excel/excel.pivottablestyle)|[delete()](/javascript/api/excel/excel.pivottablestyle#delete--)|Удаляет объект PivotTableStyle.|
||[duplicate()](/javascript/api/excel/excel.pivottablestyle#duplicate--)|Создает дубликат объекта PivotTableStyle с копиями всех элементов стиля.|
||[name](/javascript/api/excel/excel.pivottablestyle#name)|Получает имя объекта PivotTableStyle.|
||[readOnly](/javascript/api/excel/excel.pivottablestyle#readonly)|Указывает, предназначен ли объект PivotTableStyle только для чтения. Только для чтения.|
||[Set (Properties: Excel. Пивоттаблестиле)](/javascript/api/excel/excel.pivottablestyle#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Пивоттаблестилеупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.pivottablestyle#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[PivotTableStyleCollection](/javascript/api/excel/excel.pivottablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.pivottablestylecollection#add-name--makeuniquename-)|Создает пустой объект PivotTableStyle с указанным именем.|
||[getCount()](/javascript/api/excel/excel.pivottablestylecollection#getcount--)|Получает количество стилей сводных таблиц в коллекции.|
||[getDefault()](/javascript/api/excel/excel.pivottablestylecollection#getdefault--)|Получает используемый по умолчанию объект PivotTableStyle для области родительского объекта.|
||[getItem(name: string)](/javascript/api/excel/excel.pivottablestylecollection#getitem-name-)|Получает объект PivotTableStyle по имени.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.pivottablestylecollection#getitemornullobject-name-)|Получает объект PivotTableStyle по имени. Если объект PivotTableStyle не существует, возвращает пустой объект.|
||[items](/javascript/api/excel/excel.pivottablestylecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[setDefault(newDefaultStyle: PivotTableStyle \| string)](/javascript/api/excel/excel.pivottablestylecollection#setdefault-newdefaultstyle-)|Задает объект PivotTableStyle, используемый по умолчанию в области родительского объекта.|
|[Пивоттаблестилеколлектиондата](/javascript/api/excel/excel.pivottablestylecollectiondata)|[items](/javascript/api/excel/excel.pivottablestylecollectiondata#items)||
|[Пивоттаблестилеколлектионлоадоптионс](/javascript/api/excel/excel.pivottablestylecollectionloadoptions)|[$all](/javascript/api/excel/excel.pivottablestylecollectionloadoptions#$all)||
||[name](/javascript/api/excel/excel.pivottablestylecollectionloadoptions#name)|Для каждого элемента в коллекции: получает имя Пивоттаблестиле.|
||[readOnly](/javascript/api/excel/excel.pivottablestylecollectionloadoptions#readonly)|Для каждого элемента в коллекции: указывает, является ли этот объект Пивоттаблестиле доступен только для чтения. Только для чтения.|
|[Пивоттаблестилеколлектионупдатедата](/javascript/api/excel/excel.pivottablestylecollectionupdatedata)|[items](/javascript/api/excel/excel.pivottablestylecollectionupdatedata#items)||
|[Пивоттаблестиледата](/javascript/api/excel/excel.pivottablestyledata)|[name](/javascript/api/excel/excel.pivottablestyledata#name)|Получает имя объекта PivotTableStyle.|
||[readOnly](/javascript/api/excel/excel.pivottablestyledata#readonly)|Указывает, предназначен ли объект PivotTableStyle только для чтения. Только для чтения.|
|[Пивоттаблестилелоадоптионс](/javascript/api/excel/excel.pivottablestyleloadoptions)|[$all](/javascript/api/excel/excel.pivottablestyleloadoptions#$all)||
||[name](/javascript/api/excel/excel.pivottablestyleloadoptions#name)|Получает имя объекта PivotTableStyle.|
||[readOnly](/javascript/api/excel/excel.pivottablestyleloadoptions#readonly)|Указывает, предназначен ли объект PivotTableStyle только для чтения. Только для чтения.|
|[Пивоттаблестилеупдатедата](/javascript/api/excel/excel.pivottablestyleupdatedata)|[name](/javascript/api/excel/excel.pivottablestyleupdatedata#name)|Получает имя объекта PivotTableStyle.|
|[Range](/javascript/api/excel/excel.range)|[getSpillParent()](/javascript/api/excel/excel.range#getspillparent--)|Получает объект диапазона, содержащий базовую ячейку для переносимой ячейки. Возвращает ошибку, если применяется к диапазону с несколькими ячейками. Только для чтения.|
||[getSpillParentOrNullObject()](/javascript/api/excel/excel.range#getspillparentornullobject--)|Получает объект диапазона, содержащий базовую ячейку для переносимой ячейки. Только для чтения.|
||[getSpillingToRange()](/javascript/api/excel/excel.range#getspillingtorange--)|Получает объект range, содержащий диапазон переноса при вызове для базовой ячейки. Возвращает ошибку, если применяется к диапазону с несколькими ячейками. Только для чтения.|
||[getSpillingToRangeOrNullObject()](/javascript/api/excel/excel.range#getspillingtorangeornullobject--)|Получает объект range, содержащий диапазон переноса при вызове для базовой ячейки. Только для чтения.|
||[hasSpill](/javascript/api/excel/excel.range#hasspill)|Указывает, есть ли во всех ячейках граница переноса.|
||[height](/javascript/api/excel/excel.range#height)|Возвращает расстояние в пунктах (для масштаба 100 %) от верхнего до нижнего края диапазона. Только для чтения.|
||[left](/javascript/api/excel/excel.range#left)|Возвращает расстояние в пунктах (для масштаба 100 %) от левого края листа до левого края диапазона. Только для чтения.|
||[Саведасаррай](/javascript/api/excel/excel.range#savedasarray)|Указывает, следует ли сохранять все ячейки в виде формулы массива.|
||[top](/javascript/api/excel/excel.range#top)|Возвращает расстояние в пунктах для масштаба 100 % от верхнего края листа до верхнего края диапазона. Только для чтения.|
||[width](/javascript/api/excel/excel.range#width)|Возвращает расстояние в пунктах (для масштаба 100 %) от левого до правого края диапазона. Только для чтения.|
|[Ранжеколлектионлоадоптионс](/javascript/api/excel/excel.rangecollectionloadoptions)|[hasSpill](/javascript/api/excel/excel.rangecollectionloadoptions#hasspill)|Для каждого элемента в коллекции: указывает, имеет ли вся ячейка проspill границу.|
||[height](/javascript/api/excel/excel.rangecollectionloadoptions#height)|Для каждого элемента в коллекции: Возвращает расстояние в пунктах для масштаба 100% от верхнего края диапазона до нижнего края диапазона. Только для чтения.|
||[left](/javascript/api/excel/excel.rangecollectionloadoptions#left)|Для каждого элемента в коллекции: Возвращает расстояние в пунктах для масштаба 100% от левого края листа до левого края диапазона. Только для чтения.|
||[Саведасаррай](/javascript/api/excel/excel.rangecollectionloadoptions#savedasarray)|Для каждого элемента в коллекции: указывает, будут ли все ячейки сохраняться в виде формулы массива.|
||[top](/javascript/api/excel/excel.rangecollectionloadoptions#top)|Для каждого элемента в коллекции: Возвращает расстояние в пунктах для масштаба 100% от верхнего края листа до верхнего края диапазона. Только для чтения.|
||[width](/javascript/api/excel/excel.rangecollectionloadoptions#width)|Для каждого элемента в коллекции: Возвращает расстояние в пунктах для масштаба 100% от левого края диапазона до правого края диапазона. Только для чтения.|
|[Ранжедата](/javascript/api/excel/excel.rangedata)|[hasSpill](/javascript/api/excel/excel.rangedata#hasspill)|Указывает, есть ли во всех ячейках граница переноса.|
||[height](/javascript/api/excel/excel.rangedata#height)|Возвращает расстояние в пунктах (для масштаба 100 %) от верхнего до нижнего края диапазона. Только для чтения.|
||[left](/javascript/api/excel/excel.rangedata#left)|Возвращает расстояние в пунктах (для масштаба 100 %) от левого края листа до левого края диапазона. Только для чтения.|
||[Саведасаррай](/javascript/api/excel/excel.rangedata#savedasarray)|Указывает, следует ли сохранять все ячейки в виде формулы массива.|
||[top](/javascript/api/excel/excel.rangedata#top)|Возвращает расстояние в пунктах для масштаба 100 % от верхнего края листа до верхнего края диапазона. Только для чтения.|
||[width](/javascript/api/excel/excel.rangedata#width)|Возвращает расстояние в пунктах (для масштаба 100 %) от левого до правого края диапазона. Только для чтения.|
|[Ранжелоадоптионс](/javascript/api/excel/excel.rangeloadoptions)|[hasSpill](/javascript/api/excel/excel.rangeloadoptions#hasspill)|Указывает, есть ли во всех ячейках граница переноса.|
||[height](/javascript/api/excel/excel.rangeloadoptions#height)|Возвращает расстояние в пунктах (для масштаба 100 %) от верхнего до нижнего края диапазона. Только для чтения.|
||[left](/javascript/api/excel/excel.rangeloadoptions#left)|Возвращает расстояние в пунктах (для масштаба 100 %) от левого края листа до левого края диапазона. Только для чтения.|
||[Саведасаррай](/javascript/api/excel/excel.rangeloadoptions#savedasarray)|Указывает, следует ли сохранять все ячейки в виде формулы массива.|
||[top](/javascript/api/excel/excel.rangeloadoptions#top)|Возвращает расстояние в пунктах для масштаба 100 % от верхнего края листа до верхнего края диапазона. Только для чтения.|
||[width](/javascript/api/excel/excel.rangeloadoptions#width)|Возвращает расстояние в пунктах (для масштаба 100 %) от левого до правого края диапазона. Только для чтения.|
|[Shape](/javascript/api/excel/excel.shape)|[copyTo(destinationSheet?: Worksheet \| string)](/javascript/api/excel/excel.shape#copyto-destinationsheet-)|Копирует и вставляет объект Shape.|
||[placement](/javascript/api/excel/excel.shape#placement)|Представляет способ прикрепления объекта к ячейкам под ним.|
|[ShapeCollection](/javascript/api/excel/excel.shapecollection)|[addSvg(xml: string)](/javascript/api/excel/excel.shapecollection#addsvg-xml-)|Создает изображение SVG (масштабируемая векторная графика) из строки XML и добавляет его на лист. Возвращает объект Shape, представляющий новое изображение.|
|[Шапеколлектионлоадоптионс](/javascript/api/excel/excel.shapecollectionloadoptions)|[placement](/javascript/api/excel/excel.shapecollectionloadoptions#placement)|Для каждого элемента в коллекции: указывает, как объект присоединен к ячейкам, расположенным под ним.|
|[Шапедата](/javascript/api/excel/excel.shapedata)|[placement](/javascript/api/excel/excel.shapedata#placement)|Представляет способ прикрепления объекта к ячейкам под ним.|
|[Шапелоадоптионс](/javascript/api/excel/excel.shapeloadoptions)|[placement](/javascript/api/excel/excel.shapeloadoptions#placement)|Представляет способ прикрепления объекта к ячейкам под ним.|
|[Шапеупдатедата](/javascript/api/excel/excel.shapeupdatedata)|[placement](/javascript/api/excel/excel.shapeupdatedata#placement)|Представляет способ прикрепления объекта к ячейкам под ним.|
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
||[selectItems(items?: string[])](/javascript/api/excel/excel.slicer#selectitems-items-)|Выделяет элементы среза на основе их ключей. Предыдущее выделение очищается.|
||[Set (Properties: Excel. срез)](/javascript/api/excel/excel.slicer#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Слицерупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.slicer#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
||[sortBy](/javascript/api/excel/excel.slicer#sortby)|Представляет порядок сортировки элементов в срезе. Возможные значения: DataSourceOrder, Ascending, Descending.|
||[style](/javascript/api/excel/excel.slicer#style)|Постоянное значение, представляющее стиль среза. Возможные значения: "SlicerStyleLight1", "SlicerStyleLight6", "TableStyleOther1", "TableStyleOther2", "SlicerStyleDark1" и "SlicerStyleDark6". Также можно указать настраиваемый пользовательский стиль, имеющийся в книге.|
||[top](/javascript/api/excel/excel.slicer#top)|Представляет расстояние в пунктах от верхнего края среза до верхнего края листа.|
||[width](/javascript/api/excel/excel.slicer#width)|Представляет ширину среза (в пунктах).|
|[SlicerCollection](/javascript/api/excel/excel.slicercollection)|[add(slicerSource: string \| PivotTable \| Table, sourceField: string \| PivotField \| number \| TableColumn, slicerDestination?: string \| Worksheet)](/javascript/api/excel/excel.slicercollection#add-slicersource--sourcefield--slicerdestination-)|Добавляет новый срез в книгу.|
||[getCount()](/javascript/api/excel/excel.slicercollection#getcount--)|Возвращает количество срезов в коллекции.|
||[getItem(key: string)](/javascript/api/excel/excel.slicercollection#getitem-key-)|Получает объект slicer по его имени или ИД.|
||[getItemAt(index: number)](/javascript/api/excel/excel.slicercollection#getitemat-index-)|Получает срез на основе его позиции в коллекции.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.slicercollection#getitemornullobject-key-)|Получает срез по его имени или ИД. Если срез не существует, возвращает пустой объект.|
||[items](/javascript/api/excel/excel.slicercollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Слицерколлектиондата](/javascript/api/excel/excel.slicercollectiondata)|[items](/javascript/api/excel/excel.slicercollectiondata#items)||
|[Слицерколлектионлоадоптионс](/javascript/api/excel/excel.slicercollectionloadoptions)|[$all](/javascript/api/excel/excel.slicercollectionloadoptions#$all)||
||[caption](/javascript/api/excel/excel.slicercollectionloadoptions#caption)|Для каждого элемента в коллекции: представляет подпись среза.|
||[height](/javascript/api/excel/excel.slicercollectionloadoptions#height)|Для каждого элемента в коллекции: представляет высоту среза в пунктах.|
||[id](/javascript/api/excel/excel.slicercollectionloadoptions#id)|Для каждого элемента в коллекции: представляет уникальный идентификатор среза. Только для чтения.|
||[isFilterCleared](/javascript/api/excel/excel.slicercollectionloadoptions#isfiltercleared)|Для каждого элемента в коллекции: true, если все фильтры, примененные к срезу, будут очищены.|
||[left](/javascript/api/excel/excel.slicercollectionloadoptions#left)|Для каждого элемента в коллекции — представляет расстояние в пунктах от левого края среза до левого края листа.|
||[name](/javascript/api/excel/excel.slicercollectionloadoptions#name)|Для каждого элемента в коллекции: представляет имя среза.|
||[nameInFormula](/javascript/api/excel/excel.slicercollectionloadoptions#nameinformula)|Для каждого элемента в коллекции: представляет имя среза, используемое в формуле.|
||[sortBy](/javascript/api/excel/excel.slicercollectionloadoptions#sortby)|Для каждого элемента в коллекции: представляет порядок сортировки элементов в срезе. Возможные значения: DataSourceOrder, Ascending, Descending.|
||[style](/javascript/api/excel/excel.slicercollectionloadoptions#style)|Для каждого элемента в коллекции: значение константы, представляющее стиль среза. Возможные значения: "SlicerStyleLight1", "SlicerStyleLight6", "TableStyleOther1", "TableStyleOther2", "SlicerStyleDark1" и "SlicerStyleDark6". Также можно указать настраиваемый пользовательский стиль, имеющийся в книге.|
||[top](/javascript/api/excel/excel.slicercollectionloadoptions#top)|Для каждого элемента в коллекции: представляет расстояние (в пунктах) от верхнего края среза до верхнего края листа.|
||[width](/javascript/api/excel/excel.slicercollectionloadoptions#width)|Для каждого элемента в коллекции: представляет ширину (в пунктах) среза.|
||[worksheet](/javascript/api/excel/excel.slicercollectionloadoptions#worksheet)|Для каждого элемента в коллекции: представляет лист, содержащий срез.|
|[Слицерколлектионупдатедата](/javascript/api/excel/excel.slicercollectionupdatedata)|[items](/javascript/api/excel/excel.slicercollectionupdatedata#items)||
|[Слицердата](/javascript/api/excel/excel.slicerdata)|[caption](/javascript/api/excel/excel.slicerdata#caption)|Представляет подпись среза.|
||[height](/javascript/api/excel/excel.slicerdata#height)|Представляет высоту среза (в пунктах).|
||[id](/javascript/api/excel/excel.slicerdata#id)|Представляет уникальный идентификатор среза. Только для чтения.|
||[isFilterCleared](/javascript/api/excel/excel.slicerdata#isfiltercleared)|Значение true, если удалены все фильтры, примененные к срезу.|
||[left](/javascript/api/excel/excel.slicerdata#left)|Представляет расстояние в пунктах от левого края среза до левого края листа.|
||[name](/javascript/api/excel/excel.slicerdata#name)|Представляет имя среза.|
||[nameInFormula](/javascript/api/excel/excel.slicerdata#nameinformula)|Представляет имя среза, используемое в формуле.|
||[slicerItems](/javascript/api/excel/excel.slicerdata#sliceritems)|Представляет коллекцию объектов SlicerItem, которые являются частью среза. Только для чтения.|
||[sortBy](/javascript/api/excel/excel.slicerdata#sortby)|Представляет порядок сортировки элементов в срезе. Возможные значения: DataSourceOrder, Ascending, Descending.|
||[style](/javascript/api/excel/excel.slicerdata#style)|Постоянное значение, представляющее стиль среза. Возможные значения: "SlicerStyleLight1", "SlicerStyleLight6", "TableStyleOther1", "TableStyleOther2", "SlicerStyleDark1" и "SlicerStyleDark6". Также можно указать настраиваемый пользовательский стиль, имеющийся в книге.|
||[top](/javascript/api/excel/excel.slicerdata#top)|Представляет расстояние в пунктах от верхнего края среза до верхнего края листа.|
||[width](/javascript/api/excel/excel.slicerdata#width)|Представляет ширину среза (в пунктах).|
||[worksheet](/javascript/api/excel/excel.slicerdata#worksheet)|Представляет лист, содержащий срез. Только для чтения.|
|[SlicerItem](/javascript/api/excel/excel.sliceritem)|[isSelected](/javascript/api/excel/excel.sliceritem#isselected)|Значение true, если выбран элемент среза.|
||[hasData](/javascript/api/excel/excel.sliceritem#hasdata)|Значение true, если элемент среза содержит данные. |
||[key](/javascript/api/excel/excel.sliceritem#key)|Представляет уникальное значение, соответствующее элементу среза.|
||[name](/javascript/api/excel/excel.sliceritem#name)|Представляет заголовок, отображаемый в пользовательском интерфейсе.|
||[Set (Properties: Excel. SlicerItem)](/javascript/api/excel/excel.sliceritem#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Слицеритемупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.sliceritem#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection)|[getCount()](/javascript/api/excel/excel.sliceritemcollection#getcount--)|Возвращает количество элементов в срезе.|
||[getItem(key: string)](/javascript/api/excel/excel.sliceritemcollection#getitem-key-)|Получает объект элемента среза по ключу или имени.|
||[getItemAt(index: number)](/javascript/api/excel/excel.sliceritemcollection#getitemat-index-)|Получает элемент среза на основе его позиции в коллекции.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.sliceritemcollection#getitemornullobject-key-)|Получает элемент среза по ключу или имени. Если элемент среза не существует, возвращает пустой объект.|
||[items](/javascript/api/excel/excel.sliceritemcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Слицеритемколлектиондата](/javascript/api/excel/excel.sliceritemcollectiondata)|[items](/javascript/api/excel/excel.sliceritemcollectiondata#items)||
|[Слицеритемколлектионлоадоптионс](/javascript/api/excel/excel.sliceritemcollectionloadoptions)|[$all](/javascript/api/excel/excel.sliceritemcollectionloadoptions#$all)||
||[hasData](/javascript/api/excel/excel.sliceritemcollectionloadoptions#hasdata)|Для каждого элемента в коллекции: true, если у элемента среза есть данные.|
||[isSelected](/javascript/api/excel/excel.sliceritemcollectionloadoptions#isselected)|Для каждого элемента в коллекции: true, если выбран элемент среза.|
||[key](/javascript/api/excel/excel.sliceritemcollectionloadoptions#key)|Для каждого элемента в коллекции: представляет уникальное значение, представляющее элемент среза.|
||[name](/javascript/api/excel/excel.sliceritemcollectionloadoptions#name)|Для каждого элемента в коллекции: представляет название, отображаемое в пользовательском интерфейсе.|
|[Слицеритемколлектионупдатедата](/javascript/api/excel/excel.sliceritemcollectionupdatedata)|[items](/javascript/api/excel/excel.sliceritemcollectionupdatedata#items)||
|[Слицеритемдата](/javascript/api/excel/excel.sliceritemdata)|[hasData](/javascript/api/excel/excel.sliceritemdata#hasdata)|Значение true, если элемент среза содержит данные. |
||[isSelected](/javascript/api/excel/excel.sliceritemdata#isselected)|Значение true, если выбран элемент среза.|
||[key](/javascript/api/excel/excel.sliceritemdata#key)|Представляет уникальное значение, соответствующее элементу среза.|
||[name](/javascript/api/excel/excel.sliceritemdata#name)|Представляет заголовок, отображаемый в пользовательском интерфейсе.|
|[Слицеритемлоадоптионс](/javascript/api/excel/excel.sliceritemloadoptions)|[$all](/javascript/api/excel/excel.sliceritemloadoptions#$all)||
||[hasData](/javascript/api/excel/excel.sliceritemloadoptions#hasdata)|Значение true, если элемент среза содержит данные. |
||[isSelected](/javascript/api/excel/excel.sliceritemloadoptions#isselected)|Значение true, если выбран элемент среза.|
||[key](/javascript/api/excel/excel.sliceritemloadoptions#key)|Представляет уникальное значение, соответствующее элементу среза.|
||[name](/javascript/api/excel/excel.sliceritemloadoptions#name)|Представляет заголовок, отображаемый в пользовательском интерфейсе.|
|[Слицеритемупдатедата](/javascript/api/excel/excel.sliceritemupdatedata)|[isSelected](/javascript/api/excel/excel.sliceritemupdatedata#isselected)|Значение true, если выбран элемент среза.|
|[Слицерлоадоптионс](/javascript/api/excel/excel.slicerloadoptions)|[$all](/javascript/api/excel/excel.slicerloadoptions#$all)||
||[caption](/javascript/api/excel/excel.slicerloadoptions#caption)|Представляет подпись среза.|
||[height](/javascript/api/excel/excel.slicerloadoptions#height)|Представляет высоту среза (в пунктах).|
||[id](/javascript/api/excel/excel.slicerloadoptions#id)|Представляет уникальный идентификатор среза. Только для чтения.|
||[isFilterCleared](/javascript/api/excel/excel.slicerloadoptions#isfiltercleared)|Значение true, если удалены все фильтры, примененные к срезу.|
||[left](/javascript/api/excel/excel.slicerloadoptions#left)|Представляет расстояние в пунктах от левого края среза до левого края листа.|
||[name](/javascript/api/excel/excel.slicerloadoptions#name)|Представляет имя среза.|
||[nameInFormula](/javascript/api/excel/excel.slicerloadoptions#nameinformula)|Представляет имя среза, используемое в формуле.|
||[sortBy](/javascript/api/excel/excel.slicerloadoptions#sortby)|Представляет порядок сортировки элементов в срезе. Возможные значения: DataSourceOrder, Ascending, Descending.|
||[style](/javascript/api/excel/excel.slicerloadoptions#style)|Постоянное значение, представляющее стиль среза. Возможные значения: "SlicerStyleLight1", "SlicerStyleLight6", "TableStyleOther1", "TableStyleOther2", "SlicerStyleDark1" и "SlicerStyleDark6". Также можно указать настраиваемый пользовательский стиль, имеющийся в книге.|
||[top](/javascript/api/excel/excel.slicerloadoptions#top)|Представляет расстояние в пунктах от верхнего края среза до верхнего края листа.|
||[width](/javascript/api/excel/excel.slicerloadoptions#width)|Представляет ширину среза (в пунктах).|
||[worksheet](/javascript/api/excel/excel.slicerloadoptions#worksheet)|Представляет лист, содержащий срез.|
|[SlicerStyle](/javascript/api/excel/excel.slicerstyle)|[delete()](/javascript/api/excel/excel.slicerstyle#delete--)|Удаляет объект SlicerStyle.|
||[duplicate()](/javascript/api/excel/excel.slicerstyle#duplicate--)|Создает дубликат объекта SlicerStyle с копиями всех элементов стиля.|
||[name](/javascript/api/excel/excel.slicerstyle#name)|Получает имя объекта SlicerStyle.|
||[readOnly](/javascript/api/excel/excel.slicerstyle#readonly)|Указывает, предназначен ли объект SlicerStyle только для чтения. Только для чтения.|
||[Set (Properties: Excel. Слицерстиле)](/javascript/api/excel/excel.slicerstyle#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Слицерстилеупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.slicerstyle#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[SlicerStyleCollection](/javascript/api/excel/excel.slicerstylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.slicerstylecollection#add-name--makeuniquename-)|Создает пустой объект SlicerStyle с указанным именем.|
||[getCount()](/javascript/api/excel/excel.slicerstylecollection#getcount--)|Получает количество стилей срезов в коллекции.|
||[getDefault()](/javascript/api/excel/excel.slicerstylecollection#getdefault--)|Получает используемый по умолчанию объект SlicerStyle для области родительского объекта.|
||[getItem(name: string)](/javascript/api/excel/excel.slicerstylecollection#getitem-name-)|Получает объект SlicerStyle по имени.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.slicerstylecollection#getitemornullobject-name-)|Получает объект SlicerStyle по имени. Если объект SlicerStyle не существует, возвращает пустой объект.|
||[items](/javascript/api/excel/excel.slicerstylecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[setDefault(newDefaultStyle: SlicerStyle \| string)](/javascript/api/excel/excel.slicerstylecollection#setdefault-newdefaultstyle-)|Задает объект SlicerStyle, используемый по умолчанию в области родительского объекта.|
|[Слицерстилеколлектиондата](/javascript/api/excel/excel.slicerstylecollectiondata)|[items](/javascript/api/excel/excel.slicerstylecollectiondata#items)||
|[Слицерстилеколлектионлоадоптионс](/javascript/api/excel/excel.slicerstylecollectionloadoptions)|[$all](/javascript/api/excel/excel.slicerstylecollectionloadoptions#$all)||
||[name](/javascript/api/excel/excel.slicerstylecollectionloadoptions#name)|Для каждого элемента в коллекции: получает имя Слицерстиле.|
||[readOnly](/javascript/api/excel/excel.slicerstylecollectionloadoptions#readonly)|Для каждого элемента в коллекции: указывает, является ли этот объект Слицерстиле доступен только для чтения. Только для чтения.|
|[Слицерстилеколлектионупдатедата](/javascript/api/excel/excel.slicerstylecollectionupdatedata)|[items](/javascript/api/excel/excel.slicerstylecollectionupdatedata#items)||
|[Слицерстиледата](/javascript/api/excel/excel.slicerstyledata)|[name](/javascript/api/excel/excel.slicerstyledata#name)|Получает имя объекта SlicerStyle.|
||[readOnly](/javascript/api/excel/excel.slicerstyledata#readonly)|Указывает, предназначен ли объект SlicerStyle только для чтения. Только для чтения.|
|[Слицерстилелоадоптионс](/javascript/api/excel/excel.slicerstyleloadoptions)|[$all](/javascript/api/excel/excel.slicerstyleloadoptions#$all)||
||[name](/javascript/api/excel/excel.slicerstyleloadoptions#name)|Получает имя объекта SlicerStyle.|
||[readOnly](/javascript/api/excel/excel.slicerstyleloadoptions#readonly)|Указывает, предназначен ли объект SlicerStyle только для чтения. Только для чтения.|
|[Слицерстилеупдатедата](/javascript/api/excel/excel.slicerstyleupdatedata)|[name](/javascript/api/excel/excel.slicerstyleupdatedata#name)|Получает имя объекта SlicerStyle.|
|[Слицерупдатедата](/javascript/api/excel/excel.slicerupdatedata)|[caption](/javascript/api/excel/excel.slicerupdatedata#caption)|Представляет подпись среза.|
||[height](/javascript/api/excel/excel.slicerupdatedata#height)|Представляет высоту среза (в пунктах).|
||[left](/javascript/api/excel/excel.slicerupdatedata#left)|Представляет расстояние в пунктах от левого края среза до левого края листа.|
||[name](/javascript/api/excel/excel.slicerupdatedata#name)|Представляет имя среза.|
||[nameInFormula](/javascript/api/excel/excel.slicerupdatedata#nameinformula)|Представляет имя среза, используемое в формуле.|
||[sortBy](/javascript/api/excel/excel.slicerupdatedata#sortby)|Представляет порядок сортировки элементов в срезе. Возможные значения: DataSourceOrder, Ascending, Descending.|
||[style](/javascript/api/excel/excel.slicerupdatedata#style)|Постоянное значение, представляющее стиль среза. Возможные значения: "SlicerStyleLight1", "SlicerStyleLight6", "TableStyleOther1", "TableStyleOther2", "SlicerStyleDark1" и "SlicerStyleDark6". Также можно указать настраиваемый пользовательский стиль, имеющийся в книге.|
||[top](/javascript/api/excel/excel.slicerupdatedata#top)|Представляет расстояние в пунктах от верхнего края среза до верхнего края листа.|
||[width](/javascript/api/excel/excel.slicerupdatedata#width)|Представляет ширину среза (в пунктах).|
||[worksheet](/javascript/api/excel/excel.slicerupdatedata#worksheet)|Представляет лист, содержащий срез.|
|[Table](/javascript/api/excel/excel.table)|[clearStyle()](/javascript/api/excel/excel.table#clearstyle--)|Изменяет таблицу для использования стиля таблицы по умолчанию.|
||[onFiltered](/javascript/api/excel/excel.table#onfiltered)|Возникает, если применен фильтр к указанной таблице.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[onFiltered](/javascript/api/excel/excel.tablecollection#onfiltered)|Возникает, если применен фильтр к любой таблице в книге или листе.|
|[TableFilteredEventArgs](/javascript/api/excel/excel.tablefilteredeventargs)|[tableId](/javascript/api/excel/excel.tablefilteredeventargs#tableid)|Представляет идентификатор таблицы, в которой применен фильтр.|
||[type](/javascript/api/excel/excel.tablefilteredeventargs#type)|Представляет тип события. Дополнительные сведения см. в статье Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.tablefilteredeventargs#worksheetid)|Представляет идентификатор листа, содержащего таблицу.|
|[TableStyle](/javascript/api/excel/excel.tablestyle)|[delete()](/javascript/api/excel/excel.tablestyle#delete--)|Удаляет объект TableStyle.|
||[duplicate()](/javascript/api/excel/excel.tablestyle#duplicate--)|Создает дубликат объекта TableStyle с копиями всех элементов стиля.|
||[name](/javascript/api/excel/excel.tablestyle#name)|Получает имя объекта TableStyle.|
||[readOnly](/javascript/api/excel/excel.tablestyle#readonly)|Указывает, предназначен ли объект TableStyle только для чтения. Только для чтения.|
||[Set (Properties: Excel. TableStyle)](/javascript/api/excel/excel.tablestyle#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Таблестилеупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.tablestyle#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[TableStyleCollection](/javascript/api/excel/excel.tablestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.tablestylecollection#add-name--makeuniquename-)|Создает пустой объект TableStyle с указанным именем.|
||[getCount()](/javascript/api/excel/excel.tablestylecollection#getcount--)|Получает количество стилей таблиц в коллекции.|
||[getDefault()](/javascript/api/excel/excel.tablestylecollection#getdefault--)|Получает используемый по умолчанию объект TableStyle для области родительского объекта.|
||[getItem(name: string)](/javascript/api/excel/excel.tablestylecollection#getitem-name-)|Получает объект TableStyle по имени.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.tablestylecollection#getitemornullobject-name-)|Получает объект TableStyle по имени. Если объект TableStyle не существует, возвращает пустой объект.|
||[items](/javascript/api/excel/excel.tablestylecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[setDefault(newDefaultStyle: TableStyle \| string)](/javascript/api/excel/excel.tablestylecollection#setdefault-newdefaultstyle-)|Задает объект TableStyle, используемый по умолчанию в области родительского объекта.|
|[Таблестилеколлектиондата](/javascript/api/excel/excel.tablestylecollectiondata)|[items](/javascript/api/excel/excel.tablestylecollectiondata#items)||
|[Таблестилеколлектионлоадоптионс](/javascript/api/excel/excel.tablestylecollectionloadoptions)|[$all](/javascript/api/excel/excel.tablestylecollectionloadoptions#$all)||
||[name](/javascript/api/excel/excel.tablestylecollectionloadoptions#name)|Для каждого элемента в коллекции: получает имя TableStyle.|
||[readOnly](/javascript/api/excel/excel.tablestylecollectionloadoptions#readonly)|Для каждого элемента в коллекции: указывает, является ли этот объект TableStyle доступен только для чтения. Только для чтения.|
|[Таблестилеколлектионупдатедата](/javascript/api/excel/excel.tablestylecollectionupdatedata)|[items](/javascript/api/excel/excel.tablestylecollectionupdatedata#items)||
|[Таблестиледата](/javascript/api/excel/excel.tablestyledata)|[name](/javascript/api/excel/excel.tablestyledata#name)|Получает имя объекта TableStyle.|
||[readOnly](/javascript/api/excel/excel.tablestyledata#readonly)|Указывает, предназначен ли объект TableStyle только для чтения. Только для чтения.|
|[Таблестилелоадоптионс](/javascript/api/excel/excel.tablestyleloadoptions)|[$all](/javascript/api/excel/excel.tablestyleloadoptions#$all)||
||[name](/javascript/api/excel/excel.tablestyleloadoptions#name)|Получает имя объекта TableStyle.|
||[readOnly](/javascript/api/excel/excel.tablestyleloadoptions#readonly)|Указывает, предназначен ли объект TableStyle только для чтения. Только для чтения.|
|[Таблестилеупдатедата](/javascript/api/excel/excel.tablestyleupdatedata)|[name](/javascript/api/excel/excel.tablestyleupdatedata#name)|Получает имя объекта TableStyle.|
|[TimelineStyle](/javascript/api/excel/excel.timelinestyle)|[delete()](/javascript/api/excel/excel.timelinestyle#delete--)|Удаляет объект TableStyle.|
||[duplicate()](/javascript/api/excel/excel.timelinestyle#duplicate--)|Создает дубликат объекта TimelineStyle с копиями всех элементов стиля.|
||[name](/javascript/api/excel/excel.timelinestyle#name)|Получает имя объекта TimelineStyle.|
||[readOnly](/javascript/api/excel/excel.timelinestyle#readonly)|Указывает, предназначен ли объект TimelineStyle только для чтения. Только для чтения.|
||[Set (Properties: Excel. Тимелинестиле)](/javascript/api/excel/excel.timelinestyle#set-properties-)|Задает одновременно несколько свойств объекта на основе существующего загруженного объекта.|
||[Set (Properties: interfaces. Тимелинестилеупдатедата, Options?: объект officeextension. UpdateOptions)](/javascript/api/excel/excel.timelinestyle#set-properties--options-)|Задает одновременно несколько свойств объекта. Можно передать либо простой объект с соответствующими свойствами, либо другой объект API того же типа.|
|[TimelineStyleCollection](/javascript/api/excel/excel.timelinestylecollection)|[add(name: string, makeUniqueName?: boolean)](/javascript/api/excel/excel.timelinestylecollection#add-name--makeuniquename-)|Создает пустой объект TimelineStyle с указанным именем.|
||[getCount()](/javascript/api/excel/excel.timelinestylecollection#getcount--)|Получает количество стилей временной шкалы в коллекции.|
||[getDefault()](/javascript/api/excel/excel.timelinestylecollection#getdefault--)|Получает используемый по умолчанию объект TimelineStyle для области родительского объекта.|
||[getItem(name: string)](/javascript/api/excel/excel.timelinestylecollection#getitem-name-)|Получает объект TimelineStyle по имени.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.timelinestylecollection#getitemornullobject-name-)|Получает объект TimelineStyle по имени. Если объект TimelineStyle не существует, возвращает пустой объект.|
||[items](/javascript/api/excel/excel.timelinestylecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[setDefault(newDefaultStyle: TimelineStyle \| string)](/javascript/api/excel/excel.timelinestylecollection#setdefault-newdefaultstyle-)|Задает объект TimelineStyle, используемый по умолчанию в области родительского объекта.|
|[Тимелинестилеколлектиондата](/javascript/api/excel/excel.timelinestylecollectiondata)|[items](/javascript/api/excel/excel.timelinestylecollectiondata#items)||
|[Тимелинестилеколлектионлоадоптионс](/javascript/api/excel/excel.timelinestylecollectionloadoptions)|[$all](/javascript/api/excel/excel.timelinestylecollectionloadoptions#$all)||
||[name](/javascript/api/excel/excel.timelinestylecollectionloadoptions#name)|Для каждого элемента в коллекции: получает имя Тимелинестиле.|
||[readOnly](/javascript/api/excel/excel.timelinestylecollectionloadoptions#readonly)|Для каждого элемента в коллекции: указывает, является ли этот объект Тимелинестиле доступен только для чтения. Только для чтения.|
|[Тимелинестилеколлектионупдатедата](/javascript/api/excel/excel.timelinestylecollectionupdatedata)|[items](/javascript/api/excel/excel.timelinestylecollectionupdatedata#items)||
|[Тимелинестиледата](/javascript/api/excel/excel.timelinestyledata)|[name](/javascript/api/excel/excel.timelinestyledata#name)|Получает имя объекта TimelineStyle.|
||[readOnly](/javascript/api/excel/excel.timelinestyledata#readonly)|Указывает, предназначен ли объект TimelineStyle только для чтения. Только для чтения.|
|[Тимелинестилелоадоптионс](/javascript/api/excel/excel.timelinestyleloadoptions)|[$all](/javascript/api/excel/excel.timelinestyleloadoptions#$all)||
||[name](/javascript/api/excel/excel.timelinestyleloadoptions#name)|Получает имя объекта TimelineStyle.|
||[readOnly](/javascript/api/excel/excel.timelinestyleloadoptions#readonly)|Указывает, предназначен ли объект TimelineStyle только для чтения. Только для чтения.|
|[Тимелинестилеупдатедата](/javascript/api/excel/excel.timelinestyleupdatedata)|[name](/javascript/api/excel/excel.timelinestyleupdatedata#name)|Получает имя объекта TimelineStyle.|
|[Workbook](/javascript/api/excel/excel.workbook)|[close(closeBehavior?: "Save" \| "SkipSave")](/javascript/api/excel/excel.workbook#close-closebehavior-)|Закрывает текущую книгу.|
||[close(closeBehavior?: Excel.CloseBehavior)](/javascript/api/excel/excel.workbook#close-closebehavior-)|Закрывает текущую книгу.|
||[getActiveSlicer()](/javascript/api/excel/excel.workbook#getactiveslicer--)|Получает текущий активный срез в книге. Если активного среза нет, создается `ItemNotFound` исключение.|
||[getActiveSlicerOrNullObject()](/javascript/api/excel/excel.workbook#getactiveslicerornullobject--)|Получает текущий активный срез в книге. Если активный срез отсутствует, возвращается пустой объект.|
||[comments](/javascript/api/excel/excel.workbook#comments)|Представляет коллекцию примечаний, связанных с книгой. Только для чтения.|
||[pivotTableStyles](/javascript/api/excel/excel.workbook#pivottablestyles)|Представляет коллекцию объектов PivotTableStyles, связанных с книгой. Только для чтения.|
||[slicerStyles](/javascript/api/excel/excel.workbook#slicerstyles)|Представляет коллекцию объектов SlicerStyles, связанных с книгой. Только для чтения.|
||[slicers](/javascript/api/excel/excel.workbook#slicers)|Представляет коллекцию срезов, связанных с книгой. Только для чтения.|
||[tableStyles](/javascript/api/excel/excel.workbook#tablestyles)|Представляет коллекцию объектов TableStyles, связанных с книгой. Только для чтения.|
||[timelineStyles](/javascript/api/excel/excel.workbook#timelinestyles)|Представляет коллекцию объектов TimelineStyles, связанных с книгой. Только для чтения.|
||[save(saveBehavior?: "Save" \| "Prompt")](/javascript/api/excel/excel.workbook#save-savebehavior-)|Сохраняет текущую книгу.|
||[save(saveBehavior?: Excel.SaveBehavior)](/javascript/api/excel/excel.workbook#save-savebehavior-)|Сохраняет текущую книгу.|
||[use1904DateSystem](/javascript/api/excel/excel.workbook#use1904datesystem)|Значение true, если в книге используется система дат 1904.|
|[Воркбукдата](/javascript/api/excel/excel.workbookdata)|[comments](/javascript/api/excel/excel.workbookdata#comments)|Представляет коллекцию примечаний, связанных с книгой. Только для чтения.|
||[pivotTableStyles](/javascript/api/excel/excel.workbookdata#pivottablestyles)|Представляет коллекцию объектов PivotTableStyles, связанных с книгой. Только для чтения.|
||[slicerStyles](/javascript/api/excel/excel.workbookdata#slicerstyles)|Представляет коллекцию объектов SlicerStyles, связанных с книгой. Только для чтения.|
||[slicers](/javascript/api/excel/excel.workbookdata#slicers)|Представляет коллекцию срезов, связанных с книгой. Только для чтения.|
||[tableStyles](/javascript/api/excel/excel.workbookdata#tablestyles)|Представляет коллекцию объектов TableStyles, связанных с книгой. Только для чтения.|
||[timelineStyles](/javascript/api/excel/excel.workbookdata#timelinestyles)|Представляет коллекцию объектов TimelineStyles, связанных с книгой. Только для чтения.|
||[use1904DateSystem](/javascript/api/excel/excel.workbookdata#use1904datesystem)|Значение true, если в книге используется система дат 1904.|
|[Воркбуклоадоптионс](/javascript/api/excel/excel.workbookloadoptions)|[use1904DateSystem](/javascript/api/excel/excel.workbookloadoptions#use1904datesystem)|Значение true, если в книге используется система дат 1904.|
|[Воркбукупдатедата](/javascript/api/excel/excel.workbookupdatedata)|[use1904DateSystem](/javascript/api/excel/excel.workbookupdatedata#use1904datesystem)|Значение true, если в книге используется система дат 1904.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[comments](/javascript/api/excel/excel.worksheet#comments)|Возвращает коллекцию всех объектов Comments на листе. Только для чтения.|
||[onColumnSorted](/javascript/api/excel/excel.worksheet#oncolumnsorted)|Возникает при сортировке по столбцам.|
||[onFiltered](/javascript/api/excel/excel.worksheet#onfiltered)|Возникает, если применен фильтр к указанному листу.|
||[Онровхидденчанжед](/javascript/api/excel/excel.worksheet#onrowhiddenchanged)|Возникает при изменении скрытого состояния строки на определенном листе.|
||[onRowSorted](/javascript/api/excel/excel.worksheet#onrowsorted)|Возникает при сортировке по строкам.|
||[onSingleClicked](/javascript/api/excel/excel.worksheet#onsingleclicked)|Возникает, когда происходит щелчок левой кнопкой мыши или нажатие на листе.|
||[slicers](/javascript/api/excel/excel.worksheet#slicers)|Возвращает коллекцию срезов, имеющихся на листе. Только для чтения.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: "None" \| "Before" \| "After" \| "Beginning" \| "End", relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Вставляет указанные листы книги в текущую книгу.|
||[addFromBase64(base64File: string, sheetNamesToInsert?: string[], positionType?: Excel.WorksheetPositionType, relativeTo?: Worksheet \| string)](/javascript/api/excel/excel.worksheetcollection#addfrombase64-base64file--sheetnamestoinsert--positiontype--relativeto-)|Вставляет указанные листы книги в текущую книгу.|
||[onColumnSorted](/javascript/api/excel/excel.worksheetcollection#oncolumnsorted)|Возникает при сортировке по столбцам.|
||[onFiltered](/javascript/api/excel/excel.worksheetcollection#onfiltered)|Возникает при применении любого фильтра листа в книге.|
||[Онровхидденчанжед](/javascript/api/excel/excel.worksheetcollection#onrowhiddenchanged)|Происходит, когда изменилось состояние скрытой строки для любого листа в книге.|
||[onRowSorted](/javascript/api/excel/excel.worksheetcollection#onrowsorted)|Возникает при сортировке по строкам.|
||[onSingleClicked](/javascript/api/excel/excel.worksheetcollection#onsingleclicked)|Возникает, когда в коллекции листа происходит операция с нажатием и нажатием левой кнопкой мыши.|
|[WorksheetColumnSortedEventArgs](/javascript/api/excel/excel.worksheetcolumnsortedeventargs)|[address](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#address)|Получает адрес диапазона, представляющий отсортированные области конкретного листа.|
||[источник](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#source)|Получает источник события. Дополнительные сведения см. в статье Excel.EventSource.|
||[type](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#type)|Получает тип события. Дополнительные сведения см. в статье Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetcolumnsortedeventargs#worksheetid)|Получает идентификатор листа, в котором выполнена сортировка.|
|[Воркшитдата](/javascript/api/excel/excel.worksheetdata)|[comments](/javascript/api/excel/excel.worksheetdata#comments)|Возвращает коллекцию всех объектов Comments на листе. Только для чтения.|
||[slicers](/javascript/api/excel/excel.worksheetdata#slicers)|Возвращает коллекцию срезов, имеющихся на листе. Только для чтения.|
|[WorksheetFilteredEventArgs](/javascript/api/excel/excel.worksheetfilteredeventargs)|[type](/javascript/api/excel/excel.worksheetfilteredeventargs#type)|Представляет тип события. Дополнительные сведения см. в статье Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetfilteredeventargs#worksheetid)|Представляет идентификатор листа, в котором применен фильтр.|
|[Воркшитровхидденчанжедевентаргс](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs)|[address](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#address)|Получает адрес диапазона, представляющий измененную область конкретного листа.|
||[changeType](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#changetype)|Получает тип изменения, представляющий способ запуска события Changed. Дополнительные сведения см. в статье Excel. Ровхидденчанжетипе.|
||[source](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#source)|Получает источник события. Дополнительные сведения см. в статье Excel.EventSource.|
||[type](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#type)|Получает тип события. Дополнительные сведения см. в статье Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowhiddenchangedeventargs#worksheetid)|Получает идентификатор листа, в котором изменены данные.|
|[WorksheetRowSortedEventArgs](/javascript/api/excel/excel.worksheetrowsortedeventargs)|[address](/javascript/api/excel/excel.worksheetrowsortedeventargs#address)|Получает адрес диапазона, представляющий отсортированные области конкретного листа.|
||[источник](/javascript/api/excel/excel.worksheetrowsortedeventargs#source)|Получает источник события. Дополнительные сведения см. в статье Excel.EventSource.|
||[type](/javascript/api/excel/excel.worksheetrowsortedeventargs#type)|Получает тип события. Дополнительные сведения см. в статье Excel.EventType.|
||[worksheetId](/javascript/api/excel/excel.worksheetrowsortedeventargs#worksheetid)|Получает идентификатор листа, в котором выполнена сортировка.|
|[WorksheetSingleClickedEventArgs](/javascript/api/excel/excel.worksheetsingleclickedeventargs)|[address](/javascript/api/excel/excel.worksheetsingleclickedeventargs#address)|Получает адрес, представляющий ячейку, по которой выполнен щелчок левой кнопкой мыши или нажатие, для определенного листа.|
||[offsetX](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsetx)|Расстояние в пунктах от точки щелчка левой кнопкой мыши или нажатия до левого (правого при написании справа налево) края сетки ячейки, по которой выполнен щелчок левой кнопкой мыши или нажатие.|
||[offsetY](/javascript/api/excel/excel.worksheetsingleclickedeventargs#offsety)|Расстояние в пунктах от точки щелчка левой кнопкой мыши или нажатия до верхнего края сетки ячейки, по которой выполнен щелчок левой кнопкой мыши или нажатие.|
||[type](/javascript/api/excel/excel.worksheetsingleclickedeventargs#type)|Получает тип события.|
||[worksheetId](/javascript/api/excel/excel.worksheetsingleclickedeventargs#worksheetid)|Получает идентификатор листа, в котором по ячейке выполнен щелчок левой кнопкой мыши или нажатие.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-preview)
- [Наборы обязательных элементов API JavaScript для Excel](./excel-api-requirement-sets.md)
