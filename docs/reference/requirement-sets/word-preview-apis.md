---
title: API предварительного просмотра Word JavaScript
description: Сведения о предстоящих API JavaScript Word.
ms.date: 10/11/2021
ms.prod: word
ms.localizationpriority: medium
ms.openlocfilehash: 67d49528e40c3e6331a874d7589142855861b06f
ms.sourcegitcommit: fb4a55764fb60e826ad06d15d1539e41df503b65
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/14/2021
ms.locfileid: "60356359"
---
# <a name="word-javascript-preview-apis"></a>API предварительного просмотра Word JavaScript

Новые API JavaScript Word сначала вводятся в "предварительную версию", а затем становятся частью определенного набора требований с номерами после достаточного тестирования и получения отзывов пользователей.

[!INCLUDE [Information about using Word preview APIs](../../includes/word-preview-apis-note.md)]
[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

## <a name="api-list"></a>Список API

В следующей таблице перечислены API Word JavaScript, которые в настоящее время находятся в предварительном просмотре. Чтобы просмотреть полный список всех API JavaScript Word (включая API предварительного просмотра и ранее выпущенные API), см. все API [Word JavaScript.](/javascript/api/word?view=word-js-preview&preserve-view=true)

| Класс | Поля | Описание |
|:---|:---|:---|
|[Основной текст](/javascript/api/word/word.body)|[getComments()](/javascript/api/word/word.body#getComments__)|Получает комментарии, связанные с телом.|
||[endnotes](/javascript/api/word/word.body#endnotes)|Получает коллекцию endnotes в теле.|
||[сноски](/javascript/api/word/word.body#footnotes)|Получает коллекцию сносок в теле.|
||[type](/javascript/api/word/word.body#type)|Возвращает тип основного текста.|
|[Comment](/javascript/api/word/word.comment)|[content](/javascript/api/word/word.comment#content)|Gtes или задает содержимое комментария в виде простого текста.|
||[delete()](/javascript/api/word/word.comment#delete__)|Удаляет комментарий и его ответы.|
||[getRange()](/javascript/api/word/word.comment#getRange__)|Получает диапазон в основном документе, в котором находится комментарий.|
||[authorEmail](/javascript/api/word/word.comment#authorEmail)|Получает электронную почту автора примечания.|
||[authorName](/javascript/api/word/word.comment#authorName)|Получает имя автора примечания.|
||[creationDate](/javascript/api/word/word.comment#creationDate)|Получает дату создания комментария.|
||[id](/javascript/api/word/word.comment#id)|Идентификатор|
||[replies](/javascript/api/word/word.comment#replies)|Получает коллекцию объектов ответа, связанных с комментарием.|
||[reply(replyText: string)](/javascript/api/word/word.comment#reply_replyText_)|Добавляет новый ответ в конец потока комментариев.|
||[разрешено](/javascript/api/word/word.comment#resolved)|Получает или задает состояние потока комментариев.|
|[CommentCollection](/javascript/api/word/word.commentcollection)|[getFirst()](/javascript/api/word/word.commentcollection#getFirst__)|Получает первый комментарий в коллекции.|
||[getFirstOrNullObject()](/javascript/api/word/word.commentcollection#getFirstOrNullObject__)|Получает первый комментарий или объект null в коллекции.|
||[getItem(index: number)](/javascript/api/word/word.commentcollection#getItem_index_)|Получает объект комментариев по индексу в коллекции.|
||[items](/javascript/api/word/word.commentcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[CommentReply](/javascript/api/word/word.commentreply)|[content](/javascript/api/word/word.commentreply#content)|Получает или задает содержимое ответа на примечание.|
||[delete()](/javascript/api/word/word.commentreply#delete__)|Удаляет ответ на примечание.|
||[authorEmail](/javascript/api/word/word.commentreply#authorEmail)|Получает электронную почту автора ответа на примечание.|
||[authorName](/javascript/api/word/word.commentreply#authorName)|Получает имя автора ответа на примечание.|
||[creationDate](/javascript/api/word/word.commentreply#creationDate)|Получает дату создания ответа на комментарий.|
||[id](/javascript/api/word/word.commentreply#id)|Идентификатор|
||[parentComment](/javascript/api/word/word.commentreply#parentComment)|Получает родительский комментарий этого ответа.|
|[CommentReplyCollection](/javascript/api/word/word.commentreplycollection)|[getFirst()](/javascript/api/word/word.commentreplycollection#getFirst__)|Получает первый ответ комментариев в коллекции.|
||[getFirstOrNullObject()](/javascript/api/word/word.commentreplycollection#getFirstOrNullObject__)|Получает первый ответ комментария или объект null в коллекции.|
||[getItem(index: number)](/javascript/api/word/word.commentreplycollection#getItem_index_)|Получает объект ответа на комментарии по индексу в коллекции.|
||[items](/javascript/api/word/word.commentreplycollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[getComments()](/javascript/api/word/word.contentcontrol#getComments__)|Получает комментарии, связанные с телом.|
||[endnotes](/javascript/api/word/word.contentcontrol#endnotes)|Получает коллекцию endnotes в contentcontrol.|
||[сноски](/javascript/api/word/word.contentcontrol#footnotes)|Получает коллекцию сносок в contentcontrol.|
||[onDataChanged](/javascript/api/word/word.contentcontrol#onDataChanged)|Происходит при смене данных в области управления контентом.|
||[onDeleted](/javascript/api/word/word.contentcontrol#onDeleted)|Происходит при удалении управления контентом.|
||[onSelectionChanged](/javascript/api/word/word.contentcontrol#onSelectionChanged)|Возникает при смене выбора в области управления контентом.|
|[ContentControlEventArgs](/javascript/api/word/word.contentcontroleventargs)|[contentControl](/javascript/api/word/word.contentcontroleventargs#contentControl)|Объект, который поднял событие.|
||[eventType](/javascript/api/word/word.contentcontroleventargs#eventType)|Тип события.|
|[CustomXmlPart](/javascript/api/word/word.customxmlpart)|[delete()](/javascript/api/word/word.customxmlpart#delete__)|Удаляет пользовательскую XML-часть.|
||[deleteAttribute (xpath: string, namespaceMappings: any, name: string)](/javascript/api/word/word.customxmlpart#deleteAttribute_xpath__namespaceMappings__name_)|Удаляет атрибут с заданным именем из элемента, идентифицированного xpath.|
||[deleteElement (xpath: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#deleteElement_xpath__namespaceMappings_)|Удаляет элемент, идентифицированный xpath.|
||[getXml()](/javascript/api/word/word.customxmlpart#getXml__)|Получает полное XML-содержимое пользовательской части XML.|
||[insertAttribute (xpath: string, namespaceMappings: any, name: string, value: string)](/javascript/api/word/word.customxmlpart#insertAttribute_xpath__namespaceMappings__name__value_)|Вставляет атрибут с заданным именем и значением в элемент, идентифицированный xpath.|
||[insertElement (xpath: string, xml: string, namespaceMappings: any, index?: number)](/javascript/api/word/word.customxmlpart#insertElement_xpath__xml__namespaceMappings__index_)|Вставляет данный XML в родительский элемент, идентифицированный xpath в индексе положения ребенка.|
||[query(xpath: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#query_xpath__namespaceMappings_)|Запрашивает XML-содержимое пользовательской части XML.|
||[id](/javascript/api/word/word.customxmlpart#id)|Получает ID пользовательской части XML.|
||[namespaceUri](/javascript/api/word/word.customxmlpart#namespaceUri)|Получает URI пространства имен пользовательской части XML.|
||[setXml (xml: string)](/javascript/api/word/word.customxmlpart#setXml_xml_)|Задает полное XML-содержимое пользовательской части XML.|
||[updateAttribute(xpath: string, namespaceMappings: any, name: string, value: string)](/javascript/api/word/word.customxmlpart#updateAttribute_xpath__namespaceMappings__name__value_)|Обновляет значение атрибута с заданным именем элемента, идентифицированного xpath.|
||[updateElement (xpath: string, xml: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#updateElement_xpath__xml__namespaceMappings_)|Обновляет XML элемента, идентифицированного xpath.|
|[CustomXmlPartCollection](/javascript/api/word/word.customxmlpartcollection)|[add(xml: string)](/javascript/api/word/word.customxmlpartcollection#add_xml_)|Добавляет в документ новую настраиваемую часть XML.|
||[getByNamespace (namespaceUri: string)](/javascript/api/word/word.customxmlpartcollection#getByNamespace_namespaceUri_)|Получает новую ограниченную коллекцию пользовательских XML-частей, пространства имен которых совпадают с указанным пространством имен.|
||[getCount()](/javascript/api/word/word.customxmlpartcollection#getCount__)|Возвращает число элементов в коллекции.|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartcollection#getItem_id_)|Получает пользовательскую XML-часть по идентификатору.|
||[getItemOrNullObject(id: строка)](/javascript/api/word/word.customxmlpartcollection#getItemOrNullObject_id_)|Получает пользовательскую XML-часть по идентификатору.|
||[items](/javascript/api/word/word.customxmlpartcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[CustomXmlPartScopedCollection](/javascript/api/word/word.customxmlpartscopedcollection)|[getCount()](/javascript/api/word/word.customxmlpartscopedcollection#getCount__)|Возвращает число элементов в коллекции.|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#getItem_id_)|Получает пользовательскую XML-часть по идентификатору.|
||[getItemOrNullObject(id: строка)](/javascript/api/word/word.customxmlpartscopedcollection#getItemOrNullObject_id_)|Получает пользовательскую XML-часть по идентификатору.|
||[getOnlyItem()](/javascript/api/word/word.customxmlpartscopedcollection#getOnlyItem__)|Если коллекция содержит ровно один элемент, этот метод возвращает его.|
||[getOnlyItemOrNullObject()](/javascript/api/word/word.customxmlpartscopedcollection#getOnlyItemOrNullObject__)|Если коллекция содержит ровно один элемент, этот метод возвращает его.|
||[items](/javascript/api/word/word.customxmlpartscopedcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Document](/javascript/api/word/word.document)|[deleteBookmark (имя: строка)](/javascript/api/word/word.document#deleteBookmark_name_)|Удаляет закладки, если она существует, из документа.|
||[getBookmarkRange (имя: строка)](/javascript/api/word/word.document#getBookmarkRange_name_)|Получает диапазон закладок.|
||[getBookmarkRangeOrNullObject (имя: строка)](/javascript/api/word/word.document#getBookmarkRangeOrNullObject_name_)|Получает диапазон закладок.|
||[getEndnoteBody()](/javascript/api/word/word.document#getEndnoteBody__)|Получает тело endnotes документа.|
||[getFootnoteBody()](/javascript/api/word/word.document#getFootnoteBody__)|Получает тело сносок документа.|
||[ignorePunct](/javascript/api/word/word.document#ignorePunct)||
||[ignoreSpace](/javascript/api/word/word.document#ignoreSpace)||
||[matchCase](/javascript/api/word/word.document#matchCase)||
||[matchPrefix](/javascript/api/word/word.document#matchPrefix)||
||[matchSuffix](/javascript/api/word/word.document#matchSuffix)||
||[matchWholeWord](/javascript/api/word/word.document#matchWholeWord)||
||[matchWildcards](/javascript/api/word/word.document#matchWildcards)||
||[customXmlParts](/javascript/api/word/word.document#customXmlParts)|Получает настраиваемые XML-части в документе.|
||[onContentControlAdded](/javascript/api/word/word.document#onContentControlAdded)|Возникает при добавлении управления контентом.|
||[settings](/javascript/api/word/word.document#settings)|Получает параметры надстройки в документе.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.document#search_searchText__searchOptions__ignorePunct__ignoreSpace__matchCase__matchPrefix__matchSuffix__matchWholeWord__matchWildcards_)|Выполняет поиск с указанными SearchOptions в области всего документа.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[deleteBookmark (имя: строка)](/javascript/api/word/word.documentcreated#deleteBookmark_name_)|Удаляет закладки, если она существует, из документа.|
||[getBookmarkRange (имя: строка)](/javascript/api/word/word.documentcreated#getBookmarkRange_name_)|Получает диапазон закладок.|
||[getBookmarkRangeOrNullObject (имя: строка)](/javascript/api/word/word.documentcreated#getBookmarkRangeOrNullObject_name_)|Получает диапазон закладок.|
||[customXmlParts](/javascript/api/word/word.documentcreated#customXmlParts)|Получает настраиваемые XML-части в документе.|
||[settings](/javascript/api/word/word.documentcreated#settings)|Получает параметры надстройки в документе.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[imageFormat](/javascript/api/word/word.inlinepicture#imageFormat)|Получает формат inline image.|
|[Список](/javascript/api/word/word.list)|[getLevelFont (уровень: номер)](/javascript/api/word/word.list#getLevelFont_level_)|Получает шрифт пули, номера или изображения на указанном уровне в списке.|
||[getLevelPicture(level: number)](/javascript/api/word/word.list#getLevelPicture_level_)|Получает кодированное представление строки base64 на указанном уровне в списке.|
||[resetLevelFont (уровень: номер, resetFontName?: boolean)](/javascript/api/word/word.list#resetLevelFont_level__resetFontName_)|Сброс шрифта пули, номера или изображения на указанном уровне в списке.|
||[setLevelPicture(level: number, base64EncodedImage?: string)](/javascript/api/word/word.list#setLevelPicture_level__base64EncodedImage_)|Задает изображение на указанном уровне в списке.|
|[NoteItem](/javascript/api/word/word.noteitem)|[delete()](/javascript/api/word/word.noteitem#delete__)|Удаляет элемент заметки.|
||[getNext()](/javascript/api/word/word.noteitem#getNext__)|Получает следующий элемент заметки того же типа.|
||[getNextOrNullObject()](/javascript/api/word/word.noteitem#getNextOrNullObject__)|Получает следующий элемент заметки того же типа.|
||[body](/javascript/api/word/word.noteitem#body)|Представляет объект тела элемента заметки.|
||[reference](/javascript/api/word/word.noteitem#reference)|Представляет ссылку сноска/endnote в основном документе.|
||[type](/javascript/api/word/word.noteitem#type)|Представляете тип элемента примечание: сноска или endnote.|
|[NoteItemCollection](/javascript/api/word/word.noteitemcollection)|[getFirst()](/javascript/api/word/word.noteitemcollection#getFirst__)|Получает первый элемент заметки в этой коллекции.|
||[getFirstOrNullObject()](/javascript/api/word/word.noteitemcollection#getFirstOrNullObject__)|Получает первый элемент заметки в этой коллекции.|
||[items](/javascript/api/word/word.noteitemcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Paragraph](/javascript/api/word/word.paragraph)|[getComments()](/javascript/api/word/word.paragraph#getComments__)|Получает комментарии, связанные с абзацем.|
||[endnotes](/javascript/api/word/word.paragraph#endnotes)|Получает коллекцию endnotes в абзаце.|
||[сноски](/javascript/api/word/word.paragraph#footnotes)|Получает коллекцию сносок в абзаце.|
|[Range](/javascript/api/word/word.range)|[getBookmarks (includeHidden?: boolean, includeAdjacent?: boolean)](/javascript/api/word/word.range#getBookmarks_includeHidden__includeAdjacent_)|Получает имена всех закладки или перекрывает диапазон.|
||[getComments()](/javascript/api/word/word.range#getComments__)|Получает комментарии, связанные с диапазоном.|
||[insertBookmark (имя: строка)](/javascript/api/word/word.range#insertBookmark_name_)|Вставляет закладки в диапазоне.|
||[insertComment (commentText: string)](/javascript/api/word/word.range#insertComment_commentText_)|Вставьте комментарий к диапазону.|
||[insertEndnote (insertText?: string)](/javascript/api/word/word.range#insertEndnote_insertText_)|Вставляет endnote.|
||[insertFootnote (insertText?: string)](/javascript/api/word/word.range#insertFootnote_insertText_)|Вставляет сноску.|
||[endnotes](/javascript/api/word/word.range#endnotes)|Получает коллекцию endnotes в диапазоне.|
||[сноски](/javascript/api/word/word.range#footnotes)|Получает коллекцию сносок в диапазоне.|
|[Параметр](/javascript/api/word/word.setting)|[delete()](/javascript/api/word/word.setting#delete__)|Удаляет параметр.|
||[key](/javascript/api/word/word.setting#key)|Получает ключ параметра.|
||[value](/javascript/api/word/word.setting#value)|Получает или задает значение параметра.|
|[SettingCollection](/javascript/api/word/word.settingcollection)|[add(key: string, value: any)](/javascript/api/word/word.settingcollection#add_key__value_)|Создает новый параметр или задает существующий параметр.|
||[deleteAll()](/javascript/api/word/word.settingcollection#deleteAll__)|Удаляет все параметры в этой надстройки.|
||[getCount()](/javascript/api/word/word.settingcollection#getCount__)|Получает количество параметров.|
||[getItem(key: string)](/javascript/api/word/word.settingcollection#getItem_key_)|Получает объект параметра по его ключу, который является чувствительным к делу.|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.settingcollection#getItemOrNullObject_key_)|Получает объект параметра по его ключу, который является чувствительным к делу.|
||[items](/javascript/api/word/word.settingcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Table](/javascript/api/word/word.table)|[mergeCells(topRow: number, firstCell: number, bottomRow: number, lastCell: number)](/javascript/api/word/word.table#mergeCells_topRow__firstCell__bottomRow__lastCell_)|Объединяет ячейки, ограниченные включительно первой и последней ячейками.|
||[endnotes](/javascript/api/word/word.table#endnotes)|Получает коллекцию endnotes в таблице.|
||[сноски](/javascript/api/word/word.table#footnotes)|Получает коллекцию сносок в таблице.|
|[TableCell](/javascript/api/word/word.tablecell)|[split(rowCount: number, columnCount: number)](/javascript/api/word/word.tablecell#split_rowCount__columnCount_)|Разделяет ячейку на указанное количество строк и столбцов.|
|[TableRow](/javascript/api/word/word.tablerow)|[insertContentControl()](/javascript/api/word/word.tablerow#insertContentControl__)|Вставляет управление контентом в строку.|
||[merge()](/javascript/api/word/word.tablerow#merge__)|Сливает строку в одну ячейку.|
||[endnotes](/javascript/api/word/word.tablerow#endnotes)|Получает коллекцию endnotes в строке таблицы.|
||[сноски](/javascript/api/word/word.tablerow#footnotes)|Получает коллекцию сносок в строке таблицы.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Word](/javascript/api/word)
- [Наборы обязательных элементов API JavaScript для Word](word-api-requirement-sets.md)
