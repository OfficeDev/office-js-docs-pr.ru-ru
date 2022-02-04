---
title: API предварительного просмотра Word JavaScript
description: Сведения о предстоящих API JavaScript Word.
ms.date: 02/01/2022
ms.prod: word
ms.localizationpriority: medium
---

# <a name="word-javascript-preview-apis"></a>API предварительного просмотра Word JavaScript

Новые API JavaScript Word сначала вводятся в "предварительную версию", а затем становятся частью определенного набора требований с номерами после достаточного тестирования и получения отзывов пользователей.

[!INCLUDE [Information about using Word preview APIs](../../includes/word-preview-apis-note.md)]
[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

## <a name="api-list"></a>Список API

В следующей таблице перечислены API Word JavaScript, которые в настоящее время находятся в предварительном просмотре, за исключением тех, которые доступны только [в Word в Интернете](#web-only-api-list). Чтобы просмотреть полный список всех API Word JavaScript (включая API предварительного просмотра и ранее выпущенные API), см. все API [Word JavaScript](/javascript/api/word?view=word-js-preview&preserve-view=true).

| Класс | Поля | Описание |
|:---|:---|:---|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[onDataChanged](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-ondatachanged-member)|Происходит при смене данных в области управления контентом.|
||[onDeleted](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-ondeleted-member)|Происходит при удалении управления контентом.|
||[onSelectionChanged](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-onselectionchanged-member)|Возникает при смене выбора в области управления контентом.|
|[ContentControlEventArgs](/javascript/api/word/word.contentcontroleventargs)|[contentControl](/javascript/api/word/word.contentcontroleventargs#word-word-contentcontroleventargs-contentcontrol-member)|Объект, который поднял событие.|
||[eventType](/javascript/api/word/word.contentcontroleventargs#word-word-contentcontroleventargs-eventtype-member)|Тип события.|
|[CustomXmlPart](/javascript/api/word/word.customxmlpart)|[delete()](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-delete-member(1))|Удаляет пользовательскую XML-часть.|
||[deleteAttribute (xpath: string, namespaceMappings: any, name: string)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-deleteattribute-member(1))|Удаляет атрибут с заданным именем из элемента, идентифицированного xpath.|
||[deleteElement (xpath: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-deleteelement-member(1))|Удаляет элемент, идентифицированный xpath.|
||[getXml()](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-getxml-member(1))|Получает полное XML-содержимое пользовательской части XML.|
||[id](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-id-member)|Получает ID пользовательской части XML.|
||[insertAttribute (xpath: string, namespaceMappings: any, name: string, value: string)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-insertattribute-member(1))|Вставляет атрибут с заданным именем и значением в элемент, идентифицированный xpath.|
||[insertElement (xpath: string, xml: string, namespaceMappings: any, index?: number)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-insertelement-member(1))|Вставляет данный XML в родительский элемент, идентифицированный xpath в индексе положения ребенка.|
||[namespaceUri](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-namespaceuri-member)|Получает URI пространства имен пользовательской части XML.|
||[query(xpath: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-query-member(1))|Запрашивает XML-содержимое пользовательской части XML.|
||[setXml (xml: string)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-setxml-member(1))|Задает полное XML-содержимое пользовательской части XML.|
||[updateAttribute(xpath: string, namespaceMappings: any, name: string, value: string)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-updateattribute-member(1))|Обновляет значение атрибута с заданным именем элемента, идентифицированного xpath.|
||[updateElement (xpath: string, xml: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#word-word-customxmlpart-updateelement-member(1))|Обновляет XML элемента, идентифицированного xpath.|
|[CustomXmlPartCollection](/javascript/api/word/word.customxmlpartcollection)|[add(xml: string)](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-add-member(1))|Добавляет в документ новую настраиваемую часть XML.|
||[getByNamespace (namespaceUri: string)](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-getbynamespace-member(1))|Получает новую ограниченную коллекцию пользовательских XML-частей, пространства имен которых совпадают с указанным пространством имен.|
||[getCount()](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-getcount-member(1))|Возвращает число элементов в коллекции.|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-getitem-member(1))|Получает пользовательскую XML-часть по идентификатору.|
||[getItemOrNullObject(id: строка)](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-getitemornullobject-member(1))|Получает пользовательскую XML-часть по идентификатору.|
||[items](/javascript/api/word/word.customxmlpartcollection#word-word-customxmlpartcollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[CustomXmlPartScopedCollection](/javascript/api/word/word.customxmlpartscopedcollection)|[getCount()](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-getcount-member(1))|Возвращает число элементов в коллекции.|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-getitem-member(1))|Получает пользовательскую XML-часть по идентификатору.|
||[getItemOrNullObject(id: строка)](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-getitemornullobject-member(1))|Получает пользовательскую XML-часть по идентификатору.|
||[getOnlyItem()](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-getonlyitem-member(1))|Если коллекция содержит ровно один элемент, этот метод возвращает его.|
||[getOnlyItemOrNullObject()](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-getonlyitemornullobject-member(1))|Если коллекция содержит ровно один элемент, этот метод возвращает его.|
||[items](/javascript/api/word/word.customxmlpartscopedcollection#word-word-customxmlpartscopedcollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[Document](/javascript/api/word/word.document)|[customXmlParts](/javascript/api/word/word.document#word-word-document-customxmlparts-member)|Получает настраиваемые XML-части в документе.|
||[deleteBookmark (имя: строка)](/javascript/api/word/word.document#word-word-document-deletebookmark-member(1))|Удаляет закладки, если она существует, из документа.|
||[getBookmarkRange (имя: строка)](/javascript/api/word/word.document#word-word-document-getbookmarkrange-member(1))|Получает диапазон закладок.|
||[getBookmarkRangeOrNullObject (имя: строка)](/javascript/api/word/word.document#word-word-document-getbookmarkrangeornullobject-member(1))|Получает диапазон закладок.|
||[ignorePunct](/javascript/api/word/word.document#word-word-document-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.document#word-word-document-ignorespace-member)||
||[matchCase](/javascript/api/word/word.document#word-word-document-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.document#word-word-document-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.document#word-word-document-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.document#word-word-document-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.document#word-word-document-matchwildcards-member)||
||[onContentControlAdded](/javascript/api/word/word.document#word-word-document-oncontentcontroladded-member)|Возникает при добавлении управления контентом.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.document#word-word-document-search-member(1))|Выполняет поиск с указанными вариантами поиска в области всего документа.|
||[settings](/javascript/api/word/word.document#word-word-document-settings-member)|Получает параметры надстройки в документе.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[customXmlParts](/javascript/api/word/word.documentcreated#word-word-documentcreated-customxmlparts-member)|Получает настраиваемые XML-части в документе.|
||[deleteBookmark (имя: строка)](/javascript/api/word/word.documentcreated#word-word-documentcreated-deletebookmark-member(1))|Удаляет закладки, если она существует, из документа.|
||[getBookmarkRange (имя: строка)](/javascript/api/word/word.documentcreated#word-word-documentcreated-getbookmarkrange-member(1))|Получает диапазон закладок.|
||[getBookmarkRangeOrNullObject (имя: строка)](/javascript/api/word/word.documentcreated#word-word-documentcreated-getbookmarkrangeornullobject-member(1))|Получает диапазон закладок.|
||[settings](/javascript/api/word/word.documentcreated#word-word-documentcreated-settings-member)|Получает параметры надстройки в документе.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[imageFormat](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-imageformat-member)|Получает формат inline image.|
|[Список](/javascript/api/word/word.list)|[getLevelFont (уровень: номер)](/javascript/api/word/word.list#word-word-list-getlevelfont-member(1))|Получает шрифт пули, номера или изображения на указанном уровне в списке.|
||[getLevelPicture(level: number)](/javascript/api/word/word.list#word-word-list-getlevelpicture-member(1))|Получает кодированное представление строки base64 на указанном уровне в списке.|
||[resetLevelFont (уровень: номер, resetFontName?: boolean)](/javascript/api/word/word.list#word-word-list-resetlevelfont-member(1))|Сброс шрифта пули, номера или изображения на указанном уровне в списке.|
||[setLevelPicture(level: number, base64EncodedImage?: string)](/javascript/api/word/word.list#word-word-list-setlevelpicture-member(1))|Задает изображение на указанном уровне в списке.|
|[Range](/javascript/api/word/word.range)|[getBookmarks (includeHidden?: boolean, includeAdjacent?: boolean)](/javascript/api/word/word.range#word-word-range-getbookmarks-member(1))|Получает имена всех закладки или перекрывает диапазон.|
||[insertBookmark (имя: строка)](/javascript/api/word/word.range#word-word-range-insertbookmark-member(1))|Вставляет закладки в диапазоне.|
|[Параметр](/javascript/api/word/word.setting)|[delete()](/javascript/api/word/word.setting#word-word-setting-delete-member(1))|Удаляет параметр.|
||[key](/javascript/api/word/word.setting#word-word-setting-key-member)|Получает ключ параметра.|
||[value](/javascript/api/word/word.setting#word-word-setting-value-member)|Получает или задает значение параметра.|
|[SettingCollection](/javascript/api/word/word.settingcollection)|[add(key: string, value: any)](/javascript/api/word/word.settingcollection#word-word-settingcollection-add-member(1))|Создает новый параметр или задает существующий параметр.|
||[deleteAll()](/javascript/api/word/word.settingcollection#word-word-settingcollection-deleteall-member(1))|Удаляет все параметры в этой надстройки.|
||[getCount()](/javascript/api/word/word.settingcollection#word-word-settingcollection-getcount-member(1))|Получает количество параметров.|
||[getItem(key: string)](/javascript/api/word/word.settingcollection#word-word-settingcollection-getitem-member(1))|Получает объект параметра по его ключу, который является чувствительным к делу.|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.settingcollection#word-word-settingcollection-getitemornullobject-member(1))|Получает объект параметра по его ключу, который является чувствительным к делу.|
||[items](/javascript/api/word/word.settingcollection#word-word-settingcollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[Table](/javascript/api/word/word.table)|[mergeCells(topRow: number, firstCell: number, bottomRow: number, lastCell: number)](/javascript/api/word/word.table#word-word-table-mergecells-member(1))|Объединяет ячейки, ограниченные включительно первой и последней ячейками.|
|[TableCell](/javascript/api/word/word.tablecell)|[split(rowCount: number, columnCount: number)](/javascript/api/word/word.tablecell#word-word-tablecell-split-member(1))|Разделяет ячейку на указанное количество строк и столбцов.|
|[TableRow](/javascript/api/word/word.tablerow)|[insertContentControl()](/javascript/api/word/word.tablerow#word-word-tablerow-insertcontentcontrol-member(1))|Вставляет управление контентом в строку.|
||[merge()](/javascript/api/word/word.tablerow#word-word-tablerow-merge-member(1))|Сливает строку в одну ячейку.|

## <a name="web-only-api-list"></a>Список API только для веб-пользователей

В следующей таблице перечислены API Word JavaScript, которые в настоящее время находятся в предварительном просмотре только в Word в Интернете. Чтобы просмотреть полный список всех API Word JavaScript (включая API предварительного просмотра и ранее выпущенные API), см. все API [Word JavaScript](/javascript/api/word?view=word-js-preview&preserve-view=true).

| Класс | Поля | Описание |
|:---|:---|:---|
|[Основной текст](/javascript/api/word/word.body)|[endnotes](/javascript/api/word/word.body#word-word-body-endnotes-member)|Получает коллекцию endnotes в теле.|
||[сноски](/javascript/api/word/word.body#word-word-body-footnotes-member)|Получает коллекцию сносок в теле.|
||[getComments()](/javascript/api/word/word.body#word-word-body-getcomments-member(1))|Получает комментарии, связанные с телом.|
||[getReviewedText (changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.body#word-word-body-getreviewedtext-member(1))|Получает рассмотренный текст на основе выбора ChangeTrackingVersion.|
||[type](/javascript/api/word/word.body#word-word-body-type-member)|Возвращает тип основного текста.|
|[Comment](/javascript/api/word/word.comment)|[authorEmail](/javascript/api/word/word.comment#word-word-comment-authoremail-member)|Получает электронную почту автора примечания.|
||[authorName](/javascript/api/word/word.comment#word-word-comment-authorname-member)|Получает имя автора примечания.|
||[content](/javascript/api/word/word.comment#word-word-comment-content-member)|Получает или задает содержимое комментария в виде простого текста.|
||[contentRange](/javascript/api/word/word.comment#word-word-comment-contentrange-member)|Получает или задает состояние потока комментариев.|
||[creationDate](/javascript/api/word/word.comment#word-word-comment-creationdate-member)|Получает дату создания комментария.|
||[delete()](/javascript/api/word/word.comment#word-word-comment-delete-member(1))|Удаляет комментарий и его ответы.|
||[getRange()](/javascript/api/word/word.comment#word-word-comment-getrange-member(1))|Получает диапазон в основном документе, в котором находится комментарий.|
||[id](/javascript/api/word/word.comment#word-word-comment-id-member)|ID|
||[replies](/javascript/api/word/word.comment#word-word-comment-replies-member)|Получает коллекцию объектов ответа, связанных с комментарием.|
||[reply(replyText: string)](/javascript/api/word/word.comment#word-word-comment-reply-member(1))|Добавляет новый ответ в конец потока комментариев.|
||[разрешено](/javascript/api/word/word.comment#word-word-comment-resolved-member)|Получает или задает состояние потока комментариев.|
|[CommentCollection](/javascript/api/word/word.commentcollection)|[getFirst()](/javascript/api/word/word.commentcollection#word-word-commentcollection-getfirst-member(1))|Получает первый комментарий в коллекции.|
||[getFirstOrNullObject()](/javascript/api/word/word.commentcollection#word-word-commentcollection-getfirstornullobject-member(1))|Получает первый комментарий в коллекции.|
||[getItem(index: number)](/javascript/api/word/word.commentcollection#word-word-commentcollection-getitem-member(1))|Получает объект комментариев по индексу в коллекции.|
||[items](/javascript/api/word/word.commentcollection#word-word-commentcollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[CommentContentRange](/javascript/api/word/word.commentcontentrange)|[bold](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-bold-member)|Получает или задает значение, которое указывает, является ли текст комментария смелым.|
||[hyperlink](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-hyperlink-member)|Возвращает первую гиперссылку в диапазоне или задает для него гиперссылку.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-inserttext-member(1))|Вставляет текст в указанном расположении.|
||[isEmpty](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-isempty-member)|Проверяет, является ли длина диапазона нулевой.|
||[italic](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-italic-member)|Получает или задает значение, которое указывает, является ли текст комментария italicized.|
||[strikeThrough](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-strikethrough-member)|Получает или задает значение, указывающего, есть ли в тексте комментариев забастовка.|
||[text](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-text-member)|Получает текст диапазона комментариев.|
||[underline](/javascript/api/word/word.commentcontentrange#word-word-commentcontentrange-underline-member)|Получает или задает значение, указыва которое указывает на подчеркнутой тип текста комментария.|
|[CommentReply](/javascript/api/word/word.commentreply)|[authorEmail](/javascript/api/word/word.commentreply#word-word-commentreply-authoremail-member)|Получает электронную почту автора ответа на примечание.|
||[authorName](/javascript/api/word/word.commentreply#word-word-commentreply-authorname-member)|Получает имя автора ответа на примечание.|
||[content](/javascript/api/word/word.commentreply#word-word-commentreply-content-member)|Получает или задает содержимое ответа на примечание.|
||[contentRange](/javascript/api/word/word.commentreply#word-word-commentreply-contentrange-member)|Получает или задает диапазон контента commentReply.|
||[creationDate](/javascript/api/word/word.commentreply#word-word-commentreply-creationdate-member)|Получает дату создания ответа на комментарий.|
||[delete()](/javascript/api/word/word.commentreply#word-word-commentreply-delete-member(1))|Удаляет ответ на примечание.|
||[id](/javascript/api/word/word.commentreply#word-word-commentreply-id-member)|ID|
||[parentComment](/javascript/api/word/word.commentreply#word-word-commentreply-parentcomment-member)|Получает родительский комментарий этого ответа.|
|[CommentReplyCollection](/javascript/api/word/word.commentreplycollection)|[getFirst()](/javascript/api/word/word.commentreplycollection#word-word-commentreplycollection-getfirst-member(1))|Получает первый ответ комментариев в коллекции.|
||[getFirstOrNullObject()](/javascript/api/word/word.commentreplycollection#word-word-commentreplycollection-getfirstornullobject-member(1))|Получает первый ответ комментариев в коллекции.|
||[getItem(index: number)](/javascript/api/word/word.commentreplycollection#word-word-commentreplycollection-getitem-member(1))|Получает объект ответа на комментарии по индексу в коллекции.|
||[items](/javascript/api/word/word.commentreplycollection#word-word-commentreplycollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[endnotes](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-endnotes-member)|Получает коллекцию endnotes в contentcontrol.|
||[сноски](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-footnotes-member)|Получает коллекцию сносок в contentcontrol.|
||[getComments()](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-getcomments-member(1))|Получает комментарии, связанные с телом.|
||[getReviewedText (changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-getreviewedtext-member(1))|Получает рассмотренный текст на основе выбора ChangeTrackingVersion.|
|[Document](/javascript/api/word/word.document)|[changeTrackingMode](/javascript/api/word/word.document#word-word-document-changetrackingmode-member)|Получает или задает режим ChangeTracking.|
||[getEndnoteBody()](/javascript/api/word/word.document#word-word-document-getendnotebody-member(1))|Получает конечные нотки документа в одном теле.|
||[getFootnoteBody()](/javascript/api/word/word.document#word-word-document-getfootnotebody-member(1))|Получает сноски документа в одном теле.|
|[NoteItem](/javascript/api/word/word.noteitem)|[body](/javascript/api/word/word.noteitem#word-word-noteitem-body-member)|Представляет объект тела элемента заметки.|
||[delete()](/javascript/api/word/word.noteitem#word-word-noteitem-delete-member(1))|Удаляет элемент заметки.|
||[getNext()](/javascript/api/word/word.noteitem#word-word-noteitem-getnext-member(1))|Получает следующий элемент заметки того же типа.|
||[getNextOrNullObject()](/javascript/api/word/word.noteitem#word-word-noteitem-getnextornullobject-member(1))|Получает следующий элемент заметки того же типа.|
||[reference](/javascript/api/word/word.noteitem#word-word-noteitem-reference-member)|Представляет ссылку сноски или endnote в основном документе.|
||[type](/javascript/api/word/word.noteitem#word-word-noteitem-type-member)|Представляет тип элемента примечание: сноска или endnote.|
|[NoteItemCollection](/javascript/api/word/word.noteitemcollection)|[getFirst()](/javascript/api/word/word.noteitemcollection#word-word-noteitemcollection-getfirst-member(1))|Получает первый элемент заметки в этой коллекции.|
||[getFirstOrNullObject()](/javascript/api/word/word.noteitemcollection#word-word-noteitemcollection-getfirstornullobject-member(1))|Получает первый элемент заметки в этой коллекции.|
||[items](/javascript/api/word/word.noteitemcollection#word-word-noteitemcollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[Paragraph](/javascript/api/word/word.paragraph)|[endnotes](/javascript/api/word/word.paragraph#word-word-paragraph-endnotes-member)|Получает коллекцию endnotes в абзаце.|
||[сноски](/javascript/api/word/word.paragraph#word-word-paragraph-footnotes-member)|Получает коллекцию сносок в абзаце.|
||[getComments()](/javascript/api/word/word.paragraph#word-word-paragraph-getcomments-member(1))|Получает комментарии, связанные с абзацем.|
||[getReviewedText (changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.paragraph#word-word-paragraph-getreviewedtext-member(1))|Получает рассмотренный текст на основе выбора ChangeTrackingVersion.|
|[Range](/javascript/api/word/word.range)|[endnotes](/javascript/api/word/word.range#word-word-range-endnotes-member)|Получает коллекцию endnotes в диапазоне.|
||[сноски](/javascript/api/word/word.range#word-word-range-footnotes-member)|Получает коллекцию сносок в диапазоне.|
||[getComments()](/javascript/api/word/word.range#word-word-range-getcomments-member(1))|Получает комментарии, связанные с диапазоном.|
||[getReviewedText (changeTrackingVersion?: Word.ChangeTrackingVersion)](/javascript/api/word/word.range#word-word-range-getreviewedtext-member(1))|Получает рассмотренный текст на основе выбора ChangeTrackingVersion.|
||[insertComment (commentText: string)](/javascript/api/word/word.range#word-word-range-insertcomment-member(1))|Вставьте комментарий к диапазону.|
||[insertEndnote (insertText?: string)](/javascript/api/word/word.range#word-word-range-insertendnote-member(1))|Вставляет endnote.|
||[insertFootnote (insertText?: string)](/javascript/api/word/word.range#word-word-range-insertfootnote-member(1))|Вставляет сноску.|
|[Table](/javascript/api/word/word.table)|[endnotes](/javascript/api/word/word.table#word-word-table-endnotes-member)|Получает коллекцию endnotes в таблице.|
||[сноски](/javascript/api/word/word.table#word-word-table-footnotes-member)|Получает коллекцию сносок в таблице.|
|[TableRow](/javascript/api/word/word.tablerow)|[endnotes](/javascript/api/word/word.tablerow#word-word-tablerow-endnotes-member)|Получает коллекцию endnotes в строке таблицы.|
||[сноски](/javascript/api/word/word.tablerow#word-word-tablerow-footnotes-member)|Получает коллекцию сносок в строке таблицы.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Word](/javascript/api/word)
- [Наборы обязательных элементов API JavaScript для Word](word-api-requirement-sets.md)
