---
title: Набор API API Word JavaScript 1.1
description: Сведения о наборе требований WordApi 1.1
ms.date: 11/01/2021
ms.prod: word
ms.localizationpriority: medium
---

# <a name="whats-new-in-word-javascript-api-11"></a>Что нового в API JavaScript Word 1.1

WordApi 1.1 — это первый набор требований API Word JavaScript. Это единственный набор API Word, поддерживаемый Word 2016.

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в API Word JavaScript, за набором 1.1. Чтобы просмотреть справочную документацию по API для всех API, поддерживаемых требованием API Word JavaScript, установленным 1.1, см. в справке к API Word в наборе [требований 1.1](/javascript/api/word?view=word-js-1.1&preserve-view=true).

| Класс | Поля | Описание |
|:---|:---|:---|
|[Основной текст](/javascript/api/word/word.body)|[clear()](/javascript/api/word/word.body#word-word-body-clear-member(1))|Очищает объект содержимого.|
||[contentControls](/javascript/api/word/word.body#word-word-body-contentcontrols-member)|Получает коллекцию объектов управления текстовым контентом в теле.|
||[font](/javascript/api/word/word.body#word-word-body-font-member)|Получает формат текста, указанный для содержимого документа или раздела.|
||[getHtml()](/javascript/api/word/word.body#word-word-body-gethtml-member(1))|Получает HTML-представление объекта тела.|
||[getOoxml()](/javascript/api/word/word.body#word-word-body-getooxml-member(1))|Возвращает OOXML-представление (Office Open XML) объекта содержимого.|
||[ignorePunct](/javascript/api/word/word.body#word-word-body-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.body#word-word-body-ignorespace-member)||
||[inlinePictures](/javascript/api/word/word.body#word-word-body-inlinepictures-member)|Получает коллекцию объектов InlinePicture в теле.|
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#word-word-body-insertbreak-member(1))|Вставляет разрыв в указанном расположении в основном документе.|
||[insertContentControl()](/javascript/api/word/word.body#word-word-body-insertcontentcontrol-member(1))|Включает объект содержимого в элемент управления форматированным текстом.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#word-word-body-insertfilefrombase64-member(1))|Вставляет документ в содержимое в заданном расположении.|
||[insertHtml (html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#word-word-body-inserthtml-member(1))|Вставляет HTML-код в указанном расположении.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#word-word-body-insertooxml-member(1))|Вставляет OOXML-код в указанном расположении.|
||[insertParagraph (paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#word-word-body-insertparagraph-member(1))|Вставляет абзац в указанном расположении.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#word-word-body-inserttext-member(1))|Вставляет текст в содержимое в заданном расположении.|
||[matchCase](/javascript/api/word/word.body#word-word-body-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.body#word-word-body-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.body#word-word-body-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.body#word-word-body-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.body#word-word-body-matchwildcards-member)||
||[paragraphs](/javascript/api/word/word.body#word-word-body-paragraphs-member)|Получает коллекцию объектов абзаца в теле.|
||[parentContentControl](/javascript/api/word/word.body#word-word-body-parentcontentcontrol-member)|Получает элемент управления содержимым, содержащий документ или раздел.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.body#word-word-body-search-member(1))|Выполняет поиск с указанными SearchOptions в области объекта тела.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.body#word-word-body-select-member(1))|Выбирает содержимое и переходит к нему в пользовательском интерфейсе Word.|
||[style](/javascript/api/word/word.body#word-word-body-style-member)|Получает или задает имя стиля для тела.|
||[text](/javascript/api/word/word.body#word-word-body-text-member)|Возвращает текст содержимого.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[внешний вид](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-appearance-member)|Получает или задает внешний вид элемента управления содержимым.|
||[cannotDelete](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-cannotdelete-member)|Возвращает или задает значение, указывающее, может ли пользователь удалить элемент управления содержимым.|
||[cannotEdit](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-cannotedit-member)|Возвращает или задает значение, указывающее, может ли пользователь изменять содержимое элемента управления содержимым.|
||[clear()](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-clear-member(1))|Очищает содержимое элемента управления содержимым.|
||[color](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-color-member)|Возвращает или задает цвет элемента управления содержимым.|
||[contentControls](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-contentcontrols-member)|Получает коллекцию объектов элементов управления содержимым в элементе управления содержимым.|
||[delete(keepContent: boolean)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-delete-member(1))|Удаляет элемент управления содержимым и его содержимое.|
||[font](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-font-member)|Получает текстовый формат элемента управления содержимым.|
||[getHtml()](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-gethtml-member(1))|Получает ПРЕДСТАВЛЕНИЕ HTML объекта управления контентом.|
||[getOoxml()](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-getooxml-member(1))|Возвращает OOXML-представление объекта элемента управления содержимым.|
||[id](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-id-member)|Возвращает целое число, представляющее собой идентификатор элемента управления контентом.|
||[ignorePunct](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-ignorespace-member)||
||[inlinePictures](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-inlinepictures-member)|Получает коллекцию объектов inlinePicture в элементе управления содержимым.|
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-insertbreak-member(1))|Вставляет разрыв в указанном расположении в основном документе.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-insertfilefrombase64-member(1))|Вставляет документ в управление контентом в указанном расположении.|
||[insertHtml (html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-inserthtml-member(1))|Вставляет HTML-код в элемент управления содержимым в указанном расположении.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-insertooxml-member(1))|Вставляет OOXML в управление контентом в указанном расположении.|
||[insertParagraph (paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-insertparagraph-member(1))|Вставляет абзац в указанном расположении.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-inserttext-member(1))|Вставляет текст в элемент управления содержимым в указанном расположении.|
||[matchCase](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-matchwildcards-member)||
||[paragraphs](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-paragraphs-member)|Получает коллекцию объектов абзаца в области управления контентом.|
||[parentContentControl](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parentcontentcontrol-member)|Получает элемент управления содержимым, содержащий элемент управления содержимым.|
||[placeholderText](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-placeholdertext-member)|Возвращает или задает замещающий текст элемента управления содержимым.|
||[removeWhenEdited](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-removewhenedited-member)|Возвращает или задает значение, указывающее, удаляется ли элемент управления содержимым после изменения.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-search-member(1))|Выполняет поиск с указанными SearchOptions в области объекта управления контентом.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-select-member(1))|Выбирает элемент управления контентом.|
||[style](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-style-member)|Получает или задает имя стиля для управления контентом.|
||[tag](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-tag-member)|Возвращает или задает тег для определения элемента управления содержимым.|
||[text](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-text-member)|Возвращает текст элемента управления содержимым.|
||[заголовок](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-title-member)|Получает или задает заголовок для элемента управления содержимым.|
||[type](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-type-member)|Получает тип элемента управления содержимым.|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getById(id: number)](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getbyid-member(1))|Возвращает элемент управления содержимым по его идентификатору.|
||[getByTag(tag: string)](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getbytag-member(1))|Возвращает элементы управления содержимым с указанным тегом.|
||[getByTitle(title: string)](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getbytitle-member(1))|Возвращает элементы управления контентом с указанным заголовком.|
||[getItem(index: number)](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getitem-member(1))|Получает управление контентом по индексу в коллекции.|
||[items](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[Document](/javascript/api/word/word.document)|[body](/javascript/api/word/word.document#word-word-document-body-member)|Получает объект тела основного документа.|
||[contentControls](/javascript/api/word/word.document#word-word-document-contentcontrols-member)|Получает коллекцию объектов управления контентом в документе.|
||[getSelection()](/javascript/api/word/word.document#word-word-document-getselection-member(1))|Возвращает текущий выбранный фрагмент документа.|
||[save()](/javascript/api/word/word.document#word-word-document-save-member(1))|Сохраняет документ.|
||[сохранено](/javascript/api/word/word.document#word-word-document-saved-member)|Указывает, сохранены ли изменения, внесенные в документ.|
||[sections](/javascript/api/word/word.document#word-word-document-sections-member)|Получает коллекцию объектов раздела в документе.|
|[Font](/javascript/api/word/word.font)|[bold](/javascript/api/word/word.font#word-word-font-bold-member)|Возвращает или задает значение, указывающее, является ли шрифт полужирным.|
||[color](/javascript/api/word/word.font#word-word-font-color-member)|Возвращает или задает цвет для указанного шрифта.|
||[doubleStrikeThrough](/javascript/api/word/word.font#word-word-font-doublestrikethrough-member)|Получает или задает значение, которое указывает, имеет ли шрифт двойной удар.|
||[highlightColor](/javascript/api/word/word.font#word-word-font-highlightcolor-member)|Получает или задает цвет выделения.|
||[italic](/javascript/api/word/word.font#word-word-font-italic-member)|Возвращает или задает значение, указывающее, является ли шрифт курсивным.|
||[name](/javascript/api/word/word.font#word-word-font-name-member)|Получает или задает значение, представляющее имя шрифта.|
||[size](/javascript/api/word/word.font#word-word-font-size-member)|Получает или задает значение, представляющее размер шрифта в пунктах.|
||[strikeThrough](/javascript/api/word/word.font#word-word-font-strikethrough-member)|Получает или задает значение, которое указывает, есть ли у шрифта забастовка.|
||[subscript](/javascript/api/word/word.font#word-word-font-subscript-member)|Возвращает или задает значение, указывающее, является ли шрифт подстрочным.|
||[superscript](/javascript/api/word/word.font#word-word-font-superscript-member)|Возвращает или задает значение, указывающее, является ли шрифт надстрочным.|
||[underline](/javascript/api/word/word.font#word-word-font-underline-member)|Возвращает или задает значение, указывающее тип подчеркивания шрифта.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[altTextDescription](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-alttextdescription-member)|Получает или задает строку, представляюную альтернативный текст, связанный с изображением в строке.|
||[altTextTitle](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-alttexttitle-member)|Возвращает или задает строку, содержащую заголовок встроенного рисунка.|
||[getBase64ImageSrc()](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-getbase64imagesrc-member(1))|Возвращает строковое представление встроенного рисунка в кодировке base64.|
||[height](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-height-member)|Возвращает или задает число, которое описывает высоту встроенного рисунка.|
||[hyperlink](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-hyperlink-member)|Получает или задает гиперссылку на изображении.|
||[insertContentControl()](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-insertcontentcontrol-member(1))|Включает встроенный рисунок в элемент управления содержимым форматированного текста.|
||[lockAspectRatio](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-lockaspectratio-member)|Возвращает или задает значение, указывающее, сохраняет ли встроенный рисунок исходные пропорции при изменении размера.|
||[parentContentControl](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parentcontentcontrol-member)|Возвращает элемент управления содержимым, который содержит встроенный рисунок.|
||[width](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-width-member)|Возвращает или задает число, которое описывает ширину встроенного рисунка.|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[items](/javascript/api/word/word.inlinepicturecollection#word-word-inlinepicturecollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[Paragraph](/javascript/api/word/word.paragraph)|[выравнивание](/javascript/api/word/word.paragraph#word-word-paragraph-alignment-member)|Возвращает или задает выравнивание для абзаца.|
||[clear()](/javascript/api/word/word.paragraph#word-word-paragraph-clear-member(1))|Очищает содержимое объекта абзаца.|
||[contentControls](/javascript/api/word/word.paragraph#word-word-paragraph-contentcontrols-member)|Получает коллекцию объектов управления контентом в абзаце.|
||[delete()](/javascript/api/word/word.paragraph#word-word-paragraph-delete-member(1))|Удаляет абзац и его содержимое из документа.|
||[firstLineIndent](/javascript/api/word/word.paragraph#word-word-paragraph-firstlineindent-member)|Возвращает или задает значение отступа первой строки или выступа в пунктах.|
||[font](/javascript/api/word/word.paragraph#word-word-paragraph-font-member)|Возвращает формат текста абзаца.|
||[getHtml()](/javascript/api/word/word.paragraph#word-word-paragraph-gethtml-member(1))|Получает ПРЕДСТАВЛЕНИЕ HTML объекта абзаца.|
||[getOoxml()](/javascript/api/word/word.paragraph#word-word-paragraph-getooxml-member(1))|Возвращает OOXML-представление объекта абзаца.|
||[ignorePunct](/javascript/api/word/word.paragraph#word-word-paragraph-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.paragraph#word-word-paragraph-ignorespace-member)||
||[inlinePictures](/javascript/api/word/word.paragraph#word-word-paragraph-inlinepictures-member)|Получает коллекцию объектов InlinePicture в абзаце.|
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-insertbreak-member(1))|Вставляет разрыв в указанном расположении в основном документе.|
||[insertContentControl()](/javascript/api/word/word.paragraph#word-word-paragraph-insertcontentcontrol-member(1))|Включает объект абзаца в элемент управления содержимым форматированного текста.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-insertfilefrombase64-member(1))|Вставляет документ в абзац в указанном расположении.|
||[insertHtml (html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-inserthtml-member(1))|Вставляет HTML в абзац в указанном расположении.|
||[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-insertinlinepicturefrombase64-member(1))|Вставляет рисунок в абзац в указанном расположении.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-insertooxml-member(1))|Вставляет OOXML в абзац в указанном расположении.|
||[insertParagraph (paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-insertparagraph-member(1))|Вставляет абзац в указанном расположении.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-inserttext-member(1))|Вставляет текст в абзац в указанном расположении.|
||[leftIndent](/javascript/api/word/word.paragraph#word-word-paragraph-leftindent-member)|Возвращает или задает значение отступа слева для абзаца (в пунктах).|
||[lineSpacing](/javascript/api/word/word.paragraph#word-word-paragraph-linespacing-member)|Возвращает или задает междустрочный интервал для указанного абзаца (в пунктах).|
||[lineUnitAfter](/javascript/api/word/word.paragraph#word-word-paragraph-lineunitafter-member)|Получает или задает количество интервалов в строках сетки после абзаца.|
||[lineUnitBefore](/javascript/api/word/word.paragraph#word-word-paragraph-lineunitbefore-member)|Возвращает или устанавливает междустрочный интервал до абзаца (в линиях сетки).|
||[matchCase](/javascript/api/word/word.paragraph#word-word-paragraph-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.paragraph#word-word-paragraph-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.paragraph#word-word-paragraph-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.paragraph#word-word-paragraph-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.paragraph#word-word-paragraph-matchwildcards-member)||
||[outlineLevel](/javascript/api/word/word.paragraph#word-word-paragraph-outlinelevel-member)|Возвращает или задает уровень структуры абзаца.|
||[parentContentControl](/javascript/api/word/word.paragraph#word-word-paragraph-parentcontentcontrol-member)|Возвращает элемент управления содержимым, содержащий абзац.|
||[rightIndent](/javascript/api/word/word.paragraph#word-word-paragraph-rightindent-member)|Возвращает или задает значение отступа справа для абзаца (в пунктах).|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.paragraph#word-word-paragraph-search-member(1))|Выполняет поиск с указанными SearchOptions в области объекта абзаца.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.paragraph#word-word-paragraph-select-member(1))|Выбирает абзац и переходит к нему в пользовательском интерфейсе Word.|
||[SpaceAfter](/javascript/api/word/word.paragraph#word-word-paragraph-spaceafter-member)|Возвращает или задает междустрочный интервал после абзаца (в пунктах).|
||[spaceBefore](/javascript/api/word/word.paragraph#word-word-paragraph-spacebefore-member)|Возвращает или задает междустрочный интервал до абзаца (в пунктах).|
||[style](/javascript/api/word/word.paragraph#word-word-paragraph-style-member)|Получает или задает имя стиля для абзаца.|
||[text](/javascript/api/word/word.paragraph#word-word-paragraph-text-member)|Возвращает текст абзаца.|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[items](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[Range](/javascript/api/word/word.range)|[clear()](/javascript/api/word/word.range#word-word-range-clear-member(1))|Очищает содержимое объекта диапазона.|
||[contentControls](/javascript/api/word/word.range#word-word-range-contentcontrols-member)|Получает коллекцию объектов управления контентом в диапазоне.|
||[delete()](/javascript/api/word/word.range#word-word-range-delete-member(1))|Удаляет диапазон и его содержимое из документа.|
||[font](/javascript/api/word/word.range#word-word-range-font-member)|Возвращает формат текста диапазона.|
||[getHtml()](/javascript/api/word/word.range#word-word-range-gethtml-member(1))|Получает HTML-представление объекта диапазона.|
||[getOoxml()](/javascript/api/word/word.range#word-word-range-getooxml-member(1))|Возвращает OOXML-представление объекта диапазона.|
||[ignorePunct](/javascript/api/word/word.range#word-word-range-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.range#word-word-range-ignorespace-member)||
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#word-word-range-insertbreak-member(1))|Вставляет разрыв в указанном расположении в основном документе.|
||[insertContentControl()](/javascript/api/word/word.range#word-word-range-insertcontentcontrol-member(1))|Включает объект диапазона в элемент управления содержимым форматированного текста.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#word-word-range-insertfilefrombase64-member(1))|Вставляет документ в указанном расположении.|
||[insertHtml (html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#word-word-range-inserthtml-member(1))|Вставляет HTML-код в указанном расположении.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#word-word-range-insertooxml-member(1))|Вставляет OOXML-код в указанном расположении.|
||[insertParagraph (paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#word-word-range-insertparagraph-member(1))|Вставляет абзац в указанном расположении.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#word-word-range-inserttext-member(1))|Вставляет текст в заданном расположении.|
||[matchCase](/javascript/api/word/word.range#word-word-range-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.range#word-word-range-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.range#word-word-range-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.range#word-word-range-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.range#word-word-range-matchwildcards-member)||
||[paragraphs](/javascript/api/word/word.range#word-word-range-paragraphs-member)|Получает коллекцию объектов абзаца в диапазоне.|
||[parentContentControl](/javascript/api/word/word.range#word-word-range-parentcontentcontrol-member)|Возвращает элемент управления содержимым, содержащий диапазон.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.range#word-word-range-search-member(1))|Выполняет поиск с указанными SearchOptions в области объекта диапазона.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.range#word-word-range-select-member(1))|Выбор диапазона и переход к нему в пользовательском интерфейсе Word.|
||[style](/javascript/api/word/word.range#word-word-range-style-member)|Получает или задает имя стиля для диапазона.|
||[text](/javascript/api/word/word.range#word-word-range-text-member)|Возвращает текст диапазона.|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[items](/javascript/api/word/word.rangecollection#word-word-rangecollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[SearchOptions](/javascript/api/word/word.searchoptions)|[ignorePunct](/javascript/api/word/word.searchoptions#word-word-searchoptions-ignorepunct-member)|Возвращает или задает значение, которое указывает, следует ли пропустить все знаки препинания между словами.|
||[ignoreSpace](/javascript/api/word/word.searchoptions#word-word-searchoptions-ignorespace-member)|Получает или задает значение, которое указывает, следует ли игнорировать все белое пространство между словами.|
||[matchCase](/javascript/api/word/word.searchoptions#word-word-searchoptions-matchcase-member)|Возвращает или задает значение, которое указывает, следует ли выполнять поиск с учетом регистра.|
||[matchPrefix](/javascript/api/word/word.searchoptions#word-word-searchoptions-matchprefix-member)|Возвращает или задает значение, которое указывает, нужно ли учитывать слова, начинающиеся со строки поиска.|
||[matchSuffix](/javascript/api/word/word.searchoptions#word-word-searchoptions-matchsuffix-member)|Возвращает или задает значение, указывающее, нужно ли учитывать слова, которые заканчиваются строкой поиска.|
||[matchWholeWord](/javascript/api/word/word.searchoptions#word-word-searchoptions-matchwholeword-member)|Возвращает или задает значение, которое указывает, следует ли искать только целые слова, а не текст, являющийся частью большего слова.|
||[matchWildcards](/javascript/api/word/word.searchoptions#word-word-searchoptions-matchwildcards-member)|Возвращает или задает значение, которое указывает, будет ли выполняться поиск с использованием специальных операторов поиска.|
|[Section](/javascript/api/word/word.section)|[body](/javascript/api/word/word.section#word-word-section-body-member)|Получает объект тела раздела.|
||[getFooter (тип: Word.HeaderFooterType)](/javascript/api/word/word.section#word-word-section-getfooter-member(1))|Возвращает один из нижних колонтитулов раздела.|
||[getHeader (тип: Word.HeaderFooterType)](/javascript/api/word/word.section#word-word-section-getheader-member(1))|Возвращает один из верхних колонтитулов раздела.|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[items](/javascript/api/word/word.sectioncollection#word-word-sectioncollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Word](/javascript/api/word)
- [Наборы обязательных элементов API JavaScript для Word](word-api-requirement-sets.md)
