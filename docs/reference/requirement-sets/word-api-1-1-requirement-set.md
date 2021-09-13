---
title: Набор API API Word JavaScript 1.1
description: Сведения о наборе требований WordApi 1.1
ms.date: 11/09/2020
ms.prod: word
ms.localizationpriority: medium
ms.openlocfilehash: 43d2eba2180c66f4037b2f4a1742ceae61d7c353
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154829"
---
# <a name="whats-new-in-word-javascript-api-11"></a>Что нового в API JavaScript Word 1.1

WordApi 1.1 — это первый набор требований API Word JavaScript. Это единственный набор API Word, поддерживаемый Word 2016.

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в API Word JavaScript, за набором 1.1. Чтобы просмотреть справочную документацию API для всех API, поддерживаемых требованием API Word JavaScript, за набором 1.1, см. в справке к API Word в наборе [требований 1.1](/javascript/api/word?view=word-js-1.1&preserve-view=true).

| Класс | Поля | Описание |
|:---|:---|:---|
|[Основной текст](/javascript/api/word/word.body)|[clear()](/javascript/api/word/word.body#clear__)|Очищает объект содержимого.|
||[getHtml()](/javascript/api/word/word.body#getHtml__)|Получает HTML-представление объекта тела.|
||[getOoxml()](/javascript/api/word/word.body#getOoxml__)|Возвращает OOXML-представление (Office Open XML) объекта содержимого.|
||[ignorePunct](/javascript/api/word/word.body#ignorePunct)||
||[ignoreSpace](/javascript/api/word/word.body#ignoreSpace)||
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#insertBreak_breakType__insertLocation_)|Вставляет разрыв в указанном расположении в основном документе.|
||[insertContentControl()](/javascript/api/word/word.body#insertContentControl__)|Включает объект содержимого в элемент управления форматированным текстом.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#insertFileFromBase64_base64File__insertLocation_)|Вставляет документ в содержимое в заданном расположении.|
||[insertHtml (html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#insertHtml_html__insertLocation_)|Вставляет HTML-код в указанном расположении.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#insertOoxml_ooxml__insertLocation_)|Вставляет OOXML-код в указанном расположении.|
||[insertParagraph (paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#insertParagraph_paragraphText__insertLocation_)|Вставляет абзац в указанном расположении.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.body#insertText_text__insertLocation_)|Вставляет текст в содержимое в заданном расположении.|
||[matchCase](/javascript/api/word/word.body#matchCase)||
||[matchPrefix](/javascript/api/word/word.body#matchPrefix)||
||[matchSuffix](/javascript/api/word/word.body#matchSuffix)||
||[matchWholeWord](/javascript/api/word/word.body#matchWholeWord)||
||[matchWildcards](/javascript/api/word/word.body#matchWildcards)||
||[contentControls](/javascript/api/word/word.body#contentControls)|Получает коллекцию объектов управления текстовым контентом в теле.|
||[font](/javascript/api/word/word.body#font)|Получает формат текста, указанный для содержимого документа или раздела.|
||[inlinePictures](/javascript/api/word/word.body#inlinePictures)|Получает коллекцию объектов InlinePicture в теле.|
||[paragraphs](/javascript/api/word/word.body#paragraphs)|Получает коллекцию объектов абзаца в теле.|
||[parentContentControl](/javascript/api/word/word.body#parentContentControl)|Получает элемент управления содержимым, содержащий документ или раздел.|
||[text](/javascript/api/word/word.body#text)|Возвращает текст содержимого.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.body#search_searchText__searchOptions__ignorePunct__ignoreSpace__matchCase__matchPrefix__matchSuffix__matchWholeWord__matchWildcards_)|Выполняет поиск с указанными SearchOptions в области объекта тела.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.body#select_selectionMode_)|Выбирает содержимое и переходит к нему в пользовательском интерфейсе Word.|
||[style](/javascript/api/word/word.body#style)|Получает или задает имя стиля для тела.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[внешний вид](/javascript/api/word/word.contentcontrol#appearance)|Получает или задает внешний вид элемента управления содержимым.|
||[cannotDelete](/javascript/api/word/word.contentcontrol#cannotDelete)|Возвращает или задает значение, указывающее, может ли пользователь удалить элемент управления содержимым.|
||[cannotEdit](/javascript/api/word/word.contentcontrol#cannotEdit)|Возвращает или задает значение, указывающее, может ли пользователь изменять содержимое элемента управления содержимым.|
||[clear()](/javascript/api/word/word.contentcontrol#clear__)|Очищает содержимое элемента управления содержимым.|
||[color](/javascript/api/word/word.contentcontrol#color)|Возвращает или задает цвет элемента управления содержимым.|
||[delete(keepContent: boolean)](/javascript/api/word/word.contentcontrol#delete_keepContent_)|Удаляет элемент управления содержимым и его содержимое.|
||[getHtml()](/javascript/api/word/word.contentcontrol#getHtml__)|Получает ПРЕДСТАВЛЕНИЕ HTML объекта управления контентом.|
||[getOoxml()](/javascript/api/word/word.contentcontrol#getOoxml__)|Возвращает OOXML-представление объекта элемента управления содержимым.|
||[ignorePunct](/javascript/api/word/word.contentcontrol#ignorePunct)||
||[ignoreSpace](/javascript/api/word/word.contentcontrol#ignoreSpace)||
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#insertBreak_breakType__insertLocation_)|Вставляет разрыв в указанном расположении в основном документе.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#insertFileFromBase64_base64File__insertLocation_)|Вставляет документ в управление контентом в указанном расположении.|
||[insertHtml (html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#insertHtml_html__insertLocation_)|Вставляет HTML-код в элемент управления содержимым в указанном расположении.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#insertOoxml_ooxml__insertLocation_)|Вставляет OOXML в управление контентом в указанном расположении.|
||[insertParagraph (paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#insertParagraph_paragraphText__insertLocation_)|Вставляет абзац в указанном расположении.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.contentcontrol#insertText_text__insertLocation_)|Вставляет текст в элемент управления содержимым в указанном расположении.|
||[matchCase](/javascript/api/word/word.contentcontrol#matchCase)||
||[matchPrefix](/javascript/api/word/word.contentcontrol#matchPrefix)||
||[matchSuffix](/javascript/api/word/word.contentcontrol#matchSuffix)||
||[matchWholeWord](/javascript/api/word/word.contentcontrol#matchWholeWord)||
||[matchWildcards](/javascript/api/word/word.contentcontrol#matchWildcards)||
||[placeholderText](/javascript/api/word/word.contentcontrol#placeholderText)|Возвращает или задает замещающий текст элемента управления содержимым.|
||[contentControls](/javascript/api/word/word.contentcontrol#contentControls)|Получает коллекцию объектов элементов управления содержимым в элементе управления содержимым.|
||[font](/javascript/api/word/word.contentcontrol#font)|Получает текстовый формат элемента управления содержимым.|
||[id](/javascript/api/word/word.contentcontrol#id)|Возвращает целое число, представляющее собой идентификатор элемента управления контентом.|
||[inlinePictures](/javascript/api/word/word.contentcontrol#inlinePictures)|Получает коллекцию объектов inlinePicture в элементе управления содержимым.|
||[paragraphs](/javascript/api/word/word.contentcontrol#paragraphs)|Получает коллекцию объектов абзацев в элементе управления содержимым.|
||[parentContentControl](/javascript/api/word/word.contentcontrol#parentContentControl)|Получает элемент управления содержимым, содержащий элемент управления содержимым.|
||[text](/javascript/api/word/word.contentcontrol#text)|Возвращает текст элемента управления содержимым.|
||[type](/javascript/api/word/word.contentcontrol#type)|Получает тип элемента управления содержимым.|
||[removeWhenEdited](/javascript/api/word/word.contentcontrol#removeWhenEdited)|Возвращает или задает значение, указывающее, удаляется ли элемент управления содержимым после изменения.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.contentcontrol#search_searchText__searchOptions__ignorePunct__ignoreSpace__matchCase__matchPrefix__matchSuffix__matchWholeWord__matchWildcards_)|Выполняет поиск с указанными SearchOptions в области объекта управления контентом.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.contentcontrol#select_selectionMode_)|Выбирает элемент управления контентом.|
||[style](/javascript/api/word/word.contentcontrol#style)|Получает или задает имя стиля для управления контентом.|
||[tag](/javascript/api/word/word.contentcontrol#tag)|Возвращает или задает тег для определения элемента управления содержимым.|
||[заголовок](/javascript/api/word/word.contentcontrol#title)|Получает или задает заголовок для элемента управления содержимым.|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getById(id: number)](/javascript/api/word/word.contentcontrolcollection#getById_id_)|Возвращает элемент управления содержимым по его идентификатору.|
||[getByTag(tag: string)](/javascript/api/word/word.contentcontrolcollection#getByTag_tag_)|Возвращает элементы управления содержимым с указанным тегом.|
||[getByTitle(title: string)](/javascript/api/word/word.contentcontrolcollection#getByTitle_title_)|Возвращает элементы управления контентом с указанным заголовком.|
||[getItem(index: number)](/javascript/api/word/word.contentcontrolcollection#getItem_index_)|Получает управление контентом по индексу в коллекции.|
||[items](/javascript/api/word/word.contentcontrolcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Document](/javascript/api/word/word.document)|[getSelection()](/javascript/api/word/word.document#getSelection__)|Возвращает текущий выбранный фрагмент документа.|
||[body](/javascript/api/word/word.document#body)|Получает объект тела документа.|
||[contentControls](/javascript/api/word/word.document#contentControls)|Получает коллекцию объектов управления контентом в документе.|
||[сохранено](/javascript/api/word/word.document#saved)|Указывает, сохранены ли изменения, внесенные в документ.|
||[sections](/javascript/api/word/word.document#sections)|Получает коллекцию объектов раздела в документе.|
||[save()](/javascript/api/word/word.document#save__)|Сохраняет документ.|
|[Font](/javascript/api/word/word.font)|[bold](/javascript/api/word/word.font#bold)|Возвращает или задает значение, указывающее, является ли шрифт полужирным.|
||[color](/javascript/api/word/word.font#color)|Возвращает или задает цвет для указанного шрифта.|
||[doubleStrikeThrough](/javascript/api/word/word.font#doubleStrikeThrough)|Получает или задает значение, которое указывает, имеет ли шрифт двойной удар.|
||[highlightColor](/javascript/api/word/word.font#highlightColor)|Получает или задает цвет выделения.|
||[italic](/javascript/api/word/word.font#italic)|Возвращает или задает значение, указывающее, является ли шрифт курсивным.|
||[name](/javascript/api/word/word.font#name)|Получает или задает значение, представляющее имя шрифта.|
||[size](/javascript/api/word/word.font#size)|Получает или задает значение, представляющее размер шрифта в пунктах.|
||[strikeThrough](/javascript/api/word/word.font#strikeThrough)|Получает или задает значение, которое указывает, есть ли у шрифта забастовка.|
||[subscript](/javascript/api/word/word.font#subscript)|Возвращает или задает значение, указывающее, является ли шрифт подстрочным.|
||[superscript](/javascript/api/word/word.font#superscript)|Возвращает или задает значение, указывающее, является ли шрифт надстрочным.|
||[underline](/javascript/api/word/word.font#underline)|Возвращает или задает значение, указывающее тип подчеркивания шрифта.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[altTextDescription](/javascript/api/word/word.inlinepicture#altTextDescription)|Получает или задает строку, представляюную альтернативный текст, связанный с изображением в строке.|
||[altTextTitle](/javascript/api/word/word.inlinepicture#altTextTitle)|Возвращает или задает строку, содержащую заголовок встроенного рисунка.|
||[getBase64ImageSrc()](/javascript/api/word/word.inlinepicture#getBase64ImageSrc__)|Возвращает строковое представление встроенного рисунка в кодировке base64.|
||[height](/javascript/api/word/word.inlinepicture#height)|Возвращает или задает число, которое описывает высоту встроенного рисунка.|
||[hyperlink](/javascript/api/word/word.inlinepicture#hyperlink)|Получает или задает гиперссылку на изображении.|
||[insertContentControl()](/javascript/api/word/word.inlinepicture#insertContentControl__)|Включает встроенный рисунок в элемент управления содержимым форматированного текста.|
||[lockAspectRatio](/javascript/api/word/word.inlinepicture#lockAspectRatio)|Возвращает или задает значение, указывающее, сохраняет ли встроенный рисунок исходные пропорции при изменении размера.|
||[parentContentControl](/javascript/api/word/word.inlinepicture#parentContentControl)|Возвращает элемент управления содержимым, который содержит встроенный рисунок.|
||[width](/javascript/api/word/word.inlinepicture#width)|Возвращает или задает число, которое описывает ширину встроенного рисунка.|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[items](/javascript/api/word/word.inlinepicturecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Paragraph](/javascript/api/word/word.paragraph)|[выравнивание](/javascript/api/word/word.paragraph#alignment)|Возвращает или задает выравнивание для абзаца.|
||[clear()](/javascript/api/word/word.paragraph#clear__)|Очищает содержимое объекта абзаца.|
||[delete()](/javascript/api/word/word.paragraph#delete__)|Удаляет абзац и его содержимое из документа.|
||[firstLineIndent](/javascript/api/word/word.paragraph#firstLineIndent)|Возвращает или задает значение отступа первой строки или выступа в пунктах.|
||[getHtml()](/javascript/api/word/word.paragraph#getHtml__)|Получает ПРЕДСТАВЛЕНИЕ HTML объекта абзаца.|
||[getOoxml()](/javascript/api/word/word.paragraph#getOoxml__)|Возвращает OOXML-представление объекта абзаца.|
||[ignorePunct](/javascript/api/word/word.paragraph#ignorePunct)||
||[ignoreSpace](/javascript/api/word/word.paragraph#ignoreSpace)||
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#insertBreak_breakType__insertLocation_)|Вставляет разрыв в указанном расположении в основном документе.|
||[insertContentControl()](/javascript/api/word/word.paragraph#insertContentControl__)|Включает объект абзаца в элемент управления содержимым форматированного текста.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#insertFileFromBase64_base64File__insertLocation_)|Вставляет документ в абзац в указанном расположении.|
||[insertHtml (html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#insertHtml_html__insertLocation_)|Вставляет HTML в абзац в указанном расположении.|
||[insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#insertInlinePictureFromBase64_base64EncodedImage__insertLocation_)|Вставляет рисунок в абзац в указанном расположении.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#insertOoxml_ooxml__insertLocation_)|Вставляет OOXML в абзац в указанном расположении.|
||[insertParagraph (paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#insertParagraph_paragraphText__insertLocation_)|Вставляет абзац в указанном расположении.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.paragraph#insertText_text__insertLocation_)|Вставляет текст в абзац в указанном расположении.|
||[leftIndent](/javascript/api/word/word.paragraph#leftIndent)|Возвращает или задает значение отступа слева для абзаца (в пунктах).|
||[lineSpacing](/javascript/api/word/word.paragraph#lineSpacing)|Возвращает или задает междустрочный интервал для указанного абзаца (в пунктах).|
||[lineUnitAfter](/javascript/api/word/word.paragraph#lineUnitAfter)|Получает или задает количество интервалов в строках сетки после абзаца.|
||[lineUnitBefore](/javascript/api/word/word.paragraph#lineUnitBefore)|Возвращает или устанавливает междустрочный интервал до абзаца (в линиях сетки).|
||[matchCase](/javascript/api/word/word.paragraph#matchCase)||
||[matchPrefix](/javascript/api/word/word.paragraph#matchPrefix)||
||[matchSuffix](/javascript/api/word/word.paragraph#matchSuffix)||
||[matchWholeWord](/javascript/api/word/word.paragraph#matchWholeWord)||
||[matchWildcards](/javascript/api/word/word.paragraph#matchWildcards)||
||[outlineLevel](/javascript/api/word/word.paragraph#outlineLevel)|Возвращает или задает уровень структуры абзаца.|
||[contentControls](/javascript/api/word/word.paragraph#contentControls)|Получает коллекцию объектов управления контентом в абзаце.|
||[font](/javascript/api/word/word.paragraph#font)|Возвращает формат текста абзаца.|
||[inlinePictures](/javascript/api/word/word.paragraph#inlinePictures)|Получает коллекцию объектов InlinePicture в абзаце.|
||[parentContentControl](/javascript/api/word/word.paragraph#parentContentControl)|Возвращает элемент управления содержимым, содержащий абзац.|
||[text](/javascript/api/word/word.paragraph#text)|Возвращает текст абзаца.|
||[rightIndent](/javascript/api/word/word.paragraph#rightIndent)|Возвращает или задает значение отступа справа для абзаца (в пунктах).|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.paragraph#search_searchText__searchOptions__ignorePunct__ignoreSpace__matchCase__matchPrefix__matchSuffix__matchWholeWord__matchWildcards_)|Выполняет поиск с указанными SearchOptions в области объекта абзаца.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.paragraph#select_selectionMode_)|Выбирает абзац и переходит к нему в пользовательском интерфейсе Word.|
||[SpaceAfter](/javascript/api/word/word.paragraph#spaceAfter)|Возвращает или задает междустрочный интервал после абзаца (в пунктах).|
||[spaceBefore](/javascript/api/word/word.paragraph#spaceBefore)|Возвращает или задает междустрочный интервал до абзаца (в пунктах).|
||[style](/javascript/api/word/word.paragraph#style)|Получает или задает имя стиля для абзаца.|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[items](/javascript/api/word/word.paragraphcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Range](/javascript/api/word/word.range)|[clear()](/javascript/api/word/word.range#clear__)|Очищает содержимое объекта диапазона.|
||[delete()](/javascript/api/word/word.range#delete__)|Удаляет диапазон и его содержимое из документа.|
||[getHtml()](/javascript/api/word/word.range#getHtml__)|Получает HTML-представление объекта диапазона.|
||[getOoxml()](/javascript/api/word/word.range#getOoxml__)|Возвращает OOXML-представление объекта диапазона.|
||[ignorePunct](/javascript/api/word/word.range#ignorePunct)||
||[ignoreSpace](/javascript/api/word/word.range#ignoreSpace)||
||[insertBreak(breakType: Word.BreakType, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#insertBreak_breakType__insertLocation_)|Вставляет разрыв в указанном расположении в основном документе.|
||[insertContentControl()](/javascript/api/word/word.range#insertContentControl__)|Включает объект диапазона в элемент управления содержимым форматированного текста.|
||[insertFileFromBase64(base64File: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#insertFileFromBase64_base64File__insertLocation_)|Вставляет документ в указанном расположении.|
||[insertHtml (html: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#insertHtml_html__insertLocation_)|Вставляет HTML-код в указанном расположении.|
||[insertOoxml(ooxml: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#insertOoxml_ooxml__insertLocation_)|Вставляет OOXML-код в указанном расположении.|
||[insertParagraph (paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#insertParagraph_paragraphText__insertLocation_)|Вставляет абзац в указанном расположении.|
||[insertText(text: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.range#insertText_text__insertLocation_)|Вставляет текст в заданном расположении.|
||[matchCase](/javascript/api/word/word.range#matchCase)||
||[matchPrefix](/javascript/api/word/word.range#matchPrefix)||
||[matchSuffix](/javascript/api/word/word.range#matchSuffix)||
||[matchWholeWord](/javascript/api/word/word.range#matchWholeWord)||
||[matchWildcards](/javascript/api/word/word.range#matchWildcards)||
||[contentControls](/javascript/api/word/word.range#contentControls)|Получает коллекцию объектов управления контентом в диапазоне.|
||[font](/javascript/api/word/word.range#font)|Возвращает формат текста диапазона.|
||[paragraphs](/javascript/api/word/word.range#paragraphs)|Получает коллекцию объектов абзаца в диапазоне.|
||[parentContentControl](/javascript/api/word/word.range#parentContentControl)|Возвращает элемент управления содержимым, содержащий диапазон.|
||[text](/javascript/api/word/word.range#text)|Возвращает текст диапазона.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.range#search_searchText__searchOptions__ignorePunct__ignoreSpace__matchCase__matchPrefix__matchSuffix__matchWholeWord__matchWildcards_)|Выполняет поиск с указанными SearchOptions в области объекта диапазона.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.range#select_selectionMode_)|Выбор диапазона и переход к нему в пользовательском интерфейсе Word.|
||[style](/javascript/api/word/word.range#style)|Получает или задает имя стиля для диапазона.|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[items](/javascript/api/word/word.rangecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[SearchOptions](/javascript/api/word/word.searchoptions)|[ignorePunct](/javascript/api/word/word.searchoptions#ignorePunct)|Возвращает или задает значение, которое указывает, следует ли пропустить все знаки препинания между словами.|
||[ignoreSpace](/javascript/api/word/word.searchoptions#ignoreSpace)|Получает или задает значение, которое указывает, следует ли игнорировать все белое пространство между словами.|
||[matchCase](/javascript/api/word/word.searchoptions#matchCase)|Возвращает или задает значение, которое указывает, следует ли выполнять поиск с учетом регистра.|
||[matchPrefix](/javascript/api/word/word.searchoptions#matchPrefix)|Возвращает или задает значение, которое указывает, нужно ли учитывать слова, начинающиеся со строки поиска.|
||[matchSuffix](/javascript/api/word/word.searchoptions#matchSuffix)|Возвращает или задает значение, указывающее, нужно ли учитывать слова, которые заканчиваются строкой поиска.|
||[matchWholeWord](/javascript/api/word/word.searchoptions#matchWholeWord)|Возвращает или задает значение, которое указывает, следует ли искать только целые слова, а не текст, являющийся частью большего слова.|
||[matchWildcards](/javascript/api/word/word.searchoptions#matchWildcards)|Возвращает или задает значение, которое указывает, будет ли выполняться поиск с использованием специальных операторов поиска.|
|[Section](/javascript/api/word/word.section)|[getFooter (тип: Word.HeaderFooterType)](/javascript/api/word/word.section#getFooter_type_)|Возвращает один из нижних колонтитулов раздела.|
||[getHeader (тип: Word.HeaderFooterType)](/javascript/api/word/word.section#getHeader_type_)|Возвращает один из верхних колонтитулов раздела.|
||[body](/javascript/api/word/word.section#body)|Получает объект тела раздела.|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[items](/javascript/api/word/word.sectioncollection#items)|Получает загруженные дочерние элементы в этой коллекции.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Word](/javascript/api/word)
- [Наборы обязательных элементов API JavaScript для Word](word-api-requirement-sets.md)
