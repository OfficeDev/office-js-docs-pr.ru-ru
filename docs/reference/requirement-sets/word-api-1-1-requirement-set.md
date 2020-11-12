---
title: Набор обязательных элементов API JavaScript для Word 1,1
description: Сведения о наборе требований WordApi 1,1
ms.date: 11/09/2020
ms.prod: word
localization_priority: Normal
ms.openlocfilehash: 371638c18cff882f2b3907f1adedb6748761cc0c
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996440"
---
# <a name="whats-new-in-word-javascript-api-11"></a>Новые возможности API JavaScript для Word 1,1

WordApi 1,1 является первым набором требований API JavaScript для Word. Это единственный набор требований Word API, поддерживаемый Word 2016.

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в наборе обязательных элементов API JavaScript для Word 1,1. Чтобы просмотреть справочную документацию по API для всех API, поддерживаемых в наборе обязательных элементов API JavaScript для Word 1,1, обратитесь к разделам [API Word в наборе требований 1,1](/javascript/api/word?view=word-js-1.1&preserve-view=true).

| Класс | Поля | Описание |
|:---|:---|:---|
|[Основной текст](/javascript/api/word/word.body)|[clear()](/javascript/api/word/word.body#clear--)|Очищает объект содержимого.|
||[getHtml()](/javascript/api/word/word.body#gethtml--)|Возвращает HTML-представление объекта Body.|
||[getOoxml()](/javascript/api/word/word.body#getooxml--)|Возвращает OOXML-представление (Office Open XML) объекта содержимого.|
||[игнорепункт](/javascript/api/word/word.body#ignorepunct)||
||[игнореспаце](/javascript/api/word/word.body#ignorespace)||
||[insertBreak (breakType: Word. BreakType, insertLocation: Word. InsertLocation)](/javascript/api/word/word.body#insertbreak-breaktype--insertlocation-)|Вставляет разрыв в указанном расположении в основном документе.|
||[insertContentControl()](/javascript/api/word/word.body#insertcontentcontrol--)|Включает объект содержимого в элемент управления форматированным текстом.|
||[insertFileFromBase64 (base64File: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.body#insertfilefrombase64-base64file--insertlocation-)|Вставляет документ в содержимое в заданном расположении.|
||[insertHtml (HTML: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.body#inserthtml-html--insertlocation-)|Вставляет HTML-код в указанном расположении.|
||[insertOoxml (OOXML: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.body#insertooxml-ooxml--insertlocation-)|Вставляет OOXML-код в указанном расположении.|
||[insertParagraph (paragraphText: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.body#insertparagraph-paragraphtext--insertlocation-)|Вставляет абзац в указанном расположении.|
||[insertText (Text: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.body#inserttext-text--insertlocation-)|Вставляет текст в содержимое в заданном расположении.|
||[matchCase](/javascript/api/word/word.body#matchcase)||
||[матчпрефикс](/javascript/api/word/word.body#matchprefix)||
||[матчсуффикс](/javascript/api/word/word.body#matchsuffix)||
||[матчвхолеворд](/javascript/api/word/word.body#matchwholeword)||
||[матчвилдкардс](/javascript/api/word/word.body#matchwildcards)||
||[contentControls](/javascript/api/word/word.body#contentcontrols)|Возвращает коллекцию объектов элемента управления содержимым "форматированный текст" в тексте.|
||[font](/javascript/api/word/word.body#font)|Получает формат текста, указанный для содержимого документа или раздела.|
||[inlinePictures](/javascript/api/word/word.body#inlinepictures)|Получает коллекцию объектов коллекцию inlinepicture в тексте.|
||[paragraphs](/javascript/api/word/word.body#paragraphs)|Возвращает коллекцию объектов абзаца в тексте.|
||[parentContentControl](/javascript/api/word/word.body#parentcontentcontrol)|Получает элемент управления содержимым, содержащий документ или раздел.|
||[text](/javascript/api/word/word.body#text)|Возвращает текст содержимого.|
||[Search (searchText: строка, searchOptions?: Word. SearchOptions \| {игнорепункт?: Boolean игнореспаце?: Boolean matchCase?: Boolean матчпрефикс?: Boolean матчсуффикс?: Boolean матчвхолеворд?: Boolean матчвилдкардс?: Boolean})](/javascript/api/word/word.body#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Выполняет поиск с указанным SearchOptions в области объекта Body.|
||[SELECT (selectionMode?: Word. SelectionMode)](/javascript/api/word/word.body#select-selectionmode-)|Выбирает содержимое и переходит к нему в пользовательском интерфейсе Word.|
||[style](/javascript/api/word/word.body#style)|Получает или задает имя стиля для основного текста.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[виде](/javascript/api/word/word.contentcontrol#appearance)|Получает или задает внешний вид элемента управления содержимым.|
||[каннотделете](/javascript/api/word/word.contentcontrol#cannotdelete)|Возвращает или задает значение, указывающее, может ли пользователь удалить элемент управления содержимым.|
||[каннотедит](/javascript/api/word/word.contentcontrol#cannotedit)|Возвращает или задает значение, указывающее, может ли пользователь изменять содержимое элемента управления содержимым.|
||[clear()](/javascript/api/word/word.contentcontrol#clear--)|Очищает содержимое элемента управления содержимым.|
||[color](/javascript/api/word/word.contentcontrol#color)|Возвращает или задает цвет элемента управления содержимым.|
||[Delete (Кипконтент: Boolean)](/javascript/api/word/word.contentcontrol#delete-keepcontent-)|Удаляет элемент управления содержимым и его содержимое.|
||[getHtml()](/javascript/api/word/word.contentcontrol#gethtml--)|Возвращает HTML-представление объекта элемента управления содержимым.|
||[getOoxml()](/javascript/api/word/word.contentcontrol#getooxml--)|Возвращает OOXML-представление объекта элемента управления содержимым.|
||[игнорепункт](/javascript/api/word/word.contentcontrol#ignorepunct)||
||[игнореспаце](/javascript/api/word/word.contentcontrol#ignorespace)||
||[insertBreak (breakType: Word. BreakType, insertLocation: Word. InsertLocation)](/javascript/api/word/word.contentcontrol#insertbreak-breaktype--insertlocation-)|Вставляет разрыв в указанном расположении в основном документе.|
||[insertFileFromBase64 (base64File: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.contentcontrol#insertfilefrombase64-base64file--insertlocation-)|Вставляет документ в элемент управления содержимым в указанном расположении.|
||[insertHtml (HTML: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.contentcontrol#inserthtml-html--insertlocation-)|Вставляет HTML-код в элемент управления содержимым в указанном расположении.|
||[insertOoxml (OOXML: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.contentcontrol#insertooxml-ooxml--insertlocation-)|Вставляет OOXML в элемент управления содержимым в указанном расположении.|
||[insertParagraph (paragraphText: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.contentcontrol#insertparagraph-paragraphtext--insertlocation-)|Вставляет абзац в указанном расположении.|
||[insertText (Text: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.contentcontrol#inserttext-text--insertlocation-)|Вставляет текст в элемент управления содержимым в указанном расположении.|
||[matchCase](/javascript/api/word/word.contentcontrol#matchcase)||
||[матчпрефикс](/javascript/api/word/word.contentcontrol#matchprefix)||
||[матчсуффикс](/javascript/api/word/word.contentcontrol#matchsuffix)||
||[матчвхолеворд](/javascript/api/word/word.contentcontrol#matchwholeword)||
||[матчвилдкардс](/javascript/api/word/word.contentcontrol#matchwildcards)||
||[плацехолдертекст](/javascript/api/word/word.contentcontrol#placeholdertext)|Возвращает или задает замещающий текст элемента управления содержимым.|
||[contentControls](/javascript/api/word/word.contentcontrol#contentcontrols)|Получает коллекцию объектов элементов управления содержимым в элементе управления содержимым.|
||[font](/javascript/api/word/word.contentcontrol#font)|Получает текстовый формат элемента управления содержимым.|
||[id](/javascript/api/word/word.contentcontrol#id)|Возвращает целое число, представляющее собой идентификатор элемента управления контентом.|
||[inlinePictures](/javascript/api/word/word.contentcontrol#inlinepictures)|Получает коллекцию объектов inlinePicture в элементе управления содержимым.|
||[paragraphs](/javascript/api/word/word.contentcontrol#paragraphs)|Получает коллекцию объектов абзацев в элементе управления содержимым.|
||[parentContentControl](/javascript/api/word/word.contentcontrol#parentcontentcontrol)|Получает элемент управления содержимым, содержащий элемент управления содержимым.|
||[text](/javascript/api/word/word.contentcontrol#text)|Возвращает текст элемента управления содержимым.|
||[type](/javascript/api/word/word.contentcontrol#type)|Получает тип элемента управления содержимым.|
||[ремовевхенедитед](/javascript/api/word/word.contentcontrol#removewhenedited)|Возвращает или задает значение, указывающее, удаляется ли элемент управления содержимым после изменения.|
||[Search (searchText: строка, searchOptions?: Word. SearchOptions \| {игнорепункт?: Boolean игнореспаце?: Boolean matchCase?: Boolean матчпрефикс?: Boolean матчсуффикс?: Boolean матчвхолеворд?: Boolean матчвилдкардс?: Boolean})](/javascript/api/word/word.contentcontrol#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Выполняет поиск с указанным SearchOptions в области объекта элемента управления содержимым.|
||[SELECT (selectionMode?: Word. SelectionMode)](/javascript/api/word/word.contentcontrol#select-selectionmode-)|Выбирает элемент управления контентом.|
||[style](/javascript/api/word/word.contentcontrol#style)|Получает или задает имя стиля для элемента управления содержимым.|
||[мечать](/javascript/api/word/word.contentcontrol#tag)|Возвращает или задает тег для определения элемента управления содержимым.|
||[заголовок](/javascript/api/word/word.contentcontrol#title)|Получает или задает заголовок для элемента управления содержимым.|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getById(id: number)](/javascript/api/word/word.contentcontrolcollection#getbyid-id-)|Возвращает элемент управления содержимым по его идентификатору.|
||[getByTag(tag: string)](/javascript/api/word/word.contentcontrolcollection#getbytag-tag-)|Возвращает элементы управления содержимым с указанным тегом.|
||[getByTitle(title: string)](/javascript/api/word/word.contentcontrolcollection#getbytitle-title-)|Возвращает элементы управления контентом с указанным заголовком.|
||[getItem(index: number)](/javascript/api/word/word.contentcontrolcollection#getitem-index-)|Возвращает элемент управления контентом по его индексу в коллекции.|
||[items](/javascript/api/word/word.contentcontrolcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Document](/javascript/api/word/word.document)|[При выборе ()](/javascript/api/word/word.document#getselection--)|Возвращает текущий выбранный фрагмент документа.|
||[body](/javascript/api/word/word.document#body)|Возвращает объект Body документа.|
||[contentControls](/javascript/api/word/word.document#contentcontrols)|Возвращает коллекцию объектов элементов управления содержимым в документе.|
||[сохраняем](/javascript/api/word/word.document#saved)|Указывает, сохранены ли изменения, внесенные в документ.|
||[sections](/javascript/api/word/word.document#sections)|Получает коллекцию объектов Section в документе.|
||[save()](/javascript/api/word/word.document#save--)|Сохраняет документ.|
|[Font](/javascript/api/word/word.font)|[bold](/javascript/api/word/word.font#bold)|Возвращает или задает значение, указывающее, является ли шрифт полужирным.|
||[color](/javascript/api/word/word.font#color)|Возвращает или задает цвет для указанного шрифта.|
||[doubleStrikeThrough](/javascript/api/word/word.font#doublestrikethrough)|Получает или задает значение, указывающее, имеет ли шрифт двойное зачеркивание.|
||[хигхлигхтколор](/javascript/api/word/word.font#highlightcolor)|Получает или задает цвет выделения.|
||[italic](/javascript/api/word/word.font#italic)|Возвращает или задает значение, указывающее, является ли шрифт курсивным.|
||[name](/javascript/api/word/word.font#name)|Получает или задает значение, представляющее имя шрифта.|
||[size](/javascript/api/word/word.font#size)|Получает или задает значение, представляющее размер шрифта в пунктах.|
||[strikeThrough](/javascript/api/word/word.font#strikethrough)|Получает или задает значение, указывающее, имеет ли шрифт зачеркивание.|
||[subscript](/javascript/api/word/word.font#subscript)|Возвращает или задает значение, указывающее, является ли шрифт подстрочным.|
||[superscript](/javascript/api/word/word.font#superscript)|Возвращает или задает значение, указывающее, является ли шрифт надстрочным.|
||[underline](/javascript/api/word/word.font#underline)|Возвращает или задает значение, указывающее тип подчеркивания шрифта.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[altTextDescription](/javascript/api/word/word.inlinepicture#alttextdescription)|Получает или задает строку, представляющую замещающий текст, связанный с встроенным изображением.|
||[altTextTitle](/javascript/api/word/word.inlinepicture#alttexttitle)|Возвращает или задает строку, содержащую заголовок встроенного рисунка.|
||[getBase64ImageSrc()](/javascript/api/word/word.inlinepicture#getbase64imagesrc--)|Возвращает строковое представление встроенного рисунка в кодировке base64.|
||[height](/javascript/api/word/word.inlinepicture#height)|Возвращает или задает число, которое описывает высоту встроенного рисунка.|
||[hyperlink](/javascript/api/word/word.inlinepicture#hyperlink)|Получает или задает гиперссылку на изображении.|
||[insertContentControl()](/javascript/api/word/word.inlinepicture#insertcontentcontrol--)|Включает встроенный рисунок в элемент управления содержимым форматированного текста.|
||[lockAspectRatio](/javascript/api/word/word.inlinepicture#lockaspectratio)|Возвращает или задает значение, указывающее, сохраняет ли встроенный рисунок исходные пропорции при изменении размера.|
||[parentContentControl](/javascript/api/word/word.inlinepicture#parentcontentcontrol)|Возвращает элемент управления содержимым, который содержит встроенный рисунок.|
||[width](/javascript/api/word/word.inlinepicture#width)|Возвращает или задает число, которое описывает ширину встроенного рисунка.|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[items](/javascript/api/word/word.inlinepicturecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Paragraph](/javascript/api/word/word.paragraph)|[ориентации](/javascript/api/word/word.paragraph#alignment)|Возвращает или задает выравнивание для абзаца.|
||[clear()](/javascript/api/word/word.paragraph#clear--)|Очищает содержимое объекта абзаца.|
||[delete()](/javascript/api/word/word.paragraph#delete--)|Удаляет абзац и его содержимое из документа.|
||[фирстлинеиндент](/javascript/api/word/word.paragraph#firstlineindent)|Возвращает или задает значение отступа первой строки или выступа в пунктах.|
||[getHtml()](/javascript/api/word/word.paragraph#gethtml--)|Возвращает HTML-представление объекта абзаца.|
||[getOoxml()](/javascript/api/word/word.paragraph#getooxml--)|Возвращает OOXML-представление объекта абзаца.|
||[игнорепункт](/javascript/api/word/word.paragraph#ignorepunct)||
||[игнореспаце](/javascript/api/word/word.paragraph#ignorespace)||
||[insertBreak (breakType: Word. BreakType, insertLocation: Word. InsertLocation)](/javascript/api/word/word.paragraph#insertbreak-breaktype--insertlocation-)|Вставляет разрыв в указанном расположении в основном документе.|
||[insertContentControl()](/javascript/api/word/word.paragraph#insertcontentcontrol--)|Включает объект абзаца в элемент управления содержимым форматированного текста.|
||[insertFileFromBase64 (base64File: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.paragraph#insertfilefrombase64-base64file--insertlocation-)|Вставляет документ в абзац в указанном расположении.|
||[insertHtml (HTML: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.paragraph#inserthtml-html--insertlocation-)|Вставляет HTML в абзац в указанном расположении.|
||[insertInlinePictureFromBase64 (base64EncodedImage: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.paragraph#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|Вставляет рисунок в абзац в указанном расположении.|
||[insertOoxml (OOXML: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.paragraph#insertooxml-ooxml--insertlocation-)|Вставляет OOXML в абзац в указанном расположении.|
||[insertParagraph (paragraphText: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.paragraph#insertparagraph-paragraphtext--insertlocation-)|Вставляет абзац в указанном расположении.|
||[insertText (Text: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.paragraph#inserttext-text--insertlocation-)|Вставляет текст в абзац в указанном расположении.|
||[лефтиндент](/javascript/api/word/word.paragraph#leftindent)|Возвращает или задает значение отступа слева для абзаца (в пунктах).|
||[линеспаЦинг](/javascript/api/word/word.paragraph#linespacing)|Возвращает или задает междустрочный интервал для указанного абзаца (в пунктах).|
||[линеунитафтер](/javascript/api/word/word.paragraph#lineunitafter)|Возвращает или задает расстояние от абзаца до абзаца (в линиях сетки).|
||[линеунитбефоре](/javascript/api/word/word.paragraph#lineunitbefore)|Возвращает или устанавливает междустрочный интервал до абзаца (в линиях сетки).|
||[matchCase](/javascript/api/word/word.paragraph#matchcase)||
||[матчпрефикс](/javascript/api/word/word.paragraph#matchprefix)||
||[матчсуффикс](/javascript/api/word/word.paragraph#matchsuffix)||
||[матчвхолеворд](/javascript/api/word/word.paragraph#matchwholeword)||
||[матчвилдкардс](/javascript/api/word/word.paragraph#matchwildcards)||
||[аутлинелевел](/javascript/api/word/word.paragraph#outlinelevel)|Возвращает или задает уровень структуры абзаца.|
||[contentControls](/javascript/api/word/word.paragraph#contentcontrols)|Возвращает коллекцию объектов элементов управления содержимым в абзаце.|
||[font](/javascript/api/word/word.paragraph#font)|Возвращает формат текста абзаца.|
||[inlinePictures](/javascript/api/word/word.paragraph#inlinepictures)|Получает коллекцию объектов коллекцию inlinepicture в абзаце.|
||[parentContentControl](/javascript/api/word/word.paragraph#parentcontentcontrol)|Возвращает элемент управления содержимым, содержащий абзац.|
||[text](/javascript/api/word/word.paragraph#text)|Возвращает текст абзаца.|
||[ригхтиндент](/javascript/api/word/word.paragraph#rightindent)|Возвращает или задает значение отступа справа для абзаца (в пунктах).|
||[Search (searchText: строка, searchOptions?: Word. SearchOptions \| {игнорепункт?: Boolean игнореспаце?: Boolean matchCase?: Boolean матчпрефикс?: Boolean матчсуффикс?: Boolean матчвхолеворд?: Boolean матчвилдкардс?: Boolean})](/javascript/api/word/word.paragraph#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Выполняет поиск с указанным SearchOptions в области объекта абзаца.|
||[SELECT (selectionMode?: Word. SelectionMode)](/javascript/api/word/word.paragraph#select-selectionmode-)|Выбирает абзац и переходит к нему в пользовательском интерфейсе Word.|
||[спацеафтер](/javascript/api/word/word.paragraph#spaceafter)|Возвращает или задает междустрочный интервал после абзаца (в пунктах).|
||[SpaceBefore присваивается](/javascript/api/word/word.paragraph#spacebefore)|Возвращает или задает междустрочный интервал до абзаца (в пунктах).|
||[style](/javascript/api/word/word.paragraph#style)|Получает или задает имя стиля для абзаца.|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[items](/javascript/api/word/word.paragraphcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Range](/javascript/api/word/word.range)|[clear()](/javascript/api/word/word.range#clear--)|Очищает содержимое объекта диапазона.|
||[delete()](/javascript/api/word/word.range#delete--)|Удаляет диапазон и его содержимое из документа.|
||[getHtml()](/javascript/api/word/word.range#gethtml--)|Возвращает HTML-представление объекта Range.|
||[getOoxml()](/javascript/api/word/word.range#getooxml--)|Возвращает OOXML-представление объекта диапазона.|
||[игнорепункт](/javascript/api/word/word.range#ignorepunct)||
||[игнореспаце](/javascript/api/word/word.range#ignorespace)||
||[insertBreak (breakType: Word. BreakType, insertLocation: Word. InsertLocation)](/javascript/api/word/word.range#insertbreak-breaktype--insertlocation-)|Вставляет разрыв в указанном расположении в основном документе.|
||[insertContentControl()](/javascript/api/word/word.range#insertcontentcontrol--)|Включает объект диапазона в элемент управления содержимым форматированного текста.|
||[insertFileFromBase64 (base64File: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.range#insertfilefrombase64-base64file--insertlocation-)|Вставляет документ в указанном расположении.|
||[insertHtml (HTML: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.range#inserthtml-html--insertlocation-)|Вставляет HTML-код в указанном расположении.|
||[insertOoxml (OOXML: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.range#insertooxml-ooxml--insertlocation-)|Вставляет OOXML-код в указанном расположении.|
||[insertParagraph (paragraphText: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.range#insertparagraph-paragraphtext--insertlocation-)|Вставляет абзац в указанном расположении.|
||[insertText (Text: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.range#inserttext-text--insertlocation-)|Вставляет текст в заданном расположении.|
||[matchCase](/javascript/api/word/word.range#matchcase)||
||[матчпрефикс](/javascript/api/word/word.range#matchprefix)||
||[матчсуффикс](/javascript/api/word/word.range#matchsuffix)||
||[матчвхолеворд](/javascript/api/word/word.range#matchwholeword)||
||[матчвилдкардс](/javascript/api/word/word.range#matchwildcards)||
||[contentControls](/javascript/api/word/word.range#contentcontrols)|Получает коллекцию объектов элементов управления содержимым в диапазоне.|
||[font](/javascript/api/word/word.range#font)|Возвращает формат текста диапазона.|
||[paragraphs](/javascript/api/word/word.range#paragraphs)|Возвращает коллекцию объектов абзаца в диапазоне.|
||[parentContentControl](/javascript/api/word/word.range#parentcontentcontrol)|Возвращает элемент управления содержимым, содержащий диапазон.|
||[text](/javascript/api/word/word.range#text)|Возвращает текст диапазона.|
||[Search (searchText: строка, searchOptions?: Word. SearchOptions \| {игнорепункт?: Boolean игнореспаце?: Boolean matchCase?: Boolean матчпрефикс?: Boolean матчсуффикс?: Boolean матчвхолеворд?: Boolean матчвилдкардс?: Boolean})](/javascript/api/word/word.range#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Выполняет поиск с указанным SearchOptions в области объекта Range.|
||[SELECT (selectionMode?: Word. SelectionMode)](/javascript/api/word/word.range#select-selectionmode-)|Выбор диапазона и переход к нему в пользовательском интерфейсе Word.|
||[style](/javascript/api/word/word.range#style)|Получает или задает имя стиля для диапазона.|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[items](/javascript/api/word/word.rangecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[SearchOptions](/javascript/api/word/word.searchoptions)|[игнорепункт](/javascript/api/word/word.searchoptions#ignorepunct)|Возвращает или задает значение, которое указывает, следует ли пропустить все знаки препинания между словами.|
||[игнореспаце](/javascript/api/word/word.searchoptions#ignorespace)|Получает или задает значение, указывающее, следует ли игнорировать все пробелы между словами.|
||[matchCase](/javascript/api/word/word.searchoptions#matchcase)|Возвращает или задает значение, которое указывает, следует ли выполнять поиск с учетом регистра.|
||[матчпрефикс](/javascript/api/word/word.searchoptions#matchprefix)|Возвращает или задает значение, которое указывает, нужно ли учитывать слова, начинающиеся со строки поиска.|
||[матчсуффикс](/javascript/api/word/word.searchoptions#matchsuffix)|Возвращает или задает значение, указывающее, нужно ли учитывать слова, которые заканчиваются строкой поиска.|
||[матчвхолеворд](/javascript/api/word/word.searchoptions#matchwholeword)|Возвращает или задает значение, которое указывает, следует ли искать только целые слова, а не текст, являющийся частью большего слова.|
||[матчвилдкардс](/javascript/api/word/word.searchoptions#matchwildcards)|Возвращает или задает значение, которое указывает, будет ли выполняться поиск с использованием специальных операторов поиска.|
|[Section](/javascript/api/word/word.section)|[Footer (тип: Word. Хеадерфутертипе)](/javascript/api/word/word.section#getfooter-type-)|Возвращает один из нижних колонтитулов раздела.|
||[коголовка (тип: Word. Хеадерфутертипе)](/javascript/api/word/word.section#getheader-type-)|Возвращает один из верхних колонтитулов раздела.|
||[body](/javascript/api/word/word.section#body)|Возвращает объект Body раздела.|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[items](/javascript/api/word/word.sectioncollection#items)|Получает загруженные дочерние элементы в этой коллекции.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Word](/javascript/api/word)
- [Наборы обязательных элементов API JavaScript для Word](word-api-requirement-sets.md)
