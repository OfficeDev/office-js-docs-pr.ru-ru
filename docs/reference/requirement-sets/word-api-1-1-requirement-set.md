---
title: Набор обязательных элементов API JavaScript для Word 1,1
description: Сведения о наборе требований WordApi 1,1
ms.date: 07/25/2019
ms.prod: word
localization_priority: Normal
ms.openlocfilehash: ad1ec8b226bc958ed1be6e233a070108612661ad
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/11/2020
ms.locfileid: "47431509"
---
# <a name="whats-new-in-word-javascript-api-11"></a>Новые возможности API JavaScript для Word 1,1

WordApi 1,1 является первым набором требований API JavaScript для Word. Это единственный набор требований Word API, поддерживаемый Word 2016.

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в наборе обязательных элементов API JavaScript для Word 1,1. Чтобы просмотреть справочную документацию по API для всех API, поддерживаемых в наборе обязательных элементов API JavaScript для Word 1,1, обратитесь к разделам [API Word в наборе требований 1,1](/javascript/api/word?view=word-js-1.1&preserve-view=true).

| Класс | Поля | Описание |
|:---|:---|:---|
|[Основной текст](/javascript/api/word/word.body)|[clear()](/javascript/api/word/word.body#clear--)|Очищает объект содержимого. Пользователь может отменить операцию очищения для содержимого.|
||[getHtml()](/javascript/api/word/word.body#gethtml--)|Возвращает HTML-представление объекта Body. При отображении на веб-странице или в средстве просмотра HTML форматирование будет близким, но не точным, соответствующим формату документа. Этот метод не возвращает точно такой же HTML-код для одного и того же документа на различных платформах (Windows, Mac и т. д.). Если вам нужна точная точность или согласованность на различных платформах, используйте `Body.getOoxml()` и преобразуйте возвращенный XML в HTML.|
||[getOoxml()](/javascript/api/word/word.body#getooxml--)|Возвращает OOXML-представление (Office Open XML) объекта содержимого.|
||[insertBreak (breakType: Word. BreakType, insertLocation: Word. InsertLocation)](/javascript/api/word/word.body#insertbreak-breaktype--insertlocation-)|Вставляет разрыв в указанном расположении в основном документе. Возможные значения InsertLocation: Start или End.|
||[insertContentControl()](/javascript/api/word/word.body#insertcontentcontrol--)|Включает объект содержимого в элемент управления форматированным текстом.|
||[insertFileFromBase64 (base64File: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.body#insertfilefrombase64-base64file--insertlocation-)|Вставляет документ в содержимое в заданном расположении. Возможные значения insertLocation: Replace, Start и End.|
||[insertHtml (HTML: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.body#inserthtml-html--insertlocation-)|Вставляет HTML-код в указанном расположении. Возможные значения insertLocation: Replace, Start и End.|
||[insertOoxml (OOXML: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.body#insertooxml-ooxml--insertlocation-)|Вставляет OOXML-код в указанном расположении.  Возможные значения insertLocation: Replace, Start и End.|
||[insertParagraph (paragraphText: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.body#insertparagraph-paragraphtext--insertlocation-)|Вставляет абзац в указанном расположении. Возможные значения insertLocation: Start и End.|
||[insertText (Text: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.body#inserttext-text--insertlocation-)|Вставляет текст в содержимое в заданном расположении. Возможные значения insertLocation: Replace, Start и End.|
||[contentControls](/javascript/api/word/word.body#contentcontrols)|Возвращает коллекцию объектов элемента управления содержимым "форматированный текст" в тексте. Только для чтения.|
||[font](/javascript/api/word/word.body#font)|Получает формат текста, указанный для содержимого документа или раздела. Используйте этот параметр для получения и задания имени шрифта, размера, цвета и других свойств. Только для чтения.|
||[inlinePictures](/javascript/api/word/word.body#inlinepictures)|Получает коллекцию объектов коллекцию inlinepicture в тексте. Коллекция не содержит плавающие рисунки. Только для чтения.|
||[paragraphs](/javascript/api/word/word.body#paragraphs)|Возвращает коллекцию объектов абзаца в тексте. Только для чтения.|
||[parentContentControl](/javascript/api/word/word.body#parentcontentcontrol)|Получает элемент управления содержимым, содержащий документ или раздел. Вызывается, если родительский элемент управления содержимым отсутствует. Только для чтения.|
||[text](/javascript/api/word/word.body#text)|Возвращает текст содержимого. Для вставки текста используется метод insertText. Только для чтения.|
||[Search (searchText: строка, searchOptions?: Word. SearchOptions)](/javascript/api/word/word.body#search-searchtext--searchoptions-)|Выполняет поиск с указанным SearchOptions в области объекта Body. Результат поиска — это коллекция объектов диапазона.|
||[SELECT (selectionMode?: Word. SelectionMode)](/javascript/api/word/word.body#select-selectionmode-)|Выбирает содержимое и переходит к нему в пользовательском интерфейсе Word.|
||[style](/javascript/api/word/word.body#style)|Возвращает или задает имя стиля для основного текста. Используйте это свойство для пользовательских стилей и локализованных имен стилей. Чтобы использовать встроенные стили, поддерживающие несколько языковых стандартов, применяйте свойство styleBuiltIn.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[виде](/javascript/api/word/word.contentcontrol#appearance)|Получает или задает внешний вид элемента управления содержимым. Возможные значения: "BoundingBox", "Tags" или "Hidden".|
||[каннотделете](/javascript/api/word/word.contentcontrol#cannotdelete)|Возвращает или задает значение, указывающее, может ли пользователь удалить элемент управления содержимым. Является взаимоисключающим со свойством removeWhenEdited.|
||[каннотедит](/javascript/api/word/word.contentcontrol#cannotedit)|Возвращает или задает значение, указывающее, может ли пользователь изменять содержимое элемента управления содержимым.|
||[clear()](/javascript/api/word/word.contentcontrol#clear--)|Очищает содержимое элемента управления содержимым. Пользователь может отменить операцию для очищенного содержимого.|
||[color](/javascript/api/word/word.contentcontrol#color)|Возвращает или задает цвет элемента управления содержимым. Цвет задается в формате "#RRGGBB" или с помощью имени цвета.|
||[Delete (Кипконтент: Boolean)](/javascript/api/word/word.contentcontrol#delete-keepcontent-)|Удаляет элемент управления содержимым и его содержимое. Если параметру keepContent присвоено значение true, содержимое не удаляется.|
||[getHtml()](/javascript/api/word/word.contentcontrol#gethtml--)|Возвращает HTML-представление объекта элемента управления содержимым. При отображении на веб-странице или в средстве просмотра HTML форматирование будет близким, но не точным, соответствующим формату документа. Этот метод не возвращает точно такой же HTML-код для одного и того же документа на различных платформах (Windows, Mac и т. д.). Если вам нужна точная точность или согласованность на различных платформах, используйте `ContentControl.getOoxml()` и преобразуйте возвращенный XML в HTML.|
||[getOoxml()](/javascript/api/word/word.contentcontrol#getooxml--)|Возвращает OOXML-представление объекта элемента управления содержимым.|
||[insertBreak (breakType: Word. BreakType, insertLocation: Word. InsertLocation)](/javascript/api/word/word.contentcontrol#insertbreak-breaktype--insertlocation-)|Вставляет разрыв в указанном расположении в основном документе. Значение insertLocation может быть "Start", "End", "Before" или "After". Этот метод не может использоваться с элементами управления содержимым "Ричтексттабле", "Ричтексттаблеров" и "Ричтексттаблецелл".|
||[insertFileFromBase64 (base64File: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.contentcontrol#insertfilefrombase64-base64file--insertlocation-)|Вставляет документ в элемент управления содержимым в указанном расположении. Возможные значения insertLocation: Replace, Start и End.|
||[insertHtml (HTML: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.contentcontrol#inserthtml-html--insertlocation-)|Вставляет HTML-код в элемент управления содержимым в указанном расположении. Возможные значения insertLocation: Replace, Start и End.|
||[insertOoxml (OOXML: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.contentcontrol#insertooxml-ooxml--insertlocation-)|Вставляет OOXML в элемент управления содержимым в указанном расположении.  Возможные значения insertLocation: Replace, Start и End.|
||[insertParagraph (paragraphText: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.contentcontrol#insertparagraph-paragraphtext--insertlocation-)|Вставляет абзац в указанном расположении. Значение insertLocation может быть "Start", "End", "Before" или "After".|
||[insertText (Text: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.contentcontrol#inserttext-text--insertlocation-)|Вставляет текст в элемент управления содержимым в указанном расположении. Возможные значения insertLocation: Replace, Start и End.|
||[плацехолдертекст](/javascript/api/word/word.contentcontrol#placeholdertext)|Возвращает или задает замещающий текст элемента управления содержимым. Если элемент управления содержимым пуст, отображается затемненный текст.|
||[contentControls](/javascript/api/word/word.contentcontrol#contentcontrols)|Получает коллекцию объектов элементов управления содержимым в элементе управления содержимым. Только для чтения.|
||[font](/javascript/api/word/word.contentcontrol#font)|Получает текстовый формат элемента управления содержимым. Используйте это свойство для получения и установки имени, размера, цвета и других свойств шрифта. Только для чтения.|
||[id](/javascript/api/word/word.contentcontrol#id)|Возвращает целое число, представляющее собой идентификатор элемента управления контентом. Только для чтения.|
||[inlinePictures](/javascript/api/word/word.contentcontrol#inlinepictures)|Получает коллекцию объектов inlinePicture в элементе управления содержимым. Коллекция не содержит плавающие рисунки. Только для чтения.|
||[paragraphs](/javascript/api/word/word.contentcontrol#paragraphs)|Получает коллекцию объектов абзацев в элементе управления содержимым. Только для чтения.|
||[parentContentControl](/javascript/api/word/word.contentcontrol#parentcontentcontrol)|Получает элемент управления содержимым, содержащий элемент управления содержимым. Вызывается, если родительский элемент управления содержимым отсутствует. Только для чтения.|
||[text](/javascript/api/word/word.contentcontrol#text)|Возвращает текст элемента управления содержимым. Только для чтения.|
||[type](/javascript/api/word/word.contentcontrol#type)|Получает тип элемента управления содержимым. На данный момент поддерживаются только элементы управления содержимым в формате RTF. Только для чтения.|
||[ремовевхенедитед](/javascript/api/word/word.contentcontrol#removewhenedited)|Возвращает или задает значение, указывающее, удаляется ли элемент управления содержимым после изменения. Является взаимоисключающим со свойством cannotDelete.|
||[Search (searchText: строка, searchOptions?: Word. SearchOptions)](/javascript/api/word/word.contentcontrol#search-searchtext--searchoptions-)|Выполняет поиск с указанным SearchOptions в области объекта элемента управления содержимым. Результат поиска — это коллекция объектов диапазона.|
||[SELECT (selectionMode?: Word. SelectionMode)](/javascript/api/word/word.contentcontrol#select-selectionmode-)|Выбирает элемент управления контентом. При этом Word переходит к выделенному фрагменту.|
||[style](/javascript/api/word/word.contentcontrol#style)|Получает или задает имя стиля для элемента управления содержимым. Используйте это свойство для пользовательских стилей и локализованных имен стилей. Чтобы использовать встроенные стили, поддерживающие несколько языковых стандартов, применяйте свойство styleBuiltIn.|
||[мечать](/javascript/api/word/word.contentcontrol#tag)|Возвращает или задает тег для определения элемента управления содержимым.|
||[заголовок](/javascript/api/word/word.contentcontrol#title)|Получает или задает заголовок для элемента управления содержимым.|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getById(id: number)](/javascript/api/word/word.contentcontrolcollection#getbyid-id-)|Возвращает элемент управления содержимым по его идентификатору. Вызывается, если в данной коллекции нет элемента управления контентом с идентификатором.|
||[getByTag(tag: string)](/javascript/api/word/word.contentcontrolcollection#getbytag-tag-)|Возвращает элементы управления содержимым с указанным тегом.|
||[getByTitle(title: string)](/javascript/api/word/word.contentcontrolcollection#getbytitle-title-)|Возвращает элементы управления контентом с указанным заголовком.|
||[getItem(index: number)](/javascript/api/word/word.contentcontrolcollection#getitem-index-)|Возвращает элемент управления контентом по его индексу в коллекции.|
||[items](/javascript/api/word/word.contentcontrolcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Document](/javascript/api/word/word.document)|[При выборе ()](/javascript/api/word/word.document#getselection--)|Возвращает текущий выбранный фрагмент документа. Получение нескольких выбранных фрагментов не поддерживается.|
||[body](/javascript/api/word/word.document#body)|Возвращает объект Body документа. Текст — это текст, который исключает заголовки, нижние колонтитулы, сноски, текстовые поля и т. д. Только для чтения.|
||[contentControls](/javascript/api/word/word.document#contentcontrols)|Возвращает коллекцию объектов элементов управления содержимым в документе. Сюда входят элементы управления содержимым в тексте документа, верхних и нижних колонтитулов, текстовых полях и т. д. Только для чтения.|
||[сохраняем](/javascript/api/word/word.document#saved)|Указывает, сохранены ли изменения, внесенные в документ. Значение true указывает на то, что с момента последнего сохранения в документ не вносились изменения. Только для чтения.|
||[sections](/javascript/api/word/word.document#sections)|Получает коллекцию объектов Section в документе. Только для чтения.|
||[save()](/javascript/api/word/word.document#save--)|Сохраняет документ. При этом используется соглашение об именовании файлов Word по умолчанию, если документ ранее не сохранялся.|
|[Font](/javascript/api/word/word.font)|[bold](/javascript/api/word/word.font#bold)|Возвращает или задает значение, указывающее, является ли шрифт полужирным. Задайте значение true, чтобы отформатировать шрифт как полужирный, в противном случае — задайте значение false.|
||[color](/javascript/api/word/word.font#color)|Возвращает или задает цвет для указанного шрифта. Можно указать значение в формате "#RRGGBB" или имя цвета.|
||[doubleStrikeThrough](/javascript/api/word/word.font#doublestrikethrough)|Получает или задает значение, указывающее, имеет ли шрифт двойное зачеркивание. Задайте значение true, чтобы использовать двойное зачеркивание, в противном случае задайте значение false.|
||[хигхлигхтколор](/javascript/api/word/word.font#highlightcolor)|Получает или задает цвет выделения. Чтобы задать его, используйте значение либо в формате "#RRGGBB", либо в имени цвета. Чтобы удалить цвет выделения, задайте для него значение null. Возвращаемый цвет выделения может быть указан в формате "#RRGGBB", пустой строке для смешанных цветов выделения или NULL без цвета выделения.|
||[italic](/javascript/api/word/word.font#italic)|Возвращает или задает значение, указывающее, является ли шрифт курсивным. Задайте значение true, если шрифт является курсивом, в противном случае — задайте значение false.|
||[name](/javascript/api/word/word.font#name)|Получает или задает значение, представляющее имя шрифта.|
||[size](/javascript/api/word/word.font#size)|Получает или задает значение, представляющее размер шрифта в пунктах.|
||[strikeThrough](/javascript/api/word/word.font#strikethrough)|Получает или задает значение, указывающее, имеет ли шрифт зачеркивание. Задайте значение true, если зачеркивание используется, в противном случае — задайте значение false.|
||[subscript](/javascript/api/word/word.font#subscript)|Возвращает или задает значение, указывающее, является ли шрифт подстрочным. Задайте значение true, если шрифт является подстрочным, в противном случае — задайте значение false.|
||[superscript](/javascript/api/word/word.font#superscript)|Возвращает или задает значение, указывающее, является ли шрифт надстрочным. Задайте значение true, если шрифт является надстрочным, в противном случае — задайте значение false.|
||[underline](/javascript/api/word/word.font#underline)|Возвращает или задает значение, указывающее тип подчеркивания шрифта. "None", если шрифт не является подчеркиванием.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[altTextDescription](/javascript/api/word/word.inlinepicture#alttextdescription)|Получает или задает строку, представляющую замещающий текст, связанный с встроенным изображением.|
||[altTextTitle](/javascript/api/word/word.inlinepicture#alttexttitle)|Возвращает или задает строку, содержащую заголовок встроенного рисунка.|
||[getBase64ImageSrc()](/javascript/api/word/word.inlinepicture#getbase64imagesrc--)|Возвращает строковое представление встроенного рисунка в кодировке base64.|
||[height](/javascript/api/word/word.inlinepicture#height)|Возвращает или задает число, которое описывает высоту встроенного рисунка.|
||[hyperlink](/javascript/api/word/word.inlinepicture#hyperlink)|Получает или задает гиперссылку на изображении. Используйте ' # ', чтобы отделить адрес от части необязательного расположения.|
||[insertContentControl()](/javascript/api/word/word.inlinepicture#insertcontentcontrol--)|Включает встроенный рисунок в элемент управления содержимым форматированного текста.|
||[lockAspectRatio](/javascript/api/word/word.inlinepicture#lockaspectratio)|Возвращает или задает значение, указывающее, сохраняет ли встроенный рисунок исходные пропорции при изменении размера.|
||[parentContentControl](/javascript/api/word/word.inlinepicture#parentcontentcontrol)|Возвращает элемент управления содержимым, который содержит встроенный рисунок. Вызывается, если родительский элемент управления содержимым отсутствует. Только для чтения.|
||[width](/javascript/api/word/word.inlinepicture#width)|Возвращает или задает число, которое описывает ширину встроенного рисунка.|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[items](/javascript/api/word/word.inlinepicturecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Paragraph](/javascript/api/word/word.paragraph)|[ориентации](/javascript/api/word/word.paragraph#alignment)|Возвращает или задает выравнивание для абзаца. Возможные значения: left, centered, right и justified.|
||[clear()](/javascript/api/word/word.paragraph#clear--)|Очищает содержимое объекта абзаца. Пользователь может отменить операцию для очищенного содержимого.|
||[delete()](/javascript/api/word/word.paragraph#delete--)|Удаляет абзац и его содержимое из документа.|
||[фирстлинеиндент](/javascript/api/word/word.paragraph#firstlineindent)|Возвращает или задает значение отступа первой строки или выступа в пунктах. Для установки отступа первой строки укажите положительное значение и используйте отрицательное значение, чтобы задать выступ.|
||[getHtml()](/javascript/api/word/word.paragraph#gethtml--)|Возвращает HTML-представление объекта абзаца. При отображении на веб-странице или в средстве просмотра HTML форматирование будет близким, но не точным, соответствующим формату документа. Этот метод не возвращает точно такой же HTML-код для одного и того же документа на различных платформах (Windows, Mac и т. д.). Если вам нужна точная точность или согласованность на различных платформах, используйте `Paragraph.getOoxml()` и преобразуйте возвращенный XML в HTML.|
||[getOoxml()](/javascript/api/word/word.paragraph#getooxml--)|Возвращает OOXML-представление объекта абзаца.|
||[insertBreak (breakType: Word. BreakType, insertLocation: Word. InsertLocation)](/javascript/api/word/word.paragraph#insertbreak-breaktype--insertlocation-)|Вставляет разрыв в указанном расположении в основном документе. Возможные значения insertLocation: Before и After.|
||[insertContentControl()](/javascript/api/word/word.paragraph#insertcontentcontrol--)|Включает объект абзаца в элемент управления содержимым форматированного текста.|
||[insertFileFromBase64 (base64File: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.paragraph#insertfilefrombase64-base64file--insertlocation-)|Вставляет документ в абзац в указанном расположении. Возможные значения insertLocation: Replace, Start и End.|
||[insertHtml (HTML: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.paragraph#inserthtml-html--insertlocation-)|Вставляет HTML в абзац в указанном расположении. Возможные значения insertLocation: Replace, Start и End.|
||[insertInlinePictureFromBase64 (base64EncodedImage: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.paragraph#insertinlinepicturefrombase64-base64encodedimage--insertlocation-)|Вставляет рисунок в абзац в указанном расположении. Возможные значения insertLocation: Replace, Start и End.|
||[insertOoxml (OOXML: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.paragraph#insertooxml-ooxml--insertlocation-)|Вставляет OOXML в абзац в указанном расположении. Возможные значения insertLocation: Replace, Start и End.|
||[insertParagraph (paragraphText: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.paragraph#insertparagraph-paragraphtext--insertlocation-)|Вставляет абзац в указанном расположении. Возможные значения InsertLocation: Before и After.|
||[insertText (Text: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.paragraph#inserttext-text--insertlocation-)|Вставляет текст в абзац в указанном расположении. Возможные значения insertLocation: Replace, Start и End.|
||[лефтиндент](/javascript/api/word/word.paragraph#leftindent)|Возвращает или задает значение отступа слева для абзаца (в пунктах).|
||[линеспаЦинг](/javascript/api/word/word.paragraph#linespacing)|Возвращает или задает междустрочный интервал для указанного абзаца (в пунктах). В пользовательском интерфейсе Word это значение делится на 12.|
||[линеунитафтер](/javascript/api/word/word.paragraph#lineunitafter)|Возвращает или задает расстояние от абзаца до абзаца (в линиях сетки).|
||[линеунитбефоре](/javascript/api/word/word.paragraph#lineunitbefore)|Возвращает или устанавливает междустрочный интервал до абзаца (в линиях сетки).|
||[аутлинелевел](/javascript/api/word/word.paragraph#outlinelevel)|Возвращает или задает уровень структуры абзаца.|
||[contentControls](/javascript/api/word/word.paragraph#contentcontrols)|Возвращает коллекцию объектов элементов управления содержимым в абзаце. Только для чтения.|
||[font](/javascript/api/word/word.paragraph#font)|Возвращает формат текста абзаца. Используйте это свойство для получения и задания имени, размера, цвета и других свойств шрифта. Только для чтения.|
||[inlinePictures](/javascript/api/word/word.paragraph#inlinepictures)|Получает коллекцию объектов коллекцию inlinepicture в абзаце. Коллекция не содержит плавающие рисунки. Только для чтения.|
||[parentContentControl](/javascript/api/word/word.paragraph#parentcontentcontrol)|Возвращает элемент управления содержимым, содержащий абзац. Вызывается, если родительский элемент управления содержимым отсутствует. Только для чтения.|
||[text](/javascript/api/word/word.paragraph#text)|Возвращает текст абзаца. Только для чтения.|
||[ригхтиндент](/javascript/api/word/word.paragraph#rightindent)|Возвращает или задает значение отступа справа для абзаца (в пунктах).|
||[Search (searchText: строка, searchOptions?: Word. SearchOptions})](/javascript/api/word/word.paragraph#search-searchtext--searchoptions-)|Выполняет поиск с указанным SearchOptions в области объекта абзаца. Результат поиска — это коллекция объектов диапазона.|
||[SELECT (selectionMode?: Word. SelectionMode)](/javascript/api/word/word.paragraph#select-selectionmode-)|Выбирает абзац и переходит к нему в пользовательском интерфейсе Word.|
||[спацеафтер](/javascript/api/word/word.paragraph#spaceafter)|Возвращает или задает междустрочный интервал после абзаца (в пунктах).|
||[SpaceBefore присваивается](/javascript/api/word/word.paragraph#spacebefore)|Возвращает или задает междустрочный интервал до абзаца (в пунктах).|
||[style](/javascript/api/word/word.paragraph#style)|Получает или задает имя стиля для абзаца. Используйте это свойство для пользовательских стилей и локализованных имен стилей. Чтобы использовать встроенные стили, поддерживающие несколько языковых стандартов, применяйте свойство styleBuiltIn.|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[items](/javascript/api/word/word.paragraphcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Range](/javascript/api/word/word.range)|[clear()](/javascript/api/word/word.range#clear--)|Очищает содержимое объекта диапазона. Пользователь может отменить операцию для очищенного содержимого.|
||[delete()](/javascript/api/word/word.range#delete--)|Удаляет диапазон и его содержимое из документа.|
||[getHtml()](/javascript/api/word/word.range#gethtml--)|Возвращает HTML-представление объекта Range. При отображении на веб-странице или в средстве просмотра HTML форматирование будет близким, но не точным, соответствующим формату документа. Этот метод не возвращает точно такой же HTML-код для одного и того же документа на различных платформах (Windows, Mac и т. д.). Если вам нужна точная точность или согласованность на различных платформах, используйте `Range.getOoxml()` и преобразуйте возвращенный XML в HTML.|
||[getOoxml()](/javascript/api/word/word.range#getooxml--)|Возвращает OOXML-представление объекта диапазона.|
||[insertBreak (breakType: Word. BreakType, insertLocation: Word. InsertLocation)](/javascript/api/word/word.range#insertbreak-breaktype--insertlocation-)|Вставляет разрыв в указанном расположении в основном документе. Возможные значения insertLocation: Before и After.|
||[insertContentControl()](/javascript/api/word/word.range#insertcontentcontrol--)|Включает объект диапазона в элемент управления содержимым форматированного текста.|
||[insertFileFromBase64 (base64File: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.range#insertfilefrombase64-base64file--insertlocation-)|Вставляет документ в указанном расположении. Значение insertLocation может быть "Replace", "Start", "End", "Before" или "After".|
||[insertHtml (HTML: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.range#inserthtml-html--insertlocation-)|Вставляет HTML-код в указанном расположении. Значение insertLocation может быть "Replace", "Start", "End", "Before" или "After".|
||[insertOoxml (OOXML: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.range#insertooxml-ooxml--insertlocation-)|Вставляет OOXML-код в указанном расположении.  Значение insertLocation может быть "Replace", "Start", "End", "Before" или "After".|
||[insertParagraph (paragraphText: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.range#insertparagraph-paragraphtext--insertlocation-)|Вставляет абзац в указанном расположении. Возможные значения InsertLocation: Before и After.|
||[insertText (Text: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.range#inserttext-text--insertlocation-)|Вставляет текст в заданном расположении. Значение insertLocation может быть "Replace", "Start", "End", "Before" или "After".|
||[contentControls](/javascript/api/word/word.range#contentcontrols)|Получает коллекцию объектов элементов управления содержимым в диапазоне. Только для чтения.|
||[font](/javascript/api/word/word.range#font)|Возвращает формат текста диапазона. Используйте это свойство, чтобы получать и задавать имея, размер, цвет и другие свойства шрифта. Только для чтения.|
||[paragraphs](/javascript/api/word/word.range#paragraphs)|Возвращает коллекцию объектов абзаца в диапазоне. Только для чтения.|
||[parentContentControl](/javascript/api/word/word.range#parentcontentcontrol)|Возвращает элемент управления содержимым, содержащий диапазон. Вызывается, если родительский элемент управления содержимым отсутствует. Только для чтения.|
||[text](/javascript/api/word/word.range#text)|Возвращает текст диапазона. Только для чтения.|
||[Search (searchText: строка, searchOptions?: Word. SearchOptions)](/javascript/api/word/word.range#search-searchtext--searchoptions-)|Выполняет поиск с указанным SearchOptions в области объекта Range. Результат поиска — это коллекция объектов диапазона.|
||[SELECT (selectionMode?: Word. SelectionMode)](/javascript/api/word/word.range#select-selectionmode-)|Выбор диапазона и переход к нему в пользовательском интерфейсе Word.|
||[style](/javascript/api/word/word.range#style)|Получает или задает имя стиля для диапазона. Используйте это свойство для пользовательских стилей и локализованных имен стилей. Чтобы использовать встроенные стили, поддерживающие несколько языковых стандартов, применяйте свойство styleBuiltIn.|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[items](/javascript/api/word/word.rangecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[SearchOptions](/javascript/api/word/word.searchoptions)|[игнорепункт](/javascript/api/word/word.searchoptions#ignorepunct)|Возвращает или задает значение, которое указывает, следует ли пропустить все знаки препинания между словами. Соответствует установленному флажку "Не учитывать знаки препинания" в диалоговом окне "Найти и заменить".|
||[игнореспаце](/javascript/api/word/word.searchoptions#ignorespace)|Получает или задает значение, указывающее, следует ли игнорировать все пробелы между словами. Соответствует флажку игнорировать пробелы в диалоговом окне "найти и заменить".|
||[matchCase](/javascript/api/word/word.searchoptions#matchcase)|Возвращает или задает значение, которое указывает, следует ли выполнять поиск с учетом регистра. Соответствует флажку "учитывать регистр" в диалоговом окне "найти и заменить".|
||[матчпрефикс](/javascript/api/word/word.searchoptions#matchprefix)|Возвращает или задает значение, которое указывает, нужно ли учитывать слова, начинающиеся со строки поиска. Соответствует установленному флажку "Учитывать префикс" в диалоговом окне "Найти и заменить".|
||[матчсуффикс](/javascript/api/word/word.searchoptions#matchsuffix)|Возвращает или задает значение, указывающее, нужно ли учитывать слова, которые заканчиваются строкой поиска. Соответствует установленному флажку "Учитывать суффикс" в диалоговом окне "Найти и заменить".|
||[матчвхолеворд](/javascript/api/word/word.searchoptions#matchwholeword)|Возвращает или задает значение, которое указывает, следует ли искать только целые слова, а не текст, являющийся частью большего слова. Соответствует установленному флажку "Только слово целиком" в диалоговом окне "Найти и заменить".|
||[матчвилдкардс](/javascript/api/word/word.searchoptions#matchwildcards)||
||[матчвилдкардс](/javascript/api/word/word.searchoptions#matchwildcards)|Возвращает или задает значение, которое указывает, будет ли выполняться поиск с использованием специальных операторов поиска. Соответствует установленному флажку "Подстановочные знаки" в диалоговом окне "Найти и заменить".|
|[Section](/javascript/api/word/word.section)|[Footer (тип: Word. Хеадерфутертипе)](/javascript/api/word/word.section#getfooter-type-)|Возвращает один из нижних колонтитулов раздела.|
||[коголовка (тип: Word. Хеадерфутертипе)](/javascript/api/word/word.section#getheader-type-)|Возвращает один из верхних колонтитулов раздела.|
||[body](/javascript/api/word/word.section#body)|Возвращает объект Body раздела. Сюда не входят метаданные верхнего и нижнего колонтитулов, а также другие метаданные разделов. Только для чтения.|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[items](/javascript/api/word/word.sectioncollection#items)|Получает загруженные дочерние элементы в этой коллекции.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Word](/javascript/api/word)
- [Наборы обязательных элементов API JavaScript для Word](word-api-requirement-sets.md)
