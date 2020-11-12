---
title: Набор обязательных элементов API JavaScript для Word 1,3
description: Сведения о наборе требований WordApi 1,3
ms.date: 11/09/2020
ms.prod: word
localization_priority: Normal
ms.openlocfilehash: 1344d66f2a4d9a3c9ff93c042fa1f23013e1bb27
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996432"
---
# <a name="whats-new-in-word-javascript-api-13"></a>Новые возможности API JavaScript для Word 1.3

WordApi 1,3 добавлена дополнительная поддержка элементов управления содержимым, настраиваемых XML и параметров на уровне документа.

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в наборе обязательных элементов API JavaScript для Word 1,3. Чтобы просмотреть справочную документацию по API для всех API, поддерживаемых в наборе обязательных элементов API JavaScript для Word 1,3 или более ранней версии, обратитесь к разделам [API Word в наборе требований 1,3](/javascript/api/word?view=word-js-1.3&preserve-view=true)

| Класс | Поля | Описание |
|:---|:---|:---|
|[Application](/javascript/api/word/word.application)|[Креатедокумент (base64File?: строка)](/javascript/api/word/word.application#createdocument-base64file-)|Создает новый документ, используя необязательный docx файл с кодировкой base64.|
|[Основной текст](/javascript/api/word/word.body)|[Тип-Range (rangeLocation?: Word. RangeLocation)](/javascript/api/word/word.body#getrange-rangelocation-)|Возвращает весь основной текст (либо его начальную или конечную точку) в виде диапазона.|
||[insertTable (rowCount: число, columnCount: число, insertLocation: Word. InsertLocation, Values?: строка [] [])](/javascript/api/word/word.body#inserttable-rowcount--columncount--insertlocation--values-)|Вставляет таблицу с указанным количеством строк и столбцов.|
||[lists](/javascript/api/word/word.body#lists)|Возвращает коллекцию объектов списков в основном тексте.|
||[parentBody](/javascript/api/word/word.body#parentbody)|Возвращает родительский текст основного текста.|
||[parentBodyOrNullObject](/javascript/api/word/word.body#parentbodyornullobject)|Возвращает родительский текст основного текста.|
||[parentContentControlOrNullObject](/javascript/api/word/word.body#parentcontentcontrolornullobject)|Получает элемент управления содержимым, содержащий документ или раздел.|
||[parentSection](/javascript/api/word/word.body#parentsection)|Возвращает родительский раздел основного текста.|
||[parentSectionOrNullObject](/javascript/api/word/word.body#parentsectionornullobject)|Возвращает родительский раздел основного текста.|
||[Table](/javascript/api/word/word.body#tables)|Возвращает коллекцию объектов таблиц в основном тексте.|
||[type](/javascript/api/word/word.body#type)|Возвращает тип основного текста.|
||[styleBuiltIn](/javascript/api/word/word.body#stylebuiltin)|Возвращает или задает имя встроенного стиля основного текста.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[Тип-Range (rangeLocation?: Word. RangeLocation)](/javascript/api/word/word.contentcontrol#getrange-rangelocation-)|Возвращает весь элемент управления содержимым (либо его начальную или конечную точку) в виде диапазона.|
||[getTextRanges (endingMarks: строка [], trimSpacing?: Boolean)](/javascript/api/word/word.contentcontrol#gettextranges-endingmarks--trimspacing-)|Получает диапазоны текста в элементе управления содержимым с помощью знаков препинания и/или других конечных меток.|
||[insertTable (rowCount: число, columnCount: число, insertLocation: Word. InsertLocation, Values?: строка [] [])](/javascript/api/word/word.contentcontrol#inserttable-rowcount--columncount--insertlocation--values-)|Вставляет таблицу с указанным количеством строк и столбцов в элемент управления содержимым или рядом с ним.|
||[lists](/javascript/api/word/word.contentcontrol#lists)|Возвращает коллекцию объектов списков в элементе управления содержимым.|
||[parentBody](/javascript/api/word/word.contentcontrol#parentbody)|Возвращает родительский текст элемента управления содержимым.|
||[parentContentControlOrNullObject](/javascript/api/word/word.contentcontrol#parentcontentcontrolornullobject)|Получает элемент управления содержимым, содержащий элемент управления содержимым.|
||[parentTable](/javascript/api/word/word.contentcontrol#parenttable)|Возвращает таблицу, содержащую элемент управления содержимым.|
||[parentTableCell](/javascript/api/word/word.contentcontrol#parenttablecell)|Возвращает ячейку таблицы, содержащую элемент управления содержимым.|
||[паренттаблецеллорнуллобжект](/javascript/api/word/word.contentcontrol#parenttablecellornullobject)|Возвращает ячейку таблицы, содержащую элемент управления содержимым.|
||[паренттаблеорнуллобжект](/javascript/api/word/word.contentcontrol#parenttableornullobject)|Возвращает таблицу, содержащую элемент управления содержимым.|
||[Подтип](/javascript/api/word/word.contentcontrol#subtype)|Возвращает подтип элемента управления содержимым.|
||[Table](/javascript/api/word/word.contentcontrol#tables)|Возвращает коллекцию объектов таблиц в элементе управления содержимым.|
||[Split (разделители: String [], многопараграфный?: Boolean, trimDelimiters?: Boolean, trimSpacing?: Boolean)](/javascript/api/word/word.contentcontrol#split-delimiters--multiparagraphs--trimdelimiters--trimspacing-)|Разделяет элемент управления содержимым на дочерние диапазоны с помощью разделителей.|
||[styleBuiltIn](/javascript/api/word/word.contentcontrol#stylebuiltin)|Возвращает или задает имя встроенного стиля для элемента управления содержимым.|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getByIdOrNullObject (ID: число)](/javascript/api/word/word.contentcontrolcollection#getbyidornullobject-id-)|Возвращает элемент управления содержимым по его идентификатору.|
||[Жетбитипес (Types: Word. ContentControlType [])](/javascript/api/word/word.contentcontrolcollection#getbytypes-types-)|Возвращает элементы управления контентом с указанными типами и/или подтипами.|
||[getFirst()](/javascript/api/word/word.contentcontrolcollection#getfirst--)|Возвращает первый элемент управления содержимым в коллекции.|
||[Жетфирсторнуллобжект ()](/javascript/api/word/word.contentcontrolcollection#getfirstornullobject--)|Возвращает первый элемент управления содержимым в коллекции.|
|[CustomProperty](/javascript/api/word/word.customproperty)|[delete()](/javascript/api/word/word.customproperty#delete--)|Удаляет настраиваемое свойство.|
||[key](/javascript/api/word/word.customproperty#key)|Возвращает ключ настраиваемого свойства.|
||[type](/javascript/api/word/word.customproperty#type)|Получает тип значения настраиваемого свойства.|
||[value](/javascript/api/word/word.customproperty#value)|Получает или задает значение настраиваемого свойства.|
|[CustomPropertyCollection](/javascript/api/word/word.custompropertycollection)|[Add (Key: строка, Value: Any)](/javascript/api/word/word.custompropertycollection#add-key--value-)|Создает или задает настраиваемое свойство.|
||[deleteAll ()](/javascript/api/word/word.custompropertycollection#deleteall--)|Удаляет все настраиваемые свойства в коллекции.|
||[getCount()](/javascript/api/word/word.custompropertycollection#getcount--)|Получает количество настраиваемых свойств.|
||[getItem(key: string)](/javascript/api/word/word.custompropertycollection#getitem-key-)|Возвращает объект настраиваемого свойства по ключу, указываемому без учета регистра.|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.custompropertycollection#getitemornullobject-key-)|Возвращает объект настраиваемого свойства по ключу, указываемому без учета регистра.|
||[items](/javascript/api/word/word.custompropertycollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Document](/javascript/api/word/word.document)|[properties](/javascript/api/word/word.document#properties)|Получает свойства документа.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[Open ()](/javascript/api/word/word.documentcreated#open--)|Открывает документ.|
||[body](/javascript/api/word/word.documentcreated#body)|Возвращает объект Body документа.|
||[contentControls](/javascript/api/word/word.documentcreated#contentcontrols)|Возвращает коллекцию объектов элементов управления содержимым в документе.|
||[properties](/javascript/api/word/word.documentcreated#properties)|Получает свойства документа.|
||[сохраняем](/javascript/api/word/word.documentcreated#saved)|Указывает, сохранены ли изменения, внесенные в документ.|
||[sections](/javascript/api/word/word.documentcreated#sections)|Получает коллекцию объектов Section в документе.|
||[save()](/javascript/api/word/word.documentcreated#save--)|Сохраняет документ.|
|[DocumentProperties](/javascript/api/word/word.documentproperties)|[Редактирование](/javascript/api/word/word.documentproperties#author)|Возвращает или задает автора документа.|
||[категории](/javascript/api/word/word.documentproperties#category)|Возвращает или задает категорию документа.|
||[comments](/javascript/api/word/word.documentproperties#comments)|Возвращает или задает примечания к документу.|
||[company](/javascript/api/word/word.documentproperties#company)|Возвращает или задает компанию документа.|
||[format](/javascript/api/word/word.documentproperties#format)|Возвращает или задает формат документа.|
||[keyword](/javascript/api/word/word.documentproperties#keywords)|Возвращает или задает ключевые слова документа.|
||[manager](/javascript/api/word/word.documentproperties#manager)|Возвращает или задает менеджера документа.|
||[applicationName](/javascript/api/word/word.documentproperties#applicationname)|Возвращает имя приложения для документа.|
||[creationDate](/javascript/api/word/word.documentproperties#creationdate)|Возвращает дату создания документа.|
||[customProperties](/javascript/api/word/word.documentproperties#customproperties)|Возвращает коллекцию настраиваемых свойств документа.|
||[lastAuthor](/javascript/api/word/word.documentproperties#lastauthor)|Получает последнего автора документа.|
||[lastPrintDate](/javascript/api/word/word.documentproperties#lastprintdate)|Возвращает дату последней печати документа.|
||[lastSaveTime](/javascript/api/word/word.documentproperties#lastsavetime)|Возвращает время последнего сохранения документа.|
||[revisionNumber](/javascript/api/word/word.documentproperties#revisionnumber)|Возвращает номер редакции документа.|
||[защиты](/javascript/api/word/word.documentproperties#security)|Получает параметры безопасности документа.|
||[template](/javascript/api/word/word.documentproperties#template)|Возвращает шаблон документа.|
||[subject](/javascript/api/word/word.documentproperties#subject)|Возвращает или задает тему документа.|
||[заголовок](/javascript/api/word/word.documentproperties#title)|Возвращает или задает название документа.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[GetNext ()](/javascript/api/word/word.inlinepicture#getnext--)|Возвращает следующий встроенный рисунок.|
||[getNextOrNullObject ()](/javascript/api/word/word.inlinepicture#getnextornullobject--)|Возвращает следующий встроенный рисунок.|
||[Тип-Range (rangeLocation?: Word. RangeLocation)](/javascript/api/word/word.inlinepicture#getrange-rangelocation-)|Возвращает рисунок (либо его начальную или конечную точку) в виде диапазона.|
||[parentContentControlOrNullObject](/javascript/api/word/word.inlinepicture#parentcontentcontrolornullobject)|Возвращает элемент управления содержимым, который содержит встроенный рисунок.|
||[parentTable](/javascript/api/word/word.inlinepicture#parenttable)|Возвращает таблицу, содержащую встроенный рисунок.|
||[parentTableCell](/javascript/api/word/word.inlinepicture#parenttablecell)|Возвращает ячейку таблицы, содержащую встроенный рисунок.|
||[паренттаблецеллорнуллобжект](/javascript/api/word/word.inlinepicture#parenttablecellornullobject)|Возвращает ячейку таблицы, содержащую встроенный рисунок.|
||[паренттаблеорнуллобжект](/javascript/api/word/word.inlinepicture#parenttableornullobject)|Возвращает таблицу, содержащую встроенный рисунок.|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[getFirst()](/javascript/api/word/word.inlinepicturecollection#getfirst--)|Возвращает первый встроенный рисунок в коллекции.|
||[Жетфирсторнуллобжект ()](/javascript/api/word/word.inlinepicturecollection#getfirstornullobject--)|Возвращает первый встроенный рисунок в коллекции.|
|[List](/javascript/api/word/word.list)|[getLevelParagraphs (Level: число)](/javascript/api/word/word.list#getlevelparagraphs-level-)|Возвращает абзацы, обнаруженные на указанном уровне списка.|
||[getLevelString (Level: число)](/javascript/api/word/word.list#getlevelstring-level-)|Получает маркер, число или изображение на заданном уровне в виде строки.|
||[insertParagraph (paragraphText: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.list#insertparagraph-paragraphtext--insertlocation-)|Вставляет абзац в указанном расположении.|
||[id](/javascript/api/word/word.list#id)|Получает идентификатор списка.|
||[levelExistences](/javascript/api/word/word.list#levelexistences)|Проверяет наличие каждого из 9 уровней в списке.|
||[levelTypes](/javascript/api/word/word.list#leveltypes)|Возвращает типы всех 9 уровней списка.|
||[paragraphs](/javascript/api/word/word.list#paragraphs)|Возвращает абзацы в списке.|
||[setLevelAlignment (Level: число, выравнивание: Word. alignment)](/javascript/api/word/word.list#setlevelalignment-level--alignment-)|Задает выравнивание маркера, числа или рисунка на указанном уровне в списке.|
||[setLevelBullet (Level: число, listBullet: Word. ListBullet, charCode?: число, fontName?: строка)](/javascript/api/word/word.list#setlevelbullet-level--listbullet--charcode--fontname-)|Задает формат маркеров на указанном уровне списка.|
||[setLevelIndents (Level: число, textIndent: число, Буллетнумберпиктуреиндент: число)](/javascript/api/word/word.list#setlevelindents-level--textindent--bulletnumberpictureindent-)|Задает два отступа на указанном уровне списка.|
||[setLevelNumbering (Level: число, listNumbering: Word. ListNumbering, formatString?: массив<строковый \| номер>)](/javascript/api/word/word.list#setlevelnumbering-level--listnumbering--formatstring-)|Задает формат нумерации на указанном уровне списка.|
||[setLevelStartingNumber (Level: число, startingNumber: число)](/javascript/api/word/word.list#setlevelstartingnumber-level--startingnumber-)|Задает начальный номер на указанном уровне списка.|
|[ListCollection](/javascript/api/word/word.listcollection)|[getById(id: number)](/javascript/api/word/word.listcollection#getbyid-id-)|Возвращает список по идентификатору.|
||[getByIdOrNullObject (ID: число)](/javascript/api/word/word.listcollection#getbyidornullobject-id-)|Возвращает список по идентификатору.|
||[getFirst()](/javascript/api/word/word.listcollection#getfirst--)|Возвращает первый список в коллекции.|
||[Жетфирсторнуллобжект ()](/javascript/api/word/word.listcollection#getfirstornullobject--)|Возвращает первый список в коллекции.|
||[getItem(index: number)](/javascript/api/word/word.listcollection#getitem-index-)|Возвращает объект списка по индексу в коллекции.|
||[items](/javascript/api/word/word.listcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[ListItem](/javascript/api/word/word.listitem)|[тип-предок (parentOnly?: Boolean)](/javascript/api/word/word.listitem#getancestor-parentonly-)|Возвращает родительский элемент или ближайшего предка (если родительского элемента нет) для данного элемента списка.|
||[getAncestorOrNullObject (parentOnly?: Boolean)](/javascript/api/word/word.listitem#getancestorornullobject-parentonly-)|Возвращает родительский элемент или ближайшего предка (если родительского элемента нет) для данного элемента списка.|
||[дочерние элементы (directChildrenOnly?: Boolean)](/javascript/api/word/word.listitem#getdescendants-directchildrenonly-)|Возвращает всех потомков элемента списка.|
||[level](/javascript/api/word/word.listitem#level)|Возвращает или задает уровень элемента в списке.|
||[listString](/javascript/api/word/word.listitem#liststring)|Получает маркер элемента списка, число или изображение в виде строки.|
||[siblingIndex](/javascript/api/word/word.listitem#siblingindex)|Возвращает порядковый номер элемента списка относительно элементов того же уровня.|
|[Paragraph](/javascript/api/word/word.paragraph)|[attachToList (listId: число, Level: число)](/javascript/api/word/word.paragraph#attachtolist-listid--level-)|Позволяет присоединить абзац к существующему списку на указанном уровне.|
||[Детачфромлист ()](/javascript/api/word/word.paragraph#detachfromlist--)|Перемещает абзац за пределы списка (если он является элементом списка).|
||[GetNext ()](/javascript/api/word/word.paragraph#getnext--)|Возвращает следующий абзац.|
||[getNextOrNullObject ()](/javascript/api/word/word.paragraph#getnextornullobject--)|Возвращает следующий абзац.|
||[Previous ()](/javascript/api/word/word.paragraph#getprevious--)|Возвращает предыдущий абзац.|
||[getPreviousOrNullObject ()](/javascript/api/word/word.paragraph#getpreviousornullobject--)|Возвращает предыдущий абзац.|
||[Тип-Range (rangeLocation?: Word. RangeLocation)](/javascript/api/word/word.paragraph#getrange-rangelocation-)|Возвращает весь абзац (либо его начальную или конечную точку) в виде диапазона.|
||[getTextRanges (endingMarks: строка [], trimSpacing?: Boolean)](/javascript/api/word/word.paragraph#gettextranges-endingmarks--trimspacing-)|Получает текстовые диапазоны в абзаце с помощью знаков препинания и/или других конечных меток.|
||[insertTable (rowCount: число, columnCount: число, insertLocation: Word. InsertLocation, Values?: строка [] [])](/javascript/api/word/word.paragraph#inserttable-rowcount--columncount--insertlocation--values-)|Вставляет таблицу с указанным количеством строк и столбцов.|
||[isLastParagraph](/javascript/api/word/word.paragraph#islastparagraph)|Указывает, что абзац является последним в родительском тексте.|
||[isListItem](/javascript/api/word/word.paragraph#islistitem)|Проверяет, является ли абзац элементом списка.|
||[list](/javascript/api/word/word.paragraph#list)|Возвращает объект List, к которому относится абзац.|
||[listItem](/javascript/api/word/word.paragraph#listitem)|Возвращает объект ListItem для абзаца.|
||[листитеморнуллобжект](/javascript/api/word/word.paragraph#listitemornullobject)|Возвращает объект ListItem для абзаца.|
||[листорнуллобжект](/javascript/api/word/word.paragraph#listornullobject)|Возвращает объект List, к которому относится абзац.|
||[parentBody](/javascript/api/word/word.paragraph#parentbody)|Возвращает родительский текст абзаца.|
||[parentContentControlOrNullObject](/javascript/api/word/word.paragraph#parentcontentcontrolornullobject)|Возвращает элемент управления содержимым, содержащий абзац.|
||[parentTable](/javascript/api/word/word.paragraph#parenttable)|Возвращает таблицу, содержащую абзац.|
||[parentTableCell](/javascript/api/word/word.paragraph#parenttablecell)|Возвращает ячейку таблицы, содержащую абзац.|
||[паренттаблецеллорнуллобжект](/javascript/api/word/word.paragraph#parenttablecellornullobject)|Возвращает ячейку таблицы, содержащую абзац.|
||[паренттаблеорнуллобжект](/javascript/api/word/word.paragraph#parenttableornullobject)|Возвращает таблицу, содержащую абзац.|
||[tableNestingLevel](/javascript/api/word/word.paragraph#tablenestinglevel)|Возвращает уровень таблицы, содержащей абзац.|
||[Split (разделители: String [], trimDelimiters?: Boolean, trimSpacing?: Boolean)](/javascript/api/word/word.paragraph#split-delimiters--trimdelimiters--trimspacing-)|Разделяет абзац на дочерние диапазоны с помощью разделителей.|
||[Стартневлист ()](/javascript/api/word/word.paragraph#startnewlist--)|Создает список, начинающийся с данного абзаца.|
||[styleBuiltIn](/javascript/api/word/word.paragraph#stylebuiltin)|Возвращает или задает имя встроенного стиля абзаца.|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[getFirst()](/javascript/api/word/word.paragraphcollection#getfirst--)|Возвращает первый абзац в коллекции.|
||[Жетфирсторнуллобжект ()](/javascript/api/word/word.paragraphcollection#getfirstornullobject--)|Возвращает первый абзац в коллекции.|
||[-Last ()](/javascript/api/word/word.paragraphcollection#getlast--)|Возвращает последний абзац в коллекции.|
||[Жетласторнуллобжект ()](/javascript/api/word/word.paragraphcollection#getlastornullobject--)|Возвращает последний абзац в коллекции.|
|[Range](/javascript/api/word/word.range)|[compareLocationWith (Range: Word. Range)](/javascript/api/word/word.range#comparelocationwith-range-)|Сравнивает расположение данного диапазона с расположением другого диапазона.|
||[Експандто (Range: Word. Range)](/javascript/api/word/word.range#expandto-range-)|Возвращает новый диапазон, который простирается в том или ином направлении от данного диапазона и перекрывает другой диапазон.|
||[Експандтурнуллобжект (Range: Word. Range)](/javascript/api/word/word.range#expandtoornullobject-range-)|Возвращает новый диапазон, который простирается в том или ином направлении от данного диапазона и перекрывает другой диапазон.|
||[Жесиперлинкранжес ()](/javascript/api/word/word.range#gethyperlinkranges--)|Возвращает дочерние диапазоны гиперссылок в данном диапазоне.|
||[getNextTextRange (endingMarks: строка [], trimSpacing?: Boolean)](/javascript/api/word/word.range#getnexttextrange-endingmarks--trimspacing-)|Получает следующий диапазон текста с использованием знаков препинания и/или других конечных меток.|
||[getNextTextRangeOrNullObject (endingMarks: строка [], trimSpacing?: Boolean)](/javascript/api/word/word.range#getnexttextrangeornullobject-endingmarks--trimspacing-)|Получает следующий диапазон текста с использованием знаков препинания и/или других конечных меток.|
||[Тип-Range (rangeLocation?: Word. RangeLocation)](/javascript/api/word/word.range#getrange-rangelocation-)|Клонирует диапазон либо получает его начальную или конечную точку в виде нового диапазона.|
||[getTextRanges (endingMarks: строка [], trimSpacing?: Boolean)](/javascript/api/word/word.range#gettextranges-endingmarks--trimspacing-)|Возвращает дочерние диапазоны текста в диапазоне с помощью знаков препинания и/или других конечных меток.|
||[hyperlink](/javascript/api/word/word.range#hyperlink)|Возвращает первую гиперссылку в диапазоне или задает для него гиперссылку.|
||[insertTable (rowCount: число, columnCount: число, insertLocation: Word. InsertLocation, Values?: строка [] [])](/javascript/api/word/word.range#inserttable-rowcount--columncount--insertlocation--values-)|Вставляет таблицу с указанным количеством строк и столбцов.|
||[Интерсектвис (Range: Word. Range)](/javascript/api/word/word.range#intersectwith-range-)|Возвращает новый диапазон, представляющий собой пересечение данного диапазона с другим.|
||[Интерсектвисорнуллобжект (Range: Word. Range)](/javascript/api/word/word.range#intersectwithornullobject-range-)|Возвращает новый диапазон, представляющий собой пересечение данного диапазона с другим.|
||[isEmpty](/javascript/api/word/word.range#isempty)|Проверяет, является ли длина диапазона нулевой.|
||[lists](/javascript/api/word/word.range#lists)|Возвращает коллекцию объектов списков в диапазоне.|
||[parentBody](/javascript/api/word/word.range#parentbody)|Возвращает родительский текст диапазона.|
||[parentContentControlOrNullObject](/javascript/api/word/word.range#parentcontentcontrolornullobject)|Возвращает элемент управления содержимым, содержащий диапазон.|
||[parentTable](/javascript/api/word/word.range#parenttable)|Возвращает таблицу, содержащую диапазон.|
||[parentTableCell](/javascript/api/word/word.range#parenttablecell)|Возвращает ячейку таблицы, содержащую диапазон.|
||[паренттаблецеллорнуллобжект](/javascript/api/word/word.range#parenttablecellornullobject)|Возвращает ячейку таблицы, содержащую диапазон.|
||[паренттаблеорнуллобжект](/javascript/api/word/word.range#parenttableornullobject)|Возвращает таблицу, содержащую диапазон.|
||[Table](/javascript/api/word/word.range#tables)|Возвращает коллекцию объектов таблиц в диапазоне.|
||[Split (разделители: String [], многопараграфный?: Boolean, trimDelimiters?: Boolean, trimSpacing?: Boolean)](/javascript/api/word/word.range#split-delimiters--multiparagraphs--trimdelimiters--trimspacing-)|Разделяет диапазон на дочерние диапазоны с помощью разделителей.|
||[styleBuiltIn](/javascript/api/word/word.range#stylebuiltin)|Возвращает или задает имя встроенного стиля диапазона.|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[getFirst()](/javascript/api/word/word.rangecollection#getfirst--)|Возвращает первый диапазон в коллекции.|
||[Жетфирсторнуллобжект ()](/javascript/api/word/word.rangecollection#getfirstornullobject--)|Возвращает первый диапазон в коллекции.|
|[RequestContext](/javascript/api/word/word.requestcontext)|[application](/javascript/api/word/word.requestcontext#application)|[API Set: WordApi 1,3] *|
|[Section](/javascript/api/word/word.section)|[GetNext ()](/javascript/api/word/word.section#getnext--)|Возвращает следующий раздел.|
||[getNextOrNullObject ()](/javascript/api/word/word.section#getnextornullobject--)|Возвращает следующий раздел.|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[getFirst()](/javascript/api/word/word.sectioncollection#getfirst--)|Возвращает первый раздел в коллекции.|
||[Жетфирсторнуллобжект ()](/javascript/api/word/word.sectioncollection#getfirstornullobject--)|Возвращает первый раздел в коллекции.|
|[Table](/javascript/api/word/word.table)|[addColumns (insertLocation: Word. InsertLocation, columnCount: число, Values?: строка [] [])](/javascript/api/word/word.table#addcolumns-insertlocation--columncount--values-)|Добавляет столбцы в начале или в конце таблицы, используя первый или последний из имеющихся столбцов в качестве шаблона.|
||[addRows (insertLocation: Word. InsertLocation, rowCount: число, Values?: строка [] [])](/javascript/api/word/word.table#addrows-insertlocation--rowcount--values-)|Добавляет строки в начале или в конце таблицы, используя первую или последнюю из имеющихся строк в качестве шаблона.|
||[ориентации](/javascript/api/word/word.table#alignment)|Возвращает или задает выравнивание таблицы по столбцу страницы.|
||[Аутофитвиндов ()](/javascript/api/word/word.table#autofitwindow--)|Автоматически подбирает ширину столбцов таблицы в соответствии с шириной окна.|
||[clear()](/javascript/api/word/word.table#clear--)|Очищает содержимое таблицы.|
||[delete()](/javascript/api/word/word.table#delete--)|Удаляет всю таблицу.|
||[deleteColumns (columnIndex: число, columnCount?: число)](/javascript/api/word/word.table#deletecolumns-columnindex--columncount-)|Удаляет определенные столбцы.|
||[deleteRows (rowIndex: число, rowCount?: число)](/javascript/api/word/word.table#deleterows-rowindex--rowcount-)|Удаляет определенные строки.|
||[Дистрибутеколумнс ()](/javascript/api/word/word.table#distributecolumns--)|Равномерно распределяет ширину столбцов.|
||[Граница (borderLocation: Word. BorderLocation)](/javascript/api/word/word.table#getborder-borderlocation-)|Возвращает стиль указанной границы.|
||[getCell(rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#getcell-rowindex--cellindex-)|Возвращает ячейку таблицы в указанной строке и указанном столбце.|
||[getCellOrNullObject (rowIndex: число, cellIndex: число)](/javascript/api/word/word.table#getcellornullobject-rowindex--cellindex-)|Возвращает ячейку таблицы в указанной строке и указанном столбце.|
||[Жетцеллпаддинг (cellPaddingLocation: Word. CellPaddingLocation)](/javascript/api/word/word.table#getcellpadding-cellpaddinglocation-)|Возвращает размер поля ячейки в точках.|
||[GetNext ()](/javascript/api/word/word.table#getnext--)|Возвращает следующую таблицу.|
||[getNextOrNullObject ()](/javascript/api/word/word.table#getnextornullobject--)|Возвращает следующую таблицу.|
||[Жетпараграфафтер ()](/javascript/api/word/word.table#getparagraphafter--)|Возвращает абзац после таблицы.|
||[Жетпараграфафтерорнуллобжект ()](/javascript/api/word/word.table#getparagraphafterornullobject--)|Возвращает абзац после таблицы.|
||[Жетпараграфбефоре ()](/javascript/api/word/word.table#getparagraphbefore--)|Возвращает абзац перед таблицей.|
||[Жетпараграфбефореорнуллобжект ()](/javascript/api/word/word.table#getparagraphbeforeornullobject--)|Возвращает абзац перед таблицей.|
||[Тип-Range (rangeLocation?: Word. RangeLocation)](/javascript/api/word/word.table#getrange-rangelocation-)|Возвращает диапазон, содержащий данную таблицу, либо диапазон в начале или в конце таблицы.|
||[headerRowCount](/javascript/api/word/word.table#headerrowcount)|Возвращает и задает количество строк заголовков.|
||[horizontalAlignment](/javascript/api/word/word.table#horizontalalignment)|Возвращает и задает горизонтальное выравнивание для каждой ячейки в таблице.|
||[игнорепункт](/javascript/api/word/word.table#ignorepunct)||
||[игнореспаце](/javascript/api/word/word.table#ignorespace)||
||[insertContentControl()](/javascript/api/word/word.table#insertcontentcontrol--)|Вставляет в таблицу элемент управления содержимым.|
||[insertParagraph (paragraphText: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.table#insertparagraph-paragraphtext--insertlocation-)|Вставляет абзац в указанном расположении.|
||[insertTable (rowCount: число, columnCount: число, insertLocation: Word. InsertLocation, Values?: строка [] [])](/javascript/api/word/word.table#inserttable-rowcount--columncount--insertlocation--values-)|Вставляет таблицу с указанным количеством строк и столбцов.|
||[matchCase](/javascript/api/word/word.table#matchcase)||
||[матчпрефикс](/javascript/api/word/word.table#matchprefix)||
||[матчсуффикс](/javascript/api/word/word.table#matchsuffix)||
||[матчвхолеворд](/javascript/api/word/word.table#matchwholeword)||
||[матчвилдкардс](/javascript/api/word/word.table#matchwildcards)||
||[font](/javascript/api/word/word.table#font)|Возвращает шрифт.|
||[isUniform](/javascript/api/word/word.table#isuniform)|Указывает, однородны ли все строки таблицы.|
||[nestingLevel](/javascript/api/word/word.table#nestinglevel)|Возвращает уровень вложенности таблицы.|
||[parentBody](/javascript/api/word/word.table#parentbody)|Возвращает родительский текст таблицы.|
||[parentContentControl](/javascript/api/word/word.table#parentcontentcontrol)|Возвращает элемент управления содержимым, содержащий таблицу.|
||[parentContentControlOrNullObject](/javascript/api/word/word.table#parentcontentcontrolornullobject)|Возвращает элемент управления содержимым, содержащий таблицу.|
||[parentTable](/javascript/api/word/word.table#parenttable)|Возвращает таблицу, которая содержит данную таблицу.|
||[parentTableCell](/javascript/api/word/word.table#parenttablecell)|Возвращает ячейку таблицы, содержащую данную таблицу.|
||[паренттаблецеллорнуллобжект](/javascript/api/word/word.table#parenttablecellornullobject)|Возвращает ячейку таблицы, содержащую данную таблицу.|
||[паренттаблеорнуллобжект](/javascript/api/word/word.table#parenttableornullobject)|Возвращает таблицу, которая содержит данную таблицу.|
||[Стро](/javascript/api/word/word.table#rowcount)|Получает количество строк в таблице.|
||[строки](/javascript/api/word/word.table#rows)|Возвращает все строки таблицы.|
||[Table](/javascript/api/word/word.table#tables)|Возвращает дочерние таблицы, вложенные на один уровень ниже.|
||[Search (searchText: строка, searchOptions?: Word. SearchOptions \| {игнорепункт?: Boolean игнореспаце?: Boolean matchCase?: Boolean матчпрефикс?: Boolean матчсуффикс?: Boolean матчвхолеворд?: Boolean матчвилдкардс?: Boolean})](/javascript/api/word/word.table#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Выполняет поиск с указанным SearchOptions в области объекта Table.|
||[SELECT (selectionMode?: Word. SelectionMode)](/javascript/api/word/word.table#select-selectionmode-)|Выбирает таблицу либо позицию в начале или в конце таблицы, а затем переходит к ней в Word.|
||[setCellPadding (cellPaddingLocation: Word. CellPaddingLocation, cellPadding: число)](/javascript/api/word/word.table#setcellpadding-cellpaddinglocation--cellpadding-)|Задает размер поля ячейки в точках.|
||[shadingColor](/javascript/api/word/word.table#shadingcolor)|Возвращает и задает цвет заливки.|
||[style](/javascript/api/word/word.table#style)|Возвращает или задает имя стиля для таблицы.|
||[styleBandedColumns](/javascript/api/word/word.table#stylebandedcolumns)|Возвращает и задает значение, указывающее, есть ли в таблице чередующиеся столбцы.|
||[styleBandedRows](/javascript/api/word/word.table#stylebandedrows)|Возвращает и задает значение, указывающее, есть ли в таблице чередующиеся строки.|
||[styleBuiltIn](/javascript/api/word/word.table#stylebuiltin)|Возвращает или задает имя встроенного стиля таблицы.|
||[styleFirstColumn](/javascript/api/word/word.table#stylefirstcolumn)|Возвращает и задает значение, указывающее, применен ли специальный стиль к первому столбцу таблицы.|
||[styleLastColumn](/javascript/api/word/word.table#stylelastcolumn)|Возвращает и задает значение, указывающее, применен ли специальный стиль к последнему столбцу таблицы.|
||[styleTotalRow](/javascript/api/word/word.table#styletotalrow)|Возвращает и задает значение, указывающее, применен ли специальный стиль к строке итогов (последней строке) таблицы.|
||[values](/javascript/api/word/word.table#values)|Возвращает и задает текстовые значения в таблице в виде двумерного массива JavaScript.|
||[verticalAlignment](/javascript/api/word/word.table#verticalalignment)|Возвращает и задает вертикальное выравнивание для каждой ячейки в таблице.|
||[width](/javascript/api/word/word.table#width)|Возвращает и задает ширину таблицы в точках.|
|[TableBorder](/javascript/api/word/word.tableborder)|[color](/javascript/api/word/word.tableborder#color)|Получает или задает цвет границы таблицы.|
||[type](/javascript/api/word/word.tableborder#type)|Возвращает или задает тип границы таблицы.|
||[width](/javascript/api/word/word.tableborder#width)|Возвращает или задает ширину границы таблицы в точках.|
|[TableCell](/javascript/api/word/word.tablecell)|[columnWidth](/javascript/api/word/word.tablecell#columnwidth)|Возвращает и задает ширину столбца ячейки в точках.|
||[Делетеколумн ()](/javascript/api/word/word.tablecell#deletecolumn--)|Удаляет столбец, содержащий данную ячейку.|
||[deleteRow ()](/javascript/api/word/word.tablecell#deleterow--)|Удаляет строку, содержащую данную ячейку.|
||[Граница (borderLocation: Word. BorderLocation)](/javascript/api/word/word.tablecell#getborder-borderlocation-)|Возвращает стиль указанной границы.|
||[Жетцеллпаддинг (cellPaddingLocation: Word. CellPaddingLocation)](/javascript/api/word/word.tablecell#getcellpadding-cellpaddinglocation-)|Возвращает размер поля ячейки в точках.|
||[GetNext ()](/javascript/api/word/word.tablecell#getnext--)|Возвращает следующую ячейку.|
||[getNextOrNullObject ()](/javascript/api/word/word.tablecell#getnextornullobject--)|Возвращает следующую ячейку.|
||[horizontalAlignment](/javascript/api/word/word.tablecell#horizontalalignment)|Возвращает и задает горизонтальное выравнивание ячейки.|
||[insertColumns (insertLocation: Word. InsertLocation, columnCount: число, Values?: строка [] [])](/javascript/api/word/word.tablecell#insertcolumns-insertlocation--columncount--values-)|Добавляет столбцы слева или справа от ячейки, используя столбец этой ячейки в качестве шаблона.|
||[insertRows (insertLocation: Word. InsertLocation, rowCount: число, Values?: строка [] [])](/javascript/api/word/word.tablecell#insertrows-insertlocation--rowcount--values-)|Вставляет строки над ячейкой или под ней, используя строку этой ячейки в качестве шаблона.|
||[body](/javascript/api/word/word.tablecell#body)|Возвращает объект тела ячейки.|
||[cellIndex](/javascript/api/word/word.tablecell#cellindex)|Получает индекс ячейки в строке.|
||[parentRow](/javascript/api/word/word.tablecell#parentrow)|Получает родительскую строку ячейки.|
||[parentTable](/javascript/api/word/word.tablecell#parenttable)|Возвращает родительскую таблицу ячейки.|
||[rowIndex](/javascript/api/word/word.tablecell#rowindex)|Получает индекс строки ячейки в таблице.|
||[width](/javascript/api/word/word.tablecell#width)|Возвращает ширину ячейки в точках.|
||[setCellPadding (cellPaddingLocation: Word. CellPaddingLocation, cellPadding: число)](/javascript/api/word/word.tablecell#setcellpadding-cellpaddinglocation--cellpadding-)|Задает размер поля ячейки в точках.|
||[shadingColor](/javascript/api/word/word.tablecell#shadingcolor)|Возвращает или задает цвет заливки ячейки.|
||[value](/javascript/api/word/word.tablecell#value)|Возвращает и задает текст ячейки.|
||[verticalAlignment](/javascript/api/word/word.tablecell#verticalalignment)|Возвращает и задает вертикальное выравнивание ячейки.|
|[TableCellCollection](/javascript/api/word/word.tablecellcollection)|[getFirst()](/javascript/api/word/word.tablecellcollection#getfirst--)|Возвращает первую ячейку таблицы в коллекции.|
||[Жетфирсторнуллобжект ()](/javascript/api/word/word.tablecellcollection#getfirstornullobject--)|Возвращает первую ячейку таблицы в коллекции.|
||[items](/javascript/api/word/word.tablecellcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[TableCollection](/javascript/api/word/word.tablecollection)|[getFirst()](/javascript/api/word/word.tablecollection#getfirst--)|Возвращает первую таблицу в коллекции.|
||[Жетфирсторнуллобжект ()](/javascript/api/word/word.tablecollection#getfirstornullobject--)|Возвращает первую таблицу в коллекции.|
||[items](/javascript/api/word/word.tablecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[TableRow](/javascript/api/word/word.tablerow)|[clear()](/javascript/api/word/word.tablerow#clear--)|Очищает содержимое строки.|
||[delete()](/javascript/api/word/word.tablerow#delete--)|Удаляет всю строку.|
||[Граница (borderLocation: Word. BorderLocation)](/javascript/api/word/word.tablerow#getborder-borderlocation-)|Возвращает стиль границ ячеек в строке.|
||[Жетцеллпаддинг (cellPaddingLocation: Word. CellPaddingLocation)](/javascript/api/word/word.tablerow#getcellpadding-cellpaddinglocation-)|Возвращает размер поля ячейки в точках.|
||[GetNext ()](/javascript/api/word/word.tablerow#getnext--)|Возвращает следующую строку.|
||[getNextOrNullObject ()](/javascript/api/word/word.tablerow#getnextornullobject--)|Возвращает следующую строку.|
||[horizontalAlignment](/javascript/api/word/word.tablerow#horizontalalignment)|Возвращает и задает горизонтальное выравнивание для каждой ячейки в строке.|
||[игнорепункт](/javascript/api/word/word.tablerow#ignorepunct)||
||[игнореспаце](/javascript/api/word/word.tablerow#ignorespace)||
||[insertRows (insertLocation: Word. InsertLocation, rowCount: число, Values?: строка [] [])](/javascript/api/word/word.tablerow#insertrows-insertlocation--rowcount--values-)|Вставляет строки, используя данную строку в качестве шаблона.|
||[matchCase](/javascript/api/word/word.tablerow#matchcase)||
||[матчпрефикс](/javascript/api/word/word.tablerow#matchprefix)||
||[матчсуффикс](/javascript/api/word/word.tablerow#matchsuffix)||
||[матчвхолеворд](/javascript/api/word/word.tablerow#matchwholeword)||
||[матчвилдкардс](/javascript/api/word/word.tablerow#matchwildcards)||
||[preferredHeight](/javascript/api/word/word.tablerow#preferredheight)|Возвращает и задает предпочитаемую высоту строки в точках.|
||[cellCount](/javascript/api/word/word.tablerow#cellcount)|Получает количество ячеек в строке.|
||[диапазона](/javascript/api/word/word.tablerow#cells)|Возвращает ячейки.|
||[font](/javascript/api/word/word.tablerow#font)|Возвращает шрифт.|
||[isHeader](/javascript/api/word/word.tablerow#isheader)|Проверяет, является ли элемент строкой заголовков.|
||[parentTable](/javascript/api/word/word.tablerow#parenttable)|Возвращает родительскую таблицу.|
||[rowIndex](/javascript/api/word/word.tablerow#rowindex)|Получает индекс строки в родительской таблице.|
||[Search (searchText: строка, searchOptions?: Word. SearchOptions \| {игнорепункт?: Boolean игнореспаце?: Boolean matchCase?: Boolean матчпрефикс?: Boolean матчсуффикс?: Boolean матчвхолеворд?: Boolean матчвилдкардс?: Boolean})](/javascript/api/word/word.tablerow#search-searchtext--searchoptions--ignorepunct--ignorespace--matchcase--matchprefix--matchsuffix--matchwholeword--matchwildcards-)|Выполняет поиск с указанным SearchOptions в области действия строки.|
||[SELECT (selectionMode?: Word. SelectionMode)](/javascript/api/word/word.tablerow#select-selectionmode-)|Выбирает строку и переходит к ней в Word.|
||[setCellPadding (cellPaddingLocation: Word. CellPaddingLocation, cellPadding: число)](/javascript/api/word/word.tablerow#setcellpadding-cellpaddinglocation--cellpadding-)|Задает размер поля ячейки в точках.|
||[shadingColor](/javascript/api/word/word.tablerow#shadingcolor)|Возвращает и задает цвет заливки.|
||[values](/javascript/api/word/word.tablerow#values)|Возвращает и задает текстовые значения в строке в виде 2D-массива JavaScript.|
||[verticalAlignment](/javascript/api/word/word.tablerow#verticalalignment)|Возвращает и задает вертикальное выравнивание ячеек в строке.|
|[TableRowCollection](/javascript/api/word/word.tablerowcollection)|[getFirst()](/javascript/api/word/word.tablerowcollection#getfirst--)|Возвращает первую строку в коллекции.|
||[Жетфирсторнуллобжект ()](/javascript/api/word/word.tablerowcollection#getfirstornullobject--)|Возвращает первую строку в коллекции.|
||[items](/javascript/api/word/word.tablerowcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Word](/javascript/api/word)
- [Наборы обязательных элементов API JavaScript для Word](word-api-requirement-sets.md)
