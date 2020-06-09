---
title: Набор обязательных элементов API JavaScript для Word 1,3
description: Сведения о наборе требований WordApi 1,3
ms.date: 07/25/2019
ms.prod: word
localization_priority: Normal
ms.openlocfilehash: 15ec2129f53d0b408191ceb595f1fe115feb0d1a
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611297"
---
# <a name="whats-new-in-word-javascript-api-13"></a>Новые возможности API JavaScript для Word 1.3

WordApi 1,3 добавлена дополнительная поддержка элементов управления содержимым, настраиваемых XML и параметров на уровне документа.

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в наборе обязательных элементов API JavaScript для Word 1,3. Чтобы просмотреть справочную документацию по API для всех API, поддерживаемых в наборе обязательных элементов API JavaScript для Word 1,3 или более ранней версии, обратитесь к разделам [API Word в наборе требований 1,3](/javascript/api/word?view=word-js-1.3)

| Класс | Поля | Описание |
|:---|:---|:---|
|[Application](/javascript/api/word/word.application)|[Креатедокумент (base64File?: строка)](/javascript/api/word/word.application#createdocument-base64file-)|Создает новый документ, используя необязательный docx файл с кодировкой base64.|
|[Body](/javascript/api/word/word.body)|[Тип-Range (rangeLocation?: Word. RangeLocation)](/javascript/api/word/word.body#getrange-rangelocation-)|Возвращает весь основной текст (либо его начальную или конечную точку) в виде диапазона.|
||[insertTable (rowCount: число, columnCount: число, insertLocation: Word. InsertLocation, Values?: строка [] [])](/javascript/api/word/word.body#inserttable-rowcount--columncount--insertlocation--values-)|Вставляет таблицу с указанным количеством строк и столбцов. Возможные значения insertLocation: Start и End.|
||[lists](/javascript/api/word/word.body#lists)|Возвращает коллекцию объектов списков в основном тексте. Только для чтения.|
||[parentBody](/javascript/api/word/word.body#parentbody)|Возвращает родительский текст основного текста. Например, родительским текстом ячейки таблицы может быть заголовок. Выдает ошибку, если родительского текста не существует. Только для чтения.|
||[parentBodyOrNullObject](/javascript/api/word/word.body#parentbodyornullobject)|Возвращает родительский текст основного текста. Например, родительским текстом ячейки таблицы может быть заголовок. Возвращает пустой объект, если родительского текста не существует. Только для чтения.|
||[parentContentControlOrNullObject](/javascript/api/word/word.body#parentcontentcontrolornullobject)|Получает элемент управления содержимым, содержащий документ или раздел. Возвращает нулевой объект, если родительский элемент управления содержимым отсутствует. Только для чтения.|
||[parentSection](/javascript/api/word/word.body#parentsection)|Возвращает родительский раздел основного текста. Создает исключение, если родительский раздел отсутствует. Только для чтения.|
||[parentSectionOrNullObject](/javascript/api/word/word.body#parentsectionornullobject)|Возвращает родительский раздел основного текста. Возвращает нулевой объект, если родительский раздел отсутствует. Только для чтения.|
||[Table](/javascript/api/word/word.body#tables)|Возвращает коллекцию объектов таблиц в основном тексте. Только для чтения.|
||[type](/javascript/api/word/word.body#type)|Возвращает тип основного текста. Поддерживаемые типы: MainDoc, Section, Header, Footer и TableCell. Только для чтения.|
||[styleBuiltIn](/javascript/api/word/word.body#stylebuiltin)|Возвращает или задает имя встроенного стиля основного текста. Используйте это свойство для встроенных стилей, поддерживающих несколько языковых стандартов. Чтобы использовать пользовательские стили или локализованные имена стилей, применяйте свойство style.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[Тип-Range (rangeLocation?: Word. RangeLocation)](/javascript/api/word/word.contentcontrol#getrange-rangelocation-)|Возвращает весь элемент управления содержимым (либо его начальную или конечную точку) в виде диапазона.|
||[getTextRanges (endingMarks: строка [], trimSpacing?: Boolean)](/javascript/api/word/word.contentcontrol#gettextranges-endingmarks--trimspacing-)|Получает диапазоны текста в элементе управления содержимым с помощью знаков препинания и/или других конечных меток.|
||[insertTable (rowCount: число, columnCount: число, insertLocation: Word. InsertLocation, Values?: строка [] [])](/javascript/api/word/word.contentcontrol#inserttable-rowcount--columncount--insertlocation--values-)|Вставляет таблицу с указанным количеством строк и столбцов в элемент управления содержимым или рядом с ним. Значение insertLocation может быть "Start", "End", "Before" или "After".|
||[lists](/javascript/api/word/word.contentcontrol#lists)|Возвращает коллекцию объектов списков в элементе управления содержимым. Только для чтения.|
||[parentBody](/javascript/api/word/word.contentcontrol#parentbody)|Возвращает родительский текст элемента управления содержимым. Только для чтения.|
||[parentContentControlOrNullObject](/javascript/api/word/word.contentcontrol#parentcontentcontrolornullobject)|Получает элемент управления содержимым, содержащий элемент управления содержимым. Возвращает нулевой объект, если родительский элемент управления содержимым отсутствует. Только для чтения.|
||[parentTable](/javascript/api/word/word.contentcontrol#parenttable)|Возвращает таблицу, содержащую элемент управления содержимым. Вызывается, если он не включен в таблицу. Только для чтения.|
||[parentTableCell](/javascript/api/word/word.contentcontrol#parenttablecell)|Возвращает ячейку таблицы, содержащую элемент управления содержимым. Создает исключение, если оно не находится в ячейке таблицы. Только для чтения.|
||[паренттаблецеллорнуллобжект](/javascript/api/word/word.contentcontrol#parenttablecellornullobject)|Возвращает ячейку таблицы, содержащую элемент управления содержимым. Если он находится не в ячейке таблицы, возвращается пустой объект. Только для чтения.|
||[паренттаблеорнуллобжект](/javascript/api/word/word.contentcontrol#parenttableornullobject)|Возвращает таблицу, содержащую элемент управления содержимым. Если он находится не в таблице, возвращается пустой объект. Только для чтения.|
||[Подтип](/javascript/api/word/word.contentcontrol#subtype)|Возвращает подтип элемента управления содержимым. Поддерживаемые подтипы: RichTextInline, RichTextParagraphs, RichTextTableCell, RichTextTableRow и RichTextTable для элементов управления форматированным текстом. Только для чтения.|
||[Table](/javascript/api/word/word.contentcontrol#tables)|Возвращает коллекцию объектов таблиц в элементе управления содержимым. Только для чтения.|
||[Split (разделители: String [], многопараграфный?: Boolean, trimDelimiters?: Boolean, trimSpacing?: Boolean)](/javascript/api/word/word.contentcontrol#split-delimiters--multiparagraphs--trimdelimiters--trimspacing-)|Разделяет элемент управления содержимым на дочерние диапазоны с помощью разделителей.|
||[styleBuiltIn](/javascript/api/word/word.contentcontrol#stylebuiltin)|Возвращает или задает имя встроенного стиля для элемента управления содержимым. Используйте это свойство для встроенных стилей, поддерживающих несколько языковых стандартов. Чтобы использовать пользовательские стили или локализованные имена стилей, применяйте свойство style.|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getByIdOrNullObject (ID: число)](/javascript/api/word/word.contentcontrolcollection#getbyidornullobject-id-)|Возвращает элемент управления содержимым по его идентификатору. Возвращает нулевой объект, если в этой коллекции нет элемента управления контентом с идентификатором.|
||[Жетбитипес (Types: Word. ContentControlType [])](/javascript/api/word/word.contentcontrolcollection#getbytypes-types-)|Возвращает элементы управления контентом с указанными типами и/или подтипами.|
||[getFirst()](/javascript/api/word/word.contentcontrolcollection#getfirst--)|Возвращает первый элемент управления содержимым в коллекции. Вызывается, если коллекция пуста.|
||[Жетфирсторнуллобжект ()](/javascript/api/word/word.contentcontrolcollection#getfirstornullobject--)|Возвращает первый элемент управления содержимым в коллекции. Возвращает нулевой объект, если коллекция пуста.|
|[CustomProperty](/javascript/api/word/word.customproperty)|[delete()](/javascript/api/word/word.customproperty#delete--)|Удаляет настраиваемое свойство.|
||[key](/javascript/api/word/word.customproperty#key)|Возвращает ключ настраиваемого свойства. Только для чтения.|
||[type](/javascript/api/word/word.customproperty#type)|Получает тип значения настраиваемого свойства. Возможные значения: String, Number, Date, Boolean. Только для чтения.|
||[value](/javascript/api/word/word.customproperty#value)|Получает или задает значение настраиваемого свойства. Обратите внимание, что несмотря на то, что Word в Интернете и в формате docx допускает, чтобы эти свойства были произвольно длинными, настольная версия Word усекает строковые значения до 255 16-разрядных символов (возможно, создавая недопустимый символ Юникода, нарушая суррогатную пара).|
|[кустомпропертиколлектион](/javascript/api/word/word.custompropertycollection)|[Add (Key: строка, Value: Any)](/javascript/api/word/word.custompropertycollection#add-key--value-)|Создает или задает настраиваемое свойство.|
||[deleteAll ()](/javascript/api/word/word.custompropertycollection#deleteall--)|Удаляет все настраиваемые свойства в коллекции.|
||[getCount()](/javascript/api/word/word.custompropertycollection#getcount--)|Получает количество настраиваемых свойств.|
||[getItem(key: string)](/javascript/api/word/word.custompropertycollection#getitem-key-)|Возвращает объект настраиваемого свойства по ключу, указываемому без учета регистра. Вызывается, если настраиваемое свойство не существует.|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.custompropertycollection#getitemornullobject-key-)|Возвращает объект настраиваемого свойства по ключу, указываемому без учета регистра. Возвращает нулевой объект, если настраиваемое свойство не существует.|
||[items](/javascript/api/word/word.custompropertycollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Document](/javascript/api/word/word.document)|[properties](/javascript/api/word/word.document#properties)|Получает свойства документа. Только для чтения.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[Open ()](/javascript/api/word/word.documentcreated#open--)|Открывает документ.|
||[body](/javascript/api/word/word.documentcreated#body)|Возвращает объект Body документа. Текст — это текст, который исключает заголовки, нижние колонтитулы, сноски, текстовые поля и т. д. Только для чтения.|
||[contentControls](/javascript/api/word/word.documentcreated#contentcontrols)|Возвращает коллекцию объектов элементов управления содержимым в документе. Сюда входят элементы управления содержимым в тексте документа, верхних и нижних колонтитулов, текстовых полях и т. д. Только для чтения.|
||[properties](/javascript/api/word/word.documentcreated#properties)|Получает свойства документа. Только для чтения.|
||[сохраняем](/javascript/api/word/word.documentcreated#saved)|Указывает, сохранены ли изменения, внесенные в документ. Значение true указывает на то, что с момента последнего сохранения в документ не вносились изменения. Только для чтения.|
||[sections](/javascript/api/word/word.documentcreated#sections)|Получает коллекцию объектов Section в документе. Только для чтения.|
||[save()](/javascript/api/word/word.documentcreated#save--)|Сохраняет документ. При этом используется соглашение об именовании файлов Word по умолчанию, если документ ранее не сохранялся.|
|[DocumentProperties](/javascript/api/word/word.documentproperties)|[Редактирование](/javascript/api/word/word.documentproperties#author)|Возвращает или задает автора документа.|
||[категории](/javascript/api/word/word.documentproperties#category)|Возвращает или задает категорию документа.|
||[comments](/javascript/api/word/word.documentproperties#comments)|Возвращает или задает примечания к документу.|
||[company](/javascript/api/word/word.documentproperties#company)|Возвращает или задает компанию документа.|
||[format](/javascript/api/word/word.documentproperties#format)|Возвращает или задает формат документа.|
||[keyword](/javascript/api/word/word.documentproperties#keywords)|Возвращает или задает ключевые слова документа.|
||[manager](/javascript/api/word/word.documentproperties#manager)|Возвращает или задает менеджера документа.|
||[applicationName](/javascript/api/word/word.documentproperties#applicationname)|Возвращает имя приложения для документа. Только для чтения.|
||[creationDate](/javascript/api/word/word.documentproperties#creationdate)|Возвращает дату создания документа. Только для чтения.|
||[customProperties](/javascript/api/word/word.documentproperties#customproperties)|Возвращает коллекцию настраиваемых свойств документа. Только для чтения.|
||[lastAuthor](/javascript/api/word/word.documentproperties#lastauthor)|Получает последнего автора документа. Только для чтения.|
||[lastPrintDate](/javascript/api/word/word.documentproperties#lastprintdate)|Возвращает дату последней печати документа. Только для чтения.|
||[lastSaveTime](/javascript/api/word/word.documentproperties#lastsavetime)|Возвращает время последнего сохранения документа. Только для чтения.|
||[revisionNumber](/javascript/api/word/word.documentproperties#revisionnumber)|Возвращает номер редакции документа. Только для чтения.|
||[защиты](/javascript/api/word/word.documentproperties#security)|Возвращает сведения о безопасности документа. Только для чтения.|
||[template](/javascript/api/word/word.documentproperties#template)|Возвращает шаблон документа. Только для чтения.|
||[subject](/javascript/api/word/word.documentproperties#subject)|Возвращает или задает тему документа.|
||[заголовок](/javascript/api/word/word.documentproperties#title)|Возвращает или задает название документа.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[GetNext ()](/javascript/api/word/word.inlinepicture#getnext--)|Возвращает следующий встроенный рисунок. Вызывается, если данное встроенное изображение является последним.|
||[getNextOrNullObject ()](/javascript/api/word/word.inlinepicture#getnextornullobject--)|Возвращает следующий встроенный рисунок. Возвращает пустой объект, если данное встроенное изображение является последним.|
||[Тип-Range (rangeLocation?: Word. RangeLocation)](/javascript/api/word/word.inlinepicture#getrange-rangelocation-)|Возвращает рисунок (либо его начальную или конечную точку) в виде диапазона.|
||[parentContentControlOrNullObject](/javascript/api/word/word.inlinepicture#parentcontentcontrolornullobject)|Возвращает элемент управления содержимым, который содержит встроенный рисунок. Возвращает нулевой объект, если родительский элемент управления содержимым отсутствует. Только для чтения.|
||[parentTable](/javascript/api/word/word.inlinepicture#parenttable)|Возвращает таблицу, содержащую встроенный рисунок. Вызывается, если он не включен в таблицу. Только для чтения.|
||[parentTableCell](/javascript/api/word/word.inlinepicture#parenttablecell)|Возвращает ячейку таблицы, содержащую встроенный рисунок. Создает исключение, если оно не находится в ячейке таблицы. Только для чтения.|
||[паренттаблецеллорнуллобжект](/javascript/api/word/word.inlinepicture#parenttablecellornullobject)|Возвращает ячейку таблицы, содержащую встроенный рисунок. Если он находится не в ячейке таблицы, возвращается пустой объект. Только для чтения.|
||[паренттаблеорнуллобжект](/javascript/api/word/word.inlinepicture#parenttableornullobject)|Возвращает таблицу, содержащую встроенный рисунок. Если он находится не в таблице, возвращается пустой объект. Только для чтения.|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[getFirst()](/javascript/api/word/word.inlinepicturecollection#getfirst--)|Возвращает первый встроенный рисунок в коллекции. Вызывается, если коллекция пуста.|
||[Жетфирсторнуллобжект ()](/javascript/api/word/word.inlinepicturecollection#getfirstornullobject--)|Возвращает первый встроенный рисунок в коллекции. Возвращает нулевой объект, если коллекция пуста.|
|[List](/javascript/api/word/word.list)|[getLevelParagraphs (Level: число)](/javascript/api/word/word.list#getlevelparagraphs-level-)|Возвращает абзацы, обнаруженные на указанном уровне списка.|
||[getLevelString (Level: число)](/javascript/api/word/word.list#getlevelstring-level-)|Возвращает маркер, номер или рисунок на указанном уровне в виде строки.|
||[insertParagraph (paragraphText: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.list#insertparagraph-paragraphtext--insertlocation-)|Вставляет абзац в указанном расположении. Значение insertLocation может быть "Start", "End", "Before" или "After".|
||[id](/javascript/api/word/word.list#id)|Получает идентификатор списка.|
||[levelExistences](/javascript/api/word/word.list#levelexistences)|Проверяет наличие каждого из 9 уровней в списке. Значение true указывает, что уровень существует, то есть на этом уровне имеется по крайней мере один элемент списка. Только для чтения.|
||[levelTypes](/javascript/api/word/word.list#leveltypes)|Возвращает типы всех 9 уровней списка. Каждый тип может иметь вид "маркированный", "номер" или "Рисунок". Только для чтения.|
||[paragraphs](/javascript/api/word/word.list#paragraphs)|Возвращает абзацы в списке. Только для чтения.|
||[setLevelAlignment (Level: число, выравнивание: Word. alignment)](/javascript/api/word/word.list#setlevelalignment-level--alignment-)|Задает выравнивание маркера, номера или рисунка на указанном уровне списка.|
||[setLevelBullet (Level: число, listBullet: Word. ListBullet, charCode?: число, fontName?: строка)](/javascript/api/word/word.list#setlevelbullet-level--listbullet--charcode--fontname-)|Задает формат маркеров на указанном уровне списка. Если задан формат Custom, то параметр charCode является обязательным.|
||[setLevelIndents (Level: число, textIndent: число, Буллетнумберпиктуреиндент: число)](/javascript/api/word/word.list#setlevelindents-level--textindent--bulletnumberpictureindent-)|Задает два отступа на указанном уровне списка.|
||[setLevelNumbering (Level: число, listNumbering: Word. ListNumbering, formatString?: массив<строковый \| номер>)](/javascript/api/word/word.list#setlevelnumbering-level--listnumbering--formatstring-)|Задает формат нумерации на указанном уровне списка.|
||[setLevelStartingNumber (Level: число, startingNumber: число)](/javascript/api/word/word.list#setlevelstartingnumber-level--startingnumber-)|Задает начальный номер на указанном уровне списка. Значение по умолчанию: 1.|
|[ListCollection](/javascript/api/word/word.listcollection)|[getById(id: number)](/javascript/api/word/word.listcollection#getbyid-id-)|Возвращает список по идентификатору. Создает исключение, если список с идентификатором отсутствует в этой коллекции.|
||[getByIdOrNullObject (ID: число)](/javascript/api/word/word.listcollection#getbyidornullobject-id-)|Возвращает список по идентификатору. Возвращает нулевой объект, если список с идентификатором отсутствует в этой коллекции.|
||[getFirst()](/javascript/api/word/word.listcollection#getfirst--)|Возвращает первый список в коллекции. Вызывается, если коллекция пуста.|
||[Жетфирсторнуллобжект ()](/javascript/api/word/word.listcollection#getfirstornullobject--)|Возвращает первый список в коллекции. Возвращает нулевой объект, если коллекция пуста.|
||[getItem(index: number)](/javascript/api/word/word.listcollection#getitem-index-)|Возвращает объект списка по индексу в коллекции.|
||[items](/javascript/api/word/word.listcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[ListItem](/javascript/api/word/word.listitem)|[тип-предок (parentOnly?: Boolean)](/javascript/api/word/word.listitem#getancestor-parentonly-)|Возвращает родительский элемент или ближайшего предка (если родительского элемента нет) для данного элемента списка. Вызывается, если элемент списка не имеет предка.|
||[getAncestorOrNullObject (parentOnly?: Boolean)](/javascript/api/word/word.listitem#getancestorornullobject-parentonly-)|Возвращает родительский элемент или ближайшего предка (если родительского элемента нет) для данного элемента списка. Возвращает пустой объект, если элемент списка не имеет предка.|
||[дочерние элементы (directChildrenOnly?: Boolean)](/javascript/api/word/word.listitem#getdescendants-directchildrenonly-)|Возвращает всех потомков элемента списка.|
||[level](/javascript/api/word/word.listitem#level)|Возвращает или задает уровень элемента в списке.|
||[listString](/javascript/api/word/word.listitem#liststring)|Получает маркер элемента списка, число или изображение в виде строки. Только для чтения.|
||[siblingIndex](/javascript/api/word/word.listitem#siblingindex)|Возвращает порядковый номер элемента списка относительно элементов того же уровня. Только для чтения.|
|[Paragraph](/javascript/api/word/word.paragraph)|[attachToList (listId: число, Level: число)](/javascript/api/word/word.paragraph#attachtolist-listid--level-)|Позволяет присоединить абзац к существующему списку на указанном уровне. Если присоединить абзац к списку не удается или он уже является элементом списка, метод не выполняется.|
||[Детачфромлист ()](/javascript/api/word/word.paragraph#detachfromlist--)|Перемещает абзац за пределы списка (если он является элементом списка).|
||[GetNext ()](/javascript/api/word/word.paragraph#getnext--)|Возвращает следующий абзац. Вызывается, если абзац является последним.|
||[getNextOrNullObject ()](/javascript/api/word/word.paragraph#getnextornullobject--)|Возвращает следующий абзац. Возвращает нулевой объект, если абзац является последним.|
||[Previous ()](/javascript/api/word/word.paragraph#getprevious--)|Возвращает предыдущий абзац. Вызывается, если абзац первым.|
||[getPreviousOrNullObject ()](/javascript/api/word/word.paragraph#getpreviousornullobject--)|Возвращает предыдущий абзац. Возвращает нулевой объект, если абзац является первым.|
||[Тип-Range (rangeLocation?: Word. RangeLocation)](/javascript/api/word/word.paragraph#getrange-rangelocation-)|Возвращает весь абзац (либо его начальную или конечную точку) в виде диапазона.|
||[getTextRanges (endingMarks: строка [], trimSpacing?: Boolean)](/javascript/api/word/word.paragraph#gettextranges-endingmarks--trimspacing-)|Получает текстовые диапазоны в абзаце с помощью знаков препинания и/или других конечных меток.|
||[insertTable (rowCount: число, columnCount: число, insertLocation: Word. InsertLocation, Values?: строка [] [])](/javascript/api/word/word.paragraph#inserttable-rowcount--columncount--insertlocation--values-)|Вставляет таблицу с указанным количеством строк и столбцов. Возможные значения insertLocation: Before и After.|
||[isLastParagraph](/javascript/api/word/word.paragraph#islastparagraph)|Указывает, что абзац является последним в родительском тексте. Только для чтения.|
||[isListItem](/javascript/api/word/word.paragraph#islistitem)|Проверяет, является ли абзац элементом списка. Только для чтения.|
||[списка](/javascript/api/word/word.paragraph#list)|Возвращает объект List, к которому относится абзац. Вызывает исключение, если абзац не находится в списке. Только для чтения.|
||[listItem](/javascript/api/word/word.paragraph#listitem)|Возвращает объект ListItem для абзаца. Вызывается, если абзац не является частью списка. Только для чтения.|
||[листитеморнуллобжект](/javascript/api/word/word.paragraph#listitemornullobject)|Возвращает объект ListItem для абзаца. Если абзац не является частью списка, возвращается пустой объект. Только для чтения.|
||[листорнуллобжект](/javascript/api/word/word.paragraph#listornullobject)|Возвращает объект List, к которому относится абзац. Если абзац не находится в списке, возвращается пустой объект. Только для чтения.|
||[parentBody](/javascript/api/word/word.paragraph#parentbody)|Возвращает родительский текст абзаца. Только для чтения.|
||[parentContentControlOrNullObject](/javascript/api/word/word.paragraph#parentcontentcontrolornullobject)|Возвращает элемент управления содержимым, содержащий абзац. Возвращает нулевой объект, если родительский элемент управления содержимым отсутствует. Только для чтения.|
||[parentTable](/javascript/api/word/word.paragraph#parenttable)|Возвращает таблицу, содержащую абзац. Вызывается, если он не включен в таблицу. Только для чтения.|
||[parentTableCell](/javascript/api/word/word.paragraph#parenttablecell)|Возвращает ячейку таблицы, содержащую абзац. Создает исключение, если оно не находится в ячейке таблицы. Только для чтения.|
||[паренттаблецеллорнуллобжект](/javascript/api/word/word.paragraph#parenttablecellornullobject)|Возвращает ячейку таблицы, содержащую абзац. Если он находится не в ячейке таблицы, возвращается пустой объект. Только для чтения.|
||[паренттаблеорнуллобжект](/javascript/api/word/word.paragraph#parenttableornullobject)|Возвращает таблицу, содержащую абзац. Если он находится не в таблице, возвращается пустой объект. Только для чтения.|
||[tableNestingLevel](/javascript/api/word/word.paragraph#tablenestinglevel)|Возвращает уровень таблицы, содержащей абзац. Если абзац не находится в таблице, возвращается значение 0. Только для чтения.|
||[Split (разделители: String [], trimDelimiters?: Boolean, trimSpacing?: Boolean)](/javascript/api/word/word.paragraph#split-delimiters--trimdelimiters--trimspacing-)|Разделяет абзац на дочерние диапазоны с помощью разделителей.|
||[Стартневлист ()](/javascript/api/word/word.paragraph#startnewlist--)|Создает список, начинающийся с данного абзаца. Если абзац уже является элементом списка, метод не выполняется.|
||[styleBuiltIn](/javascript/api/word/word.paragraph#stylebuiltin)|Возвращает или задает имя встроенного стиля абзаца. Используйте это свойство для встроенных стилей, поддерживающих несколько языковых стандартов. Чтобы использовать пользовательские стили или локализованные имена стилей, применяйте свойство style.|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[getFirst()](/javascript/api/word/word.paragraphcollection#getfirst--)|Возвращает первый абзац в коллекции. Вызывается, если коллекция пуста.|
||[Жетфирсторнуллобжект ()](/javascript/api/word/word.paragraphcollection#getfirstornullobject--)|Возвращает первый абзац в коллекции. Возвращает нулевой объект, если коллекция пуста.|
||[-Last ()](/javascript/api/word/word.paragraphcollection#getlast--)|Возвращает последний абзац в коллекции. Вызывается, если коллекция пуста.|
||[Жетласторнуллобжект ()](/javascript/api/word/word.paragraphcollection#getlastornullobject--)|Возвращает последний абзац в коллекции. Возвращает нулевой объект, если коллекция пуста.|
|[Range](/javascript/api/word/word.range)|[compareLocationWith (Range: Word. Range)](/javascript/api/word/word.range#comparelocationwith-range-)|Сравнивает расположение данного диапазона с расположением другого диапазона.|
||[Експандто (Range: Word. Range)](/javascript/api/word/word.range#expandto-range-)|Возвращает новый диапазон, который простирается в том или ином направлении от данного диапазона и перекрывает другой диапазон. Данный диапазон не меняется. Вызывается, если два диапазона не имеют объединения.|
||[Експандтурнуллобжект (Range: Word. Range)](/javascript/api/word/word.range#expandtoornullobject-range-)|Возвращает новый диапазон, который простирается в том или ином направлении от данного диапазона и перекрывает другой диапазон. Данный диапазон не меняется. Возвращает нулевой объект, если два диапазона не имеют объединения.|
||[Жесиперлинкранжес ()](/javascript/api/word/word.range#gethyperlinkranges--)|Возвращает дочерние диапазоны гиперссылок в данном диапазоне.|
||[getNextTextRange (endingMarks: строка [], trimSpacing?: Boolean)](/javascript/api/word/word.range#getnexttextrange-endingmarks--trimspacing-)|Получает следующий диапазон текста с использованием знаков препинания и/или других конечных меток. Вызывается, если этот диапазон текста является последним.|
||[getNextTextRangeOrNullObject (endingMarks: строка [], trimSpacing?: Boolean)](/javascript/api/word/word.range#getnexttextrangeornullobject-endingmarks--trimspacing-)|Получает следующий диапазон текста с использованием знаков препинания и/или других конечных меток. Возвращает нулевой объект, если этот диапазон текста является последним.|
||[Тип-Range (rangeLocation?: Word. RangeLocation)](/javascript/api/word/word.range#getrange-rangelocation-)|Клонирует диапазон либо получает его начальную или конечную точку в виде нового диапазона.|
||[getTextRanges (endingMarks: строка [], trimSpacing?: Boolean)](/javascript/api/word/word.range#gettextranges-endingmarks--trimspacing-)|Возвращает дочерние диапазоны текста в диапазоне с помощью знаков препинания и/или других конечных меток.|
||[hyperlink](/javascript/api/word/word.range#hyperlink)|Возвращает первую гиперссылку в диапазоне или задает для него гиперссылку. При добавлении в диапазон новой гиперссылки из него удаляются все имеющиеся гиперссылки. Используйте ' # ', чтобы отделить адрес от части необязательного расположения.|
||[insertTable (rowCount: число, columnCount: число, insertLocation: Word. InsertLocation, Values?: строка [] [])](/javascript/api/word/word.range#inserttable-rowcount--columncount--insertlocation--values-)|Вставляет таблицу с указанным количеством строк и столбцов. Возможные значения insertLocation: Before и After.|
||[Интерсектвис (Range: Word. Range)](/javascript/api/word/word.range#intersectwith-range-)|Возвращает новый диапазон, представляющий собой пересечение данного диапазона с другим. Данный диапазон не меняется. Вызывается, если два диапазона не перекрываются или не являются смежными.|
||[Интерсектвисорнуллобжект (Range: Word. Range)](/javascript/api/word/word.range#intersectwithornullobject-range-)|Возвращает новый диапазон, представляющий собой пересечение данного диапазона с другим. Данный диапазон не меняется. Возвращает нулевой объект, если два диапазона не перекрываются или не являются смежными.|
||[isEmpty](/javascript/api/word/word.range#isempty)|Проверяет, является ли длина диапазона нулевой. Только для чтения.|
||[lists](/javascript/api/word/word.range#lists)|Возвращает коллекцию объектов списков в диапазоне. Только для чтения.|
||[parentBody](/javascript/api/word/word.range#parentbody)|Возвращает родительский текст диапазона. Только для чтения.|
||[parentContentControlOrNullObject](/javascript/api/word/word.range#parentcontentcontrolornullobject)|Возвращает элемент управления содержимым, содержащий диапазон. Возвращает нулевой объект, если родительский элемент управления содержимым отсутствует. Только для чтения.|
||[parentTable](/javascript/api/word/word.range#parenttable)|Возвращает таблицу, содержащую диапазон. Вызывается, если он не включен в таблицу. Только для чтения.|
||[parentTableCell](/javascript/api/word/word.range#parenttablecell)|Возвращает ячейку таблицы, содержащую диапазон. Создает исключение, если оно не находится в ячейке таблицы. Только для чтения.|
||[паренттаблецеллорнуллобжект](/javascript/api/word/word.range#parenttablecellornullobject)|Возвращает ячейку таблицы, содержащую диапазон. Если он находится не в ячейке таблицы, возвращается пустой объект. Только для чтения.|
||[паренттаблеорнуллобжект](/javascript/api/word/word.range#parenttableornullobject)|Возвращает таблицу, содержащую диапазон. Если она находится не в таблице, возвращается пустой объект. Только для чтения.|
||[Table](/javascript/api/word/word.range#tables)|Возвращает коллекцию объектов таблиц в диапазоне. Только для чтения.|
||[Split (разделители: String [], многопараграфный?: Boolean, trimDelimiters?: Boolean, trimSpacing?: Boolean)](/javascript/api/word/word.range#split-delimiters--multiparagraphs--trimdelimiters--trimspacing-)|Разделяет диапазон на дочерние диапазоны с помощью разделителей.|
||[styleBuiltIn](/javascript/api/word/word.range#stylebuiltin)|Возвращает или задает имя встроенного стиля диапазона. Используйте это свойство для встроенных стилей, поддерживающих несколько языковых стандартов. Чтобы использовать пользовательские стили или локализованные имена стилей, применяйте свойство style.|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[getFirst()](/javascript/api/word/word.rangecollection#getfirst--)|Возвращает первый диапазон в коллекции. Вызывается, если коллекция пуста.|
||[Жетфирсторнуллобжект ()](/javascript/api/word/word.rangecollection#getfirstornullobject--)|Возвращает первый диапазон в коллекции. Возвращает нулевой объект, если коллекция пуста.|
|[Section](/javascript/api/word/word.section)|[GetNext ()](/javascript/api/word/word.section#getnext--)|Возвращает следующий раздел. Вызывается, если этот раздел является последним.|
||[getNextOrNullObject ()](/javascript/api/word/word.section#getnextornullobject--)|Возвращает следующий раздел. Возвращает нулевой объект, если этот раздел является последним.|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[getFirst()](/javascript/api/word/word.sectioncollection#getfirst--)|Возвращает первый раздел в коллекции. Вызывается, если коллекция пуста.|
||[Жетфирсторнуллобжект ()](/javascript/api/word/word.sectioncollection#getfirstornullobject--)|Возвращает первый раздел в коллекции. Возвращает нулевой объект, если коллекция пуста.|
|[Table](/javascript/api/word/word.table)|[addColumns (insertLocation: Word. InsertLocation, columnCount: число, Values?: строка [] [])](/javascript/api/word/word.table#addcolumns-insertlocation--columncount--values-)|Добавляет столбцы в начале или в конце таблицы, используя первый или последний из имеющихся столбцов в качестве шаблона. Применим к однородным таблицам. Строковые значения (если они указаны) добавляются в новые строки.|
||[addRows (insertLocation: Word. InsertLocation, rowCount: число, Values?: строка [] [])](/javascript/api/word/word.table#addrows-insertlocation--rowcount--values-)|Добавляет строки в начале или в конце таблицы, используя первую или последнюю из имеющихся строк в качестве шаблона. Строковые значения (если они указаны) добавляются в новые строки.|
||[ориентации](/javascript/api/word/word.table#alignment)|Возвращает или задает выравнивание таблицы по столбцу страницы. Значение может быть "Left", "Centerd" или "Right".|
||[Аутофитвиндов ()](/javascript/api/word/word.table#autofitwindow--)|Автоматически подбирает ширину столбцов таблицы в соответствии с шириной окна.|
||[clear()](/javascript/api/word/word.table#clear--)|Очищает содержимое таблицы.|
||[delete()](/javascript/api/word/word.table#delete--)|Удаляет всю таблицу.|
||[deleteColumns (columnIndex: число, columnCount?: число)](/javascript/api/word/word.table#deletecolumns-columnindex--columncount-)|Удаляет определенные столбцы. Применим к однородным таблицам.|
||[deleteRows (rowIndex: число, rowCount?: число)](/javascript/api/word/word.table#deleterows-rowindex--rowcount-)|Удаляет определенные строки.|
||[Дистрибутеколумнс ()](/javascript/api/word/word.table#distributecolumns--)|Равномерно распределяет ширину столбцов. Применим к однородным таблицам.|
||[Граница (borderLocation: Word. BorderLocation)](/javascript/api/word/word.table#getborder-borderlocation-)|Возвращает стиль указанной границы.|
||[getCell(rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#getcell-rowindex--cellindex-)|Возвращает ячейку таблицы в указанной строке и указанном столбце. Вызывается, если указанная ячейка таблицы не существует.|
||[getCellOrNullObject (rowIndex: число, cellIndex: число)](/javascript/api/word/word.table#getcellornullobject-rowindex--cellindex-)|Возвращает ячейку таблицы в указанной строке и указанном столбце. Возвращает нулевой объект, если указанная ячейка таблицы не существует.|
||[Жетцеллпаддинг (cellPaddingLocation: Word. CellPaddingLocation)](/javascript/api/word/word.table#getcellpadding-cellpaddinglocation-)|Возвращает размер поля ячейки в точках.|
||[GetNext ()](/javascript/api/word/word.table#getnext--)|Возвращает следующую таблицу. Вызывается, если эта таблица является последней.|
||[getNextOrNullObject ()](/javascript/api/word/word.table#getnextornullobject--)|Возвращает следующую таблицу. Возвращает нулевой объект, если эта таблица является последней.|
||[Жетпараграфафтер ()](/javascript/api/word/word.table#getparagraphafter--)|Возвращает абзац после таблицы. Вызывается, если после таблицы нет абзаца.|
||[Жетпараграфафтерорнуллобжект ()](/javascript/api/word/word.table#getparagraphafterornullobject--)|Возвращает абзац после таблицы. Возвращает нулевой объект, если после таблицы нет абзаца.|
||[Жетпараграфбефоре ()](/javascript/api/word/word.table#getparagraphbefore--)|Возвращает абзац перед таблицей. Создает исключение, если перед таблицей нет абзаца.|
||[Жетпараграфбефореорнуллобжект ()](/javascript/api/word/word.table#getparagraphbeforeornullobject--)|Возвращает абзац перед таблицей. Возвращает нулевой объект, если перед таблицей нет абзаца.|
||[Тип-Range (rangeLocation?: Word. RangeLocation)](/javascript/api/word/word.table#getrange-rangelocation-)|Возвращает диапазон, содержащий данную таблицу, либо диапазон в начале или в конце таблицы.|
||[headerRowCount](/javascript/api/word/word.table#headerrowcount)|Возвращает и задает количество строк заголовков.|
||[horizontalAlignment](/javascript/api/word/word.table#horizontalalignment)|Возвращает и задает горизонтальное выравнивание для каждой ячейки в таблице. Допустимые значения: "Left", "Centerd", "Right" и "по ширине".|
||[insertContentControl()](/javascript/api/word/word.table#insertcontentcontrol--)|Вставляет в таблицу элемент управления содержимым.|
||[insertParagraph (paragraphText: строка, insertLocation: Word. InsertLocation)](/javascript/api/word/word.table#insertparagraph-paragraphtext--insertlocation-)|Вставляет абзац в указанном расположении. Возможные значения InsertLocation: Before и After.|
||[insertTable (rowCount: число, columnCount: число, insertLocation: Word. InsertLocation, Values?: строка [] [])](/javascript/api/word/word.table#inserttable-rowcount--columncount--insertlocation--values-)|Вставляет таблицу с указанным количеством строк и столбцов. Возможные значения InsertLocation: Before и After.|
||[font](/javascript/api/word/word.table#font)|Возвращает шрифт. Используйте эту связь, чтобы получать и задавать имя, размер, цвет и другие свойства шрифта. Только для чтения.|
||[isUniform](/javascript/api/word/word.table#isuniform)|Указывает, однородны ли все строки таблицы. Только для чтения.|
||[nestingLevel](/javascript/api/word/word.table#nestinglevel)|Возвращает уровень вложенности таблицы. Таблицам верхнего уровня соответствует значение 1. Только для чтения.|
||[parentBody](/javascript/api/word/word.table#parentbody)|Возвращает родительский текст таблицы. Только для чтения.|
||[parentContentControl](/javascript/api/word/word.table#parentcontentcontrol)|Возвращает элемент управления содержимым, содержащий таблицу. Вызывается, если родительский элемент управления содержимым отсутствует. Только для чтения.|
||[parentContentControlOrNullObject](/javascript/api/word/word.table#parentcontentcontrolornullobject)|Возвращает элемент управления содержимым, содержащий таблицу. Возвращает нулевой объект, если родительский элемент управления содержимым отсутствует. Только для чтения.|
||[parentTable](/javascript/api/word/word.table#parenttable)|Возвращает таблицу, которая содержит данную таблицу. Вызывается, если он не включен в таблицу. Только для чтения.|
||[parentTableCell](/javascript/api/word/word.table#parenttablecell)|Возвращает ячейку таблицы, содержащую данную таблицу. Создает исключение, если оно не находится в ячейке таблицы. Только для чтения.|
||[паренттаблецеллорнуллобжект](/javascript/api/word/word.table#parenttablecellornullobject)|Возвращает ячейку таблицы, содержащую данную таблицу. Если она находится не в ячейке таблицы, возвращается пустой объект. Только для чтения.|
||[паренттаблеорнуллобжект](/javascript/api/word/word.table#parenttableornullobject)|Возвращает таблицу, которая содержит данную таблицу. Если она находится не в таблице, возвращается пустой объект. Только для чтения.|
||[Стро](/javascript/api/word/word.table#rowcount)|Получает количество строк в таблице. Только для чтения.|
||[строки](/javascript/api/word/word.table#rows)|Возвращает все строки таблицы. Только для чтения.|
||[Table](/javascript/api/word/word.table#tables)|Возвращает дочерние таблицы, вложенные на один уровень ниже. Только для чтения.|
||[Search (searchText: строка, searchOptions?: Word. SearchOptions](/javascript/api/word/word.table#search-searchtext--searchoptions-)|Выполняет поиск с указанным SearchOptions в области объекта Table. Результат поиска — это коллекция объектов диапазона.|
||[SELECT (selectionMode?: Word. SelectionMode)](/javascript/api/word/word.table#select-selectionmode-)|Выбирает таблицу либо позицию в начале или в конце таблицы, а затем переходит к ней в Word.|
||[setCellPadding (cellPaddingLocation: Word. CellPaddingLocation, cellPadding: число)](/javascript/api/word/word.table#setcellpadding-cellpaddinglocation--cellpadding-)|Задает размер поля ячейки в точках.|
||[shadingColor](/javascript/api/word/word.table#shadingcolor)|Возвращает и задает цвет заливки. Цвет задается в формате "#RRGGBB" или по имени.|
||[style](/javascript/api/word/word.table#style)|Возвращает или задает имя стиля для таблицы. Используйте это свойство для пользовательских стилей и локализованных имен стилей. Чтобы использовать встроенные стили, поддерживающие несколько языковых стандартов, применяйте свойство styleBuiltIn.|
||[styleBandedColumns](/javascript/api/word/word.table#stylebandedcolumns)|Возвращает и задает значение, указывающее, есть ли в таблице чередующиеся столбцы.|
||[styleBandedRows](/javascript/api/word/word.table#stylebandedrows)|Возвращает и задает значение, указывающее, есть ли в таблице чередующиеся строки.|
||[styleBuiltIn](/javascript/api/word/word.table#stylebuiltin)|Возвращает или задает имя встроенного стиля таблицы. Используйте это свойство для встроенных стилей, поддерживающих несколько языковых стандартов. Чтобы использовать пользовательские стили или локализованные имена стилей, применяйте свойство style.|
||[styleFirstColumn](/javascript/api/word/word.table#stylefirstcolumn)|Возвращает и задает значение, указывающее, применен ли специальный стиль к первому столбцу таблицы.|
||[styleLastColumn](/javascript/api/word/word.table#stylelastcolumn)|Возвращает и задает значение, указывающее, применен ли специальный стиль к последнему столбцу таблицы.|
||[styleTotalRow](/javascript/api/word/word.table#styletotalrow)|Возвращает и задает значение, указывающее, применен ли специальный стиль к строке итогов (последней строке) таблицы.|
||[values](/javascript/api/word/word.table#values)|Возвращает и задает текстовые значения в таблице в виде двумерного массива JavaScript.|
||[verticalAlignment](/javascript/api/word/word.table#verticalalignment)|Возвращает и задает вертикальное выравнивание для каждой ячейки в таблице. Значение может быть "Top", "Center" или "Bottom".|
||[width](/javascript/api/word/word.table#width)|Возвращает и задает ширину таблицы в точках.|
|[таблебордер](/javascript/api/word/word.tableborder)|[color](/javascript/api/word/word.tableborder#color)|Получает или задает цвет границы таблицы.|
||[type](/javascript/api/word/word.tableborder#type)|Возвращает или задает тип границы таблицы.|
||[width](/javascript/api/word/word.tableborder#width)|Возвращает или задает ширину границы таблицы в точках. Не применимо к типам границ с фиксированной шириной.|
|[TableCell](/javascript/api/word/word.tablecell)|[columnWidth](/javascript/api/word/word.tablecell#columnwidth)|Возвращает и задает ширину столбца ячейки в точках. Применимо к однородным таблицам.|
||[Делетеколумн ()](/javascript/api/word/word.tablecell#deletecolumn--)|Удаляет столбец, содержащий данную ячейку. Применим к однородным таблицам.|
||[deleteRow ()](/javascript/api/word/word.tablecell#deleterow--)|Удаляет строку, содержащую данную ячейку.|
||[Граница (borderLocation: Word. BorderLocation)](/javascript/api/word/word.tablecell#getborder-borderlocation-)|Возвращает стиль указанной границы.|
||[Жетцеллпаддинг (cellPaddingLocation: Word. CellPaddingLocation)](/javascript/api/word/word.tablecell#getcellpadding-cellpaddinglocation-)|Возвращает размер поля ячейки в точках.|
||[GetNext ()](/javascript/api/word/word.tablecell#getnext--)|Возвращает следующую ячейку. Вызывается, если ячейка является последней.|
||[getNextOrNullObject ()](/javascript/api/word/word.tablecell#getnextornullobject--)|Возвращает следующую ячейку. Возвращает нулевой объект, если ячейка является последней.|
||[horizontalAlignment](/javascript/api/word/word.tablecell#horizontalalignment)|Возвращает и задает горизонтальное выравнивание ячейки. Допустимые значения: "Left", "Centerd", "Right" и "по ширине".|
||[insertColumns (insertLocation: Word. InsertLocation, columnCount: число, Values?: строка [] [])](/javascript/api/word/word.tablecell#insertcolumns-insertlocation--columncount--values-)|Добавляет столбцы слева или справа от ячейки, используя столбец этой ячейки в качестве шаблона. Применим к однородным таблицам. Строковые значения (если они указаны) добавляются в новые строки.|
||[insertRows (insertLocation: Word. InsertLocation, rowCount: число, Values?: строка [] [])](/javascript/api/word/word.tablecell#insertrows-insertlocation--rowcount--values-)|Вставляет строки над ячейкой или под ней, используя строку этой ячейки в качестве шаблона. Строковые значения (если они указаны) добавляются в новые строки.|
||[body](/javascript/api/word/word.tablecell#body)|Возвращает объект тела ячейки. Только для чтения.|
||[cellIndex](/javascript/api/word/word.tablecell#cellindex)|Получает индекс ячейки в строке. Только для чтения.|
||[parentRow](/javascript/api/word/word.tablecell#parentrow)|Получает родительскую строку ячейки. Только для чтения.|
||[parentTable](/javascript/api/word/word.tablecell#parenttable)|Возвращает родительскую таблицу ячейки. Только для чтения.|
||[rowIndex](/javascript/api/word/word.tablecell#rowindex)|Получает индекс строки ячейки в таблице. Только для чтения.|
||[width](/javascript/api/word/word.tablecell#width)|Возвращает ширину ячейки в точках. Только для чтения.|
||[setCellPadding (cellPaddingLocation: Word. CellPaddingLocation, cellPadding: число)](/javascript/api/word/word.tablecell#setcellpadding-cellpaddinglocation--cellpadding-)|Задает размер поля ячейки в точках.|
||[shadingColor](/javascript/api/word/word.tablecell#shadingcolor)|Возвращает или задает цвет заливки ячейки. Цвет задается в формате "#RRGGBB" или по имени.|
||[value](/javascript/api/word/word.tablecell#value)|Возвращает и задает текст ячейки.|
||[verticalAlignment](/javascript/api/word/word.tablecell#verticalalignment)|Возвращает и задает вертикальное выравнивание ячейки. Значение может быть "Top", "Center" или "Bottom".|
|[TableCellCollection](/javascript/api/word/word.tablecellcollection)|[getFirst()](/javascript/api/word/word.tablecellcollection#getfirst--)|Возвращает первую ячейку таблицы в коллекции. Вызывается, если коллекция пуста.|
||[Жетфирсторнуллобжект ()](/javascript/api/word/word.tablecellcollection#getfirstornullobject--)|Возвращает первую ячейку таблицы в коллекции. Возвращает нулевой объект, если коллекция пуста.|
||[items](/javascript/api/word/word.tablecellcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[TableCollection](/javascript/api/word/word.tablecollection)|[getFirst()](/javascript/api/word/word.tablecollection#getfirst--)|Возвращает первую таблицу в коллекции. Вызывается, если коллекция пуста.|
||[Жетфирсторнуллобжект ()](/javascript/api/word/word.tablecollection#getfirstornullobject--)|Возвращает первую таблицу в коллекции. Возвращает нулевой объект, если коллекция пуста.|
||[items](/javascript/api/word/word.tablecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[TableRow](/javascript/api/word/word.tablerow)|[clear()](/javascript/api/word/word.tablerow#clear--)|Очищает содержимое строки.|
||[delete()](/javascript/api/word/word.tablerow#delete--)|Удаляет всю строку.|
||[Граница (borderLocation: Word. BorderLocation)](/javascript/api/word/word.tablerow#getborder-borderlocation-)|Возвращает стиль границ ячеек в строке.|
||[Жетцеллпаддинг (cellPaddingLocation: Word. CellPaddingLocation)](/javascript/api/word/word.tablerow#getcellpadding-cellpaddinglocation-)|Возвращает размер поля ячейки в точках.|
||[GetNext ()](/javascript/api/word/word.tablerow#getnext--)|Возвращает следующую строку. Вызывается, если эта строка является последней.|
||[getNextOrNullObject ()](/javascript/api/word/word.tablerow#getnextornullobject--)|Возвращает следующую строку. Возвращает нулевой объект, если эта строка является последней.|
||[horizontalAlignment](/javascript/api/word/word.tablerow#horizontalalignment)|Возвращает и задает горизонтальное выравнивание для каждой ячейки в строке. Допустимые значения: "Left", "Centerd", "Right" и "по ширине".|
||[insertRows (insertLocation: Word. InsertLocation, rowCount: число, Values?: строка [] [])](/javascript/api/word/word.tablerow#insertrows-insertlocation--rowcount--values-)|Вставляет строки, используя данную строку в качестве шаблона. Если указаны значения, они вставляются в новые строки.|
||[preferredHeight](/javascript/api/word/word.tablerow#preferredheight)|Возвращает и задает предпочитаемую высоту строки в точках.|
||[cellCount](/javascript/api/word/word.tablerow#cellcount)|Получает количество ячеек в строке. Только для чтения.|
||[диапазона](/javascript/api/word/word.tablerow#cells)|Возвращает ячейки. Только для чтения.|
||[font](/javascript/api/word/word.tablerow#font)|Возвращает шрифт. Используйте эту связь, чтобы получать и задавать имя, размер, цвет и другие свойства шрифта. Только для чтения.|
||[isHeader](/javascript/api/word/word.tablerow#isheader)|Проверяет, является ли элемент строкой заголовков. Только для чтения. Чтобы задать количество строк заголовков, используйте свойство HeaderRowCount объекта Table.|
||[parentTable](/javascript/api/word/word.tablerow#parenttable)|Возвращает родительскую таблицу. Только для чтения.|
||[rowIndex](/javascript/api/word/word.tablerow#rowindex)|Получает индекс строки в родительской таблице. Только для чтения.|
||[Search (searchText: строка, searchOptions?: Word. SearchOptions)](/javascript/api/word/word.tablerow#search-searchtext--searchoptions-)|Выполняет поиск с указанным SearchOptions в области действия строки. Результат поиска — это коллекция объектов диапазона.|
||[SELECT (selectionMode?: Word. SelectionMode)](/javascript/api/word/word.tablerow#select-selectionmode-)|Выбирает строку и переходит к ней в Word.|
||[setCellPadding (cellPaddingLocation: Word. CellPaddingLocation, cellPadding: число)](/javascript/api/word/word.tablerow#setcellpadding-cellpaddinglocation--cellpadding-)|Задает размер поля ячейки в точках.|
||[shadingColor](/javascript/api/word/word.tablerow#shadingcolor)|Возвращает и задает цвет заливки. Цвет задается в формате "#RRGGBB" или по имени.|
||[values](/javascript/api/word/word.tablerow#values)|Возвращает и задает текстовые значения в строке в виде 2D-массива JavaScript.|
||[verticalAlignment](/javascript/api/word/word.tablerow#verticalalignment)|Возвращает и задает вертикальное выравнивание ячеек в строке. Значение может быть "Top", "Center" или "Bottom".|
|[TableRowCollection](/javascript/api/word/word.tablerowcollection)|[getFirst()](/javascript/api/word/word.tablerowcollection#getfirst--)|Возвращает первую строку в коллекции. Вызывается, если коллекция пуста.|
||[Жетфирсторнуллобжект ()](/javascript/api/word/word.tablerowcollection#getfirstornullobject--)|Возвращает первую строку в коллекции. Возвращает нулевой объект, если коллекция пуста.|
||[items](/javascript/api/word/word.tablerowcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Word](/javascript/api/word)
- [Наборы обязательных элементов API JavaScript для Word](word-api-requirement-sets.md)
