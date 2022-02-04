---
title: Набор API API Word JavaScript 1.3
description: Сведения о наборе требований WordApi 1.3.
ms.date: 03/09/2021
ms.prod: word
ms.localizationpriority: medium
---

# <a name="whats-new-in-word-javascript-api-13"></a>Новые возможности API JavaScript для Word 1.3

WordApi 1.3 добавила больше поддержки элементов управления контентом и параметров уровня документов.

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в API Word JavaScript, за набором 1.3. Чтобы просмотреть справочную документацию API для всех API, поддерживаемых требованием API Word JavaScript, установленным 1.3 или ранее, см. в справке Word API в наборе требований [1.3 или ранее](/javascript/api/word?view=word-js-1.3&preserve-view=true).

| Класс | Поля | Описание |
|:---|:---|:---|
|[Application](/javascript/api/word/word.application)|[createDocument(base64File?: string)](/javascript/api/word/word.application#word-word-application-createdocument-member(1))|Создает новый документ с помощью дополнительного файла base64, закодированного .docx.|
|[Основной текст](/javascript/api/word/word.body)|[getRange (rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.body#word-word-body-getrange-member(1))|Возвращает весь основной текст (либо его начальную или конечную точку) в виде диапазона.|
||[insertTable (rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.body#word-word-body-inserttable-member(1))|Вставляет таблицу с указанным количеством строк и столбцов.|
||[lists](/javascript/api/word/word.body#word-word-body-lists-member)|Возвращает коллекцию объектов списков в основном тексте.|
||[parentBody](/javascript/api/word/word.body#word-word-body-parentbody-member)|Возвращает родительский текст основного текста.|
||[parentBodyOrNullObject](/javascript/api/word/word.body#word-word-body-parentbodyornullobject-member)|Возвращает родительский текст основного текста.|
||[parentContentControlOrNullObject](/javascript/api/word/word.body#word-word-body-parentcontentcontrolornullobject-member)|Получает элемент управления содержимым, содержащий документ или раздел.|
||[parentSection](/javascript/api/word/word.body#word-word-body-parentsection-member)|Возвращает родительский раздел основного текста.|
||[parentSectionOrNullObject](/javascript/api/word/word.body#word-word-body-parentsectionornullobject-member)|Возвращает родительский раздел основного текста.|
||[styleBuiltIn](/javascript/api/word/word.body#word-word-body-stylebuiltin-member)|Возвращает или задает имя встроенного стиля основного текста.|
||[таблицы](/javascript/api/word/word.body#word-word-body-tables-member)|Возвращает коллекцию объектов таблиц в основном тексте.|
||[type](/javascript/api/word/word.body#word-word-body-type-member)|Возвращает тип основного текста.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[getRange (rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-getrange-member(1))|Возвращает весь элемент управления содержимым (либо его начальную или конечную точку) в виде диапазона.|
||[getTextRanges (endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-gettextranges-member(1))|Получает диапазоны текстов в области управления контентом с помощью знаков препинания и/или других знаков окончания.|
||[insertTable (rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-inserttable-member(1))|Вставляет таблицу с указанным количеством строк и столбцов в элемент управления содержимым или рядом с ним.|
||[lists](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-lists-member)|Возвращает коллекцию объектов списков в элементе управления содержимым.|
||[parentBody](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parentbody-member)|Возвращает родительский текст элемента управления содержимым.|
||[parentContentControlOrNullObject](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parentcontentcontrolornullobject-member)|Получает элемент управления содержимым, содержащий элемент управления содержимым.|
||[parentTable](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parenttable-member)|Возвращает таблицу, содержащую элемент управления содержимым.|
||[parentTableCell](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parenttablecell-member)|Возвращает ячейку таблицы, содержащую элемент управления содержимым.|
||[parentTableCellOrNullObject](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parenttablecellornullobject-member)|Возвращает ячейку таблицы, содержащую элемент управления содержимым.|
||[parentTableOrNullObject](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-parenttableornullobject-member)|Возвращает таблицу, содержащую элемент управления содержимым.|
||[split(delimiters: string[], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-split-member(1))|Разделяет элемент управления содержимым на дочерние диапазоны с помощью разделителей.|
||[styleBuiltIn](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-stylebuiltin-member)|Возвращает или задает имя встроенного стиля для элемента управления содержимым.|
||[подтип](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-subtype-member)|Возвращает подтип элемента управления содержимым.|
||[таблицы](/javascript/api/word/word.contentcontrol#word-word-contentcontrol-tables-member)|Возвращает коллекцию объектов таблиц в элементе управления содержимым.|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getByIdOrNullObject (id: number)](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getbyidornullobject-member(1))|Возвращает элемент управления содержимым по его идентификатору.|
||[getByTypes (типы: Word.ContentControlType[])](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getbytypes-member(1))|Получает элементы управления контентом, которые имеют указанные типы и/или подтипы.|
||[getFirst()](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getfirst-member(1))|Возвращает первый элемент управления содержимым в коллекции.|
||[getFirstOrNullObject()](/javascript/api/word/word.contentcontrolcollection#word-word-contentcontrolcollection-getfirstornullobject-member(1))|Возвращает первый элемент управления содержимым в коллекции.|
|[CustomProperty](/javascript/api/word/word.customproperty)|[delete()](/javascript/api/word/word.customproperty#word-word-customproperty-delete-member(1))|Удаляет настраиваемое свойство.|
||[key](/javascript/api/word/word.customproperty#word-word-customproperty-key-member)|Возвращает ключ настраиваемого свойства.|
||[type](/javascript/api/word/word.customproperty#word-word-customproperty-type-member)|Получает тип значения настраиваемого свойства.|
||[value](/javascript/api/word/word.customproperty#word-word-customproperty-value-member)|Получает или задает значение настраиваемого свойства.|
|[CustomPropertyCollection](/javascript/api/word/word.custompropertycollection)|[add(key: string, value: any)](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-add-member(1))|Создает или задает настраиваемое свойство.|
||[deleteAll()](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-deleteall-member(1))|Удаляет все настраиваемые свойства в коллекции.|
||[getCount()](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-getcount-member(1))|Получает количество настраиваемых свойств.|
||[getItem(key: string)](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-getitem-member(1))|Возвращает объект настраиваемого свойства по ключу, указываемому без учета регистра.|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-getitemornullobject-member(1))|Возвращает объект настраиваемого свойства по ключу, указываемому без учета регистра.|
||[items](/javascript/api/word/word.custompropertycollection#word-word-custompropertycollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[Document](/javascript/api/word/word.document)|[properties](/javascript/api/word/word.document#word-word-document-properties-member)|Получает свойства документа.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[body](/javascript/api/word/word.documentcreated#word-word-documentcreated-body-member)|Получает объект тела документа.|
||[contentControls](/javascript/api/word/word.documentcreated#word-word-documentcreated-contentcontrols-member)|Получает коллекцию объектов управления контентом в документе.|
||[open()](/javascript/api/word/word.documentcreated#word-word-documentcreated-open-member(1))|Открывает документ.|
||[properties](/javascript/api/word/word.documentcreated#word-word-documentcreated-properties-member)|Получает свойства документа.|
||[save()](/javascript/api/word/word.documentcreated#word-word-documentcreated-save-member(1))|Сохраняет документ.|
||[сохранено](/javascript/api/word/word.documentcreated#word-word-documentcreated-saved-member)|Указывает, сохранены ли изменения, внесенные в документ.|
||[sections](/javascript/api/word/word.documentcreated#word-word-documentcreated-sections-member)|Получает коллекцию объектов раздела в документе.|
|[DocumentProperties](/javascript/api/word/word.documentproperties)|[applicationName](/javascript/api/word/word.documentproperties#word-word-documentproperties-applicationname-member)|Возвращает имя приложения для документа.|
||[автор](/javascript/api/word/word.documentproperties#word-word-documentproperties-author-member)|Возвращает или задает автора документа.|
||[категория](/javascript/api/word/word.documentproperties#word-word-documentproperties-category-member)|Возвращает или задает категорию документа.|
||[comments](/javascript/api/word/word.documentproperties#word-word-documentproperties-comments-member)|Возвращает или задает примечания к документу.|
||[company](/javascript/api/word/word.documentproperties#word-word-documentproperties-company-member)|Возвращает или задает компанию документа.|
||[creationDate](/javascript/api/word/word.documentproperties#word-word-documentproperties-creationdate-member)|Возвращает дату создания документа.|
||[customProperties](/javascript/api/word/word.documentproperties#word-word-documentproperties-customproperties-member)|Возвращает коллекцию настраиваемых свойств документа.|
||[format](/javascript/api/word/word.documentproperties#word-word-documentproperties-format-member)|Возвращает или задает формат документа.|
||[ключевые слова](/javascript/api/word/word.documentproperties#word-word-documentproperties-keywords-member)|Возвращает или задает ключевые слова документа.|
||[lastAuthor](/javascript/api/word/word.documentproperties#word-word-documentproperties-lastauthor-member)|Получает последнего автора документа.|
||[lastPrintDate](/javascript/api/word/word.documentproperties#word-word-documentproperties-lastprintdate-member)|Возвращает дату последней печати документа.|
||[lastSaveTime](/javascript/api/word/word.documentproperties#word-word-documentproperties-lastsavetime-member)|Возвращает время последнего сохранения документа.|
||[manager](/javascript/api/word/word.documentproperties#word-word-documentproperties-manager-member)|Возвращает или задает менеджера документа.|
||[revisionNumber](/javascript/api/word/word.documentproperties#word-word-documentproperties-revisionnumber-member)|Возвращает номер редакции документа.|
||[безопасность](/javascript/api/word/word.documentproperties#word-word-documentproperties-security-member)|Получает параметры безопасности документа.|
||[subject](/javascript/api/word/word.documentproperties#word-word-documentproperties-subject-member)|Возвращает или задает тему документа.|
||[template](/javascript/api/word/word.documentproperties#word-word-documentproperties-template-member)|Возвращает шаблон документа.|
||[заголовок](/javascript/api/word/word.documentproperties#word-word-documentproperties-title-member)|Возвращает или задает название документа.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[getNext()](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-getnext-member(1))|Возвращает следующий встроенный рисунок.|
||[getNextOrNullObject()](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-getnextornullobject-member(1))|Возвращает следующий встроенный рисунок.|
||[getRange (rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-getrange-member(1))|Возвращает рисунок (либо его начальную или конечную точку) в виде диапазона.|
||[parentContentControlOrNullObject](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parentcontentcontrolornullobject-member)|Возвращает элемент управления содержимым, который содержит встроенный рисунок.|
||[parentTable](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parenttable-member)|Возвращает таблицу, содержащую встроенный рисунок.|
||[parentTableCell](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parenttablecell-member)|Возвращает ячейку таблицы, содержащую встроенный рисунок.|
||[parentTableCellOrNullObject](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parenttablecellornullobject-member)|Возвращает ячейку таблицы, содержащую встроенный рисунок.|
||[parentTableOrNullObject](/javascript/api/word/word.inlinepicture#word-word-inlinepicture-parenttableornullobject-member)|Возвращает таблицу, содержащую встроенный рисунок.|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[getFirst()](/javascript/api/word/word.inlinepicturecollection#word-word-inlinepicturecollection-getfirst-member(1))|Возвращает первый встроенный рисунок в коллекции.|
||[getFirstOrNullObject()](/javascript/api/word/word.inlinepicturecollection#word-word-inlinepicturecollection-getfirstornullobject-member(1))|Возвращает первый встроенный рисунок в коллекции.|
|[Список](/javascript/api/word/word.list)|[getLevelParagraphs(level: number)](/javascript/api/word/word.list#word-word-list-getlevelparagraphs-member(1))|Возвращает абзацы, обнаруженные на указанном уровне списка.|
||[getLevelString (уровень: номер)](/javascript/api/word/word.list#word-word-list-getlevelstring-member(1))|Получает пулю, номер или изображение на указанном уровне в качестве строки.|
||[id](/javascript/api/word/word.list#word-word-list-id-member)|Получает id списка.|
||[insertParagraph (paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.list#word-word-list-insertparagraph-member(1))|Вставляет абзац в указанном расположении.|
||[levelExistences](/javascript/api/word/word.list#word-word-list-levelexistences-member)|Проверяет наличие каждого из 9 уровней в списке.|
||[levelTypes](/javascript/api/word/word.list#word-word-list-leveltypes-member)|Возвращает типы всех 9 уровней списка.|
||[paragraphs](/javascript/api/word/word.list#word-word-list-paragraphs-member)|Возвращает абзацы в списке.|
||[setLevelAlignment (уровень: номер, выравнивание: Word.Alignment)](/javascript/api/word/word.list#word-word-list-setlevelalignment-member(1))|Задает выравнивание пули, номера или изображения на указанном уровне в списке.|
||[setLevelBullet(level: number, listBullet: Word.ListBullet, charCode?: number, fontName?: string)](/javascript/api/word/word.list#word-word-list-setlevelbullet-member(1))|Задает формат маркеров на указанном уровне списка.|
||[setLevelIndents(level: number, textIndent: number, bulletNumberPictureIndent: number)](/javascript/api/word/word.list#word-word-list-setlevelindents-member(1))|Задает два отступа на указанном уровне списка.|
||[setLevelNumbering(level: number, listNumbering: Word.ListNumbering, formatString?: Array<string \| number>)](/javascript/api/word/word.list#word-word-list-setlevelnumbering-member(1))|Задает формат нумерации на указанном уровне списка.|
||[setLevelStartingNumber (уровень: номер, startingNumber: number)](/javascript/api/word/word.list#word-word-list-setlevelstartingnumber-member(1))|Задает начальный номер на указанном уровне списка.|
|[ListCollection](/javascript/api/word/word.listcollection)|[getById(id: number)](/javascript/api/word/word.listcollection#word-word-listcollection-getbyid-member(1))|Возвращает список по идентификатору.|
||[getByIdOrNullObject (id: number)](/javascript/api/word/word.listcollection#word-word-listcollection-getbyidornullobject-member(1))|Возвращает список по идентификатору.|
||[getFirst()](/javascript/api/word/word.listcollection#word-word-listcollection-getfirst-member(1))|Возвращает первый список в коллекции.|
||[getFirstOrNullObject()](/javascript/api/word/word.listcollection#word-word-listcollection-getfirstornullobject-member(1))|Возвращает первый список в коллекции.|
||[getItem(index: number)](/javascript/api/word/word.listcollection#word-word-listcollection-getitem-member(1))|Возвращает объект списка по индексу в коллекции.|
||[items](/javascript/api/word/word.listcollection#word-word-listcollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[ListItem](/javascript/api/word/word.listitem)|[getAncestor(parentOnly?: boolean)](/javascript/api/word/word.listitem#word-word-listitem-getancestor-member(1))|Возвращает родительский элемент или ближайшего предка (если родительского элемента нет) для данного элемента списка.|
||[getAncestorOrNullObject(parentOnly?: boolean)](/javascript/api/word/word.listitem#word-word-listitem-getancestorornullobject-member(1))|Возвращает родительский элемент или ближайшего предка (если родительского элемента нет) для данного элемента списка.|
||[getDescendants (directChildrenOnly?: boolean)](/javascript/api/word/word.listitem#word-word-listitem-getdescendants-member(1))|Возвращает всех потомков элемента списка.|
||[level](/javascript/api/word/word.listitem#word-word-listitem-level-member)|Возвращает или задает уровень элемента в списке.|
||[listString](/javascript/api/word/word.listitem#word-word-listitem-liststring-member)|Получает пулю элемента списка, номер или изображение в качестве строки.|
||[siblingIndex](/javascript/api/word/word.listitem#word-word-listitem-siblingindex-member)|Возвращает порядковый номер элемента списка относительно элементов того же уровня.|
|[Paragraph](/javascript/api/word/word.paragraph)|[attachToList(listId: number, level: number)](/javascript/api/word/word.paragraph#word-word-paragraph-attachtolist-member(1))|Позволяет присоединить абзац к существующему списку на указанном уровне.|
||[detachFromList()](/javascript/api/word/word.paragraph#word-word-paragraph-detachfromlist-member(1))|Перемещает абзац за пределы списка (если он является элементом списка).|
||[getNext()](/javascript/api/word/word.paragraph#word-word-paragraph-getnext-member(1))|Возвращает следующий абзац.|
||[getNextOrNullObject()](/javascript/api/word/word.paragraph#word-word-paragraph-getnextornullobject-member(1))|Возвращает следующий абзац.|
||[getPrevious()](/javascript/api/word/word.paragraph#word-word-paragraph-getprevious-member(1))|Возвращает предыдущий абзац.|
||[getPreviousOrNullObject()](/javascript/api/word/word.paragraph#word-word-paragraph-getpreviousornullobject-member(1))|Возвращает предыдущий абзац.|
||[getRange (rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.paragraph#word-word-paragraph-getrange-member(1))|Возвращает весь абзац (либо его начальную или конечную точку) в виде диапазона.|
||[getTextRanges (endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.paragraph#word-word-paragraph-gettextranges-member(1))|Получает диапазоны текста в абзаце, используя знаки препинания и/или другие знаки окончания.|
||[insertTable (rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.paragraph#word-word-paragraph-inserttable-member(1))|Вставляет таблицу с указанным количеством строк и столбцов.|
||[isLastParagraph](/javascript/api/word/word.paragraph#word-word-paragraph-islastparagraph-member)|Указывает, что абзац является последним в родительском тексте.|
||[isListItem](/javascript/api/word/word.paragraph#word-word-paragraph-islistitem-member)|Проверяет, является ли абзац элементом списка.|
||[list](/javascript/api/word/word.paragraph#word-word-paragraph-list-member)|Возвращает объект List, к которому относится абзац.|
||[listItem](/javascript/api/word/word.paragraph#word-word-paragraph-listitem-member)|Возвращает объект ListItem для абзаца.|
||[listItemOrNullObject](/javascript/api/word/word.paragraph#word-word-paragraph-listitemornullobject-member)|Возвращает объект ListItem для абзаца.|
||[listOrNullObject](/javascript/api/word/word.paragraph#word-word-paragraph-listornullobject-member)|Возвращает объект List, к которому относится абзац.|
||[parentBody](/javascript/api/word/word.paragraph#word-word-paragraph-parentbody-member)|Возвращает родительский текст абзаца.|
||[parentContentControlOrNullObject](/javascript/api/word/word.paragraph#word-word-paragraph-parentcontentcontrolornullobject-member)|Возвращает элемент управления содержимым, содержащий абзац.|
||[parentTable](/javascript/api/word/word.paragraph#word-word-paragraph-parenttable-member)|Возвращает таблицу, содержащую абзац.|
||[parentTableCell](/javascript/api/word/word.paragraph#word-word-paragraph-parenttablecell-member)|Возвращает ячейку таблицы, содержащую абзац.|
||[parentTableCellOrNullObject](/javascript/api/word/word.paragraph#word-word-paragraph-parenttablecellornullobject-member)|Возвращает ячейку таблицы, содержащую абзац.|
||[parentTableOrNullObject](/javascript/api/word/word.paragraph#word-word-paragraph-parenttableornullobject-member)|Возвращает таблицу, содержащую абзац.|
||[split(delimiters: string[], trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.paragraph#word-word-paragraph-split-member(1))|Разделяет абзац на дочерние диапазоны с помощью разделителей.|
||[startNewList()](/javascript/api/word/word.paragraph#word-word-paragraph-startnewlist-member(1))|Создает список, начинающийся с данного абзаца.|
||[styleBuiltIn](/javascript/api/word/word.paragraph#word-word-paragraph-stylebuiltin-member)|Возвращает или задает имя встроенного стиля абзаца.|
||[tableNestingLevel](/javascript/api/word/word.paragraph#word-word-paragraph-tablenestinglevel-member)|Возвращает уровень таблицы, содержащей абзац.|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[getFirst()](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-getfirst-member(1))|Возвращает первый абзац в коллекции.|
||[getFirstOrNullObject()](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-getfirstornullobject-member(1))|Возвращает первый абзац в коллекции.|
||[getLast()](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-getlast-member(1))|Возвращает последний абзац в коллекции.|
||[getLastOrNullObject()](/javascript/api/word/word.paragraphcollection#word-word-paragraphcollection-getlastornullobject-member(1))|Возвращает последний абзац в коллекции.|
|[Range](/javascript/api/word/word.range)|[compareLocationWith (диапазон: Word.Range)](/javascript/api/word/word.range#word-word-range-comparelocationwith-member(1))|Сравнивает расположение данного диапазона с расположением другого диапазона.|
||[expandTo (диапазон: Word.Range)](/javascript/api/word/word.range#word-word-range-expandto-member(1))|Возвращает новый диапазон, который простирается в том или ином направлении от данного диапазона и перекрывает другой диапазон.|
||[expandToOrNullObject (диапазон: Word.Range)](/javascript/api/word/word.range#word-word-range-expandtoornullobject-member(1))|Возвращает новый диапазон, который простирается в том или ином направлении от данного диапазона и перекрывает другой диапазон.|
||[getHyperlinkRanges()](/javascript/api/word/word.range#word-word-range-gethyperlinkranges-member(1))|Возвращает дочерние диапазоны гиперссылок в данном диапазоне.|
||[getNextTextRange (endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#word-word-range-getnexttextrange-member(1))|Получает следующий диапазон текста, используя знаки препинания и/или другие знаки окончания.|
||[getNextTextRangeOrNullObject (endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#word-word-range-getnexttextrangeornullobject-member(1))|Получает следующий диапазон текста, используя знаки препинания и/или другие знаки окончания.|
||[getRange (rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.range#word-word-range-getrange-member(1))|Клонирует диапазон либо получает его начальную или конечную точку в виде нового диапазона.|
||[getTextRanges (endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#word-word-range-gettextranges-member(1))|Получает текстовые детские диапазоны в диапазоне, используя знаки препинания и/или другие знаки окончания.|
||[hyperlink](/javascript/api/word/word.range#word-word-range-hyperlink-member)|Возвращает первую гиперссылку в диапазоне или задает для него гиперссылку.|
||[insertTable (rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.range#word-word-range-inserttable-member(1))|Вставляет таблицу с указанным количеством строк и столбцов.|
||[intersectWith (диапазон: Word.Range)](/javascript/api/word/word.range#word-word-range-intersectwith-member(1))|Возвращает новый диапазон, представляющий собой пересечение данного диапазона с другим.|
||[intersectWithOrNullObject (диапазон: Word.Range)](/javascript/api/word/word.range#word-word-range-intersectwithornullobject-member(1))|Возвращает новый диапазон, представляющий собой пересечение данного диапазона с другим.|
||[isEmpty](/javascript/api/word/word.range#word-word-range-isempty-member)|Проверяет, является ли длина диапазона нулевой.|
||[lists](/javascript/api/word/word.range#word-word-range-lists-member)|Возвращает коллекцию объектов списков в диапазоне.|
||[parentBody](/javascript/api/word/word.range#word-word-range-parentbody-member)|Возвращает родительский текст диапазона.|
||[parentContentControlOrNullObject](/javascript/api/word/word.range#word-word-range-parentcontentcontrolornullobject-member)|Возвращает элемент управления содержимым, содержащий диапазон.|
||[parentTable](/javascript/api/word/word.range#word-word-range-parenttable-member)|Возвращает таблицу, содержащую диапазон.|
||[parentTableCell](/javascript/api/word/word.range#word-word-range-parenttablecell-member)|Возвращает ячейку таблицы, содержащую диапазон.|
||[parentTableCellOrNullObject](/javascript/api/word/word.range#word-word-range-parenttablecellornullobject-member)|Возвращает ячейку таблицы, содержащую диапазон.|
||[parentTableOrNullObject](/javascript/api/word/word.range#word-word-range-parenttableornullobject-member)|Возвращает таблицу, содержащую диапазон.|
||[split(delimiters: string[], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.range#word-word-range-split-member(1))|Разделяет диапазон на дочерние диапазоны с помощью разделителей.|
||[styleBuiltIn](/javascript/api/word/word.range#word-word-range-stylebuiltin-member)|Возвращает или задает имя встроенного стиля диапазона.|
||[таблицы](/javascript/api/word/word.range#word-word-range-tables-member)|Возвращает коллекцию объектов таблиц в диапазоне.|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[getFirst()](/javascript/api/word/word.rangecollection#word-word-rangecollection-getfirst-member(1))|Возвращает первый диапазон в коллекции.|
||[getFirstOrNullObject()](/javascript/api/word/word.rangecollection#word-word-rangecollection-getfirstornullobject-member(1))|Возвращает первый диапазон в коллекции.|
|[RequestContext](/javascript/api/word/word.requestcontext)|[application](/javascript/api/word/word.requestcontext#word-word-requestcontext-application-member)|[Набор API: WordApi 1.3] *|
|[Section](/javascript/api/word/word.section)|[getNext()](/javascript/api/word/word.section#word-word-section-getnext-member(1))|Возвращает следующий раздел.|
||[getNextOrNullObject()](/javascript/api/word/word.section#word-word-section-getnextornullobject-member(1))|Возвращает следующий раздел.|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[getFirst()](/javascript/api/word/word.sectioncollection#word-word-sectioncollection-getfirst-member(1))|Возвращает первый раздел в коллекции.|
||[getFirstOrNullObject()](/javascript/api/word/word.sectioncollection#word-word-sectioncollection-getfirstornullobject-member(1))|Возвращает первый раздел в коллекции.|
|[Table](/javascript/api/word/word.table)|[addColumns(insertLocation: Word.InsertLocation, columnCount: number, values?: string[][])](/javascript/api/word/word.table#word-word-table-addcolumns-member(1))|Добавляет столбцы в начале или в конце таблицы, используя первый или последний из имеющихся столбцов в качестве шаблона.|
||[addRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.table#word-word-table-addrows-member(1))|Добавляет строки в начале или в конце таблицы, используя первую или последнюю из имеющихся строк в качестве шаблона.|
||[выравнивание](/javascript/api/word/word.table#word-word-table-alignment-member)|Получает или задает выравнивание таблицы со столбцом страницы.|
||[autoFitWindow()](/javascript/api/word/word.table#word-word-table-autofitwindow-member(1))|Автоматически подбирает ширину столбцов таблицы в соответствии с шириной окна.|
||[clear()](/javascript/api/word/word.table#word-word-table-clear-member(1))|Очищает содержимое таблицы.|
||[delete()](/javascript/api/word/word.table#word-word-table-delete-member(1))|Удаляет всю таблицу.|
||[deleteColumns(columnIndex: number, columnCount?: number)](/javascript/api/word/word.table#word-word-table-deletecolumns-member(1))|Удаляет определенные столбцы.|
||[deleteRows(rowIndex: number, rowCount?: number)](/javascript/api/word/word.table#word-word-table-deleterows-member(1))|Удаляет определенные строки.|
||[distributeColumns()](/javascript/api/word/word.table#word-word-table-distributecolumns-member(1))|Равномерно распределяет ширину столбцов.|
||[font](/javascript/api/word/word.table#word-word-table-font-member)|Возвращает шрифт.|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.table#word-word-table-getborder-member(1))|Возвращает стиль указанной границы.|
||[getCell(rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#word-word-table-getcell-member(1))|Возвращает ячейку таблицы в указанной строке и указанном столбце.|
||[getCellOrNullObject (rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#word-word-table-getcellornullobject-member(1))|Возвращает ячейку таблицы в указанной строке и указанном столбце.|
||[getCellPadding (cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.table#word-word-table-getcellpadding-member(1))|Возвращает размер поля ячейки в точках.|
||[getNext()](/javascript/api/word/word.table#word-word-table-getnext-member(1))|Возвращает следующую таблицу.|
||[getNextOrNullObject()](/javascript/api/word/word.table#word-word-table-getnextornullobject-member(1))|Возвращает следующую таблицу.|
||[getParagraphAfter()](/javascript/api/word/word.table#word-word-table-getparagraphafter-member(1))|Возвращает абзац после таблицы.|
||[getParagraphAfterOrNullObject()](/javascript/api/word/word.table#word-word-table-getparagraphafterornullobject-member(1))|Возвращает абзац после таблицы.|
||[getParagraphBefore()](/javascript/api/word/word.table#word-word-table-getparagraphbefore-member(1))|Возвращает абзац перед таблицей.|
||[getParagraphBeforeOrNullObject()](/javascript/api/word/word.table#word-word-table-getparagraphbeforeornullobject-member(1))|Возвращает абзац перед таблицей.|
||[getRange (rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.table#word-word-table-getrange-member(1))|Возвращает диапазон, содержащий данную таблицу, либо диапазон в начале или в конце таблицы.|
||[headerRowCount](/javascript/api/word/word.table#word-word-table-headerrowcount-member)|Возвращает и задает количество строк заголовков.|
||[horizontalAlignment](/javascript/api/word/word.table#word-word-table-horizontalalignment-member)|Возвращает и задает горизонтальное выравнивание для каждой ячейки в таблице.|
||[ignorePunct](/javascript/api/word/word.table#word-word-table-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.table#word-word-table-ignorespace-member)||
||[insertContentControl()](/javascript/api/word/word.table#word-word-table-insertcontentcontrol-member(1))|Вставляет в таблицу элемент управления содержимым.|
||[insertParagraph (paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.table#word-word-table-insertparagraph-member(1))|Вставляет абзац в указанном расположении.|
||[insertTable (rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.table#word-word-table-inserttable-member(1))|Вставляет таблицу с указанным количеством строк и столбцов.|
||[isUniform](/javascript/api/word/word.table#word-word-table-isuniform-member)|Указывает, однородны ли все строки таблицы.|
||[matchCase](/javascript/api/word/word.table#word-word-table-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.table#word-word-table-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.table#word-word-table-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.table#word-word-table-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.table#word-word-table-matchwildcards-member)||
||[nestingLevel](/javascript/api/word/word.table#word-word-table-nestinglevel-member)|Возвращает уровень вложенности таблицы.|
||[parentBody](/javascript/api/word/word.table#word-word-table-parentbody-member)|Возвращает родительский текст таблицы.|
||[parentContentControl](/javascript/api/word/word.table#word-word-table-parentcontentcontrol-member)|Возвращает элемент управления содержимым, содержащий таблицу.|
||[parentContentControlOrNullObject](/javascript/api/word/word.table#word-word-table-parentcontentcontrolornullobject-member)|Возвращает элемент управления содержимым, содержащий таблицу.|
||[parentTable](/javascript/api/word/word.table#word-word-table-parenttable-member)|Возвращает таблицу, которая содержит данную таблицу.|
||[parentTableCell](/javascript/api/word/word.table#word-word-table-parenttablecell-member)|Возвращает ячейку таблицы, содержащую данную таблицу.|
||[parentTableCellOrNullObject](/javascript/api/word/word.table#word-word-table-parenttablecellornullobject-member)|Возвращает ячейку таблицы, содержащую данную таблицу.|
||[parentTableOrNullObject](/javascript/api/word/word.table#word-word-table-parenttableornullobject-member)|Возвращает таблицу, которая содержит данную таблицу.|
||[rowCount](/javascript/api/word/word.table#word-word-table-rowcount-member)|Получает количество строк в таблице.|
||[строки](/javascript/api/word/word.table#word-word-table-rows-member)|Возвращает все строки таблицы.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.table#word-word-table-search-member(1))|Выполняет поиск с указанными SearchOptions в области объекта таблицы.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.table#word-word-table-select-member(1))|Выбирает таблицу либо позицию в начале или в конце таблицы, а затем переходит к ней в Word.|
||[setCellPadding (cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.table#word-word-table-setcellpadding-member(1))|Задает размер поля ячейки в точках.|
||[shadingColor](/javascript/api/word/word.table#word-word-table-shadingcolor-member)|Возвращает и задает цвет заливки.|
||[style](/javascript/api/word/word.table#word-word-table-style-member)|Возвращает или задает имя стиля для таблицы.|
||[styleBandedColumns](/javascript/api/word/word.table#word-word-table-stylebandedcolumns-member)|Возвращает и задает значение, указывающее, есть ли в таблице чередующиеся столбцы.|
||[styleBandedRows](/javascript/api/word/word.table#word-word-table-stylebandedrows-member)|Возвращает и задает значение, указывающее, есть ли в таблице чередующиеся строки.|
||[styleBuiltIn](/javascript/api/word/word.table#word-word-table-stylebuiltin-member)|Возвращает или задает имя встроенного стиля таблицы.|
||[styleFirstColumn](/javascript/api/word/word.table#word-word-table-stylefirstcolumn-member)|Возвращает и задает значение, указывающее, применен ли специальный стиль к первому столбцу таблицы.|
||[styleLastColumn](/javascript/api/word/word.table#word-word-table-stylelastcolumn-member)|Возвращает и задает значение, указывающее, применен ли специальный стиль к последнему столбцу таблицы.|
||[styleTotalRow](/javascript/api/word/word.table#word-word-table-styletotalrow-member)|Возвращает и задает значение, указывающее, применен ли специальный стиль к строке итогов (последней строке) таблицы.|
||[таблицы](/javascript/api/word/word.table#word-word-table-tables-member)|Возвращает дочерние таблицы, вложенные на один уровень ниже.|
||[values](/javascript/api/word/word.table#word-word-table-values-member)|Возвращает и задает текстовые значения в таблице в виде двумерного массива JavaScript.|
||[verticalAlignment](/javascript/api/word/word.table#word-word-table-verticalalignment-member)|Возвращает и задает вертикальное выравнивание для каждой ячейки в таблице.|
||[width](/javascript/api/word/word.table#word-word-table-width-member)|Возвращает и задает ширину таблицы в точках.|
|[TableBorder](/javascript/api/word/word.tableborder)|[color](/javascript/api/word/word.tableborder#word-word-tableborder-color-member)|Получает или задает цвет границы таблицы.|
||[type](/javascript/api/word/word.tableborder#word-word-tableborder-type-member)|Возвращает или задает тип границы таблицы.|
||[width](/javascript/api/word/word.tableborder#word-word-tableborder-width-member)|Возвращает или задает ширину границы таблицы в точках.|
|[TableCell](/javascript/api/word/word.tablecell)|[body](/javascript/api/word/word.tablecell#word-word-tablecell-body-member)|Возвращает объект тела ячейки.|
||[cellIndex](/javascript/api/word/word.tablecell#word-word-tablecell-cellindex-member)|Получает индекс ячейки в строке.|
||[columnWidth](/javascript/api/word/word.tablecell#word-word-tablecell-columnwidth-member)|Возвращает и задает ширину столбца ячейки в точках.|
||[deleteColumn()](/javascript/api/word/word.tablecell#word-word-tablecell-deletecolumn-member(1))|Удаляет столбец, содержащий данную ячейку.|
||[deleteRow()](/javascript/api/word/word.tablecell#word-word-tablecell-deleterow-member(1))|Удаляет строку, содержащую данную ячейку.|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.tablecell#word-word-tablecell-getborder-member(1))|Возвращает стиль указанной границы.|
||[getCellPadding (cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.tablecell#word-word-tablecell-getcellpadding-member(1))|Возвращает размер поля ячейки в точках.|
||[getNext()](/javascript/api/word/word.tablecell#word-word-tablecell-getnext-member(1))|Возвращает следующую ячейку.|
||[getNextOrNullObject()](/javascript/api/word/word.tablecell#word-word-tablecell-getnextornullobject-member(1))|Возвращает следующую ячейку.|
||[horizontalAlignment](/javascript/api/word/word.tablecell#word-word-tablecell-horizontalalignment-member)|Возвращает и задает горизонтальное выравнивание ячейки.|
||[insertColumns(insertLocation: Word.InsertLocation, columnCount: number, values?: string[][])](/javascript/api/word/word.tablecell#word-word-tablecell-insertcolumns-member(1))|Добавляет столбцы слева или справа от ячейки, используя столбец этой ячейки в качестве шаблона.|
||[insertRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[]])](/javascript/api/word/word.tablecell#word-word-tablecell-insertrows-member(1))|Вставляет строки над ячейкой или под ней, используя строку этой ячейки в качестве шаблона.|
||[parentRow](/javascript/api/word/word.tablecell#word-word-tablecell-parentrow-member)|Получает родительскую строку ячейки.|
||[parentTable](/javascript/api/word/word.tablecell#word-word-tablecell-parenttable-member)|Возвращает родительскую таблицу ячейки.|
||[rowIndex](/javascript/api/word/word.tablecell#word-word-tablecell-rowindex-member)|Получает индекс строки ячейки в таблице.|
||[setCellPadding (cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.tablecell#word-word-tablecell-setcellpadding-member(1))|Задает размер поля ячейки в точках.|
||[shadingColor](/javascript/api/word/word.tablecell#word-word-tablecell-shadingcolor-member)|Возвращает или задает цвет заливки ячейки.|
||[value](/javascript/api/word/word.tablecell#word-word-tablecell-value-member)|Возвращает и задает текст ячейки.|
||[verticalAlignment](/javascript/api/word/word.tablecell#word-word-tablecell-verticalalignment-member)|Возвращает и задает вертикальное выравнивание ячейки.|
||[width](/javascript/api/word/word.tablecell#word-word-tablecell-width-member)|Возвращает ширину ячейки в точках.|
|[TableCellCollection](/javascript/api/word/word.tablecellcollection)|[getFirst()](/javascript/api/word/word.tablecellcollection#word-word-tablecellcollection-getfirst-member(1))|Возвращает первую ячейку таблицы в коллекции.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablecellcollection#word-word-tablecellcollection-getfirstornullobject-member(1))|Возвращает первую ячейку таблицы в коллекции.|
||[items](/javascript/api/word/word.tablecellcollection#word-word-tablecellcollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[TableCollection](/javascript/api/word/word.tablecollection)|[getFirst()](/javascript/api/word/word.tablecollection#word-word-tablecollection-getfirst-member(1))|Возвращает первую таблицу в коллекции.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablecollection#word-word-tablecollection-getfirstornullobject-member(1))|Возвращает первую таблицу в коллекции.|
||[items](/javascript/api/word/word.tablecollection#word-word-tablecollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[TableRow](/javascript/api/word/word.tablerow)|[cellCount](/javascript/api/word/word.tablerow#word-word-tablerow-cellcount-member)|Получает количество ячеек в строке.|
||[ячейки](/javascript/api/word/word.tablerow#word-word-tablerow-cells-member)|Возвращает ячейки.|
||[clear()](/javascript/api/word/word.tablerow#word-word-tablerow-clear-member(1))|Очищает содержимое строки.|
||[delete()](/javascript/api/word/word.tablerow#word-word-tablerow-delete-member(1))|Удаляет всю строку.|
||[font](/javascript/api/word/word.tablerow#word-word-tablerow-font-member)|Возвращает шрифт.|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.tablerow#word-word-tablerow-getborder-member(1))|Возвращает стиль границ ячеек в строке.|
||[getCellPadding (cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.tablerow#word-word-tablerow-getcellpadding-member(1))|Возвращает размер поля ячейки в точках.|
||[getNext()](/javascript/api/word/word.tablerow#word-word-tablerow-getnext-member(1))|Возвращает следующую строку.|
||[getNextOrNullObject()](/javascript/api/word/word.tablerow#word-word-tablerow-getnextornullobject-member(1))|Возвращает следующую строку.|
||[horizontalAlignment](/javascript/api/word/word.tablerow#word-word-tablerow-horizontalalignment-member)|Возвращает и задает горизонтальное выравнивание для каждой ячейки в строке.|
||[ignorePunct](/javascript/api/word/word.tablerow#word-word-tablerow-ignorepunct-member)||
||[ignoreSpace](/javascript/api/word/word.tablerow#word-word-tablerow-ignorespace-member)||
||[insertRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[]])](/javascript/api/word/word.tablerow#word-word-tablerow-insertrows-member(1))|Вставляет строки, используя данную строку в качестве шаблона.|
||[isHeader](/javascript/api/word/word.tablerow#word-word-tablerow-isheader-member)|Проверяет, является ли элемент строкой заголовков.|
||[matchCase](/javascript/api/word/word.tablerow#word-word-tablerow-matchcase-member)||
||[matchPrefix](/javascript/api/word/word.tablerow#word-word-tablerow-matchprefix-member)||
||[matchSuffix](/javascript/api/word/word.tablerow#word-word-tablerow-matchsuffix-member)||
||[matchWholeWord](/javascript/api/word/word.tablerow#word-word-tablerow-matchwholeword-member)||
||[matchWildcards](/javascript/api/word/word.tablerow#word-word-tablerow-matchwildcards-member)||
||[parentTable](/javascript/api/word/word.tablerow#word-word-tablerow-parenttable-member)|Возвращает родительскую таблицу.|
||[preferredHeight](/javascript/api/word/word.tablerow#word-word-tablerow-preferredheight-member)|Возвращает и задает предпочитаемую высоту строки в точках.|
||[rowIndex](/javascript/api/word/word.tablerow#word-word-tablerow-rowindex-member)|Получает индекс строки в родительской таблице.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.tablerow#word-word-tablerow-search-member(1))|Выполняет поиск с указанными SearchOptions в области строки.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.tablerow#word-word-tablerow-select-member(1))|Выбирает строку и переходит к ней в Word.|
||[setCellPadding (cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.tablerow#word-word-tablerow-setcellpadding-member(1))|Задает размер поля ячейки в точках.|
||[shadingColor](/javascript/api/word/word.tablerow#word-word-tablerow-shadingcolor-member)|Возвращает и задает цвет заливки.|
||[values](/javascript/api/word/word.tablerow#word-word-tablerow-values-member)|Получает и задает текстовые значения строки в виде массива Javascript 2D.|
||[verticalAlignment](/javascript/api/word/word.tablerow#word-word-tablerow-verticalalignment-member)|Возвращает и задает вертикальное выравнивание ячеек в строке.|
|[TableRowCollection](/javascript/api/word/word.tablerowcollection)|[getFirst()](/javascript/api/word/word.tablerowcollection#word-word-tablerowcollection-getfirst-member(1))|Возвращает первую строку в коллекции.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablerowcollection#word-word-tablerowcollection-getfirstornullobject-member(1))|Возвращает первую строку в коллекции.|
||[items](/javascript/api/word/word.tablerowcollection#word-word-tablerowcollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Word](/javascript/api/word)
- [Наборы обязательных элементов API JavaScript для Word](word-api-requirement-sets.md)
