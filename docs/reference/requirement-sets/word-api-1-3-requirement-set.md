---
title: Набор API API Word JavaScript 1.3
description: Сведения о наборе требований WordApi 1.3.
ms.date: 03/09/2021
ms.prod: word
localization_priority: Normal
ms.openlocfilehash: 4943eeb020e99f9a87d77996c59ea838e84ec6eecf705cb483930dc948d4e8c1
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57092166"
---
# <a name="whats-new-in-word-javascript-api-13"></a>Новые возможности API JavaScript для Word 1.3

WordApi 1.3 добавила больше поддержки элементов управления контентом и параметров уровня документов.

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в API Word JavaScript, за набором 1.3. Чтобы просмотреть справочную документацию API для всех API, поддерживаемых требованием API Word JavaScript, установленным 1.3 или ранее, см. в справке Word API в наборе требований [1.3 или более ранних](/javascript/api/word?view=word-js-1.3&preserve-view=true).

| Класс | Поля | Описание |
|:---|:---|:---|
|[Application](/javascript/api/word/word.application)|[createDocument(base64File?: string)](/javascript/api/word/word.application#createDocument_base64File_)|Создает новый документ с помощью дополнительного файла base64, закодированного .docx.|
|[Основной текст](/javascript/api/word/word.body)|[getRange (rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.body#getRange_rangeLocation_)|Возвращает весь основной текст (либо его начальную или конечную точку) в виде диапазона.|
||[insertTable (rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.body#insertTable_rowCount__columnCount__insertLocation__values_)|Вставляет таблицу с указанным количеством строк и столбцов.|
||[lists](/javascript/api/word/word.body#lists)|Возвращает коллекцию объектов списков в основном тексте.|
||[parentBody](/javascript/api/word/word.body#parentBody)|Возвращает родительский текст основного текста.|
||[parentBodyOrNullObject](/javascript/api/word/word.body#parentBodyOrNullObject)|Возвращает родительский текст основного текста.|
||[parentContentControlOrNullObject](/javascript/api/word/word.body#parentContentControlOrNullObject)|Получает элемент управления содержимым, содержащий документ или раздел.|
||[parentSection](/javascript/api/word/word.body#parentSection)|Возвращает родительский раздел основного текста.|
||[parentSectionOrNullObject](/javascript/api/word/word.body#parentSectionOrNullObject)|Возвращает родительский раздел основного текста.|
||[таблицы](/javascript/api/word/word.body#tables)|Возвращает коллекцию объектов таблиц в основном тексте.|
||[type](/javascript/api/word/word.body#type)|Возвращает тип основного текста.|
||[styleBuiltIn](/javascript/api/word/word.body#styleBuiltIn)|Возвращает или задает имя встроенного стиля основного текста.|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[getRange (rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.contentcontrol#getRange_rangeLocation_)|Возвращает весь элемент управления содержимым (либо его начальную или конечную точку) в виде диапазона.|
||[getTextRanges (endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.contentcontrol#getTextRanges_endingMarks__trimSpacing_)|Получает диапазоны текстов в области управления контентом с помощью знаков препинания и/или других знаков окончания.|
||[insertTable (rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.contentcontrol#insertTable_rowCount__columnCount__insertLocation__values_)|Вставляет таблицу с указанным количеством строк и столбцов в элемент управления содержимым или рядом с ним.|
||[lists](/javascript/api/word/word.contentcontrol#lists)|Возвращает коллекцию объектов списков в элементе управления содержимым.|
||[parentBody](/javascript/api/word/word.contentcontrol#parentBody)|Возвращает родительский текст элемента управления содержимым.|
||[parentContentControlOrNullObject](/javascript/api/word/word.contentcontrol#parentContentControlOrNullObject)|Получает элемент управления содержимым, содержащий элемент управления содержимым.|
||[parentTable](/javascript/api/word/word.contentcontrol#parentTable)|Возвращает таблицу, содержащую элемент управления содержимым.|
||[parentTableCell](/javascript/api/word/word.contentcontrol#parentTableCell)|Возвращает ячейку таблицы, содержащую элемент управления содержимым.|
||[parentTableCellOrNullObject](/javascript/api/word/word.contentcontrol#parentTableCellOrNullObject)|Возвращает ячейку таблицы, содержащую элемент управления содержимым.|
||[parentTableOrNullObject](/javascript/api/word/word.contentcontrol#parentTableOrNullObject)|Возвращает таблицу, содержащую элемент управления содержимым.|
||[подтип](/javascript/api/word/word.contentcontrol#subtype)|Возвращает подтип элемента управления содержимым.|
||[таблицы](/javascript/api/word/word.contentcontrol#tables)|Возвращает коллекцию объектов таблиц в элементе управления содержимым.|
||[split(delimiters: string[], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.contentcontrol#split_delimiters__multiParagraphs__trimDelimiters__trimSpacing_)|Разделяет элемент управления содержимым на дочерние диапазоны с помощью разделителей.|
||[styleBuiltIn](/javascript/api/word/word.contentcontrol#styleBuiltIn)|Возвращает или задает имя встроенного стиля для элемента управления содержимым.|
|[ContentControlCollection](/javascript/api/word/word.contentcontrolcollection)|[getByIdOrNullObject (id: number)](/javascript/api/word/word.contentcontrolcollection#getByIdOrNullObject_id_)|Возвращает элемент управления содержимым по его идентификатору.|
||[getByTypes (типы: Word.ContentControlType[])](/javascript/api/word/word.contentcontrolcollection#getByTypes_types_)|Получает элементы управления контентом, которые имеют указанные типы и/или подтипы.|
||[getFirst()](/javascript/api/word/word.contentcontrolcollection#getFirst__)|Возвращает первый элемент управления содержимым в коллекции.|
||[getFirstOrNullObject()](/javascript/api/word/word.contentcontrolcollection#getFirstOrNullObject__)|Возвращает первый элемент управления содержимым в коллекции.|
|[CustomProperty](/javascript/api/word/word.customproperty)|[delete()](/javascript/api/word/word.customproperty#delete__)|Удаляет настраиваемое свойство.|
||[key](/javascript/api/word/word.customproperty#key)|Возвращает ключ настраиваемого свойства.|
||[type](/javascript/api/word/word.customproperty#type)|Получает тип значения настраиваемого свойства.|
||[value](/javascript/api/word/word.customproperty#value)|Получает или задает значение настраиваемого свойства.|
|[CustomPropertyCollection](/javascript/api/word/word.custompropertycollection)|[add(key: string, value: any)](/javascript/api/word/word.custompropertycollection#add_key__value_)|Создает или задает настраиваемое свойство.|
||[deleteAll()](/javascript/api/word/word.custompropertycollection#deleteAll__)|Удаляет все настраиваемые свойства в коллекции.|
||[getCount()](/javascript/api/word/word.custompropertycollection#getCount__)|Получает количество настраиваемых свойств.|
||[getItem(key: string)](/javascript/api/word/word.custompropertycollection#getItem_key_)|Возвращает объект настраиваемого свойства по ключу, указываемому без учета регистра.|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.custompropertycollection#getItemOrNullObject_key_)|Возвращает объект настраиваемого свойства по ключу, указываемому без учета регистра.|
||[items](/javascript/api/word/word.custompropertycollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Document](/javascript/api/word/word.document)|[properties](/javascript/api/word/word.document#properties)|Получает свойства документа.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[open()](/javascript/api/word/word.documentcreated#open__)|Открывает документ.|
||[body](/javascript/api/word/word.documentcreated#body)|Получает объект тела документа.|
||[contentControls](/javascript/api/word/word.documentcreated#contentControls)|Получает коллекцию объектов управления контентом в документе.|
||[properties](/javascript/api/word/word.documentcreated#properties)|Получает свойства документа.|
||[сохранено](/javascript/api/word/word.documentcreated#saved)|Указывает, сохранены ли изменения, внесенные в документ.|
||[sections](/javascript/api/word/word.documentcreated#sections)|Получает коллекцию объектов раздела в документе.|
||[save()](/javascript/api/word/word.documentcreated#save__)|Сохраняет документ.|
|[DocumentProperties](/javascript/api/word/word.documentproperties)|[автор](/javascript/api/word/word.documentproperties#author)|Возвращает или задает автора документа.|
||[категория](/javascript/api/word/word.documentproperties#category)|Возвращает или задает категорию документа.|
||[comments](/javascript/api/word/word.documentproperties#comments)|Возвращает или задает примечания к документу.|
||[company](/javascript/api/word/word.documentproperties#company)|Возвращает или задает компанию документа.|
||[format](/javascript/api/word/word.documentproperties#format)|Возвращает или задает формат документа.|
||[ключевые слова](/javascript/api/word/word.documentproperties#keywords)|Возвращает или задает ключевые слова документа.|
||[manager](/javascript/api/word/word.documentproperties#manager)|Возвращает или задает менеджера документа.|
||[applicationName](/javascript/api/word/word.documentproperties#applicationName)|Возвращает имя приложения для документа.|
||[creationDate](/javascript/api/word/word.documentproperties#creationDate)|Возвращает дату создания документа.|
||[customProperties](/javascript/api/word/word.documentproperties#customProperties)|Возвращает коллекцию настраиваемых свойств документа.|
||[lastAuthor](/javascript/api/word/word.documentproperties#lastAuthor)|Получает последнего автора документа.|
||[lastPrintDate](/javascript/api/word/word.documentproperties#lastPrintDate)|Возвращает дату последней печати документа.|
||[lastSaveTime](/javascript/api/word/word.documentproperties#lastSaveTime)|Возвращает время последнего сохранения документа.|
||[revisionNumber](/javascript/api/word/word.documentproperties#revisionNumber)|Возвращает номер редакции документа.|
||[безопасность](/javascript/api/word/word.documentproperties#security)|Получает параметры безопасности документа.|
||[template](/javascript/api/word/word.documentproperties#template)|Возвращает шаблон документа.|
||[subject](/javascript/api/word/word.documentproperties#subject)|Возвращает или задает тему документа.|
||[заголовок](/javascript/api/word/word.documentproperties#title)|Возвращает или задает название документа.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[getNext()](/javascript/api/word/word.inlinepicture#getNext__)|Возвращает следующий встроенный рисунок.|
||[getNextOrNullObject()](/javascript/api/word/word.inlinepicture#getNextOrNullObject__)|Возвращает следующий встроенный рисунок.|
||[getRange (rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.inlinepicture#getRange_rangeLocation_)|Возвращает рисунок (либо его начальную или конечную точку) в виде диапазона.|
||[parentContentControlOrNullObject](/javascript/api/word/word.inlinepicture#parentContentControlOrNullObject)|Возвращает элемент управления содержимым, который содержит встроенный рисунок.|
||[parentTable](/javascript/api/word/word.inlinepicture#parentTable)|Возвращает таблицу, содержащую встроенный рисунок.|
||[parentTableCell](/javascript/api/word/word.inlinepicture#parentTableCell)|Возвращает ячейку таблицы, содержащую встроенный рисунок.|
||[parentTableCellOrNullObject](/javascript/api/word/word.inlinepicture#parentTableCellOrNullObject)|Возвращает ячейку таблицы, содержащую встроенный рисунок.|
||[parentTableOrNullObject](/javascript/api/word/word.inlinepicture#parentTableOrNullObject)|Возвращает таблицу, содержащую встроенный рисунок.|
|[InlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|[getFirst()](/javascript/api/word/word.inlinepicturecollection#getFirst__)|Возвращает первый встроенный рисунок в коллекции.|
||[getFirstOrNullObject()](/javascript/api/word/word.inlinepicturecollection#getFirstOrNullObject__)|Возвращает первый встроенный рисунок в коллекции.|
|[Список](/javascript/api/word/word.list)|[getLevelParagraphs(level: number)](/javascript/api/word/word.list#getLevelParagraphs_level_)|Возвращает абзацы, обнаруженные на указанном уровне списка.|
||[getLevelString (уровень: номер)](/javascript/api/word/word.list#getLevelString_level_)|Получает пулю, номер или изображение на указанном уровне в качестве строки.|
||[insertParagraph (paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.list#insertParagraph_paragraphText__insertLocation_)|Вставляет абзац в указанном расположении.|
||[id](/javascript/api/word/word.list#id)|Получает id списка.|
||[levelExistences](/javascript/api/word/word.list#levelExistences)|Проверяет наличие каждого из 9 уровней в списке.|
||[levelTypes](/javascript/api/word/word.list#levelTypes)|Возвращает типы всех 9 уровней списка.|
||[paragraphs](/javascript/api/word/word.list#paragraphs)|Возвращает абзацы в списке.|
||[setLevelAlignment (уровень: номер, выравнивание: Word.Alignment)](/javascript/api/word/word.list#setLevelAlignment_level__alignment_)|Задает выравнивание пули, номера или изображения на указанном уровне в списке.|
||[setLevelBullet(level: number, listBullet: Word.ListBullet, charCode?: number, fontName?: string)](/javascript/api/word/word.list#setLevelBullet_level__listBullet__charCode__fontName_)|Задает формат маркеров на указанном уровне списка.|
||[setLevelIndents(level: number, textIndent: number, bulletNumberPictureIndent: number)](/javascript/api/word/word.list#setLevelIndents_level__textIndent__bulletNumberPictureIndent_)|Задает два отступа на указанном уровне списка.|
||[setLevelNumbering(level: number, listNumbering: Word.ListNumbering, formatString?: Array<string \| number>)](/javascript/api/word/word.list#setLevelNumbering_level__listNumbering__formatString_)|Задает формат нумерации на указанном уровне списка.|
||[setLevelStartingNumber (уровень: номер, startingNumber: number)](/javascript/api/word/word.list#setLevelStartingNumber_level__startingNumber_)|Задает начальный номер на указанном уровне списка.|
|[ListCollection](/javascript/api/word/word.listcollection)|[getById(id: number)](/javascript/api/word/word.listcollection#getById_id_)|Возвращает список по идентификатору.|
||[getByIdOrNullObject (id: number)](/javascript/api/word/word.listcollection#getByIdOrNullObject_id_)|Возвращает список по идентификатору.|
||[getFirst()](/javascript/api/word/word.listcollection#getFirst__)|Возвращает первый список в коллекции.|
||[getFirstOrNullObject()](/javascript/api/word/word.listcollection#getFirstOrNullObject__)|Возвращает первый список в коллекции.|
||[getItem(index: number)](/javascript/api/word/word.listcollection#getItem_index_)|Возвращает объект списка по индексу в коллекции.|
||[items](/javascript/api/word/word.listcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[ListItem](/javascript/api/word/word.listitem)|[getAncestor(parentOnly?: boolean)](/javascript/api/word/word.listitem#getAncestor_parentOnly_)|Возвращает родительский элемент или ближайшего предка (если родительского элемента нет) для данного элемента списка.|
||[getAncestorOrNullObject(parentOnly?: boolean)](/javascript/api/word/word.listitem#getAncestorOrNullObject_parentOnly_)|Возвращает родительский элемент или ближайшего предка (если родительского элемента нет) для данного элемента списка.|
||[getDescendants (directChildrenOnly?: boolean)](/javascript/api/word/word.listitem#getDescendants_directChildrenOnly_)|Возвращает всех потомков элемента списка.|
||[level](/javascript/api/word/word.listitem#level)|Возвращает или задает уровень элемента в списке.|
||[listString](/javascript/api/word/word.listitem#listString)|Получает пулю элемента списка, номер или изображение в качестве строки.|
||[siblingIndex](/javascript/api/word/word.listitem#siblingIndex)|Возвращает порядковый номер элемента списка относительно элементов того же уровня.|
|[Paragraph](/javascript/api/word/word.paragraph)|[attachToList(listId: number, level: number)](/javascript/api/word/word.paragraph#attachToList_listId__level_)|Позволяет присоединить абзац к существующему списку на указанном уровне.|
||[detachFromList()](/javascript/api/word/word.paragraph#detachFromList__)|Перемещает абзац за пределы списка (если он является элементом списка).|
||[getNext()](/javascript/api/word/word.paragraph#getNext__)|Возвращает следующий абзац.|
||[getNextOrNullObject()](/javascript/api/word/word.paragraph#getNextOrNullObject__)|Возвращает следующий абзац.|
||[getPrevious()](/javascript/api/word/word.paragraph#getPrevious__)|Возвращает предыдущий абзац.|
||[getPreviousOrNullObject()](/javascript/api/word/word.paragraph#getPreviousOrNullObject__)|Возвращает предыдущий абзац.|
||[getRange (rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.paragraph#getRange_rangeLocation_)|Возвращает весь абзац (либо его начальную или конечную точку) в виде диапазона.|
||[getTextRanges (endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.paragraph#getTextRanges_endingMarks__trimSpacing_)|Получает диапазоны текста в абзаце, используя знаки препинания и/или другие знаки окончания.|
||[insertTable (rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.paragraph#insertTable_rowCount__columnCount__insertLocation__values_)|Вставляет таблицу с указанным количеством строк и столбцов.|
||[isLastParagraph](/javascript/api/word/word.paragraph#isLastParagraph)|Указывает, что абзац является последним в родительском тексте.|
||[isListItem](/javascript/api/word/word.paragraph#isListItem)|Проверяет, является ли абзац элементом списка.|
||[list](/javascript/api/word/word.paragraph#list)|Возвращает объект List, к которому относится абзац.|
||[listItem](/javascript/api/word/word.paragraph#listItem)|Возвращает объект ListItem для абзаца.|
||[listItemOrNullObject](/javascript/api/word/word.paragraph#listItemOrNullObject)|Возвращает объект ListItem для абзаца.|
||[listOrNullObject](/javascript/api/word/word.paragraph#listOrNullObject)|Возвращает объект List, к которому относится абзац.|
||[parentBody](/javascript/api/word/word.paragraph#parentBody)|Возвращает родительский текст абзаца.|
||[parentContentControlOrNullObject](/javascript/api/word/word.paragraph#parentContentControlOrNullObject)|Возвращает элемент управления содержимым, содержащий абзац.|
||[parentTable](/javascript/api/word/word.paragraph#parentTable)|Возвращает таблицу, содержащую абзац.|
||[parentTableCell](/javascript/api/word/word.paragraph#parentTableCell)|Возвращает ячейку таблицы, содержащую абзац.|
||[parentTableCellOrNullObject](/javascript/api/word/word.paragraph#parentTableCellOrNullObject)|Возвращает ячейку таблицы, содержащую абзац.|
||[parentTableOrNullObject](/javascript/api/word/word.paragraph#parentTableOrNullObject)|Возвращает таблицу, содержащую абзац.|
||[tableNestingLevel](/javascript/api/word/word.paragraph#tableNestingLevel)|Возвращает уровень таблицы, содержащей абзац.|
||[split(delimiters: string[], trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.paragraph#split_delimiters__trimDelimiters__trimSpacing_)|Разделяет абзац на дочерние диапазоны с помощью разделителей.|
||[startNewList()](/javascript/api/word/word.paragraph#startNewList__)|Создает список, начинающийся с данного абзаца.|
||[styleBuiltIn](/javascript/api/word/word.paragraph#styleBuiltIn)|Возвращает или задает имя встроенного стиля абзаца.|
|[ParagraphCollection](/javascript/api/word/word.paragraphcollection)|[getFirst()](/javascript/api/word/word.paragraphcollection#getFirst__)|Возвращает первый абзац в коллекции.|
||[getFirstOrNullObject()](/javascript/api/word/word.paragraphcollection#getFirstOrNullObject__)|Возвращает первый абзац в коллекции.|
||[getLast()](/javascript/api/word/word.paragraphcollection#getLast__)|Возвращает последний абзац в коллекции.|
||[getLastOrNullObject()](/javascript/api/word/word.paragraphcollection#getLastOrNullObject__)|Возвращает последний абзац в коллекции.|
|[Range](/javascript/api/word/word.range)|[compareLocationWith (диапазон: Word.Range)](/javascript/api/word/word.range#compareLocationWith_range_)|Сравнивает расположение данного диапазона с расположением другого диапазона.|
||[expandTo (диапазон: Word.Range)](/javascript/api/word/word.range#expandTo_range_)|Возвращает новый диапазон, который простирается в том или ином направлении от данного диапазона и перекрывает другой диапазон.|
||[expandToOrNullObject (диапазон: Word.Range)](/javascript/api/word/word.range#expandToOrNullObject_range_)|Возвращает новый диапазон, который простирается в том или ином направлении от данного диапазона и перекрывает другой диапазон.|
||[getHyperlinkRanges()](/javascript/api/word/word.range#getHyperlinkRanges__)|Возвращает дочерние диапазоны гиперссылок в данном диапазоне.|
||[getNextTextRange (endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#getNextTextRange_endingMarks__trimSpacing_)|Получает следующий диапазон текста, используя знаки препинания и/или другие знаки окончания.|
||[getNextTextRangeOrNullObject (endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#getNextTextRangeOrNullObject_endingMarks__trimSpacing_)|Получает следующий диапазон текста, используя знаки препинания и/или другие знаки окончания.|
||[getRange (rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.range#getRange_rangeLocation_)|Клонирует диапазон либо получает его начальную или конечную точку в виде нового диапазона.|
||[getTextRanges (endingMarks: string[], trimSpacing?: boolean)](/javascript/api/word/word.range#getTextRanges_endingMarks__trimSpacing_)|Получает текстовые детские диапазоны в диапазоне, используя знаки препинания и/или другие знаки окончания.|
||[hyperlink](/javascript/api/word/word.range#hyperlink)|Возвращает первую гиперссылку в диапазоне или задает для него гиперссылку.|
||[insertTable (rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.range#insertTable_rowCount__columnCount__insertLocation__values_)|Вставляет таблицу с указанным количеством строк и столбцов.|
||[intersectWith (диапазон: Word.Range)](/javascript/api/word/word.range#intersectWith_range_)|Возвращает новый диапазон, представляющий собой пересечение данного диапазона с другим.|
||[intersectWithOrNullObject (диапазон: Word.Range)](/javascript/api/word/word.range#intersectWithOrNullObject_range_)|Возвращает новый диапазон, представляющий собой пересечение данного диапазона с другим.|
||[isEmpty](/javascript/api/word/word.range#isEmpty)|Проверяет, является ли длина диапазона нулевой.|
||[lists](/javascript/api/word/word.range#lists)|Возвращает коллекцию объектов списков в диапазоне.|
||[parentBody](/javascript/api/word/word.range#parentBody)|Возвращает родительский текст диапазона.|
||[parentContentControlOrNullObject](/javascript/api/word/word.range#parentContentControlOrNullObject)|Возвращает элемент управления содержимым, содержащий диапазон.|
||[parentTable](/javascript/api/word/word.range#parentTable)|Возвращает таблицу, содержащую диапазон.|
||[parentTableCell](/javascript/api/word/word.range#parentTableCell)|Возвращает ячейку таблицы, содержащую диапазон.|
||[parentTableCellOrNullObject](/javascript/api/word/word.range#parentTableCellOrNullObject)|Возвращает ячейку таблицы, содержащую диапазон.|
||[parentTableOrNullObject](/javascript/api/word/word.range#parentTableOrNullObject)|Возвращает таблицу, содержащую диапазон.|
||[таблицы](/javascript/api/word/word.range#tables)|Возвращает коллекцию объектов таблиц в диапазоне.|
||[split(delimiters: string[], multiParagraphs?: boolean, trimDelimiters?: boolean, trimSpacing?: boolean)](/javascript/api/word/word.range#split_delimiters__multiParagraphs__trimDelimiters__trimSpacing_)|Разделяет диапазон на дочерние диапазоны с помощью разделителей.|
||[styleBuiltIn](/javascript/api/word/word.range#styleBuiltIn)|Возвращает или задает имя встроенного стиля диапазона.|
|[RangeCollection](/javascript/api/word/word.rangecollection)|[getFirst()](/javascript/api/word/word.rangecollection#getFirst__)|Возвращает первый диапазон в коллекции.|
||[getFirstOrNullObject()](/javascript/api/word/word.rangecollection#getFirstOrNullObject__)|Возвращает первый диапазон в коллекции.|
|[RequestContext](/javascript/api/word/word.requestcontext)|[application](/javascript/api/word/word.requestcontext#application)|[Набор API: WordApi 1.3] *|
|[Section](/javascript/api/word/word.section)|[getNext()](/javascript/api/word/word.section#getNext__)|Возвращает следующий раздел.|
||[getNextOrNullObject()](/javascript/api/word/word.section#getNextOrNullObject__)|Возвращает следующий раздел.|
|[SectionCollection](/javascript/api/word/word.sectioncollection)|[getFirst()](/javascript/api/word/word.sectioncollection#getFirst__)|Возвращает первый раздел в коллекции.|
||[getFirstOrNullObject()](/javascript/api/word/word.sectioncollection#getFirstOrNullObject__)|Возвращает первый раздел в коллекции.|
|[Table](/javascript/api/word/word.table)|[addColumns(insertLocation: Word.InsertLocation, columnCount: number, values?: string[][])](/javascript/api/word/word.table#addColumns_insertLocation__columnCount__values_)|Добавляет столбцы в начале или в конце таблицы, используя первый или последний из имеющихся столбцов в качестве шаблона.|
||[addRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[][])](/javascript/api/word/word.table#addRows_insertLocation__rowCount__values_)|Добавляет строки в начале или в конце таблицы, используя первую или последнюю из имеющихся строк в качестве шаблона.|
||[выравнивание](/javascript/api/word/word.table#alignment)|Получает или задает выравнивание таблицы со столбцом страницы.|
||[autoFitWindow()](/javascript/api/word/word.table#autoFitWindow__)|Автоматически подбирает ширину столбцов таблицы в соответствии с шириной окна.|
||[clear()](/javascript/api/word/word.table#clear__)|Очищает содержимое таблицы.|
||[delete()](/javascript/api/word/word.table#delete__)|Удаляет всю таблицу.|
||[deleteColumns(columnIndex: number, columnCount?: number)](/javascript/api/word/word.table#deleteColumns_columnIndex__columnCount_)|Удаляет определенные столбцы.|
||[deleteRows(rowIndex: number, rowCount?: number)](/javascript/api/word/word.table#deleteRows_rowIndex__rowCount_)|Удаляет определенные строки.|
||[distributeColumns()](/javascript/api/word/word.table#distributeColumns__)|Равномерно распределяет ширину столбцов.|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.table#getBorder_borderLocation_)|Возвращает стиль указанной границы.|
||[getCell(rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#getCell_rowIndex__cellIndex_)|Возвращает ячейку таблицы в указанной строке и указанном столбце.|
||[getCellOrNullObject (rowIndex: number, cellIndex: number)](/javascript/api/word/word.table#getCellOrNullObject_rowIndex__cellIndex_)|Возвращает ячейку таблицы в указанной строке и указанном столбце.|
||[getCellPadding (cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.table#getCellPadding_cellPaddingLocation_)|Возвращает размер поля ячейки в точках.|
||[getNext()](/javascript/api/word/word.table#getNext__)|Возвращает следующую таблицу.|
||[getNextOrNullObject()](/javascript/api/word/word.table#getNextOrNullObject__)|Возвращает следующую таблицу.|
||[getParagraphAfter()](/javascript/api/word/word.table#getParagraphAfter__)|Возвращает абзац после таблицы.|
||[getParagraphAfterOrNullObject()](/javascript/api/word/word.table#getParagraphAfterOrNullObject__)|Возвращает абзац после таблицы.|
||[getParagraphBefore()](/javascript/api/word/word.table#getParagraphBefore__)|Возвращает абзац перед таблицей.|
||[getParagraphBeforeOrNullObject()](/javascript/api/word/word.table#getParagraphBeforeOrNullObject__)|Возвращает абзац перед таблицей.|
||[getRange (rangeLocation?: Word.RangeLocation)](/javascript/api/word/word.table#getRange_rangeLocation_)|Возвращает диапазон, содержащий данную таблицу, либо диапазон в начале или в конце таблицы.|
||[headerRowCount](/javascript/api/word/word.table#headerRowCount)|Возвращает и задает количество строк заголовков.|
||[horizontalAlignment](/javascript/api/word/word.table#horizontalAlignment)|Возвращает и задает горизонтальное выравнивание для каждой ячейки в таблице.|
||[ignorePunct](/javascript/api/word/word.table#ignorePunct)||
||[ignoreSpace](/javascript/api/word/word.table#ignoreSpace)||
||[insertContentControl()](/javascript/api/word/word.table#insertContentControl__)|Вставляет в таблицу элемент управления содержимым.|
||[insertParagraph (paragraphText: string, insertLocation: Word.InsertLocation)](/javascript/api/word/word.table#insertParagraph_paragraphText__insertLocation_)|Вставляет абзац в указанном расположении.|
||[insertTable (rowCount: number, columnCount: number, insertLocation: Word.InsertLocation, values?: string[][])](/javascript/api/word/word.table#insertTable_rowCount__columnCount__insertLocation__values_)|Вставляет таблицу с указанным количеством строк и столбцов.|
||[matchCase](/javascript/api/word/word.table#matchCase)||
||[matchPrefix](/javascript/api/word/word.table#matchPrefix)||
||[matchSuffix](/javascript/api/word/word.table#matchSuffix)||
||[matchWholeWord](/javascript/api/word/word.table#matchWholeWord)||
||[matchWildcards](/javascript/api/word/word.table#matchWildcards)||
||[font](/javascript/api/word/word.table#font)|Возвращает шрифт.|
||[isUniform](/javascript/api/word/word.table#isUniform)|Указывает, однородны ли все строки таблицы.|
||[nestingLevel](/javascript/api/word/word.table#nestingLevel)|Возвращает уровень вложенности таблицы.|
||[parentBody](/javascript/api/word/word.table#parentBody)|Возвращает родительский текст таблицы.|
||[parentContentControl](/javascript/api/word/word.table#parentContentControl)|Возвращает элемент управления содержимым, содержащий таблицу.|
||[parentContentControlOrNullObject](/javascript/api/word/word.table#parentContentControlOrNullObject)|Возвращает элемент управления содержимым, содержащий таблицу.|
||[parentTable](/javascript/api/word/word.table#parentTable)|Возвращает таблицу, которая содержит данную таблицу.|
||[parentTableCell](/javascript/api/word/word.table#parentTableCell)|Возвращает ячейку таблицы, содержащую данную таблицу.|
||[parentTableCellOrNullObject](/javascript/api/word/word.table#parentTableCellOrNullObject)|Возвращает ячейку таблицы, содержащую данную таблицу.|
||[parentTableOrNullObject](/javascript/api/word/word.table#parentTableOrNullObject)|Возвращает таблицу, которая содержит данную таблицу.|
||[rowCount](/javascript/api/word/word.table#rowCount)|Получает количество строк в таблице.|
||[строки](/javascript/api/word/word.table#rows)|Возвращает все строки таблицы.|
||[таблицы](/javascript/api/word/word.table#tables)|Возвращает дочерние таблицы, вложенные на один уровень ниже.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.table#search_searchText__searchOptions__ignorePunct__ignoreSpace__matchCase__matchPrefix__matchSuffix__matchWholeWord__matchWildcards_)|Выполняет поиск с указанными SearchOptions в области объекта таблицы.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.table#select_selectionMode_)|Выбирает таблицу либо позицию в начале или в конце таблицы, а затем переходит к ней в Word.|
||[setCellPadding (cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.table#setCellPadding_cellPaddingLocation__cellPadding_)|Задает размер поля ячейки в точках.|
||[shadingColor](/javascript/api/word/word.table#shadingColor)|Возвращает и задает цвет заливки.|
||[style](/javascript/api/word/word.table#style)|Возвращает или задает имя стиля для таблицы.|
||[styleBandedColumns](/javascript/api/word/word.table#styleBandedColumns)|Возвращает и задает значение, указывающее, есть ли в таблице чередующиеся столбцы.|
||[styleBandedRows](/javascript/api/word/word.table#styleBandedRows)|Возвращает и задает значение, указывающее, есть ли в таблице чередующиеся строки.|
||[styleBuiltIn](/javascript/api/word/word.table#styleBuiltIn)|Возвращает или задает имя встроенного стиля таблицы.|
||[styleFirstColumn](/javascript/api/word/word.table#styleFirstColumn)|Возвращает и задает значение, указывающее, применен ли специальный стиль к первому столбцу таблицы.|
||[styleLastColumn](/javascript/api/word/word.table#styleLastColumn)|Возвращает и задает значение, указывающее, применен ли специальный стиль к последнему столбцу таблицы.|
||[styleTotalRow](/javascript/api/word/word.table#styleTotalRow)|Возвращает и задает значение, указывающее, применен ли специальный стиль к строке итогов (последней строке) таблицы.|
||[values](/javascript/api/word/word.table#values)|Возвращает и задает текстовые значения в таблице в виде двумерного массива JavaScript.|
||[verticalAlignment](/javascript/api/word/word.table#verticalAlignment)|Возвращает и задает вертикальное выравнивание для каждой ячейки в таблице.|
||[width](/javascript/api/word/word.table#width)|Возвращает и задает ширину таблицы в точках.|
|[TableBorder](/javascript/api/word/word.tableborder)|[color](/javascript/api/word/word.tableborder#color)|Получает или задает цвет границы таблицы.|
||[type](/javascript/api/word/word.tableborder#type)|Возвращает или задает тип границы таблицы.|
||[width](/javascript/api/word/word.tableborder#width)|Возвращает или задает ширину границы таблицы в точках.|
|[TableCell](/javascript/api/word/word.tablecell)|[columnWidth](/javascript/api/word/word.tablecell#columnWidth)|Возвращает и задает ширину столбца ячейки в точках.|
||[deleteColumn()](/javascript/api/word/word.tablecell#deleteColumn__)|Удаляет столбец, содержащий данную ячейку.|
||[deleteRow()](/javascript/api/word/word.tablecell#deleteRow__)|Удаляет строку, содержащую данную ячейку.|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.tablecell#getBorder_borderLocation_)|Возвращает стиль указанной границы.|
||[getCellPadding (cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.tablecell#getCellPadding_cellPaddingLocation_)|Возвращает размер поля ячейки в точках.|
||[getNext()](/javascript/api/word/word.tablecell#getNext__)|Возвращает следующую ячейку.|
||[getNextOrNullObject()](/javascript/api/word/word.tablecell#getNextOrNullObject__)|Возвращает следующую ячейку.|
||[horizontalAlignment](/javascript/api/word/word.tablecell#horizontalAlignment)|Возвращает и задает горизонтальное выравнивание ячейки.|
||[insertColumns(insertLocation: Word.InsertLocation, columnCount: number, values?: string[][])](/javascript/api/word/word.tablecell#insertColumns_insertLocation__columnCount__values_)|Добавляет столбцы слева или справа от ячейки, используя столбец этой ячейки в качестве шаблона.|
||[insertRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[]])](/javascript/api/word/word.tablecell#insertRows_insertLocation__rowCount__values_)|Вставляет строки над ячейкой или под ней, используя строку этой ячейки в качестве шаблона.|
||[body](/javascript/api/word/word.tablecell#body)|Возвращает объект тела ячейки.|
||[cellIndex](/javascript/api/word/word.tablecell#cellIndex)|Получает индекс ячейки в строке.|
||[parentRow](/javascript/api/word/word.tablecell#parentRow)|Получает родительскую строку ячейки.|
||[parentTable](/javascript/api/word/word.tablecell#parentTable)|Возвращает родительскую таблицу ячейки.|
||[rowIndex](/javascript/api/word/word.tablecell#rowIndex)|Получает индекс строки ячейки в таблице.|
||[width](/javascript/api/word/word.tablecell#width)|Возвращает ширину ячейки в точках.|
||[setCellPadding (cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.tablecell#setCellPadding_cellPaddingLocation__cellPadding_)|Задает размер поля ячейки в точках.|
||[shadingColor](/javascript/api/word/word.tablecell#shadingColor)|Возвращает или задает цвет заливки ячейки.|
||[value](/javascript/api/word/word.tablecell#value)|Возвращает и задает текст ячейки.|
||[verticalAlignment](/javascript/api/word/word.tablecell#verticalAlignment)|Возвращает и задает вертикальное выравнивание ячейки.|
|[TableCellCollection](/javascript/api/word/word.tablecellcollection)|[getFirst()](/javascript/api/word/word.tablecellcollection#getFirst__)|Возвращает первую ячейку таблицы в коллекции.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablecellcollection#getFirstOrNullObject__)|Возвращает первую ячейку таблицы в коллекции.|
||[items](/javascript/api/word/word.tablecellcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[TableCollection](/javascript/api/word/word.tablecollection)|[getFirst()](/javascript/api/word/word.tablecollection#getFirst__)|Возвращает первую таблицу в коллекции.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablecollection#getFirstOrNullObject__)|Возвращает первую таблицу в коллекции.|
||[items](/javascript/api/word/word.tablecollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[TableRow](/javascript/api/word/word.tablerow)|[clear()](/javascript/api/word/word.tablerow#clear__)|Очищает содержимое строки.|
||[delete()](/javascript/api/word/word.tablerow#delete__)|Удаляет всю строку.|
||[getBorder(borderLocation: Word.BorderLocation)](/javascript/api/word/word.tablerow#getBorder_borderLocation_)|Возвращает стиль границ ячеек в строке.|
||[getCellPadding (cellPaddingLocation: Word.CellPaddingLocation)](/javascript/api/word/word.tablerow#getCellPadding_cellPaddingLocation_)|Возвращает размер поля ячейки в точках.|
||[getNext()](/javascript/api/word/word.tablerow#getNext__)|Возвращает следующую строку.|
||[getNextOrNullObject()](/javascript/api/word/word.tablerow#getNextOrNullObject__)|Возвращает следующую строку.|
||[horizontalAlignment](/javascript/api/word/word.tablerow#horizontalAlignment)|Возвращает и задает горизонтальное выравнивание для каждой ячейки в строке.|
||[ignorePunct](/javascript/api/word/word.tablerow#ignorePunct)||
||[ignoreSpace](/javascript/api/word/word.tablerow#ignoreSpace)||
||[insertRows(insertLocation: Word.InsertLocation, rowCount: number, values?: string[]])](/javascript/api/word/word.tablerow#insertRows_insertLocation__rowCount__values_)|Вставляет строки, используя данную строку в качестве шаблона.|
||[matchCase](/javascript/api/word/word.tablerow#matchCase)||
||[matchPrefix](/javascript/api/word/word.tablerow#matchPrefix)||
||[matchSuffix](/javascript/api/word/word.tablerow#matchSuffix)||
||[matchWholeWord](/javascript/api/word/word.tablerow#matchWholeWord)||
||[matchWildcards](/javascript/api/word/word.tablerow#matchWildcards)||
||[preferredHeight](/javascript/api/word/word.tablerow#preferredHeight)|Возвращает и задает предпочитаемую высоту строки в точках.|
||[cellCount](/javascript/api/word/word.tablerow#cellCount)|Получает количество ячеек в строке.|
||[ячейки](/javascript/api/word/word.tablerow#cells)|Возвращает ячейки.|
||[font](/javascript/api/word/word.tablerow#font)|Возвращает шрифт.|
||[isHeader](/javascript/api/word/word.tablerow#isHeader)|Проверяет, является ли элемент строкой заголовков.|
||[parentTable](/javascript/api/word/word.tablerow#parentTable)|Возвращает родительскую таблицу.|
||[rowIndex](/javascript/api/word/word.tablerow#rowIndex)|Получает индекс строки в родительской таблице.|
||[search(searchText: string, searchOptions?: Word.SearchOptions \| {ignorePunct?: boolean ignoreSpace?: boolean matchCase?: boolean matchPrefix?: boolean matchSuffix?: boolean matchWholeWord?: boolean matchWildcards?: boolean })](/javascript/api/word/word.tablerow#search_searchText__searchOptions__ignorePunct__ignoreSpace__matchCase__matchPrefix__matchSuffix__matchWholeWord__matchWildcards_)|Выполняет поиск с указанными SearchOptions в области строки.|
||[select(selectionMode?: Word.SelectionMode)](/javascript/api/word/word.tablerow#select_selectionMode_)|Выбирает строку и переходит к ней в Word.|
||[setCellPadding (cellPaddingLocation: Word.CellPaddingLocation, cellPadding: number)](/javascript/api/word/word.tablerow#setCellPadding_cellPaddingLocation__cellPadding_)|Задает размер поля ячейки в точках.|
||[shadingColor](/javascript/api/word/word.tablerow#shadingColor)|Возвращает и задает цвет заливки.|
||[values](/javascript/api/word/word.tablerow#values)|Получает и задает текстовые значения строки в виде массива Javascript 2D.|
||[verticalAlignment](/javascript/api/word/word.tablerow#verticalAlignment)|Возвращает и задает вертикальное выравнивание ячеек в строке.|
|[TableRowCollection](/javascript/api/word/word.tablerowcollection)|[getFirst()](/javascript/api/word/word.tablerowcollection#getFirst__)|Возвращает первую строку в коллекции.|
||[getFirstOrNullObject()](/javascript/api/word/word.tablerowcollection#getFirstOrNullObject__)|Возвращает первую строку в коллекции.|
||[items](/javascript/api/word/word.tablerowcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Word](/javascript/api/word)
- [Наборы обязательных элементов API JavaScript для Word](word-api-requirement-sets.md)
