# <a name="word-javascript-api-requirement-sets"></a>Наборы требований API JavaScript для Word

Наборы требований — это именованные группы членов API. С помощью наборов обязательных элементов, указанных в манифесте, или проверки в среде выполнения надстройки Office определяют, поддерживает ли ведущее приложение Office необходимые API. Дополнительные сведения см. в статье [Версии Office и наборы требований](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).

Надстройки Word выполняются в нескольких версиях Office, включая Office 2016 или более поздние версии для Windows, Office для iPad, Office для Mac и Office Online. В следующей таблице перечислены наборы требований Word, ведущие приложения Office, которые поддерживают этот набор требований, и номера построения или версии для этих приложений.

> [!NOTE]
> Для наборов требований, помеченных как бета, используйте указанную (или более позднюю) версию программного обеспечения Office и используйте бета-библиотеку на CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.
> 
> Операции, которые не перечислены в качестве бета-версии, как правило, доступны, и вы можете продолжить использовать сеть CDN рабочей библиотеки: https://appsforoffice.microsoft.com/lib/1/hosted/office.js

|  Набор требований  |   Office 365 для Windows\*  |  Office для iPad  |  Office 365 для Mac  | Office Online  | Office Online Server  |
|:-----|-----|:-----|:-----|:-----|:-----|
| WordApi 1.3 | Версия 1612 (сборка 7668.1000) или более поздняя| Марта 2017 г., 2.22 или более поздняя | Март 2017 г., 15.32 или более поздняя| Март 2017 г. ||
| WordApi 1.2  | Обновление за декабрь 2015 г., версия 1601 (сборка 6568.1000) или более поздняя | Январь 2016 г., версия 1.18 или более поздняя | Январь 2016 г., версия 15.19 или более поздняя| Сентябрь 2016 г. | |
| WordApi 1.1  | Версия 1509 (сборка 4266.1001) или более поздняя| Январь 2016 г., версия 1.18 или более поздняя | Январь 2016 г., версия 15.19 или более поздняя| Сентябрь 2016 г. | |

> [!NOTE]
> Номер построения Office 2016, установленный с помощью MSI, — 16.0.4266.1001. Эта версия содержит только наборы требований WordApi 1.1.

Статьи и разделы с дополнительными сведениями о версиях, номерах построений и Office Online Server:

- [Номера версий и построений выпусков канала обновления для клиентов Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Какую версию Office я использую?](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [Где можно найти номера версии и построения клиентского приложения Office 365](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [Обзор Office Online Server](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a>Стандартные наборы требований API для Office

Сведения о стандартных наборах требований API см. в статье [Стандартные наборы требований API для Office](office-add-in-requirement-sets.md).

## <a name="whats-new-in-word-javascript-api-13"></a>Новые возможности API JavaScript для Word 1.3 

Ниже перечислены новые дополнения API JavaScript для Word в наборе требований 1.3. 

|Объект| Что нового| Описание|Набор требований| 
|:-----|-----|:----|:----| 
|[application](/javascript/api/word/word.application)|_Метод_ > createDocument(base64File: string) | Создает документ с помощью DOCX-файла с кодировкой base64. Только для чтения.|1.3|
|[body](/javascript/api/word/word.body)|_Связь_ > lists|Получает коллекцию объектов списка в тексте. Только для чтения.|1.3|
|[body](/javascript/api/word/word.body)|_Связь_ > parentBody|Получает родительский текст текста. Например, родительским текстом текста ячейки таблицы может быть заголовок. Только для чтения.|1.3|
|[body](/javascript/api/word/word.body)|_Связь_ > parentSection|Получает родительский раздел текста. Только для чтения.|1.3|
|[body](/javascript/api/word/word.body)|_Связь_ > styleBuiltIn|Получает или задает имя встроенного стиля текста. Используйте это свойство для встроенных стилей, поддерживающих несколько языковых стандартов. Чтобы использовать пользовательские стили или локализованные имена стилей, применяйте свойство style.|1.3|
|[body](/javascript/api/word/word.body)|_Связь_ > tables|Получает коллекцию объектов таблицы в тексте. Только для чтения.|1.3|
|[body](/javascript/api/word/word.body)|_Связь_ > type|Получает тип текста. Поддерживаемые типы: MainDoc, Section, Header, Footer и TableCell. Только для чтения.|1.3|
|[body](/javascript/api/word/word.body)|_Метод_ > getRange(rangeLocation: RangeLocation)|Получает весь текст (либо его начальную или конечную точку) в виде диапазона.|1.3|
|[body](/javascript/api/word/word.body)|_Метод_ > insertTable(rowCount: number, columnCount: number, insertLocation: InsertLocation, values: string)|Вставляет таблицу с указанным количеством строк и столбцов. Возможные значения insertLocation: Start и End.|1.3|
|[breaktype](/javascript/api/word/word.breaktype)|_Связь_ > breaks|Указывает форму разрыва: тип line, page или section. Только для чтения.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Связь_ > lists|Получает коллекцию объектов списка в элемент управления содержимым. Только для чтения.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Связь_ > parentBody|Получает родительский текст элемента управления содержимым. Только для чтения.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Связь_ > parentTable|Получает таблицу, содержащую элемент управления содержимым. Если он находится не в таблице, возвращается пустой объект. Только для чтения.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Связь_ > parentTableCell|Получает ячейку таблицы, содержащую элемент управления содержимым. Если он находится не в ячейке таблицы, возвращается пустой объект. Только для чтения.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Связь_ > styleBuiltIn|Получает или задает имя встроенного стиля для элемента управления содержимым. Используйте это свойство для встроенных стилей, поддерживающих несколько языковых стандартов. Чтобы использовать пользовательские стили или локализованные имена стилей, применяйте свойство style.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Связь_ > subtype|Получает подтип элемента управления содержимым. Поддерживаемые подтипы: RichTextInline, RichTextParagraphs, RichTextTableCell, RichTextTableRow и RichTextTable для элементов управления форматированным текстом. Только для чтения.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Связь_ > tables|Получает коллекцию объектов таблицы в элемент управления содержимым. Только для чтения.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Метод_ > getRange(rangeLocation: RangeLocation)|Получает весь элемент управления содержимым (либо его начальную или конечную точку) в виде диапазона.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Метод_ > getTextRanges (endingMarks: string, trimSpacing: bool)|Получает текстовые диапазоны в элемент управления содержимым с помощью знаков препинания и/или других маркеров конца.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Метод_ > insertTable(rowCount: number, columnCount: number, insertLocation: InsertLocation, values: string)|Вставляет таблицу с указанным количеством строк и столбцов в элемент управления содержимым или рядом с ним. Возможные значения insertLocation: Start, End, Before и After.|1.3|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Метод_ > split(delimiters: string, multiParagraphs: bool, trimDelimiters: bool, trimSpacing: bool)|Разделяет элемент управления содержимым на дочерние диапазоны с помощью разделителей.|1.3|
|[contentControlCollection](/javascript/api/word/word.contentcontrolcollection)|_Метод_ > getByTypes(types: ContentControlType)|Получает элементы управления содержимым с указанными типами и/или подтипами.|1.3|
|[contentControlCollection](/javascript/api/word/word.contentcontrolcollection)|_Метод_ > getFirst()|Получает первый элемент управления содержимым в этой коллекции.|1.3|
|[customProperty](/javascript/api/word/word.customproperty)|_Свойство_ > key|Получает ключ настраиваемого свойства. Только для чтения. |1.3|
|[customProperty](/javascript/api/word/word.customproperty)|_Свойство_ > value|Получает или задает значение настраиваемого свойства.|1.3|
|[customProperty](/javascript/api/word/word.customproperty)|_Связь_ > type|Получает тип значения настраиваемого свойства. Только для чтения.|1.3|
|[customProperty](/javascript/api/word/word.customproperty)|_Метод_ > delete()|Удаляет настраиваемое свойство.|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_Свойство_ > items|Коллекция объектов customProperty. Только для чтения.|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_Метод_ > deleteAll()|Удаляет все настраиваемые свойства в коллекции.|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_Метод_ > getCount()|Получает количество настраиваемых свойств.|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_Метод_ > getItem(key: string)|Получает объект настраиваемого свойства по ключу, указываемому без учета регистра.|1.3|
|[customPropertyCollection](/javascript/api/word/word.custompropertycollection)|_Метод_ > add(key: string, value: object)|Создает или задает настраиваемое свойство.|1.3|
|[document](/javascript/api/word/word.document)|_Связь_ > properties|Получает свойства текущего документа. Только для чтения.|1.3|
|[document](/javascript/api/word/word.document)|_Метод_ > open()|Открывает документ.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Свойство_ > applicationName|Получает имя приложения документа. Только для чтения.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Свойство_ > author|Получает или задает автора документа.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Свойство_ > category|Получает или задает категорию документа.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Свойство_ > comments|Получает или задает комментарии к документу.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Свойство_ > company|Получает или задает организацию документа.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Свойство_ > format|Получает или задает формат документа.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Свойство_ > keywords|Получает или задает ключевые слова документа.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Свойство_ > lastAuthor|Получает или задает последнего автора документа.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Свойство_ > manager|Получает или задает менеджера документа.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Свойство_ > revisionNumber|Получает номер редакции документа. Только для чтения.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Свойство_ > security|Получает защиту документа. Только для чтения.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Свойство_ > subject|Получает или задает тему документа.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Свойство_ > template|Получает шаблон документа. Только для чтения.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Свойство_ > title|Получает или задает заголовок документа.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Свойство_ > creationDate|Получает дату создания документа. Только для чтения.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Связь_ > customProperties|Получает коллекцию настраиваемых свойств документа. Только для чтения.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Связь_ > lastPrintDate|Получает дату последней печати документа. Только для чтения.|1.3|
|[documentProperties](/javascript/api/word/word.documentproperties)|_Связь_ > lastSaveTime|Получает время последнего сохранения документа. Только для чтения.|1.3|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Связь_ > parentTable|Получает таблицу, содержащую встроенный рисунок. Если он находится не в таблице, возвращается пустой объект. Только для чтения.|1.3|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Связь_ > parentTableCell|Получает ячейку таблицы, содержащую встроенный рисунок. Если он находится не в ячейке таблицы, возвращается пустой объект. Только для чтения.|1.3|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Метод_ > getNext()|Получает следующий встроенный рисунок.|1.3|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Метод_ > getRange(rangeLocation: RangeLocation)|Получает рисунок (либо его начальную или конечную точку) в виде диапазона.|1.3|
|[inlinePictureCollection](/javascript/api/word/word.inlinepicturecollection)|_Метод_ > getFirst()|Получает первый встроенный рисунок в коллекции.|1.3|
|[list](/javascript/api/word/word.list)|_Свойство_ > id|Получает идентификатор списка. Только для чтения.|1.3|
|[list](/javascript/api/word/word.list)|_Свойство_ > levelExistences|Проверяет наличие каждого из 9 уровней в списке. Значение true указывает, что уровень существует, то есть на этом уровне имеется по крайней мере один элемент списка. Только для чтения.|1.3|
|[list](/javascript/api/word/word.list)|_Связь_ > levelTypes|Получает типы всех 9 уровней списка. Поддерживаемые типы: Bullet, Number и Picture. Только для чтения.|1.3|
|[list](/javascript/api/word/word.list)|_Связь_ > paragraphs|Получает абзацы в списке. Только для чтения.|1.3|
|[list](/javascript/api/word/word.list)|_Метод_ > getLevelParagraphs(level: number)|Получает абзацы, обнаруженные на указанном уровне списка.|1.3|
|[list](/javascript/api/word/word.list)|_Метод_ > getLevelString(level: number)|Получает маркер, номер или рисунок на указанном уровне в виде строки.|1.3|
|[list](/javascript/api/word/word.list)|_Метод_ > insertParagraph(paragraphText: string, insertLocation: InsertLocation)|Вставляет абзац в указанное расположение. Возможные значения insertLocation: Start, End, Before и After.|1.3|
|[list](/javascript/api/word/word.list)|_Метод_ > setLevelAlignment(level: number, alignment: Alignment)|Задает выравнивание маркера, номера или картинки на указанном уровне списка.|1.3|
|[list](/javascript/api/word/word.list)|_Метод_ > setLevelBullet(level: number, listBullet: ListBullet, charCode: number, fontName: string)|Задает формат маркера на указанном уровне списка. Если задан формат Custom, то параметр charCode является обязательным.|1.3|
|[list](/javascript/api/word/word.list)|_Метод_ > setLevelIndents(level: number, textIndent: float, textIndent: float)|Задает два отступа на указанном уровне списка.|1.3|
|[list](/javascript/api/word/word.list)|_Метод_ > setLevelNumbering(level: number, listNumbering: ListNumbering, formatString: object)|Задает формат нумерации на указанном уровне списка.|1.3|
|[list](/javascript/api/word/word.list)|_Метод_ > setLevelStartingNumber(level: number, startingNumber: number)|Задает начальный номер на указанном уровне списка. Значение по умолчанию: 1.|1.3|
|[listCollection](/javascript/api/word/word.listcollection)|_Свойство_ > items|Коллекция объектов списков. Только для чтения.|1.3|
|[listCollection](/javascript/api/word/word.listcollection)|_Метод_ > getById(id: number)|Получает список по идентификатору.|1.3|
|[listCollection](/javascript/api/word/word.listcollection)|_Метод_ > getFirst()|Получает первый список в коллекции.|1.3|
|[listCollection](/javascript/api/word/word.listcollection)|_Метод_ > getItem(index: number)|Получает объект списка по индексу в коллекции.|1.3|
|[listItem](/javascript/api/word/word.listitem)|_Свойство_ > level|Получает или задает уровень элемента в списке.|1.3|
|[listItem](/javascript/api/word/word.listitem)|_Свойство_ > listString|Получает маркер, номер или рисунок элемента списка в виде строки. Только для чтения.|1.3|
|[listItem](/javascript/api/word/word.listitem)|_Свойство_ > siblingIndex|Получает порядковый номер элемента списка относительно элементов того же уровня. Только для чтения.|1.3|
|[listItem](/javascript/api/word/word.listitem)|_Метод_ > getAncestor(parentOnly: bool)|Получает родительский элемент или ближайшего предка (если родительского элемента нет) для данного элемента списка.|1.3|
|[listItem](/javascript/api/word/word.listitem)|_Метод_ > getDescendants(directChildrenOnly: bool)|Получает всех потомков элемента списка.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Свойство_ > isLastParagraph|Указывает, что абзац является последним в родительском тексте. Только для чтения.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Свойство_ > isListItem|Проверяет, является ли абзац элементом списка. Только для чтения.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Свойство_ > tableNestingLevel|Получает уровень таблицы, содержащей абзац. Если абзац не находится в таблице, возвращается значение 0. Только для чтения.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Связь_ > list|Получает объект List, к которому относится абзац. Если абзац не находится в списке, возвращается пустой объект. Только для чтения.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Связь_ > listItem|Получает объект ListItem для абзаца. Если абзац не является частью списка, возвращается пустой объект. Только для чтения.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Связь_ > parentBody|Получает родительский текст абзаца. Только для чтения.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Связь_ > parentTable|Получает таблицу, содержащую абзац. Если он находится не в таблице, возвращается пустой объект. Только для чтения.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Связь_ > parentTableCell|Получает ячейку таблицы, содержащую абзац. Если он находится не в ячейке таблицы, возвращается пустой объект. Только для чтения.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Связь_ > styleBuiltIn|Получает или задает имя встроенного стиля абзаца. Используйте это свойство для встроенных стилей, поддерживающих несколько языковых стандартов. Чтобы использовать пользовательские стили или локализованные имена стилей, применяйте свойство style.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Метод_ > attachToList(listId: number, level: number)|Позволяет присоединить абзац к существующему списку на указанном уровне. Если присоединить абзац к списку не удается или он уже является элементом списка, метод не выполняется.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Метод_ > detachFromList()|Перемещает абзац за пределы списка (если он является элементом списка).|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Метод_ > getNext()|Получает следующий абзац.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Метод_ > getPrevious()|Получает предыдущий абзац.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Метод_ > getRange(rangeLocation: RangeLocation)|Получает весь абзац (либо его начальную или конечную точку) в виде диапазона.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Метод_ > getTextRanges (endingMarks: string, trimSpacing: bool)|Получает текстовые диапазоны абзаца с помощью знаков препинания и/или других маркеров конца.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Метод_ > insertTable(rowCount: number, columnCount: number, insertLocation: InsertLocation, values: string)|Вставляет таблицу с указанным количеством строк и столбцов. Возможные значения insertLocation: Before и After.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Метод_ > split(delimiters: string, multiParagraphs: bool, trimDelimiters: bool, trimSpacing: bool)|Разделяет абзац на дочерние диапазоны с помощью разделителей.|1.3|
|[paragraph](/javascript/api/word/word.paragraph)|_Метод_ > startNewList()|Создает список, начинающийся с данного абзаца. Если абзац уже является элементом списка, метод не выполняется.|1.3|
|[paragraphCollection](/javascript/api/word/word.paragraphcollection)|_Метод_ > getFirst()|Получает первый абзац в коллекции.|1.3|
|[paragraphCollection](/javascript/api/word/word.paragraphcollection)|_Метод_ > getLast()|Получает последний абзац в коллекции.|1.3|
|[range](/javascript/api/word/word.range)|_Свойство_ > hyperlink|Получает первую гиперссылку в диапазоне или задает для диапазона гиперссылку. При добавлении в диапазон новой гиперссылки из него удаляются все имеющиеся гиперссылки. Используйте символ новой строки ('\n'), чтобы отделить часть адреса от необязательной части расположения.|1.3|
|[range](/javascript/api/word/word.range)|_Свойство_ > isEmpty|Проверяет, является ли длина диапазона нулевой. Только для чтения.|1.3|
|[range](/javascript/api/word/word.range)|_Связь_ > lists|Получает коллекцию объектов списков в диапазоне. Только для чтения.|1.3|
|[range](/javascript/api/word/word.range)|_Связь_ > parentBody|Получает родительский текст диапазона. Только для чтения.|1.3|
|[range](/javascript/api/word/word.range)|_Связь_ > parentTable|Получает таблицу, содержащую диапазон. Если он находится не в таблице, возвращается пустой объект. Только для чтения.|1.3|
|[range](/javascript/api/word/word.range)|_Связь_ > parentTableCell|Получает ячейку таблицы, содержащую диапазон. Если он находится не в ячейке таблицы, возвращается пустой объект. Только для чтения.|1.3|
|[range](/javascript/api/word/word.range)|_Связь_ > styleBuiltIn|Получает или задает имя встроенного стиля диапазона. Используйте это свойство для встроенных стилей, поддерживающих несколько языковых стандартов. Чтобы использовать пользовательские стили или локализованные имена стилей, применяйте свойство style.|1.3|
|[range](/javascript/api/word/word.range)|_Связь_ > tables|Получает коллекцию объектов таблиц в диапазоне. Только для чтения.|1.3|
|[range](/javascript/api/word/word.range)|_Метод_ > compareLocationWith(range: Range)|Сравнивает расположение данного диапазона с расположением другого диапазона.|1.3|
|[range](/javascript/api/word/word.range)|_Метод_ > expandTo(range: Range)|Получает новый диапазон, который простирается в том или ином направлении от данного диапазона и перекрывает другой диапазон. Данный диапазон не меняется.|1.3|
|[range](/javascript/api/word/word.range)|_Метод_ > getHyperlinkRanges()|Получает дочерние диапазоны гиперссылок в данном диапазоне.|1.3|
|[range](/javascript/api/word/word.range)|_Метод_ > getNextTextRange(endingMarks: string, trimSpacing: bool)|Получает следующий текстовый диапазон с помощью знаков препинания и/или других маркеров конца.|1.3|
|[range](/javascript/api/word/word.range)|_Метод_ > getRange(rangeLocation: RangeLocation)|Клонирует диапазон либо получает его начальную или конечную точку в виде нового диапазона.|1.3|
|[range](/javascript/api/word/word.range)|_Метод_ > getTextRanges (endingMarks: string, trimSpacing: bool)|Получает дочерние текстовые диапазоны данного диапазона с помощью знаков препинания и/или других маркеров конца.|1.3|
|[range](/javascript/api/word/word.range)|_Метод_ > insertTable(rowCount: number, columnCount: number, insertLocation: InsertLocation, values: string)|Вставляет таблицу с указанным количеством строк и столбцов. Возможные значения insertLocation: Before и After.|1.3|
|[range](/javascript/api/word/word.range)|_Метод_ > intersectWith(range: Range)|Получает новый диапазон, представляющий собой пересечение данного диапазона с другим. Данный диапазон не меняется.|1.3|
|[range](/javascript/api/word/word.range)|_Метод_ > split(delimiters: string, multiParagraphs: bool, trimDelimiters: bool, trimSpacing: bool)|Разделяет диапазон на дочерние диапазоны с помощью разделителей.|1.3|
|[rangeCollection](/javascript/api/word/word.rangecollection)|_Свойство_ > items|Коллекция объектов диапазонов. Только для чтения.|1.3|
|[rangeCollection](/javascript/api/word/word.rangecollection)|_Метод_ > getFirst()|Получает первый диапазон в коллекции.|1.3|
|[rangeCollection](/javascript/api/word/word.rangecollection)|_Метод_ > getItem(index: number)|Получает объект диапазона по индексу в коллекции.|1.3|
|[RequestContext](/javascript/api/word/word.requestcontext)|_Метод_ > load(object: object, option: object)|Заполняет объект прокси, созданный на уровне JavaScrypt, свойством и параметрами, которые указаны в параметре. |1.3|
|[RequestContext](/javascript/api/word/word.requestcontext)|_Метод_ > sync()|Отправляет очередь запросов в Word и возвращает объект Promise, который может использоваться для построения цепочки дальнейших действий.|1.3|
|[section](/javascript/api/word/word.section)|_Метод_ > getNext()|Получает следующий раздел.|1.3|
|[sectionCollection](/javascript/api/word/word.sectioncollection)|_Метод_ > getFirst()|Получает первый раздел в коллекции.|1.3|
|[table](/javascript/api/word/word.table)|_Свойство_ > headerRowCount|Получает и задает количество строк заголовков.|1.3|
|[table](/javascript/api/word/word.table)|_Свойство_ > height|Получает высоту таблицы в точках. Только для чтения.|1.3|
|[table](/javascript/api/word/word.table)|_Свойство_ > isUniform|Указывает, однородны ли все строки таблицы. Только для чтения.|1.3|
|[table](/javascript/api/word/word.table)|_Свойство_ > nestingLevel|Получает уровень вложения таблицы. Таблицам верхнего уровня соответствует значение 1. Только для чтения.|1.3|
|[table](/javascript/api/word/word.table)|_Свойство_ > rowCount|Получает количество строк в таблице. Только для чтения.|1.3|
|[table](/javascript/api/word/word.table)|_Свойство_ > shadingColor|Получает и задает цвет заливки.|1.3|
|[table](/javascript/api/word/word.table)|_Свойство_ > style|Получает или задает имя стиля для таблицы. Используйте это свойство для пользовательских стилей и локализованных имен стилей. Чтобы использовать встроенные стили, поддерживающие несколько языковых стандартов, применяйте свойство styleBuiltIn.|1.3|
|[table](/javascript/api/word/word.table)|_Свойство_ > styleBandedColumns|Получает и задает значение, указывающее, есть ли в таблице чередующиеся столбцы.|1.3|
|[table](/javascript/api/word/word.table)|_Свойство_ > styleBandedRows|Получает и задает значение, указывающее, есть ли в таблице чередующиеся строки.|1.3|
|[table](/javascript/api/word/word.table)|_Свойство_ > styleFirstColumn|Получает и задает значение, указывающее, применен ли специальный стиль к первому столбцу таблицы.|1.3|
|[table](/javascript/api/word/word.table)|_Свойство_ > styleLastColumn|Получает и задает значение, указывающее, применен ли специальный стиль к последнему столбцу таблицы.|1.3|
|[table](/javascript/api/word/word.table)|_Свойство_ > styleTotalRow|Получает и задает значение, указывающее, применен ли специальный стиль к строке итогов (последней строке) таблицы.|1.3|
|[table](/javascript/api/word/word.table)|_Свойство_ > values|Получает и задает текстовые значения в таблице в виде двумерного массива JavaScript.|1.3|
|[table](/javascript/api/word/word.table)|_Свойство_ > width|Получает и задает ширину таблицы в точках.|1.3|
|[table](/javascript/api/word/word.table)|_Связь_ > font|Получает шрифт. Используйте эту связь, чтобы получать и задавать имя, размер, цвет и другие свойства шрифта. Только для чтения.|1.3|
|[table](/javascript/api/word/word.table)|_Связь_ > horizontalAlignment|Получает и задает горизонтальное выравнивание для каждой ячейки в таблице. Возможные значения: left, centered, right и justified.|1.3|
|[table](/javascript/api/word/word.table)|_Связь_ > paragraphAfter|Получает абзац после таблицы. Только для чтения.|1.3|
|[table](/javascript/api/word/word.table)|_Связь_ > paragraphBefore|Получает абзац перед таблицей. Только для чтения.|1.3|
|[table](/javascript/api/word/word.table)|_Связь_ > parentBody|Получает родительский текст таблицы. Только для чтения.|1.3|
|[table](/javascript/api/word/word.table)|_Связь_ > parentContentControl|Получает элемент управления содержимым, содержащий таблицу. Только для чтения.|1.3|
|[table](/javascript/api/word/word.table)|_Связь_ > parentTable|Получает таблицу, которая содержит данную таблицу. Если она находится не в таблице, возвращается пустой объект. Только для чтения.|1.3|
|[table](/javascript/api/word/word.table)|_Связь_ > parentTableCell|Получает ячейку таблицы, содержащую данную таблицу. Если она находится не в ячейке таблицы, возвращается пустой объект. Только для чтения.|1.3|
|[table](/javascript/api/word/word.table)|_Связь_ > rows|Получает все строки таблицы. Только для чтения.|1.3|
|[table](/javascript/api/word/word.table)|_Связь_ > styleBuiltIn|Получает или задает имя встроенного стиля таблицы. Используйте это свойство для встроенных стилей, поддерживающих несколько языковых стандартов. Чтобы использовать пользовательские стили или локализованные имена стилей, применяйте свойство style.|1.3|
|[table](/javascript/api/word/word.table)|_Связь_ > tables|Получает дочерние таблицы, вложенные на один уровень ниже. Только для чтения.|1.3|
|[table](/javascript/api/word/word.table)|_Связь_ > verticalAlignment|Получает и задает вертикальное выравнивание для каждой ячейки в таблице. Возможные значения: top, center и bottom.|1.3|
|[table](/javascript/api/word/word.table)|_Метод_ > addColumns(insertLocation: InsertLocation, columnCount: number, values: string)|Добавляет столбцы в начале или в конце таблицы, используя первый или последний из имеющихся столбцов в качестве шаблона. Применим к однородным таблицам. Строковые значения (если они указаны) добавляются в новые строки.|1.3|
|[table](/javascript/api/word/word.table)|_Метод_ > addRows(insertLocation: InsertLocation, rowCount: number, values: string)|Добавляет строки в начале или в конце таблицы, используя первую или последнюю из имеющихся строк в качестве шаблона. Строковые значения (если они указаны) добавляются в новые строки.|1.3|
|[table](/javascript/api/word/word.table)|_Метод_ > autoFitContents()|Автоматически подбирает ширину столбцов таблицы в соответствии с их содержимым.|1.3|
|[table](/javascript/api/word/word.table)|_Метод_ > autoFitWindow()|Автоматически подбирает ширину столбцов таблицы в соответствии с шириной окна.|1.3|
|[table](/javascript/api/word/word.table)|_Метод_ > clear()|Очищает содержимое таблицы.|1.3|
|[table](/javascript/api/word/word.table)|_Метод_ > delete()|Удаляет всю таблицу.|1.3|
|[table](/javascript/api/word/word.table)|_Метод_ > deleteColumns(columnIndex: number, columnCount: number)|Удаляет определенные столбцы. Применим к однородным таблицам.|1.3|
|[table](/javascript/api/word/word.table)|_Метод_ > deleteRows(rowIndex: number, rowCount: number)|Удаляет определенные строки.|1.3|
|[table](/javascript/api/word/word.table)|_Метод_ > distributeColumns()|Равномерно распределяет ширину столбцов.|1.3|
|[table](/javascript/api/word/word.table)|_Метод_ > distributeRows()|Равномерно распределяет высоту строк.|1.3|
|[table](/javascript/api/word/word.table)|_Метод_ > getBorder(borderLocation: BorderLocation)|Получает стиль указанной границы.|1.3|
|[table](/javascript/api/word/word.table)|_Метод_ > getCell(rowIndex: number, cellIndex: number)|Получает ячейку таблицы в указанных строке и столбце.|1.3|
|[table](/javascript/api/word/word.table)|_Метод_ > getCellPadding(cellPaddingLocation: CellPaddingLocation)|Получает размер поля ячейки в точках.|1.3|
|[table](/javascript/api/word/word.table)|_Метод_ > getNext()|Получает следующую таблицу.|1.3|
|[table](/javascript/api/word/word.table)|_Метод_ > getRange(rangeLocation: RangeLocation)|Получает диапазон, содержащий данную таблицу, либо диапазон в начале или в конце таблицы.|1.3|
|[table](/javascript/api/word/word.table)|_Метод_ > insertContentControl()|Вставляет в таблицу элемент управления содержимым.|1.3|
|[table](/javascript/api/word/word.table)|_Метод_ > insertParagraph(paragraphText: string, insertLocation: InsertLocation)|Вставляет абзац в указанном расположении. Возможные значения insertLocation: Before и After.|1.3|
|[table](/javascript/api/word/word.table)|_Метод_ > insertTable(rowCount: number, columnCount: number, insertLocation: InsertLocation, values: string)|Вставляет таблицу с указанным количеством строк и столбцов. Возможные значения insertLocation: Before и After.|1.3|
|[table](/javascript/api/word/word.table)|_Метод_ > search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)|Выполняет поиск с помощью указанного объекта searchOptions в области объекта таблицы. Результат поиска — коллекция объектов диапазонов.|1.3|
|[table](/javascript/api/word/word.table)|_Метод_ > select(selectionMode: SelectionMode)|Выбирает таблицу либо позицию в начале или в конце таблицы, а затем переходит к ней в Word.|1.3|
|[table](/javascript/api/word/word.table)|_Метод_ > setCellPadding(cellPaddingLocation: CellPaddingLocation, cellPadding: float)|Задает размер поля ячейки в точках.|1.3|
|[tableBorder](/javascript/api/word/word.tableborder)|_Свойство_ > color|Получает или задает цвет границы таблицы по шестнадцатеричному значению или имени.|1.3|
|[tableBorder](/javascript/api/word/word.tableborder)|_Свойство_ > width|Получает или задает ширину границы таблицы в точках. Не применимо к типам границ с фиксированной шириной.|1.3|
|[tableBorder](/javascript/api/word/word.tableborder)|_Связь_ > type|Получает или задает тип границы таблицы.|1.3|
|[TableCell](/javascript/api/word/word.tablecell)|_Свойство_ > cellIndex|Получает индекс ячейки в строке. Только для чтения.|1.3|
|[TableCell](/javascript/api/word/word.tablecell)|_Свойство_ > columnWidth|Получает и задает ширину столбца ячейки в точках. Применимо к однородным таблицам.|1.3|
|[TableCell](/javascript/api/word/word.tablecell)|_Свойство_ > rowIndex|Получает индекс строки ячейки в таблице. Только для чтения.|1.3|
|[TableCell](/javascript/api/word/word.tablecell)|_Свойство_ > shadingColor|Получает или задает цвет заливки ячейки. Цвет задается в формате "#RRGGBB" или по имени.|1.3|
|[TableCell](/javascript/api/word/word.tablecell)|_Свойство_ > value|Получает и задает текст ячейки.|1.3|
|[TableCell](/javascript/api/word/word.tablecell)|_Свойство_ > width|Получает ширину ячейки в точках. Только для чтения.|1.3|
|[TableCell](/javascript/api/word/word.tablecell)|_Связь_ > body|Получает объект текста ячейки. Только для чтения.|1.3|
|[TableCell](/javascript/api/word/word.tablecell)|_Связь_ > horizontalAlignment|Получает и задает горизонтальное выравнивание ячейки. Возможные значения: left, centered, right и justified.|1.3|
|[TableCell](/javascript/api/word/word.tablecell)|_Связь_ > parentRow|Получает родительскую строку ячейки. Только для чтения.|1.3|
|[TableCell](/javascript/api/word/word.tablecell)|_Связь_ > parentTable|Получает родительскую таблицу ячейки. Только для чтения.|1.3|
|[TableCell](/javascript/api/word/word.tablecell)|_Связь_ > verticalAlignment|Получает и задает вертикальное выравнивание ячейки. Возможные значения: top, center и bottom.|1.3|
|[TableCell](/javascript/api/word/word.tablecell)|_Метод_ > deleteColumn()|Удаляет столбец, содержащий данную ячейку. Применим к однородным таблицам.|1.3|
|[TableCell](/javascript/api/word/word.tablecell)|_Метод_ > deleteRow()|Удаляет строку, содержащую данную ячейку.|1.3|
|[TableCell](/javascript/api/word/word.tablecell)|_Метод_ > getBorder(borderLocation: BorderLocation)|Получает стиль указанной границы.|1.3|
|[TableCell](/javascript/api/word/word.tablecell)|_Метод_ > getCellPadding(cellPaddingLocation: CellPaddingLocation)|Получает размер поля ячейки в точках.|1.3|
|[TableCell](/javascript/api/word/word.tablecell)|_Метод_ > getNext()|Получает следующую ячейку.|1.3|
|[TableCell](/javascript/api/word/word.tablecell)|_Метод_ > addColumns(insertLocation: InsertLocation, columnCount: number, values: string)|Добавляет столбцы слева или справа от ячейки, используя столбец этой ячейки в качестве шаблона. Применим к однородным таблицам. Строковые значения (если они указаны) добавляются в новые строки.|1.3|
|[TableCell](/javascript/api/word/word.tablecell)|_Метод_ > addRows(insertLocation: InsertLocation, rowCount: number, values: string)|Вставляет строки над ячейкой или под ней, используя строку этой ячейки в качестве шаблона. Строковые значения (если они указаны) добавляются в новые строки.|1.3|
|[TableCell](/javascript/api/word/word.tablecell)|_Метод_ > setCellPadding(cellPaddingLocation: CellPaddingLocation, cellPadding: float)|Задает размер поля ячейки в точках.|1.3|
|[tableCellCollection](/javascript/api/word/word.tablecellcollection)|_Свойство_ > items|Коллекция объектов TableCell. Только для чтения.|1.3|
|[tableCellCollection](/javascript/api/word/word.tablecellcollection)|_Метод_ > getFirst()|Получает первую ячейку таблицы в коллекции.|1.3|
|[tableCellCollection](/javascript/api/word/word.tablecellcollection)|_Метод_ > getItem(index: number)|Получает объект ячейки таблицы по индексу в коллекции.|1.3|
|[tableCollection](/javascript/api/word/word.tablecollection)|_Свойство_ > items|Коллекция объектов таблиц. Только для чтения.|1.3|
|[tableCollection](/javascript/api/word/word.tablecollection)|_Метод_ > getFirst()|Получает первую таблицу в коллекции.|1.3|
|[tableCollection](/javascript/api/word/word.tablecollection)|_Метод_ > getItem(index: number)|Получает объект таблицы по индексу в коллекции.|1.3|
|[TableRow](/javascript/api/word/word.tablerow)|_Свойство_ > cellCount|Получает количество ячеек в строке. Только для чтения.|1.3|
|[TableRow](/javascript/api/word/word.tablerow)|_Свойство_ > isHeader|Проверяет, является ли элемент строкой заголовков. Только для чтения. Чтобы задать количество строк заголовков, используйте свойство HeaderRowCount объекта Table. Только для чтения.|1.3|
|[TableRow](/javascript/api/word/word.tablerow)|_Свойство_ > preferredHeight|Получает и задает предпочитаемую высоту строки в точках.|1.3|
|[TableRow](/javascript/api/word/word.tablerow)|_Свойство_ > rowIndex|Получает индекс строки в родительской таблице. Только для чтения.|1.3|
|[TableRow](/javascript/api/word/word.tablerow)|_Свойство_ > shadingColor|Получает и задает цвет заливки.|1.3|
|[TableRow](/javascript/api/word/word.tablerow)|_Свойство_ > values|Получает и задает текстовые значения в строке в виде одномерного массива JavaScript.|1.3|
|[TableRow](/javascript/api/word/word.tablerow)|_Связь_ > cells|Получает ячейки. Только для чтения.|1.3|
|[TableRow](/javascript/api/word/word.tablerow)|_Связь_ > font|Получает шрифт. Используйте эту связь, чтобы получать и задавать имя, размер, цвет и другие свойства шрифта. Только для чтения.|1.3|
|[TableRow](/javascript/api/word/word.tablerow)|_Связь_ > horizontalAlignment|Получает и задает горизонтальное выравнивание для каждой ячейки в строке. Возможные значения: left, centered, right и justified.|1.3|
|[TableRow](/javascript/api/word/word.tablerow)|_Связь_ > parentTable|Получает родительскую таблицу. Только для чтения.|1.3|
|[TableRow](/javascript/api/word/word.tablerow)|_Связь_ > verticalAlignment|Получает и задает вертикальное выравнивание ячеек в строке. Возможные значения: top, center и bottom.|1.3|
|[TableRow](/javascript/api/word/word.tablerow)|_Метод_ > clear()|Очищает содержимое строки.|1.3|
|[TableRow](/javascript/api/word/word.tablerow)|_Метод_ > delete()|Удаляет всю строку.|1.3|
|[TableRow](/javascript/api/word/word.tablerow)|_Метод_ > getBorder(borderLocation: BorderLocation)|Получает стиль границ ячеек в строке.|1.3|
|[TableRow](/javascript/api/word/word.tablerow)|_Метод_ > getCellPadding(cellPaddingLocation: CellPaddingLocation)|Получает размер поля ячейки в точках.|1.3|
|[TableRow](/javascript/api/word/word.tablerow)|_Метод_ > getNext()|Получает следующую строку.|1.3|
|[TableRow](/javascript/api/word/word.tablerow)|_Метод_ > addRows(insertLocation: InsertLocation, rowCount: number, values: string)|Вставляет строки, используя данную строку в качестве шаблона. Если указаны значения, они вставляются в новые строки.|1.3|
|[TableRow](/javascript/api/word/word.tablerow)|_Метод_ > search(searchText: string, searchOptions: ParamTypeStrings.SearchOptions)|Выполняет поиск с помощью указанного параметра searchOptions в области строки. Результат поиска — коллекция объектов диапазонов.|1.3|
|[TableRow](/javascript/api/word/word.tablerow)|_Метод_ > select(selectionMode: SelectionMode)|Выбирает строку и переходит к ней в пользовательском интерфейсе Word.|1.3|
|[TableRow](/javascript/api/word/word.tablerow)|_Метод_ > setCellPadding(cellPaddingLocation: CellPaddingLocation, cellPadding: float)|Задает размер поля ячейки в точках.|1.3|
|[tableRowCollection](/javascript/api/word/word.tablerowcollection)|_Свойство_ > items|Коллекция объектов TableRow. Только для чтения.|1.3|
|[tableRowCollection](/javascript/api/word/word.tablerowcollection)|_Метод_ > getFirst()|Получает первую строку в коллекции.|1.3|
|[tableRowCollection](/javascript/api/word/word.tablerowcollection)|_Метод_ > getItem(index: number)|Получает объект строки таблицы по индексу в коллекции.|1.3|


## <a name="whats-new-in-word-javascript-api-12"></a>Новые возможности API JavaScript для Word 1.2

Ниже перечислены новые возможности API JavaScript для Word в наборе требований 1.2. 

|Объект| Что нового| Описание|Набор требований|
|:-----|-----|:----|:----|
|[contentControl](/javascript/api/word/word.contentcontrol)|_Метод_ > insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)|Вставляет встроенный рисунок в элемент управления содержимым в указанном расположении. Возможные значения insertLocation: Replace, Start и End.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Связь_ > paragraph|Получает родительский абзац, который содержит встроенный рисунок. Только для чтения.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Метод_ > delete()|Удаляет встроенный рисунок из документа.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Метод_ > insertBreak(breakType: BreakType, insertLocation: InsertLocation)|Вставляет разрыв в указанном расположении в основном документе. Возможные значения insertLocation: Before и After.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Метод_ > insertFileFromBase64(base64File: string, insertLocation: InsertLocation)|Вставляет документ в указанном расположении. Возможные значения insertLocation: Before и After.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Метод_ > insertHtml(html: string, insertLocation: InsertLocation)|Вставляет HTML-код в заданном расположении. Возможные значения insertLocation: Before и After.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Метод_ > insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)|Вставляет встроенный рисунок в указанном расположении. Возможные значения InsertLocation: Replace, Before и After.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Метод_ > insertOoxml(ooxml: string, insertLocation: InsertLocation)|Вставляет OOXML-код в заданном расположении.  Возможные значения insertLocation: Before и After.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Метод_ > insertParagraph(paragraphText: string, insertLocation: InsertLocation)|Вставляет абзац в указанном расположении. Возможные значения insertLocation: Before и After.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Метод_ > insertText(text: string, insertLocation: InsertLocation)|Вставляет текст в заданном расположении. Возможные значения insertLocation: Before и After.|1.2|
|[inlinePicture](/javascript/api/word/word.inlinepicture)|_Метод_ > select(selectionMode: SelectionMode)|Выбирает встроенный рисунок. При этом Word переходит к выделенному объекту.|1.2|
|[range](/javascript/api/word/word.range)|_Связь_ > inlinePictures|Получает коллекцию объектов встроенных рисунков в диапазоне. Только для чтения.|1.2|
|[range](/javascript/api/word/word.range)|_Метод_ > insertInlinePictureFromBase64(base64EncodedImage: string, insertLocation: InsertLocation)|Вставляет рисунок в указанном расположении. Возможные значения insertLocation: Replace, Start, End, Before и After.|1.2|

## <a name="word-javascript-api-11"></a>API JavaScript для Word 1.1

API JavaScript для Word 1.1 — это первая версия API. Дополнительные сведения об API см. в разделах справки по [API JavaScript для Word](/javascript/api/word). 

## <a name="see-also"></a>См. также

- [Версии Office и наборы обязательных элементов](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [Указание ведущих приложений Office и обязательных элементов API](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [XML-манифест надстроек Office](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
