---
title: API предварительного просмотра Word JavaScript
description: Сведения о предстоящих API JavaScript Word
ms.date: 11/09/2020
ms.prod: word
ms.localizationpriority: medium
ms.openlocfilehash: c6aa7b8107e0443091f876baa8bd66ccb8db7061
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59154032"
---
# <a name="word-javascript-preview-apis"></a>API предварительного просмотра Word JavaScript

Новые API JavaScript Word сначала вводятся в "предварительную версию", а затем становятся частью определенного набора требований с номерами после достаточного тестирования и получения отзывов пользователей.

[!INCLUDE [Information about using Word preview APIs](../../includes/word-preview-apis-note.md)]
[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

## <a name="api-list"></a>Список API

В следующей таблице перечислены API Word JavaScript, которые в настоящее время находятся в предварительном просмотре. Чтобы просмотреть полный список всех API JavaScript Word (включая API предварительного просмотра и ранее выпущенные API), см. все API [Word JavaScript.](/javascript/api/word?view=word-js-preview&preserve-view=true)

| Класс | Поля | Описание |
|:---|:---|:---|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[onDataChanged](/javascript/api/word/word.contentcontrol#ondatachanged)|Происходит при смене данных в области управления контентом.|
||[onDeleted](/javascript/api/word/word.contentcontrol#ondeleted)|Происходит при удалении управления контентом.|
||[onSelectionChanged](/javascript/api/word/word.contentcontrol#onselectionchanged)|Возникает при смене выбора в области управления контентом.|
|[ContentControlEventArgs](/javascript/api/word/word.contentcontroleventargs)|[contentControl](/javascript/api/word/word.contentcontroleventargs#contentcontrol)|Объект, который поднял событие.|
||[eventType](/javascript/api/word/word.contentcontroleventargs#eventtype)|Тип события.|
|[CustomXmlPart](/javascript/api/word/word.customxmlpart)|[delete()](/javascript/api/word/word.customxmlpart#delete--)|Удаляет пользовательскую XML-часть.|
||[deleteAttribute (xpath: string, namespaceMappings: any, name: string)](/javascript/api/word/word.customxmlpart#deleteattribute-xpath--namespacemappings--name-)|Удаляет атрибут с заданным именем из элемента, идентифицированного xpath.|
||[deleteElement (xpath: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#deleteelement-xpath--namespacemappings-)|Удаляет элемент, идентифицированный xpath.|
||[getXml()](/javascript/api/word/word.customxmlpart#getxml--)|Получает полное XML-содержимое пользовательской части XML.|
||[insertAttribute (xpath: string, namespaceMappings: any, name: string, value: string)](/javascript/api/word/word.customxmlpart#insertattribute-xpath--namespacemappings--name--value-)|Вставляет атрибут с заданным именем и значением в элемент, идентифицированный xpath.|
||[insertElement (xpath: string, xml: string, namespaceMappings: any, index?: number)](/javascript/api/word/word.customxmlpart#insertelement-xpath--xml--namespacemappings--index-)|Вставляет данный XML в родительский элемент, идентифицированный xpath в индексе положения ребенка.|
||[query(xpath: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#query-xpath--namespacemappings-)|Запрашивает XML-содержимое пользовательской части XML.|
||[id](/javascript/api/word/word.customxmlpart#id)|Получает ID пользовательской части XML.|
||[namespaceUri](/javascript/api/word/word.customxmlpart#namespaceuri)|Получает URI пространства имен пользовательской части XML.|
||[setXml (xml: string)](/javascript/api/word/word.customxmlpart#setxml-xml-)|Задает полное XML-содержимое пользовательской части XML.|
||[updateAttribute(xpath: string, namespaceMappings: any, name: string, value: string)](/javascript/api/word/word.customxmlpart#updateattribute-xpath--namespacemappings--name--value-)|Обновляет значение атрибута с заданным именем элемента, идентифицированного xpath.|
||[updateElement (xpath: string, xml: string, namespaceMappings: any)](/javascript/api/word/word.customxmlpart#updateelement-xpath--xml--namespacemappings-)|Обновляет XML элемента, идентифицированного xpath.|
|[CustomXmlPartCollection](/javascript/api/word/word.customxmlpartcollection)|[add(xml: string)](/javascript/api/word/word.customxmlpartcollection#add-xml-)|Добавляет в документ новую настраиваемую часть XML.|
||[getByNamespace (namespaceUri: string)](/javascript/api/word/word.customxmlpartcollection#getbynamespace-namespaceuri-)|Получает новую ограниченную коллекцию пользовательских XML-частей, пространства имен которых совпадают с указанным пространством имен.|
||[getCount()](/javascript/api/word/word.customxmlpartcollection#getcount--)|Возвращает число элементов в коллекции.|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartcollection#getitem-id-)|Получает пользовательскую XML-часть по идентификатору.|
||[getItemOrNullObject(id: строка)](/javascript/api/word/word.customxmlpartcollection#getitemornullobject-id-)|Получает пользовательскую XML-часть по идентификатору.|
||[items](/javascript/api/word/word.customxmlpartcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[CustomXmlPartScopedCollection](/javascript/api/word/word.customxmlpartscopedcollection)|[getCount()](/javascript/api/word/word.customxmlpartscopedcollection#getcount--)|Возвращает число элементов в коллекции.|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#getitem-id-)|Получает пользовательскую XML-часть по идентификатору.|
||[getItemOrNullObject(id: строка)](/javascript/api/word/word.customxmlpartscopedcollection#getitemornullobject-id-)|Получает пользовательскую XML-часть по идентификатору.|
||[getOnlyItem()](/javascript/api/word/word.customxmlpartscopedcollection#getonlyitem--)|Если коллекция содержит ровно один элемент, этот метод возвращает его.|
||[getOnlyItemOrNullObject()](/javascript/api/word/word.customxmlpartscopedcollection#getonlyitemornullobject--)|Если коллекция содержит ровно один элемент, этот метод возвращает его.|
||[items](/javascript/api/word/word.customxmlpartscopedcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Document](/javascript/api/word/word.document)|[deleteBookmark (имя: строка)](/javascript/api/word/word.document#deletebookmark-name-)|Удаляет закладки, если она существует, из документа.|
||[getBookmarkRange (имя: строка)](/javascript/api/word/word.document#getbookmarkrange-name-)|Получает диапазон закладок.|
||[getBookmarkRangeOrNullObject (имя: строка)](/javascript/api/word/word.document#getbookmarkrangeornullobject-name-)|Получает диапазон закладок.|
||[customXmlParts](/javascript/api/word/word.document#customxmlparts)|Получает настраиваемые XML-части в документе.|
||[onContentControlAdded](/javascript/api/word/word.document#oncontentcontroladded)|Возникает при добавлении управления контентом.|
||[settings](/javascript/api/word/word.document#settings)|Получает параметры надстройки в документе.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[deleteBookmark (имя: строка)](/javascript/api/word/word.documentcreated#deletebookmark-name-)|Удаляет закладки, если она существует, из документа.|
||[getBookmarkRange (имя: строка)](/javascript/api/word/word.documentcreated#getbookmarkrange-name-)|Получает диапазон закладок.|
||[getBookmarkRangeOrNullObject (имя: строка)](/javascript/api/word/word.documentcreated#getbookmarkrangeornullobject-name-)|Получает диапазон закладок.|
||[customXmlParts](/javascript/api/word/word.documentcreated#customxmlparts)|Получает настраиваемые XML-части в документе.|
||[settings](/javascript/api/word/word.documentcreated#settings)|Получает параметры надстройки в документе.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[imageFormat](/javascript/api/word/word.inlinepicture#imageformat)|Получает формат inline image.|
|[Перечисление](/javascript/api/word/word.list)|[getLevelFont (уровень: номер)](/javascript/api/word/word.list#getlevelfont-level-)|Получает шрифт пули, номера или изображения на указанном уровне в списке.|
||[getLevelPicture(level: number)](/javascript/api/word/word.list#getlevelpicture-level-)|Получает кодированное представление строки base64 на указанном уровне в списке.|
||[resetLevelFont (уровень: номер, resetFontName?: boolean)](/javascript/api/word/word.list#resetlevelfont-level--resetfontname-)|Сброс шрифта пули, номера или изображения на указанном уровне в списке.|
||[setLevelPicture(level: number, base64EncodedImage?: string)](/javascript/api/word/word.list#setlevelpicture-level--base64encodedimage-)|Задает изображение на указанном уровне в списке.|
|[Range](/javascript/api/word/word.range)|[getBookmarks (includeHidden?: boolean, includeAdjacent?: boolean)](/javascript/api/word/word.range#getbookmarks-includehidden--includeadjacent-)|Получает имена всех закладки или перекрывает диапазон.|
||[insertBookmark (имя: строка)](/javascript/api/word/word.range#insertbookmark-name-)|Вставляет закладки в диапазоне.|
|[Параметр](/javascript/api/word/word.setting)|[delete()](/javascript/api/word/word.setting#delete--)|Удаляет параметр.|
||[key](/javascript/api/word/word.setting#key)|Получает ключ параметра.|
||[value](/javascript/api/word/word.setting#value)|Получает или задает значение параметра.|
|[SettingCollection](/javascript/api/word/word.settingcollection)|[add(key: string, value: any)](/javascript/api/word/word.settingcollection#add-key--value-)|Создает новый параметр или задает существующий параметр.|
||[deleteAll()](/javascript/api/word/word.settingcollection#deleteall--)|Удаляет все параметры в этой надстройки.|
||[getCount()](/javascript/api/word/word.settingcollection#getcount--)|Получает количество параметров.|
||[getItem(key: string)](/javascript/api/word/word.settingcollection#getitem-key-)|Получает объект параметра по его ключу, который является чувствительным к делу.|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.settingcollection#getitemornullobject-key-)|Получает объект параметра по его ключу, который является чувствительным к делу.|
||[items](/javascript/api/word/word.settingcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Table](/javascript/api/word/word.table)|[mergeCells(topRow: number, firstCell: number, bottomRow: number, lastCell: number)](/javascript/api/word/word.table#mergecells-toprow--firstcell--bottomrow--lastcell-)|Объединяет ячейки, ограниченные включительно первой и последней ячейками.|
|[TableCell](/javascript/api/word/word.tablecell)|[split(rowCount: number, columnCount: number)](/javascript/api/word/word.tablecell#split-rowcount--columncount-)|Разделяет ячейку на указанное количество строк и столбцов.|
|[TableRow](/javascript/api/word/word.tablerow)|[insertContentControl()](/javascript/api/word/word.tablerow#insertcontentcontrol--)|Вставляет управление контентом в строку.|
||[merge()](/javascript/api/word/word.tablerow#merge--)|Сливает строку в одну ячейку.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Word](/javascript/api/word)
- [Наборы обязательных элементов API JavaScript для Word](word-api-requirement-sets.md)
