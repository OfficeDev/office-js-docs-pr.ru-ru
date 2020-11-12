---
title: API предварительного просмотра для Word JavaScript
description: Сведения о предстоящих API JavaScript для Word
ms.date: 11/09/2020
ms.prod: word
localization_priority: Normal
ms.openlocfilehash: 6a3b67e65c4ced3f1b89d98afe45d5d6c33f63b6
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996405"
---
# <a name="word-javascript-preview-apis"></a>API предварительного просмотра для Word JavaScript

Новые API JavaScript для Word впервые представлены в слове Preview и далее становятся частью определенного набора обязательных требований после выполнения тестирования и получения отзывов пользователей.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

## <a name="api-list"></a>Список API

В следующей таблице перечислены API JavaScript для Word, находящиеся в предварительной версии. Чтобы просмотреть полный список всех API JavaScript для Word (включая предварительные API и ранее выпущенные API), ознакомьтесь со статьями [все API JavaScript для Word](/javascript/api/word?view=word-js-preview&preserve-view=true).

| Класс | Поля | Описание |
|:---|:---|:---|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[onDataChanged](/javascript/api/word/word.contentcontrol#ondatachanged)|Возникает при изменении данных в элементе управления содержимым.|
||[onDeleted](/javascript/api/word/word.contentcontrol#ondeleted)|Возникает при удалении элемента управления содержимым.|
||[onSelectionChanged](/javascript/api/word/word.contentcontrol#onselectionchanged)|Возникает при изменении выделенного фрагмента в элементе управления содержимым.|
|[ContentControlEventArgs](/javascript/api/word/word.contentcontroleventargs)|[contentControl](/javascript/api/word/word.contentcontroleventargs#contentcontrol)|Объект, который вызвал событие.|
||[eventType](/javascript/api/word/word.contentcontroleventargs#eventtype)|Тип события.|
|[CustomXmlPart](/javascript/api/word/word.customxmlpart)|[delete()](/javascript/api/word/word.customxmlpart#delete--)|Удаляет пользовательскую XML-часть.|
||[Делетеаттрибуте (XPath: строка, Намеспацемаппингс: Any, Name: строка)](/javascript/api/word/word.customxmlpart#deleteattribute-xpath--namespacemappings--name-)|Удаляет атрибут с указанным именем из элемента, указанного с помощью XPath.|
||[Делетилемент (XPath: String, Намеспацемаппингс: Any)](/javascript/api/word/word.customxmlpart#deleteelement-xpath--namespacemappings-)|Удаляет элемент, указанный с помощью XPath.|
||[Жетксмл ()](/javascript/api/word/word.customxmlpart#getxml--)|Получает полное XML-содержимое пользовательской XML-части.|
||[Инсертаттрибуте (XPath: String, Намеспацемаппингс: Any, Name: String, Value: String)](/javascript/api/word/word.customxmlpart#insertattribute-xpath--namespacemappings--name--value-)|Вставляет атрибут с заданным именем и значением в элемент, указанный с помощью XPath.|
||[Инсертелемент (XPath: строка, XML: строка, Намеспацемаппингс: Any, index?: число)](/javascript/api/word/word.customxmlpart#insertelement-xpath--xml--namespacemappings--index-)|Вставляет заданный XML-код в родительский элемент, определенный с помощью XPath в индексе позиции дочернего элемента.|
||[запрос (XPath: String, Намеспацемаппингс: Any)](/javascript/api/word/word.customxmlpart#query-xpath--namespacemappings-)|Запрашивает XML-содержимое пользовательской XML-части.|
||[id](/javascript/api/word/word.customxmlpart#id)|Получает идентификатор пользовательской XML-части.|
||[Пространства](/javascript/api/word/word.customxmlpart#namespaceuri)|Получает URI пространства имен настраиваемой XML-части.|
||[setXml (XML: строка)](/javascript/api/word/word.customxmlpart#setxml-xml-)|Задает полное XML-содержимое пользовательской XML-части.|
||[Упдатеаттрибуте (XPath: String, Намеспацемаппингс: Any, Name: String, Value: String)](/javascript/api/word/word.customxmlpart#updateattribute-xpath--namespacemappings--name--value-)|Обновляет значение атрибута, используя заданное имя элемента, указанного с помощью XPath.|
||[Упдатилемент (XPath: строка, XML: строка, Намеспацемаппингс: Any)](/javascript/api/word/word.customxmlpart#updateelement-xpath--xml--namespacemappings-)|Обновляет XML элемента, указанного с помощью XPath.|
|[CustomXmlPartCollection](/javascript/api/word/word.customxmlpartcollection)|[Add (XML: String)](/javascript/api/word/word.customxmlpartcollection#add-xml-)|Добавляет новую пользовательскую XML-часть в документ.|
||[getByNamespace (namespaceUri: строка)](/javascript/api/word/word.customxmlpartcollection#getbynamespace-namespaceuri-)|Получает новую ограниченную коллекцию пользовательских XML-частей, пространства имен которых совпадают с указанным пространством имен.|
||[getCount()](/javascript/api/word/word.customxmlpartcollection#getcount--)|Возвращает число элементов в коллекции.|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartcollection#getitem-id-)|Получает пользовательскую XML-часть по идентификатору.|
||[getItemOrNullObject(id: строка)](/javascript/api/word/word.customxmlpartcollection#getitemornullobject-id-)|Получает пользовательскую XML-часть по идентификатору.|
||[items](/javascript/api/word/word.customxmlpartcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[CustomXmlPartScopedCollection](/javascript/api/word/word.customxmlpartscopedcollection)|[getCount()](/javascript/api/word/word.customxmlpartscopedcollection#getcount--)|Возвращает число элементов в коллекции.|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#getitem-id-)|Получает пользовательскую XML-часть по идентификатору.|
||[getItemOrNullObject(id: строка)](/javascript/api/word/word.customxmlpartscopedcollection#getitemornullobject-id-)|Получает пользовательскую XML-часть по идентификатору.|
||[Жетонлитем ()](/javascript/api/word/word.customxmlpartscopedcollection#getonlyitem--)|Если коллекция содержит ровно один элемент, этот метод возвращает его.|
||[Жетонлитеморнуллобжект ()](/javascript/api/word/word.customxmlpartscopedcollection#getonlyitemornullobject--)|Если коллекция содержит ровно один элемент, этот метод возвращает его.|
||[items](/javascript/api/word/word.customxmlpartscopedcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Document](/javascript/api/word/word.document)|[Делетебукмарк (имя: строка)](/javascript/api/word/word.document#deletebookmark-name-)|Удаляет закладку (если она существует) из документа.|
||[Жетбукмаркранже (имя: строка)](/javascript/api/word/word.document#getbookmarkrange-name-)|Возвращает диапазон закладок.|
||[Жетбукмаркранжеорнуллобжект (имя: строка)](/javascript/api/word/word.document#getbookmarkrangeornullobject-name-)|Возвращает диапазон закладок.|
||[customXmlParts](/javascript/api/word/word.document#customxmlparts)|Возвращает пользовательские XML-части в документе.|
||[онконтентконтроладдед](/javascript/api/word/word.document#oncontentcontroladded)|Возникает при добавлении элемента управления содержимым.|
||[settings](/javascript/api/word/word.document#settings)|Получает параметры надстройки в документе.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[Делетебукмарк (имя: строка)](/javascript/api/word/word.documentcreated#deletebookmark-name-)|Удаляет закладку (если она существует) из документа.|
||[Жетбукмаркранже (имя: строка)](/javascript/api/word/word.documentcreated#getbookmarkrange-name-)|Возвращает диапазон закладок.|
||[Жетбукмаркранжеорнуллобжект (имя: строка)](/javascript/api/word/word.documentcreated#getbookmarkrangeornullobject-name-)|Возвращает диапазон закладок.|
||[customXmlParts](/javascript/api/word/word.documentcreated#customxmlparts)|Возвращает пользовательские XML-части в документе.|
||[settings](/javascript/api/word/word.documentcreated#settings)|Получает параметры надстройки в документе.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[имажеформат](/javascript/api/word/word.inlinepicture#imageformat)|Получает формат встроенного изображения.|
|[List](/javascript/api/word/word.list)|[Жетлевелфонт (Level: число)](/javascript/api/word/word.list#getlevelfont-level-)|Получает или задает значение, указывающее, указаны ли в списке.|
||[Жетлевелпиктуре (Level: число)](/javascript/api/word/word.list#getlevelpicture-level-)|Получает строковое представление изображения в кодировке Base64 на указанном уровне в списке.|
||[Ресетлевелфонт (Level: число, Ресетфонтнаме?: Boolean)](/javascript/api/word/word.list#resetlevelfont-level--resetfontname-)|Сбрасывает шрифт маркера, числа или изображения на указанном уровне списка.|
||[Сетлевелпиктуре (Level: число, base64EncodedImage?: строка)](/javascript/api/word/word.list#setlevelpicture-level--base64encodedimage-)|Задает рисунок на указанном уровне в списке.|
|[Range](/javascript/api/word/word.range)|[Закладки (Инклудехидден?: Boolean, Инклудеаджацент?: Boolean)](/javascript/api/word/word.range#getbookmarks-includehidden--includeadjacent-)|Получает имена всех закладок в диапазоне или перекрывают их.|
||[Инсертбукмарк (имя: строка)](/javascript/api/word/word.range#insertbookmark-name-)|Вставляет закладку в диапазон.|
|[Параметр](/javascript/api/word/word.setting)|[delete()](/javascript/api/word/word.setting#delete--)|Удаляет параметр.|
||[key](/javascript/api/word/word.setting#key)|Получает ключ параметра.|
||[value](/javascript/api/word/word.setting#value)|Получает или задает значение параметра.|
|[SettingCollection](/javascript/api/word/word.settingcollection)|[Add (Key: строка, Value: Any)](/javascript/api/word/word.settingcollection#add-key--value-)|Создает новый параметр или устанавливает существующий параметр.|
||[deleteAll ()](/javascript/api/word/word.settingcollection#deleteall--)|Удаляет все параметры в этой надстройке.|
||[getCount()](/javascript/api/word/word.settingcollection#getcount--)|Получает количество параметров.|
||[getItem(key: string)](/javascript/api/word/word.settingcollection#getitem-key-)|Получает объект Setting по ключу, для которого учитывается регистр.|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.settingcollection#getitemornullobject-key-)|Получает объект Setting по ключу, для которого учитывается регистр.|
||[items](/javascript/api/word/word.settingcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Table](/javascript/api/word/word.table)|[Мержецеллс (Топров: число, Фирстцелл: число, Боттомров: число, Ластцелл: число)](/javascript/api/word/word.table#mergecells-toprow--firstcell--bottomrow--lastcell-)|Объединяет ячейки, ограниченные в первой и последней ячейках.|
|[TableCell](/javascript/api/word/word.tablecell)|[Split (rowCount: число, columnCount: число)](/javascript/api/word/word.tablecell#split-rowcount--columncount-)|Разделяет ячейку на указанное количество строк и столбцов.|
|[TableRow](/javascript/api/word/word.tablerow)|[insertContentControl()](/javascript/api/word/word.tablerow#insertcontentcontrol--)|Вставляет в строку элемент управления содержимым.|
||[Merge ()](/javascript/api/word/word.tablerow#merge--)|Объединяет строку в одну ячейку.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Word](/javascript/api/word)
- [Наборы обязательных элементов API JavaScript для Word](word-api-requirement-sets.md)
