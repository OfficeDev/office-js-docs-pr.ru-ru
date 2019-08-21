---
title: API предварительного просмотра для Word JavaScript
description: Сведения о предстоящих API JavaScript для Word
ms.date: 08/15/2019
ms.prod: word
localization_priority: Normal
ms.openlocfilehash: 1bc6cf2f4b8d8bf876d0b28ead9643f14c81fde1
ms.sourcegitcommit: da8e6148f4bd9884ab9702db3033273a383d15f0
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/20/2019
ms.locfileid: "36477906"
---
# <a name="word-javascript-preview-apis"></a>API предварительного просмотра для Word JavaScript

Новые API JavaScript для Word впервые представлены в слове Preview и далее становятся частью определенного набора обязательных требований после выполнения тестирования и получения отзывов пользователей.

[!INCLUDE [Information about using preview APIs](../../includes/using-preview-apis-host.md)]

## <a name="api-list"></a>Список API

В следующей таблице перечислены API JavaScript для Word, находящиеся в предварительной версии. Чтобы просмотреть полный список всех API JavaScript для Word (включая предварительные API и ранее выпущенные API), ознакомьтесь со статьями [все API JavaScript для Word](/javascript/api/word?view=word-js-preview).

| Класс | Поля | Описание |
|:---|:---|:---|
|[ContentControl](/javascript/api/word/word.contentcontrol)|[onDataChanged](/javascript/api/word/word.contentcontrol#ondatachanged)|Возникает при изменении данных в элементе управления содержимым. Чтобы получить новый текст, загрузите этот элемент управления содержимым в обработчике. Чтобы получить старый текст, не загружайте его.|
||[onDeleted](/javascript/api/word/word.contentcontrol#ondeleted)|Возникает при удалении элемента управления содержимым. Не загружайте этот элемент управления содержимым в обработчике, иначе вы не сможете получить исходные свойства.|
||[onSelectionChanged](/javascript/api/word/word.contentcontrol#onselectionchanged)|Возникает при изменении выделенного фрагмента в элементе управления содержимым.|
|[контентконтролевентаргс](/javascript/api/word/word.contentcontroleventargs)|[contentControl](/javascript/api/word/word.contentcontroleventargs#contentcontrol)|Объект, который вызвал событие. Загрузите этот объект, чтобы получить его свойства.|
||[eventType](/javascript/api/word/word.contentcontroleventargs#eventtype)|Тип события. Дополнительные сведения см. в Word. EventType.|
|[CustomXmlPart](/javascript/api/word/word.customxmlpart)|[delete()](/javascript/api/word/word.customxmlpart#delete--)|Удаляет пользовательскую XML-часть.|
||[Делетеаттрибуте (XPath: строка, Намеспацемаппингс: Any, Name: строка)](/javascript/api/word/word.customxmlpart#deleteattribute-xpath--namespacemappings--name-)|Удаляет атрибут с указанным именем из элемента, указанного с помощью XPath.|
||[Делетилемент (XPath: String, Намеспацемаппингс: Any)](/javascript/api/word/word.customxmlpart#deleteelement-xpath--namespacemappings-)|Удаляет элемент, указанный с помощью XPath.|
||[Жетксмл ()](/javascript/api/word/word.customxmlpart#getxml--)|Получает полное XML-содержимое пользовательской XML-части.|
||[Инсертаттрибуте (XPath: String, Намеспацемаппингс: Any, Name: String, Value: String)](/javascript/api/word/word.customxmlpart#insertattribute-xpath--namespacemappings--name--value-)|Вставляет атрибут с заданным именем и значением в элемент, указанный с помощью XPath.|
||[Инсертелемент (XPath: строка, XML: строка, Намеспацемаппингс: Any, index?: число)](/javascript/api/word/word.customxmlpart#insertelement-xpath--xml--namespacemappings--index-)|Вставляет заданный XML-код в родительский элемент, определенный с помощью XPath в индексе позиции дочернего элемента.|
||[запрос (XPath: String, Намеспацемаппингс: Any)](/javascript/api/word/word.customxmlpart#query-xpath--namespacemappings-)|Запрашивает XML-содержимое пользовательской XML-части.|
||[id](/javascript/api/word/word.customxmlpart#id)|Получает идентификатор пользовательской XML-части. Только для чтения.|
||[Пространства](/javascript/api/word/word.customxmlpart#namespaceuri)|Получает URI пространства имен настраиваемой XML-части. Только для чтения.|
||[setXml (XML: строка)](/javascript/api/word/word.customxmlpart#setxml-xml-)|Задает полное XML-содержимое пользовательской XML-части.|
||[Упдатеаттрибуте (XPath: String, Намеспацемаппингс: Any, Name: String, Value: String)](/javascript/api/word/word.customxmlpart#updateattribute-xpath--namespacemappings--name--value-)|Обновляет значение атрибута, используя заданное имя элемента, указанного с помощью XPath.|
||[Упдатилемент (XPath: строка, XML: строка, Намеспацемаппингс: Any)](/javascript/api/word/word.customxmlpart#updateelement-xpath--xml--namespacemappings-)|Обновляет XML элемента, указанного с помощью XPath.|
|[CustomXmlPartCollection](/javascript/api/word/word.customxmlpartcollection)|[Add (XML: String)](/javascript/api/word/word.customxmlpartcollection#add-xml-)|Добавляет новую пользовательскую XML-часть в документ.|
||[getByNamespace (namespaceUri: строка)](/javascript/api/word/word.customxmlpartcollection#getbynamespace-namespaceuri-)|Получает новую ограниченную коллекцию пользовательских XML-частей, пространства имен которых совпадают с указанным пространством имен.|
||[getCount()](/javascript/api/word/word.customxmlpartcollection#getcount--)|Возвращает число элементов в коллекции.|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartcollection#getitem-id-)|Получает пользовательскую XML-часть по идентификатору. Только для чтения.|
||[getItemOrNullObject(id: строка)](/javascript/api/word/word.customxmlpartcollection#getitemornullobject-id-)|Получает пользовательскую XML-часть по идентификатору. Возвращает нулевой объект, если CustomXmlPart не существует.|
||[items](/javascript/api/word/word.customxmlpartcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[кустомксмлпартскопедколлектион](/javascript/api/word/word.customxmlpartscopedcollection)|[getCount()](/javascript/api/word/word.customxmlpartscopedcollection#getcount--)|Возвращает число элементов в коллекции.|
||[getItem(id: string)](/javascript/api/word/word.customxmlpartscopedcollection#getitem-id-)|Получает пользовательскую XML-часть по идентификатору. Только для чтения.|
||[getItemOrNullObject(id: строка)](/javascript/api/word/word.customxmlpartscopedcollection#getitemornullobject-id-)|Получает пользовательскую XML-часть по идентификатору. Возвращает нулевой объект, если CustomXmlPart не существует в коллекции.|
||[Жетонлитем ()](/javascript/api/word/word.customxmlpartscopedcollection#getonlyitem--)|Если коллекция содержит ровно один элемент, этот метод возвращает его. В противном случае этот метод выдает ошибку.|
||[Жетонлитеморнуллобжект ()](/javascript/api/word/word.customxmlpartscopedcollection#getonlyitemornullobject--)|Если коллекция содержит ровно один элемент, этот метод возвращает его. В противном случае этот метод возвращает пустой объект.|
||[items](/javascript/api/word/word.customxmlpartscopedcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Document](/javascript/api/word/word.document)|[Делетебукмарк (имя: строка)](/javascript/api/word/word.document#deletebookmark-name-)|Удаляет закладку, если она существует, из документа.|
||[Жетбукмаркранже (имя: строка)](/javascript/api/word/word.document#getbookmarkrange-name-)|Возвращает диапазон закладок. Вызывается, если закладка не существует.|
||[Жетбукмаркранжеорнуллобжект (имя: строка)](/javascript/api/word/word.document#getbookmarkrangeornullobject-name-)|Возвращает диапазон закладок. Возвращает нулевой объект, если закладка не существует.|
||[customXmlParts](/javascript/api/word/word.document#customxmlparts)|Возвращает пользовательские XML-части в документе. Только для чтения.|
||[онконтентконтроладдед](/javascript/api/word/word.document#oncontentcontroladded)|Возникает при добавлении элемента управления содержимым. Выполните context. Sync () в обработчике, чтобы получить свойства нового элемента управления содержимым.|
||[settings](/javascript/api/word/word.document#settings)|Получает параметры надстройки в документе. Только для чтения.|
|[DocumentCreated](/javascript/api/word/word.documentcreated)|[Делетебукмарк (имя: строка)](/javascript/api/word/word.documentcreated#deletebookmark-name-)|Удаляет закладку, если она существует, из документа.|
||[Жетбукмаркранже (имя: строка)](/javascript/api/word/word.documentcreated#getbookmarkrange-name-)|Возвращает диапазон закладок. Вызывается, если закладка не существует.|
||[Жетбукмаркранжеорнуллобжект (имя: строка)](/javascript/api/word/word.documentcreated#getbookmarkrangeornullobject-name-)|Возвращает диапазон закладок. Возвращает нулевой объект, если закладка не существует.|
||[customXmlParts](/javascript/api/word/word.documentcreated#customxmlparts)|Возвращает пользовательские XML-части в документе. Только для чтения.|
||[settings](/javascript/api/word/word.documentcreated#settings)|Получает параметры надстройки в документе. Только для чтения.|
|[InlinePicture](/javascript/api/word/word.inlinepicture)|[имажеформат](/javascript/api/word/word.inlinepicture#imageformat)|Получает формат встроенного изображения. Только для чтения.|
|[List](/javascript/api/word/word.list)|[Жетлевелфонт (Level: число)](/javascript/api/word/word.list#getlevelfont-level-)|Получает или задает значение, указывающее, указаны ли в списке.|
||[Жетлевелпиктуре (Level: число)](/javascript/api/word/word.list#getlevelpicture-level-)|Получает строковое представление изображения в кодировке Base64 на указанном уровне в списке.|
||[Ресетлевелфонт (Level: число, Ресетфонтнаме?: Boolean)](/javascript/api/word/word.list#resetlevelfont-level--resetfontname-)|Сбрасывает шрифт маркера, номера или изображения на указанном уровне списка.|
||[Сетлевелпиктуре (Level: число, base64EncodedImage?: строка)](/javascript/api/word/word.list#setlevelpicture-level--base64encodedimage-)|Задает рисунок на указанном уровне в списке.|
|[Range](/javascript/api/word/word.range)|[Закладки (Инклудехидден?: Boolean, Инклудеаджацент?: Boolean)](/javascript/api/word/word.range#getbookmarks-includehidden--includeadjacent-)|Получает имена всех закладок в диапазоне или перекрывают их. Закладка скрывается, если ее имя начинается с символа подчеркивания.|
||[Инсертбукмарк (имя: строка)](/javascript/api/word/word.range#insertbookmark-name-)|Вставляет закладку в диапазон. Если закладка с таким же именем существует в другом месте, она будет удалена первыми.|
|[Параметр](/javascript/api/word/word.setting)|[delete()](/javascript/api/word/word.setting#delete--)|Удаляет параметр.|
||[key](/javascript/api/word/word.setting#key)|Получает ключ параметра. Только для чтения.|
||[value](/javascript/api/word/word.setting#value)|Получает или задает значение параметра.|
|[SettingCollection](/javascript/api/word/word.settingcollection)|[Add (Key: строка, Value: Any)](/javascript/api/word/word.settingcollection#add-key--value-)|Создает новый параметр или устанавливает существующий параметр.|
||[deleteAll ()](/javascript/api/word/word.settingcollection#deleteall--)|Удаляет все параметры в этой надстройке.|
||[getCount()](/javascript/api/word/word.settingcollection#getcount--)|Получает количество параметров.|
||[getItem(key: string)](/javascript/api/word/word.settingcollection#getitem-key-)|Получает объект Setting по ключу, для которого учитывается регистр. Вызывается, если параметр не существует.|
||[getItemOrNullObject(key: string)](/javascript/api/word/word.settingcollection#getitemornullobject-key-)|Получает объект Setting по ключу, для которого учитывается регистр. Возвращает нулевой объект, если параметр не существует.|
||[items](/javascript/api/word/word.settingcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[Table](/javascript/api/word/word.table)|[Мержецеллс (Топров: число, Фирстцелл: число, Боттомров: число, Ластцелл: число)](/javascript/api/word/word.table#mergecells-toprow--firstcell--bottomrow--lastcell-)|Объединяет ячейки, ограниченные в первой и последней ячейках.|
|[TableCell](/javascript/api/word/word.tablecell)|[Split (rowCount: число, columnCount: число)](/javascript/api/word/word.tablecell#split-rowcount--columncount-)|Разделяет ячейку на указанное количество строк и столбцов.|
|[TableRow](/javascript/api/word/word.tablerow)|[insertContentControl()](/javascript/api/word/word.tablerow#insertcontentcontrol--)|Вставляет в строку элемент управления содержимым.|
||[Merge ()](/javascript/api/word/word.tablerow#merge--)|Объединяет строку в одну ячейку.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Word](/javascript/api/word)
- [Наборы обязательных элементов API JavaScript для Word](word-api-requirement-sets.md)
