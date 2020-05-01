---
title: Набор обязательных элементов API JavaScript для Excel 1,5
description: Сведения о наборе требований ExcelApi 1,5
ms.date: 07/26/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 4860ef7153abea059d04cda5e8709c7bdbd88855
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064762"
---
# <a name="whats-new-in-excel-javascript-api-15"></a>Новые возможности API JavaScript для Excel 1.5

ExcelApi 1,5 добавляет пользовательские XML-части. Они доступны через [настраиваемую коллекцию XML-частей](/javascript/api/excel/excel.workbook#customxmlparts) в объекте Workbook.

## <a name="custom-xml-part"></a>Пользовательская XML-часть

* Получение настраиваемых XML-частей с помощью идентификатора.
* Получение новой ограниченной коллекции пользовательских XML-частей, пространства имен которых совпадают с указанным пространством имен.
* Получение XML-строки, связанной с частью.
* Укажите идентификатор и пространство имен части.
* Добавьте в книгу новую пользовательскую XML-часть.
* Задавайте всю XML-часть.
* Удаление пользовательской XML-части.
* Удаление атрибута с указанным именем из элемента, указанного по XPath.
* Запрос содержимого XML по XPath.
* Атрибуты вставки, обновления и удаления.

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в наборе обязательных элементов API JavaScript для Excel 1,5. Чтобы просмотреть справочную документацию по API для всех API, поддерживаемых набором обязательных элементов API JavaScript для Excel 1,5 или более ранней версии, обратитесь к разделам [API Excel в наборе требований 1,5](/javascript/api/excel?view=excel-js-1.5)

| Класс | Поля | Описание |
|:---|:---|:---|
|[CustomXmlPart](/javascript/api/excel/excel.customxmlpart)|[delete()](/javascript/api/excel/excel.customxmlpart#delete--)|Удаляет пользовательскую XML-часть.|
||[Жетксмл ()](/javascript/api/excel/excel.customxmlpart#getxml--)|Получает полное содержимое пользовательской XML-части.|
||[id](/javascript/api/excel/excel.customxmlpart#id)|ИДЕНТИФИКАТОР пользовательской XML-части. Только для чтения.|
||[Пространства](/javascript/api/excel/excel.customxmlpart#namespaceuri)|URI пространства имен настраиваемой части XML. Только для чтения.|
||[setXml (XML: строка)](/javascript/api/excel/excel.customxmlpart#setxml-xml-)|Задает полное содержимое пользовательской XML-части.|
|[CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|[Add (XML: String)](/javascript/api/excel/excel.customxmlpartcollection#add-xml-)|Добавляет новую пользовательскую XML-часть в книгу.|
||[getByNamespace (namespaceUri: строка)](/javascript/api/excel/excel.customxmlpartcollection#getbynamespace-namespaceuri-)|Получает новую ограниченную коллекцию пользовательских XML-частей, пространства имен которых совпадают с указанным пространством имен.|
||[getCount()](/javascript/api/excel/excel.customxmlpartcollection#getcount--)|Получает количество частей CustomXml в коллекции.|
||[getItem(id: string)](/javascript/api/excel/excel.customxmlpartcollection#getitem-id-)|Получает пользовательскую XML-часть по идентификатору.|
||[getItemOrNullObject(id: строка)](/javascript/api/excel/excel.customxmlpartcollection#getitemornullobject-id-)|Получает пользовательскую XML-часть по идентификатору.|
||[items](/javascript/api/excel/excel.customxmlpartcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[кустомксмлпартскопедколлектион](/javascript/api/excel/excel.customxmlpartscopedcollection)|[getCount()](/javascript/api/excel/excel.customxmlpartscopedcollection#getcount--)|Получает количество частей CustomXML в этой коллекции.|
||[getItem(id: string)](/javascript/api/excel/excel.customxmlpartscopedcollection#getitem-id-)|Получает пользовательскую XML-часть по идентификатору.|
||[getItemOrNullObject(id: строка)](/javascript/api/excel/excel.customxmlpartscopedcollection#getitemornullobject-id-)|Получает пользовательскую XML-часть по идентификатору.|
||[Жетонлитем ()](/javascript/api/excel/excel.customxmlpartscopedcollection#getonlyitem--)|Если коллекция содержит ровно один элемент, этот метод возвращает его.|
||[Жетонлитеморнуллобжект ()](/javascript/api/excel/excel.customxmlpartscopedcollection#getonlyitemornullobject--)|Если коллекция содержит ровно один элемент, этот метод возвращает его.|
||[items](/javascript/api/excel/excel.customxmlpartscopedcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[id](/javascript/api/excel/excel.pivottable#id)|Идентификатор сводной таблицы. Только для чтения.|
|[Среда выполнения](/javascript/api/excel/excel.runtime)||[Workbook](/javascript/api/excel/excel.workbook)|[customXmlParts](/javascript/api/excel/excel.workbook#customxmlparts)|Представляет коллекцию настраиваемых XML-частей, которые содержит эта книга. Только для чтения.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[GetNext (visibleOnly?: Boolean)](/javascript/api/excel/excel.worksheet#getnext-visibleonly-)|Получает лист, следующий по отношению к элементу. При отсутствии листов, указанных ниже, этот метод вызовет ошибку.|
||[getNextOrNullObject (visibleOnly?: Boolean)](/javascript/api/excel/excel.worksheet#getnextornullobject-visibleonly-)|Получает лист, следующий по отношению к элементу. Если после этого листа нет ни одного листа, этот метод возвратит нулевой объект.|
||[Previous (visibleOnly?: Boolean)](/javascript/api/excel/excel.worksheet#getprevious-visibleonly-)|Получает лист, который предшествует этому. Если нет предыдущих листов, этот метод выдаст ошибку.|
||[getPreviousOrNullObject (visibleOnly?: Boolean)](/javascript/api/excel/excel.worksheet#getpreviousornullobject-visibleonly-)|Получает лист, который предшествует этому. Если нет предыдущих листов, этот метод возвратит значение NULL обжет.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[-First (visibleOnly?: Boolean)](/javascript/api/excel/excel.worksheetcollection#getfirst-visibleonly-)|Получает первый лист в коллекции.|
||[-Last (visibleOnly?: Boolean)](/javascript/api/excel/excel.worksheetcollection#getlast-visibleonly-)|Получает последний лист в коллекции.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.5)
- [Наборы обязательных элементов API JavaScript для Excel](./excel-api-requirement-sets.md)
