---
title: Набор требований к API JavaScript Excel 1.5
description: Сведения о наборе требований ExcelApi 1.5.
ms.date: 03/19/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 9c2ebf7230bec32f5036f2fc530bb82f492f2246
ms.sourcegitcommit: 7482ab6bc258d98acb9ba9b35c7dd3b5cc5bed21
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/24/2021
ms.locfileid: "51178099"
---
# <a name="whats-new-in-excel-javascript-api-15"></a>Новые возможности API JavaScript для Excel 1.5

ExcelApi 1.5 добавляет пользовательские XML-части. Они доступны через настраиваемую [коллекцию частей XML](/javascript/api/excel/excel.workbook#customxmlparts) в объекте книги.

## <a name="custom-xml-part"></a>Пользовательская XML-часть

* Получите настраиваемые XML-части с помощью их ID.
* Получение новой ограниченной коллекции пользовательских XML-частей, пространства имен которых совпадают с указанным пространством имен.
* Получите строку XML, связанную с частью.
* Предоставление ID и пространства имен части.
* Добавьте новую настраиваемую часть XML в книгу.
* Установите всю часть XML.
* Удаление пользовательской XML-части.
* Удаление атрибута с указанным именем из элемента, указанного по XPath.
* Запрос содержимого XML по XPath.
* Вставка, обновление и удаление атрибутов.

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в API Excel JavaScript, за набором 1.5. Чтобы просмотреть справочную документацию API для всех API, поддерживаемых требованием API Excel JavaScript, установленным 1.5 или ранее, см. в справке Об API Excel в наборе требований [1.5](/javascript/api/excel?view=excel-js-1.5&preserve-view=true)или более ранних .

| Класс | Поля | Описание |
|:---|:---|:---|
|[CustomXmlPart](/javascript/api/excel/excel.customxmlpart)|[delete()](/javascript/api/excel/excel.customxmlpart#delete--)|Удаляет пользовательскую XML-часть.|
||[getXml()](/javascript/api/excel/excel.customxmlpart#getxml--)|Получает полное содержимое пользовательской XML-части.|
||[id](/javascript/api/excel/excel.customxmlpart#id)|Пользовательский ID части XML.|
||[namespaceUri](/javascript/api/excel/excel.customxmlpart#namespaceuri)|Пользовательское пространство имен XML-части URI.|
||[setXml (xml: string)](/javascript/api/excel/excel.customxmlpart#setxml-xml-)|Задает полное содержимое пользовательской XML-части.|
|[CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|[add(xml: string)](/javascript/api/excel/excel.customxmlpartcollection#add-xml-)|Добавляет новую пользовательскую XML-часть в книгу.|
||[getByNamespace (namespaceUri: string)](/javascript/api/excel/excel.customxmlpartcollection#getbynamespace-namespaceuri-)|Получает новую ограниченную коллекцию пользовательских XML-частей, пространства имен которых совпадают с указанным пространством имен.|
||[getCount()](/javascript/api/excel/excel.customxmlpartcollection#getcount--)|Получает количество частей CustomXml в коллекции.|
||[getItem(id: string)](/javascript/api/excel/excel.customxmlpartcollection#getitem-id-)|Получает пользовательскую XML-часть по идентификатору.|
||[getItemOrNullObject(id: строка)](/javascript/api/excel/excel.customxmlpartcollection#getitemornullobject-id-)|Получает пользовательскую XML-часть по идентификатору.|
||[items](/javascript/api/excel/excel.customxmlpartcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[CustomXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|[getCount()](/javascript/api/excel/excel.customxmlpartscopedcollection#getcount--)|Получает количество частей CustomXML в этой коллекции.|
||[getItem(id: string)](/javascript/api/excel/excel.customxmlpartscopedcollection#getitem-id-)|Получает пользовательскую XML-часть по идентификатору.|
||[getItemOrNullObject(id: строка)](/javascript/api/excel/excel.customxmlpartscopedcollection#getitemornullobject-id-)|Получает пользовательскую XML-часть по идентификатору.|
||[getOnlyItem()](/javascript/api/excel/excel.customxmlpartscopedcollection#getonlyitem--)|Если коллекция содержит ровно один элемент, этот метод возвращает его.|
||[getOnlyItemOrNullObject()](/javascript/api/excel/excel.customxmlpartscopedcollection#getonlyitemornullobject--)|Если коллекция содержит ровно один элемент, этот метод возвращает его.|
||[items](/javascript/api/excel/excel.customxmlpartscopedcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[id](/javascript/api/excel/excel.pivottable#id)|Идентификатор сводной таблицы.|
|[RequestContext](/javascript/api/excel/excel.requestcontext)|[runtime](/javascript/api/excel/excel.requestcontext#runtime)||
|[Время выполнения](/javascript/api/excel/excel.runtime)|||
|[Workbook](/javascript/api/excel/excel.workbook)|[customXmlParts](/javascript/api/excel/excel.workbook#customxmlparts)|Представляет коллекцию пользовательских частей XML, содержащихся в этой книге.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[getNext (visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#getnext-visibleonly-)|Получает таблицу, которая следует за этим.|
||[getNextOrNullObject (visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#getnextornullobject-visibleonly-)|Получает таблицу, которая следует за этим.|
||[getPrevious (visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#getprevious-visibleonly-)|Получает таблицу, предшествующего этому.|
||[getPreviousOrNullObject (visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#getpreviousornullobject-visibleonly-)|Получает таблицу, предшествующего этому.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[getFirst (visibleOnly?: boolean)](/javascript/api/excel/excel.worksheetcollection#getfirst-visibleonly-)|Получает первый лист в коллекции.|
||[getLast (visibleOnly?: boolean)](/javascript/api/excel/excel.worksheetcollection#getlast-visibleonly-)|Получает последний лист в коллекции.|

## <a name="see-also"></a>См. также

* [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.5&preserve-view=true)
* [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)
