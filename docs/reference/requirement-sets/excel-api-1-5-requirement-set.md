---
title: Excel Набор API JavaScript 1.5
description: Сведения о наборе требований ExcelApi 1.5.
ms.date: 03/19/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 01a13a0f531eae9eea2c213ba0da764fbe51ee15
ms.sourcegitcommit: 3fa8c754a47bab909e559ae3e5d4237ba27fdbe4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/30/2021
ms.locfileid: "53671809"
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

В следующей таблице перечислены API в Excel API JavaScript, за набором 1.5. Чтобы просмотреть справочную документацию API для всех API, поддерживаемых Excel API JavaScript, за набором 1.5 или более ранних, см. в Excel API в наборе требований [1.5](/javascript/api/excel?view=excel-js-1.5&preserve-view=true)или более ранних .

| Класс | Поля | Описание |
|:---|:---|:---|
|[CustomXmlPart](/javascript/api/excel/excel.customxmlpart)|[delete()](/javascript/api/excel/excel.customxmlpart#delete__)|Удаляет пользовательскую XML-часть.|
||[getXml()](/javascript/api/excel/excel.customxmlpart#getXml__)|Получает полное содержимое пользовательской XML-части.|
||[id](/javascript/api/excel/excel.customxmlpart#id)|Пользовательский ID части XML.|
||[namespaceUri](/javascript/api/excel/excel.customxmlpart#namespaceUri)|Пользовательское пространство имен XML-части URI.|
||[setXml (xml: string)](/javascript/api/excel/excel.customxmlpart#setXml_xml_)|Задает полное содержимое пользовательской XML-части.|
|[CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|[add(xml: string)](/javascript/api/excel/excel.customxmlpartcollection#add_xml_)|Добавляет новую пользовательскую XML-часть в книгу.|
||[getByNamespace (namespaceUri: string)](/javascript/api/excel/excel.customxmlpartcollection#getByNamespace_namespaceUri_)|Получает новую ограниченную коллекцию пользовательских XML-частей, пространства имен которых совпадают с указанным пространством имен.|
||[getCount()](/javascript/api/excel/excel.customxmlpartcollection#getCount__)|Получает количество пользовательских частей XML в коллекции.|
||[getItem(id: string)](/javascript/api/excel/excel.customxmlpartcollection#getItem_id_)|Получает пользовательскую XML-часть по идентификатору.|
||[getItemOrNullObject(id: строка)](/javascript/api/excel/excel.customxmlpartcollection#getItemOrNullObject_id_)|Получает пользовательскую XML-часть по идентификатору.|
||[items](/javascript/api/excel/excel.customxmlpartcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[CustomXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|[getCount()](/javascript/api/excel/excel.customxmlpartscopedcollection#getCount__)|Получает количество частей CustomXML в этой коллекции.|
||[getItem(id: string)](/javascript/api/excel/excel.customxmlpartscopedcollection#getItem_id_)|Получает пользовательскую XML-часть по идентификатору.|
||[getItemOrNullObject(id: строка)](/javascript/api/excel/excel.customxmlpartscopedcollection#getItemOrNullObject_id_)|Получает пользовательскую XML-часть по идентификатору.|
||[getOnlyItem()](/javascript/api/excel/excel.customxmlpartscopedcollection#getOnlyItem__)|Если коллекция содержит ровно один элемент, этот метод возвращает его.|
||[getOnlyItemOrNullObject()](/javascript/api/excel/excel.customxmlpartscopedcollection#getOnlyItemOrNullObject__)|Если коллекция содержит ровно один элемент, этот метод возвращает его.|
||[items](/javascript/api/excel/excel.customxmlpartscopedcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[id](/javascript/api/excel/excel.pivottable#id)|ID of the PivotTable.|
|[RequestContext](/javascript/api/excel/excel.requestcontext)|[runtime](/javascript/api/excel/excel.requestcontext#runtime)||
|[Runtime](/javascript/api/excel/excel.runtime)|||
|[Workbook](/javascript/api/excel/excel.workbook)|[customXmlParts](/javascript/api/excel/excel.workbook#customXmlParts)|Представляет коллекцию пользовательских частей XML, содержащихся в этой книге.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[getNext (visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#getNext_visibleOnly_)|Получает таблицу, которая следует за этим.|
||[getNextOrNullObject (visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#getNextOrNullObject_visibleOnly_)|Получает таблицу, которая следует за этим.|
||[getPrevious (visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#getPrevious_visibleOnly_)|Получает таблицу, предшествующего этому.|
||[getPreviousOrNullObject (visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#getPreviousOrNullObject_visibleOnly_)|Получает таблицу, предшествующего этому.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[getFirst (visibleOnly?: boolean)](/javascript/api/excel/excel.worksheetcollection#getFirst_visibleOnly_)|Получает первый лист в коллекции.|
||[getLast (visibleOnly?: boolean)](/javascript/api/excel/excel.worksheetcollection#getLast_visibleOnly_)|Получает последний лист в коллекции.|

## <a name="see-also"></a>См. также

* [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.5&preserve-view=true)
* [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)
