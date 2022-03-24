---
title: Excel API JavaScript установлено 1.5
description: Сведения о наборе требований ExcelApi 1.5.
ms.date: 03/19/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 60da29607a8c8a22b38c9e19345a574e4f923922
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745210"
---
# <a name="whats-new-in-excel-javascript-api-15"></a>Новые возможности API JavaScript для Excel 1.5

ExcelApi 1.5 добавляет пользовательские XML-части. Они доступны через настраиваемую [коллекцию частей XML](/javascript/api/excel/excel.workbook#excel-excel-workbook-customxmlparts-member) в объекте книги.

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

В следующей таблице перечислены API в Excel API JavaScript, установленный 1.5. Чтобы просмотреть справочную документацию API для всех API, поддерживаемых Excel API JavaScript, за набором 1.5 или более ранних, см. в Excel API в наборе требований [1.5 или ранее](/javascript/api/excel?view=excel-js-1.5&preserve-view=true).

| Класс | Поля | Описание |
|:---|:---|:---|
|[CustomXmlPart](/javascript/api/excel/excel.customxmlpart)|[delete()](/javascript/api/excel/excel.customxmlpart#excel-excel-customxmlpart-delete-member(1))|Удаляет пользовательскую XML-часть.|
||[getXml()](/javascript/api/excel/excel.customxmlpart#excel-excel-customxmlpart-getxml-member(1))|Получает полное содержимое пользовательской XML-части.|
||[id](/javascript/api/excel/excel.customxmlpart#excel-excel-customxmlpart-id-member)|Пользовательский ID части XML.|
||[namespaceUri](/javascript/api/excel/excel.customxmlpart#excel-excel-customxmlpart-namespaceuri-member)|Пользовательское пространство имен XML-части URI.|
||[setXml (xml: string)](/javascript/api/excel/excel.customxmlpart#excel-excel-customxmlpart-setxml-member(1))|Задает полное содержимое пользовательской XML-части.|
|[CustomXmlPartCollection](/javascript/api/excel/excel.customxmlpartcollection)|[add(xml: string)](/javascript/api/excel/excel.customxmlpartcollection#excel-excel-customxmlpartcollection-add-member(1))|Добавляет новую пользовательскую XML-часть в книгу.|
||[getByNamespace (namespaceUri: string)](/javascript/api/excel/excel.customxmlpartcollection#excel-excel-customxmlpartcollection-getbynamespace-member(1))|Получает новую ограниченную коллекцию пользовательских XML-частей, пространства имен которых совпадают с указанным пространством имен.|
||[getCount()](/javascript/api/excel/excel.customxmlpartcollection#excel-excel-customxmlpartcollection-getcount-member(1))|Получает количество пользовательских частей XML в коллекции.|
||[getItem(id: string)](/javascript/api/excel/excel.customxmlpartcollection#excel-excel-customxmlpartcollection-getitem-member(1))|Получает пользовательскую XML-часть по идентификатору.|
||[getItemOrNullObject(id: строка)](/javascript/api/excel/excel.customxmlpartcollection#excel-excel-customxmlpartcollection-getitemornullobject-member(1))|Получает пользовательскую XML-часть по идентификатору.|
||[items](/javascript/api/excel/excel.customxmlpartcollection#excel-excel-customxmlpartcollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[CustomXmlPartScopedCollection](/javascript/api/excel/excel.customxmlpartscopedcollection)|[getCount()](/javascript/api/excel/excel.customxmlpartscopedcollection#excel-excel-customxmlpartscopedcollection-getcount-member(1))|Получает количество частей CustomXML в этой коллекции.|
||[getItem(id: string)](/javascript/api/excel/excel.customxmlpartscopedcollection#excel-excel-customxmlpartscopedcollection-getitem-member(1))|Получает пользовательскую XML-часть по идентификатору.|
||[getItemOrNullObject(id: строка)](/javascript/api/excel/excel.customxmlpartscopedcollection#excel-excel-customxmlpartscopedcollection-getitemornullobject-member(1))|Получает пользовательскую XML-часть по идентификатору.|
||[getOnlyItem()](/javascript/api/excel/excel.customxmlpartscopedcollection#excel-excel-customxmlpartscopedcollection-getonlyitem-member(1))|Если коллекция содержит ровно один элемент, этот метод возвращает его.|
||[getOnlyItemOrNullObject()](/javascript/api/excel/excel.customxmlpartscopedcollection#excel-excel-customxmlpartscopedcollection-getonlyitemornullobject-member(1))|Если коллекция содержит ровно один элемент, этот метод возвращает его.|
||[items](/javascript/api/excel/excel.customxmlpartscopedcollection#excel-excel-customxmlpartscopedcollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
|[PivotTable](/javascript/api/excel/excel.pivottable)|[id](/javascript/api/excel/excel.pivottable#excel-excel-pivottable-id-member)|ID of the PivotTable.|
|[RequestContext](/javascript/api/excel/excel.requestcontext)|[runtime](/javascript/api/excel/excel.requestcontext#excel-excel-requestcontext-runtime-member)||
|[Runtime](/javascript/api/excel/excel.runtime)|||
|[Workbook](/javascript/api/excel/excel.workbook)|[customXmlParts](/javascript/api/excel/excel.workbook#excel-excel-workbook-customxmlparts-member)|Представляет коллекцию пользовательских частей XML, содержащихся в этой книге.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[getNext (visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getnext-member(1))|Получает таблицу, которая следует за этим.|
||[getNextOrNullObject (visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getnextornullobject-member(1))|Получает таблицу, которая следует за этим.|
||[getPrevious (visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getprevious-member(1))|Получает таблицу, предшествующего этому.|
||[getPreviousOrNullObject (visibleOnly?: boolean)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getpreviousornullobject-member(1))|Получает таблицу, предшествующего этому.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[getFirst (visibleOnly?: boolean)](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-getfirst-member(1))|Получает первый лист в коллекции.|
||[getLast (visibleOnly?: boolean)](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-getlast-member(1))|Получает последний лист в коллекции.|

## <a name="see-also"></a>См. также

* [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.5&preserve-view=true)
* [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)
