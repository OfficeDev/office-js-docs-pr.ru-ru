---
title: Excel Набор API JavaScript 1.4
description: Сведения о наборе требований ExcelApi 1.4.
ms.date: 11/09/2020
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: e8a38d45a73560cade580cf5f2f9a0892790b38d
ms.sourcegitcommit: e4d98eb90e516b9c90e3832f3212caf48691acf6
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/22/2021
ms.locfileid: "60537480"
---
# <a name="whats-new-in-excel-javascript-api-14"></a>Новые возможности API JavaScript для Excel 1.4

Ниже перечислено то, что было недавно добавлено в набор обязательных элементов 1.4, относящийся к API JavaScript для Excel.

## <a name="named-item-add-and-new-properties"></a>Именованный элемент add и новые свойства

Новые свойства:

* `comment`
* `scope` — таблицы или элементы с объемной областью книги.
* `worksheet` - Возвращает таблицу, в которой область действия именуемого элемента.

Новые методы:

* `add(name: string, reference: Range or string, comment: string)` - Добавляет новое имя в коллекцию данной области.
* `addFormulaLocal(name: string, formula: string, comment: string)` - Добавляет новое имя в коллекцию данной области с помощью пользовательского локаула для формулы.

## <a name="settings-api-in-the-excel-namespace"></a>Параметры API в пространстве имен Excel

Объект [Setting](/javascript/api/excel/excel.setting) представляет пару "ключ-значение" для параметра, хранящегося в документе. Функциональные возможности объекта `Excel.Setting` аналогичны `Office.Settings`, но он использует пакетный синтаксис API, а не модель обратного вызова общего API.

API включают в себя создание записи с помощью ключа и добавление указанной пары параметров `getItem()` `add()` ключа:значения в книгу.

## <a name="others"></a>Другие

* Установите имя столбца таблицы.
* Добавьте столбец таблицы в конец таблицы.
* Добавьте несколько строк в таблицу одновременно.
* `range.getColumnsAfter(count: number)` и `range.getColumnsBefore(count: number)`, чтобы вернуть определенное количество столбцов справа/слева от текущего объекта Range.
* Методы и свойства [ \* OrNullObject.](../../develop/application-specific-api-model.md#ornullobject-methods-and-properties)Эта функция позволяет получать объект с помощью ключа. Если объекта не существует, свойство возвращенного объекта `isNullObject` будет верным. Это позволяет разработчикам проверять, существует ли объект без обработки исключений. Метод `*OrNullObject` доступен на большинстве объектов коллекции.

```js
worksheet.getItemOrNullObject("itemName")
```

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в Excel API JavaScript, за набором 1.4. Чтобы просмотреть справочную документацию API для всех API, поддерживаемых Excel API JavaScript, установленного 1.4 или ранее, см. в Excel API в наборе требований [1.4](/javascript/api/excel?view=excel-js-1.4&preserve-view=true)или ранее .

| Класс | Поля | Описание |
|:---|:---|:---|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[getCount()](/javascript/api/excel/excel.bindingcollection#getCount__)|Получает количество привязок в коллекции.|
||[getItemOrNullObject(id: строка)](/javascript/api/excel/excel.bindingcollection#getItemOrNullObject_id_)|Возвращает объект привязки по идентификатору.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[getCount()](/javascript/api/excel/excel.chartcollection#getCount__)|Возвращает количество диаграмм на листе.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.chartcollection#getItemOrNullObject_name_)|Возвращает диаграмму по ее имени.|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[getCount()](/javascript/api/excel/excel.chartpointscollection#getCount__)|Возвращает количество точек диаграммы в ряду.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[getCount()](/javascript/api/excel/excel.chartseriescollection#getCount__)|Возвращает количество рядов в коллекции.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[comment](/javascript/api/excel/excel.nameditem#comment)|Указывает комментарий, связанный с этим именем.|
||[delete()](/javascript/api/excel/excel.nameditem#delete__)|Удаляет заданное имя.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.nameditem#getRangeOrNullObject__)|Возвращает объект Range, сопоставленный с именем.|
||[scope](/javascript/api/excel/excel.nameditem#scope)|Указывает, задано ли имя в книге или в определенной таблице.|
||[worksheet](/javascript/api/excel/excel.nameditem#worksheet)|Возвращает лист, к которому относится именованный элемент.|
||[worksheetOrNullObject](/javascript/api/excel/excel.nameditem#worksheetOrNullObject)|Возвращает таблицу, в которую область действия именуемой номенклатуры.|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[add(name: string, reference: Range \| string, comment?: string)](/javascript/api/excel/excel.nameditemcollection#add_name__reference__comment_)|Добавляет новое имя в определенную коллекцию.|
||[addFormulaLocal (имя: строка, формула: строка, комментарий?: строка)](/javascript/api/excel/excel.nameditemcollection#addFormulaLocal_name__formula__comment_)|Добавляет новое имя в определенную коллекцию, используя языковой стандарт пользователя для формулы.|
||[getCount()](/javascript/api/excel/excel.nameditemcollection#getCount__)|Получает количество именованных элементов в коллекции.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.nameditemcollection#getItemOrNullObject_name_)|Получает объект `NamedItem` с его именем.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getCount()](/javascript/api/excel/excel.pivottablecollection#getCount__)|Получает количество сводных таблиц в коллекции.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.pivottablecollection#getItemOrNullObject_name_)|Получает сводную таблицу по имени.|
|[Range](/javascript/api/excel/excel.range)|[getIntersectionOrNullObject (anotherRange: Range \| string)](/javascript/api/excel/excel.range#getIntersectionOrNullObject_anotherRange_)|Возвращает объект диапазона, представляющий прямоугольное пересечение заданных диапазонов.|
||[getUsedRangeOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.range#getUsedRangeOrNullObject_valuesOnly_)|Возвращает используемый диапазон заданного объекта диапазона.|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getCount()](/javascript/api/excel/excel.rangeviewcollection#getCount__)|Получает количество `RangeView` объектов в коллекции.|
|[Параметр](/javascript/api/excel/excel.setting)|[delete()](/javascript/api/excel/excel.setting#delete__)|Удаляет параметр.|
||[key](/javascript/api/excel/excel.setting#key)|Ключ, который представляет ID параметра.|
||[value](/javascript/api/excel/excel.setting#value)|Представляет значение, сохраненное для этого параметра.|
|[SettingCollection](/javascript/api/excel/excel.settingcollection)|[add(key: string, value: string \| number \| boolean \| Date Array \| \| any)](/javascript/api/excel/excel.settingcollection#add_key__value_)|Задает или добавляет указанный параметр в книгу.|
||[getCount()](/javascript/api/excel/excel.settingcollection#getCount__)|Получает количество параметров в коллекции.|
||[getItem(key: string)](/javascript/api/excel/excel.settingcollection#getItem_key_)|Получает запись параметра с помощью ключа.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.settingcollection#getItemOrNullObject_key_)|Получает запись параметра с помощью ключа.|
||[items](/javascript/api/excel/excel.settingcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[onSettingsChanged](/javascript/api/excel/excel.settingcollection#onSettingsChanged)|Возникает при смене параметров документа.|
|[SettingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|[settings](/javascript/api/excel/excel.settingschangedeventargs#settings)|Получает `Setting` объект, представляюющий привязку, которая подняла событие изменения параметров|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[getCount()](/javascript/api/excel/excel.tablecollection#getCount__)|Получает количество таблиц в коллекции.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablecollection#getItemOrNullObject_key_)|Получает таблицу по имени или ИД.|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[getCount()](/javascript/api/excel/excel.tablecolumncollection#getCount__)|Получает количество столбцов в таблице.|
||[getItemOrNullObject(key: number \| string)](/javascript/api/excel/excel.tablecolumncollection#getItemOrNullObject_key_)|Возвращает объект столбца по имени или идентификатору.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[getCount()](/javascript/api/excel/excel.tablerowcollection#getCount__)|Получает количество строк в таблице.|
|[Workbook](/javascript/api/excel/excel.workbook)|[settings](/javascript/api/excel/excel.workbook#settings)|Представляет коллекцию параметров, связанных с книгой.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[getUsedRangeOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.worksheet#getUsedRangeOrNullObject_valuesOnly_)|Используемый диапазон — это наименьший диапазон, включающий в себя все ячейки с определенным значением или форматированием.|
||[имена](/javascript/api/excel/excel.worksheet#names)|Коллекция имен, относящих к текущему листу.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[getCount (visibleOnly?: boolean)](/javascript/api/excel/excel.worksheetcollection#getCount_visibleOnly_)|Получает количество листов в коллекции.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcollection#getItemOrNullObject_key_)|Получает объект листа по его имени или ИД.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.4&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)
