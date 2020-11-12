---
title: Набор обязательных элементов API JavaScript для Excel 1,4
description: Сведения о наборе требований ExcelApi 1,4.
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 17e915eea2cddffc8c48735e5c9f628fffb4d072
ms.sourcegitcommit: ca66ff7462bfdf4ed7ae04f43d1388c24de63bf9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/11/2020
ms.locfileid: "48996468"
---
# <a name="whats-new-in-excel-javascript-api-14"></a>Новые возможности API JavaScript для Excel 1.4

Ниже перечислено то, что было недавно добавлено в набор обязательных элементов 1.4, относящийся к API JavaScript для Excel.

## <a name="named-item-add-and-new-properties"></a>Именованный элемент add и новые свойства

Новые свойства:

* `comment`
* `scope` — Элементы листа или книги.
* `worksheet` — Возвращает лист, на который распространяется именованный элемент.

Новые методы:

* `add(name: string, reference: Range or string, comment: string)` — Добавляет новое имя в коллекцию заданной области.
* `addFormulaLocal(name: string, formula: string, comment: string)` -Добавляет новое имя в коллекцию заданной области, используя языковой стандарт пользователя для формулы.

## <a name="settings-api-in-the-excel-namespace"></a>Параметры API в пространстве имен Excel

Объект [Setting](/javascript/api/excel/excel.setting) представляет пару "ключ-значение" для параметра, хранящегося в документе. Функциональные возможности объекта `Excel.Setting` аналогичны `Office.Settings`, но он использует пакетный синтаксис API, а не модель обратного вызова общего API.

Интерфейсы API включают `getItem()` в себя получение записи параметров с помощью ключа и `add()` Добавление указанной записи параметра key: value в книгу.

## <a name="others"></a>Другие

* Задайте имя столбца таблицы.
* Добавление столбца таблицы в конец таблицы.
* Добавление нескольких строк в таблицу за раз.
* `range.getColumnsAfter(count: number)` и `range.getColumnsBefore(count: number)`, чтобы вернуть определенное количество столбцов справа/слева от текущего объекта Range.
* [ \* Методы и свойства орнуллобжект](../../develop/application-specific-api-model.md#ornullobject-methods-and-properties): Эта функция позволяет получает объект с помощью ключа. Если объект не существует, свойство возвращаемого объекта `isNullObject` будет иметь значение true. Это позволяет разработчикам проверять, существует ли объект, без необходимости его обрабатывать с помощью обработки исключений. `*OrNullObject`Метод доступен для большинства объектов Collection.

```js
worksheet.getItemOrNullObject("itemName")
```

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в наборе обязательных элементов API JavaScript для Excel 1,4. Чтобы просмотреть справочную документацию по API для всех API, поддерживаемых набором обязательных элементов API JavaScript для Excel 1,4 или более ранней версии, обратитесь к разделам [API Excel в наборе требований 1,4](/javascript/api/excel?view=excel-js-1.4&preserve-view=true)

| Класс | Поля | Описание |
|:---|:---|:---|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[getCount()](/javascript/api/excel/excel.bindingcollection#getcount--)|Получает количество привязок в коллекции.|
||[getItemOrNullObject(id: строка)](/javascript/api/excel/excel.bindingcollection#getitemornullobject-id-)|Возвращает объект привязки по идентификатору.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[getCount()](/javascript/api/excel/excel.chartcollection#getcount--)|Возвращает количество диаграмм на листе.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.chartcollection#getitemornullobject-name-)|Возвращает диаграмму по ее имени.|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[getCount()](/javascript/api/excel/excel.chartpointscollection#getcount--)|Возвращает количество точек диаграммы в ряду.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[getCount()](/javascript/api/excel/excel.chartseriescollection#getcount--)|Возвращает количество рядов в коллекции.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[comment](/javascript/api/excel/excel.nameditem#comment)|Задает комментарий, связанный с этим именем.|
||[delete()](/javascript/api/excel/excel.nameditem#delete--)|Удаляет заданное имя.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.nameditem#getrangeornullobject--)|Возвращает объект Range, сопоставленный с именем.|
||[scope](/javascript/api/excel/excel.nameditem#scope)|Указывает, ограничивается ли имя книгой или определенным листом.|
||[worksheet](/javascript/api/excel/excel.nameditem#worksheet)|Возвращает лист, к которому относится именованный элемент.|
||[worksheetOrNullObject](/javascript/api/excel/excel.nameditem#worksheetornullobject)|Возвращает лист, к которому относится именованный элемент.|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[Add (имя: строка, ссылка: \| строка диапазона, комментарий?: строка)](/javascript/api/excel/excel.nameditemcollection#add-name--reference--comment-)|Добавляет новое имя в определенную коллекцию.|
||[addFormulaLocal (имя: строка, формула: строка, Примечание?: строка)](/javascript/api/excel/excel.nameditemcollection#addformulalocal-name--formula--comment-)|Добавляет новое имя в определенную коллекцию, используя языковой стандарт пользователя для формулы.|
||[getCount()](/javascript/api/excel/excel.nameditemcollection#getcount--)|Получает количество именованных элементов в коллекции.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.nameditemcollection#getitemornullobject-name-)|Возвращает объект NamedItem, используя его имя.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getCount()](/javascript/api/excel/excel.pivottablecollection#getcount--)|Получает количество сводных таблиц в коллекции.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.pivottablecollection#getitemornullobject-name-)|Получает сводную таблицу по имени.|
|[Range](/javascript/api/excel/excel.range)|[getIntersectionOrNullObject (anotherRange: \| строка Range)](/javascript/api/excel/excel.range#getintersectionornullobject-anotherrange-)|Возвращает объект диапазона, представляющий прямоугольное пересечение заданных диапазонов.|
||[getUsedRangeOrNullObject (valuesOnly?: Boolean)](/javascript/api/excel/excel.range#getusedrangeornullobject-valuesonly-)|Возвращает используемый диапазон заданного объекта диапазона.|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getCount()](/javascript/api/excel/excel.rangeviewcollection#getcount--)|Получает количество объектов RangeView в коллекции.|
|[Параметр](/javascript/api/excel/excel.setting)|[delete()](/javascript/api/excel/excel.setting#delete--)|Удаляет параметр.|
||[key](/javascript/api/excel/excel.setting#key)|Ключ, представляющий идентификатор параметра.|
||[value](/javascript/api/excel/excel.setting#value)|Представляет значение, сохраненное для этого параметра.|
|[SettingCollection](/javascript/api/excel/excel.settingcollection)|[Add (ключ: строка, значение: строка \| Number \| Boolean \| \| массив дат <any> \| Any)](/javascript/api/excel/excel.settingcollection#add-key--value-)|Задает или добавляет указанный параметр в книгу.|
||[getCount()](/javascript/api/excel/excel.settingcollection#getcount--)|Получает количество параметров в коллекции.|
||[getItem(key: string)](/javascript/api/excel/excel.settingcollection#getitem-key-)|Получает запись Setting по ключу.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.settingcollection#getitemornullobject-key-)|Получает запись Setting по ключу.|
||[items](/javascript/api/excel/excel.settingcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[onSettingsChanged](/javascript/api/excel/excel.settingcollection#onsettingschanged)|Возникает при изменении параметров в документе.|
|[SettingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|[settings](/javascript/api/excel/excel.settingschangedeventargs#settings)|Получает объект Setting, представляющий привязку, которая вызвала событие SettingsChanged.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[getCount()](/javascript/api/excel/excel.tablecollection#getcount--)|Получает количество таблиц в коллекции.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablecollection#getitemornullobject-key-)|Получает таблицу по имени или идентификатору.|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[getCount()](/javascript/api/excel/excel.tablecolumncollection#getcount--)|Получает количество столбцов в таблице.|
||[getItemOrNullObject (Key: номер \| строки)](/javascript/api/excel/excel.tablecolumncollection#getitemornullobject-key-)|Возвращает объект column по имени или идентификатору.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[getCount()](/javascript/api/excel/excel.tablerowcollection#getcount--)|Получает количество строк в таблице.|
|[Workbook](/javascript/api/excel/excel.workbook)|[settings](/javascript/api/excel/excel.workbook#settings)|Представляет коллекцию параметров, сопоставленных с книгой.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[getUsedRangeOrNullObject (valuesOnly?: Boolean)](/javascript/api/excel/excel.worksheet#getusedrangeornullobject-valuesonly-)|Используемый диапазон — это наименьший диапазон, включающий в себя все ячейки с определенным значением или форматированием.|
||[псевдоним](/javascript/api/excel/excel.worksheet#names)|Коллекция имен, относящих к текущему листу.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[NOCOUNT (visibleOnly?: Boolean)](/javascript/api/excel/excel.worksheetcollection#getcount-visibleonly-)|Получает количество листов в коллекции.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcollection#getitemornullobject-key-)|Получает объект листа по его имени или ИД.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.4&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)
