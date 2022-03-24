---
title: Excel API JavaScript установлено 1.4
description: Сведения о наборе требований ExcelApi 1.4.
ms.date: 11/09/2020
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: bcdbd044c5de562b7c2cc2bc9971af31179f8a9b
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/23/2022
ms.locfileid: "63746542"
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

API включают в `getItem()` себя создание записи с помощью ключа `add()` и добавление указанной пары параметров ключа:значения в книгу.

## <a name="others"></a>Другие

* Установите имя столбца таблицы.
* Добавьте столбец таблицы в конец таблицы.
* Добавьте несколько строк в таблицу одновременно.
* `range.getColumnsAfter(count: number)` и `range.getColumnsBefore(count: number)`, чтобы вернуть определенное количество столбцов справа/слева от текущего объекта Range.
* Методы [\*и свойства OrNullObject](../../develop/application-specific-api-model.md#ornullobject-methods-and-properties). Эта функция позволяет получать объект с помощью ключа. Если объекта не существует, `isNullObject` свойство возвращенного объекта будет верным. Это позволяет разработчикам проверять, существует ли объект без обработки исключений. Метод `*OrNullObject` доступен на большинстве объектов коллекции.

```js
worksheet.getItemOrNullObject("itemName")
```

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в Excel API JavaScript, за набором 1.4. Чтобы просмотреть справочную документацию API для всех API, поддерживаемых Excel API JavaScript, установленного 1.4 или ранее, см. Excel API в наборе требований [1.4 или ранее](/javascript/api/excel?view=excel-js-1.4&preserve-view=true).

| Класс | Поля | Описание |
|:---|:---|:---|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[getCount()](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-getcount-member(1))|Получает количество привязок в коллекции.|
||[getItemOrNullObject(id: строка)](/javascript/api/excel/excel.bindingcollection#excel-excel-bindingcollection-getitemornullobject-member(1))|Возвращает объект привязки по идентификатору.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[getCount()](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-getcount-member(1))|Возвращает количество диаграмм на листе.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.chartcollection#excel-excel-chartcollection-getitemornullobject-member(1))|Возвращает диаграмму по ее имени.|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[getCount()](/javascript/api/excel/excel.chartpointscollection#excel-excel-chartpointscollection-getcount-member(1))|Возвращает количество точек диаграммы в ряду.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[getCount()](/javascript/api/excel/excel.chartseriescollection#excel-excel-chartseriescollection-getcount-member(1))|Возвращает количество рядов в коллекции.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[comment](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-comment-member)|Указывает комментарий, связанный с этим именем.|
||[delete()](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-delete-member(1))|Удаляет заданное имя.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-getrangeornullobject-member(1))|Возвращает объект Range, сопоставленный с именем.|
||[scope](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-scope-member)|Указывает, задано ли имя в книге или в определенной таблице.|
||[worksheet](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-worksheet-member)|Возвращает лист, к которому относится именованный элемент.|
||[worksheetOrNullObject](/javascript/api/excel/excel.nameditem#excel-excel-nameditem-worksheetornullobject-member)|Возвращает таблицу, в которую область действия именуемой номенклатуры.|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[add(name: string, reference: Range \| string, comment?: string)](/javascript/api/excel/excel.nameditemcollection#excel-excel-nameditemcollection-add-member(1))|Добавляет новое имя в определенную коллекцию.|
||[addFormulaLocal (имя: строка, формула: строка, комментарий?: строка)](/javascript/api/excel/excel.nameditemcollection#excel-excel-nameditemcollection-addformulalocal-member(1))|Добавляет новое имя в определенную коллекцию, используя языковой стандарт пользователя для формулы.|
||[getCount()](/javascript/api/excel/excel.nameditemcollection#excel-excel-nameditemcollection-getcount-member(1))|Получает количество именованных элементов в коллекции.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.nameditemcollection#excel-excel-nameditemcollection-getitemornullobject-member(1))|Получает объект `NamedItem` с его именем.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getCount()](/javascript/api/excel/excel.pivottablecollection#excel-excel-pivottablecollection-getcount-member(1))|Получает количество сводных таблиц в коллекции.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.pivottablecollection#excel-excel-pivottablecollection-getitemornullobject-member(1))|Получает сводную таблицу по имени.|
|[Range](/javascript/api/excel/excel.range)|[getIntersectionOrNullObject (anotherRange: Range \| string)](/javascript/api/excel/excel.range#excel-excel-range-getintersectionornullobject-member(1))|Возвращает объект диапазона, представляющий прямоугольное пересечение заданных диапазонов.|
||[getUsedRangeOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.range#excel-excel-range-getusedrangeornullobject-member(1))|Возвращает используемый диапазон заданного объекта диапазона.|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getCount()](/javascript/api/excel/excel.rangeviewcollection#excel-excel-rangeviewcollection-getcount-member(1))|Получает количество объектов `RangeView` в коллекции.|
|[Параметр](/javascript/api/excel/excel.setting)|[delete()](/javascript/api/excel/excel.setting#excel-excel-setting-delete-member(1))|Удаляет параметр.|
||[key](/javascript/api/excel/excel.setting#excel-excel-setting-key-member)|Ключ, который представляет ID параметра.|
||[value](/javascript/api/excel/excel.setting#excel-excel-setting-value-member)|Представляет значение, сохраненное для этого параметра.|
|[SettingCollection](/javascript/api/excel/excel.settingcollection)|[add(key: string, value: string \| number \| boolean \| Date \| Array \| any)](/javascript/api/excel/excel.settingcollection#excel-excel-settingcollection-add-member(1))|Задает или добавляет указанный параметр в книгу.|
||[getCount()](/javascript/api/excel/excel.settingcollection#excel-excel-settingcollection-getcount-member(1))|Получает количество параметров в коллекции.|
||[getItem(key: string)](/javascript/api/excel/excel.settingcollection#excel-excel-settingcollection-getitem-member(1))|Получает запись параметра с помощью ключа.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.settingcollection#excel-excel-settingcollection-getitemornullobject-member(1))|Получает запись параметра с помощью ключа.|
||[items](/javascript/api/excel/excel.settingcollection#excel-excel-settingcollection-items-member)|Получает загруженные дочерние элементы в этой коллекции.|
||[onSettingsChanged](/javascript/api/excel/excel.settingcollection#excel-excel-settingcollection-onsettingschanged-member)|Возникает при смене параметров документа.|
|[SettingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|[settings](/javascript/api/excel/excel.settingschangedeventargs#excel-excel-settingschangedeventargs-settings-member)|Получает объект `Setting` , представляюющий привязку, которая подняла событие изменения параметров|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[getCount()](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-getcount-member(1))|Получает количество таблиц в коллекции.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablecollection#excel-excel-tablecollection-getitemornullobject-member(1))|Получает таблицу по имени или ИД.|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[getCount()](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-getcount-member(1))|Получает количество столбцов в таблице.|
||[getItemOrNullObject(key: number \| string)](/javascript/api/excel/excel.tablecolumncollection#excel-excel-tablecolumncollection-getitemornullobject-member(1))|Возвращает объект столбца по имени или идентификатору.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[getCount()](/javascript/api/excel/excel.tablerowcollection#excel-excel-tablerowcollection-getcount-member(1))|Получает количество строк в таблице.|
|[Workbook](/javascript/api/excel/excel.workbook)|[settings](/javascript/api/excel/excel.workbook#excel-excel-workbook-settings-member)|Представляет коллекцию параметров, связанных с книгой.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[getUsedRangeOrNullObject(valuesOnly?: boolean)](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-getusedrangeornullobject-member(1))|Используемый диапазон — это наименьший диапазон, включающий в себя все ячейки с определенным значением или форматированием.|
||[имена](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-names-member)|Коллекция имен, относящих к текущему листу.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[getCount (visibleOnly?: boolean)](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-getcount-member(1))|Получает количество листов в коллекции.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcollection#excel-excel-worksheetcollection-getitemornullobject-member(1))|Получает объект листа по его имени или ИД.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.4&preserve-view=true)
- [Наборы обязательных элементов API JavaScript для Excel](excel-api-requirement-sets.md)
