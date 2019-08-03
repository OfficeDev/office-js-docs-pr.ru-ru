---
title: Набор обязательных элементов API JavaScript для Excel 1,4
description: Сведения о наборе требований ExcelApi 1,4
ms.date: 07/26/2019
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 8dd672dc8ff30a47f7858934d995d809070eb3cf
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064741"
---
# <a name="whats-new-in-excel-javascript-api-14"></a>Новые возможности API JavaScript для Excel 1.4

Ниже перечислено то, что было недавно добавлено в набор обязательных элементов 1.4, относящийся к API JavaScript для Excel.

## <a name="named-item-add-and-new-properties"></a>Именованный элемент add и новые свойства

Новые свойства:

* `comment`
* `scope`— Элементы листа или книги.
* `worksheet`— Возвращает лист, на который распространяется именованный элемент.

Новые методы:

* `add(name: string, reference: Range or string, comment: string)`— Добавляет новое имя в коллекцию заданной области.
* `addFormulaLocal(name: string, formula: string, comment: string)`-Добавляет новое имя в коллекцию заданной области, используя языковой стандарт пользователя для формулы.

## <a name="settings-api-in-the-excel-namespace"></a>Параметры API в пространстве имен Excel

Объект [Setting](/javascript/api/excel/excel.setting) представляет пару "ключ-значение" для параметра, хранящегося в документе. Функциональные возможности объекта `Excel.Setting` аналогичны `Office.Settings`, но он использует пакетный синтаксис API, а не модель обратного вызова общего API.

Интерфейсы API `getItem()` включают в себя получение записи параметров с помощью `add()` ключа и добавление указанной записи параметра key: value в книгу.

## <a name="others"></a>Другие

* Задайте имя столбца таблицы.
* Добавление столбца таблицы в конец таблицы.
* Добавление нескольких строк в таблицу за раз.
* `range.getColumnsAfter(count: number)` и `range.getColumnsBefore(count: number)`, чтобы вернуть определенное количество столбцов справа/слева от текущего объекта Range.
* [Функция "получить элемент" или "null Object](../../excel/excel-add-ins-advanced-concepts.md#ornullobject-methods)": Эта функция позволяет получать объект с помощью ключа. Если объект не существует, `isNullObject` свойство возвращаемого объекта будет иметь значение true. Это позволяет разработчикам проверять, существует ли объект, но не обрабатывать его с помощью обработки исключений. `*OrNullObject` Метод доступен для большинства объектов Collection.

```javascript
worksheet.getItemOrNullObject("itemName")
```

## <a name="api-list"></a>Список API

В следующей таблице перечислены API в наборе обязательных элементов API JavaScript для Excel 1,4. Чтобы просмотреть справочную документацию по API для всех API, поддерживаемых набором обязательных элементов API JavaScript для Excel 1,4 или более ранней версии, обратитесь к разделам [API Excel в наборе требований 1,4](/javascript/api/excel?view=excel-js-1.4)

| Класс | Поля | Описание |
|:---|:---|:---|
|[BindingCollection](/javascript/api/excel/excel.bindingcollection)|[getCount()](/javascript/api/excel/excel.bindingcollection#getcount--)|Получает количество привязок в коллекции.|
||[getItemOrNullObject(id: строка)](/javascript/api/excel/excel.bindingcollection#getitemornullobject-id-)|Получает объект привязки по идентификатору. Если объект привязки не существует, возвращает пустой объект.|
|[ChartCollection](/javascript/api/excel/excel.chartcollection)|[getCount()](/javascript/api/excel/excel.chartcollection#getcount--)|Возвращает количество диаграмм на листе.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.chartcollection#getitemornullobject-name-)|Возвращает диаграмму по ее имени. Если одно и то же имя принадлежит нескольким диаграммам, будет возвращена первая из них.|
|[ChartPointsCollection](/javascript/api/excel/excel.chartpointscollection)|[getCount()](/javascript/api/excel/excel.chartpointscollection#getcount--)|Возвращает количество точек диаграммы в ряду.|
|[ChartSeriesCollection](/javascript/api/excel/excel.chartseriescollection)|[getCount()](/javascript/api/excel/excel.chartseriescollection#getcount--)|Возвращает количество рядов в коллекции.|
|[NamedItem](/javascript/api/excel/excel.nameditem)|[comment](/javascript/api/excel/excel.nameditem#comment)|Представляет примечание, связанное с этим именем.|
||[delete()](/javascript/api/excel/excel.nameditem#delete--)|Удаляет заданное имя.|
||[getRangeOrNullObject()](/javascript/api/excel/excel.nameditem#getrangeornullobject--)|Возвращает объект диапазона, связанный с именем. Возвращает пустой объект, если именованный элемент не является диапазоном.|
||[scope](/javascript/api/excel/excel.nameditem#scope)|Указывает, относится ли имя к книге или определенному листу. Возможные значения: лист, книга. Только для чтения.|
||[worksheet](/javascript/api/excel/excel.nameditem#worksheet)|Возвращает лист, к которому относится именованный элемент. Выдает ошибку, если элемент находится в области действия книги.|
||[worksheetOrNullObject](/javascript/api/excel/excel.nameditem#worksheetornullobject)|Возвращает лист, к которому относится именованный элемент. Возвращает пустой объект, если элемент относится к книге.|
|[NamedItemCollection](/javascript/api/excel/excel.nameditemcollection)|[Add (имя: строка, ссылка: строка \| диапазона, комментарий?: строка)](/javascript/api/excel/excel.nameditemcollection#add-name--reference--comment-)|Добавляет новое имя в определенную коллекцию.|
||[addFormulaLocal (имя: строка, формула: строка, Примечание?: строка)](/javascript/api/excel/excel.nameditemcollection#addformulalocal-name--formula--comment-)|Добавляет новое имя в определенную коллекцию, используя языковой стандарт пользователя для формулы.|
||[getCount()](/javascript/api/excel/excel.nameditemcollection#getcount--)|Получает количество именованных элементов в коллекции.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.nameditemcollection#getitemornullobject-name-)|Возвращает объект NamedItem, используя его имя. Если объект nameditem не существует, возвращает пустой объект.|
|[PivotTableCollection](/javascript/api/excel/excel.pivottablecollection)|[getCount()](/javascript/api/excel/excel.pivottablecollection#getcount--)|Получает количество сводных таблиц в коллекции.|
||[getItemOrNullObject(имя: строка)](/javascript/api/excel/excel.pivottablecollection#getitemornullobject-name-)|Получает сводную таблицу по имени. Если сводная таблица не существует, возвращает пустой объект.|
|[Range](/javascript/api/excel/excel.range)|[getIntersectionOrNullObject (anotherRange: строка \| Range)](/javascript/api/excel/excel.range#getintersectionornullobject-anotherrange-)|Возвращает объект диапазона, представляющий прямоугольное пересечение заданных диапазонов. Если пересечение не найдено, возвращает пустой объект.|
||[getUsedRangeOrNullObject (valuesOnly?: Boolean)](/javascript/api/excel/excel.range#getusedrangeornullobject-valuesonly-)|Возвращает используемый диапазон заданного объекта диапазона. Если в диапазоне нет используемых ячеек, эта функция возвращает пустой объект.|
|[RangeViewCollection](/javascript/api/excel/excel.rangeviewcollection)|[getCount()](/javascript/api/excel/excel.rangeviewcollection#getcount--)|Получает количество объектов RangeView в коллекции.|
|[Параметр](/javascript/api/excel/excel.setting)|[delete()](/javascript/api/excel/excel.setting#delete--)|Удаляет параметр.|
||[key](/javascript/api/excel/excel.setting#key)|Возвращает ключ, представляющий идентификатор setting. Только для чтения.|
||[value](/javascript/api/excel/excel.setting#value)|Представляет значение, сохраненное для этого параметра.|
|[SettingCollection](/javascript/api/excel/excel.settingcollection)|[Add (ключ: строка, значение: строка \| Number \| Boolean \| массив \| <any> \| дат Any)](/javascript/api/excel/excel.settingcollection#add-key--value-)|Задает или добавляет указанный параметр в книгу.|
||[getCount()](/javascript/api/excel/excel.settingcollection#getcount--)|Получает количество параметров в коллекции.|
||[getItem(key: string)](/javascript/api/excel/excel.settingcollection#getitem-key-)|Получает запись Setting по ключу.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.settingcollection#getitemornullobject-key-)|Возвращает объект Setting по ключу. Если параметр не существует, возвращает пустой объект.|
||[items](/javascript/api/excel/excel.settingcollection#items)|Получает загруженные дочерние элементы в этой коллекции.|
||[onSettingsChanged](/javascript/api/excel/excel.settingcollection#onsettingschanged)|Возникает при изменении параметров в документе.|
|[SettingsChangedEventArgs](/javascript/api/excel/excel.settingschangedeventargs)|[settings](/javascript/api/excel/excel.settingschangedeventargs#settings)|Получает объект Setting, представляющий привязку, которая вызвала событие SettingsChanged.|
|[TableCollection](/javascript/api/excel/excel.tablecollection)|[getCount()](/javascript/api/excel/excel.tablecollection#getcount--)|Получает количество таблиц в коллекции.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.tablecollection#getitemornullobject-key-)|Получает таблицу по имени или ИД. Если таблица не существует, возвращает пустой объект.|
|[TableColumnCollection](/javascript/api/excel/excel.tablecolumncollection)|[getCount()](/javascript/api/excel/excel.tablecolumncollection#getcount--)|Получает количество столбцов в таблице.|
||[getItemOrNullObject (Key: номер \| строки)](/javascript/api/excel/excel.tablecolumncollection#getitemornullobject-key-)|Возвращает объект столбца по имени или ИД. Если столбец не существует, возвращает пустой объект.|
|[TableRowCollection](/javascript/api/excel/excel.tablerowcollection)|[getCount()](/javascript/api/excel/excel.tablerowcollection#getcount--)|Получает количество строк в таблице.|
|[Workbook](/javascript/api/excel/excel.workbook)|[settings](/javascript/api/excel/excel.workbook#settings)|Представляет коллекцию параметров, сопоставленных с книгой. Только для чтения.|
|[Worksheet](/javascript/api/excel/excel.worksheet)|[getUsedRangeOrNullObject (valuesOnly?: Boolean)](/javascript/api/excel/excel.worksheet#getusedrangeornullobject-valuesonly-)|Используемый диапазон — это наименьший диапазон, включающий в себя все ячейки, которые содержат значение или форматирование. Если весь лист пустой, эта функция возвращает пустой объект.|
||[псевдоним](/javascript/api/excel/excel.worksheet#names)|Коллекция имен, относящих к текущему листу. Только для чтения.|
|[WorksheetCollection](/javascript/api/excel/excel.worksheetcollection)|[NOCOUNT (visibleOnly?: Boolean)](/javascript/api/excel/excel.worksheetcollection#getcount-visibleonly-)|Получает количество листов в коллекции.|
||[getItemOrNullObject(key: string)](/javascript/api/excel/excel.worksheetcollection#getitemornullobject-key-)|Получает объект листа по его имени или ИД. Если лист не существует, возвращает пустой объект.|

## <a name="see-also"></a>См. также

- [Справочная документация по API JavaScript для Excel](/javascript/api/excel?view=excel-js-1.4)
- [Наборы обязательных элементов API JavaScript для Excel](./excel-api-requirement-sets.md)
