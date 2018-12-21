---
title: Работа с несколькими диапазонами одновременно в надстройках Excel
description: ''
ms.date: 09/04/2018
ms.openlocfilehash: f1217fc76d14269882a73ec5eb7758e519563456
ms.sourcegitcommit: 6870f0d96ed3da2da5a08652006c077a72d811b6
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/21/2018
ms.locfileid: "27383227"
---
# <a name="work-with-multiple-ranges-simultaneously-in-excel-add-ins-preview"></a>Работа с несколькими диапазонами одновременно в надстройках Excel (предварительная версия)

Библиотека JavaScript для Excel позволяет вашей надстройке выполнять операции и устанавливать свойства одновременно для нескольких диапазонов. Диапазоны необязательно должны быть смежными. Этот способ установки свойства не только упрощает код, но и выполняется намного быстрее, чем установка одинакового свойства отдельно для каждого диапазона.

> [!NOTE]
> Для работы с API-интерфейсами, описанными в этой статье, требуется **Office 2016 "нажми и работай" версии 1809 сборки 10820.20000** или более поздней версии (возможно, вам придется принять участие в [программе предварительной оценки Office](https://products.office.com/office-insider) для получения нужной сборки). Кроме того, необходимо загрузить бета-версию библиотеки JavaScript для Office из сети [CDN Office.js](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js). В настоящее время нет справочных страниц для этих API. Но следующий файл типа определения содержит их описания: [бета-версия office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).

## <a name="rangeareas"></a>RangeAreas

Набор диапазонов (возможно, несмежных) представлен объектом `Excel.RangeAreas`. Его свойства и методы аналогичны типу `Range` (многие с одинаковыми или похожими именами), но с изменением указанных ниже параметров:

- Типы данных для свойств и поведений методов задания и методов получения.
- Типы данных параметров метода и поведений метода.
- Типы данных возвращаемых значений метода.

Примеры:

- У `RangeAreas` есть свойство `address`, возвращающее строку с адресами диапазона, разделенными запятой, а не только один адрес, как в случае со свойством `Range.address`.
- У `RangeAreas` есть свойство `dataValidation`, которое возвращает объект `DataValidation`, представляющий проверку данных всех диапазонов в `RangeAreas` при соответствии. Значение этого свойства будет равно `null`, если ко всем диапазонам в `RangeAreas` не применяются одинаковые объекты `DataValidation`. Это общий, но не универсальный принцип для объекта `RangeAreas`: *если у свойства нет согласованных значений во всех диапазонах в `RangeAreas`, его значением будет `null`.* Дополнительные сведения и некоторые исключения см. в разделе [Чтение свойств RangeAreas](#read-properties-of-rangeareas).
- `RangeAreas.cellCount` получает общее количество ячеек во всех диапазонах в `RangeAreas`.
- `RangeAreas.calculate` пересчитывает ячейки всех диапазонов в `RangeAreas`.
- `RangeAreas.getEntireColumn` и `RangeAreas.getEntireRow` возвращают другой объект `RangeAreas`, представляющий все столбцы (или строки) во всех диапазонах в `RangeAreas`. Например, если `RangeAreas` представляет "A1:C4" и "F14:L15", то `RangeAreas.getEntireColumn` возвращает объект `RangeAreas`, представляющий "A:C" и "F:L".
- `RangeAreas.copyFrom` может использовать параметр `Range` или `RangeAreas`, представляющий исходный диапазон (или диапазоны) операции копирования.

#### <a name="complete-list-of-range-members-that-are-also-available-on-rangeareas"></a>Полный список элементов Range, также доступных в RangeAreas

##### <a name="properties"></a>Свойства

Ознакомьтесь со статьей [Чтение свойств RangeAreas](#read-properties-of-rangeareas) перед написанием кода, считывающего любое из перечисленных свойств. Возвращаемое значение зависит от ряда факторов.

- address
- addressLocal
- cellCount
- conditionalFormats
- context
- dataValidation
- format
- isEntireColumn
- isEntireRow
- style
- worksheet

##### <a name="methods"></a>Методы

Методы Range в предварительной версии помечены.

- calculate()
- clear()
- convertDataTypeToText() (предварительная версия)
- convertToLinkedDataType() (предварительная версия)
- copyFrom() (предварительная версия)
- getEntireColumn()
- getEntireRow()
- getIntersection()
- getIntersectionOrNullObject()
- getOffsetRange() (называется getOffsetRangeAreas в объекте RangeAreas)
- getSpecialCells() (предварительная версия)
- getSpecialCellsOrNullObject() (предварительная версия)
- getTables() (предварительная версия)
- getUsedRange() (называется getUsedRangeAreas в объекте RangeAreas)
- getUsedRangeOrNullObject() (называется getUsedRangeAreasOrNullObject в объекте RangeAreas)
- load()
- set()
- setDirty() (предварительная версия)
- toJSON()
- track()
- untrack()

### <a name="rangearea-specific-properties-and-methods"></a>Свойства и методы, характерные для объекта RangeArea

Для типа `RangeAreas` существуют несколько свойств и методов, отсутствующих в объекте `Range`. Ниже приведены некоторые из них.

- `areas`. Объект `RangeCollection`, содержащий все диапазоны, которые представлены объектом `RangeAreas`. Объект `RangeCollection` — еще один новый объект, аналогичный другим объектам коллекции Excel. У него есть свойство `items`, являющееся массивом объектов `Range`, которые представляют диапазоны.
- `areaCount`. Общее количество диапазонов в `RangeAreas`.
- `getOffsetRangeAreas`. Действует аналогично методу [Range.getOffsetRange](/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), за исключением того, что возвращается объект `RangeAreas`, содержащий диапазоны, каждый из которых смещен относительно одного из диапазонов в исходном объекте `RangeAreas`.

## <a name="create-rangeareas-and-set-properties"></a>Создание RangeAreas и установка свойств

Объект `RangeAreas` можно создать двумя основными способами:

- Вызвать метод `Worksheet.getRanges()` и передать ему строку с адресами диапазона, разделенными запятыми. Если диапазон, который нужно включить, был преобразован [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem), в строку можно включить имя вместо адреса.
- Вызвать метод `Workbook.getSelectedRanges()`. Этот метод возвращает объект `RangeAreas`, представляющий все диапазоны, выбранные в активном на данный момент листе.

После получения объекта `RangeAreas` можно создать другие с помощью методов объекта, возвращающих `RangeAreas`, например `getOffsetRangeAreas` и `getIntersection`.

> [!NOTE]
> Нельзя напрямую добавить дополнительные диапазоны к объекту `RangeAreas`. Например, у коллекции в `RangeAreas.areas` нет метода `add`.

> [!WARNING]
> Не пытайтесь напрямую добавлять или удалять элементы из массива `RangeAreas.areas.items`. Это приведет к нежелательному поведению кода.  Например, существует возможность принудительно добавить дополнительный объект `Range` в массив, но это приведет к ошибкам, поскольку свойства и методы `RangeAreas` действуют, как будто новый элемент не был добавлен. Например, свойство `areaCount` не включает диапазоны, принудительно добавленные таким образом, а `RangeAreas.getItemAt(index)` вызывает ошибку, если `index` больше, чем `areasCount-1`.  Аналогичным образом, удаление объекта `Range` в массиве `RangeAreas.areas.items` путем получения ссылки на него и вызова его метода `Range.delete` приводит к ошибкам: хотя объект `Range` *удален*, свойства и методы родительского объекта `RangeAreas` будут действовать (или пытаться действовать), как будто он еще существует. Например, если код вызывает метод `RangeAreas.calculate`, Office попытается рассчитать диапазон, но это завершится ошибкой, поскольку объект range отсутствует.

Установка свойства для `RangeAreas` задает соответствующее свойство для всех диапазонов в коллекции `RangeAreas.areas`.

Ниже приведен пример установки свойства в нескольких диапазонах. Функция выделяет диапазоны **F3:F5** и **H3:H5**.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    return context.sync();
})
```

Этот пример применяется к сценариям, в которых можно жестко задать адреса диапазонов, передаваемых в `getRanges`, или легко рассчитать их во время выполнения. Ниже перечислены некоторые сценарии, в которых это возможно: 

- Код выполняется в контексте известного шаблона.
- Код выполняется в контексте импортированных данных, в котором известна схема данных.

Если при создании кода неизвестно, с какими диапазонами придется работать, необходимо обнаружить их во время выполнения. В следующем разделе рассматриваются эти сценарии.

### <a name="discover-range-areas-programmatically"></a>Обнаружение областей диапазона программным способом

Методы `Range.getSpecialCells()` и `Range.getSpecialCellsOrNullObject()` позволяют во время выполнения обнаруживать диапазоны, с которыми нужно работать, на основе характеристик ячеек и типа значений в ячейках. Подписи методов из файла типов данных TypeScript:

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

Ниже приведен пример использования первого из них. Вот что нужно знать об этом коде:

- Он ограничивает часть листа, в которой требуется выполнять поиск, путем вызова сначала метода `Worksheet.getUsedRange`, а затем метода `getSpecialCells` только для этого диапазона.
- В качестве параметра для `getSpecialCells` он передает строковое представление значения из перечисления `Excel.SpecialCellType`. Некоторые другие значения, которые могут быть переданы вместо этого: "Blanks" для пустых ячеек, "Constants" для ячеек со значениями литералов вместо формул и "SameConditionalFormat" для ячеек, у которых такое же условное форматирование, как и у первой ячейки в `usedRange`. Первая ячейка — это верхняя крайняя ячейка слева. Полный список значений перечисления см. в [бета-версии office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).
- Метод `getSpecialCells` возвращает объект `RangeAreas`, поэтому все ячейки с формулами окрашены розовым цветом даже в том случае, если они не являются смежными. 

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var formulaRanges = usedRange.getSpecialCells("Formulas");
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

В некоторых случаях диапазон не содержит *ни одной* ячейки с целевой характеристикой. Если метод `getSpecialCells` не находит ни одной такой ячейки, он выдает ошибку **ItemNotFound**. Это приведет к переадресации потока управления к блоку/методу `catch`, если таковой существует. Если нет, ошибка останавливает исполнение функции. Могут существовать сценарии, в которых выдача ошибки – это именно то, что должно происходить при отсутствии ячейки с целевой характеристикой. 

Если в сценариях отсутствие соответствующих ячеек является нормальной, но редкой ситуацией, ваш код должен проверить наличие такой возможности и корректно выполнить действие без выдачи ошибки. Для этих сценариев следует использовать метод `getSpecialCellsOrNullObject` и протестировать свойство `RangeAreas.isNullObject`. Ниже приведен пример. Вот что нужно знать об этом коде:

- Метод `getSpecialCellsOrNullObject` всегда возвращает прокси-объект, поэтому он не может иметь значение `null` в обычном смысле JavaScript. Но если соответствующие ячейки не обнаружены, свойству `isNullObject` объекта присваивается значение `true`.
- Он вызывает `context.sync` *перед* тестированием свойства `isNullObject`. Это требование для всех методов и свойств `*OrNullObject`, так как всегда нужно загружать и синхронизировать свойство, чтобы его прочесть. Однако необязательно *явно* загружать свойство `isNullObject`. Оно автоматически загружается с помощью `context.sync`, даже если `load` не вызывается для объекта. Дополнительные сведения см. в разделе [\*OrNullObject](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods).
- Этот код можно проверить, выбрав сначала диапазон без ячеек с формулами и запустив его. Затем следует выбрать диапазон, содержащий по крайней мере одну ячейку с формулой, и снова запустить его.

```js
Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    var formulaRanges = range.getSpecialCellsOrNullObject("Formulas");
    return context.sync()
        .then(function() {
            if (formulaRanges.isNullObject) {
                console.log("No cells have formulas");
            }
            else {
                formulaRanges.format.fill.color = "pink";
            }
        })
        .then(context.sync);
})
```

Для удобства во всех других примерах в этой статье используйте метод `getSpecialCells` вместо `getSpecialCellsOrNullObject`.

#### <a name="narrow-the-target-cells-with-cell-value-types"></a>Ограничение целевых ячеек с помощью типа значений ячеек

Существует необязательный второй параметр типа перечисления `Excel.SpecialCellValueType`, который дополнительно ограничивает целевые ячейки. Его можно использовать только в том случае, если передается значение "Formulas" или "Constants" для `getSpecialCells` или `getSpecialCellsOrNullObject`. Этот параметр указывает, что требуются только ячейки с определенными типами значений. Существует четыре основных типа: "Error", "Logical" (логический), "Numbers" и "Text" (у перечисления есть другие значения помимо этих четырех, которые рассматриваются ниже). Ниже приведен пример. Вот что нужно знать об этом коде:

- Он выделяет только ячейки с числовым значением литерала. Он не выделяет ячейки с формулой (даже если результат является числом), логическим значением, текстовым значением или ячейки с состоянием ошибки.
- Чтобы протестировать код, убедитесь, что в листе есть ячейки с числовыми значениями литералов, ячейки с другими значениями литералов и ячейки с формулами.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var constantNumberRanges = usedRange.getSpecialCells("Constants", "Numbers");
    constantNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

Иногда требуется работать с ячейками, имеющими несколько типов значений, например со всеми ячейками с текстовыми значениями и всеми ячейками с логическими значениями ("Logical"). Перечисление `Excel.SpecialCellValueType` содержит значения, позволяющие объединять типы. Например, "LogicalText" обрабатывает все ячейки с логическими и текстовыми значениями. Можно объединить любые два или три из четырех основных типов. Имена этих значений перечисления, объединяющих основные типы, всегда располагаются в алфавитном порядке. Поэтому для объединения ячеек со значениями ошибок, текстовыми и логическими значениями используйте параметр "ErrorLogicalText", а не "LogicalErrorText" или "TextErrorLogical". Параметр по умолчанию "All" объединяет все четыре типа. В приведенном ниже примере выделены все ячейки с формулами, которые производят числовые или логические значения:

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var formulaLogicalNumberRanges = usedRange.getSpecialCells("Formulas", "LogicalNumbers");
    formulaLogicalNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

> [!NOTE]
> Параметр `Excel.SpecialCellValueType` можно использовать, только если параметру `Excel.SpecialCellType` присвоено значение "Formulas" или "Constants".

### <a name="get-rangeareas-within-rangeareas"></a>Получение объектов RangeAreas в RangeAreas

У типа `RangeAreas` также есть методы `getSpecialCells` и `getSpecialCellsOrNullObject`, которые используют те же два параметра. Эти методы возвращают все целевые ячейки из всех диапазонов в коллекции `RangeAreas.areas`. Существует одно небольшое отличие в поведении методов при вызове объекта `RangeAreas` вместо объекта `Range`: если вы передаете "SameConditionalFormat" в качестве первого параметра, метод возвращает все ячейки с таким же условным форматированием, как у крайней левой верхней ячейки *первого диапазона в коллекции `RangeAreas.areas`*. То же касается и "SameDataValidation": при передаче к `Range.getSpecialCells` возвращаются все ячейки с таким же правилом проверки данных, как у крайней левой верхней ячейки *диапазона*.  Но при передаче к `RangeAreas.getSpecialCells` возвращаются все ячейки с таким же правилом проверки данных, как у крайней левой верхней ячейки *первого диапазона в коллекции `RangeAreas.areas`*.

## <a name="read-properties-of-rangeareas"></a>Чтение свойств RangeAreas

Чтение значений свойств `RangeAreas` требует внимания, так как определенное свойство может иметь разные значения для разных диапазонов в `RangeAreas`. Общее правило заключается в том, что если соответствующее значение *может* быть возвращено, оно будет возвращено. Например, в приведенном ниже коде RGB-код для розового цвета (`#FFC0CB`) и `true` будут записаны в консоль, так как оба диапазона в объекте `RangeAreas` имеют розовую заливку и оба являются целыми столбцами.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();

    // The ranges are the F column and the H column.
    var rangeAreas = sheet.getRanges("F:F, H:H");  
    rangeAreas.format.fill.color = "pink";

    rangeAreas.load("format/fill/color, isEntireColumn");

    return context.sync()
        .then(function () {
            console.log(rangeAreas.format.fill.color); // #FFC0CB
            console.log(rangeAreas.isEntireColumn); // true
        })
        .then(context.sync);
})
```

Все усложняется, если согласование невозможно. Свойства `RangeAreas` действуют в соответствии с приведенными ниже тремя принципами:

- Логическое свойство объекта `RangeAreas` возвращает значение `false`, кроме случаев, когда свойство имеет значение true для всех диапазонов элементов.
- Свойства, не являющиеся логическими, за исключением свойства `address`, возвращают значение `null`, кроме тех случаев, когда соответствующее свойство для всех диапазонов элементов обладает тем же значением.
- Свойство `address` возвращает строку с адресами диапазонов элементов, разделенными запятыми.

Например, в приведенном ниже коде создается объект `RangeAreas`, в котором только один диапазон является целым столбцом и только один залит розовым цветом. Консоль отобразит значение `null` для цвета заливки, `false` для свойства `isEntireRow` и "Sheet1!F3:F5, Sheet1!H:H" (при условии, что имя листа — "Sheet1") для свойства `address`. 

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var rangeAreas = sheet.getRanges("F3:F5, H:H");

    var pinkColumnRange = sheet.getRange("H:H");
    pinkColumnRange.format.fill.color = "pink";

    rangeAreas.load("format/fill/color, isEntireColumn, address");

    return context.sync()
        .then(function () {
            console.log(rangeAreas.format.fill.color); // null
            console.log(rangeAreas.isEntireColumn); // false
            console.log(rangeAreas.address); // "Sheet1!F3:F5, Sheet1!H:H"
        })
        .then(context.sync);
})
```

## <a name="see-also"></a>См. также

- [Основные концепции программирования с помощью API JavaScript для Excel](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
- [Объект Range (API JavaScript для Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range)
- [Объект RangeAreas (API JavaScript для Excel)](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas) (эта ссылка может не работать, пока API находится в предварительной версии. В качестве альтернативы см. [бета-версию office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts)).