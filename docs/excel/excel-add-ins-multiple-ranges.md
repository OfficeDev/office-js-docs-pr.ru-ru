---
title: Работа с несколькими диапазонами одновременно в надстройках Excel
description: ''
ms.date: 9/4/2018
ms.openlocfilehash: a00bbf15b53649147fb2c2b1dfa590f15c5739be
ms.sourcegitcommit: c53f05bbd4abdfe1ee2e42fdd4f82b318b363ad7
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/12/2018
ms.locfileid: "25506296"
---
# <a name="work-with-multiple-ranges-simultaneously-in-excel-add-ins-preview"></a>Работа с несколькими диапазонами одновременно в надстройках Excel (предварительная версия)

Библиотека JavaScript для Excel позволяет вашей надстройке выполнять операции и устанавливать свойства одновременно для нескольких диапазонов. Диапазоны необязательно должны быть непрерывными. Этот способ установки свойства не только упрощает код, но и выполняется намного быстрее, чем установка каждого отдельного свойства для каждого диапазона.

> [!NOTE]
> Для работы с API-интерфейсами, описанными в этой статье, требуется **версия Office 2016 Click-to-Run 1809 сборки 10820.20000** или более поздняя версия (возможно, вам придется принять участие в [программе предварительной оценки Office](https://products.office.com/office-insider) для получения нужной сборки). Кроме того, необходимо загрузить бета-версию библиотеки Office JavaScript из [сети CDN Office.js](https://appsforoffice.microsoft.com/lib/beta/hosted/office.js). Наконец, в настоящее время у нас еще нет страниц ссылки для этих API. Но следующий файл типа определения содержит их описания: [бета-версию office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).

## <a name="rangeareas"></a>RangeAreas

Набор диапазонов (возможно, разобщенных) представлен `Excel.RangeAreas`объектом . Он имеет свойства и методы, аналогичные `Range`типу  (многие из которых имеют одинаковые или похожие имена), но следующие параметры были изменены:

- Типы данных для свойств и поведений методов задания и методов получения.
- Типы данных параметров метода и поведений метода.
- Типы данных возвращаемых значений метода.

Некоторые примеры:

- `RangeAreas` имеет свойство `address`, которое возвращает строку с адресами диапазона, разделенными запятой, а не только один адрес, как в случае со свойством `Range.address`.
- `RangeAreas` имеет < свойство `dataValidation`, которое возвращает `DataValidation`объект  , представляющий собой проверку данных всех диапазонов в `RangeAreas` при соответствии. Этим свойством будет `null`, если идентичные `DataValidation`объекты  не применяются ко всем диапазонам в `RangeAreas`. Это общие, но не универсальные принципы для`RangeAreas` объекта : *если свойство не имеет согласованных значений для каждого из всех диапазонов в `RangeAreas`, то это свойство будет `null`.*  Дополнительные сведения и некоторые исключения см. в статье [Чтение свойств RangeAreas](#reading-properties-of-rangeareas).
- `RangeAreas.cellCount` возвращает общее число ячеек во всех диапазонах в `RangeAreas`.
- `RangeAreas.calculate` пересчитывает ячейки всех диапазонов в `RangeAreas`.
- `RangeAreas.getEntireColumn` и `RangeAreas.getEntireRow` возвращает другой объект `RangeAreas`, представляющий все столбцы (или строки) во всех диапазонах в `RangeAreas`. Например, если `RangeAreas` представляет «A1:C4» и «F14:L15», то `RangeAreas.getEntireColumn` возвращает объект `RangeAreas`, представляющий «A:C» и «F:L».
- `RangeAreas.copyFrom` может использовать параметр `Range` или `RangeAreas`, представляющий диапазон(ы) источника операции копирования.

#### <a name="complete-list-of-range-members-that-are-also-available-on-rangeareas"></a>Полный список элементов диапазона Range, которые также доступны на RangeAreas

##### <a name="properties"></a>Свойства

Ознакомьтесь со статьей [Чтение свойств RangeAreas](#reading-properties-of-rangeareas) до написания кода, который считывает все свойства из списка. Возвращаемое значение зависит от ряда факторов.

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

Помеченные методы диапазона в режиме предварительной версии.

- calculate()
- clear()
- convertDataTypeToText() (предварительная версия)
- convertToLinkedDataType() (предварительная версия)
- copyFrom() (предварительная версия)
- getEntireColumn()
- getEntireRow()
- getIntersection()
- getIntersectionOrNullObject()
- getOffsetRange() (с именем getOffsetRangeAreas на объекте RangeAreas)
- getSpecialCells() (предварительная версия)
- getSpecialCellsOrNullObject() (предварительная версия)
- getTables() (предварительная версия)
- getUsedRange() (с именем getUsedRangeAreas на объекте RangeAreas)
- getUsedRangeOrNullObject() (с именем getUsedRangeAreasOrNullObject на объекте RangeAreas)
- load()
- set()
- setDirty() (предварительная версия)
- toJSON()
- track()
- untrack()

### <a name="rangearea-specific-properties-and-methods"></a>Свойства и методы, характерные для объекта RangeArea

Тип `RangeAreas` имеет некоторые свойства и методы, которые не входят в объект `Range`. Ниже приведены некоторые из них:

- `areas`: объект `RangeCollection`, содержащий все диапазоны, представленные объектом `RangeAreas`. Объект `RangeCollection`  — также новый и аналогичен другим объектам коллекции Excel. Он имеет свойство `items`, которое представляет собой массив объектов `Range`, представляющих диапазоны.
- `areaCount`: общее число диапазонов в `RangeAreas`.
- `getOffsetRangeAreas`: работает так же, как [Range.getOffsetRange](https://docs.microsoft.com/javascript/api/excel/excel.range#getoffsetrange-rowoffset--columnoffset-), за исключением того, что `RangeAreas` возвращается и содержит диапазоны, каждый из которых смещен от одного из диапазонов в исходном `RangeAreas`.

## <a name="create-rangeareas-and-set-properties"></a>Создание RangeAreas и установка свойств

Вы можете создать объект `RangeAreas` двумя основными способами:

- Вызвать `Worksheet.getRanges()`  и передать его в строку с адресами диапазона, разделенными запятыми. Если диапазон, который вы хотите включить, был переделан в [NamedItem](https://docs.microsoft.com/javascript/api/excel/excel.nameditem), вы можете включить в строку имя вместо адреса.
- Вызвать `Workbook.getSelectedRanges()`. Этот метод возвращает `RangeAreas`, представляющий все диапазоны, выбранные в активном на данный момент листе.

После получения объекта `RangeAreas` можно создать другие с помощью методов, применяемых к объекту, который возвращает `RangeAreas`, такие как `getOffsetRangeAreas` и `getIntersection`.

> [!NOTE]
> Невозможно напрямую добавить дополнительные диапазоны к объекту `RangeAreas`. Например, коллекция в `RangeAreas.areas` не имеет метода `add`.


> [!WARNING] 
> Не пытайтесь напрямую добавлять или удалять элементы из массива `RangeAreas.areas.items`. Это приведет к нежелательному функционированию кода. Например, существует возможность принудительно добавить дополнительный объект `Range` в массив, но это приведет к ошибкам, так как свойства `RangeAreas` и методы функционируют так, как если бы новый элемент не был добавлен. Например, свойство `areaCount` не включает диапазоны, принудительно добавленные таким образом, а `RangeAreas.getItemAt(index)`  вызывает ошибку, если `index` больше, чем `areasCount-1`. Аналогичным образом, удаление объекта `Range`  в массиве `RangeAreas.areas.items` путем получения ссылки на него и вызов его метода `Range.delete`  вызывает ошибки: хотя объект `Range` * будет* удален, свойства и методы родительского объекта `RangeAreas`  будут функционировать (или пытаться функционировать) так, как если бы он все еще присутствовал. Например, если код вызывает метод `RangeAreas.calculate`, Office попытается рассчитать диапазон, но это завершится ошибкой, поскольку объект range отсутствует.

Установка свойства для `RangeAreas` задает соответствующее свойство для всех диапазонов в коллекции `RangeAreas.areas`.

Ниже приведен пример установки свойства в нескольких диапазонах. Функция выделяет диапазоны **F3:F5** и **H3:H5**.

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const rangeAreas = sheet.getRanges("F3:F5, H3:H5");
    rangeAreas.format.fill.color = "pink";

    return context.sync();
})
```

Этот пример применяется к сценариям, в которых можно создать серьезный код адресов диапазона, передаваемых в `getRanges`, или легко рассчитать их во время выполнения. Ниже перечислены некоторые сценарии, в которых это возможно: 

- Код выполняется в контексте известного шаблона.
- Код выполняется в контексте импортированных данных, в котором известна схема данных.

Когда во время создания кода не известно, с какими диапазонами вам придется работать, необходимо обнаружить их во время выполнения. В следующем разделе описываются эти сценарии.

### <a name="discover-range-areas-programmatically"></a>Обнаружение областей диапазона с помощью программных средств

Методы `Range.getSpecialCells()` и `Range.getSpecialCellsOrNullObject()` можно использовать для поиска во время выполнения диапазонов, с которыми вы хотите работать, на основе характеристик ячеек и типа значений в ячейках. Вот подписи методов из файла типов данных TypeScript:

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

Ниже приведен пример использования первого из них. Вот что нужно знать об этом коде:

- Он ограничивает часть листа, которую нужно искать, вызвав сначала `Worksheet.getUsedRange`, а затем вызвав `getSpecialCells` только для этого диапазона.
- Передает в качестве параметра `getSpecialCells` строчную версию значения из перечисления `Excel.SpecialCellType`. Некоторые другие значения, которые могут быть переданы вместо этого, — это "Blanks" для пустых ячеек, "Constants" для ячейки со значениями литералов вместо формул и "SameConditionalFormat" для ячеек, у которых такое же условное форматирование, как и у первой ячейки в `usedRange`. Первая ячейка — это самая верхняя ячейка слева. Полный список значений перечисления см. в [бета-версии office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).
- Метод `getSpecialCells` возвращает объект `RangeAreas`, поэтому все ячейки с формулами залиты розовым цветом даже в том случае, если они не последовательны. 

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const formulaRanges = usedRange.getSpecialCells("Formulas");
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

В некоторых случаях диапазон не содержит *ни одной* ячейки с целевой характеристикой. Если `getSpecialCells` не находит ни одной такой ячеки, он выдает ошибку **ItemNotFound** . Это приведет к переадресации потока управления к блоку/методу `catch`, если таковой существует. Если нет, ошибка приведет к прекращению исполнения функции. Могут существовать сценарии, в которых выдача ошибки – это именно то, что должно происходить при отсутствии ячейки с целевой характеристикой. 

Но в некоторых сценариях отсутствие соответствующих ячеек нормально, хотя и, возможно, необычно; ваш код должен проверить наличие такой возможности и аккуратно провести работу со сценарием без выдачи ошибки. Для этих сценариев следует использовать метод `getSpecialCellsOrNullObject`  и протестировать свойство `RangeAreas.isNullObject`. Пример см. ниже. Что нужно знать об этом коде:

- `getSpecialCellsOrNullObject`Метод  всегда возвращает объект прокси-сервера, поэтому он не может быть `null` в обычном смысле JavaScript. Но если соответствующие ячейки не обнаружены, `isNullObject`свойству  объекта присваивается значение `true`.
- Он вызывает `context.sync` *перед* тестированием `isNullObject`свойства . Это требование для всех `*OrNullObject`методов и свойств, так как всегда нужно загружать и синхронизировать свойство, чтобы его прочесть. Тем не менее, необязательно *прямо* загружать`isNullObject` свойство. Оно автоматически загружается `context.sync` даже в том случае, если `load` не вызывается для объекта. Дополнительные сведения см. в разделе [\*OrNullObject](https://docs.microsoft.com/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#42ornullobject-methods).
- Этот код можно проверить, выбрав сначала диапазон, у которого нет ячеек формулы, и запустив его. Затем следует выбрать диапазон, содержащий по крайней мере одну ячейку с формулой, и снова запустить его.

```js
Excel.run(function (context) {
    const range = context.workbook.getSelectedRange();
    const formulaRanges = range.getSpecialCellsOrNullObject("Formulas");
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

Для простоты во всех других примерах в этой статье используйте метод `getSpecialCells` вместо `getSpecialCellsOrNullObject`.

#### <a name="narrow-the-target-cells-with-cell-value-types"></a>Сужение целевых ячеек с типом значений ячеек

Также существует необязательный второй параметр типа перечисления `Excel.SpecialCellValueType`, который в дальнейшем сужает ячейки до целевого объекта.  Его можно использовать только в том случае, если передается значение "Formulas" или "Constants" для `getSpecialCells` или `getSpecialCellsOrNullObject`. Этот параметр указывает, что требуются только ячейки с определенными типами значений. Существует четыре основных типа: "Error", "Logical" (то же самое, что и boolean — логический), "Numbers" и "Text" (перечисление имеет другие значения помимо этих четырех, которые рассматриваются ниже). См. пример ниже. Что нужно знать об этом коде:

- Он выделяет только ячейки, имеющие числовое значение литерала и не выделяет ячейки, в которых содержится формула (даже в том случае, если результат является числом), логическое, текстовое значение, или ячейки с состоянием ошибки.
- Чтобы протестировать код, убедитесь, что в листе есть ячейки с числовыми значениями литералов, ячейки с другими значениями литералов и ячейки с формулами.

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const constantNumberRanges = usedRange.getSpecialCells("Constants", "Numbers");
    constantNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

В некоторых случаях вам нужно работать с ячейками, имеющими более одного типа значения, например, со всеми ячейками с текстовыми значениями и всеми ячейками с логическими значениями ("Logical"). Перечисление `Excel.SpecialCellValueType` содержит значения, которые позволяют объединять различные типы. Например, "LogicalText" обрабатывает все логические и все текстовые ячейки. Вы можете использовать любые два или три из четырех основных типов. Имена этих значений перечисления, которые объединяют основные типы, всегда располагаются в алфавитном порядке. Поэтому для объединения ячеек со значениями ошибок, текстовыми и логическими значениями используйте "ErrorLogicalText", а не "LogicalErrorText" или "TextErrorLogical". Параметр по умолчанию "All" объединяет все четыре типа. В следующем примере выделены все ячейки с формулами, которые производят числовые или логические значения:

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    const formulaLogicalNumberRanges = usedRange.getSpecialCells("Formulas", "LogicalNumbers");
    formulaLogicalNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

> [!NOTE]
> Параметр `Excel.SpecialCellValueType` можно использовать, только если параметр `Excel.SpecialCellType` — это "Formulas" или "Constants".

### <a name="get-rangeareas-within-rangeareas"></a>Получение объектов RangeAreas в RangeAreas

Тип  `RangeAreas` также имеет методы  `getSpecialCells` и `getSpecialCellsOrNullObject`, которые принимают те же два параметра. Эти методы возвращают все целевые ячейки из всех диапазонов в коллекции `RangeAreas.areas`. Существует одно небольшое отличие в поведении методов при вызове объекта `RangeAreas`вместо объекта `Range`: когда вы передаете "SameConditionalFormat" в качестве первого параметра, метод возвращает все ячейки, имеющие одинаковое условное форматирование, как верхнюю крайнюю слева ячейку *в первом диапазоне в коллекции `getSpecialCellsOrNullObject`*. То же касается и "SameDataValidation": при передаче к `Range.getSpecialCells`он возвращает все ячейки, которые имеют такое же правило проверки данных, как верхнюю крайнюю слева ячейку *в диапазоне*. Но при передаче к `RangeAreas.getSpecialCells` он возвращает все ячейки, которые имеют такое же правило проверки данных, как верхнюю крайнюю слева ячейку * в первом диапазоне в коллекции`RangeAreas.areas`*.

## <a name="read-properties-of-rangeareas"></a>Чтение свойств RangeAreas

Чтение значения свойств `RangeAreas` требует внимания, так как то или иное свойство может иметь разные значения для разных диапазонов в `RangeAreas`. Общее правило заключается в том, что если соответствующее значение *может*  быть возвращено, оно будет возвращено. Например, в следующем коде RGB-код для розовой заливки (`#FFC0CB`) и `true`  будет выполнять вход в консоль, так как оба диапазона в объекте `RangeAreas`  имеют розовую заливку и оба являются целыми столбцами.

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    // The ranges are the F column and the H column.
    const rangeAreas = sheet.getRanges("F:F, H:H");  
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

Все усложняется, когда согласование невозможно. Свойство `RangeAreas` работает в соответствии со следующими тремя принципами:

- Логическое свойство объекта `RangeAreas` возвращает `false`, кроме случаев, когда свойство имеет значение true для всех диапазонов элементов.
- Свойства, не являющиеся логическими, за исключением свойства `address`, возвращают `null`, кроме тех случаев, когда соответствующее свойство для всех элементов диапазона обладает тем же значением.
- Свойство `address` возвращает строку с адресами диапазонов элементов, разделенными запятыми.

Например, следующий код создает `RangeAreas`, в котором только один диапазон — это целый столбец, и только один заполнен розовым цветом. Консоль покажет `null`для цвета заливки, `false`для `isEntireRow`свойства  и "Sheet1! F3:F5 Sheet1! H:H" (при условии, что имя листа – это "Sheet1") для `address`свойства . 

```js
Excel.run(function (context) {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const rangeAreas = sheet.getRanges("F3:F5, H:H");

    const pinkColumnRange = sheet.getRange("H:H");
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

- [Основные принципы программирования с помощью API JavaScript для Excel](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
- [Объект Range (API JavaScript для Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range)
- Объект[RangeAreas (JavaScript API для Excel)](https://docs.microsoft.com/javascript/api/excel/excel.rangeareas) (эта ссылка может не работать, пока API находится в режиме предварительной версии. В качестве альтернативы см. [бета-версию office.d.ts](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts).)