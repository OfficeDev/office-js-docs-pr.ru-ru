---
title: Работа с несколькими диапазонами одновременно в надстройках Excel
description: Узнайте, как Excel JavaScript позволяет вашей надстройки выполнять операции и устанавливать свойства одновременно на нескольких диапазонах.
ms.date: 04/01/2021
localization_priority: Normal
ms.openlocfilehash: 4f1661d07432d6072649cb6db7315fd39fee5b4f
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937978"
---
# <a name="work-with-multiple-ranges-simultaneously-in-excel-add-ins"></a>Работа с несколькими диапазонами одновременно в надстройках Excel

Библиотека JavaScript для Excel позволяет вашей надстройке выполнять операции и устанавливать свойства одновременно для нескольких диапазонов. Диапазоны необязательно должны быть смежными. Этот способ установки свойства не только упрощает код, но и выполняется намного быстрее, чем установка одинакового свойства отдельно для каждого диапазона.

## <a name="rangeareas"></a>RangeAreas

Набор (возможно, дисконтных) диапазонов представлен [объектом RangeAreas.](/javascript/api/excel/excel.rangeareas) Его свойства и методы аналогичны типу `Range` (многие с одинаковыми или похожими именами), но с изменением указанных ниже параметров:

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

- `address`
- `addressLocal`
- `cellCount`
- `conditionalFormats`
- `context`
- `dataValidation`
- `format`
- `isEntireColumn`
- `isEntireRow`
- `style`
- `worksheet`

##### <a name="methods"></a>Методы

- `calculate()`
- `clear()`
- `convertDataTypeToText()`
- `convertToLinkedDataType()`
- `copyFrom()`
- `getEntireColumn()`
- `getEntireRow()`
- `getIntersection()`
- `getIntersectionOrNullObject()`
- `getOffsetRange()``getOffsetRangeAreas`(названо на `RangeAreas` объекте)
- `getSpecialCells()`
- `getSpecialCellsOrNullObject()`
- `getTables()`
- `getUsedRange()``getUsedRangeAreas`(названо на `RangeAreas` объекте)
- `getUsedRangeOrNullObject()``getUsedRangeAreasOrNullObject`(названо на `RangeAreas` объекте)
- `load()`
- `set()`
- `setDirty()`
- `toJSON()`
- `track()`
- `untrack()`

### <a name="rangearea-specific-properties-and-methods"></a>Свойства и методы, характерные для объекта RangeArea

Для типа `RangeAreas` существуют несколько свойств и методов, отсутствующих в объекте `Range`. Ниже приводится их выбор.

- `areas`. Объект `RangeCollection`, содержащий все диапазоны, которые представлены объектом `RangeAreas`. Объект `RangeCollection` — еще один новый объект, аналогичный другим объектам коллекции Excel. У него есть свойство `items`, являющееся массивом объектов `Range`, которые представляют диапазоны.
- `areaCount`. Общее количество диапазонов в `RangeAreas`.
- `getOffsetRangeAreas`. Действует аналогично методу [Range.getOffsetRange](/javascript/api/excel/excel.range#getOffsetRange_rowOffset__columnOffset_), за исключением того, что возвращается объект `RangeAreas`, содержащий диапазоны, каждый из которых смещен относительно одного из диапазонов в исходном объекте `RangeAreas`.

## <a name="create-rangeareas"></a>Создание RangeAreas

Объект `RangeAreas` можно создать двумя основными способами:

- Вызвать метод `Worksheet.getRanges()` и передать ему строку с адресами диапазона, разделенными запятыми. Если диапазон, который нужно включить, был преобразован [NamedItem](/javascript/api/excel/excel.nameditem), в строку можно включить имя вместо адреса.
- Вызвать метод `Workbook.getSelectedRanges()`. Этот метод возвращает объект `RangeAreas`, представляющий все диапазоны, выбранные в активном на данный момент листе.

После получения объекта `RangeAreas` можно создать другие с помощью методов объекта, возвращающих `RangeAreas`, например `getOffsetRangeAreas` и `getIntersection`.

> [!NOTE]
> Нельзя напрямую добавить дополнительные диапазоны к объекту `RangeAreas`. Например, у коллекции в `RangeAreas.areas` нет метода `add`.

> [!WARNING]
> Не пытайтесь напрямую добавлять или удалять элементы из массива `RangeAreas.areas.items`. Это приведет к нежелательному поведению кода.  Например, существует возможность принудительно добавить дополнительный объект `Range` в массив, но это приведет к ошибкам, поскольку свойства и методы `RangeAreas` действуют, как будто новый элемент не был добавлен. Например, свойство `areaCount` не включает диапазоны, принудительно добавленные таким образом, а `RangeAreas.getItemAt(index)` вызывает ошибку, если `index` больше, чем `areasCount-1`.  Аналогичным образом, удаление объекта `Range` в массиве `RangeAreas.areas.items` путем получения ссылки на него и вызова его метода `Range.delete` приводит к ошибкам: хотя объект `Range` *удален*, свойства и методы родительского объекта `RangeAreas` будут действовать (или пытаться действовать), как будто он еще существует. Например, если код вызывает метод `RangeAreas.calculate`, Office попытается рассчитать диапазон, но это завершится ошибкой, поскольку объект range отсутствует.

## <a name="set-properties-on-multiple-ranges"></a>Задание свойств для нескольких диапазонов

Установка свойства для объекта `RangeAreas` задает соответствующее свойство для всех диапазонов в коллекции `RangeAreas.areas`.

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

## <a name="get-special-cells-from-multiple-ranges"></a>Получение специальных ячеек из нескольких диапазонов

Методы `getSpecialCells` и `getSpecialCellsOrNullObject` для объекта `RangeAreas` действуют аналогично методам с теми же названиями для объекта `Range`. Эти методы возвращают ячейки с указанными характеристиками из всех диапазонов в коллекции `RangeAreas.areas`. Дополнительные сведения о специальных ячейках см. в материале [Find special cells within a range.](excel-add-ins-ranges-special-cells.md)

При вызове метода `getSpecialCells` или `getSpecialCellsOrNullObject` для объекта `RangeAreas`:

- Если в качестве первого параметра передается `Excel.SpecialCellType.sameConditionalFormat`, метод возвращает все ячейки с таким же условным форматированием, как у крайней левой верхней ячейки первого диапазона в коллекции `RangeAreas.areas`.
- Если в качестве первого параметра передается `Excel.SpecialCellType.sameDataValidation`, метод возвращает все ячейки с таким же правилом проверки данных, как у крайней левой верхней ячейки первого диапазона в коллекции `RangeAreas.areas`.

## <a name="read-properties-of-rangeareas"></a>Чтение свойств RangeAreas

Чтение значений свойств `RangeAreas` требует внимания, так как определенное свойство может иметь разные значения для разных диапазонов в `RangeAreas`. Общее правило заключается в том, что если соответствующее значение *может* быть возвращено, оно будет возвращено. Например, в следующем коде код RGB для розового () и будет вошел в консоль, так как оба диапазона в объекте имеют розовую заливка и оба являются целые `#FFC0CB` `true` `RangeAreas` столбцы.

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

- [Основные концепции программирования с помощью API JavaScript для Excel](../reference/overview/excel-add-ins-reference-overview.md)
- [Чтение или написание в большом диапазоне с Excel API JavaScript](excel-add-ins-ranges-large.md)
