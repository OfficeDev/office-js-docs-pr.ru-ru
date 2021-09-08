---
title: Поиск специальных ячеек в диапазоне с Excel API JavaScript
description: Узнайте, как использовать API Excel JavaScript для поиска специальных ячеек, таких как ячейки с формулами, ошибками или числами.
ms.date: 07/08/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: f1562351b045b5c8df1edb3c22f651883a836ad9
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937767"
---
# <a name="find-special-cells-within-a-range-using-the-excel-javascript-api"></a>Поиск специальных ячеек в диапазоне с Excel API JavaScript

В этой статье данная статья содержит примеры кода, которые находят специальные ячейки в диапазоне с Excel API JavaScript. Полный список свойств и методов, поддерживаемый объектом, см. в `Range` [Excel. Класс Range](/javascript/api/excel/excel.range).

## <a name="find-ranges-with-special-cells"></a>Поиск диапазонов с помощью специальных ячеек

Методы [Range.getSpecialCells](/javascript/api/excel/excel.range#getSpecialCells_cellType__cellValueType_) и [Range.getSpecialCellsOrNullObject](/javascript/api/excel/excel.range#getSpecialCellsOrNullObject_cellType__cellValueType_) находят диапазоны, основанные на характеристиках их клеток и типах значений их клеток. Оба этих метода возвращают объекты `RangeAreas`. Подписи методов из файла типов данных TypeScript:

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

В следующем примере кода используется `getSpecialCells` метод для поиска всех ячеек с формулами. Вот что нужно знать об этом коде:

- Он ограничивает часть листа, в которой требуется выполнять поиск, путем вызова сначала метода `Worksheet.getUsedRange`, а затем метода `getSpecialCells` только для этого диапазона.
- Метод `getSpecialCells` возвращает объект `RangeAreas`, поэтому все ячейки с формулами окрашены розовым цветом даже в том случае, если они не являются смежными.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var formulaRanges = usedRange.getSpecialCells(Excel.SpecialCellType.formulas);
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

Если в диапазоне нет ячеек с целевыми характеристиками, метод `getSpecialCells` выдает ошибку **ItemNotFound**. Это приведет к переадресации потока управления к блоку `catch`, если таковой существует. Если блокировки `catch` нет, ошибка останавливает метод.

Если ожидается, что всегда должны существовать ячейки с целевыми характеристиками, скорее всего вы захотите, чтобы код выдавал ошибку при их отсутствии. Если отсутствие соответствующих ячеек является допустимым сценарием, ваш код должен проверить наличие такой возможности и корректно выполнить действие без выдачи ошибки. Добиться такого поведения можно с помощью метода `getSpecialCellsOrNullObject` и возвращаемого им свойства `isNullObject`. В следующем примере кода используется этот шаблон. Вот что нужно знать об этом коде:

- Метод всегда возвращает прокси-объект, поэтому он никогда не находится в `getSpecialCellsOrNullObject` `null` обычном смысле JavaScript. Но если соответствующие ячейки не обнаружены, свойству `isNullObject` объекта присваивается значение `true`.
- Он вызывает `context.sync` *перед* тестированием свойства `isNullObject`. Это требование для всех методов и свойств `*OrNullObject`, так как всегда нужно загружать и синхронизировать свойство, чтобы его прочесть. Однако не нужно явно *загружать* `isNullObject` свойство. Он автоматически загружается объектом, даже если он не `context.sync` `load` вызван. Дополнительные сведения см. в дополнительных сведениях о методах [ \* и свойствах OrNullObject.](../develop/application-specific-api-model.md#ornullobject-methods-and-properties)
- Этот код можно проверить, выбрав сначала диапазон без ячеек с формулами и запустив его. Затем следует выбрать диапазон, содержащий по крайней мере одну ячейку с формулой, и снова запустить его.

```js
Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    var formulaRanges = range.getSpecialCellsOrNullObject(Excel.SpecialCellType.formulas);
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

Для простоты все остальные образцы кода в этой статье используют `getSpecialCells` метод вместо  `getSpecialCellsOrNullObject` .

## <a name="narrow-the-target-cells-with-cell-value-types"></a>Ограничение целевых ячеек с помощью типа значений ячеек

Методы `Range.getSpecialCells()` и `Range.getSpecialCellsOrNullObject()` принимают необязательный второй параметр, используемый для дополнительного ограничения целевых ячеек. Этот второй параметр `Excel.SpecialCellValueType` используется для указания того, что требуются только ячейки, содержащие определенные типы значений.

> [!NOTE]
> Параметр `Excel.SpecialCellValueType` можно использовать, только если для параметра `Excel.SpecialCellType` задано значение `Excel.SpecialCellType.formulas` или `Excel.SpecialCellType.constants`.

### <a name="test-for-a-single-cell-value-type"></a>Тестирование для ячеек с одним типом значений

Для перечисления `Excel.SpecialCellValueType` существует четыре основных типа (в дополнение к другим объединенным значениям, описанным ниже в этом разделе):

- `Excel.SpecialCellValueType.errors`
- `Excel.SpecialCellValueType.logical` (означает логическое значение)
- `Excel.SpecialCellValueType.numbers`
- `Excel.SpecialCellValueType.text`

В следующем примере кода находятся специальные ячейки, которые являются числовыми константами, и цвета этих клеток розовыми. Вот что нужно знать об этом коде:

- Он выделяет только ячейки, которые имеют буквальное значение числа. В нем не будут выделены ячейки, у них есть формула (даже если результат — число) или клеток состояния boolean, text или error.
- Чтобы протестировать код, убедитесь, что в листе есть ячейки с числовыми значениями литералов, ячейки с другими значениями литералов и ячейки с формулами.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var constantNumberRanges = usedRange.getSpecialCells(
        Excel.SpecialCellType.constants,
        Excel.SpecialCellValueType.numbers);
    constantNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

### <a name="test-for-multiple-cell-value-types"></a>Тестирование для ячеек с несколькими типами значений

Иногда требуется работать с ячейками, имеющими несколько типов значений, например со всеми ячейками с текстовыми значениями и всеми ячейками с логическими значениями (`Excel.SpecialCellValueType.logical`). Для перечисления `Excel.SpecialCellValueType` существуют значения с объединенными типами. Например, `Excel.SpecialCellValueType.logicalText` обрабатывает все ячейки с логическими и текстовыми значениями. `Excel.SpecialCellValueType.all` является значением по умолчанию, которое не ограничивает возвращаемые типы значений ячеек. В следующем примере кода цвета всех ячеек с формулами, которые производят количество или значение boolean.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var formulaLogicalNumberRanges = usedRange.getSpecialCells(
        Excel.SpecialCellType.formulas,
        Excel.SpecialCellValueType.logicalNumbers);
    formulaLogicalNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

## <a name="see-also"></a>См. также

- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
- [Работа с ячейками с Excel API JavaScript](excel-add-ins-cells.md)
- [Поиск строки с Excel API JavaScript](excel-add-ins-ranges-string-match.md)
- [Работа с несколькими диапазонами одновременно в надстройках Excel](excel-add-ins-multiple-ranges.md)
