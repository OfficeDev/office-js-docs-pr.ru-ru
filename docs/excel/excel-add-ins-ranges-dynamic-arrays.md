---
title: Обработка динамических массивов и разлива диапазона с Excel API JavaScript
description: Узнайте, как обрабатывать динамические массивы и разливать диапазоны с помощью Excel API JavaScript.
ms.date: 04/02/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 4601cd6a901243cadab0e7c5ead6061e28806377
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151485"
---
# <a name="handle-dynamic-arrays-and-spilling-using-the-excel-javascript-api"></a>Обработка динамических массивов и разлив с помощью Excel API JavaScript

В этой статье приводится пример кода, который обрабатывает динамические массивы и разлив диапазонов с Excel API JavaScript. Полный список свойств и методов, поддерживаемый объектом, см. в `Range` [Excel. Класс Range](/javascript/api/excel/excel.range).

## <a name="dynamic-arrays"></a>Динамические массивы

Некоторые Excel возвращают [динамические массивы.](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531) Они заполняют значения нескольких ячеек за пределами исходной ячейки формулы. Это переполнение значения называется "разлив". Надстройка может найти диапазон, используемый для разлива с помощью метода [Range.getSpillingToRange.](/javascript/api/excel/excel.range#getSpillingToRange__) Существует также [версия *OrNullObject](../develop/application-specific-api-model.md#ornullobject-methods-and-properties), `Range.getSpillingToRangeOrNullObject` .

В следующем примере показана базовая формула, которая копирует содержимое диапазона в ячейку, которая разливается в соседние ячейки. Затем надстройка регистрит диапазон, содержащий разлив.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    // Set G4 to a formula that returns a dynamic array.
    var targetCell = sheet.getRange("G4");
    targetCell.formulas = [["=A4:D4"]];

    // Get the address of the cells that the dynamic array spilled into.
    var spillRange = targetCell.getSpillingToRange();
    spillRange.load("address");

    // Sync and log the spilled-to range.
    return context.sync().then(function () {
        // This will log the range as "G4:J4".
        console.log(`Copying the table headers spilled into ${spillRange.address}.`);
    });
}).catch(errorHandlerFunction);
```

## <a name="range-spilling"></a>Разлиение диапазона

Найдите ячейку, ответственную за разлив в заданную ячейку с помощью метода [Range.getSpillParent.](/javascript/api/excel/excel.range#getSpillParent__) Обратите `getSpillParent` внимание, что работает только в том случае, если объект диапазона является одной ячейкой. Вызов диапазона с несколькими ячейками приведет к ошибке, которая будет выброшена (или возвращается диапазон `getSpillParent` `Range.getSpillParentOrNullObject` null).

## <a name="see-also"></a>Дополнительные материалы

- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
- [Работа с ячейками с Excel API JavaScript](excel-add-ins-cells.md)
- [Работа с несколькими диапазонами одновременно в надстройках Excel](excel-add-ins-multiple-ranges.md)
