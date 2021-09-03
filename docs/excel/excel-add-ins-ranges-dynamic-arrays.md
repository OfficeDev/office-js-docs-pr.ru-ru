---
title: Обработка динамических массивов и разлива диапазона с Excel API JavaScript
description: Узнайте, как обрабатывать динамические массивы и разливать диапазоны с помощью Excel API JavaScript.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: d95546b4cff3f0ba7410d9ceaa73e19b7e684985
ms.sourcegitcommit: 69f6492de8a4c91e734250c76681c44b3f349440
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/03/2021
ms.locfileid: "58868689"
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

## <a name="see-also"></a>См. также

- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
- [Работа с ячейками с Excel API JavaScript](excel-add-ins-cells.md)
- [Работа с несколькими диапазонами одновременно в надстройках Excel](excel-add-ins-multiple-ranges.md)
