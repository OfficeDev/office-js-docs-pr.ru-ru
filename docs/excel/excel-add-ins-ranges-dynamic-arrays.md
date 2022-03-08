---
title: Обработка динамических массивов и разлива диапазона с Excel API JavaScript
description: Узнайте, как обрабатывать динамические массивы и разливать диапазоны с помощью Excel API JavaScript.
ms.date: 02/17/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: b096225a7f4582f15b5707dcd0059e8e8869ad8d
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340696"
---
# <a name="handle-dynamic-arrays-and-spilling-using-the-excel-javascript-api"></a>Обработка динамических массивов и разлив с помощью Excel API JavaScript

В этой статье приводится пример кода, который обрабатывает динамические массивы и разлив диапазонов с Excel API JavaScript. Полный список свойств `Range` и методов, поддерживаемый объектом, см. в Excel[. Класс Range](/javascript/api/excel/excel.range).

## <a name="dynamic-arrays"></a>Динамические массивы

Некоторые Excel возвращают [динамические массивы](https://support.microsoft.com/office/205c6b06-03ba-4151-89a1-87a7eb36e531). Они заполняют значения нескольких ячеек за пределами исходной ячейки формулы. Это переполнение значения называется "разлив". Надстройка может найти диапазон, используемый для разлива с помощью метода [Range.getSpillingToRange](/javascript/api/excel/excel.range#excel-excel-range-getspillingtorange-member(1)) . Существует также [версия *OrNullObject](../develop/application-specific-api-model.md#ornullobject-methods-and-properties), `Range.getSpillingToRangeOrNullObject`.

В следующем примере показана базовая формула, которая копирует содержимое диапазона в ячейку, которая разливается в соседние ячейки. Затем надстройка регистрит диапазон, содержащий разлив.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    // Set G4 to a formula that returns a dynamic array.
    let targetCell = sheet.getRange("G4");
    targetCell.formulas = [["=A4:D4"]];

    // Get the address of the cells that the dynamic array spilled into.
    let spillRange = targetCell.getSpillingToRange();
    spillRange.load("address");

    // Sync and log the spilled-to range.
    await context.sync();

    // This will log the range as "G4:J4".
    console.log(`Copying the table headers spilled into ${spillRange.address}.`);
});
```

## <a name="range-spilling"></a>Разлиение диапазона

Найдите ячейку, ответственную за разлив в заданную ячейку с помощью метода [Range.getSpillParent](/javascript/api/excel/excel.range#excel-excel-range-getspillparent-member(1)) . Обратите внимание, `getSpillParent` что работает только в том случае, если объект диапазона является одной ячейкой. Вызов `getSpillParent` диапазона с несколькими ячейками приведет к ошибке, которая будет выброшена (или возвращается диапазон null).`Range.getSpillParentOrNullObject`

## <a name="see-also"></a>См. также

- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
- [Работа с ячейками с Excel API JavaScript](excel-add-ins-cells.md)
- [Работа с несколькими диапазонами одновременно в надстройках Excel](excel-add-ins-multiple-ranges.md)
