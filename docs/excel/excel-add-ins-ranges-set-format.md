---
title: Установите формат диапазона с помощью API Excel JavaScript
description: Узнайте, как использовать Excel API JavaScript для набора формата диапазона.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: a09d3b4d79584e186c0be37d4a30954c4d4d0086
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937775"
---
# <a name="set-range-format-using-the-excel-javascript-api"></a>Настройка формата диапазона с Excel API JavaScript

В этой статье данная статья содержит примеры кода, которые устанавливают цвет шрифта, заполняют цвет и формат номеров для ячеек в диапазоне с Excel API JavaScript. Полный список свойств и методов, поддерживаемый объектом, см. в `Range` [Excel. Класс Range](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-font-color-and-fill-color"></a>Задание цвета шрифта и цвета заливки

В примере ниже показано, как задать цвет шрифта и цвет заливки для ячеек в диапазоне **B2: E2**.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var range = sheet.getRange("B2:E2");
    range.format.fill.color = "#4472C4";
    range.format.font.color = "white";

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-in-range-before-font-color-and-fill-color-are-set"></a>Данные в диапазоне перед заданием цвета шрифта и цвета заливки

![Данные в Excel перед набором формата.](../images/excel-ranges-format-before.png)

### <a name="data-in-range-after-font-color-and-fill-color-are-set"></a>Данные в диапазоне после задания цвета шрифта и цвета заливки

![Данные в Excel после набора формата.](../images/excel-ranges-format-font-and-fill.png)

## <a name="set-number-format"></a>Задание формата чисел

В примере ниже показано, как задать формат чисел для ячеек в диапазоне **D3:E5**.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var formats = [
        ["0.00", "0.00"],
        ["0.00", "0.00"],
        ["0.00", "0.00"]
    ];

    var range = sheet.getRange("D3:E5");
    range.numberFormat = formats;

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-in-range-before-number-format-is-set"></a>Данные в диапазоне перед заданием формата чисел

![Данные в Excel перед набором формата номеров.](../images/excel-ranges-format-font-and-fill.png)

### <a name="data-in-range-after-number-format-is-set"></a>Данные в диапазоне после задания формата чисел

![Данные в Excel после набора формата номеров.](../images/excel-ranges-format-numbers.png)

## <a name="see-also"></a>См. также

- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
- [Работа с ячейками с Excel API JavaScript](excel-add-ins-cells.md)
- [Настройка и получения диапазонов с Excel API JavaScript](excel-add-ins-ranges-set-get.md)
- [Установите и получите значения диапазона, текст или формулы с Excel API JavaScript](excel-add-ins-ranges-set-get-values.md)
