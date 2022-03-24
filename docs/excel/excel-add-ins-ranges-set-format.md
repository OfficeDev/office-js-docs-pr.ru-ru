---
title: Установите формат диапазона с помощью API Excel JavaScript
description: Узнайте, как использовать Excel API JavaScript для набора формата диапазона.
ms.date: 02/17/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: b3cd37aa4415627c2d549b1d11bdb6276388868b
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745352"
---
# <a name="set-range-format-using-the-excel-javascript-api"></a>Настройка формата диапазона с Excel API JavaScript

В этой статье данная статья содержит примеры кода, которые устанавливают цвет шрифта, цвет и формат номеров для ячеек в диапазоне с Excel API JavaScript. Полный список свойств `Range` и методов, поддерживаемый объектом, см. в Excel[. Класс Range](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-font-color-and-fill-color"></a>Задание цвета шрифта и цвета заливки

В примере ниже показано, как задать цвет шрифта и цвет заливки для ячеек в диапазоне **B2: E2**.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let range = sheet.getRange("B2:E2");
    range.format.fill.color = "#4472C4";
    range.format.font.color = "white";

    await context.sync();
});
```

### <a name="data-in-range-before-font-color-and-fill-color-are-set"></a>Данные в диапазоне перед заданием цвета шрифта и цвета заливки

![Данные в Excel перед набором формата.](../images/excel-ranges-format-before.png)

### <a name="data-in-range-after-font-color-and-fill-color-are-set"></a>Данные в диапазоне после задания цвета шрифта и цвета заливки

![Данные в Excel после набора формата.](../images/excel-ranges-format-font-and-fill.png)

## <a name="set-number-format"></a>Задание формата чисел

В примере ниже показано, как задать формат чисел для ячеек в диапазоне **D3:E5**.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    let formats = [
        ["0.00", "0.00"],
        ["0.00", "0.00"],
        ["0.00", "0.00"]
    ];

    let range = sheet.getRange("D3:E5");
    range.numberFormat = formats;

    await context.sync();
});
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
