---
title: Вставьте диапазоны с Excel API JavaScript
description: Узнайте, как вставить ряд ячеек с Excel API JavaScript.
ms.date: 02/17/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 0e1ed6d2302bcdb4a11688cd6d77448811f8a93b
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340549"
---
# <a name="insert-a-range-of-cells-using-the-excel-javascript-api"></a>Вставьте диапазон ячеек с Excel API JavaScript

В этой статье приводится пример кода, который вставляет ряд ячеек с Excel API JavaScript. Полный список свойств `Range` и методов, поддерживаемых объектом, см. [в Excel. Класс Range](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="insert-a-range-of-cells"></a>Вставка диапазона ячеек

В примере кода ниже показано, как вставить диапазон ячеек в расположение **B4:E4** и сдвинуть другие ячейки вниз, чтобы освободить место для новых ячеек.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let range = sheet.getRange("B4:E4");

    range.insert(Excel.InsertShiftDirection.down);

    await context.sync();
});
```

### <a name="data-before-range-is-inserted"></a>Данные перед вставкой диапазона

![Данные в Excel перед вставкой диапазона.](../images/excel-ranges-start.png)

### <a name="data-after-range-is-inserted"></a>Данные после вставки диапазона

![Данные в Excel после вставки диапазона.](../images/excel-ranges-after-insert.png)

## <a name="see-also"></a>См. также

- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
- [Работа с ячейками с Excel API JavaScript](excel-add-ins-cells.md)
- [Очистить или удалить диапазоны с Excel API JavaScript](excel-add-ins-ranges-clear-delete.md)
