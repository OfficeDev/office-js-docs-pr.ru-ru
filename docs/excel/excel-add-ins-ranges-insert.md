---
title: Вставьте диапазоны с Excel API JavaScript
description: Узнайте, как вставить ряд ячеек с Excel API JavaScript.
ms.date: 04/02/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: e14aeb030e01dbf170d3acc1edd4952b4989a557
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59150862"
---
# <a name="insert-a-range-of-cells-using-the-excel-javascript-api"></a>Вставьте диапазон ячеек с Excel API JavaScript

В этой статье содержится пример кода, который вставляет ряд ячеек с Excel API JavaScript. Полный список свойств и методов, поддерживаемых объектом, см. `Range` [в Excel. Класс Range](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="insert-a-range-of-cells"></a>Вставка диапазона ячеек

В примере кода ниже показано, как вставить диапазон ячеек в расположение **B4:E4** и сдвинуть другие ячейки вниз, чтобы освободить место для новых ячеек.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.insert(Excel.InsertShiftDirection.down);

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-before-range-is-inserted"></a>Данные перед вставкой диапазона

![Данные в Excel перед вставкой диапазона.](../images/excel-ranges-start.png)

### <a name="data-after-range-is-inserted"></a>Данные после вставки диапазона

![Данные в Excel после вставки диапазона.](../images/excel-ranges-after-insert.png)

## <a name="see-also"></a>Дополнительные материалы

- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
- [Работа с ячейками с Excel API JavaScript](excel-add-ins-cells.md)
- [Очистить или удалить диапазоны с Excel API JavaScript](excel-add-ins-ranges-clear-delete.md)
