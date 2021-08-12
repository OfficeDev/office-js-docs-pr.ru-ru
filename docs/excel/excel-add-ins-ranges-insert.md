---
title: Вставьте диапазоны с Excel API JavaScript
description: Узнайте, как вставить ряд ячеек с Excel API JavaScript.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: ad75b8c107005777047418ff9a1824665552cb5cca06c1e858f3645172f12e7c
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57084651"
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

## <a name="see-also"></a>См. также

- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
- [Работа с ячейками с Excel API JavaScript](excel-add-ins-cells.md)
- [Очистить или удалить диапазоны с Excel API JavaScript](excel-add-ins-ranges-clear-delete.md)
