---
title: Очистить или удалить диапазоны с Excel API JavaScript
description: Узнайте, как очистить или удалить диапазоны с помощью Excel API JavaScript.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: a1bd99db3aa9af3903552d9cefc6ec6d21701136
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075833"
---
# <a name="clear-or-delete-ranges-using-the-excel-javascript-api"></a>Очистить или удалить диапазоны с Excel API JavaScript

В этой статье данная статья содержит примеры кода, которые очищают и удаляют диапазоны с Excel API JavaScript. Полный список свойств и методов, поддерживаемых объектом, см. в `Range` [Excel. Класс Range](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="clear-a-range-of-cells"></a>Очистка диапазона ячеек

В примере кода ниже показано, как удалить все содержимое и форматирование ячеек в диапазоне **E2:E5**.  

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("E2:E5");

    range.clear();

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-before-range-is-cleared"></a>Данные перед очисткой диапазона

![Данные в Excel перед очисткой диапазона.](../images/excel-ranges-start.png)

### <a name="data-after-range-is-cleared"></a>Данные после очистки диапазона

![Данные в Excel после очистки диапазона.](../images/excel-ranges-after-clear.png)

## <a name="delete-a-range-of-cells"></a>Удаление диапазона ячеек

В следующем примере кода удаляются ячейки в диапазоне **B4:E4** и перемещаются другие ячейки для заполнения пространства, освобождаемого удаленными ячейками.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B4:E4");

    range.delete(Excel.DeleteShiftDirection.up);

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-before-range-is-deleted"></a>Данные перед удалением диапазона

![Данные в Excel перед удалением диапазона.](../images/excel-ranges-start.png)

### <a name="data-after-range-is-deleted"></a>Данные после удаления диапазона

![Данные в Excel после удаления диапазона.](../images/excel-ranges-after-delete.png)


## <a name="see-also"></a>См. также

- [Работа с ячейками с Excel API JavaScript](excel-add-ins-cells.md)
- [Настройка и получения диапазонов с Excel API JavaScript](excel-add-ins-ranges-set-get.md)
- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
