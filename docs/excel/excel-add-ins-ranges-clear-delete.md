---
title: Очистить или удалить диапазоны с Excel API JavaScript
description: Узнайте, как очистить или удалить диапазоны с Excel API JavaScript.
ms.date: 02/16/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 7336a0e6485ce502216818b4a8cd077fed0069c3
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340710"
---
# <a name="clear-or-delete-ranges-using-the-excel-javascript-api"></a>Очистить или удалить диапазоны с Excel API JavaScript

В этой статье данная статья содержит примеры кода, которые очищают и удаляют диапазоны с Excel API JavaScript. Полный список свойств и `Range` методов, поддерживаемых объектом, см. в Excel[. Класс Range](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="clear-a-range-of-cells"></a>Очистка диапазона ячеек

В примере кода ниже показано, как удалить все содержимое и форматирование ячеек в диапазоне **E2:E5**.  

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let range = sheet.getRange("E2:E5");

    range.clear();

    await context.sync();
});
```

### <a name="data-before-range-is-cleared"></a>Данные перед очисткой диапазона

![Данные в Excel до очистки диапазона.](../images/excel-ranges-start.png)

### <a name="data-after-range-is-cleared"></a>Данные после очистки диапазона

![Данные в Excel после очистки диапазона.](../images/excel-ranges-after-clear.png)

## <a name="delete-a-range-of-cells"></a>Удаление диапазона ячеек

В следующем примере кода удаляются ячейки в диапазоне **B4:E4** и перемещаются другие ячейки для заполнения пространства, освобождаемого удаленными ячейками.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let range = sheet.getRange("B4:E4");

    range.delete(Excel.DeleteShiftDirection.up);

    await context.sync();
});
```

### <a name="data-before-range-is-deleted"></a>Данные перед удалением диапазона

![Данные в Excel перед удалением диапазона.](../images/excel-ranges-start.png)

### <a name="data-after-range-is-deleted"></a>Данные после удаления диапазона

![Данные в Excel после удаления диапазона.](../images/excel-ranges-after-delete.png)

## <a name="see-also"></a>См. также

- [Работа с ячейками с Excel API JavaScript](excel-add-ins-cells.md)
- [Настройка и получения диапазонов с Excel API JavaScript](excel-add-ins-ranges-set-get.md)
- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
