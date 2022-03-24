---
title: Диапазоны групп с Excel API JavaScript
description: Узнайте, как сгруппить строки или столбцы диапазона вместе, чтобы создать контур с Excel API JavaScript.
ms.date: 02/17/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: f1150bdbe198fc2ebc7e026dedd49358006f6bc6
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745203"
---
# <a name="group-ranges-for-an-outline-using-the-excel-javascript-api"></a>Диапазоны групп для контура с Excel API JavaScript

В этой статье приводится пример кода, который показывает, как группировать диапазоны для контура с Excel API JavaScript. Полный список свойств `Range` и методов, поддерживаемый объектом, см. в Excel[. Класс Range](/javascript/api/excel/excel.range).

## <a name="group-rows-or-columns-of-a-range-for-an-outline"></a>Групповые строки или столбцы диапазона для контура

Строки или столбцы диапазона можно сгруппить для создания [контура](https://support.microsoft.com/office/08ce98c4-0063-4d42-8ac7-8278c49e9aff). Эти группы можно свернуть и расширить, чтобы скрыть и показать соответствующие ячейки. Это упрощает быстрый анализ данных верхнего верхней строки. Чтобы сделать эти группы контуров, используйте [Range.group](/javascript/api/excel/excel.range#excel-excel-range-group-member(1)) .

Контур может иметь иерархию, в которой небольшие группы вложены в более крупные группы. Это позволяет просматривать контуры на разных уровнях. Изменение уровня видимых контуров можно сделать программным путем с помощью метода [Worksheet.showOutlineLevels](/javascript/api/excel/excel.worksheet#excel-excel-worksheet-showoutlinelevels-member(1)) . Обратите внимание, Excel поддерживает только восемь уровней групп контуров.

В следующем примере кода создается контур с двумя уровнями групп для строк и столбцов. На последующем изображении показаны группировки этого контура. В примере кода диапазоны, которые группуются, не включают строку или столбец управления контурами (в этом примере "Итоги"). Группа определяет, что будет свернуто, а не строка или столбец с управлением.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");

    // Group the larger, main level. Note that the outline controls
    // will be on row 10, meaning 4-9 will collapse and expand.
    sheet.getRange("4:9").group(Excel.GroupOption.byRows);

    // Group the smaller, sublevels. Note that the outline controls
    // will be on rows 6 and 9, meaning 4-5 and 7-8 will collapse and expand.
    sheet.getRange("4:5").group(Excel.GroupOption.byRows);
    sheet.getRange("7:8").group(Excel.GroupOption.byRows);

    // Group the larger, main level. Note that the outline controls
    // will be on column R, meaning C-Q will collapse and expand.
    sheet.getRange("C:Q").group(Excel.GroupOption.byColumns);

    // Group the smaller, sublevels. Note that the outline controls
    // will be on columns G, L, and R, meaning C-F, H-K, and M-P will collapse and expand.
    sheet.getRange("C:F").group(Excel.GroupOption.byColumns);
    sheet.getRange("H:K").group(Excel.GroupOption.byColumns);
    sheet.getRange("M:P").group(Excel.GroupOption.byColumns);
    await context.sync();
});
```

![Диапазон с двухуровневой двухмерной схемой.](../images/excel-outline.png)

## <a name="remove-grouping-from-rows-or-columns-of-a-range"></a>Удаление группировки из строк или столбцов диапазона

Чтобы разгруппировать строку или группу столбцов, используйте [метод Range.ungroup](/javascript/api/excel/excel.range#excel-excel-range-ungroup-member(1)) . Это удаляет внешний уровень из контура. Если несколько групп одного и того же типа строки или столбца находятся на одном уровне в указанном диапазоне, все эти группы негруппировываются.

## <a name="see-also"></a>См. также

- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
- [Работа с ячейками с Excel API JavaScript](excel-add-ins-cells.md)
- [Работа с несколькими диапазонами одновременно в надстройках Excel](excel-add-ins-multiple-ranges.md)
