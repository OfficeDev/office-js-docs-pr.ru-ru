---
title: Работа с прецедентами формул с помощью API JavaScript Excel
description: Узнайте, как использовать API JavaScript Excel для получения прецедентов формул.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 0d21ae411615a22873a0f4dda185984f6191ac8e
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652915"
---
# <a name="get-formula-precedents-using-the-excel-javascript-api"></a>Получите прецеденты формул с помощью API JavaScript Excel

В этой статье приводится пример кода, который извлекает прецеденты формул с помощью API JavaScript Excel. Полный список свойств и методов, поддерживаемых объектом, см. в `Range` [класс Excel.Range.](/javascript/api/excel/excel.range)

## <a name="get-formula-precedents"></a>Получить прецеденты формул

Формула Excel часто ссылается на другие ячейки. Когда ячейка предоставляет данные формуле, она называется формулой "прецедент". Дополнительные новости о свойствах Excel, связанных с отношениями между ячейками, см. в дополнительных подробностях отображения взаимосвязей между [формулами и ячейками.](https://support.microsoft.com/office/display-the-relationships-between-formulas-and-cells-a59bef2b-3701-46bf-8ff1-d3518771d507) 

С [помощью Range.getDirectPrecedents](/javascript/api/excel/excel.range#getdirectprecedents--)надстройка может найти прямые ячейки прецедента формулы. `Range.getDirectPrecedents` возвращает `WorkbookRangeAreas` объект. Этот объект содержит адреса всех прецедентов в книге. Для каждого таблицы имеется отдельный объект, содержащий по `RangeAreas` крайней мере один прецедент формулы. Дополнительные сведения о работе с объектом см. в совместной работе с несколькими диапазонами в `RangeAreas` [надстройки Excel.](excel-add-ins-multiple-ranges.md)

В пользовательском интерфейсе Excel кнопка **Trace Precedents** рисует стрелку из ячеек-прецедентов в выбранную формулу. В отличие от кнопки пользовательского интерфейса Excel, `getDirectPrecedents` метод не рисует стрелки. 

> [!IMPORTANT]
> Метод `getDirectPrecedents` не может получить ячейки прецедента в книгах. 

В следующем примере кода получаются прямые прецеденты для активного диапазона, а затем изменяется фоновый цвет этих ячеек-прецедентов на желтый. 

> [!NOTE]
> Активный диапазон должен содержать формулу, которая ссылается на другие ячейки в той же книге, чтобы выделение работало правильно. 

```js
Excel.run(function (context) {
    // Precedents are cells that provide data to the selected formula.
    var range = context.workbook.getActiveCell();
    var directPrecedents = range.getDirectPrecedents();
    range.load("address");
    directPrecedents.areas.load("address");
    
    return context.sync()
        .then(function () {
            console.log(`Direct precedent cells of ${range.address}:`);

            // Use the direct precedents API to loop through precedents of the active cell.
            for (var i = 0; i < directPrecedents.areas.items.length; i++) {
              // Highlight and print out the address of each precedent cell.
              directPrecedents.areas.items[i].format.fill.color = "Yellow";
              console.log(`  ${directPrecedents.areas.items[i].address}`);
            }
        })
        .then(context.sync);
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a>См. также

- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
- [Работа с ячейками с помощью API JavaScript Excel](excel-add-ins-cells.md)
- [Работа с несколькими диапазонами одновременно в надстройках Excel](excel-add-ins-multiple-ranges.md)
