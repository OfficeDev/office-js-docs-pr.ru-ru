---
title: Работа с прецедентами формул и зависимыми с помощью Excel API JavaScript
description: Узнайте, как использовать API Excel JavaScript для получения прецедентов формул и зависимых.
ms.date: 11/30/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 60da910879fc48f1564d43cf3f87c2a5bf930fbe
ms.sourcegitcommit: 5daf91eb3be99c88b250348186189f4dc1270956
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/01/2021
ms.locfileid: "61242063"
---
# <a name="get-formula-precedents-and-dependents-using-the-excel-javascript-api"></a>Получите прецеденты формул и иждивенцев с Excel API JavaScript

Excel часто ссылаются на другие ячейки. Эти межклеточные ссылки называются "прецедентами" и "зависимыми". Прецедент — это ячейка, которая предоставляет данные формуле. Зависимая ячейка содержит формулу, которая ссылается на другие ячейки. Дополнительные дополнительные Excel, связанные с отношениями между ячейками, см. в руб. Отображение взаимосвязей между [формулами и ячейками.](https://support.microsoft.com/office/a59bef2b-3701-46bf-8ff1-d3518771d507)

Ячейка прецедента может иметь свои собственные ячейки прецедента. Каждая ячейка прецедентов в этой цепочке прецедентов по-прежнему является прецедентом исходной ячейки. Одинаковые отношения существуют и для иждивенцев. Любая ячейка, затрагиваемая другой ячейкой, зависит от этой ячейки. "Прямой прецедент" является первой предыдущей группой ячеек в этой последовательности, аналогичной концепции родителей в родительских отношениях с ребенком. "Прямая зависимость" — это первая зависимая группа ячеек в последовательности, похожая на детей в отношениях между родителем и ребенком.

В этой статье приводится пример кода, который извлекает прецеденты и зависит от формул с Excel API JavaScript. Полный список свойств и методов, поддерживаемых объектом, см. в руб. `Range` [Range Object (API JavaScript для Excel).](/javascript/api/excel/excel.range)

## <a name="get-the-precedents-of-a-formula"></a>Получить прецеденты формулы

Найдите прецедентные ячейки формулы [с помощью Range.getPrecedents.](/javascript/api/excel/excel.range#getPrecedents__) `Range.getPrecedents` возвращает `WorkbookRangeAreas` объект. Этот объект содержит адреса всех прецедентов в книге. Для каждого таблицы имеется отдельный объект, содержащий по `RangeAreas` крайней мере один прецедент формулы. Дополнительные новости об объекте см. в добавлении `RangeAreas` [Work with multiple ranges Excel надстройки.](excel-add-ins-multiple-ranges.md)

Чтобы найти только прямые ячейки-прецеденты формулы, используйте [Range.getDirectPrecedents.](/javascript/api/excel/excel.range#getDirectPrecedents__) `Range.getDirectPrecedents` работает как `Range.getPrecedents` и возвращает `WorkbookRangeAreas` объект, содержащий адреса прямых прецедентов.

На следующем скриншоте показан результат выбора кнопки **Trace Precedents** в пользовательском Excel интерфейсе. Эта кнопка рисует стрелку из ячеек-прецедентов в выбранную ячейку. Выбранная ячейка **E3** содержит формулу "=C3 * D3", поэтому **C3** и **D3** являются прецедентными ячейками. В отличие Excel кнопки пользовательского интерфейса, стрелки и методы не `getPrecedents` `getDirectPrecedents` рисуют.

![Отслеживание прецедентных ячеек стрелки в Excel пользовательского интерфейса.](../images/excel-ranges-trace-precedents.png)

> [!IMPORTANT]
> Эти `getPrecedents` методы и методы не `getDirectPrecedents` извлекать ячейки прецедентов в книгах.

В следующем примере кода показано, как работать с `Range.getPrecedents` методами и `Range.getDirectPrecedents` методами. В примере получаются прецеденты для активного диапазона, а затем изменяется фоновый цвет этих ячеек-прецедентов. Фоновый цвет прямых ячеек-прецедентов за установлен желтым, а цвет фона других ячеек прецедента — оранжевым.

```js
// This code sample shows how to find and highlight the precedents 
// and direct precedents of the currently selected cell.
Excel.run(function (context) {
  var range = context.workbook.getActiveCell();
  // Precedents are all cells that provide data to the selected formula.
  var precedents = range.getPrecedents();
  // Direct precedents are the parent cells, or the first preceding group of cells that provide data to the selected formula.    
  var directPrecedents = range.getDirectPrecedents();

  range.load("address");
  precedents.areas.load("address");
  directPrecedents.areas.load("address");
  
  return context.sync()
    .then(function () {
      console.log(`All precedent cells of ${range.address}:`);
      
      // Use the precedents API to loop through all precedents of the active cell.
      for (var i = 0; i < precedents.areas.items.length; i++) {
        // Highlight and print out the address of all precedent cells.
        precedents.areas.items[i].format.fill.color = "Orange";
        console.log(`  ${precedents.areas.items[i].address}`);
      }

      console.log(`Direct precedent cells of ${range.address}:`);

      // Use the direct precedents API to loop through direct precedents of the active cell.
      for (var i = 0; i < directPrecedents.areas.items.length; i++) {
        // Highlight and print out the address of each direct precedent cell.
        directPrecedents.areas.items[i].format.fill.color = "Yellow";
        console.log(`  ${directPrecedents.areas.items[i].address}`);
      }
    });
}).catch(errorHandlerFunction);
```

## <a name="get-the-direct-dependents-of-a-formula"></a>Получить прямые иждивенцы формулы

Найдите прямые зависимые ячейки формулы [с помощью Range.getDirectDependents.](/javascript/api/excel/excel.range#getDirectDependents__) Как `Range.getDirectPrecedents` , также возвращает `Range.getDirectDependents` `WorkbookRangeAreas` объект. Этот объект содержит адреса всех прямых иждивенцев в книге. Он имеет отдельный `RangeAreas` объект для каждого таблицы, содержащего по крайней мере одну зависимую формулу. Дополнительные сведения о работе с объектом см. в совместной работе с несколькими диапазонами `RangeAreas` [Excel надстройки.](excel-add-ins-multiple-ranges.md)

На следующем скриншоте показан результат выбора кнопки **Trace Dependents** в пользовательском Excel интерфейсе. Эта кнопка рисует стрелку из зависимых ячеек в выбранную ячейку. Выбранная ячейка **D3** имеет ячейку **E3** в качестве зависимой. **E3** содержит формулу "=C3 * D3". В отличие Excel пользовательского интерфейса, `getDirectDependents` метод не рисует стрелки.

![Отслеживание зависимых ячеек стрелки Excel пользовательского интерфейса.](../images/excel-ranges-trace-dependents.png)

> [!IMPORTANT]
> Метод `getDirectDependents` не извлекает зависимые ячейки в книгах.

В следующем примере кода получаются прямые иждивенцы для активного диапазона, а затем изменяется фоновый цвет этих зависимых ячеек на желтый.

```js
// This code sample shows how to find and highlight the dependents of the currently selected cell.
Excel.run(function (context) {
    // Direct dependents are cells that contain formulas that refer to other cells.
    var range = context.workbook.getActiveCell();
    var directDependents = range.getDirectDependents();
    range.load("address");
    directDependents.areas.load("address");
    
    return context.sync()
        .then(function () {
            console.log(`Direct dependent cells of ${range.address}:`);
    
            // Use the direct dependents API to loop through direct dependents of the active cell.
            for (var i = 0; i < directDependents.areas.items.length; i++) {
              // Highlight and print the address of each dependent cell.
              directDependents.areas.items[i].format.fill.color = "Yellow";
              console.log(`  ${directDependents.areas.items[i].address}`);
            }
        });
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a>См. также

- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
- [Работа с ячейками с Excel API JavaScript](excel-add-ins-cells.md)
- [Работа с несколькими диапазонами одновременно в надстройках Excel](excel-add-ins-multiple-ranges.md)
