---
title: Работа с прецедентами формул и зависимыми с помощью Excel API JavaScript
description: Узнайте, как использовать API Excel JavaScript для получения прецедентов формул и зависимых.
ms.date: 02/17/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 9d686b416b271dce81ee072a98f8cb9e1dac65b2
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/23/2022
ms.locfileid: "63744087"
---
# <a name="get-formula-precedents-and-dependents-using-the-excel-javascript-api"></a>Получите прецеденты формул и иждивенцев с Excel API JavaScript

Excel часто ссылаются на другие ячейки. Эти межклеточные ссылки называются "прецедентами" и "зависимыми". Прецедент — это ячейка, которая предоставляет данные формуле. Зависимая ячейка содержит формулу, которая ссылается на другие ячейки. Дополнительные информацию о Excel, связанных с отношениями между ячейками, см[](https://support.microsoft.com/office/a59bef2b-3701-46bf-8ff1-d3518771d507). в этой информации.

Ячейка прецедента может иметь свои собственные ячейки прецедента. Каждая ячейка прецедентов в этой цепочке прецедентов по-прежнему является прецедентом исходной ячейки. Одинаковые отношения существуют и для иждивенцев. Любая ячейка, затрагиваемая другой ячейкой, зависит от этой ячейки. "Прямой прецедент" является первой предыдущей группой ячеек в этой последовательности, аналогичной концепции родителей в родительских отношениях с ребенком. "Прямая зависимость" — это первая зависимая группа ячеек в последовательности, похожая на детей в отношениях между родителем и ребенком.

В этой статье приводится пример кода, который извлекает прецеденты и зависит от формул с Excel API JavaScript. Полный список свойств `Range` и методов, поддерживаемых объектом, см. в списке [Range Object (API JavaScript для Excel)](/javascript/api/excel/excel.range).

## <a name="get-the-precedents-of-a-formula"></a>Получить прецеденты формулы

Найдите ячейки-прецеденты формулы [с помощью Range.getPrecedents](/javascript/api/excel/excel.range#excel-excel-range-getprecedents-member(1)). `Range.getPrecedents` возвращает объект `WorkbookRangeAreas` . Этот объект содержит адреса всех прецедентов в книге. Для каждого таблицы `RangeAreas` имеется отдельный объект, содержащий по крайней мере один прецедент формулы. Дополнительные новости об объекте `RangeAreas` см. в добавлении [Work with multiple ranges Excel надстройки](excel-add-ins-multiple-ranges.md).

Чтобы найти только прямые ячейки-прецеденты формулы, используйте [Range.getDirectPrecedents](/javascript/api/excel/excel.range#excel-excel-range-getdirectprecedents-member(1)). `Range.getDirectPrecedents` работает как `Range.getPrecedents` и возвращает объект `WorkbookRangeAreas` , содержащий адреса прямых прецедентов.

На следующем скриншоте показан результат выбора кнопки **Trace Precedents** в пользовательском Excel интерфейсе. Эта кнопка рисует стрелку из ячеек-прецедентов в выбранную ячейку. Выбранная ячейка **E3** содержит формулу "=C3 * D3", поэтому **C3** и **D3** являются прецедентными ячейками. В отличие Excel пользовательского интерфейса, стрелки `getPrecedents` и `getDirectPrecedents` методы не рисуют.

![Отслеживание прецедентных ячеек стрелки Excel пользовательского интерфейса.](../images/excel-ranges-trace-precedents.png)

> [!IMPORTANT]
> Эти `getPrecedents` методы `getDirectPrecedents` и методы не извлекать ячейки прецедентов в книгах.

В следующем примере кода показано, как работать с методами `Range.getPrecedents` и методами `Range.getDirectPrecedents` . В примере получаются прецеденты для активного диапазона, а затем изменяется фоновый цвет этих ячеек-прецедентов. Фоновый цвет прямых ячеек-прецедентов за установлен желтым, а цвет фона других ячеек прецедента — оранжевым.

```js
// This code sample shows how to find and highlight the precedents 
// and direct precedents of the currently selected cell.
await Excel.run(async (context) => {
  let range = context.workbook.getActiveCell();
  // Precedents are all cells that provide data to the selected formula.
  let precedents = range.getPrecedents();
  // Direct precedents are the parent cells, or the first preceding group of cells that provide data to the selected formula.    
  let directPrecedents = range.getDirectPrecedents();

  range.load("address");
  precedents.areas.load("address");
  directPrecedents.areas.load("address");
  
  await context.sync();

  console.log(`All precedent cells of ${range.address}:`);
  
  // Use the precedents API to loop through all precedents of the active cell.
  for (let i = 0; i < precedents.areas.items.length; i++) {
    // Highlight and print out the address of all precedent cells.
    precedents.areas.items[i].format.fill.color = "Orange";
    console.log(`  ${precedents.areas.items[i].address}`);
  }

  console.log(`Direct precedent cells of ${range.address}:`);

  // Use the direct precedents API to loop through direct precedents of the active cell.
  for (let i = 0; i < directPrecedents.areas.items.length; i++) {
    // Highlight and print out the address of each direct precedent cell.
    directPrecedents.areas.items[i].format.fill.color = "Yellow";
    console.log(`  ${directPrecedents.areas.items[i].address}`);
  }
});
```

## <a name="get-the-direct-dependents-of-a-formula"></a>Получить прямые иждивенцы формулы

Найдите прямые зависимые ячейки формулы [с помощью Range.getDirectDependents](/javascript/api/excel/excel.range#excel-excel-range-getdirectdependents-member(1)). Как `Range.getDirectPrecedents`, `Range.getDirectDependents` также возвращает `WorkbookRangeAreas` объект. Этот объект содержит адреса всех прямых иждивенцев в книге. Он имеет отдельный `RangeAreas` объект для каждого таблицы, содержащего по крайней мере одну зависимую формулу. Дополнительные сведения о работе с объектом `RangeAreas` см. в совместной работе с несколькими [диапазонами Excel надстройки](excel-add-ins-multiple-ranges.md).

На следующем скриншоте показан результат выбора кнопки **Trace Dependents** в Excel пользовательском интерфейсе. Эта кнопка рисует стрелку из зависимых ячеек в выбранную ячейку. Выбранная ячейка **D3** имеет ячейку **E3** в качестве зависимой. **E3** содержит формулу "=C3 * D3". В отличие Excel пользовательского интерфейса, метод `getDirectDependents` не рисует стрелки.

![Отслеживание зависимых ячеек стрелки Excel пользовательского интерфейса.](../images/excel-ranges-trace-dependents.png)

> [!IMPORTANT]
> Метод `getDirectDependents` не извлекает зависимые ячейки в книгах.

В следующем примере кода получаются прямые иждивенцы для активного диапазона, а затем изменяется фоновый цвет этих зависимых ячеек на желтый.

```js
// This code sample shows how to find and highlight the dependents of the currently selected cell.
await Excel.run(async (context) => {
    // Direct dependents are cells that contain formulas that refer to other cells.
    let range = context.workbook.getActiveCell();
    let directDependents = range.getDirectDependents();
    range.load("address");
    directDependents.areas.load("address");
    
    await context.sync();
    console.log(`Direct dependent cells of ${range.address}:`);

    // Use the direct dependents API to loop through direct dependents of the active cell.
    for (let i = 0; i < directDependents.areas.items.length; i++) {
      // Highlight and print the address of each dependent cell.
      directDependents.areas.items[i].format.fill.color = "Yellow";
      console.log(`  ${directDependents.areas.items[i].address}`);
    }
});
```

## <a name="see-also"></a>См. также

- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
- [Работа с ячейками с Excel API JavaScript](excel-add-ins-cells.md)
- [Работа с несколькими диапазонами одновременно в надстройках Excel](excel-add-ins-multiple-ranges.md)
