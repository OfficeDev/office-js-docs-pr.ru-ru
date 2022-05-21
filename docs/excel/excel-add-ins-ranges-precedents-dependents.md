---
title: Работа с зависимыми и зависимыми формулами с помощью Excel API JavaScript
description: Узнайте, как использовать API JavaScript Excel для извлечения приоритетов формул и зависимых элементов.
ms.date: 05/19/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: ca432b7eb6825781960e995af2ed2193c7caa5e2
ms.sourcegitcommit: 4ca3334f3cefa34e6b391eb92a429a308229fe89
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/21/2022
ms.locfileid: "65628098"
---
# <a name="get-formula-precedents-and-dependents-using-the-excel-javascript-api"></a>Получение приоритетов формул и зависимых элементов с Excel API JavaScript

Excel формулы часто ссылаются на другие ячейки. Эти ссылки между ячейками называются "влияющими" и "зависимыми". Приоритетом является ячейка, которая предоставляет данные формуле. Зависимая — это ячейка, содержащая формулу, которая ссылается на другие ячейки. Дополнительные сведения о функциях Excel, связанных со связями между ячейками, см. в статье "Отображение связей между [формулами и ячейками"](https://support.microsoft.com/office/a59bef2b-3701-46bf-8ff1-d3518771d507).

У ячейки, которая имеется в качестве приоритета, могут быть собственные ячейки- Каждая ячейка, которая имеет приоритет в этой цепочке, по-прежнему является приоритетом исходной ячейки. Для зависимых элементов существует та же связь. Любая ячейка, затронутая другой ячейкой, является зависимой от этой ячейки. "Прямой приоритет" — это первая предшествущая группа ячеек в этой последовательности, аналогичная концепции родительских элементов в связи "родители-потомки". "Прямая зависимость" — это первая зависимая группа ячеек в последовательности, аналогичная дочерним элементам в связи "родители-потомки".

В этой статье приводятся примеры кода, которые получают приоритеты и зависимые формулы с помощью Excel API JavaScript. Полный список свойств и `Range` методов, поддерживаемых объектом, см. в разделе "Объект [Range" (API JavaScript](/javascript/api/excel/excel.range) для Excel).

## <a name="get-the-precedents-of-a-formula"></a>Получение приоритетов формулы

Найдите ячейки-ячейки формулы с [помощью Range.getPrecedents](/javascript/api/excel/excel.range#excel-excel-range-getprecedents-member(1)). `Range.getPrecedents` возвращает объект `WorkbookRangeAreas` . Этот объект содержит адреса всех приоритетов в книге. Он имеет отдельный `RangeAreas` объект для каждого листа, содержащий по крайней мере один приоритет формулы. Дополнительные сведения об объекте `RangeAreas` см. в статье "Работа с несколькими [диапазонами Excel надстроек"](excel-add-ins-multiple-ranges.md).

Чтобы найти только прямые ячейки-ячейки формулы, используйте [Range.getDirectPrecedents](/javascript/api/excel/excel.range#excel-excel-range-getdirectprecedents-member(1)). `Range.getDirectPrecedents` работает аналогично `Range.getPrecedents` и возвращает объект `WorkbookRangeAreas` , содержащий адреса прямых приоритетов.

На следующем снимке экрана показан результат нажатия кнопки **Trace Precedents** в Excel пользовательском интерфейсе. Эта кнопка рисует стрелку из ячеек-контейнеров в выделенную ячейку. Выделенная ячейка **E3** содержит формулу "=C3 * D3", поэтому **ячейки C3** и **D3** являются ячейками-приоритетами. В отличие Excel пользовательского интерфейса, `getPrecedents` методы и методы `getDirectPrecedents` не рисуют стрелки.

![Трассировка ячеек со стрелками в Excel пользовательского интерфейса.](../images/excel-ranges-trace-precedents.png)

> [!IMPORTANT]
> Методы `getPrecedents` и `getDirectPrecedents` методы не извлекаются ячейки-ячейки в книгах.

В следующем примере кода показано, как работать с методами `Range.getPrecedents` и методами `Range.getDirectPrecedents` . Пример получает приоритеты активного диапазона, а затем изменяет цвет фона этих ячеек. Цвет фона непосредственных ячеек-приоритетов имеет желтый цвет, а цвет фона других ячеек- в качестве оранжевого.

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

## <a name="get-the-dependents-of-a-formula"></a>Получение зависимости формулы

Найдите зависимые ячейки формулы с [помощью Range.getDependents](/javascript/api/excel/excel.range#excel-excel-range-getdependents-member(1)). Like `Range.getPrecedents`, `Range.getDependents` также возвращает `WorkbookRangeAreas` объект. Этот объект содержит адреса всех зависимых элементов в книге. Он имеет отдельный объект `RangeAreas` для каждого листа, содержащий по крайней мере одну зависимую формулу. Дополнительные сведения о работе с объектом `RangeAreas` см. в статье "Работа с несколькими диапазонами Excel [надстроек"](excel-add-ins-multiple-ranges.md).

Чтобы найти только прямые зависимые ячейки формулы, используйте [Range.getDirectDependents](/javascript/api/excel/excel.range#excel-excel-range-getdirectdependents-member(1)). `Range.getDirectDependents` работает аналогично `Range.getDependents` и возвращает объект `WorkbookRangeAreas` , содержащий адреса прямых зависимых элементов.

На следующем снимке экрана показан результат нажатия кнопки **"** Зависимые от трассировки" Excel пользовательского интерфейса. Эта кнопка рисует стрелку из выделенной ячейки в зависимые ячейки. Выделенная ячейка **D3** имеет ячейку **E3** в качестве зависимой. **E3** содержит формулу "=C3 * D3". В отличие Excel пользовательского интерфейса, `getDependents` методы и методы `getDirectDependents` не рисуют стрелки.

![Трассировка зависимых ячеек со стрелками в Excel пользовательского интерфейса.](../images/excel-ranges-trace-dependents.png)

> [!IMPORTANT]
> Методы `getDependents` и `getDirectDependents` зависимые ячейки не извлекаются между книгами.

Следующий пример кода получает прямые зависимые элементы активного диапазона, а затем изменяет цвет фона этих зависимых ячеек на желтый.

В следующем примере кода показано, как работать с методами `Range.getDependents` и методами `Range.getDirectDependents` . Образец получает зависимые элементы активного диапазона, а затем изменяет цвет фона этих зависимых ячеек. Цвет фона прямых зависимых ячеек имеет желтый цвет, а цвет фона других зависимых ячеек — оранжевый.

```js
// This code sample shows how to find and highlight the dependents 
// and direct dependents of the currently selected cell.
await Excel.run(async (context) => {
    let range = context.workbook.getActiveCell();
    // Dependents are all cells that contain formulas that refer to other cells.
    let dependents = range.getDependents();  
    // Direct dependents are the child cells, or the first succeeding group of cells in a sequence of cells that refer to other cells.
    let directDependents = range.getDirectDependents();

    range.load("address");
    dependents.areas.load("address");    
    directDependents.areas.load("address");
    
    await context.sync();

    console.log(`All dependent cells of ${range.address}:`);
    
    // Use the dependents API to loop through all dependents of the active cell.
    for (let i = 0; i < dependents.areas.items.length; i++) {
      // Highlight and print out the addresses of all dependent cells.
      dependents.areas.items[i].format.fill.color = "Orange";
      console.log(`  ${dependents.areas.items[i].address}`);
    }

    console.log(`Direct dependent cells of ${range.address}:`);

    // Use the direct dependents API to loop through direct dependents of the active cell.
    for (let i = 0; i < directDependents.areas.items.length; i++) {
      // Highlight and print the address of each dependent cell.
      directDependents.areas.items[i].format.fill.color = "Yellow";
      console.log(`  ${directDependents.areas.items[i].address}`);
    }
});
```

## <a name="see-also"></a>Дополнительные ресурсы

- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
- [Работа с ячейками с помощью Excel API JavaScript](excel-add-ins-cells.md)
- [Работа с несколькими диапазонами одновременно в надстройках Excel](excel-add-ins-multiple-ranges.md)
