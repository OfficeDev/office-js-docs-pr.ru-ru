---
title: Удаление дубликатов с Excel API JavaScript
description: Узнайте, как использовать API Excel JavaScript для удаления дубликатов.
ms.date: 02/17/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 5bd6ff6a6c0889f447e362d200e73bba499f572c
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745348"
---
# <a name="remove-duplicates-using-the-excel-javascript-api"></a>Удаление дубликатов с Excel API JavaScript

В этой статье приводится пример кода, который удаляет дублирующиеся записи в диапазоне с Excel API JavaScript. Полный список свойств `Range` и методов, поддерживаемый объектом, см. в Excel[. Класс Range](/javascript/api/excel/excel.range).

## <a name="remove-rows-with-duplicate-entries"></a>Удаление строк с дублирующими записями

Метод [Range.removeDuplicates](/javascript/api/excel/excel.range#excel-excel-range-removeduplicates-member(1)) удаляет строки с дублирующимися записями в указанных столбцах. Метод проходит через каждую строку в диапазоне от самого низкого значения индекса до индекса с самым высоким значением в диапазоне (сверху донизу). Строка удаляется, если значение в ее указанном столбце или столбцах уже встречалось в диапазоне. Строки в диапазоне под удаленной строкой сдвигаются вверх. Функция `removeDuplicates` не влияет на положение ячеек вне диапазона.

Функция `removeDuplicates` использует параметр `number[]`, представляющий индексы столбцов, которые проверяются на наличие дубликатов. Этот массив отсчитывается от нуля относительно диапазона, а не листа. Метод также принимает параметр boolean, который указывает, является ли первая строка загонщиком. При значении **true** верхняя строка игнорируется при поиске дубликатов. Метод `removeDuplicates` возвращает объект, `RemoveDuplicatesResult` который указывает количество удаленных строк и количество оставшихся уникальных строк.

При использовании метода диапазона `removeDuplicates` имейте в виду следующее.

- Функция `removeDuplicates` рассматривает значения ячеек, а не результаты функций. Если две разные функции вычисляют одинаковый результат, значения ячеек не считаются повторяющимися.
- Пустые ячейки не игнорируются функцией `removeDuplicates`. Значение пустой ячейки обрабатывается как любое другое значение. Это означает, что пустые строки, содержащиеся в диапазоне, будут включены в объект `RemoveDuplicatesResult`.

В следующем примере кода показано удаление записей с дублирующими значениями в первом столбце.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let range = sheet.getRange("B2:D11");

    let deleteResult = range.removeDuplicates([0],true);
    deleteResult.load();

    await context.sync();

    console.log(deleteResult.removed + " entries with duplicate names removed.");
    console.log(deleteResult.uniqueRemaining + " entries with unique names remain in the range.");
});
```

### <a name="data-before-duplicate-entries-are-removed"></a>Данные перед удалением дублирующих записей

![Данные в Excel перед запуском метода удаления дубликатов диапазона.](../images/excel-ranges-remove-duplicates-before.png)

### <a name="data-after-duplicate-entries-are-removed"></a>Данные после удаления дублирующих записей

![Данные в Excel после запуска метода удаления дубликатов диапазона.](../images/excel-ranges-remove-duplicates-after.png)

## <a name="see-also"></a>См. также

- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
- [Работа с ячейками с Excel API JavaScript](excel-add-ins-cells.md)
- [Диапазоны вырезать, скопировать и вклеить с Excel API JavaScript](excel-add-ins-ranges-cut-copy-paste.md)
- [Работа с несколькими диапазонами одновременно в надстройках Excel](excel-add-ins-multiple-ranges.md)
