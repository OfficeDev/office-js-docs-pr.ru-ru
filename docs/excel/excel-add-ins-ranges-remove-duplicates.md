---
title: Удаление дубликатов с Excel API JavaScript
description: Узнайте, как использовать API Excel JavaScript для удаления дубликатов.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 626676dae57854baf5e8fe195c669a0432eee832c8099715144c2b569ae1055a
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57086139"
---
# <a name="remove-duplicates-using-the-excel-javascript-api"></a>Удаление дубликатов с Excel API JavaScript

В этой статье содержится пример кода, который удаляет дублирующиеся записи в диапазоне с Excel API JavaScript. Полный список свойств и методов, поддерживаемый объектом, см. в `Range` [Excel. Класс Range](/javascript/api/excel/excel.range).

## <a name="remove-rows-with-duplicate-entries"></a>Удаление строк с дублирующими записями

Метод [Range.removeDuplicates](/javascript/api/excel/excel.range#removeDuplicates_columns__includesHeader_) удаляет строки с дублирующимися записями в указанных столбцах. Метод проходит через каждую строку в диапазоне от самого низкого значения индекса до индекса с самым высоким значением в диапазоне (сверху донизу). Строка удаляется, если значение в ее указанном столбце или столбцах уже встречалось в диапазоне. Строки в диапазоне под удаленной строкой сдвигаются вверх. Функция `removeDuplicates` не влияет на положение ячеек вне диапазона.

Функция `removeDuplicates` использует параметр `number[]`, представляющий индексы столбцов, которые проверяются на наличие дубликатов. Этот массив отсчитывается от нуля относительно диапазона, а не листа. Метод также принимает параметр boolean, который указывает, является ли первая строка загонщиком. При значении **true** верхняя строка игнорируется при поиске дубликатов. Метод возвращает объект, который указывает количество удаленных строк и количество `removeDuplicates` `RemoveDuplicatesResult` оставшихся уникальных строк.

При использовании метода диапазона имейте в виду `removeDuplicates` следующее.

- Функция `removeDuplicates` рассматривает значения ячеек, а не результаты функций. Если две разные функции вычисляют одинаковый результат, значения ячеек не считаются повторяющимися.
- Пустые ячейки не игнорируются функцией `removeDuplicates`. Значение пустой ячейки обрабатывается как любое другое значение. Это означает, что пустые строки, содержащиеся в диапазоне, будут включены в объект `RemoveDuplicatesResult`.

В следующем примере кода показано удаление записей с дублирующими значениями в первом столбце.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:D11");

    var deleteResult = range.removeDuplicates([0],true);
    deleteResult.load();

    return context.sync().then(function () {
        console.log(deleteResult.removed + " entries with duplicate names removed.");
        console.log(deleteResult.uniqueRemaining + " entries with unique names remain in the range.");
    });
}).catch(errorHandlerFunction);
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
