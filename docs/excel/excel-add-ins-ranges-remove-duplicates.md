---
title: Удаление дубликатов с помощью API JavaScript для Excel
description: Узнайте, как использовать API JavaScript для Excel для удаления дубликатов.
ms.date: 02/17/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 9ece7c9f35b341dbb8d0d90e8ca4bda5215580ed
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889144"
---
# <a name="remove-duplicates-using-the-excel-javascript-api"></a>Удаление дубликатов с помощью API JavaScript для Excel

В этой статье приведен пример кода, который удаляет повторяющиеся записи в диапазоне с помощью API JavaScript для Excel. Полный список свойств и методов `Range` , поддерживаемых объектом, см. в разделе ["Класс Excel.Range"](/javascript/api/excel/excel.range).

## <a name="remove-rows-with-duplicate-entries"></a>Удаление строк с повторяющимися записями

Метод [Range.removeDuplicates](/javascript/api/excel/excel.range#excel-excel-range-removeduplicates-member(1)) удаляет строки с повторяющимися записями в указанных столбцах. Метод проходит через каждую строку в диапазоне от индекса с наименьшим значением до индекса с наибольшим значением в диапазоне (сверху вниз). Строка удаляется, если значение в ее указанном столбце или столбцах уже встречалось в диапазоне. Строки в диапазоне под удаленной строкой сдвигаются вверх. Функция `removeDuplicates` не влияет на положение ячеек вне диапазона.

Функция `removeDuplicates` использует параметр `number[]`, представляющий индексы столбцов, которые проверяются на наличие дубликатов. Этот массив отсчитывается от нуля относительно диапазона, а не листа. Метод также принимает логический параметр, указывающий, является ли первая строка заголовком. Если `true`верхняя строка игнорируется при рассмотрении дубликатов. Метод `removeDuplicates` возвращает объект, `RemoveDuplicatesResult` указывающий количество удаленных строк и количество оставшихся уникальных строк.

При использовании метода диапазона `removeDuplicates` учитывайте следующее.

- Функция `removeDuplicates` рассматривает значения ячеек, а не результаты функций. Если две разные функции вычисляют одинаковый результат, значения ячеек не считаются повторяющимися.
- Пустые ячейки не игнорируются функцией `removeDuplicates`. Значение пустой ячейки обрабатывается как любое другое значение. Это означает, что пустые строки, содержащиеся в диапазоне, будут включены в объект `RemoveDuplicatesResult`.

В следующем примере кода показано удаление записей с повторяющимися значениями в первом столбце.

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

### <a name="data-before-duplicate-entries-are-removed"></a>Данные перед удалением повторяющихся записей

![Данные в Excel перед запуском метода удаления дубликатов диапазона.](../images/excel-ranges-remove-duplicates-before.png)

### <a name="data-after-duplicate-entries-are-removed"></a>Данные после удаления повторяющихся записей

![Данные в Excel после выполнения метода удаления дубликатов диапазона.](../images/excel-ranges-remove-duplicates-after.png)

## <a name="see-also"></a>Дополнительные ресурсы

- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
- [Работа с ячейками с помощью API JavaScript для Excel](excel-add-ins-cells.md)
- [Вырезание, копирование и вставка диапазонов с помощью API JavaScript для Excel](excel-add-ins-ranges-cut-copy-paste.md)
- [Работа с несколькими диапазонами одновременно в надстройках Excel](excel-add-ins-multiple-ranges.md)
