---
title: Чтение или написание в неограниченый диапазон с Excel API JavaScript
description: Узнайте, как использовать API Excel JavaScript для чтения или записи в неограниченый диапазон.
ms.date: 04/05/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: f7be2efc3e069ea3451088608ca5255a632ef863
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937762"
---
# <a name="read-or-write-to-an-unbounded-range-using-the-excel-javascript-api"></a>Чтение или написание в неограниченый диапазон с Excel API JavaScript

В этой статье описывается, как читать и писать в диапазоне с Excel API JavaScript. Полный список свойств и методов, поддерживаемый объектом, см. в `Range` [Excel. Класс Range](/javascript/api/excel/excel.range).

Адрес неограниченого диапазона — это адрес диапазона, который указывает целые столбцы или целые строки. Например.

- Адреса диапазона, состоящие из целых столбцов:<ul><li>`C:C`</li><li>`A:F`</li></ul>
- Адреса диапазона, состоящие из целых строк:<ul><li>`2:2`</li><li>`1:4`</li></ul>

## <a name="read-an-unbounded-range"></a>Чтение из неограниченного диапазона

Когда API отправляет запрос на получение неограниченного диапазона (например, `getRange('C:C')`), ответ будет содержать значения `null` для свойств уровня ячейки, например свойств `values`, `text`, `numberFormat` и `formula`. Другие свойства диапазона, например `address` и `cellCount`, будут содержать допустимые значения для неограниченного диапазона.

## <a name="write-to-an-unbounded-range"></a>Запись в неограниченный диапазон

Вы не можете установить свойства уровня ячейки, такие как , и на неограниченый диапазон, так как запрос ввода `values` `numberFormat` слишком `formula` велик. Например, следующий пример кода недостоверный, так как он пытается указать для `values` неограниченого диапазона. API возвращает ошибку, если вы попытаетесь установить свойства уровня ячейки для неограниченого диапазона.

```js
// Note: This code sample attempts to specify `values` for an unbounded range, which is not a valid request. The sample will return an error. 
var range = context.workbook.worksheets.getActiveWorksheet().getRange('A:B');
range.values = 'Due Date';
```

## <a name="see-also"></a>См. также

- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
- [Работа с ячейками с Excel API JavaScript](excel-add-ins-cells.md)
- [Чтение или написание в большом диапазоне с Excel API JavaScript](excel-add-ins-ranges-large.md)
- [Работа с несколькими диапазонами одновременно в надстройках Excel](excel-add-ins-multiple-ranges.md)
