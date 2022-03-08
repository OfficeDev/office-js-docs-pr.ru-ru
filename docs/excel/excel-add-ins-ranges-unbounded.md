---
title: Чтение или написание в неограниченый диапазон с Excel API JavaScript
description: Узнайте, как использовать API Excel JavaScript для чтения или записи в неограниченый диапазон.
ms.date: 02/17/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 6e9b0c56dfd04cd53e01c41fea23fbf826a6fa14
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340955"
---
# <a name="read-or-write-to-an-unbounded-range-using-the-excel-javascript-api"></a>Чтение или написание в неограниченый диапазон с Excel API JavaScript

В этой статье описывается, как читать и писать в диапазоне с Excel API JavaScript. Полный список свойств `Range` и методов, поддерживаемый объектом, см. в Excel[. Класс Range](/javascript/api/excel/excel.range).

Адрес неограниченого диапазона — это адрес диапазона, который указывает целые столбцы или целые строки. Пример:

- Адреса диапазона, состоящие из целых столбцов.
  - `C:C`
  - `A:F`
- Адреса диапазона, состоящие из целых строк.
  - `2:2`
  - `1:4`

## <a name="read-an-unbounded-range"></a>Чтение из неограниченного диапазона

Когда API отправляет запрос на получение неограниченного диапазона (например, `getRange('C:C')`), ответ будет содержать значения `null` для свойств уровня ячейки, например свойств `values`, `text`, `numberFormat` и `formula`. Другие свойства диапазона, например `address` и `cellCount`, будут содержать допустимые значения для неограниченного диапазона.

## <a name="write-to-an-unbounded-range"></a>Запись в неограниченный диапазон

Вы не можете установить свойства `values`уровня ячейки, такие как , и `formula` `numberFormat`на неограниченый диапазон, так как запрос ввода слишком велик. Например, следующий пример кода недостоверный, так как он пытается `values` указать для неограниченого диапазона. API возвращает ошибку, если вы попытаетесь установить свойства уровня ячейки для неограниченого диапазона.

```js
// Note: This code sample attempts to specify `values` for an unbounded range, which is not a valid request. The sample will return an error. 
let range = context.workbook.worksheets.getActiveWorksheet().getRange('A:B');
range.values = 'Due Date';
```

## <a name="see-also"></a>См. также

- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
- [Работа с ячейками с Excel API JavaScript](excel-add-ins-cells.md)
- [Чтение или написание в большом диапазоне с Excel API JavaScript](excel-add-ins-ranges-large.md)
- [Работа с несколькими диапазонами одновременно в надстройках Excel](excel-add-ins-multiple-ranges.md)
