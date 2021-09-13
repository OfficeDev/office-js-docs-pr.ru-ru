---
title: Поиск строки с Excel API JavaScript
description: Узнайте, как найти строку в диапазоне с Excel API JavaScript.
ms.date: 04/02/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: c143acdfb94928b3c59e4fa92eab41ca635f021a
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59153846"
---
# <a name="find-a-string-within-a-range-using-the-excel-javascript-api"></a>Поиск строки в диапазоне с Excel API JavaScript

В этой статье приводится пример кода, который находит строку в диапазоне с Excel API JavaScript. Полный список свойств и методов, поддерживаемый объектом, см. в `Range` [Excel. Класс Range](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="match-a-string-within-a-range"></a>Соответствие строке в диапазоне

У объекта `Range` есть метод `find` для поиска указанной строки в диапазоне. Он возвращает диапазон первой ячейки с текстом, соответствующим критериям.

Приведенный ниже пример кода находит первую ячейку со значением, соответствующим строке **Food** (Еда), и заносит ее адрес в консоль. Обратите внимание, что метод `find` выдает ошибку `ItemNotFound`, если указанной строки не существует в диапазоне. Если ожидается, что указанная строка может отсутствовать в диапазоне, используйте вместо этого метод [findOrNullObject](../develop/application-specific-api-model.md#ornullobject-methods-and-properties), чтобы ваш код корректно обработал этот сценарий.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var table = sheet.tables.getItem("ExpensesTable");
    var searchRange = table.getRange();
    var foundRange = searchRange.find("Food", {
        completeMatch: true, // find will match the whole cell value
        matchCase: false, // find will not match case
        searchDirection: Excel.SearchDirection.forward // find will start searching at the beginning of the range
    });

    foundRange.load("address");
    return context.sync()
        .then(function() {
            console.log(foundRange.address);
    });
}).catch(errorHandlerFunction);
```

Если метод `find` вызывается для диапазона, представляющего одну ячейку, поиск выполняется во всем листе. Поиск начинается в этой ячейке и продолжается в направлении, которое определяется параметром `SearchCriteria.searchDirection`, охватывающим концы листа при необходимости.

## <a name="see-also"></a>См. также

- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
- [Работа с ячейками с Excel API JavaScript](excel-add-ins-cells.md)
- [Поиск специальных ячеек в диапазоне с Excel API JavaScript](excel-add-ins-ranges-special-cells.md)
