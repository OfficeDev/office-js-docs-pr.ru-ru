---
title: Поиск строки с Excel API JavaScript
description: Узнайте, как найти строку в диапазоне с Excel API JavaScript.
ms.date: 02/17/2022
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: 042465e01af55bbb3f4325ea44edc27174d558f2
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340990"
---
# <a name="find-a-string-within-a-range-using-the-excel-javascript-api"></a>Поиск строки в диапазоне с Excel API JavaScript

В этой статье приводится пример кода, который находит строку в диапазоне с Excel API JavaScript. Полный список свойств `Range` и методов, поддерживаемый объектом, см. в Excel[. Класс Range](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="match-a-string-within-a-range"></a>Соответствие строке в диапазоне

У объекта `Range` есть метод `find` для поиска указанной строки в диапазоне. Он возвращает диапазон первой ячейки с текстом, соответствующим критериям.

Приведенный ниже пример кода находит первую ячейку со значением, соответствующим строке **Food** (Еда), и заносит ее адрес в консоль. Обратите внимание, что метод `find` выдает ошибку `ItemNotFound`, если указанной строки не существует в диапазоне. Если ожидается, что указанная строка может отсутствовать в диапазоне, используйте вместо этого метод [findOrNullObject](../develop/application-specific-api-model.md#ornullobject-methods-and-properties), чтобы ваш код корректно обработал этот сценарий.

```js
await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("Sample");
    let table = sheet.tables.getItem("ExpensesTable");
    let searchRange = table.getRange();
    let foundRange = searchRange.find("Food", {
        completeMatch: true, // Match the whole cell value.
        matchCase: false, // Don't match case.
        searchDirection: Excel.SearchDirection.forward // Start search at the beginning of the range.
    });

    foundRange.load("address");
    await context.sync();

    console.log(foundRange.address);
});
```

Если метод `find` вызывается для диапазона, представляющего одну ячейку, поиск выполняется во всем листе. Поиск начинается в этой ячейке и продолжается в направлении, которое определяется параметром `SearchCriteria.searchDirection`, охватывающим концы листа при необходимости.

## <a name="see-also"></a>См. также

- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
- [Работа с ячейками с Excel API JavaScript](excel-add-ins-cells.md)
- [Поиск специальных ячеек в диапазоне с Excel API JavaScript](excel-add-ins-ranges-special-cells.md)
