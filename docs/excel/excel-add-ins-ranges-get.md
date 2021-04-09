---
title: Получите диапазон с помощью API JavaScript Excel
description: Узнайте, как получить диапазон с помощью API JavaScript Excel.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 6aa9bb00bc9d24aeee5f1fef9e8d1531525e9d1f
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652927"
---
# <a name="get-a-range-using-the-excel-javascript-api"></a>Получите диапазон с помощью API JavaScript Excel

В этой статье приводится ряд примеров получения диапазона в листах с помощью API JavaScript Excel. Полный список свойств и методов, поддерживаемых объектом, см. в `Range` [класс Excel.Range.](/javascript/api/excel/excel.range)

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="get-range-by-address"></a>Получение диапазона по адресу

Следующий пример кода получает диапазон с адресом **B2:C5** из таблицы с именем **Sample,** загружает ее свойство и пишет сообщение на `address` консоль.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:C5");
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the range B2:C5 is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="get-range-by-name"></a>Получение диапазона по имени

Следующий пример кода получает диапазон, названный из таблицы с именем Sample, загружает его свойство и пишет `MyRange` сообщение на  `address` консоль.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("MyRange");
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the range "MyRange" is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="get-used-range"></a>Получение используемого диапазона

В следующем примере кода используется диапазон от таблицы с именем **Sample,** загружается его свойство и записывает сообщение `address` на консоль. Используемый диапазон — это наименьший диапазон, включающий в себя все ячейки листа, которые содержат значение или форматирование. Если весь лист пустой, метод возвращает диапазон, состоящий только из `getUsedRange()` верхнего левого элемента.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getUsedRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the used range in the worksheet is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="get-entire-range"></a>Получение всего диапазона

Следующий пример кода получает весь диапазон таблицы из таблицы с именем **Sample,** загружает ее свойство и пишет сообщение `address` на консоль.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the entire worksheet range is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a>См. также

- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
- [Работа с ячейками с помощью API JavaScript Excel](excel-add-ins-cells.md)
- [Вставьте диапазон с помощью API JavaScript Excel](excel-add-ins-ranges-insert.md)
