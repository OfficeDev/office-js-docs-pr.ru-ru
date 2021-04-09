---
title: Установите и получите выбранный диапазон с помощью API JavaScript Excel
description: Узнайте, как использовать API JavaScript Excel для набора и получения диапазонов с помощью API JavaScript Excel.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 06b6219924f0667ecef57d608cb417a76ef8031d
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652882"
---
# <a name="set-and-get-ranges-using-the-excel-javascript-api"></a>Настройка и получения диапазонов с помощью API JavaScript Excel

В этой статье данная статья содержит примеры кода, которые устанавливают и получают диапазоны с API JavaScript Excel. Полный список свойств и методов, поддерживаемых объектом, см. в `Range` [класс Excel.Range.](/javascript/api/excel/excel.range)

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-the-selected-range"></a>Задание выделенного диапазона

В примере кода ниже показано, как выделить диапазон **B2:E6** на активном листе.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:E6");

    range.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="selected-range-b2e6"></a>Выделенный диапазон B2:E6

![Выделенный диапазон в Excel](../images/excel-ranges-set-selection.png)

## <a name="get-the-selected-range"></a>Получение выделенного диапазона

Следующий пример кода получает выбранный диапазон, загружает его `address` свойство и пишет сообщение на консоль.

```js
Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the selected range is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a>См. также

- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
- [Работа с ячейками с помощью API JavaScript Excel](excel-add-ins-cells.md)
- [Установите и получите значения диапазона, текст или формулы с помощью API JavaScript Excel](excel-add-ins-ranges-set-get-values.md)
- [Настройка формата диапазона с помощью API JavaScript Excel](excel-add-ins-ranges-set-format.md)
