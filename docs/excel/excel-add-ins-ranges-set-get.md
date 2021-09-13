---
title: Установите и получите выбранный диапазон с Excel API JavaScript
description: Узнайте, как использовать API Excel JavaScript для набора и получения выбранного диапазона с Excel API JavaScript.
ms.date: 07/02/2021
ms.prod: excel
ms.localizationpriority: medium
ms.openlocfilehash: e6a61e200482d48691c38845598987a1719b457f
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59153850"
---
# <a name="set-and-get-the-selected-range-using-the-excel-javascript-api"></a>Установите и получите выбранный диапазон с Excel API JavaScript

В этой статье данная статья содержит примеры кода, которые устанавливают и получают выбранный диапазон с Excel API JavaScript. Полный список свойств и методов, поддерживаемый объектом, см. в `Range` [Excel. Класс Range](/javascript/api/excel/excel.range).

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

![Выбранный диапазон в Excel.](../images/excel-ranges-set-selection.png)

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

## <a name="select-the-edge-of-a-used-range"></a>Выберите край используемого диапазона

Методы [Range.getRangeEdge](/javascript/api/excel/excel.range#getRangeEdge_direction__activeCell_) и [Range.getExtendedRange](/javascript/api/excel/excel.range#getExtendedRange_directionString__activeCell_) позволяют надстройке реплицировать поведение ярлыков выбора клавиатуры, выбрав край используемого диапазона на основе выбранного диапазона. Дополнительные дополнительные новости об используемых диапазонах см. в [руб. Get used range.](excel-add-ins-ranges-get.md#get-used-range)

На следующем скриншоте используется диапазон таблицы со значениями в каждой ячейке **C5:F12**. Пустые ячейки за пределами этой таблицы находятся за пределами используемого диапазона.

![Таблица с данными C5:F12 в Excel.](../images/excel-ranges-used-range.png)

### <a name="select-the-cell-at-the-edge-of-the-current-used-range"></a>Выберите ячейку на краю текущего используемого диапазона

В следующем примере кода показано, как использовать метод для выбора ячейки на самом дальнем краю используемого диапазона тока `Range.getRangeEdge` в направлении вверх. Это действие соответствует результату использования клавиши Ctrl+Up при выборе диапазона.

```js
Excel.run(function (context) {
    // Get the selected range.
    var range = context.workbook.getSelectedRange();

    // Specify the direction with the `KeyboardDirection` enum.
    var direction = Excel.KeyboardDirection.up;

    // Get the active cell in the workbook.
    var activeCell = context.workbook.getActiveCell();

    // Get the top-most cell of the current used range.
    // This method acts like the Ctrl+Up arrow key keyboard shortcut while a range is selected.
    var rangeEdge = range.getRangeEdge(
      direction,
      activeCell
    );
    rangeEdge.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

#### <a name="before-selecting-the-cell-at-the-edge-of-the-used-range"></a>Перед выбором ячейки на краю используемого диапазона

На следующем скриншоте показан используемый диапазон и выбранный диапазон в используемом диапазоне. Используемый диапазон — это таблица с данными **на C5:F12**. В этой таблице выбирается **диапазон D8:E9.** Этот выбор является *состоянием до* запуска `Range.getRangeEdge` метода.

![Таблица с данными C5:F12 в Excel. Выбран диапазон D8:E9.](../images/excel-ranges-used-range-d8-e9.png)

#### <a name="after-selecting-the-cell-at-the-edge-of-the-used-range"></a>После выбора ячейки на краю используемого диапазона

На следующем скриншоте показана та же таблица, что и на предыдущем скриншоте, с данными в диапазоне **C5:F12**. В этой таблице выбирается **диапазон D5.** Этот выбор после *состояния,* после запуска метода, чтобы выбрать ячейку на краю используемого диапазона `Range.getRangeEdge` в направлении вверх.

![Таблица с данными C5:F12 в Excel. Выбран диапазон D5.](../images/excel-ranges-used-range-d5.png)

### <a name="select-all-cells-from-current-range-to-furthest-edge-of-used-range"></a>Выберите все ячейки от текущего диапазона до дальнего края используемого диапазона

В следующем примере кода показано, как использовать метод для выбора всех ячеек из выбранного диапазона до самого дальнего края используемого диапазона в направлении `Range.getExtendedRange` вниз. Это действие соответствует результату использования клавиши Ctrl+Shift+Down при выборе диапазона.

```js
Excel.run(function (context) {
    // Get the selected range.
    var range = context.workbook.getSelectedRange();

    // Specify the direction with the `KeyboardDirection` enum.
    var direction = Excel.KeyboardDirection.down;

    // Get the active cell in the workbook.
    var activeCell = context.workbook.getActiveCell();

    // Get all the cells from the currently selected range to the bottom-most edge of the used range.
    // This method acts like the Ctrl+Shift+Down arrow key keyboard shortcut while a range is selected.
    var extendedRange = range.getExtendedRange(
      direction,
      activeCell
    );
    extendedRange.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

#### <a name="before-selecting-all-the-cells-from-the-current-range-to-the-edge-of-the-used-range"></a>Перед выбором всех ячеек от текущего диапазона до края используемого диапазона

На следующем скриншоте показан используемый диапазон и выбранный диапазон в используемом диапазоне. Используемый диапазон — это таблица с данными **на C5:F12**. В этой таблице выбирается **диапазон D8:E9.** Этот выбор является *состоянием до* запуска `Range.getExtendedRange` метода.

![Таблица с данными C5:F12 в Excel. Выбран диапазон D8:E9.](../images/excel-ranges-used-range-d8-e9.png)

#### <a name="after-selecting-all-the-cells-from-the-current-range-to-the-edge-of-the-used-range"></a>После выбора всех ячеек от текущего диапазона до края используемого диапазона

На следующем скриншоте показана та же таблица, что и на предыдущем скриншоте, с данными в диапазоне **C5:F12**. В этой таблице выбирается **диапазон D8:E12.** Этот выбор *после* состояния после запуска метода для выбора всех ячеек от текущего диапазона до края используемого диапазона `Range.getExtendedRange` в направлении вниз.

![Таблица с данными C5:F12 в Excel. Выбран диапазон D8:E12.](../images/excel-ranges-used-range-d8-e12.png)

## <a name="see-also"></a>Дополнительные материалы

- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
- [Работа с ячейками с Excel API JavaScript](excel-add-ins-cells.md)
- [Установите и получите значения диапазона, текст или формулы с Excel API JavaScript](excel-add-ins-ranges-set-get-values.md)
- [Настройка формата диапазона с Excel API JavaScript](excel-add-ins-ranges-set-format.md)
