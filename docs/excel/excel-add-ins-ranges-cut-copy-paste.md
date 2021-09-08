---
title: Диапазоны вырезать, скопировать и вклеить с Excel API JavaScript
description: Узнайте, как вырезать, скопировать и вклеить диапазоны с Excel API JavaScript.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: a48d726e517899249652d857d9e79d2201f3bfc3
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/08/2021
ms.locfileid: "58937147"
---
# <a name="cut-copy-and-paste-ranges-using-the-excel-javascript-api"></a>Диапазоны вырезать, скопировать и вклеить с Excel API JavaScript

В этой статье данная статья содержит примеры кода, которые вырезали, копируют и вклеили диапазоны с Excel API JavaScript. Полный список свойств и методов, поддерживаемый объектом, см. в `Range` [Excel. Класс Range](/javascript/api/excel/excel.range).

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="copy-and-paste"></a>Copy and paste

Метод [Range.copyFrom](/javascript/api/excel/excel.range#copyFrom_sourceRange__copyType__skipBlanks__transpose_) реплицирует  действия copy и **Paste** Excel пользовательского интерфейса. Назначение — это `Range` объект, `copyFrom` который вызван. Источник для копирования передается как диапазон или адрес строки, представляющий диапазон.

В следующем примере кода копируются данные из **A1:E1** в диапазон, начиная с **G1** (который заканчивается вставкой в **G1:K1**).

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy everything from "A1:E1" into "G1" and the cells afterwards ("G1:K1")
    sheet.getRange("G1").copyFrom("A1:E1");
    return context.sync();
}).catch(errorHandlerFunction);
```

У функции `Range.copyFrom` есть три необязательных параметра.

```TypeScript
copyFrom(sourceRange: Range | RangeAreas | string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean): void;
```

`copyType` указывает, какие данные копируются из источника в назначение.

- `Excel.RangeCopyType.formulas` передает формулы в исходных ячейках и сохраняет относительное расположение диапазонов этих формул. Все записи, не являющиеся формулами, копируются в исходном виде.
- `Excel.RangeCopyType.values` копирует значения данных, а в случае формул — результат формулы.
- `Excel.RangeCopyType.formats` копирует форматирование диапазона, включая шрифт, цвет и другие параметры форматирования, но без значений.
- `Excel.RangeCopyType.all` (параметр по умолчанию) копирует данные и форматирование, сохраняя формулы ячеек при обнаружении.

`skipBlanks` устанавливает, будут ли копироваться пустые ячейки в назначение. Если значение равно true, `copyFrom` пропускает пустые ячейки в диапазоне источника.
Пропущенные ячейки не перезапишут существующие данные в соответствующих им ячейках конечного диапазона. Значение по умолчанию: false.

`transpose` определяет, переставляются ли данные в исходное расположение, то есть переключаются ли строки и столбцы.
Переставленный диапазон переключается на главной диагонали, поэтому строки **1**, **2** и **3** становятся столбцами **A**, **B** и **C**.

В приведенном ниже примере кода и изображениях демонстрируется это поведение в простом сценарии.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy a range, omitting the blank cells so existing data is not overwritten in those cells
    sheet.getRange("D1").copyFrom("A1:C1",
        Excel.RangeCopyType.all,
        true, // skipBlanks
        false); // transpose
    // copy a range, including the blank cells which will overwrite existing data in the target cells
    sheet.getRange("D2").copyFrom("A2:C2",
        Excel.RangeCopyType.all,
        false, // skipBlanks
        false); // transpose
    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-before-range-is-copied-and-pasted"></a>Данные перед копированием и вклейка диапазона

![Данные в Excel перед запуском метода копирования диапазона.](../images/excel-range-copyfrom-skipblanks-before.png)

### <a name="data-after-range-is-copied-and-pasted"></a>Данные после копирования и вклейки данных после диапазона

![Данные в Excel после запуска метода копирования диапазона.](../images/excel-range-copyfrom-skipblanks-after.png)

## <a name="cut-and-paste-move-cells"></a>Вырезать и вклеить (переместить) ячейки

Метод [Range.moveTo](/javascript/api/excel/excel.range#moveTo_destinationRange_) перемещает ячейки в новое расположение в книге. Это поведение движения клеток работает так [](https://support.microsoft.com/office/803d65eb-6a3e-4534-8c6f-ff12d1c4139e) же, как при перемещении ячеек путем перетаскивание границы диапазона или при принятии действий **Cut** и **Paste.** Форматирование и значения диапазона перемещаются в указанное в качестве параметра `destinationRange` расположение.

Следующий пример кода перемещает диапазон с помощью `Range.moveTo` метода. Обратите внимание, что если диапазон назначения меньше источника, он будет расширен, чтобы охватить исходный контент.

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.getRange("F1").values = [["Moved Range"]];

    // Move the cells "A1:E1" to "G1" (which fills the range "G1:K1").
    sheet.getRange("A1:E1").moveTo("G1");
    return context.sync();
});
```

## <a name="see-also"></a>См. также

- [Объектная модель JavaScript для Excel в надстройках Office](excel-add-ins-core-concepts.md)
- [Работа с ячейками с Excel API JavaScript](excel-add-ins-cells.md)
- [Удаление дубликатов с Excel API JavaScript](excel-add-ins-ranges-remove-duplicates.md)
- [Работа с несколькими диапазонами одновременно в надстройках Excel](excel-add-ins-multiple-ranges.md)
