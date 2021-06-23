---
title: Диапазоны вырезать, скопировать и вклеить с Excel API JavaScript
description: Узнайте, как вырезать, скопировать и вклеить диапазоны с Excel API JavaScript.
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 2112702110b72e0020ed72090ce495abb3ff5366
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075826"
---
# <a name="cut-copy-and-paste-ranges-using-the-excel-javascript-api"></a><span data-ttu-id="c6d85-103">Диапазоны вырезать, скопировать и вклеить с Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="c6d85-103">Cut, copy, and paste ranges using the Excel JavaScript API</span></span>

<span data-ttu-id="c6d85-104">В этой статье данная статья содержит примеры кода, которые вырезали, копируют и вклеили диапазоны с Excel API JavaScript.</span><span class="sxs-lookup"><span data-stu-id="c6d85-104">This article provides code samples that cut, copy, and paste ranges using the Excel JavaScript API.</span></span> <span data-ttu-id="c6d85-105">Полный список свойств и методов, поддерживаемый объектом, см. в `Range` [Excel. Класс Range](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="c6d85-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="copy-and-paste"></a><span data-ttu-id="c6d85-106">Copy and paste</span><span class="sxs-lookup"><span data-stu-id="c6d85-106">Copy and paste</span></span>

<span data-ttu-id="c6d85-107">Метод [Range.copyFrom](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-) реплицирует  действия copy и **Paste** Excel пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="c6d85-107">The [Range.copyFrom](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-) method replicates the **Copy** and **Paste** actions of the Excel UI.</span></span> <span data-ttu-id="c6d85-108">Назначение — это `Range` объект, `copyFrom` который вызван.</span><span class="sxs-lookup"><span data-stu-id="c6d85-108">The destination is the `Range` object that `copyFrom` is called on.</span></span> <span data-ttu-id="c6d85-109">Источник для копирования передается как диапазон или адрес строки, представляющий диапазон.</span><span class="sxs-lookup"><span data-stu-id="c6d85-109">The source to be copied is passed as a range or a string address representing a range.</span></span>

<span data-ttu-id="c6d85-110">В следующем примере кода копируются данные из **A1:E1** в диапазон, начиная с **G1** (который заканчивается вставкой в **G1:K1**).</span><span class="sxs-lookup"><span data-stu-id="c6d85-110">The following code sample copies the data from **A1:E1** into the range starting at **G1** (which ends up pasting into **G1:K1**).</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy everything from "A1:E1" into "G1" and the cells afterwards ("G1:K1")
    sheet.getRange("G1").copyFrom("A1:E1");
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="c6d85-111">У функции `Range.copyFrom` есть три необязательных параметра.</span><span class="sxs-lookup"><span data-stu-id="c6d85-111">`Range.copyFrom` has three optional parameters.</span></span>

```TypeScript
copyFrom(sourceRange: Range | RangeAreas | string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean): void;
```

<span data-ttu-id="c6d85-112">`copyType` указывает, какие данные копируются из источника в назначение.</span><span class="sxs-lookup"><span data-stu-id="c6d85-112">`copyType` specifies what data gets copied from the source to the destination.</span></span>

- <span data-ttu-id="c6d85-113">`Excel.RangeCopyType.formulas` передает формулы в исходных ячейках и сохраняет относительное расположение диапазонов этих формул.</span><span class="sxs-lookup"><span data-stu-id="c6d85-113">`Excel.RangeCopyType.formulas` transfers the formulas in the source cells and preserves the relative positioning of those formulas' ranges.</span></span> <span data-ttu-id="c6d85-114">Все записи, не являющиеся формулами, копируются в исходном виде.</span><span class="sxs-lookup"><span data-stu-id="c6d85-114">Any non-formula entries are copied as-is.</span></span>
- <span data-ttu-id="c6d85-115">`Excel.RangeCopyType.values` копирует значения данных, а в случае формул — результат формулы.</span><span class="sxs-lookup"><span data-stu-id="c6d85-115">`Excel.RangeCopyType.values` copies the data values and, in the case of formulas, the result of the formula.</span></span>
- <span data-ttu-id="c6d85-116">`Excel.RangeCopyType.formats` копирует форматирование диапазона, включая шрифт, цвет и другие параметры форматирования, но без значений.</span><span class="sxs-lookup"><span data-stu-id="c6d85-116">`Excel.RangeCopyType.formats` copies the formatting of the range, including font, color, and other format settings, but no values.</span></span>
- <span data-ttu-id="c6d85-117">`Excel.RangeCopyType.all` (параметр по умолчанию) копирует данные и форматирование, сохраняя формулы ячеек при обнаружении.</span><span class="sxs-lookup"><span data-stu-id="c6d85-117">`Excel.RangeCopyType.all` (the default option) copies both data and formatting, preserving cells' formulas if found.</span></span>

<span data-ttu-id="c6d85-118">`skipBlanks` устанавливает, будут ли копироваться пустые ячейки в назначение.</span><span class="sxs-lookup"><span data-stu-id="c6d85-118">`skipBlanks` sets whether blank cells are copied into the destination.</span></span> <span data-ttu-id="c6d85-119">Если значение равно true, `copyFrom` пропускает пустые ячейки в диапазоне источника.</span><span class="sxs-lookup"><span data-stu-id="c6d85-119">When true, `copyFrom` skips blank cells in the source range.</span></span>
<span data-ttu-id="c6d85-120">Пропущенные ячейки не перезапишут существующие данные в соответствующих им ячейках конечного диапазона.</span><span class="sxs-lookup"><span data-stu-id="c6d85-120">Skipped cells will not overwrite the existing data of their corresponding cells in the destination range.</span></span> <span data-ttu-id="c6d85-121">Значение по умолчанию: false.</span><span class="sxs-lookup"><span data-stu-id="c6d85-121">The default is false.</span></span>

<span data-ttu-id="c6d85-122">`transpose` определяет, переставляются ли данные в исходное расположение, то есть переключаются ли строки и столбцы.</span><span class="sxs-lookup"><span data-stu-id="c6d85-122">`transpose` determines whether or not the data is transposed, meaning its rows and columns are switched, into the source location.</span></span>
<span data-ttu-id="c6d85-123">Переставленный диапазон переключается на главной диагонали, поэтому строки **1**, **2** и **3** становятся столбцами **A**, **B** и **C**.</span><span class="sxs-lookup"><span data-stu-id="c6d85-123">A transposed range is flipped along the main diagonal, so rows **1**, **2**, and **3** will become columns **A**, **B**, and **C**.</span></span>

<span data-ttu-id="c6d85-124">В приведенном ниже примере кода и изображениях демонстрируется это поведение в простом сценарии.</span><span class="sxs-lookup"><span data-stu-id="c6d85-124">The following code sample and images demonstrate this behavior in a simple scenario.</span></span>

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

### <a name="data-before-range-is-copied-and-pasted"></a><span data-ttu-id="c6d85-125">Данные перед копированием и вклейка диапазона</span><span class="sxs-lookup"><span data-stu-id="c6d85-125">Data before range is copied and pasted</span></span>

![Данные в Excel перед запуском метода копирования диапазона.](../images/excel-range-copyfrom-skipblanks-before.png)

### <a name="data-after-range-is-copied-and-pasted"></a><span data-ttu-id="c6d85-127">Данные после копирования и вклейки данных после диапазона</span><span class="sxs-lookup"><span data-stu-id="c6d85-127">Data after range is copied and pasted</span></span>

![Данные в Excel после запуска метода копирования диапазона.](../images/excel-range-copyfrom-skipblanks-after.png)

## <a name="cut-and-paste-move-cells"></a><span data-ttu-id="c6d85-129">Вырезать и вклеить (переместить) ячейки</span><span class="sxs-lookup"><span data-stu-id="c6d85-129">Cut and paste (move) cells</span></span>

<span data-ttu-id="c6d85-130">Метод [Range.moveTo](/javascript/api/excel/excel.range#moveto-destinationrange-) перемещает ячейки в новое расположение в книге.</span><span class="sxs-lookup"><span data-stu-id="c6d85-130">The [Range.moveTo](/javascript/api/excel/excel.range#moveto-destinationrange-) method moves cells to a new location in the workbook.</span></span> <span data-ttu-id="c6d85-131">Это поведение движения клеток работает так [](https://support.office.com/article/Move-or-copy-cells-and-cell-contents-803d65eb-6a3e-4534-8c6f-ff12d1c4139e) же, как при перемещении ячеек путем перетаскивание границы диапазона или при принятии действий **Cut** и **Paste.**</span><span class="sxs-lookup"><span data-stu-id="c6d85-131">This cell movement behavior works the same as when cells are moved by [dragging the range border](https://support.office.com/article/Move-or-copy-cells-and-cell-contents-803d65eb-6a3e-4534-8c6f-ff12d1c4139e) or when taking the **Cut** and **Paste** actions.</span></span> <span data-ttu-id="c6d85-132">Форматирование и значения диапазона перемещаются в указанное в качестве параметра `destinationRange` расположение.</span><span class="sxs-lookup"><span data-stu-id="c6d85-132">Both the formatting and values of the range are moved to the location specified as the `destinationRange` parameter.</span></span>

<span data-ttu-id="c6d85-133">Следующий пример кода перемещает диапазон с помощью `Range.moveTo` метода.</span><span class="sxs-lookup"><span data-stu-id="c6d85-133">The following code sample moves a range with the `Range.moveTo` method.</span></span> <span data-ttu-id="c6d85-134">Обратите внимание, что если диапазон назначения меньше источника, он будет расширен, чтобы охватить исходный контент.</span><span class="sxs-lookup"><span data-stu-id="c6d85-134">Note that if the destination range is smaller than the source, it will be expanded to encompass the source content.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.getRange("F1").values = [["Moved Range"]];

    // Move the cells "A1:E1" to "G1" (which fills the range "G1:K1").
    sheet.getRange("A1:E1").moveTo("G1");
    return context.sync();
});
```

## <a name="see-also"></a><span data-ttu-id="c6d85-135">См. также</span><span class="sxs-lookup"><span data-stu-id="c6d85-135">See also</span></span>

- [<span data-ttu-id="c6d85-136">Объектная модель JavaScript для Excel в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="c6d85-136">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="c6d85-137">Работа с ячейками с Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="c6d85-137">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="c6d85-138">Удаление дубликатов с Excel API JavaScript</span><span class="sxs-lookup"><span data-stu-id="c6d85-138">Remove duplicates using the Excel JavaScript API</span></span>](excel-add-ins-ranges-remove-duplicates.md)
- [<span data-ttu-id="c6d85-139">Работа с несколькими диапазонами одновременно в надстройках Excel</span><span class="sxs-lookup"><span data-stu-id="c6d85-139">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
