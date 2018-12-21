---
title: Работа с диапазонами с использованием API JavaScript для Excel (дополнительные задачи)
description: ''
ms.date: 12/18/2018
ms.openlocfilehash: 6d3da1e7eff4e61ae1b88213d0b432581d8f6a8a
ms.sourcegitcommit: 6870f0d96ed3da2da5a08652006c077a72d811b6
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/21/2018
ms.locfileid: "27383241"
---
# <a name="work-with-ranges-using-the-excel-javascript-api-advanced"></a><span data-ttu-id="2bf45-102">Работа с диапазонами с использованием API JavaScript для Excel (дополнительные задачи)</span><span class="sxs-lookup"><span data-stu-id="2bf45-102">Work with ranges using the Excel JavaScript API (advanced)</span></span>

<span data-ttu-id="2bf45-103">Эта статья основана на сведениях из статьи [Работа с диапазонами с использованием API JavaScript для Excel (основные задачи)](excel-add-ins-ranges.md) с предоставлением примеров кода, демонстрирующих способы выполнения более сложных задач с диапазонами с использованием API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="2bf45-103">This article builds upon information in [Work with ranges using the Excel JavaScript API (fundamental)](excel-add-ins-ranges.md) by providing code samples that show how to perform more advanced tasks with ranges using the Excel JavaScript API.</span></span> <span data-ttu-id="2bf45-104">Полный список свойств и методов, поддерживаемых объектом **Range**, см. в статье [Объект Range (API JavaScript для Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="2bf45-104">For the complete list of properties and methods that the **Range** object supports, see [Range Object (JavaScript API for Excel)](https://docs.microsoft.com/javascript/api/excel/excel.range).</span></span>

## <a name="work-with-dates-using-the-moment-msdate-plug-in"></a><span data-ttu-id="2bf45-105">Работа с датами с использованием подключаемого модуля Moment-MSDate</span><span class="sxs-lookup"><span data-stu-id="2bf45-105">Work with dates using the Moment-MSDate plug-in</span></span>

<span data-ttu-id="2bf45-106">[Библиотека JavaScript Moment](https://momentjs.com/) предоставляет удобный способ использования дат и меток времени.</span><span class="sxs-lookup"><span data-stu-id="2bf45-106">The [Moment JavaScript library](https://momentjs.com/) provides a convenient way to use dates and timestamps.</span></span> <span data-ttu-id="2bf45-107">[Подключаемый модуль Moment-MSDate](https://www.npmjs.com/package/moment-msdate) преобразует формат моментов времени в предпочитаемый для Excel.</span><span class="sxs-lookup"><span data-stu-id="2bf45-107">The [Moment-MSDate plug-in](https://www.npmjs.com/package/moment-msdate) converts the format of moments into one preferable for Excel.</span></span> <span data-ttu-id="2bf45-108">Это тот же формат, который возвращает [функция ТДАТА](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46).</span><span class="sxs-lookup"><span data-stu-id="2bf45-108">This is the same format the [NOW function](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) returns.</span></span>

<span data-ttu-id="2bf45-109">В приведенном ниже коде показано, как установить для диапазона в **B4** метку момента времени.</span><span class="sxs-lookup"><span data-stu-id="2bf45-109">The following code shows how to set the range at **B4** to a moment's timestamp:</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var now = Date.now();
    var nowMoment = moment(now);
    var nowMS = nowMoment.toOADate();

    var dateRange = sheet.getRange("B4");
    dateRange.values = [[nowMS]];

    dateRange.numberFormat = [["[$-409]m/d/yy h:mm AM/PM;@"]];

    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="2bf45-110">Это похоже на способ получения даты из ячейки и ее преобразования в формат момента времени или другой формат, как показано в приведенном ниже коде:</span><span class="sxs-lookup"><span data-stu-id="2bf45-110">It is a similar technique to get the date back out of the cell and convert it to a moment or other format, as demonstrated in the following code:</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    var dateRange = sheet.getRange("B4");
    dateRange.load("values");

    return context.sync().then(function () {
        var nowMS = dateRange.values[0][0];

        // log the date as a moment
        var nowMoment = moment.fromOADate(nowMS);
        console.log(`get (moment): ${JSON.stringify(nowMoment)}`);

        // log the date as a UNIX-style timestamp 
        var now = nowMoment.unix();
        console.log(`get (timestamp): ${now}`);
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="2bf45-111">Вашей надстройке потребуется отформатировать диапазоны, чтобы отобразить даты в более понятной для человека форме.</span><span class="sxs-lookup"><span data-stu-id="2bf45-111">Your add-in will have to format the ranges to display the dates in a more human-readable form.</span></span> <span data-ttu-id="2bf45-112">В примере `"[$-409]m/d/yy h:mm AM/PM;@"` время отобразится как "12/3/18 3:57 PM".</span><span class="sxs-lookup"><span data-stu-id="2bf45-112">The example of `"[$-409]m/d/yy h:mm AM/PM;@"` displays a time like "12/3/18 3:57 PM".</span></span> <span data-ttu-id="2bf45-113">Дополнительные сведения о форматах чисел даты и времени см. в разделе "Рекомендации по форматам даты и времени" статьи [Рекомендации по настройке числовых форматов](https://support.office.com/article/review-guidelines-for-customizing-a-number-format-c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5).</span><span class="sxs-lookup"><span data-stu-id="2bf45-113">For more information about date and time number formats, please see the "Guidelines for date and time formats" in the [Review guidelines for customizing a number format](https://support.office.com/article/review-guidelines-for-customizing-a-number-format-c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5) article.</span></span>

## <a name="work-with-multiple-ranges-simultaneously-preview"></a><span data-ttu-id="2bf45-114">Работа с несколькими диапазонами одновременно (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="2bf45-114">Work with multiple ranges simultaneously in Excel add-ins</span></span>

> [!NOTE]
> <span data-ttu-id="2bf45-115">Объект `RangeAreas` в настоящее время доступен только в общедоступной предварительной версии (бета-версии).</span><span class="sxs-lookup"><span data-stu-id="2bf45-115">The `RangeAreas` function is currently available only in public preview (beta).</span></span> <span data-ttu-id="2bf45-116">Для применения этой функции необходимо использовать бета-версию библиотеки в CDN Office.js: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span><span class="sxs-lookup"><span data-stu-id="2bf45-116">To use this feature, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span>
> <span data-ttu-id="2bf45-117">Если вы используете TypeScript или ваш редактор кода использует файлы определения типа TypeScript для IntelliSense, воспользуйтесь https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span><span class="sxs-lookup"><span data-stu-id="2bf45-117">If you are using TypeScript or your code editor uses TypeScript type definition files for IntelliSense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span></span>

<span data-ttu-id="2bf45-118">Объект `RangeAreas` позволяет вашей надстройке выполнять операции над несколькими диапазонами одновременно.</span><span class="sxs-lookup"><span data-stu-id="2bf45-118">The `RangeAreas` object lets your add-in perform operations on multiple ranges at once.</span></span> <span data-ttu-id="2bf45-119">Эти диапазоны могут быть смежными, но это необязательно.</span><span class="sxs-lookup"><span data-stu-id="2bf45-119">These ranges may be contiguous, but do not have to be.</span></span> <span data-ttu-id="2bf45-120">Объект `RangeAreas` подробнее рассматривается в статье [Работа с несколькими диапазонами одновременно в надстройках Excel](excel-add-ins-multiple-ranges.md).</span><span class="sxs-lookup"><span data-stu-id="2bf45-120">`RangeAreas` are further discussed in the article [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

## <a name="copy-and-paste-preview"></a><span data-ttu-id="2bf45-121">Копирование и вставка (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="2bf45-121">Copy and paste</span></span>

> [!NOTE]
> <span data-ttu-id="2bf45-122">Функция `Range.copyFrom` в настоящее время доступна только в общедоступной предварительной версии (бета-версии).</span><span class="sxs-lookup"><span data-stu-id="2bf45-122">The `Range.copyFrom` function is currently available only in public preview (beta).</span></span> <span data-ttu-id="2bf45-123">Для применения этой функции необходимо использовать бета-версию библиотеки в CDN Office.js: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span><span class="sxs-lookup"><span data-stu-id="2bf45-123">To use this feature, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span>
> <span data-ttu-id="2bf45-124">Если вы используете TypeScript или ваш редактор кода использует файлы определения типа TypeScript для IntelliSense, воспользуйтесь https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span><span class="sxs-lookup"><span data-stu-id="2bf45-124">If you are using TypeScript or your code editor uses TypeScript type definition files for IntelliSense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span></span>

<span data-ttu-id="2bf45-125">Функция `copyFrom` диапазона реплицирует поведение копирования и вставки пользовательского интерфейса Excel.</span><span class="sxs-lookup"><span data-stu-id="2bf45-125">Range’s `copyFrom` function replicates the copy-and-paste behavior of the Excel UI.</span></span> <span data-ttu-id="2bf45-126">Диапазон объекта, который вызывается `copyFrom`, является назначением.</span><span class="sxs-lookup"><span data-stu-id="2bf45-126">The range object that `copyFrom` is called on is the destination.</span></span>
<span data-ttu-id="2bf45-127">Источник для копирования передается как диапазон или адрес строки, представляющий диапазон.</span><span class="sxs-lookup"><span data-stu-id="2bf45-127">The source to be copied is passed as a range or a string address representing a range.</span></span> <span data-ttu-id="2bf45-128">В следующем примере кода копируются данные из **A1:E1** в диапазон, начиная с **G1** (который заканчивается вставкой в **G1:K1**).</span><span class="sxs-lookup"><span data-stu-id="2bf45-128">The following code sample copies the data from **A1:E1** into the range starting at **G1** (which ends up pasting into **G1:K1**).</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy a range starting at a single cell destination
    sheet.getRange("G1").copyFrom("A1:E1");
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="2bf45-129">У функции `Range.copyFrom` есть три необязательных параметра.</span><span class="sxs-lookup"><span data-stu-id="2bf45-129">`Range.copyFrom` has three optional parameters.</span></span>

```TypeScript
copyFrom(sourceRange: Range | string, copyType?: "All" | "Formulas" | "Values" | "Formats", skipBlanks?: boolean, transpose?: boolean): void;
```

<span data-ttu-id="2bf45-130">`copyType` указывает, какие данные копируются из источника в назначение.</span><span class="sxs-lookup"><span data-stu-id="2bf45-130">`copyType` specifies what data gets copied from the source to the destination.</span></span>

- <span data-ttu-id="2bf45-131">`"Formulas"` переносит формулы в ячейках источника и сохраняет относительное положение диапазонов этих формул.</span><span class="sxs-lookup"><span data-stu-id="2bf45-131">`"Formulas"` transfers the formulas in the source cells and preserves the relative positioning of those formulas’ ranges.</span></span> <span data-ttu-id="2bf45-132">Все записи, не являющиеся формулами, копируются в исходном виде.</span><span class="sxs-lookup"><span data-stu-id="2bf45-132">Any non-formula entries are copied as-is.</span></span>
- <span data-ttu-id="2bf45-133">`"Values"` копирует значения данных, а в случае формул — результат формулы.</span><span class="sxs-lookup"><span data-stu-id="2bf45-133">`"Values"` copies the data values and, in the case of formulas, the result of the formula.</span></span>
- <span data-ttu-id="2bf45-134">`"Formats"` копирует форматирование диапазона, включая шрифт, цвет и другие параметры форматирования, но без значений.</span><span class="sxs-lookup"><span data-stu-id="2bf45-134">`"Formats"` copies the formatting of the range, including font, color, and other format settings, but no values.</span></span>
- <span data-ttu-id="2bf45-135">`"All"` (вариант по умолчанию) копирует данные и форматирование, сохраняя формулы ячеек при их обнаружении.</span><span class="sxs-lookup"><span data-stu-id="2bf45-135">`"All"` (the default option) copies both data and formatting, preserving cells’ formulas if found.</span></span>

<span data-ttu-id="2bf45-136">`skipBlanks` устанавливает, будут ли копироваться пустые ячейки в назначение.</span><span class="sxs-lookup"><span data-stu-id="2bf45-136">`skipBlanks` sets whether blank cells are copied into the destination.</span></span> <span data-ttu-id="2bf45-137">Если значение равно true, `copyFrom` пропускает пустые ячейки в диапазоне источника.</span><span class="sxs-lookup"><span data-stu-id="2bf45-137">When true, `copyFrom` skips blank cells in the source range.</span></span>
<span data-ttu-id="2bf45-138">Пропущенные ячейки не перезапишут существующие данные в соответствующих им ячейках конечного диапазона.</span><span class="sxs-lookup"><span data-stu-id="2bf45-138">Skipped cells will not overwrite the existing data of their corresponding cells in the destination range.</span></span> <span data-ttu-id="2bf45-139">Значение по умолчанию: false.</span><span class="sxs-lookup"><span data-stu-id="2bf45-139">The default is false.</span></span>

<span data-ttu-id="2bf45-140">`transpose` определяет, переставляются ли данные в исходное расположение, то есть переключаются ли строки и столбцы.</span><span class="sxs-lookup"><span data-stu-id="2bf45-140">`transpose` determines whether or not the data is transposed, meaning its rows and columns are switched, into the source location.</span></span>
<span data-ttu-id="2bf45-141">Переставленный диапазон переключается на главной диагонали, поэтому строки **1**, **2** и **3** становятся столбцами **A**, **B** и **C**.</span><span class="sxs-lookup"><span data-stu-id="2bf45-141">A transposed range is flipped along the main diagonal, so rows **1**, **2**, and **3** will become columns **A**, **B**, and **C**.</span></span>

<span data-ttu-id="2bf45-142">В приведенном ниже примере кода и изображениях демонстрируется это поведение в простом сценарии.</span><span class="sxs-lookup"><span data-stu-id="2bf45-142">The following code sample and images demonstrate this behavior in a simple scenario.</span></span>

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

<span data-ttu-id="2bf45-143">*Прежде чем предыдущая функция была запущена.*</span><span class="sxs-lookup"><span data-stu-id="2bf45-143">*Before the preceding function has been run.*</span></span>

![Данные в Excel перед запуском метода копирования диапазона](../images/excel-range-copyfrom-skipblanks-before.png)

<span data-ttu-id="2bf45-145">*После запуска предыдущей функции.*</span><span class="sxs-lookup"><span data-stu-id="2bf45-145">*After the preceding function has been run.*</span></span>

![Данные в Excel после запуска метода копирования диапазона](../images/excel-range-copyfrom-skipblanks-after.png)

## <a name="remove-duplicates-preview"></a><span data-ttu-id="2bf45-147">Удаление дубликатов (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="2bf45-147">Remove duplicates (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="2bf45-148">Функция `removeDuplicates` объекта Range в настоящее время доступна только в общедоступной предварительной версии (бета-версии).</span><span class="sxs-lookup"><span data-stu-id="2bf45-148">The Range object's `removeDuplicates` function is currently available only in public preview (beta).</span></span> <span data-ttu-id="2bf45-149">Для применения этой функции необходимо использовать бета-версию библиотеки в CDN Office.js: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span><span class="sxs-lookup"><span data-stu-id="2bf45-149">To use this feature, you must use the beta library of the Office.js CDN: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span>
> <span data-ttu-id="2bf45-150">Если вы используете TypeScript или ваш редактор кода использует файлы определения типа TypeScript для IntelliSense, воспользуйтесь https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span><span class="sxs-lookup"><span data-stu-id="2bf45-150">If you are using TypeScript or your code editor uses TypeScript type definition files for IntelliSense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span></span>

<span data-ttu-id="2bf45-151">Функция `removeDuplicates` объекта Range удаляет строки с повторяющимися записями в указанных столбцах.</span><span class="sxs-lookup"><span data-stu-id="2bf45-151">The Range object's `removeDuplicates` function removes rows with duplicate entries in the specified columns.</span></span> <span data-ttu-id="2bf45-152">Функция проверяет каждую строку в диапазоне от индекса с наименьшим значением до индекса с наибольшим значением (сверху вниз).</span><span class="sxs-lookup"><span data-stu-id="2bf45-152">The function goes through each row in the range from the lowest-valued index to the highest-valued index in the range (from top to bottom).</span></span> <span data-ttu-id="2bf45-153">Строка удаляется, если значение в ее указанном столбце или столбцах уже встречалось в диапазоне.</span><span class="sxs-lookup"><span data-stu-id="2bf45-153">A row is deleted if a value in its specified column or columns appeared earlier in the range.</span></span> <span data-ttu-id="2bf45-154">Строки в диапазоне под удаленной строкой сдвигаются вверх.</span><span class="sxs-lookup"><span data-stu-id="2bf45-154">Rows in the range below the deleted row are shifted up.</span></span> <span data-ttu-id="2bf45-155">Функция `removeDuplicates` не влияет на положение ячеек вне диапазона.</span><span class="sxs-lookup"><span data-stu-id="2bf45-155">`removeDuplicates` does not affect the position of cells outside of the range.</span></span>

<span data-ttu-id="2bf45-156">Функция `removeDuplicates` использует параметр `number[]`, представляющий индексы столбцов, которые проверяются на наличие дубликатов.</span><span class="sxs-lookup"><span data-stu-id="2bf45-156">`removeDuplicates` takes in a `number[]` representing the column indices which are checked for duplicates.</span></span> <span data-ttu-id="2bf45-157">Этот массив отсчитывается от нуля относительно диапазона, а не листа.</span><span class="sxs-lookup"><span data-stu-id="2bf45-157">This array is zero-based and relative to the range, not the worksheet.</span></span> <span data-ttu-id="2bf45-158">Функция также использует логический параметр, который определяет, является ли первая строка заголовком.</span><span class="sxs-lookup"><span data-stu-id="2bf45-158">The function also takes in a boolean parameter that specifies whether the first row is a header.</span></span> <span data-ttu-id="2bf45-159">При значении **true** верхняя строка игнорируется при поиске дубликатов.</span><span class="sxs-lookup"><span data-stu-id="2bf45-159">When **true**, the top row is ignored when considering duplicates.</span></span> <span data-ttu-id="2bf45-160">Функция `removeDuplicates` возвращает объект `RemoveDuplicatesResult`, указывающий количество удаленных строк и количество оставшихся уникальных строк.</span><span class="sxs-lookup"><span data-stu-id="2bf45-160">The `removeDuplicates` function returns a `RemoveDuplicatesResult` object that specifies the number of rows removed and the number of unique rows remaining.</span></span>

<span data-ttu-id="2bf45-161">При использовании функции `removeDuplicates` диапазона, учитывайте следующее:</span><span class="sxs-lookup"><span data-stu-id="2bf45-161">When using a range's `removeDuplicates` function, keep the following in mind:</span></span>

- <span data-ttu-id="2bf45-162">Функция `removeDuplicates` рассматривает значения ячеек, а не результаты функций.</span><span class="sxs-lookup"><span data-stu-id="2bf45-162">`removeDuplicates` considers cell values, not function results.</span></span> <span data-ttu-id="2bf45-163">Если две разные функции вычисляют одинаковый результат, значения ячеек не считаются повторяющимися.</span><span class="sxs-lookup"><span data-stu-id="2bf45-163">If two different functions evaluate to the same result, the cell values are not considered duplicates.</span></span>
- <span data-ttu-id="2bf45-164">Пустые ячейки не игнорируются функцией `removeDuplicates`.</span><span class="sxs-lookup"><span data-stu-id="2bf45-164">Empty cells are not ignored by `removeDuplicates`.</span></span> <span data-ttu-id="2bf45-165">Значение пустой ячейки обрабатывается как любое другое значение.</span><span class="sxs-lookup"><span data-stu-id="2bf45-165">The value of an empty cell is treated like any other value.</span></span> <span data-ttu-id="2bf45-166">Это означает, что пустые строки, содержащиеся в диапазоне, будут включены в объект `RemoveDuplicatesResult`.</span><span class="sxs-lookup"><span data-stu-id="2bf45-166">This means empty rows contained within in the range will be included in the `RemoveDuplicatesResult`.</span></span>

<span data-ttu-id="2bf45-167">В приведенном ниже примере показано удаление записей с повторяющимися значениями в первом столбце.</span><span class="sxs-lookup"><span data-stu-id="2bf45-167">The following sample shows the removal of entries with duplicate values in the first column.</span></span>

```js
Excel.run(async (context) => {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:D11");

    var deleteResult = range.removeDuplicates([0],true);
    deleteResult.load();

    return context.sync().then(function () {
        console.log(deleteResult.removed + " entries with duplicate names removed.");
        console.log(deleteResult.uniqueRemaining + " entries with unique names remain in the range.");
    });
}).catch(errorHandlerFunction);
```

<span data-ttu-id="2bf45-168">*Прежде чем предыдущая функция была запущена.*</span><span class="sxs-lookup"><span data-stu-id="2bf45-168">*Before the preceding function has been run.*</span></span>

![Данные в Excel перед запуском метода удаления дубликатов](../images/excel-ranges-remove-duplicates-before.png)

<span data-ttu-id="2bf45-170">*После запуска предыдущей функции.*</span><span class="sxs-lookup"><span data-stu-id="2bf45-170">*After the preceding function has been run.*</span></span>

![Данные в Excel после запуска метода удаления дубликатов](../images/excel-ranges-remove-duplicates-after.png)

## <a name="see-also"></a><span data-ttu-id="2bf45-172">См. также</span><span class="sxs-lookup"><span data-stu-id="2bf45-172">See also</span></span>

- [<span data-ttu-id="2bf45-173">Работа с диапазонами с использованием API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="2bf45-173">Work with ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges.md)
- [<span data-ttu-id="2bf45-174">Основные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="2bf45-174">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)