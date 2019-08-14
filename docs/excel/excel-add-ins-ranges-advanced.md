---
title: Работа с диапазонами с использованием API JavaScript для Excel (дополнительные задачи)
description: ''
ms.date: 04/30/2019
localization_priority: Normal
ms.openlocfilehash: c8fbe1dcc75080c932b4c3e2946fe62747d35c6b
ms.sourcegitcommit: 1c7e555733ee6d5a08e444a3c4c16635d998e032
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/14/2019
ms.locfileid: "36395598"
---
# <a name="work-with-ranges-using-the-excel-javascript-api-advanced"></a><span data-ttu-id="a8881-102">Работа с диапазонами с использованием API JavaScript для Excel (дополнительные задачи)</span><span class="sxs-lookup"><span data-stu-id="a8881-102">Work with ranges using the Excel JavaScript API (advanced)</span></span>

<span data-ttu-id="a8881-103">Эта статья основана на сведениях из статьи [Работа с диапазонами с использованием API JavaScript для Excel (основные задачи)](excel-add-ins-ranges.md) с предоставлением примеров кода, демонстрирующих способы выполнения более сложных задач с диапазонами с использованием API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="a8881-103">This article builds upon information in [Work with ranges using the Excel JavaScript API (fundamental)](excel-add-ins-ranges.md) by providing code samples that show how to perform more advanced tasks with ranges using the Excel JavaScript API.</span></span> <span data-ttu-id="a8881-104">Полный список свойств и методов, поддерживаемых объектом **Range**, см. в статье [Объект Range (API JavaScript для Excel)](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="a8881-104">For the complete list of properties and methods that the **Range** object supports, see [Range Object (JavaScript API for Excel)](/javascript/api/excel/excel.range).</span></span>

## <a name="work-with-dates-using-the-moment-msdate-plug-in"></a><span data-ttu-id="a8881-105">Работа с датами с использованием подключаемого модуля Moment-MSDate</span><span class="sxs-lookup"><span data-stu-id="a8881-105">Work with dates using the Moment-MSDate plug-in</span></span>

<span data-ttu-id="a8881-106">[Библиотека JavaScript Moment](https://momentjs.com/) предоставляет удобный способ использования дат и меток времени.</span><span class="sxs-lookup"><span data-stu-id="a8881-106">The [Moment JavaScript library](https://momentjs.com/) provides a convenient way to use dates and timestamps.</span></span> <span data-ttu-id="a8881-107">[Подключаемый модуль Moment-MSDate](https://www.npmjs.com/package/moment-msdate) преобразует формат моментов времени в предпочитаемый для Excel.</span><span class="sxs-lookup"><span data-stu-id="a8881-107">The [Moment-MSDate plug-in](https://www.npmjs.com/package/moment-msdate) converts the format of moments into one preferable for Excel.</span></span> <span data-ttu-id="a8881-108">Это тот же формат, который возвращает [функция ТДАТА](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46).</span><span class="sxs-lookup"><span data-stu-id="a8881-108">This is the same format the [NOW function](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) returns.</span></span>

<span data-ttu-id="a8881-109">В приведенном ниже коде показано, как установить для диапазона в **B4** метку момента времени.</span><span class="sxs-lookup"><span data-stu-id="a8881-109">The following code shows how to set the range at **B4** to a moment's timestamp:</span></span>

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

<span data-ttu-id="a8881-110">Это похоже на способ получения даты из ячейки и ее преобразования в формат момента времени или другой формат, как показано в приведенном ниже коде:</span><span class="sxs-lookup"><span data-stu-id="a8881-110">It is a similar technique to get the date back out of the cell and convert it to a moment or other format, as demonstrated in the following code:</span></span>

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

<span data-ttu-id="a8881-111">Вашей надстройке потребуется отформатировать диапазоны, чтобы отобразить даты в более понятной для человека форме.</span><span class="sxs-lookup"><span data-stu-id="a8881-111">Your add-in will have to format the ranges to display the dates in a more human-readable form.</span></span> <span data-ttu-id="a8881-112">В примере `"[$-409]m/d/yy h:mm AM/PM;@"` время отобразится как "12/3/18 3:57 PM".</span><span class="sxs-lookup"><span data-stu-id="a8881-112">The example of `"[$-409]m/d/yy h:mm AM/PM;@"` displays a time like "12/3/18 3:57 PM".</span></span> <span data-ttu-id="a8881-113">Дополнительные сведения о форматах чисел даты и времени см. в разделе "Рекомендации по форматам даты и времени" статьи [Рекомендации по настройке числовых форматов](https://support.office.com/article/review-guidelines-for-customizing-a-number-format-c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5).</span><span class="sxs-lookup"><span data-stu-id="a8881-113">For more information about date and time number formats, please see the "Guidelines for date and time formats" in the [Review guidelines for customizing a number format](https://support.office.com/article/review-guidelines-for-customizing-a-number-format-c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5) article.</span></span>

## <a name="work-with-multiple-ranges-simultaneously"></a><span data-ttu-id="a8881-114">Одновременное работу с несколькими диапазонами</span><span class="sxs-lookup"><span data-stu-id="a8881-114">Work with multiple ranges simultaneously</span></span>

<span data-ttu-id="a8881-115">Объект `RangeAreas` позволяет вашей надстройке выполнять операции над несколькими диапазонами одновременно.</span><span class="sxs-lookup"><span data-stu-id="a8881-115">The `RangeAreas` object lets your add-in perform operations on multiple ranges at once.</span></span> <span data-ttu-id="a8881-116">Эти диапазоны могут быть смежными, но это необязательно.</span><span class="sxs-lookup"><span data-stu-id="a8881-116">These ranges may be contiguous, but do not have to be.</span></span> <span data-ttu-id="a8881-117">Объект `RangeAreas` подробнее рассматривается в статье [Работа с несколькими диапазонами одновременно в надстройках Excel](excel-add-ins-multiple-ranges.md).</span><span class="sxs-lookup"><span data-stu-id="a8881-117">`RangeAreas` are further discussed in the article [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

## <a name="find-special-cells-within-a-range"></a><span data-ttu-id="a8881-118">Поиск специальных ячеек в диапазоне</span><span class="sxs-lookup"><span data-stu-id="a8881-118">Find special cells within a range</span></span>

<span data-ttu-id="a8881-119">Методы `Range.getSpecialCells()` и `Range.getSpecialCellsOrNullObject()` находят диапазоны с учетом характеристик ячеек и типов значений ячеек.</span><span class="sxs-lookup"><span data-stu-id="a8881-119">The `Range.getSpecialCells()` and `Range.getSpecialCellsOrNullObject()` methods find ranges based on the characteristics of their cells and the types of values of their cells.</span></span> <span data-ttu-id="a8881-120">Оба этих метода возвращают объекты `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="a8881-120">Both of these methods return `RangeAreas` objects.</span></span> <span data-ttu-id="a8881-121">Подписи методов из файла типов данных TypeScript:</span><span class="sxs-lookup"><span data-stu-id="a8881-121">Here are the signatures of the methods from the TypeScript data types file:</span></span>

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

<span data-ttu-id="a8881-122">В приведенном ниже примере используется метод `getSpecialCells`, чтобы найти все ячейки с формулами.</span><span class="sxs-lookup"><span data-stu-id="a8881-122">The following example uses the `getSpecialCells` method to find all the cells with formulas.</span></span> <span data-ttu-id="a8881-123">Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="a8881-123">About this code, note:</span></span>

- <span data-ttu-id="a8881-124">Он ограничивает часть листа, в которой требуется выполнять поиск, путем вызова сначала метода `Worksheet.getUsedRange`, а затем метода `getSpecialCells` только для этого диапазона.</span><span class="sxs-lookup"><span data-stu-id="a8881-124">It limits the part of the sheet that needs to be searched by first calling `Worksheet.getUsedRange` and calling `getSpecialCells` for only that range.</span></span>
- <span data-ttu-id="a8881-125">Метод `getSpecialCells` возвращает объект `RangeAreas`, поэтому все ячейки с формулами окрашены розовым цветом даже в том случае, если они не являются смежными.</span><span class="sxs-lookup"><span data-stu-id="a8881-125">The `getSpecialCells` method returns a `RangeAreas` object, so all of the cells with formulas will be colored pink even if they are not all contiguous.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var formulaRanges = usedRange.getSpecialCells(Excel.SpecialCellType.formulas);
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="a8881-126">Если в диапазоне нет ячеек с целевыми характеристиками, метод `getSpecialCells` выдает ошибку **ItemNotFound**.</span><span class="sxs-lookup"><span data-stu-id="a8881-126">If no cells with the targeted characteristic exist in the range, `getSpecialCells` throws an **ItemNotFound** error.</span></span> <span data-ttu-id="a8881-127">Это приведет к переадресации потока управления к блоку `catch`, если таковой существует.</span><span class="sxs-lookup"><span data-stu-id="a8881-127">This diverts the flow of control to a `catch` block, if there is one.</span></span> <span data-ttu-id="a8881-128">Если блок `catch` отсутствует, ошибка останавливает исполнение функции.</span><span class="sxs-lookup"><span data-stu-id="a8881-128">If there isn't a `catch` block, the error halts the function.</span></span>

<span data-ttu-id="a8881-129">Если ожидается, что всегда должны существовать ячейки с целевыми характеристиками, скорее всего вы захотите, чтобы код выдавал ошибку при их отсутствии.</span><span class="sxs-lookup"><span data-stu-id="a8881-129">If you expect that cells with the targeted characteristic should always exist, you'll likely want your code to throw an error if those cells aren't there.</span></span> <span data-ttu-id="a8881-130">Если отсутствие соответствующих ячеек является допустимым сценарием, ваш код должен проверить наличие такой возможности и корректно выполнить действие без выдачи ошибки.</span><span class="sxs-lookup"><span data-stu-id="a8881-130">If it's a valid scenario that there aren't any matching cells, your code should check for this possibility and handle it gracefully without throwing an error.</span></span> <span data-ttu-id="a8881-131">Добиться такого поведения можно с помощью метода `getSpecialCellsOrNullObject` и возвращаемого им свойства `isNullObject`.</span><span class="sxs-lookup"><span data-stu-id="a8881-131">You can achieve this behavior with the `getSpecialCellsOrNullObject` method and its returned `isNullObject` property.</span></span> <span data-ttu-id="a8881-132">Этот шаблон используется в приведенном ниже примере.</span><span class="sxs-lookup"><span data-stu-id="a8881-132">The following example uses this pattern.</span></span> <span data-ttu-id="a8881-133">Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="a8881-133">About this code, note:</span></span>

- <span data-ttu-id="a8881-134">Метод `getSpecialCellsOrNullObject` всегда возвращает прокси-объект, поэтому он не может иметь значение `null` в обычном смысле JavaScript.</span><span class="sxs-lookup"><span data-stu-id="a8881-134">The `getSpecialCellsOrNullObject` method always returns a proxy object, so it is never `null` in the ordinary JavaScript sense.</span></span> <span data-ttu-id="a8881-135">Но если соответствующие ячейки не обнаружены, свойству `isNullObject` объекта присваивается значение `true`.</span><span class="sxs-lookup"><span data-stu-id="a8881-135">But if no matching cells are found, the `isNullObject` property of the object is set to `true`.</span></span>
- <span data-ttu-id="a8881-136">Он вызывает `context.sync` *перед* тестированием свойства `isNullObject`.</span><span class="sxs-lookup"><span data-stu-id="a8881-136">It calls `context.sync` *before* it tests the `isNullObject` property.</span></span> <span data-ttu-id="a8881-137">Это требование для всех методов и свойств `*OrNullObject`, так как всегда нужно загружать и синхронизировать свойство, чтобы его прочесть.</span><span class="sxs-lookup"><span data-stu-id="a8881-137">This is a requirement with all `*OrNullObject` methods and properties, because you always have to load and sync a property in order to read it.</span></span> <span data-ttu-id="a8881-138">Однако необязательно *явно* загружать свойство `isNullObject`.</span><span class="sxs-lookup"><span data-stu-id="a8881-138">However, it is not necessary to *explicitly* load the `isNullObject` property.</span></span> <span data-ttu-id="a8881-139">Оно автоматически загружается с помощью `context.sync`, даже если `load` не вызывается для объекта.</span><span class="sxs-lookup"><span data-stu-id="a8881-139">It is automatically loaded by the `context.sync` even if `load` is not called on the object.</span></span> <span data-ttu-id="a8881-140">Дополнительные сведения см. в разделе [\*OrNullObject](/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#ornullobject-methods).</span><span class="sxs-lookup"><span data-stu-id="a8881-140">For more information, see [\*OrNullObject](/office/dev/add-ins/excel/excel-add-ins-advanced-concepts#ornullobject-methods).</span></span>
- <span data-ttu-id="a8881-141">Этот код можно проверить, выбрав сначала диапазон без ячеек с формулами и запустив его.</span><span class="sxs-lookup"><span data-stu-id="a8881-141">You can test this code by first selecting a range that has no formula cells and running it.</span></span> <span data-ttu-id="a8881-142">Затем следует выбрать диапазон, содержащий по крайней мере одну ячейку с формулой, и снова запустить его.</span><span class="sxs-lookup"><span data-stu-id="a8881-142">Then select a range that has at least one cell with a formula and run it again.</span></span>

```js
Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    var formulaRanges = range.getSpecialCellsOrNullObject(Excel.SpecialCellType.formulas);
    return context.sync()
        .then(function() {
            if (formulaRanges.isNullObject) {
                console.log("No cells have formulas");
            }
            else {
                formulaRanges.format.fill.color = "pink";
            }
        })
        .then(context.sync);
})
```

<span data-ttu-id="a8881-143">Для удобства во всех других примерах в этой статье используйте метод `getSpecialCells` вместо `getSpecialCellsOrNullObject`.</span><span class="sxs-lookup"><span data-stu-id="a8881-143">For simplicity, all other examples in this article use the `getSpecialCells` method instead of  `getSpecialCellsOrNullObject`.</span></span>

### <a name="narrow-the-target-cells-with-cell-value-types"></a><span data-ttu-id="a8881-144">Ограничение целевых ячеек с помощью типа значений ячеек</span><span class="sxs-lookup"><span data-stu-id="a8881-144">Narrow the target cells with cell value types</span></span>

<span data-ttu-id="a8881-145">Методы `Range.getSpecialCells()` и `Range.getSpecialCellsOrNullObject()` принимают необязательный второй параметр, используемый для дополнительного ограничения целевых ячеек.</span><span class="sxs-lookup"><span data-stu-id="a8881-145">The `Range.getSpecialCells()` and `Range.getSpecialCellsOrNullObject()` methods accept an optional second parameter used to further narrow down the targeted cells.</span></span> <span data-ttu-id="a8881-146">Этот второй параметр `Excel.SpecialCellValueType` используется для указания того, что требуются только ячейки, содержащие определенные типы значений.</span><span class="sxs-lookup"><span data-stu-id="a8881-146">This second parameter is an `Excel.SpecialCellValueType` you use to specify that you only want cells that contain certain types of values.</span></span>

> [!NOTE]
> <span data-ttu-id="a8881-147">Параметр `Excel.SpecialCellValueType` можно использовать, только если для параметра `Excel.SpecialCellType` задано значение `Excel.SpecialCellType.formulas` или `Excel.SpecialCellType.constants`.</span><span class="sxs-lookup"><span data-stu-id="a8881-147">The `Excel.SpecialCellValueType` parameter can only be used if the `Excel.SpecialCellType` is `Excel.SpecialCellType.formulas` or `Excel.SpecialCellType.constants`.</span></span>

#### <a name="test-for-a-single-cell-value-type"></a><span data-ttu-id="a8881-148">Тестирование для ячеек с одним типом значений</span><span class="sxs-lookup"><span data-stu-id="a8881-148">Test for a single cell value type</span></span>

<span data-ttu-id="a8881-149">Для перечисления `Excel.SpecialCellValueType` существует четыре основных типа (в дополнение к другим объединенным значениям, описанным ниже в этом разделе):</span><span class="sxs-lookup"><span data-stu-id="a8881-149">The `Excel.SpecialCellValueType` enum has these four basic types (in addition to the other combined values described later in this section):</span></span>

- `Excel.SpecialCellValueType.errors`
- <span data-ttu-id="a8881-150">`Excel.SpecialCellValueType.logical` (означает логическое значение)</span><span class="sxs-lookup"><span data-stu-id="a8881-150">`Excel.SpecialCellValueType.logical` (which means boolean)</span></span>
- `Excel.SpecialCellValueType.numbers`
- `Excel.SpecialCellValueType.text`

<span data-ttu-id="a8881-151">В приведенном ниже примере выполняется поиск специальных ячеек, являющихся числовыми константами, и их окрашивание в розовый цвет.</span><span class="sxs-lookup"><span data-stu-id="a8881-151">The following example finds special cells that are numerical constants and colors those cells pink.</span></span> <span data-ttu-id="a8881-152">Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="a8881-152">About this code, note:</span></span>

- <span data-ttu-id="a8881-153">Он выделяет только ячейки с числовым значением литерала.</span><span class="sxs-lookup"><span data-stu-id="a8881-153">It will only highlight cells that have a literal number value.</span></span> <span data-ttu-id="a8881-154">Он не выделяет ячейки с формулой (даже если результат является числом), логическим значением, текстовым значением или ячейки с состоянием ошибки.</span><span class="sxs-lookup"><span data-stu-id="a8881-154">It will not highlight cells that have a formula (even if the result is a number) or a boolean, text, or error state cells.</span></span>
- <span data-ttu-id="a8881-155">Чтобы протестировать код, убедитесь, что в листе есть ячейки с числовыми значениями литералов, ячейки с другими значениями литералов и ячейки с формулами.</span><span class="sxs-lookup"><span data-stu-id="a8881-155">To test the code, be sure the worksheet has some cells with literal number values, some with other kinds of literal values, and some with formulas.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var constantNumberRanges = usedRange.getSpecialCells(
        Excel.SpecialCellType.constants,
        Excel.SpecialCellValueType.numbers);
    constantNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

#### <a name="test-for-multiple-cell-value-types"></a><span data-ttu-id="a8881-156">Тестирование для ячеек с несколькими типами значений</span><span class="sxs-lookup"><span data-stu-id="a8881-156">Test for multiple cell value types</span></span>

<span data-ttu-id="a8881-157">Иногда требуется работать с ячейками, имеющими несколько типов значений, например со всеми ячейками с текстовыми значениями и всеми ячейками с логическими значениями (`Excel.SpecialCellValueType.logical`).</span><span class="sxs-lookup"><span data-stu-id="a8881-157">Sometimes you need to operate on more than one cell value type, such as all text-valued and all boolean-valued (`Excel.SpecialCellValueType.logical`) cells.</span></span> <span data-ttu-id="a8881-158">Для перечисления `Excel.SpecialCellValueType` существуют значения с объединенными типами.</span><span class="sxs-lookup"><span data-stu-id="a8881-158">The `Excel.SpecialCellValueType` enum has values with combined types.</span></span> <span data-ttu-id="a8881-159">Например, `Excel.SpecialCellValueType.logicalText` обрабатывает все ячейки с логическими и текстовыми значениями.</span><span class="sxs-lookup"><span data-stu-id="a8881-159">For example, `Excel.SpecialCellValueType.logicalText` targets all boolean and all text-valued cells.</span></span> <span data-ttu-id="a8881-160">`Excel.SpecialCellValueType.all` является значением по умолчанию, которое не ограничивает возвращаемые типы значений ячеек.</span><span class="sxs-lookup"><span data-stu-id="a8881-160">`Excel.SpecialCellValueType.all` is the default value, which does not limit the cell value types returned.</span></span> <span data-ttu-id="a8881-161">В приведенном ниже примере окрашены все ячейки с формулами, которые производят числовое или логическое значение.</span><span class="sxs-lookup"><span data-stu-id="a8881-161">The following example colors all cells with formulas that produce number or boolean value.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var formulaLogicalNumberRanges = usedRange.getSpecialCells(
        Excel.SpecialCellType.formulas,
        Excel.SpecialCellValueType.logicalNumbers);
    formulaLogicalNumberRanges.format.fill.color = "pink";

    return context.sync();
})
```

## <a name="copy-and-paste"></a><span data-ttu-id="a8881-162">Copy and paste</span><span class="sxs-lookup"><span data-stu-id="a8881-162">Copy and paste</span></span>

<span data-ttu-id="a8881-163">Функция `copyFrom` диапазона реплицирует поведение копирования и вставки пользовательского интерфейса Excel.</span><span class="sxs-lookup"><span data-stu-id="a8881-163">Range’s `copyFrom` function replicates the copy-and-paste behavior of the Excel UI.</span></span> <span data-ttu-id="a8881-164">Диапазон объекта, который вызывается `copyFrom`, является назначением.</span><span class="sxs-lookup"><span data-stu-id="a8881-164">The range object that `copyFrom` is called on is the destination.</span></span>
<span data-ttu-id="a8881-165">Источник для копирования передается как диапазон или адрес строки, представляющий диапазон.</span><span class="sxs-lookup"><span data-stu-id="a8881-165">The source to be copied is passed as a range or a string address representing a range.</span></span>
<span data-ttu-id="a8881-166">В следующем примере кода копируются данные из **A1:E1** в диапазон, начиная с **G1** (который заканчивается вставкой в **G1:K1**).</span><span class="sxs-lookup"><span data-stu-id="a8881-166">The following code sample copies the data from **A1:E1** into the range starting at **G1** (which ends up pasting into **G1:K1**).</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy everything from "A1:E1" into "G1" and the cells afterwards ("G1:K1")
    sheet.getRange("G1").copyFrom("A1:E1");
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="a8881-167">У функции `Range.copyFrom` есть три необязательных параметра.</span><span class="sxs-lookup"><span data-stu-id="a8881-167">`Range.copyFrom` has three optional parameters.</span></span>

```TypeScript
copyFrom(sourceRange: Range | RangeAreas | string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean): void;
```

<span data-ttu-id="a8881-168">`copyType` указывает, какие данные копируются из источника в назначение.</span><span class="sxs-lookup"><span data-stu-id="a8881-168">`copyType` specifies what data gets copied from the source to the destination.</span></span>

- <span data-ttu-id="a8881-169">`Excel.RangeCopyType.formulas` переносит формулы в ячейках источника и сохраняет относительное положение диапазонов этих формул.</span><span class="sxs-lookup"><span data-stu-id="a8881-169">`Excel.RangeCopyType.formulas` transfers the formulas in the source cells and preserves the relative positioning of those formulas’ ranges.</span></span> <span data-ttu-id="a8881-170">Все записи, не являющиеся формулами, копируются в исходном виде.</span><span class="sxs-lookup"><span data-stu-id="a8881-170">Any non-formula entries are copied as-is.</span></span>
- <span data-ttu-id="a8881-171">`Excel.RangeCopyType.values` копирует значения данных, а в случае формул — результат формулы.</span><span class="sxs-lookup"><span data-stu-id="a8881-171">`Excel.RangeCopyType.values` copies the data values and, in the case of formulas, the result of the formula.</span></span>
- <span data-ttu-id="a8881-172">`Excel.RangeCopyType.formats` копирует форматирование диапазона, включая шрифт, цвет и другие параметры форматирования, но без значений.</span><span class="sxs-lookup"><span data-stu-id="a8881-172">`Excel.RangeCopyType.formats` copies the formatting of the range, including font, color, and other format settings, but no values.</span></span>
- <span data-ttu-id="a8881-173">`Excel.RangeCopyType.all` (вариант по умолчанию) копирует данные и форматирование, сохраняя формулы ячеек при их обнаружении.</span><span class="sxs-lookup"><span data-stu-id="a8881-173">`Excel.RangeCopyType.all` (the default option) copies both data and formatting, preserving cells’ formulas if found.</span></span>

<span data-ttu-id="a8881-174">`skipBlanks` устанавливает, будут ли копироваться пустые ячейки в назначение.</span><span class="sxs-lookup"><span data-stu-id="a8881-174">`skipBlanks` sets whether blank cells are copied into the destination.</span></span> <span data-ttu-id="a8881-175">Если значение равно true, `copyFrom` пропускает пустые ячейки в диапазоне источника.</span><span class="sxs-lookup"><span data-stu-id="a8881-175">When true, `copyFrom` skips blank cells in the source range.</span></span>
<span data-ttu-id="a8881-176">Пропущенные ячейки не перезапишут существующие данные в соответствующих им ячейках конечного диапазона.</span><span class="sxs-lookup"><span data-stu-id="a8881-176">Skipped cells will not overwrite the existing data of their corresponding cells in the destination range.</span></span> <span data-ttu-id="a8881-177">Значение по умолчанию: false.</span><span class="sxs-lookup"><span data-stu-id="a8881-177">The default is false.</span></span>

<span data-ttu-id="a8881-178">`transpose` определяет, переставляются ли данные в исходное расположение, то есть переключаются ли строки и столбцы.</span><span class="sxs-lookup"><span data-stu-id="a8881-178">`transpose` determines whether or not the data is transposed, meaning its rows and columns are switched, into the source location.</span></span>
<span data-ttu-id="a8881-179">Переставленный диапазон переключается на главной диагонали, поэтому строки **1**, **2** и **3** становятся столбцами **A**, **B** и **C**.</span><span class="sxs-lookup"><span data-stu-id="a8881-179">A transposed range is flipped along the main diagonal, so rows **1**, **2**, and **3** will become columns **A**, **B**, and **C**.</span></span>

<span data-ttu-id="a8881-180">В приведенном ниже примере кода и изображениях демонстрируется это поведение в простом сценарии.</span><span class="sxs-lookup"><span data-stu-id="a8881-180">The following code sample and images demonstrate this behavior in a simple scenario.</span></span>

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

<span data-ttu-id="a8881-181">*Прежде чем предыдущая функция была запущена.*</span><span class="sxs-lookup"><span data-stu-id="a8881-181">*Before the preceding function has been run.*</span></span>

![Данные в Excel перед запуском метода копирования диапазона](../images/excel-range-copyfrom-skipblanks-before.png)

<span data-ttu-id="a8881-183">*После запуска предыдущей функции.*</span><span class="sxs-lookup"><span data-stu-id="a8881-183">*After the preceding function has been run.*</span></span>

![Данные в Excel после запуска метода копирования диапазона](../images/excel-range-copyfrom-skipblanks-after.png)

## <a name="remove-duplicates"></a><span data-ttu-id="a8881-185">Удаление дубликатов</span><span class="sxs-lookup"><span data-stu-id="a8881-185">Remove duplicates</span></span>

<span data-ttu-id="a8881-186">Функция `removeDuplicates` объекта Range удаляет строки с повторяющимися записями в указанных столбцах.</span><span class="sxs-lookup"><span data-stu-id="a8881-186">The Range object's `removeDuplicates` function removes rows with duplicate entries in the specified columns.</span></span> <span data-ttu-id="a8881-187">Функция проверяет каждую строку в диапазоне от индекса с наименьшим значением до индекса с наибольшим значением (сверху вниз).</span><span class="sxs-lookup"><span data-stu-id="a8881-187">The function goes through each row in the range from the lowest-valued index to the highest-valued index in the range (from top to bottom).</span></span> <span data-ttu-id="a8881-188">Строка удаляется, если значение в ее указанном столбце или столбцах уже встречалось в диапазоне.</span><span class="sxs-lookup"><span data-stu-id="a8881-188">A row is deleted if a value in its specified column or columns appeared earlier in the range.</span></span> <span data-ttu-id="a8881-189">Строки в диапазоне под удаленной строкой сдвигаются вверх.</span><span class="sxs-lookup"><span data-stu-id="a8881-189">Rows in the range below the deleted row are shifted up.</span></span> <span data-ttu-id="a8881-190">Функция `removeDuplicates` не влияет на положение ячеек вне диапазона.</span><span class="sxs-lookup"><span data-stu-id="a8881-190">`removeDuplicates` does not affect the position of cells outside of the range.</span></span>

<span data-ttu-id="a8881-191">Функция `removeDuplicates` использует параметр `number[]`, представляющий индексы столбцов, которые проверяются на наличие дубликатов.</span><span class="sxs-lookup"><span data-stu-id="a8881-191">`removeDuplicates` takes in a `number[]` representing the column indices which are checked for duplicates.</span></span> <span data-ttu-id="a8881-192">Этот массив отсчитывается от нуля относительно диапазона, а не листа.</span><span class="sxs-lookup"><span data-stu-id="a8881-192">This array is zero-based and relative to the range, not the worksheet.</span></span> <span data-ttu-id="a8881-193">Функция также использует логический параметр, который определяет, является ли первая строка заголовком.</span><span class="sxs-lookup"><span data-stu-id="a8881-193">The function also takes in a boolean parameter that specifies whether the first row is a header.</span></span> <span data-ttu-id="a8881-194">При значении **true** верхняя строка игнорируется при поиске дубликатов.</span><span class="sxs-lookup"><span data-stu-id="a8881-194">When **true**, the top row is ignored when considering duplicates.</span></span> <span data-ttu-id="a8881-195">Функция `removeDuplicates` возвращает объект `RemoveDuplicatesResult`, указывающий количество удаленных строк и количество оставшихся уникальных строк.</span><span class="sxs-lookup"><span data-stu-id="a8881-195">The `removeDuplicates` function returns a `RemoveDuplicatesResult` object that specifies the number of rows removed and the number of unique rows remaining.</span></span>

<span data-ttu-id="a8881-196">При использовании функции `removeDuplicates` диапазона, учитывайте следующее:</span><span class="sxs-lookup"><span data-stu-id="a8881-196">When using a range's `removeDuplicates` function, keep the following in mind:</span></span>

- <span data-ttu-id="a8881-197">Функция `removeDuplicates` рассматривает значения ячеек, а не результаты функций.</span><span class="sxs-lookup"><span data-stu-id="a8881-197">`removeDuplicates` considers cell values, not function results.</span></span> <span data-ttu-id="a8881-198">Если две разные функции вычисляют одинаковый результат, значения ячеек не считаются повторяющимися.</span><span class="sxs-lookup"><span data-stu-id="a8881-198">If two different functions evaluate to the same result, the cell values are not considered duplicates.</span></span>
- <span data-ttu-id="a8881-199">Пустые ячейки не игнорируются функцией `removeDuplicates`.</span><span class="sxs-lookup"><span data-stu-id="a8881-199">Empty cells are not ignored by `removeDuplicates`.</span></span> <span data-ttu-id="a8881-200">Значение пустой ячейки обрабатывается как любое другое значение.</span><span class="sxs-lookup"><span data-stu-id="a8881-200">The value of an empty cell is treated like any other value.</span></span> <span data-ttu-id="a8881-201">Это означает, что пустые строки, содержащиеся в диапазоне, будут включены в объект `RemoveDuplicatesResult`.</span><span class="sxs-lookup"><span data-stu-id="a8881-201">This means empty rows contained within in the range will be included in the `RemoveDuplicatesResult`.</span></span>

<span data-ttu-id="a8881-202">В приведенном ниже примере показано удаление записей с повторяющимися значениями в первом столбце.</span><span class="sxs-lookup"><span data-stu-id="a8881-202">The following sample shows the removal of entries with duplicate values in the first column.</span></span>

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

<span data-ttu-id="a8881-203">*Прежде чем предыдущая функция была запущена.*</span><span class="sxs-lookup"><span data-stu-id="a8881-203">*Before the preceding function has been run.*</span></span>

![Данные в Excel перед запуском метода удаления дубликатов](../images/excel-ranges-remove-duplicates-before.png)

<span data-ttu-id="a8881-205">*После запуска предыдущей функции.*</span><span class="sxs-lookup"><span data-stu-id="a8881-205">*After the preceding function has been run.*</span></span>

![Данные в Excel после запуска метода удаления дубликатов](../images/excel-ranges-remove-duplicates-after.png)

## <a name="see-also"></a><span data-ttu-id="a8881-207">См. также</span><span class="sxs-lookup"><span data-stu-id="a8881-207">See also</span></span>

- [<span data-ttu-id="a8881-208">Работа с диапазонами с использованием API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="a8881-208">Work with ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges.md)
- [<span data-ttu-id="a8881-209">Основные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="a8881-209">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="a8881-210">Работа с несколькими диапазонами одновременно в надстройках Excel</span><span class="sxs-lookup"><span data-stu-id="a8881-210">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
