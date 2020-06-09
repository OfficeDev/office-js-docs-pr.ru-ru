---
title: Работа с диапазонами с использованием API JavaScript для Excel (дополнительные задачи)
description: Расширенные функции и сценарии объектов Range, такие как специальные ячейки, удаление дубликатов и работа с датами.
ms.date: 05/06/2020
localization_priority: Normal
ms.openlocfilehash: 0a185551bf0ddd6b5d4d5a90e4faac7ce78e2cc9
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609750"
---
# <a name="work-with-ranges-using-the-excel-javascript-api-advanced"></a><span data-ttu-id="90480-103">Работа с диапазонами с использованием API JavaScript для Excel (дополнительные задачи)</span><span class="sxs-lookup"><span data-stu-id="90480-103">Work with ranges using the Excel JavaScript API (advanced)</span></span>

<span data-ttu-id="90480-104">Эта статья основана на сведениях из статьи [Работа с диапазонами с использованием API JavaScript для Excel (основные задачи)](excel-add-ins-ranges.md) с предоставлением примеров кода, демонстрирующих способы выполнения более сложных задач с диапазонами с использованием API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="90480-104">This article builds upon information in [Work with ranges using the Excel JavaScript API (fundamental)](excel-add-ins-ranges.md) by providing code samples that show how to perform more advanced tasks with ranges using the Excel JavaScript API.</span></span> <span data-ttu-id="90480-105">Полный список свойств и методов, `Range` поддерживаемых объектом, представлен в разделе [объект Range (API JavaScript для Excel)](/javascript/api/excel/excel.range).</span><span class="sxs-lookup"><span data-stu-id="90480-105">For the complete list of properties and methods that the `Range` object supports, see [Range Object (JavaScript API for Excel)](/javascript/api/excel/excel.range).</span></span>

## <a name="work-with-dates-using-the-moment-msdate-plug-in"></a><span data-ttu-id="90480-106">Работа с датами с использованием подключаемого модуля Moment-MSDate</span><span class="sxs-lookup"><span data-stu-id="90480-106">Work with dates using the Moment-MSDate plug-in</span></span>

<span data-ttu-id="90480-107">[Библиотека JavaScript Moment](https://momentjs.com/) предоставляет удобный способ использования дат и меток времени.</span><span class="sxs-lookup"><span data-stu-id="90480-107">The [Moment JavaScript library](https://momentjs.com/) provides a convenient way to use dates and timestamps.</span></span> <span data-ttu-id="90480-108">[Подключаемый модуль Moment-MSDate](https://www.npmjs.com/package/moment-msdate) преобразует формат моментов времени в предпочитаемый для Excel.</span><span class="sxs-lookup"><span data-stu-id="90480-108">The [Moment-MSDate plug-in](https://www.npmjs.com/package/moment-msdate) converts the format of moments into one preferable for Excel.</span></span> <span data-ttu-id="90480-109">Это тот же формат, который возвращает [функция ТДАТА](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46).</span><span class="sxs-lookup"><span data-stu-id="90480-109">This is the same format the [NOW function](https://support.office.com/article/now-function-3337fd29-145a-4347-b2e6-20c904739c46) returns.</span></span>

<span data-ttu-id="90480-110">В приведенном ниже коде показано, как установить для диапазона в **B4** метку момента времени.</span><span class="sxs-lookup"><span data-stu-id="90480-110">The following code shows how to set the range at **B4** to a moment's timestamp:</span></span>

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

<span data-ttu-id="90480-111">Это похоже на способ получения даты из ячейки и ее преобразования в формат момента времени или другой формат, как показано в приведенном ниже коде:</span><span class="sxs-lookup"><span data-stu-id="90480-111">It is a similar technique to get the date back out of the cell and convert it to a moment or other format, as demonstrated in the following code:</span></span>

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

<span data-ttu-id="90480-112">Вашей надстройке потребуется отформатировать диапазоны, чтобы отобразить даты в более понятной для человека форме.</span><span class="sxs-lookup"><span data-stu-id="90480-112">Your add-in will have to format the ranges to display the dates in a more human-readable form.</span></span> <span data-ttu-id="90480-113">В примере `"[$-409]m/d/yy h:mm AM/PM;@"` время отобразится как "12/3/18 3:57 PM".</span><span class="sxs-lookup"><span data-stu-id="90480-113">The example of `"[$-409]m/d/yy h:mm AM/PM;@"` displays a time like "12/3/18 3:57 PM".</span></span> <span data-ttu-id="90480-114">Дополнительные сведения о форматах чисел даты и времени см. в разделе "Рекомендации по форматам даты и времени" статьи [Рекомендации по настройке числовых форматов](https://support.microsoft.com/office/c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5).</span><span class="sxs-lookup"><span data-stu-id="90480-114">For more information about date and time number formats, please see the "Guidelines for date and time formats" in the [Review guidelines for customizing a number format](https://support.microsoft.com/office/c0a1d1fa-d3f4-4018-96b7-9c9354dd99f5) article.</span></span>

## <a name="work-with-multiple-ranges-simultaneously"></a><span data-ttu-id="90480-115">Одновременное работу с несколькими диапазонами</span><span class="sxs-lookup"><span data-stu-id="90480-115">Work with multiple ranges simultaneously</span></span>

<span data-ttu-id="90480-116">Объект [RangeAreas](/javascript/api/excel/excel.rangeareas) позволяет надстройке выполнять операции над несколькими диапазонами одновременно.</span><span class="sxs-lookup"><span data-stu-id="90480-116">The [RangeAreas](/javascript/api/excel/excel.rangeareas) object lets your add-in perform operations on multiple ranges at once.</span></span> <span data-ttu-id="90480-117">Эти диапазоны могут быть смежными, но это необязательно.</span><span class="sxs-lookup"><span data-stu-id="90480-117">These ranges may be contiguous, but do not have to be.</span></span> <span data-ttu-id="90480-118">Объект `RangeAreas` подробнее рассматривается в статье [Работа с несколькими диапазонами одновременно в надстройках Excel](excel-add-ins-multiple-ranges.md).</span><span class="sxs-lookup"><span data-stu-id="90480-118">`RangeAreas` are further discussed in the article [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

## <a name="find-special-cells-within-a-range"></a><span data-ttu-id="90480-119">Поиск специальных ячеек в диапазоне</span><span class="sxs-lookup"><span data-stu-id="90480-119">Find special cells within a range</span></span>

<span data-ttu-id="90480-120">Методы [Range. жетспеЦиалцеллс](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-) и [Range. жетспеЦиалцеллсорнуллобжект](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-) находят диапазоны на основе характеристик их ячеек и типов значений их ячеек.</span><span class="sxs-lookup"><span data-stu-id="90480-120">The [Range.getSpecialCells](/javascript/api/excel/excel.range#getspecialcells-celltype--cellvaluetype-) and [Range.getSpecialCellsOrNullObject](/javascript/api/excel/excel.range#getspecialcellsornullobject-celltype--cellvaluetype-) methods find ranges based on the characteristics of their cells and the types of values of their cells.</span></span> <span data-ttu-id="90480-121">Оба этих метода возвращают объекты `RangeAreas`.</span><span class="sxs-lookup"><span data-stu-id="90480-121">Both of these methods return `RangeAreas` objects.</span></span> <span data-ttu-id="90480-122">Подписи методов из файла типов данных TypeScript:</span><span class="sxs-lookup"><span data-stu-id="90480-122">Here are the signatures of the methods from the TypeScript data types file:</span></span>

```typescript
getSpecialCells(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

```typescript
getSpecialCellsOrNullObject(cellType: Excel.SpecialCellType, cellValueType?: Excel.SpecialCellValueType): Excel.RangeAreas;
```

<span data-ttu-id="90480-123">В приведенном ниже примере используется метод `getSpecialCells`, чтобы найти все ячейки с формулами.</span><span class="sxs-lookup"><span data-stu-id="90480-123">The following example uses the `getSpecialCells` method to find all the cells with formulas.</span></span> <span data-ttu-id="90480-124">Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="90480-124">About this code, note:</span></span>

- <span data-ttu-id="90480-125">Он ограничивает часть листа, в которой требуется выполнять поиск, путем вызова сначала метода `Worksheet.getUsedRange`, а затем метода `getSpecialCells` только для этого диапазона.</span><span class="sxs-lookup"><span data-stu-id="90480-125">It limits the part of the sheet that needs to be searched by first calling `Worksheet.getUsedRange` and calling `getSpecialCells` for only that range.</span></span>
- <span data-ttu-id="90480-126">Метод `getSpecialCells` возвращает объект `RangeAreas`, поэтому все ячейки с формулами окрашены розовым цветом даже в том случае, если они не являются смежными.</span><span class="sxs-lookup"><span data-stu-id="90480-126">The `getSpecialCells` method returns a `RangeAreas` object, so all of the cells with formulas will be colored pink even if they are not all contiguous.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var usedRange = sheet.getUsedRange();
    var formulaRanges = usedRange.getSpecialCells(Excel.SpecialCellType.formulas);
    formulaRanges.format.fill.color = "pink";

    return context.sync();
})
```

<span data-ttu-id="90480-127">Если в диапазоне нет ячеек с целевыми характеристиками, метод `getSpecialCells` выдает ошибку **ItemNotFound**.</span><span class="sxs-lookup"><span data-stu-id="90480-127">If no cells with the targeted characteristic exist in the range, `getSpecialCells` throws an **ItemNotFound** error.</span></span> <span data-ttu-id="90480-128">Это приведет к переадресации потока управления к блоку `catch`, если таковой существует.</span><span class="sxs-lookup"><span data-stu-id="90480-128">This diverts the flow of control to a `catch` block, if there is one.</span></span> <span data-ttu-id="90480-129">Если `catch` блок отсутствует, то ошибка приостанавливается для метода.</span><span class="sxs-lookup"><span data-stu-id="90480-129">If there isn't a `catch` block, the error halts the method.</span></span>

<span data-ttu-id="90480-130">Если ожидается, что всегда должны существовать ячейки с целевыми характеристиками, скорее всего вы захотите, чтобы код выдавал ошибку при их отсутствии.</span><span class="sxs-lookup"><span data-stu-id="90480-130">If you expect that cells with the targeted characteristic should always exist, you'll likely want your code to throw an error if those cells aren't there.</span></span> <span data-ttu-id="90480-131">Если отсутствие соответствующих ячеек является допустимым сценарием, ваш код должен проверить наличие такой возможности и корректно выполнить действие без выдачи ошибки.</span><span class="sxs-lookup"><span data-stu-id="90480-131">If it's a valid scenario that there aren't any matching cells, your code should check for this possibility and handle it gracefully without throwing an error.</span></span> <span data-ttu-id="90480-132">Добиться такого поведения можно с помощью метода `getSpecialCellsOrNullObject` и возвращаемого им свойства `isNullObject`.</span><span class="sxs-lookup"><span data-stu-id="90480-132">You can achieve this behavior with the `getSpecialCellsOrNullObject` method and its returned `isNullObject` property.</span></span> <span data-ttu-id="90480-133">Этот шаблон используется в приведенном ниже примере.</span><span class="sxs-lookup"><span data-stu-id="90480-133">The following example uses this pattern.</span></span> <span data-ttu-id="90480-134">Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="90480-134">About this code, note:</span></span>

- <span data-ttu-id="90480-135">Метод `getSpecialCellsOrNullObject` всегда возвращает прокси-объект, поэтому он не может иметь значение `null` в обычном смысле JavaScript.</span><span class="sxs-lookup"><span data-stu-id="90480-135">The `getSpecialCellsOrNullObject` method always returns a proxy object, so it is never `null` in the ordinary JavaScript sense.</span></span> <span data-ttu-id="90480-136">Но если соответствующие ячейки не обнаружены, свойству `isNullObject` объекта присваивается значение `true`.</span><span class="sxs-lookup"><span data-stu-id="90480-136">But if no matching cells are found, the `isNullObject` property of the object is set to `true`.</span></span>
- <span data-ttu-id="90480-137">Он вызывает `context.sync` *перед* тестированием свойства `isNullObject`.</span><span class="sxs-lookup"><span data-stu-id="90480-137">It calls `context.sync` *before* it tests the `isNullObject` property.</span></span> <span data-ttu-id="90480-138">Это требование для всех методов и свойств `*OrNullObject`, так как всегда нужно загружать и синхронизировать свойство, чтобы его прочесть.</span><span class="sxs-lookup"><span data-stu-id="90480-138">This is a requirement with all `*OrNullObject` methods and properties, because you always have to load and sync a property in order to read it.</span></span> <span data-ttu-id="90480-139">Однако необязательно *явно* загружать свойство `isNullObject`.</span><span class="sxs-lookup"><span data-stu-id="90480-139">However, it is not necessary to *explicitly* load the `isNullObject` property.</span></span> <span data-ttu-id="90480-140">Оно автоматически загружается с помощью `context.sync`, даже если `load` не вызывается для объекта.</span><span class="sxs-lookup"><span data-stu-id="90480-140">It is automatically loaded by the `context.sync` even if `load` is not called on the object.</span></span> <span data-ttu-id="90480-141">Дополнительные сведения см. в разделе [\*OrNullObject](../excel/excel-add-ins-advanced-concepts.md#ornullobject-methods).</span><span class="sxs-lookup"><span data-stu-id="90480-141">For more information, see [\*OrNullObject](../excel/excel-add-ins-advanced-concepts.md#ornullobject-methods).</span></span>
- <span data-ttu-id="90480-142">Этот код можно проверить, выбрав сначала диапазон без ячеек с формулами и запустив его.</span><span class="sxs-lookup"><span data-stu-id="90480-142">You can test this code by first selecting a range that has no formula cells and running it.</span></span> <span data-ttu-id="90480-143">Затем следует выбрать диапазон, содержащий по крайней мере одну ячейку с формулой, и снова запустить его.</span><span class="sxs-lookup"><span data-stu-id="90480-143">Then select a range that has at least one cell with a formula and run it again.</span></span>

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

<span data-ttu-id="90480-144">Для удобства во всех других примерах в этой статье используйте метод `getSpecialCells` вместо `getSpecialCellsOrNullObject`.</span><span class="sxs-lookup"><span data-stu-id="90480-144">For simplicity, all other examples in this article use the `getSpecialCells` method instead of  `getSpecialCellsOrNullObject`.</span></span>

### <a name="narrow-the-target-cells-with-cell-value-types"></a><span data-ttu-id="90480-145">Ограничение целевых ячеек с помощью типа значений ячеек</span><span class="sxs-lookup"><span data-stu-id="90480-145">Narrow the target cells with cell value types</span></span>

<span data-ttu-id="90480-146">Методы `Range.getSpecialCells()` и `Range.getSpecialCellsOrNullObject()` принимают необязательный второй параметр, используемый для дополнительного ограничения целевых ячеек.</span><span class="sxs-lookup"><span data-stu-id="90480-146">The `Range.getSpecialCells()` and `Range.getSpecialCellsOrNullObject()` methods accept an optional second parameter used to further narrow down the targeted cells.</span></span> <span data-ttu-id="90480-147">Этот второй параметр `Excel.SpecialCellValueType` используется для указания того, что требуются только ячейки, содержащие определенные типы значений.</span><span class="sxs-lookup"><span data-stu-id="90480-147">This second parameter is an `Excel.SpecialCellValueType` you use to specify that you only want cells that contain certain types of values.</span></span>

> [!NOTE]
> <span data-ttu-id="90480-148">Параметр `Excel.SpecialCellValueType` можно использовать, только если для параметра `Excel.SpecialCellType` задано значение `Excel.SpecialCellType.formulas` или `Excel.SpecialCellType.constants`.</span><span class="sxs-lookup"><span data-stu-id="90480-148">The `Excel.SpecialCellValueType` parameter can only be used if the `Excel.SpecialCellType` is `Excel.SpecialCellType.formulas` or `Excel.SpecialCellType.constants`.</span></span>

#### <a name="test-for-a-single-cell-value-type"></a><span data-ttu-id="90480-149">Тестирование для ячеек с одним типом значений</span><span class="sxs-lookup"><span data-stu-id="90480-149">Test for a single cell value type</span></span>

<span data-ttu-id="90480-150">Для перечисления `Excel.SpecialCellValueType` существует четыре основных типа (в дополнение к другим объединенным значениям, описанным ниже в этом разделе):</span><span class="sxs-lookup"><span data-stu-id="90480-150">The `Excel.SpecialCellValueType` enum has these four basic types (in addition to the other combined values described later in this section):</span></span>

- `Excel.SpecialCellValueType.errors`
- <span data-ttu-id="90480-151">`Excel.SpecialCellValueType.logical` (означает логическое значение)</span><span class="sxs-lookup"><span data-stu-id="90480-151">`Excel.SpecialCellValueType.logical` (which means boolean)</span></span>
- `Excel.SpecialCellValueType.numbers`
- `Excel.SpecialCellValueType.text`

<span data-ttu-id="90480-152">В приведенном ниже примере выполняется поиск специальных ячеек, являющихся числовыми константами, и их окрашивание в розовый цвет.</span><span class="sxs-lookup"><span data-stu-id="90480-152">The following example finds special cells that are numerical constants and colors those cells pink.</span></span> <span data-ttu-id="90480-153">Вот что нужно знать об этом коде:</span><span class="sxs-lookup"><span data-stu-id="90480-153">About this code, note:</span></span>

- <span data-ttu-id="90480-154">Он выделяет только ячейки с числовым значением литерала.</span><span class="sxs-lookup"><span data-stu-id="90480-154">It will only highlight cells that have a literal number value.</span></span> <span data-ttu-id="90480-155">Он не выделяет ячейки с формулой (даже если результат является числом), логическим значением, текстовым значением или ячейки с состоянием ошибки.</span><span class="sxs-lookup"><span data-stu-id="90480-155">It will not highlight cells that have a formula (even if the result is a number) or a boolean, text, or error state cells.</span></span>
- <span data-ttu-id="90480-156">Чтобы протестировать код, убедитесь, что в листе есть ячейки с числовыми значениями литералов, ячейки с другими значениями литералов и ячейки с формулами.</span><span class="sxs-lookup"><span data-stu-id="90480-156">To test the code, be sure the worksheet has some cells with literal number values, some with other kinds of literal values, and some with formulas.</span></span>

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

#### <a name="test-for-multiple-cell-value-types"></a><span data-ttu-id="90480-157">Тестирование для ячеек с несколькими типами значений</span><span class="sxs-lookup"><span data-stu-id="90480-157">Test for multiple cell value types</span></span>

<span data-ttu-id="90480-158">Иногда требуется работать с ячейками, имеющими несколько типов значений, например со всеми ячейками с текстовыми значениями и всеми ячейками с логическими значениями (`Excel.SpecialCellValueType.logical`).</span><span class="sxs-lookup"><span data-stu-id="90480-158">Sometimes you need to operate on more than one cell value type, such as all text-valued and all boolean-valued (`Excel.SpecialCellValueType.logical`) cells.</span></span> <span data-ttu-id="90480-159">Для перечисления `Excel.SpecialCellValueType` существуют значения с объединенными типами.</span><span class="sxs-lookup"><span data-stu-id="90480-159">The `Excel.SpecialCellValueType` enum has values with combined types.</span></span> <span data-ttu-id="90480-160">Например, `Excel.SpecialCellValueType.logicalText` обрабатывает все ячейки с логическими и текстовыми значениями.</span><span class="sxs-lookup"><span data-stu-id="90480-160">For example, `Excel.SpecialCellValueType.logicalText` targets all boolean and all text-valued cells.</span></span> <span data-ttu-id="90480-161">`Excel.SpecialCellValueType.all` является значением по умолчанию, которое не ограничивает возвращаемые типы значений ячеек.</span><span class="sxs-lookup"><span data-stu-id="90480-161">`Excel.SpecialCellValueType.all` is the default value, which does not limit the cell value types returned.</span></span> <span data-ttu-id="90480-162">В приведенном ниже примере окрашены все ячейки с формулами, которые производят числовое или логическое значение.</span><span class="sxs-lookup"><span data-stu-id="90480-162">The following example colors all cells with formulas that produce number or boolean value.</span></span>

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

## <a name="cut-copy-and-paste"></a><span data-ttu-id="90480-163">Команды "Вырезать", "Копировать" и "Вставить"</span><span class="sxs-lookup"><span data-stu-id="90480-163">Cut, copy, and paste</span></span>

### <a name="copy-and-paste"></a><span data-ttu-id="90480-164">Copy and paste</span><span class="sxs-lookup"><span data-stu-id="90480-164">Copy and paste</span></span>

<span data-ttu-id="90480-165">Метод [Range. copyFrom](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-) реплицирует действия **копирования** и **вставки** пользовательского интерфейса Excel.</span><span class="sxs-lookup"><span data-stu-id="90480-165">The [Range.copyFrom](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-) method replicates the **Copy** and **Paste** actions of the Excel UI.</span></span> <span data-ttu-id="90480-166">Диапазон объекта, который вызывается `copyFrom`, является назначением.</span><span class="sxs-lookup"><span data-stu-id="90480-166">The range object that `copyFrom` is called on is the destination.</span></span> <span data-ttu-id="90480-167">Источник для копирования передается как диапазон или адрес строки, представляющий диапазон.</span><span class="sxs-lookup"><span data-stu-id="90480-167">The source to be copied is passed as a range or a string address representing a range.</span></span>

<span data-ttu-id="90480-168">В следующем примере кода копируются данные из **A1:E1** в диапазон, начиная с **G1** (который заканчивается вставкой в **G1:K1**).</span><span class="sxs-lookup"><span data-stu-id="90480-168">The following code sample copies the data from **A1:E1** into the range starting at **G1** (which ends up pasting into **G1:K1**).</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy everything from "A1:E1" into "G1" and the cells afterwards ("G1:K1")
    sheet.getRange("G1").copyFrom("A1:E1");
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="90480-169">У функции `Range.copyFrom` есть три необязательных параметра.</span><span class="sxs-lookup"><span data-stu-id="90480-169">`Range.copyFrom` has three optional parameters.</span></span>

```TypeScript
copyFrom(sourceRange: Range | RangeAreas | string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean): void;
```

<span data-ttu-id="90480-170">`copyType` указывает, какие данные копируются из источника в назначение.</span><span class="sxs-lookup"><span data-stu-id="90480-170">`copyType` specifies what data gets copied from the source to the destination.</span></span>

- <span data-ttu-id="90480-171">`Excel.RangeCopyType.formulas`передает формулы в исходных ячейках и сохраняет относительное расположение диапазонов этих формул.</span><span class="sxs-lookup"><span data-stu-id="90480-171">`Excel.RangeCopyType.formulas` transfers the formulas in the source cells and preserves the relative positioning of those formulas' ranges.</span></span> <span data-ttu-id="90480-172">Все записи, не являющиеся формулами, копируются в исходном виде.</span><span class="sxs-lookup"><span data-stu-id="90480-172">Any non-formula entries are copied as-is.</span></span>
- <span data-ttu-id="90480-173">`Excel.RangeCopyType.values` копирует значения данных, а в случае формул — результат формулы.</span><span class="sxs-lookup"><span data-stu-id="90480-173">`Excel.RangeCopyType.values` copies the data values and, in the case of formulas, the result of the formula.</span></span>
- <span data-ttu-id="90480-174">`Excel.RangeCopyType.formats` копирует форматирование диапазона, включая шрифт, цвет и другие параметры форматирования, но без значений.</span><span class="sxs-lookup"><span data-stu-id="90480-174">`Excel.RangeCopyType.formats` copies the formatting of the range, including font, color, and other format settings, but no values.</span></span>
- <span data-ttu-id="90480-175">`Excel.RangeCopyType.all`(параметр по умолчанию) копирует данные и форматирование, сохраняя формулы ячеек, если они найдены.</span><span class="sxs-lookup"><span data-stu-id="90480-175">`Excel.RangeCopyType.all` (the default option) copies both data and formatting, preserving cells' formulas if found.</span></span>

<span data-ttu-id="90480-176">`skipBlanks` устанавливает, будут ли копироваться пустые ячейки в назначение.</span><span class="sxs-lookup"><span data-stu-id="90480-176">`skipBlanks` sets whether blank cells are copied into the destination.</span></span> <span data-ttu-id="90480-177">Если значение равно true, `copyFrom` пропускает пустые ячейки в диапазоне источника.</span><span class="sxs-lookup"><span data-stu-id="90480-177">When true, `copyFrom` skips blank cells in the source range.</span></span>
<span data-ttu-id="90480-178">Пропущенные ячейки не перезапишут существующие данные в соответствующих им ячейках конечного диапазона.</span><span class="sxs-lookup"><span data-stu-id="90480-178">Skipped cells will not overwrite the existing data of their corresponding cells in the destination range.</span></span> <span data-ttu-id="90480-179">Значение по умолчанию: false.</span><span class="sxs-lookup"><span data-stu-id="90480-179">The default is false.</span></span>

<span data-ttu-id="90480-180">`transpose` определяет, переставляются ли данные в исходное расположение, то есть переключаются ли строки и столбцы.</span><span class="sxs-lookup"><span data-stu-id="90480-180">`transpose` determines whether or not the data is transposed, meaning its rows and columns are switched, into the source location.</span></span>
<span data-ttu-id="90480-181">Переставленный диапазон переключается на главной диагонали, поэтому строки **1**, **2** и **3** становятся столбцами **A**, **B** и **C**.</span><span class="sxs-lookup"><span data-stu-id="90480-181">A transposed range is flipped along the main diagonal, so rows **1**, **2**, and **3** will become columns **A**, **B**, and **C**.</span></span>

<span data-ttu-id="90480-182">В приведенном ниже примере кода и изображениях демонстрируется это поведение в простом сценарии.</span><span class="sxs-lookup"><span data-stu-id="90480-182">The following code sample and images demonstrate this behavior in a simple scenario.</span></span>

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

<span data-ttu-id="90480-183">*Прежде чем предыдущая функция была запущена.*</span><span class="sxs-lookup"><span data-stu-id="90480-183">*Before the preceding function has been run.*</span></span>

![Данные в Excel перед запуском метода копирования диапазона](../images/excel-range-copyfrom-skipblanks-before.png)

<span data-ttu-id="90480-185">*После запуска предыдущей функции.*</span><span class="sxs-lookup"><span data-stu-id="90480-185">*After the preceding function has been run.*</span></span>

![Данные в Excel после запуска метода копирования диапазона](../images/excel-range-copyfrom-skipblanks-after.png)

### <a name="cut-and-paste-move-cells"></a><span data-ttu-id="90480-187">Вырезание и вставка (перемещение) ячеек</span><span class="sxs-lookup"><span data-stu-id="90480-187">Cut and paste (move) cells</span></span>

<span data-ttu-id="90480-188">Метод [Range. moveTo](/javascript/api/excel/excel.range#moveto-destinationrange-) перемещает ячейки в новое расположение в книге.</span><span class="sxs-lookup"><span data-stu-id="90480-188">The [Range.moveTo](/javascript/api/excel/excel.range#moveto-destinationrange-) method moves cells to a new location in the workbook.</span></span> <span data-ttu-id="90480-189">Поведение перемещения ячейки работает так же, как при перемещении ячеек, [перетаскивая границу диапазона](https://support.office.com/article/Move-or-copy-cells-and-cell-contents-803d65eb-6a3e-4534-8c6f-ff12d1c4139e) или при выполнении действий по **вырезанию** и **вставке** .</span><span class="sxs-lookup"><span data-stu-id="90480-189">This cell movement behavior works the same as when cells are moved by [dragging the range border](https://support.office.com/article/Move-or-copy-cells-and-cell-contents-803d65eb-6a3e-4534-8c6f-ff12d1c4139e) or when taking the **Cut** and **Paste** actions.</span></span> <span data-ttu-id="90480-190">Форматирование и значения диапазона перемещаются в расположение, указанное в качестве `destinationRange` параметра.</span><span class="sxs-lookup"><span data-stu-id="90480-190">Both the formatting and values of the range are moved to the location specified as the `destinationRange` parameter.</span></span>

<span data-ttu-id="90480-191">В следующем примере кода показан диапазон, перемещенный с `Range.moveTo` методом.</span><span class="sxs-lookup"><span data-stu-id="90480-191">The following code sample shows a range being moved with the `Range.moveTo` method.</span></span> <span data-ttu-id="90480-192">Обратите внимание, что если конечный диапазон меньше исходного, он будет развернут, чтобы охватывать исходное содержимое.</span><span class="sxs-lookup"><span data-stu-id="90480-192">Note that if the destination range is smaller than the source, it will be expanded to encompass the source content.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.getRange("F1").values = [["Moved Range"]];

    // Move the cells "A1:E1" to "G1" (which fills the range "G1:K1").
    sheet.getRange("A1:E1").moveTo("G1");
    return context.sync();
});
```

## <a name="remove-duplicates"></a><span data-ttu-id="90480-193">Удаление дубликатов</span><span class="sxs-lookup"><span data-stu-id="90480-193">Remove duplicates</span></span>

<span data-ttu-id="90480-194">Метод [Range. removeDuplicates](/javascript/api/excel/excel.range#removeduplicates-columns--includesheader-) удаляет строки с повторяющимися записями в указанных столбцах.</span><span class="sxs-lookup"><span data-stu-id="90480-194">The [Range.removeDuplicates](/javascript/api/excel/excel.range#removeduplicates-columns--includesheader-) method removes rows with duplicate entries in the specified columns.</span></span> <span data-ttu-id="90480-195">Метод проходит через каждую строку в диапазоне от самого низкого значения до индекса с максимальным значением в диапазоне (сверху вниз).</span><span class="sxs-lookup"><span data-stu-id="90480-195">The method goes through each row in the range from the lowest-valued index to the highest-valued index in the range (from top to bottom).</span></span> <span data-ttu-id="90480-196">Строка удаляется, если значение в ее указанном столбце или столбцах уже встречалось в диапазоне.</span><span class="sxs-lookup"><span data-stu-id="90480-196">A row is deleted if a value in its specified column or columns appeared earlier in the range.</span></span> <span data-ttu-id="90480-197">Строки в диапазоне под удаленной строкой сдвигаются вверх.</span><span class="sxs-lookup"><span data-stu-id="90480-197">Rows in the range below the deleted row are shifted up.</span></span> <span data-ttu-id="90480-198">Функция `removeDuplicates` не влияет на положение ячеек вне диапазона.</span><span class="sxs-lookup"><span data-stu-id="90480-198">`removeDuplicates` does not affect the position of cells outside of the range.</span></span>

<span data-ttu-id="90480-199">Функция `removeDuplicates` использует параметр `number[]`, представляющий индексы столбцов, которые проверяются на наличие дубликатов.</span><span class="sxs-lookup"><span data-stu-id="90480-199">`removeDuplicates` takes in a `number[]` representing the column indices which are checked for duplicates.</span></span> <span data-ttu-id="90480-200">Этот массив отсчитывается от нуля относительно диапазона, а не листа.</span><span class="sxs-lookup"><span data-stu-id="90480-200">This array is zero-based and relative to the range, not the worksheet.</span></span> <span data-ttu-id="90480-201">Метод также использует логический параметр, указывающий, является ли первая строка заголовком.</span><span class="sxs-lookup"><span data-stu-id="90480-201">The method also takes in a boolean parameter that specifies whether the first row is a header.</span></span> <span data-ttu-id="90480-202">При значении **true** верхняя строка игнорируется при поиске дубликатов.</span><span class="sxs-lookup"><span data-stu-id="90480-202">When **true**, the top row is ignored when considering duplicates.</span></span> <span data-ttu-id="90480-203">`removeDuplicates`Метод возвращает `RemoveDuplicatesResult` объект, указывающий количество удаленных строк и количество оставшихся уникальных строк.</span><span class="sxs-lookup"><span data-stu-id="90480-203">The `removeDuplicates` method returns a `RemoveDuplicatesResult` object that specifies the number of rows removed and the number of unique rows remaining.</span></span>

<span data-ttu-id="90480-204">При использовании метода диапазона учитывайте `removeDuplicates` следующее:</span><span class="sxs-lookup"><span data-stu-id="90480-204">When using a range's `removeDuplicates` method, keep the following in mind:</span></span>

- <span data-ttu-id="90480-205">Функция `removeDuplicates` рассматривает значения ячеек, а не результаты функций.</span><span class="sxs-lookup"><span data-stu-id="90480-205">`removeDuplicates` considers cell values, not function results.</span></span> <span data-ttu-id="90480-206">Если две разные функции вычисляют одинаковый результат, значения ячеек не считаются повторяющимися.</span><span class="sxs-lookup"><span data-stu-id="90480-206">If two different functions evaluate to the same result, the cell values are not considered duplicates.</span></span>
- <span data-ttu-id="90480-207">Пустые ячейки не игнорируются функцией `removeDuplicates`.</span><span class="sxs-lookup"><span data-stu-id="90480-207">Empty cells are not ignored by `removeDuplicates`.</span></span> <span data-ttu-id="90480-208">Значение пустой ячейки обрабатывается как любое другое значение.</span><span class="sxs-lookup"><span data-stu-id="90480-208">The value of an empty cell is treated like any other value.</span></span> <span data-ttu-id="90480-209">Это означает, что пустые строки, содержащиеся в диапазоне, будут включены в объект `RemoveDuplicatesResult`.</span><span class="sxs-lookup"><span data-stu-id="90480-209">This means empty rows contained within in the range will be included in the `RemoveDuplicatesResult`.</span></span>

<span data-ttu-id="90480-210">В приведенном ниже примере показано удаление записей с повторяющимися значениями в первом столбце.</span><span class="sxs-lookup"><span data-stu-id="90480-210">The following sample shows the removal of entries with duplicate values in the first column.</span></span>

```js
Excel.run(function (context) {
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

<span data-ttu-id="90480-211">*Прежде чем предыдущая функция была запущена.*</span><span class="sxs-lookup"><span data-stu-id="90480-211">*Before the preceding function has been run.*</span></span>

![Данные в Excel перед выполнением метода удаления дубликатов в диапазоне](../images/excel-ranges-remove-duplicates-before.png)

<span data-ttu-id="90480-213">*После запуска предыдущей функции.*</span><span class="sxs-lookup"><span data-stu-id="90480-213">*After the preceding function has been run.*</span></span>

![Данные в Excel после запуска метода удаления повторяющихся значений диапазона](../images/excel-ranges-remove-duplicates-after.png)

## <a name="group-data-for-an-outline"></a><span data-ttu-id="90480-215">Группирование данных для структуры</span><span class="sxs-lookup"><span data-stu-id="90480-215">Group data for an outline</span></span>

<span data-ttu-id="90480-216">Строки или столбцы диапазона можно объединять для создания [структуры](https://support.office.com/article/Outline-group-data-in-a-worksheet-08CE98C4-0063-4D42-8AC7-8278C49E9AFF).</span><span class="sxs-lookup"><span data-stu-id="90480-216">Rows or columns of a range can be grouped together to create an [outline](https://support.office.com/article/Outline-group-data-in-a-worksheet-08CE98C4-0063-4D42-8AC7-8278C49E9AFF).</span></span> <span data-ttu-id="90480-217">Эти группы можно сворачивать и разворачивать для скрытия и отображения соответствующих ячеек.</span><span class="sxs-lookup"><span data-stu-id="90480-217">These groups can be collapsed and expanded to hide and show the corresponding cells.</span></span> <span data-ttu-id="90480-218">Это упрощает быстрый анализ данных в верхней строке.</span><span class="sxs-lookup"><span data-stu-id="90480-218">This makes quick analysis of top-line data easier.</span></span> <span data-ttu-id="90480-219">Используйте [Range. Group](/javascript/api/excel/excel.range#group-groupoption-) , чтобы сделать эти группы структуры.</span><span class="sxs-lookup"><span data-stu-id="90480-219">Use [Range.group](/javascript/api/excel/excel.range#group-groupoption-) to make these outline groups.</span></span>

<span data-ttu-id="90480-220">Структура может иметь иерархию, где небольшие группы вложены в крупные группы.</span><span class="sxs-lookup"><span data-stu-id="90480-220">An outline can have a hierarchy, where smaller groups are nested under larger groups.</span></span> <span data-ttu-id="90480-221">Это позволяет просматривать структуру на разных уровнях.</span><span class="sxs-lookup"><span data-stu-id="90480-221">This allows the outline to be viewed at different levels.</span></span> <span data-ttu-id="90480-222">Изменение видимого уровня структуры можно выполнить программным способом с помощью метода [листа. шоваутлинелевелс](/javascript/api/excel/excel.worksheet#showoutlinelevels-rowlevels--columnlevels-) .</span><span class="sxs-lookup"><span data-stu-id="90480-222">Changing the visible outline level can be done programmatically through the [Worksheet.showOutlineLevels](/javascript/api/excel/excel.worksheet#showoutlinelevels-rowlevels--columnlevels-) method.</span></span> <span data-ttu-id="90480-223">Обратите внимание, что Excel поддерживает только восемь уровней групп структуры.</span><span class="sxs-lookup"><span data-stu-id="90480-223">Note that Excel only supports eight levels of outline groups.</span></span>

<span data-ttu-id="90480-224">В приведенном ниже примере кода показано, как создать структуру с двумя уровнями групп для строк и столбцов.</span><span class="sxs-lookup"><span data-stu-id="90480-224">The following code sample shows how to create an outline with two levels of groups for both the rows and columns.</span></span> <span data-ttu-id="90480-225">На следующем изображении показаны группирования этой структуры.</span><span class="sxs-lookup"><span data-stu-id="90480-225">The subsequent image shows the groupings of that outline.</span></span> <span data-ttu-id="90480-226">Обратите внимание, что в примере кода сгруппированные диапазоны не включают строку или столбец элемента управления структуры (итоговые значения для этого примера).</span><span class="sxs-lookup"><span data-stu-id="90480-226">Note that in the code sample, the ranges being grouped do not include the row or column of the outline control (the "Totals" for this example).</span></span> <span data-ttu-id="90480-227">Группа определяет, что будет свернуто, а не как строка или столбец с элементом управления.</span><span class="sxs-lookup"><span data-stu-id="90480-227">A group defines what will be collapsed, not the row or column with the control.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    // Group the larger, main level. Note that the outline controls
    // will be on row 10, meaning 4-9 will collapse and expand.
    sheet.getRange("4:9").group(Excel.GroupOption.byRows);

    // Group the smaller, sublevels. Note that the outline controls
    // will be on rows 6 and 9, meaning 4-5 and 7-8 will collapse and expand.
    sheet.getRange("4:5").group(Excel.GroupOption.byRows);
    sheet.getRange("7:8").group(Excel.GroupOption.byRows);

    // Group the larger, main level. Note that the outline controls
    // will be on column R, meaning C-Q will collapse and expand.
    sheet.getRange("C:Q").group(Excel.GroupOption.byColumns);

    // Group the smaller, sublevels. Note that the outline controls
    // will be on columns G, L, and R, meaning C-F, H-K, and M-P will collapse and expand.
    sheet.getRange("C:F").group(Excel.GroupOption.byColumns);
    sheet.getRange("H:K").group(Excel.GroupOption.byColumns);
    sheet.getRange("M:P").group(Excel.GroupOption.byColumns);
    return context.sync();
}).catch(errorHandlerFunction);

```

![Диапазон с двумя уровнями структуры с двумя измерениями](../images/excel-outline.png)

<span data-ttu-id="90480-229">Чтобы разгруппировать группу строк или столбцов, используйте метод [Range. Ungroup](/javascript/api/excel/excel.range#ungroup-groupoption-) .</span><span class="sxs-lookup"><span data-stu-id="90480-229">To ungroup a row or column group, use the [Range.ungroup](/javascript/api/excel/excel.range#ungroup-groupoption-) method.</span></span> <span data-ttu-id="90480-230">Это приведет к удалению внешнего уровня структуры.</span><span class="sxs-lookup"><span data-stu-id="90480-230">This removes the outermost level from the outline.</span></span> <span data-ttu-id="90480-231">Если несколько групп одного и того же типа строк или столбцов находятся на одном уровне в пределах указанного диапазона, все эти группы размещаются в разгруппировании.</span><span class="sxs-lookup"><span data-stu-id="90480-231">If multiple groups of the same row or column type are at the same level within the specified range, all of those groups are ungrouped.</span></span>

## <a name="see-also"></a><span data-ttu-id="90480-232">См. также</span><span class="sxs-lookup"><span data-stu-id="90480-232">See also</span></span>

- [<span data-ttu-id="90480-233">Работа с диапазонами с использованием API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="90480-233">Work with ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges.md)
- [<span data-ttu-id="90480-234">Основные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="90480-234">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="90480-235">Работа с несколькими диапазонами одновременно в надстройках Excel</span><span class="sxs-lookup"><span data-stu-id="90480-235">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
