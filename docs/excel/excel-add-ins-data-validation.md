---
title: Добавление проверки данных в диапазоны Excel
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: b0b2d886ceb9026ebe41414fed4ef8be1b59cc95
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449213"
---
# <a name="add-data-validation-to-excel-ranges"></a><span data-ttu-id="cefbc-102">Добавление проверки данных в диапазоны Excel</span><span class="sxs-lookup"><span data-stu-id="cefbc-102">Add data validation to Excel ranges</span></span>

<span data-ttu-id="cefbc-103">Библиотека JavaScript Excel предоставляет API, позволяющие вашей надстройке добавлять функцию автоматической проверки данных для таблиц, столбцов, строк и других диапазонов в книге. </span><span class="sxs-lookup"><span data-stu-id="cefbc-103">The Excel JavaScript Library provides APIs to enable your add-in to add automatic data validation to tables, columns, rows, and other ranges in a workbook.</span></span> <span data-ttu-id="cefbc-104">Чтобы понять принципы и терминологию проверки данных, ознакомьтесь со следующими статьями о том, как пользователи добавляют проверку данных в пользовательском интерфейсе Excel:</span><span class="sxs-lookup"><span data-stu-id="cefbc-104">To understand the concepts and the terminology of data validation, please see the following articles about how users add data validation through the Excel UI:</span></span>

- [<span data-ttu-id="cefbc-105">Применение проверки данных к ячейкам</span><span class="sxs-lookup"><span data-stu-id="cefbc-105">Apply data validation to cells</span></span>](https://support.office.com/article/Apply-data-validation-to-cells-29FECBCC-D1B9-42C1-9D76-EFF3CE5F7249)
- [<span data-ttu-id="cefbc-106">Подробнее о проверке данных</span><span class="sxs-lookup"><span data-stu-id="cefbc-106">More on data validation</span></span>](https://support.office.com/article/More-on-data-validation-f38dee73-9900-4ca6-9301-8a5f6e1f0c4c)
- [<span data-ttu-id="cefbc-107">Описание и примеры проверки данных в Excel</span><span class="sxs-lookup"><span data-stu-id="cefbc-107">Description and examples of data validation in Excel</span></span>](https://support.microsoft.com/help/211485/description-and-examples-of-data-validation-in-excel)

## <a name="programmatic-control-of-data-validation"></a><span data-ttu-id="cefbc-108">Программное управление проверкой данных</span><span class="sxs-lookup"><span data-stu-id="cefbc-108">Programmatic control of data validation</span></span>

<span data-ttu-id="cefbc-109">Свойство `Range.dataValidation`, которое получает объект [DataValidation](/javascript/api/excel/excel.datavalidation), является точкой входа для программного управления проверкой данных в Excel.</span><span class="sxs-lookup"><span data-stu-id="cefbc-109">The `Range.dataValidation` property, which takes a [DataValidation](/javascript/api/excel/excel.datavalidation) object, is the entry point for programmatic control of data validation in Excel.</span></span> <span data-ttu-id="cefbc-110">Существует пять свойств объекта `DataValidation`:</span><span class="sxs-lookup"><span data-stu-id="cefbc-110">There are five properties to the `DataValidation` object:</span></span>

- <span data-ttu-id="cefbc-111">`rule` — определяет, какие данные для диапазона являются допустимыми.</span><span class="sxs-lookup"><span data-stu-id="cefbc-111">`rule` &#8212; Defines what constitutes valid data for the range.</span></span> <span data-ttu-id="cefbc-112">См. статью [DataValidationRule](/javascript/api/excel/excel.datavalidationrule).</span><span class="sxs-lookup"><span data-stu-id="cefbc-112">See [DataValidationRule](/javascript/api/excel/excel.datavalidationrule).</span></span>
- <span data-ttu-id="cefbc-113">`errorAlert` — указывает, появляется ли ошибка, если пользователь вводит недопустимые данные, и определяет текст, название и стиль оповещения, например **Informational** (информирование), **Warning** (предупреждение) и **Stop** (остановка). </span><span class="sxs-lookup"><span data-stu-id="cefbc-113">`errorAlert` &#8212; Specifies whether an error pops up if the user enters invalid data, and defines the alert text, title, and style; for example, **Informational**, **Warning**, and **Stop**.</span></span> <span data-ttu-id="cefbc-114">См. статью [DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert).</span><span class="sxs-lookup"><span data-stu-id="cefbc-114">See [DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert).</span></span>
- <span data-ttu-id="cefbc-115">`prompt` — указывает, появляется ли подсказка, когда пользователь наводит указатель мыши на диапазон, и определяет текст подсказки.</span><span class="sxs-lookup"><span data-stu-id="cefbc-115">`prompt` &#8212; Specifies whether a prompt appears when the user hovers over the range and defines the prompt message.</span></span> <span data-ttu-id="cefbc-116">См. статью [DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt).</span><span class="sxs-lookup"><span data-stu-id="cefbc-116">See [DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt).</span></span>
- <span data-ttu-id="cefbc-117">`ignoreBlanks` — указывает, применяется ли правило проверки данных к пустым ячейкам в диапазоне.</span><span class="sxs-lookup"><span data-stu-id="cefbc-117">`ignoreBlanks` &#8212; Specifies whether the data validation rule applies to blank cells in the range.</span></span> <span data-ttu-id="cefbc-118">Значение по умолчанию: `true`.</span><span class="sxs-lookup"><span data-stu-id="cefbc-118">Defaults to `true`.</span></span>
- <span data-ttu-id="cefbc-119">`type` — идентификация типа проверки "только для чтения", например WholeNumber, Date, TextLength и т. д. Это свойство устанавливается неявно при задании свойства `rule`.</span><span class="sxs-lookup"><span data-stu-id="cefbc-119">`type` &#8212; A read-only identification of the validation type, such as WholeNumber, Date, TextLength, etc. It is set indirectly when you set the `rule` property.</span></span>

> [!NOTE]
> <span data-ttu-id="cefbc-120">Проверка данных, добавляемая программно, ведет себя так же, как проверка данных, добавляемая вручную. </span><span class="sxs-lookup"><span data-stu-id="cefbc-120">Data validation added programmatically behaves just like manually added data validation.</span></span> <span data-ttu-id="cefbc-121">В частности, обратите внимание на то, что проверка данных запускается только в том случае, если пользователь вводит значение в ячейку или копирует и вставляет ячейки из другого источника в книге и выбирает параметр вставки **Значения**.</span><span class="sxs-lookup"><span data-stu-id="cefbc-121">In particular, note that data validation is triggered only if the user directly enters a value into a cell or copies and pastes a cell from elsewhere in the workbook and chooses the **Values** paste option.</span></span> <span data-ttu-id="cefbc-122">Если пользователь копирует ячейку и выполняет простую вставку в диапазон проверки данных, проверка не выполняется.</span><span class="sxs-lookup"><span data-stu-id="cefbc-122">If the user copies a cell and does a plain paste into a range with data validation, validation is not triggered.</span></span>

## <a name="creating-validation-rules"></a><span data-ttu-id="cefbc-123">Создание правил проверки</span><span class="sxs-lookup"><span data-stu-id="cefbc-123">Creating validation rules</span></span>

<span data-ttu-id="cefbc-124">Чтобы добавить проверку данных в диапазон, ваш код должен установить свойство `rule` объекта `DataValidation` в `Range.dataValidation`.</span><span class="sxs-lookup"><span data-stu-id="cefbc-124">To add data validation to a range, your code must set the `rule` property of the `DataValidation` object in `Range.dataValidation`.</span></span> <span data-ttu-id="cefbc-125">Это приводит к получению объекта [DataValidationRule](/javascript/api/excel/excel.datavalidationrule), который имеет семь дополнительных свойств.</span><span class="sxs-lookup"><span data-stu-id="cefbc-125">This takes a [DataValidationRule](/javascript/api/excel/excel.datavalidationrule) object which has seven optional properties.</span></span> <span data-ttu-id="cefbc-126">*Максимум одно свойство может присутствовать в любом объекте `DataValidationRule`.*</span><span class="sxs-lookup"><span data-stu-id="cefbc-126">*No more than one of these properties may be present in any `DataValidationRule` object.*</span></span> <span data-ttu-id="cefbc-127">Указываемое свойство определяет тип выполняемой проверки.</span><span class="sxs-lookup"><span data-stu-id="cefbc-127">The property that you include determines the type of validation.</span></span>

### <a name="basic-and-datetime-validation-rule-types"></a><span data-ttu-id="cefbc-128">Типы правил проверки Basic и DateTime</span><span class="sxs-lookup"><span data-stu-id="cefbc-128">Basic and DateTime validation rule types</span></span>

<span data-ttu-id="cefbc-129">Первые три свойства `DataValidationRule` (т. е. типы правил проверки) в качестве своего значения принимают объект [BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation).</span><span class="sxs-lookup"><span data-stu-id="cefbc-129">The first three `DataValidationRule` properties (i.e., validation rule types) take a [BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation) object as their value.</span></span>

- <span data-ttu-id="cefbc-130">`wholeNumber` — требует целое число в дополнение к другим проверкам, указанным объектом `BasicDataValidation`.</span><span class="sxs-lookup"><span data-stu-id="cefbc-130">`wholeNumber` &#8212; Requires a whole number in addition to any other validation specified by the `BasicDataValidation` object.</span></span>
- <span data-ttu-id="cefbc-131">`decimal` — требует десятичное число в дополнение к другим проверкам, указанным объектом `BasicDataValidation`.</span><span class="sxs-lookup"><span data-stu-id="cefbc-131">`decimal` &#8212; Requires a decimal number in addition to any other validation specified by the `BasicDataValidation` object.</span></span>
- <span data-ttu-id="cefbc-132">`textLength` — применяет сведения проверки объекта `BasicDataValidation` к *длине* значения ячейки.</span><span class="sxs-lookup"><span data-stu-id="cefbc-132">`textLength` &#8212; Applies the validation details in the `BasicDataValidation` object to the *length* of the cell's value.</span></span>

<span data-ttu-id="cefbc-133">Ниже приведен пример создания правила проверки. </span><span class="sxs-lookup"><span data-stu-id="cefbc-133">Here is an example of creating a validation rule.</span></span> <span data-ttu-id="cefbc-134">Обратите внимание на следующие особенности этого кода:</span><span class="sxs-lookup"><span data-stu-id="cefbc-134">Note the following about this code:</span></span>

- <span data-ttu-id="cefbc-135">`operator` — это бинарный оператор "GreaterThan".</span><span class="sxs-lookup"><span data-stu-id="cefbc-135">The `operator` is the binary operator "GreaterThan".</span></span> <span data-ttu-id="cefbc-136">При использовании бинарного оператора значение, которое пользователь пытается ввести в ячейку, — это левый операнд, а значение, указанное в `formula1`, — это правый операнд.</span><span class="sxs-lookup"><span data-stu-id="cefbc-136">Whenever you use a binary operator, the value that the user tries to enter in the cell is the left-hand operand and the value specified in `formula1` is the right-hand operand.</span></span> <span data-ttu-id="cefbc-137">Поэтому согласно этому правилу только целые числа больше 0 являются допустимыми.</span><span class="sxs-lookup"><span data-stu-id="cefbc-137">So this rule says that only whole numbers that are greater than 0 are valid.</span></span> 
- <span data-ttu-id="cefbc-138">`formula1` — это жестко заданное число.</span><span class="sxs-lookup"><span data-stu-id="cefbc-138">The `formula1` is a hard-coded number.</span></span> <span data-ttu-id="cefbc-139">Если во время кодирования вы не знаете, какое значение должно быть задано, можно также использовать формулу Excel (в виде строки) для значения.</span><span class="sxs-lookup"><span data-stu-id="cefbc-139">If you don't know at coding time what the value should be, you can also use an Excel formula (as a string) for the value.</span></span> <span data-ttu-id="cefbc-140">Например, "= A3" и "= SUM(A4,B5)" могут также быть значениями `formula1`.</span><span class="sxs-lookup"><span data-stu-id="cefbc-140">For example, "=A3" and "=SUM(A4,B5)" could also be values of `formula1`.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");

    range.dataValidation.rule = {
            wholeNumber: {
                formula1: 0,
                operator: "GreaterThan"
            }
        };

    return context.sync();
})
```

<span data-ttu-id="cefbc-141">Перечень других бинарных операторов см. в статье [BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation).</span><span class="sxs-lookup"><span data-stu-id="cefbc-141">See [BasicDataValidation](/javascript/api/excel/excel.basicdatavalidation) for a list of the other binary operators.</span></span> 

<span data-ttu-id="cefbc-142">Существует также два тернарных оператора: "Between" и "NotBetween".</span><span class="sxs-lookup"><span data-stu-id="cefbc-142">There are also two ternary operators: "Between" and "NotBetween".</span></span> <span data-ttu-id="cefbc-143">Для их использования необходимо указать необязательное свойство `formula2`. </span><span class="sxs-lookup"><span data-stu-id="cefbc-143">To use these, you must specify the optional `formula2` property.</span></span> <span data-ttu-id="cefbc-144">Значения `formula1` и `formula2` — это ограничивающие операнды.</span><span class="sxs-lookup"><span data-stu-id="cefbc-144">The `formula1` and `formula2` values are the bounding operands.</span></span> <span data-ttu-id="cefbc-145">Значение, которое пользователь пытается ввести в ячейку, — это третий (вычисленный) операнд.</span><span class="sxs-lookup"><span data-stu-id="cefbc-145">The value that the user tries to enter in the cell is the third (evaluated) operand.</span></span> <span data-ttu-id="cefbc-146">Ниже приведены примеры использования оператора "Between":</span><span class="sxs-lookup"><span data-stu-id="cefbc-146">The following is an example of using the "Between" operator:</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");

    range.dataValidation.rule = {
            decimal: {
                formula1: 0,
                formula2: 100,
                operator: "Between"
            }
        };

    return context.sync();
})
```

<span data-ttu-id="cefbc-147">Следующие два свойства правила в качестве своего значения принимают объект [DateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation).</span><span class="sxs-lookup"><span data-stu-id="cefbc-147">The next two rule properties take a [DateTimeDataValidation](/javascript/api/excel/excel.datetimedatavalidation) object as their value.</span></span>

- `date`
- `time`

<span data-ttu-id="cefbc-148">Объект `DateTimeDataValidation` структурирован так же, как и `BasicDataValidation`: он имеет свойства `formula1`, `formula2` и `operator` и используется аналогичным образом.</span><span class="sxs-lookup"><span data-stu-id="cefbc-148">The `DateTimeDataValidation` object is structured similarly to the `BasicDataValidation`: it has the properties `formula1`, `formula2`, and `operator`, and is used in the same way.</span></span> <span data-ttu-id="cefbc-149">Различие состоит в том, что в свойствах формулы нельзя использовать число, но можно ввести строку [даты и времени ISO 8606](https://www.iso.org/iso-8601-date-and-time-format.html) (или формулу Excel).</span><span class="sxs-lookup"><span data-stu-id="cefbc-149">The difference is that you cannot use a number in the formula properties, but you can enter a [ISO 8606 datetime](https://www.iso.org/iso-8601-date-and-time-format.html) string (or an Excel formula).</span></span> <span data-ttu-id="cefbc-150">Ниже приведен пример, в котором определяются допустимые значения для дат в первую неделю апреля 2018 года.</span><span class="sxs-lookup"><span data-stu-id="cefbc-150">The following is an example that defines valid values as dates in the first week of April, 2018.</span></span> 

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");

    range.dataValidation.rule = {
            date: {
                formula1: "2018-04-01",
                formula2: "2018-04-08",
                operator: "Between"
            }
        };

    return context.sync();
})
```

### <a name="list-validation-rule-type"></a><span data-ttu-id="cefbc-151">Тип правила проверки для списка</span><span class="sxs-lookup"><span data-stu-id="cefbc-151">List validation rule type</span></span>

<span data-ttu-id="cefbc-152">Используйте свойство `list` в объекте `DataValidationRule`, чтобы указать, что единственными допустимыми значениями являются значения из конечного списка.</span><span class="sxs-lookup"><span data-stu-id="cefbc-152">Use the `list` property in the `DataValidationRule` object to specify that the only valid values are those from a finite list.</span></span> <span data-ttu-id="cefbc-153">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="cefbc-153">The following is an example.</span></span> <span data-ttu-id="cefbc-154">Обратите внимание на следующие особенности этого кода:</span><span class="sxs-lookup"><span data-stu-id="cefbc-154">Note the following about this code:</span></span>

- <span data-ttu-id="cefbc-155">Предполагается, что существует лист с именем "Имена", а значения в диапазоне "A1:A3" являются именами.</span><span class="sxs-lookup"><span data-stu-id="cefbc-155">It assumes that there is a worksheet named "Names" and that the values in the range "A1:A3" are names.</span></span>
- <span data-ttu-id="cefbc-156">Свойство `source` определяет список допустимых значений.</span><span class="sxs-lookup"><span data-stu-id="cefbc-156">The `source` property specifies the list of valid values.</span></span> <span data-ttu-id="cefbc-157">Строковый аргумент ссылается на диапазон с именами.</span><span class="sxs-lookup"><span data-stu-id="cefbc-157">The string argument refers to a range containing the names.</span></span> <span data-ttu-id="cefbc-158">Можно также назначить разделенный запятыми список, например "Регина, Сергей, Анна".</span><span class="sxs-lookup"><span data-stu-id="cefbc-158">You can also assign a comma-delimited list; for example: "Sue, Ricky, Liz".</span></span> 
- <span data-ttu-id="cefbc-159">Свойство `inCellDropDown` указывает, будет ли раскрывающийся элемент управления отображаться в ячейке, когда пользователь выбирает ее.</span><span class="sxs-lookup"><span data-stu-id="cefbc-159">The `inCellDropDown` property specifies whether a drop-down control will appear in the cell when the user selects it.</span></span> <span data-ttu-id="cefbc-160">Если свойству присвоено значение `true`, то раскрывающийся список отображается со списком значений из `source`.</span><span class="sxs-lookup"><span data-stu-id="cefbc-160">If set to `true`, then the drop-down appears with the list of values from the `source`.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");   
    var nameSourceRange = context.workbook.worksheets.getItem("Names").getRange("A1:A3");

    range.dataValidation.rule = {
        list: {
            inCellDropDown: true,
            source: "=Names!$A$1:$A$3"
        }
    };

    return context.sync();
})
```

### <a name="custom-validation-rule-type"></a><span data-ttu-id="cefbc-161">Настраиваемый тип правила проверки</span><span class="sxs-lookup"><span data-stu-id="cefbc-161">Custom validation rule type</span></span>

<span data-ttu-id="cefbc-162">Используйте свойство `custom` в объекте `DataValidationRule`, чтобы задать настраиваемую формулу проверки.</span><span class="sxs-lookup"><span data-stu-id="cefbc-162">Use the `custom` property in the `DataValidationRule` object to specify a custom validation formula.</span></span> <span data-ttu-id="cefbc-163">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="cefbc-163">The following is an example.</span></span> <span data-ttu-id="cefbc-164">Обратите внимание на следующие особенности этого кода:</span><span class="sxs-lookup"><span data-stu-id="cefbc-164">Note the following about this code:</span></span>

- <span data-ttu-id="cefbc-165">Предполагается, что на листе расположена таблица с двумя столбцами A и B: **Имя спортсмена** и **Комментарии**.</span><span class="sxs-lookup"><span data-stu-id="cefbc-165">It assumes there is a two-column table with columns **Athlete Name** and **Comments** in the A and B columns of the worksheet.</span></span>
- <span data-ttu-id="cefbc-166">Чтобы исключить многословие в столбце **Комментарии**, данные, содержащие имя спортсмена, определяются недопустимыми.</span><span class="sxs-lookup"><span data-stu-id="cefbc-166">To reduce verbosity in the **Comments** column, it makes data that includes the athlete's name invalid.</span></span>
- <span data-ttu-id="cefbc-167">`SEARCH(A2,B2)` возвращает стартовую позицию строки в ячейке A2 в строку в ячейке B2.</span><span class="sxs-lookup"><span data-stu-id="cefbc-167">`SEARCH(A2,B2)` returns the starting position, in string in B2, of the string in A2.</span></span> <span data-ttu-id="cefbc-168">Если A2 не находится в ячейке B2, не возвращается числовое значение.</span><span class="sxs-lookup"><span data-stu-id="cefbc-168">If A2 is not contained in B2, it does not return a number.</span></span> <span data-ttu-id="cefbc-169">`ISNUMBER()` возвращает логическое значение.</span><span class="sxs-lookup"><span data-stu-id="cefbc-169">`ISNUMBER()` returns a boolean.</span></span> <span data-ttu-id="cefbc-170">Поэтому свойство `formula` указывает, что допустимые данные для столбца **Комментарии** — это данные, которые не содержат строку в столбце **Имя спортсмена**.</span><span class="sxs-lookup"><span data-stu-id="cefbc-170">So the `formula` property says that valid data for the **Comment** column is data that does not include the string in the **Athlete Name** column.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");
    var commentsRange = sheet.tables.getItem("AthletesTable").columns.getItem("Comments").getDataBodyRange();

    commentsRange.dataValidation.rule = {
            custom: {
                formula: "=NOT(ISNUMBER(SEARCH(A2,B2)))"
            }
        };

    return context.sync();
})
```

## <a name="create-validation-error-alerts"></a><span data-ttu-id="cefbc-171">Создание оповещений об ошибках проверки</span><span class="sxs-lookup"><span data-stu-id="cefbc-171">Create validation error alerts</span></span>

<span data-ttu-id="cefbc-172">Вы можете создать настраиваемое оповещение об ошибке, которое отображается, если пользователь пытается ввести недопустимые данные в ячейке.</span><span class="sxs-lookup"><span data-stu-id="cefbc-172">You can a create custom error alert that appears when a user tries to enter invalid data in a cell.</span></span> <span data-ttu-id="cefbc-173">Ниже приведен простой пример.</span><span class="sxs-lookup"><span data-stu-id="cefbc-173">The following is a simple example.</span></span> <span data-ttu-id="cefbc-174">Обратите внимание на следующие особенности этого кода:</span><span class="sxs-lookup"><span data-stu-id="cefbc-174">Note the following about this code:</span></span>

- <span data-ttu-id="cefbc-175">Свойство `style` определяет, получает ли пользователь информационное уведомление, предупреждение или оповещение "stop".</span><span class="sxs-lookup"><span data-stu-id="cefbc-175">The `style` property determines whether the user gets an informational alert, a warning, or a "stop" alert.</span></span> <span data-ttu-id="cefbc-176">Только `Stop` действительно не позволяет пользователю добавлять недопустимые данные. </span><span class="sxs-lookup"><span data-stu-id="cefbc-176">Only `Stop` actually prevents the user from adding invalid data.</span></span> <span data-ttu-id="cefbc-177">Всплывающее окно для `Warning` и `Information` содержит параметры, позволяющие пользователю в любом случае ввести недопустимые данные.</span><span class="sxs-lookup"><span data-stu-id="cefbc-177">The pop-up for `Warning` and `Information` has options that allow the user enter the invalid data anyway.</span></span>
- <span data-ttu-id="cefbc-178">Свойству `showAlert` по умолчанию присвоено значение `true`. </span><span class="sxs-lookup"><span data-stu-id="cefbc-178">The `showAlert` property defaults to `true`.</span></span> <span data-ttu-id="cefbc-179">Это означает, что узел Excel будет выдавать всплывающее окно с универсальным оповещением (типа `Stop`), если вы не создадите настраиваемое оповещение, которое либо устанавливает для `showAlert` значение `false`, либо задает настраиваемое сообщение, заголовок и стиль. </span><span class="sxs-lookup"><span data-stu-id="cefbc-179">This means that the Excel host will pop-up a generic alert (of type `Stop`) unless you create a custom alert which either sets `showAlert` to `false` or sets a custom message, title, and style.</span></span> <span data-ttu-id="cefbc-180">Этот код задает настраиваемое сообщение и заголовок.</span><span class="sxs-lookup"><span data-stu-id="cefbc-180">This code sets a custom message and title.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");

    range.dataValidation.errorAlert = {
            message: "Sorry, only positive whole numbers are allowed",
            showAlert: true, // default is 'true'
            style: "Stop", // other possible values: Warning, Information
            title: "Negative or Decimal Number Entered"
        };

    // Set range.dataValidation.rule and optionally .prompt here.

    return context.sync();
})
```

<span data-ttu-id="cefbc-181">Дополнительные сведения см. в статье [DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert).</span><span class="sxs-lookup"><span data-stu-id="cefbc-181">For more information, see [DataValidationErrorAlert](/javascript/api/excel/excel.datavalidationerroralert).</span></span>

## <a name="create-validation-prompts"></a><span data-ttu-id="cefbc-182">Создание запросов проверки</span><span class="sxs-lookup"><span data-stu-id="cefbc-182">Create validation prompts</span></span>

<span data-ttu-id="cefbc-183">Вы можете создать пояснительную подсказку, которая появляется, когда пользователь наводит указатель мыши на ячейку, к которой была применена проверка данных, или выбирает ее.</span><span class="sxs-lookup"><span data-stu-id="cefbc-183">You can create an instructional prompt that appears when a user hovers over, or selects, a cell to which data validation has been applied.</span></span> <span data-ttu-id="cefbc-184">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="cefbc-184">The following is an example:</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");

    range.dataValidation.prompt = {
            message: "Please enter a positive whole number.",
            showPrompt: true, // default is 'false'
            title: "Positive Whole Numbers Only."
        };

    // Set range.dataValidation.rule and optionally .errorAlert here.

    return context.sync();
})
```

<span data-ttu-id="cefbc-185">Дополнительные сведения см. в статье [DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt).</span><span class="sxs-lookup"><span data-stu-id="cefbc-185">For more information, see [DataValidationPrompt](/javascript/api/excel/excel.datavalidationprompt).</span></span>

## <a name="remove-data-validation-from-a-range"></a><span data-ttu-id="cefbc-186">Удаление проверки данных из диапазона</span><span class="sxs-lookup"><span data-stu-id="cefbc-186">Remove data validation from a range</span></span>

<span data-ttu-id="cefbc-187">Чтобы удалить проверку данных из диапазона, вызовите метод [Range.dataValidation.clear()](/javascript/api/excel/excel.datavalidation#clear--).</span><span class="sxs-lookup"><span data-stu-id="cefbc-187">To remove data validation from a range, call the  [Range.dataValidation.clear()](/javascript/api/excel/excel.datavalidation#clear--) method.</span></span>

```js
myrange.dataValidation.clear()
```

<span data-ttu-id="cefbc-188">Необязательно, чтобы очищаемый диапазон был тем же диапазоном, к которому вы применили проверку данных.</span><span class="sxs-lookup"><span data-stu-id="cefbc-188">It isn't necessary that the range you clear is exactly the same range as a range on which you added data validation.</span></span> <span data-ttu-id="cefbc-189">Если это не один и тот же диапазон, удаляются только перекрывающиеся ячейки двух диапазонов (при их наличии).</span><span class="sxs-lookup"><span data-stu-id="cefbc-189">If it isn't, only the overlapping cells, if any, of the two ranges are cleared.</span></span> 

> [!NOTE]
> <span data-ttu-id="cefbc-190">Удаление проверки данных из диапазона также распространяется на любую проверку данных, которую пользователь добавил вручную в диапазон.</span><span class="sxs-lookup"><span data-stu-id="cefbc-190">Clearing data validation from a range will also clear any data validation that a user has added manually to the range.</span></span>

## <a name="see-also"></a><span data-ttu-id="cefbc-191">См. также</span><span class="sxs-lookup"><span data-stu-id="cefbc-191">See also</span></span>

- [<span data-ttu-id="cefbc-192">Основные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="cefbc-192">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="cefbc-193">Объект DataValidation (API JavaScript для Excel)</span><span class="sxs-lookup"><span data-stu-id="cefbc-193">DataValidation Object (JavaScript API for Excel)</span></span>](/javascript/api/excel/excel.datavalidation)
- [<span data-ttu-id="cefbc-194">Объект Range (API JavaScript для Excel)</span><span class="sxs-lookup"><span data-stu-id="cefbc-194">Range Object (JavaScript API for Excel)</span></span>](/javascript/api/excel/excel.range)
