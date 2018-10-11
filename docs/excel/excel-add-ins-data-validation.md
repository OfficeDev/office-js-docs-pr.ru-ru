---
title: Добавление проверки данных в диапазоны Excel
description: ''
ms.date: 10/03/2018
ms.openlocfilehash: 9e3aba8d87e84405bb3e1ae35a8d35d60ce8e2b6
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459156"
---
# <a name="add-data-validation-to-excel-ranges"></a><span data-ttu-id="f6567-102">Добавление проверки данных в диапазоны Excel</span><span class="sxs-lookup"><span data-stu-id="f6567-102">Add data validation to Excel ranges</span></span>

<span data-ttu-id="f6567-p101">Библиотека JavaScript Excel предоставляет API, позволяющие вашей надстройке добавлять функцию автоматической проверки данных для таблиц, столбцов, строк и других диапазонов в книге. Чтобы понять принципы и терминологию проверки данных, ознакомьтесь со следующими статьями о том, как пользователи добавляют проверку данных в пользовательском интерфейсе Excel:</span><span class="sxs-lookup"><span data-stu-id="f6567-p101">The Excel JavaScript Library provides APIs to enable your add-in to add automatic data validation to tables, columns, rows, and other ranges in a workbook. To understand the concepts and the terminology of data validation, please see the following articles about how users add data validation through the Excel UI:</span></span>

- [<span data-ttu-id="f6567-105">Применение проверки данных к ячейкам</span><span class="sxs-lookup"><span data-stu-id="f6567-105">Apply data validation to cells</span></span>](https://support.office.com/article/Apply-data-validation-to-cells-29FECBCC-D1B9-42C1-9D76-EFF3CE5F7249)
- [<span data-ttu-id="f6567-106">Подробнее о проверке данных</span><span class="sxs-lookup"><span data-stu-id="f6567-106">More on data validation</span></span>](https://support.office.com/article/More-on-data-validation-f38dee73-9900-4ca6-9301-8a5f6e1f0c4c)
- [<span data-ttu-id="f6567-107">Описание и примеры проверки данных в Excel</span><span class="sxs-lookup"><span data-stu-id="f6567-107">Description and examples of data validation in Excel</span></span>](https://support.microsoft.com/help/211485/description-and-examples-of-data-validation-in-excel)

## <a name="programmatic-control-of-data-validation"></a><span data-ttu-id="f6567-108">Программное управление проверкой данных</span><span class="sxs-lookup"><span data-stu-id="f6567-108">Programmatic control of data validation</span></span>

<span data-ttu-id="f6567-p102">Свойство `Range.dataValidation` , которое принимает объект [DataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation), — это точка входа для программного управления проверкой данных в Excel. Существует пять свойств объекта `DataValidation` :</span><span class="sxs-lookup"><span data-stu-id="f6567-p102">The `Range.dataValidation` property, which takes a [DataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation) object, is the entry point for programmatic control of data validation in Excel. There are five properties to the `DataValidation` object:</span></span>

- <span data-ttu-id="f6567-p103">`rule` — Определяет, какие данные для диапазона являются допустимыми. См. статью [DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule).</span><span class="sxs-lookup"><span data-stu-id="f6567-p103">`rule` &#8212; Defines what constitutes valid data for the range. See [DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule).</span></span>
- <span data-ttu-id="f6567-p104">`errorAlert` — Указывает, появляется ли ошибка, если пользователь вводит недопустимые данные, и определяет текст, название и стиль оповещения, например, **Informational** (информирование), **Warning** (предупреждение), а также **Stop** (остановка). См. статью[DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert).</span><span class="sxs-lookup"><span data-stu-id="f6567-p104">`errorAlert` &#8212; Specifies whether an error pops up if the user enters invalid data, and defines the alert text, title, and style; for example, **Informational**, **Warning**, and **Stop**. See [DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert).</span></span>
- <span data-ttu-id="f6567-p105">`prompt` — Указывает, появляется ли подсказка, когда пользователь наводит курсор мыши на диапазон, и определяет текст подсказки. См. статью[DataValidationPrompt ](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt).</span><span class="sxs-lookup"><span data-stu-id="f6567-p105">`prompt` &#8212; Specifies whether a prompt appears when the user hovers over the range and defines the prompt message. See [DataValidationPrompt](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt).</span></span>
- <span data-ttu-id="f6567-p106">`ignoreBlanks` — Указывает, применяется ли правило проверки данных к пустым ячейкам в диапазоне. По умолчанию `true`.</span><span class="sxs-lookup"><span data-stu-id="f6567-p106">`ignoreBlanks` &#8212; Specifies whether the data validation rule applies to blank cells in the range. Defaults to `true`.</span></span>
- <span data-ttu-id="f6567-119">`type` — Идентификация типа проверки "только для чтения", например, WholeNumber, Date, TextLength и т. д. Это свойство устанавливается неявно при установке свойства `rule`.</span><span class="sxs-lookup"><span data-stu-id="f6567-119">`type` &#8212; A read-only identification of the validation type, such as WholeNumber, Date, TextLength, etc. It is set indirectly when you set the `rule` property.</span></span>

> [!NOTE]
> <span data-ttu-id="f6567-p107">Проверка данных, добавляемая программно, ведет себя так же, как проверка данных, добавляемая вручную. В частности, обратите внимание на то, что проверка данных запускается только в том случае, если пользователь вводит значение в ячейку или копирует и вставляет ячейки из другого источника в книге и выбирает вариант вставки **Значения**. Если пользователь копирует ячейку и выполняет простую вставку в диапазон проверки данных, проверка не выполняется.</span><span class="sxs-lookup"><span data-stu-id="f6567-p107">Data validation added programmatically behaves just like manually added data validation. In particular, note that data validation is triggered only if the user directly enters a value into a cell or copies and pastes a cell from elsewhere in the workbook and chooses the **Values** paste option. If the user copies a cell and does a plain paste into a range with data validation, validation is not triggered.</span></span>

## <a name="creating-validation-rules"></a><span data-ttu-id="f6567-123">Создание правил проверки</span><span class="sxs-lookup"><span data-stu-id="f6567-123">Creating validation rules</span></span>

<span data-ttu-id="f6567-p108">Чтобы добавить проверку данных в диапазон, ваш код должен установить свойство `rule`объекта `DataValidation` в `Range.dataValidation`. Это получает объект [DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule), который имеет семь дополнительных свойств. *Максимум одно свойство может присутствовать в любом `DataValidationRule` объекте.* Свойство, которое вы указываете, определяет тип выполняемой проверки.</span><span class="sxs-lookup"><span data-stu-id="f6567-p108">To add data validation to a range, your code must set the `rule` property of the `DataValidation` object in `Range.dataValidation`. This takes a [DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule) object which has seven optional properties. *No more than one of these properties may be present in any `DataValidationRule` object.* The property that you include determines the type of validation.</span></span>

### <a name="basic-and-datetime-validation-rule-types"></a><span data-ttu-id="f6567-128">Типы правил проверки Basic и DateTime</span><span class="sxs-lookup"><span data-stu-id="f6567-128">Basic and DateTime validation rule types</span></span>

<span data-ttu-id="f6567-129">Первые три свойства `DataValidationRule` (т. е. типы правил проверки) в качестве своего значения принимают объект [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation).</span><span class="sxs-lookup"><span data-stu-id="f6567-129">The first three `DataValidationRule` properties (i.e., validation rule types) take a [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation) object as their value.</span></span>

- <span data-ttu-id="f6567-130">`wholeNumber` — Требует целое число в дополнение к другим проверкам, указанным объектом `BasicDataValidation`.</span><span class="sxs-lookup"><span data-stu-id="f6567-130">`wholeNumber` &#8212; Requires a whole number in addition to any other validation specified by the `BasicDataValidation` object.</span></span>
- <span data-ttu-id="f6567-131">`decimal` — Требует десятичное число в дополнение к любой другой проверке, определенной в объекте `BasicDataValidation`.</span><span class="sxs-lookup"><span data-stu-id="f6567-131">`decimal` &#8212; Requires a decimal number in addition to any other validation specified by the `BasicDataValidation` object.</span></span>
- <span data-ttu-id="f6567-132">`textLength` — Применяет сведения проверки объекта `BasicDataValidation` к *длине* значения ячейки.</span><span class="sxs-lookup"><span data-stu-id="f6567-132">`textLength` &#8212; Applies the validation details in the `BasicDataValidation` object to the *length* of the cell's value.</span></span>

<span data-ttu-id="f6567-p109">Ниже приведен пример создания правила проверки. Обратите внимание на следующие аспекты этого кода:</span><span class="sxs-lookup"><span data-stu-id="f6567-p109">Here is an example of creating a validation rule. Note the following about this code:</span></span>

- <span data-ttu-id="f6567-p110"> `operator` — это бинарный оператор "GreaterThan". При использовании бинарного оператора значение, которое пользователь пытается ввести в ячейку, — это левый операнд, а значение, заданное в `formula1\`, — это правый операнд. Поэтому согласно этому правилу только целые числа больше 0 являются допустимыми.</span><span class="sxs-lookup"><span data-stu-id="f6567-p110">The `operator` is the binary operator "GreaterThan". Whenever you use a binary operator, the value that the user tries to enter in the cell is the left-hand operand and the value specified in `formula1` is the right-hand operand. So this rule says that only whole numbers that are greater than 0 are valid.</span></span> 
- <span data-ttu-id="f6567-p111"> `formula1` — это жестко заданное число. Если во время кодирования вы не знаете, какое значение должно быть задано, можно также использовать формулу Excel (в виде строки) для значения. Например, "= A3" и "= SUM(A4,B5)" могут также быть значениями `formula1\`.</span><span class="sxs-lookup"><span data-stu-id="f6567-p111">The `formula1` is a hard-coded number. If you don't know at coding time what the value should be, you can also use an Excel formula (as a string) for the value. For example, "=A3" and "=SUM(A4,B5)" could also be values of `formula1`.</span></span>

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

<span data-ttu-id="f6567-141">Перечень других бинарных операторов см. в статье [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation).</span><span class="sxs-lookup"><span data-stu-id="f6567-141">See [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation) for a list of the other binary operators.</span></span> 

<span data-ttu-id="f6567-p112">Существует также два тернарных оператора: "Between" и "NotBetween". Для их использования необходимо указать необязательное свойство `formula2`. Значения `formula1` и `formula2` — это ограничивающие операнды. Значение, которое пользователь пытается ввести в ячейку, — это третий (вычисленный) операнд. Ниже приведены примеры использования оператора "Between":</span><span class="sxs-lookup"><span data-stu-id="f6567-p112">There are also two ternary operators: "Between" and "NotBetween". To use these, you must specify the optional `formula2` property. The `formula1` and `formula2` values are the bounding operands. The value that the user tries to enter in the cell is the third (evaluated) operand. The following is an example of using the "Between" operator:</span></span>

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

<span data-ttu-id="f6567-147">Следующие два свойства правила в качестве своего значения принимают объект [DateTimeDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datetimedatavalidation).</span><span class="sxs-lookup"><span data-stu-id="f6567-147">The next two rule properties take a [DateTimeDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datetimedatavalidation) object as their value.</span></span>

- `date`
- `time`

<span data-ttu-id="f6567-p113">Объект `DateTimeDataValidation` структурирован так же, как и `BasicDataValidation`: он имеет свойства `formula1`, `formula2` и `operator` и используется аналогичным образом. Различие состоит в том, что в свойствах формулы нельзя использовать число, но можно ввести строку [даты и времени ISO 8606](https://www.iso.org/iso-8601-date-and-time-format.html) (или формулу Excel). Ниже приведен пример, в котором определяются допустимые значения для дат в первую неделю апреля 2018 года.</span><span class="sxs-lookup"><span data-stu-id="f6567-p113">The `DateTimeDataValidation` object is structured similarly to the `BasicDataValidation`: it has the properties `formula1`, `formula2`, and `operator`, and is used in the same way. The difference is that you cannot use a number in the formula properties, but you can enter a [ISO 8606 datetime](https://www.iso.org/iso-8601-date-and-time-format.html) string (or an Excel formula). The following is an example that defines valid values as dates in the first week of April, 2018.</span></span> 

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

### <a name="list-validation-rule-type"></a><span data-ttu-id="f6567-151">Тип правила проверки для списка</span><span class="sxs-lookup"><span data-stu-id="f6567-151">List validation rule type</span></span>

<span data-ttu-id="f6567-p114">Используйте свойство `list` в объекте`DataValidationRule` , чтобы указать, что единственные допустимые значения — это значения из конечного списка. Ниже приведен пример. Обратите внимание на следующие аспекты этого кода:</span><span class="sxs-lookup"><span data-stu-id="f6567-p114">Use the `list` property in the `DataValidationRule` object to specify that the only valid values are those from a finite list. The following is an example. Note the following about this code:</span></span>

- <span data-ttu-id="f6567-155">Предполагается, что существует лист с именем "Имена", а значения в диапазоне "A1: A3" являются именами.</span><span class="sxs-lookup"><span data-stu-id="f6567-155">It assumes that there is a worksheet named "Names" and that the values in the range "A1:A3" are names.</span></span>
- <span data-ttu-id="f6567-p115">Свойство `source` определяет список допустимых значений. Ему был назначен диапазон с именами. Можно также назначить разделенный запятыми список, например, "Сью, Рикки, Лиз".</span><span class="sxs-lookup"><span data-stu-id="f6567-p115">The `source` property specifies the list of valid values. The range with the names has been assigned to it. You can also assign a comma-delimited list; for example: "Sue, Ricky, Liz".</span></span> 
- <span data-ttu-id="f6567-p116">Свойство `inCellDropDown` указывает, будет ли раскрывающийся элемент управления отображаться в ячейке, когда пользователь выбирает ее. Если свойство задано как `true`, то раскрывающийся список отображается со списком значений, полученных из `source`.</span><span class="sxs-lookup"><span data-stu-id="f6567-p116">The `inCellDropDown` property specifies whether a drop-down control will appear in the cell when the user selects it. If set to `true`, then the drop-down appears with the list of values from the `source`.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:C5");   
    var nameSourceRange = context.workbook.worksheets.getItem("Names").getRange("A1:A3");

    range.dataValidation.rule = {
        list: {
            inCellDropDown: true,
            source: nameSourceRange
        }
    };

    return context.sync();
})
```

### <a name="custom-validation-rule-type"></a><span data-ttu-id="f6567-161">Настраиваемый тип правила проверки</span><span class="sxs-lookup"><span data-stu-id="f6567-161">Custom validation rule type</span></span>

<span data-ttu-id="f6567-p117">Используйте свойство `custom` в объекте `DataValidationRule`, чтобы задать настраиваемую формулу проверки. Ниже приведен пример. Обратите внимание на следующие аспекты этого кода:</span><span class="sxs-lookup"><span data-stu-id="f6567-p117">Use the `custom` property in the `DataValidationRule` object to specify a custom validation formula. The following is an example. Note the following about this code:</span></span>

- <span data-ttu-id="f6567-165">Предполагается, что на листе расположена таблица с двумя столбцами A и B: **Имя спортсмена** и **Комментарии**.</span><span class="sxs-lookup"><span data-stu-id="f6567-165">It assumes there is a two-column table with columns **Athlete Name** and **Comments** in the A and B columns of the worksheet.</span></span>
- <span data-ttu-id="f6567-166">Для исключения многословия в столбце **Комментарии** код определяет недопустимыми данные, которые содержат имя спортсмена.</span><span class="sxs-lookup"><span data-stu-id="f6567-166">To reduce verbosity in the **Comments** column, it makes data that includes the athlete's name invalid.</span></span>
- <span data-ttu-id="f6567-p118">`SEARCH(A2,B2)` возвращает стартовую позицию строки в ячейке A2 в строку в ячейке B2. Если A2 не находится в ячейке B2, не возвращается числовое значение. `ISNUMBER()` возвращает значение логического типа. Поэтому свойство `formula` указывает, что допустимые данные для столбца **Комментарии**  — это данные, которые не содержат строку в столбце **Имя атлета** .</span><span class="sxs-lookup"><span data-stu-id="f6567-p118">`SEARCH(A2,B2)` returns the starting position, in string in B2, of the string in A2. If A2 is not contained in B2, it does not return a number. `ISNUMBER()` returns a boolean. So the `formula` property says that valid data for the **Comment** column is data that does not include the string in the **Athlete Name** column.</span></span>

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

## <a name="create-validation-error-alerts"></a><span data-ttu-id="f6567-171">Создание предупреждений об ошибках проверки</span><span class="sxs-lookup"><span data-stu-id="f6567-171">Create validation error alerts</span></span>

<span data-ttu-id="f6567-p119">Вы можете создать настраиваемое оповещение об ошибке, которое отображается, когда пользователь пытается ввести недопустимые данные в ячейке. Ниже приведен простой пример. Обратите внимание на следующие аспекты этого кода:</span><span class="sxs-lookup"><span data-stu-id="f6567-p119">You can a create custom error alert that appears when a user tries to enter invalid data in a cell. The following is a simple example. Note the following about this code:</span></span>

- <span data-ttu-id="f6567-p120">Свойство `style` определяет, получает ли пользователь информационное уведомление, предупреждение или оповещение "stop". Только `Stop` фактически не позволяет пользователю добавлять недопустимые данные. Всплывающее окно для `Warning` и `Information` содержит параметры, позволяющие пользователю в любом случае ввести недопустимые данные.</span><span class="sxs-lookup"><span data-stu-id="f6567-p120">The `style` property determines whether the user gets an informational alert, a warning, or a "stop" alert. Only `Stop` actually prevents the user from adding invalid data. The pop-up for `Warning` and `Information` has options that allow the user enter the invalid data anyway.</span></span>
- <span data-ttu-id="f6567-p121">Свойство `showAlert` по умолчанию `true`. Это означает, что узел Excel будет выдавать всплывающее окно с универсальным оповещением (типа `Stop`), если вы не создадите настраиваемое предупреждение, которое либо определяет `showAlert` как `false`, либо задает настраиваемое сообщение, заголовок и стиль. Этот код задает настраиваемое сообщение и заголовок.</span><span class="sxs-lookup"><span data-stu-id="f6567-p121">The `showAlert` property defaults to `true`. This means that the Excel host will pop-up a generic alert (of type `Stop`) unless you create a custom alert which either sets `showAlert` to `false` or sets a custom message, title, and style. This code sets a custom message and title.</span></span>


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

<span data-ttu-id="f6567-181">Дополнительные сведения см. в статье [DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert).</span><span class="sxs-lookup"><span data-stu-id="f6567-181">For more information, see [](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert).</span></span>

## <a name="create-validation-prompts"></a><span data-ttu-id="f6567-182">Создание запросов проверки</span><span class="sxs-lookup"><span data-stu-id="f6567-182">Create validation prompts</span></span>

<span data-ttu-id="f6567-p122">Вы можете создать пояснительную подсказку, которая появляется, когда пользователь наводит курсор мыши на ячейку, к которой была применена проверка данных, или выбирает ее. Ниже приведен пример:</span><span class="sxs-lookup"><span data-stu-id="f6567-p122">You can create an instructional prompt that appears when a user hovers over, or selects, a cell to which data validation has been applied. The following is an example:</span></span>

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

<span data-ttu-id="f6567-185">Дополнительные сведения см. в статье [DataValidationPrompt](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt).</span><span class="sxs-lookup"><span data-stu-id="f6567-185">For more information, see [](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt).</span></span>

## <a name="remove-data-validation-from-a-range"></a><span data-ttu-id="f6567-186">Удаление проверки данных из диапазона</span><span class="sxs-lookup"><span data-stu-id="f6567-186">Remove data validation from a range</span></span>

<span data-ttu-id="f6567-187">Чтобы удалить проверку данных из диапазона, вызовите метод [Range.dataValidation.clear ()](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation#clear--).</span><span class="sxs-lookup"><span data-stu-id="f6567-187">To remove data validation from a range, call the  [Range.dataValidation.clear()](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation#clear--) method.</span></span>

```js
myrange.dataValidation.clear()
```

<span data-ttu-id="f6567-p123">Не обязательно, чтобы диапазон, который вы очищаете,был тем же диапазоном, к которому вы применили проверку данных. Если это не один и тот же диапазон, удаляются только перекрывающиеся ячейки двух диапазонов (при их наличии).</span><span class="sxs-lookup"><span data-stu-id="f6567-p123">It isn't necessary that the range you clear is exactly the same range as a range on which you added data validation. If it isn't, only the overlapping cells, if any, of the two ranges are cleared.</span></span> 

> [!NOTE]
> <span data-ttu-id="f6567-190">Очистка проверки данных из диапазона также распространяется на любую проверку данных, которую пользователь добавил вручную в диапазон.</span><span class="sxs-lookup"><span data-stu-id="f6567-190">Clearing data validation from a range will also clear any data validation that a user has added manually to the range.</span></span>

## <a name="see-also"></a><span data-ttu-id="f6567-191">См. также</span><span class="sxs-lookup"><span data-stu-id="f6567-191">See also</span></span>

- [<span data-ttu-id="f6567-192">Основные принципы программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="f6567-192">Fundamental programming concepts with the Excel JavaScript API</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="f6567-193">Объект DataValidation (API JavaScript для Excel)</span><span class="sxs-lookup"><span data-stu-id="f6567-193">Chart Object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation)
- [<span data-ttu-id="f6567-194">Объект Range (API JavaScript для Excel)</span><span class="sxs-lookup"><span data-stu-id="f6567-194">Range Object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.range)



 
