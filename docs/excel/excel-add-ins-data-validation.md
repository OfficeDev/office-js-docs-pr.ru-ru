---
title: Добавление проверки данных в диапазоны Excel
description: ''
ms.date: 04/13/2018
ms.openlocfilehash: fd40cab045da0472a060752651a27f0b26028b4b
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2018
ms.locfileid: "23944879"
---
# <a name="add-data-validation-to-excel-ranges-preview"></a><span data-ttu-id="130ad-102">Добавление проверки данных в диапазоны Excel (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="130ad-102">Add data validation to Excel ranges (Preview)</span></span>

> [!NOTE]
> <span data-ttu-id="130ad-103">Пока API проверки данных являются предварительной версией, для их использования вы должны загрузить бета-версию библиотеки JavaScript для Office.</span><span class="sxs-lookup"><span data-stu-id="130ad-103">While the data validation APIs are in preview, you must load the beta version of the Office JavaScript library to use them.</span></span> <span data-ttu-id="130ad-104">URL-адрес: https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span><span class="sxs-lookup"><span data-stu-id="130ad-104">The full URL is https://appsforoffice.microsoft.com/lib/beta/hosted/office.js.</span></span> <span data-ttu-id="130ad-105">Если вы используете TypeScript или редактор кода использует файл определения типа TypeScript для IntelliSense, воспользуйтесь https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span><span class="sxs-lookup"><span data-stu-id="130ad-105">If you are using TypeScript or your code editor uses a TypeScript type definition file for intellisense, use https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts.</span></span>

> [!NOTE]
> <span data-ttu-id="130ad-106">Поскольку API проверки достоверности данных находятся в режиме предварительного просмотра, в этой статье ссылки на API работать не будут.</span><span class="sxs-lookup"><span data-stu-id="130ad-106">While the data validation APIs are in preview, the links in this article to API reference will not work.</span></span> <span data-ttu-id="130ad-107">Тем временем вы можете использовать [черновой вариант справки по API Excel](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec/reference/excel).</span><span class="sxs-lookup"><span data-stu-id="130ad-107">In the meantime, you can use the [draft Excel API reference](https://github.com/OfficeDev/office-js-docs/tree/ExcelJs_OpenSpec/reference/excel).</span></span>

<span data-ttu-id="130ad-108">Библиотека JavaScript Excel предоставляет API, позволяющие вашей надстройке добавлять автоматическую проверку данных в таблицы, столбцы, строки и другие диапазоны в книге.</span><span class="sxs-lookup"><span data-stu-id="130ad-108">The Excel JavaScript Library provides APIs to enable your add-in to add automatic data validation to tables, columns, rows, and other ranges in a workbook.</span></span> <span data-ttu-id="130ad-109">Для ознакомления с понятиями и терминологией проверки данных смотрите следующие статьи о том, как пользователи добавляют проверку данных через интерфейс Excel:</span><span class="sxs-lookup"><span data-stu-id="130ad-109">To understand the concepts and the terminology of data validation, please see the following articles about how users add data validation through the Excel UI:</span></span>

- [<span data-ttu-id="130ad-110">Применение проверки данных к ячейкам</span><span class="sxs-lookup"><span data-stu-id="130ad-110">Apply data validation to cells</span></span>](https://support.office.com/article/Apply-data-validation-to-cells-29FECBCC-D1B9-42C1-9D76-EFF3CE5F7249)
- [<span data-ttu-id="130ad-111">Подробнее о проверке данных</span><span class="sxs-lookup"><span data-stu-id="130ad-111">More on data validation</span></span>](https://support.office.com/article/More-on-data-validation-f38dee73-9900-4ca6-9301-8a5f6e1f0c4c)
- [<span data-ttu-id="130ad-112">Описание и примеры проверки данных в Excel</span><span class="sxs-lookup"><span data-stu-id="130ad-112">Description and examples of data validation in Excel</span></span>](https://support.microsoft.com/help/211485/description-and-examples-of-data-validation-in-excel)

## <a name="programmatic-control-of-data-validation"></a><span data-ttu-id="130ad-113">Программный элемент управления проверкой данных</span><span class="sxs-lookup"><span data-stu-id="130ad-113">Programmatic control of data validation</span></span>

<span data-ttu-id="130ad-114">Свойство `Range.dataValidation`, которое принимает объект [DataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation), является точкой входа для программного управления проверкой данных в Excel.</span><span class="sxs-lookup"><span data-stu-id="130ad-114">The `Range.dataValidation` property, which takes a [DataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation) object, is the entry point for programmatic control of data validation in Excel.</span></span> <span data-ttu-id="130ad-115">Существует пять свойств объекта `DataValidation`:</span><span class="sxs-lookup"><span data-stu-id="130ad-115">There are five properties to the `DataValidation` object:</span></span>

- <span data-ttu-id="130ad-116">`rule` — Определяет, какие данные для диапазона являются достоверными.</span><span class="sxs-lookup"><span data-stu-id="130ad-116">`rule` &#8212; Defines what constitutes valid data for the range.</span></span> <span data-ttu-id="130ad-117">См. [DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule).</span><span class="sxs-lookup"><span data-stu-id="130ad-117">See [DataValidationRule](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationrule).</span></span>
- <span data-ttu-id="130ad-118">`errorAlert` — Указывает, появляется ли ошибка, если пользователь вводит неверные данные, и определяет текст, название и стиль оповещения, например, **Informational** (информирование), **Warning** (предупреждение), а также **Stop** (остановка).</span><span class="sxs-lookup"><span data-stu-id="130ad-118">`errorAlert` &#8212; Specifies whether an error pops up if the user enters invalid data, and defines the alert text, title, and style; for example, **Informational**, **Warning**, and **Stop**.</span></span> <span data-ttu-id="130ad-119">См. [DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert).</span><span class="sxs-lookup"><span data-stu-id="130ad-119">See [DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert).</span></span>
- <span data-ttu-id="130ad-120">`prompt` — Указывает, появляется ли подсказка, когда пользователь наводит курсор мыши на диапазон, и определяет текст подсказки.</span><span class="sxs-lookup"><span data-stu-id="130ad-120">`prompt` &#8212; Specifies whether a prompt appears when the user hovers over the range and defines the prompt message.</span></span> <span data-ttu-id="130ad-121">См. [DataValidationPrompt](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt).</span><span class="sxs-lookup"><span data-stu-id="130ad-121">See [DataValidationPrompt](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt).</span></span>
- <span data-ttu-id="130ad-122">`ignoreBlanks` — Указывает, применяется ли правило проверки данных к пустым ячейкам в диапазоне.</span><span class="sxs-lookup"><span data-stu-id="130ad-122">`ignoreBlanks` &#8212; Specifies whether the data validation rule applies to blank cells in the range.</span></span> <span data-ttu-id="130ad-123">Значение по умолчанию: `true`.</span><span class="sxs-lookup"><span data-stu-id="130ad-123">Defaults to `true`.</span></span>
- <span data-ttu-id="130ad-124">`type` — Идентификация типа проверки "только для чтения", например, WholeNumber, Date, TextLength и т. д. Это свойство устанавливается неявно при установке `rule`.</span><span class="sxs-lookup"><span data-stu-id="130ad-124">`type` &#8212; A read-only identification of the validation type, such as WholeNumber, Date, TextLength, etc. It is set indirectly when you set the `rule` property.</span></span>

> [!NOTE]
> <span data-ttu-id="130ad-125">Добавленная программно проверка данных ведет себя так же, как и добавленная вручную.</span><span class="sxs-lookup"><span data-stu-id="130ad-125">Data validation added programmatically behaves just like manually added data validation.</span></span> <span data-ttu-id="130ad-126">Например, обратите внимание, что проверка данных запускается только в том случае, если пользователь непосредственно вводит значение в ячейку или копирует и вставляет ячейку из другого места в книге с параметром вставки **Значения**.</span><span class="sxs-lookup"><span data-stu-id="130ad-126">In particular, note that data validation is triggered only if the user directly enters a value into a cell or copies and pastes a cell from elsewhere in the workbook and chooses the **Values** paste option.</span></span> <span data-ttu-id="130ad-127">Если пользователь копирует ячейку и выполняет простую вставку в диапазон с проверкой данных, проверка не запускается.</span><span class="sxs-lookup"><span data-stu-id="130ad-127">If the user copies a cell and does a plain paste into a range with data validation, validation is not triggered.</span></span>

### <a name="creating-validation-rules"></a><span data-ttu-id="130ad-128">Создание правил проверки</span><span class="sxs-lookup"><span data-stu-id="130ad-128">Creating validation rules</span></span>

<span data-ttu-id="130ad-129">Чтобы добавить проверку данных в диапазон, в код нужно добавить свойство `rule` объекта `DataValidation` в `Range.dataValidation`.</span><span class="sxs-lookup"><span data-stu-id="130ad-129">To add data validation to a range, your code must set the `rule` property of the `DataValidation` object in `Range.dataValidation`.</span></span> <span data-ttu-id="130ad-130">При этом используется объект [DataValidationRule](https://docs.microsoft.com/javascript/api/excel?view=office-js), который имеет семь необязательных свойств.</span><span class="sxs-lookup"><span data-stu-id="130ad-130">This takes a [DataValidationRule](https://docs.microsoft.com/javascript/api/excel?view=office-js) object which has seven optional properties.</span></span> <span data-ttu-id="130ad-131">*В любом объекте `DataValidationRule` может использоваться не более одного из этих свойств.*</span><span class="sxs-lookup"><span data-stu-id="130ad-131">*No more than one of these properties may be present in any `DataValidationRule` object.*</span></span> <span data-ttu-id="130ad-132">Включенное свойство определяет тип проверки.</span><span class="sxs-lookup"><span data-stu-id="130ad-132">The property that you include determines the type of validation.</span></span>

#### <a name="basic-and-datetime-validation-rule-types"></a><span data-ttu-id="130ad-133">Типы правил проверки Basic и DateTime</span><span class="sxs-lookup"><span data-stu-id="130ad-133">Basic and DateTime validation rule types</span></span>

<span data-ttu-id="130ad-134">Первые три свойства `DataValidationRule` (т. е. типа правил проверки) в качестве своего значения принимают объект [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel?view=office-js).</span><span class="sxs-lookup"><span data-stu-id="130ad-134">The first three `DataValidationRule` properties (i.e., validation rule types) take a [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel?view=office-js) object as their value.</span></span>

- <span data-ttu-id="130ad-135">`wholeNumber` — Требует целое число в дополнение к другим проверкам, определенным в объекте `BasicDataValidation`.</span><span class="sxs-lookup"><span data-stu-id="130ad-135">`wholeNumber` &#8212; Requires a whole number in addition to any other validation specified by the `BasicDataValidation` object.</span></span>
- <span data-ttu-id="130ad-136">`decimal` — Требует десятичное число в дополнение к другим условиям проверки, определенным в объекте `BasicDataValidation`.</span><span class="sxs-lookup"><span data-stu-id="130ad-136">`decimal` &#8212; Requires a decimal number in addition to any other validation specified by the `BasicDataValidation` object.</span></span>
- <span data-ttu-id="130ad-137">`textLength` — Применяет сведения проверки объекта `BasicDataValidation` к *длине* значения ячейки.</span><span class="sxs-lookup"><span data-stu-id="130ad-137">`textLength` &#8212; Applies the validation details in the `BasicDataValidation` object to the *length* of the cell's value.</span></span>

<span data-ttu-id="130ad-138">Вот пример создания правила проверки.</span><span class="sxs-lookup"><span data-stu-id="130ad-138">Here is an example of creating a validation rule.</span></span> <span data-ttu-id="130ad-139">Обратите внимание на следующие особенности этого кода:</span><span class="sxs-lookup"><span data-stu-id="130ad-139">Note the following about this code:</span></span>

- <span data-ttu-id="130ad-140">— это бинарный оператор "GreaterThan".`operator`</span><span class="sxs-lookup"><span data-stu-id="130ad-140">The `operator` is the binary operator "GreaterThan".</span></span> <span data-ttu-id="130ad-141">Всякий раз, когда вы используете бинарный оператор, значение, которое пользователь пытается ввести в ячейку, — это левый операнд, а значение, указанное в `formula1` — правый операнд.</span><span class="sxs-lookup"><span data-stu-id="130ad-141">Whenever you use a binary operator, the value that the user tries to enter in the cell is the left-hand operand and the value specified in `formula1` is the right-hand operand.</span></span> <span data-ttu-id="130ad-142">Таким образом, это правило устанавливает, что действительны только целые числа, которые больше 0.</span><span class="sxs-lookup"><span data-stu-id="130ad-142">So this rule says that only whole numbers that are greater than 0 are valid.</span></span> 
- <span data-ttu-id="130ad-143">— жестко заданное число.`formula1`</span><span class="sxs-lookup"><span data-stu-id="130ad-143">The `formula1` is a hard-coded number.</span></span> <span data-ttu-id="130ad-144">Если во время написания кода вы не знаете, каким должно быть значение, для него можно использовать формулу Excel (как строку).</span><span class="sxs-lookup"><span data-stu-id="130ad-144">If you don't know at coding time what the value should be, you can also use an Excel formula (as a string) for the value.</span></span> <span data-ttu-id="130ad-145">Например, "= A3" и "= SUM (A4, B5)" также могут быть значениями `formula1`.</span><span class="sxs-lookup"><span data-stu-id="130ad-145">For example, "=A3" and "=SUM(A4,B5)" could also be values of `formula1`.</span></span>

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

<span data-ttu-id="130ad-146">Перечень других бинарных операторов см. в статье [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation).</span><span class="sxs-lookup"><span data-stu-id="130ad-146">See [BasicDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.basicdatavalidation) for a list of the other binary operators.</span></span> 

<span data-ttu-id="130ad-147">Есть также два тернарных оператора: "Between" и "NotBetween".</span><span class="sxs-lookup"><span data-stu-id="130ad-147">There are also two ternary operators: "Between" and "NotBetween".</span></span> <span data-ttu-id="130ad-148">Чтобы их использовать, необходимо указать необязательное свойство `formula2`.</span><span class="sxs-lookup"><span data-stu-id="130ad-148">To use these, you must specify the optional `formula2` property.</span></span> <span data-ttu-id="130ad-149">Значения `formula1` и `formula2` — ограничивающие операнды.</span><span class="sxs-lookup"><span data-stu-id="130ad-149">The `formula1` and `formula2` values are the bounding operands.</span></span> <span data-ttu-id="130ad-150">Значение, которое пользователь вводит в ячейку, является третьим (оцениваемым) операндом.</span><span class="sxs-lookup"><span data-stu-id="130ad-150">The value that the user tries to enter in the cell is the third (evaluated) operand.</span></span> <span data-ttu-id="130ad-151">Ниже приведен пример использования оператора "Between":</span><span class="sxs-lookup"><span data-stu-id="130ad-151">The following is an example of using the "Between" operator:</span></span>

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

<span data-ttu-id="130ad-152">Следующие два свойства правила в качестве своего значения принимают объект [DateTimeDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datetimedatavalidation).</span><span class="sxs-lookup"><span data-stu-id="130ad-152">The next two rule properties take a [DateTimeDataValidation](https://docs.microsoft.com/javascript/api/excel/excel.datetimedatavalidation) object as their value.</span></span>

- `date`
- `time`

<span data-ttu-id="130ad-153">Объект `DateTimeDataValidation` структурирован аналогично `BasicDataValidation`: он содержит свойства `formula1`, `formula2` и `operator` и используется таким же образом.</span><span class="sxs-lookup"><span data-stu-id="130ad-153">The `DateTimeDataValidation` object is structured similarly to the `BasicDataValidation`: it has the properties `formula1`, `formula2`, and `operator`, and is used in the same way.</span></span> <span data-ttu-id="130ad-154">Разница состоит в том, что число в свойствах формулы использовать нельзя, но можно ввести строку [с датой и временем по ISO 8606](https://www.iso.org/iso-8601-date-and-time-format.html) (или формулу Excel).</span><span class="sxs-lookup"><span data-stu-id="130ad-154">The difference is that you cannot use a number in the formula properties, but you can enter a [ISO 8606 datetime](https://www.iso.org/iso-8601-date-and-time-format.html) string (or an Excel formula).</span></span> <span data-ttu-id="130ad-155">Ниже приведен пример, в котором определены допустимые значения дат для первой недели апреля 2018 года.</span><span class="sxs-lookup"><span data-stu-id="130ad-155">The following is an example that defines valid values as dates in the first week of April, 2018.</span></span> 

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

#### <a name="list-validation-rule-type"></a><span data-ttu-id="130ad-156">Тип правила проверки List</span><span class="sxs-lookup"><span data-stu-id="130ad-156">List validation rule type</span></span>

<span data-ttu-id="130ad-157">Используйте свойство `list` для объекта `DataValidationRule` для указания того, что единственно допустимыми являются значения из ограниченного списка.</span><span class="sxs-lookup"><span data-stu-id="130ad-157">Use the `list` property in the `DataValidationRule` object to specify that the only valid values are those from a finite list.</span></span> <span data-ttu-id="130ad-158">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="130ad-158">The following is an example.</span></span> <span data-ttu-id="130ad-159">Обратите внимание на следующие особенности этого кода:</span><span class="sxs-lookup"><span data-stu-id="130ad-159">Note the following about this code:</span></span>

- <span data-ttu-id="130ad-160">Предполагается, что существует лист с именем "Names", и значения в диапазоне "A1: A3" являются именами.</span><span class="sxs-lookup"><span data-stu-id="130ad-160">It assumes that there is a worksheet named "Names" and that the values in the range "A1:A3" are names.</span></span>
- <span data-ttu-id="130ad-161">Свойство `source`задает список допустимых значений.</span><span class="sxs-lookup"><span data-stu-id="130ad-161">The `source` property specifies the list of valid values.</span></span> <span data-ttu-id="130ad-162">Ему присвоен диапазон с именами.</span><span class="sxs-lookup"><span data-stu-id="130ad-162">The range with the names has been assigned to it.</span></span> <span data-ttu-id="130ad-163">Также можно назначить список с разделителями-запятыми, например, "Сью, Рики, Лиз".</span><span class="sxs-lookup"><span data-stu-id="130ad-163">You can also assign a comma-delimited list; for example: "Sue, Ricky, Liz".</span></span> 
- <span data-ttu-id="130ad-164">Свойство `inCellDropDown` определяет, будет ли выпадающий элемент управления появляться в ячейке, когда пользователь ее выберет.</span><span class="sxs-lookup"><span data-stu-id="130ad-164">The `inCellDropDown` property specifies whether a drop-down control will appear in the cell when the user selects it.</span></span> <span data-ttu-id="130ad-165">Если установлено значение `true`, появится выпадающий список значений из `source`.</span><span class="sxs-lookup"><span data-stu-id="130ad-165">If set to `true`, then the drop-down appears with the list of values from the `source`.</span></span>

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

#### <a name="custom-validation-rule-type"></a><span data-ttu-id="130ad-166">Тип правила проверки Custom</span><span class="sxs-lookup"><span data-stu-id="130ad-166">Custom validation rule type</span></span>

<span data-ttu-id="130ad-167">Используйте свойство `custom` для объекта `DataValidationRule`, чтобы указать настраиваемую формулу проверки.</span><span class="sxs-lookup"><span data-stu-id="130ad-167">Use the `custom` property in the `DataValidationRule` object to specify a custom validation formula.</span></span> <span data-ttu-id="130ad-168">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="130ad-168">The following is an example.</span></span> <span data-ttu-id="130ad-169">Обратите внимание на следующие особенности этого кода:</span><span class="sxs-lookup"><span data-stu-id="130ad-169">Note the following about this code:</span></span>

- <span data-ttu-id="130ad-170">Предполагается, что на листе расположена таблица с двумя столбцами, A и B: **Athlete Name** (имя спортсмена) и **Comments** .</span><span class="sxs-lookup"><span data-stu-id="130ad-170">It assumes there is a two-column table with columns **Athlete Name** and **Comments** in the A and B columns of the worksheet.</span></span>
- <span data-ttu-id="130ad-171">Для исключения многословия в столбце **Комментарии** код определяет недопустимыми данные, которые содержат имя спортсмена.</span><span class="sxs-lookup"><span data-stu-id="130ad-171">To reduce verbosity in the **Comments** column, it makes data that includes the athlete's name invalid.</span></span>
- <span data-ttu-id="130ad-172">`SEARCH(A2,B2)` возвращает начальную позицию строки в A2 в строке в B2.</span><span class="sxs-lookup"><span data-stu-id="130ad-172">`SEARCH(A2,B2)` returns the starting position, in string in B2, of the string in A2.</span></span> <span data-ttu-id="130ad-173">Если A2 не содержится в B2, число не возвращается.</span><span class="sxs-lookup"><span data-stu-id="130ad-173">If A2 is not contained in B2, it does not return a number.</span></span> <span data-ttu-id="130ad-174">`ISNUMBER()` возвращает логическое значение.</span><span class="sxs-lookup"><span data-stu-id="130ad-174">`ISNUMBER()` returns a boolean.</span></span> <span data-ttu-id="130ad-175">Итак, свойство `formula` говорит, что данные в столбце**Comment** действительны, если в них не включена строка из столбца **Имя атлета**.</span><span class="sxs-lookup"><span data-stu-id="130ad-175">So the `formula` property says that valid data for the **Comment** column is data that does not include the string in the **Athlete Name** column.</span></span>

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

### <a name="create-validation-error-alerts"></a><span data-ttu-id="130ad-176">Создание предупреждений об ошибках проверки</span><span class="sxs-lookup"><span data-stu-id="130ad-176">Create validation error alerts</span></span>

<span data-ttu-id="130ad-177">Вы можете создать настраиваемое предупреждение об ошибке, которое появляется, когда пользователь пытается ввести недопустимые данные в ячейку.</span><span class="sxs-lookup"><span data-stu-id="130ad-177">You can a create custom error alert that appears when a user tries to enter invalid data in a cell.</span></span> <span data-ttu-id="130ad-178">Ниже приведен простой пример.</span><span class="sxs-lookup"><span data-stu-id="130ad-178">The following is a simple example:</span></span> <span data-ttu-id="130ad-179">Обратите внимание на следующие особенности этого кода:</span><span class="sxs-lookup"><span data-stu-id="130ad-179">Note the following about this code:</span></span>

- <span data-ttu-id="130ad-180">Свойство `style` определяет, какое сообщение получит пользователь: alert (оповещение), warning (предупреждение) или "stop" (стоп-оповещение).</span><span class="sxs-lookup"><span data-stu-id="130ad-180">The `style` property determines whether the user gets an informational alert, a warning, or a "stop" alert.</span></span> <span data-ttu-id="130ad-181">Только `Stop` действительно предотвращает добавление пользователем недопустимых данных.</span><span class="sxs-lookup"><span data-stu-id="130ad-181">Only `Stop` actually prevents the user from adding invalid data.</span></span> <span data-ttu-id="130ad-182">Всплывающее окна `Warning` и `Information` обладают параметрами, которые позволяют пользователю все равно ввести недопустимые данные.</span><span class="sxs-lookup"><span data-stu-id="130ad-182">The pop-up for `Warning` and `Information` has options that allow the user enter the invalid data anyway.</span></span>
- <span data-ttu-id="130ad-183">Свойство `showAlert` по умолчанию имеет значение `true`.</span><span class="sxs-lookup"><span data-stu-id="130ad-183">The `showAlert` property defaults to `true`.</span></span> <span data-ttu-id="130ad-184">Это означает, что в ведущем приложении Excel появится общее оповещение (типа `Stop`), если вы не создали настраиваемое оповещение, которое либо устанавливает `showAlert` значение `false`, либо устанавливает настраиваемое сообщение, заголовок и стиль.</span><span class="sxs-lookup"><span data-stu-id="130ad-184">This means that the Excel host will pop-up a generic alert (of type `Stop`) unless you create a custom alert which either sets `showAlert` to `false` or sets a custom message, title, and style.</span></span> <span data-ttu-id="130ad-185">Этот код устанавливает настраиваемое сообщение и заголовок.</span><span class="sxs-lookup"><span data-stu-id="130ad-185">This code sets a custom message and title.</span></span>


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

<span data-ttu-id="130ad-186">Дополнительные сведения см. в статье [DataValidationErrorAlert](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert).</span><span class="sxs-lookup"><span data-stu-id="130ad-186">For more information, see [](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationerroralert).</span></span>

### <a name="create-validation-prompts"></a><span data-ttu-id="130ad-187">Создание запросов проверки</span><span class="sxs-lookup"><span data-stu-id="130ad-187">Create validation prompts</span></span>

<span data-ttu-id="130ad-188">Вы можете создать подсказку, которая появляется, когда пользователь наводит курсор мыши на ячейку, к которой применяется проверка данные, или выбирает ее.</span><span class="sxs-lookup"><span data-stu-id="130ad-188">You can create an instructional prompt that appears when a user hovers over, or selects, a cell to which data validation has been applied.</span></span> <span data-ttu-id="130ad-189">Ниже приведен пример.</span><span class="sxs-lookup"><span data-stu-id="130ad-189">The following is an example:</span></span>

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

<span data-ttu-id="130ad-190">Дополнительные сведения см. в статье [DataValidationPrompt](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt).</span><span class="sxs-lookup"><span data-stu-id="130ad-190">For more information, see [](https://docs.microsoft.com/javascript/api/excel/excel.datavalidationprompt).</span></span>

### <a name="remove-data-validation-from-a-range"></a><span data-ttu-id="130ad-191">Удаление проверки данных из диапазона</span><span class="sxs-lookup"><span data-stu-id="130ad-191">Remove data validation from a range</span></span>

<span data-ttu-id="130ad-192">Чтобы удалить проверку данных из диапазона, вызовите метод [Range.dataValidation.clear ()](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation#clear).</span><span class="sxs-lookup"><span data-stu-id="130ad-192">To remove data validation from a range, call the  [Range.dataValidation.clear()](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation#clear) method.</span></span>

```js
myrange.dataValidation.clear()
```

<span data-ttu-id="130ad-193">Не обязательно, чтобы диапазон, который вы очищаете, полностью совпадал с диапазоном, для которого вы добавили проверку данных.</span><span class="sxs-lookup"><span data-stu-id="130ad-193">It isn't necessary that the range you clear is exactly the same range as a range on which you added data validation.</span></span> <span data-ttu-id="130ad-194">Если они не совпадают, очищаются только из двух диапазонов, которые совпадают.</span><span class="sxs-lookup"><span data-stu-id="130ad-194">If it isn't, only the overlapping cells, if any, of the two ranges are cleared.</span></span> 

> [!NOTE]
> <span data-ttu-id="130ad-195">Очистка проверки данных из диапазона также распространяется на любую проверку данных, которую пользователь добавил вручную в диапазон.</span><span class="sxs-lookup"><span data-stu-id="130ad-195">Clearing data validation from a range will also clear any data validation that a user has added manually to the range.</span></span>

## <a name="see-also"></a><span data-ttu-id="130ad-196">См. также</span><span class="sxs-lookup"><span data-stu-id="130ad-196">See also</span></span>

- [<span data-ttu-id="130ad-197">Основные понятия API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="130ad-197">Excel JavaScript API core concepts</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="130ad-198">Объект DataValidation (API JavaScript для Excel)</span><span class="sxs-lookup"><span data-stu-id="130ad-198">Chart Object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.datavalidation)
- [<span data-ttu-id="130ad-199">Объект Range (API JavaScript для Excel)</span><span class="sxs-lookup"><span data-stu-id="130ad-199">Range Object (JavaScript API for Excel)</span></span>](https://docs.microsoft.com/javascript/api/excel/excel.range)



 
