---
ms.date: 07/15/2019
description: Узнайте, как использовать различные параметры в пользовательских функциях, таких как диапазоны Excel, необязательные параметры, контекст вызова и многое другое.
title: Параметры для пользовательских функций Excel
localization_priority: Normal
ms.openlocfilehash: e5b75b098d64d5998b0393d5995896f0289337fc
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771426"
---
# <a name="custom-functions-parameter-options"></a><span data-ttu-id="29895-103">Параметры параметров пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="29895-103">Custom functions parameter options</span></span>

<span data-ttu-id="29895-104">Настраиваемые функции можно настраивать с помощью различных параметров.</span><span class="sxs-lookup"><span data-stu-id="29895-104">Custom functions are configurable with many different options for parameters.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="optional-parameters"></a><span data-ttu-id="29895-105">Необязательные параметры</span><span class="sxs-lookup"><span data-stu-id="29895-105">Optional parameters</span></span>

<span data-ttu-id="29895-106">В то время как обычные параметры являются обязательными, необязательные параметры — нет.</span><span class="sxs-lookup"><span data-stu-id="29895-106">Whereas regular parameters are required, optional parameters are not.</span></span> <span data-ttu-id="29895-107">Если пользователь вызывает функцию в Excel, необязательные параметры отображаются в квадратных скобках.</span><span class="sxs-lookup"><span data-stu-id="29895-107">When a user invokes a function in Excel, optional parameters appear in brackets.</span></span> <span data-ttu-id="29895-108">В приведенном ниже примере функция Add может дополнительно добавить третий номер.</span><span class="sxs-lookup"><span data-stu-id="29895-108">In the following sample, the add function can optionally add a third number.</span></span> <span data-ttu-id="29895-109">Эта функция отображается как `=CONTOSO.ADD(first, second, [third])` в Excel.</span><span class="sxs-lookup"><span data-stu-id="29895-109">This function appears as `=CONTOSO.ADD(first, second, [third])` in Excel.</span></span>

#### <a name="javascripttabjavascript"></a>[<span data-ttu-id="29895-110">JavaScript</span><span class="sxs-lookup"><span data-stu-id="29895-110">JavaScript</span></span>](#tab/javascript)

```js
/**
 * Calculates the sum of the specified numbers
 * @customfunction
 * @param {number} first First number.
 * @param {number} second Second number.
 * @param {number} [third] Third number to add. If omitted, third = 0.
 * @returns {number} The sum of the numbers.
 */
function add(first, second, third) {
  if (third === null) {
    third = 0;
  }
  return first + second + third;
}
```

#### <a name="typescripttabtypescript"></a>[<span data-ttu-id="29895-111">TypeScript</span><span class="sxs-lookup"><span data-stu-id="29895-111">TypeScript</span></span>](#tab/typescript)

```typescript
/**
 * Calculates the sum of the specified numbers
 * @customfunction
 * @param first First number.
 * @param second Second number.
 * @param [third] Third number to add. If omitted, third = 0.
 * @returns The sum of the numbers.
 */
function add(first: number, second: number, third?: number): number {
  if (third === null) {
    third = 0;
  }
  return first + second + third;
}
```

---

> [!NOTE]
> <span data-ttu-id="29895-112">Если для необязательного параметра не указано значение, Excel присваивает ему значение `null`.</span><span class="sxs-lookup"><span data-stu-id="29895-112">When no value is specified for an optional parameter, Excel assigns it the value `null`.</span></span> <span data-ttu-id="29895-113">Это означает, что параметры, инициализированные по умолчанию в TypeScript, не будут работать должным образом.</span><span class="sxs-lookup"><span data-stu-id="29895-113">This means default-initialized parameters in TypeScript will not work as expected.</span></span> <span data-ttu-id="29895-114">Поэтому не следует использовать синтаксис `function add(first:number, second:number, third=0):number` , так как он не инициализируется `third` до 0.</span><span class="sxs-lookup"><span data-stu-id="29895-114">Therefore, don't use the syntax `function add(first:number, second:number, third=0):number` because it will not initialize `third` to 0.</span></span> <span data-ttu-id="29895-115">Вместо этого используйте синтаксис TypeScript, как показано в предыдущем примере.</span><span class="sxs-lookup"><span data-stu-id="29895-115">Instead use the TypeScript syntax as shown in the previous example.</span></span>

<span data-ttu-id="29895-116">При определении функции, которая содержит один или несколько необязательных параметров, следует указать, что происходит, если необязательные параметры имеют значение null.</span><span class="sxs-lookup"><span data-stu-id="29895-116">When you define a function that contains one or more optional parameters, you should specify what happens when the optional parameters are null.</span></span> <span data-ttu-id="29895-117">В приведенном ниже примере `zipCode` и `dayOfWeek` являются необязательными параметрами для функции `getWeatherReport`.</span><span class="sxs-lookup"><span data-stu-id="29895-117">In the following example, `zipCode` and `dayOfWeek` are both optional parameters for the `getWeatherReport` function.</span></span> <span data-ttu-id="29895-118">Если `zipCode` параметр имеет значение null, для `98052`него устанавливается значение по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="29895-118">If the `zipCode` parameter is null, the default value is set to `98052`.</span></span> <span data-ttu-id="29895-119">Если `dayOfWeek` параметр имеет значение null, ему присваивается значение среда.</span><span class="sxs-lookup"><span data-stu-id="29895-119">If the `dayOfWeek` parameter is null, it is set to Wednesday.</span></span>

#### <a name="javascripttabjavascript"></a>[<span data-ttu-id="29895-120">JavaScript</span><span class="sxs-lookup"><span data-stu-id="29895-120">JavaScript</span></span>](#tab/javascript)

```js
/**
 * Gets a weather report for a specified zipCode and dayOfWeek
 * @customfunction
 * @param {number} [zipCode] Zip code. If omitted, zipCode = 98052.
 * @param {string} [dayOfWeek] Day of the week. If omitted, dayOfWeek = Wednesday.
 * @returns {string} Weather report for the day of the week in that zip code.
 */
function getWeatherReport(zipCode, dayOfWeek) {
  if (zipCode === null) {
    zipCode = 98052;
  }

  if (dayOfWeek === null) {
    dayOfWeek = "Wednesday";
  }

  // Get weather report for specified zipCode and dayOfWeek.
  // ...
}
```

#### <a name="typescripttabtypescript"></a>[<span data-ttu-id="29895-121">TypeScript</span><span class="sxs-lookup"><span data-stu-id="29895-121">TypeScript</span></span>](#tab/typescript)

```typescript
/**
 * Gets a weather report for a specified zipCode and dayOfWeek
 * @customfunction
 * @param zipCode Zip code. If omitted, zipCode = 98052.
 * @param [dayOfWeek] Day of the week. If omitted, dayOfWeek = Wednesday.
 * @returns Weather report for the day of the week in that zip code.
 */
function getWeatherReport(zipCode?: number, dayOfWeek?: string): string {
  if (zipCode === null) {
    zipCode = 98052;
  }

  if (dayOfWeek === null) {
    dayOfWeek = "Wednesday";
  }

  // Get weather report for specified zipCode and dayOfWeek.
  // ...
}
```

---

## <a name="range-parameters"></a><span data-ttu-id="29895-122">Параметры Range</span><span class="sxs-lookup"><span data-stu-id="29895-122">Range parameters</span></span>

<span data-ttu-id="29895-123">Настраиваемая функция может принимать диапазон данных ячейки в качестве входного параметра.</span><span class="sxs-lookup"><span data-stu-id="29895-123">Your custom function may accept a range of cell data as an input parameter.</span></span> <span data-ttu-id="29895-124">Функция также может возвращать диапазон данных.</span><span class="sxs-lookup"><span data-stu-id="29895-124">A function can also return a range of data.</span></span> <span data-ttu-id="29895-125">Excel передает диапазон данных ячейки в виде двумерного массива.</span><span class="sxs-lookup"><span data-stu-id="29895-125">Excel will pass a range of cell data as a two-dimensional array.</span></span>

<span data-ttu-id="29895-126">Например, предположим, что функция возвращает второе по величине значение из диапазона значений, хранящихся в Excel.</span><span class="sxs-lookup"><span data-stu-id="29895-126">For example, suppose that your function returns the second highest value from a range of numbers stored in Excel.</span></span> <span data-ttu-id="29895-127">Приведенная ниже функция принимает параметр `values`, относящийся к типу `Excel.CustomFunctionDimensionality.matrix`.</span><span class="sxs-lookup"><span data-stu-id="29895-127">The following function accepts the parameter `values`, which is of type `Excel.CustomFunctionDimensionality.matrix`.</span></span> <span data-ttu-id="29895-128">Обратите внимание, что в метаданных JSON для этой функции для `type` свойства параметра задано значение `matrix`.</span><span class="sxs-lookup"><span data-stu-id="29895-128">Note that in the JSON metadata for this function, the parameter's `type` property is set to `matrix`.</span></span>

```js
/**
 * Returns the second highest value in a matrixed range of values.
 * @customfunction
 * @param {number[][]} values Multiple ranges of values.
 */
function secondHighest(values) {
  let highest = values[0][0],
    secondHighest = values[0][0];
  for (var i = 0; i < values.length; i++) {
    for (var j = 0; j < values[i].length; j++) {
      if (values[i][j] >= highest) {
        secondHighest = highest;
        highest = values[i][j];
      } else if (values[i][j] >= secondHighest) {
        secondHighest = values[i][j];
      }
    }
  }
  return secondHighest;
}
```

## <a name="repeating-parameters"></a><span data-ttu-id="29895-129">Повторяющиеся параметры</span><span class="sxs-lookup"><span data-stu-id="29895-129">Repeating parameters</span></span>

<span data-ttu-id="29895-130">Повторяющийся параметр позволяет пользователю ввести ряд необязательных аргументов функции.</span><span class="sxs-lookup"><span data-stu-id="29895-130">A repeating parameter allows a user to enter a series of optional of arguments to a function.</span></span> <span data-ttu-id="29895-131">При вызове функции значения задаются в массиве для параметра.</span><span class="sxs-lookup"><span data-stu-id="29895-131">When the function is called, the values are provided in an array for the parameter.</span></span> <span data-ttu-id="29895-132">Если имя параметра заканчивается числом, каждый аргумент увеличит значение, например `ADD(number1, [number2], [number3],…)`.</span><span class="sxs-lookup"><span data-stu-id="29895-132">If the parameter name ends with a number, each argument will increment the number, such as `ADD(number1, [number2], [number3],…)`.</span></span> <span data-ttu-id="29895-133">Это соответствует соглашению, используемому для встроенных функций Excel.</span><span class="sxs-lookup"><span data-stu-id="29895-133">This matches the convention used for built-in Excel functions.</span></span>

<span data-ttu-id="29895-134">Приведенная ниже функция суммирует сумму чисел, адресов ячеек, а также диапазонов, если они введены.</span><span class="sxs-lookup"><span data-stu-id="29895-134">The following function sums the total of numbers, cell addresses, as well as ranges, if entered.</span></span>

```TS
/**
* The sum of all of the numbers.
* @customfunction
* @param operands A number (such as 1 or 3.1415), a cell address (such as A1 or $E$11), or a range of cell addresses (such as B3:F12)
*/

function ADD(operands: number[][][]): number {
  let total: number = 0;

  operands.forEach(range => {
    range.forEach(row => {
      row.forEach(num => {
        total += num;
      });
    });
  });

  return total;
}
```

<span data-ttu-id="29895-135">Эта функция отображается `=CONTOSO.ADD([operands], [operands]...)` в книге Excel.</span><span class="sxs-lookup"><span data-stu-id="29895-135">This function shows `=CONTOSO.ADD([operands], [operands]...)` in the Excel workbook.</span></span>

<img alt="The ADD custom function being entered into cell of an Excel worksheet" src="../images/operands.png" />

### <a name="repeating-single-value-parameter"></a><span data-ttu-id="29895-136">Повторяющийся параметр с одним значением</span><span class="sxs-lookup"><span data-stu-id="29895-136">Repeating single value parameter</span></span>

<span data-ttu-id="29895-137">Повторяющийся одиночный параметр значения позволяет передавать несколько отдельных значений.</span><span class="sxs-lookup"><span data-stu-id="29895-137">A repeating single value parameter allows multiple single values to be passed.</span></span> <span data-ttu-id="29895-138">Например, пользователь может ввести ADD (1, B2, 3).</span><span class="sxs-lookup"><span data-stu-id="29895-138">For example, the user could enter ADD(1,B2,3).</span></span> <span data-ttu-id="29895-139">В следующем примере показано, как объявить параметр с одним значением.</span><span class="sxs-lookup"><span data-stu-id="29895-139">The following sample shows how to declare a single value parameter.</span></span>

```JS
/**
 * @customfunction
 * @param {number[]} singleValue An array of numbers that are repeating parameters.
 */
function addSingleValue(singleValue) {
  let total = 0;
  singleValue.forEach(value => {
    total += value;
  })

  return total;
}
```

### <a name="single-range-parameter"></a><span data-ttu-id="29895-140">Один параметр Range</span><span class="sxs-lookup"><span data-stu-id="29895-140">Single range parameter</span></span>

<span data-ttu-id="29895-141">Один параметр диапазона технически не является повторяющимся параметром, но включается здесь, так как объявление очень похоже на повторяющиеся параметры.</span><span class="sxs-lookup"><span data-stu-id="29895-141">A single range parameter is not technically a repeating parameter, but is included here because the declaration is very similar to repeating parameters.</span></span> <span data-ttu-id="29895-142">Она будет выглядеть как ADD (a2: B3), где один диапазон передается из Excel.</span><span class="sxs-lookup"><span data-stu-id="29895-142">It would appear to the user as ADD(A2:B3) where a single range is passed from Excel.</span></span> <span data-ttu-id="29895-143">В следующем примере показано, как объявить один параметр Range.</span><span class="sxs-lookup"><span data-stu-id="29895-143">The following sample shows how to declare a single range parameter.</span></span>

```JS
/**
 * @customfunction
 * @param {number[][]} singleRange
 */
function addSingleRange(singleRange) {
  let total = 0;
  singleRange.forEach(setOfSingleValues => {
    setOfSingleValues.forEach(value => {
      total += value;
    })
  })
  return total;
}
```

### <a name="repeating-range-parameter"></a><span data-ttu-id="29895-144">Параметр повторяющегося диапазона</span><span class="sxs-lookup"><span data-stu-id="29895-144">Repeating range parameter</span></span>

<span data-ttu-id="29895-145">Параметр повторяющегося диапазона позволяет передавать несколько диапазонов или номеров.</span><span class="sxs-lookup"><span data-stu-id="29895-145">A repeating range parameter allows multiple ranges or numbers to be passed.</span></span> <span data-ttu-id="29895-146">Например, пользователь может ввести ADD (5, B2, C3, 8, No5: E8).</span><span class="sxs-lookup"><span data-stu-id="29895-146">For example, the user could enter ADD(5,B2,C3,8,E5:E8).</span></span> <span data-ttu-id="29895-147">Повторяющиеся диапазоны обычно указываются с `number[][][]` типом, так как они представляют собой трехмерные матрицы.</span><span class="sxs-lookup"><span data-stu-id="29895-147">Repeating ranges are usually specified with the type `number[][][]` as they are three-dimensional matrices.</span></span> <span data-ttu-id="29895-148">Пример приведен в основном примере для повторяющихся параметров (#repeating-Parameters).</span><span class="sxs-lookup"><span data-stu-id="29895-148">For a sample, see the main sample listed for repeating parameters(#repeating-parameters).</span></span>


### <a name="declaring-repeating-parameters"></a><span data-ttu-id="29895-149">Объявление повторяющихся параметров</span><span class="sxs-lookup"><span data-stu-id="29895-149">Declaring repeating parameters</span></span>
<span data-ttu-id="29895-150">В typescript укажите, что параметр является многомерным.</span><span class="sxs-lookup"><span data-stu-id="29895-150">In Typescript, indicate that the parameter is multi-dimensional.</span></span> <span data-ttu-id="29895-151">Например, `ADD(values: number[])` указывает на одномерный массив, `ADD(values:number[][])` который указывает на двухмерный массив и т. д.</span><span class="sxs-lookup"><span data-stu-id="29895-151">For example,  `ADD(values: number[])` would indicate a one-dimensional array, `ADD(values:number[][])` would indicate a two-dimensional array, and so on.</span></span>

<span data-ttu-id="29895-152">В JavaScript используйте `@param values {number[]}` одномерные массивы, `@param <name> {number[][]}` для двумерных массивов и т. д. для дополнительных измерений.</span><span class="sxs-lookup"><span data-stu-id="29895-152">In JavaScript, use `@param values {number[]}` for one-dimensional arrays, `@param <name> {number[][]}` for two-dimensional arrays, and so on for more dimensions.</span></span>

<span data-ttu-id="29895-153">Для созданного вручную JSON убедитесь, что параметр указан как `"repeating": true` в файле JSON, а также проверьте, что параметры помечены как. `"dimensionality”: matrix`</span><span class="sxs-lookup"><span data-stu-id="29895-153">For hand-authored JSON, ensure your parameter is specified as `"repeating": true` in your JSON file, as well as check that your parameters are marked as `"dimensionality”: matrix`.</span></span>

>[!NOTE]
><span data-ttu-id="29895-154">Функции, содержащие повторяющиеся параметры, автоматически содержат параметр вызова в качестве последнего параметра.</span><span class="sxs-lookup"><span data-stu-id="29895-154">Functions containing repeating parameters automatically contain an invocation parameter as the last parameter.</span></span> <span data-ttu-id="29895-155">Дополнительные сведения о параметрах вызова можно найти в следующем разделе.</span><span class="sxs-lookup"><span data-stu-id="29895-155">For more information on invocation parameters, see the following section.</span></span>

## <a name="invocation-parameter"></a><span data-ttu-id="29895-156">Параметр вызова</span><span class="sxs-lookup"><span data-stu-id="29895-156">Invocation parameter</span></span>

<span data-ttu-id="29895-157">Каждая пользовательская функция автоматически передает `invocation` аргумент в качестве последнего аргумента.</span><span class="sxs-lookup"><span data-stu-id="29895-157">Every custom function is automatically passed an `invocation` argument as the last argument.</span></span> <span data-ttu-id="29895-158">Этот аргумент можно использовать для получения дополнительного контекста, например адреса вызывающей ячейки.</span><span class="sxs-lookup"><span data-stu-id="29895-158">This argument can be used to retrieve additional context, such as the address of the calling cell.</span></span> <span data-ttu-id="29895-159">Или его можно использовать для отправки в Excel данных, например обработчика функции для [отмены функции](custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="29895-159">Or it can be used to send information to Excel, such as a function handler for [canceling a function](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> <span data-ttu-id="29895-160">Даже если вы не объявили параметры, у настраиваемой функции есть этот параметр.</span><span class="sxs-lookup"><span data-stu-id="29895-160">Even if you declare no parameters, your custom function has this parameter.</span></span> <span data-ttu-id="29895-161">Этот аргумент не отображается для пользователя в Excel.</span><span class="sxs-lookup"><span data-stu-id="29895-161">This argument doesn't appear for a user in Excel.</span></span> <span data-ttu-id="29895-162">Если вы хотите использовать `invocation` пользовательскую функцию, объявите ее в качестве последнего параметра.</span><span class="sxs-lookup"><span data-stu-id="29895-162">If you want to use `invocation` in your custom function, declare it as the last parameter.</span></span>

<span data-ttu-id="29895-163">В следующем примере кода `invocation` контекст явно указывается для ссылки.</span><span class="sxs-lookup"><span data-stu-id="29895-163">In the following code sample, the `invocation` context is explicitly stated for your reference.</span></span>

```js
/**
 * Add two numbers.
 * @customfunction
 * @param {number} first First number.
 * @param {number} second Second number.
 * @returns {number} The sum of the two (or optionally three) numbers.
 */
function add(first, second, invocation) {
  return first + second;
}
```

<span data-ttu-id="29895-164">Параметр позволяет получить контекст вызывающей ячейки, который может быть полезен в некоторых сценариях, в том числе [Обнаружение адреса ячейки, которая вызывает настраиваемую функцию](#addressing-cells-context-parameter).</span><span class="sxs-lookup"><span data-stu-id="29895-164">The parameter allows you to get the context of the invoking cell, which can be helpful in some scenarios including [discovering the address of a cell which invoke a custom function](#addressing-cells-context-parameter).</span></span>

### <a name="addressing-cells-context-parameter"></a><span data-ttu-id="29895-165">Параметр контекста ячейки адресации</span><span class="sxs-lookup"><span data-stu-id="29895-165">Addressing cell's context parameter</span></span>

<span data-ttu-id="29895-166">В некоторых случаях необходимо получить адрес ячейки, которая вызвала пользовательскую функцию.</span><span class="sxs-lookup"><span data-stu-id="29895-166">In some cases you need to get the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="29895-167">Это полезно в следующих сценариях:</span><span class="sxs-lookup"><span data-stu-id="29895-167">This is useful in the following scenarios:</span></span>

- <span data-ttu-id="29895-168">Диапазоны форматирования: используйте адрес ячейки в качестве ключа для хранения информации в [оффицерунтиме. Storage](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data).</span><span class="sxs-lookup"><span data-stu-id="29895-168">Formatting ranges: Use the cell's address as the key to store information in [OfficeRuntime.storage](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data).</span></span> <span data-ttu-id="29895-169">После этого используйте событие [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) в Excel, чтобы загрузить ключ из `OfficeRuntime.storage`.</span><span class="sxs-lookup"><span data-stu-id="29895-169">Then, use [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) in Excel to load the key from `OfficeRuntime.storage`.</span></span>
- <span data-ttu-id="29895-170">Отображение кэшированных значений. Если функция используется в автономном режиме, отображайте сохраненные в кэше значения из `OfficeRuntime.storage` с помощью `onCalculated`.</span><span class="sxs-lookup"><span data-stu-id="29895-170">Displaying cached values: If your function is used offline, display stored cached values from `OfficeRuntime.storage` using `onCalculated`.</span></span>
- <span data-ttu-id="29895-171">Сверка: используйте адрес ячейки, чтобы найти исходную ячейку, чтобы упростить сверку при выполнении обработки.</span><span class="sxs-lookup"><span data-stu-id="29895-171">Reconciliation: Use the cell's address to discover an origin cell to help you reconcile where processing is occurring.</span></span>

<span data-ttu-id="29895-172">Чтобы запросить контекст ячейки адресации в функции, необходимо использовать функцию для поиска адреса ячейки, например, в приведенном ниже примере.</span><span class="sxs-lookup"><span data-stu-id="29895-172">To request an addressing cell's context in a function, you need to use a function to find the cell's address, such as the one in the following example.</span></span> <span data-ttu-id="29895-173">Сведения об адресе ячейки отображаются только в том случае, `@requiresAddress` если она помечена комментариями функции.</span><span class="sxs-lookup"><span data-stu-id="29895-173">The information about a cell's address is exposed only if `@requiresAddress` is tagged in the function's comments.</span></span>

```js
/**
 * Function that gets the address of a cell.
 * @customfunction
 * @param {CustomFunctions.Invocation} invocation Uses the invocation parameter present in each cell.
 * @requiresAddress
 * @returns {string} Returns address of cell.
 */

function getAddress(invocation) {
  return invocation.address;
}
```

<span data-ttu-id="29895-174">По умолчанию значения, возвращаемые из функции `getAddress`, соответствуют следующему формату: `SheetName!CellNumber`.</span><span class="sxs-lookup"><span data-stu-id="29895-174">By default, values returned from a `getAddress` function follow the following format: `SheetName!CellNumber`.</span></span> <span data-ttu-id="29895-175">Например, если функция вызвана с листа с названием Expenses (Расходы) в ячейке B2, возвращаемым значением будет `Expenses!B2`.</span><span class="sxs-lookup"><span data-stu-id="29895-175">For example, if a function was called from a sheet called Expenses in cell B2, the returned value would be `Expenses!B2`.</span></span>

## <a name="next-steps"></a><span data-ttu-id="29895-176">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="29895-176">Next steps</span></span>

<span data-ttu-id="29895-177">Сведения о том, как [сохранить состояние в пользовательских функциях](custom-functions-save-state.md) или использовать [переменные значения в пользовательских функциях](custom-functions-volatile.md).</span><span class="sxs-lookup"><span data-stu-id="29895-177">Learn how to [save state in your custom functions](custom-functions-save-state.md) or use [volatile values in your custom functions](custom-functions-volatile.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="29895-178">См. также</span><span class="sxs-lookup"><span data-stu-id="29895-178">See also</span></span>

* [<span data-ttu-id="29895-179">Получение и обработка данных с помощью пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="29895-179">Receive and handle data with custom functions</span></span>](custom-functions-web-reqs.md)
* [<span data-ttu-id="29895-180">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="29895-180">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="29895-181">Автоматическое генерирование метаданных JSON для пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="29895-181">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="29895-182">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="29895-182">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="29895-183">Руководство по пользовательским функциям в Excel</span><span class="sxs-lookup"><span data-stu-id="29895-183">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)