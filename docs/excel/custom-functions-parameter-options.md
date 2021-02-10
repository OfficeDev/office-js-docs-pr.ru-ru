---
ms.date: 02/04/2021
description: Узнайте, как использовать различные параметры в пользовательских функциях, такие как диапазоны Excel, необязательные параметры, контекст вызовов и другие.
title: Параметры пользовательских функций Excel
localization_priority: Normal
ms.openlocfilehash: afe6947b1a1b9022a0284535b9ab1d68c9777c14
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173908"
---
# <a name="custom-functions-parameter-options"></a><span data-ttu-id="8de3a-103">Параметры пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="8de3a-103">Custom functions parameter options</span></span>

<span data-ttu-id="8de3a-104">Настраиваемые функции можно настраивать с помощью множества различных параметров.</span><span class="sxs-lookup"><span data-stu-id="8de3a-104">Custom functions are configurable with many different parameter options.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="optional-parameters"></a><span data-ttu-id="8de3a-105">Необязательные параметры</span><span class="sxs-lookup"><span data-stu-id="8de3a-105">Optional parameters</span></span>

<span data-ttu-id="8de3a-106">Если пользователь вызывает функцию в Excel, необязательные параметры отображаются в квадратных скобках.</span><span class="sxs-lookup"><span data-stu-id="8de3a-106">When a user invokes a function in Excel, optional parameters appear in brackets.</span></span> <span data-ttu-id="8de3a-107">В следующем примере функция добавления при желании может добавить третий номер.</span><span class="sxs-lookup"><span data-stu-id="8de3a-107">In the following sample, the add function can optionally add a third number.</span></span> <span data-ttu-id="8de3a-108">Эта функция отображается, как `=CONTOSO.ADD(first, second, [third])` в Excel.</span><span class="sxs-lookup"><span data-stu-id="8de3a-108">This function appears as `=CONTOSO.ADD(first, second, [third])` in Excel.</span></span>

#### <a name="javascript"></a>[<span data-ttu-id="8de3a-109">JavaScript</span><span class="sxs-lookup"><span data-stu-id="8de3a-109">JavaScript</span></span>](#tab/javascript)

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

#### <a name="typescript"></a>[<span data-ttu-id="8de3a-110">TypeScript</span><span class="sxs-lookup"><span data-stu-id="8de3a-110">TypeScript</span></span>](#tab/typescript)

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
> <span data-ttu-id="8de3a-111">Если для необязательного параметра не задано значение, Excel назначает ему `null` значение.</span><span class="sxs-lookup"><span data-stu-id="8de3a-111">When no value is specified for an optional parameter, Excel assigns it the value `null`.</span></span> <span data-ttu-id="8de3a-112">Это означает, что инициализированные по умолчанию параметры в TypeScript не будут работать ожидаемым образом.</span><span class="sxs-lookup"><span data-stu-id="8de3a-112">This means default-initialized parameters in TypeScript will not work as expected.</span></span> <span data-ttu-id="8de3a-113">Не используйте синтаксис, так как он не будет инициализироваться `function add(first:number, second:number, third=0):number` `third` до 0.</span><span class="sxs-lookup"><span data-stu-id="8de3a-113">Don't use the syntax `function add(first:number, second:number, third=0):number` because it will not initialize `third` to 0.</span></span> <span data-ttu-id="8de3a-114">Вместо этого используйте синтаксис TypeScript, как показано в предыдущем примере.</span><span class="sxs-lookup"><span data-stu-id="8de3a-114">Instead use the TypeScript syntax as shown in the previous example.</span></span>

<span data-ttu-id="8de3a-115">При указании функции, которая содержит один или несколько необязательных параметров, укажите, что происходит, если необязательные параметры имеют null.</span><span class="sxs-lookup"><span data-stu-id="8de3a-115">When you define a function that contains one or more optional parameters, specify what happens when the optional parameters are null.</span></span> <span data-ttu-id="8de3a-116">В приведенном ниже примере `zipCode` и `dayOfWeek` являются необязательными параметрами для функции `getWeatherReport`.</span><span class="sxs-lookup"><span data-stu-id="8de3a-116">In the following example, `zipCode` and `dayOfWeek` are both optional parameters for the `getWeatherReport` function.</span></span> <span data-ttu-id="8de3a-117">Если параметр `zipCode` имеет значение NULL, по умолчанию задано значение `98052` .</span><span class="sxs-lookup"><span data-stu-id="8de3a-117">If the `zipCode` parameter is null, the default value is set to `98052`.</span></span> <span data-ttu-id="8de3a-118">Если параметр `dayOfWeek` имеет null, ему задана среда.</span><span class="sxs-lookup"><span data-stu-id="8de3a-118">If the `dayOfWeek` parameter is null, it's set to Wednesday.</span></span>

#### <a name="javascript"></a>[<span data-ttu-id="8de3a-119">JavaScript</span><span class="sxs-lookup"><span data-stu-id="8de3a-119">JavaScript</span></span>](#tab/javascript)

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

#### <a name="typescript"></a>[<span data-ttu-id="8de3a-120">TypeScript</span><span class="sxs-lookup"><span data-stu-id="8de3a-120">TypeScript</span></span>](#tab/typescript)

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

## <a name="range-parameters"></a><span data-ttu-id="8de3a-121">Параметры range</span><span class="sxs-lookup"><span data-stu-id="8de3a-121">Range parameters</span></span>

<span data-ttu-id="8de3a-122">Пользовательская функция может принимать диапазон данных ячейки в качестве входного параметра.</span><span class="sxs-lookup"><span data-stu-id="8de3a-122">Your custom function may accept a range of cell data as an input parameter.</span></span> <span data-ttu-id="8de3a-123">Функция также может возвращать диапазон данных.</span><span class="sxs-lookup"><span data-stu-id="8de3a-123">A function can also return a range of data.</span></span> <span data-ttu-id="8de3a-124">Excel передает диапазон данных ячейки в качестве двумерного массива.</span><span class="sxs-lookup"><span data-stu-id="8de3a-124">Excel will pass a range of cell data as a two-dimensional array.</span></span>

<span data-ttu-id="8de3a-125">Например, предположим, что функция возвращает второе по величине значение из диапазона значений, хранящихся в Excel.</span><span class="sxs-lookup"><span data-stu-id="8de3a-125">For example, suppose that your function returns the second highest value from a range of numbers stored in Excel.</span></span> <span data-ttu-id="8de3a-126">Следующая функция принимает параметр, а синтаксис JSDOC задает свойство параметра в метаданных `values` `number[][]` `dimensionality` `matrix` JSON для этой функции.</span><span class="sxs-lookup"><span data-stu-id="8de3a-126">The following function accepts the parameter `values`, and the JSDOC syntax `number[][]` sets the parameter's `dimensionality` property to `matrix` in the JSON metadata for this function.</span></span> 

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

## <a name="repeating-parameters"></a><span data-ttu-id="8de3a-127">Повторяющиеся параметры</span><span class="sxs-lookup"><span data-stu-id="8de3a-127">Repeating parameters</span></span>

<span data-ttu-id="8de3a-128">Повторяюющийся параметр позволяет пользователю ввести ряд необязательных аргументов в функцию.</span><span class="sxs-lookup"><span data-stu-id="8de3a-128">A repeating parameter allows a user to enter a series of optional arguments to a function.</span></span> <span data-ttu-id="8de3a-129">Когда функция вызвана, значения предоставляются в массиве для параметра.</span><span class="sxs-lookup"><span data-stu-id="8de3a-129">When the function is called, the values are provided in an array for the parameter.</span></span> <span data-ttu-id="8de3a-130">Если имя параметра заканчивается числом, число каждого аргумента увеличивается постепенно, например `ADD(number1, [number2], [number3],…)` .</span><span class="sxs-lookup"><span data-stu-id="8de3a-130">If the parameter name ends with a number, each argument's number will increase incrementally, such as `ADD(number1, [number2], [number3],…)`.</span></span> <span data-ttu-id="8de3a-131">Это соответствует соглашению, используемого для встроенных функций Excel.</span><span class="sxs-lookup"><span data-stu-id="8de3a-131">This matches the convention used for built-in Excel functions.</span></span>

<span data-ttu-id="8de3a-132">Следующая функция суммирует сумму чисел, адресов ячеей, а также диапазонов, если они введены.</span><span class="sxs-lookup"><span data-stu-id="8de3a-132">The following function sums the total of numbers, cell addresses, as well as ranges, if entered.</span></span>

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

<span data-ttu-id="8de3a-133">Эта функция `=CONTOSO.ADD([operands], [operands]...)` показана в книге Excel.</span><span class="sxs-lookup"><span data-stu-id="8de3a-133">This function shows `=CONTOSO.ADD([operands], [operands]...)` in the Excel workbook.</span></span>

<img alt="The ADD custom function being entered into cell of an Excel worksheet" src="../images/operands.png" />

### <a name="repeating-single-value-parameter"></a><span data-ttu-id="8de3a-134">Повторяюющийся параметр с одним значением</span><span class="sxs-lookup"><span data-stu-id="8de3a-134">Repeating single value parameter</span></span>

<span data-ttu-id="8de3a-135">Повторяющийся параметр с одним значением позволяет передавать несколько одно значений.</span><span class="sxs-lookup"><span data-stu-id="8de3a-135">A repeating single value parameter allows multiple single values to be passed.</span></span> <span data-ttu-id="8de3a-136">Например, пользователь может ввести ADD(1,B2,3).</span><span class="sxs-lookup"><span data-stu-id="8de3a-136">For example, the user could enter ADD(1,B2,3).</span></span> <span data-ttu-id="8de3a-137">В следующем примере показано, как объявить один параметр значения.</span><span class="sxs-lookup"><span data-stu-id="8de3a-137">The following sample shows how to declare a single value parameter.</span></span>

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

### <a name="single-range-parameter"></a><span data-ttu-id="8de3a-138">Параметр одиночного диапазона</span><span class="sxs-lookup"><span data-stu-id="8de3a-138">Single range parameter</span></span>

<span data-ttu-id="8de3a-139">С технической точки000 г. один параметр диапазона не является повторяются, но он включен в него, так как объявление очень похоже на повторяющие параметры.</span><span class="sxs-lookup"><span data-stu-id="8de3a-139">A single range parameter isn't technically a repeating parameter, but is included here because the declaration is very similar to repeating parameters.</span></span> <span data-ttu-id="8de3a-140">Пользователю будет отображаться как ADD(A2:B3), где из Excel передается один диапазон.</span><span class="sxs-lookup"><span data-stu-id="8de3a-140">It would appear to the user as ADD(A2:B3) where a single range is passed from Excel.</span></span> <span data-ttu-id="8de3a-141">В следующем примере показано, как объявить один параметр диапазона.</span><span class="sxs-lookup"><span data-stu-id="8de3a-141">The following sample shows how to declare a single range parameter.</span></span>

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

### <a name="repeating-range-parameter"></a><span data-ttu-id="8de3a-142">Параметр повторяют диапазон</span><span class="sxs-lookup"><span data-stu-id="8de3a-142">Repeating range parameter</span></span>

<span data-ttu-id="8de3a-143">Параметр повторяют диапазон позволяет передавать несколько диапазонов или чисел.</span><span class="sxs-lookup"><span data-stu-id="8de3a-143">A repeating range parameter allows multiple ranges or numbers to be passed.</span></span> <span data-ttu-id="8de3a-144">Например, пользователь может ввести ADD(5,B2,C3,8,E5:E8).</span><span class="sxs-lookup"><span data-stu-id="8de3a-144">For example, the user could enter ADD(5,B2,C3,8,E5:E8).</span></span> <span data-ttu-id="8de3a-145">Повторяющиеся диапазоны обычно заданы с типом, так как `number[][][]` они являются трехмерными матрицами.</span><span class="sxs-lookup"><span data-stu-id="8de3a-145">Repeating ranges are usually specified with the type `number[][][]` as they are three-dimensional matrices.</span></span> <span data-ttu-id="8de3a-146">Пример см. в основном примере, в списке повторяюющихся параметров (#repeating-parameters).</span><span class="sxs-lookup"><span data-stu-id="8de3a-146">For a sample, see the main sample listed for repeating parameters(#repeating-parameters).</span></span>


### <a name="declaring-repeating-parameters"></a><span data-ttu-id="8de3a-147">Объявление повторяюющихся параметров</span><span class="sxs-lookup"><span data-stu-id="8de3a-147">Declaring repeating parameters</span></span>
<span data-ttu-id="8de3a-148">В Typescript указать, что параметр многомерный.</span><span class="sxs-lookup"><span data-stu-id="8de3a-148">In Typescript, indicate that the parameter is multi-dimensional.</span></span> <span data-ttu-id="8de3a-149">Например,  `ADD(values: number[])` можно указать одномерный массив, указать двумерный массив и так `ADD(values:number[][])` далее.</span><span class="sxs-lookup"><span data-stu-id="8de3a-149">For example,  `ADD(values: number[])` would indicate a one-dimensional array, `ADD(values:number[][])` would indicate a two-dimensional array, and so on.</span></span>

<span data-ttu-id="8de3a-150">В JavaScript используйте одномерные массивы, двумерные массивы и так далее для `@param values {number[]}` `@param <name> {number[][]}` большего размера.</span><span class="sxs-lookup"><span data-stu-id="8de3a-150">In JavaScript, use `@param values {number[]}` for one-dimensional arrays, `@param <name> {number[][]}` for two-dimensional arrays, and so on for more dimensions.</span></span>

<span data-ttu-id="8de3a-151">Для JSON, от руки, убедитесь, что параметр указан как в файле JSON, а также убедитесь, что параметры `"repeating": true` помечены как `"dimensionality": matrix` .</span><span class="sxs-lookup"><span data-stu-id="8de3a-151">For hand-authored JSON, ensure your parameter is specified as `"repeating": true` in your JSON file, as well as check that your parameters are marked as `"dimensionality": matrix`.</span></span>

## <a name="invocation-parameter"></a><span data-ttu-id="8de3a-152">Параметр вызовов</span><span class="sxs-lookup"><span data-stu-id="8de3a-152">Invocation parameter</span></span>

<span data-ttu-id="8de3a-153">Каждая пользовательская функция автоматически передает аргумент в качестве последнего входного параметра, даже если она не `invocation` объявлена явным образом.</span><span class="sxs-lookup"><span data-stu-id="8de3a-153">Every custom function is automatically passed an `invocation` argument as the last input parameter, even if it's not explicitly declared.</span></span> <span data-ttu-id="8de3a-154">Этот `invocation` параметр соответствует объекту [Invocation.](/javascript/api/custom-functions-runtime/customfunctions.invocation)</span><span class="sxs-lookup"><span data-stu-id="8de3a-154">This `invocation` parameter corresponds to the [Invocation](/javascript/api/custom-functions-runtime/customfunctions.invocation) object.</span></span> <span data-ttu-id="8de3a-155">Объект можно использовать для получения дополнительного контекста, например адреса ячейки, которая `Invocation` вызывалась пользовательской функцией.</span><span class="sxs-lookup"><span data-stu-id="8de3a-155">The `Invocation` object can be used to retrieve additional context, such as the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="8de3a-156">Чтобы получить доступ `Invocation` к объекту, необходимо объявить `invocation` как последний параметр в пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="8de3a-156">To access the `Invocation` object, you must declare `invocation` as the last parameter in your custom function.</span></span> 

> [!NOTE]
> <span data-ttu-id="8de3a-157">Параметр не появляется в Excel в качестве `invocation` аргумента пользовательской функции для пользователей.</span><span class="sxs-lookup"><span data-stu-id="8de3a-157">The `invocation` parameter doesn't appear as a custom function argument for users in Excel.</span></span>

<span data-ttu-id="8de3a-158">В следующем примере показано, как использовать параметр для возврата адреса ячейки, которая `invocation` вызывает пользовательскую функцию.</span><span class="sxs-lookup"><span data-stu-id="8de3a-158">The following sample shows how to use the `invocation` parameter to return the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="8de3a-159">В этом примере [используется свойство address](/javascript/api/custom-functions-runtime/customfunctions.invocation#address) `Invocation` объекта.</span><span class="sxs-lookup"><span data-stu-id="8de3a-159">This sample uses the [address](/javascript/api/custom-functions-runtime/customfunctions.invocation#address) property of the `Invocation` object.</span></span> <span data-ttu-id="8de3a-160">Чтобы получить доступ `Invocation` к объекту, сначала `CustomFunctions.Invocation` объявите в качестве параметра в JSDoc.</span><span class="sxs-lookup"><span data-stu-id="8de3a-160">To access the `Invocation` object, first declare `CustomFunctions.Invocation` as a parameter in your JSDoc.</span></span> <span data-ttu-id="8de3a-161">Затем `@requiresAddress` объявите в JSDoc доступ к `address` свойству `Invocation` объекта.</span><span class="sxs-lookup"><span data-stu-id="8de3a-161">Next, declare `@requiresAddress` in your JSDoc to access the `address` property of the `Invocation` object.</span></span> <span data-ttu-id="8de3a-162">Наконец, в функции извлекаете и возвращаете `address` свойство.</span><span class="sxs-lookup"><span data-stu-id="8de3a-162">Finally, within the function, retrieve and then return the `address` property.</span></span> 

```js
/**
 * Return the address of the cell that invoked the custom function. 
 * @customfunction
 * @param {number} first First parameter.
 * @param {number} second Second parameter.
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @requiresAddress 
 */
function getAddress(first, second, invocation) {
  var address = invocation.address;
  return address;
}
```

<span data-ttu-id="8de3a-163">В Excel пользовательская функция, вызываемая свойство объекта, возвращает абсолютный адрес, следующий за форматом в ячейке, вызываемой `address` `Invocation` `SheetName!RelativeCellAddress` функцией.</span><span class="sxs-lookup"><span data-stu-id="8de3a-163">In Excel, a custom function calling the `address` property of the `Invocation` object will return the absolute address following the format `SheetName!RelativeCellAddress` in the cell that invoked the function.</span></span> <span data-ttu-id="8de3a-164">Например, если входной параметр находится на листе с и названием **"Цены"** в ячейке F6, возвращается значение адреса параметра `Prices!F6` .</span><span class="sxs-lookup"><span data-stu-id="8de3a-164">For example, if the input parameter is located on a sheet called **Prices** in cell F6, the returned parameter address value will be `Prices!F6`.</span></span> 

<span data-ttu-id="8de3a-165">Этот `invocation` параметр также можно использовать для отправки сведений в Excel.</span><span class="sxs-lookup"><span data-stu-id="8de3a-165">The `invocation` parameter can also be used to send information to Excel.</span></span> <span data-ttu-id="8de3a-166">Дополнительные [дополнительные данные см.](custom-functions-web-reqs.md#make-a-streaming-function) в подкатовной функции "Сделать потоковой передачей".</span><span class="sxs-lookup"><span data-stu-id="8de3a-166">See [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function) to learn more.</span></span>

## <a name="detect-the-address-of-a-parameter"></a><span data-ttu-id="8de3a-167">Обнаружение адреса параметра</span><span class="sxs-lookup"><span data-stu-id="8de3a-167">Detect the address of a parameter</span></span>

<span data-ttu-id="8de3a-168">В сочетании с [параметром вызовов](#invocation-parameter)можно использовать объект ["Вызов"](/javascript/api/custom-functions-runtime/customfunctions.invocation) для получения адреса параметра ввода пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="8de3a-168">In combination with the [invocation parameter](#invocation-parameter), you can use the [Invocation](/javascript/api/custom-functions-runtime/customfunctions.invocation) object to retrieve the address of a custom function input parameter.</span></span> <span data-ttu-id="8de3a-169">При вызове свойство [parameterAddresses](/javascript/api/custom-functions-runtime/customfunctions.invocation#parameterAddresses) объекта позволяет функции возвращать адреса `Invocation` всех входных параметров.</span><span class="sxs-lookup"><span data-stu-id="8de3a-169">When invoked, the [parameterAddresses](/javascript/api/custom-functions-runtime/customfunctions.invocation#parameterAddresses) property of the `Invocation` object allows a function to return the addresses of all input parameters.</span></span> 

<span data-ttu-id="8de3a-170">Это полезно в сценариях, где типы входных данных могут отличаться.</span><span class="sxs-lookup"><span data-stu-id="8de3a-170">This is useful in scenarios where input data types may vary.</span></span> <span data-ttu-id="8de3a-171">Адрес входного параметра можно использовать для проверки формата номера входного значения.</span><span class="sxs-lookup"><span data-stu-id="8de3a-171">The address of an input parameter can be used to check the number format of the input value.</span></span> <span data-ttu-id="8de3a-172">Формат номера можно при необходимости скорректировать до ввода.</span><span class="sxs-lookup"><span data-stu-id="8de3a-172">The number format can then be adjusted prior to input, if necessary.</span></span> <span data-ttu-id="8de3a-173">Адрес входного параметра также можно использовать, чтобы определить, есть ли у входного значения какие-либо связанные свойства, которые могут быть релевантны для последующих вычислений.</span><span class="sxs-lookup"><span data-stu-id="8de3a-173">The address of an input parameter can also be used to detect whether the input value has any related properties that may be relevant to subsequent calculations.</span></span> 

>[!NOTE]
> <span data-ttu-id="8de3a-174">Если вы работаете с вручную созданными метаданными [JSON](custom-functions-json.md) для возврата адресов параметров вместо генератора Yo Office, для объекта должно быть задано свойство , а для объекта должно быть задано свойство `options` `requiresParameterAddresses` `true` `result` `dimensionality` `matrix` .</span><span class="sxs-lookup"><span data-stu-id="8de3a-174">If you're working with [manually-created JSON metadata](custom-functions-json.md) to return parameter addresses instead of the Yo Office generator, the `options` object must have the `requiresParameterAddresses` property set to `true`, and the `result` object must have the `dimensionality` property set to `matrix`.</span></span>

<span data-ttu-id="8de3a-175">Следующая пользовательская функция принимает три входных параметра, извлекает свойство объекта для каждого параметра и возвращает `parameterAddresses` `Invocation` адреса.</span><span class="sxs-lookup"><span data-stu-id="8de3a-175">The following custom function takes in three input parameters, retrieves the `parameterAddresses` property of the `Invocation` object for each parameter, and then returns the addresses.</span></span> 

```js
/**
 * Return the address of three parameters. 
 * @customfunction
 * @param {string} firstParameter First parameter.
 * @param {string} secondParameter Second parameter.
 * @param {string} thirdParameter Third parameter
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @requiresParameterAddresses
 */
function getParameterAddresses(firstParameter, secondParameter, thirdParameter, invocation) {
  var addresses = [
    [invocation.parameterAddresses[0]],
    [invocation.parameterAddresses[1]],
    [invocation.parameterAddresses[2]]
  ];
  return addresses;
}
```

<span data-ttu-id="8de3a-176">При запуске настраиваемой функции, вызываемой свойством, адрес параметра возвращается в соответствии с форматом в ячейке, `parameterAddresses` `SheetName!RelativeCellAddress` вызываемой функцией.</span><span class="sxs-lookup"><span data-stu-id="8de3a-176">When a custom function calling the `parameterAddresses` property runs, the parameter address is returned following the format `SheetName!RelativeCellAddress` in the cell that invoked the function.</span></span> <span data-ttu-id="8de3a-177">Например, если входной параметр расположен на листе **"Затраты"** в ячейке D8, возвращается значение адреса параметра `Costs!D8` .</span><span class="sxs-lookup"><span data-stu-id="8de3a-177">For example, if the input parameter is located on a sheet called **Costs** in cell D8, the returned parameter address value will be `Costs!D8`.</span></span> <span data-ttu-id="8de3a-178">Если настраиваемая функция имеет несколько параметров и возвращается несколько адресов параметров, возвращаемые адреса будут перетекать по нескольким ячейкам, убывая по вертикали из ячейки, которая вызывает функцию.</span><span class="sxs-lookup"><span data-stu-id="8de3a-178">If the custom function has multiple parameters and more than one parameter address is returned, the returned addresses will spill across multiple cells, descending vertically from the cell that invoked the function.</span></span> 

## <a name="next-steps"></a><span data-ttu-id="8de3a-179">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="8de3a-179">Next steps</span></span>

<span data-ttu-id="8de3a-180">Узнайте, как использовать [переменные значения в пользовательских функциях.](custom-functions-volatile.md)</span><span class="sxs-lookup"><span data-stu-id="8de3a-180">Learn how to use [volatile values in your custom functions](custom-functions-volatile.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="8de3a-181">См. также</span><span class="sxs-lookup"><span data-stu-id="8de3a-181">See also</span></span>

* [<span data-ttu-id="8de3a-182">Получение и обработка данных с помощью пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="8de3a-182">Receive and handle data with custom functions</span></span>](custom-functions-web-reqs.md)
* [<span data-ttu-id="8de3a-183">Автоматическое генерирование метаданных JSON для пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="8de3a-183">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="8de3a-184">Создание метаданных JSON для пользовательских функций вручную</span><span class="sxs-lookup"><span data-stu-id="8de3a-184">Manually create JSON metadata for custom functions</span></span>](custom-functions-json.md)
* [<span data-ttu-id="8de3a-185">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="8de3a-185">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="8de3a-186">Руководство по пользовательским функциям в Excel</span><span class="sxs-lookup"><span data-stu-id="8de3a-186">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
