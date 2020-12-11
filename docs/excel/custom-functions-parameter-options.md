---
ms.date: 12/09/2020
description: Узнайте, как использовать различные параметры в пользовательских функциях, такие как диапазоны Excel, необязательные параметры, контекст вызовов и другие.
title: Параметры пользовательских функций Excel
localization_priority: Normal
ms.openlocfilehash: 9f43955324c148a0af030fb796b82f6d72f429c5
ms.sourcegitcommit: b300e63a96019bdcf5d9f856497694dbd24bfb11
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/11/2020
ms.locfileid: "49624668"
---
# <a name="custom-functions-parameter-options"></a><span data-ttu-id="acac9-103">Параметры пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="acac9-103">Custom functions parameter options</span></span>

<span data-ttu-id="acac9-104">Настраиваемые функции можно настраивать с помощью множества различных параметров.</span><span class="sxs-lookup"><span data-stu-id="acac9-104">Custom functions are configurable with many different parameter options.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="optional-parameters"></a><span data-ttu-id="acac9-105">Необязательные параметры</span><span class="sxs-lookup"><span data-stu-id="acac9-105">Optional parameters</span></span>

<span data-ttu-id="acac9-106">Если пользователь вызывает функцию в Excel, необязательные параметры отображаются в квадратных скобках.</span><span class="sxs-lookup"><span data-stu-id="acac9-106">When a user invokes a function in Excel, optional parameters appear in brackets.</span></span> <span data-ttu-id="acac9-107">В следующем примере функция добавления при желании может добавить третий номер.</span><span class="sxs-lookup"><span data-stu-id="acac9-107">In the following sample, the add function can optionally add a third number.</span></span> <span data-ttu-id="acac9-108">Эта функция отображается, как `=CONTOSO.ADD(first, second, [third])` в Excel.</span><span class="sxs-lookup"><span data-stu-id="acac9-108">This function appears as `=CONTOSO.ADD(first, second, [third])` in Excel.</span></span>

#### <a name="javascript"></a>[<span data-ttu-id="acac9-109">JavaScript</span><span class="sxs-lookup"><span data-stu-id="acac9-109">JavaScript</span></span>](#tab/javascript)

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

#### <a name="typescript"></a>[<span data-ttu-id="acac9-110">TypeScript</span><span class="sxs-lookup"><span data-stu-id="acac9-110">TypeScript</span></span>](#tab/typescript)

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
> <span data-ttu-id="acac9-111">Если для необязательного параметра не задано значение, Excel назначает ему `null` значение.</span><span class="sxs-lookup"><span data-stu-id="acac9-111">When no value is specified for an optional parameter, Excel assigns it the value `null`.</span></span> <span data-ttu-id="acac9-112">Это означает, что параметры, инициализированные по умолчанию в TypeScript, не будут работать ожидаемым образом.</span><span class="sxs-lookup"><span data-stu-id="acac9-112">This means default-initialized parameters in TypeScript will not work as expected.</span></span> <span data-ttu-id="acac9-113">Не используйте синтаксис, так как он не будет инициализироваться `function add(first:number, second:number, third=0):number` `third` до 0.</span><span class="sxs-lookup"><span data-stu-id="acac9-113">Don't use the syntax `function add(first:number, second:number, third=0):number` because it will not initialize `third` to 0.</span></span> <span data-ttu-id="acac9-114">Вместо этого используйте синтаксис TypeScript, как показано в предыдущем примере.</span><span class="sxs-lookup"><span data-stu-id="acac9-114">Instead use the TypeScript syntax as shown in the previous example.</span></span>

<span data-ttu-id="acac9-115">При указании функции, которая содержит один или несколько необязательных параметров, укажите, что происходит, если необязательные параметры имеют null.</span><span class="sxs-lookup"><span data-stu-id="acac9-115">When you define a function that contains one or more optional parameters, specify what happens when the optional parameters are null.</span></span> <span data-ttu-id="acac9-116">В приведенном ниже примере `zipCode` и `dayOfWeek` являются необязательными параметрами для функции `getWeatherReport`.</span><span class="sxs-lookup"><span data-stu-id="acac9-116">In the following example, `zipCode` and `dayOfWeek` are both optional parameters for the `getWeatherReport` function.</span></span> <span data-ttu-id="acac9-117">Если параметр `zipCode` имеет значение NULL, по умолчанию задано значение `98052` .</span><span class="sxs-lookup"><span data-stu-id="acac9-117">If the `zipCode` parameter is null, the default value is set to `98052`.</span></span> <span data-ttu-id="acac9-118">Если параметр `dayOfWeek` имеет null, ему задана среда.</span><span class="sxs-lookup"><span data-stu-id="acac9-118">If the `dayOfWeek` parameter is null, it's set to Wednesday.</span></span>

#### <a name="javascript"></a>[<span data-ttu-id="acac9-119">JavaScript</span><span class="sxs-lookup"><span data-stu-id="acac9-119">JavaScript</span></span>](#tab/javascript)

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

#### <a name="typescript"></a>[<span data-ttu-id="acac9-120">TypeScript</span><span class="sxs-lookup"><span data-stu-id="acac9-120">TypeScript</span></span>](#tab/typescript)

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

## <a name="range-parameters"></a><span data-ttu-id="acac9-121">Параметры range</span><span class="sxs-lookup"><span data-stu-id="acac9-121">Range parameters</span></span>

<span data-ttu-id="acac9-122">Пользовательская функция может принимать диапазон данных ячейки в качестве входного параметра.</span><span class="sxs-lookup"><span data-stu-id="acac9-122">Your custom function may accept a range of cell data as an input parameter.</span></span> <span data-ttu-id="acac9-123">Функция также может возвращать диапазон данных.</span><span class="sxs-lookup"><span data-stu-id="acac9-123">A function can also return a range of data.</span></span> <span data-ttu-id="acac9-124">Excel передает диапазон данных ячейки в качестве двумерного массива.</span><span class="sxs-lookup"><span data-stu-id="acac9-124">Excel will pass a range of cell data as a two-dimensional array.</span></span>

<span data-ttu-id="acac9-125">Например, предположим, что функция возвращает второе по величине значение из диапазона значений, хранящихся в Excel.</span><span class="sxs-lookup"><span data-stu-id="acac9-125">For example, suppose that your function returns the second highest value from a range of numbers stored in Excel.</span></span> <span data-ttu-id="acac9-126">Следующая функция принимает параметр, а синтаксис JSDOC задает свойство параметра в метаданных `values` `number[][]` `dimensionality` `matrix` JSON для этой функции.</span><span class="sxs-lookup"><span data-stu-id="acac9-126">The following function accepts the parameter `values`, and the JSDOC syntax `number[][]` sets the parameter's `dimensionality` property to `matrix` in the JSON metadata for this function.</span></span> 

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

## <a name="repeating-parameters"></a><span data-ttu-id="acac9-127">Повторяющиеся параметры</span><span class="sxs-lookup"><span data-stu-id="acac9-127">Repeating parameters</span></span>

<span data-ttu-id="acac9-128">Повторяюющийся параметр позволяет пользователю ввести ряд необязательных аргументов в функцию.</span><span class="sxs-lookup"><span data-stu-id="acac9-128">A repeating parameter allows a user to enter a series of optional arguments to a function.</span></span> <span data-ttu-id="acac9-129">Когда функция вызвана, значения предоставляются в массиве для параметра.</span><span class="sxs-lookup"><span data-stu-id="acac9-129">When the function is called, the values are provided in an array for the parameter.</span></span> <span data-ttu-id="acac9-130">Если имя параметра заканчивается числом, число каждого аргумента увеличивается постепенно, например `ADD(number1, [number2], [number3],…)` .</span><span class="sxs-lookup"><span data-stu-id="acac9-130">If the parameter name ends with a number, each argument's number will increase incrementally, such as `ADD(number1, [number2], [number3],…)`.</span></span> <span data-ttu-id="acac9-131">Это соответствует соглашению, используемого для встроенных функций Excel.</span><span class="sxs-lookup"><span data-stu-id="acac9-131">This matches the convention used for built-in Excel functions.</span></span>

<span data-ttu-id="acac9-132">Следующая функция суммирует сумму чисел, адресов ячеей, а также диапазонов, если они введены.</span><span class="sxs-lookup"><span data-stu-id="acac9-132">The following function sums the total of numbers, cell addresses, as well as ranges, if entered.</span></span>

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

<span data-ttu-id="acac9-133">Эта функция `=CONTOSO.ADD([operands], [operands]...)` показана в книге Excel.</span><span class="sxs-lookup"><span data-stu-id="acac9-133">This function shows `=CONTOSO.ADD([operands], [operands]...)` in the Excel workbook.</span></span>

<img alt="The ADD custom function being entered into cell of an Excel worksheet" src="../images/operands.png" />

### <a name="repeating-single-value-parameter"></a><span data-ttu-id="acac9-134">Повторяюющийся параметр с одним значением</span><span class="sxs-lookup"><span data-stu-id="acac9-134">Repeating single value parameter</span></span>

<span data-ttu-id="acac9-135">Повторяющийся параметр с одним значением позволяет передавать несколько одно значений.</span><span class="sxs-lookup"><span data-stu-id="acac9-135">A repeating single value parameter allows multiple single values to be passed.</span></span> <span data-ttu-id="acac9-136">Например, пользователь может ввести ADD(1,B2,3).</span><span class="sxs-lookup"><span data-stu-id="acac9-136">For example, the user could enter ADD(1,B2,3).</span></span> <span data-ttu-id="acac9-137">В следующем примере показано, как объявить один параметр значения.</span><span class="sxs-lookup"><span data-stu-id="acac9-137">The following sample shows how to declare a single value parameter.</span></span>

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

### <a name="single-range-parameter"></a><span data-ttu-id="acac9-138">Параметр одиночного диапазона</span><span class="sxs-lookup"><span data-stu-id="acac9-138">Single range parameter</span></span>

<span data-ttu-id="acac9-139">С технической точки000 г. один параметр диапазона не является повторяются, но он включен в него, так как объявление очень похоже на повторяющие параметры.</span><span class="sxs-lookup"><span data-stu-id="acac9-139">A single range parameter isn't technically a repeating parameter, but is included here because the declaration is very similar to repeating parameters.</span></span> <span data-ttu-id="acac9-140">Пользователю будет отображаться как ADD(A2:B3), где из Excel передается один диапазон.</span><span class="sxs-lookup"><span data-stu-id="acac9-140">It would appear to the user as ADD(A2:B3) where a single range is passed from Excel.</span></span> <span data-ttu-id="acac9-141">В следующем примере показано, как объявить один параметр диапазона.</span><span class="sxs-lookup"><span data-stu-id="acac9-141">The following sample shows how to declare a single range parameter.</span></span>

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

### <a name="repeating-range-parameter"></a><span data-ttu-id="acac9-142">Параметр повторяют диапазон</span><span class="sxs-lookup"><span data-stu-id="acac9-142">Repeating range parameter</span></span>

<span data-ttu-id="acac9-143">Параметр повторяют диапазон позволяет передавать несколько диапазонов или чисел.</span><span class="sxs-lookup"><span data-stu-id="acac9-143">A repeating range parameter allows multiple ranges or numbers to be passed.</span></span> <span data-ttu-id="acac9-144">Например, пользователь может ввести ADD(5,B2,C3,8,E5:E8).</span><span class="sxs-lookup"><span data-stu-id="acac9-144">For example, the user could enter ADD(5,B2,C3,8,E5:E8).</span></span> <span data-ttu-id="acac9-145">Повторяющиеся диапазоны обычно заданы с типом, так как `number[][][]` они являются трехмерными матрицами.</span><span class="sxs-lookup"><span data-stu-id="acac9-145">Repeating ranges are usually specified with the type `number[][][]` as they are three-dimensional matrices.</span></span> <span data-ttu-id="acac9-146">Пример см. в основном примере, в списке повторяюющихся параметров (#repeating-parameters).</span><span class="sxs-lookup"><span data-stu-id="acac9-146">For a sample, see the main sample listed for repeating parameters(#repeating-parameters).</span></span>


### <a name="declaring-repeating-parameters"></a><span data-ttu-id="acac9-147">Объявление повторяюющихся параметров</span><span class="sxs-lookup"><span data-stu-id="acac9-147">Declaring repeating parameters</span></span>
<span data-ttu-id="acac9-148">В Typescript указать, что параметр многомерный.</span><span class="sxs-lookup"><span data-stu-id="acac9-148">In Typescript, indicate that the parameter is multi-dimensional.</span></span> <span data-ttu-id="acac9-149">Например,  `ADD(values: number[])` можно указать одномерный массив, указать двумерный массив и так `ADD(values:number[][])` далее.</span><span class="sxs-lookup"><span data-stu-id="acac9-149">For example,  `ADD(values: number[])` would indicate a one-dimensional array, `ADD(values:number[][])` would indicate a two-dimensional array, and so on.</span></span>

<span data-ttu-id="acac9-150">В JavaScript используйте одномерные массивы, двумерные массивы и так далее для `@param values {number[]}` `@param <name> {number[][]}` большего размера.</span><span class="sxs-lookup"><span data-stu-id="acac9-150">In JavaScript, use `@param values {number[]}` for one-dimensional arrays, `@param <name> {number[][]}` for two-dimensional arrays, and so on for more dimensions.</span></span>

<span data-ttu-id="acac9-151">Для JSON, от руки, убедитесь, что параметр указан как в файле JSON, а также убедитесь, что параметры `"repeating": true` помечены как `"dimensionality": matrix` .</span><span class="sxs-lookup"><span data-stu-id="acac9-151">For hand-authored JSON, ensure your parameter is specified as `"repeating": true` in your JSON file, as well as check that your parameters are marked as `"dimensionality": matrix`.</span></span>

## <a name="invocation-parameter"></a><span data-ttu-id="acac9-152">Параметр вызовов</span><span class="sxs-lookup"><span data-stu-id="acac9-152">Invocation parameter</span></span>

<span data-ttu-id="acac9-153">Каждая пользовательская функция автоматически передает аргумент `invocation` в качестве последнего аргумента.</span><span class="sxs-lookup"><span data-stu-id="acac9-153">Every custom function is automatically passed an `invocation` argument as the last argument.</span></span> <span data-ttu-id="acac9-154">Этот аргумент можно использовать для получения дополнительного контекста, например адреса вызываемой ячейки.</span><span class="sxs-lookup"><span data-stu-id="acac9-154">This argument can be used to retrieve additional context, such as the address of the calling cell.</span></span> <span data-ttu-id="acac9-155">Или его можно использовать для отправки сведений в Excel, таких как обработитель функции для [отмены функции.](custom-functions-web-reqs.md#make-a-streaming-function)</span><span class="sxs-lookup"><span data-stu-id="acac9-155">Or it can be used to send information to Excel, such as a function handler for [canceling a function](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> <span data-ttu-id="acac9-156">Даже если параметры не объявлены, этот параметр имеется в пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="acac9-156">Even if you declare no parameters, your custom function has this parameter.</span></span> <span data-ttu-id="acac9-157">Этот аргумент не появляется для пользователя в Excel.</span><span class="sxs-lookup"><span data-stu-id="acac9-157">This argument doesn't appear for a user in Excel.</span></span> <span data-ttu-id="acac9-158">Если вы хотите использовать `invocation` настраиваемую функцию, объявите ее в качестве последнего параметра.</span><span class="sxs-lookup"><span data-stu-id="acac9-158">If you want to use `invocation` in your custom function, declare it as the last parameter.</span></span>

<span data-ttu-id="acac9-159">В следующем примере кода контекст `invocation` явно заявим для ссылки.</span><span class="sxs-lookup"><span data-stu-id="acac9-159">In the following code sample, the `invocation` context is explicitly stated for your reference.</span></span>

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

## <a name="next-steps"></a><span data-ttu-id="acac9-160">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="acac9-160">Next steps</span></span>

<span data-ttu-id="acac9-161">Узнайте, как использовать [переменные значения в пользовательских функциях.](custom-functions-volatile.md)</span><span class="sxs-lookup"><span data-stu-id="acac9-161">Learn how to use [volatile values in your custom functions](custom-functions-volatile.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="acac9-162">См. также</span><span class="sxs-lookup"><span data-stu-id="acac9-162">See also</span></span>

* [<span data-ttu-id="acac9-163">Получение и обработка данных с помощью пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="acac9-163">Receive and handle data with custom functions</span></span>](custom-functions-web-reqs.md)
* [<span data-ttu-id="acac9-164">Автоматическое генерирование метаданных JSON для пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="acac9-164">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="acac9-165">Создание метаданных JSON для пользовательских функций вручную</span><span class="sxs-lookup"><span data-stu-id="acac9-165">Manually create JSON metadata for custom functions</span></span>](custom-functions-json.md)
* [<span data-ttu-id="acac9-166">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="acac9-166">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="acac9-167">Руководство по пользовательским функциям в Excel</span><span class="sxs-lookup"><span data-stu-id="acac9-167">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
