---
ms.date: 04/29/2020
description: Узнайте, как использовать различные параметры в пользовательских функциях, таких как диапазоны Excel, необязательные параметры, контекст вызова и многое другое.
title: Параметры для пользовательских функций Excel
localization_priority: Normal
ms.openlocfilehash: ee193ed68ef59bfd9068bc43cd30721d6bb7b86a
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609277"
---
# <a name="custom-functions-parameter-options"></a><span data-ttu-id="96702-103">Параметры параметров пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="96702-103">Custom functions parameter options</span></span>

<span data-ttu-id="96702-104">Настраиваемые функции можно настраивать с помощью множества различных параметров.</span><span class="sxs-lookup"><span data-stu-id="96702-104">Custom functions are configurable with many different parameter options.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="optional-parameters"></a><span data-ttu-id="96702-105">Необязательные параметры</span><span class="sxs-lookup"><span data-stu-id="96702-105">Optional parameters</span></span>

<span data-ttu-id="96702-106">Если пользователь вызывает функцию в Excel, необязательные параметры отображаются в квадратных скобках.</span><span class="sxs-lookup"><span data-stu-id="96702-106">When a user invokes a function in Excel, optional parameters appear in brackets.</span></span> <span data-ttu-id="96702-107">В приведенном ниже примере функция Add может дополнительно добавить третий номер.</span><span class="sxs-lookup"><span data-stu-id="96702-107">In the following sample, the add function can optionally add a third number.</span></span> <span data-ttu-id="96702-108">Эта функция отображается как `=CONTOSO.ADD(first, second, [third])` в Excel.</span><span class="sxs-lookup"><span data-stu-id="96702-108">This function appears as `=CONTOSO.ADD(first, second, [third])` in Excel.</span></span>

#### <a name="javascript"></a>[<span data-ttu-id="96702-109">JavaScript</span><span class="sxs-lookup"><span data-stu-id="96702-109">JavaScript</span></span>](#tab/javascript)

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

#### <a name="typescript"></a>[<span data-ttu-id="96702-110">TypeScript</span><span class="sxs-lookup"><span data-stu-id="96702-110">TypeScript</span></span>](#tab/typescript)

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
> <span data-ttu-id="96702-111">Если для необязательного параметра не указано значение, Excel присваивает ему значение `null` .</span><span class="sxs-lookup"><span data-stu-id="96702-111">When no value is specified for an optional parameter, Excel assigns it the value `null`.</span></span> <span data-ttu-id="96702-112">Это означает, что параметры, инициализированные по умолчанию в TypeScript, не будут работать должным образом.</span><span class="sxs-lookup"><span data-stu-id="96702-112">This means default-initialized parameters in TypeScript will not work as expected.</span></span> <span data-ttu-id="96702-113">Не следует использовать синтаксис, `function add(first:number, second:number, third=0):number` так как он не инициализируется `third` до 0.</span><span class="sxs-lookup"><span data-stu-id="96702-113">Don't use the syntax `function add(first:number, second:number, third=0):number` because it will not initialize `third` to 0.</span></span> <span data-ttu-id="96702-114">Вместо этого используйте синтаксис TypeScript, как показано в предыдущем примере.</span><span class="sxs-lookup"><span data-stu-id="96702-114">Instead use the TypeScript syntax as shown in the previous example.</span></span>

<span data-ttu-id="96702-115">При определении функции, которая содержит один или несколько необязательных параметров, укажите, что происходит, если необязательные параметры имеют значение null.</span><span class="sxs-lookup"><span data-stu-id="96702-115">When you define a function that contains one or more optional parameters, specify what happens when the optional parameters are null.</span></span> <span data-ttu-id="96702-116">В приведенном ниже примере `zipCode` и `dayOfWeek` являются необязательными параметрами для функции `getWeatherReport`.</span><span class="sxs-lookup"><span data-stu-id="96702-116">In the following example, `zipCode` and `dayOfWeek` are both optional parameters for the `getWeatherReport` function.</span></span> <span data-ttu-id="96702-117">Если `zipCode` параметр имеет значение null, для него устанавливается значение по умолчанию `98052` .</span><span class="sxs-lookup"><span data-stu-id="96702-117">If the `zipCode` parameter is null, the default value is set to `98052`.</span></span> <span data-ttu-id="96702-118">Если `dayOfWeek` параметр имеет значение null, ему присваивается значение среда.</span><span class="sxs-lookup"><span data-stu-id="96702-118">If the `dayOfWeek` parameter is null, it's set to Wednesday.</span></span>

#### <a name="javascript"></a>[<span data-ttu-id="96702-119">JavaScript</span><span class="sxs-lookup"><span data-stu-id="96702-119">JavaScript</span></span>](#tab/javascript)

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

#### <a name="typescript"></a>[<span data-ttu-id="96702-120">TypeScript</span><span class="sxs-lookup"><span data-stu-id="96702-120">TypeScript</span></span>](#tab/typescript)

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

## <a name="range-parameters"></a><span data-ttu-id="96702-121">Параметры Range</span><span class="sxs-lookup"><span data-stu-id="96702-121">Range parameters</span></span>

<span data-ttu-id="96702-122">Настраиваемая функция может принимать диапазон данных ячейки в качестве входного параметра.</span><span class="sxs-lookup"><span data-stu-id="96702-122">Your custom function may accept a range of cell data as an input parameter.</span></span> <span data-ttu-id="96702-123">Функция также может возвращать диапазон данных.</span><span class="sxs-lookup"><span data-stu-id="96702-123">A function can also return a range of data.</span></span> <span data-ttu-id="96702-124">Excel передает диапазон данных ячейки в виде двумерного массива.</span><span class="sxs-lookup"><span data-stu-id="96702-124">Excel will pass a range of cell data as a two-dimensional array.</span></span>

<span data-ttu-id="96702-125">Например, предположим, что функция возвращает второе по величине значение из диапазона значений, хранящихся в Excel.</span><span class="sxs-lookup"><span data-stu-id="96702-125">For example, suppose that your function returns the second highest value from a range of numbers stored in Excel.</span></span> <span data-ttu-id="96702-126">Приведенная ниже функция принимает параметр `values`, относящийся к типу `Excel.CustomFunctionDimensionality.matrix`.</span><span class="sxs-lookup"><span data-stu-id="96702-126">The following function accepts the parameter `values`, which is of type `Excel.CustomFunctionDimensionality.matrix`.</span></span> <span data-ttu-id="96702-127">Обратите внимание, что в метаданных JSON для этой функции для `type` Свойства параметра задано значение `matrix` .</span><span class="sxs-lookup"><span data-stu-id="96702-127">Note that in the JSON metadata for this function, the parameter's `type` property is set to `matrix`.</span></span>

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

## <a name="repeating-parameters"></a><span data-ttu-id="96702-128">Повторяющиеся параметры</span><span class="sxs-lookup"><span data-stu-id="96702-128">Repeating parameters</span></span>

<span data-ttu-id="96702-129">Повторяющийся параметр позволяет пользователю ввести ряд необязательных аргументов функции.</span><span class="sxs-lookup"><span data-stu-id="96702-129">A repeating parameter allows a user to enter a series of optional arguments to a function.</span></span> <span data-ttu-id="96702-130">При вызове функции значения задаются в массиве для параметра.</span><span class="sxs-lookup"><span data-stu-id="96702-130">When the function is called, the values are provided in an array for the parameter.</span></span> <span data-ttu-id="96702-131">Если имя параметра заканчивается числом, каждый номер аргумента будет увеличиваться инкрементно, например `ADD(number1, [number2], [number3],…)` .</span><span class="sxs-lookup"><span data-stu-id="96702-131">If the parameter name ends with a number, each argument's number will increase incrementally, such as `ADD(number1, [number2], [number3],…)`.</span></span> <span data-ttu-id="96702-132">Это соответствует соглашению, используемому для встроенных функций Excel.</span><span class="sxs-lookup"><span data-stu-id="96702-132">This matches the convention used for built-in Excel functions.</span></span>

<span data-ttu-id="96702-133">Приведенная ниже функция суммирует сумму чисел, адресов ячеек, а также диапазонов, если они введены.</span><span class="sxs-lookup"><span data-stu-id="96702-133">The following function sums the total of numbers, cell addresses, as well as ranges, if entered.</span></span>

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

<span data-ttu-id="96702-134">Эта функция отображается `=CONTOSO.ADD([operands], [operands]...)` в книге Excel.</span><span class="sxs-lookup"><span data-stu-id="96702-134">This function shows `=CONTOSO.ADD([operands], [operands]...)` in the Excel workbook.</span></span>

<img alt="The ADD custom function being entered into cell of an Excel worksheet" src="../images/operands.png" />

### <a name="repeating-single-value-parameter"></a><span data-ttu-id="96702-135">Повторяющийся параметр с одним значением</span><span class="sxs-lookup"><span data-stu-id="96702-135">Repeating single value parameter</span></span>

<span data-ttu-id="96702-136">Повторяющийся одиночный параметр значения позволяет передавать несколько отдельных значений.</span><span class="sxs-lookup"><span data-stu-id="96702-136">A repeating single value parameter allows multiple single values to be passed.</span></span> <span data-ttu-id="96702-137">Например, пользователь может ввести ADD (1, B2, 3).</span><span class="sxs-lookup"><span data-stu-id="96702-137">For example, the user could enter ADD(1,B2,3).</span></span> <span data-ttu-id="96702-138">В следующем примере показано, как объявить параметр с одним значением.</span><span class="sxs-lookup"><span data-stu-id="96702-138">The following sample shows how to declare a single value parameter.</span></span>

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

### <a name="single-range-parameter"></a><span data-ttu-id="96702-139">Один параметр Range</span><span class="sxs-lookup"><span data-stu-id="96702-139">Single range parameter</span></span>

<span data-ttu-id="96702-140">Один параметр range технически не является повторяющимся параметром, но он включен здесь, так как объявление очень похоже на повторяющиеся параметры.</span><span class="sxs-lookup"><span data-stu-id="96702-140">A single range parameter isn't technically a repeating parameter, but is included here because the declaration is very similar to repeating parameters.</span></span> <span data-ttu-id="96702-141">Она будет выглядеть как ADD (a2: B3), где один диапазон передается из Excel.</span><span class="sxs-lookup"><span data-stu-id="96702-141">It would appear to the user as ADD(A2:B3) where a single range is passed from Excel.</span></span> <span data-ttu-id="96702-142">В следующем примере показано, как объявить один параметр Range.</span><span class="sxs-lookup"><span data-stu-id="96702-142">The following sample shows how to declare a single range parameter.</span></span>

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

### <a name="repeating-range-parameter"></a><span data-ttu-id="96702-143">Параметр повторяющегося диапазона</span><span class="sxs-lookup"><span data-stu-id="96702-143">Repeating range parameter</span></span>

<span data-ttu-id="96702-144">Параметр повторяющегося диапазона позволяет передавать несколько диапазонов или номеров.</span><span class="sxs-lookup"><span data-stu-id="96702-144">A repeating range parameter allows multiple ranges or numbers to be passed.</span></span> <span data-ttu-id="96702-145">Например, пользователь может ввести ADD (5, B2, C3, 8, No5: E8).</span><span class="sxs-lookup"><span data-stu-id="96702-145">For example, the user could enter ADD(5,B2,C3,8,E5:E8).</span></span> <span data-ttu-id="96702-146">Повторяющиеся диапазоны обычно указываются с типом, `number[][][]` так как они представляют собой трехмерные матрицы.</span><span class="sxs-lookup"><span data-stu-id="96702-146">Repeating ranges are usually specified with the type `number[][][]` as they are three-dimensional matrices.</span></span> <span data-ttu-id="96702-147">Пример приведен в основном примере для повторяющихся параметров (#repeating-Parameters).</span><span class="sxs-lookup"><span data-stu-id="96702-147">For a sample, see the main sample listed for repeating parameters(#repeating-parameters).</span></span>


### <a name="declaring-repeating-parameters"></a><span data-ttu-id="96702-148">Объявление повторяющихся параметров</span><span class="sxs-lookup"><span data-stu-id="96702-148">Declaring repeating parameters</span></span>
<span data-ttu-id="96702-149">В typescript укажите, что параметр является многомерным.</span><span class="sxs-lookup"><span data-stu-id="96702-149">In Typescript, indicate that the parameter is multi-dimensional.</span></span> <span data-ttu-id="96702-150">Например, `ADD(values: number[])` указывает на одномерный массив, который указывает на `ADD(values:number[][])` двухмерный массив и т. д.</span><span class="sxs-lookup"><span data-stu-id="96702-150">For example,  `ADD(values: number[])` would indicate a one-dimensional array, `ADD(values:number[][])` would indicate a two-dimensional array, and so on.</span></span>

<span data-ttu-id="96702-151">В JavaScript используйте одномерные `@param values {number[]}` массивы, `@param <name> {number[][]}` для двумерных массивов и т. д. для дополнительных измерений.</span><span class="sxs-lookup"><span data-stu-id="96702-151">In JavaScript, use `@param values {number[]}` for one-dimensional arrays, `@param <name> {number[][]}` for two-dimensional arrays, and so on for more dimensions.</span></span>

<span data-ttu-id="96702-152">Для созданного вручную JSON убедитесь, что параметр указан как `"repeating": true` в файле JSON, а также проверьте, что параметры помечены как `"dimensionality": matrix` .</span><span class="sxs-lookup"><span data-stu-id="96702-152">For hand-authored JSON, ensure your parameter is specified as `"repeating": true` in your JSON file, as well as check that your parameters are marked as `"dimensionality": matrix`.</span></span>

## <a name="invocation-parameter"></a><span data-ttu-id="96702-153">Параметр вызова</span><span class="sxs-lookup"><span data-stu-id="96702-153">Invocation parameter</span></span>

<span data-ttu-id="96702-154">Каждая пользовательская функция автоматически передает `invocation` аргумент в качестве последнего аргумента.</span><span class="sxs-lookup"><span data-stu-id="96702-154">Every custom function is automatically passed an `invocation` argument as the last argument.</span></span> <span data-ttu-id="96702-155">Этот аргумент можно использовать для получения дополнительного контекста, например адреса вызывающей ячейки.</span><span class="sxs-lookup"><span data-stu-id="96702-155">This argument can be used to retrieve additional context, such as the address of the calling cell.</span></span> <span data-ttu-id="96702-156">Или его можно использовать для отправки в Excel данных, например обработчика функции для [отмены функции](custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="96702-156">Or it can be used to send information to Excel, such as a function handler for [canceling a function](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> <span data-ttu-id="96702-157">Даже если вы не объявили параметры, у настраиваемой функции есть этот параметр.</span><span class="sxs-lookup"><span data-stu-id="96702-157">Even if you declare no parameters, your custom function has this parameter.</span></span> <span data-ttu-id="96702-158">Этот аргумент не отображается для пользователя в Excel.</span><span class="sxs-lookup"><span data-stu-id="96702-158">This argument doesn't appear for a user in Excel.</span></span> <span data-ttu-id="96702-159">Если вы хотите использовать `invocation` пользовательскую функцию, объявите ее в качестве последнего параметра.</span><span class="sxs-lookup"><span data-stu-id="96702-159">If you want to use `invocation` in your custom function, declare it as the last parameter.</span></span>

<span data-ttu-id="96702-160">В следующем примере кода `invocation` контекст явно указывается для ссылки.</span><span class="sxs-lookup"><span data-stu-id="96702-160">In the following code sample, the `invocation` context is explicitly stated for your reference.</span></span>

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

## <a name="next-steps"></a><span data-ttu-id="96702-161">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="96702-161">Next steps</span></span>

<span data-ttu-id="96702-162">Сведения о том, как использовать [переменные значения в пользовательских функциях](custom-functions-volatile.md).</span><span class="sxs-lookup"><span data-stu-id="96702-162">Learn how to use [volatile values in your custom functions](custom-functions-volatile.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="96702-163">См. также</span><span class="sxs-lookup"><span data-stu-id="96702-163">See also</span></span>

* [<span data-ttu-id="96702-164">Получение и обработка данных с помощью пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="96702-164">Receive and handle data with custom functions</span></span>](custom-functions-web-reqs.md)
* [<span data-ttu-id="96702-165">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="96702-165">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="96702-166">Автоматическое генерирование метаданных JSON для пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="96702-166">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="96702-167">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="96702-167">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="96702-168">Руководство по пользовательским функциям в Excel</span><span class="sxs-lookup"><span data-stu-id="96702-168">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
