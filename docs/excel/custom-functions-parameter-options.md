---
ms.date: 07/01/2019
description: Узнайте, как использовать различные параметры в пользовательских функциях, таких как диапазоны Excel, необязательные параметры, контекст вызова и многое другое.
title: Параметры для пользовательских функций Excel
localization_priority: Normal
ms.openlocfilehash: 9416653d697bdf36ca698271e00d9742ff0e75a9
ms.sourcegitcommit: 9c5a836d4464e49846c9795bf44cfe23e9fc8fbe
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/10/2019
ms.locfileid: "35617046"
---
# <a name="custom-functions-parameter-options"></a><span data-ttu-id="c3091-103">Параметры параметров пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="c3091-103">Custom functions parameter options</span></span>

<span data-ttu-id="c3091-104">Настраиваемые функции можно настраивать с помощью различных параметров.</span><span class="sxs-lookup"><span data-stu-id="c3091-104">Custom functions are configurable with many different options for parameters.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="optional-parameters"></a><span data-ttu-id="c3091-105">Необязательные параметры</span><span class="sxs-lookup"><span data-stu-id="c3091-105">Optional parameters</span></span>

<span data-ttu-id="c3091-106">В то время как обычные параметры являются обязательными, необязательные параметры — нет.</span><span class="sxs-lookup"><span data-stu-id="c3091-106">Whereas regular parameters are required, optional parameters are not.</span></span> <span data-ttu-id="c3091-107">Если пользователь вызывает функцию в Excel, необязательные параметры отображаются в квадратных скобках.</span><span class="sxs-lookup"><span data-stu-id="c3091-107">When a user invokes a function in Excel, optional parameters appear in brackets.</span></span> <span data-ttu-id="c3091-108">В приведенном ниже примере функция Add может дополнительно добавить третий номер.</span><span class="sxs-lookup"><span data-stu-id="c3091-108">In the following sample, the add function can optionally add a third number.</span></span> <span data-ttu-id="c3091-109">Эта функция отображается как `=CONTOSO.ADD(first, second, [third])` в Excel.</span><span class="sxs-lookup"><span data-stu-id="c3091-109">This function appears as `=CONTOSO.ADD(first, second, [third])` in Excel.</span></span>

#### <a name="javascripttabjavascript"></a>[<span data-ttu-id="c3091-110">JavaScript</span><span class="sxs-lookup"><span data-stu-id="c3091-110">JavaScript</span></span>](#tab/javascript)

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
CustomFunctions.associate("ADD", add);
```

#### <a name="typescripttabtypescript"></a>[<span data-ttu-id="c3091-111">TypeScript</span><span class="sxs-lookup"><span data-stu-id="c3091-111">TypeScript</span></span>](#tab/typescript)

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
CustomFunctions.associate("ADD", add);
```

---

> [!NOTE]
> <span data-ttu-id="c3091-112">Если для необязательного параметра не указано значение, Excel присваивает ему значение `null`.</span><span class="sxs-lookup"><span data-stu-id="c3091-112">When no value is specified for an optional parameter, Excel assigns it the value `null`.</span></span> <span data-ttu-id="c3091-113">Это означает, что параметры, инициализированные по умолчанию в TypeScript, не будут работать должным образом.</span><span class="sxs-lookup"><span data-stu-id="c3091-113">This means default-initialized parameters in TypeScript will not work as expected.</span></span> <span data-ttu-id="c3091-114">Поэтому не следует использовать синтаксис `function add(first:number, second:number, third=0):number` , так как он не инициализируется `third` до 0.</span><span class="sxs-lookup"><span data-stu-id="c3091-114">Therefore, don't use the syntax `function add(first:number, second:number, third=0):number` because it will not initialize `third` to 0.</span></span> <span data-ttu-id="c3091-115">Вместо этого используйте синтаксис TypeScript, как показано в предыдущем примере.</span><span class="sxs-lookup"><span data-stu-id="c3091-115">Instead use the TypeScript syntax as shown in the previous example.</span></span>

<span data-ttu-id="c3091-116">При определении функции, которая содержит один или несколько необязательных параметров, следует указать, что происходит, если необязательные параметры имеют значение null.</span><span class="sxs-lookup"><span data-stu-id="c3091-116">When you define a function that contains one or more optional parameters, you should specify what happens when the optional parameters are null.</span></span> <span data-ttu-id="c3091-117">В приведенном ниже примере `zipCode` и `dayOfWeek` являются необязательными параметрами для функции `getWeatherReport`.</span><span class="sxs-lookup"><span data-stu-id="c3091-117">In the following example, `zipCode` and `dayOfWeek` are both optional parameters for the `getWeatherReport` function.</span></span> <span data-ttu-id="c3091-118">Если `zipCode` параметр имеет значение null, для `98052`него устанавливается значение по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="c3091-118">If the `zipCode` parameter is null, the default value is set to `98052`.</span></span> <span data-ttu-id="c3091-119">Если `dayOfWeek` параметр имеет значение null, ему присваивается значение среда.</span><span class="sxs-lookup"><span data-stu-id="c3091-119">If the `dayOfWeek` parameter is null, it is set to Wednesday.</span></span>

#### <a name="javascripttabjavascript"></a>[<span data-ttu-id="c3091-120">JavaScript</span><span class="sxs-lookup"><span data-stu-id="c3091-120">JavaScript</span></span>](#tab/javascript)

```js
/**
 * Gets a weather report for a specified zipCode and dayOfWeek
 * @customfunction
 * @param {number} [zipCode] Zip code. If omitted, zipCode = 98052.
 * @param {string} [dayOfWeek] Day of the week. If omitted, dayOfWeek = Wednesday.
 * @returns {string} Weather report for the day of the week in that zip code.
 */
function getWeatherReport(zipCode, dayOfWeek)
{
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

#### <a name="typescripttabtypescript"></a>[<span data-ttu-id="c3091-121">TypeScript</span><span class="sxs-lookup"><span data-stu-id="c3091-121">TypeScript</span></span>](#tab/typescript)

```typescript
/**
 * Gets a weather report for a specified zipCode and dayOfWeek
 * @customfunction
 * @param zipCode Zip code. If omitted, zipCode = 98052.
 * @param [dayOfWeek] Day of the week. If omitted, dayOfWeek = Wednesday.
 * @returns Weather report for the day of the week in that zip code.
 */
function getWeatherReport(zipCode?: number, dayOfWeek?: string): string
{
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

## <a name="range-parameters"></a><span data-ttu-id="c3091-122">Параметры Range</span><span class="sxs-lookup"><span data-stu-id="c3091-122">Range parameters</span></span>

<span data-ttu-id="c3091-123">Настраиваемая функция может принимать диапазон данных ячейки в качестве входного параметра.</span><span class="sxs-lookup"><span data-stu-id="c3091-123">Your custom function may accept a range of cell data as an input parameter.</span></span> <span data-ttu-id="c3091-124">Функция также может возвращать диапазон данных.</span><span class="sxs-lookup"><span data-stu-id="c3091-124">A function can also return a range of data.</span></span> <span data-ttu-id="c3091-125">Excel передает диапазон данных ячейки в виде двумерного массива.</span><span class="sxs-lookup"><span data-stu-id="c3091-125">Excel will pass a range of cell data as a two-dimensional array.</span></span>

<span data-ttu-id="c3091-126">Например, предположим, что функция возвращает второе по величине значение из диапазона значений, хранящихся в Excel.</span><span class="sxs-lookup"><span data-stu-id="c3091-126">For example, suppose that your function returns the second highest value from a range of numbers stored in Excel.</span></span> <span data-ttu-id="c3091-127">Приведенная ниже функция принимает параметр `values`, относящийся к типу `Excel.CustomFunctionDimensionality.matrix`.</span><span class="sxs-lookup"><span data-stu-id="c3091-127">The following function accepts the parameter `values`, which is of type `Excel.CustomFunctionDimensionality.matrix`.</span></span> <span data-ttu-id="c3091-128">Обратите внимание, что в метаданных JSON для этой функции для `type` свойства параметра задано значение `matrix`.</span><span class="sxs-lookup"><span data-stu-id="c3091-128">Note that in the JSON metadata for this function, the parameter's `type` property is set to `matrix`.</span></span>

```js
/**
 * Returns the second highest value in a matrixed range of values.
 * @customfunction
 * @param {number[][]} values Multiple ranges of values.  
 */
function secondHighest(values){
  let highest = values[0][0], secondHighest = values[0][0];
  for(var i = 0; i < values.length; i++){
    for(var j = 0; j < values[i].length; j++){
      if(values[i][j] >= highest){
        secondHighest = highest;
        highest = values[i][j];
      }
      else if(values[i][j] >= secondHighest){
        secondHighest = values[i][j];
      }
    }
  }
  return secondHighest;
}
CustomFunctions.associate("SECONDHIGHEST", secondHighest);
```

## <a name="invocation-parameter"></a><span data-ttu-id="c3091-129">Параметр вызова</span><span class="sxs-lookup"><span data-stu-id="c3091-129">Invocation parameter</span></span>

<span data-ttu-id="c3091-130">Каждая пользовательская функция автоматически передает `invocation` аргумент в качестве последнего аргумента.</span><span class="sxs-lookup"><span data-stu-id="c3091-130">Every custom function is automatically passed an `invocation` argument as the last argument.</span></span> <span data-ttu-id="c3091-131">Этот аргумент можно использовать для получения дополнительного контекста, например адреса вызывающей ячейки.</span><span class="sxs-lookup"><span data-stu-id="c3091-131">This argument can be used to retrieve additional context, such as the address of the calling cell.</span></span> <span data-ttu-id="c3091-132">Или его можно использовать для отправки в Excel данных, например обработчика функции для [отмены функции](custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="c3091-132">Or it can be used to send information to Excel, such as a function handler for [canceling a function](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> <span data-ttu-id="c3091-133">Даже если вы не объявили параметры, у настраиваемой функции есть этот параметр.</span><span class="sxs-lookup"><span data-stu-id="c3091-133">Even if you declare no parameters, your custom function has this parameter.</span></span> <span data-ttu-id="c3091-134">Этот аргумент не отображается для пользователя в Excel.</span><span class="sxs-lookup"><span data-stu-id="c3091-134">This argument doesn't appear for a user in Excel.</span></span> <span data-ttu-id="c3091-135">Если вы хотите использовать `invocation` пользовательскую функцию, объявите ее в качестве последнего параметра.</span><span class="sxs-lookup"><span data-stu-id="c3091-135">If you want to use `invocation` in your custom function, declare it as the last parameter.</span></span>

<span data-ttu-id="c3091-136">В следующем примере кода `invocation` контекст явно указывается для ссылки.</span><span class="sxs-lookup"><span data-stu-id="c3091-136">In the following code sample, the `invocation` context is explicitly stated for your reference.</span></span>

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
CustomFunctions.associate("ADD", add);
```

<span data-ttu-id="c3091-137">Параметр позволяет получить контекст вызывающей ячейки, который может быть полезен в некоторых сценариях, в том числе [Обнаружение адреса ячейки, которая вызывает настраиваемую функцию](#addressing-cells-context-parameter).</span><span class="sxs-lookup"><span data-stu-id="c3091-137">The parameter allows you to get the context of the invoking cell, which can be helpful in some scenarios including [discovering the address of a cell which invoke a custom function](#addressing-cells-context-parameter).</span></span>

### <a name="addressing-cells-context-parameter"></a><span data-ttu-id="c3091-138">Параметр контекста ячейки адресации</span><span class="sxs-lookup"><span data-stu-id="c3091-138">Addressing cell's context parameter</span></span>

<span data-ttu-id="c3091-139">В некоторых случаях необходимо получить адрес ячейки, которая вызвала пользовательскую функцию.</span><span class="sxs-lookup"><span data-stu-id="c3091-139">In some cases you need to get the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="c3091-140">Это полезно в следующих сценариях:</span><span class="sxs-lookup"><span data-stu-id="c3091-140">This is useful in the following scenarios:</span></span>

- <span data-ttu-id="c3091-141">Диапазоны форматирования: используйте адрес ячейки в качестве ключа для хранения информации в [оффицерунтиме. Storage](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data).</span><span class="sxs-lookup"><span data-stu-id="c3091-141">Formatting ranges: Use the cell's address as the key to store information in [OfficeRuntime.storage](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data).</span></span> <span data-ttu-id="c3091-142">После этого используйте событие [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) в Excel, чтобы загрузить ключ из `OfficeRuntime.storage`.</span><span class="sxs-lookup"><span data-stu-id="c3091-142">Then, use [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) in Excel to load the key from `OfficeRuntime.storage`.</span></span>
- <span data-ttu-id="c3091-143">Отображение кэшированных значений. Если функция используется в автономном режиме, отображайте сохраненные в кэше значения из `OfficeRuntime.storage` с помощью `onCalculated`.</span><span class="sxs-lookup"><span data-stu-id="c3091-143">Displaying cached values: If your function is used offline, display stored cached values from `OfficeRuntime.storage` using `onCalculated`.</span></span>
- <span data-ttu-id="c3091-144">Сверка: используйте адрес ячейки, чтобы найти исходную ячейку, чтобы упростить сверку при выполнении обработки.</span><span class="sxs-lookup"><span data-stu-id="c3091-144">Reconciliation: Use the cell's address to discover an origin cell to help you reconcile where processing is occurring.</span></span>

<span data-ttu-id="c3091-145">Чтобы запросить контекст ячейки адресации в функции, необходимо использовать функцию для поиска адреса ячейки, например, в приведенном ниже примере.</span><span class="sxs-lookup"><span data-stu-id="c3091-145">To request an addressing cell's context in a function, you need to use a function to find the cell's address, such as the one in the following example.</span></span> <span data-ttu-id="c3091-146">Сведения об адресе ячейки отображаются только в том случае, `@requiresAddress` если она помечена комментариями функции.</span><span class="sxs-lookup"><span data-stu-id="c3091-146">The information about a cell's address is exposed only if `@requiresAddress` is tagged in the function's comments.</span></span>

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
CustomFunctions.associate("GETADDRESS", getAddress);
```

<span data-ttu-id="c3091-147">По умолчанию значения, возвращаемые из функции `getAddress`, соответствуют следующему формату: `SheetName!CellNumber`.</span><span class="sxs-lookup"><span data-stu-id="c3091-147">By default, values returned from a `getAddress` function follow the following format: `SheetName!CellNumber`.</span></span> <span data-ttu-id="c3091-148">Например, если функция вызвана с листа с названием Expenses (Расходы) в ячейке B2, возвращаемым значением будет `Expenses!B2`.</span><span class="sxs-lookup"><span data-stu-id="c3091-148">For example, if a function was called from a sheet called Expenses in cell B2, the returned value would be `Expenses!B2`.</span></span>

## <a name="next-steps"></a><span data-ttu-id="c3091-149">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="c3091-149">Next steps</span></span>
<span data-ttu-id="c3091-150">Сведения о том, как [сохранить состояние в пользовательских функциях](custom-functions-save-state.md) или использовать [переменные значения в пользовательских функциях](custom-functions-volatile.md).</span><span class="sxs-lookup"><span data-stu-id="c3091-150">Learn how to [save state in your custom functions](custom-functions-save-state.md) or use [volatile values in your custom functions](custom-functions-volatile.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="c3091-151">См. также</span><span class="sxs-lookup"><span data-stu-id="c3091-151">See also</span></span>

* [<span data-ttu-id="c3091-152">Получение и обработка данных с помощью пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="c3091-152">Receive and handle data with custom functions</span></span>](custom-functions-web-reqs.md)
* [<span data-ttu-id="c3091-153">Рекомендации по пользовательским функциям</span><span class="sxs-lookup"><span data-stu-id="c3091-153">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="c3091-154">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="c3091-154">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="c3091-155">Автоматическое генерирование метаданных JSON для пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="c3091-155">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="c3091-156">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="c3091-156">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="c3091-157">Руководство по пользовательским функциям в Excel</span><span class="sxs-lookup"><span data-stu-id="c3091-157">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
