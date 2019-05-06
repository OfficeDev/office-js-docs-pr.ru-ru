---
ms.date: 04/30/2019
description: Узнайте, как использовать различные параметры в пользовательских функциях, таких как диапазоны Excel, необязательные параметры, контекст вызова и многое другое.
title: Параметры для пользовательских функций Excel (Предварительная версия)
localization_priority: Normal
ms.openlocfilehash: b5dba59431f4c6ec4ee08c563e7cb3affeb06608
ms.sourcegitcommit: 68872372d181cca5bee37ade73c2250c4a56bab6
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/01/2019
ms.locfileid: "33527324"
---
# <a name="custom-functions-parameter-options"></a><span data-ttu-id="f43d0-103">Параметры параметров пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="f43d0-103">Custom functions parameter options</span></span>

<span data-ttu-id="f43d0-104">Настраиваемые функции можно настраивать с помощью различных параметров:</span><span class="sxs-lookup"><span data-stu-id="f43d0-104">Custom functions are configurable with many different options for parameters:</span></span> 
- [<span data-ttu-id="f43d0-105">Необязательные параметры</span><span class="sxs-lookup"><span data-stu-id="f43d0-105">Optional parameters</span></span>](#custom-functions-optional-parameters)
- [<span data-ttu-id="f43d0-106">Параметры Range</span><span class="sxs-lookup"><span data-stu-id="f43d0-106">Range parameters</span></span>](#range-parameters)
- [<span data-ttu-id="f43d0-107">Параметр контекста вызова</span><span class="sxs-lookup"><span data-stu-id="f43d0-107">Invocation context parameter</span></span>](#invocation-parameter)

## <a name="custom-functions-optional-parameters"></a><span data-ttu-id="f43d0-108">Необязательные параметры настраиваемых функций</span><span class="sxs-lookup"><span data-stu-id="f43d0-108">Custom functions optional parameters</span></span>

<span data-ttu-id="f43d0-109">В то время как обычные параметры являются обязательными, необязательные параметры — нет.</span><span class="sxs-lookup"><span data-stu-id="f43d0-109">Whereas regular parameters are required, optional parameters are not.</span></span> <span data-ttu-id="f43d0-110">Если пользователь вызывает функцию в Excel, необязательные параметры отображаются в квадратных скобках.</span><span class="sxs-lookup"><span data-stu-id="f43d0-110">When a user invokes a function in Excel, optional parameters appear in brackets.</span></span> <span data-ttu-id="f43d0-111">В приведенном ниже примере функция Add может дополнительно добавить третий номер.</span><span class="sxs-lookup"><span data-stu-id="f43d0-111">In the following sample, the add function can optionally add a third number.</span></span> <span data-ttu-id="f43d0-112">Эта функция отображается как `=CONTOSO.ADD(first, second, [third])` в Excel.</span><span class="sxs-lookup"><span data-stu-id="f43d0-112">This function appears as `=CONTOSO.ADD(first, second, [third])` in Excel.</span></span>

```js
/**
 * Add two numbers
 * @customfunction 
 * @param {number} first First number.
 * @param {number} second Second number.
 * @param {number} [third] Third number to add. If omitted, third = 0.
 * @returns {number} The sum of the numbers.
 */
function add(first, second, third) {
  if (third === undefined) {
    return first + second + third;
  }
  return first + second;
}
CustomFunctions.associate("ADD", add);
```

<span data-ttu-id="f43d0-113">Если вы определяете функцию, содержащую один или несколько необязательных параметров, нужно указать, что происходит, когда необязательный параметр не задан.</span><span class="sxs-lookup"><span data-stu-id="f43d0-113">When you define a function that contains one or more optional parameters, you should specify what happens when the optional parameters are undefined.</span></span> <span data-ttu-id="f43d0-114">В приведенном ниже примере `zipCode` и `dayOfWeek` являются необязательными параметрами для функции `getWeatherReport`.</span><span class="sxs-lookup"><span data-stu-id="f43d0-114">In the following example, `zipCode` and `dayOfWeek` are both optional parameters for the `getWeatherReport` function.</span></span> <span data-ttu-id="f43d0-115">Если `zipCode` параметр не определен, для `98052`него устанавливается значение по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="f43d0-115">If the `zipCode` parameter is undefined, the default value is set to `98052`.</span></span> <span data-ttu-id="f43d0-116">Если параметр `dayOfWeek` не определен, ему присваивается значение Wednesday (Среда).</span><span class="sxs-lookup"><span data-stu-id="f43d0-116">If the `dayOfWeek` parameter is undefined, it is set to Wednesday.</span></span>

```js
/**
 * Gets a weather report for a specified zipCode and dayOfWeek
 * @customfunction
 * @param {number} zipCode Zip code. If omitted, zipCode = 98052.
 * @param {string} dayOfWeek Day of the week. If omitted, dayOfWeek = Wednesday.
 * @returns {string} Weather report for the day of the week in that zip code.
 */
function getWeatherReport(zipCode, dayOfWeek)
{
  if (zipCode === undefined) {
      zipCode = "98052";
  }

  if (dayOfWeek === undefined) {
    dayOfWeek = "Wednesday";
  }

  // Get weather report for specified zipCode and dayOfWeek.
  // ...
}
```

## <a name="range-parameters"></a><span data-ttu-id="f43d0-117">Параметры Range</span><span class="sxs-lookup"><span data-stu-id="f43d0-117">Range parameters</span></span>

<span data-ttu-id="f43d0-118">Настраиваемая функция может принимать диапазон данных ячейки в качестве входного параметра.</span><span class="sxs-lookup"><span data-stu-id="f43d0-118">Your custom function may accept a range of cell data as an input parameter.</span></span> <span data-ttu-id="f43d0-119">Функция также может возвращать диапазон данных.</span><span class="sxs-lookup"><span data-stu-id="f43d0-119">A function can also return a range of data.</span></span> <span data-ttu-id="f43d0-120">Excel передает диапазон данных ячейки в виде двумерного массива.</span><span class="sxs-lookup"><span data-stu-id="f43d0-120">Excel will pass a range of cell data as a two-dimensional array.</span></span>

<span data-ttu-id="f43d0-121">Например, предположим, что функция возвращает второе по величине значение из диапазона значений, хранящихся в Excel.</span><span class="sxs-lookup"><span data-stu-id="f43d0-121">For example, suppose that your function returns the second highest value from a range of numbers stored in Excel.</span></span> <span data-ttu-id="f43d0-122">Приведенная ниже функция принимает параметр `values`, относящийся к типу `Excel.CustomFunctionDimensionality.matrix`.</span><span class="sxs-lookup"><span data-stu-id="f43d0-122">The following function accepts the parameter `values`, which is of type `Excel.CustomFunctionDimensionality.matrix`.</span></span> <span data-ttu-id="f43d0-123">Обратите внимание, что в метаданных JSON для этой функции для `type` свойства параметра задано значение `matrix`.</span><span class="sxs-lookup"><span data-stu-id="f43d0-123">Note that in the JSON metadata for this function, the parameter's `type` property is set to `matrix`.</span></span>

```js
/**
 * Returns the second highest value in a matrixed range of values.
 * @customfunction
 * @param {[][]} values Multiple ranges of values.  
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

## <a name="invocation-parameter"></a><span data-ttu-id="f43d0-124">Параметр вызова</span><span class="sxs-lookup"><span data-stu-id="f43d0-124">Invocation parameter</span></span>

<span data-ttu-id="f43d0-125">Каждая пользовательская функция автоматически передает `invocation` аргумент в качестве последнего аргумента.</span><span class="sxs-lookup"><span data-stu-id="f43d0-125">Every custom function is automatically passed an `invocation` argument as the last argument.</span></span> <span data-ttu-id="f43d0-126">Этот аргумент можно использовать для получения дополнительного контекста, например адреса вызывающей ячейки.</span><span class="sxs-lookup"><span data-stu-id="f43d0-126">This argument can be used to retrieve additional context, such as the address of the calling cell.</span></span> <span data-ttu-id="f43d0-127">Или его можно использовать для отправки в Excel данных, например обработчика функции для [отмены функции](custom-functions-web-reqs.md#stream-and-cancel-functions).</span><span class="sxs-lookup"><span data-stu-id="f43d0-127">Or it can be used to send information to Excel, such as a function handler for [canceling a function](custom-functions-web-reqs.md#stream-and-cancel-functions).</span></span> <span data-ttu-id="f43d0-128">Даже если вы не объявили параметры, у настраиваемой функции есть этот параметр.</span><span class="sxs-lookup"><span data-stu-id="f43d0-128">Even if you declare no parameters, your custom function has this parameter.</span></span> <span data-ttu-id="f43d0-129">Этот аргумент не отображается для пользователя в Excel.</span><span class="sxs-lookup"><span data-stu-id="f43d0-129">This argument doesn't appear for a user in Excel.</span></span> <span data-ttu-id="f43d0-130">Если вы хотите использовать `invocation` пользовательскую функцию, объявите ее в качестве последнего параметра.</span><span class="sxs-lookup"><span data-stu-id="f43d0-130">If you want to use `invocation` in your custom function, declare it as the last parameter.</span></span>

<span data-ttu-id="f43d0-131">В следующем примере кода `invocation` контекст явно указывается для ссылки.</span><span class="sxs-lookup"><span data-stu-id="f43d0-131">In the following code sample, the `invocation` context is explicitly stated for your reference.</span></span>

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

<span data-ttu-id="f43d0-132">Параметр позволяет получить контекст вызывающей ячейки, который может быть полезен в некоторых сценариях, в том числе [Обнаружение адреса ячейки, которая вызывает настраиваемую функцию](#addressing-cells-context-parameter).</span><span class="sxs-lookup"><span data-stu-id="f43d0-132">The parameter allows you to get the context of the invoking cell, which can be helpful in some scenarios including [discovering the address of a cell which invoke a custom function](#addressing-cells-context-parameter).</span></span>

### <a name="addressing-cells-context-parameter"></a><span data-ttu-id="f43d0-133">Параметр контекста ячейки адресации</span><span class="sxs-lookup"><span data-stu-id="f43d0-133">Addressing cell's context parameter</span></span>

<span data-ttu-id="f43d0-134">В некоторых случаях необходимо получить адрес ячейки, которая вызвала пользовательскую функцию.</span><span class="sxs-lookup"><span data-stu-id="f43d0-134">In some cases you need to get the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="f43d0-135">Это полезно в следующих типах сценариев:</span><span class="sxs-lookup"><span data-stu-id="f43d0-135">This is useful in the following types of scenarios:</span></span>

- <span data-ttu-id="f43d0-136">Диапазоны форматирования: используйте адрес ячейки в качестве ключа для хранения сведений в [Office. Storage](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data).</span><span class="sxs-lookup"><span data-stu-id="f43d0-136">Formatting ranges: Use the cell's address as the key to store information in [Office.storage](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data).</span></span> <span data-ttu-id="f43d0-137">После этого используйте событие [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) в Excel, чтобы загрузить ключ из `Office.storage`.</span><span class="sxs-lookup"><span data-stu-id="f43d0-137">Then, use [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) in Excel to load the key from `Office.storage`.</span></span>
- <span data-ttu-id="f43d0-138">Отображение кэшированных значений. Если функция используется в автономном режиме, отображайте сохраненные в кэше значения из `Office.storage` с помощью `onCalculated`.</span><span class="sxs-lookup"><span data-stu-id="f43d0-138">Displaying cached values: If your function is used offline, display stored cached values from `Office.storage` using `onCalculated`.</span></span>
- <span data-ttu-id="f43d0-139">Сверка: используйте адрес ячейки, чтобы найти исходную ячейку, чтобы упростить сверку при выполнении обработки.</span><span class="sxs-lookup"><span data-stu-id="f43d0-139">Reconciliation: Use the cell's address to discover an origin cell to help you reconcile where processing is occurring.</span></span>

<span data-ttu-id="f43d0-140">Чтобы запросить контекст ячейки адресации в функции, необходимо использовать функцию для поиска адреса ячейки, например, в приведенном ниже примере.</span><span class="sxs-lookup"><span data-stu-id="f43d0-140">To request an addressing cell's context in a function, you need to use a function to find the cell's address, such as the one in the following example.</span></span> <span data-ttu-id="f43d0-141">Сведения об адресе ячейки отображаются только в том случае, `@requiresAddress` если она помечена комментариями функции.</span><span class="sxs-lookup"><span data-stu-id="f43d0-141">The information about a cell's address is exposed only if `@requiresAddress` is tagged in the function's comments.</span></span>

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

<span data-ttu-id="f43d0-142">По умолчанию значения, возвращаемые из функции `getAddress`, соответствуют следующему формату: `SheetName!CellNumber`.</span><span class="sxs-lookup"><span data-stu-id="f43d0-142">By default, values returned from a `getAddress` function follow the following format: `SheetName!CellNumber`.</span></span> <span data-ttu-id="f43d0-143">Например, если функция вызвана с листа с названием Expenses (Расходы) в ячейке B2, возвращаемым значением будет `Expenses!B2`.</span><span class="sxs-lookup"><span data-stu-id="f43d0-143">For example, if a function was called from a sheet called Expenses in cell B2, the returned value would be `Expenses!B2`.</span></span>

## <a name="see-also"></a><span data-ttu-id="f43d0-144">См. также</span><span class="sxs-lookup"><span data-stu-id="f43d0-144">See also</span></span>

* [<span data-ttu-id="f43d0-145">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="f43d0-145">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="f43d0-146">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="f43d0-146">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="f43d0-147">Рекомендации по пользовательским функциям</span><span class="sxs-lookup"><span data-stu-id="f43d0-147">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="f43d0-148">Журнал изменений пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="f43d0-148">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="f43d0-149">Руководство по настраиваемым функциям в Excel</span><span class="sxs-lookup"><span data-stu-id="f43d0-149">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)