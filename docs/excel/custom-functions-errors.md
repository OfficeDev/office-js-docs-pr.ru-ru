---
ms.date: 09/21/2020
description: 'Обработка и возврат таких ошибок, как #ПУСТО!, из пользовательской функции.'
title: Обработка и возврат ошибок пользовательской функции
localization_priority: Normal
ms.openlocfilehash: 58c2ab432a4525f660e2d89735fd3add6e76fa7f
ms.sourcegitcommit: 4a03d8b3f676ee2d91114813cb81bce5da3c8d6b
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/22/2020
ms.locfileid: "48175530"
---
# <a name="handle-and-return-errors-from-your-custom-function"></a><span data-ttu-id="feceb-104">Обработка и возврат ошибок пользовательской функции</span><span class="sxs-lookup"><span data-stu-id="feceb-104">Handle and return errors from your custom function</span></span>

<span data-ttu-id="feceb-105">Если при выполнении пользовательской функции возникла проблема, возвращайте ошибку, чтобы уведомить пользователя.</span><span class="sxs-lookup"><span data-stu-id="feceb-105">If something goes wrong while your custom function runs, return an error to inform the user.</span></span> <span data-ttu-id="feceb-106">Если у вас есть особые требования к параметрам, например, только положительные числа, проверьте параметры и вызовите ошибку, если они неправильные.</span><span class="sxs-lookup"><span data-stu-id="feceb-106">If you have specific parameter requirements, such as only positive numbers, test the parameters and throw an error if they aren't correct.</span></span> <span data-ttu-id="feceb-107">Можно также использовать блок `try`-`catch`, чтобы отслеживать любые ошибки, возникающие при выполнении пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="feceb-107">You can also use a `try`-`catch` block to catch any errors that occur while your custom function runs.</span></span>

## <a name="detect-and-throw-an-error"></a><span data-ttu-id="feceb-108">Обнаружение и возвращение ошибки</span><span class="sxs-lookup"><span data-stu-id="feceb-108">Detect and throw an error</span></span>

<span data-ttu-id="feceb-109">Рассмотрим ситуацию, в которой необходимо убедиться, что параметр ZIP-кода имеет правильный формат, чтобы пользовательская функция работала.</span><span class="sxs-lookup"><span data-stu-id="feceb-109">Let's look at a case where you need to ensure that a zip code parameter is in the correct format for the custom function to work.</span></span> <span data-ttu-id="feceb-110">В следующей пользовательской функции используется регулярное выражение для проверки почтового индекса.</span><span class="sxs-lookup"><span data-stu-id="feceb-110">The following custom function uses a regular expression to check the zip code.</span></span> <span data-ttu-id="feceb-111">Если формат ZIP-кода правильный, то он будет искать город с помощью другой функции и возвращать значение.</span><span class="sxs-lookup"><span data-stu-id="feceb-111">If the zip code format is correct, then it will look up the city using another function and return the value.</span></span> <span data-ttu-id="feceb-112">Если формат не является допустимым, функция возвращает `#VALUE!` ошибку в ячейку.</span><span class="sxs-lookup"><span data-stu-id="feceb-112">If the format isn't valid, the function returns a `#VALUE!` error to the cell.</span></span>

```typescript
/**
* Gets a city name for the given U.S. zip code.
* @customfunction
* @param {string} zipCode
* @returns The city of the zip code.
*/
function getCity(zipCode: string): string {
  let isValidZip = /(^\d{5}$)|(^\d{5}-\d{4}$)/.test(zipCode);
  if (isValidZip) return cityLookup(zipCode);
  let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, "Please provide a valid U.S. zip code.");
  throw error;
}
```

## <a name="the-customfunctionserror-object"></a><span data-ttu-id="feceb-113">Объект CustomFunctions.Error</span><span class="sxs-lookup"><span data-stu-id="feceb-113">The CustomFunctions.Error object</span></span>

<span data-ttu-id="feceb-114">Объект [CustomFunctions. Error](/javascript/api/custom-functions-runtime/customfunctions.error) используется для возврата к ячейке ошибки.</span><span class="sxs-lookup"><span data-stu-id="feceb-114">The [CustomFunctions.Error](/javascript/api/custom-functions-runtime/customfunctions.error) object is used to return an error back to the cell.</span></span> <span data-ttu-id="feceb-115">При создании объекта укажите, какую ошибку следует использовать, выбрав одно из следующих `ErrorCode` значений перечисления.</span><span class="sxs-lookup"><span data-stu-id="feceb-115">When you create the object, specify which error you want to use by choosing one of the following `ErrorCode` enum values.</span></span>


|<span data-ttu-id="feceb-116">Значение перечисления ErrorCode</span><span class="sxs-lookup"><span data-stu-id="feceb-116">ErrorCode enum value</span></span>  |<span data-ttu-id="feceb-117">Значение ячейки Excel</span><span class="sxs-lookup"><span data-stu-id="feceb-117">Excel cell value</span></span>  |<span data-ttu-id="feceb-118">Смысл</span><span class="sxs-lookup"><span data-stu-id="feceb-118">Meaning</span></span>  |
|---------------|---------|---------|
|`divisionByZero` | `#DIV/0`  | <span data-ttu-id="feceb-119">Обратите внимание, что JavaScript позволяет делить на нуль, поэтому при создании обработчика ошибок нужно внимательно определить это условие.</span><span class="sxs-lookup"><span data-stu-id="feceb-119">Be aware that JavaScript allows division by zero so you need to write an error handler carefully to detect this condition.</span></span> |
|`invalidName`    | `#NAME?`  | <span data-ttu-id="feceb-120">В имени функции присутствует опечатка.</span><span class="sxs-lookup"><span data-stu-id="feceb-120">There is a typo in the function name.</span></span> <span data-ttu-id="feceb-121">Обратите внимание, что эта ошибка поддерживается как ошибка ввода пользовательской функции, но не в качестве ошибки вывода пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="feceb-121">Note that this error is supported as a custom function input error, but not as a custom function output error.</span></span> | 
|`invalidNumber`  | `#NUM!`   | <span data-ttu-id="feceb-122">Возникла проблема с числом в формуле.</span><span class="sxs-lookup"><span data-stu-id="feceb-122">There is a problem with a number in the formula.</span></span> |
|`invalidReference` | `#REF!` | <span data-ttu-id="feceb-123">Функция ссылается на недопустимую ячейку.</span><span class="sxs-lookup"><span data-stu-id="feceb-123">The function refers to an invalid cell.</span></span> <span data-ttu-id="feceb-124">Обратите внимание, что эта ошибка поддерживается как ошибка ввода пользовательской функции, но не в качестве ошибки вывода пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="feceb-124">Note that this error is supported as a custom function input error, but not as a custom function output error.</span></span>|
|`invalidValue`   | `#VALUE!` | <span data-ttu-id="feceb-125">Недопустимый тип значения в формуле.</span><span class="sxs-lookup"><span data-stu-id="feceb-125">A value in the formula is of the wrong type.</span></span> |
|`notAvailable`   | `#N/A`    | <span data-ttu-id="feceb-126">Функция или служба недоступна.</span><span class="sxs-lookup"><span data-stu-id="feceb-126">The function or service isn't available.</span></span> |
|`nullReference`  | `#NULL!`  | <span data-ttu-id="feceb-127">Диапазоны в формуле не пересекаются.</span><span class="sxs-lookup"><span data-stu-id="feceb-127">The ranges in the formula don't intersect.</span></span> |

<span data-ttu-id="feceb-128">В следующем примере кода показано, как создать и вернуть ошибку для неверного числа (`#NUM!`).</span><span class="sxs-lookup"><span data-stu-id="feceb-128">The following code sample shows how to create and return an error for an invalid number (`#NUM!`).</span></span>

```typescript
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidNumber);
throw error;
```

<span data-ttu-id="feceb-129">`#VALUE!` `#N/A` Кроме того, ошибки также поддерживают настраиваемые сообщения об ошибках.</span><span class="sxs-lookup"><span data-stu-id="feceb-129">The `#VALUE!` and `#N/A` errors also support custom error messages.</span></span> <span data-ttu-id="feceb-130">Настраиваемые сообщения об ошибках отображаются в меню индикации ошибки, доступ к которому осуществляется при наведении курсора на флаг ошибки в каждой ячейке с ошибкой.</span><span class="sxs-lookup"><span data-stu-id="feceb-130">Custom error messages are displayed in the error indicator menu, which is accessed by hovering over the error flag on each cell with an error.</span></span> <span data-ttu-id="feceb-131">В приведенном ниже примере показано, как вернуть настраиваемое сообщение об ошибке с `#VALUE!` ошибкой.</span><span class="sxs-lookup"><span data-stu-id="feceb-131">The following example shows how to return a custom error message with the `#VALUE!` error.</span></span>

```typescript
// You can only return a custom error message with the #VALUE! and #N/A errors.
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, "The parameter can only contain lowercase characters.");
throw error;
```

## <a name="use-try-catch-blocks"></a><span data-ttu-id="feceb-132">Использование блоков try-catch</span><span class="sxs-lookup"><span data-stu-id="feceb-132">Use try-catch blocks</span></span>

<span data-ttu-id="feceb-133">В общем случае `try` - `catch` для перехвата возможных ошибок используйте блоки в пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="feceb-133">In general, use `try`-`catch` blocks in your custom function to catch any potential errors that occur.</span></span> <span data-ttu-id="feceb-134">Если в коде не обрабатываются исключения, они будут возвращаться в Excel.</span><span class="sxs-lookup"><span data-stu-id="feceb-134">If you do not handle exceptions in your code, they will be returned to Excel.</span></span> <span data-ttu-id="feceb-135">По умолчанию Excel возвращает `#VALUE!` для необработанных ошибок или исключений.</span><span class="sxs-lookup"><span data-stu-id="feceb-135">By default, Excel returns `#VALUE!` for unhandled errors or exceptions.</span></span>

<span data-ttu-id="feceb-136">В следующем примере кода пользовательская функция создает запрос fetch в службу REST.</span><span class="sxs-lookup"><span data-stu-id="feceb-136">In the following code sample, the custom function makes a fetch call to a REST service.</span></span> <span data-ttu-id="feceb-137">Возможно, что вызов завершится сбоем (например, если служба REST возвращает ошибку или не работает сеть).</span><span class="sxs-lookup"><span data-stu-id="feceb-137">It's possible that the call will fail, for example, if the REST service returns an error or the network goes down.</span></span> <span data-ttu-id="feceb-138">В этом случае пользовательская функция вернется, `#N/A` чтобы указать, что веб-вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="feceb-138">If this happens, the custom function will return `#N/A` to indicate that the web call failed.</span></span>


```typescript
/**
 * Gets a comment from the hypothetical contoso.com/comments API.
 * @customfunction
 * @param {number} commentID ID of a comment.
 */
function getComment(commentID) {
  let url = "https://www.contoso.com/comments/" + commentID;
  return fetch(url)
    .then(function (data) {
      return data.json();
    })
    .then(function (json) {
      return json.body;
    })
    .catch(function (error) {
      throw new CustomFunctions.Error(CustomFunctions.ErrorCode.notAvailable);
    })
}
```

## <a name="next-steps"></a><span data-ttu-id="feceb-139">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="feceb-139">Next steps</span></span>

<span data-ttu-id="feceb-140">Узнайте, как [устранять проблемы с пользовательскими функциями](custom-functions-troubleshooting.md).</span><span class="sxs-lookup"><span data-stu-id="feceb-140">Learn how to [troubleshoot problems with your custom functions](custom-functions-troubleshooting.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="feceb-141">Дополнительные ресурсы</span><span class="sxs-lookup"><span data-stu-id="feceb-141">See also</span></span>

* [<span data-ttu-id="feceb-142">Отладка пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="feceb-142">Custom functions debugging</span></span>](custom-functions-debugging.md)
* [<span data-ttu-id="feceb-143">Требования к настраиваемым функциям</span><span class="sxs-lookup"><span data-stu-id="feceb-143">Custom functions requirements</span></span>](custom-functions-requirement-sets.md)
* [<span data-ttu-id="feceb-144">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="feceb-144">Create custom functions in Excel</span></span>](custom-functions-overview.md)
