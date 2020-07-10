---
ms.date: 05/06/2020
description: 'Обработка и возврат таких ошибок, как #ПУСТО!, из пользовательской функции'
title: Обработка и возврат ошибок из пользовательской функции (предварительная версия)
localization_priority: Normal
ms.openlocfilehash: 5b1efcdc22a4efc59304bbe76f8d3f2d09979bc1
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093471"
---
# <a name="handle-and-return-errors-from-your-custom-function-preview"></a><span data-ttu-id="cbf4b-104">Обработка и возврат ошибок из пользовательской функции (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="cbf4b-104">Handle and return errors from your custom function (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="cbf4b-105">Возможности, описанные в этой статье, в настоящее время доступны в предварительной версии и могут изменяться.</span><span class="sxs-lookup"><span data-stu-id="cbf4b-105">The features described in this article are currently in preview and subject to change.</span></span> <span data-ttu-id="cbf4b-106">В настоящее время их нельзя использовать в рабочих средах.</span><span class="sxs-lookup"><span data-stu-id="cbf4b-106">They are not currently supported for use in production environments.</span></span> <span data-ttu-id="cbf4b-107">Для ознакомления с предварительными возможностями необходимо присоединиться к программе [предварительной оценки Office](https://insider.office.com/join) .</span><span class="sxs-lookup"><span data-stu-id="cbf4b-107">You will need to join the [Office Insider](https://insider.office.com/join) program to try the preview features.</span></span>  <span data-ttu-id="cbf4b-108">Хороший способ испытать ознакомительные функции — использовать подписку на Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="cbf4b-108">A good way to try out preview features is by using a Microsoft 365 subscription.</span></span> <span data-ttu-id="cbf4b-109">Если у вас еще нет подписки на Microsoft 365, вы можете получить бесплатную, 90 день реневабле подписку на Microsoft 365, присоединяясь к [программе microsoft 365 Developer Program](https://developer.microsoft.com/office/dev-program).</span><span class="sxs-lookup"><span data-stu-id="cbf4b-109">If you don't already have a Microsoft 365 subscription, you can get a free, 90-day renewable Microsoft 365 subscription by joining the [Microsoft 365 developer program](https://developer.microsoft.com/office/dev-program).</span></span>

<span data-ttu-id="cbf4b-110">Если при выполнении пользовательской функции возникла проблема, возвращайте ошибку, чтобы уведомить пользователя.</span><span class="sxs-lookup"><span data-stu-id="cbf4b-110">If something goes wrong while your custom function runs, return an error to inform the user.</span></span> <span data-ttu-id="cbf4b-111">Если у вас есть особые требования к параметрам, например, только положительные числа, проверьте параметры и вызовите ошибку, если они неправильные.</span><span class="sxs-lookup"><span data-stu-id="cbf4b-111">If you have specific parameter requirements, such as only positive numbers, test the parameters and throw an error if they aren't correct.</span></span> <span data-ttu-id="cbf4b-112">Можно также использовать блок `try`-`catch`, чтобы отслеживать любые ошибки, возникающие при выполнении пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="cbf4b-112">You can also use a `try`-`catch` block to catch any errors that occur while your custom function runs.</span></span>

## <a name="detect-and-throw-an-error"></a><span data-ttu-id="cbf4b-113">Обнаружение и возвращение ошибки</span><span class="sxs-lookup"><span data-stu-id="cbf4b-113">Detect and throw an error</span></span>

<span data-ttu-id="cbf4b-114">Рассмотрим ситуацию, в которой необходимо убедиться, что параметр ZIP-кода имеет правильный формат, чтобы пользовательская функция работала.</span><span class="sxs-lookup"><span data-stu-id="cbf4b-114">Let's look at a case where you need to ensure that a zip code parameter is in the correct format for the custom function to work.</span></span> <span data-ttu-id="cbf4b-115">В следующей пользовательской функции используется регулярное выражение для проверки почтового индекса.</span><span class="sxs-lookup"><span data-stu-id="cbf4b-115">The following custom function uses a regular expression to check the zip code.</span></span> <span data-ttu-id="cbf4b-116">Если он правильный, то он будет искать город с помощью другой функции и возвращать значение.</span><span class="sxs-lookup"><span data-stu-id="cbf4b-116">If it is correct, then it will look up the city using another function, and return the value.</span></span> <span data-ttu-id="cbf4b-117">Если это не так, то возвращается `#VALUE!` Ошибка в ячейке.</span><span class="sxs-lookup"><span data-stu-id="cbf4b-117">If it isn't correct, it returns a `#VALUE!` error to the cell.</span></span>

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

## <a name="the-customfunctionserror-object"></a><span data-ttu-id="cbf4b-118">Объект CustomFunctions.Error</span><span class="sxs-lookup"><span data-stu-id="cbf4b-118">The CustomFunctions.Error object</span></span>

<span data-ttu-id="cbf4b-119">Объект `CustomFunctions.Error` используется для возвращения ошибки в ячейку.</span><span class="sxs-lookup"><span data-stu-id="cbf4b-119">The `CustomFunctions.Error` object is used to return an error back to the cell.</span></span> <span data-ttu-id="cbf4b-120">При создании объекта укажите, какую ошибку нужно использовать, применив одно из следующих значений перечисления `ErrorCode`.</span><span class="sxs-lookup"><span data-stu-id="cbf4b-120">When you create the object, specify which error you want to use by using one of the following `ErrorCode` enum values.</span></span>


|<span data-ttu-id="cbf4b-121">Значение перечисления ErrorCode</span><span class="sxs-lookup"><span data-stu-id="cbf4b-121">ErrorCode enum value</span></span>  |<span data-ttu-id="cbf4b-122">Значение ячейки Excel</span><span class="sxs-lookup"><span data-stu-id="cbf4b-122">Excel cell value</span></span>  |<span data-ttu-id="cbf4b-123">Смысл</span><span class="sxs-lookup"><span data-stu-id="cbf4b-123">Meaning</span></span>  |
|---------------|---------|---------|
|`invalidValue`   | `#VALUE!` | <span data-ttu-id="cbf4b-124">В формуле используется значение неправильного типа.</span><span class="sxs-lookup"><span data-stu-id="cbf4b-124">A value used in the formula is the wrong type.</span></span> |
|`notAvailable`   | `#N/A`    | <span data-ttu-id="cbf4b-125">Функция или служба недоступна.</span><span class="sxs-lookup"><span data-stu-id="cbf4b-125">The function or service isn't available.</span></span> |
|`divisionByZero` | `#DIV/0`  | <span data-ttu-id="cbf4b-126">Обратите внимание, что JavaScript позволяет делить на нуль, поэтому при создании обработчика ошибок нужно внимательно определить это условие.</span><span class="sxs-lookup"><span data-stu-id="cbf4b-126">Be aware that JavaScript allows division by zero so you need to write an error handler carefully to detect this condition.</span></span> |
|`invalidNumber`  | `#NUM!`   | <span data-ttu-id="cbf4b-127">Обнаружена проблема с числом, используемым в формуле</span><span class="sxs-lookup"><span data-stu-id="cbf4b-127">There is a problem with the number used in the formula</span></span> |
|`nullReference`  | `#NULL!`  | <span data-ttu-id="cbf4b-128">Диапазоны в формуле не пересекаются.</span><span class="sxs-lookup"><span data-stu-id="cbf4b-128">The ranges in the formula don't intersect.</span></span> |

<span data-ttu-id="cbf4b-129">В следующем примере кода показано, как создать и вернуть ошибку для неверного числа (`#NUM!`).</span><span class="sxs-lookup"><span data-stu-id="cbf4b-129">The following code sample shows how to create and return an error for an invalid number (`#NUM!`).</span></span>

```typescript
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidNumber);
throw error;
```

<span data-ttu-id="cbf4b-130">При возврате ошибки `#VALUE!` также можно включить настраиваемое сообщение, отображаемое во всплывающем окне, когда пользователь наводит на ячейку указатель мыши.</span><span class="sxs-lookup"><span data-stu-id="cbf4b-130">When you return a `#VALUE!` error you can also include a custom message that will be shown in a popup when the user hovers over the cell.</span></span> <span data-ttu-id="cbf4b-131">В следующем примере показано, как вернуть настраиваемое сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="cbf4b-131">The following example shows how to return a custom error message.</span></span>

```typescript
// You can only return a custom error message with the #VALUE! error
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, "The parameter can only contain lowercase characters.");
throw error;
```

## <a name="use-try-catch-blocks"></a><span data-ttu-id="cbf4b-132">Использование блоков try-catch</span><span class="sxs-lookup"><span data-stu-id="cbf4b-132">Use try-catch blocks</span></span>

<span data-ttu-id="cbf4b-133">В общем случае `try` - `catch` для перехвата возможных ошибок используйте блоки в пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="cbf4b-133">In general, use `try`-`catch` blocks in your custom function to catch any potential errors that occur.</span></span> <span data-ttu-id="cbf4b-134">Если в коде не обрабатываются исключения, они будут возвращаться в Excel.</span><span class="sxs-lookup"><span data-stu-id="cbf4b-134">If you do not handle exceptions in your code, they will be returned to Excel.</span></span> <span data-ttu-id="cbf4b-135">По умолчанию Excel возвращает `#VALUE!` для необработанного исключения.</span><span class="sxs-lookup"><span data-stu-id="cbf4b-135">By default, Excel returns `#VALUE!` for an unhandled exception.</span></span>

<span data-ttu-id="cbf4b-136">В следующем примере кода пользовательская функция создает запрос fetch в службу REST.</span><span class="sxs-lookup"><span data-stu-id="cbf4b-136">In the following code sample, the custom function makes a fetch call to a REST service.</span></span> <span data-ttu-id="cbf4b-137">Возможно, что вызов завершится сбоем (например, если служба REST возвращает ошибку или не работает сеть).</span><span class="sxs-lookup"><span data-stu-id="cbf4b-137">It's possible that the call will fail, for example, if the REST service returns an error or the network goes down.</span></span> <span data-ttu-id="cbf4b-138">В этом случае пользовательская функция возвращает `#N/A`, чтобы указать на сбой веб-вызова.</span><span class="sxs-lookup"><span data-stu-id="cbf4b-138">If this happens, the custom function will return `#N/A` to indicate the web call failed.</span></span>


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

## <a name="next-steps"></a><span data-ttu-id="cbf4b-139">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="cbf4b-139">Next steps</span></span>

<span data-ttu-id="cbf4b-140">Узнайте, как [устранять проблемы с пользовательскими функциями](custom-functions-troubleshooting.md).</span><span class="sxs-lookup"><span data-stu-id="cbf4b-140">Learn how to [troubleshoot problems with your custom functions](custom-functions-troubleshooting.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="cbf4b-141">Дополнительные ресурсы</span><span class="sxs-lookup"><span data-stu-id="cbf4b-141">See also</span></span>

* [<span data-ttu-id="cbf4b-142">Отладка пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="cbf4b-142">Custom functions debugging</span></span>](custom-functions-debugging.md)
* [<span data-ttu-id="cbf4b-143">Требования к настраиваемым функциям</span><span class="sxs-lookup"><span data-stu-id="cbf4b-143">Custom functions requirements</span></span>](custom-functions-requirement-sets.md)
* [<span data-ttu-id="cbf4b-144">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="cbf4b-144">Create custom functions in Excel</span></span>](custom-functions-overview.md)
