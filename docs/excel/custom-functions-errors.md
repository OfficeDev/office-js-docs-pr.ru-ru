---
ms.date: 03/11/2020
description: 'Обработка и возврат таких ошибок, как #ПУСТО!, из пользовательской функции'
title: Обработка и возврат ошибок из пользовательской функции (предварительная версия)
localization_priority: Normal
ms.openlocfilehash: 10bb7ca6ff612ef38b26b88fed5ce9ce81ed7edb
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717049"
---
# <a name="handle-and-return-errors-from-your-custom-function-preview"></a><span data-ttu-id="4d0a1-104">Обработка и возврат ошибок из пользовательской функции (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="4d0a1-104">Handle and return errors from your custom function (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="4d0a1-105">Возможности, описанные в этой статье, в настоящее время доступны в предварительной версии и могут изменяться.</span><span class="sxs-lookup"><span data-stu-id="4d0a1-105">The features described in this article are currently in preview and subject to change.</span></span> <span data-ttu-id="4d0a1-106">В настоящее время их нельзя использовать в рабочих средах.</span><span class="sxs-lookup"><span data-stu-id="4d0a1-106">They are not currently supported for use in production environments.</span></span> <span data-ttu-id="4d0a1-107">Для ознакомления с предварительными возможностями необходимо присоединиться к программе [предварительной оценки Office](https://insider.office.com/join) .</span><span class="sxs-lookup"><span data-stu-id="4d0a1-107">You will need to join the [Office Insider](https://insider.office.com/join) program to try the preview features.</span></span>  <span data-ttu-id="4d0a1-108">Хороший способ ознакомиться с такими возможностями — использование подписки на Office 365.</span><span class="sxs-lookup"><span data-stu-id="4d0a1-108">A good way to try out preview features is by using an Office 365 subscription.</span></span> <span data-ttu-id="4d0a1-109">Если у вас еще нет подписки на Office 365, вы можете оформить бесплатную возобновляемую подписку на Office 365 на 90 дней, присоединившись к [программе для разработчиков Office 365](https://developer.microsoft.com/office/dev-program).</span><span class="sxs-lookup"><span data-stu-id="4d0a1-109">If you don't already have an Office 365 subscription, you can get a free, 90-day renewable Office 365 subscription by joining the [Office 365 Developer Program](https://developer.microsoft.com/office/dev-program).</span></span>

<span data-ttu-id="4d0a1-110">Если при выполнении пользовательской функции возникает ошибка, потребуется возвратить сообщение об ошибке, чтобы уведомить пользователя.</span><span class="sxs-lookup"><span data-stu-id="4d0a1-110">If something goes wrong while your custom function runs, you will need to return an error to inform the user.</span></span> <span data-ttu-id="4d0a1-111">Если у вас есть конкретные требования к параметрам, например применение только положительных чисел, нужно протестировать параметры и вернуть ошибку, если они неверны.</span><span class="sxs-lookup"><span data-stu-id="4d0a1-111">If you have specific parameter requirements, such as only positive numbers, you will need to test the parameters and throw an error if they are not correct.</span></span> <span data-ttu-id="4d0a1-112">Можно также использовать блок `try`-`catch`, чтобы отслеживать любые ошибки, возникающие при выполнении пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="4d0a1-112">You can also use a `try`-`catch` block to catch any errors that occur while your custom function runs.</span></span>

## <a name="detect-and-throw-an-error"></a><span data-ttu-id="4d0a1-113">Обнаружение и возвращение ошибки</span><span class="sxs-lookup"><span data-stu-id="4d0a1-113">Detect and throw an error</span></span>

<span data-ttu-id="4d0a1-114">Рассмотрим ситуацию, в которой необходимо убедиться, что параметр ZIP-кода имеет правильный формат, чтобы пользовательская функция работала.</span><span class="sxs-lookup"><span data-stu-id="4d0a1-114">Let's look at a case where you need to ensure that a zip code parameter is in the correct format for the custom function to work.</span></span> <span data-ttu-id="4d0a1-115">В следующей пользовательской функции используется регулярное выражение для проверки почтового индекса.</span><span class="sxs-lookup"><span data-stu-id="4d0a1-115">The following custom function uses a regular expression to check the zip code.</span></span> <span data-ttu-id="4d0a1-116">Если он правильный, будет подставлен город (в другой функции) и вернется значение.</span><span class="sxs-lookup"><span data-stu-id="4d0a1-116">If it is correct, then it will look up the city (in another function) and return the value.</span></span> <span data-ttu-id="4d0a1-117">В противном случае в ячейке возвращается ошибка `#VALUE!`.</span><span class="sxs-lookup"><span data-stu-id="4d0a1-117">If it is not correct, it returns a `#VALUE!` error to the cell.</span></span>

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

## <a name="the-customfunctionserror-object"></a><span data-ttu-id="4d0a1-118">Объект CustomFunctions.Error</span><span class="sxs-lookup"><span data-stu-id="4d0a1-118">The CustomFunctions.Error object</span></span>

<span data-ttu-id="4d0a1-119">Объект `CustomFunctions.Error` используется для возвращения ошибки в ячейку.</span><span class="sxs-lookup"><span data-stu-id="4d0a1-119">The `CustomFunctions.Error` object is used to return an error back to the cell.</span></span> <span data-ttu-id="4d0a1-120">При создании объекта укажите, какую ошибку нужно использовать, применив одно из следующих значений перечисления `ErrorCode`.</span><span class="sxs-lookup"><span data-stu-id="4d0a1-120">When you create the object, specify which error you want to use by using one of the following `ErrorCode` enum values.</span></span>


|<span data-ttu-id="4d0a1-121">Значение перечисления ErrorCode</span><span class="sxs-lookup"><span data-stu-id="4d0a1-121">ErrorCode enum value</span></span>  |<span data-ttu-id="4d0a1-122">Значение ячейки Excel</span><span class="sxs-lookup"><span data-stu-id="4d0a1-122">Excel cell value</span></span>  |<span data-ttu-id="4d0a1-123">Смысл</span><span class="sxs-lookup"><span data-stu-id="4d0a1-123">Meaning</span></span>  |
|---------------|---------|---------|
|`invalidValue`   | `#VALUE!` | <span data-ttu-id="4d0a1-124">В формуле используется значение неправильного типа.</span><span class="sxs-lookup"><span data-stu-id="4d0a1-124">A value used in the formula is the wrong type.</span></span> |
|`notAvailable`   | `#N/A`    | <span data-ttu-id="4d0a1-125">Функция или служба недоступна.</span><span class="sxs-lookup"><span data-stu-id="4d0a1-125">The function or service is not available.</span></span> |
|`divisionByZero` | `#DIV/0`  | <span data-ttu-id="4d0a1-126">Обратите внимание, что JavaScript позволяет делить на нуль, поэтому при создании обработчика ошибок нужно внимательно определить это условие.</span><span class="sxs-lookup"><span data-stu-id="4d0a1-126">Be aware that JavaScript allows division by zero so you need to write an error handler carefully to detect this condition.</span></span> |
|`invalidNumber`  | `#NUM!`   | <span data-ttu-id="4d0a1-127">Обнаружена проблема с числом, используемым в формуле</span><span class="sxs-lookup"><span data-stu-id="4d0a1-127">There is a problem with the number used in the formula</span></span> |
|`nullReference`  | `#NULL!`  | <span data-ttu-id="4d0a1-128">Диапазоны формулы не пересекаются.</span><span class="sxs-lookup"><span data-stu-id="4d0a1-128">The ranges in the formula do not intersect.</span></span> |

<span data-ttu-id="4d0a1-129">В следующем примере кода показано, как создать и вернуть ошибку для неверного числа (`#NUM!`).</span><span class="sxs-lookup"><span data-stu-id="4d0a1-129">The following code sample shows how to create and return an error for an invalid number (`#NUM!`).</span></span>

```typescript
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidNumber);
throw error;
```

<span data-ttu-id="4d0a1-130">При возврате ошибки `#VALUE!` также можно включить настраиваемое сообщение, отображаемое во всплывающем окне, когда пользователь наводит на ячейку указатель мыши.</span><span class="sxs-lookup"><span data-stu-id="4d0a1-130">When you return a `#VALUE!` error you can also include a custom message that will be shown in a popup when the user hovers over the cell.</span></span> <span data-ttu-id="4d0a1-131">В следующем примере показано, как вернуть настраиваемое сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="4d0a1-131">The following example shows how to return a custom error message.</span></span>

```typescript
// You can only return a custom error message with the #VALUE! error
let error = new CustomFunctions.Error(CustomFunctions.ErrorCode.invalidValue, "The parameter can only contain lowercase characters.");
throw error;
```

## <a name="use-try-catch-blocks"></a><span data-ttu-id="4d0a1-132">Использование блоков try-catch</span><span class="sxs-lookup"><span data-stu-id="4d0a1-132">Use try-catch blocks</span></span>

<span data-ttu-id="4d0a1-133">Как правило, для отслеживания любых возможных ошибок следует использовать блоки `try`-`catch` в пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="4d0a1-133">In general, you should use `try`-`catch` blocks in your custom function to catch any potential errors that occur.</span></span> <span data-ttu-id="4d0a1-134">Если в коде не обрабатываются исключения, они будут возвращаться в Excel.</span><span class="sxs-lookup"><span data-stu-id="4d0a1-134">If you do not handle exceptions in your code, they will be returned to Excel.</span></span> <span data-ttu-id="4d0a1-135">По умолчанию Excel возвращает `#VALUE!` для необработанного исключения.</span><span class="sxs-lookup"><span data-stu-id="4d0a1-135">By default, Excel returns `#VALUE!` for an unhandled exception.</span></span>

<span data-ttu-id="4d0a1-136">В следующем примере кода пользовательская функция создает запрос fetch в службу REST.</span><span class="sxs-lookup"><span data-stu-id="4d0a1-136">In the following code sample, the custom function makes a fetch call to a REST service.</span></span> <span data-ttu-id="4d0a1-137">Возможно, что вызов завершится сбоем (например, если служба REST возвращает ошибку или не работает сеть).</span><span class="sxs-lookup"><span data-stu-id="4d0a1-137">It's possible that the call will fail, for example, if the REST service returns an error or the network goes down.</span></span> <span data-ttu-id="4d0a1-138">В этом случае пользовательская функция возвращает `#N/A`, чтобы указать на сбой веб-вызова.</span><span class="sxs-lookup"><span data-stu-id="4d0a1-138">If this happens, the custom function will return `#N/A` to indicate the web call failed.</span></span>


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

## <a name="next-steps"></a><span data-ttu-id="4d0a1-139">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="4d0a1-139">Next steps</span></span>

<span data-ttu-id="4d0a1-140">Узнайте, как [устранять проблемы с пользовательскими функциями](custom-functions-troubleshooting.md).</span><span class="sxs-lookup"><span data-stu-id="4d0a1-140">Learn how to [troubleshoot problems with your custom functions](custom-functions-troubleshooting.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="4d0a1-141">Дополнительные ресурсы</span><span class="sxs-lookup"><span data-stu-id="4d0a1-141">See also</span></span>

* [<span data-ttu-id="4d0a1-142">Отладка пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="4d0a1-142">Custom functions debugging</span></span>](custom-functions-debugging.md)
* [<span data-ttu-id="4d0a1-143">Требования к настраиваемым функциям</span><span class="sxs-lookup"><span data-stu-id="4d0a1-143">Custom functions requirements</span></span>](custom-functions-requirement-sets.md)
* [<span data-ttu-id="4d0a1-144">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="4d0a1-144">Create custom functions in Excel</span></span>](custom-functions-overview.md)
