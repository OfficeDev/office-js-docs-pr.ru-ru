---
ms.date: 04/03/2019
description: Использование тегов JSDOC для динамического создания метаданных JSON пользовательских функций.
title: Создание метаданных JSON для пользовательских функций (предварительная версия)
localization_priority: Priority
ms.openlocfilehash: 2efe2a9a5a83ba60ef327273d5bd599f82916d48
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449255"
---
# <a name="create-json-metadata-for-custom-functions-preview"></a><span data-ttu-id="2e1e8-103">Создание метаданных JSON для пользовательских функций (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="2e1e8-103">Create JSON metadata for custom functions (preview)</span></span>

<span data-ttu-id="2e1e8-104">Если пользовательская функция Excel написана в JavaScript или TypeScript, теги JSDoc используются для предоставления дополнительной информации о пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-104">When an Excel custom function is written in JavaScript or TypeScript, JSDoc tags are used to provide extra information about the custom function.</span></span> <span data-ttu-id="2e1e8-105">Теги JSDoc используются при сборке для создания [файла метаданных JSON](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="2e1e8-105">The JSDoc tags are then used at build time to create the [JSON metadata file](custom-functions-json.md).</span></span> <span data-ttu-id="2e1e8-106">Использование тегов JSDoc освобождает вас от необходимости редактировать файл метаданных JSON вручную.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-106">Using JSDoc tags saves you from the effort of manually editing the JSON metadata file.</span></span>

<span data-ttu-id="2e1e8-107">Добавьте тег `@customfunction` в примечаниях к коду для функции JavaScript или TypeScript, чтобы пометить ее как пользовательскую.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-107">Add the `@customfunction` tag in the code comments for a JavaScript or TypeScript function to mark it as a custom function.</span></span>

<span data-ttu-id="2e1e8-108">Типы параметров функции можно получить с помощью тега [@param](#param) в JavaScript или из раздела [Тип функции](https://www.typescriptlang.org/docs/handbook/functions.html) в TypeScript.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-108">The function parameter types may be provided using the [@param](#param) tag in JavaScript, or from the [Function type](https://www.typescriptlang.org/docs/handbook/functions.html) in TypeScript.</span></span> <span data-ttu-id="2e1e8-109">Дополнительную информацию см. в теге [@param](#param) и разделе [Типы](#types).</span><span class="sxs-lookup"><span data-stu-id="2e1e8-109">For more information, see the [@param](#param) tag and [Types](#types) section.</span></span>

## <a name="jsdoc-tags"></a><span data-ttu-id="2e1e8-110">Теги JSDoc</span><span class="sxs-lookup"><span data-stu-id="2e1e8-110">JSDoc Tags</span></span>
<span data-ttu-id="2e1e8-111">Ниже приведены теги JSDoc, которые поддерживаются в пользовательских функциях Excel:</span><span class="sxs-lookup"><span data-stu-id="2e1e8-111">The following JSDoc tags are supported in Excel custom functions:</span></span>
* [<span data-ttu-id="2e1e8-112">@cancelable</span><span class="sxs-lookup"><span data-stu-id="2e1e8-112">@cancelable</span></span>](#cancelable)
* <span data-ttu-id="2e1e8-113">[@customfunction](#customfunction) id name</span><span class="sxs-lookup"><span data-stu-id="2e1e8-113">[@customfunction](#customfunction) id name</span></span>
* <span data-ttu-id="2e1e8-114">[@helpurl](#helpurl) url</span><span class="sxs-lookup"><span data-stu-id="2e1e8-114">[@helpurl](#helpurl) url</span></span>
* <span data-ttu-id="2e1e8-115">[@param](#param) _{type}_ name description</span><span class="sxs-lookup"><span data-stu-id="2e1e8-115">[@param](#param) _{type}_ name description</span></span>
* [<span data-ttu-id="2e1e8-116">@requiresAddress</span><span class="sxs-lookup"><span data-stu-id="2e1e8-116">@requiresAddress</span></span>](#requiresAddress)
* <span data-ttu-id="2e1e8-117">[@returns](#returns) _{type}_</span><span class="sxs-lookup"><span data-stu-id="2e1e8-117">[@returns](#returns) _{type}_</span></span>
* [<span data-ttu-id="2e1e8-118">@streaming</span><span class="sxs-lookup"><span data-stu-id="2e1e8-118">@streaming</span></span>](#streaming)
* [<span data-ttu-id="2e1e8-119">@volatile</span><span class="sxs-lookup"><span data-stu-id="2e1e8-119">@volatile</span></span>](#volatile)

---
### <a name="cancelable"></a><span data-ttu-id="2e1e8-120">@cancelable</span><span class="sxs-lookup"><span data-stu-id="2e1e8-120">@cancelable</span></span>
<a id="cancelable"/>

<span data-ttu-id="2e1e8-121">При отмене функции указывает, что пользовательская функция стремится к выполнению действия.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-121">Indicates that a custom function wants to perform an action when the function is canceled.</span></span>

<span data-ttu-id="2e1e8-122">В качестве типа последнего параметра функции должно быть указано `CustomFunctions.CancelableInvocation`.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-122">The last function parameter must be of type `CustomFunctions.CancelableInvocation`.</span></span> <span data-ttu-id="2e1e8-123">Функция может назначить функцию свойству `oncanceled`, чтобы обозначить действия для выполнения в случае отмены функции.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-123">The function can assign a function to the `oncanceled` property to denote the action to perform when the function is canceled.</span></span>

<span data-ttu-id="2e1e8-124">Если тип последнего параметра функции `CustomFunctions.CancelableInvocation`, он будет рассматриваться как `@cancelable`, даже если тег отсутствует.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-124">If the last function parameter is of type `CustomFunctions.CancelableInvocation`, it will be considered `@cancelable` even if the tag is not present.</span></span>

<span data-ttu-id="2e1e8-125">Функция не может содержать одновременно теги `@cancelable` и `@streaming`.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-125">A function cannot have both `@cancelable` and `@streaming` tags.</span></span>

---
### <a name="customfunction"></a><span data-ttu-id="2e1e8-126">@customfunction</span><span class="sxs-lookup"><span data-stu-id="2e1e8-126">@customfunction</span></span>
<a id="customfunction"/>

<span data-ttu-id="2e1e8-127">Синтаксис: @customfunction _id_ _name_</span><span class="sxs-lookup"><span data-stu-id="2e1e8-127">Syntax: @customfunction _id_ _name_</span></span>

<span data-ttu-id="2e1e8-128">Укажите этот тег, чтобы рассматривать функцию JavaScript или TypeScript как пользовательскую функцию Excel.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-128">Specify this tag to treat the JavaScript/TypeScript function as an Excel custom function.</span></span>

<span data-ttu-id="2e1e8-129">Этот тег необходим, чтобы создать метаданные для пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-129">This tag is required to create metadata for the custom function.</span></span>

<span data-ttu-id="2e1e8-130">Кроме того, требуется вызов функции `CustomFunctions.associate("id", functionName);`</span><span class="sxs-lookup"><span data-stu-id="2e1e8-130">There should also be a call to `CustomFunctions.associate("id", functionName);`</span></span>

#### <a name="id"></a><span data-ttu-id="2e1e8-131">id</span><span class="sxs-lookup"><span data-stu-id="2e1e8-131">id</span></span> 

<span data-ttu-id="2e1e8-132">Идентификатор используется как инвариантный идентификатор для пользовательских функций, которые хранятся в документе.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-132">The id is used as the invariant identifier for the custom function stored in the document.</span></span> <span data-ttu-id="2e1e8-133">Его не следует менять.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-133">It should not change.</span></span>

* <span data-ttu-id="2e1e8-134">Если идентификатор не указан, название функции JavaScript или TypeScript преобразуется в верхний регистр, а недопустимые символы удаляются.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-134">If id is not provided, the JavaScript/TypeScript function name is converted to uppercase, disallowed characters are removed.</span></span>
* <span data-ttu-id="2e1e8-135">Идентификатор должен быть уникальным для всех пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-135">The id must be unique for all custom functions.</span></span>
* <span data-ttu-id="2e1e8-136">Допустимые символы: A–Z, a–z, 0–9 и точки (.).</span><span class="sxs-lookup"><span data-stu-id="2e1e8-136">The characters allowed are limited to: A-Z, a-z, 0-9, and period (.).</span></span>

#### <a name="name"></a><span data-ttu-id="2e1e8-137">name</span><span class="sxs-lookup"><span data-stu-id="2e1e8-137">name</span></span>

<span data-ttu-id="2e1e8-138">Предоставляет отображаемое имя для пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-138">Provides the display name for the custom function.</span></span> 

* <span data-ttu-id="2e1e8-139">Если имя не указано, идентификатор тоже используется как имя.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-139">If name is not provided, the id is also used as the name.</span></span>
* <span data-ttu-id="2e1e8-140">Допустимые символы: буквы [буквенные символы Юникод](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), числа, точки (.) и подчеркивания (\_).</span><span class="sxs-lookup"><span data-stu-id="2e1e8-140">Allowed characters: Letters [Unicode Alphabetic character](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), numbers, period (.), and underscore (\_).</span></span>
* <span data-ttu-id="2e1e8-141">Имя должно начинаться с буквы.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-141">Must start with a letter.</span></span>
* <span data-ttu-id="2e1e8-142">Максимальная длина: 128 символов.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-142">Maximum length is 128 characters.</span></span>

---
### <a name="helpurl"></a><span data-ttu-id="2e1e8-143">@helpurl</span><span class="sxs-lookup"><span data-stu-id="2e1e8-143">@helpurl</span></span>
<a id="helpurl"/>

<span data-ttu-id="2e1e8-144">Синтаксис: @helpurl _url_</span><span class="sxs-lookup"><span data-stu-id="2e1e8-144">Syntax: @helpurl _url_</span></span>

<span data-ttu-id="2e1e8-145">Предоставленный _url_-адрес отображается в Excel.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-145">The provided _url_ is displayed in Excel.</span></span>

---
### <a name="param"></a><span data-ttu-id="2e1e8-146">@param</span><span class="sxs-lookup"><span data-stu-id="2e1e8-146">@param</span></span>
<a id="param"/>

#### <a name="javascript"></a><span data-ttu-id="2e1e8-147">JavaScript</span><span class="sxs-lookup"><span data-stu-id="2e1e8-147">JavaScript</span></span>

<span data-ttu-id="2e1e8-148">Синтаксис JavaScript: @param {type} name _description_</span><span class="sxs-lookup"><span data-stu-id="2e1e8-148">JavaScript Syntax: @param {type} name _description_</span></span>

* <span data-ttu-id="2e1e8-149">`{type}` должен указывать информацию о типе в фигурных скобках.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-149">`{type}` should specify the type info within curly braces.</span></span> <span data-ttu-id="2e1e8-150">Дополнительную информацию о типах, которые могут использоваться, см. в разделе [Типы](##types).</span><span class="sxs-lookup"><span data-stu-id="2e1e8-150">See the [Types](##types) for more information about the types which may be used.</span></span> <span data-ttu-id="2e1e8-151">Необязательно: если тип не указан, будет использоваться тип `any`.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-151">Optional: if not specified, the type `any` will be used.</span></span>
* <span data-ttu-id="2e1e8-152">`name` указывает, к какому параметру относится тег @param.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-152">`name` specifies which parameter the @param tag applies to.</span></span> <span data-ttu-id="2e1e8-153">Обязательно.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-153">Required.</span></span>
* <span data-ttu-id="2e1e8-154">`description` предоставляет описание, которое отображается в Excel для параметра функции.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-154">`description` provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="2e1e8-155">Необязательно.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-155">Optional.</span></span>

<span data-ttu-id="2e1e8-156">Чтобы обозначить параметр пользовательской функции как необязательный:</span><span class="sxs-lookup"><span data-stu-id="2e1e8-156">To denote a custom function parameter as optional:</span></span>
* <span data-ttu-id="2e1e8-157">Поместите имя параметра в квадратные скобки.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-157">Put square brackets around the parameter name.</span></span> <span data-ttu-id="2e1e8-158">Пример: `@param {string} [text] Optional text`.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-158">For example: `@param {string} [text] Optional text`.</span></span>

#### <a name="typescript"></a><span data-ttu-id="2e1e8-159">TypeScript</span><span class="sxs-lookup"><span data-stu-id="2e1e8-159">TypeScript</span></span>

<span data-ttu-id="2e1e8-160">Синтаксис TypeScript: @param name _description_</span><span class="sxs-lookup"><span data-stu-id="2e1e8-160">TypeScript Syntax: @param name _description_</span></span>

* <span data-ttu-id="2e1e8-161">`name` указывает, к какому параметру относится тег @param.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-161">`name` specifies which parameter the @param tag applies to.</span></span> <span data-ttu-id="2e1e8-162">Обязательно.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-162">Required.</span></span>
* <span data-ttu-id="2e1e8-163">`description` предоставляет описание, которое отображается в Excel для параметра функции.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-163">`description` provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="2e1e8-164">Необязательно.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-164">Optional.</span></span>

<span data-ttu-id="2e1e8-165">Дополнительную информацию о типах параметров функций, которые могут использоваться, см. в разделе [Типы](##types).</span><span class="sxs-lookup"><span data-stu-id="2e1e8-165">See the [Types](##types) for more information about the function parameter types which may be used.</span></span>

<span data-ttu-id="2e1e8-166">Чтобы обозначить параметр пользовательской функции как необязательный, выполните одно из указанных ниже действий.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-166">To denote a custom function parameter as optional, do one of the following:</span></span>
* <span data-ttu-id="2e1e8-167">Используйте необязательный параметр.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-167">Use an optional parameter.</span></span> <span data-ttu-id="2e1e8-168">Пример: `function f(text?: string)`</span><span class="sxs-lookup"><span data-stu-id="2e1e8-168">For example: `function f(text?: string)`</span></span>
* <span data-ttu-id="2e1e8-169">Задайте для параметра значение по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-169">Give the parameter a default value.</span></span> <span data-ttu-id="2e1e8-170">Пример: `function f(text: string = "abc")`</span><span class="sxs-lookup"><span data-stu-id="2e1e8-170">For example: `function f(text: string = "abc")`</span></span>

<span data-ttu-id="2e1e8-171">Подробное описание @param см. в [JSDoc](http://usejsdoc.org/tags-param.html)</span><span class="sxs-lookup"><span data-stu-id="2e1e8-171">For detailed description of the @param see: [JSDoc](http://usejsdoc.org/tags-param.html)</span></span>

---
### <a name="requiresaddress"></a><span data-ttu-id="2e1e8-172">@requiresAddress</span><span class="sxs-lookup"><span data-stu-id="2e1e8-172">@requiresAddress</span></span>
<a id="requiresAddress"/>

<span data-ttu-id="2e1e8-173">Указывает, что следует предоставлять адрес ячейки, в которой вычисляется функция.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-173">Indicates that the address of the cell where the function is being evaluated should be provided.</span></span> 

<span data-ttu-id="2e1e8-174">Тип последнего параметра функции должен быть `CustomFunctions.Invocation` или производной от него.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-174">The last function parameter must be of type `CustomFunctions.Invocation` or a derived type.</span></span> <span data-ttu-id="2e1e8-175">При вызове функции свойство `address` будет содержать адрес.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-175">When the function is called, the `address` property will contain the address.</span></span>

---
### <a name="returns"></a><span data-ttu-id="2e1e8-176">@returns</span><span class="sxs-lookup"><span data-stu-id="2e1e8-176">@returns</span></span>
<a id="returns"/>

<span data-ttu-id="2e1e8-177">Синтаксис: @returns {_type_}</span><span class="sxs-lookup"><span data-stu-id="2e1e8-177">Syntax: @returns {_type_}</span></span>

<span data-ttu-id="2e1e8-178">Предоставляет тип для возвращаемого значения.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-178">Provides the type for the return value.</span></span>

<span data-ttu-id="2e1e8-179">Если `{type}` не указан, будет использоваться информация о типе TypeScript.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-179">If `{type}` is omitted, the TypeScript type info will be used.</span></span> <span data-ttu-id="2e1e8-180">Если информация о типе отсутствует, будет использоваться тип `any`.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-180">If there is no type info, the type will be `any`.</span></span>

---
### <a name="streaming"></a><span data-ttu-id="2e1e8-181">@streaming</span><span class="sxs-lookup"><span data-stu-id="2e1e8-181">@streaming</span></span>
<a id="streaming"/>

<span data-ttu-id="2e1e8-182">Используется для обозначения того, что пользовательская функция является потоковой передачей функции.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-182">Used to indicate that a custom function is a streaming function.</span></span> 

<span data-ttu-id="2e1e8-183">Тип последнего параметра должен быть `CustomFunctions.StreamingInvocation<ResultType>`.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-183">The last parameter should be of type `CustomFunctions.StreamingInvocation<ResultType>`.</span></span>
<span data-ttu-id="2e1e8-184">Функция должна вернуть значение `void`.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-184">The function should return `void`.</span></span>

<span data-ttu-id="2e1e8-185">Потоковые передачи функций непосредственно не возвращают значения, для этого необходимо вызывать `setResult(result: ResultType)` с помощью последнего параметра.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-185">Streaming functions do not return values directly, but rather should call `setResult(result: ResultType)` using the last parameter.</span></span>

<span data-ttu-id="2e1e8-186">Исключения, которые возникают при потоковой передаче функций, игнорируются.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-186">Exceptions thrown by a streaming function are ignored.</span></span> <span data-ttu-id="2e1e8-187">`setResult()` при вызове может вернуть ошибку в качестве результата.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-187">`setResult()` may be called with Error to indicate an error result.</span></span>

<span data-ttu-id="2e1e8-188">Потоковые передачи функций невозможно пометить как [@volatile](#volatile).</span><span class="sxs-lookup"><span data-stu-id="2e1e8-188">Streaming functions cannot be marked as [@volatile](#volatile).</span></span>

---
### <a name="volatile"></a><span data-ttu-id="2e1e8-189">@volatile</span><span class="sxs-lookup"><span data-stu-id="2e1e8-189">@volatile</span></span>
<a id="volatile"/>

<span data-ttu-id="2e1e8-190">Переменные функции — это такие функции, чей результат не остается неизменным в каждый период времени, даже если они не содержат аргументов или их аргументы не меняются.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-190">A volatile function is one whose result cannot be assumed to be the same from one moment to the next even if it takes no arguments or the arguments have not changed.</span></span> <span data-ttu-id="2e1e8-191">Excel повторно проводит вычисления в ячейках, которые содержат переменные функции, вместе со всеми зависимыми функциями при каждом вычислении.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-191">Excel re-evaluates cells that contain volatile functions, together with all dependents, every time that a calculation is done.</span></span> <span data-ttu-id="2e1e8-192">По этой причине чрезмерное использование переменных функций может замедлить пересчет, поэтому используйте их умеренно.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-192">For this reason, too much reliance on volatile functions can make recalculation times slow, so use them sparingly.</span></span>

<span data-ttu-id="2e1e8-193">Потоковые передачи функций не могут быть переменными.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-193">Streaming functions cannot be volatile.</span></span>

---

## <a name="types"></a><span data-ttu-id="2e1e8-194">Типы</span><span class="sxs-lookup"><span data-stu-id="2e1e8-194">Types</span></span>

<span data-ttu-id="2e1e8-195">Указывая тип параметра, Excel преобразует значения в этот тип, прежде чем вызывать функцию.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-195">By specifying a parameter type, Excel will convert values into that type before calling the function.</span></span> <span data-ttu-id="2e1e8-196">Если указан тип `any`, преобразование выполняться не будет.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-196">If the type is `any`, no conversion will be performed.</span></span>

### <a name="value-types"></a><span data-ttu-id="2e1e8-197">Типы значений</span><span class="sxs-lookup"><span data-stu-id="2e1e8-197">Value types</span></span>

<span data-ttu-id="2e1e8-198">Одно значение может быть представлено с помощью одного из приведенных ниже типов: `boolean`, `number`, `string`.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-198">A single value may be represented using one of the following types: `boolean`, `number`, `string`.</span></span>

### <a name="matrix-type"></a><span data-ttu-id="2e1e8-199">Тип "матрица"</span><span class="sxs-lookup"><span data-stu-id="2e1e8-199">Matrix type</span></span>

<span data-ttu-id="2e1e8-200">Используйте тип двумерного массива, чтобы параметр или возвращаемое значение представляли собой матрицу значений.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-200">Use a two-dimensional array type to have the parameter or return value be a matrix of values.</span></span> <span data-ttu-id="2e1e8-201">Например, тип `number[][]` указывает на матрицу чисел.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-201">For example, the type `number[][]` indicates a matrix of numbers.</span></span> <span data-ttu-id="2e1e8-202">`string[][]` указывает на матрицу строк.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-202">`string[][]` indicates a matrix of strings.</span></span> 

### <a name="error-type"></a><span data-ttu-id="2e1e8-203">Тип "ошибка"</span><span class="sxs-lookup"><span data-stu-id="2e1e8-203">Error type</span></span>

<span data-ttu-id="2e1e8-204">Функция непотоковой передачи может указывать на ошибку, возвращая тип "Ошибка".</span><span class="sxs-lookup"><span data-stu-id="2e1e8-204">A non-streaming function can indicate an error by returning an Error type.</span></span>

<span data-ttu-id="2e1e8-205">Функция потоковой передачи может указывать на ошибку, вызывая метод setResult() типа "Ошибка".</span><span class="sxs-lookup"><span data-stu-id="2e1e8-205">A streaming function can indicate an error by calling setResult() with an Error type.</span></span>

### <a name="promise"></a><span data-ttu-id="2e1e8-206">Обещание</span><span class="sxs-lookup"><span data-stu-id="2e1e8-206">Promise</span></span>

<span data-ttu-id="2e1e8-207">Функция может вернуть тип "Обещание", который задаст значение, когда обещание будет разрешено.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-207">A function can return a Promise, which will provide the value when the promise is resolved.</span></span> <span data-ttu-id="2e1e8-208">В случае отклонения обещания возникнет ошибка.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-208">If the promise is rejected, then it is an error.</span></span>

### <a name="other-types"></a><span data-ttu-id="2e1e8-209">Другие типы</span><span class="sxs-lookup"><span data-stu-id="2e1e8-209">Other types</span></span>

<span data-ttu-id="2e1e8-210">Любой другой тип будет рассматриваться как ошибка.</span><span class="sxs-lookup"><span data-stu-id="2e1e8-210">Any other type will be treated as an error.</span></span>

## <a name="see-also"></a><span data-ttu-id="2e1e8-211">См. также</span><span class="sxs-lookup"><span data-stu-id="2e1e8-211">See also</span></span>

* [<span data-ttu-id="2e1e8-212">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="2e1e8-212">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="2e1e8-213">Среда выполнения для пользовательских функций Excel</span><span class="sxs-lookup"><span data-stu-id="2e1e8-213">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="2e1e8-214">Рекомендации по пользовательским функциям</span><span class="sxs-lookup"><span data-stu-id="2e1e8-214">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="2e1e8-215">Журнал изменений пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="2e1e8-215">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="2e1e8-216">Руководство по пользовательским функциям в Excel</span><span class="sxs-lookup"><span data-stu-id="2e1e8-216">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="2e1e8-217">Отладка пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="2e1e8-217">Custom functions debugging</span></span>](custom-functions-debugging.md)
