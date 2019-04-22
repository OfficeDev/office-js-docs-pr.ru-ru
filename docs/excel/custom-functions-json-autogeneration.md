---
ms.date: 04/03/2019
description: Использование тегов JSDOC для динамического создания метаданных JSON пользовательских функций.
title: Создание метаданных JSON для пользовательских функций (предварительная версия)
localization_priority: Priority
ms.openlocfilehash: 2efe2a9a5a83ba60ef327273d5bd599f82916d48
ms.sourcegitcommit: 6d375518c119d09c8d3fb5f0cc4583ba5b20ac03
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/18/2019
ms.locfileid: "31914286"
---
# <a name="create-json-metadata-for-custom-functions-preview"></a><span data-ttu-id="374ca-103">Создание метаданных JSON для пользовательских функций (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="374ca-103">Create JSON metadata for custom functions (preview)</span></span>

<span data-ttu-id="374ca-104">Если пользовательская функция Excel написана в JavaScript или TypeScript, теги JSDoc используются для предоставления дополнительной информации о пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="374ca-104">When an Excel custom function is written in JavaScript or TypeScript, JSDoc tags are used to provide extra information about the custom function.</span></span> <span data-ttu-id="374ca-105">Теги JSDoc используются при сборке для создания [файла метаданных JSON](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="374ca-105">The JSDoc tags are then used at build time to create the [JSON metadata file](custom-functions-json.md).</span></span> <span data-ttu-id="374ca-106">Использование тегов JSDoc освобождает вас от необходимости редактировать файл метаданных JSON вручную.</span><span class="sxs-lookup"><span data-stu-id="374ca-106">Using JSDoc tags saves you from the effort of manually editing the JSON metadata file.</span></span>

<span data-ttu-id="374ca-107">Добавьте тег `@customfunction` в примечаниях к коду для функции JavaScript или TypeScript, чтобы пометить ее как пользовательскую.</span><span class="sxs-lookup"><span data-stu-id="374ca-107">Add the `@customfunction` tag in the code comments for a JavaScript or TypeScript function to mark it as a custom function.</span></span>

<span data-ttu-id="374ca-108">Типы параметров функции можно получить с помощью тега [@param](#param) в JavaScript или из раздела [Тип функции](https://www.typescriptlang.org/docs/handbook/functions.html) в TypeScript.</span><span class="sxs-lookup"><span data-stu-id="374ca-108">The function parameter types may be provided using the    tag in JavaScript, or from the Function type in TypeScript.</span></span> <span data-ttu-id="374ca-109">Дополнительную информацию см. в теге [@param](#param) и разделе [Типы](#types).</span><span class="sxs-lookup"><span data-stu-id="374ca-109">For more information, see the    tag and Types section.</span></span>

## <a name="jsdoc-tags"></a><span data-ttu-id="374ca-110">Теги JSDoc</span><span class="sxs-lookup"><span data-stu-id="374ca-110">JSDoc Tags</span></span>
<span data-ttu-id="374ca-111">Ниже приведены теги JSDoc, которые поддерживаются в пользовательских функциях Excel:</span><span class="sxs-lookup"><span data-stu-id="374ca-111">The following JSDoc tags are supported in Excel custom functions:</span></span>
* [<span data-ttu-id="374ca-112">@cancelable</span><span class="sxs-lookup"><span data-stu-id="374ca-112">@cancelable</span></span>](#cancelable)
* <span data-ttu-id="374ca-113">[@customfunction](#customfunction) id name</span><span class="sxs-lookup"><span data-stu-id="374ca-113">[@customfunction](#customfunction) id name</span></span>
* <span data-ttu-id="374ca-114">[@helpurl](#helpurl) url</span><span class="sxs-lookup"><span data-stu-id="374ca-114">[@helpurl](#helpurl) url</span></span>
* <span data-ttu-id="374ca-115">[@param](#param) _{type}_ name description</span><span class="sxs-lookup"><span data-stu-id="374ca-115">   {type} name description</span></span>
* [<span data-ttu-id="374ca-116">@requiresAddress</span><span class="sxs-lookup"><span data-stu-id="374ca-116">@requiresAddress</span></span>](#requiresAddress)
* <span data-ttu-id="374ca-117">[@returns](#returns) _{type}_</span><span class="sxs-lookup"><span data-stu-id="374ca-117">[@returns](#returns) _{type}_</span></span>
* [<span data-ttu-id="374ca-118">@streaming</span><span class="sxs-lookup"><span data-stu-id="374ca-118">streaming</span></span>](#streaming)
* [<span data-ttu-id="374ca-119">@volatile</span><span class="sxs-lookup"><span data-stu-id="374ca-119">Volatile</span></span>](#volatile)

---
### <a name="cancelable"></a><span data-ttu-id="374ca-120">@cancelable</span><span class="sxs-lookup"><span data-stu-id="374ca-120">@cancelable</span></span>
<a id="cancelable"/>

<span data-ttu-id="374ca-121">При отмене функции указывает, что пользовательская функция стремится к выполнению действия.</span><span class="sxs-lookup"><span data-stu-id="374ca-121">Indicates that a custom function wants to perform an action when the function is canceled.</span></span>

<span data-ttu-id="374ca-122">В качестве типа последнего параметра функции должно быть указано `CustomFunctions.CancelableInvocation`.</span><span class="sxs-lookup"><span data-stu-id="374ca-122">The last function parameter must be of type `CustomFunctions.CancelableInvocation`.</span></span> <span data-ttu-id="374ca-123">Функция может назначить функцию свойству `oncanceled`, чтобы обозначить действия для выполнения в случае отмены функции.</span><span class="sxs-lookup"><span data-stu-id="374ca-123">The function can assign a function to the `oncanceled` property to denote the action to perform when the function is canceled.</span></span>

<span data-ttu-id="374ca-124">Если тип последнего параметра функции `CustomFunctions.CancelableInvocation`, он будет рассматриваться как `@cancelable`, даже если тег отсутствует.</span><span class="sxs-lookup"><span data-stu-id="374ca-124">If the last function parameter is of type `CustomFunctions.CancelableInvocation`, it will be considered `@cancelable` even if the tag is not present.</span></span>

<span data-ttu-id="374ca-125">Функция не может содержать одновременно теги `@cancelable` и `@streaming`.</span><span class="sxs-lookup"><span data-stu-id="374ca-125">A function cannot have both `@cancelable` and `@streaming` tags.</span></span>

---
### <a name="customfunction"></a><span data-ttu-id="374ca-126">@customfunction</span><span class="sxs-lookup"><span data-stu-id="374ca-126">@customfunction</span></span>
<a id="customfunction"/>

<span data-ttu-id="374ca-127">Синтаксис: @customfunction _id_ _name_</span><span class="sxs-lookup"><span data-stu-id="374ca-127">Syntax:  id name</span></span>

<span data-ttu-id="374ca-128">Укажите этот тег, чтобы рассматривать функцию JavaScript или TypeScript как пользовательскую функцию Excel.</span><span class="sxs-lookup"><span data-stu-id="374ca-128">Specify this tag to treat the JavaScript/TypeScript function as an Excel custom function.</span></span>

<span data-ttu-id="374ca-129">Этот тег необходим, чтобы создать метаданные для пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="374ca-129">This tag is required to create metadata for the custom function.</span></span>

<span data-ttu-id="374ca-130">Кроме того, требуется вызов функции `CustomFunctions.associate("id", functionName);`</span><span class="sxs-lookup"><span data-stu-id="374ca-130">There should also be a call to`CustomFunctions.associate("id", functionName);`</span></span>

#### <a name="id"></a><span data-ttu-id="374ca-131">id</span><span class="sxs-lookup"><span data-stu-id="374ca-131">id</span></span> 

<span data-ttu-id="374ca-132">Идентификатор используется как инвариантный идентификатор для пользовательских функций, которые хранятся в документе.</span><span class="sxs-lookup"><span data-stu-id="374ca-132">The id is used as the invariant identifier for the custom function stored in the document.</span></span> <span data-ttu-id="374ca-133">Его не следует менять.</span><span class="sxs-lookup"><span data-stu-id="374ca-133">It should not change.</span></span>

* <span data-ttu-id="374ca-134">Если идентификатор не указан, название функции JavaScript или TypeScript преобразуется в верхний регистр, а недопустимые символы удаляются.</span><span class="sxs-lookup"><span data-stu-id="374ca-134">If id is not provided, the JavaScript/TypeScript function name is converted to uppercase, disallowed characters are removed.</span></span>
* <span data-ttu-id="374ca-135">Идентификатор должен быть уникальным для всех пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="374ca-135">The id must be unique for all custom functions.</span></span>
* <span data-ttu-id="374ca-136">Допустимые символы: A–Z, a–z, 0–9 и точки (.).</span><span class="sxs-lookup"><span data-stu-id="374ca-136">The characters allowed are limited to: A-Z, a-z, 0-9, and period (.).</span></span>

#### <a name="name"></a><span data-ttu-id="374ca-137">name</span><span class="sxs-lookup"><span data-stu-id="374ca-137">name</span></span>

<span data-ttu-id="374ca-138">Предоставляет отображаемое имя для пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="374ca-138">Provides the display name for the custom function.</span></span> 

* <span data-ttu-id="374ca-139">Если имя не указано, идентификатор тоже используется как имя.</span><span class="sxs-lookup"><span data-stu-id="374ca-139">If name is not provided, the id is also used as the name.</span></span>
* <span data-ttu-id="374ca-140">Допустимые символы: буквы [буквенные символы Юникод](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), числа, точки (.) и подчеркивания (\_).</span><span class="sxs-lookup"><span data-stu-id="374ca-140">Allowed characters: Letters [Unicode Alphabetic character](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), numbers, period (.), and underscore (\_).</span></span>
* <span data-ttu-id="374ca-141">Имя должно начинаться с буквы.</span><span class="sxs-lookup"><span data-stu-id="374ca-141">Must start with a letter.</span></span>
* <span data-ttu-id="374ca-142">Максимальная длина: 128 символов.</span><span class="sxs-lookup"><span data-stu-id="374ca-142">Maximum length is 128 characters.</span></span>

---
### <a name="helpurl"></a><span data-ttu-id="374ca-143">@helpurl</span><span class="sxs-lookup"><span data-stu-id="374ca-143">helpUrl</span></span>
<a id="helpurl"/>

<span data-ttu-id="374ca-144">Синтаксис: @helpurl _url_</span><span class="sxs-lookup"><span data-stu-id="374ca-144">Syntax: @helpurl _url_</span></span>

<span data-ttu-id="374ca-145">Предоставленный _url_-адрес отображается в Excel.</span><span class="sxs-lookup"><span data-stu-id="374ca-145">The provided _url_ is displayed in Excel.</span></span>

---
### <a name="param"></a><span data-ttu-id="374ca-146">@param</span><span class="sxs-lookup"><span data-stu-id="374ca-146">param</span></span>
<a id="param"/>

#### <a name="javascript"></a><span data-ttu-id="374ca-147">JavaScript</span><span class="sxs-lookup"><span data-stu-id="374ca-147">JavaScript</span></span>

<span data-ttu-id="374ca-148">Синтаксис JavaScript: @param {type} name _description_</span><span class="sxs-lookup"><span data-stu-id="374ca-148">JavaScript Syntax:  {type} name description</span></span>

* <span data-ttu-id="374ca-149">`{type}` должен указывать информацию о типе в фигурных скобках.</span><span class="sxs-lookup"><span data-stu-id="374ca-149">`{type}`should specify the type info within curly braces.</span></span> <span data-ttu-id="374ca-150">Дополнительную информацию о типах, которые могут использоваться, см. в разделе [Типы](##types).</span><span class="sxs-lookup"><span data-stu-id="374ca-150">See the [Types](##types) for more information about the types which may be used.</span></span> <span data-ttu-id="374ca-151">Необязательно: если тип не указан, будет использоваться тип `any`.</span><span class="sxs-lookup"><span data-stu-id="374ca-151">Optional: if not specified, the type `any` will be used.</span></span>
* <span data-ttu-id="374ca-152">`name` указывает, к какому параметру относится тег @param.</span><span class="sxs-lookup"><span data-stu-id="374ca-152">specifies which parameter the `name` tag applies to.</span></span> <span data-ttu-id="374ca-153">Обязательно.</span><span class="sxs-lookup"><span data-stu-id="374ca-153">Required.</span></span>
* <span data-ttu-id="374ca-154">`description` предоставляет описание, которое отображается в Excel для параметра функции.</span><span class="sxs-lookup"><span data-stu-id="374ca-154">`description`provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="374ca-155">Необязательно.</span><span class="sxs-lookup"><span data-stu-id="374ca-155">Optional.</span></span>

<span data-ttu-id="374ca-156">Чтобы обозначить параметр пользовательской функции как необязательный:</span><span class="sxs-lookup"><span data-stu-id="374ca-156">To denote a custom function parameter as optional:</span></span>
* <span data-ttu-id="374ca-157">Поместите имя параметра в квадратные скобки.</span><span class="sxs-lookup"><span data-stu-id="374ca-157">Put square brackets around the parameter name.</span></span> <span data-ttu-id="374ca-158">Пример: `@param {string} [text] Optional text`.</span><span class="sxs-lookup"><span data-stu-id="374ca-158">For example: `@param {string} [text] Optional text`.</span></span>

#### <a name="typescript"></a><span data-ttu-id="374ca-159">TypeScript</span><span class="sxs-lookup"><span data-stu-id="374ca-159">TypeScript</span></span>

<span data-ttu-id="374ca-160">Синтаксис TypeScript: @param name _description_</span><span class="sxs-lookup"><span data-stu-id="374ca-160">TypeScript Syntax:  name description</span></span>

* <span data-ttu-id="374ca-161">`name` указывает, к какому параметру относится тег @param.</span><span class="sxs-lookup"><span data-stu-id="374ca-161">specifies which parameter the `name` tag applies to.</span></span> <span data-ttu-id="374ca-162">Обязательно.</span><span class="sxs-lookup"><span data-stu-id="374ca-162">Required.</span></span>
* <span data-ttu-id="374ca-163">`description` предоставляет описание, которое отображается в Excel для параметра функции.</span><span class="sxs-lookup"><span data-stu-id="374ca-163">`description`provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="374ca-164">Необязательно.</span><span class="sxs-lookup"><span data-stu-id="374ca-164">Optional.</span></span>

<span data-ttu-id="374ca-165">Дополнительную информацию о типах параметров функций, которые могут использоваться, см. в разделе [Типы](##types).</span><span class="sxs-lookup"><span data-stu-id="374ca-165">See the [Types](##types) for more information about the function parameter types which may be used.</span></span>

<span data-ttu-id="374ca-166">Чтобы обозначить параметр пользовательской функции как необязательный, выполните одно из указанных ниже действий.</span><span class="sxs-lookup"><span data-stu-id="374ca-166">To denote a custom function parameter as optional, do one of the following:</span></span>
* <span data-ttu-id="374ca-167">Используйте необязательный параметр.</span><span class="sxs-lookup"><span data-stu-id="374ca-167">Use an optional parameter.</span></span> <span data-ttu-id="374ca-168">Пример: `function f(text?: string)`</span><span class="sxs-lookup"><span data-stu-id="374ca-168">For example: `function f(text?: string)`</span></span>
* <span data-ttu-id="374ca-169">Задайте для параметра значение по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="374ca-169">Give the parameter a default value.</span></span> <span data-ttu-id="374ca-170">Пример: `function f(text: string = "abc")`</span><span class="sxs-lookup"><span data-stu-id="374ca-170">For example: `function f(text: string = "abc")`</span></span>

<span data-ttu-id="374ca-171">Подробное описание @param см. в [JSDoc](http://usejsdoc.org/tags-param.html)</span><span class="sxs-lookup"><span data-stu-id="374ca-171">For detailed description of the  see: JSDoc</span></span>

---
### <a name="requiresaddress"></a><span data-ttu-id="374ca-172">@requiresAddress</span><span class="sxs-lookup"><span data-stu-id="374ca-172">@requiresAddress</span></span>
<a id="requiresAddress"/>

<span data-ttu-id="374ca-173">Указывает, что следует предоставлять адрес ячейки, в которой вычисляется функция.</span><span class="sxs-lookup"><span data-stu-id="374ca-173">Indicates that the address of the cell where the function is being evaluated should be provided.</span></span> 

<span data-ttu-id="374ca-174">Тип последнего параметра функции должен быть `CustomFunctions.Invocation` или производной от него.</span><span class="sxs-lookup"><span data-stu-id="374ca-174">The last function parameter must be of type `CustomFunctions.Invocation` or a derived type.</span></span> <span data-ttu-id="374ca-175">При вызове функции свойство `address` будет содержать адрес.</span><span class="sxs-lookup"><span data-stu-id="374ca-175">When the function is called, the `address` property will contain the address.</span></span>

---
### <a name="returns"></a><span data-ttu-id="374ca-176">@returns</span><span class="sxs-lookup"><span data-stu-id="374ca-176">Returns:</span></span>
<a id="returns"/>

<span data-ttu-id="374ca-177">Синтаксис: @returns {_type_}</span><span class="sxs-lookup"><span data-stu-id="374ca-177">Syntax: @returns {_type_}</span></span>

<span data-ttu-id="374ca-178">Предоставляет тип для возвращаемого значения.</span><span class="sxs-lookup"><span data-stu-id="374ca-178">Provides the type for the return value.</span></span>

<span data-ttu-id="374ca-179">Если `{type}` не указан, будет использоваться информация о типе TypeScript.</span><span class="sxs-lookup"><span data-stu-id="374ca-179">If `{type}` is omitted, the TypeScript type info will be used.</span></span> <span data-ttu-id="374ca-180">Если информация о типе отсутствует, будет использоваться тип `any`.</span><span class="sxs-lookup"><span data-stu-id="374ca-180">If there is no type info, the type will be `any`.</span></span>

---
### <a name="streaming"></a><span data-ttu-id="374ca-181">@streaming</span><span class="sxs-lookup"><span data-stu-id="374ca-181">streaming</span></span>
<a id="streaming"/>

<span data-ttu-id="374ca-182">Используется для обозначения того, что пользовательская функция является потоковой передачей функции.</span><span class="sxs-lookup"><span data-stu-id="374ca-182">Used to indicate that a custom function is a streaming function.</span></span> 

<span data-ttu-id="374ca-183">Тип последнего параметра должен быть `CustomFunctions.StreamingInvocation<ResultType>`.</span><span class="sxs-lookup"><span data-stu-id="374ca-183">The last parameter should be of type `CustomFunctions.StreamingInvocation<ResultType>`.</span></span>
<span data-ttu-id="374ca-184">Функция должна вернуть значение `void`.</span><span class="sxs-lookup"><span data-stu-id="374ca-184">The function should return `void`.</span></span>

<span data-ttu-id="374ca-185">Потоковые передачи функций непосредственно не возвращают значения, для этого необходимо вызывать `setResult(result: ResultType)` с помощью последнего параметра.</span><span class="sxs-lookup"><span data-stu-id="374ca-185">Streaming functions do not return values directly, but rather should call `setResult(result: ResultType)` using the last parameter.</span></span>

<span data-ttu-id="374ca-186">Исключения, которые возникают при потоковой передаче функций, игнорируются.</span><span class="sxs-lookup"><span data-stu-id="374ca-186">Exceptions thrown by a streaming function are ignored.</span></span> <span data-ttu-id="374ca-187">`setResult()` при вызове может вернуть ошибку в качестве результата.</span><span class="sxs-lookup"><span data-stu-id="374ca-187">`setResult()`may be called with Error to indicate an error result.</span></span>

<span data-ttu-id="374ca-188">Потоковые передачи функций невозможно пометить как [@volatile](#volatile).</span><span class="sxs-lookup"><span data-stu-id="374ca-188">Streaming functions cannot be marked as   .</span></span>

---
### <a name="volatile"></a><span data-ttu-id="374ca-189">@volatile</span><span class="sxs-lookup"><span data-stu-id="374ca-189">Volatile</span></span>
<a id="volatile"/>

<span data-ttu-id="374ca-190">Переменные функции — это такие функции, чей результат не остается неизменным в каждый период времени, даже если они не содержат аргументов или их аргументы не меняются.</span><span class="sxs-lookup"><span data-stu-id="374ca-190">A volatile function is one whose result cannot be assumed to be the same from one moment to the next even if it takes no arguments or the arguments have not changed.</span></span> <span data-ttu-id="374ca-191">Excel повторно проводит вычисления в ячейках, которые содержат переменные функции, вместе со всеми зависимыми функциями при каждом вычислении.</span><span class="sxs-lookup"><span data-stu-id="374ca-191">Excel re-evaluates cells that contain volatile functions, together with all dependents, every time that a calculation is done.</span></span> <span data-ttu-id="374ca-192">По этой причине чрезмерное использование переменных функций может замедлить пересчет, поэтому используйте их умеренно.</span><span class="sxs-lookup"><span data-stu-id="374ca-192">For this reason, too much reliance on volatile functions can make recalculation times slow, so use them sparingly.</span></span>

<span data-ttu-id="374ca-193">Потоковые передачи функций не могут быть переменными.</span><span class="sxs-lookup"><span data-stu-id="374ca-193">Streaming functions cannot be volatile.</span></span>

---

## <a name="types"></a><span data-ttu-id="374ca-194">Типы</span><span class="sxs-lookup"><span data-stu-id="374ca-194">Types</span></span>

<span data-ttu-id="374ca-195">Указывая тип параметра, Excel преобразует значения в этот тип, прежде чем вызывать функцию.</span><span class="sxs-lookup"><span data-stu-id="374ca-195">By specifying a parameter type, Excel will convert values into that type before calling the function.</span></span> <span data-ttu-id="374ca-196">Если указан тип `any`, преобразование выполняться не будет.</span><span class="sxs-lookup"><span data-stu-id="374ca-196">If the type is `any`, no conversion will be performed.</span></span>

### <a name="value-types"></a><span data-ttu-id="374ca-197">Типы значений</span><span class="sxs-lookup"><span data-stu-id="374ca-197">Value types</span></span>

<span data-ttu-id="374ca-198">Одно значение может быть представлено с помощью одного из приведенных ниже типов: `boolean`, `number`, `string`.</span><span class="sxs-lookup"><span data-stu-id="374ca-198">A single value may be represented using one of the following types: `boolean`, `number`, `string`.</span></span>

### <a name="matrix-type"></a><span data-ttu-id="374ca-199">Тип "матрица"</span><span class="sxs-lookup"><span data-stu-id="374ca-199">Matrix type</span></span>

<span data-ttu-id="374ca-200">Используйте тип двумерного массива, чтобы параметр или возвращаемое значение представляли собой матрицу значений.</span><span class="sxs-lookup"><span data-stu-id="374ca-200">Use a two-dimensional array type to have the parameter or return value be a matrix of values.</span></span> <span data-ttu-id="374ca-201">Например, тип `number[][]` указывает на матрицу чисел.</span><span class="sxs-lookup"><span data-stu-id="374ca-201">For example, the type `number[][]` indicates a matrix of numbers.</span></span> <span data-ttu-id="374ca-202">`string[][]` указывает на матрицу строк.</span><span class="sxs-lookup"><span data-stu-id="374ca-202">`string[][]`indicates a matrix of strings.</span></span> 

### <a name="error-type"></a><span data-ttu-id="374ca-203">Тип "ошибка"</span><span class="sxs-lookup"><span data-stu-id="374ca-203">Error type</span></span>

<span data-ttu-id="374ca-204">Функция непотоковой передачи может указывать на ошибку, возвращая тип "Ошибка".</span><span class="sxs-lookup"><span data-stu-id="374ca-204">A non-streaming function can indicate an error by returning an Error type.</span></span>

<span data-ttu-id="374ca-205">Функция потоковой передачи может указывать на ошибку, вызывая метод setResult() типа "Ошибка".</span><span class="sxs-lookup"><span data-stu-id="374ca-205">A streaming function can indicate an error by calling setResult() with an Error type.</span></span>

### <a name="promise"></a><span data-ttu-id="374ca-206">Обещание</span><span class="sxs-lookup"><span data-stu-id="374ca-206">Promise</span></span>

<span data-ttu-id="374ca-207">Функция может вернуть тип "Обещание", который задаст значение, когда обещание будет разрешено.</span><span class="sxs-lookup"><span data-stu-id="374ca-207">A function can return a Promise, which will provide the value when the promise is resolved.</span></span> <span data-ttu-id="374ca-208">В случае отклонения обещания возникнет ошибка.</span><span class="sxs-lookup"><span data-stu-id="374ca-208">If the promise is rejected, then it is an error.</span></span>

### <a name="other-types"></a><span data-ttu-id="374ca-209">Другие типы</span><span class="sxs-lookup"><span data-stu-id="374ca-209">Other types</span></span>

<span data-ttu-id="374ca-210">Любой другой тип будет рассматриваться как ошибка.</span><span class="sxs-lookup"><span data-stu-id="374ca-210">Any other type will be treated as an error.</span></span>

## <a name="see-also"></a><span data-ttu-id="374ca-211">См. также</span><span class="sxs-lookup"><span data-stu-id="374ca-211">See also</span></span>

* [<span data-ttu-id="374ca-212">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="374ca-212">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="374ca-213">Среда выполнения для пользовательских функций Excel</span><span class="sxs-lookup"><span data-stu-id="374ca-213">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="374ca-214">Рекомендации по пользовательским функциям</span><span class="sxs-lookup"><span data-stu-id="374ca-214">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="374ca-215">Журнал изменений пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="374ca-215">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="374ca-216">Руководство по пользовательским функциям в Excel</span><span class="sxs-lookup"><span data-stu-id="374ca-216">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="374ca-217">Отладка пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="374ca-217">Custom functions debugging</span></span>](custom-functions-debugging.md)
