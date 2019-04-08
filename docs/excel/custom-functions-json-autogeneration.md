---
ms.date: 04/03/2019
description: Использование тегов JSDOC для динамического создания метаданных JSON пользовательских функций.
title: Создание метаданных JSON для пользовательских функций (предварительная версия)
localization_priority: Priority
ms.openlocfilehash: c6d89684da2d0773ccfb1763e5e3e426e647523b
ms.sourcegitcommit: 14ceac067e0e130869b861d289edb438b5e3eff9
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/04/2019
ms.locfileid: "31478967"
---
# <a name="create-json-metadata-for-custom-functions-preview"></a><span data-ttu-id="949e1-103">Создание метаданных JSON для пользовательских функций (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="949e1-103">Create JSON metadata for custom functions (preview)</span></span>

<span data-ttu-id="949e1-104">Если пользовательская функция Excel написана в JavaScript или TypeScript, теги JSDoc используются для предоставления дополнительной информации о пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="949e1-104">When an Excel custom function is written in JavaScript or TypeScript, JSDoc tags are used to provide extra information about the custom function.</span></span> <span data-ttu-id="949e1-105">Теги JSDoc используются при сборке для создания [файла метаданных JSON](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="949e1-105">The JSDoc tags are then used at build time to create the [JSON metadata file](custom-functions-json.md).</span></span> <span data-ttu-id="949e1-106">Использование тегов JSDoc освобождает вас от необходимости редактировать файл метаданных JSON вручную.</span><span class="sxs-lookup"><span data-stu-id="949e1-106">Using JSDoc tags saves you from the effort of manually editing the JSON metadata file.</span></span>

<span data-ttu-id="949e1-107">Добавьте тег `@customfunction` в примечаниях к коду для функции JavaScript или TypeScript, чтобы пометить ее как пользовательскую.</span><span class="sxs-lookup"><span data-stu-id="949e1-107">Add the `@customfunction` tag in the code comments for a JavaScript or TypeScript function to mark it as a custom function.</span></span>

<span data-ttu-id="949e1-108">Типы параметров функции можно получить с помощью тега [@param](#param) в JavaScript или из раздела [Типа функции](http://www.typescriptlang.org/docs/handbook/functions.html) в TypeScript.</span><span class="sxs-lookup"><span data-stu-id="949e1-108">The function parameter types may be provided using the [@param](#param) tag in JavaScript, or from the [Function type](http://www.typescriptlang.org/docs/handbook/functions.html) in TypeScript.</span></span> <span data-ttu-id="949e1-109">Дополнительную информацию см. в теге [@param](#param) и разделе [Типы](#Types).</span><span class="sxs-lookup"><span data-stu-id="949e1-109">For more information, see the [@param](#param) tag and [Types](#Types) section.</span></span>

## <a name="jsdoc-tags"></a><span data-ttu-id="949e1-110">Теги JSDoc</span><span class="sxs-lookup"><span data-stu-id="949e1-110">JSDoc Tags</span></span>
<span data-ttu-id="949e1-111">Ниже приведены теги JSDoc, которые поддерживаются в пользовательских функциях Excel:</span><span class="sxs-lookup"><span data-stu-id="949e1-111">The following JSDoc tags are supported in Excel custom functions:</span></span>
* [@cancelable](#cancelable)
* <span data-ttu-id="949e1-112">[@customfunction](#customfunction) id name</span><span class="sxs-lookup"><span data-stu-id="949e1-112">[@customfunction](#customfunction) id name</span></span>
* <span data-ttu-id="949e1-113">[@helpurl](#helpurl) url</span><span class="sxs-lookup"><span data-stu-id="949e1-113">URL</span></span>
* <span data-ttu-id="949e1-114">[@param](#param) _{type}_ name description</span><span class="sxs-lookup"><span data-stu-id="949e1-114">[@param](#param) _{type}_ name description</span></span>
* [@requiresAddress](#requiresAddress)
* <span data-ttu-id="949e1-115">[@returns](#returns) _{type}_</span><span class="sxs-lookup"><span data-stu-id="949e1-115">Type</span></span>
* [@streaming](#streaming)
* [@volatile](#volatile)

---
### <a name="cancelable"></a>@cancelable
<a id="cancelable"/>

<span data-ttu-id="949e1-116">При отмене функции указывает, что пользовательская функция стремится к выполнению действия.</span><span class="sxs-lookup"><span data-stu-id="949e1-116">Indicates that a custom function wants to perform an action when the function is canceled.</span></span>

<span data-ttu-id="949e1-117">В качестве типа последнего параметра функции должно быть указано `CustomFunctions.CancelableInvocation`.</span><span class="sxs-lookup"><span data-stu-id="949e1-117">The last function parameter must be of type `CustomFunctions.CancelableInvocation`.</span></span> <span data-ttu-id="949e1-118">Функция может назначить функцию свойству `oncanceled`, чтобы обозначить действия для выполнения в случае отмены функции.</span><span class="sxs-lookup"><span data-stu-id="949e1-118">The function can assign a function to the `oncanceled` property to denote the action to perform when the function is canceled.</span></span>

<span data-ttu-id="949e1-119">Если тип последнего параметра функции `CustomFunctions.CancelableInvocation`, он будет рассматриваться как `@cancelable`, даже если тег отсутствует.</span><span class="sxs-lookup"><span data-stu-id="949e1-119">If the last function parameter is of type `CustomFunctions.CancelableInvocation`, it will be considered `@cancelable` even if the tag is not present.</span></span>

<span data-ttu-id="949e1-120">Функция не может содержать одновременно теги `@cancelable` и `@streaming`.</span><span class="sxs-lookup"><span data-stu-id="949e1-120">A function cannot have both `@cancelable` and `@streaming` tags.</span></span>

---
### <a name="customfunction"></a>@customfunction
<a id="customfunction"/>

<span data-ttu-id="949e1-121">Синтаксис: @customfunction _id_ _name_</span><span class="sxs-lookup"><span data-stu-id="949e1-121">Syntax: @customfunction _id_ _name_</span></span>

<span data-ttu-id="949e1-122">Укажите этот тег, чтобы рассматривать функцию JavaScript или TypeScript как пользовательскую функцию Excel.</span><span class="sxs-lookup"><span data-stu-id="949e1-122">Specify this tag to treat the JavaScript/TypeScript function as an Excel custom function.</span></span>

<span data-ttu-id="949e1-123">Этот тег необходим, чтобы создать метаданные для пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="949e1-123">This tag is required to create metadata for the custom function.</span></span>

<span data-ttu-id="949e1-124">Кроме того, требуется вызов функции</span><span class="sxs-lookup"><span data-stu-id="949e1-124">There should also be a call to</span></span> `CustomFunctions.associate("id", functionName);`

#### <a name="id"></a><span data-ttu-id="949e1-125">id</span><span class="sxs-lookup"><span data-stu-id="949e1-125">id</span></span> 

<span data-ttu-id="949e1-126">Идентификатор используется как инвариантный идентификатор для пользовательских функций, которые хранятся в документе.</span><span class="sxs-lookup"><span data-stu-id="949e1-126">The id is used as the invariant identifier for the custom function stored in the document.</span></span> <span data-ttu-id="949e1-127">Его не следует менять.</span><span class="sxs-lookup"><span data-stu-id="949e1-127">It should not change.</span></span>

* <span data-ttu-id="949e1-128">Если идентификатор не указан, название функции JavaScript или TypeScript преобразуется в верхний регистр, а недопустимые символы удаляются.</span><span class="sxs-lookup"><span data-stu-id="949e1-128">If id is not provided, the JavaScript/TypeScript function name is converted to uppercase, disallowed characters are removed.</span></span>
* <span data-ttu-id="949e1-129">Идентификатор должен быть уникальным для всех пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="949e1-129">The id must be unique for all custom functions.</span></span>
* <span data-ttu-id="949e1-130">Допустимые символы: A–Z, a–z, 0–9 и точки (.).</span><span class="sxs-lookup"><span data-stu-id="949e1-130">The characters allowed are limited to: A-Z, a-z, 0-9, and period (.).</span></span>

#### <a name="name"></a><span data-ttu-id="949e1-131">name</span><span class="sxs-lookup"><span data-stu-id="949e1-131">name</span></span>

<span data-ttu-id="949e1-132">Предоставляет отображаемое имя для пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="949e1-132">Provides the display name of a custom category for the property.</span></span> 

* <span data-ttu-id="949e1-133">Если имя не указано, идентификатор тоже используется как имя.</span><span class="sxs-lookup"><span data-stu-id="949e1-133">If name is not provided, the id is also used as the name.</span></span>
* <span data-ttu-id="949e1-134">Допустимые символы: буквы [буквенные символы Юникод](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), числа, точки (.) и подчеркивания (\_).</span><span class="sxs-lookup"><span data-stu-id="949e1-134">Allowed characters: Letters [Unicode Alphabetic character](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), numbers, period (.), and underscore (\_).</span></span>
* <span data-ttu-id="949e1-135">Имя должно начинаться с буквы.</span><span class="sxs-lookup"><span data-stu-id="949e1-135">Must start with a letter.</span></span>
* <span data-ttu-id="949e1-136">Максимальная длина: 128 символов.</span><span class="sxs-lookup"><span data-stu-id="949e1-136">Maximum length is 255 characters.</span></span>

---
### <a name="helpurl"></a>@helpurl
<a id="helpurl"/>

<span data-ttu-id="949e1-137">Синтаксис: @helpurl _url_</span><span class="sxs-lookup"><span data-stu-id="949e1-137">Syntax: @helpurl _url_</span></span>

<span data-ttu-id="949e1-138">Предоставленный _url_-адрес отображается в Excel.</span><span class="sxs-lookup"><span data-stu-id="949e1-138">The provided _url_ is displayed in Excel.</span></span>

---
### <a name="param"></a>@param
<a id="param"/>

#### <a name="javascript"></a><span data-ttu-id="949e1-139">JavaScript</span><span class="sxs-lookup"><span data-stu-id="949e1-139">JavaScript</span></span>

<span data-ttu-id="949e1-140">Синтаксис JavaScript: @param {type} name _description_</span><span class="sxs-lookup"><span data-stu-id="949e1-140">JavaScript Syntax: @param {type} name _description_</span></span>

* `{type}` <span data-ttu-id="949e1-141">следует указать информацию о типе в фигурных скобках.</span><span class="sxs-lookup"><span data-stu-id="949e1-141">should specify the type info within curly braces.</span></span> <span data-ttu-id="949e1-142">Дополнительную информацию о типах, которые могут использоваться, см. в разделе [Типы](##types).</span><span class="sxs-lookup"><span data-stu-id="949e1-142">See the [Types](##types) for more information about the types which may be used.</span></span> <span data-ttu-id="949e1-143">Необязательно: если тип не указан, будет использоваться тип `any`.</span><span class="sxs-lookup"><span data-stu-id="949e1-143">Optional: if not specified, the type `any` will be used.</span></span>
* `name` <span data-ttu-id="949e1-144">указывает, к какому параметру относится тег @param.</span><span class="sxs-lookup"><span data-stu-id="949e1-144">specifies which parameter the @param tag applies to.</span></span> <span data-ttu-id="949e1-145">Обязательно.</span><span class="sxs-lookup"><span data-stu-id="949e1-145">Required.</span></span>
* `description` <span data-ttu-id="949e1-146">предоставляет описание, которое отображается в Excel для параметра функции.</span><span class="sxs-lookup"><span data-stu-id="949e1-146">provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="949e1-147">Необязательно.</span><span class="sxs-lookup"><span data-stu-id="949e1-147">Optional.</span></span>

<span data-ttu-id="949e1-148">Чтобы обозначить параметр пользовательской функции как необязательный:</span><span class="sxs-lookup"><span data-stu-id="949e1-148">To denote a custom function parameter as optional:</span></span>
* <span data-ttu-id="949e1-149">Поместите имя параметра в квадратные скобки.</span><span class="sxs-lookup"><span data-stu-id="949e1-149">Put square brackets around the parameter name.</span></span> <span data-ttu-id="949e1-150">Пример: `@param {string} [text] Optional text`.</span><span class="sxs-lookup"><span data-stu-id="949e1-150">For example: `@param {string} [text] Optional text`.</span></span>

#### <a name="typescript"></a><span data-ttu-id="949e1-151">TypeScript</span><span class="sxs-lookup"><span data-stu-id="949e1-151">TypeScript</span></span>

<span data-ttu-id="949e1-152">Синтаксис TypeScript: @param name _description_</span><span class="sxs-lookup"><span data-stu-id="949e1-152">TypeScript Syntax: @param name _description_</span></span>

* `name` <span data-ttu-id="949e1-153">указывает, к какому параметру относится тег @param.</span><span class="sxs-lookup"><span data-stu-id="949e1-153">specifies which parameter the @param tag applies to.</span></span> <span data-ttu-id="949e1-154">Обязательно.</span><span class="sxs-lookup"><span data-stu-id="949e1-154">Required.</span></span>
* `description` <span data-ttu-id="949e1-155">предоставляет описание, которое отображается в Excel для параметра функции.</span><span class="sxs-lookup"><span data-stu-id="949e1-155">provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="949e1-156">Необязательно.</span><span class="sxs-lookup"><span data-stu-id="949e1-156">Optional.</span></span>

<span data-ttu-id="949e1-157">Дополнительную информацию о типах параметров функций, которые могут использоваться, см. в разделе [Типы](##types).</span><span class="sxs-lookup"><span data-stu-id="949e1-157">See the [Types](##types) for more information about the function parameter types which may be used.</span></span>

<span data-ttu-id="949e1-158">Чтобы обозначить параметр пользовательской функции как необязательный, выполните одно из указанных ниже действий.</span><span class="sxs-lookup"><span data-stu-id="949e1-158">To denote a custom function parameter as optional, do one of the following:</span></span>
* <span data-ttu-id="949e1-159">Используйте необязательный параметр.</span><span class="sxs-lookup"><span data-stu-id="949e1-159">Use an optional parameter.</span></span> <span data-ttu-id="949e1-160">Пример:</span><span class="sxs-lookup"><span data-stu-id="949e1-160">For example:</span></span> `function f(text?: string)`
* <span data-ttu-id="949e1-161">Задайте для параметра значение по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="949e1-161">Give the parameter a default value.</span></span> <span data-ttu-id="949e1-162">Пример:</span><span class="sxs-lookup"><span data-stu-id="949e1-162">For example:</span></span> `function f(text: string = "abc")`

<span data-ttu-id="949e1-163">Подробное описание @param см. в [JSDoc](http://usejsdoc.org/tags-param.html)</span><span class="sxs-lookup"><span data-stu-id="949e1-163">For detailed description of the @param see: [JSDoc](http://usejsdoc.org/tags-param.html)</span></span>

---
### <a name="requiresaddress"></a>@requiresAddress
<a id="requiresAddress"/>

<span data-ttu-id="949e1-164">Указывает, что следует предоставлять адрес ячейки, в которой вычисляется функция.</span><span class="sxs-lookup"><span data-stu-id="949e1-164">Indicates that the address of the cell where the function is being evaluated should be provided.</span></span> 

<span data-ttu-id="949e1-165">Тип последнего параметра функции должен быть `CustomFunctions.Invocation` или производной от него.</span><span class="sxs-lookup"><span data-stu-id="949e1-165">The last function parameter must be of type `CustomFunctions.Invocation` or a derived type.</span></span> <span data-ttu-id="949e1-166">При вызове функции свойство `address` будет содержать адрес.</span><span class="sxs-lookup"><span data-stu-id="949e1-166">When the function is called, the `address` property will contain the address.</span></span>

---
### <a name="returns"></a>@returns
<a id="returns"/>

<span data-ttu-id="949e1-167">Синтаксис: @returns {_type_}</span><span class="sxs-lookup"><span data-stu-id="949e1-167">Syntax: @returns {_type_}</span></span>

<span data-ttu-id="949e1-168">Предоставляет тип для возвращаемого значения.</span><span class="sxs-lookup"><span data-stu-id="949e1-168">Provides the type for the return value.</span></span>

<span data-ttu-id="949e1-169">Если `{type}` не указан, будет использоваться информация о типе TypeScript.</span><span class="sxs-lookup"><span data-stu-id="949e1-169">If `{type}` is omitted, the TypeScript type info will be used.</span></span> <span data-ttu-id="949e1-170">Если информация о типе отсутствует, будет использоваться тип `any`.</span><span class="sxs-lookup"><span data-stu-id="949e1-170">If there is no type info, the type will be `any`.</span></span>

---
### <a name="streaming"></a>@streaming
<a id="streaming"/>

<span data-ttu-id="949e1-171">Используется для обозначения того, что пользовательская функция является потоковой передачей функции.</span><span class="sxs-lookup"><span data-stu-id="949e1-171">Used to indicate that a custom function is a streaming function.</span></span> 

<span data-ttu-id="949e1-172">Тип последнего параметра должен быть `CustomFunctions.StreamingInvocation<ResultType>`.</span><span class="sxs-lookup"><span data-stu-id="949e1-172">The last parameter should be of type `CustomFunctions.StreamingInvocation<ResultType>`.</span></span>
<span data-ttu-id="949e1-173">Функция должна вернуть значение `void`.</span><span class="sxs-lookup"><span data-stu-id="949e1-173">The function should return `void`.</span></span>

<span data-ttu-id="949e1-174">Потоковые передачи функций непосредственно не возвращают значения, для этого необходимо вызывать `setResult(result: ResultType)` с помощью последнего параметра.</span><span class="sxs-lookup"><span data-stu-id="949e1-174">Streaming functions do not return values directly, but rather should call `setResult(result: ResultType)` using the last parameter.</span></span>

<span data-ttu-id="949e1-175">Исключения, которые возникают при потоковой передаче функций, игнорируются.</span><span class="sxs-lookup"><span data-stu-id="949e1-175">Exceptions thrown by a streaming function are ignored.</span></span> `setResult()` <span data-ttu-id="949e1-176">при вызове может вернуть ошибку в качестве результата.</span><span class="sxs-lookup"><span data-stu-id="949e1-176">may be called with Error to indicate an error result.</span></span>

<span data-ttu-id="949e1-177">Потоковые передачи функций невозможно пометить как [@volatile](#volatile).</span><span class="sxs-lookup"><span data-stu-id="949e1-177">Streaming functions cannot be marked as [@volatile](#volatile).</span></span>

---
### <a name="volatile"></a>@volatile
<a id="volatile"/>

<span data-ttu-id="949e1-178">Переменные функции — это такие функции, чей результат не остается неизменным в каждый период времени, даже если они не содержат аргументов или их аргументы не меняются.</span><span class="sxs-lookup"><span data-stu-id="949e1-178">A volatile function is one whose result cannot be assumed to be the same from one moment to the next even if it takes no arguments or the arguments have not changed.</span></span> <span data-ttu-id="949e1-179">Excel повторно проводит вычисления в ячейках, которые содержат переменные функции, вместе со всеми зависимыми функциями при каждом вычислении.</span><span class="sxs-lookup"><span data-stu-id="949e1-179">Excel reevaluates cells that contain volatile functions, together with all dependents, every time that it recalculates.</span></span> <span data-ttu-id="949e1-180">По этой причине чрезмерное использование переменных функций может замедлить пересчет, поэтому используйте их умеренно.</span><span class="sxs-lookup"><span data-stu-id="949e1-180">For this reason, too much reliance on volatile functions can make recalculation times slow.</span></span>

<span data-ttu-id="949e1-181">Потоковые передачи функций не могут быть переменными.</span><span class="sxs-lookup"><span data-stu-id="949e1-181">Streaming functions cannot be volatile.</span></span>

---

## <a name="types"></a><span data-ttu-id="949e1-182">Типы</span><span class="sxs-lookup"><span data-stu-id="949e1-182">Types</span></span>

<span data-ttu-id="949e1-183">Указывая тип параметра, Excel преобразует значения в этот тип, прежде чем вызывать функцию.</span><span class="sxs-lookup"><span data-stu-id="949e1-183">By specifying a parameter type, Excel will convert values into that type before calling the function.</span></span> <span data-ttu-id="949e1-184">Если указан тип `any`, преобразование выполняться не будет.</span><span class="sxs-lookup"><span data-stu-id="949e1-184">If the type is `any`, no conversion will be performed.</span></span>

### <a name="value-types"></a><span data-ttu-id="949e1-185">Типы значений</span><span class="sxs-lookup"><span data-stu-id="949e1-185">Value types</span></span>

<span data-ttu-id="949e1-186">Одно значение может быть представлено с помощью одного из приведенных ниже типов: `boolean`, `number`, `string`.</span><span class="sxs-lookup"><span data-stu-id="949e1-186">A single value may be represented using one of the following types: `boolean`, `number`, `string`.</span></span>

### <a name="matrix-type"></a><span data-ttu-id="949e1-187">Тип "матрица"</span><span class="sxs-lookup"><span data-stu-id="949e1-187">Matrix type</span></span>

<span data-ttu-id="949e1-188">Используйте тип двумерного массива, чтобы параметр или возвращаемое значение представляли собой матрицу значений.</span><span class="sxs-lookup"><span data-stu-id="949e1-188">Use a two-dimensional array type to have the parameter or return value be a matrix of values.</span></span> <span data-ttu-id="949e1-189">Например, тип `number[][]` указывает на матрицу чисел.</span><span class="sxs-lookup"><span data-stu-id="949e1-189">For example, the type `number[][]` indicates a matrix of numbers.</span></span> `string[][]` <span data-ttu-id="949e1-190">указывает на матрицу строк.</span><span class="sxs-lookup"><span data-stu-id="949e1-190">indicates a matrix of strings.</span></span> 

### <a name="error-type"></a><span data-ttu-id="949e1-191">Тип "ошибка"</span><span class="sxs-lookup"><span data-stu-id="949e1-191">Error Type</span></span>

<span data-ttu-id="949e1-192">Функция непотоковой передачи может указывать на ошибку, возвращая тип "Ошибка".</span><span class="sxs-lookup"><span data-stu-id="949e1-192">A non-streaming function can indicate an error by returning an Error type.</span></span>

<span data-ttu-id="949e1-193">Функция потоковой передачи может указывать на ошибку, вызывая метод setResult() типа "Ошибка".</span><span class="sxs-lookup"><span data-stu-id="949e1-193">A streaming function can indicate an error by calling setResult() with an Error type.</span></span>

### <a name="promise"></a><span data-ttu-id="949e1-194">Обещание</span><span class="sxs-lookup"><span data-stu-id="949e1-194">Promise object.</span></span>

<span data-ttu-id="949e1-195">Функция может вернуть тип "Обещание", который задаст значение, когда обещание будет разрешено.</span><span class="sxs-lookup"><span data-stu-id="949e1-195">A function can return a Promise, which will provide the value when the promise is resolved.</span></span> <span data-ttu-id="949e1-196">В случае отклонения обещания возникнет ошибка.</span><span class="sxs-lookup"><span data-stu-id="949e1-196">If the promise is rejected, then it is an error.</span></span>

### <a name="other-types"></a><span data-ttu-id="949e1-197">Другие типы</span><span class="sxs-lookup"><span data-stu-id="949e1-197">Other solution types</span></span>

<span data-ttu-id="949e1-198">Любой другой тип будет рассматриваться как ошибка.</span><span class="sxs-lookup"><span data-stu-id="949e1-198">Any other type will be treated as an error.</span></span>

## <a name="see-also"></a><span data-ttu-id="949e1-199">См. также</span><span class="sxs-lookup"><span data-stu-id="949e1-199">See also</span></span>

* [<span data-ttu-id="949e1-200">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="949e1-200">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="949e1-201">Среда выполнения для пользовательских функций Excel</span><span class="sxs-lookup"><span data-stu-id="949e1-201">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="949e1-202">Рекомендации в отношении пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="949e1-202">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="949e1-203">Журнал изменений пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="949e1-203">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="949e1-204">Руководство по пользовательским функциям в Excel</span><span class="sxs-lookup"><span data-stu-id="949e1-204">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
* [<span data-ttu-id="949e1-205">Отладка пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="949e1-205">Custom functions debugging</span></span>](custom-functions-debugging.md)
