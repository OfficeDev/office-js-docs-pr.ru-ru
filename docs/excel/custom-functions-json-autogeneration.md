---
ms.date: 06/10/2019
description: Использование тегов JSDoc для динамического создания метаданных JSON пользовательских функций.
title: Автоматическое генерирование метаданных JSON для пользовательских функций
localization_priority: Priority
ms.openlocfilehash: 960e1eca1e01aec21967733d802a5fdd48122cbc
ms.sourcegitcommit: 3f84b2caa73d7fe1eb0d15e32ea4dec459e2ff53
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/12/2019
ms.locfileid: "34910303"
---
# <a name="autogenerate-json-metadata-for-custom-functions"></a><span data-ttu-id="beb2c-103">Автоматическое генерирование метаданных JSON для пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="beb2c-103">Autogenerate JSON metadata for custom functions</span></span>

<span data-ttu-id="beb2c-104">Если пользовательская функция Excel написана в JavaScript или TypeScript, теги JSDoc используются для предоставления дополнительной информации о пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="beb2c-104">When an Excel custom function is written in JavaScript or TypeScript, JSDoc tags are used to provide extra information about the custom function.</span></span> <span data-ttu-id="beb2c-105">Теги JSDoc используются при сборке для создания [файла метаданных JSON](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="beb2c-105">The JSDoc tags are then used at build time to create the [JSON metadata file](custom-functions-json.md).</span></span> <span data-ttu-id="beb2c-106">Использование тегов JSDoc освобождает вас от необходимости редактировать файл метаданных JSON вручную.</span><span class="sxs-lookup"><span data-stu-id="beb2c-106">Using JSDoc tags saves you from the effort of manually editing the JSON metadata file.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="beb2c-107">Добавьте тег `@customfunction` в примечаниях к коду для функции JavaScript или TypeScript, чтобы пометить ее как пользовательскую.</span><span class="sxs-lookup"><span data-stu-id="beb2c-107">Add the `@customfunction` tag in the code comments for a JavaScript or TypeScript function to mark it as a custom function.</span></span>

<span data-ttu-id="beb2c-108">Типы параметров функции можно получить с помощью тега [@param](#param) в JavaScript или из раздела [Тип функции](https://www.typescriptlang.org/docs/handbook/functions.html) в TypeScript.</span><span class="sxs-lookup"><span data-stu-id="beb2c-108">The function parameter types may be provided using the [@param](#param) tag in JavaScript, or from the [Function type](https://www.typescriptlang.org/docs/handbook/functions.html) in TypeScript.</span></span> <span data-ttu-id="beb2c-109">Дополнительную информацию см. в теге [@param](#param) и разделе [Типы](#types).</span><span class="sxs-lookup"><span data-stu-id="beb2c-109">For more information, see the [@param](#param) tag and [Types](#types) section.</span></span>

### <a name="adding-a-description-to-a-function"></a><span data-ttu-id="beb2c-110">Добавление описания функции</span><span class="sxs-lookup"><span data-stu-id="beb2c-110">Adding a description to a function</span></span>

<span data-ttu-id="beb2c-111">Описание отображается пользователю в качестве текста справки, если ему непонятно действие пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="beb2c-111">The description is displayed to the user as help text when they need help to understand what your custom function does.</span></span> <span data-ttu-id="beb2c-112">Описанию не требуется какой-либо конкретный тег.</span><span class="sxs-lookup"><span data-stu-id="beb2c-112">The description doesn't require any specific tag.</span></span> <span data-ttu-id="beb2c-113">Просто введите краткий текст описания в комментарии JSDoc.</span><span class="sxs-lookup"><span data-stu-id="beb2c-113">Just enter a short text description in the JSDoc comment.</span></span> <span data-ttu-id="beb2c-114">Обычно описание размещается в начале раздела комментариев JSDoc, но оно поддерживается независимо от места размещения.</span><span class="sxs-lookup"><span data-stu-id="beb2c-114">In general the description is placed at the start of the JSDoc comment section, but it will work no matter where it is placed.</span></span>

<span data-ttu-id="beb2c-115">Чтобы просмотреть примеры описаний встроенных функций, откройте Excel, перейдите на вкладку **Формулы** и нажмите кнопку **Вставить функцию**.</span><span class="sxs-lookup"><span data-stu-id="beb2c-115">To see examples of the built-in function descriptions, open Excel, go to the **Formulas** tab, and choose **Insert function**.</span></span> <span data-ttu-id="beb2c-116">Вы сможете просмотреть все описания функций, а также список собственных пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="beb2c-116">You can then browse through all the function descriptions, and also see your own custom functions listed.</span></span>

<span data-ttu-id="beb2c-117">В следующем примере фраза "Calculates the volume of a sphere." (Вычисляет объем сферы)</span><span class="sxs-lookup"><span data-stu-id="beb2c-117">In the following example, the phrase "Calculates the volume of a sphere."</span></span> <span data-ttu-id="beb2c-118">является описанием пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="beb2c-118">is the description for the custom function.</span></span>

```JS
/**
/* Calculates the volume of a sphere
/* @customfunction VOLUME
...
 */
```


## <a name="jsdoc-tags"></a><span data-ttu-id="beb2c-119">Теги JSDoc</span><span class="sxs-lookup"><span data-stu-id="beb2c-119">JSDoc Tags</span></span>
<span data-ttu-id="beb2c-120">Ниже приведены теги JSDoc, которые поддерживаются в пользовательских функциях Excel:</span><span class="sxs-lookup"><span data-stu-id="beb2c-120">The following JSDoc tags are supported in Excel custom functions:</span></span>
* [<span data-ttu-id="beb2c-121">@cancelable</span><span class="sxs-lookup"><span data-stu-id="beb2c-121">@cancelable</span></span>](#cancelable)
* <span data-ttu-id="beb2c-122">[@customfunction](#customfunction) id name</span><span class="sxs-lookup"><span data-stu-id="beb2c-122">[@customfunction](#customfunction) id name</span></span>
* <span data-ttu-id="beb2c-123">[@helpurl](#helpurl) url</span><span class="sxs-lookup"><span data-stu-id="beb2c-123">[@helpurl](#helpurl) url</span></span>
* <span data-ttu-id="beb2c-124">[@param](#param) _{type}_ name description</span><span class="sxs-lookup"><span data-stu-id="beb2c-124">[@param](#param) _{type}_ name description</span></span>
* [<span data-ttu-id="beb2c-125">@requiresAddress</span><span class="sxs-lookup"><span data-stu-id="beb2c-125">@requiresAddress</span></span>](#requiresAddress)
* <span data-ttu-id="beb2c-126">[@returns](#returns) _{type}_</span><span class="sxs-lookup"><span data-stu-id="beb2c-126">[@returns](#returns) _{type}_</span></span>
* [<span data-ttu-id="beb2c-127">@streaming</span><span class="sxs-lookup"><span data-stu-id="beb2c-127">@streaming</span></span>](#streaming)
* [<span data-ttu-id="beb2c-128">@volatile</span><span class="sxs-lookup"><span data-stu-id="beb2c-128">@volatile</span></span>](#volatile)

---
### <a name="cancelable"></a><span data-ttu-id="beb2c-129">@cancelable</span><span class="sxs-lookup"><span data-stu-id="beb2c-129">@cancelable</span></span>
<a id="cancelable"/>

<span data-ttu-id="beb2c-130">При отмене функции указывает, что пользовательская функция стремится к выполнению действия.</span><span class="sxs-lookup"><span data-stu-id="beb2c-130">Indicates that a custom function wants to perform an action when the function is canceled.</span></span>

<span data-ttu-id="beb2c-131">В качестве типа последнего параметра функции должно быть указано `CustomFunctions.CancelableInvocation`.</span><span class="sxs-lookup"><span data-stu-id="beb2c-131">The last function parameter must be of type `CustomFunctions.CancelableInvocation`.</span></span> <span data-ttu-id="beb2c-132">Функция может назначить функцию свойству `oncanceled`, чтобы обозначить действия для выполнения в случае отмены функции.</span><span class="sxs-lookup"><span data-stu-id="beb2c-132">The function can assign a function to the `oncanceled` property to denote the action to perform when the function is canceled.</span></span>

<span data-ttu-id="beb2c-133">Если тип последнего параметра функции `CustomFunctions.CancelableInvocation`, он будет рассматриваться как `@cancelable`, даже если тег отсутствует.</span><span class="sxs-lookup"><span data-stu-id="beb2c-133">If the last function parameter is of type `CustomFunctions.CancelableInvocation`, it will be considered `@cancelable` even if the tag is not present.</span></span>

<span data-ttu-id="beb2c-134">Функция не может содержать одновременно теги `@cancelable` и `@streaming`.</span><span class="sxs-lookup"><span data-stu-id="beb2c-134">A function cannot have both `@cancelable` and `@streaming` tags.</span></span>

---
### <a name="customfunction"></a><span data-ttu-id="beb2c-135">@customfunction</span><span class="sxs-lookup"><span data-stu-id="beb2c-135">@customfunction</span></span>
<a id="customfunction"/>

<span data-ttu-id="beb2c-136">Синтаксис: @customfunction _id_ _name_</span><span class="sxs-lookup"><span data-stu-id="beb2c-136">Syntax: @customfunction _id_ _name_</span></span>

<span data-ttu-id="beb2c-137">Укажите этот тег, чтобы рассматривать функцию JavaScript или TypeScript как пользовательскую функцию Excel.</span><span class="sxs-lookup"><span data-stu-id="beb2c-137">Specify this tag to treat the JavaScript/TypeScript function as an Excel custom function.</span></span>

<span data-ttu-id="beb2c-138">Этот тег необходим, чтобы создать метаданные для пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="beb2c-138">This tag is required to create metadata for the custom function.</span></span>

<span data-ttu-id="beb2c-139">Кроме того, требуется вызов функции `CustomFunctions.associate("id", functionName);`</span><span class="sxs-lookup"><span data-stu-id="beb2c-139">There should also be a call to `CustomFunctions.associate("id", functionName);`</span></span>

#### <a name="id"></a><span data-ttu-id="beb2c-140">id</span><span class="sxs-lookup"><span data-stu-id="beb2c-140">id</span></span>

<span data-ttu-id="beb2c-141">`id` является инвариантным идентификатором для пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="beb2c-141">The id is used as the invariant identifier for the custom function stored in the document.</span></span>

* <span data-ttu-id="beb2c-142">Если `id` не указан, название функции JavaScript или TypeScript преобразуется в верхний регистр, а недопустимые символы удаляются.</span><span class="sxs-lookup"><span data-stu-id="beb2c-142">If id is not provided, the JavaScript/TypeScript function name is converted to uppercase, disallowed characters are removed.</span></span>
* <span data-ttu-id="beb2c-143">`id` должен быть уникальным для всех пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="beb2c-143">The id must be unique for all custom functions.</span></span>
* <span data-ttu-id="beb2c-144">Допустимые символы: A — Z, a — z, 0–9, символ подчеркивания (\_) и точка (.).</span><span class="sxs-lookup"><span data-stu-id="beb2c-144">The characters allowed are limited to: A-Z, a-z, 0-9, and period (.).</span></span>

#### <a name="name"></a><span data-ttu-id="beb2c-145">name</span><span class="sxs-lookup"><span data-stu-id="beb2c-145">name</span></span>

<span data-ttu-id="beb2c-146">Предоставляет отображаемый параметр `name` для пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="beb2c-146">Provides the display name for the custom function.</span></span>

* <span data-ttu-id="beb2c-147">Если имя не указано, идентификатор также используется как имя.</span><span class="sxs-lookup"><span data-stu-id="beb2c-147">If name is not provided, the id is also used as the name.</span></span>
* <span data-ttu-id="beb2c-148">Допустимые символы: буквы [буквенные символы Юникод](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), числа, точки (.) и подчеркивания (\_).</span><span class="sxs-lookup"><span data-stu-id="beb2c-148">Allowed characters: Letters [Unicode Alphabetic character](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), numbers, period (.), and underscore (\_).</span></span>
* <span data-ttu-id="beb2c-149">Имя должно начинаться с буквы.</span><span class="sxs-lookup"><span data-stu-id="beb2c-149">Must start with a letter.</span></span>
* <span data-ttu-id="beb2c-150">Максимальная длина: 128 символов.</span><span class="sxs-lookup"><span data-stu-id="beb2c-150">Maximum length is 128 characters.</span></span>

### <a name="description"></a><span data-ttu-id="beb2c-151">description</span><span class="sxs-lookup"><span data-stu-id="beb2c-151">description</span></span>

<span data-ttu-id="beb2c-152">Описанию не требуется какой-либо конкретный тег.</span><span class="sxs-lookup"><span data-stu-id="beb2c-152">A description doesn't require any specific tag.</span></span> <span data-ttu-id="beb2c-153">Создайте описание для пользовательской функции, добавив в комментарии JSDoc фразу, описывающую действие функции.</span><span class="sxs-lookup"><span data-stu-id="beb2c-153">Add a description to a custom function by adding a phrase to describe what the function does inside the JSDoc comment.</span></span> <span data-ttu-id="beb2c-154">По умолчанию любой текст без тегов в разделе комментариев JSDoc является описанием функции.</span><span class="sxs-lookup"><span data-stu-id="beb2c-154">By default, whatever text is untagged in the JSDoc comment section will be the description of the function.</span></span> <span data-ttu-id="beb2c-155">Описание отображается для пользователей в Excel при вводе функции.</span><span class="sxs-lookup"><span data-stu-id="beb2c-155">The description appears to users in Excel as they are entering the function.</span></span> <span data-ttu-id="beb2c-156">В следующем примере фраза "A function that sums two numbers" (Функция, суммирующая два числа) — это описание пользовательской функции со свойством id, имеющим значение `SUM`.</span><span class="sxs-lookup"><span data-stu-id="beb2c-156">In the following example, the phrase "A function that sums two numbers" is the description for the custom function with the id property of `SUM`.</span></span>

```JS
/**
/* @customfunction SUM
/* A function that sums two numbers
...
 */
```

---
### <a name="helpurl"></a><span data-ttu-id="beb2c-157">@helpurl</span><span class="sxs-lookup"><span data-stu-id="beb2c-157">@helpurl</span></span>
<a id="helpurl"/>

<span data-ttu-id="beb2c-158">Синтаксис: @helpurl _url_</span><span class="sxs-lookup"><span data-stu-id="beb2c-158">Syntax: @helpurl _url_</span></span>

<span data-ttu-id="beb2c-159">Предоставленный _url_-адрес отображается в Excel.</span><span class="sxs-lookup"><span data-stu-id="beb2c-159">The provided _url_ is displayed in Excel.</span></span>

---
### <a name="param"></a><span data-ttu-id="beb2c-160">@param</span><span class="sxs-lookup"><span data-stu-id="beb2c-160">@param</span></span>
<a id="param"/>

#### <a name="javascript"></a><span data-ttu-id="beb2c-161">JavaScript</span><span class="sxs-lookup"><span data-stu-id="beb2c-161">JavaScript</span></span>

<span data-ttu-id="beb2c-162">Синтаксис JavaScript: @param {type} name _description_</span><span class="sxs-lookup"><span data-stu-id="beb2c-162">JavaScript Syntax: @param {type} name _description_</span></span>

* <span data-ttu-id="beb2c-163">`{type}` должен указывать информацию о типе в фигурных скобках.</span><span class="sxs-lookup"><span data-stu-id="beb2c-163">`{type}` should specify the type info within curly braces.</span></span> <span data-ttu-id="beb2c-164">Дополнительную информацию о типах, которые могут использоваться, см. в разделе [Типы](##types).</span><span class="sxs-lookup"><span data-stu-id="beb2c-164">See the [Types](##types) for more information about the types which may be used.</span></span> <span data-ttu-id="beb2c-165">Необязательно: если тип не указан, будет использоваться тип `any`.</span><span class="sxs-lookup"><span data-stu-id="beb2c-165">Optional: if not specified, the type `any` will be used.</span></span>
* <span data-ttu-id="beb2c-166">`name` указывает, к какому параметру относится тег @param.</span><span class="sxs-lookup"><span data-stu-id="beb2c-166">`name` specifies which parameter the @param tag applies to.</span></span> <span data-ttu-id="beb2c-167">Обязательно.</span><span class="sxs-lookup"><span data-stu-id="beb2c-167">Required.</span></span>
* <span data-ttu-id="beb2c-168">`description` предоставляет описание, которое отображается в Excel для параметра функции.</span><span class="sxs-lookup"><span data-stu-id="beb2c-168">`description` provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="beb2c-169">Необязательно.</span><span class="sxs-lookup"><span data-stu-id="beb2c-169">Optional.</span></span>

<span data-ttu-id="beb2c-170">Чтобы обозначить параметр пользовательской функции как необязательный:</span><span class="sxs-lookup"><span data-stu-id="beb2c-170">To denote a custom function parameter as optional:</span></span>
* <span data-ttu-id="beb2c-171">Поместите имя параметра в квадратные скобки.</span><span class="sxs-lookup"><span data-stu-id="beb2c-171">Put square brackets around the parameter name.</span></span> <span data-ttu-id="beb2c-172">Пример: `@param {string} [text] Optional text`.</span><span class="sxs-lookup"><span data-stu-id="beb2c-172">For example: `@param {string} [text] Optional text`.</span></span>

> [!NOTE]
> <span data-ttu-id="beb2c-173">Значение по умолчанию для дополнительных параметров — `null`.</span><span class="sxs-lookup"><span data-stu-id="beb2c-173">The default value for optional parameters is `null`.</span></span>

#### <a name="typescript"></a><span data-ttu-id="beb2c-174">TypeScript</span><span class="sxs-lookup"><span data-stu-id="beb2c-174">TypeScript</span></span>

<span data-ttu-id="beb2c-175">Синтаксис TypeScript: @param name _description_</span><span class="sxs-lookup"><span data-stu-id="beb2c-175">TypeScript Syntax: @param name _description_</span></span>

* <span data-ttu-id="beb2c-176">`name` указывает, к какому параметру относится тег @param.</span><span class="sxs-lookup"><span data-stu-id="beb2c-176">`name` specifies which parameter the @param tag applies to.</span></span> <span data-ttu-id="beb2c-177">Обязательно.</span><span class="sxs-lookup"><span data-stu-id="beb2c-177">Required.</span></span>
* <span data-ttu-id="beb2c-178">`description` предоставляет описание, которое отображается в Excel для параметра функции.</span><span class="sxs-lookup"><span data-stu-id="beb2c-178">`description` provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="beb2c-179">Необязательно.</span><span class="sxs-lookup"><span data-stu-id="beb2c-179">Optional.</span></span>

<span data-ttu-id="beb2c-180">Дополнительную информацию о типах параметров функций, которые могут использоваться, см. в разделе [Типы](##types).</span><span class="sxs-lookup"><span data-stu-id="beb2c-180">See the [Types](##types) for more information about the function parameter types which may be used.</span></span>

<span data-ttu-id="beb2c-181">Чтобы обозначить параметр пользовательской функции как необязательный, выполните одно из указанных ниже действий.</span><span class="sxs-lookup"><span data-stu-id="beb2c-181">To denote a custom function parameter as optional, do one of the following:</span></span>
* <span data-ttu-id="beb2c-182">Используйте необязательный параметр.</span><span class="sxs-lookup"><span data-stu-id="beb2c-182">Use an optional parameter.</span></span> <span data-ttu-id="beb2c-183">Пример: `function f(text?: string)`</span><span class="sxs-lookup"><span data-stu-id="beb2c-183">For example: `function f(text?: string)`</span></span>
* <span data-ttu-id="beb2c-184">Задайте для параметра значение по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="beb2c-184">Give the parameter a default value.</span></span> <span data-ttu-id="beb2c-185">Пример: `function f(text: string = "abc")`</span><span class="sxs-lookup"><span data-stu-id="beb2c-185">For example: `function f(text: string = "abc")`</span></span>

<span data-ttu-id="beb2c-186">Подробное описание @param см. в [JSDoc](https://usejsdoc.org/tags-param.html)</span><span class="sxs-lookup"><span data-stu-id="beb2c-186">For detailed description of the @param see: [JSDoc](https://usejsdoc.org/tags-param.html)</span></span>

> [!NOTE]
> <span data-ttu-id="beb2c-187">Значение по умолчанию для дополнительных параметров — `null`.</span><span class="sxs-lookup"><span data-stu-id="beb2c-187">The default value for optional parameters is `null`.</span></span>

---
### <a name="requiresaddress"></a><span data-ttu-id="beb2c-188">@requiresAddress</span><span class="sxs-lookup"><span data-stu-id="beb2c-188">@requiresAddress</span></span>
<a id="requiresAddress"/>

<span data-ttu-id="beb2c-189">Указывает, что следует предоставлять адрес ячейки, в которой вычисляется функция.</span><span class="sxs-lookup"><span data-stu-id="beb2c-189">Indicates that the address of the cell where the function is being evaluated should be provided.</span></span> 

<span data-ttu-id="beb2c-190">Тип последнего параметра функции должен быть `CustomFunctions.Invocation` или производной от него.</span><span class="sxs-lookup"><span data-stu-id="beb2c-190">The last function parameter must be of type `CustomFunctions.Invocation` or a derived type.</span></span> <span data-ttu-id="beb2c-191">При вызове функции свойство `address` будет содержать адрес.</span><span class="sxs-lookup"><span data-stu-id="beb2c-191">When the function is called, the `address` property will contain the address.</span></span>

---
### <a name="returns"></a><span data-ttu-id="beb2c-192">@returns</span><span class="sxs-lookup"><span data-stu-id="beb2c-192">@returns</span></span>
<a id="returns"/>

<span data-ttu-id="beb2c-193">Синтаксис: @returns {_type_}</span><span class="sxs-lookup"><span data-stu-id="beb2c-193">Syntax: @returns {_type_}</span></span>

<span data-ttu-id="beb2c-194">Предоставляет тип для возвращаемого значения.</span><span class="sxs-lookup"><span data-stu-id="beb2c-194">Provides the type for the return value.</span></span>

<span data-ttu-id="beb2c-195">Если `{type}` не указан, будет использоваться информация о типе TypeScript.</span><span class="sxs-lookup"><span data-stu-id="beb2c-195">If `{type}` is omitted, the TypeScript type info will be used.</span></span> <span data-ttu-id="beb2c-196">Если информация о типе отсутствует, будет использоваться тип `any`.</span><span class="sxs-lookup"><span data-stu-id="beb2c-196">If there is no type info, the type will be `any`.</span></span>

---
### <a name="streaming"></a><span data-ttu-id="beb2c-197">@streaming</span><span class="sxs-lookup"><span data-stu-id="beb2c-197">@streaming</span></span>
<a id="streaming"/>

<span data-ttu-id="beb2c-198">Используется для обозначения того, что пользовательская функция является потоковой передачей функции.</span><span class="sxs-lookup"><span data-stu-id="beb2c-198">Used to indicate that a custom function is a streaming function.</span></span> 

<span data-ttu-id="beb2c-199">Тип последнего параметра должен быть `CustomFunctions.StreamingInvocation<ResultType>`.</span><span class="sxs-lookup"><span data-stu-id="beb2c-199">The last parameter should be of type `CustomFunctions.StreamingInvocation<ResultType>`.</span></span>
<span data-ttu-id="beb2c-200">Функция должна вернуть значение `void`.</span><span class="sxs-lookup"><span data-stu-id="beb2c-200">The function should return `void`.</span></span>

<span data-ttu-id="beb2c-201">Потоковые передачи функций непосредственно не возвращают значения, для этого необходимо вызывать `setResult(result: ResultType)` с помощью последнего параметра.</span><span class="sxs-lookup"><span data-stu-id="beb2c-201">Streaming functions do not return values directly, but rather should call `setResult(result: ResultType)` using the last parameter.</span></span>

<span data-ttu-id="beb2c-202">Исключения, которые возникают при потоковой передаче функций, игнорируются.</span><span class="sxs-lookup"><span data-stu-id="beb2c-202">Exceptions thrown by a streaming function are ignored.</span></span> <span data-ttu-id="beb2c-203">`setResult()` при вызове может вернуть ошибку в качестве результата.</span><span class="sxs-lookup"><span data-stu-id="beb2c-203">`setResult()` may be called with Error to indicate an error result.</span></span>

<span data-ttu-id="beb2c-204">Потоковые передачи функций невозможно пометить как [@volatile](#volatile).</span><span class="sxs-lookup"><span data-stu-id="beb2c-204">Streaming functions cannot be marked as [@volatile](#volatile).</span></span>

---
### <a name="volatile"></a><span data-ttu-id="beb2c-205">@volatile</span><span class="sxs-lookup"><span data-stu-id="beb2c-205">@volatile</span></span>
<a id="volatile"/>

<span data-ttu-id="beb2c-206">Переменные функции — это такие функции, чей результат не остается неизменным в каждый период времени, даже если они не содержат аргументов или их аргументы не меняются.</span><span class="sxs-lookup"><span data-stu-id="beb2c-206">A volatile function is one whose result cannot be assumed to be the same from one moment to the next even if it takes no arguments or the arguments have not changed.</span></span> <span data-ttu-id="beb2c-207">Excel повторно проводит вычисления в ячейках, которые содержат переменные функции, вместе со всеми зависимыми функциями при каждом вычислении.</span><span class="sxs-lookup"><span data-stu-id="beb2c-207">Excel re-evaluates cells that contain volatile functions, together with all dependents, every time that a calculation is done.</span></span> <span data-ttu-id="beb2c-208">По этой причине чрезмерное использование переменных функций может замедлить пересчет, поэтому используйте их умеренно.</span><span class="sxs-lookup"><span data-stu-id="beb2c-208">For this reason, too much reliance on volatile functions can make recalculation times slow, so use them sparingly.</span></span>

<span data-ttu-id="beb2c-209">Потоковые передачи функций не могут быть переменными.</span><span class="sxs-lookup"><span data-stu-id="beb2c-209">Streaming functions cannot be volatile.</span></span>

---

## <a name="types"></a><span data-ttu-id="beb2c-210">Типы</span><span class="sxs-lookup"><span data-stu-id="beb2c-210">Types</span></span>

<span data-ttu-id="beb2c-211">Указывая тип параметра, Excel преобразует значения в этот тип, прежде чем вызывать функцию.</span><span class="sxs-lookup"><span data-stu-id="beb2c-211">By specifying a parameter type, Excel will convert values into that type before calling the function.</span></span> <span data-ttu-id="beb2c-212">Если указан тип `any`, преобразование выполняться не будет.</span><span class="sxs-lookup"><span data-stu-id="beb2c-212">If the type is `any`, no conversion will be performed.</span></span>

### <a name="value-types"></a><span data-ttu-id="beb2c-213">Типы значений</span><span class="sxs-lookup"><span data-stu-id="beb2c-213">Value types</span></span>

<span data-ttu-id="beb2c-214">Одно значение может быть представлено с помощью одного из приведенных ниже типов: `boolean`, `number`, `string`.</span><span class="sxs-lookup"><span data-stu-id="beb2c-214">A single value may be represented using one of the following types: `boolean`, `number`, `string`.</span></span>

### <a name="matrix-type"></a><span data-ttu-id="beb2c-215">Тип "матрица"</span><span class="sxs-lookup"><span data-stu-id="beb2c-215">Matrix type</span></span>

<span data-ttu-id="beb2c-216">Используйте тип двумерного массива, чтобы параметр или возвращаемое значение представляли собой матрицу значений.</span><span class="sxs-lookup"><span data-stu-id="beb2c-216">Use a two-dimensional array type to have the parameter or return value be a matrix of values.</span></span> <span data-ttu-id="beb2c-217">Например, тип `number[][]` указывает на матрицу чисел.</span><span class="sxs-lookup"><span data-stu-id="beb2c-217">For example, the type `number[][]` indicates a matrix of numbers.</span></span> <span data-ttu-id="beb2c-218">`string[][]` указывает на матрицу строк.</span><span class="sxs-lookup"><span data-stu-id="beb2c-218">`string[][]` indicates a matrix of strings.</span></span> 

### <a name="error-type"></a><span data-ttu-id="beb2c-219">Тип "ошибка"</span><span class="sxs-lookup"><span data-stu-id="beb2c-219">Error type</span></span>

<span data-ttu-id="beb2c-220">Функция непотоковой передачи может указывать на ошибку, возвращая тип "Ошибка".</span><span class="sxs-lookup"><span data-stu-id="beb2c-220">A non-streaming function can indicate an error by returning an Error type.</span></span>

<span data-ttu-id="beb2c-221">Функция потоковой передачи может указывать на ошибку, вызывая метод `setResult()` типа "Ошибка".</span><span class="sxs-lookup"><span data-stu-id="beb2c-221">A streaming function can indicate an error by calling setResult() with an Error type.</span></span>

### <a name="promise"></a><span data-ttu-id="beb2c-222">Обещание</span><span class="sxs-lookup"><span data-stu-id="beb2c-222">Promise</span></span>

<span data-ttu-id="beb2c-223">Функция может вернуть тип "Обещание", который задаст значение, когда обещание будет разрешено.</span><span class="sxs-lookup"><span data-stu-id="beb2c-223">A function can return a Promise, which will provide the value when the promise is resolved.</span></span> <span data-ttu-id="beb2c-224">В случае отклонения обещания возникнет ошибка.</span><span class="sxs-lookup"><span data-stu-id="beb2c-224">If the promise is rejected, then it is an error.</span></span>

### <a name="other-types"></a><span data-ttu-id="beb2c-225">Другие типы</span><span class="sxs-lookup"><span data-stu-id="beb2c-225">Other types</span></span>

<span data-ttu-id="beb2c-226">Любой другой тип будет рассматриваться как ошибка.</span><span class="sxs-lookup"><span data-stu-id="beb2c-226">Any other type will be treated as an error.</span></span>

## <a name="next-steps"></a><span data-ttu-id="beb2c-227">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="beb2c-227">Next steps</span></span>
<span data-ttu-id="beb2c-228">Узнайте о [соглашениях именования для пользовательских функций](custom-functions-naming.md).</span><span class="sxs-lookup"><span data-stu-id="beb2c-228">Learn about [naming conventions for custom functions](custom-functions-naming.md).</span></span> <span data-ttu-id="beb2c-229">Или же узнайте, как [локализовать свои функции](custom-functions-localize.md), для чего нужно [записать файл JSON вручную](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="beb2c-229">Alternatively, learn how to [localize your functions](custom-functions-localize.md) which requires you to [write your JSON file by hand](custom-functions-json.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="beb2c-230">Дополнительные ресурсы</span><span class="sxs-lookup"><span data-stu-id="beb2c-230">See also</span></span>

* [<span data-ttu-id="beb2c-231">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="beb2c-231">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="beb2c-232">Рекомендации по пользовательским функциям</span><span class="sxs-lookup"><span data-stu-id="beb2c-232">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="beb2c-233">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="beb2c-233">Create custom functions in Excel</span></span>](custom-functions-overview.md)
