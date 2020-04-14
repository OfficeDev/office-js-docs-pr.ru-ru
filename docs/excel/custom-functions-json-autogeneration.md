---
ms.date: 04/13/2020
description: Использование тегов JSDoc для динамического создания метаданных JSON пользовательских функций.
title: Автоматическое генерирование метаданных JSON для пользовательских функций
localization_priority: Normal
ms.openlocfilehash: d62def1d107d4472d2e89623c4f10b5cd012f28c
ms.sourcegitcommit: 118e8bcbcfb73c93e2053bda67fe8dd20799b170
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/13/2020
ms.locfileid: "43241072"
---
# <a name="autogenerate-json-metadata-for-custom-functions"></a><span data-ttu-id="c2683-103">Автоматическое генерирование метаданных JSON для пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="c2683-103">Autogenerate JSON metadata for custom functions</span></span>

<span data-ttu-id="c2683-104">Если пользовательская функция Excel написана в JavaScript или TypeScript, [теги JSDoc](https://jsdoc.app/) используются для предоставления дополнительной информации о пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="c2683-104">When an Excel custom function is written in JavaScript or TypeScript, [JSDoc tags](https://jsdoc.app/) are used to provide extra information about the custom function.</span></span> <span data-ttu-id="c2683-105">Теги JSDoc используются при сборке для создания [файла метаданных JSON](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="c2683-105">The JSDoc tags are then used at build time to create the [JSON metadata file](custom-functions-json.md).</span></span> <span data-ttu-id="c2683-106">Использование тегов JSDoc освобождает вас от необходимости редактировать файл метаданных JSON вручную.</span><span class="sxs-lookup"><span data-stu-id="c2683-106">Using JSDoc tags saves you from the effort of manually editing the JSON metadata file.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="c2683-107">Добавьте тег `@customfunction` в примечаниях к коду для функции JavaScript или TypeScript, чтобы пометить ее как пользовательскую.</span><span class="sxs-lookup"><span data-stu-id="c2683-107">Add the `@customfunction` tag in the code comments for a JavaScript or TypeScript function to mark it as a custom function.</span></span>

<span data-ttu-id="c2683-108">Типы параметров функции можно получить с помощью тега [@param](#param) в JavaScript или из раздела [Тип функции](https://www.typescriptlang.org/docs/handbook/functions.html) в TypeScript.</span><span class="sxs-lookup"><span data-stu-id="c2683-108">The function parameter types may be provided using the [@param](#param) tag in JavaScript, or from the [Function type](https://www.typescriptlang.org/docs/handbook/functions.html) in TypeScript.</span></span> <span data-ttu-id="c2683-109">Дополнительные сведения см. в разделах, посвященных тегу [@param](#param) и [типам](#types).</span><span class="sxs-lookup"><span data-stu-id="c2683-109">For more information, see the [@param](#param) tag and [Types](#types) sections.</span></span>

### <a name="adding-a-description-to-a-function"></a><span data-ttu-id="c2683-110">Добавление описания функции</span><span class="sxs-lookup"><span data-stu-id="c2683-110">Adding a description to a function</span></span>

<span data-ttu-id="c2683-111">Описание отображается пользователю в качестве текста справки, если ему непонятно действие пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="c2683-111">The description is displayed to the user as help text when they need help to understand what your custom function does.</span></span> <span data-ttu-id="c2683-112">Описанию не требуется какой-либо конкретный тег.</span><span class="sxs-lookup"><span data-stu-id="c2683-112">The description doesn't require any specific tag.</span></span> <span data-ttu-id="c2683-113">Просто введите краткий текст описания в комментарии JSDoc.</span><span class="sxs-lookup"><span data-stu-id="c2683-113">Just enter a short text description in the JSDoc comment.</span></span> <span data-ttu-id="c2683-114">Обычно описание размещается в начале раздела комментариев JSDoc, но оно поддерживается независимо от места размещения.</span><span class="sxs-lookup"><span data-stu-id="c2683-114">In general the description is placed at the start of the JSDoc comment section, but it will work no matter where it is placed.</span></span>

<span data-ttu-id="c2683-115">Чтобы просмотреть примеры описаний встроенных функций, откройте Excel, перейдите на вкладку **Формулы** и нажмите кнопку **Вставить функцию**.</span><span class="sxs-lookup"><span data-stu-id="c2683-115">To see examples of the built-in function descriptions, open Excel, go to the **Formulas** tab, and choose **Insert function**.</span></span> <span data-ttu-id="c2683-116">Вы сможете просмотреть все описания функций, а также список собственных пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="c2683-116">You can then browse through all the function descriptions, and also see your own custom functions listed.</span></span>

<span data-ttu-id="c2683-117">В следующем примере фраза "Calculates the volume of a sphere." (Вычисляет объем сферы)</span><span class="sxs-lookup"><span data-stu-id="c2683-117">In the following example, the phrase "Calculates the volume of a sphere."</span></span> <span data-ttu-id="c2683-118">является описанием пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="c2683-118">is the description for the custom function.</span></span>

```js
/**
/* Calculates the volume of a sphere.
/* @customfunction VOLUME
...
 */
```


## <a name="jsdoc-tags"></a><span data-ttu-id="c2683-119">Теги JSDoc</span><span class="sxs-lookup"><span data-stu-id="c2683-119">JSDoc Tags</span></span>
<span data-ttu-id="c2683-120">Ниже приведены теги JSDoc, которые поддерживаются в пользовательских функциях Excel:</span><span class="sxs-lookup"><span data-stu-id="c2683-120">The following JSDoc tags are supported in Excel custom functions:</span></span>
* [<span data-ttu-id="c2683-121">@cancelable</span><span class="sxs-lookup"><span data-stu-id="c2683-121">@cancelable</span></span>](#cancelable)
* <span data-ttu-id="c2683-122">[@customfunction](#customfunction) id name</span><span class="sxs-lookup"><span data-stu-id="c2683-122">[@customfunction](#customfunction) id name</span></span>
* <span data-ttu-id="c2683-123">[@helpurl](#helpurl) url</span><span class="sxs-lookup"><span data-stu-id="c2683-123">[@helpurl](#helpurl) url</span></span>
* <span data-ttu-id="c2683-124">[@param](#param) _{type}_ name description</span><span class="sxs-lookup"><span data-stu-id="c2683-124">[@param](#param) _{type}_ name description</span></span>
* [<span data-ttu-id="c2683-125">@requiresAddress</span><span class="sxs-lookup"><span data-stu-id="c2683-125">@requiresAddress</span></span>](#requiresAddress)
* <span data-ttu-id="c2683-126">[@returns](#returns) _{type}_</span><span class="sxs-lookup"><span data-stu-id="c2683-126">[@returns](#returns) _{type}_</span></span>
* [<span data-ttu-id="c2683-127">@streaming</span><span class="sxs-lookup"><span data-stu-id="c2683-127">@streaming</span></span>](#streaming)
* [<span data-ttu-id="c2683-128">@volatile</span><span class="sxs-lookup"><span data-stu-id="c2683-128">@volatile</span></span>](#volatile)

---
### <a name="cancelable"></a><span data-ttu-id="c2683-129">@cancelable</span><span class="sxs-lookup"><span data-stu-id="c2683-129">@cancelable</span></span>
<a id="cancelable"/>

<span data-ttu-id="c2683-130">При отмене функции указывает, что пользовательская функция стремится к выполнению действия.</span><span class="sxs-lookup"><span data-stu-id="c2683-130">Indicates that a custom function wants to perform an action when the function is canceled.</span></span>

<span data-ttu-id="c2683-131">В качестве типа последнего параметра функции должно быть указано `CustomFunctions.CancelableInvocation`.</span><span class="sxs-lookup"><span data-stu-id="c2683-131">The last function parameter must be of type `CustomFunctions.CancelableInvocation`.</span></span> <span data-ttu-id="c2683-132">Функция может назначить функцию свойству `oncanceled`, чтобы обозначить действия для выполнения в случае отмены функции.</span><span class="sxs-lookup"><span data-stu-id="c2683-132">The function can assign a function to the `oncanceled` property to denote the action to perform when the function is canceled.</span></span>

<span data-ttu-id="c2683-133">Если тип последнего параметра функции `CustomFunctions.CancelableInvocation`, он будет рассматриваться как `@cancelable`, даже если тег отсутствует.</span><span class="sxs-lookup"><span data-stu-id="c2683-133">If the last function parameter is of type `CustomFunctions.CancelableInvocation`, it will be considered `@cancelable` even if the tag isn't present.</span></span>

<span data-ttu-id="c2683-134">Функция не может содержать одновременно теги `@cancelable` и `@streaming`.</span><span class="sxs-lookup"><span data-stu-id="c2683-134">A function can't have both `@cancelable` and `@streaming` tags.</span></span>

---
### <a name="customfunction"></a><span data-ttu-id="c2683-135">@customfunction</span><span class="sxs-lookup"><span data-stu-id="c2683-135">@customfunction</span></span>
<a id="customfunction"/>

<span data-ttu-id="c2683-136">Синтаксис: @customfunction _id_ _name_</span><span class="sxs-lookup"><span data-stu-id="c2683-136">Syntax: @customfunction _id_ _name_</span></span>

<span data-ttu-id="c2683-137">Укажите этот тег, чтобы рассматривать функцию JavaScript или TypeScript как пользовательскую функцию Excel.</span><span class="sxs-lookup"><span data-stu-id="c2683-137">Specify this tag to treat the JavaScript/TypeScript function as an Excel custom function.</span></span>

<span data-ttu-id="c2683-138">Этот тег необходим, чтобы создать метаданные для пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="c2683-138">This tag is required to create metadata for the custom function.</span></span>

<span data-ttu-id="c2683-139">В следующем примере показан самый простой способ объявления пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="c2683-139">The following example shows the simplest way to declare a custom function.</span></span>

```js
/**
 * Increments a value once a second.
 * @customfunction
 * ...
 */
```

#### <a name="id"></a><span data-ttu-id="c2683-140">id</span><span class="sxs-lookup"><span data-stu-id="c2683-140">id</span></span>

<span data-ttu-id="c2683-141">`id` является инвариантным идентификатором для пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="c2683-141">The `id` is an invariant identifier for the custom function.</span></span>

* <span data-ttu-id="c2683-142">Если `id` не указан, название функции JavaScript или TypeScript преобразуется в верхний регистр, а недопустимые символы удаляются.</span><span class="sxs-lookup"><span data-stu-id="c2683-142">If `id` isn't provided, the JavaScript/TypeScript function name is converted to uppercase and disallowed characters are removed.</span></span>
* <span data-ttu-id="c2683-143">`id` должен быть уникальным для всех пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="c2683-143">The `id` must be unique for all custom functions.</span></span>
* <span data-ttu-id="c2683-144">Допустимые символы: A — Z, a — z, 0–9, символ подчеркивания (\_) и точка (.).</span><span class="sxs-lookup"><span data-stu-id="c2683-144">The allowed characters are limited to: A-Z, a-z, 0-9, underscores (\_), and period (.).</span></span>

<span data-ttu-id="c2683-145">В следующем примере increment — это параметр `id` и `name` функции.</span><span class="sxs-lookup"><span data-stu-id="c2683-145">In the following example, increment is the `id` and the `name` of the function.</span></span>

```js
/**
 * Increments a value once a second.
 * @customfunction INCREMENT
 * ...
 */
```

#### <a name="name"></a><span data-ttu-id="c2683-146">name</span><span class="sxs-lookup"><span data-stu-id="c2683-146">name</span></span>

<span data-ttu-id="c2683-147">Предоставляет отображаемый параметр `name` для пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="c2683-147">Provides the display `name` for the custom function.</span></span>

* <span data-ttu-id="c2683-148">Если имя не указано, идентификатор также используется как имя.</span><span class="sxs-lookup"><span data-stu-id="c2683-148">If name isn't provided, the id is also used as the name.</span></span>
* <span data-ttu-id="c2683-149">Допустимые символы: буквы [буквенные символы Юникод](http://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), числа, точки (.) и подчеркивания (\_).</span><span class="sxs-lookup"><span data-stu-id="c2683-149">Allowed characters: Letters [Unicode Alphabetic character](http://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), numbers, period (.), and underscore (\_).</span></span>
* <span data-ttu-id="c2683-150">Имя должно начинаться с буквы.</span><span class="sxs-lookup"><span data-stu-id="c2683-150">Must start with a letter.</span></span>
* <span data-ttu-id="c2683-151">Максимальная длина: 128 символов.</span><span class="sxs-lookup"><span data-stu-id="c2683-151">Maximum length is 128 characters.</span></span>

<span data-ttu-id="c2683-152">В следующем примере INC — это параметр `id` функции, а `increment` — параметр `name`.</span><span class="sxs-lookup"><span data-stu-id="c2683-152">In the following example, INC is the `id` of the function and `increment` is the `name`.</span></span>

```js
/**
 * Increments a value once a second.
 * @customfunction INC INCREMENT
 * ...
 */
```

### <a name="description"></a><span data-ttu-id="c2683-153">description</span><span class="sxs-lookup"><span data-stu-id="c2683-153">description</span></span>

<span data-ttu-id="c2683-154">Описанию не требуется какой-либо конкретный тег.</span><span class="sxs-lookup"><span data-stu-id="c2683-154">A description doesn't require any specific tag.</span></span> <span data-ttu-id="c2683-155">Создайте описание для пользовательской функции, добавив в комментарии JSDoc фразу, описывающую действие функции.</span><span class="sxs-lookup"><span data-stu-id="c2683-155">Add a description to a custom function by adding a phrase to describe what the function does inside the JSDoc comment.</span></span> <span data-ttu-id="c2683-156">По умолчанию любой текст без тегов в разделе комментариев JSDoc является описанием функции.</span><span class="sxs-lookup"><span data-stu-id="c2683-156">By default, whatever text is untagged in the JSDoc comment section will be the description of the function.</span></span> <span data-ttu-id="c2683-157">Описание отображается для пользователей в Excel при вводе функции.</span><span class="sxs-lookup"><span data-stu-id="c2683-157">The description appears to users in Excel as they are entering the function.</span></span> <span data-ttu-id="c2683-158">В следующем примере фраза "A function that adds two numbers" (Функция, складывающая два числа) — это описание пользовательской функции со свойством id, имеющим значение `ADD`.</span><span class="sxs-lookup"><span data-stu-id="c2683-158">In the following example, the phrase "A function that adds two numbers" is the description for the custom function with the id property of `ADD`.</span></span>

<span data-ttu-id="c2683-159">В следующем примере ADD — это параметр `id` и `name` функции с указанным описанием.</span><span class="sxs-lookup"><span data-stu-id="c2683-159">In the following example, ADD is the `id` and `name` of the function and a description is given.</span></span>

```js
/**
 * A function that adds two numbers.
 * @customfunction ADD
 * ...
 */
```

---
### <a name="helpurl"></a><span data-ttu-id="c2683-160">@helpurl</span><span class="sxs-lookup"><span data-stu-id="c2683-160">@helpurl</span></span>
<a id="helpurl"/>

<span data-ttu-id="c2683-161">Синтаксис: @helpurl _url_</span><span class="sxs-lookup"><span data-stu-id="c2683-161">Syntax: @helpurl _url_</span></span>

<span data-ttu-id="c2683-162">Предоставленный _url_-адрес отображается в Excel.</span><span class="sxs-lookup"><span data-stu-id="c2683-162">The provided _url_ is displayed in Excel.</span></span>

<span data-ttu-id="c2683-163">В следующем примере `helpurl` используется `www.contoso.com/weatherhelp`значение.</span><span class="sxs-lookup"><span data-stu-id="c2683-163">In the following example, the `helpurl` is `www.contoso.com/weatherhelp`.</span></span>

```js
/**
 * A function which streams the temperature in a town you specify.
 * @customfunction getTemperature
 * @helpurl www.contoso.com/weatherhelp
 * ...
 */
```

---
### <a name="param"></a><span data-ttu-id="c2683-164">@param</span><span class="sxs-lookup"><span data-stu-id="c2683-164">@param</span></span>
<a id="param"/>

#### <a name="javascript"></a><span data-ttu-id="c2683-165">JavaScript</span><span class="sxs-lookup"><span data-stu-id="c2683-165">JavaScript</span></span>

<span data-ttu-id="c2683-166">Синтаксис JavaScript: @param {type} name _description_</span><span class="sxs-lookup"><span data-stu-id="c2683-166">JavaScript Syntax: @param {type} name _description_</span></span>

* <span data-ttu-id="c2683-167">`{type}` должен указывать информацию о типе в фигурных скобках.</span><span class="sxs-lookup"><span data-stu-id="c2683-167">`{type}` should specify the type info within curly braces.</span></span> <span data-ttu-id="c2683-168">Дополнительную информацию о типах, которые могут использоваться, см. в разделе [Типы](#types).</span><span class="sxs-lookup"><span data-stu-id="c2683-168">See the [Types](#types) section for more information about the types which may be used.</span></span> <span data-ttu-id="c2683-169">Необязательно: если тип не указан, будет использоваться тип `any`.</span><span class="sxs-lookup"><span data-stu-id="c2683-169">Optional: if not specified, the type `any` will be used.</span></span>
* <span data-ttu-id="c2683-170">`name` указывает, к какому параметру относится тег @param.</span><span class="sxs-lookup"><span data-stu-id="c2683-170">`name` specifies which parameter the @param tag applies to.</span></span> <span data-ttu-id="c2683-171">Обязательно.</span><span class="sxs-lookup"><span data-stu-id="c2683-171">Required.</span></span>
* <span data-ttu-id="c2683-172">`description` предоставляет описание, которое отображается в Excel для параметра функции.</span><span class="sxs-lookup"><span data-stu-id="c2683-172">`description` provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="c2683-173">Необязательно.</span><span class="sxs-lookup"><span data-stu-id="c2683-173">Optional.</span></span>

<span data-ttu-id="c2683-174">Чтобы обозначить параметр пользовательской функции как необязательный:</span><span class="sxs-lookup"><span data-stu-id="c2683-174">To denote a custom function parameter as optional:</span></span>
* <span data-ttu-id="c2683-175">Поместите имя параметра в квадратные скобки.</span><span class="sxs-lookup"><span data-stu-id="c2683-175">Put square brackets around the parameter name.</span></span> <span data-ttu-id="c2683-176">Пример: `@param {string} [text] Optional text`.</span><span class="sxs-lookup"><span data-stu-id="c2683-176">For example: `@param {string} [text] Optional text`.</span></span>

> [!NOTE]
> <span data-ttu-id="c2683-177">Значение по умолчанию для дополнительных параметров — `null`.</span><span class="sxs-lookup"><span data-stu-id="c2683-177">The default value for optional parameters is `null`.</span></span>

<span data-ttu-id="c2683-178">В следующем примере показана функция ADD, складывающая два или три числа, где третье число является необязательным параметром.</span><span class="sxs-lookup"><span data-stu-id="c2683-178">The following example shows a ADD function which adds two or three numbers, with the third number as an optional parameter.</span></span>

```js
/**
 * A function which sums two, or optionally three, numbers.
 * @customfunction ADDNUMBERS
 * @param firstNumber {number} First number to add.
 * @param secondNumber {number} Second number to add.
 * @param [thirdNumber] {number} Optional third number you wish to add.
 * ...
 */
```

#### <a name="typescript"></a><span data-ttu-id="c2683-179">TypeScript</span><span class="sxs-lookup"><span data-stu-id="c2683-179">TypeScript</span></span>

<span data-ttu-id="c2683-180">Синтаксис TypeScript: @param name _description_</span><span class="sxs-lookup"><span data-stu-id="c2683-180">TypeScript Syntax: @param name _description_</span></span>

* <span data-ttu-id="c2683-181">`name` указывает, к какому параметру относится тег @param.</span><span class="sxs-lookup"><span data-stu-id="c2683-181">`name` specifies which parameter the @param tag applies to.</span></span> <span data-ttu-id="c2683-182">Обязательно.</span><span class="sxs-lookup"><span data-stu-id="c2683-182">Required.</span></span>
* <span data-ttu-id="c2683-183">`description` предоставляет описание, которое отображается в Excel для параметра функции.</span><span class="sxs-lookup"><span data-stu-id="c2683-183">`description` provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="c2683-184">Необязательное свойство.</span><span class="sxs-lookup"><span data-stu-id="c2683-184">Optional.</span></span>

<span data-ttu-id="c2683-185">Дополнительные сведения о типах параметров функций, которые могут использоваться, см. в разделе [Типы](#types).</span><span class="sxs-lookup"><span data-stu-id="c2683-185">See the [Types](#types) section for more information about the function parameter types which may be used.</span></span>

<span data-ttu-id="c2683-186">Чтобы обозначить параметр пользовательской функции как необязательный, выполните одно из указанных ниже действий.</span><span class="sxs-lookup"><span data-stu-id="c2683-186">To denote a custom function parameter as optional, do one of the following:</span></span>
* <span data-ttu-id="c2683-187">Используйте необязательный параметр.</span><span class="sxs-lookup"><span data-stu-id="c2683-187">Use an optional parameter.</span></span> <span data-ttu-id="c2683-188">Пример: `function f(text?: string)`</span><span class="sxs-lookup"><span data-stu-id="c2683-188">For example: `function f(text?: string)`</span></span>
* <span data-ttu-id="c2683-189">Задайте для параметра значение по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="c2683-189">Give the parameter a default value.</span></span> <span data-ttu-id="c2683-190">Пример: `function f(text: string = "abc")`</span><span class="sxs-lookup"><span data-stu-id="c2683-190">For example: `function f(text: string = "abc")`</span></span>

<span data-ttu-id="c2683-191">Подробное описание @param см. в [JSDoc](https://jsdoc.app/tags-param.html)</span><span class="sxs-lookup"><span data-stu-id="c2683-191">For detailed description of the @param see: [JSDoc](https://jsdoc.app/tags-param.html)</span></span>

> [!NOTE]
> <span data-ttu-id="c2683-192">Значение по умолчанию для дополнительных параметров — `null`.</span><span class="sxs-lookup"><span data-stu-id="c2683-192">The default value for optional parameters is `null`.</span></span>

<span data-ttu-id="c2683-193">В следующем примере показана функция `add`, складывающая два числа.</span><span class="sxs-lookup"><span data-stu-id="c2683-193">The following example shows the `add` function that adds two numbers.</span></span>

```ts
/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
function add(first: number, second: number): number {
  return first + second;
}
```

---
### <a name="requiresaddress"></a><span data-ttu-id="c2683-194">@requiresAddress</span><span class="sxs-lookup"><span data-stu-id="c2683-194">@requiresAddress</span></span>
<a id="requiresAddress"/>

<span data-ttu-id="c2683-195">Указывает, что следует предоставлять адрес ячейки, в которой вычисляется функция.</span><span class="sxs-lookup"><span data-stu-id="c2683-195">Indicates that the address of the cell where the function is being evaluated should be provided.</span></span>

<span data-ttu-id="c2683-196">Тип последнего параметра функции должен быть `CustomFunctions.Invocation` или производной от него.</span><span class="sxs-lookup"><span data-stu-id="c2683-196">The last function parameter must be of type `CustomFunctions.Invocation` or a derived type.</span></span> <span data-ttu-id="c2683-197">При вызове функции свойство `address` будет содержать адрес.</span><span class="sxs-lookup"><span data-stu-id="c2683-197">When the function is called, the `address` property will contain the address.</span></span> <span data-ttu-id="c2683-198">Пример функции, использующей тег `@requiresAddress`, см. в разделе [Обращение к параметру контекста ячейки](custom-functions-parameter-options.md#addressing-cells-context-parameter)</span><span class="sxs-lookup"><span data-stu-id="c2683-198">For an example of a function that uses the `@requiresAddress` tag, see [Addressing cell's context parameter](custom-functions-parameter-options.md#addressing-cells-context-parameter).</span></span>

---
### <a name="returns"></a><span data-ttu-id="c2683-199">@returns</span><span class="sxs-lookup"><span data-stu-id="c2683-199">@returns</span></span>
<a id="returns"/>

<span data-ttu-id="c2683-200">Синтаксис: @returns {_type_}</span><span class="sxs-lookup"><span data-stu-id="c2683-200">Syntax: @returns {_type_}</span></span>

<span data-ttu-id="c2683-201">Предоставляет тип для возвращаемого значения.</span><span class="sxs-lookup"><span data-stu-id="c2683-201">Provides the type for the return value.</span></span>

<span data-ttu-id="c2683-202">Если `{type}` не указан, будет использоваться информация о типе TypeScript.</span><span class="sxs-lookup"><span data-stu-id="c2683-202">If `{type}` is omitted, the TypeScript type info will be used.</span></span> <span data-ttu-id="c2683-203">Если информация о типе отсутствует, будет использоваться тип `any`.</span><span class="sxs-lookup"><span data-stu-id="c2683-203">If there is no type info, the type will be `any`.</span></span>

<span data-ttu-id="c2683-204">В следующем примере показана функция `add`, использующая тег `@returns`.</span><span class="sxs-lookup"><span data-stu-id="c2683-204">The following example shows the `add` function that uses the `@returns` tag.</span></span>

```ts
/**
 * Adds two numbers.
 * @customfunction 
 * @param first First number
 * @param second Second number
 * @returns The sum of the two numbers.
 */
function add(first: number, second: number): number {
  return first + second;
}
```

---
### <a name="streaming"></a><span data-ttu-id="c2683-205">@streaming</span><span class="sxs-lookup"><span data-stu-id="c2683-205">@streaming</span></span>
<a id="streaming"/>

<span data-ttu-id="c2683-206">Используется для обозначения того, что пользовательская функция является потоковой передачей функции.</span><span class="sxs-lookup"><span data-stu-id="c2683-206">Used to indicate that a custom function is a streaming function.</span></span> 

<span data-ttu-id="c2683-207">Тип последнего параметра должен быть `CustomFunctions.StreamingInvocation<ResultType>`.</span><span class="sxs-lookup"><span data-stu-id="c2683-207">The last parameter should be of type `CustomFunctions.StreamingInvocation<ResultType>`.</span></span>
<span data-ttu-id="c2683-208">Функция должна вернуть значение `void`.</span><span class="sxs-lookup"><span data-stu-id="c2683-208">The function should return `void`.</span></span>

<span data-ttu-id="c2683-209">Потоковые передачи функций непосредственно не возвращают значения, для этого необходимо вызывать `setResult(result: ResultType)` с помощью последнего параметра.</span><span class="sxs-lookup"><span data-stu-id="c2683-209">Streaming functions don't return values directly, but rather should call `setResult(result: ResultType)` using the last parameter.</span></span>

<span data-ttu-id="c2683-210">Исключения, которые возникают при потоковой передаче функций, игнорируются.</span><span class="sxs-lookup"><span data-stu-id="c2683-210">Exceptions thrown by a streaming function are ignored.</span></span> <span data-ttu-id="c2683-211">`setResult()` при вызове может вернуть ошибку в качестве результата.</span><span class="sxs-lookup"><span data-stu-id="c2683-211">`setResult()` may be called with Error to indicate an error result.</span></span> <span data-ttu-id="c2683-212">Пример функции потоковой передачи и дополнительные сведения см. в разделе [Создание функции потоковой передачи](./custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="c2683-212">For an example of a streaming function and more information, see [Make a streaming function](./custom-functions-web-reqs.md#make-a-streaming-function).</span></span>

<span data-ttu-id="c2683-213">Потоковые передачи функций невозможно пометить как [@volatile](#volatile).</span><span class="sxs-lookup"><span data-stu-id="c2683-213">Streaming functions can't be marked as [@volatile](#volatile).</span></span>

---
### <a name="volatile"></a><span data-ttu-id="c2683-214">@volatile</span><span class="sxs-lookup"><span data-stu-id="c2683-214">@volatile</span></span>
<a id="volatile"/>

<span data-ttu-id="c2683-215">Переменные функции — это такие функции, чей результат не остается неизменным в каждый период времени, даже если они не содержат аргументов или их аргументы не меняются.</span><span class="sxs-lookup"><span data-stu-id="c2683-215">A volatile function is one whose result isn't the same from one moment to the next, even if it takes no arguments or the arguments haven't changed.</span></span> <span data-ttu-id="c2683-216">Excel повторно проводит вычисления в ячейках, которые содержат переменные функции, вместе со всеми зависимыми функциями при каждом вычислении.</span><span class="sxs-lookup"><span data-stu-id="c2683-216">Excel re-evaluates cells that contain volatile functions, together with all dependents, every time that a calculation is done.</span></span> <span data-ttu-id="c2683-217">По этой причине чрезмерное использование переменных функций может замедлить пересчет, поэтому используйте их умеренно.</span><span class="sxs-lookup"><span data-stu-id="c2683-217">For this reason, too much reliance on volatile functions can make recalculation times slow, so use them sparingly.</span></span>

<span data-ttu-id="c2683-218">Потоковые передачи функций не могут быть переменными.</span><span class="sxs-lookup"><span data-stu-id="c2683-218">Streaming functions can't be volatile.</span></span>

<span data-ttu-id="c2683-219">Следующая функция является переменной и использует тег `@volatile`.</span><span class="sxs-lookup"><span data-stu-id="c2683-219">The following function is volatile and uses the `@volatile` tag.</span></span>

```js
/**
 * Simulates rolling a 6-sided dice.
 * @customfunction
 * @volatile
 */
function roll6sided(): number {
  return Math.floor(Math.random() * 6) + 1;
}
```

---

## <a name="types"></a><span data-ttu-id="c2683-220">Типы</span><span class="sxs-lookup"><span data-stu-id="c2683-220">Types</span></span>

<span data-ttu-id="c2683-221">Указывая тип параметра, Excel преобразует значения в этот тип, прежде чем вызывать функцию.</span><span class="sxs-lookup"><span data-stu-id="c2683-221">By specifying a parameter type, Excel will convert values into that type before calling the function.</span></span> <span data-ttu-id="c2683-222">Если указан тип `any`, преобразование выполняться не будет.</span><span class="sxs-lookup"><span data-stu-id="c2683-222">If the type is `any`, no conversion will be performed.</span></span>

### <a name="value-types"></a><span data-ttu-id="c2683-223">Типы значений</span><span class="sxs-lookup"><span data-stu-id="c2683-223">Value types</span></span>

<span data-ttu-id="c2683-224">Одно значение может быть представлено с помощью одного из приведенных ниже типов: `boolean`, `number`, `string`.</span><span class="sxs-lookup"><span data-stu-id="c2683-224">A single value may be represented using one of the following types: `boolean`, `number`, `string`.</span></span>

### <a name="matrix-type"></a><span data-ttu-id="c2683-225">Тип "матрица"</span><span class="sxs-lookup"><span data-stu-id="c2683-225">Matrix type</span></span>

<span data-ttu-id="c2683-226">Используйте тип двумерного массива, чтобы параметр или возвращаемое значение представляли собой матрицу значений.</span><span class="sxs-lookup"><span data-stu-id="c2683-226">Use a two-dimensional array type to have the parameter or return value be a matrix of values.</span></span> <span data-ttu-id="c2683-227">Например, тип `number[][]` указывает на матрицу чисел.</span><span class="sxs-lookup"><span data-stu-id="c2683-227">For example, the type `number[][]` indicates a matrix of numbers.</span></span> <span data-ttu-id="c2683-228">`string[][]` указывает на матрицу строк.</span><span class="sxs-lookup"><span data-stu-id="c2683-228">`string[][]` indicates a matrix of strings.</span></span>

### <a name="error-type"></a><span data-ttu-id="c2683-229">Тип "ошибка"</span><span class="sxs-lookup"><span data-stu-id="c2683-229">Error type</span></span>

<span data-ttu-id="c2683-230">Функция непотоковой передачи может указывать на ошибку, возвращая тип "Ошибка".</span><span class="sxs-lookup"><span data-stu-id="c2683-230">A non-streaming function can indicate an error by returning an Error type.</span></span>

<span data-ttu-id="c2683-231">Функция потоковой передачи может указывать на ошибку, вызывая метод `setResult()` типа "Ошибка".</span><span class="sxs-lookup"><span data-stu-id="c2683-231">A streaming function can indicate an error by calling `setResult()` with an Error type.</span></span>

### <a name="promise"></a><span data-ttu-id="c2683-232">Обещание</span><span class="sxs-lookup"><span data-stu-id="c2683-232">Promise</span></span>

<span data-ttu-id="c2683-233">Функция может вернуть тип "Обещание", который задаст значение, когда обещание будет разрешено.</span><span class="sxs-lookup"><span data-stu-id="c2683-233">A function can return a Promise, which will provide the value when the promise is resolved.</span></span> <span data-ttu-id="c2683-234">В случае отклонения обещания возникнет ошибка.</span><span class="sxs-lookup"><span data-stu-id="c2683-234">If the promise is rejected, then it is an error.</span></span>

### <a name="other-types"></a><span data-ttu-id="c2683-235">Другие типы</span><span class="sxs-lookup"><span data-stu-id="c2683-235">Other types</span></span>

<span data-ttu-id="c2683-236">Любой другой тип будет рассматриваться как ошибка.</span><span class="sxs-lookup"><span data-stu-id="c2683-236">Any other type will be treated as an error.</span></span>

## <a name="next-steps"></a><span data-ttu-id="c2683-237">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="c2683-237">Next steps</span></span>
<span data-ttu-id="c2683-238">Узнайте о [соглашениях именования для пользовательских функций](custom-functions-naming.md).</span><span class="sxs-lookup"><span data-stu-id="c2683-238">Learn about [naming conventions for custom functions](custom-functions-naming.md).</span></span> <span data-ttu-id="c2683-239">Или же узнайте, как [локализовать свои функции](custom-functions-localize.md), для чего нужно [записать файл JSON вручную](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="c2683-239">Alternatively, learn how to [localize your functions](custom-functions-localize.md) which requires you to [write your JSON file by hand](custom-functions-json.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="c2683-240">Дополнительные ресурсы</span><span class="sxs-lookup"><span data-stu-id="c2683-240">See also</span></span>

* [<span data-ttu-id="c2683-241">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="c2683-241">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="c2683-242">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="c2683-242">Create custom functions in Excel</span></span>](custom-functions-overview.md)
