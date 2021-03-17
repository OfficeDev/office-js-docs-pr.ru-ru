---
ms.date: 03/15/2021
description: Использование тегов JSDoc для динамического создания метаданных JSON пользовательских функций.
title: Автоматическое генерирование метаданных JSON для пользовательских функций
localization_priority: Normal
ms.openlocfilehash: 344239c35e38bd88bfee5338289d1c2a929ea14c
ms.sourcegitcommit: c0c61fe84f3c5de88bd7eac29120056bb1224fc8
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2021
ms.locfileid: "50836867"
---
# <a name="autogenerate-json-metadata-for-custom-functions"></a><span data-ttu-id="53f7e-103">Автоматическое генерирование метаданных JSON для пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="53f7e-103">Autogenerate JSON metadata for custom functions</span></span>

<span data-ttu-id="53f7e-104">Если пользовательская функция Excel написана в JavaScript или TypeScript, [теги JSDoc](https://jsdoc.app/) используются для предоставления дополнительной информации о пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="53f7e-104">When an Excel custom function is written in JavaScript or TypeScript, [JSDoc tags](https://jsdoc.app/) are used to provide extra information about the custom function.</span></span> <span data-ttu-id="53f7e-105">Теги JSDoc используются при сборке для создания файла метаданных JSON.</span><span class="sxs-lookup"><span data-stu-id="53f7e-105">The JSDoc tags are then used at build time to create the JSON metadata file.</span></span> <span data-ttu-id="53f7e-106">Использование тегов JSDoc спасает вас от попытки вручную редактировать [файл метаданных JSON.](custom-functions-json.md)</span><span class="sxs-lookup"><span data-stu-id="53f7e-106">Using JSDoc tags saves you from the effort of [manually editing the JSON metadata file](custom-functions-json.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="53f7e-107">Добавьте тег `@customfunction` в примечаниях к коду для функции JavaScript или TypeScript, чтобы пометить ее как пользовательскую.</span><span class="sxs-lookup"><span data-stu-id="53f7e-107">Add the `@customfunction` tag in the code comments for a JavaScript or TypeScript function to mark it as a custom function.</span></span>

<span data-ttu-id="53f7e-108">Типы параметров функции можно получить с помощью тега [@param](#param) в JavaScript или из раздела [Тип функции](https://www.typescriptlang.org/docs/handbook/functions.html) в TypeScript.</span><span class="sxs-lookup"><span data-stu-id="53f7e-108">The function parameter types may be provided using the [@param](#param) tag in JavaScript, or from the [Function type](https://www.typescriptlang.org/docs/handbook/functions.html) in TypeScript.</span></span> <span data-ttu-id="53f7e-109">Дополнительные сведения см. в разделах, посвященных тегу [@param](#param) и [типам](#types).</span><span class="sxs-lookup"><span data-stu-id="53f7e-109">For more information, see the [@param](#param) tag and [Types](#types) sections.</span></span>

### <a name="adding-a-description-to-a-function"></a><span data-ttu-id="53f7e-110">Добавление описания функции</span><span class="sxs-lookup"><span data-stu-id="53f7e-110">Adding a description to a function</span></span>

<span data-ttu-id="53f7e-111">Описание отображается пользователю в качестве текста справки, если ему непонятно действие пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="53f7e-111">The description is displayed to the user as help text when they need help to understand what your custom function does.</span></span> <span data-ttu-id="53f7e-112">Описанию не требуется какой-либо конкретный тег.</span><span class="sxs-lookup"><span data-stu-id="53f7e-112">The description doesn't require any specific tag.</span></span> <span data-ttu-id="53f7e-113">Просто введите краткий текст описания в комментарии JSDoc.</span><span class="sxs-lookup"><span data-stu-id="53f7e-113">Just enter a short text description in the JSDoc comment.</span></span> <span data-ttu-id="53f7e-114">Обычно описание размещается в начале раздела комментариев JSDoc, но оно поддерживается независимо от места размещения.</span><span class="sxs-lookup"><span data-stu-id="53f7e-114">In general the description is placed at the start of the JSDoc comment section, but it will work no matter where it is placed.</span></span>

<span data-ttu-id="53f7e-115">Чтобы просмотреть примеры описаний встроенных функций, откройте Excel, перейдите на вкладку **Формулы** и нажмите кнопку **Вставить функцию**.</span><span class="sxs-lookup"><span data-stu-id="53f7e-115">To see examples of the built-in function descriptions, open Excel, go to the **Formulas** tab, and choose **Insert function**.</span></span> <span data-ttu-id="53f7e-116">Вы сможете просмотреть все описания функций, а также список собственных пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="53f7e-116">You can then browse through all the function descriptions, and also see your own custom functions listed.</span></span>

<span data-ttu-id="53f7e-117">В следующем примере фраза "Calculates the volume of a sphere." (Вычисляет объем сферы)</span><span class="sxs-lookup"><span data-stu-id="53f7e-117">In the following example, the phrase "Calculates the volume of a sphere."</span></span> <span data-ttu-id="53f7e-118">является описанием пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="53f7e-118">is the description for the custom function.</span></span>

```js
/**
/* Calculates the volume of a sphere.
/* @customfunction VOLUME
...
 */
```


## <a name="jsdoc-tags"></a><span data-ttu-id="53f7e-119">Теги JSDoc</span><span class="sxs-lookup"><span data-stu-id="53f7e-119">JSDoc Tags</span></span>

<span data-ttu-id="53f7e-120">Следующие теги JSDoc поддерживаются в пользовательских функциях Excel.</span><span class="sxs-lookup"><span data-stu-id="53f7e-120">The following JSDoc tags are supported in Excel custom functions.</span></span>

* [<span data-ttu-id="53f7e-121">@cancelable</span><span class="sxs-lookup"><span data-stu-id="53f7e-121">@cancelable</span></span>](#cancelable)
* <span data-ttu-id="53f7e-122">[@customfunction](#customfunction) id name</span><span class="sxs-lookup"><span data-stu-id="53f7e-122">[@customfunction](#customfunction) id name</span></span>
* <span data-ttu-id="53f7e-123">[@helpurl](#helpurl) url</span><span class="sxs-lookup"><span data-stu-id="53f7e-123">[@helpurl](#helpurl) url</span></span>
* <span data-ttu-id="53f7e-124">[@param](#param) _{type}_ name description</span><span class="sxs-lookup"><span data-stu-id="53f7e-124">[@param](#param) _{type}_ name description</span></span>
* [<span data-ttu-id="53f7e-125">@requiresAddress</span><span class="sxs-lookup"><span data-stu-id="53f7e-125">@requiresAddress</span></span>](#requiresAddress)
* [<span data-ttu-id="53f7e-126">@requiresParameterAddresses</span><span class="sxs-lookup"><span data-stu-id="53f7e-126">@requiresParameterAddresses</span></span>](#requiresParameterAddresses)
* <span data-ttu-id="53f7e-127">[@returns](#returns) _{type}_</span><span class="sxs-lookup"><span data-stu-id="53f7e-127">[@returns](#returns) _{type}_</span></span>
* [<span data-ttu-id="53f7e-128">@streaming</span><span class="sxs-lookup"><span data-stu-id="53f7e-128">@streaming</span></span>](#streaming)
* [<span data-ttu-id="53f7e-129">@volatile</span><span class="sxs-lookup"><span data-stu-id="53f7e-129">@volatile</span></span>](#volatile)

---
<a id="cancelable"></a>
### <a name="cancelable"></a><span data-ttu-id="53f7e-130">@cancelable</span><span class="sxs-lookup"><span data-stu-id="53f7e-130">@cancelable</span></span>

<span data-ttu-id="53f7e-131">Указывает, что настраиваемая функция выполняет действие при отмене функции.</span><span class="sxs-lookup"><span data-stu-id="53f7e-131">Indicates that a custom function performs an action when the function is canceled.</span></span>

<span data-ttu-id="53f7e-132">В качестве типа последнего параметра функции должно быть указано `CustomFunctions.CancelableInvocation`.</span><span class="sxs-lookup"><span data-stu-id="53f7e-132">The last function parameter must be of type `CustomFunctions.CancelableInvocation`.</span></span> <span data-ttu-id="53f7e-133">Функция может назначить свойству функцию, чтобы обозначить результат `oncanceled` при отмене функции.</span><span class="sxs-lookup"><span data-stu-id="53f7e-133">The function can assign a function to the `oncanceled` property to denote the result when the function is canceled.</span></span>

<span data-ttu-id="53f7e-134">Если тип последнего параметра функции `CustomFunctions.CancelableInvocation`, он будет рассматриваться как `@cancelable`, даже если тег отсутствует.</span><span class="sxs-lookup"><span data-stu-id="53f7e-134">If the last function parameter is of type `CustomFunctions.CancelableInvocation`, it will be considered `@cancelable` even if the tag isn't present.</span></span>

<span data-ttu-id="53f7e-135">Функция не может содержать одновременно теги `@cancelable` и `@streaming`.</span><span class="sxs-lookup"><span data-stu-id="53f7e-135">A function can't have both `@cancelable` and `@streaming` tags.</span></span>

<a id="customfunction"></a>

### <a name="customfunction"></a><span data-ttu-id="53f7e-136">@customfunction</span><span class="sxs-lookup"><span data-stu-id="53f7e-136">@customfunction</span></span>

<span data-ttu-id="53f7e-137">Синтаксис: @customfunction _id_ _name_</span><span class="sxs-lookup"><span data-stu-id="53f7e-137">Syntax: @customfunction _id_ _name_</span></span>

<span data-ttu-id="53f7e-138">Этот тег указывает, что функция JavaScript/TypeScript — это настраиваемая функция Excel.</span><span class="sxs-lookup"><span data-stu-id="53f7e-138">This tag indicates that the JavaScript/TypeScript function is an Excel custom function.</span></span> <span data-ttu-id="53f7e-139">Необходимо создать метаданные для настраиваемой функции.</span><span class="sxs-lookup"><span data-stu-id="53f7e-139">It is required to create metadata for the custom function.</span></span>

<span data-ttu-id="53f7e-140">Ниже показан пример этого тега.</span><span class="sxs-lookup"><span data-stu-id="53f7e-140">The following shows an example of this tag.</span></span>

```js
/**
 * Increments a value once a second.
 * @customfunction
 * ...
 */
```

#### <a name="id"></a><span data-ttu-id="53f7e-141">id</span><span class="sxs-lookup"><span data-stu-id="53f7e-141">id</span></span>

<span data-ttu-id="53f7e-142">Эта `id` функция определяет настраиваемую функцию.</span><span class="sxs-lookup"><span data-stu-id="53f7e-142">The `id` identifies a custom function.</span></span>

* <span data-ttu-id="53f7e-143">Если `id` не указан, название функции JavaScript или TypeScript преобразуется в верхний регистр, а недопустимые символы удаляются.</span><span class="sxs-lookup"><span data-stu-id="53f7e-143">If `id` isn't provided, the JavaScript/TypeScript function name is converted to uppercase and disallowed characters are removed.</span></span>
* <span data-ttu-id="53f7e-144">`id` должен быть уникальным для всех пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="53f7e-144">The `id` must be unique for all custom functions.</span></span>
* <span data-ttu-id="53f7e-145">Допустимые символы: A — Z, a — z, 0–9, символ подчеркивания (\_) и точка (.).</span><span class="sxs-lookup"><span data-stu-id="53f7e-145">The allowed characters are limited to: A-Z, a-z, 0-9, underscores (\_), and period (.).</span></span>

<span data-ttu-id="53f7e-146">В следующем примере increment — это параметр `id` и `name` функции.</span><span class="sxs-lookup"><span data-stu-id="53f7e-146">In the following example, increment is the `id` and the `name` of the function.</span></span>

```js
/**
 * Increments a value once a second.
 * @customfunction INCREMENT
 * ...
 */
```

#### <a name="name"></a><span data-ttu-id="53f7e-147">name</span><span class="sxs-lookup"><span data-stu-id="53f7e-147">name</span></span>

<span data-ttu-id="53f7e-148">Предоставляет отображаемый параметр `name` для пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="53f7e-148">Provides the display `name` for the custom function.</span></span>

* <span data-ttu-id="53f7e-149">Если имя не указано, идентификатор также используется как имя.</span><span class="sxs-lookup"><span data-stu-id="53f7e-149">If name isn't provided, the id is also used as the name.</span></span>
* <span data-ttu-id="53f7e-150">Допустимые символы: буквы [буквенные символы Юникод](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), числа, точки (.) и подчеркивания (\_).</span><span class="sxs-lookup"><span data-stu-id="53f7e-150">Allowed characters: Letters [Unicode Alphabetic character](https://www.unicode.org/reports/tr44/tr44-22.html#Alphabetic), numbers, period (.), and underscore (\_).</span></span>
* <span data-ttu-id="53f7e-151">Имя должно начинаться с буквы.</span><span class="sxs-lookup"><span data-stu-id="53f7e-151">Must start with a letter.</span></span>
* <span data-ttu-id="53f7e-152">Максимальная длина: 128 символов.</span><span class="sxs-lookup"><span data-stu-id="53f7e-152">Maximum length is 128 characters.</span></span>

<span data-ttu-id="53f7e-153">В следующем примере INC — это параметр `id` функции, а `increment` — параметр `name`.</span><span class="sxs-lookup"><span data-stu-id="53f7e-153">In the following example, INC is the `id` of the function and `increment` is the `name`.</span></span>

```js
/**
 * Increments a value once a second.
 * @customfunction INC INCREMENT
 * ...
 */
```

### <a name="description"></a><span data-ttu-id="53f7e-154">description</span><span class="sxs-lookup"><span data-stu-id="53f7e-154">description</span></span>

<span data-ttu-id="53f7e-155">Описание отображается пользователям в Excel при вводе функции и указывает, что делает эта функция.</span><span class="sxs-lookup"><span data-stu-id="53f7e-155">A description appears to users in Excel as they are entering the function and specifies what the function does.</span></span> <span data-ttu-id="53f7e-156">Описанию не требуется какой-либо конкретный тег.</span><span class="sxs-lookup"><span data-stu-id="53f7e-156">A description doesn't require any specific tag.</span></span> <span data-ttu-id="53f7e-157">Создайте описание для пользовательской функции, добавив в комментарии JSDoc фразу, описывающую действие функции.</span><span class="sxs-lookup"><span data-stu-id="53f7e-157">Add a description to a custom function by adding a phrase to describe what the function does inside the JSDoc comment.</span></span> <span data-ttu-id="53f7e-158">По умолчанию любой текст без тегов в разделе комментариев JSDoc является описанием функции.</span><span class="sxs-lookup"><span data-stu-id="53f7e-158">By default, whatever text is untagged in the JSDoc comment section will be the description of the function.</span></span>

<span data-ttu-id="53f7e-159">В следующем примере фраза "A function that adds two numbers" (Функция, складывающая два числа) — это описание пользовательской функции со свойством id, имеющим значение `ADD`.</span><span class="sxs-lookup"><span data-stu-id="53f7e-159">In the following example, the phrase "A function that adds two numbers" is the description for the custom function with the id property of `ADD`.</span></span>

```js
/**
 * A function that adds two numbers.
 * @customfunction ADD
 * ...
 */
```

<a id="helpurl"></a>
### <a name="helpurl"></a><span data-ttu-id="53f7e-160">@helpurl</span><span class="sxs-lookup"><span data-stu-id="53f7e-160">@helpurl</span></span>

<span data-ttu-id="53f7e-161">Синтаксис: @helpurl _url_</span><span class="sxs-lookup"><span data-stu-id="53f7e-161">Syntax: @helpurl _url_</span></span>

<span data-ttu-id="53f7e-162">Предоставленный _url_-адрес отображается в Excel.</span><span class="sxs-lookup"><span data-stu-id="53f7e-162">The provided _url_ is displayed in Excel.</span></span>

<span data-ttu-id="53f7e-163">В следующем примере `helpurl` это `www.contoso.com/weatherhelp` .</span><span class="sxs-lookup"><span data-stu-id="53f7e-163">In the following example, the `helpurl` is `www.contoso.com/weatherhelp`.</span></span>

```js
/**
 * A function which streams the temperature in a town you specify.
 * @customfunction getTemperature
 * @helpurl www.contoso.com/weatherhelp
 * ...
 */
```

<a id="param"></a>
### <a name="param"></a><span data-ttu-id="53f7e-164">@param</span><span class="sxs-lookup"><span data-stu-id="53f7e-164">@param</span></span>

#### <a name="javascript"></a><span data-ttu-id="53f7e-165">JavaScript</span><span class="sxs-lookup"><span data-stu-id="53f7e-165">JavaScript</span></span>

<span data-ttu-id="53f7e-166">Синтаксис JavaScript: @param {type} name _description_</span><span class="sxs-lookup"><span data-stu-id="53f7e-166">JavaScript Syntax: @param {type} name _description_</span></span>

* <span data-ttu-id="53f7e-167">`{type}` указывает сведения о типе в фигурных скобки.</span><span class="sxs-lookup"><span data-stu-id="53f7e-167">`{type}` specifies the type info within curly braces.</span></span> <span data-ttu-id="53f7e-168">Дополнительную информацию о типах, которые могут использоваться, см. в разделе [Типы](#types).</span><span class="sxs-lookup"><span data-stu-id="53f7e-168">See the [Types](#types) section for more information about the types which may be used.</span></span> <span data-ttu-id="53f7e-169">Если не указан тип, будет использоваться тип `any` по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="53f7e-169">If no type is specified, the default type `any` will be used.</span></span>
* <span data-ttu-id="53f7e-170">`name` указывает параметр, к @param тег.</span><span class="sxs-lookup"><span data-stu-id="53f7e-170">`name` specifies the parameter that the @param tag applies to.</span></span> <span data-ttu-id="53f7e-171">Это необходимо.</span><span class="sxs-lookup"><span data-stu-id="53f7e-171">It is required.</span></span>
* <span data-ttu-id="53f7e-172">`description` предоставляет описание, которое отображается в Excel для параметра функции.</span><span class="sxs-lookup"><span data-stu-id="53f7e-172">`description` provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="53f7e-173">Это необязательно.</span><span class="sxs-lookup"><span data-stu-id="53f7e-173">It is optional.</span></span>

<span data-ttu-id="53f7e-174">Чтобы обозначить параметр пользовательской функции как необязательный:</span><span class="sxs-lookup"><span data-stu-id="53f7e-174">To denote a custom function parameter as optional:</span></span>

* <span data-ttu-id="53f7e-175">Поместите имя параметра в квадратные скобки.</span><span class="sxs-lookup"><span data-stu-id="53f7e-175">Put square brackets around the parameter name.</span></span> <span data-ttu-id="53f7e-176">Пример: `@param {string} [text] Optional text`.</span><span class="sxs-lookup"><span data-stu-id="53f7e-176">For example: `@param {string} [text] Optional text`.</span></span>

> [!NOTE]
> <span data-ttu-id="53f7e-177">Значение по умолчанию для дополнительных параметров — `null`.</span><span class="sxs-lookup"><span data-stu-id="53f7e-177">The default value for optional parameters is `null`.</span></span>

<span data-ttu-id="53f7e-178">В следующем примере показана функция ADD, которая добавляет два или три номера, а третий номер — необязательный параметр.</span><span class="sxs-lookup"><span data-stu-id="53f7e-178">The following example shows an ADD function that adds two or three numbers, with the third number as an optional parameter.</span></span>

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

#### <a name="typescript"></a><span data-ttu-id="53f7e-179">TypeScript</span><span class="sxs-lookup"><span data-stu-id="53f7e-179">TypeScript</span></span>

<span data-ttu-id="53f7e-180">Синтаксис TypeScript: @param name _description_</span><span class="sxs-lookup"><span data-stu-id="53f7e-180">TypeScript Syntax: @param name _description_</span></span>

* <span data-ttu-id="53f7e-181">`name` указывает параметр, к @param тег.</span><span class="sxs-lookup"><span data-stu-id="53f7e-181">`name` specifies the parameter that the @param tag applies to.</span></span> <span data-ttu-id="53f7e-182">Это необходимо.</span><span class="sxs-lookup"><span data-stu-id="53f7e-182">It is required.</span></span>
* <span data-ttu-id="53f7e-183">`description` предоставляет описание, которое отображается в Excel для параметра функции.</span><span class="sxs-lookup"><span data-stu-id="53f7e-183">`description` provides the description which appears in Excel for the function parameter.</span></span> <span data-ttu-id="53f7e-184">Это необязательно.</span><span class="sxs-lookup"><span data-stu-id="53f7e-184">It is optional.</span></span>

<span data-ttu-id="53f7e-185">Дополнительные сведения о типах параметров функций, которые могут использоваться, см. в разделе [Типы](#types).</span><span class="sxs-lookup"><span data-stu-id="53f7e-185">See the [Types](#types) section for more information about the function parameter types which may be used.</span></span>

<span data-ttu-id="53f7e-186">Чтобы обозначить параметр пользовательской функции как необязательный, выполните одно из указанных ниже действий.</span><span class="sxs-lookup"><span data-stu-id="53f7e-186">To denote a custom function parameter as optional, do one of the following:</span></span>

* <span data-ttu-id="53f7e-187">Используйте необязательный параметр.</span><span class="sxs-lookup"><span data-stu-id="53f7e-187">Use an optional parameter.</span></span> <span data-ttu-id="53f7e-188">Пример: `function f(text?: string)`</span><span class="sxs-lookup"><span data-stu-id="53f7e-188">For example: `function f(text?: string)`</span></span>
* <span data-ttu-id="53f7e-189">Задайте для параметра значение по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="53f7e-189">Give the parameter a default value.</span></span> <span data-ttu-id="53f7e-190">Пример: `function f(text: string = "abc")`</span><span class="sxs-lookup"><span data-stu-id="53f7e-190">For example: `function f(text: string = "abc")`</span></span>

<span data-ttu-id="53f7e-191">Подробное описание @param см. в [JSDoc](https://jsdoc.app/tags-param.html)</span><span class="sxs-lookup"><span data-stu-id="53f7e-191">For detailed description of the @param see: [JSDoc](https://jsdoc.app/tags-param.html)</span></span>

> [!NOTE]
> <span data-ttu-id="53f7e-192">Значение по умолчанию для дополнительных параметров — `null`.</span><span class="sxs-lookup"><span data-stu-id="53f7e-192">The default value for optional parameters is `null`.</span></span>

<span data-ttu-id="53f7e-193">В следующем примере показана функция `add`, складывающая два числа.</span><span class="sxs-lookup"><span data-stu-id="53f7e-193">The following example shows the `add` function that adds two numbers.</span></span>

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

<a id="requiresAddress"></a>

### <a name="requiresaddress"></a><span data-ttu-id="53f7e-194">@requiresAddress</span><span class="sxs-lookup"><span data-stu-id="53f7e-194">@requiresAddress</span></span>

<span data-ttu-id="53f7e-195">Указывает, что следует предоставлять адрес ячейки, в которой вычисляется функция.</span><span class="sxs-lookup"><span data-stu-id="53f7e-195">Indicates that the address of the cell where the function is being evaluated should be provided.</span></span>

<span data-ttu-id="53f7e-196">Последний параметр функции должен быть типом или производным типом `CustomFunctions.Invocation` для `@requiresAddress` использования.</span><span class="sxs-lookup"><span data-stu-id="53f7e-196">The last function parameter must be of type `CustomFunctions.Invocation` or a derived type to use `@requiresAddress`.</span></span> <span data-ttu-id="53f7e-197">При вызове функции свойство `address` будет содержать адрес.</span><span class="sxs-lookup"><span data-stu-id="53f7e-197">When the function is called, the `address` property will contain the address.</span></span>

<span data-ttu-id="53f7e-198">В следующем примере показано, как использовать параметр в сочетании с возвращением адреса ячейки, вызываемой `invocation` `@requiresAddress` вашей настраиваемой функцией.</span><span class="sxs-lookup"><span data-stu-id="53f7e-198">The following sample shows how to use the `invocation` parameter in combination with `@requiresAddress` to return the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="53f7e-199">Дополнительные [сведения см. в параметре Вызов.](custom-functions-parameter-options.md#invocation-parameter)</span><span class="sxs-lookup"><span data-stu-id="53f7e-199">See [Invocation parameter](custom-functions-parameter-options.md#invocation-parameter) for more information.</span></span>

```js
/**
 * Return the address of the cell that invoked the custom function. 
 * @customfunction
 * @param {number} first First parameter.
 * @param {number} second Second parameter.
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @requiresAddress 
 */
function getAddress(first, second, invocation) {
  var address = invocation.address;
  return address;
}
```

<a id="requiresParameterAddresses"></a>
### <a name="requiresparameteraddresses"></a><span data-ttu-id="53f7e-200">@requiresParameterAddresses</span><span class="sxs-lookup"><span data-stu-id="53f7e-200">@requiresParameterAddresses</span></span>

<span data-ttu-id="53f7e-201">Указывает, что функция должна возвращать адреса параметров ввода.</span><span class="sxs-lookup"><span data-stu-id="53f7e-201">Indicates that the function should return the addresses of input parameters.</span></span> 

<span data-ttu-id="53f7e-202">Последний параметр функции должен быть типом или производным типом `CustomFunctions.Invocation` для  `@requiresParameterAddresses` использования.</span><span class="sxs-lookup"><span data-stu-id="53f7e-202">The last function parameter must be of type `CustomFunctions.Invocation` or a derived type to use  `@requiresParameterAddresses`.</span></span> <span data-ttu-id="53f7e-203">Комментарий JSDoc также должен включать тег, указывающий, что возвращаемая величина — это `@returns` матрица, например `@returns {string[][]}` или `@returns {number[][]}` .</span><span class="sxs-lookup"><span data-stu-id="53f7e-203">The JSDoc comment must also include an `@returns` tag specifying that the return value be a matrix, such as `@returns {string[][]}` or `@returns {number[][]}`.</span></span> <span data-ttu-id="53f7e-204">Дополнительные [сведения см.](/office/dev/add-ins/excel/custom-functions-json-autogeneration#matrix-type) в матричных типах.</span><span class="sxs-lookup"><span data-stu-id="53f7e-204">See [Matrix types](/office/dev/add-ins/excel/custom-functions-json-autogeneration#matrix-type) for additional information.</span></span> 

<span data-ttu-id="53f7e-205">Когда функция называется, `parameterAddresses` свойство будет содержать адреса параметров ввода.</span><span class="sxs-lookup"><span data-stu-id="53f7e-205">When the function is called, the `parameterAddresses` property will contain the addresses of the input parameters.</span></span>

<span data-ttu-id="53f7e-206">В следующем примере показано, как использовать параметр в сочетании с возвращением `invocation` `@requiresParameterAddresses` адресов трех параметров ввода.</span><span class="sxs-lookup"><span data-stu-id="53f7e-206">The following sample shows how to use the `invocation` parameter in combination with `@requiresParameterAddresses` to return the addresses of three input parameters.</span></span> <span data-ttu-id="53f7e-207">Дополнительные [сведения см. в](/office/dev/add-ins/excel/custom-functions-parameter-options#detect-the-address-of-a-parameter) дополнительных сведениях Об обнаружении адреса параметра.</span><span class="sxs-lookup"><span data-stu-id="53f7e-207">See [Detect the address of a parameter](/office/dev/add-ins/excel/custom-functions-parameter-options#detect-the-address-of-a-parameter) for more information.</span></span> 

```js
/**
 * Return the addresses of three parameters. 
 * @customfunction
 * @param {string} firstParameter First parameter.
 * @param {string} secondParameter Second parameter.
 * @param {string} thirdParameter Third parameter.
 * @param {CustomFunctions.Invocation} invocation Invocation object. 
 * @returns {string[][]} The addresses of the parameters, as a 2-dimensional array.
 * @requiresParameterAddresses
 */
function getParameterAddresses(firstParameter, secondParameter, thirdParameter, invocation) {
  var addresses = [
    [invocation.parameterAddresses[0]],
    [invocation.parameterAddresses[1]],
    [invocation.parameterAddresses[2]]
  ];
  return addresses;
}
```

<a id="returns"></a>
### <a name="returns"></a><span data-ttu-id="53f7e-208">@returns</span><span class="sxs-lookup"><span data-stu-id="53f7e-208">@returns</span></span>

<span data-ttu-id="53f7e-209">Синтаксис: @returns {_type_}</span><span class="sxs-lookup"><span data-stu-id="53f7e-209">Syntax: @returns {_type_}</span></span>

<span data-ttu-id="53f7e-210">Предоставляет тип для возвращаемого значения.</span><span class="sxs-lookup"><span data-stu-id="53f7e-210">Provides the type for the return value.</span></span>

<span data-ttu-id="53f7e-211">Если `{type}` не указан, будет использоваться информация о типе TypeScript.</span><span class="sxs-lookup"><span data-stu-id="53f7e-211">If `{type}` is omitted, the TypeScript type info will be used.</span></span> <span data-ttu-id="53f7e-212">Если информация о типе отсутствует, будет использоваться тип `any`.</span><span class="sxs-lookup"><span data-stu-id="53f7e-212">If there is no type info, the type will be `any`.</span></span>

<span data-ttu-id="53f7e-213">В следующем примере показана функция `add`, использующая тег `@returns`.</span><span class="sxs-lookup"><span data-stu-id="53f7e-213">The following example shows the `add` function that uses the `@returns` tag.</span></span>

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

<a id="streaming"></a>
### <a name="streaming"></a><span data-ttu-id="53f7e-214">@streaming</span><span class="sxs-lookup"><span data-stu-id="53f7e-214">@streaming</span></span>

<span data-ttu-id="53f7e-215">Используется для обозначения того, что пользовательская функция является потоковой передачей функции.</span><span class="sxs-lookup"><span data-stu-id="53f7e-215">Used to indicate that a custom function is a streaming function.</span></span> 

<span data-ttu-id="53f7e-216">Последний параметр — тип `CustomFunctions.StreamingInvocation<ResultType>` .</span><span class="sxs-lookup"><span data-stu-id="53f7e-216">The last parameter is of type `CustomFunctions.StreamingInvocation<ResultType>`.</span></span>
<span data-ttu-id="53f7e-217">Функция `void` возвращается.</span><span class="sxs-lookup"><span data-stu-id="53f7e-217">The function returns `void`.</span></span>

<span data-ttu-id="53f7e-218">Функции потоковой передачи не возвращают значения напрямую, а звонят `setResult(result: ResultType)` с помощью последнего параметра.</span><span class="sxs-lookup"><span data-stu-id="53f7e-218">Streaming functions don't return values directly, instead they call `setResult(result: ResultType)` using the last parameter.</span></span>

<span data-ttu-id="53f7e-219">Исключения, которые возникают при потоковой передаче функций, игнорируются.</span><span class="sxs-lookup"><span data-stu-id="53f7e-219">Exceptions thrown by a streaming function are ignored.</span></span> <span data-ttu-id="53f7e-220">`setResult()` при вызове может вернуть ошибку в качестве результата.</span><span class="sxs-lookup"><span data-stu-id="53f7e-220">`setResult()` may be called with Error to indicate an error result.</span></span> <span data-ttu-id="53f7e-221">Пример функции потоковой передачи и дополнительные сведения см. в разделе [Создание функции потоковой передачи](custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="53f7e-221">For an example of a streaming function and more information, see [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function).</span></span>

<span data-ttu-id="53f7e-222">Потоковые передачи функций невозможно пометить как [@volatile](#volatile).</span><span class="sxs-lookup"><span data-stu-id="53f7e-222">Streaming functions can't be marked as [@volatile](#volatile).</span></span>

<a id="volatile"></a>
### <a name="volatile"></a><span data-ttu-id="53f7e-223">@volatile</span><span class="sxs-lookup"><span data-stu-id="53f7e-223">@volatile</span></span>

<span data-ttu-id="53f7e-224">Переменные функции — это такие функции, чей результат не остается неизменным в каждый период времени, даже если они не содержат аргументов или их аргументы не меняются.</span><span class="sxs-lookup"><span data-stu-id="53f7e-224">A volatile function is one whose result isn't the same from one moment to the next, even if it takes no arguments or the arguments haven't changed.</span></span> <span data-ttu-id="53f7e-225">Excel повторно проводит вычисления в ячейках, которые содержат переменные функции, вместе со всеми зависимыми функциями при каждом вычислении.</span><span class="sxs-lookup"><span data-stu-id="53f7e-225">Excel re-evaluates cells that contain volatile functions, together with all dependents, every time that a calculation is done.</span></span> <span data-ttu-id="53f7e-226">По этой причине чрезмерное использование переменных функций может замедлить пересчет, поэтому используйте их умеренно.</span><span class="sxs-lookup"><span data-stu-id="53f7e-226">For this reason, too much reliance on volatile functions can make recalculation times slow, so use them sparingly.</span></span>

<span data-ttu-id="53f7e-227">Потоковые передачи функций не могут быть переменными.</span><span class="sxs-lookup"><span data-stu-id="53f7e-227">Streaming functions can't be volatile.</span></span>

<span data-ttu-id="53f7e-228">Следующая функция является переменной и использует тег `@volatile`.</span><span class="sxs-lookup"><span data-stu-id="53f7e-228">The following function is volatile and uses the `@volatile` tag.</span></span>

```js
/**
 * Simulates rolling a 6-sided die.
 * @customfunction
 * @volatile
 */
function roll6sided(): number {
  return Math.floor(Math.random() * 6) + 1;
}
```

---

## <a name="types"></a><span data-ttu-id="53f7e-229">Типы</span><span class="sxs-lookup"><span data-stu-id="53f7e-229">Types</span></span>

<span data-ttu-id="53f7e-230">Указывая тип параметра, Excel преобразует значения в этот тип, прежде чем вызывать функцию.</span><span class="sxs-lookup"><span data-stu-id="53f7e-230">By specifying a parameter type, Excel will convert values into that type before calling the function.</span></span> <span data-ttu-id="53f7e-231">Если указан тип `any`, преобразование выполняться не будет.</span><span class="sxs-lookup"><span data-stu-id="53f7e-231">If the type is `any`, no conversion will be performed.</span></span>

### <a name="value-types"></a><span data-ttu-id="53f7e-232">Типы значений</span><span class="sxs-lookup"><span data-stu-id="53f7e-232">Value types</span></span>

<span data-ttu-id="53f7e-233">Одно значение может быть представлено с помощью одного из приведенных ниже типов: `boolean`, `number`, `string`.</span><span class="sxs-lookup"><span data-stu-id="53f7e-233">A single value may be represented using one of the following types: `boolean`, `number`, `string`.</span></span>

### <a name="matrix-type"></a><span data-ttu-id="53f7e-234">Тип "матрица"</span><span class="sxs-lookup"><span data-stu-id="53f7e-234">Matrix type</span></span>

<span data-ttu-id="53f7e-235">Используйте тип двумерного массива, чтобы параметр или возвращаемое значение представляли собой матрицу значений.</span><span class="sxs-lookup"><span data-stu-id="53f7e-235">Use a two-dimensional array type to have the parameter or return value be a matrix of values.</span></span> <span data-ttu-id="53f7e-236">Например, тип указывает матрицу чисел и указывает `number[][]` `string[][]` матрицу строк.</span><span class="sxs-lookup"><span data-stu-id="53f7e-236">For example, the type `number[][]` indicates a matrix of numbers and `string[][]` indicates a matrix of strings.</span></span>

### <a name="error-type"></a><span data-ttu-id="53f7e-237">Тип "ошибка"</span><span class="sxs-lookup"><span data-stu-id="53f7e-237">Error type</span></span>

<span data-ttu-id="53f7e-238">Функция непотоковой передачи может указывать на ошибку, возвращая тип "Ошибка".</span><span class="sxs-lookup"><span data-stu-id="53f7e-238">A non-streaming function can indicate an error by returning an Error type.</span></span>

<span data-ttu-id="53f7e-239">Функция потоковой передачи может указывать на ошибку, вызывая метод `setResult()` типа "Ошибка".</span><span class="sxs-lookup"><span data-stu-id="53f7e-239">A streaming function can indicate an error by calling `setResult()` with an Error type.</span></span>

### <a name="promise"></a><span data-ttu-id="53f7e-240">Обещание</span><span class="sxs-lookup"><span data-stu-id="53f7e-240">Promise</span></span>

<span data-ttu-id="53f7e-241">Настраиваемая функция может вернуть обещание, которое предоставляет значение при его ок.</span><span class="sxs-lookup"><span data-stu-id="53f7e-241">A custom function can return a promise that provides the value when the promise is resolved.</span></span> <span data-ttu-id="53f7e-242">Если обещание отклоняется, то настраиваемая функция будет бросать ошибку.</span><span class="sxs-lookup"><span data-stu-id="53f7e-242">If the promise is rejected, then the custom function will throw an error.</span></span>

### <a name="other-types"></a><span data-ttu-id="53f7e-243">Другие типы</span><span class="sxs-lookup"><span data-stu-id="53f7e-243">Other types</span></span>

<span data-ttu-id="53f7e-244">Любой другой тип будет рассматриваться как ошибка.</span><span class="sxs-lookup"><span data-stu-id="53f7e-244">Any other type will be treated as an error.</span></span>

## <a name="next-steps"></a><span data-ttu-id="53f7e-245">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="53f7e-245">Next steps</span></span>

<span data-ttu-id="53f7e-246">Узнайте о [соглашениях именования для пользовательских функций](custom-functions-naming.md).</span><span class="sxs-lookup"><span data-stu-id="53f7e-246">Learn about [naming conventions for custom functions](custom-functions-naming.md).</span></span> <span data-ttu-id="53f7e-247">Или же узнайте, как [локализовать свои функции](custom-functions-localize.md), для чего нужно [записать файл JSON вручную](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="53f7e-247">Alternatively, learn how to [localize your functions](custom-functions-localize.md) which requires you to [write your JSON file by hand](custom-functions-json.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="53f7e-248">Дополнительные ресурсы</span><span class="sxs-lookup"><span data-stu-id="53f7e-248">See also</span></span>

* [<span data-ttu-id="53f7e-249">Вручную создайте метаданные JSON для пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="53f7e-249">Manually create JSON metadata for custom functions</span></span>](custom-functions-json.md)
* [<span data-ttu-id="53f7e-250">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="53f7e-250">Create custom functions in Excel</span></span>](custom-functions-overview.md)
