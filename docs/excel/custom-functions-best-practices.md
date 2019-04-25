---
ms.date: 01/08/2019
description: Ознакомьтесь с рекомендациями по разработке пользовательских функций в Excel.
title: Рекомендации в отношении пользовательских функций (предварительная версия)
localization_priority: Normal
ms.openlocfilehash: 4efcd0ba5efb0dc7450192694e8f0750de43b8a8
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32448611"
---
# <a name="custom-functions-best-practices-preview"></a><span data-ttu-id="77dec-103">Рекомендации в отношении пользовательских функций (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="77dec-103">Custom functions best practices (preview)</span></span>

<span data-ttu-id="77dec-104">В этой статье описаны рекомендации по разработке пользовательских функций в Excel.</span><span class="sxs-lookup"><span data-stu-id="77dec-104">This article describes best practices for developing custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="troubleshooting"></a><span data-ttu-id="77dec-105">Устранение неполадок</span><span class="sxs-lookup"><span data-stu-id="77dec-105">Troubleshooting</span></span>

1. <span data-ttu-id="77dec-106">Если вы проверяете надстройку в Office для Windows, нужно включить **[ведение журнала в среде выполнения](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in)**, чтобы устранять проблемы с XML-файлом манифеста надстройки, а также с некоторыми условиями установки и среды выполнения.</span><span class="sxs-lookup"><span data-stu-id="77dec-106">If you are testing your add-in in Office on Windows, you should enable **[runtime logging](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in)** to troubleshoot issues with your add-in's XML manifest file, as well as several installation and runtime conditions.</span></span> <span data-ttu-id="77dec-107">В файл журнала в среде выполнения записываются операторы `console.log`, что облегчает выявление проблем.</span><span class="sxs-lookup"><span data-stu-id="77dec-107">Runtime logging writes `console.log` statements to a log file to help you uncover issues.</span></span>

2. <span data-ttu-id="77dec-108">Ваша надстройка не будет загружена, если одна или несколько пользовательских функций конфликтуют с пользовательскими функциями, зарегистрированными в ранее зарегистрированной надстройке.</span><span class="sxs-lookup"><span data-stu-id="77dec-108">Your add-in will not load if one or more custom functions conflicts with a previously registered add-in's custom functions.</span></span> <span data-ttu-id="77dec-109">В этом случае вы можете удалить существующую надстройку или при возникновении этой ошибки при разработке надстройки, в манифесте можно указать другое имя пространства имен.</span><span class="sxs-lookup"><span data-stu-id="77dec-109">In this case, you can either remove the existing add-in, or if you encounter this error while developing an add-in, you can specify a different namespace name in your manifest.</span></span>

3. <span data-ttu-id="77dec-110">Чтобы поделиться своим мнением об этом методе устранения неполадок с группой, отвечающей за пользовательские функции Excel, отправьте отзыв группе.</span><span class="sxs-lookup"><span data-stu-id="77dec-110">To report feedback to the Excel Custom Functions team about this method of troubleshooting, send the team feedback.</span></span> <span data-ttu-id="77dec-111">Для этого выберите **Файл | Отзыв | Отправить нахмуренный смайлик**.</span><span class="sxs-lookup"><span data-stu-id="77dec-111">To do this, select **File | Feedback | Send a Frown**.</span></span> <span data-ttu-id="77dec-112">Отправка нахмуренного смайлика предоставит необходимые журналы для понимания проблемы, на которую вы указываете.</span><span class="sxs-lookup"><span data-stu-id="77dec-112">Sending a frown will provide the necessary logs to understand the issue you are hitting.</span></span>

## <a name="associating-function-names-with-json-metadata"></a><span data-ttu-id="77dec-113">Сопоставление имен функций с метаданными JSON</span><span class="sxs-lookup"><span data-stu-id="77dec-113">Associating function names with JSON metadata</span></span>

<span data-ttu-id="77dec-114">Как описано в статье [Обзор пользовательских функций](custom-functions-overview.md) проект пользовательских функций должен содержать как файл метаданных JSON, так и файл сценария (JavaScript или TypeScript) для образования готовой функции.</span><span class="sxs-lookup"><span data-stu-id="77dec-114">As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include both a JSON metadata file and a script (either JavaScript or TypeScript) file to form a complete function.</span></span> <span data-ttu-id="77dec-115">Чтобы функция работала должным образом, необходимо связать идентификатор с реализацией JavaScript.</span><span class="sxs-lookup"><span data-stu-id="77dec-115">For a function to work properly, you need to associate the id with the JavaScript implementation.</span></span> <span data-ttu-id="77dec-116">Убедитесь, что существует связь, иначе функция не будет вызываться.</span><span class="sxs-lookup"><span data-stu-id="77dec-116">Make sure there is an association, otherwise the function will not be called.</span></span>

<span data-ttu-id="77dec-117">В следующем примере показано, как выполнить данное сопоставление.</span><span class="sxs-lookup"><span data-stu-id="77dec-117">The following code sample shows how to do this association.</span></span> <span data-ttu-id="77dec-118">Пример определяет пользовательскую функцию `add` и связывает ее с объектом в файле метаданных JSON, где для свойства `id` установлено значение **ADD**.</span><span class="sxs-lookup"><span data-stu-id="77dec-118">The sample defines the custom function `add` and associates it with the object in the JSON metadata file where the value of the `id` property is **ADD**.</span></span>

```js
function add(first, second){
  return first + second;
}

CustomFunctions.associate("ADD", add);
```

<span data-ttu-id="77dec-119">Имейте в виду приведенные ниже рекомендации при создании пользовательских функций в файле JavaScript и указании соответствующих сведений в файле метаданных JSON.</span><span class="sxs-lookup"><span data-stu-id="77dec-119">Keep in mind the following best practices when creating custom functions in your JavaScript file and specifying corresponding information in the JSON metadata file.</span></span>

* <span data-ttu-id="77dec-120">Используйте только заглавные буквы для `name` и `id` функции в файле метаданных JSON.</span><span class="sxs-lookup"><span data-stu-id="77dec-120">Only use uppercase letters for a function's `name` and `id` in the JSON metadata file.</span></span> <span data-ttu-id="77dec-121">Не используйте буквы разного регистра или только строчные буквы.</span><span class="sxs-lookup"><span data-stu-id="77dec-121">Do not use a mix of cases or only lowercase letters.</span></span> <span data-ttu-id="77dec-122">Если вы сделаете это, вы можете в итоге получить два значения, которые отличаются только регистром, что будет приводить к непреднамеренной перезаписи функции.</span><span class="sxs-lookup"><span data-stu-id="77dec-122">If you do, you may end up with two values that only differ by case which will cause unintentional overwriting of your functions.</span></span> <span data-ttu-id="77dec-123">Например, объект функции со значением `id` **add** может быть перезаписан в объявлении позже в файле объекта функция со значением `id` **ADD**.</span><span class="sxs-lookup"><span data-stu-id="77dec-123">For example, a function object with an `id` value of **add** could be overwritten by declaration later in the file of function object with an `id` value of **ADD**.</span></span> <span data-ttu-id="77dec-124">Кроме того, свойство `name` определяет имя функции, которое отображается для конечных пользователей в Excel.</span><span class="sxs-lookup"><span data-stu-id="77dec-124">Additionally, the `name` property defines the function name that end users will see in Excel.</span></span> <span data-ttu-id="77dec-125">Использование прописных букв для имени каждой пользовательской функции обеспечивает единый интерфейс Excel, где все имена встроенных функций записаны прописными буквами.</span><span class="sxs-lookup"><span data-stu-id="77dec-125">Using uppercase letters for the name of each custom function provides a consistent experience in Excel, where all built-in function names are uppercase.</span></span>

* <span data-ttu-id="77dec-126">Убедитесь, что в файле метаданных JSON значение каждого свойства `id` содержит только буквы, цифры и точки.</span><span class="sxs-lookup"><span data-stu-id="77dec-126">In the JSON metadata file, ensure that the value of each `id` property contains only alphanumeric characters and periods.</span></span>

* <span data-ttu-id="77dec-127">Убедитесь, что в файле метаданных JSON значение каждого свойства `id` уникально в пределах файла.</span><span class="sxs-lookup"><span data-stu-id="77dec-127">In the JSON metadata file, ensure that the value of each `id` property is unique within the scope of the file.</span></span> <span data-ttu-id="77dec-128">То есть никакие два объекта функций в файле метаданных не должны иметь одинаковое значение `id`.</span><span class="sxs-lookup"><span data-stu-id="77dec-128">That is, no two function objects in the metadata file should have the same `id` value.</span></span> 

* <span data-ttu-id="77dec-129">Не изменяйте значение свойства `id` в файле метаданных JSON после его сопоставления с соответствующим именем функции JavaScript.</span><span class="sxs-lookup"><span data-stu-id="77dec-129">Do not change the value of an `id` property in the JSON metadata file after it's been associated with a corresponding JavaScript function name.</span></span> <span data-ttu-id="77dec-130">Вы можете изменить имя функции, которое отображается для конечных пользователей в Excel, путем обновления свойства `name` в файле метаданных JSON, но никогда не следует изменять значение свойства `id` после его установления.</span><span class="sxs-lookup"><span data-stu-id="77dec-130">You can change the function name that end users see in Excel by updating the `name` property within the JSON metadata file, but you should never change the value of an `id` property after it's been established.</span></span>

* <span data-ttu-id="77dec-131">В файле JavaScript указывайте все сопоставления пользовательских функций в одном расположении.</span><span class="sxs-lookup"><span data-stu-id="77dec-131">In the JavaScript file, specify all custom function associations in the same location.</span></span> <span data-ttu-id="77dec-132">Например, в приведенном ниже примере кода определяются две пользовательские функции и указываются сведения о сопоставлении для обеих функций.</span><span class="sxs-lookup"><span data-stu-id="77dec-132">For example, the following code sample defines two custom functions and then specifies the association information for both functions.</span></span>

    ```js
    function add(first, second){
      return first + second;
    }

    function increment(incrementBy, callback) {
      var result = 0;
      var timer = setInterval(function() {
        result += incrementBy;
        callback.setResult(result);
      }, 1000);

      callback.onCanceled = function() {
        clearInterval(timer);
      };
    }

    // associate `id` values in the JSON metadata file to JavaScript function names
    CustomFunctions.associate("ADD", add);
    CustomFunctions.associate("INCREMENT", increment);
    ```

    <span data-ttu-id="77dec-133">В приведенном ниже примере показаны метаданные JSON, соответствующие функциям, определенным в этом примере кода JavaScript.</span><span class="sxs-lookup"><span data-stu-id="77dec-133">The following sample shows the JSON metadata that corresponds to the functions defined in this JavaScript code sample.</span></span> <span data-ttu-id="77dec-134">Обратите внимание, что свойства `id` и `name` содержат заглавные буквы в этом файле.</span><span class="sxs-lookup"><span data-stu-id="77dec-134">Note that the `id` and `name` properties are in uppercase letters in this file.</span></span> 

    ```json
    {
      "$schema": "https://developer.microsoft.com/en-us/json-schemas/office-js/custom-functions.schema.json",
      "functions": [
        {
          "id": "ADD",
          "name": "ADD",
          ...
        },
        {
          "id": "INCREMENT",
          "name": "INCREMENT",
          ...
        }
      ]
    }
    ```

## <a name="declaring-optional-parameters"></a><span data-ttu-id="77dec-135">Объявление необязательных параметров</span><span class="sxs-lookup"><span data-stu-id="77dec-135">Declaring optional parameters</span></span> 

<span data-ttu-id="77dec-136">В Excel для Windows (версии 1812 или более поздней) можно объявлять необязательные параметры для пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="77dec-136">In Excel for Windows (version 1812 or later), you can declare optional parameters for your custom functions.</span></span> <span data-ttu-id="77dec-137">Если пользователь вызывает функцию в Excel, необязательные параметры отображаются в квадратных скобках.</span><span class="sxs-lookup"><span data-stu-id="77dec-137">When a user invokes a function in Excel, optional parameters appear in brackets.</span></span> <span data-ttu-id="77dec-138">Например, функция `FOO` с одним обязательным параметром `parameter1` и одним необязательным параметром `parameter2` будет отображаться в Excel как `=FOO(parameter1, [parameter2])`.</span><span class="sxs-lookup"><span data-stu-id="77dec-138">For example, a function `FOO` with one required parameter called `parameter1` and one optional parameter called `parameter2` would appear as `=FOO(parameter1, [parameter2])` in Excel.</span></span>

<span data-ttu-id="77dec-139">Чтобы сделать параметр необязательным, добавьте `"optional": true` к параметру в файле метаданных JSON, определяющем функцию.</span><span class="sxs-lookup"><span data-stu-id="77dec-139">To make a parameter optional, add `"optional": true` to the parameter in the JSON metadata file that defines the function.</span></span> <span data-ttu-id="77dec-140">В приведенном ниже примере показано, как это может выглядеть для функции `=ADD(first, second, [third])`.</span><span class="sxs-lookup"><span data-stu-id="77dec-140">The following example shows what this might look like for the function `=ADD(first, second, [third])`.</span></span> <span data-ttu-id="77dec-141">Обратите внимание, что необязательный параметр `[third]` расположен после двух обязательных параметров.</span><span class="sxs-lookup"><span data-stu-id="77dec-141">Notice that the optional `[third]` parameter follows the two required parameters.</span></span> <span data-ttu-id="77dec-142">Обязательные параметры отображаются первыми в интерфейсе формулы Excel.</span><span class="sxs-lookup"><span data-stu-id="77dec-142">Required parameters will appear first in Excel’s Formula UI.</span></span>

```json
{
    "id": "ADD",
    "name": "ADD",
    "description": "Add two numbers",
    "helpUrl": "http://www.contoso.com",
    "result": {
        "type": "number",
        "dimensionality": "scalar"
        },
    "parameters": [
        {
            "name": "first",
            "description": "first number to add",
            "type": "number",
            "dimensionality": "scalar"
        },
        {
            "name": "second",
            "description": "second number to add",
            "type": "number",
            "dimensionality": "scalar",
        },
        {
            "name": "third",
            "description": "third optional number to add",
            "type": "number",
            "dimensionality": "scalar",
            "optional": true
        }
    ],
    "options": {
        "sync": false
    }
}
```

<span data-ttu-id="77dec-143">Если вы определяете функцию, содержащую один или несколько необязательных параметров, нужно указать, что происходит, когда необязательный параметр не задан.</span><span class="sxs-lookup"><span data-stu-id="77dec-143">When you define a function that contains one or more optional parameters, you should specify what happens when the optional parameters are undefined.</span></span> <span data-ttu-id="77dec-144">В приведенном ниже примере `zipCode` и `dayOfWeek` являются необязательными параметрами для функции `getWeatherReport`.</span><span class="sxs-lookup"><span data-stu-id="77dec-144">In the following example, `zipCode` and `dayOfWeek` are both optional parameters for the `getWeatherReport` function.</span></span> <span data-ttu-id="77dec-145">Если параметр `zipCode` не определен, значение по умолчанию устанавливается равным 98052.</span><span class="sxs-lookup"><span data-stu-id="77dec-145">If the `zipCode` parameter is undefined, the default value is set to 98052.</span></span> <span data-ttu-id="77dec-146">Если параметр `dayOfWeek` не определен, ему присваивается значение Wednesday (Среда).</span><span class="sxs-lookup"><span data-stu-id="77dec-146">If the `dayOfWeek` parameter is undefined, it is set to Wednesday.</span></span>

```js
function getWeatherReport(zipCode, dayOfWeek)
{
  if (zipCode === undefined) {
      zipCode = "98052";
  }

  if (dayOfWeek === undefined) {
    dayOfWeek = "Wednesday";
  }

  // Get weather report for specified zipCode and dayOfWeek
  // ...
}
```

## <a name="additional-considerations"></a><span data-ttu-id="77dec-147">Дополнительные рекомендации</span><span class="sxs-lookup"><span data-stu-id="77dec-147">Additional considerations</span></span>

<span data-ttu-id="77dec-148">Чтобы создать надстройку, которая будет работать на различных платформах (один из основных клиентов надстроек Office), вам не следует выполнять доступ к модели DOM в пользовательских функциях или использовать библиотеки, такие как jQuery, которые используют DOM.</span><span class="sxs-lookup"><span data-stu-id="77dec-148">In order to create an add-in that will run on multiple platforms (one of the key tenants of Office Add-ins), you should not access the Document Object Model (DOM) in custom functions or use libraries like jQuery that rely on the DOM.</span></span> <span data-ttu-id="77dec-149">В Excel для Windows, где пользовательские функции используют [среду выполнения JavaScript](custom-functions-runtime.md), пользовательские функции не могут выполнять доступ к DOM.</span><span class="sxs-lookup"><span data-stu-id="77dec-149">On Excel for Windows, where custom functions use the [JavaScript runtime](custom-functions-runtime.md), custom functions cannot access the DOM.</span></span>

## <a name="see-also"></a><span data-ttu-id="77dec-150">См. также</span><span class="sxs-lookup"><span data-stu-id="77dec-150">See also</span></span>

* [<span data-ttu-id="77dec-151">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="77dec-151">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="77dec-152">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="77dec-152">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="77dec-153">Среда выполнения для пользовательских функций Excel</span><span class="sxs-lookup"><span data-stu-id="77dec-153">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="77dec-154">Журнал изменений пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="77dec-154">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="77dec-155">Руководство по настраиваемым функциям в Excel</span><span class="sxs-lookup"><span data-stu-id="77dec-155">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
