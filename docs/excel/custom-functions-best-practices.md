---
ms.date: 01/08/2019
description: Ознакомьтесь с рекомендациями по разработке пользовательских функций в Excel.
title: Рекомендации в отношении пользовательских функций (предварительная версия)
ms.openlocfilehash: 45618a61d0d1fdd0398ecec3aa0db21e493787fd
ms.sourcegitcommit: 9afcb1bb295ec0c8940ed3a8364dbac08ef6b382
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/08/2019
ms.locfileid: "27770653"
---
# <a name="custom-functions-best-practices-preview"></a><span data-ttu-id="015c0-103">Рекомендации в отношении пользовательских функций (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="015c0-103">Custom functions best practices (preview)</span></span>

<span data-ttu-id="015c0-104">В этой статье описаны рекомендации по разработке пользовательских функций в Excel.</span><span class="sxs-lookup"><span data-stu-id="015c0-104">This article describes best practices for developing custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="error-handling"></a><span data-ttu-id="015c0-105">Обработка ошибок</span><span class="sxs-lookup"><span data-stu-id="015c0-105">Error handling</span></span>

<span data-ttu-id="015c0-106">При создании надстройки, определяющей пользовательские функции, не забудьте включить логику для обработки ошибок, возникающих в среде выполнения.</span><span class="sxs-lookup"><span data-stu-id="015c0-106">When you build an add-in that defines custom functions, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="015c0-107">Обработка ошибок для пользовательских функций в значительной степени совпадает с [обработкой ошибок для API JavaScript в Excel](excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="015c0-107">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="015c0-108">В приведенном ниже примере кода `.catch` обрабатывает любые ошибки, возникающие в коде ранее.</span><span class="sxs-lookup"><span data-stu-id="015c0-108">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

```js
function getComment(x) {
  let url = "https://www.contoso.com/comments/" + x;
  return fetch(url)
    .then(function (data) {
      return data.json();
    })
    .then((json) => {
      return json.body;
    })
    .catch(function (error) {
      throw error;
    })
}
```

## <a name="troubleshooting"></a><span data-ttu-id="015c0-109">Устранение неполадок</span><span class="sxs-lookup"><span data-stu-id="015c0-109">Troubleshooting</span></span>

<span data-ttu-id="015c0-110">Если вы проверяете надстройку в Office для Windows, нужно включить **[ведение журнала в среде выполнения](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in)**, чтобы устранять проблемы с XML-файлом манифеста надстройки, а также с некоторыми условиями установки и среды выполнения.</span><span class="sxs-lookup"><span data-stu-id="015c0-110">If you are testing your add-in in Office on Windows, you should enable **[runtime logging](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in)** to troubleshoot issues with your add-in's XML manifest file, as well as several installation and runtime conditions.</span></span> <span data-ttu-id="015c0-111">В файл журнала в среде выполнения записываются операторы `console.log`, что облегчает выявление проблем.</span><span class="sxs-lookup"><span data-stu-id="015c0-111">Runtime logging writes `console.log` statements to a log file to help you uncover issues.</span></span>

<span data-ttu-id="015c0-112">Чтобы поделиться своим мнением об этом методе устранения неполадок с группой, отвечающей за пользовательские функции Excel, отправьте отзыв группе.</span><span class="sxs-lookup"><span data-stu-id="015c0-112">To report feedback to the Excel Custom Functions team about this method of troubleshooting, send the team feedback.</span></span> <span data-ttu-id="015c0-113">Для этого выберите **Файл | Отзыв | Отправить нахмуренный смайлик**.</span><span class="sxs-lookup"><span data-stu-id="015c0-113">To do this, select **File | Feedback | Send a Frown**.</span></span> <span data-ttu-id="015c0-114">Отправка нахмуренного смайлика предоставит необходимые журналы для понимания проблемы, на которую вы указываете.</span><span class="sxs-lookup"><span data-stu-id="015c0-114">Sending a frown will provide the necessary logs to understand the issue you are hitting.</span></span>

## <a name="debugging"></a><span data-ttu-id="015c0-115">Отладка</span><span class="sxs-lookup"><span data-stu-id="015c0-115">Debugging</span></span>

<span data-ttu-id="015c0-116">В настоящее время наилучшим способом отладки пользовательских функций Excel является [загрузка неопубликованной](../testing/sideload-office-add-ins-for-testing.md) надстройки в **Excel Online**.</span><span class="sxs-lookup"><span data-stu-id="015c0-116">Currently, the best method for debugging Excel custom functions is to first [sideload](../testing/sideload-office-add-ins-for-testing.md) your add-in within **Excel Online**.</span></span> <span data-ttu-id="015c0-117">Затем вы сможете выполнить отладку пользовательских функций с помощью [собственного средства отладки вашего браузера, вызываемого клавишей F12,](../testing/debug-add-ins-in-office-online.md) в сочетании с указанными ниже методами.</span><span class="sxs-lookup"><span data-stu-id="015c0-117">You can then debug your custom functions by using the [F12 debugging tool native to your browser](../testing/debug-add-ins-in-office-online.md) in combination with the following techniques:</span></span>

- <span data-ttu-id="015c0-118">Используйте операторы `console.log` в коде пользовательских функций, чтобы отправлять результаты в консоль в режиме реального времени.</span><span class="sxs-lookup"><span data-stu-id="015c0-118">Use `console.log` statements within your custom functions code to send output to the console in real time.</span></span>

- <span data-ttu-id="015c0-119">Используйте операторы `debugger;` в коде пользовательских функций, чтобы указать точки останова для приостановки выполнения при открытом окне, вызываемом клавишей F12.</span><span class="sxs-lookup"><span data-stu-id="015c0-119">Use `debugger;` statements within your custom functions code to specify breakpoints where execution will pause when the F12 window is open.</span></span> <span data-ttu-id="015c0-120">Например, если выполняется следующая функция при открытом окне F12, выполнение приостанавливается на операторе `debugger;`, что позволяет вручную проверить значения параметров перед возвратом функции.</span><span class="sxs-lookup"><span data-stu-id="015c0-120">For example, if the following function runs while the F12 window is open, execution will pause on the `debugger;` statement, enabling you to manually inspect parameter values before the function returns.</span></span> <span data-ttu-id="015c0-121">Оператор `debugger;` не выполняет никаких действий в Excel Online, если не открыто окно F12.</span><span class="sxs-lookup"><span data-stu-id="015c0-121">The `debugger;` statement has no effect in Excel Online when the F12 window is not open.</span></span> <span data-ttu-id="015c0-122">В настоящее время оператор `debugger;` не выполняет никаких действий в Excel для Windows.</span><span class="sxs-lookup"><span data-stu-id="015c0-122">Currently, the `debugger;` statement has no effect in Excel for Windows.</span></span>

    ```js
    function add(first, second){
      debugger;
      return first + second;
    }
    ```

<span data-ttu-id="015c0-123">Если не удается зарегистрировать надстройку, [проверьте, что SSL-сертификаты правильно настроены](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) для веб-сервера, на котором размещается ваше приложение надстройки.</span><span class="sxs-lookup"><span data-stu-id="015c0-123">If your add-in fails to register, [verify that SSL certificates are correctly configured](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) for the web server that's hosting your add-in application.</span></span>

## <a name="associating-function-names-with-json-metadata"></a><span data-ttu-id="015c0-124">Сопоставление имен функций с метаданными JSON</span><span class="sxs-lookup"><span data-stu-id="015c0-124">Associating function names with JSON metadata</span></span>

<span data-ttu-id="015c0-125">Как описано в статье [Обзор пользовательских функций](custom-functions-overview.md) проект пользовательских функций должен содержать как файл метаданных JSON, так и файл сценария (JavaScript или TypeScript) для образования готовой функции.</span><span class="sxs-lookup"><span data-stu-id="015c0-125">As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include both a JSON metadata file and a script (either JavaScript or TypeScript) file to form a complete function.</span></span> <span data-ttu-id="015c0-126">Для корректной работы функции вам потребуется связать имя функции в файле сценария с идентификатором, указанным в файле JSON.</span><span class="sxs-lookup"><span data-stu-id="015c0-126">For a function to work properly, you'll need to bind the name of the function in the script file to the id listed in the JSON file.</span></span> <span data-ttu-id="015c0-127">Данный процесс называется сопоставлением.</span><span class="sxs-lookup"><span data-stu-id="015c0-127">This process is called association.</span></span> <span data-ttu-id="015c0-128">Сделайте заметку о необходимости включения сопоставления в конце файлов кода JavaScript, иначе функция не будет работать.</span><span class="sxs-lookup"><span data-stu-id="015c0-128">Make a note to include associations at the end of your JavaScript code files; otherwise, your functions will not work.</span></span>

<span data-ttu-id="015c0-129">В следующем примере показано, как выполнить данное сопоставление.</span><span class="sxs-lookup"><span data-stu-id="015c0-129">The following multi-part POST code shows how to do this.</span></span> <span data-ttu-id="015c0-130">Пример определяет пользовательскую функцию `add` и связывает ее с объектом в файле метаданных JSON, где для свойства `id` установлено значение **ADD**.</span><span class="sxs-lookup"><span data-stu-id="015c0-130">For example, the following code sample defines the custom function  and then specifies that the function  corresponds to the object in the JSON metadata file where the value of the  property is ADD.</span></span>

```js
function add(first, second){
  return first + second;
}

CustomFunctions.associate("ADD", add); 
```

<span data-ttu-id="015c0-131">Имейте в виду приведенные ниже рекомендации при создании пользовательских функций в файле JavaScript и указании соответствующих сведений в файле метаданных JSON.</span><span class="sxs-lookup"><span data-stu-id="015c0-131">Keep in mind the following best practices when creating custom functions in your JavaScript file and specifying corresponding information in the JSON metadata file.</span></span>

* <span data-ttu-id="015c0-132">Используйте только заглавные буквы для `name` и `id` функции в файле метаданных JSON.</span><span class="sxs-lookup"><span data-stu-id="015c0-132">Only use uppercase letters for a function's `name` and `id` in the JSON metadata file.</span></span> <span data-ttu-id="015c0-133">Не используйте буквы разного регистра или только строчные буквы.</span><span class="sxs-lookup"><span data-stu-id="015c0-133">Do not use a mix of cases or only lowercase letters.</span></span> <span data-ttu-id="015c0-134">Если вы сделаете это, вы можете в итоге получить два значения, которые отличаются только регистром, что будет приводить к непреднамеренной перезаписи функции.</span><span class="sxs-lookup"><span data-stu-id="015c0-134">If you do, you may end up with two values that only differ by case which will cause unintentional overwriting of your functions.</span></span> <span data-ttu-id="015c0-135">Например, объект функции со значением `id` **add** может быть перезаписан в объявлении позже в файле объекта функция со значением `id` **ADD**.</span><span class="sxs-lookup"><span data-stu-id="015c0-135">For example, a function object with an `id` value of **add** could be overwritten by declaration later in the file of function object with an `id` value of **ADD**.</span></span> <span data-ttu-id="015c0-136">Кроме того, свойство `name` определяет имя функции, которое отображается для конечных пользователей в Excel.</span><span class="sxs-lookup"><span data-stu-id="015c0-136">The `name` property defines the function name that end users will see in Excel.</span></span> <span data-ttu-id="015c0-137">Использование прописных букв для имени каждой пользовательской функции обеспечивает единый интерфейс Excel, где все имена встроенных функций записаны прописными буквами.</span><span class="sxs-lookup"><span data-stu-id="015c0-137">Using uppercase letters for the name of each custom function provides a consistent experience for end users in Excel, where all built-in function names are uppercase.</span></span>

* <span data-ttu-id="015c0-138">Тем не менее, нет необходимости использовать заглавную букву для `name` функции при сопоставлении.</span><span class="sxs-lookup"><span data-stu-id="015c0-138">However, it is not necessary to capitalize the function's `name` when associating.</span></span> <span data-ttu-id="015c0-139">Например, `CustomFunctions.associate("add", add)` является эквивалентом `CustomFunctions.associate("ADD", add)`.</span><span class="sxs-lookup"><span data-stu-id="015c0-139">For example, `CustomFunctions.associate("add", add)` is equivalent to `CustomFunctions.associate("ADD", add)`.</span></span>

* <span data-ttu-id="015c0-140">Убедитесь, что в файле метаданных JSON значение каждого свойства `id` содержит только буквы, цифры и точки.</span><span class="sxs-lookup"><span data-stu-id="015c0-140">In the JSON metadata file, ensure that the value of each `id` property contains only alphanumeric characters and periods.</span></span>

* <span data-ttu-id="015c0-141">Убедитесь, что в файле метаданных JSON значение каждого свойства `id` уникально в пределах файла.</span><span class="sxs-lookup"><span data-stu-id="015c0-141">In the JSON metadata file, ensure that the value of each `id` property is unique within the scope of the file.</span></span> <span data-ttu-id="015c0-142">То есть никакие два объекта функций в файле метаданных не должны иметь одинаковое значение `id`.</span><span class="sxs-lookup"><span data-stu-id="015c0-142">That is, no two function objects in the metadata file should have the same `id` value.</span></span> 

* <span data-ttu-id="015c0-143">Не изменяйте значение свойства `id` в файле метаданных JSON после его сопоставления с соответствующим именем функции JavaScript.</span><span class="sxs-lookup"><span data-stu-id="015c0-143">Do not change the value of an `id` property in the JSON metadata file after it's been mapped to a corresponding JavaScript function name.</span></span> <span data-ttu-id="015c0-144">Вы можете изменить имя функции, которое отображается для конечных пользователей в Excel, путем обновления свойства `name` в файле метаданных JSON, но никогда не следует изменять значение свойства `id` после его установления.</span><span class="sxs-lookup"><span data-stu-id="015c0-144">You can change the function name that end users see in Excel by updating the `name` property within the JSON metadata file, but you should never change the value of an `id` property after it's been established.</span></span>

* <span data-ttu-id="015c0-145">В файле JavaScript указывайте все сопоставления пользовательских функций в одном расположении.</span><span class="sxs-lookup"><span data-stu-id="015c0-145">In the JavaScript file, specify all custom function mappings in the same location.</span></span> <span data-ttu-id="015c0-146">Например, в приведенном ниже примере кода определяются две пользовательские функции и указываются сведения о сопоставлении для обеих функций.</span><span class="sxs-lookup"><span data-stu-id="015c0-146">For example, the following code sample defines two custom functions and then specifies the mapping information for both functions.</span></span>

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

    <span data-ttu-id="015c0-147">В приведенном ниже примере показаны метаданные JSON, соответствующие функциям, определенным в этом примере кода JavaScript.</span><span class="sxs-lookup"><span data-stu-id="015c0-147">The following sample shows the JSON metadata that corresponds to the functions defined in this JavaScript code sample.</span></span> <span data-ttu-id="015c0-148">Обратите внимание, что свойства `id` и `name` содержат заглавные буквы в этом файле.</span><span class="sxs-lookup"><span data-stu-id="015c0-148">Note that the `id` and `name` properties are in uppercase letters in this file.</span></span> 

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

## <a name="declaring-optional-parameters"></a><span data-ttu-id="015c0-149">Объявление необязательных параметров</span><span class="sxs-lookup"><span data-stu-id="015c0-149">Declaring optional parameters</span></span> 
<span data-ttu-id="015c0-150">В Excel для Windows (версии 1812 или более поздней) можно объявлять необязательные параметры для пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="015c0-150">In Excel for Windows (version 1812 or later), you can declare optional parameters for your custom functions.</span></span> <span data-ttu-id="015c0-151">Если пользователь вызывает функцию в Excel, необязательные параметры отображаются в квадратных скобках.</span><span class="sxs-lookup"><span data-stu-id="015c0-151">When a user invokes a function in Excel, optional parameters appear in brackets.</span></span> <span data-ttu-id="015c0-152">Например, функция `FOO` с одним обязательным параметром `parameter1` и одним необязательным параметром `parameter2` будет отображаться в Excel как `=FOO(parameter1, [parameter2])`.</span><span class="sxs-lookup"><span data-stu-id="015c0-152">For example, a function `FOO` with one required parameter called `parameter1` and one optional parameter called `parameter2` would appear as `=FOO(parameter1, [parameter2])` in Excel.</span></span>

<span data-ttu-id="015c0-153">Чтобы сделать параметр необязательным, добавьте `"optional": true` к параметру в файле метаданных JSON, определяющем функцию.</span><span class="sxs-lookup"><span data-stu-id="015c0-153">To make a parameter optional, add `"optional": true` to the parameter in the JSON metadata file that defines the function.</span></span> <span data-ttu-id="015c0-154">В приведенном ниже примере показано, как это может выглядеть для функции `=ADD(first, second, [third])`.</span><span class="sxs-lookup"><span data-stu-id="015c0-154">The following example shows what this might look like for the function `=ADD(first, second, [third])`.</span></span> <span data-ttu-id="015c0-155">Обратите внимание, что необязательный параметр `[third]` расположен после двух обязательных параметров.</span><span class="sxs-lookup"><span data-stu-id="015c0-155">Notice that the optional `[third]` parameter follows the two required parameters.</span></span> <span data-ttu-id="015c0-156">Обязательные параметры отображаются первыми в интерфейсе формулы Excel.</span><span class="sxs-lookup"><span data-stu-id="015c0-156">Required parameters will appear first in Excel’s Formula UI.</span></span>

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

<span data-ttu-id="015c0-157">Если вы определяете функцию, содержащую один или несколько необязательных параметров, нужно указать, что происходит, когда необязательный параметр не задан.</span><span class="sxs-lookup"><span data-stu-id="015c0-157">When you define a function that contains one or more optional parameters, you should specify what happens when the optional parameters are undefined.</span></span> <span data-ttu-id="015c0-158">В приведенном ниже примере `zipCode` и `dayOfWeek` являются необязательными параметрами для функции `getWeatherReport`.</span><span class="sxs-lookup"><span data-stu-id="015c0-158">In the following example, `zipCode` and `dayOfWeek` are both optional parameters for the `getWeatherReport` function.</span></span> <span data-ttu-id="015c0-159">Если параметр `zipCode` не определен, значение по умолчанию устанавливается равным 98052.</span><span class="sxs-lookup"><span data-stu-id="015c0-159">If the `zipCode` parameter is undefined, the default value is set to 98052.</span></span> <span data-ttu-id="015c0-160">Если параметр `dayOfWeek` не определен, ему присваивается значение Wednesday (Среда).</span><span class="sxs-lookup"><span data-stu-id="015c0-160">If the `dayOfWeek` parameter is undefined, it is set to Wednesday.</span></span>

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

## <a name="additional-considerations"></a><span data-ttu-id="015c0-161">Дополнительные рекомендации</span><span class="sxs-lookup"><span data-stu-id="015c0-161">Additional considerations</span></span>

<span data-ttu-id="015c0-162">Чтобы создать надстройку, которая будет работать на различных платформах (один из основных клиентов надстроек Office), вам не следует выполнять доступ к модели DOM в пользовательских функциях или использовать библиотеки, такие как jQuery, которые используют DOM.</span><span class="sxs-lookup"><span data-stu-id="015c0-162">In order to create an add-in that will run on multiple platforms (one of the key tenants of Office Add-ins), you should not access the Document Object Model (DOM) in custom functions or use libraries like jQuery that rely on the DOM.</span></span> <span data-ttu-id="015c0-163">В Excel для Windows, где пользовательские функции используют [среду выполнения JavaScript](custom-functions-runtime.md), пользовательские функции не могут выполнять доступ к DOM.</span><span class="sxs-lookup"><span data-stu-id="015c0-163">On Excel for Windows, where custom functions use the [JavaScript runtime](custom-functions-runtime.md), custom functions cannot access the DOM.</span></span>

## <a name="see-also"></a><span data-ttu-id="015c0-164">См. также</span><span class="sxs-lookup"><span data-stu-id="015c0-164">See also</span></span>

* [<span data-ttu-id="015c0-165">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="015c0-165">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="015c0-166">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="015c0-166">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="015c0-167">Среда выполнения для пользовательских функций Excel</span><span class="sxs-lookup"><span data-stu-id="015c0-167">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="015c0-168">Журнал изменений пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="015c0-168">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="015c0-169">Руководство по настраиваемым функциям в Excel</span><span class="sxs-lookup"><span data-stu-id="015c0-169">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
