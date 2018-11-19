---
ms.date: 10/24/2018
description: Ознакомьтесь с советами и рекомендованными шаблонами в отношении пользовательских функций Excel.
title: Рекомендации в отношении пользовательских функций
ms.openlocfilehash: 0408318227e1f89726ed7c0e4dfbb8e6340abef4
ms.sourcegitcommit: 52d18dd8a60e0cec1938394669d577570700e61e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/26/2018
ms.locfileid: "25797401"
---
# <a name="custom-functions-best-practices-preview"></a><span data-ttu-id="801e3-103">Рекомендации в отношении пользовательских функций (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="801e3-103">Custom functions best practices (preview)</span></span>

<span data-ttu-id="801e3-104">В этой статье описаны рекомендации по разработке пользовательских функций в Excel.</span><span class="sxs-lookup"><span data-stu-id="801e3-104">This article describes best practices for developing custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="error-handling"></a><span data-ttu-id="801e3-105">Обработка ошибок</span><span class="sxs-lookup"><span data-stu-id="801e3-105">Error handling</span></span>

<span data-ttu-id="801e3-106">При создании надстройки, определяющей пользовательские функции, не забудьте включить логику для обработки ошибок, возникающих в среде выполнения.</span><span class="sxs-lookup"><span data-stu-id="801e3-106">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="801e3-107">Обработка ошибок для пользовательских функций в значительной степени совпадает с [обработкой ошибок для API JavaScript в Excel](excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="801e3-107">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="801e3-108">В приведенном ниже примере кода `.catch` обрабатывает любые ошибки, возникающие в коде ранее.</span><span class="sxs-lookup"><span data-stu-id="801e3-108">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

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

## <a name="troubleshooting"></a><span data-ttu-id="801e3-109">Устранение неполадок</span><span class="sxs-lookup"><span data-stu-id="801e3-109">Troubleshooting</span></span>

<span data-ttu-id="801e3-110">Если вы проверяете надстройку в Office для Windows, нужно включить **[ведение журнала в среде выполнения](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in)**, чтобы устранять проблемы с XML-файлом манифеста надстройки, а также с некоторыми условиями установки и среды выполнения.</span><span class="sxs-lookup"><span data-stu-id="801e3-110">If you are testing your add-in in Office on Windows, you should enable **[runtime logging](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in)** to troubleshoot issues with your add-in's XML manifest file, as well as several installation and runtime conditions.</span></span> <span data-ttu-id="801e3-111">В файл журнала в среде выполнения записываются операторы `console.log`, что облегчает выявление проблем.</span><span class="sxs-lookup"><span data-stu-id="801e3-111">Runtime logging writes `console.log` statements to a log file to help you uncover issues.</span></span>

<span data-ttu-id="801e3-112">Чтобы поделиться своим мнением об этом методе устранения неполадок с группой, отвечающей за пользовательские функции Excel, отправьте отзыв группе.</span><span class="sxs-lookup"><span data-stu-id="801e3-112">To report feedback to the Excel Custom Functions team about this method of troubleshooting, send the team feedback.</span></span> <span data-ttu-id="801e3-113">Для этого выберите **Файл | Отзыв | Отправить нахмуренный смайлик**.</span><span class="sxs-lookup"><span data-stu-id="801e3-113">To do this, select **File | Feedback | Send a Frown**.</span></span> <span data-ttu-id="801e3-114">Отправка нахмуренного смайлика предоставит необходимые журналы для понимания проблемы, на которую вы указываете.</span><span class="sxs-lookup"><span data-stu-id="801e3-114">Sending a frown will provide the necessary logs to understand the issue you are hitting.</span></span> 

## <a name="debugging"></a><span data-ttu-id="801e3-115">Отладка</span><span class="sxs-lookup"><span data-stu-id="801e3-115">Debugging</span></span>

<span data-ttu-id="801e3-116">В настоящее время наилучшим способом отладки пользовательских функций Excel является [загрузка неопубликованной](../testing/sideload-office-add-ins-for-testing.md) надстройки в **Excel Online**.</span><span class="sxs-lookup"><span data-stu-id="801e3-116">Currently, the best method for debugging Excel custom functions is to first [sideload](../testing/sideload-office-add-ins-for-testing.md) your add-in within **Excel Online**.</span></span> <span data-ttu-id="801e3-117">Затем вы сможете выполнить отладку пользовательских функций с помощью [собственного средства отладки вашего браузера, вызываемого клавишей F12,](../testing/debug-add-ins-in-office-online.md) в сочетании с указанными ниже методами.</span><span class="sxs-lookup"><span data-stu-id="801e3-117">You can then debug your custom functions by using the [F12 debugging tool native to your browser](../testing/debug-add-ins-in-office-online.md) in combination with the following techniques:</span></span>

- <span data-ttu-id="801e3-118">Используйте операторы `console.log` в коде пользовательских функций, чтобы отправлять результаты в консоль в режиме реального времени.</span><span class="sxs-lookup"><span data-stu-id="801e3-118">Use `console.log` statements within your custom functions code to send output to the console in real time.</span></span>

- <span data-ttu-id="801e3-119">Используйте операторы `debugger;` в коде пользовательских функций, чтобы указать точки останова для приостановки выполнения при открытом окне, вызываемом клавишей F12.</span><span class="sxs-lookup"><span data-stu-id="801e3-119">Use `debugger;` statements within your custom functions code to specify breakpoints where execution will pause when the F12 window is open.</span></span> <span data-ttu-id="801e3-120">Например, если выполняется следующая функция при открытом окне F12, выполнение приостанавливается на операторе `debugger;`, что позволяет вручную проверить значения параметров перед возвратом функции.</span><span class="sxs-lookup"><span data-stu-id="801e3-120">For example, if the following function runs while the F12 window is open, execution will pause on the `debugger;` statement, enabling you to manually inspect parameter values before the function returns.</span></span> <span data-ttu-id="801e3-121">Оператор `debugger;` не выполняет никаких действий в Excel Online, если не открыто окно F12.</span><span class="sxs-lookup"><span data-stu-id="801e3-121">The `debugger;` statement has no effect in Excel Online when the F12 window is not open.</span></span> <span data-ttu-id="801e3-122">В настоящее время оператор `debugger;` не выполняет никаких действий в Excel для Windows.</span><span class="sxs-lookup"><span data-stu-id="801e3-122">Currently, the `debugger;` statement has no effect in Excel for Windows.</span></span>

    ```js
    function add(first, second){
      debugger;
      return first + second;
    }
    ```

<span data-ttu-id="801e3-123">Если не удается зарегистрировать надстройку, [проверьте, что SSL-сертификаты правильно настроены](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) для веб-сервера, на котором размещается ваше приложение надстройки.</span><span class="sxs-lookup"><span data-stu-id="801e3-123">If your add-in fails to register, [verify that SSL certificates are correctly configured](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) for the web server that's hosting your add-in application.</span></span>

## <a name="mapping-function-names-to-json-metadata"></a><span data-ttu-id="801e3-124">Сопоставление имен функций с метаданными JSON</span><span class="sxs-lookup"><span data-stu-id="801e3-124">Mapping function names to JSON metadata</span></span>

<span data-ttu-id="801e3-125">Как указано в статье [Обзор пользовательских функций](custom-functions-overview.md), проект пользовательских функций должен содержать файл метаданных JSON, который предоставляет сведения, необходимые Excel для регистрации пользовательских функций и обеспечения их доступности для конечных пользователей.</span><span class="sxs-lookup"><span data-stu-id="801e3-125">As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users.</span></span> <span data-ttu-id="801e3-126">Кроме того, в файле JavaScript, определяющем пользовательские функции, необходимо предоставить сведения, которые указывают, какой объект функции в файле метаданных JSON соответствует каждой пользовательской функции в файле JavaScript.</span><span class="sxs-lookup"><span data-stu-id="801e3-126">Additionally, within the JavaScript file that defines your custom functions, you must provide information to specify which function object in the JSON metadata file corresponds to each custom function in the JavaScript file.</span></span>

<span data-ttu-id="801e3-127">Например, в приведенном ниже примера кода определяется пользовательская функция `add` и указывается, что функция `add` соответствует объекту в файле метаданных JSON, при этом свойству `id` присваивается значение **ADD**.</span><span class="sxs-lookup"><span data-stu-id="801e3-127">For example, the following code sample defines the custom function `add` and then specifies that the function `add` corresponds to the object in the JSON metadata file where the value of the `id` property is **ADD**.</span></span>

```js
function add(first, second){
  return first + second;
}

CustomFunctionMappings.ADD = add;
```

<span data-ttu-id="801e3-128">Имейте в виду приведенные ниже рекомендации при создании пользовательских функций в файле JavaScript и указании соответствующих сведений в файле метаданных JSON.</span><span class="sxs-lookup"><span data-stu-id="801e3-128">Keep in mind the following best practices when creating custom functions in your JavaScript file and specifying corresponding information in the JSON metadata file.</span></span>

* <span data-ttu-id="801e3-129">В файле JavaScript укажите имена функций в стиле camelCase.</span><span class="sxs-lookup"><span data-stu-id="801e3-129">In the JavaScript file, specify function names in camelCase.</span></span> <span data-ttu-id="801e3-130">Например, имя функции `addTenToInput` записано в стиле camelCase: первое слово имени начинается со строчной буквы, а каждое последующее слово начинается с прописной буквы.</span><span class="sxs-lookup"><span data-stu-id="801e3-130">For example, the function name `addTenToInput` is written in camelCase: the first word in the name starts with a lowercase letter and each subsequent word in the name starts with an uppercase letter.</span></span>

* <span data-ttu-id="801e3-131">В файле метаданных JSON укажите значение каждого свойства `name` прописными буквами.</span><span class="sxs-lookup"><span data-stu-id="801e3-131">In the JSON metadata file, specify the value of each `name` property in uppercase.</span></span> <span data-ttu-id="801e3-132">Свойство `name` определяет имя функции, которое отображается для конечных пользователей в Excel.</span><span class="sxs-lookup"><span data-stu-id="801e3-132">The `name` property defines the function name that end users will see in Excel.</span></span> <span data-ttu-id="801e3-133">Использование прописных букв для имени каждой пользовательской функции обеспечивает единый интерфейс для конечных пользователей в Excel, где все имена встроенных функций записаны прописными буквами.</span><span class="sxs-lookup"><span data-stu-id="801e3-133">Using uppercase letters for the name of each custom function provides a consistent experience for end users in Excel, where all built-in function names are uppercase.</span></span>

* <span data-ttu-id="801e3-134">В файле метаданных JSON укажите значение каждого свойства `id` прописными буквами.</span><span class="sxs-lookup"><span data-stu-id="801e3-134">In the JSON metadata file, specify the value of each `id` property in uppercase.</span></span> <span data-ttu-id="801e3-135">Благодаря этому становится понятно, какая часть оператора `CustomFunctionMappings` в коде JavaScript соответствует свойству `id` в файле метаданных JSON (при условии, что для имени функции используется стиль camelCase, как рекомендуется выше).</span><span class="sxs-lookup"><span data-stu-id="801e3-135">Doing so makes it obvious which part of the `CustomFunctionMappings` statement in your JavaScript code corresponds to the `id` property in the JSON metadata file (provided that your function name uses camelCase, as recommended earlier).</span></span>

* <span data-ttu-id="801e3-136">Убедитесь, что в файле метаданных JSON значение каждого свойства `id` содержит только буквы, цифры и точки.</span><span class="sxs-lookup"><span data-stu-id="801e3-136">In the JSON metadata file, ensure that the value of each `id` property contains only alphanumeric characters and periods.</span></span> 

* <span data-ttu-id="801e3-137">Убедитесь, что в файле метаданных JSON значение каждого свойства `id` уникально в пределах файла.</span><span class="sxs-lookup"><span data-stu-id="801e3-137">In the JSON metadata file, ensure that the value of each `id` property is unique within the scope of the file.</span></span> <span data-ttu-id="801e3-138">То есть никакие два объекта функций в файле метаданных не должны иметь одинаковое значение `id`.</span><span class="sxs-lookup"><span data-stu-id="801e3-138">That is, no two function objects in the metadata file should have the same `id` value.</span></span> <span data-ttu-id="801e3-139">Кроме того, не указывайте два значения `id` в файле метаданных, которые отличаются только регистром.</span><span class="sxs-lookup"><span data-stu-id="801e3-139">Additionally, do not specify two `id` values in the metadata file that only differ by case.</span></span> <span data-ttu-id="801e3-140">Например, не определяйте один объект функции со значением `id` равным **add**, а другой объект функции со значением `id` равным **ADD**.</span><span class="sxs-lookup"><span data-stu-id="801e3-140">For example, do not define one function object with an `id` value of **add** and another function object with an `id` value of **ADD**.</span></span>

* <span data-ttu-id="801e3-141">Не изменяйте значение свойства `id` в файле метаданных JSON после его сопоставления с соответствующим именем функции JavaScript.</span><span class="sxs-lookup"><span data-stu-id="801e3-141">Do not change the value of an `id` property in the JSON metadata file after it's been mapped to a corresponding JavaScript function name.</span></span> <span data-ttu-id="801e3-142">Вы можете изменить имя функции, которое отображается для конечных пользователей в Excel, путем обновления свойства `name` в файле метаданных JSON, но никогда не следует изменять значение свойства `id` после его установления.</span><span class="sxs-lookup"><span data-stu-id="801e3-142">You can change the function name that end users see in Excel by updating the `name` property within the JSON metadata file, but you should never change the value of an `id` property after it's been established.</span></span>

* <span data-ttu-id="801e3-143">В файле JavaScript указывайте все сопоставления пользовательских функций в одном расположении.</span><span class="sxs-lookup"><span data-stu-id="801e3-143">In the JavaScript file, specify all custom function mappings in the same location.</span></span> <span data-ttu-id="801e3-144">Например, в приведенном ниже примере кода определяются две пользовательские функции и указываются сведения о сопоставлении для обеих функций.</span><span class="sxs-lookup"><span data-stu-id="801e3-144">For example, the following code sample defines two custom functions and then specifies the mapping information for both functions.</span></span>

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

    // map `id` values in the JSON metadata file to JavaScript function names
    CustomFunctionMappings.ADD = add;
    CustomFunctionMappings.INCREMENT = increment;
    ```

    <span data-ttu-id="801e3-145">В приведенном ниже примере показаны метаданные JSON, соответствующие функциям, определенным в этом примере кода JavaScript.</span><span class="sxs-lookup"><span data-stu-id="801e3-145">The following sample shows the JSON metadata that corresponds to the functions defined in this JavaScript code sample.</span></span>

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

## <a name="additional-considerations"></a><span data-ttu-id="801e3-146">Дополнительные рекомендации</span><span class="sxs-lookup"><span data-stu-id="801e3-146">Additional considerations</span></span>

<span data-ttu-id="801e3-147">Чтобы создать надстройку, которая будет работать на различных платформах (один из основных клиентов надстроек Office), вам не следует выполнять доступ к модели DOM в пользовательских функциях или использовать библиотеки, такие как jQuery, которые используют DOM.</span><span class="sxs-lookup"><span data-stu-id="801e3-147">In order to create an add-in that will run on multiple platforms (one of the key tenants of Office Add-ins), you should not access the Document Object Model (DOM) in custom functions or use libraries like jQuery that rely on the DOM.</span></span> <span data-ttu-id="801e3-148">В Excel для Windows, где пользовательские функции используют [среду выполнения JavaScript](custom-functions-runtime.md), пользовательские функции не могут выполнять доступ к DOM.</span><span class="sxs-lookup"><span data-stu-id="801e3-148">On Excel for Windows, where custom functions use the [JavaScript runtime](custom-functions-runtime.md), custom functions cannot access the DOM.</span></span>

## <a name="see-also"></a><span data-ttu-id="801e3-149">См. также</span><span class="sxs-lookup"><span data-stu-id="801e3-149">See also</span></span>

* [<span data-ttu-id="801e3-150">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="801e3-150">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="801e3-151">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="801e3-151">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="801e3-152">Среда выполнения для пользовательских функций Excel</span><span class="sxs-lookup"><span data-stu-id="801e3-152">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="801e3-153">Руководство по пользовательским функциям в Excel</span><span class="sxs-lookup"><span data-stu-id="801e3-153">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)
