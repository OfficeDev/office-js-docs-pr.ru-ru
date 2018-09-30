---
ms.date: 09/27/2018
description: Рекомендации и рекомендуемые шаблоны для настраиваемых функций Excel.
title: Рекомендации по настраиваемым функциям
ms.openlocfilehash: d157464a3a8bf453cd0970281f1a4fdd27df5d25
ms.sourcegitcommit: 1852ae367de53deb91d03ca55d16eb69709340d3
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/29/2018
ms.locfileid: "25348789"
---
# <a name="custom-functions-best-practices-preview"></a><span data-ttu-id="6ec63-103">Рекомендации по настраиваемым функциям (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="6ec63-103">Custom functions best practices</span></span>

<span data-ttu-id="6ec63-104">В этой статье приводятся рекомендации по разработке настраиваемых функций в Excel.</span><span class="sxs-lookup"><span data-stu-id="6ec63-104">This article describes best practices for developing custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="error-handling"></a><span data-ttu-id="6ec63-105">Обработка ошибок</span><span class="sxs-lookup"><span data-stu-id="6ec63-105">Error handling</span></span>

<span data-ttu-id="6ec63-106">При построении надстройки, которая определяет настраиваемые функции, не забудьте включить логику обработки ошибок, возникающих в среде выполнения.</span><span class="sxs-lookup"><span data-stu-id="6ec63-106">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="6ec63-107">Обработка ошибок для настраиваемых функций такая же, как и в случае [обработки ошибок для API JavaScript Excel в целом](excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="6ec63-107">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="6ec63-108">В следующем примере кода `.catch` будет обрабатывать все ошибки, возникшие ранее в этом коде.</span><span class="sxs-lookup"><span data-stu-id="6ec63-108">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

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

## <a name="debugging"></a><span data-ttu-id="6ec63-109">Отладка</span><span class="sxs-lookup"><span data-stu-id="6ec63-109">Debugging</span></span>

<span data-ttu-id="6ec63-p102">На данный момент наилучшим способом отладки настраиваемых функций Excel является [загрузка неопубликованных](../testing/sideload-office-add-ins-for-testing.md) надстроек в **Excel Online**. После этого отладку настраиваемых функций можно выполнить с помощью [входящего в состав веб-обозревателя средства отладки, вызываемого при нажатии на F12](../testing/debug-add-ins-in-office-online.md), используемого в сочетании со следующими методами:</span><span class="sxs-lookup"><span data-stu-id="6ec63-p102">Currently, the best method for debugging Excel custom functions is to first [sideload](../testing/sideload-office-add-ins-for-testing.md) your add-in within **Excel Online**. You can then debug your custom functions by using the [F12 debugging tool native to your browser](../testing/debug-add-ins-in-office-online.md). Use  statements within your custom functions code to send output to the console in real time.</span></span>

- <span data-ttu-id="6ec63-112">Используйте операторы `console.log` в коде настраиваемых функций для отправки выходных данных в консоль в режиме реального времени.</span><span class="sxs-lookup"><span data-stu-id="6ec63-112">Use `console.log` statements within your custom functions code to send output to the console in real time.</span></span>

- <span data-ttu-id="6ec63-113">Используйте операторы `debugger;` в коде настраиваемых функций для указания точек останова, в которых выполнение будет приостановлено, если открыто окно F12.</span><span class="sxs-lookup"><span data-stu-id="6ec63-113">Use `debugger;` statements within your custom functions code to specify breakpoints where execution will pause when the F12 window is open.</span></span> <span data-ttu-id="6ec63-114">Так, к примеру, если следующая функция выполняется при открытом окне F12, то ее выполнение будет приостановлено при достижнии оператора `debugger;`, что позволит вручную проверить значения параметров до возврата данных функцией.</span><span class="sxs-lookup"><span data-stu-id="6ec63-114">For example, if the following function runs while the F12 window is open, execution will pause on the `debugger;` statement, enabling you to manually inspect parameter values before the function returns.</span></span> <span data-ttu-id="6ec63-115">Если окно F12 не открыто, то оператор `debugger;` в Excel Online не влияет на выполнение функции.</span><span class="sxs-lookup"><span data-stu-id="6ec63-115">The `debugger;` statement has no effect in Excel Online when the F12 window is not open.</span></span> <span data-ttu-id="6ec63-116">В настоящий момент оператор `debugger;` в Excel для Windows не работает.</span><span class="sxs-lookup"><span data-stu-id="6ec63-116">Currently, the `debugger;` statement has no effect in Excel for Windows.</span></span>

    ```js
    function add(first, second){
      debugger;
      return first + second;
    }
    ```

<span data-ttu-id="6ec63-117">Если надстройку не удалось зарегистрировать, [проверьте правильность настройки сертификатов SSL](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) для веб-сервера, на котором размещено приложение надстройки.</span><span class="sxs-lookup"><span data-stu-id="6ec63-117">If your add-in fails to register, [verify that SSL certificates are correctly configured](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) for the web server that's hosting your add-in application.</span></span>

<span data-ttu-id="6ec63-118">При тестировании надстройки в классическом приложении Office 2016 можно включить [регистрацию времени выполнения](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in) для отладки проблем, связанных с XML-файлом манифеста надстройки, а также использовать несколько условий установки и выполнения.</span><span class="sxs-lookup"><span data-stu-id="6ec63-118">If you are testing your add-in in Office 2016 desktop, you can enable [runtime logging](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in) to debug issues with your add-in's XML manifest file as well as several installation and runtime conditions.</span></span>

## <a name="mapping-function-names-to-json-metadata"></a><span data-ttu-id="6ec63-119">Сопоставление имен функций с метаданными JSON</span><span class="sxs-lookup"><span data-stu-id="6ec63-119">Mapping function names to JSON metadata</span></span>

<span data-ttu-id="6ec63-120">Как описано в статье [Обзор настраиваемых функций](custom-functions-overview.md), проект настраиваемых функций должен включать в себя файл метаданных JSON, который содержит информацию, необходимую Excel для регистрации настраиваемых функций и их предоставления конечным пользователям.</span><span class="sxs-lookup"><span data-stu-id="6ec63-120">As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users.</span></span> <span data-ttu-id="6ec63-121">Кроме того, в файле JavaScript, которым определяются настраиваемые функции, должна содержаться информация о том, какие объекты функции в файле метаданных JSON сответствуют каждой из настраиваемых функций в файле JavaScript.</span><span class="sxs-lookup"><span data-stu-id="6ec63-121">Additionally, within the JavaScript file that defines your custom functions, you must provide information to specify which function object in the JSON metadata file corresponds to each custom function in the JavaScript file.</span></span>

<span data-ttu-id="6ec63-122">Так, к примеру, приведенный ниже пример кода определяет настраиваемую функцию `add` и указывает, какая функция `add` соответствует объекту в файле метаданных JSON, где значение свойства `id` равно **ADD**.</span><span class="sxs-lookup"><span data-stu-id="6ec63-122">For example, the following code sample defines the custom function `add` and then specifies that the function `add` corresponds to the object in the JSON metadata file where the value of the `id` property is **ADD**.</span></span>

```js
function add(first, second){
  return first + second;
}

CustomFunctionMappings.ADD = add;
```

<span data-ttu-id="6ec63-123">При создании настраиваемых функций в файле JavaScript и указании соответствующей информации в файле метаданных JSON принимайте во внимание следующие рекомендации.</span><span class="sxs-lookup"><span data-stu-id="6ec63-123">Keep in mind the following best practices when creating custom functions in your JavaScript file and specifying corresponding information in the JSON metadata file.</span></span>

* <span data-ttu-id="6ec63-124">В файле JavaScript укажите имена функций в camelCase.</span><span class="sxs-lookup"><span data-stu-id="6ec63-124">In the JavaScript file, specify function names in camelCase.</span></span> <span data-ttu-id="6ec63-125">Примером может служить запись имени функции `addTenToInput` в camelCase: первое слово в имени начинается со строчной буквы нижнего регистра, а все последующие — с прописной буквы.</span><span class="sxs-lookup"><span data-stu-id="6ec63-125">For example, the function name `addTenToInput` is written in camelCase: the first word in the name starts with a lowercase letter and each subsequent word in the name starts with an uppercase letter.</span></span>

* <span data-ttu-id="6ec63-126">В файле метаданных JSON укажите прописными буквами значение каждого свойства `name`.</span><span class="sxs-lookup"><span data-stu-id="6ec63-126">In the JSON metadata file, specify the value of each `name` property in uppercase.</span></span> <span data-ttu-id="6ec63-127">Свойство `name` определяет имя функции, которое конечные пользователи видят в Excel.</span><span class="sxs-lookup"><span data-stu-id="6ec63-127">The `name` property defines the function name that end users will see in Excel.</span></span> <span data-ttu-id="6ec63-128">Использование прописных букв для имен всех настраиваемых функций позволяет сформировать согласованное представление в Excel для конечных пользователей, при котором все имена встроенных функций показываются прописными буквами.</span><span class="sxs-lookup"><span data-stu-id="6ec63-128">Using uppercase letters for the name of each custom function provides a consistent experience for end users in Excel, where all built-in function names are uppercase.</span></span>

* <span data-ttu-id="6ec63-129">В файле метаданных JSON укажите прописными буквами значение каждого свойства `id`.</span><span class="sxs-lookup"><span data-stu-id="6ec63-129">In the JSON metadata file, specify the value of each `id` property in uppercase.</span></span> <span data-ttu-id="6ec63-130">Такой подход позволяет четко обозначить, какая часть оператора `CustomFunctionMappings` в коде JavaScript соответствует свойству `id` в файле метаданных JSON (при условии, что именем функции используется camelCase в соответствии с приведенными выше рекомендациями).</span><span class="sxs-lookup"><span data-stu-id="6ec63-130">Doing so makes it obvious which part of the `CustomFunctionMappings` statement in your JavaScript code corresponds to the `id` property in the JSON metadata file (provided that your function name uses camelCase, as recommended earlier).</span></span>

* <span data-ttu-id="6ec63-131">Убедитесь в том, что в файле метаданных JSON значение каждого из свойств `id` является уникальным для этого файла.</span><span class="sxs-lookup"><span data-stu-id="6ec63-131">In the JSON metadata file, ensure that the value of each `id` property is unique within the scope of the file.</span></span> <span data-ttu-id="6ec63-132">При этом в файле метаданных не должно быть двух объектов функции, имеющих одинаковое значение `id`.</span><span class="sxs-lookup"><span data-stu-id="6ec63-132">That is, no two function objects in the metadata file should have the same `id` value.</span></span> <span data-ttu-id="6ec63-133">Кроме того, не указывайте в файле метаданных два значения `id`, которые различаются только регистром.</span><span class="sxs-lookup"><span data-stu-id="6ec63-133">Additionally, do not specify two `id` values in the metadata file that only differ by case.</span></span> <span data-ttu-id="6ec63-134">К примеру, не используйте для определения одного объекта функции значение `id` **add**, а для определения другого — значение `id` **ADD**.</span><span class="sxs-lookup"><span data-stu-id="6ec63-134">For example, do not define one function object with an `id` value of **add** and another function object with an `id` value of **ADD**.</span></span>

* <span data-ttu-id="6ec63-135">Не изменяйте значение свойства `id` в файле метаданных JSON после его сопоставления с соответствующим именем функции JavaScript.</span><span class="sxs-lookup"><span data-stu-id="6ec63-135">Do not change the value of an `id` property in the JSON metadata file after it's been mapped to a corresponding JavaScript function name.</span></span> <span data-ttu-id="6ec63-136">Имя функции, отображаемое в Excel для конечных пользователей, можно изменить, обновив свойство `name` в файле метаданных JSON, но изменять значение свойства `id` после его установки не следует ни при каких обстоятельствах.</span><span class="sxs-lookup"><span data-stu-id="6ec63-136">You can change the function name that end users see in Excel by updating the `name` property within the JSON metadata file, but you should never change the value of an `id` property after it's been established.</span></span>

* <span data-ttu-id="6ec63-137">В файле JavaScript укажите все сопоставления настраиваемых функций для одного и того же расположения.</span><span class="sxs-lookup"><span data-stu-id="6ec63-137">In the JavaScript file, specify all custom function mappings in the same location.</span></span> <span data-ttu-id="6ec63-138">Так, к примеру, в приведенном далее примере кода определяются две настраиваемые функции, а затем указывается информация о сопоставлении для обеих функций.</span><span class="sxs-lookup"><span data-stu-id="6ec63-138">For example, the following code sample defines two custom functions and then specifies the mapping information for both functions.</span></span>

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

    <span data-ttu-id="6ec63-139">В следующем примере показаны метаданные JSON, соответствующие функциям, определенным в этом примере кода JavaScript.</span><span class="sxs-lookup"><span data-stu-id="6ec63-139">The following sample shows the JSON metadata that corresponds to the functions defined in this JavaScript code sample.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="6ec63-140">См. также</span><span class="sxs-lookup"><span data-stu-id="6ec63-140">See also</span></span>

* [<span data-ttu-id="6ec63-141">Создание настраиваемых функций в Excel</span><span class="sxs-lookup"><span data-stu-id="6ec63-141">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="6ec63-142">Настраиваемые функции метаданных</span><span class="sxs-lookup"><span data-stu-id="6ec63-142">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="6ec63-143">Среда выполнения для настраиваемых функций Excel</span><span class="sxs-lookup"><span data-stu-id="6ec63-143">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="6ec63-144">Руководство по настраиваемых функциях Excel</span><span class="sxs-lookup"><span data-stu-id="6ec63-144">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)
