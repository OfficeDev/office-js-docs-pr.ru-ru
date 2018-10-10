---
ms.date: 10/03/2018
description: Рекомендации и рекомендуемые шаблоны для настраиваемых функций Excel.
title: Рекомендации по настраиваемым функциям
ms.openlocfilehash: f6781de97f912df70800532032162187ae9f9344
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459114"
---
# <a name="custom-functions-best-practices-preview"></a><span data-ttu-id="0e512-103">Рекомендации по настраиваемым функциям (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="0e512-103">Custom functions best practices</span></span>

<span data-ttu-id="0e512-104">В этой статье приводятся рекомендации по разработке настраиваемых функций в Excel.</span><span class="sxs-lookup"><span data-stu-id="0e512-104">This article describes best practices for developing custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="error-handling"></a><span data-ttu-id="0e512-105">Обработка ошибок</span><span class="sxs-lookup"><span data-stu-id="0e512-105">Error handling</span></span>

<span data-ttu-id="0e512-p101">При построении надстройки, который определяет настраиваемые функции, не забудьте включить логику  обработки ошибок, возникающих в среде выполнения, для учетной записи. Обработка ошибок для настраиваемых функций  в целом совпадает с [обработкой ошибок для API JavaScript Excel](excel-add-ins-error-handling.md). В следующем образце кода `.catch` будут обработаны все ошибки, возникшие  ранее в коде.</span><span class="sxs-lookup"><span data-stu-id="0e512-p101">When you build an add-in that defines custom functions, be sure to include error handling logic to account for runtime errors. Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md). In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

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

## <a name="debugging"></a><span data-ttu-id="0e512-109">Отладка</span><span class="sxs-lookup"><span data-stu-id="0e512-109">Debugging</span></span>

<span data-ttu-id="0e512-110">На данный момент наилучшим методом отладки настраиваемых функций Excel является предварительная [загрузка неопубликованной надстройки](../testing/sideload-office-add-ins-for-testing.md) в рамках  **Excel Online**.</span><span class="sxs-lookup"><span data-stu-id="0e512-110">Currently, the best method for debugging Excel custom functions is to first [sideload](../testing/sideload-office-add-ins-for-testing.md) your add-in within **Excel Online**.</span></span> <span data-ttu-id="0e512-111">Затем можно выполнить отладку настраиваемых функций с помощью [встроенного в веб-обозреватель средства отладки F12](../testing/debug-add-ins-in-office-online.md) в сочетании со следующими методами:</span><span class="sxs-lookup"><span data-stu-id="0e512-111">You can then debug your custom functions by using the [F12 debugging tool native to your browser](../testing/debug-add-ins-in-office-online.md) in combination with the following techniques:</span></span>

- <span data-ttu-id="0e512-112">Используйте `console.log` операторы в коде настраиваемых функций для отправки выходных данных в консоль в режиме реального времени.</span><span class="sxs-lookup"><span data-stu-id="0e512-112">Use `console.log` statements within your custom functions code to send output to the console in real time.</span></span>

- <span data-ttu-id="0e512-p103">Используйте `debugger;` операторы в коде настраиваемых функций для указания точек останова, в которых выполнение будет приостановлено, если открыто окно F12. Например, если следующая функция выполняется при открытом окне F12, то ее выполнение будет приостановлено при достижении оператора `debugger;` , что  позволит вручную проверить значения параметров до возврата данных функцией. `debugger;` Оператор не оказывает влияния в Excel Online, когда не открыто окно F12. На данный момент `debugger;` оператор не оказывает никакого влияния в Excel для Windows.</span><span class="sxs-lookup"><span data-stu-id="0e512-p103">Use `debugger;` statements within your custom functions code to specify breakpoints where execution will pause when the F12 window is open. For example, if the following function runs while the F12 window is open, execution will pause on the `debugger;` statement, enabling you to manually inspect parameter values before the function returns. The `debugger;` statement has no effect in Excel Online when the F12 window is not open. Currently, the `debugger;` statement has no effect in Excel for Windows.</span></span>

    ```js
    function add(first, second){
      debugger;
      return first + second;
    }
    ```

<span data-ttu-id="0e512-117">Если надстройку не удалось зарегистрировать, [проверьте правильность настройки сертификатов SSL](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) для веб-сервера, на котором размещено приложение надстройки.</span><span class="sxs-lookup"><span data-stu-id="0e512-117">If your add-in fails to register, [verify that SSL certificates are correctly configured](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md) for the web server that's hosting your add-in application.</span></span>

<span data-ttu-id="0e512-118">При тестировании надстройки в классическом приложении Office 2016 можно включить [регистрацию времени выполнения](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in) для отладки проблем, связанных с XML-файлом манифеста надстройки, а также использовать несколько условий установки и выполнения.</span><span class="sxs-lookup"><span data-stu-id="0e512-118">If you are testing your add-in in Office 2016 desktop, you can enable [runtime logging](../testing/troubleshoot-manifest.md#use-runtime-logging-to-debug-your-add-in) to debug issues with your add-in's XML manifest file as well as several installation and runtime conditions.</span></span>

## <a name="mapping-function-names-to-json-metadata"></a><span data-ttu-id="0e512-119">Сопоставление имен функций с метаданными JSON</span><span class="sxs-lookup"><span data-stu-id="0e512-119">Mapping function names to JSON metadata</span></span>

<span data-ttu-id="0e512-p104">Как описано в статье [Обзор настраиваемых функций](custom-functions-overview.md) , проект настраиваемых функций должен включать в себя файл метаданных JSON, который содержит информацию, необходимую Excel для регистрации настраиваемых функций и их предоставления конечным пользователям.  Кроме того, в файле JavaScript, которым определяются настраиваемые функции, должна содержаться информация о том, какие объекты функции в файле метаданных JSON соответствуют каждой из настраиваемых функций в файле JavaScript.</span><span class="sxs-lookup"><span data-stu-id="0e512-p104">As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users. Additionally, within the JavaScript file that defines your custom functions, you must provide information to specify which function object in the JSON metadata file corresponds to each custom function in the JavaScript file.</span></span>

<span data-ttu-id="0e512-122">Например, приведенный ниже пример кода определяет настраиваемую функцию `add` и указывает, какая функция `add` соответствует объекту в файле метаданных JSON, где значение свойства `id` равно **ADD**.</span><span class="sxs-lookup"><span data-stu-id="0e512-122">For example, the following code sample defines the custom function `add` and then specifies that the function `add` corresponds to the object in the JSON metadata file where the value of the `id` property is **ADD**.</span></span>

```js
function add(first, second){
  return first + second;
}

CustomFunctionMappings.ADD = add;
```

<span data-ttu-id="0e512-123">При создании настраиваемых функций в файле JavaScript и указании соответствующей информации в файле метаданных JSON принимайте во внимание следующие рекомендации.</span><span class="sxs-lookup"><span data-stu-id="0e512-123">Keep in mind the following best practices when creating custom functions in your JavaScript file and specifying corresponding information in the JSON metadata file.</span></span>

* <span data-ttu-id="0e512-p105">В файле JavaScript укажите имена функций в camelCase. Примером может служить запись имени функции `addTenToInput` в camelCase: первое слово в имени начинается со строчной буквы нижнего регистра, а все последующие — с прописной буквы.</span><span class="sxs-lookup"><span data-stu-id="0e512-p105">In the JavaScript file, specify function names in camelCase. For example, the function name `addTenToInput` is written in camelCase: the first word in the name starts with a lowercase letter and each subsequent word in the name starts with an uppercase letter.</span></span>

* <span data-ttu-id="0e512-p106">В файле метаданных JSON укажите прописными буквами значение каждого  свойства`name`.  Свойство `name`определяет имя функции, которое конечные пользователи видят в Excel. Использование прописных букв для имен всех настраиваемых функций позволяет сформировать согласованное представление в Excel для конечных пользователей, при котором все имена встроенных функций показываются прописными буквами.</span><span class="sxs-lookup"><span data-stu-id="0e512-p106">In the JSON metadata file, specify the value of each `name` property in uppercase. The `name` property defines the function name that end users will see in Excel. Using uppercase letters for the name of each custom function provides a consistent experience for end users in Excel, where all built-in function names are uppercase.</span></span>

* <span data-ttu-id="0e512-p107">В файле метаданных JSON укажите прописными буквами значение каждого  `id` свойства. Такой подход позволяет четко обозначить, какая часть оператора `CustomFunctionMappings` в коде JavaScript соответствует свойству  `id` в  файле метаданных JSON (при условии, что именем функции используется camelCase в соответствии с приведенными выше рекомендациями).</span><span class="sxs-lookup"><span data-stu-id="0e512-p107">In the JSON metadata file, specify the value of each `id` property in uppercase. Doing so makes it obvious which part of the `CustomFunctionMappings` statement in your JavaScript code corresponds to the `id` property in the JSON metadata file (provided that your function name uses camelCase, as recommended earlier).</span></span>

* <span data-ttu-id="0e512-p108">Убедитесь в том, что в файле метаданных JSON значение каждого из свойств `id`  является уникальным для этого файла. При этом в файле метаданных не должно быть двух объектов функции, имеющих одинаковое `id` значение. Кроме того, не указывайте в файле метаданных два значения `id` которые различаются только регистром.  Например, не используйте для определения одного объекта функции `id` значение **add**, а для определения другого — `id` значение **ADD**.</span><span class="sxs-lookup"><span data-stu-id="0e512-p108">In the JSON metadata file, ensure that the value of each `id` property is unique within the scope of the file. That is, no two function objects in the metadata file should have the same `id` value. Additionally, do not specify two `id` values in the metadata file that only differ by case. For example, do not define one function object with an `id` value of **add** and another function object with an `id` value of **ADD**.</span></span>

* <span data-ttu-id="0e512-p109">Не изменяйте значение свойства  `id` в файле метаданных JSON после его сопоставления с соответствующим именем функции JavaScript. Имя функции, отображаемое в Excel для конечных пользователей, можно изменить, обновив свойство `name` в файле метаданных JSON, но не следует ни при каких обстоятельствах изменять значение свойства `id` после его установки.</span><span class="sxs-lookup"><span data-stu-id="0e512-p109">Do not change the value of an `id` property in the JSON metadata file after it's been mapped to a corresponding JavaScript function name. You can change the function name that end users see in Excel by updating the `name` property within the JSON metadata file, but you should never change the value of an `id` property after it's been established.</span></span>

* <span data-ttu-id="0e512-p110">В файле JavaScript укажите все сопоставления настраиваемых функций для одного и того же расположения. Так, к примеру, в приведенном далее примере кода определяются две настраиваемые функции, а затем указывается информация о сопоставлении для обеих функций.</span><span class="sxs-lookup"><span data-stu-id="0e512-p110">In the JavaScript file, specify all custom function mappings in the same location. For example, the following code sample defines two custom functions and then specifies the mapping information for both functions.</span></span>

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

    <span data-ttu-id="0e512-139">В следующем примере показаны метаданные JSON, соответствующие функциям, определенным в этом примере кода JavaScript.</span><span class="sxs-lookup"><span data-stu-id="0e512-139">The following sample shows the JSON metadata that corresponds to the functions defined in this JavaScript code sample.</span></span>

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

## <a name="additional-considerations"></a><span data-ttu-id="0e512-140">Дополнительные рекомендации</span><span class="sxs-lookup"><span data-stu-id="0e512-140">Additional considerations</span></span>

<span data-ttu-id="0e512-141">Чтобы создать надстройку, которая будет работать на нескольких платформах (для одного из основных клиентов надстроек Office), вы не должны запрашивать доступ к объектной модели документа (DOM) в настраиваемых функциях или использовать библиотеки, такие как jQuery, которые полагаются на DOM.</span><span class="sxs-lookup"><span data-stu-id="0e512-141">In order to create an add-in that will run on multiple platforms (one of the key tenants of Office Add-ins), you should not access the Document Object Model (DOM) in custom functions or use libraries like jQuery that rely on the DOM.</span></span> <span data-ttu-id="0e512-142">В Excel для Windows настраиваемые функции, использующие [среду выполнения JavaScript](custom-functions-runtime.md), не могут получить доступ к DOM.</span><span class="sxs-lookup"><span data-stu-id="0e512-142">On Excel for Windows, where custom functions use the [JavaScript runtime](custom-functions-runtime.md), custom functions cannot access the DOM.</span></span>

## <a name="see-also"></a><span data-ttu-id="0e512-143">См. также</span><span class="sxs-lookup"><span data-stu-id="0e512-143">See also</span></span>

* [<span data-ttu-id="0e512-144">Создание настраиваемых функций в Excel</span><span class="sxs-lookup"><span data-stu-id="0e512-144">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="0e512-145">Настраиваемые функции метаданных</span><span class="sxs-lookup"><span data-stu-id="0e512-145">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="0e512-146">Среда выполнения для настраиваемых функций Excel</span><span class="sxs-lookup"><span data-stu-id="0e512-146">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="0e512-147">Руководство по настраиваемых функциях Excel</span><span class="sxs-lookup"><span data-stu-id="0e512-147">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)
