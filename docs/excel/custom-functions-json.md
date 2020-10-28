---
ms.date: 10/22/2020
description: Определите метаданные JSON для пользовательских функций в Excel и свяжите свойства идентификатора и имени функции.
title: Создание метаданных JSON для пользовательских функций в Excel
localization_priority: Normal
ms.openlocfilehash: c676abc3115082fa861a4650b11869009f168e7f
ms.sourcegitcommit: a4e09546fd59579439025aca9cc58474b5ae7676
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/27/2020
ms.locfileid: "48774749"
---
# <a name="create-json-metadata-for-custom-functions"></a><span data-ttu-id="0dbae-103">Создание метаданных JSON для пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="0dbae-103">Create JSON metadata for custom functions</span></span>

<span data-ttu-id="0dbae-104">Как описано в статье [Обзор пользовательских функций](custom-functions-overview.md) , проект пользовательских функций должен включать файл метаданных JSON и файл скрипта (JavaScript или TypeScript) для регистрации функции, делая ее доступной для использования.</span><span class="sxs-lookup"><span data-stu-id="0dbae-104">As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include both a JSON metadata file and a script (either JavaScript or TypeScript) file to register a function, making it available for use.</span></span> <span data-ttu-id="0dbae-105">Пользовательские функции регистрируются при первом запуске надстройки и после их появления для одного и того же пользователя во всех книгах.</span><span class="sxs-lookup"><span data-stu-id="0dbae-105">Custom functions are registered when the user runs the add-in for the first time and after that are available to the same user in all workbooks.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="0dbae-106">Рекомендуется использовать автоматическое создание JSON, когда это возможно, вместо создания файла JSON.</span><span class="sxs-lookup"><span data-stu-id="0dbae-106">We recommend using JSON autogeneration when possible instead of creating your own JSON file.</span></span> <span data-ttu-id="0dbae-107">Автоматическое создание менее подвержено ошибкам пользователей, и в шаблоне `yo office` уже есть файл шаблона.</span><span class="sxs-lookup"><span data-stu-id="0dbae-107">Autogeneration is less prone to user error and the `yo office` scaffolded files already include this.</span></span> <span data-ttu-id="0dbae-108">Дополнительные сведения о тегах Жсдок и процессе автоматического формирования JSON приведены в статье Автоматическое [Создание МЕТАДАННЫХ JSON для пользовательских функций](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="0dbae-108">For more information on JSDoc tags and the JSON autogeneration process, see [Autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

<span data-ttu-id="0dbae-109">Тем не менее, вы можете сделать проект пользовательских функций с нуля, но необходимо выполнить следующие действия:</span><span class="sxs-lookup"><span data-stu-id="0dbae-109">However, you can make a custom functions project from scratch but it requires you to:</span></span>

- <span data-ttu-id="0dbae-110">Создайте файл JSON.</span><span class="sxs-lookup"><span data-stu-id="0dbae-110">Write your JSON file.</span></span>
- <span data-ttu-id="0dbae-111">Убедитесь, что файл манифеста подключен к файлу JSON.</span><span class="sxs-lookup"><span data-stu-id="0dbae-111">Check that your manifest file is connected to your JSON file.</span></span>
- <span data-ttu-id="0dbae-112">Свяжите функции `id` и `name` свойства в файле скрипта, чтобы зарегистрировать функции.</span><span class="sxs-lookup"><span data-stu-id="0dbae-112">Associate your functions' `id` and `name` properties in the script file in order to register your functions.</span></span>

<span data-ttu-id="0dbae-113">На следующем рисунке показано различие между `yo office` файлами формирования шаблонов и записью JSON с нуля.</span><span class="sxs-lookup"><span data-stu-id="0dbae-113">The following image explains the differences between using `yo office` scaffold files and writing JSON from scratch.</span></span>

![Изображение различий при использовании Yo Office и написании собственного JSON](../images/custom-functions-json.png)

> [!NOTE]
> <span data-ttu-id="0dbae-115">Не забудьте подключить манифест к созданному файлу JSON, используя `<Resources>` раздел XML-файла манифеста, если генератор не используется `yo office` .</span><span class="sxs-lookup"><span data-stu-id="0dbae-115">Remember to connect your manifest to the JSON file you create, through the `<Resources>` section in your XML manifest file if you do not use the `yo office` generator.</span></span>

## <a name="authoring-metadata-and-connecting-to-the-manifest"></a><span data-ttu-id="0dbae-116">Создание метаданных и подключение к манифесту</span><span class="sxs-lookup"><span data-stu-id="0dbae-116">Authoring metadata and connecting to the manifest</span></span>

<span data-ttu-id="0dbae-117">Создайте файл JSON в проекте и предоставьте все подробные сведения о функциях, таких как параметры функции.</span><span class="sxs-lookup"><span data-stu-id="0dbae-117">Create a JSON file in your project and provide all the details about your functions in it, such as the function's parameters.</span></span> <span data-ttu-id="0dbae-118">В [приведенном ниже примере метаданных](#json-metadata-example) и [справочнике по метаданным](#metadata-reference) представлен полный список свойств функций.</span><span class="sxs-lookup"><span data-stu-id="0dbae-118">See the [following metadata example](#json-metadata-example) and [the metadata reference](#metadata-reference) for a complete list of function properties.</span></span>

<span data-ttu-id="0dbae-119">Убедитесь, что XML-файл манифеста ссылается на JSON-файл в `<Resources>` разделе, как показано в следующем примере.</span><span class="sxs-lookup"><span data-stu-id="0dbae-119">Ensure your XML manifest file references your JSON file in the `<Resources>` section, similar to the following example.</span></span>

```json
<Resources>
    <bt:Urls>
        <bt:Url id="JSON-URL" DefaultValue="https://subdomain.contoso.com/config/customfunctions.json"/>
        <bt:Url id="JS-URL" DefaultValue="https://subdomain.contoso.com/dist/win32/ship/index.win32.bundle"/>
            <bt:Url id="HTML-URL" DefaultValue="https://subdomain.contoso.com/index.html"/>
    </bt:Urls>
    <bt:ShortStrings>
        <bt:String id="namespace" DefaultValue="CONTOSO"/>
    </bt:ShortStrings>
</Resources>
```

## <a name="json-metadata-example"></a><span data-ttu-id="0dbae-120">Пример метаданных JSON</span><span class="sxs-lookup"><span data-stu-id="0dbae-120">JSON metadata example</span></span>

<span data-ttu-id="0dbae-121">В примере кода ниже показано содержимое JSON-файла метаданных для надстройки, определяющей настраиваемые функции.</span><span class="sxs-lookup"><span data-stu-id="0dbae-121">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="0dbae-122">В следующих за этим примером разделах приводятся подробные сведения об отдельных свойствах, представленных в этом примере JSON.</span><span class="sxs-lookup"><span data-stu-id="0dbae-122">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

```json
{
  "functions": [
    {
      "id": "ADD",
      "name": "ADD",
      "description": "Add two numbers",
      "helpUrl": "http://www.contoso.com/help",
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
          "dimensionality": "scalar"
        }
      ]
    },
    {
      "id": "GETDAY",
      "name": "GETDAY",
      "description": "Get the day of the week",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
        "dimensionality": "scalar"
      },
      "parameters": []
    },
    {
      "id": "INCREMENTVALUE",
      "name": "INCREMENTVALUE",
      "description": "Count up from zero",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
        "dimensionality": "scalar"
      },
      "parameters": [
        {
          "name": "increment",
          "description": "the number to be added each time",
          "type": "number",
          "dimensionality": "scalar"
        }
      ],
      "options": {
        "stream": true,
        "cancelable": true
      }
    },
    {
      "id": "SECONDHIGHEST",
      "name": "SECONDHIGHEST",
      "description": "Get the second highest number from a range",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
        "dimensionality": "scalar"
      },
      "parameters": [
        {
          "name": "range",
          "description": "the input range",
          "type": "number",
          "dimensionality": "matrix"
        }
      ]
    }
  ]
}
```

> [!NOTE]
> <span data-ttu-id="0dbae-123">Полный пример JSON-файла доступен в журнале транзакций [OfficeDev/Excel-Custom-functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) репозитория GitHub.</span><span class="sxs-lookup"><span data-stu-id="0dbae-123">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) GitHub repository's commit history.</span></span> <span data-ttu-id="0dbae-124">Так как проект был скорректирован для автоматического создания JSON, полный пример рукописного кода JSON доступен только в предыдущих версиях проекта.</span><span class="sxs-lookup"><span data-stu-id="0dbae-124">As the project has been adjusted to automatically generate JSON, a full sample of handwritten JSON is only available in previous versions of the project.</span></span>

## <a name="metadata-reference"></a><span data-ttu-id="0dbae-125">Справка по метаданным</span><span class="sxs-lookup"><span data-stu-id="0dbae-125">Metadata reference</span></span>

### <a name="functions"></a><span data-ttu-id="0dbae-126">functions</span><span class="sxs-lookup"><span data-stu-id="0dbae-126">functions</span></span>

<span data-ttu-id="0dbae-127">Свойство `functions` представляет собой массив объектов настраиваемых функций.</span><span class="sxs-lookup"><span data-stu-id="0dbae-127">The `functions` property is an array of custom function objects.</span></span> <span data-ttu-id="0dbae-128">В таблице ниже приведены свойства каждого объекта.</span><span class="sxs-lookup"><span data-stu-id="0dbae-128">The following table lists the properties of each object.</span></span>

| <span data-ttu-id="0dbae-129">Свойство</span><span class="sxs-lookup"><span data-stu-id="0dbae-129">Property</span></span>      | <span data-ttu-id="0dbae-130">Тип данных</span><span class="sxs-lookup"><span data-stu-id="0dbae-130">Data type</span></span> | <span data-ttu-id="0dbae-131">Обязательный</span><span class="sxs-lookup"><span data-stu-id="0dbae-131">Required</span></span> | <span data-ttu-id="0dbae-132">Описание</span><span class="sxs-lookup"><span data-stu-id="0dbae-132">Description</span></span>                                                                                                                                                                      |
| :------------ | :-------- | :------- | :------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `description` | <span data-ttu-id="0dbae-133">string</span><span class="sxs-lookup"><span data-stu-id="0dbae-133">string</span></span>    | <span data-ttu-id="0dbae-134">Нет</span><span class="sxs-lookup"><span data-stu-id="0dbae-134">No</span></span>       | <span data-ttu-id="0dbae-135">Описание функции, которое отображается пользователям в Excel</span><span class="sxs-lookup"><span data-stu-id="0dbae-135">The description of the function that end users see in Excel.</span></span> <span data-ttu-id="0dbae-136">(например, **преобразует значение по шкале Цельсия в температуру по шкале Фаренгейта** ).</span><span class="sxs-lookup"><span data-stu-id="0dbae-136">For example, **Converts a Celsius value to Fahrenheit** .</span></span>                                                            |
| `helpUrl`     | <span data-ttu-id="0dbae-137">string</span><span class="sxs-lookup"><span data-stu-id="0dbae-137">string</span></span>    | <span data-ttu-id="0dbae-138">Нет</span><span class="sxs-lookup"><span data-stu-id="0dbae-138">No</span></span>       | <span data-ttu-id="0dbae-139">URL-адрес, по которому можно получить сведения о функции</span><span class="sxs-lookup"><span data-stu-id="0dbae-139">URL that provides information about the function.</span></span> <span data-ttu-id="0dbae-140">(отображается в области задач). Пример: `http://contoso.com/help/convertcelsiustofahrenheit.html`.</span><span class="sxs-lookup"><span data-stu-id="0dbae-140">(It is displayed in a task pane.) For example, `http://contoso.com/help/convertcelsiustofahrenheit.html`.</span></span>                      |
| `id`          | <span data-ttu-id="0dbae-141">string</span><span class="sxs-lookup"><span data-stu-id="0dbae-141">string</span></span>    | <span data-ttu-id="0dbae-142">Да</span><span class="sxs-lookup"><span data-stu-id="0dbae-142">Yes</span></span>      | <span data-ttu-id="0dbae-143">Уникальный идентификатор для функции.</span><span class="sxs-lookup"><span data-stu-id="0dbae-143">A unique ID for the function.</span></span> <span data-ttu-id="0dbae-144">Этот идентификатор может содержать только буквы, цифры и точки и не может изменяться после настройки.</span><span class="sxs-lookup"><span data-stu-id="0dbae-144">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span>                                            |
| `name`        | <span data-ttu-id="0dbae-145">string</span><span class="sxs-lookup"><span data-stu-id="0dbae-145">string</span></span>    | <span data-ttu-id="0dbae-146">Да</span><span class="sxs-lookup"><span data-stu-id="0dbae-146">Yes</span></span>      | <span data-ttu-id="0dbae-147">Имя функции, которое отображается пользователям в Excel.</span><span class="sxs-lookup"><span data-stu-id="0dbae-147">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="0dbae-148">В Excel это имя функции предваряется пространством имен пользовательских функций, указанным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="0dbae-148">In Excel, this function name is prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
| `options`     | <span data-ttu-id="0dbae-149">object</span><span class="sxs-lookup"><span data-stu-id="0dbae-149">object</span></span>    | <span data-ttu-id="0dbae-150">Нет</span><span class="sxs-lookup"><span data-stu-id="0dbae-150">No</span></span>       | <span data-ttu-id="0dbae-151">Позволяет настроить некоторые аспекты того, как и когда Excel выполняет функцию.</span><span class="sxs-lookup"><span data-stu-id="0dbae-151">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="0dbae-152">Дополнительные сведения см. в разделе [options](#options).</span><span class="sxs-lookup"><span data-stu-id="0dbae-152">See [options](#options) for details.</span></span>                                                          |
| `parameters`  | <span data-ttu-id="0dbae-153">array</span><span class="sxs-lookup"><span data-stu-id="0dbae-153">array</span></span>     | <span data-ttu-id="0dbae-154">Да</span><span class="sxs-lookup"><span data-stu-id="0dbae-154">Yes</span></span>      | <span data-ttu-id="0dbae-155">Массив, который определяет входные параметры для функции.</span><span class="sxs-lookup"><span data-stu-id="0dbae-155">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="0dbae-156">Дополнительные сведения см. в разделе [Parameters](#parameters) .</span><span class="sxs-lookup"><span data-stu-id="0dbae-156">See [parameters](#parameters) for details.</span></span>                                                                             |
| `result`      | <span data-ttu-id="0dbae-157">object</span><span class="sxs-lookup"><span data-stu-id="0dbae-157">object</span></span>    | <span data-ttu-id="0dbae-158">Да</span><span class="sxs-lookup"><span data-stu-id="0dbae-158">Yes</span></span>      | <span data-ttu-id="0dbae-159">Объект, который определяет тип информации, возвращаемый функцией.</span><span class="sxs-lookup"><span data-stu-id="0dbae-159">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="0dbae-160">Дополнительные сведения см. в разделе [result](#result).</span><span class="sxs-lookup"><span data-stu-id="0dbae-160">See [result](#result) for details.</span></span>                                                                 |

### <a name="options"></a><span data-ttu-id="0dbae-161">options</span><span class="sxs-lookup"><span data-stu-id="0dbae-161">options</span></span>

<span data-ttu-id="0dbae-162">Объект `options` позволяет настроить некоторые аспекты того, как и когда Excel выполняет функцию.</span><span class="sxs-lookup"><span data-stu-id="0dbae-162">The `options` object enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="0dbae-163">В таблице ниже приведены свойства объекта `options`.</span><span class="sxs-lookup"><span data-stu-id="0dbae-163">The following table lists the properties of the `options` object.</span></span>

| <span data-ttu-id="0dbae-164">Свойство</span><span class="sxs-lookup"><span data-stu-id="0dbae-164">Property</span></span>          | <span data-ttu-id="0dbae-165">Тип данных</span><span class="sxs-lookup"><span data-stu-id="0dbae-165">Data type</span></span> | <span data-ttu-id="0dbae-166">Обязательный</span><span class="sxs-lookup"><span data-stu-id="0dbae-166">Required</span></span>                               | <span data-ttu-id="0dbae-167">Описание</span><span class="sxs-lookup"><span data-stu-id="0dbae-167">Description</span></span> |
| :---------------- | :-------- | :------------------------------------- | :---------- |
| `cancelable`      | <span data-ttu-id="0dbae-168">boolean</span><span class="sxs-lookup"><span data-stu-id="0dbae-168">boolean</span></span>   | <span data-ttu-id="0dbae-169">Нет</span><span class="sxs-lookup"><span data-stu-id="0dbae-169">No</span></span><br/><br/><span data-ttu-id="0dbae-170">Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="0dbae-170">Default value is `false`.</span></span>  | <span data-ttu-id="0dbae-171">Если это свойство имеет значение `true`, Excel будет вызывать обработчик `CancelableInvocation` каждый раз, когда пользователь будет предпринимать действия, которые приводят к отмене функции (например, вручную вызывает пересчет или редактирует ячейку, на которую ссылается функция).</span><span class="sxs-lookup"><span data-stu-id="0dbae-171">If `true`, Excel calls the `CancelableInvocation` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="0dbae-172">Функции, которые можно отменять, обычно используются только для асинхронных функций, которые возвращают один результат и нуждаются в обработке отмены запроса данных.</span><span class="sxs-lookup"><span data-stu-id="0dbae-172">Cancelable functions are typically only used for asynchronous functions that return a single result and need to handle the cancellation of a request for data.</span></span> <span data-ttu-id="0dbae-173">Функция не может быть одновременно потоковой и отмены.</span><span class="sxs-lookup"><span data-stu-id="0dbae-173">A function cannot be both streaming and cancelable.</span></span> <span data-ttu-id="0dbae-174">Более подробную информацию можно найти в заметке около конца [функции потоковой передачи](custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="0dbae-174">For more information, see the note near the end of [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> |
| `requiresAddress` | <span data-ttu-id="0dbae-175">boolean</span><span class="sxs-lookup"><span data-stu-id="0dbae-175">boolean</span></span>   | <span data-ttu-id="0dbae-176">Нет</span><span class="sxs-lookup"><span data-stu-id="0dbae-176">No</span></span> <br/><br/><span data-ttu-id="0dbae-177">Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="0dbae-177">Default value is `false`.</span></span> | <span data-ttu-id="0dbae-178">Если `true` Пользовательская функция может получить доступ к адресу ячейки, которая вызвала пользовательскую функцию.</span><span class="sxs-lookup"><span data-stu-id="0dbae-178">If `true`, your custom function can access the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="0dbae-179">Чтобы получить адрес ячейки, которая вызвала пользовательскую функцию, используйте context. Address в пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="0dbae-179">To get the address of the cell that invoked your custom function, use context.address in your custom function.</span></span> <span data-ttu-id="0dbae-180">Пользовательские функции не могут быть заданы как потоковые, так и Рекуиресаддресс.</span><span class="sxs-lookup"><span data-stu-id="0dbae-180">Custom functions cannot be set as both streaming and requiresAddress.</span></span> <span data-ttu-id="0dbae-181">При использовании этого параметра параметр "вызов" должен быть последним параметром, переданным в параметрах.</span><span class="sxs-lookup"><span data-stu-id="0dbae-181">When using this option, the 'invocation' parameter must be the last parameter passed in options.</span></span> |
| `stream`          | <span data-ttu-id="0dbae-182">boolean</span><span class="sxs-lookup"><span data-stu-id="0dbae-182">boolean</span></span>   | <span data-ttu-id="0dbae-183">Нет</span><span class="sxs-lookup"><span data-stu-id="0dbae-183">No</span></span><br/><br/><span data-ttu-id="0dbae-184">Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="0dbae-184">Default value is `false`.</span></span>  | <span data-ttu-id="0dbae-185">Если это свойство имеет значение `true`, функция может выводить значение в ячейку несколько раз, даже если вызвана всего единожды.</span><span class="sxs-lookup"><span data-stu-id="0dbae-185">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="0dbae-186">Этот параметр полезен для быстро изменяющихся источников данных, таких как цена акций.</span><span class="sxs-lookup"><span data-stu-id="0dbae-186">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="0dbae-187">Функция не должна содержать оператор `return`.</span><span class="sxs-lookup"><span data-stu-id="0dbae-187">The function should have no `return` statement.</span></span> <span data-ttu-id="0dbae-188">Вместо этого результирующее значение передается как аргумент метода обратного вызова `StreamingInvocation.setResult`.</span><span class="sxs-lookup"><span data-stu-id="0dbae-188">Instead, the result value is passed as the argument of the `StreamingInvocation.setResult` callback method.</span></span> <span data-ttu-id="0dbae-189">Дополнительные сведения см. в разделе [Потоковые функции](custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="0dbae-189">For more information, see [Streaming functions](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> |
| `volatile`        | <span data-ttu-id="0dbae-190">boolean</span><span class="sxs-lookup"><span data-stu-id="0dbae-190">boolean</span></span>   | <span data-ttu-id="0dbae-191">Нет</span><span class="sxs-lookup"><span data-stu-id="0dbae-191">No</span></span> <br/><br/><span data-ttu-id="0dbae-192">Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="0dbae-192">Default value is `false`.</span></span> | <span data-ttu-id="0dbae-193">Если `true` функция пересчитывает при каждом пересчете Excel, функция пересчитывается, а не только при изменении зависимых значений формулы.</span><span class="sxs-lookup"><span data-stu-id="0dbae-193">If `true`, the function recalculates each time Excel recalculates, instead of only when the formula's dependent values have changed.</span></span> <span data-ttu-id="0dbae-194">Функция не может быть одновременно потоковой и переменной.</span><span class="sxs-lookup"><span data-stu-id="0dbae-194">A function cannot be both streaming and volatile.</span></span> <span data-ttu-id="0dbae-195">Если обоим свойствам `stream` и `volatile` присвоено значение `true`, параметр переменности будет игнорироваться.</span><span class="sxs-lookup"><span data-stu-id="0dbae-195">If the `stream` and `volatile` properties are both set to `true`, the volatile option will be ignored.</span></span> |

### <a name="parameters"></a><span data-ttu-id="0dbae-196">parameters</span><span class="sxs-lookup"><span data-stu-id="0dbae-196">parameters</span></span>

<span data-ttu-id="0dbae-197">Свойство `parameters` представляет собой массив объектов параметров.</span><span class="sxs-lookup"><span data-stu-id="0dbae-197">The `parameters` property is an array of parameter objects.</span></span> <span data-ttu-id="0dbae-198">В таблице ниже приведены свойства каждого объекта.</span><span class="sxs-lookup"><span data-stu-id="0dbae-198">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="0dbae-199">Свойство</span><span class="sxs-lookup"><span data-stu-id="0dbae-199">Property</span></span>  |  <span data-ttu-id="0dbae-200">Тип данных</span><span class="sxs-lookup"><span data-stu-id="0dbae-200">Data type</span></span>  |  <span data-ttu-id="0dbae-201">Обязательный</span><span class="sxs-lookup"><span data-stu-id="0dbae-201">Required</span></span>  |  <span data-ttu-id="0dbae-202">Описание</span><span class="sxs-lookup"><span data-stu-id="0dbae-202">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="0dbae-203">string</span><span class="sxs-lookup"><span data-stu-id="0dbae-203">string</span></span>  |  <span data-ttu-id="0dbae-204">Нет</span><span class="sxs-lookup"><span data-stu-id="0dbae-204">No</span></span> |  <span data-ttu-id="0dbae-205">Описание параметра.</span><span class="sxs-lookup"><span data-stu-id="0dbae-205">A description of the parameter.</span></span> <span data-ttu-id="0dbae-206">Это отображается в IntelliSense Excel.</span><span class="sxs-lookup"><span data-stu-id="0dbae-206">This is displayed in Excel's IntelliSense.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="0dbae-207">string</span><span class="sxs-lookup"><span data-stu-id="0dbae-207">string</span></span>  |  <span data-ttu-id="0dbae-208">Нет</span><span class="sxs-lookup"><span data-stu-id="0dbae-208">No</span></span>  |  <span data-ttu-id="0dbae-209">Должно быть **скалярным** (значение, отличное от массива) или **матричным** (двухмерный массив).</span><span class="sxs-lookup"><span data-stu-id="0dbae-209">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="0dbae-210">string</span><span class="sxs-lookup"><span data-stu-id="0dbae-210">string</span></span>  |  <span data-ttu-id="0dbae-211">Да</span><span class="sxs-lookup"><span data-stu-id="0dbae-211">Yes</span></span>  |  <span data-ttu-id="0dbae-212">Имя параметра.</span><span class="sxs-lookup"><span data-stu-id="0dbae-212">The name of the parameter.</span></span> <span data-ttu-id="0dbae-213">Это имя отображается в IntelliSense Excel.</span><span class="sxs-lookup"><span data-stu-id="0dbae-213">This name is displayed in Excel's IntelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="0dbae-214">string</span><span class="sxs-lookup"><span data-stu-id="0dbae-214">string</span></span>  |  <span data-ttu-id="0dbae-215">Нет</span><span class="sxs-lookup"><span data-stu-id="0dbae-215">No</span></span>  |  <span data-ttu-id="0dbae-216">Тип данных параметра.</span><span class="sxs-lookup"><span data-stu-id="0dbae-216">The data type of the parameter.</span></span> <span data-ttu-id="0dbae-217">Может иметь значение **boolean** , **number** , **string** или **any** , что позволяет использовать любой из трех предыдущих типов.</span><span class="sxs-lookup"><span data-stu-id="0dbae-217">Can be **boolean** , **number** , **string** , or **any** , which allows you to use of any of the previous three types.</span></span> <span data-ttu-id="0dbae-218">Если это свойство не задано, по умолчанию устанавливается тип данных **any** .</span><span class="sxs-lookup"><span data-stu-id="0dbae-218">If this property is not specified, the data type defaults to **any** .</span></span> |
|  `optional`  | <span data-ttu-id="0dbae-219">boolean</span><span class="sxs-lookup"><span data-stu-id="0dbae-219">boolean</span></span> | <span data-ttu-id="0dbae-220">Нет</span><span class="sxs-lookup"><span data-stu-id="0dbae-220">No</span></span> | <span data-ttu-id="0dbae-221">Если присвоено значение `true`, параметр не обязателен.</span><span class="sxs-lookup"><span data-stu-id="0dbae-221">If `true`, the parameter is optional.</span></span> |
|`repeating`| <span data-ttu-id="0dbae-222">boolean</span><span class="sxs-lookup"><span data-stu-id="0dbae-222">boolean</span></span> | <span data-ttu-id="0dbae-223">Нет</span><span class="sxs-lookup"><span data-stu-id="0dbae-223">No</span></span> | <span data-ttu-id="0dbae-224">Если `true` параметры заполняются из указанного массива.</span><span class="sxs-lookup"><span data-stu-id="0dbae-224">If `true`, parameters populate from a specified array.</span></span> <span data-ttu-id="0dbae-225">Обратите внимание, что функции все повторяющиеся параметры считаются необязательными параметрами по определению.</span><span class="sxs-lookup"><span data-stu-id="0dbae-225">Note that functions all repeating parameters are considered optional parameters by definition.</span></span>  |

### <a name="result"></a><span data-ttu-id="0dbae-226">result</span><span class="sxs-lookup"><span data-stu-id="0dbae-226">result</span></span>

<span data-ttu-id="0dbae-227">Объект `result` определяет тип информации, возвращаемый функцией.</span><span class="sxs-lookup"><span data-stu-id="0dbae-227">The `result` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="0dbae-228">В таблице ниже приведены свойства объекта `result`.</span><span class="sxs-lookup"><span data-stu-id="0dbae-228">The following table lists the properties of the `result` object.</span></span>

| <span data-ttu-id="0dbae-229">Свойство</span><span class="sxs-lookup"><span data-stu-id="0dbae-229">Property</span></span>         | <span data-ttu-id="0dbae-230">Тип данных</span><span class="sxs-lookup"><span data-stu-id="0dbae-230">Data type</span></span> | <span data-ttu-id="0dbae-231">Обязательный</span><span class="sxs-lookup"><span data-stu-id="0dbae-231">Required</span></span> | <span data-ttu-id="0dbae-232">Описание</span><span class="sxs-lookup"><span data-stu-id="0dbae-232">Description</span></span>                                                                          |
| :--------------- | :-------- | :------- | :----------------------------------------------------------------------------------- |
| `dimensionality` | <span data-ttu-id="0dbae-233">string</span><span class="sxs-lookup"><span data-stu-id="0dbae-233">string</span></span>    | <span data-ttu-id="0dbae-234">Нет</span><span class="sxs-lookup"><span data-stu-id="0dbae-234">No</span></span>       | <span data-ttu-id="0dbae-235">Должно быть **скалярным** (значение, отличное от массива) или **матричным** (двухмерный массив).</span><span class="sxs-lookup"><span data-stu-id="0dbae-235">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |

## <a name="associating-function-names-with-json-metadata"></a><span data-ttu-id="0dbae-236">Сопоставление имен функций с метаданными JSON</span><span class="sxs-lookup"><span data-stu-id="0dbae-236">Associating function names with JSON metadata</span></span>

<span data-ttu-id="0dbae-237">Чтобы функция работала должным образом, необходимо связать `id` свойство функции с реализацией JavaScript.</span><span class="sxs-lookup"><span data-stu-id="0dbae-237">For a function to work properly, you need to associate the function's `id` property with the JavaScript implementation.</span></span> <span data-ttu-id="0dbae-238">Убедитесь, что существует связь, в противном случае функция не будет зарегистрирована и непригодна для работы в Excel.</span><span class="sxs-lookup"><span data-stu-id="0dbae-238">Make sure there is an association, otherwise the function won't be registered and isn't useable in Excel.</span></span> <span data-ttu-id="0dbae-239">В приведенном ниже примере кода показано, как выполнить связь с помощью `CustomFunctions.associate()` метода.</span><span class="sxs-lookup"><span data-stu-id="0dbae-239">The following code sample shows how to make the association using the `CustomFunctions.associate()` method.</span></span> <span data-ttu-id="0dbae-240">Пример определяет пользовательскую функцию `add` и связывает ее с объектом в файле метаданных JSON, где для свойства `id` установлено значение **ADD** .</span><span class="sxs-lookup"><span data-stu-id="0dbae-240">The sample defines the custom function `add` and associates it with the object in the JSON metadata file where the value of the `id` property is **ADD** .</span></span>

```js
/**
 * Add two numbers
 * @customfunction
 * @param {number} first First number
 * @param {number} second Second number
 * @returns {number} The sum of the two numbers.
 */
function add(first, second) {
  return first + second;
}

CustomFunctions.associate("ADD", add);
```

<span data-ttu-id="0dbae-241">В следующем JSON показаны метаданные JSON, связанные с предыдущим кодом пользовательской функции JavaScript.</span><span class="sxs-lookup"><span data-stu-id="0dbae-241">The following JSON shows the JSON metadata that is associated with the previous custom function JavaScript code.</span></span>

```json
{
  "functions": [
    {
      "description": "Add two numbers",
      "id": "ADD",
      "name": "ADD",
      "parameters": [
        {
          "description": "First number",
          "name": "first",
          "type": "number"
        },
        {
          "description": "Second number",
          "name": "second",
          "type": "number"
        }
      ],
      "result": {
        "type": "number"
      }
    }
  ]
}
```

<span data-ttu-id="0dbae-242">Имейте в виду приведенные ниже рекомендации при создании пользовательских функций в файле JavaScript и указании соответствующих сведений в файле метаданных JSON.</span><span class="sxs-lookup"><span data-stu-id="0dbae-242">Keep in mind the following best practices when creating custom functions in your JavaScript file and specifying corresponding information in the JSON metadata file.</span></span>

- <span data-ttu-id="0dbae-243">Убедитесь, что в файле метаданных JSON значение каждого свойства `id` содержит только буквы, цифры и точки.</span><span class="sxs-lookup"><span data-stu-id="0dbae-243">In the JSON metadata file, ensure that the value of each `id` property contains only alphanumeric characters and periods.</span></span>

- <span data-ttu-id="0dbae-244">Убедитесь, что в файле метаданных JSON значение каждого свойства `id` уникально в пределах файла.</span><span class="sxs-lookup"><span data-stu-id="0dbae-244">In the JSON metadata file, ensure that the value of each `id` property is unique within the scope of the file.</span></span> <span data-ttu-id="0dbae-245">То есть никакие два объекта функций в файле метаданных не должны иметь одинаковое значение `id`.</span><span class="sxs-lookup"><span data-stu-id="0dbae-245">That is, no two function objects in the metadata file should have the same `id` value.</span></span>

- <span data-ttu-id="0dbae-246">Не изменяйте значение свойства `id` в файле метаданных JSON после его сопоставления с соответствующим именем функции JavaScript.</span><span class="sxs-lookup"><span data-stu-id="0dbae-246">Do not change the value of an `id` property in the JSON metadata file after it's been associated with a corresponding JavaScript function name.</span></span> <span data-ttu-id="0dbae-247">Вы можете изменить имя функции, которое отображается для конечных пользователей в Excel, путем обновления свойства `name` в файле метаданных JSON, но никогда не следует изменять значение свойства `id` после его установления.</span><span class="sxs-lookup"><span data-stu-id="0dbae-247">You can change the function name that end users see in Excel by updating the `name` property within the JSON metadata file, but you should never change the value of an `id` property after it's been established.</span></span>

- <span data-ttu-id="0dbae-248">В файле JavaScript укажите настраиваемое сопоставление функций с помощью `CustomFunctions.associate` каждой функции.</span><span class="sxs-lookup"><span data-stu-id="0dbae-248">In the JavaScript file, specify a custom function association using `CustomFunctions.associate` after each function.</span></span>

<span data-ttu-id="0dbae-249">В следующем примере показаны метаданные JSON, которые соответствуют функциям, определенным в предыдущем примере кода JavaScript.</span><span class="sxs-lookup"><span data-stu-id="0dbae-249">The following sample shows the JSON metadata that corresponds to the functions defined in the preceding JavaScript code sample.</span></span> <span data-ttu-id="0dbae-250">`id` `name` Значения свойств и представлены в верхнем регистре, что является лучшим вариантом при описании пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="0dbae-250">The `id` and `name` property values are in uppercase, which is a best practice when describing your custom functions.</span></span> <span data-ttu-id="0dbae-251">Этот код JSON необходимо добавить только в том случае, если вы готовите собственный файл JSON вручную и не используете автоматическое создание.</span><span class="sxs-lookup"><span data-stu-id="0dbae-251">You only need to add this JSON if you are preparing your own JSON file manually and not using autogeneration.</span></span> <span data-ttu-id="0dbae-252">Дополнительные сведения об автоформировании приведены в статье Автоматическое [Создание МЕТАДАННЫХ JSON для пользовательских функций](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="0dbae-252">For more information on autogeneration, see [Autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

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

## <a name="next-steps"></a><span data-ttu-id="0dbae-253">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="0dbae-253">Next steps</span></span>

<span data-ttu-id="0dbae-254">Ознакомьтесь с рекомендациями [по именованию функции](custom-functions-naming.md) или [локализации функции](custom-functions-localize.md) с помощью ранее описанного рукописного метода JSON.</span><span class="sxs-lookup"><span data-stu-id="0dbae-254">Learn the [best practices for naming your function](custom-functions-naming.md) or discover how to [localize your function](custom-functions-localize.md) using the previously described handwritten JSON method.</span></span>

## <a name="see-also"></a><span data-ttu-id="0dbae-255">См. также</span><span class="sxs-lookup"><span data-stu-id="0dbae-255">See also</span></span>

- [<span data-ttu-id="0dbae-256">Автоматическое генерирование метаданных JSON для пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="0dbae-256">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
- [<span data-ttu-id="0dbae-257">Параметры параметров пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="0dbae-257">Custom functions parameter options</span></span>](custom-functions-parameter-options.md)
- [<span data-ttu-id="0dbae-258">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="0dbae-258">Create custom functions in Excel</span></span>](custom-functions-overview.md)
