---
ms.date: 11/06/2020
description: Определите метаданные JSON для пользовательских функций в Excel и свяжите свойства идентификатора и имени функции.
title: Создание метаданных JSON для пользовательских функций вручную в Excel
localization_priority: Normal
ms.openlocfilehash: adbcbb9d2705a38b1ed9ff5cdffa6162b9d93a9c
ms.sourcegitcommit: 5bfd1e9956485c140179dfcc9d210c4c5a49a789
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/13/2020
ms.locfileid: "49071643"
---
# <a name="manually-create-json-metadata-for-custom-functions"></a><span data-ttu-id="214fb-103">Создание метаданных JSON для пользовательских функций вручную</span><span class="sxs-lookup"><span data-stu-id="214fb-103">Manually create JSON metadata for custom functions</span></span>

<span data-ttu-id="214fb-104">Как описано в статье [Обзор пользовательских функций](custom-functions-overview.md) , проект пользовательских функций должен включать файл метаданных JSON и файл скрипта (JavaScript или TypeScript) для регистрации функции, делая ее доступной для использования.</span><span class="sxs-lookup"><span data-stu-id="214fb-104">As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include both a JSON metadata file and a script (either JavaScript or TypeScript) file to register a function, making it available for use.</span></span> <span data-ttu-id="214fb-105">Пользовательские функции регистрируются при первом запуске надстройки и после их появления для одного и того же пользователя во всех книгах.</span><span class="sxs-lookup"><span data-stu-id="214fb-105">Custom functions are registered when the user runs the add-in for the first time and after that are available to the same user in all workbooks.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="214fb-106">Рекомендуется использовать автоматическое создание JSON, когда это возможно, вместо создания файла JSON.</span><span class="sxs-lookup"><span data-stu-id="214fb-106">We recommend using JSON autogeneration when possible instead of creating your own JSON file.</span></span> <span data-ttu-id="214fb-107">Автоматическое создание менее подвержено ошибкам пользователей, и в шаблоне `yo office` уже есть файл шаблона.</span><span class="sxs-lookup"><span data-stu-id="214fb-107">Autogeneration is less prone to user error and the `yo office` scaffolded files already include this.</span></span> <span data-ttu-id="214fb-108">Дополнительные сведения о тегах Жсдок и процессе автоматического формирования JSON приведены в статье Автоматическое [Создание МЕТАДАННЫХ JSON для пользовательских функций](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="214fb-108">For more information on JSDoc tags and the JSON autogeneration process, see [Autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

<span data-ttu-id="214fb-109">Тем не менее, проект настраиваемых функций можно сделать с нуля.</span><span class="sxs-lookup"><span data-stu-id="214fb-109">However, you can make a custom functions project from scratch.</span></span> <span data-ttu-id="214fb-110">Этот процесс требует выполнения следующих действий:</span><span class="sxs-lookup"><span data-stu-id="214fb-110">This process requires you to:</span></span>

- <span data-ttu-id="214fb-111">Создайте файл JSON.</span><span class="sxs-lookup"><span data-stu-id="214fb-111">Write your JSON file.</span></span>
- <span data-ttu-id="214fb-112">Убедитесь, что файл манифеста подключен к файлу JSON.</span><span class="sxs-lookup"><span data-stu-id="214fb-112">Check that your manifest file is connected to your JSON file.</span></span>
- <span data-ttu-id="214fb-113">Свяжите функции `id` и `name` свойства в файле скрипта, чтобы зарегистрировать функции.</span><span class="sxs-lookup"><span data-stu-id="214fb-113">Associate your functions' `id` and `name` properties in the script file in order to register your functions.</span></span>

<span data-ttu-id="214fb-114">На следующем рисунке показано различие между `yo office` файлами формирования шаблонов и записью JSON с нуля.</span><span class="sxs-lookup"><span data-stu-id="214fb-114">The following image explains the differences between using `yo office` scaffold files and writing JSON from scratch.</span></span>

![Изображение различий при использовании Yo Office и написании собственного JSON](../images/custom-functions-json.png)

> [!NOTE]
> <span data-ttu-id="214fb-116">Не забудьте подключить манифест к созданному файлу JSON, используя `<Resources>` раздел XML-файла манифеста, если генератор не используется `yo office` .</span><span class="sxs-lookup"><span data-stu-id="214fb-116">Remember to connect your manifest to the JSON file you create, through the `<Resources>` section in your XML manifest file if you do not use the `yo office` generator.</span></span>

## <a name="authoring-metadata-and-connecting-to-the-manifest"></a><span data-ttu-id="214fb-117">Создание метаданных и подключение к манифесту</span><span class="sxs-lookup"><span data-stu-id="214fb-117">Authoring metadata and connecting to the manifest</span></span>

<span data-ttu-id="214fb-118">Создайте файл JSON в проекте и предоставьте все подробные сведения о функциях, таких как параметры функции.</span><span class="sxs-lookup"><span data-stu-id="214fb-118">Create a JSON file in your project and provide all the details about your functions in it, such as the function's parameters.</span></span> <span data-ttu-id="214fb-119">В [приведенном ниже примере метаданных](#json-metadata-example) и [справочнике по метаданным](#metadata-reference) представлен полный список свойств функций.</span><span class="sxs-lookup"><span data-stu-id="214fb-119">See the [following metadata example](#json-metadata-example) and [the metadata reference](#metadata-reference) for a complete list of function properties.</span></span>

<span data-ttu-id="214fb-120">Убедитесь, что XML-файл манифеста ссылается на JSON-файл в `<Resources>` разделе, как показано в следующем примере.</span><span class="sxs-lookup"><span data-stu-id="214fb-120">Ensure your XML manifest file references your JSON file in the `<Resources>` section, similar to the following example.</span></span>

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

## <a name="json-metadata-example"></a><span data-ttu-id="214fb-121">Пример метаданных JSON</span><span class="sxs-lookup"><span data-stu-id="214fb-121">JSON metadata example</span></span>

<span data-ttu-id="214fb-122">В примере кода ниже показано содержимое JSON-файла метаданных для надстройки, определяющей настраиваемые функции.</span><span class="sxs-lookup"><span data-stu-id="214fb-122">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="214fb-123">В следующих за этим примером разделах приводятся подробные сведения об отдельных свойствах, представленных в этом примере JSON.</span><span class="sxs-lookup"><span data-stu-id="214fb-123">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
> <span data-ttu-id="214fb-124">Полный пример JSON-файла доступен в журнале транзакций [OfficeDev/Excel-Custom-functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) репозитория GitHub.</span><span class="sxs-lookup"><span data-stu-id="214fb-124">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) GitHub repository's commit history.</span></span> <span data-ttu-id="214fb-125">Так как проект был скорректирован для автоматического создания JSON, полный пример рукописного кода JSON доступен только в предыдущих версиях проекта.</span><span class="sxs-lookup"><span data-stu-id="214fb-125">As the project has been adjusted to automatically generate JSON, a full sample of handwritten JSON is only available in previous versions of the project.</span></span>

## <a name="metadata-reference"></a><span data-ttu-id="214fb-126">Справка по метаданным</span><span class="sxs-lookup"><span data-stu-id="214fb-126">Metadata reference</span></span>

### <a name="functions"></a><span data-ttu-id="214fb-127">functions</span><span class="sxs-lookup"><span data-stu-id="214fb-127">functions</span></span>

<span data-ttu-id="214fb-128">Свойство `functions` представляет собой массив объектов настраиваемых функций.</span><span class="sxs-lookup"><span data-stu-id="214fb-128">The `functions` property is an array of custom function objects.</span></span> <span data-ttu-id="214fb-129">В таблице ниже приведены свойства каждого объекта.</span><span class="sxs-lookup"><span data-stu-id="214fb-129">The following table lists the properties of each object.</span></span>

| <span data-ttu-id="214fb-130">Свойство</span><span class="sxs-lookup"><span data-stu-id="214fb-130">Property</span></span>      | <span data-ttu-id="214fb-131">Тип данных</span><span class="sxs-lookup"><span data-stu-id="214fb-131">Data type</span></span> | <span data-ttu-id="214fb-132">Обязательный</span><span class="sxs-lookup"><span data-stu-id="214fb-132">Required</span></span> | <span data-ttu-id="214fb-133">Описание</span><span class="sxs-lookup"><span data-stu-id="214fb-133">Description</span></span>                                                                                                                                                                      |
| :------------ | :-------- | :------- | :------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `description` | <span data-ttu-id="214fb-134">строка</span><span class="sxs-lookup"><span data-stu-id="214fb-134">string</span></span>    | <span data-ttu-id="214fb-135">Нет</span><span class="sxs-lookup"><span data-stu-id="214fb-135">No</span></span>       | <span data-ttu-id="214fb-136">Описание функции, которое отображается пользователям в Excel</span><span class="sxs-lookup"><span data-stu-id="214fb-136">The description of the function that end users see in Excel.</span></span> <span data-ttu-id="214fb-137">(например, **преобразует значение по шкале Цельсия в температуру по шкале Фаренгейта** ).</span><span class="sxs-lookup"><span data-stu-id="214fb-137">For example, **Converts a Celsius value to Fahrenheit**.</span></span>                                                            |
| `helpUrl`     | <span data-ttu-id="214fb-138">string</span><span class="sxs-lookup"><span data-stu-id="214fb-138">string</span></span>    | <span data-ttu-id="214fb-139">Нет</span><span class="sxs-lookup"><span data-stu-id="214fb-139">No</span></span>       | <span data-ttu-id="214fb-140">URL-адрес, по которому можно получить сведения о функции</span><span class="sxs-lookup"><span data-stu-id="214fb-140">URL that provides information about the function.</span></span> <span data-ttu-id="214fb-141">(отображается в области задач). Пример: `http://contoso.com/help/convertcelsiustofahrenheit.html`.</span><span class="sxs-lookup"><span data-stu-id="214fb-141">(It is displayed in a task pane.) For example, `http://contoso.com/help/convertcelsiustofahrenheit.html`.</span></span>                      |
| `id`          | <span data-ttu-id="214fb-142">string</span><span class="sxs-lookup"><span data-stu-id="214fb-142">string</span></span>    | <span data-ttu-id="214fb-143">Да</span><span class="sxs-lookup"><span data-stu-id="214fb-143">Yes</span></span>      | <span data-ttu-id="214fb-144">Уникальный идентификатор для функции.</span><span class="sxs-lookup"><span data-stu-id="214fb-144">A unique ID for the function.</span></span> <span data-ttu-id="214fb-145">Этот идентификатор может содержать только буквы, цифры и точки и не может изменяться после настройки.</span><span class="sxs-lookup"><span data-stu-id="214fb-145">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span>                                            |
| `name`        | <span data-ttu-id="214fb-146">string</span><span class="sxs-lookup"><span data-stu-id="214fb-146">string</span></span>    | <span data-ttu-id="214fb-147">Да</span><span class="sxs-lookup"><span data-stu-id="214fb-147">Yes</span></span>      | <span data-ttu-id="214fb-148">Имя функции, которое отображается пользователям в Excel.</span><span class="sxs-lookup"><span data-stu-id="214fb-148">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="214fb-149">В Excel это имя функции предваряется пространством имен пользовательских функций, указанным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="214fb-149">In Excel, this function name is prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
| `options`     | <span data-ttu-id="214fb-150">object</span><span class="sxs-lookup"><span data-stu-id="214fb-150">object</span></span>    | <span data-ttu-id="214fb-151">Нет</span><span class="sxs-lookup"><span data-stu-id="214fb-151">No</span></span>       | <span data-ttu-id="214fb-152">Позволяет настроить некоторые аспекты того, как и когда Excel выполняет функцию.</span><span class="sxs-lookup"><span data-stu-id="214fb-152">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="214fb-153">Дополнительные сведения см. в разделе [options](#options).</span><span class="sxs-lookup"><span data-stu-id="214fb-153">See [options](#options) for details.</span></span>                                                          |
| `parameters`  | <span data-ttu-id="214fb-154">array</span><span class="sxs-lookup"><span data-stu-id="214fb-154">array</span></span>     | <span data-ttu-id="214fb-155">Да</span><span class="sxs-lookup"><span data-stu-id="214fb-155">Yes</span></span>      | <span data-ttu-id="214fb-156">Массив, который определяет входные параметры для функции.</span><span class="sxs-lookup"><span data-stu-id="214fb-156">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="214fb-157">Дополнительные сведения см. в разделе [Parameters](#parameters) .</span><span class="sxs-lookup"><span data-stu-id="214fb-157">See [parameters](#parameters) for details.</span></span>                                                                             |
| `result`      | <span data-ttu-id="214fb-158">object</span><span class="sxs-lookup"><span data-stu-id="214fb-158">object</span></span>    | <span data-ttu-id="214fb-159">Да</span><span class="sxs-lookup"><span data-stu-id="214fb-159">Yes</span></span>      | <span data-ttu-id="214fb-160">Объект, который определяет тип информации, возвращаемый функцией.</span><span class="sxs-lookup"><span data-stu-id="214fb-160">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="214fb-161">Дополнительные сведения см. в разделе [result](#result).</span><span class="sxs-lookup"><span data-stu-id="214fb-161">See [result](#result) for details.</span></span>                                                                 |

### <a name="options"></a><span data-ttu-id="214fb-162">options</span><span class="sxs-lookup"><span data-stu-id="214fb-162">options</span></span>

<span data-ttu-id="214fb-163">Объект `options` позволяет настроить некоторые аспекты того, как и когда Excel выполняет функцию.</span><span class="sxs-lookup"><span data-stu-id="214fb-163">The `options` object enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="214fb-164">В таблице ниже приведены свойства объекта `options`.</span><span class="sxs-lookup"><span data-stu-id="214fb-164">The following table lists the properties of the `options` object.</span></span>

| <span data-ttu-id="214fb-165">Свойство</span><span class="sxs-lookup"><span data-stu-id="214fb-165">Property</span></span>          | <span data-ttu-id="214fb-166">Тип данных</span><span class="sxs-lookup"><span data-stu-id="214fb-166">Data type</span></span> | <span data-ttu-id="214fb-167">Обязательный</span><span class="sxs-lookup"><span data-stu-id="214fb-167">Required</span></span>                               | <span data-ttu-id="214fb-168">Описание</span><span class="sxs-lookup"><span data-stu-id="214fb-168">Description</span></span> |
| :---------------- | :-------- | :------------------------------------- | :---------- |
| `cancelable`      | <span data-ttu-id="214fb-169">boolean</span><span class="sxs-lookup"><span data-stu-id="214fb-169">boolean</span></span>   | <span data-ttu-id="214fb-170">Нет</span><span class="sxs-lookup"><span data-stu-id="214fb-170">No</span></span><br/><br/><span data-ttu-id="214fb-171">Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="214fb-171">Default value is `false`.</span></span>  | <span data-ttu-id="214fb-172">Если это свойство имеет значение `true`, Excel будет вызывать обработчик `CancelableInvocation` каждый раз, когда пользователь будет предпринимать действия, которые приводят к отмене функции (например, вручную вызывает пересчет или редактирует ячейку, на которую ссылается функция).</span><span class="sxs-lookup"><span data-stu-id="214fb-172">If `true`, Excel calls the `CancelableInvocation` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="214fb-173">Функции, которые можно отменять, обычно используются только для асинхронных функций, которые возвращают один результат и нуждаются в обработке отмены запроса данных.</span><span class="sxs-lookup"><span data-stu-id="214fb-173">Cancelable functions are typically only used for asynchronous functions that return a single result and need to handle the cancellation of a request for data.</span></span> <span data-ttu-id="214fb-174">Функция не может быть одновременно потоковой и отмены.</span><span class="sxs-lookup"><span data-stu-id="214fb-174">A function cannot be both streaming and cancelable.</span></span> <span data-ttu-id="214fb-175">Более подробную информацию можно найти в заметке около конца [функции потоковой передачи](custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="214fb-175">For more information, see the note near the end of [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> |
| `requiresAddress` | <span data-ttu-id="214fb-176">boolean</span><span class="sxs-lookup"><span data-stu-id="214fb-176">boolean</span></span>   | <span data-ttu-id="214fb-177">Нет</span><span class="sxs-lookup"><span data-stu-id="214fb-177">No</span></span> <br/><br/><span data-ttu-id="214fb-178">Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="214fb-178">Default value is `false`.</span></span> | <span data-ttu-id="214fb-179">Если `true` Пользовательская функция может получить доступ к адресу ячейки, которая вызвала пользовательскую функцию.</span><span class="sxs-lookup"><span data-stu-id="214fb-179">If `true`, your custom function can access the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="214fb-180">Чтобы получить адрес ячейки, которая вызвала пользовательскую функцию, используйте context. Address в пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="214fb-180">To get the address of the cell that invoked your custom function, use context.address in your custom function.</span></span> <span data-ttu-id="214fb-181">Пользовательские функции не могут быть заданы как потоковые, так и Рекуиресаддресс.</span><span class="sxs-lookup"><span data-stu-id="214fb-181">Custom functions cannot be set as both streaming and requiresAddress.</span></span> <span data-ttu-id="214fb-182">При использовании этого параметра параметр "вызов" должен быть последним параметром, переданным в параметрах.</span><span class="sxs-lookup"><span data-stu-id="214fb-182">When using this option, the 'invocation' parameter must be the last parameter passed in options.</span></span> |
| `stream`          | <span data-ttu-id="214fb-183">boolean</span><span class="sxs-lookup"><span data-stu-id="214fb-183">boolean</span></span>   | <span data-ttu-id="214fb-184">Нет</span><span class="sxs-lookup"><span data-stu-id="214fb-184">No</span></span><br/><br/><span data-ttu-id="214fb-185">Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="214fb-185">Default value is `false`.</span></span>  | <span data-ttu-id="214fb-186">Если это свойство имеет значение `true`, функция может выводить значение в ячейку несколько раз, даже если вызвана всего единожды.</span><span class="sxs-lookup"><span data-stu-id="214fb-186">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="214fb-187">Этот параметр полезен для быстро изменяющихся источников данных, таких как цена акций.</span><span class="sxs-lookup"><span data-stu-id="214fb-187">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="214fb-188">Функция не должна содержать оператор `return`.</span><span class="sxs-lookup"><span data-stu-id="214fb-188">The function should have no `return` statement.</span></span> <span data-ttu-id="214fb-189">Вместо этого результирующее значение передается как аргумент метода обратного вызова `StreamingInvocation.setResult`.</span><span class="sxs-lookup"><span data-stu-id="214fb-189">Instead, the result value is passed as the argument of the `StreamingInvocation.setResult` callback method.</span></span> <span data-ttu-id="214fb-190">Дополнительные сведения см. в разделе [Потоковые функции](custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="214fb-190">For more information, see [Streaming functions](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> |
| `volatile`        | <span data-ttu-id="214fb-191">boolean</span><span class="sxs-lookup"><span data-stu-id="214fb-191">boolean</span></span>   | <span data-ttu-id="214fb-192">Нет</span><span class="sxs-lookup"><span data-stu-id="214fb-192">No</span></span> <br/><br/><span data-ttu-id="214fb-193">Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="214fb-193">Default value is `false`.</span></span> | <span data-ttu-id="214fb-194">Если `true` функция пересчитывает при каждом пересчете Excel, функция пересчитывается, а не только при изменении зависимых значений формулы.</span><span class="sxs-lookup"><span data-stu-id="214fb-194">If `true`, the function recalculates each time Excel recalculates, instead of only when the formula's dependent values have changed.</span></span> <span data-ttu-id="214fb-195">Функция не может быть одновременно потоковой и переменной.</span><span class="sxs-lookup"><span data-stu-id="214fb-195">A function cannot be both streaming and volatile.</span></span> <span data-ttu-id="214fb-196">Если обоим свойствам `stream` и `volatile` присвоено значение `true`, параметр переменности будет игнорироваться.</span><span class="sxs-lookup"><span data-stu-id="214fb-196">If the `stream` and `volatile` properties are both set to `true`, the volatile option will be ignored.</span></span> |

### <a name="parameters"></a><span data-ttu-id="214fb-197">parameters</span><span class="sxs-lookup"><span data-stu-id="214fb-197">parameters</span></span>

<span data-ttu-id="214fb-198">Свойство `parameters` представляет собой массив объектов параметров.</span><span class="sxs-lookup"><span data-stu-id="214fb-198">The `parameters` property is an array of parameter objects.</span></span> <span data-ttu-id="214fb-199">В таблице ниже приведены свойства каждого объекта.</span><span class="sxs-lookup"><span data-stu-id="214fb-199">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="214fb-200">Свойство</span><span class="sxs-lookup"><span data-stu-id="214fb-200">Property</span></span>  |  <span data-ttu-id="214fb-201">Тип данных</span><span class="sxs-lookup"><span data-stu-id="214fb-201">Data type</span></span>  |  <span data-ttu-id="214fb-202">Обязательный</span><span class="sxs-lookup"><span data-stu-id="214fb-202">Required</span></span>  |  <span data-ttu-id="214fb-203">Описание</span><span class="sxs-lookup"><span data-stu-id="214fb-203">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="214fb-204">строка</span><span class="sxs-lookup"><span data-stu-id="214fb-204">string</span></span>  |  <span data-ttu-id="214fb-205">Нет</span><span class="sxs-lookup"><span data-stu-id="214fb-205">No</span></span> |  <span data-ttu-id="214fb-206">Описание параметра.</span><span class="sxs-lookup"><span data-stu-id="214fb-206">A description of the parameter.</span></span> <span data-ttu-id="214fb-207">Это отображается в IntelliSense Excel.</span><span class="sxs-lookup"><span data-stu-id="214fb-207">This is displayed in Excel's IntelliSense.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="214fb-208">string</span><span class="sxs-lookup"><span data-stu-id="214fb-208">string</span></span>  |  <span data-ttu-id="214fb-209">Нет</span><span class="sxs-lookup"><span data-stu-id="214fb-209">No</span></span>  |  <span data-ttu-id="214fb-210">Должно быть **скалярным** (значение, отличное от массива) или **матричным** (двухмерный массив).</span><span class="sxs-lookup"><span data-stu-id="214fb-210">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="214fb-211">string</span><span class="sxs-lookup"><span data-stu-id="214fb-211">string</span></span>  |  <span data-ttu-id="214fb-212">Да</span><span class="sxs-lookup"><span data-stu-id="214fb-212">Yes</span></span>  |  <span data-ttu-id="214fb-213">Имя параметра.</span><span class="sxs-lookup"><span data-stu-id="214fb-213">The name of the parameter.</span></span> <span data-ttu-id="214fb-214">Это имя отображается в IntelliSense Excel.</span><span class="sxs-lookup"><span data-stu-id="214fb-214">This name is displayed in Excel's IntelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="214fb-215">string</span><span class="sxs-lookup"><span data-stu-id="214fb-215">string</span></span>  |  <span data-ttu-id="214fb-216">Нет</span><span class="sxs-lookup"><span data-stu-id="214fb-216">No</span></span>  |  <span data-ttu-id="214fb-217">Тип данных параметра.</span><span class="sxs-lookup"><span data-stu-id="214fb-217">The data type of the parameter.</span></span> <span data-ttu-id="214fb-218">Может иметь значение **boolean** , **number** , **string** или **any** , что позволяет использовать любой из трех предыдущих типов.</span><span class="sxs-lookup"><span data-stu-id="214fb-218">Can be **boolean** , **number** , **string** , or **any** , which allows you to use of any of the previous three types.</span></span> <span data-ttu-id="214fb-219">Если это свойство не задано, по умолчанию устанавливается тип данных **any**.</span><span class="sxs-lookup"><span data-stu-id="214fb-219">If this property is not specified, the data type defaults to **any**.</span></span> |
|  `optional`  | <span data-ttu-id="214fb-220">boolean</span><span class="sxs-lookup"><span data-stu-id="214fb-220">boolean</span></span> | <span data-ttu-id="214fb-221">Нет</span><span class="sxs-lookup"><span data-stu-id="214fb-221">No</span></span> | <span data-ttu-id="214fb-222">Если присвоено значение `true`, параметр не обязателен.</span><span class="sxs-lookup"><span data-stu-id="214fb-222">If `true`, the parameter is optional.</span></span> |
|`repeating`| <span data-ttu-id="214fb-223">boolean</span><span class="sxs-lookup"><span data-stu-id="214fb-223">boolean</span></span> | <span data-ttu-id="214fb-224">Нет</span><span class="sxs-lookup"><span data-stu-id="214fb-224">No</span></span> | <span data-ttu-id="214fb-225">Если `true` параметры заполняются из указанного массива.</span><span class="sxs-lookup"><span data-stu-id="214fb-225">If `true`, parameters populate from a specified array.</span></span> <span data-ttu-id="214fb-226">Обратите внимание, что функции все повторяющиеся параметры считаются необязательными параметрами по определению.</span><span class="sxs-lookup"><span data-stu-id="214fb-226">Note that functions all repeating parameters are considered optional parameters by definition.</span></span>  |

### <a name="result"></a><span data-ttu-id="214fb-227">result</span><span class="sxs-lookup"><span data-stu-id="214fb-227">result</span></span>

<span data-ttu-id="214fb-228">Объект `result` определяет тип информации, возвращаемый функцией.</span><span class="sxs-lookup"><span data-stu-id="214fb-228">The `result` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="214fb-229">В таблице ниже приведены свойства объекта `result`.</span><span class="sxs-lookup"><span data-stu-id="214fb-229">The following table lists the properties of the `result` object.</span></span>

| <span data-ttu-id="214fb-230">Свойство</span><span class="sxs-lookup"><span data-stu-id="214fb-230">Property</span></span>         | <span data-ttu-id="214fb-231">Тип данных</span><span class="sxs-lookup"><span data-stu-id="214fb-231">Data type</span></span> | <span data-ttu-id="214fb-232">Обязательный</span><span class="sxs-lookup"><span data-stu-id="214fb-232">Required</span></span> | <span data-ttu-id="214fb-233">Описание</span><span class="sxs-lookup"><span data-stu-id="214fb-233">Description</span></span>                                                                          |
| :--------------- | :-------- | :------- | :----------------------------------------------------------------------------------- |
| `dimensionality` | <span data-ttu-id="214fb-234">строка</span><span class="sxs-lookup"><span data-stu-id="214fb-234">string</span></span>    | <span data-ttu-id="214fb-235">Нет</span><span class="sxs-lookup"><span data-stu-id="214fb-235">No</span></span>       | <span data-ttu-id="214fb-236">Должно быть **скалярным** (значение, отличное от массива) или **матричным** (двухмерный массив).</span><span class="sxs-lookup"><span data-stu-id="214fb-236">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |

## <a name="associating-function-names-with-json-metadata"></a><span data-ttu-id="214fb-237">Сопоставление имен функций с метаданными JSON</span><span class="sxs-lookup"><span data-stu-id="214fb-237">Associating function names with JSON metadata</span></span>

<span data-ttu-id="214fb-238">Чтобы функция работала должным образом, необходимо связать `id` свойство функции с реализацией JavaScript.</span><span class="sxs-lookup"><span data-stu-id="214fb-238">For a function to work properly, you need to associate the function's `id` property with the JavaScript implementation.</span></span> <span data-ttu-id="214fb-239">Убедитесь, что существует связь, в противном случае функция не будет зарегистрирована и непригодна для работы в Excel.</span><span class="sxs-lookup"><span data-stu-id="214fb-239">Make sure there is an association, otherwise the function won't be registered and isn't useable in Excel.</span></span> <span data-ttu-id="214fb-240">В приведенном ниже примере кода показано, как выполнить связь с помощью `CustomFunctions.associate()` метода.</span><span class="sxs-lookup"><span data-stu-id="214fb-240">The following code sample shows how to make the association using the `CustomFunctions.associate()` method.</span></span> <span data-ttu-id="214fb-241">Пример определяет пользовательскую функцию `add` и связывает ее с объектом в файле метаданных JSON, где для свойства `id` установлено значение **ADD**.</span><span class="sxs-lookup"><span data-stu-id="214fb-241">The sample defines the custom function `add` and associates it with the object in the JSON metadata file where the value of the `id` property is **ADD**.</span></span>

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

<span data-ttu-id="214fb-242">В следующем JSON показаны метаданные JSON, связанные с предыдущим кодом пользовательской функции JavaScript.</span><span class="sxs-lookup"><span data-stu-id="214fb-242">The following JSON shows the JSON metadata that is associated with the previous custom function JavaScript code.</span></span>

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

<span data-ttu-id="214fb-243">Имейте в виду приведенные ниже рекомендации при создании пользовательских функций в файле JavaScript и указании соответствующих сведений в файле метаданных JSON.</span><span class="sxs-lookup"><span data-stu-id="214fb-243">Keep in mind the following best practices when creating custom functions in your JavaScript file and specifying corresponding information in the JSON metadata file.</span></span>

- <span data-ttu-id="214fb-244">Убедитесь, что в файле метаданных JSON значение каждого свойства `id` содержит только буквы, цифры и точки.</span><span class="sxs-lookup"><span data-stu-id="214fb-244">In the JSON metadata file, ensure that the value of each `id` property contains only alphanumeric characters and periods.</span></span>

- <span data-ttu-id="214fb-245">Убедитесь, что в файле метаданных JSON значение каждого свойства `id` уникально в пределах файла.</span><span class="sxs-lookup"><span data-stu-id="214fb-245">In the JSON metadata file, ensure that the value of each `id` property is unique within the scope of the file.</span></span> <span data-ttu-id="214fb-246">То есть никакие два объекта функций в файле метаданных не должны иметь одинаковое значение `id`.</span><span class="sxs-lookup"><span data-stu-id="214fb-246">That is, no two function objects in the metadata file should have the same `id` value.</span></span>

- <span data-ttu-id="214fb-247">Не изменяйте значение свойства `id` в файле метаданных JSON после его сопоставления с соответствующим именем функции JavaScript.</span><span class="sxs-lookup"><span data-stu-id="214fb-247">Do not change the value of an `id` property in the JSON metadata file after it's been associated with a corresponding JavaScript function name.</span></span> <span data-ttu-id="214fb-248">Вы можете изменить имя функции, которое отображается для конечных пользователей в Excel, путем обновления свойства `name` в файле метаданных JSON, но никогда не следует изменять значение свойства `id` после его установления.</span><span class="sxs-lookup"><span data-stu-id="214fb-248">You can change the function name that end users see in Excel by updating the `name` property within the JSON metadata file, but you should never change the value of an `id` property after it's been established.</span></span>

- <span data-ttu-id="214fb-249">В файле JavaScript укажите настраиваемое сопоставление функций с помощью `CustomFunctions.associate` каждой функции.</span><span class="sxs-lookup"><span data-stu-id="214fb-249">In the JavaScript file, specify a custom function association using `CustomFunctions.associate` after each function.</span></span>

<span data-ttu-id="214fb-250">В следующем примере показаны метаданные JSON, которые соответствуют функциям, определенным в предыдущем примере кода JavaScript.</span><span class="sxs-lookup"><span data-stu-id="214fb-250">The following sample shows the JSON metadata that corresponds to the functions defined in the preceding JavaScript code sample.</span></span> <span data-ttu-id="214fb-251">`id` `name` Значения свойств и представлены в верхнем регистре, что является лучшим вариантом при описании пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="214fb-251">The `id` and `name` property values are in uppercase, which is a best practice when describing your custom functions.</span></span> <span data-ttu-id="214fb-252">Этот код JSON необходимо добавить только в том случае, если вы готовите собственный файл JSON вручную и не используете автоматическое создание.</span><span class="sxs-lookup"><span data-stu-id="214fb-252">You only need to add this JSON if you are preparing your own JSON file manually and not using autogeneration.</span></span> <span data-ttu-id="214fb-253">Дополнительные сведения об автоформировании приведены в статье Автоматическое [Создание МЕТАДАННЫХ JSON для пользовательских функций](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="214fb-253">For more information on autogeneration, see [Autogenerate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

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

## <a name="next-steps"></a><span data-ttu-id="214fb-254">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="214fb-254">Next steps</span></span>

<span data-ttu-id="214fb-255">Ознакомьтесь с рекомендациями [по именованию функции](custom-functions-naming.md) или [локализации функции](custom-functions-localize.md) с помощью ранее описанного рукописного метода JSON.</span><span class="sxs-lookup"><span data-stu-id="214fb-255">Learn the [best practices for naming your function](custom-functions-naming.md) or discover how to [localize your function](custom-functions-localize.md) using the previously described handwritten JSON method.</span></span>

## <a name="see-also"></a><span data-ttu-id="214fb-256">См. также</span><span class="sxs-lookup"><span data-stu-id="214fb-256">See also</span></span>

- [<span data-ttu-id="214fb-257">Автоматическое генерирование метаданных JSON для пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="214fb-257">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
- [<span data-ttu-id="214fb-258">Параметры параметров пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="214fb-258">Custom functions parameter options</span></span>](custom-functions-parameter-options.md)
- [<span data-ttu-id="214fb-259">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="214fb-259">Create custom functions in Excel</span></span>](custom-functions-overview.md)
