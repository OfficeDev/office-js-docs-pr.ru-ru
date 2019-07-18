---
ms.date: 07/15/2019
description: Определите метаданные JSON для пользовательских функций в Excel и свяжите свойства идентификатора и имени функции.
title: Метаданные для пользовательских функций в Excel
localization_priority: Normal
ms.openlocfilehash: b0e015cfa439651420487db4885647f5c7de7da8
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771563"
---
# <a name="custom-functions-metadata"></a><span data-ttu-id="66a54-103">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="66a54-103">Custom functions metadata</span></span>

<span data-ttu-id="66a54-104">Как описано в статье [Обзор пользовательских функций](custom-functions-overview.md) , проект пользовательских функций должен включать файл метаданных JSON и файл скрипта (JavaScript или TypeScript) для регистрации функции, делая ее доступной для использования.</span><span class="sxs-lookup"><span data-stu-id="66a54-104">As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include both a JSON metadata file and a script (either JavaScript or TypeScript) file to register a function, making it available for use.</span></span> <span data-ttu-id="66a54-105">Пользовательские функции регистрируются при первом запуске надстройки и после их появления для одного и того же пользователя во всех книгах.</span><span class="sxs-lookup"><span data-stu-id="66a54-105">Custom functions are registered when the user runs the add-in for the first time and after that are available to the same user in all workbooks.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="66a54-106">Рекомендуется использовать автоматическое создание JSON, когда это возможно, используя файлы `yo office` шаблонов, похожие на процесс, показанный в руководстве по настраиваемым функциям [Excel](../tutorials/excel-tutorial-create-custom-functions.md) , так как этот процесс проще и менее подвержен ошибкам пользователя.</span><span class="sxs-lookup"><span data-stu-id="66a54-106">It is recommended that you use JSON autogeneration when possible, using the `yo office` scaffold files, similar to the process shown in the [Excel Custom Function tutorial](../tutorials/excel-tutorial-create-custom-functions.md) because this process is easier and less prone to user error.</span></span> <span data-ttu-id="66a54-107">Дополнительные сведения о процессе создания JSON файла Жсдок Comment можно найти в статье [Создание МЕТАДАННЫХ JSON для пользовательских функций](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="66a54-107">For more information on the process of JSDoc comment JSON file generation, see [Generate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

<span data-ttu-id="66a54-108">Тем не менее, проект пользовательских функций можно сделать с нуля. для этого необходимо выполнить следующие действия:</span><span class="sxs-lookup"><span data-stu-id="66a54-108">However, you can make a custom functions project from scratch; it requires that you:</span></span>

- <span data-ttu-id="66a54-109">Создание файла JSON вручную</span><span class="sxs-lookup"><span data-stu-id="66a54-109">Write your JSON file by hand</span></span>
- <span data-ttu-id="66a54-110">Убедитесь, что файл манифеста подключен к файлу JSON, созданному вручную</span><span class="sxs-lookup"><span data-stu-id="66a54-110">Check that your manifest file is connected to your hand-authored JSON file</span></span>
- <span data-ttu-id="66a54-111">Свяжите функции `id` и `name` свойства в файле скрипта, чтобы зарегистрировать функции</span><span class="sxs-lookup"><span data-stu-id="66a54-111">Associate your functions' `id` and `name` properties in the script file in order to register your functions</span></span>

<span data-ttu-id="66a54-112">В этой статье рассказывается, как выполнить все три из этих действий.</span><span class="sxs-lookup"><span data-stu-id="66a54-112">This article will show you how to do all three of these steps.</span></span>

> [!NOTE]
> <span data-ttu-id="66a54-113">В отличие от файлов `yo office` шаблонов, необходимо подключить манифест к созданному файлу JSON с помощью `<Resources>` раздела в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="66a54-113">In contrast with the `yo office` scaffold files, you need to hook your manifest up with the JSON file you create, through the `<Resources>` section in your XML manifest file.</span></span> <span data-ttu-id="66a54-114">Обратите внимание, что параметры [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) для сервера, на котором размещается JSON-файл, должны быть включены, чтобы пользовательские функции правильно работали в Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="66a54-114">Note that the server settings on the server that hosts the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel on the web.</span></span>

## <a name="authoring-metadata-and-hooking-up-to-the-manifest"></a><span data-ttu-id="66a54-115">Создание метаданных и подключение к манифесту</span><span class="sxs-lookup"><span data-stu-id="66a54-115">Authoring metadata and hooking up to the manifest</span></span>

<span data-ttu-id="66a54-116">Необходимо создать файл JSON в проекте и предоставить все сведения о функциях, в том числе параметры функции.</span><span class="sxs-lookup"><span data-stu-id="66a54-116">You need to create a JSON file in your project and provide all the details about your functions in it, such as the function's parameters.</span></span> <span data-ttu-id="66a54-117">В [приведенном ниже примере метаданных](#json-metadata-example) и Справочнике по [метаданным](#metadata-reference) представлен полный список свойств функций.</span><span class="sxs-lookup"><span data-stu-id="66a54-117">See the [following metadata example](#json-metadata-example) and [the metadata reference](#metadata-reference) for a complete list of function properties.</span></span>

<span data-ttu-id="66a54-118">Также необходимо убедиться, что XML-файл манифеста ссылается на JSON-файл в `<Resources>` разделе, как в следующем примере.</span><span class="sxs-lookup"><span data-stu-id="66a54-118">You also need to make sure your XML manifest file references your JSON file in the `<Resources>` section, similar to the following example.</span></span>

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

## <a name="json-metadata-example"></a><span data-ttu-id="66a54-119">Пример метаданных JSON</span><span class="sxs-lookup"><span data-stu-id="66a54-119">JSON metadata example</span></span>

<span data-ttu-id="66a54-120">В примере кода ниже показано содержимое JSON-файла метаданных для надстройки, определяющей настраиваемые функции.</span><span class="sxs-lookup"><span data-stu-id="66a54-120">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="66a54-121">В следующих за этим примером разделах приводятся подробные сведения об отдельных свойствах, представленных в этом примере JSON.</span><span class="sxs-lookup"><span data-stu-id="66a54-121">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
> <span data-ttu-id="66a54-122">Полный пример JSON-файла доступен в журнале транзакций [OfficeDev/Excel-Custom-functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) репозитория GitHub.</span><span class="sxs-lookup"><span data-stu-id="66a54-122">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) GitHub repository's commit history.</span></span> <span data-ttu-id="66a54-123">Так как проект был скорректирован для автоматического создания JSON, полный пример рукописного кода JSON доступен только в предыдущих версиях проекта.</span><span class="sxs-lookup"><span data-stu-id="66a54-123">As the project has been adjusted to automatically generate JSON, a full sample of handwritten JSON is only available in previous versions of the project.</span></span>

## <a name="metadata-reference"></a><span data-ttu-id="66a54-124">Справка по метаданным</span><span class="sxs-lookup"><span data-stu-id="66a54-124">Metadata reference</span></span>

### <a name="functions"></a><span data-ttu-id="66a54-125">functions</span><span class="sxs-lookup"><span data-stu-id="66a54-125">functions</span></span>

<span data-ttu-id="66a54-126">Свойство `functions` представляет собой массив объектов настраиваемых функций.</span><span class="sxs-lookup"><span data-stu-id="66a54-126">The `functions` property is an array of custom function objects.</span></span> <span data-ttu-id="66a54-127">В таблице ниже приведены свойства каждого объекта.</span><span class="sxs-lookup"><span data-stu-id="66a54-127">The following table lists the properties of each object.</span></span>

| <span data-ttu-id="66a54-128">Свойство</span><span class="sxs-lookup"><span data-stu-id="66a54-128">Property</span></span>      | <span data-ttu-id="66a54-129">Тип данных</span><span class="sxs-lookup"><span data-stu-id="66a54-129">Data type</span></span> | <span data-ttu-id="66a54-130">Обязательный</span><span class="sxs-lookup"><span data-stu-id="66a54-130">Required</span></span> | <span data-ttu-id="66a54-131">Описание</span><span class="sxs-lookup"><span data-stu-id="66a54-131">Description</span></span>                                                                                                                                                                      |
| :------------ | :-------- | :------- | :------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `description` | <span data-ttu-id="66a54-132">string</span><span class="sxs-lookup"><span data-stu-id="66a54-132">string</span></span>    | <span data-ttu-id="66a54-133">Нет</span><span class="sxs-lookup"><span data-stu-id="66a54-133">No</span></span>       | <span data-ttu-id="66a54-134">Описание функции, которое отображается пользователям в Excel</span><span class="sxs-lookup"><span data-stu-id="66a54-134">The description of the function that end users see in Excel.</span></span> <span data-ttu-id="66a54-135">(например, **преобразует значение по шкале Цельсия в температуру по шкале Фаренгейта**).</span><span class="sxs-lookup"><span data-stu-id="66a54-135">For example, **Converts a Celsius value to Fahrenheit**.</span></span>                                                            |
| `helpUrl`     | <span data-ttu-id="66a54-136">string</span><span class="sxs-lookup"><span data-stu-id="66a54-136">string</span></span>    | <span data-ttu-id="66a54-137">Нет</span><span class="sxs-lookup"><span data-stu-id="66a54-137">No</span></span>       | <span data-ttu-id="66a54-138">URL-адрес, по которому можно получить сведения о функции</span><span class="sxs-lookup"><span data-stu-id="66a54-138">URL that provides information about the function.</span></span> <span data-ttu-id="66a54-139">(отображается в области задач). Пример: `http://contoso.com/help/convertcelsiustofahrenheit.html`.</span><span class="sxs-lookup"><span data-stu-id="66a54-139">(It is displayed in a task pane.) For example, `http://contoso.com/help/convertcelsiustofahrenheit.html`.</span></span>                      |
| `id`          | <span data-ttu-id="66a54-140">string</span><span class="sxs-lookup"><span data-stu-id="66a54-140">string</span></span>    | <span data-ttu-id="66a54-141">Да</span><span class="sxs-lookup"><span data-stu-id="66a54-141">Yes</span></span>      | <span data-ttu-id="66a54-142">Уникальный идентификатор для функции.</span><span class="sxs-lookup"><span data-stu-id="66a54-142">A unique ID for the function.</span></span> <span data-ttu-id="66a54-143">Этот идентификатор может содержать только буквы, цифры и точки и не может изменяться после настройки.</span><span class="sxs-lookup"><span data-stu-id="66a54-143">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span>                                            |
| `name`        | <span data-ttu-id="66a54-144">string</span><span class="sxs-lookup"><span data-stu-id="66a54-144">string</span></span>    | <span data-ttu-id="66a54-145">Да</span><span class="sxs-lookup"><span data-stu-id="66a54-145">Yes</span></span>      | <span data-ttu-id="66a54-146">Имя функции, которое отображается пользователям в Excel.</span><span class="sxs-lookup"><span data-stu-id="66a54-146">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="66a54-147">В Excel имя этой функции будет присоединено в качестве префикса пространством имен настраиваемой функции, указанным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="66a54-147">In Excel, this function name will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
| `options`     | <span data-ttu-id="66a54-148">object</span><span class="sxs-lookup"><span data-stu-id="66a54-148">object</span></span>    | <span data-ttu-id="66a54-149">Нет</span><span class="sxs-lookup"><span data-stu-id="66a54-149">No</span></span>       | <span data-ttu-id="66a54-150">Позволяет настроить некоторые аспекты того, как и когда Excel выполняет функцию.</span><span class="sxs-lookup"><span data-stu-id="66a54-150">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="66a54-151">Дополнительные сведения см. в разделе [options](#options).</span><span class="sxs-lookup"><span data-stu-id="66a54-151">See [options](#options) for details.</span></span>                                                          |
| `parameters`  | <span data-ttu-id="66a54-152">array</span><span class="sxs-lookup"><span data-stu-id="66a54-152">array</span></span>     | <span data-ttu-id="66a54-153">Да</span><span class="sxs-lookup"><span data-stu-id="66a54-153">Yes</span></span>      | <span data-ttu-id="66a54-154">Массив, который определяет входные параметры для функции.</span><span class="sxs-lookup"><span data-stu-id="66a54-154">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="66a54-155">Дополнительные сведения см. в разделе [Parameters](#parameters) .</span><span class="sxs-lookup"><span data-stu-id="66a54-155">See [parameters](#parameters) for details.</span></span>                                                                             |
| `result`      | <span data-ttu-id="66a54-156">object</span><span class="sxs-lookup"><span data-stu-id="66a54-156">object</span></span>    | <span data-ttu-id="66a54-157">Да</span><span class="sxs-lookup"><span data-stu-id="66a54-157">Yes</span></span>      | <span data-ttu-id="66a54-158">Объект, который определяет тип информации, возвращаемый функцией.</span><span class="sxs-lookup"><span data-stu-id="66a54-158">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="66a54-159">Дополнительные сведения см. в разделе [result](#result).</span><span class="sxs-lookup"><span data-stu-id="66a54-159">See [result](#result) for details.</span></span>                                                                 |

### <a name="options"></a><span data-ttu-id="66a54-160">options</span><span class="sxs-lookup"><span data-stu-id="66a54-160">options</span></span>

<span data-ttu-id="66a54-161">Объект `options` позволяет настроить некоторые аспекты того, как и когда Excel выполняет функцию.</span><span class="sxs-lookup"><span data-stu-id="66a54-161">The `options` object enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="66a54-162">В таблице ниже приведены свойства объекта `options`.</span><span class="sxs-lookup"><span data-stu-id="66a54-162">The following table lists the properties of the `options` object.</span></span>

| <span data-ttu-id="66a54-163">Свойство</span><span class="sxs-lookup"><span data-stu-id="66a54-163">Property</span></span>          | <span data-ttu-id="66a54-164">Тип данных</span><span class="sxs-lookup"><span data-stu-id="66a54-164">Data type</span></span> | <span data-ttu-id="66a54-165">Обязательный</span><span class="sxs-lookup"><span data-stu-id="66a54-165">Required</span></span>                               | <span data-ttu-id="66a54-166">Описание</span><span class="sxs-lookup"><span data-stu-id="66a54-166">Description</span></span>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                |
| :---------------- | :-------- | :------------------------------------- | :--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `cancelable`      | <span data-ttu-id="66a54-167">boolean</span><span class="sxs-lookup"><span data-stu-id="66a54-167">boolean</span></span>   | <span data-ttu-id="66a54-168">Нет</span><span class="sxs-lookup"><span data-stu-id="66a54-168">No</span></span><br/><br/><span data-ttu-id="66a54-169">Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="66a54-169">Default value is `false`.</span></span>  | <span data-ttu-id="66a54-170">Если это свойство имеет значение `true`, Excel будет вызывать обработчик `CancelableInvocation` каждый раз, когда пользователь будет предпринимать действия, которые приводят к отмене функции (например, вручную вызывает пересчет или редактирует ячейку, на которую ссылается функция).</span><span class="sxs-lookup"><span data-stu-id="66a54-170">If `true`, Excel calls the `CancelableInvocation` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="66a54-171">Функции, которые можно отменять, обычно используются только для асинхронных функций, которые возвращают один результат и нуждаются в обработке отмены запроса данных.</span><span class="sxs-lookup"><span data-stu-id="66a54-171">Cancelable functions are typically only used for asynchronous functions that return a single result and need to handle the cancellation of a request for data.</span></span> <span data-ttu-id="66a54-172">Функция не может быть одновременно потоковой и отмены.</span><span class="sxs-lookup"><span data-stu-id="66a54-172">A function cannot be both streaming and cancelable.</span></span> <span data-ttu-id="66a54-173">Более подробную информацию можно найти в заметке около конца [функции потоковой передачи](custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="66a54-173">For more information, see the note near the end of [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> |
| `requiresAddress` | <span data-ttu-id="66a54-174">boolean</span><span class="sxs-lookup"><span data-stu-id="66a54-174">boolean</span></span>   | <span data-ttu-id="66a54-175">Нет</span><span class="sxs-lookup"><span data-stu-id="66a54-175">No</span></span> <br/><br/><span data-ttu-id="66a54-176">Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="66a54-176">Default value is `false`.</span></span> | <span data-ttu-id="66a54-177">Если `true`пользовательская функция может получить доступ к адресу ячейки, которая вызвала пользовательскую функцию.</span><span class="sxs-lookup"><span data-stu-id="66a54-177">If `true`, your custom function can access the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="66a54-178">Чтобы получить адрес ячейки, которая вызвала пользовательскую функцию, используйте context. Address в пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="66a54-178">To get the address of the cell that invoked your custom function, use context.address in your custom function.</span></span> <span data-ttu-id="66a54-179">Более подробную информацию можно узнать в разделе [Address Parameter Cell](/office/dev/add-ins/excel/custom-functions-parameter-options#addressing-cells-context-parameter).</span><span class="sxs-lookup"><span data-stu-id="66a54-179">For more information, see [Addressing cell's context parameter](/office/dev/add-ins/excel/custom-functions-parameter-options#addressing-cells-context-parameter).</span></span> <span data-ttu-id="66a54-180">Пользовательские функции не могут быть заданы как потоковые, так и Рекуиресаддресс.</span><span class="sxs-lookup"><span data-stu-id="66a54-180">Custom functions cannot be set as both streaming and requiresAddress.</span></span> <span data-ttu-id="66a54-181">При использовании этого параметра параметр "вызов" должен быть последним параметром, переданным в параметрах.</span><span class="sxs-lookup"><span data-stu-id="66a54-181">When using this option, the 'invocation' parameter must be the last parameter passed in options.</span></span>                                              |
| `stream`          | <span data-ttu-id="66a54-182">boolean</span><span class="sxs-lookup"><span data-stu-id="66a54-182">boolean</span></span>   | <span data-ttu-id="66a54-183">Нет</span><span class="sxs-lookup"><span data-stu-id="66a54-183">No</span></span><br/><br/><span data-ttu-id="66a54-184">Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="66a54-184">Default value is `false`.</span></span>  | <span data-ttu-id="66a54-185">Если это свойство имеет значение `true`, функция может выводить значение в ячейку несколько раз, даже если вызвана всего единожды.</span><span class="sxs-lookup"><span data-stu-id="66a54-185">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="66a54-186">Этот параметр полезен для быстро изменяющихся источников данных, таких как цена акций.</span><span class="sxs-lookup"><span data-stu-id="66a54-186">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="66a54-187">Функция не должна содержать оператор `return`.</span><span class="sxs-lookup"><span data-stu-id="66a54-187">The function should have no `return` statement.</span></span> <span data-ttu-id="66a54-188">Вместо этого результирующее значение передается как аргумент метода обратного вызова `StreamingInvocation.setResult`.</span><span class="sxs-lookup"><span data-stu-id="66a54-188">Instead, the result value is passed as the argument of the `StreamingInvocation.setResult` callback method.</span></span> <span data-ttu-id="66a54-189">Дополнительные сведения см. в разделе [Потоковые функции](custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="66a54-189">For more information, see [Streaming functions](custom-functions-web-reqs.md#make-a-streaming-function).</span></span>                                                                                                                                                                |
| `volatile`        | <span data-ttu-id="66a54-190">boolean</span><span class="sxs-lookup"><span data-stu-id="66a54-190">boolean</span></span>   | <span data-ttu-id="66a54-191">Нет</span><span class="sxs-lookup"><span data-stu-id="66a54-191">No</span></span> <br/><br/><span data-ttu-id="66a54-192">Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="66a54-192">Default value is `false`.</span></span> | <br /><br /> <span data-ttu-id="66a54-193">Если присвоено значение `true`, функция пересчитывается при каждом выполнении пересчета в Excel, а не только при изменении зависимых значений формулы.</span><span class="sxs-lookup"><span data-stu-id="66a54-193">If `true`, the function will recalculate each time Excel recalculates, instead of only when the formula's dependent values have changed.</span></span> <span data-ttu-id="66a54-194">Функция не может быть одновременно потоковой и переменной.</span><span class="sxs-lookup"><span data-stu-id="66a54-194">A function cannot be both streaming and volatile.</span></span> <span data-ttu-id="66a54-195">Если обоим свойствам `stream` и `volatile` присвоено значение `true`, параметр переменности будет игнорироваться.</span><span class="sxs-lookup"><span data-stu-id="66a54-195">If the `stream` and `volatile` properties are both set to `true`, the volatile option will be ignored.</span></span>                                                                                                                                                                                                                                                                                             |

### <a name="parameters"></a><span data-ttu-id="66a54-196">parameters</span><span class="sxs-lookup"><span data-stu-id="66a54-196">parameters</span></span>

<span data-ttu-id="66a54-197">Свойство `parameters` представляет собой массив объектов параметров.</span><span class="sxs-lookup"><span data-stu-id="66a54-197">The `parameters` property is an array of parameter objects.</span></span> <span data-ttu-id="66a54-198">В таблице ниже приведены свойства каждого объекта.</span><span class="sxs-lookup"><span data-stu-id="66a54-198">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="66a54-199">Свойство</span><span class="sxs-lookup"><span data-stu-id="66a54-199">Property</span></span>  |  <span data-ttu-id="66a54-200">Тип данных</span><span class="sxs-lookup"><span data-stu-id="66a54-200">Data type</span></span>  |  <span data-ttu-id="66a54-201">Обязательный</span><span class="sxs-lookup"><span data-stu-id="66a54-201">Required</span></span>  |  <span data-ttu-id="66a54-202">Описание</span><span class="sxs-lookup"><span data-stu-id="66a54-202">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="66a54-203">string</span><span class="sxs-lookup"><span data-stu-id="66a54-203">string</span></span>  |  <span data-ttu-id="66a54-204">Нет</span><span class="sxs-lookup"><span data-stu-id="66a54-204">No</span></span> |  <span data-ttu-id="66a54-205">Описание параметра.</span><span class="sxs-lookup"><span data-stu-id="66a54-205">A description of the parameter.</span></span> <span data-ttu-id="66a54-206">Отображается в IntelliSense Excel.</span><span class="sxs-lookup"><span data-stu-id="66a54-206">This is displayed in Excel's intelliSense.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="66a54-207">string</span><span class="sxs-lookup"><span data-stu-id="66a54-207">string</span></span>  |  <span data-ttu-id="66a54-208">Нет</span><span class="sxs-lookup"><span data-stu-id="66a54-208">No</span></span>  |  <span data-ttu-id="66a54-209">Должно быть **скалярным** (значение, отличное от массива) или **матричным** (двухмерный массив).</span><span class="sxs-lookup"><span data-stu-id="66a54-209">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="66a54-210">string</span><span class="sxs-lookup"><span data-stu-id="66a54-210">string</span></span>  |  <span data-ttu-id="66a54-211">Да</span><span class="sxs-lookup"><span data-stu-id="66a54-211">Yes</span></span>  |  <span data-ttu-id="66a54-212">Имя параметра.</span><span class="sxs-lookup"><span data-stu-id="66a54-212">The name of the parameter.</span></span> <span data-ttu-id="66a54-213">Оно отображается в IntelliSense Excel.</span><span class="sxs-lookup"><span data-stu-id="66a54-213">This name is displayed in Excel's intelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="66a54-214">string</span><span class="sxs-lookup"><span data-stu-id="66a54-214">string</span></span>  |  <span data-ttu-id="66a54-215">Нет</span><span class="sxs-lookup"><span data-stu-id="66a54-215">No</span></span>  |  <span data-ttu-id="66a54-216">Тип данных параметра.</span><span class="sxs-lookup"><span data-stu-id="66a54-216">The data type of the parameter.</span></span> <span data-ttu-id="66a54-217">Может иметь значение **boolean**, **number**, **string** или **any**, что позволяет использовать любой из трех предыдущих типов.</span><span class="sxs-lookup"><span data-stu-id="66a54-217">Can be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> <span data-ttu-id="66a54-218">Если это свойство не задано, по умолчанию устанавливается тип данных **any**.</span><span class="sxs-lookup"><span data-stu-id="66a54-218">If this property is not specified, the data type defaults to **any**.</span></span> |
|  `optional`  | <span data-ttu-id="66a54-219">boolean</span><span class="sxs-lookup"><span data-stu-id="66a54-219">boolean</span></span> | <span data-ttu-id="66a54-220">Нет</span><span class="sxs-lookup"><span data-stu-id="66a54-220">No</span></span> | <span data-ttu-id="66a54-221">Если присвоено значение `true`, параметр не обязателен.</span><span class="sxs-lookup"><span data-stu-id="66a54-221">If `true`, the parameter is optional.</span></span> |
|`repeating`| <span data-ttu-id="66a54-222">boolean</span><span class="sxs-lookup"><span data-stu-id="66a54-222">boolean</span></span> | <span data-ttu-id="66a54-223">Нет</span><span class="sxs-lookup"><span data-stu-id="66a54-223">No</span></span> | <span data-ttu-id="66a54-224">If `true`, параметры заполняются из указанного массива.</span><span class="sxs-lookup"><span data-stu-id="66a54-224">If `true`, parameters will populate from a specified array.</span></span> <span data-ttu-id="66a54-225">Обратите внимание, что функции все повторяющиеся параметры считаются необязательными параметрами по определению.</span><span class="sxs-lookup"><span data-stu-id="66a54-225">Note that functions all repeating parameters are considered optional parameters by definition.</span></span>  |

### <a name="result"></a><span data-ttu-id="66a54-226">result</span><span class="sxs-lookup"><span data-stu-id="66a54-226">result</span></span>

<span data-ttu-id="66a54-227">Объект `result` определяет тип информации, возвращаемый функцией.</span><span class="sxs-lookup"><span data-stu-id="66a54-227">The `result` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="66a54-228">В таблице ниже приведены свойства объекта `result`.</span><span class="sxs-lookup"><span data-stu-id="66a54-228">The following table lists the properties of the `result` object.</span></span>

| <span data-ttu-id="66a54-229">Свойство</span><span class="sxs-lookup"><span data-stu-id="66a54-229">Property</span></span>         | <span data-ttu-id="66a54-230">Тип данных</span><span class="sxs-lookup"><span data-stu-id="66a54-230">Data type</span></span> | <span data-ttu-id="66a54-231">Обязательный</span><span class="sxs-lookup"><span data-stu-id="66a54-231">Required</span></span> | <span data-ttu-id="66a54-232">Описание</span><span class="sxs-lookup"><span data-stu-id="66a54-232">Description</span></span>                                                                          |
| :--------------- | :-------- | :------- | :----------------------------------------------------------------------------------- |
| `dimensionality` | <span data-ttu-id="66a54-233">string</span><span class="sxs-lookup"><span data-stu-id="66a54-233">string</span></span>    | <span data-ttu-id="66a54-234">Нет</span><span class="sxs-lookup"><span data-stu-id="66a54-234">No</span></span>       | <span data-ttu-id="66a54-235">Должно быть **скалярным** (значение, отличное от массива) или **матричным** (двухмерный массив).</span><span class="sxs-lookup"><span data-stu-id="66a54-235">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |

## <a name="associating-function-names-with-json-metadata"></a><span data-ttu-id="66a54-236">Сопоставление имен функций с метаданными JSON</span><span class="sxs-lookup"><span data-stu-id="66a54-236">Associating function names with JSON metadata</span></span>

<span data-ttu-id="66a54-237">Чтобы функция работала должным образом, необходимо связать `id` свойство функции с реализацией JavaScript.</span><span class="sxs-lookup"><span data-stu-id="66a54-237">For a function to work properly, you need to associate the function's `id` property with the JavaScript implementation.</span></span> <span data-ttu-id="66a54-238">Убедитесь, что существует связь, в противном случае функция не будет зарегистрирована и непригодна для работы в Excel.</span><span class="sxs-lookup"><span data-stu-id="66a54-238">Make sure there is an association, otherwise the function will not be registered and not useable in Excel.</span></span> <span data-ttu-id="66a54-239">В приведенном ниже примере кода показано, как выполнить связь с `CustomFunctions.associate()` помощью метода.</span><span class="sxs-lookup"><span data-stu-id="66a54-239">The following code sample shows how to make the association using the `CustomFunctions.associate()` method.</span></span> <span data-ttu-id="66a54-240">Пример определяет пользовательскую функцию `add` и связывает ее с объектом в файле метаданных JSON, где для свойства `id` установлено значение **ADD**.</span><span class="sxs-lookup"><span data-stu-id="66a54-240">The sample defines the custom function `add` and associates it with the object in the JSON metadata file where the value of the `id` property is **ADD**.</span></span>

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

<span data-ttu-id="66a54-241">В следующем JSON показаны метаданные JSON, связанные с предыдущим кодом пользовательской функции JavaScript.</span><span class="sxs-lookup"><span data-stu-id="66a54-241">The following JSON shows the JSON metadata that is associated with the previous custom function JavaScript code.</span></span>

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

<span data-ttu-id="66a54-242">Имейте в виду приведенные ниже рекомендации при создании пользовательских функций в файле JavaScript и указании соответствующих сведений в файле метаданных JSON.</span><span class="sxs-lookup"><span data-stu-id="66a54-242">Keep in mind the following best practices when creating custom functions in your JavaScript file and specifying corresponding information in the JSON metadata file.</span></span>

- <span data-ttu-id="66a54-243">Убедитесь, что в файле метаданных JSON значение каждого свойства `id` содержит только буквы, цифры и точки.</span><span class="sxs-lookup"><span data-stu-id="66a54-243">In the JSON metadata file, ensure that the value of each `id` property contains only alphanumeric characters and periods.</span></span>

- <span data-ttu-id="66a54-244">Убедитесь, что в файле метаданных JSON значение каждого свойства `id` уникально в пределах файла.</span><span class="sxs-lookup"><span data-stu-id="66a54-244">In the JSON metadata file, ensure that the value of each `id` property is unique within the scope of the file.</span></span> <span data-ttu-id="66a54-245">То есть никакие два объекта функций в файле метаданных не должны иметь одинаковое значение `id`.</span><span class="sxs-lookup"><span data-stu-id="66a54-245">That is, no two function objects in the metadata file should have the same `id` value.</span></span>

- <span data-ttu-id="66a54-246">Не изменяйте значение свойства `id` в файле метаданных JSON после его сопоставления с соответствующим именем функции JavaScript.</span><span class="sxs-lookup"><span data-stu-id="66a54-246">Do not change the value of an `id` property in the JSON metadata file after it's been associated with a corresponding JavaScript function name.</span></span> <span data-ttu-id="66a54-247">Вы можете изменить имя функции, которое отображается для конечных пользователей в Excel, путем обновления свойства `name` в файле метаданных JSON, но никогда не следует изменять значение свойства `id` после его установления.</span><span class="sxs-lookup"><span data-stu-id="66a54-247">You can change the function name that end users see in Excel by updating the `name` property within the JSON metadata file, but you should never change the value of an `id` property after it's been established.</span></span>

- <span data-ttu-id="66a54-248">В файле JavaScript укажите настраиваемое сопоставление функций с помощью `CustomFunctions.associate` каждой функции.</span><span class="sxs-lookup"><span data-stu-id="66a54-248">In the JavaScript file, specify a custom function association using `CustomFunctions.associate` after each function.</span></span>

<span data-ttu-id="66a54-249">В приведенном ниже примере показаны метаданные JSON, соответствующие функциям, определенным в этом примере кода JavaScript.</span><span class="sxs-lookup"><span data-stu-id="66a54-249">The following sample shows the JSON metadata that corresponds to the functions defined in this JavaScript code sample.</span></span> <span data-ttu-id="66a54-250">Значения `id` свойств `name` и представлены в верхнем регистре, что является лучшим вариантом при описании пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="66a54-250">The `id` and `name` property values are in uppercase, which is a best practice when describing your custom functions.</span></span> <span data-ttu-id="66a54-251">Этот код JSON необходимо добавить только в том случае, если вы готовите собственный файл JSON вручную и не используете автоматическое создание.</span><span class="sxs-lookup"><span data-stu-id="66a54-251">You only need to add this JSON if you are preparing your own JSON file manually and not using autogeneration.</span></span> <span data-ttu-id="66a54-252">Для получения дополнительной информации об автоформировании, ознакомьтесь со статьей [Создание МЕТАДАННЫХ JSON для пользовательских функций](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="66a54-252">For more information on autogeneration, see [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

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

## <a name="next-steps"></a><span data-ttu-id="66a54-253">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="66a54-253">Next steps</span></span>

<span data-ttu-id="66a54-254">Ознакомьтесь с рекомендациями [по именованию функции](custom-functions-naming.md) или [локализации функции](custom-functions-localize.md) с помощью ранее описанного рукописного метода JSON.</span><span class="sxs-lookup"><span data-stu-id="66a54-254">Learn the [best practices for naming your function](custom-functions-naming.md) or discover how to [localize your function](custom-functions-localize.md) using the previously described handwritten JSON method.</span></span>

## <a name="see-also"></a><span data-ttu-id="66a54-255">См. также</span><span class="sxs-lookup"><span data-stu-id="66a54-255">See also</span></span>

- [<span data-ttu-id="66a54-256">Автоматическое генерирование метаданных JSON для пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="66a54-256">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
- [<span data-ttu-id="66a54-257">Параметры параметров пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="66a54-257">Custom functions parameter options</span></span>](custom-functions-parameter-options.md)
- [<span data-ttu-id="66a54-258">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="66a54-258">Create custom functions in Excel</span></span>](custom-functions-overview.md)
