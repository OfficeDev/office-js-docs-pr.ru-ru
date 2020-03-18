---
ms.date: 01/14/2020
description: Определите метаданные JSON для пользовательских функций в Excel и свяжите свойства идентификатора и имени функции.
title: Метаданные для пользовательских функций в Excel
localization_priority: Normal
ms.openlocfilehash: 679087336fc7aea741c98d0104514ab96068ffbf
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719464"
---
# <a name="custom-functions-metadata"></a><span data-ttu-id="20c5a-103">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="20c5a-103">Custom functions metadata</span></span>

<span data-ttu-id="20c5a-104">Как описано в статье [Обзор пользовательских функций](custom-functions-overview.md) , проект пользовательских функций должен включать файл метаданных JSON и файл скрипта (JavaScript или TypeScript) для регистрации функции, делая ее доступной для использования.</span><span class="sxs-lookup"><span data-stu-id="20c5a-104">As described in the [custom functions overview](custom-functions-overview.md) article, a custom functions project must include both a JSON metadata file and a script (either JavaScript or TypeScript) file to register a function, making it available for use.</span></span> <span data-ttu-id="20c5a-105">Пользовательские функции регистрируются при первом запуске надстройки и после их появления для одного и того же пользователя во всех книгах.</span><span class="sxs-lookup"><span data-stu-id="20c5a-105">Custom functions are registered when the user runs the add-in for the first time and after that are available to the same user in all workbooks.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="20c5a-106">Рекомендуется использовать автоматическое создание JSON, когда это возможно, используя файлы `yo office` шаблонов, похожие на процесс, показанный в руководстве по [настраиваемым функциям Excel](../tutorials/excel-tutorial-create-custom-functions.md) , так как этот процесс проще и менее подвержен ошибкам пользователя.</span><span class="sxs-lookup"><span data-stu-id="20c5a-106">It is recommended that you use JSON autogeneration when possible, using the `yo office` scaffold files, similar to the process shown in the [Excel Custom Function tutorial](../tutorials/excel-tutorial-create-custom-functions.md) because this process is easier and less prone to user error.</span></span> <span data-ttu-id="20c5a-107">Дополнительные сведения о процессе создания JSON файла Жсдок Comment можно найти в статье [Создание МЕТАДАННЫХ JSON для пользовательских функций](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="20c5a-107">For more information on the process of JSDoc comment JSON file generation, see [Generate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

<span data-ttu-id="20c5a-108">Тем не менее, проект пользовательских функций можно сделать с нуля. для этого необходимо выполнить следующие действия:</span><span class="sxs-lookup"><span data-stu-id="20c5a-108">However, you can make a custom functions project from scratch; it requires that you:</span></span>

- <span data-ttu-id="20c5a-109">Создание файла JSON вручную</span><span class="sxs-lookup"><span data-stu-id="20c5a-109">Write your JSON file by hand</span></span>
- <span data-ttu-id="20c5a-110">Убедитесь, что файл манифеста подключен к файлу JSON, созданному вручную</span><span class="sxs-lookup"><span data-stu-id="20c5a-110">Check that your manifest file is connected to your hand-authored JSON file</span></span>
- <span data-ttu-id="20c5a-111">Свяжите функции `id` и `name` свойства в файле скрипта, чтобы зарегистрировать функции</span><span class="sxs-lookup"><span data-stu-id="20c5a-111">Associate your functions' `id` and `name` properties in the script file in order to register your functions</span></span>

<span data-ttu-id="20c5a-112">В этой статье рассказывается, как выполнить все три из этих действий.</span><span class="sxs-lookup"><span data-stu-id="20c5a-112">This article will show you how to do all three of these steps.</span></span>

<span data-ttu-id="20c5a-113">На следующем рисунке показано различие между файлами `yo office` формирования шаблонов и записью JSON с нуля.</span><span class="sxs-lookup"><span data-stu-id="20c5a-113">The following image explains the differences between using `yo office` scaffold files and writing JSON from scratch.</span></span>
<span data-ttu-id="20c5a-114">![Изображение различий при использовании Yo Office и написании собственного JSON](../images/custom-functions-json.png)</span><span class="sxs-lookup"><span data-stu-id="20c5a-114">![Image of differences between using Yo Office and writing your own JSON](../images/custom-functions-json.png)</span></span>

> [!NOTE]
> <span data-ttu-id="20c5a-115">В отличие от файлов `yo office` шаблонов, необходимо подключить манифест к созданному файлу JSON с помощью `<Resources>` раздела в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="20c5a-115">In contrast with the `yo office` scaffold files, you need to connect your manifest to the JSON file you create, through the `<Resources>` section in your XML manifest file.</span></span> <span data-ttu-id="20c5a-116">Обратите внимание, что параметры [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) для сервера, на котором размещается JSON-файл, должны быть включены, чтобы пользовательские функции правильно работали в Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="20c5a-116">Note that the server settings on the server that hosts the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel on the web.</span></span>

## <a name="authoring-metadata-and-connecting-to-the-manifest"></a><span data-ttu-id="20c5a-117">Создание метаданных и подключение к манифесту</span><span class="sxs-lookup"><span data-stu-id="20c5a-117">Authoring metadata and connecting to the manifest</span></span>

<span data-ttu-id="20c5a-118">Необходимо создать файл JSON в проекте и предоставить все сведения о функциях, в том числе параметры функции.</span><span class="sxs-lookup"><span data-stu-id="20c5a-118">You need to create a JSON file in your project and provide all the details about your functions in it, such as the function's parameters.</span></span> <span data-ttu-id="20c5a-119">В [приведенном ниже примере метаданных](#json-metadata-example) и [справочнике по метаданным](#metadata-reference) представлен полный список свойств функций.</span><span class="sxs-lookup"><span data-stu-id="20c5a-119">See the [following metadata example](#json-metadata-example) and [the metadata reference](#metadata-reference) for a complete list of function properties.</span></span>

<span data-ttu-id="20c5a-120">Также необходимо убедиться, что XML-файл манифеста ссылается на JSON-файл в `<Resources>` разделе, как в следующем примере.</span><span class="sxs-lookup"><span data-stu-id="20c5a-120">You also need to make sure your XML manifest file references your JSON file in the `<Resources>` section, similar to the following example.</span></span>

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

## <a name="json-metadata-example"></a><span data-ttu-id="20c5a-121">Пример метаданных JSON</span><span class="sxs-lookup"><span data-stu-id="20c5a-121">JSON metadata example</span></span>

<span data-ttu-id="20c5a-122">В примере кода ниже показано содержимое JSON-файла метаданных для надстройки, определяющей настраиваемые функции.</span><span class="sxs-lookup"><span data-stu-id="20c5a-122">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="20c5a-123">В следующих за этим примером разделах приводятся подробные сведения об отдельных свойствах, представленных в этом примере JSON.</span><span class="sxs-lookup"><span data-stu-id="20c5a-123">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
> <span data-ttu-id="20c5a-124">Полный пример JSON-файла доступен в журнале транзакций [OfficeDev/Excel-Custom-functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) репозитория GitHub.</span><span class="sxs-lookup"><span data-stu-id="20c5a-124">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) GitHub repository's commit history.</span></span> <span data-ttu-id="20c5a-125">Так как проект был скорректирован для автоматического создания JSON, полный пример рукописного кода JSON доступен только в предыдущих версиях проекта.</span><span class="sxs-lookup"><span data-stu-id="20c5a-125">As the project has been adjusted to automatically generate JSON, a full sample of handwritten JSON is only available in previous versions of the project.</span></span>

## <a name="metadata-reference"></a><span data-ttu-id="20c5a-126">Справка по метаданным</span><span class="sxs-lookup"><span data-stu-id="20c5a-126">Metadata reference</span></span>

### <a name="functions"></a><span data-ttu-id="20c5a-127">functions</span><span class="sxs-lookup"><span data-stu-id="20c5a-127">functions</span></span>

<span data-ttu-id="20c5a-128">Свойство `functions` представляет собой массив объектов настраиваемых функций.</span><span class="sxs-lookup"><span data-stu-id="20c5a-128">The `functions` property is an array of custom function objects.</span></span> <span data-ttu-id="20c5a-129">В таблице ниже приведены свойства каждого объекта.</span><span class="sxs-lookup"><span data-stu-id="20c5a-129">The following table lists the properties of each object.</span></span>

| <span data-ttu-id="20c5a-130">Свойство</span><span class="sxs-lookup"><span data-stu-id="20c5a-130">Property</span></span>      | <span data-ttu-id="20c5a-131">Тип данных</span><span class="sxs-lookup"><span data-stu-id="20c5a-131">Data type</span></span> | <span data-ttu-id="20c5a-132">Обязательный</span><span class="sxs-lookup"><span data-stu-id="20c5a-132">Required</span></span> | <span data-ttu-id="20c5a-133">Описание</span><span class="sxs-lookup"><span data-stu-id="20c5a-133">Description</span></span>                                                                                                                                                                      |
| :------------ | :-------- | :------- | :------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `description` | <span data-ttu-id="20c5a-134">string</span><span class="sxs-lookup"><span data-stu-id="20c5a-134">string</span></span>    | <span data-ttu-id="20c5a-135">Нет</span><span class="sxs-lookup"><span data-stu-id="20c5a-135">No</span></span>       | <span data-ttu-id="20c5a-136">Описание функции, которое отображается пользователям в Excel</span><span class="sxs-lookup"><span data-stu-id="20c5a-136">The description of the function that end users see in Excel.</span></span> <span data-ttu-id="20c5a-137">(например, **преобразует значение по шкале Цельсия в температуру по шкале Фаренгейта**).</span><span class="sxs-lookup"><span data-stu-id="20c5a-137">For example, **Converts a Celsius value to Fahrenheit**.</span></span>                                                            |
| `helpUrl`     | <span data-ttu-id="20c5a-138">string</span><span class="sxs-lookup"><span data-stu-id="20c5a-138">string</span></span>    | <span data-ttu-id="20c5a-139">Нет</span><span class="sxs-lookup"><span data-stu-id="20c5a-139">No</span></span>       | <span data-ttu-id="20c5a-140">URL-адрес, по которому можно получить сведения о функции</span><span class="sxs-lookup"><span data-stu-id="20c5a-140">URL that provides information about the function.</span></span> <span data-ttu-id="20c5a-141">(отображается в области задач). Пример: `http://contoso.com/help/convertcelsiustofahrenheit.html`.</span><span class="sxs-lookup"><span data-stu-id="20c5a-141">(It is displayed in a task pane.) For example, `http://contoso.com/help/convertcelsiustofahrenheit.html`.</span></span>                      |
| `id`          | <span data-ttu-id="20c5a-142">string</span><span class="sxs-lookup"><span data-stu-id="20c5a-142">string</span></span>    | <span data-ttu-id="20c5a-143">Да</span><span class="sxs-lookup"><span data-stu-id="20c5a-143">Yes</span></span>      | <span data-ttu-id="20c5a-144">Уникальный идентификатор для функции.</span><span class="sxs-lookup"><span data-stu-id="20c5a-144">A unique ID for the function.</span></span> <span data-ttu-id="20c5a-145">Этот идентификатор может содержать только буквы, цифры и точки и не может изменяться после настройки.</span><span class="sxs-lookup"><span data-stu-id="20c5a-145">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span>                                            |
| `name`        | <span data-ttu-id="20c5a-146">string</span><span class="sxs-lookup"><span data-stu-id="20c5a-146">string</span></span>    | <span data-ttu-id="20c5a-147">Да</span><span class="sxs-lookup"><span data-stu-id="20c5a-147">Yes</span></span>      | <span data-ttu-id="20c5a-148">Имя функции, которое отображается пользователям в Excel.</span><span class="sxs-lookup"><span data-stu-id="20c5a-148">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="20c5a-149">В Excel имя этой функции будет присоединено в качестве префикса пространством имен настраиваемой функции, указанным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="20c5a-149">In Excel, this function name will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
| `options`     | <span data-ttu-id="20c5a-150">object</span><span class="sxs-lookup"><span data-stu-id="20c5a-150">object</span></span>    | <span data-ttu-id="20c5a-151">Нет</span><span class="sxs-lookup"><span data-stu-id="20c5a-151">No</span></span>       | <span data-ttu-id="20c5a-152">Позволяет настроить некоторые аспекты того, как и когда Excel выполняет функцию.</span><span class="sxs-lookup"><span data-stu-id="20c5a-152">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="20c5a-153">Дополнительные сведения см. в разделе [options](#options).</span><span class="sxs-lookup"><span data-stu-id="20c5a-153">See [options](#options) for details.</span></span>                                                          |
| `parameters`  | <span data-ttu-id="20c5a-154">array</span><span class="sxs-lookup"><span data-stu-id="20c5a-154">array</span></span>     | <span data-ttu-id="20c5a-155">Да</span><span class="sxs-lookup"><span data-stu-id="20c5a-155">Yes</span></span>      | <span data-ttu-id="20c5a-156">Массив, который определяет входные параметры для функции.</span><span class="sxs-lookup"><span data-stu-id="20c5a-156">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="20c5a-157">Дополнительные сведения см. в разделе [Parameters](#parameters) .</span><span class="sxs-lookup"><span data-stu-id="20c5a-157">See [parameters](#parameters) for details.</span></span>                                                                             |
| `result`      | <span data-ttu-id="20c5a-158">object</span><span class="sxs-lookup"><span data-stu-id="20c5a-158">object</span></span>    | <span data-ttu-id="20c5a-159">Да</span><span class="sxs-lookup"><span data-stu-id="20c5a-159">Yes</span></span>      | <span data-ttu-id="20c5a-160">Объект, который определяет тип информации, возвращаемый функцией.</span><span class="sxs-lookup"><span data-stu-id="20c5a-160">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="20c5a-161">Дополнительные сведения см. в разделе [result](#result).</span><span class="sxs-lookup"><span data-stu-id="20c5a-161">See [result](#result) for details.</span></span>                                                                 |

### <a name="options"></a><span data-ttu-id="20c5a-162">options</span><span class="sxs-lookup"><span data-stu-id="20c5a-162">options</span></span>

<span data-ttu-id="20c5a-163">Объект `options` позволяет настроить некоторые аспекты того, как и когда Excel выполняет функцию.</span><span class="sxs-lookup"><span data-stu-id="20c5a-163">The `options` object enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="20c5a-164">В таблице ниже приведены свойства объекта `options`.</span><span class="sxs-lookup"><span data-stu-id="20c5a-164">The following table lists the properties of the `options` object.</span></span>

| <span data-ttu-id="20c5a-165">Свойство</span><span class="sxs-lookup"><span data-stu-id="20c5a-165">Property</span></span>          | <span data-ttu-id="20c5a-166">Тип данных</span><span class="sxs-lookup"><span data-stu-id="20c5a-166">Data type</span></span> | <span data-ttu-id="20c5a-167">Обязательный</span><span class="sxs-lookup"><span data-stu-id="20c5a-167">Required</span></span>                               | <span data-ttu-id="20c5a-168">Описание</span><span class="sxs-lookup"><span data-stu-id="20c5a-168">Description</span></span>                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                |
| :---------------- | :-------- | :------------------------------------- | :--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| `cancelable`      | <span data-ttu-id="20c5a-169">boolean</span><span class="sxs-lookup"><span data-stu-id="20c5a-169">boolean</span></span>   | <span data-ttu-id="20c5a-170">Нет</span><span class="sxs-lookup"><span data-stu-id="20c5a-170">No</span></span><br/><br/><span data-ttu-id="20c5a-171">Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="20c5a-171">Default value is `false`.</span></span>  | <span data-ttu-id="20c5a-172">Если это свойство имеет значение `true`, Excel будет вызывать обработчик `CancelableInvocation` каждый раз, когда пользователь будет предпринимать действия, которые приводят к отмене функции (например, вручную вызывает пересчет или редактирует ячейку, на которую ссылается функция).</span><span class="sxs-lookup"><span data-stu-id="20c5a-172">If `true`, Excel calls the `CancelableInvocation` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="20c5a-173">Функции, которые можно отменять, обычно используются только для асинхронных функций, которые возвращают один результат и нуждаются в обработке отмены запроса данных.</span><span class="sxs-lookup"><span data-stu-id="20c5a-173">Cancelable functions are typically only used for asynchronous functions that return a single result and need to handle the cancellation of a request for data.</span></span> <span data-ttu-id="20c5a-174">Функция не может быть одновременно потоковой и отмены.</span><span class="sxs-lookup"><span data-stu-id="20c5a-174">A function cannot be both streaming and cancelable.</span></span> <span data-ttu-id="20c5a-175">Более подробную информацию можно найти в заметке около конца [функции потоковой передачи](custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="20c5a-175">For more information, see the note near the end of [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> |
| `requiresAddress` | <span data-ttu-id="20c5a-176">boolean</span><span class="sxs-lookup"><span data-stu-id="20c5a-176">boolean</span></span>   | <span data-ttu-id="20c5a-177">Нет</span><span class="sxs-lookup"><span data-stu-id="20c5a-177">No</span></span> <br/><br/><span data-ttu-id="20c5a-178">Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="20c5a-178">Default value is `false`.</span></span> | <span data-ttu-id="20c5a-179">Если `true`пользовательская функция может получить доступ к адресу ячейки, которая вызвала пользовательскую функцию.</span><span class="sxs-lookup"><span data-stu-id="20c5a-179">If `true`, your custom function can access the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="20c5a-180">Чтобы получить адрес ячейки, которая вызвала пользовательскую функцию, используйте context. Address в пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="20c5a-180">To get the address of the cell that invoked your custom function, use context.address in your custom function.</span></span> <span data-ttu-id="20c5a-181">Более подробную информацию можно узнать в разделе [Address Parameter Cell](../excel/custom-functions-parameter-options.md#addressing-cells-context-parameter).</span><span class="sxs-lookup"><span data-stu-id="20c5a-181">For more information, see [Addressing cell's context parameter](../excel/custom-functions-parameter-options.md#addressing-cells-context-parameter).</span></span> <span data-ttu-id="20c5a-182">Пользовательские функции не могут быть заданы как потоковые, так и Рекуиресаддресс.</span><span class="sxs-lookup"><span data-stu-id="20c5a-182">Custom functions cannot be set as both streaming and requiresAddress.</span></span> <span data-ttu-id="20c5a-183">При использовании этого параметра параметр "вызов" должен быть последним параметром, переданным в параметрах.</span><span class="sxs-lookup"><span data-stu-id="20c5a-183">When using this option, the 'invocation' parameter must be the last parameter passed in options.</span></span>                                              |
| `stream`          | <span data-ttu-id="20c5a-184">boolean</span><span class="sxs-lookup"><span data-stu-id="20c5a-184">boolean</span></span>   | <span data-ttu-id="20c5a-185">Нет</span><span class="sxs-lookup"><span data-stu-id="20c5a-185">No</span></span><br/><br/><span data-ttu-id="20c5a-186">Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="20c5a-186">Default value is `false`.</span></span>  | <span data-ttu-id="20c5a-187">Если это свойство имеет значение `true`, функция может выводить значение в ячейку несколько раз, даже если вызвана всего единожды.</span><span class="sxs-lookup"><span data-stu-id="20c5a-187">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="20c5a-188">Этот параметр полезен для быстро изменяющихся источников данных, таких как цена акций.</span><span class="sxs-lookup"><span data-stu-id="20c5a-188">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="20c5a-189">Функция не должна содержать оператор `return`.</span><span class="sxs-lookup"><span data-stu-id="20c5a-189">The function should have no `return` statement.</span></span> <span data-ttu-id="20c5a-190">Вместо этого результирующее значение передается как аргумент метода обратного вызова `StreamingInvocation.setResult`.</span><span class="sxs-lookup"><span data-stu-id="20c5a-190">Instead, the result value is passed as the argument of the `StreamingInvocation.setResult` callback method.</span></span> <span data-ttu-id="20c5a-191">Дополнительные сведения см. в разделе [Потоковые функции](custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="20c5a-191">For more information, see [Streaming functions](custom-functions-web-reqs.md#make-a-streaming-function).</span></span>                                                                                                                                                                |
| `volatile`        | <span data-ttu-id="20c5a-192">boolean</span><span class="sxs-lookup"><span data-stu-id="20c5a-192">boolean</span></span>   | <span data-ttu-id="20c5a-193">Нет</span><span class="sxs-lookup"><span data-stu-id="20c5a-193">No</span></span> <br/><br/><span data-ttu-id="20c5a-194">Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="20c5a-194">Default value is `false`.</span></span> | <br /><br /> <span data-ttu-id="20c5a-195">Если присвоено значение `true`, функция пересчитывается при каждом выполнении пересчета в Excel, а не только при изменении зависимых значений формулы.</span><span class="sxs-lookup"><span data-stu-id="20c5a-195">If `true`, the function will recalculate each time Excel recalculates, instead of only when the formula's dependent values have changed.</span></span> <span data-ttu-id="20c5a-196">Функция не может быть одновременно потоковой и переменной.</span><span class="sxs-lookup"><span data-stu-id="20c5a-196">A function cannot be both streaming and volatile.</span></span> <span data-ttu-id="20c5a-197">Если обоим свойствам `stream` и `volatile` присвоено значение `true`, параметр переменности будет игнорироваться.</span><span class="sxs-lookup"><span data-stu-id="20c5a-197">If the `stream` and `volatile` properties are both set to `true`, the volatile option will be ignored.</span></span>                                                                                                                                                                                                                                                                                             |

### <a name="parameters"></a><span data-ttu-id="20c5a-198">parameters</span><span class="sxs-lookup"><span data-stu-id="20c5a-198">parameters</span></span>

<span data-ttu-id="20c5a-199">Свойство `parameters` представляет собой массив объектов параметров.</span><span class="sxs-lookup"><span data-stu-id="20c5a-199">The `parameters` property is an array of parameter objects.</span></span> <span data-ttu-id="20c5a-200">В таблице ниже приведены свойства каждого объекта.</span><span class="sxs-lookup"><span data-stu-id="20c5a-200">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="20c5a-201">Свойство</span><span class="sxs-lookup"><span data-stu-id="20c5a-201">Property</span></span>  |  <span data-ttu-id="20c5a-202">Тип данных</span><span class="sxs-lookup"><span data-stu-id="20c5a-202">Data type</span></span>  |  <span data-ttu-id="20c5a-203">Обязательный</span><span class="sxs-lookup"><span data-stu-id="20c5a-203">Required</span></span>  |  <span data-ttu-id="20c5a-204">Описание</span><span class="sxs-lookup"><span data-stu-id="20c5a-204">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="20c5a-205">string</span><span class="sxs-lookup"><span data-stu-id="20c5a-205">string</span></span>  |  <span data-ttu-id="20c5a-206">Нет</span><span class="sxs-lookup"><span data-stu-id="20c5a-206">No</span></span> |  <span data-ttu-id="20c5a-207">Описание параметра.</span><span class="sxs-lookup"><span data-stu-id="20c5a-207">A description of the parameter.</span></span> <span data-ttu-id="20c5a-208">Отображается в IntelliSense Excel.</span><span class="sxs-lookup"><span data-stu-id="20c5a-208">This is displayed in Excel's intelliSense.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="20c5a-209">string</span><span class="sxs-lookup"><span data-stu-id="20c5a-209">string</span></span>  |  <span data-ttu-id="20c5a-210">Нет</span><span class="sxs-lookup"><span data-stu-id="20c5a-210">No</span></span>  |  <span data-ttu-id="20c5a-211">Должно быть **скалярным** (значение, отличное от массива) или **матричным** (двухмерный массив).</span><span class="sxs-lookup"><span data-stu-id="20c5a-211">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="20c5a-212">string</span><span class="sxs-lookup"><span data-stu-id="20c5a-212">string</span></span>  |  <span data-ttu-id="20c5a-213">Да</span><span class="sxs-lookup"><span data-stu-id="20c5a-213">Yes</span></span>  |  <span data-ttu-id="20c5a-214">Имя параметра.</span><span class="sxs-lookup"><span data-stu-id="20c5a-214">The name of the parameter.</span></span> <span data-ttu-id="20c5a-215">Оно отображается в IntelliSense Excel.</span><span class="sxs-lookup"><span data-stu-id="20c5a-215">This name is displayed in Excel's intelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="20c5a-216">string</span><span class="sxs-lookup"><span data-stu-id="20c5a-216">string</span></span>  |  <span data-ttu-id="20c5a-217">Нет</span><span class="sxs-lookup"><span data-stu-id="20c5a-217">No</span></span>  |  <span data-ttu-id="20c5a-218">Тип данных параметра.</span><span class="sxs-lookup"><span data-stu-id="20c5a-218">The data type of the parameter.</span></span> <span data-ttu-id="20c5a-219">Может иметь значение **boolean**, **number**, **string** или **any**, что позволяет использовать любой из трех предыдущих типов.</span><span class="sxs-lookup"><span data-stu-id="20c5a-219">Can be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> <span data-ttu-id="20c5a-220">Если это свойство не задано, по умолчанию устанавливается тип данных **any**.</span><span class="sxs-lookup"><span data-stu-id="20c5a-220">If this property is not specified, the data type defaults to **any**.</span></span> |
|  `optional`  | <span data-ttu-id="20c5a-221">boolean</span><span class="sxs-lookup"><span data-stu-id="20c5a-221">boolean</span></span> | <span data-ttu-id="20c5a-222">Нет</span><span class="sxs-lookup"><span data-stu-id="20c5a-222">No</span></span> | <span data-ttu-id="20c5a-223">Если присвоено значение `true`, параметр не обязателен.</span><span class="sxs-lookup"><span data-stu-id="20c5a-223">If `true`, the parameter is optional.</span></span> |
|`repeating`| <span data-ttu-id="20c5a-224">boolean</span><span class="sxs-lookup"><span data-stu-id="20c5a-224">boolean</span></span> | <span data-ttu-id="20c5a-225">Нет</span><span class="sxs-lookup"><span data-stu-id="20c5a-225">No</span></span> | <span data-ttu-id="20c5a-226">If `true`, параметры заполняются из указанного массива.</span><span class="sxs-lookup"><span data-stu-id="20c5a-226">If `true`, parameters will populate from a specified array.</span></span> <span data-ttu-id="20c5a-227">Обратите внимание, что функции все повторяющиеся параметры считаются необязательными параметрами по определению.</span><span class="sxs-lookup"><span data-stu-id="20c5a-227">Note that functions all repeating parameters are considered optional parameters by definition.</span></span>  |

### <a name="result"></a><span data-ttu-id="20c5a-228">result</span><span class="sxs-lookup"><span data-stu-id="20c5a-228">result</span></span>

<span data-ttu-id="20c5a-229">Объект `result` определяет тип информации, возвращаемый функцией.</span><span class="sxs-lookup"><span data-stu-id="20c5a-229">The `result` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="20c5a-230">В таблице ниже приведены свойства объекта `result`.</span><span class="sxs-lookup"><span data-stu-id="20c5a-230">The following table lists the properties of the `result` object.</span></span>

| <span data-ttu-id="20c5a-231">Свойство</span><span class="sxs-lookup"><span data-stu-id="20c5a-231">Property</span></span>         | <span data-ttu-id="20c5a-232">Тип данных</span><span class="sxs-lookup"><span data-stu-id="20c5a-232">Data type</span></span> | <span data-ttu-id="20c5a-233">Обязательный</span><span class="sxs-lookup"><span data-stu-id="20c5a-233">Required</span></span> | <span data-ttu-id="20c5a-234">Описание</span><span class="sxs-lookup"><span data-stu-id="20c5a-234">Description</span></span>                                                                          |
| :--------------- | :-------- | :------- | :----------------------------------------------------------------------------------- |
| `dimensionality` | <span data-ttu-id="20c5a-235">string</span><span class="sxs-lookup"><span data-stu-id="20c5a-235">string</span></span>    | <span data-ttu-id="20c5a-236">Нет</span><span class="sxs-lookup"><span data-stu-id="20c5a-236">No</span></span>       | <span data-ttu-id="20c5a-237">Должно быть **скалярным** (значение, отличное от массива) или **матричным** (двухмерный массив).</span><span class="sxs-lookup"><span data-stu-id="20c5a-237">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |

## <a name="associating-function-names-with-json-metadata"></a><span data-ttu-id="20c5a-238">Сопоставление имен функций с метаданными JSON</span><span class="sxs-lookup"><span data-stu-id="20c5a-238">Associating function names with JSON metadata</span></span>

<span data-ttu-id="20c5a-239">Чтобы функция работала должным образом, необходимо связать `id` свойство функции с реализацией JavaScript.</span><span class="sxs-lookup"><span data-stu-id="20c5a-239">For a function to work properly, you need to associate the function's `id` property with the JavaScript implementation.</span></span> <span data-ttu-id="20c5a-240">Убедитесь, что существует связь, в противном случае функция не будет зарегистрирована и непригодна для работы в Excel.</span><span class="sxs-lookup"><span data-stu-id="20c5a-240">Make sure there is an association, otherwise the function will not be registered and not useable in Excel.</span></span> <span data-ttu-id="20c5a-241">В приведенном ниже примере кода показано, как выполнить связь с `CustomFunctions.associate()` помощью метода.</span><span class="sxs-lookup"><span data-stu-id="20c5a-241">The following code sample shows how to make the association using the `CustomFunctions.associate()` method.</span></span> <span data-ttu-id="20c5a-242">Пример определяет пользовательскую функцию `add` и связывает ее с объектом в файле метаданных JSON, где для свойства `id` установлено значение **ADD**.</span><span class="sxs-lookup"><span data-stu-id="20c5a-242">The sample defines the custom function `add` and associates it with the object in the JSON metadata file where the value of the `id` property is **ADD**.</span></span>

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

<span data-ttu-id="20c5a-243">В следующем JSON показаны метаданные JSON, связанные с предыдущим кодом пользовательской функции JavaScript.</span><span class="sxs-lookup"><span data-stu-id="20c5a-243">The following JSON shows the JSON metadata that is associated with the previous custom function JavaScript code.</span></span>

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

<span data-ttu-id="20c5a-244">Имейте в виду приведенные ниже рекомендации при создании пользовательских функций в файле JavaScript и указании соответствующих сведений в файле метаданных JSON.</span><span class="sxs-lookup"><span data-stu-id="20c5a-244">Keep in mind the following best practices when creating custom functions in your JavaScript file and specifying corresponding information in the JSON metadata file.</span></span>

- <span data-ttu-id="20c5a-245">Убедитесь, что в файле метаданных JSON значение каждого свойства `id` содержит только буквы, цифры и точки.</span><span class="sxs-lookup"><span data-stu-id="20c5a-245">In the JSON metadata file, ensure that the value of each `id` property contains only alphanumeric characters and periods.</span></span>

- <span data-ttu-id="20c5a-246">Убедитесь, что в файле метаданных JSON значение каждого свойства `id` уникально в пределах файла.</span><span class="sxs-lookup"><span data-stu-id="20c5a-246">In the JSON metadata file, ensure that the value of each `id` property is unique within the scope of the file.</span></span> <span data-ttu-id="20c5a-247">То есть никакие два объекта функций в файле метаданных не должны иметь одинаковое значение `id`.</span><span class="sxs-lookup"><span data-stu-id="20c5a-247">That is, no two function objects in the metadata file should have the same `id` value.</span></span>

- <span data-ttu-id="20c5a-248">Не изменяйте значение свойства `id` в файле метаданных JSON после его сопоставления с соответствующим именем функции JavaScript.</span><span class="sxs-lookup"><span data-stu-id="20c5a-248">Do not change the value of an `id` property in the JSON metadata file after it's been associated with a corresponding JavaScript function name.</span></span> <span data-ttu-id="20c5a-249">Вы можете изменить имя функции, которое отображается для конечных пользователей в Excel, путем обновления свойства `name` в файле метаданных JSON, но никогда не следует изменять значение свойства `id` после его установления.</span><span class="sxs-lookup"><span data-stu-id="20c5a-249">You can change the function name that end users see in Excel by updating the `name` property within the JSON metadata file, but you should never change the value of an `id` property after it's been established.</span></span>

- <span data-ttu-id="20c5a-250">В файле JavaScript укажите настраиваемое сопоставление функций с помощью `CustomFunctions.associate` каждой функции.</span><span class="sxs-lookup"><span data-stu-id="20c5a-250">In the JavaScript file, specify a custom function association using `CustomFunctions.associate` after each function.</span></span>

<span data-ttu-id="20c5a-251">В приведенном ниже примере показаны метаданные JSON, соответствующие функциям, определенным в этом примере кода JavaScript.</span><span class="sxs-lookup"><span data-stu-id="20c5a-251">The following sample shows the JSON metadata that corresponds to the functions defined in this JavaScript code sample.</span></span> <span data-ttu-id="20c5a-252">Значения `id` свойств `name` и представлены в верхнем регистре, что является лучшим вариантом при описании пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="20c5a-252">The `id` and `name` property values are in uppercase, which is a best practice when describing your custom functions.</span></span> <span data-ttu-id="20c5a-253">Этот код JSON необходимо добавить только в том случае, если вы готовите собственный файл JSON вручную и не используете автоматическое создание.</span><span class="sxs-lookup"><span data-stu-id="20c5a-253">You only need to add this JSON if you are preparing your own JSON file manually and not using autogeneration.</span></span> <span data-ttu-id="20c5a-254">Для получения дополнительной информации об автоформировании, ознакомьтесь со статьей [Создание МЕТАДАННЫХ JSON для пользовательских функций](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="20c5a-254">For more information on autogeneration, see [Create JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

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

## <a name="next-steps"></a><span data-ttu-id="20c5a-255">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="20c5a-255">Next steps</span></span>

<span data-ttu-id="20c5a-256">Ознакомьтесь с рекомендациями [по именованию функции](custom-functions-naming.md) или [локализации функции](custom-functions-localize.md) с помощью ранее описанного рукописного метода JSON.</span><span class="sxs-lookup"><span data-stu-id="20c5a-256">Learn the [best practices for naming your function](custom-functions-naming.md) or discover how to [localize your function](custom-functions-localize.md) using the previously described handwritten JSON method.</span></span>

## <a name="see-also"></a><span data-ttu-id="20c5a-257">См. также</span><span class="sxs-lookup"><span data-stu-id="20c5a-257">See also</span></span>

- [<span data-ttu-id="20c5a-258">Автоматическое генерирование метаданных JSON для пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="20c5a-258">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
- [<span data-ttu-id="20c5a-259">Параметры параметров пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="20c5a-259">Custom functions parameter options</span></span>](custom-functions-parameter-options.md)
- [<span data-ttu-id="20c5a-260">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="20c5a-260">Create custom functions in Excel</span></span>](custom-functions-overview.md)
