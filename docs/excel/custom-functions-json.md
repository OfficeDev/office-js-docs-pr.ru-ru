---
ms.date: 06/20/2019
description: Определение метаданных для настраиваемых функций в Excel.
title: Метаданные для пользовательских функций в Excel
localization_priority: Normal
ms.openlocfilehash: f97a339972a8ac134bd30c87b86c4701cb4b5fc4
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127872"
---
# <a name="custom-functions-metadata"></a><span data-ttu-id="cee7e-103">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="cee7e-103">Custom functions metadata</span></span>

<span data-ttu-id="cee7e-104">При определении [пользовательских функций](custom-functions-overview.md) в надстройке Excel проект надстройки содержит файл метаданных JSON, который предоставляет сведения, необходимые Excel для регистрации настраиваемых функций и предоставления доступа к ним конечным пользователям.</span><span class="sxs-lookup"><span data-stu-id="cee7e-104">When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project includes a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="cee7e-105">Этот файл создается следующим образом:</span><span class="sxs-lookup"><span data-stu-id="cee7e-105">This file is generated either:</span></span>

- <span data-ttu-id="cee7e-106">В рукописном файле JSON</span><span class="sxs-lookup"><span data-stu-id="cee7e-106">By you, in a handwritten JSON file</span></span>
- <span data-ttu-id="cee7e-107">Из комментариев Жсдок, вводимых в начале функции;</span><span class="sxs-lookup"><span data-stu-id="cee7e-107">From the JSDoc comments you enter at the beginning of your function</span></span>

<span data-ttu-id="cee7e-108">Пользовательские функции регистрируются при первом запуске надстройки и после их появления для одного и того же пользователя во всех книгах.</span><span class="sxs-lookup"><span data-stu-id="cee7e-108">Custom functions are registered when the user runs the add-in for the first time and after that are available to the same user in all workbooks.</span></span>

<span data-ttu-id="cee7e-109">В этой статье описывается формат файла метаданных JSON, предполагая, что он пишется вручную.</span><span class="sxs-lookup"><span data-stu-id="cee7e-109">This article describes the format of the JSON metadata file, assuming you are writing it by hand.</span></span> <span data-ttu-id="cee7e-110">Дополнительные сведения о создании файла Жсдок комментариев JSON можно узнать в статье [Создание МЕТАДАННЫХ JSON для пользовательских функций](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="cee7e-110">For information about JSDoc comment JSON file generation, see [Generate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

<span data-ttu-id="cee7e-111">Сведения о других файлах, которые необходимо добавить в проект надстройки для включения настраиваемых функций, см. в статье [Создание пользовательских функций в Excel](custom-functions-overview.md).</span><span class="sxs-lookup"><span data-stu-id="cee7e-111">For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md).</span></span>

<span data-ttu-id="cee7e-112">Для правильной работы пользовательских функций в Excel в Интернете параметры сервера на сервере, на котором размещается JSON-файл, должны быть включены [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) .</span><span class="sxs-lookup"><span data-stu-id="cee7e-112">Server settings on the server that hosts the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel on the web.</span></span>

## <a name="example-metadata"></a><span data-ttu-id="cee7e-113">Пример метаданных</span><span class="sxs-lookup"><span data-stu-id="cee7e-113">Example metadata</span></span>

<span data-ttu-id="cee7e-114">В примере кода ниже показано содержимое JSON-файла метаданных для надстройки, определяющей настраиваемые функции.</span><span class="sxs-lookup"><span data-stu-id="cee7e-114">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="cee7e-115">В следующих за этим примером разделах приводятся подробные сведения об отдельных свойствах, представленных в этом примере JSON.</span><span class="sxs-lookup"><span data-stu-id="cee7e-115">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
      "description":  "Count up from zero",
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
      "description":  "Get the second highest number from a range",
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
> <span data-ttu-id="cee7e-116">Полный пример JSON-файла доступен в журнале транзакций [OfficeDev/Excel-Custom-functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) репозитория GitHub.</span><span class="sxs-lookup"><span data-stu-id="cee7e-116">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) GitHub repository's commit history.</span></span> <span data-ttu-id="cee7e-117">Так как проект был скорректирован для автоматического создания JSON, полный пример рукописного кода JSON доступен только в предыдущих версиях проекта.</span><span class="sxs-lookup"><span data-stu-id="cee7e-117">As the project has been adjusted to automatically generate JSON, a full sample of handwritten JSON is only available in previous versions of the project.</span></span>

## <a name="functions"></a><span data-ttu-id="cee7e-118">functions</span><span class="sxs-lookup"><span data-stu-id="cee7e-118">functions</span></span> 

<span data-ttu-id="cee7e-119">Свойство `functions` представляет собой массив объектов настраиваемых функций.</span><span class="sxs-lookup"><span data-stu-id="cee7e-119">The `functions` property is an array of custom function objects.</span></span> <span data-ttu-id="cee7e-120">В таблице ниже приведены свойства каждого объекта.</span><span class="sxs-lookup"><span data-stu-id="cee7e-120">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="cee7e-121">Свойство</span><span class="sxs-lookup"><span data-stu-id="cee7e-121">Property</span></span>  |  <span data-ttu-id="cee7e-122">Тип данных</span><span class="sxs-lookup"><span data-stu-id="cee7e-122">Data type</span></span>  |  <span data-ttu-id="cee7e-123">Обязательный</span><span class="sxs-lookup"><span data-stu-id="cee7e-123">Required</span></span>  |  <span data-ttu-id="cee7e-124">Описание</span><span class="sxs-lookup"><span data-stu-id="cee7e-124">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="cee7e-125">string</span><span class="sxs-lookup"><span data-stu-id="cee7e-125">string</span></span>  |  <span data-ttu-id="cee7e-126">Нет</span><span class="sxs-lookup"><span data-stu-id="cee7e-126">No</span></span>  |  <span data-ttu-id="cee7e-127">Описание функции, которое отображается пользователям в Excel</span><span class="sxs-lookup"><span data-stu-id="cee7e-127">The description of the function that end users see in Excel.</span></span> <span data-ttu-id="cee7e-128">(например, **преобразует значение по шкале Цельсия в температуру по шкале Фаренгейта**).</span><span class="sxs-lookup"><span data-stu-id="cee7e-128">For example, **Converts a Celsius value to Fahrenheit**.</span></span> |
|  `helpUrl`  |  <span data-ttu-id="cee7e-129">string</span><span class="sxs-lookup"><span data-stu-id="cee7e-129">string</span></span>  |   <span data-ttu-id="cee7e-130">Нет</span><span class="sxs-lookup"><span data-stu-id="cee7e-130">No</span></span>  |  <span data-ttu-id="cee7e-131">URL-адрес, по которому можно получить сведения о функции</span><span class="sxs-lookup"><span data-stu-id="cee7e-131">URL that provides information about the function.</span></span> <span data-ttu-id="cee7e-132">(отображается в области задач). Пример: `http://contoso.com/help/convertcelsiustofahrenheit.html`.</span><span class="sxs-lookup"><span data-stu-id="cee7e-132">(It is displayed in a task pane.) For example, `http://contoso.com/help/convertcelsiustofahrenheit.html`.</span></span> |
| `id`     | <span data-ttu-id="cee7e-133">string</span><span class="sxs-lookup"><span data-stu-id="cee7e-133">string</span></span> | <span data-ttu-id="cee7e-134">Да</span><span class="sxs-lookup"><span data-stu-id="cee7e-134">Yes</span></span> | <span data-ttu-id="cee7e-135">Уникальный идентификатор для функции.</span><span class="sxs-lookup"><span data-stu-id="cee7e-135">A unique ID for the function.</span></span> <span data-ttu-id="cee7e-136">Этот идентификатор может содержать только буквы, цифры и точки и не может изменяться после настройки.</span><span class="sxs-lookup"><span data-stu-id="cee7e-136">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span> |
|  `name`  |  <span data-ttu-id="cee7e-137">string</span><span class="sxs-lookup"><span data-stu-id="cee7e-137">string</span></span>  |  <span data-ttu-id="cee7e-138">Да</span><span class="sxs-lookup"><span data-stu-id="cee7e-138">Yes</span></span>  |  <span data-ttu-id="cee7e-139">Имя функции, которое отображается пользователям в Excel.</span><span class="sxs-lookup"><span data-stu-id="cee7e-139">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="cee7e-140">В Excel имя этой функции будет присоединено в качестве префикса пространством имен настраиваемой функции, указанным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="cee7e-140">In Excel, this function name will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
|  `options`  |  <span data-ttu-id="cee7e-141">объект</span><span class="sxs-lookup"><span data-stu-id="cee7e-141">object</span></span>  |  <span data-ttu-id="cee7e-142">Нет</span><span class="sxs-lookup"><span data-stu-id="cee7e-142">No</span></span>  |  <span data-ttu-id="cee7e-143">Позволяет настроить некоторые аспекты того, как и когда Excel выполняет функцию.</span><span class="sxs-lookup"><span data-stu-id="cee7e-143">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="cee7e-144">Дополнительные сведения см. в разделе [options](#options).</span><span class="sxs-lookup"><span data-stu-id="cee7e-144">See [options](#options) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="cee7e-145">array</span><span class="sxs-lookup"><span data-stu-id="cee7e-145">array</span></span>  |  <span data-ttu-id="cee7e-146">Да</span><span class="sxs-lookup"><span data-stu-id="cee7e-146">Yes</span></span>  |  <span data-ttu-id="cee7e-147">Массив, который определяет входные параметры для функции.</span><span class="sxs-lookup"><span data-stu-id="cee7e-147">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="cee7e-148">Дополнительные сведения см. в разделе [parameters](#parameters).</span><span class="sxs-lookup"><span data-stu-id="cee7e-148">See [parameters](#parameters)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="cee7e-149">object</span><span class="sxs-lookup"><span data-stu-id="cee7e-149">object</span></span>  |  <span data-ttu-id="cee7e-150">Да</span><span class="sxs-lookup"><span data-stu-id="cee7e-150">Yes</span></span>  |  <span data-ttu-id="cee7e-151">Объект, который определяет тип информации, возвращаемый функцией.</span><span class="sxs-lookup"><span data-stu-id="cee7e-151">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="cee7e-152">Дополнительные сведения см. в разделе [result](#result).</span><span class="sxs-lookup"><span data-stu-id="cee7e-152">See [result](#result) for details.</span></span> |

## <a name="options"></a><span data-ttu-id="cee7e-153">options</span><span class="sxs-lookup"><span data-stu-id="cee7e-153">options</span></span>

<span data-ttu-id="cee7e-154">Объект `options` позволяет настроить некоторые аспекты того, как и когда Excel выполняет функцию.</span><span class="sxs-lookup"><span data-stu-id="cee7e-154">The `options` object enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="cee7e-155">В таблице ниже приведены свойства объекта `options`.</span><span class="sxs-lookup"><span data-stu-id="cee7e-155">The following table lists the properties of the `options` object.</span></span>

|  <span data-ttu-id="cee7e-156">Свойство</span><span class="sxs-lookup"><span data-stu-id="cee7e-156">Property</span></span>  |  <span data-ttu-id="cee7e-157">Тип данных</span><span class="sxs-lookup"><span data-stu-id="cee7e-157">Data type</span></span>  |  <span data-ttu-id="cee7e-158">Обязательный</span><span class="sxs-lookup"><span data-stu-id="cee7e-158">Required</span></span>  |  <span data-ttu-id="cee7e-159">Описание</span><span class="sxs-lookup"><span data-stu-id="cee7e-159">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="cee7e-160">boolean</span><span class="sxs-lookup"><span data-stu-id="cee7e-160">boolean</span></span>  |  <span data-ttu-id="cee7e-161">Нет</span><span class="sxs-lookup"><span data-stu-id="cee7e-161">No</span></span><br/><br/><span data-ttu-id="cee7e-162">Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="cee7e-162">Default value is `false`.</span></span>  |  <span data-ttu-id="cee7e-163">Если это свойство имеет значение `true`, Excel будет вызывать обработчик `CancelableInvocation` каждый раз, когда пользователь будет предпринимать действия, которые приводят к отмене функции (например, вручную вызывает пересчет или редактирует ячейку, на которую ссылается функция).</span><span class="sxs-lookup"><span data-stu-id="cee7e-163">If `true`, Excel calls the `CancelableInvocation` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="cee7e-164">Функции, которые можно отменять, обычно используются только для асинхронных функций, которые возвращают один результат и нуждаются в обработке отмены запроса данных.</span><span class="sxs-lookup"><span data-stu-id="cee7e-164">Cancelable functions are typically only used for asynchronous functions that return a single result and need to handle the cancellation of a request for data.</span></span> <span data-ttu-id="cee7e-165">Функция не может быть одновременно потоковой и отмены.</span><span class="sxs-lookup"><span data-stu-id="cee7e-165">A function cannot be both streaming and cancelable.</span></span> <span data-ttu-id="cee7e-166">Более подробную информацию можно найти в заметке около конца [функции потоковой передачи](custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="cee7e-166">For more information, see the note near the end of [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> |
|  `requiresAddress`  | <span data-ttu-id="cee7e-167">boolean</span><span class="sxs-lookup"><span data-stu-id="cee7e-167">boolean</span></span> | <span data-ttu-id="cee7e-168">Нет</span><span class="sxs-lookup"><span data-stu-id="cee7e-168">No</span></span> <br/><br/><span data-ttu-id="cee7e-169">Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="cee7e-169">Default value is `false`.</span></span> | <br /><br /> <span data-ttu-id="cee7e-170">Если этот параметр имеет значение true, пользовательская функция может получить доступ к адресу ячейки, которая вызвала пользовательскую функцию.</span><span class="sxs-lookup"><span data-stu-id="cee7e-170">If true, your custom function can access the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="cee7e-171">Чтобы получить адрес ячейки, которая вызвала пользовательскую функцию, используйте context. Address в пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="cee7e-171">To get the address of the cell that invoked your custom function, use context.address in your custom function.</span></span> <span data-ttu-id="cee7e-172">Дополнительные сведения см. в статье [Определение того, какая ячейка вызывала пользовательскую функцию](/office/dev/add-ins/excel/custom-functions-overview#determine-which-cell-invoked-your-custom-function).</span><span class="sxs-lookup"><span data-stu-id="cee7e-172">For more information, see [Determine which cell invoked your custom function](/office/dev/add-ins/excel/custom-functions-overview#determine-which-cell-invoked-your-custom-function).</span></span> <span data-ttu-id="cee7e-173">Пользовательские функции не могут быть заданы как потоковые, так и Рекуиресаддресс.</span><span class="sxs-lookup"><span data-stu-id="cee7e-173">Custom functions cannot be set as both streaming and requiresAddress.</span></span> <span data-ttu-id="cee7e-174">При использовании этого параметра параметр "вызов" должен быть последним параметром, переданным в параметрах.</span><span class="sxs-lookup"><span data-stu-id="cee7e-174">When using this option, the 'invocation' parameter must be the last parameter passed in options.</span></span> |
|  `stream`  |  <span data-ttu-id="cee7e-175">boolean</span><span class="sxs-lookup"><span data-stu-id="cee7e-175">boolean</span></span>  |  <span data-ttu-id="cee7e-176">Нет</span><span class="sxs-lookup"><span data-stu-id="cee7e-176">No</span></span><br/><br/><span data-ttu-id="cee7e-177">Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="cee7e-177">Default value is `false`.</span></span>  |  <span data-ttu-id="cee7e-178">Если это свойство имеет значение `true`, функция может выводить значение в ячейку несколько раз, даже если вызвана всего единожды.</span><span class="sxs-lookup"><span data-stu-id="cee7e-178">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="cee7e-179">Этот параметр полезен для быстро изменяющихся источников данных, таких как цена акций.</span><span class="sxs-lookup"><span data-stu-id="cee7e-179">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="cee7e-180">Функция не должна содержать оператор `return`.</span><span class="sxs-lookup"><span data-stu-id="cee7e-180">The function should have no `return` statement.</span></span> <span data-ttu-id="cee7e-181">Вместо этого результирующее значение передается как аргумент метода обратного вызова `StreamingInvocation.setResult`.</span><span class="sxs-lookup"><span data-stu-id="cee7e-181">Instead, the result value is passed as the argument of the `StreamingInvocation.setResult` callback method.</span></span> <span data-ttu-id="cee7e-182">Дополнительные сведения см. в разделе [Потоковые функции](custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="cee7e-182">For more information, see [Streaming functions](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> |
|  `volatile`  | <span data-ttu-id="cee7e-183">boolean</span><span class="sxs-lookup"><span data-stu-id="cee7e-183">boolean</span></span> | <span data-ttu-id="cee7e-184">Нет</span><span class="sxs-lookup"><span data-stu-id="cee7e-184">No</span></span> <br/><br/><span data-ttu-id="cee7e-185">Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="cee7e-185">Default value is `false`.</span></span> | <br /><br /> <span data-ttu-id="cee7e-186">Если присвоено значение `true`, функция пересчитывается при каждом выполнении пересчета в Excel, а не только при изменении зависимых значений формулы.</span><span class="sxs-lookup"><span data-stu-id="cee7e-186">If `true`, the function will recalculate each time Excel recalculates, instead of only when the formula's dependent values have changed.</span></span> <span data-ttu-id="cee7e-187">Функция не может быть одновременно потоковой и переменной.</span><span class="sxs-lookup"><span data-stu-id="cee7e-187">A function cannot be both streaming and volatile.</span></span> <span data-ttu-id="cee7e-188">Если обоим свойствам `stream` и `volatile` присвоено значение `true`, параметр переменности будет игнорироваться.</span><span class="sxs-lookup"><span data-stu-id="cee7e-188">If the `stream` and `volatile` properties are both set to `true`, the volatile option will be ignored.</span></span> |

## <a name="parameters"></a><span data-ttu-id="cee7e-189">parameters</span><span class="sxs-lookup"><span data-stu-id="cee7e-189">parameters</span></span>

<span data-ttu-id="cee7e-190">Свойство `parameters` представляет собой массив объектов параметров.</span><span class="sxs-lookup"><span data-stu-id="cee7e-190">The `parameters` property is an array of parameter objects.</span></span> <span data-ttu-id="cee7e-191">В таблице ниже приведены свойства каждого объекта.</span><span class="sxs-lookup"><span data-stu-id="cee7e-191">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="cee7e-192">Свойство</span><span class="sxs-lookup"><span data-stu-id="cee7e-192">Property</span></span>  |  <span data-ttu-id="cee7e-193">Тип данных</span><span class="sxs-lookup"><span data-stu-id="cee7e-193">Data type</span></span>  |  <span data-ttu-id="cee7e-194">Обязательный</span><span class="sxs-lookup"><span data-stu-id="cee7e-194">Required</span></span>  |  <span data-ttu-id="cee7e-195">Описание</span><span class="sxs-lookup"><span data-stu-id="cee7e-195">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="cee7e-196">string</span><span class="sxs-lookup"><span data-stu-id="cee7e-196">string</span></span>  |  <span data-ttu-id="cee7e-197">Нет</span><span class="sxs-lookup"><span data-stu-id="cee7e-197">No</span></span> |  <span data-ttu-id="cee7e-198">Описание параметра.</span><span class="sxs-lookup"><span data-stu-id="cee7e-198">A description of the parameter.</span></span> <span data-ttu-id="cee7e-199">Отображается в IntelliSense Excel.</span><span class="sxs-lookup"><span data-stu-id="cee7e-199">This is displayed in Excel's intelliSense.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="cee7e-200">string</span><span class="sxs-lookup"><span data-stu-id="cee7e-200">string</span></span>  |  <span data-ttu-id="cee7e-201">Нет</span><span class="sxs-lookup"><span data-stu-id="cee7e-201">No</span></span>  |  <span data-ttu-id="cee7e-202">Должно быть **скалярным** (значение, отличное от массива) или **матричным** (двухмерный массив).</span><span class="sxs-lookup"><span data-stu-id="cee7e-202">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="cee7e-203">string</span><span class="sxs-lookup"><span data-stu-id="cee7e-203">string</span></span>  |  <span data-ttu-id="cee7e-204">Да</span><span class="sxs-lookup"><span data-stu-id="cee7e-204">Yes</span></span>  |  <span data-ttu-id="cee7e-205">Имя параметра.</span><span class="sxs-lookup"><span data-stu-id="cee7e-205">The name of the parameter.</span></span> <span data-ttu-id="cee7e-206">Оно отображается в IntelliSense Excel.</span><span class="sxs-lookup"><span data-stu-id="cee7e-206">This name is displayed in Excel's intelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="cee7e-207">string</span><span class="sxs-lookup"><span data-stu-id="cee7e-207">string</span></span>  |  <span data-ttu-id="cee7e-208">Нет</span><span class="sxs-lookup"><span data-stu-id="cee7e-208">No</span></span>  |  <span data-ttu-id="cee7e-209">Тип данных параметра.</span><span class="sxs-lookup"><span data-stu-id="cee7e-209">The data type of the parameter.</span></span> <span data-ttu-id="cee7e-210">Может иметь значение **boolean**, **number**, **string** или **any**, что позволяет использовать любой из трех предыдущих типов.</span><span class="sxs-lookup"><span data-stu-id="cee7e-210">Can be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> <span data-ttu-id="cee7e-211">Если это свойство не задано, по умолчанию устанавливается тип данных **any**.</span><span class="sxs-lookup"><span data-stu-id="cee7e-211">If this property is not specified, the data type defaults to **any**.</span></span> |
|  `optional`  | <span data-ttu-id="cee7e-212">boolean</span><span class="sxs-lookup"><span data-stu-id="cee7e-212">boolean</span></span> | <span data-ttu-id="cee7e-213">Нет</span><span class="sxs-lookup"><span data-stu-id="cee7e-213">No</span></span> | <span data-ttu-id="cee7e-214">Если присвоено значение `true`, параметр не обязателен.</span><span class="sxs-lookup"><span data-stu-id="cee7e-214">If `true`, the parameter is optional.</span></span> |

## <a name="result"></a><span data-ttu-id="cee7e-215">result</span><span class="sxs-lookup"><span data-stu-id="cee7e-215">result</span></span>

<span data-ttu-id="cee7e-216">Объект `result` определяет тип информации, возвращаемый функцией.</span><span class="sxs-lookup"><span data-stu-id="cee7e-216">The `result` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="cee7e-217">В таблице ниже приведены свойства объекта `result`.</span><span class="sxs-lookup"><span data-stu-id="cee7e-217">The following table lists the properties of the `result` object.</span></span>

|  <span data-ttu-id="cee7e-218">Свойство</span><span class="sxs-lookup"><span data-stu-id="cee7e-218">Property</span></span>  |  <span data-ttu-id="cee7e-219">Тип данных</span><span class="sxs-lookup"><span data-stu-id="cee7e-219">Data type</span></span>  |  <span data-ttu-id="cee7e-220">Обязательный</span><span class="sxs-lookup"><span data-stu-id="cee7e-220">Required</span></span>  |  <span data-ttu-id="cee7e-221">Описание</span><span class="sxs-lookup"><span data-stu-id="cee7e-221">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="cee7e-222">string</span><span class="sxs-lookup"><span data-stu-id="cee7e-222">string</span></span>  |  <span data-ttu-id="cee7e-223">Нет</span><span class="sxs-lookup"><span data-stu-id="cee7e-223">No</span></span>  |  <span data-ttu-id="cee7e-224">Должно быть **скалярным** (значение, отличное от массива) или **матричным** (двухмерный массив).</span><span class="sxs-lookup"><span data-stu-id="cee7e-224">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |

## <a name="next-steps"></a><span data-ttu-id="cee7e-225">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="cee7e-225">Next steps</span></span>
<span data-ttu-id="cee7e-226">Ознакомьтесь с рекомендациями [по именованию функции](custom-functions-naming.md) или [локализации функции](custom-functions-localize.md) с помощью ранее описанного рукописного метода JSON.</span><span class="sxs-lookup"><span data-stu-id="cee7e-226">Learn the [best practices for naming your function](custom-functions-naming.md) or discover how to [localize your function](custom-functions-localize.md) using the previously described handwritten JSON method.</span></span>

## <a name="see-also"></a><span data-ttu-id="cee7e-227">См. также</span><span class="sxs-lookup"><span data-stu-id="cee7e-227">See also</span></span>

* [<span data-ttu-id="cee7e-228">Автоматическое генерирование метаданных JSON для пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="cee7e-228">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="cee7e-229">Параметры параметров пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="cee7e-229">Custom functions parameter options</span></span>](custom-functions-parameter-options.md)
* [<span data-ttu-id="cee7e-230">Рекомендации по пользовательским функциям</span><span class="sxs-lookup"><span data-stu-id="cee7e-230">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="cee7e-231">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="cee7e-231">Create custom functions in Excel</span></span>](custom-functions-overview.md)