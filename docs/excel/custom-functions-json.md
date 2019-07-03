---
ms.date: 06/20/2019
description: Определение метаданных для настраиваемых функций в Excel.
title: Метаданные для пользовательских функций в Excel
localization_priority: Normal
ms.openlocfilehash: a9fbefb7ea1c5474d26b668d3a4f64ed68ae36f7
ms.sourcegitcommit: 90c2d8236c6b30d80ac2b13950028a208ef60973
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/02/2019
ms.locfileid: "35454638"
---
# <a name="custom-functions-metadata"></a><span data-ttu-id="e167b-103">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="e167b-103">Custom functions metadata</span></span>

<span data-ttu-id="e167b-104">При определении [пользовательских функций](custom-functions-overview.md) в надстройке Excel проект надстройки содержит файл метаданных JSON, который предоставляет сведения, необходимые Excel для регистрации настраиваемых функций и предоставления доступа к ним конечным пользователям.</span><span class="sxs-lookup"><span data-stu-id="e167b-104">When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project includes a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="e167b-105">Этот файл создается следующим образом:</span><span class="sxs-lookup"><span data-stu-id="e167b-105">This file is generated either:</span></span>

- <span data-ttu-id="e167b-106">В рукописном файле JSON</span><span class="sxs-lookup"><span data-stu-id="e167b-106">By you, in a handwritten JSON file</span></span>
- <span data-ttu-id="e167b-107">Из комментариев Жсдок, вводимых в начале функции;</span><span class="sxs-lookup"><span data-stu-id="e167b-107">From the JSDoc comments you enter at the beginning of your function</span></span>

<span data-ttu-id="e167b-108">Пользовательские функции регистрируются при первом запуске надстройки и после их появления для одного и того же пользователя во всех книгах.</span><span class="sxs-lookup"><span data-stu-id="e167b-108">Custom functions are registered when the user runs the add-in for the first time and after that are available to the same user in all workbooks.</span></span>

<span data-ttu-id="e167b-109">В этой статье описывается формат файла метаданных JSON, предполагая, что он пишется вручную.</span><span class="sxs-lookup"><span data-stu-id="e167b-109">This article describes the format of the JSON metadata file, assuming you are writing it by hand.</span></span> <span data-ttu-id="e167b-110">Дополнительные сведения о создании файла Жсдок комментариев JSON можно узнать в статье [Создание МЕТАДАННЫХ JSON для пользовательских функций](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="e167b-110">For information about JSDoc comment JSON file generation, see [Generate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

<span data-ttu-id="e167b-111">Сведения о других файлах, которые необходимо добавить в проект надстройки для включения настраиваемых функций, см. в статье [Создание пользовательских функций в Excel](custom-functions-overview.md).</span><span class="sxs-lookup"><span data-stu-id="e167b-111">For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md).</span></span>

<span data-ttu-id="e167b-112">Для правильной работы пользовательских функций в Excel в Интернете параметры сервера на сервере, на котором размещается JSON-файл, должны быть включены [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) .</span><span class="sxs-lookup"><span data-stu-id="e167b-112">Server settings on the server that hosts the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel on the web.</span></span>

## <a name="example-metadata"></a><span data-ttu-id="e167b-113">Пример метаданных</span><span class="sxs-lookup"><span data-stu-id="e167b-113">Example metadata</span></span>

<span data-ttu-id="e167b-114">В примере кода ниже показано содержимое JSON-файла метаданных для надстройки, определяющей настраиваемые функции.</span><span class="sxs-lookup"><span data-stu-id="e167b-114">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="e167b-115">В следующих за этим примером разделах приводятся подробные сведения об отдельных свойствах, представленных в этом примере JSON.</span><span class="sxs-lookup"><span data-stu-id="e167b-115">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
> <span data-ttu-id="e167b-116">Полный пример JSON-файла доступен в журнале транзакций [OfficeDev/Excel-Custom-functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) репозитория GitHub.</span><span class="sxs-lookup"><span data-stu-id="e167b-116">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/77760adb1dcc53469183049bea08196734dbc114/config/customfunctions.json) GitHub repository's commit history.</span></span> <span data-ttu-id="e167b-117">Так как проект был скорректирован для автоматического создания JSON, полный пример рукописного кода JSON доступен только в предыдущих версиях проекта.</span><span class="sxs-lookup"><span data-stu-id="e167b-117">As the project has been adjusted to automatically generate JSON, a full sample of handwritten JSON is only available in previous versions of the project.</span></span>

## <a name="functions"></a><span data-ttu-id="e167b-118">functions</span><span class="sxs-lookup"><span data-stu-id="e167b-118">functions</span></span> 

<span data-ttu-id="e167b-119">Свойство `functions` представляет собой массив объектов настраиваемых функций.</span><span class="sxs-lookup"><span data-stu-id="e167b-119">The `functions` property is an array of custom function objects.</span></span> <span data-ttu-id="e167b-120">В таблице ниже приведены свойства каждого объекта.</span><span class="sxs-lookup"><span data-stu-id="e167b-120">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="e167b-121">Свойство</span><span class="sxs-lookup"><span data-stu-id="e167b-121">Property</span></span>  |  <span data-ttu-id="e167b-122">Тип данных</span><span class="sxs-lookup"><span data-stu-id="e167b-122">Data type</span></span>  |  <span data-ttu-id="e167b-123">Обязательный</span><span class="sxs-lookup"><span data-stu-id="e167b-123">Required</span></span>  |  <span data-ttu-id="e167b-124">Описание</span><span class="sxs-lookup"><span data-stu-id="e167b-124">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="e167b-125">string</span><span class="sxs-lookup"><span data-stu-id="e167b-125">string</span></span>  |  <span data-ttu-id="e167b-126">Нет</span><span class="sxs-lookup"><span data-stu-id="e167b-126">No</span></span>  |  <span data-ttu-id="e167b-127">Описание функции, которое отображается пользователям в Excel</span><span class="sxs-lookup"><span data-stu-id="e167b-127">The description of the function that end users see in Excel.</span></span> <span data-ttu-id="e167b-128">(например, **преобразует значение по шкале Цельсия в температуру по шкале Фаренгейта**).</span><span class="sxs-lookup"><span data-stu-id="e167b-128">For example, **Converts a Celsius value to Fahrenheit**.</span></span> |
|  `helpUrl`  |  <span data-ttu-id="e167b-129">string</span><span class="sxs-lookup"><span data-stu-id="e167b-129">string</span></span>  |   <span data-ttu-id="e167b-130">Нет</span><span class="sxs-lookup"><span data-stu-id="e167b-130">No</span></span>  |  <span data-ttu-id="e167b-131">URL-адрес, по которому можно получить сведения о функции</span><span class="sxs-lookup"><span data-stu-id="e167b-131">URL that provides information about the function.</span></span> <span data-ttu-id="e167b-132">(отображается в области задач). Пример: `http://contoso.com/help/convertcelsiustofahrenheit.html`.</span><span class="sxs-lookup"><span data-stu-id="e167b-132">(It is displayed in a task pane.) For example, `http://contoso.com/help/convertcelsiustofahrenheit.html`.</span></span> |
| `id`     | <span data-ttu-id="e167b-133">string</span><span class="sxs-lookup"><span data-stu-id="e167b-133">string</span></span> | <span data-ttu-id="e167b-134">Да</span><span class="sxs-lookup"><span data-stu-id="e167b-134">Yes</span></span> | <span data-ttu-id="e167b-135">Уникальный идентификатор для функции.</span><span class="sxs-lookup"><span data-stu-id="e167b-135">A unique ID for the function.</span></span> <span data-ttu-id="e167b-136">Этот идентификатор может содержать только буквы, цифры и точки и не может изменяться после настройки.</span><span class="sxs-lookup"><span data-stu-id="e167b-136">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span> |
|  `name`  |  <span data-ttu-id="e167b-137">string</span><span class="sxs-lookup"><span data-stu-id="e167b-137">string</span></span>  |  <span data-ttu-id="e167b-138">Да</span><span class="sxs-lookup"><span data-stu-id="e167b-138">Yes</span></span>  |  <span data-ttu-id="e167b-139">Имя функции, которое отображается пользователям в Excel.</span><span class="sxs-lookup"><span data-stu-id="e167b-139">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="e167b-140">В Excel имя этой функции будет присоединено в качестве префикса пространством имен настраиваемой функции, указанным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="e167b-140">In Excel, this function name will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
|  `options`  |  <span data-ttu-id="e167b-141">объект</span><span class="sxs-lookup"><span data-stu-id="e167b-141">object</span></span>  |  <span data-ttu-id="e167b-142">Нет</span><span class="sxs-lookup"><span data-stu-id="e167b-142">No</span></span>  |  <span data-ttu-id="e167b-143">Позволяет настроить некоторые аспекты того, как и когда Excel выполняет функцию.</span><span class="sxs-lookup"><span data-stu-id="e167b-143">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="e167b-144">Дополнительные сведения см. в разделе [options](#options).</span><span class="sxs-lookup"><span data-stu-id="e167b-144">See [options](#options) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="e167b-145">array</span><span class="sxs-lookup"><span data-stu-id="e167b-145">array</span></span>  |  <span data-ttu-id="e167b-146">Да</span><span class="sxs-lookup"><span data-stu-id="e167b-146">Yes</span></span>  |  <span data-ttu-id="e167b-147">Массив, который определяет входные параметры для функции.</span><span class="sxs-lookup"><span data-stu-id="e167b-147">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="e167b-148">Дополнительные сведения см. в разделе [parameters](#parameters).</span><span class="sxs-lookup"><span data-stu-id="e167b-148">See [parameters](#parameters)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="e167b-149">object</span><span class="sxs-lookup"><span data-stu-id="e167b-149">object</span></span>  |  <span data-ttu-id="e167b-150">Да</span><span class="sxs-lookup"><span data-stu-id="e167b-150">Yes</span></span>  |  <span data-ttu-id="e167b-151">Объект, который определяет тип информации, возвращаемый функцией.</span><span class="sxs-lookup"><span data-stu-id="e167b-151">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="e167b-152">Дополнительные сведения см. в разделе [result](#result).</span><span class="sxs-lookup"><span data-stu-id="e167b-152">See [result](#result) for details.</span></span> |

## <a name="options"></a><span data-ttu-id="e167b-153">options</span><span class="sxs-lookup"><span data-stu-id="e167b-153">options</span></span>

<span data-ttu-id="e167b-154">Объект `options` позволяет настроить некоторые аспекты того, как и когда Excel выполняет функцию.</span><span class="sxs-lookup"><span data-stu-id="e167b-154">The `options` object enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="e167b-155">В таблице ниже приведены свойства объекта `options`.</span><span class="sxs-lookup"><span data-stu-id="e167b-155">The following table lists the properties of the `options` object.</span></span>

|  <span data-ttu-id="e167b-156">Свойство</span><span class="sxs-lookup"><span data-stu-id="e167b-156">Property</span></span>  |  <span data-ttu-id="e167b-157">Тип данных</span><span class="sxs-lookup"><span data-stu-id="e167b-157">Data type</span></span>  |  <span data-ttu-id="e167b-158">Обязательный</span><span class="sxs-lookup"><span data-stu-id="e167b-158">Required</span></span>  |  <span data-ttu-id="e167b-159">Описание</span><span class="sxs-lookup"><span data-stu-id="e167b-159">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="e167b-160">boolean</span><span class="sxs-lookup"><span data-stu-id="e167b-160">boolean</span></span>  |  <span data-ttu-id="e167b-161">Нет</span><span class="sxs-lookup"><span data-stu-id="e167b-161">No</span></span><br/><br/><span data-ttu-id="e167b-162">Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="e167b-162">Default value is `false`.</span></span>  |  <span data-ttu-id="e167b-163">Если это свойство имеет значение `true`, Excel будет вызывать обработчик `CancelableInvocation` каждый раз, когда пользователь будет предпринимать действия, которые приводят к отмене функции (например, вручную вызывает пересчет или редактирует ячейку, на которую ссылается функция).</span><span class="sxs-lookup"><span data-stu-id="e167b-163">If `true`, Excel calls the `CancelableInvocation` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="e167b-164">Функции, которые можно отменять, обычно используются только для асинхронных функций, которые возвращают один результат и нуждаются в обработке отмены запроса данных.</span><span class="sxs-lookup"><span data-stu-id="e167b-164">Cancelable functions are typically only used for asynchronous functions that return a single result and need to handle the cancellation of a request for data.</span></span> <span data-ttu-id="e167b-165">Функция не может быть одновременно потоковой и отмены.</span><span class="sxs-lookup"><span data-stu-id="e167b-165">A function cannot be both streaming and cancelable.</span></span> <span data-ttu-id="e167b-166">Более подробную информацию можно найти в заметке около конца [функции потоковой передачи](custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="e167b-166">For more information, see the note near the end of [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> |
|  `requiresAddress`  | <span data-ttu-id="e167b-167">boolean</span><span class="sxs-lookup"><span data-stu-id="e167b-167">boolean</span></span> | <span data-ttu-id="e167b-168">Нет</span><span class="sxs-lookup"><span data-stu-id="e167b-168">No</span></span> <br/><br/><span data-ttu-id="e167b-169">Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="e167b-169">Default value is `false`.</span></span> | <span data-ttu-id="e167b-170">Если `true`пользовательская функция может получить доступ к адресу ячейки, которая вызвала пользовательскую функцию.</span><span class="sxs-lookup"><span data-stu-id="e167b-170">If `true`, your custom function can access the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="e167b-171">Чтобы получить адрес ячейки, которая вызвала пользовательскую функцию, используйте context. Address в пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="e167b-171">To get the address of the cell that invoked your custom function, use context.address in your custom function.</span></span> <span data-ttu-id="e167b-172">Более подробную информацию можно узнать в разделе [Address Parameter Cell](/office/dev/add-ins/excel/custom-functions-parameter-options#addressing-cells-context-parameter).</span><span class="sxs-lookup"><span data-stu-id="e167b-172">For more information, see [Addressing cell's context parameter](/office/dev/add-ins/excel/custom-functions-parameter-options#addressing-cells-context-parameter).</span></span> <span data-ttu-id="e167b-173">Пользовательские функции не могут быть заданы как потоковые, так и Рекуиресаддресс.</span><span class="sxs-lookup"><span data-stu-id="e167b-173">Custom functions cannot be set as both streaming and requiresAddress.</span></span> <span data-ttu-id="e167b-174">При использовании этого параметра параметр "вызов" должен быть последним параметром, переданным в параметрах.</span><span class="sxs-lookup"><span data-stu-id="e167b-174">When using this option, the 'invocation' parameter must be the last parameter passed in options.</span></span> |
|  `stream`  |  <span data-ttu-id="e167b-175">boolean</span><span class="sxs-lookup"><span data-stu-id="e167b-175">boolean</span></span>  |  <span data-ttu-id="e167b-176">Нет</span><span class="sxs-lookup"><span data-stu-id="e167b-176">No</span></span><br/><br/><span data-ttu-id="e167b-177">Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="e167b-177">Default value is `false`.</span></span>  |  <span data-ttu-id="e167b-178">Если это свойство имеет значение `true`, функция может выводить значение в ячейку несколько раз, даже если вызвана всего единожды.</span><span class="sxs-lookup"><span data-stu-id="e167b-178">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="e167b-179">Этот параметр полезен для быстро изменяющихся источников данных, таких как цена акций.</span><span class="sxs-lookup"><span data-stu-id="e167b-179">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="e167b-180">Функция не должна содержать оператор `return`.</span><span class="sxs-lookup"><span data-stu-id="e167b-180">The function should have no `return` statement.</span></span> <span data-ttu-id="e167b-181">Вместо этого результирующее значение передается как аргумент метода обратного вызова `StreamingInvocation.setResult`.</span><span class="sxs-lookup"><span data-stu-id="e167b-181">Instead, the result value is passed as the argument of the `StreamingInvocation.setResult` callback method.</span></span> <span data-ttu-id="e167b-182">Дополнительные сведения см. в разделе [Потоковые функции](custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="e167b-182">For more information, see [Streaming functions](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> |
|  `volatile`  | <span data-ttu-id="e167b-183">boolean</span><span class="sxs-lookup"><span data-stu-id="e167b-183">boolean</span></span> | <span data-ttu-id="e167b-184">Нет</span><span class="sxs-lookup"><span data-stu-id="e167b-184">No</span></span> <br/><br/><span data-ttu-id="e167b-185">Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="e167b-185">Default value is `false`.</span></span> | <br /><br /> <span data-ttu-id="e167b-186">Если присвоено значение `true`, функция пересчитывается при каждом выполнении пересчета в Excel, а не только при изменении зависимых значений формулы.</span><span class="sxs-lookup"><span data-stu-id="e167b-186">If `true`, the function will recalculate each time Excel recalculates, instead of only when the formula's dependent values have changed.</span></span> <span data-ttu-id="e167b-187">Функция не может быть одновременно потоковой и переменной.</span><span class="sxs-lookup"><span data-stu-id="e167b-187">A function cannot be both streaming and volatile.</span></span> <span data-ttu-id="e167b-188">Если обоим свойствам `stream` и `volatile` присвоено значение `true`, параметр переменности будет игнорироваться.</span><span class="sxs-lookup"><span data-stu-id="e167b-188">If the `stream` and `volatile` properties are both set to `true`, the volatile option will be ignored.</span></span> |

## <a name="parameters"></a><span data-ttu-id="e167b-189">parameters</span><span class="sxs-lookup"><span data-stu-id="e167b-189">parameters</span></span>

<span data-ttu-id="e167b-190">Свойство `parameters` представляет собой массив объектов параметров.</span><span class="sxs-lookup"><span data-stu-id="e167b-190">The `parameters` property is an array of parameter objects.</span></span> <span data-ttu-id="e167b-191">В таблице ниже приведены свойства каждого объекта.</span><span class="sxs-lookup"><span data-stu-id="e167b-191">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="e167b-192">Свойство</span><span class="sxs-lookup"><span data-stu-id="e167b-192">Property</span></span>  |  <span data-ttu-id="e167b-193">Тип данных</span><span class="sxs-lookup"><span data-stu-id="e167b-193">Data type</span></span>  |  <span data-ttu-id="e167b-194">Обязательный</span><span class="sxs-lookup"><span data-stu-id="e167b-194">Required</span></span>  |  <span data-ttu-id="e167b-195">Описание</span><span class="sxs-lookup"><span data-stu-id="e167b-195">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="e167b-196">string</span><span class="sxs-lookup"><span data-stu-id="e167b-196">string</span></span>  |  <span data-ttu-id="e167b-197">Нет</span><span class="sxs-lookup"><span data-stu-id="e167b-197">No</span></span> |  <span data-ttu-id="e167b-198">Описание параметра.</span><span class="sxs-lookup"><span data-stu-id="e167b-198">A description of the parameter.</span></span> <span data-ttu-id="e167b-199">Отображается в IntelliSense Excel.</span><span class="sxs-lookup"><span data-stu-id="e167b-199">This is displayed in Excel's intelliSense.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="e167b-200">string</span><span class="sxs-lookup"><span data-stu-id="e167b-200">string</span></span>  |  <span data-ttu-id="e167b-201">Нет</span><span class="sxs-lookup"><span data-stu-id="e167b-201">No</span></span>  |  <span data-ttu-id="e167b-202">Должно быть **скалярным** (значение, отличное от массива) или **матричным** (двухмерный массив).</span><span class="sxs-lookup"><span data-stu-id="e167b-202">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="e167b-203">string</span><span class="sxs-lookup"><span data-stu-id="e167b-203">string</span></span>  |  <span data-ttu-id="e167b-204">Да</span><span class="sxs-lookup"><span data-stu-id="e167b-204">Yes</span></span>  |  <span data-ttu-id="e167b-205">Имя параметра.</span><span class="sxs-lookup"><span data-stu-id="e167b-205">The name of the parameter.</span></span> <span data-ttu-id="e167b-206">Оно отображается в IntelliSense Excel.</span><span class="sxs-lookup"><span data-stu-id="e167b-206">This name is displayed in Excel's intelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="e167b-207">string</span><span class="sxs-lookup"><span data-stu-id="e167b-207">string</span></span>  |  <span data-ttu-id="e167b-208">Нет</span><span class="sxs-lookup"><span data-stu-id="e167b-208">No</span></span>  |  <span data-ttu-id="e167b-209">Тип данных параметра.</span><span class="sxs-lookup"><span data-stu-id="e167b-209">The data type of the parameter.</span></span> <span data-ttu-id="e167b-210">Может иметь значение **boolean**, **number**, **string** или **any**, что позволяет использовать любой из трех предыдущих типов.</span><span class="sxs-lookup"><span data-stu-id="e167b-210">Can be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> <span data-ttu-id="e167b-211">Если это свойство не задано, по умолчанию устанавливается тип данных **any**.</span><span class="sxs-lookup"><span data-stu-id="e167b-211">If this property is not specified, the data type defaults to **any**.</span></span> |
|  `optional`  | <span data-ttu-id="e167b-212">boolean</span><span class="sxs-lookup"><span data-stu-id="e167b-212">boolean</span></span> | <span data-ttu-id="e167b-213">Нет</span><span class="sxs-lookup"><span data-stu-id="e167b-213">No</span></span> | <span data-ttu-id="e167b-214">Если присвоено значение `true`, параметр не обязателен.</span><span class="sxs-lookup"><span data-stu-id="e167b-214">If `true`, the parameter is optional.</span></span> |

## <a name="result"></a><span data-ttu-id="e167b-215">result</span><span class="sxs-lookup"><span data-stu-id="e167b-215">result</span></span>

<span data-ttu-id="e167b-216">Объект `result` определяет тип информации, возвращаемый функцией.</span><span class="sxs-lookup"><span data-stu-id="e167b-216">The `result` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="e167b-217">В таблице ниже приведены свойства объекта `result`.</span><span class="sxs-lookup"><span data-stu-id="e167b-217">The following table lists the properties of the `result` object.</span></span>

|  <span data-ttu-id="e167b-218">Свойство</span><span class="sxs-lookup"><span data-stu-id="e167b-218">Property</span></span>  |  <span data-ttu-id="e167b-219">Тип данных</span><span class="sxs-lookup"><span data-stu-id="e167b-219">Data type</span></span>  |  <span data-ttu-id="e167b-220">Обязательный</span><span class="sxs-lookup"><span data-stu-id="e167b-220">Required</span></span>  |  <span data-ttu-id="e167b-221">Описание</span><span class="sxs-lookup"><span data-stu-id="e167b-221">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="e167b-222">string</span><span class="sxs-lookup"><span data-stu-id="e167b-222">string</span></span>  |  <span data-ttu-id="e167b-223">Нет</span><span class="sxs-lookup"><span data-stu-id="e167b-223">No</span></span>  |  <span data-ttu-id="e167b-224">Должно быть **скалярным** (значение, отличное от массива) или **матричным** (двухмерный массив).</span><span class="sxs-lookup"><span data-stu-id="e167b-224">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |

## <a name="next-steps"></a><span data-ttu-id="e167b-225">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="e167b-225">Next steps</span></span>
<span data-ttu-id="e167b-226">Ознакомьтесь с рекомендациями [по именованию функции](custom-functions-naming.md) или [локализации функции](custom-functions-localize.md) с помощью ранее описанного рукописного метода JSON.</span><span class="sxs-lookup"><span data-stu-id="e167b-226">Learn the [best practices for naming your function](custom-functions-naming.md) or discover how to [localize your function](custom-functions-localize.md) using the previously described handwritten JSON method.</span></span>

## <a name="see-also"></a><span data-ttu-id="e167b-227">См. также</span><span class="sxs-lookup"><span data-stu-id="e167b-227">See also</span></span>

* [<span data-ttu-id="e167b-228">Автоматическое генерирование метаданных JSON для пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="e167b-228">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="e167b-229">Параметры параметров пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="e167b-229">Custom functions parameter options</span></span>](custom-functions-parameter-options.md)
* [<span data-ttu-id="e167b-230">Рекомендации по пользовательским функциям</span><span class="sxs-lookup"><span data-stu-id="e167b-230">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="e167b-231">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="e167b-231">Create custom functions in Excel</span></span>](custom-functions-overview.md)
