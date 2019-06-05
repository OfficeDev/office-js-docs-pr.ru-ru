---
ms.date: 05/30/2019
description: Определение метаданных для настраиваемых функций в Excel.
title: Метаданные для пользовательских функций в Excel
localization_priority: Normal
ms.openlocfilehash: e51e4e8ee89eb1f345ee0c564e9b2ff8119806b2
ms.sourcegitcommit: 567aa05d6ee6b3639f65c50188df2331b7685857
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/04/2019
ms.locfileid: "34706125"
---
# <a name="custom-functions-metadata"></a><span data-ttu-id="286c7-103">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="286c7-103">Custom functions metadata</span></span>

<span data-ttu-id="286c7-104">При определении [пользовательских функций](custom-functions-overview.md) в надстройке Excel проект надстройки содержит файл метаданных JSON, который предоставляет сведения, необходимые Excel для регистрации настраиваемых функций и предоставления доступа к ним конечным пользователям.</span><span class="sxs-lookup"><span data-stu-id="286c7-104">When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project includes a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="286c7-105">Этот файл создается следующим образом:</span><span class="sxs-lookup"><span data-stu-id="286c7-105">This file is generated either:</span></span>

- <span data-ttu-id="286c7-106">В рукописном файле JSON</span><span class="sxs-lookup"><span data-stu-id="286c7-106">By you, in a handwritten JSON file</span></span>
- <span data-ttu-id="286c7-107">Из комментариев Жсдок, вводимых в начале функции;</span><span class="sxs-lookup"><span data-stu-id="286c7-107">From the JSDoc comments you enter at the beginning of your function</span></span>

<span data-ttu-id="286c7-108">Пользовательские функции регистрируются при первом запуске надстройки и после их появления для одного и того же пользователя во всех книгах.</span><span class="sxs-lookup"><span data-stu-id="286c7-108">Custom functions are registered when the user runs the add-in for the first time and after that are available to the same user in all workbooks.</span></span>

<span data-ttu-id="286c7-109">В этой статье описывается формат файла метаданных JSON, предполагая, что он пишется вручную.</span><span class="sxs-lookup"><span data-stu-id="286c7-109">This article describes the format of the JSON metadata file, assuming you are writing it by hand.</span></span> <span data-ttu-id="286c7-110">Дополнительные сведения о создании файла Жсдок комментариев JSON можно узнать в статье [Создание МЕТАДАННЫХ JSON для пользовательских функций](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="286c7-110">For information about JSDoc comment JSON file generation, see [Generate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

<span data-ttu-id="286c7-111">Сведения о других файлах, которые необходимо добавить в проект надстройки для включения настраиваемых функций, см. в статье [Создание пользовательских функций в Excel](custom-functions-overview.md).</span><span class="sxs-lookup"><span data-stu-id="286c7-111">For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md).</span></span>

<span data-ttu-id="286c7-112">Настройки сервера на сервере, на котором размещен JSON-файл, должны включать активацию [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS), чтобы пользовательские функции сработали надлежащим образом в Excel Online.</span><span class="sxs-lookup"><span data-stu-id="286c7-112">Server settings on the server that hosts the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>

## <a name="example-metadata"></a><span data-ttu-id="286c7-113">Пример метаданных</span><span class="sxs-lookup"><span data-stu-id="286c7-113">Example metadata</span></span>

<span data-ttu-id="286c7-114">В примере кода ниже показано содержимое JSON-файла метаданных для надстройки, определяющей настраиваемые функции.</span><span class="sxs-lookup"><span data-stu-id="286c7-114">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="286c7-115">В следующих за этим примером разделах приводятся подробные сведения об отдельных свойствах, представленных в этом примере JSON.</span><span class="sxs-lookup"><span data-stu-id="286c7-115">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
> <span data-ttu-id="286c7-116">Пример готового JSON-файла приводится в репозитории GitHub [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/src/functions/functions.json).</span><span class="sxs-lookup"><span data-stu-id="286c7-116">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/src/functions/functions.json) GitHub repository.</span></span>

## <a name="functions"></a><span data-ttu-id="286c7-117">functions</span><span class="sxs-lookup"><span data-stu-id="286c7-117">functions</span></span> 

<span data-ttu-id="286c7-118">Свойство `functions` представляет собой массив объектов настраиваемых функций.</span><span class="sxs-lookup"><span data-stu-id="286c7-118">The `functions` property is an array of custom function objects.</span></span> <span data-ttu-id="286c7-119">В таблице ниже приведены свойства каждого объекта.</span><span class="sxs-lookup"><span data-stu-id="286c7-119">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="286c7-120">Свойство</span><span class="sxs-lookup"><span data-stu-id="286c7-120">Property</span></span>  |  <span data-ttu-id="286c7-121">Тип данных</span><span class="sxs-lookup"><span data-stu-id="286c7-121">Data type</span></span>  |  <span data-ttu-id="286c7-122">Обязательный</span><span class="sxs-lookup"><span data-stu-id="286c7-122">Required</span></span>  |  <span data-ttu-id="286c7-123">Описание</span><span class="sxs-lookup"><span data-stu-id="286c7-123">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="286c7-124">string</span><span class="sxs-lookup"><span data-stu-id="286c7-124">string</span></span>  |  <span data-ttu-id="286c7-125">Нет</span><span class="sxs-lookup"><span data-stu-id="286c7-125">No</span></span>  |  <span data-ttu-id="286c7-126">Описание функции, которое отображается пользователям в Excel</span><span class="sxs-lookup"><span data-stu-id="286c7-126">The description of the function that end users see in Excel.</span></span> <span data-ttu-id="286c7-127">(например, **преобразует значение по шкале Цельсия в температуру по шкале Фаренгейта**).</span><span class="sxs-lookup"><span data-stu-id="286c7-127">For example, **Converts a Celsius value to Fahrenheit**.</span></span> |
|  `helpUrl`  |  <span data-ttu-id="286c7-128">string</span><span class="sxs-lookup"><span data-stu-id="286c7-128">string</span></span>  |   <span data-ttu-id="286c7-129">Нет</span><span class="sxs-lookup"><span data-stu-id="286c7-129">No</span></span>  |  <span data-ttu-id="286c7-130">URL-адрес, по которому можно получить сведения о функции</span><span class="sxs-lookup"><span data-stu-id="286c7-130">URL that provides information about the function.</span></span> <span data-ttu-id="286c7-131">(отображается в области задач). Пример: `http://contoso.com/help/convertcelsiustofahrenheit.html`.</span><span class="sxs-lookup"><span data-stu-id="286c7-131">(It is displayed in a task pane.) For example, `http://contoso.com/help/convertcelsiustofahrenheit.html`.</span></span> |
| `id`     | <span data-ttu-id="286c7-132">string</span><span class="sxs-lookup"><span data-stu-id="286c7-132">string</span></span> | <span data-ttu-id="286c7-133">Да</span><span class="sxs-lookup"><span data-stu-id="286c7-133">Yes</span></span> | <span data-ttu-id="286c7-134">Уникальный идентификатор для функции.</span><span class="sxs-lookup"><span data-stu-id="286c7-134">A unique ID for the function.</span></span> <span data-ttu-id="286c7-135">Этот идентификатор может содержать только буквы, цифры и точки и не может изменяться после настройки.</span><span class="sxs-lookup"><span data-stu-id="286c7-135">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span> |
|  `name`  |  <span data-ttu-id="286c7-136">string</span><span class="sxs-lookup"><span data-stu-id="286c7-136">string</span></span>  |  <span data-ttu-id="286c7-137">Да</span><span class="sxs-lookup"><span data-stu-id="286c7-137">Yes</span></span>  |  <span data-ttu-id="286c7-138">Имя функции, которое отображается пользователям в Excel.</span><span class="sxs-lookup"><span data-stu-id="286c7-138">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="286c7-139">В Excel имя этой функции будет присоединено в качестве префикса пространством имен настраиваемой функции, указанным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="286c7-139">In Excel, this function name will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
|  `options`  |  <span data-ttu-id="286c7-140">объект</span><span class="sxs-lookup"><span data-stu-id="286c7-140">object</span></span>  |  <span data-ttu-id="286c7-141">Нет</span><span class="sxs-lookup"><span data-stu-id="286c7-141">No</span></span>  |  <span data-ttu-id="286c7-142">Позволяет настроить некоторые аспекты того, как и когда Excel выполняет функцию.</span><span class="sxs-lookup"><span data-stu-id="286c7-142">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="286c7-143">Дополнительные сведения см. в разделе [options](#options).</span><span class="sxs-lookup"><span data-stu-id="286c7-143">See [options](#options) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="286c7-144">array</span><span class="sxs-lookup"><span data-stu-id="286c7-144">array</span></span>  |  <span data-ttu-id="286c7-145">Да</span><span class="sxs-lookup"><span data-stu-id="286c7-145">Yes</span></span>  |  <span data-ttu-id="286c7-146">Массив, который определяет входные параметры для функции.</span><span class="sxs-lookup"><span data-stu-id="286c7-146">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="286c7-147">Дополнительные сведения см. в разделе [parameters](#parameters).</span><span class="sxs-lookup"><span data-stu-id="286c7-147">See [parameters](#parameters)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="286c7-148">object</span><span class="sxs-lookup"><span data-stu-id="286c7-148">object</span></span>  |  <span data-ttu-id="286c7-149">Да</span><span class="sxs-lookup"><span data-stu-id="286c7-149">Yes</span></span>  |  <span data-ttu-id="286c7-150">Объект, который определяет тип информации, возвращаемый функцией.</span><span class="sxs-lookup"><span data-stu-id="286c7-150">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="286c7-151">Дополнительные сведения см. в разделе [result](#result).</span><span class="sxs-lookup"><span data-stu-id="286c7-151">See [result](#result) for details.</span></span> |

## <a name="options"></a><span data-ttu-id="286c7-152">options</span><span class="sxs-lookup"><span data-stu-id="286c7-152">options</span></span>

<span data-ttu-id="286c7-153">Объект `options` позволяет настроить некоторые аспекты того, как и когда Excel выполняет функцию.</span><span class="sxs-lookup"><span data-stu-id="286c7-153">The `options` object enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="286c7-154">В таблице ниже приведены свойства объекта `options`.</span><span class="sxs-lookup"><span data-stu-id="286c7-154">The following table lists the properties of the `options` object.</span></span>

|  <span data-ttu-id="286c7-155">Свойство</span><span class="sxs-lookup"><span data-stu-id="286c7-155">Property</span></span>  |  <span data-ttu-id="286c7-156">Тип данных</span><span class="sxs-lookup"><span data-stu-id="286c7-156">Data type</span></span>  |  <span data-ttu-id="286c7-157">Обязательный</span><span class="sxs-lookup"><span data-stu-id="286c7-157">Required</span></span>  |  <span data-ttu-id="286c7-158">Описание</span><span class="sxs-lookup"><span data-stu-id="286c7-158">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="286c7-159">boolean</span><span class="sxs-lookup"><span data-stu-id="286c7-159">boolean</span></span>  |  <span data-ttu-id="286c7-160">Нет</span><span class="sxs-lookup"><span data-stu-id="286c7-160">No</span></span><br/><br/><span data-ttu-id="286c7-161">Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="286c7-161">Default value is `false`.</span></span>  |  <span data-ttu-id="286c7-162">Если это свойство имеет значение `true`, Excel будет вызывать обработчик `CancelableInvocation` каждый раз, когда пользователь будет предпринимать действия, которые приводят к отмене функции (например, вручную вызывает пересчет или редактирует ячейку, на которую ссылается функция).</span><span class="sxs-lookup"><span data-stu-id="286c7-162">If `true`, Excel calls the `CancelableInvocation` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="286c7-163">Функции, которые можно отменять, обычно используются только для асинхронных функций, которые возвращают один результат и нуждаются в обработке отмены запроса данных.</span><span class="sxs-lookup"><span data-stu-id="286c7-163">Cancelable functions are typically only used for asynchronous functions that return a single result and need to handle the cancellation of a request for data.</span></span> <span data-ttu-id="286c7-164">Функция не может быть одновременно потоковой и отмены.</span><span class="sxs-lookup"><span data-stu-id="286c7-164">A function cannot be both streaming and cancelable.</span></span> <span data-ttu-id="286c7-165">Более подробную информацию можно найти в заметке около конца [функции потоковой передачи](custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="286c7-165">For more information, see the note near the end of [Make a streaming function](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> |
|  `requiresAddress`  | <span data-ttu-id="286c7-166">boolean</span><span class="sxs-lookup"><span data-stu-id="286c7-166">boolean</span></span> | <span data-ttu-id="286c7-167">Нет</span><span class="sxs-lookup"><span data-stu-id="286c7-167">No</span></span> <br/><br/><span data-ttu-id="286c7-168">Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="286c7-168">Default value is `false`.</span></span> | <br /><br /> <span data-ttu-id="286c7-169">Если этот параметр имеет значение true, пользовательская функция может получить доступ к адресу ячейки, которая вызвала пользовательскую функцию.</span><span class="sxs-lookup"><span data-stu-id="286c7-169">If true, your custom function can access the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="286c7-170">Чтобы получить адрес ячейки, которая вызвала пользовательскую функцию, используйте context. Address в пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="286c7-170">To get the address of the cell that invoked your custom function, use context.address in your custom function.</span></span> <span data-ttu-id="286c7-171">Дополнительные сведения см. в статье [Определение того, какая ячейка вызывала пользовательскую функцию](/office/dev/add-ins/excel/custom-functions-overview#determine-which-cell-invoked-your-custom-function).</span><span class="sxs-lookup"><span data-stu-id="286c7-171">For more information, see [Determine which cell invoked your custom function](/office/dev/add-ins/excel/custom-functions-overview#determine-which-cell-invoked-your-custom-function).</span></span> <span data-ttu-id="286c7-172">Пользовательские функции не могут быть заданы как потоковые, так и Рекуиресаддресс.</span><span class="sxs-lookup"><span data-stu-id="286c7-172">Custom functions cannot be set as both streaming and requiresAddress.</span></span> <span data-ttu-id="286c7-173">При использовании этого параметра параметр "вызов" должен быть последним параметром, переданным в параметрах.</span><span class="sxs-lookup"><span data-stu-id="286c7-173">When using this option, the 'invocation' parameter must be the last parameter passed in options.</span></span> |
|  `stream`  |  <span data-ttu-id="286c7-174">boolean</span><span class="sxs-lookup"><span data-stu-id="286c7-174">boolean</span></span>  |  <span data-ttu-id="286c7-175">Нет</span><span class="sxs-lookup"><span data-stu-id="286c7-175">No</span></span><br/><br/><span data-ttu-id="286c7-176">Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="286c7-176">Default value is `false`.</span></span>  |  <span data-ttu-id="286c7-177">Если это свойство имеет значение `true`, функция может выводить значение в ячейку несколько раз, даже если вызвана всего единожды.</span><span class="sxs-lookup"><span data-stu-id="286c7-177">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="286c7-178">Этот параметр полезен для быстро изменяющихся источников данных, таких как цена акций.</span><span class="sxs-lookup"><span data-stu-id="286c7-178">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="286c7-179">Функция не должна содержать оператор `return`.</span><span class="sxs-lookup"><span data-stu-id="286c7-179">The function should have no `return` statement.</span></span> <span data-ttu-id="286c7-180">Вместо этого результирующее значение передается как аргумент метода обратного вызова `StreamingInvocation.setResult`.</span><span class="sxs-lookup"><span data-stu-id="286c7-180">Instead, the result value is passed as the argument of the `StreamingInvocation.setResult` callback method.</span></span> <span data-ttu-id="286c7-181">Дополнительные сведения см. в разделе [Потоковые функции](custom-functions-web-reqs.md#make-a-streaming-function).</span><span class="sxs-lookup"><span data-stu-id="286c7-181">For more information, see [Streaming functions](custom-functions-web-reqs.md#make-a-streaming-function).</span></span> |
|  `volatile`  | <span data-ttu-id="286c7-182">boolean</span><span class="sxs-lookup"><span data-stu-id="286c7-182">boolean</span></span> | <span data-ttu-id="286c7-183">Нет</span><span class="sxs-lookup"><span data-stu-id="286c7-183">No</span></span> <br/><br/><span data-ttu-id="286c7-184">Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="286c7-184">Default value is `false`.</span></span> | <br /><br /> <span data-ttu-id="286c7-185">Если присвоено значение `true`, функция пересчитывается при каждом выполнении пересчета в Excel, а не только при изменении зависимых значений формулы.</span><span class="sxs-lookup"><span data-stu-id="286c7-185">If `true`, the function will recalculate each time Excel recalculates, instead of only when the formula's dependent values have changed.</span></span> <span data-ttu-id="286c7-186">Функция не может быть одновременно потоковой и переменной.</span><span class="sxs-lookup"><span data-stu-id="286c7-186">A function cannot be both streaming and volatile.</span></span> <span data-ttu-id="286c7-187">Если обоим свойствам `stream` и `volatile` присвоено значение `true`, параметр переменности будет игнорироваться.</span><span class="sxs-lookup"><span data-stu-id="286c7-187">If the `stream` and `volatile` properties are both set to `true`, the volatile option will be ignored.</span></span> |

## <a name="parameters"></a><span data-ttu-id="286c7-188">parameters</span><span class="sxs-lookup"><span data-stu-id="286c7-188">parameters</span></span>

<span data-ttu-id="286c7-189">Свойство `parameters` представляет собой массив объектов параметров.</span><span class="sxs-lookup"><span data-stu-id="286c7-189">The `parameters` property is an array of parameter objects.</span></span> <span data-ttu-id="286c7-190">В таблице ниже приведены свойства каждого объекта.</span><span class="sxs-lookup"><span data-stu-id="286c7-190">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="286c7-191">Свойство</span><span class="sxs-lookup"><span data-stu-id="286c7-191">Property</span></span>  |  <span data-ttu-id="286c7-192">Тип данных</span><span class="sxs-lookup"><span data-stu-id="286c7-192">Data type</span></span>  |  <span data-ttu-id="286c7-193">Обязательный</span><span class="sxs-lookup"><span data-stu-id="286c7-193">Required</span></span>  |  <span data-ttu-id="286c7-194">Описание</span><span class="sxs-lookup"><span data-stu-id="286c7-194">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="286c7-195">string</span><span class="sxs-lookup"><span data-stu-id="286c7-195">string</span></span>  |  <span data-ttu-id="286c7-196">Нет</span><span class="sxs-lookup"><span data-stu-id="286c7-196">No</span></span> |  <span data-ttu-id="286c7-197">Описание параметра.</span><span class="sxs-lookup"><span data-stu-id="286c7-197">A description of the parameter.</span></span> <span data-ttu-id="286c7-198">Отображается в IntelliSense Excel.</span><span class="sxs-lookup"><span data-stu-id="286c7-198">This is displayed in Excel's intelliSense.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="286c7-199">string</span><span class="sxs-lookup"><span data-stu-id="286c7-199">string</span></span>  |  <span data-ttu-id="286c7-200">Нет</span><span class="sxs-lookup"><span data-stu-id="286c7-200">No</span></span>  |  <span data-ttu-id="286c7-201">Должно быть **скалярным** (значение, отличное от массива) или **матричным** (двухмерный массив).</span><span class="sxs-lookup"><span data-stu-id="286c7-201">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="286c7-202">string</span><span class="sxs-lookup"><span data-stu-id="286c7-202">string</span></span>  |  <span data-ttu-id="286c7-203">Да</span><span class="sxs-lookup"><span data-stu-id="286c7-203">Yes</span></span>  |  <span data-ttu-id="286c7-204">Имя параметра.</span><span class="sxs-lookup"><span data-stu-id="286c7-204">The name of the parameter.</span></span> <span data-ttu-id="286c7-205">Оно отображается в IntelliSense Excel.</span><span class="sxs-lookup"><span data-stu-id="286c7-205">This name is displayed in Excel's intelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="286c7-206">string</span><span class="sxs-lookup"><span data-stu-id="286c7-206">string</span></span>  |  <span data-ttu-id="286c7-207">Нет</span><span class="sxs-lookup"><span data-stu-id="286c7-207">No</span></span>  |  <span data-ttu-id="286c7-208">Тип данных параметра.</span><span class="sxs-lookup"><span data-stu-id="286c7-208">The data type of the parameter.</span></span> <span data-ttu-id="286c7-209">Может иметь значение **boolean**, **number**, **string** или **any**, что позволяет использовать любой из трех предыдущих типов.</span><span class="sxs-lookup"><span data-stu-id="286c7-209">Can be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> <span data-ttu-id="286c7-210">Если это свойство не задано, по умолчанию устанавливается тип данных **any**.</span><span class="sxs-lookup"><span data-stu-id="286c7-210">If this property is not specified, the data type defaults to **any**.</span></span> |
|  `optional`  | <span data-ttu-id="286c7-211">boolean</span><span class="sxs-lookup"><span data-stu-id="286c7-211">boolean</span></span> | <span data-ttu-id="286c7-212">Нет</span><span class="sxs-lookup"><span data-stu-id="286c7-212">No</span></span> | <span data-ttu-id="286c7-213">Если присвоено значение `true`, параметр не обязателен.</span><span class="sxs-lookup"><span data-stu-id="286c7-213">If `true`, the parameter is optional.</span></span> |

## <a name="result"></a><span data-ttu-id="286c7-214">result</span><span class="sxs-lookup"><span data-stu-id="286c7-214">result</span></span>

<span data-ttu-id="286c7-215">Объект `result` определяет тип информации, возвращаемый функцией.</span><span class="sxs-lookup"><span data-stu-id="286c7-215">The `result` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="286c7-216">В таблице ниже приведены свойства объекта `result`.</span><span class="sxs-lookup"><span data-stu-id="286c7-216">The following table lists the properties of the `result` object.</span></span>

|  <span data-ttu-id="286c7-217">Свойство</span><span class="sxs-lookup"><span data-stu-id="286c7-217">Property</span></span>  |  <span data-ttu-id="286c7-218">Тип данных</span><span class="sxs-lookup"><span data-stu-id="286c7-218">Data type</span></span>  |  <span data-ttu-id="286c7-219">Обязательный</span><span class="sxs-lookup"><span data-stu-id="286c7-219">Required</span></span>  |  <span data-ttu-id="286c7-220">Описание</span><span class="sxs-lookup"><span data-stu-id="286c7-220">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="286c7-221">string</span><span class="sxs-lookup"><span data-stu-id="286c7-221">string</span></span>  |  <span data-ttu-id="286c7-222">Нет</span><span class="sxs-lookup"><span data-stu-id="286c7-222">No</span></span>  |  <span data-ttu-id="286c7-223">Должно быть **скалярным** (значение, отличное от массива) или **матричным** (двухмерный массив).</span><span class="sxs-lookup"><span data-stu-id="286c7-223">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |

## <a name="next-steps"></a><span data-ttu-id="286c7-224">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="286c7-224">Next steps</span></span>
<span data-ttu-id="286c7-225">Ознакомьтесь с рекомендациями [по именованию функции](custom-functions-naming.md) или [локализации функции](custom-functions-localize.md) с помощью ранее описанного рукописного метода JSON.</span><span class="sxs-lookup"><span data-stu-id="286c7-225">Learn the [best practices for naming your function](custom-functions-naming.md) or discover how to [localize your function](custom-functions-localize.md) using the previously described handwritten JSON method.</span></span>

## <a name="see-also"></a><span data-ttu-id="286c7-226">См. также</span><span class="sxs-lookup"><span data-stu-id="286c7-226">See also</span></span>

* [<span data-ttu-id="286c7-227">Автоматическое генерирование метаданных JSON для пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="286c7-227">Autogenerate JSON metadata for custom functions</span></span>](custom-functions-json-autogeneration.md)
* [<span data-ttu-id="286c7-228">Параметры параметров пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="286c7-228">Custom functions parameter options</span></span>](custom-functions-parameter-options.md)
* [<span data-ttu-id="286c7-229">Рекомендации по пользовательским функциям</span><span class="sxs-lookup"><span data-stu-id="286c7-229">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="286c7-230">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="286c7-230">Create custom functions in Excel</span></span>](custom-functions-overview.md)