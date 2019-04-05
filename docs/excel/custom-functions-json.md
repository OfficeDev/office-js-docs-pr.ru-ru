---
ms.date: 03/29/2019
description: Определение метаданных для настраиваемых функций в Excel.
title: Метаданные для настраиваемых функций в Excel (предварительная версия)
localization_priority: Normal
ms.openlocfilehash: 28a9a0207f7439af164eb9ca7c4b9ed9e966b3ed
ms.sourcegitcommit: 14ceac067e0e130869b861d289edb438b5e3eff9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/04/2019
ms.locfileid: "31477553"
---
# <a name="custom-functions-metadata-preview"></a><span data-ttu-id="fa98f-103">Метаданные для настраиваемых функций (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="fa98f-103">Custom functions metadata (preview)</span></span>

<span data-ttu-id="fa98f-104">При определении [пользовательских функций](custom-functions-overview.md) в надстройке Excel проект надстройки содержит файл метаданных JSON, который предоставляет сведения, необходимые Excel для регистрации настраиваемых функций и предоставления доступа к ним конечным пользователям.</span><span class="sxs-lookup"><span data-stu-id="fa98f-104">When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project includes a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users.</span></span> <span data-ttu-id="fa98f-105">Этот файл создается следующим образом:</span><span class="sxs-lookup"><span data-stu-id="fa98f-105">This file is generated either:</span></span>

- <span data-ttu-id="fa98f-106">в рукописном файле JSON</span><span class="sxs-lookup"><span data-stu-id="fa98f-106">by you, in a handwritten JSON file</span></span>
- <span data-ttu-id="fa98f-107">из комментариев Жсдок, вводимых в начале функции;</span><span class="sxs-lookup"><span data-stu-id="fa98f-107">from the JSDoc comments you enter at the beginning of your function</span></span>

<span data-ttu-id="fa98f-108">Пользовательские функции регистрируются при первом запуске надстройки и после их появления для одного и того же пользователя во всех книгах.</span><span class="sxs-lookup"><span data-stu-id="fa98f-108">Custom functions are registered when the user runs the add-in for the first time and after that are available to the same user in all workbooks.</span></span>

<span data-ttu-id="fa98f-109">В этой статье описывается формат файла метаданных JSON, предполагая, что он пишется вручную.</span><span class="sxs-lookup"><span data-stu-id="fa98f-109">This article describes the format of the JSON metadata file, assuming you are writing it by hand.</span></span> <span data-ttu-id="fa98f-110">Дополнительные сведения о создании файла Жсдок комментариев JSON можно узнать в статье [Создание МЕТАДАННЫХ JSON для пользовательских функций](custom-functions-json-autogeneration.md).</span><span class="sxs-lookup"><span data-stu-id="fa98f-110">For information about JSDoc comment JSON file generation, see [Generate JSON metadata for custom functions](custom-functions-json-autogeneration.md).</span></span>

<span data-ttu-id="fa98f-111">Сведения о других файлах, которые необходимо добавить в проект надстройки для включения настраиваемых функций, см. в статье [Создание пользовательских функций в Excel](custom-functions-overview.md).</span><span class="sxs-lookup"><span data-stu-id="fa98f-111">For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

> <span data-ttu-id="fa98f-112">Настройки сервера на сервере, на котором размещен JSON-файл, должны включать активацию [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS), чтобы пользовательские функции сработали надлежащим образом в Excel Online.</span><span class="sxs-lookup"><span data-stu-id="fa98f-112">Server settings on the server that hosts the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>

## <a name="example-metadata"></a><span data-ttu-id="fa98f-113">Пример метаданных</span><span class="sxs-lookup"><span data-stu-id="fa98f-113">Example metadata</span></span>

<span data-ttu-id="fa98f-114">В примере кода ниже показано содержимое JSON-файла метаданных для надстройки, определяющей настраиваемые функции.</span><span class="sxs-lookup"><span data-stu-id="fa98f-114">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="fa98f-115">В следующих за этим примером разделах приводятся подробные сведения об отдельных свойствах, представленных в этом примере JSON.</span><span class="sxs-lookup"><span data-stu-id="fa98f-115">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
        "type": "string"
      },
      "parameters": []
    },
    {
      "id": "INCREMENTVALUE",
      "name": "INCREMENTVALUE", 
      "description":  "Count up from zero",
      "helpUrl": "http://www.contoso.com/help",
      "result": {
        "type": "number",
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
        "type": "number",
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
> <span data-ttu-id="fa98f-116">Пример готового JSON-файла приводится в репозитории GitHub [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/src/functions/functions.json).</span><span class="sxs-lookup"><span data-stu-id="fa98f-116">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/src/functions/functions.json) GitHub repository.</span></span>

## <a name="functions"></a><span data-ttu-id="fa98f-117">functions</span><span class="sxs-lookup"><span data-stu-id="fa98f-117">functions</span></span> 

<span data-ttu-id="fa98f-118">Свойство `functions` представляет собой массив объектов настраиваемых функций.</span><span class="sxs-lookup"><span data-stu-id="fa98f-118">The `functions` property is an array of custom function objects.</span></span> <span data-ttu-id="fa98f-119">В таблице ниже приведены свойства каждого объекта.</span><span class="sxs-lookup"><span data-stu-id="fa98f-119">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="fa98f-120">Свойство</span><span class="sxs-lookup"><span data-stu-id="fa98f-120">Property</span></span>  |  <span data-ttu-id="fa98f-121">Тип данных</span><span class="sxs-lookup"><span data-stu-id="fa98f-121">Data type</span></span>  |  <span data-ttu-id="fa98f-122">Обязательный</span><span class="sxs-lookup"><span data-stu-id="fa98f-122">Required</span></span>  |  <span data-ttu-id="fa98f-123">Описание</span><span class="sxs-lookup"><span data-stu-id="fa98f-123">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="fa98f-124">string</span><span class="sxs-lookup"><span data-stu-id="fa98f-124">string</span></span>  |  <span data-ttu-id="fa98f-125">Нет</span><span class="sxs-lookup"><span data-stu-id="fa98f-125">No</span></span>  |  <span data-ttu-id="fa98f-126">Описание функции, которое отображается пользователям в Excel</span><span class="sxs-lookup"><span data-stu-id="fa98f-126">The description of the function that end users see in Excel.</span></span> <span data-ttu-id="fa98f-127">(например, **преобразует значение по шкале Цельсия в температуру по шкале Фаренгейта**).</span><span class="sxs-lookup"><span data-stu-id="fa98f-127">For example, **Converts a Celsius value to Fahrenheit**.</span></span> |
|  `helpUrl`  |  <span data-ttu-id="fa98f-128">string</span><span class="sxs-lookup"><span data-stu-id="fa98f-128">string</span></span>  |   <span data-ttu-id="fa98f-129">Нет</span><span class="sxs-lookup"><span data-stu-id="fa98f-129">No</span></span>  |  <span data-ttu-id="fa98f-130">URL-адрес, по которому можно получить сведения о функции</span><span class="sxs-lookup"><span data-stu-id="fa98f-130">URL that provides information about the function.</span></span> <span data-ttu-id="fa98f-131">(отображается в области задач). Пример: **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span><span class="sxs-lookup"><span data-stu-id="fa98f-131">(It is displayed in a task pane.) For example, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span></span> |
| `id`     | <span data-ttu-id="fa98f-132">string</span><span class="sxs-lookup"><span data-stu-id="fa98f-132">string</span></span> | <span data-ttu-id="fa98f-133">Да</span><span class="sxs-lookup"><span data-stu-id="fa98f-133">Yes</span></span> | <span data-ttu-id="fa98f-134">Уникальный идентификатор для функции.</span><span class="sxs-lookup"><span data-stu-id="fa98f-134">A unique ID for the function.</span></span> <span data-ttu-id="fa98f-135">Этот идентификатор может содержать только буквы, цифры и точки и не может изменяться после настройки.</span><span class="sxs-lookup"><span data-stu-id="fa98f-135">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span> |
|  `name`  |  <span data-ttu-id="fa98f-136">string</span><span class="sxs-lookup"><span data-stu-id="fa98f-136">string</span></span>  |  <span data-ttu-id="fa98f-137">Да</span><span class="sxs-lookup"><span data-stu-id="fa98f-137">Yes</span></span>  |  <span data-ttu-id="fa98f-138">Имя функции, которое отображается пользователям в Excel.</span><span class="sxs-lookup"><span data-stu-id="fa98f-138">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="fa98f-139">В Excel имя этой функции будет присоединено в качестве префикса пространством имен настраиваемой функции, указанным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="fa98f-139">In Excel, this function name will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
|  `options`  |  <span data-ttu-id="fa98f-140">объект</span><span class="sxs-lookup"><span data-stu-id="fa98f-140">object</span></span>  |  <span data-ttu-id="fa98f-141">Нет</span><span class="sxs-lookup"><span data-stu-id="fa98f-141">No</span></span>  |  <span data-ttu-id="fa98f-142">Позволяет настроить некоторые аспекты того, как и когда Excel выполняет функцию.</span><span class="sxs-lookup"><span data-stu-id="fa98f-142">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="fa98f-143">Дополнительные сведения см. в разделе [options](#options).</span><span class="sxs-lookup"><span data-stu-id="fa98f-143">See [options](#options) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="fa98f-144">array</span><span class="sxs-lookup"><span data-stu-id="fa98f-144">array</span></span>  |  <span data-ttu-id="fa98f-145">Да</span><span class="sxs-lookup"><span data-stu-id="fa98f-145">Yes</span></span>  |  <span data-ttu-id="fa98f-146">Массив, который определяет входные параметры для функции.</span><span class="sxs-lookup"><span data-stu-id="fa98f-146">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="fa98f-147">Дополнительные сведения см. в разделе [parameters](#parameters).</span><span class="sxs-lookup"><span data-stu-id="fa98f-147">See [parameters](#parameters)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="fa98f-148">object</span><span class="sxs-lookup"><span data-stu-id="fa98f-148">object</span></span>  |  <span data-ttu-id="fa98f-149">Да</span><span class="sxs-lookup"><span data-stu-id="fa98f-149">Yes</span></span>  |  <span data-ttu-id="fa98f-150">Объект, который определяет тип информации, возвращаемый функцией.</span><span class="sxs-lookup"><span data-stu-id="fa98f-150">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="fa98f-151">Дополнительные сведения см. в разделе [result](#result).</span><span class="sxs-lookup"><span data-stu-id="fa98f-151">See [result](#result) for details.</span></span> |

## <a name="options"></a><span data-ttu-id="fa98f-152">options</span><span class="sxs-lookup"><span data-stu-id="fa98f-152">options</span></span>

<span data-ttu-id="fa98f-153">Объект `options` позволяет настроить некоторые аспекты того, как и когда Excel выполняет функцию.</span><span class="sxs-lookup"><span data-stu-id="fa98f-153">The `options` object enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="fa98f-154">В таблице ниже приведены свойства объекта `options`.</span><span class="sxs-lookup"><span data-stu-id="fa98f-154">The following table lists the properties of the `options` object.</span></span>

|  <span data-ttu-id="fa98f-155">Свойство</span><span class="sxs-lookup"><span data-stu-id="fa98f-155">Property</span></span>  |  <span data-ttu-id="fa98f-156">Тип данных</span><span class="sxs-lookup"><span data-stu-id="fa98f-156">Data type</span></span>  |  <span data-ttu-id="fa98f-157">Обязательный</span><span class="sxs-lookup"><span data-stu-id="fa98f-157">Required</span></span>  |  <span data-ttu-id="fa98f-158">Описание</span><span class="sxs-lookup"><span data-stu-id="fa98f-158">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="fa98f-159">boolean</span><span class="sxs-lookup"><span data-stu-id="fa98f-159">boolean</span></span>  |  <span data-ttu-id="fa98f-160">Нет</span><span class="sxs-lookup"><span data-stu-id="fa98f-160">No</span></span><br/><br/><span data-ttu-id="fa98f-161">Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="fa98f-161">Default value is `false`.</span></span>  |  <span data-ttu-id="fa98f-162">Если это свойство имеет значение `true`, Excel будет вызывать обработчик `onCanceled` каждый раз, когда пользователь будет предпринимать действия, которые приводят к отмене функции (например, вручную вызывает пересчет или редактирует ячейку, на которую ссылается функция).</span><span class="sxs-lookup"><span data-stu-id="fa98f-162">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="fa98f-163">Если вы используете это свойство, Excel вызовет функцию JavaScript с дополнительным параметром `caller`.</span><span class="sxs-lookup"><span data-stu-id="fa98f-163">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="fa98f-164">(***Не*** регистрируйте этот параметр в свойстве `parameters`.)</span><span class="sxs-lookup"><span data-stu-id="fa98f-164">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="fa98f-165">В тексте функции обработчик необходимо назначить элементу `caller.onCanceled`.</span><span class="sxs-lookup"><span data-stu-id="fa98f-165">In the body of the function, a handler must be assigned to the `caller.onCanceled` member.</span></span> <span data-ttu-id="fa98f-166">Дополнительные сведения см. в разделе [Отмена функции](custom-functions-web-reqs.md#canceling-a-function).</span><span class="sxs-lookup"><span data-stu-id="fa98f-166">For more information, see [Canceling a function](custom-functions-web-reqs.md#canceling-a-function).</span></span> |
|  `stream`  |  <span data-ttu-id="fa98f-167">boolean</span><span class="sxs-lookup"><span data-stu-id="fa98f-167">boolean</span></span>  |  <span data-ttu-id="fa98f-168">Нет</span><span class="sxs-lookup"><span data-stu-id="fa98f-168">No</span></span><br/><br/><span data-ttu-id="fa98f-169">Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="fa98f-169">Default value is `false`.</span></span>  |  <span data-ttu-id="fa98f-170">Если это свойство имеет значение `true`, функция может выводить значение в ячейку несколько раз, даже если вызвана всего единожды.</span><span class="sxs-lookup"><span data-stu-id="fa98f-170">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="fa98f-171">Этот параметр полезен для быстро изменяющихся источников данных, таких как цена акций.</span><span class="sxs-lookup"><span data-stu-id="fa98f-171">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="fa98f-172">Если вы используете это свойство, Excel вызовет функцию JavaScript с дополнительным параметром `caller`.</span><span class="sxs-lookup"><span data-stu-id="fa98f-172">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="fa98f-173">(***Не*** регистрируйте этот параметр в свойстве `parameters`.)</span><span class="sxs-lookup"><span data-stu-id="fa98f-173">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="fa98f-174">Функция не должна содержать оператор `return`.</span><span class="sxs-lookup"><span data-stu-id="fa98f-174">The function should have no `return` statement.</span></span> <span data-ttu-id="fa98f-175">Вместо этого результирующее значение передается как аргумент метода обратного вызова `caller.setResult`.</span><span class="sxs-lookup"><span data-stu-id="fa98f-175">Instead, the result value is passed as the argument of the `caller.setResult` callback method.</span></span> <span data-ttu-id="fa98f-176">Дополнительные сведения см. в разделе [Потоковые функции](custom-functions-web-reqs.md#streaming-functions).</span><span class="sxs-lookup"><span data-stu-id="fa98f-176">For more information, see [Streaming functions](custom-functions-web-reqs.md#streaming-functions).</span></span> |
|  `volatile`  | <span data-ttu-id="fa98f-177">boolean</span><span class="sxs-lookup"><span data-stu-id="fa98f-177">boolean</span></span> | <span data-ttu-id="fa98f-178">Нет</span><span class="sxs-lookup"><span data-stu-id="fa98f-178">No</span></span> <br/><br/><span data-ttu-id="fa98f-179">Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="fa98f-179">Default value is `false`.</span></span> | <br /><br /> <span data-ttu-id="fa98f-180">Если присвоено значение `true`, функция пересчитывается при каждом выполнении пересчета в Excel, а не только при изменении зависимых значений формулы.</span><span class="sxs-lookup"><span data-stu-id="fa98f-180">If `true`, the function will recalculate each time Excel recalculates, instead of only when the formula's dependent values have changed.</span></span> <span data-ttu-id="fa98f-181">Функция не может быть одновременно потоковой и переменной.</span><span class="sxs-lookup"><span data-stu-id="fa98f-181">A function cannot be both streaming and volatile.</span></span> <span data-ttu-id="fa98f-182">Если обоим свойствам `stream` и `volatile` присвоено значение `true`, параметр переменности будет игнорироваться.</span><span class="sxs-lookup"><span data-stu-id="fa98f-182">If the `stream` and `volatile` properties are both set to `true`, the volatile option will be ignored.</span></span> |

## <a name="parameters"></a><span data-ttu-id="fa98f-183">parameters</span><span class="sxs-lookup"><span data-stu-id="fa98f-183">parameters</span></span>

<span data-ttu-id="fa98f-184">Свойство `parameters` представляет собой массив объектов параметров.</span><span class="sxs-lookup"><span data-stu-id="fa98f-184">The `parameters` property is an array of parameter objects.</span></span> <span data-ttu-id="fa98f-185">В таблице ниже приведены свойства каждого объекта.</span><span class="sxs-lookup"><span data-stu-id="fa98f-185">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="fa98f-186">Свойство</span><span class="sxs-lookup"><span data-stu-id="fa98f-186">Property</span></span>  |  <span data-ttu-id="fa98f-187">Тип данных</span><span class="sxs-lookup"><span data-stu-id="fa98f-187">Data type</span></span>  |  <span data-ttu-id="fa98f-188">Обязательный</span><span class="sxs-lookup"><span data-stu-id="fa98f-188">Required</span></span>  |  <span data-ttu-id="fa98f-189">Описание</span><span class="sxs-lookup"><span data-stu-id="fa98f-189">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="fa98f-190">string</span><span class="sxs-lookup"><span data-stu-id="fa98f-190">string</span></span>  |  <span data-ttu-id="fa98f-191">Нет</span><span class="sxs-lookup"><span data-stu-id="fa98f-191">No</span></span> |  <span data-ttu-id="fa98f-192">Описание параметра.</span><span class="sxs-lookup"><span data-stu-id="fa98f-192">A description of the parameter.</span></span> <span data-ttu-id="fa98f-193">Отображается в IntelliSense Excel.</span><span class="sxs-lookup"><span data-stu-id="fa98f-193">This is displayed in Excel's intelliSense.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="fa98f-194">string</span><span class="sxs-lookup"><span data-stu-id="fa98f-194">string</span></span>  |  <span data-ttu-id="fa98f-195">Нет</span><span class="sxs-lookup"><span data-stu-id="fa98f-195">No</span></span>  |  <span data-ttu-id="fa98f-196">Должно быть **скалярным** (значение, отличное от массива) или **матричным** (двухмерный массив).</span><span class="sxs-lookup"><span data-stu-id="fa98f-196">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="fa98f-197">string</span><span class="sxs-lookup"><span data-stu-id="fa98f-197">string</span></span>  |  <span data-ttu-id="fa98f-198">Да</span><span class="sxs-lookup"><span data-stu-id="fa98f-198">Yes</span></span>  |  <span data-ttu-id="fa98f-199">Имя параметра.</span><span class="sxs-lookup"><span data-stu-id="fa98f-199">The name of the parameter.</span></span> <span data-ttu-id="fa98f-200">Оно отображается в IntelliSense Excel.</span><span class="sxs-lookup"><span data-stu-id="fa98f-200">This name is displayed in Excel's intelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="fa98f-201">string</span><span class="sxs-lookup"><span data-stu-id="fa98f-201">string</span></span>  |  <span data-ttu-id="fa98f-202">Нет</span><span class="sxs-lookup"><span data-stu-id="fa98f-202">No</span></span>  |  <span data-ttu-id="fa98f-203">Тип данных параметра.</span><span class="sxs-lookup"><span data-stu-id="fa98f-203">The data type of the parameter.</span></span> <span data-ttu-id="fa98f-204">Может иметь значение **boolean**, **number**, **string** или **any**, что позволяет использовать любой из трех предыдущих типов.</span><span class="sxs-lookup"><span data-stu-id="fa98f-204">Can be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> <span data-ttu-id="fa98f-205">Если это свойство не задано, по умолчанию устанавливается тип данных **any**.</span><span class="sxs-lookup"><span data-stu-id="fa98f-205">If this property is not specified, the data type defaults to **any**.</span></span> |
|  `optional`  | <span data-ttu-id="fa98f-206">boolean</span><span class="sxs-lookup"><span data-stu-id="fa98f-206">boolean</span></span> | <span data-ttu-id="fa98f-207">Нет</span><span class="sxs-lookup"><span data-stu-id="fa98f-207">No</span></span> | <span data-ttu-id="fa98f-208">Если присвоено значение `true`, параметр не обязателен.</span><span class="sxs-lookup"><span data-stu-id="fa98f-208">If `true`, the parameter is optional.</span></span> |

>[!NOTE]
> <span data-ttu-id="fa98f-209">Если свойство `type` необязательного параметра не указано или равно `any`, вы можете заметить проблемы, например ошибки линтинга в интегрированной среде разработки (IDE) и отсутствие необязательных параметров при вводе функции в ячейке Excel.</span><span class="sxs-lookup"><span data-stu-id="fa98f-209">If the `type` property of an optional parameter is either not specified or set to `any`, you may notice issues such as linting errors in your IDE and optional parameters not being displayed when the function is being entered into a cell in Excel.</span></span> <span data-ttu-id="fa98f-210">Это планируется изменить в декабре 2018 г.</span><span class="sxs-lookup"><span data-stu-id="fa98f-210">This is projected to change in December of 2018.</span></span>

## <a name="result"></a><span data-ttu-id="fa98f-211">result</span><span class="sxs-lookup"><span data-stu-id="fa98f-211">result</span></span>

<span data-ttu-id="fa98f-212">Объект `result` определяет тип информации, возвращаемый функцией.</span><span class="sxs-lookup"><span data-stu-id="fa98f-212">The `result` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="fa98f-213">В таблице ниже приведены свойства объекта `result`.</span><span class="sxs-lookup"><span data-stu-id="fa98f-213">The following table lists the properties of the `result` object.</span></span>

|  <span data-ttu-id="fa98f-214">Свойство</span><span class="sxs-lookup"><span data-stu-id="fa98f-214">Property</span></span>  |  <span data-ttu-id="fa98f-215">Тип данных</span><span class="sxs-lookup"><span data-stu-id="fa98f-215">Data type</span></span>  |  <span data-ttu-id="fa98f-216">Обязательный</span><span class="sxs-lookup"><span data-stu-id="fa98f-216">Required</span></span>  |  <span data-ttu-id="fa98f-217">Описание</span><span class="sxs-lookup"><span data-stu-id="fa98f-217">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="fa98f-218">string</span><span class="sxs-lookup"><span data-stu-id="fa98f-218">string</span></span>  |  <span data-ttu-id="fa98f-219">Нет</span><span class="sxs-lookup"><span data-stu-id="fa98f-219">No</span></span>  |  <span data-ttu-id="fa98f-220">Должно быть **скалярным** (значение, отличное от массива) или **матричным** (двухмерный массив).</span><span class="sxs-lookup"><span data-stu-id="fa98f-220">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |
|  `type`  |  <span data-ttu-id="fa98f-221">string</span><span class="sxs-lookup"><span data-stu-id="fa98f-221">string</span></span>  |  <span data-ttu-id="fa98f-222">Да</span><span class="sxs-lookup"><span data-stu-id="fa98f-222">Yes</span></span>  |  <span data-ttu-id="fa98f-223">Тип данных параметра.</span><span class="sxs-lookup"><span data-stu-id="fa98f-223">The data type of the parameter.</span></span> <span data-ttu-id="fa98f-224">Должен иметь значение **boolean**, **number**, **string** или **any**, что позволяет использовать любой из трех предыдущих типов.</span><span class="sxs-lookup"><span data-stu-id="fa98f-224">Must be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> |

## <a name="see-also"></a><span data-ttu-id="fa98f-225">См. также</span><span class="sxs-lookup"><span data-stu-id="fa98f-225">See also</span></span>

* [<span data-ttu-id="fa98f-226">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="fa98f-226">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="fa98f-227">Среда выполнения для пользовательских функций Excel</span><span class="sxs-lookup"><span data-stu-id="fa98f-227">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="fa98f-228">Рекомендации по пользовательским функциям</span><span class="sxs-lookup"><span data-stu-id="fa98f-228">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="fa98f-229">Журнал изменений пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="fa98f-229">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="fa98f-230">Руководство по пользовательским функциям в Excel</span><span class="sxs-lookup"><span data-stu-id="fa98f-230">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
