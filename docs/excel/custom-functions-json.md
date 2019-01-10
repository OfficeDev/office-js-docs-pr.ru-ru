---
ms.date: 01/08/2019
description: Определение метаданных для настраиваемых функций в Excel.
title: Метаданные для настраиваемых функций в Excel (предварительная версия)
ms.openlocfilehash: 79fb5ac0ebcbe91b154c48cb10aad65451b6e5ee
ms.sourcegitcommit: 9afcb1bb295ec0c8940ed3a8364dbac08ef6b382
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/08/2019
ms.locfileid: "27770632"
---
# <a name="custom-functions-metadata-preview"></a><span data-ttu-id="e9c79-103">Метаданные для настраиваемых функций (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="e9c79-103">Custom functions metadata (preview)</span></span>

<span data-ttu-id="e9c79-104">При определении [настраиваемых функций](custom-functions-overview.md) в надстройке Excel в проект надстройки необходимо включить JSON-файл метаданных, содержащий информацию, необходимую Excel для регистрации настраиваемых функций и предоставления пользователям доступа к ним.</span><span class="sxs-lookup"><span data-stu-id="e9c79-104">When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project must include a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users.</span></span> <span data-ttu-id="e9c79-105">В этой статье описан формат JSON-файла метаданных.</span><span class="sxs-lookup"><span data-stu-id="e9c79-105">This article describes the format of the JSON metadata file.</span></span>

<span data-ttu-id="e9c79-106">Сведения о других файлах, которые необходимо добавить в проект надстройки для включения настраиваемых функций, см. в статье [Создание пользовательских функций в Excel](custom-functions-overview.md).</span><span class="sxs-lookup"><span data-stu-id="e9c79-106">For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="example-metadata"></a><span data-ttu-id="e9c79-107">Пример метаданных</span><span class="sxs-lookup"><span data-stu-id="e9c79-107">Example metadata</span></span>

<span data-ttu-id="e9c79-108">В примере кода ниже показано содержимое JSON-файла метаданных для надстройки, определяющей настраиваемые функции.</span><span class="sxs-lookup"><span data-stu-id="e9c79-108">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="e9c79-109">В следующих за этим примером разделах приводятся подробные сведения об отдельных свойствах, представленных в этом примере JSON.</span><span class="sxs-lookup"><span data-stu-id="e9c79-109">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
> <span data-ttu-id="e9c79-110">Пример готового JSON-файла приводится в репозитории GitHub [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).</span><span class="sxs-lookup"><span data-stu-id="e9c79-110">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json) GitHub repository.</span></span>

## <a name="functions"></a><span data-ttu-id="e9c79-111">functions</span><span class="sxs-lookup"><span data-stu-id="e9c79-111">functions</span></span> 

<span data-ttu-id="e9c79-112">Свойство `functions` представляет собой массив объектов настраиваемых функций.</span><span class="sxs-lookup"><span data-stu-id="e9c79-112">The `functions` property is an array of custom function objects.</span></span> <span data-ttu-id="e9c79-113">В таблице ниже приведены свойства каждого объекта.</span><span class="sxs-lookup"><span data-stu-id="e9c79-113">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="e9c79-114">Свойство</span><span class="sxs-lookup"><span data-stu-id="e9c79-114">Property</span></span>  |  <span data-ttu-id="e9c79-115">Тип данных</span><span class="sxs-lookup"><span data-stu-id="e9c79-115">Data type</span></span>  |  <span data-ttu-id="e9c79-116">Обязательное</span><span class="sxs-lookup"><span data-stu-id="e9c79-116">Required</span></span>  |  <span data-ttu-id="e9c79-117">Описание</span><span class="sxs-lookup"><span data-stu-id="e9c79-117">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="e9c79-118">string</span><span class="sxs-lookup"><span data-stu-id="e9c79-118">string</span></span>  |  <span data-ttu-id="e9c79-119">Нет</span><span class="sxs-lookup"><span data-stu-id="e9c79-119">No</span></span>  |  <span data-ttu-id="e9c79-120">Описание функции, которое отображается пользователям в Excel</span><span class="sxs-lookup"><span data-stu-id="e9c79-120">The description of the function that end users see in Excel.</span></span> <span data-ttu-id="e9c79-121">(например, **преобразует значение по шкале Цельсия в температуру по шкале Фаренгейта**).</span><span class="sxs-lookup"><span data-stu-id="e9c79-121">For example, **Converts a Celsius value to Fahrenheit**.</span></span> |
|  `helpUrl`  |  <span data-ttu-id="e9c79-122">string</span><span class="sxs-lookup"><span data-stu-id="e9c79-122">string</span></span>  |   <span data-ttu-id="e9c79-123">Нет</span><span class="sxs-lookup"><span data-stu-id="e9c79-123">No</span></span>  |  <span data-ttu-id="e9c79-124">URL-адрес, по которому можно получить сведения о функции</span><span class="sxs-lookup"><span data-stu-id="e9c79-124">URL that provides information about the function.</span></span> <span data-ttu-id="e9c79-125">(отображается в области задач). Пример: **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span><span class="sxs-lookup"><span data-stu-id="e9c79-125">(It is displayed in a task pane.) For example, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span></span> |
| `id`     | <span data-ttu-id="e9c79-126">string</span><span class="sxs-lookup"><span data-stu-id="e9c79-126">string</span></span> | <span data-ttu-id="e9c79-127">Да</span><span class="sxs-lookup"><span data-stu-id="e9c79-127">Yes</span></span> | <span data-ttu-id="e9c79-128">Уникальный идентификатор для функции.</span><span class="sxs-lookup"><span data-stu-id="e9c79-128">A unique ID for the function.</span></span> <span data-ttu-id="e9c79-129">Этот идентификатор может содержать только буквы, цифры и точки, а его изменение после настройки не допускается.</span><span class="sxs-lookup"><span data-stu-id="e9c79-129">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span> |
|  `name`  |  <span data-ttu-id="e9c79-130">string</span><span class="sxs-lookup"><span data-stu-id="e9c79-130">string</span></span>  |  <span data-ttu-id="e9c79-131">Да</span><span class="sxs-lookup"><span data-stu-id="e9c79-131">Yes</span></span>  |  <span data-ttu-id="e9c79-132">Имя функции, которое отображается пользователям в Excel.</span><span class="sxs-lookup"><span data-stu-id="e9c79-132">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="e9c79-133">В Excel имя этой функции будет присоединено в качестве префикса пространством имен настраиваемой функции, указанным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="e9c79-133">In Excel, this function name will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
|  `options`  |  <span data-ttu-id="e9c79-134">object</span><span class="sxs-lookup"><span data-stu-id="e9c79-134">object</span></span>  |  <span data-ttu-id="e9c79-135">Нет</span><span class="sxs-lookup"><span data-stu-id="e9c79-135">No</span></span>  |  <span data-ttu-id="e9c79-136">Позволяет настроить некоторые аспекты того, как и когда Excel выполняет функцию.</span><span class="sxs-lookup"><span data-stu-id="e9c79-136">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="e9c79-137">Дополнительные сведения см. в разделе [options](#options).</span><span class="sxs-lookup"><span data-stu-id="e9c79-137">See [options](#options) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="e9c79-138">array</span><span class="sxs-lookup"><span data-stu-id="e9c79-138">array</span></span>  |  <span data-ttu-id="e9c79-139">Да</span><span class="sxs-lookup"><span data-stu-id="e9c79-139">Yes</span></span>  |  <span data-ttu-id="e9c79-140">Массив, который определяет входные параметры для функции.</span><span class="sxs-lookup"><span data-stu-id="e9c79-140">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="e9c79-141">Дополнительные сведения см. в разделе [parameters](#parameters).</span><span class="sxs-lookup"><span data-stu-id="e9c79-141">See [parameters](#parameters)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="e9c79-142">object</span><span class="sxs-lookup"><span data-stu-id="e9c79-142">object</span></span>  |  <span data-ttu-id="e9c79-143">Да</span><span class="sxs-lookup"><span data-stu-id="e9c79-143">Yes</span></span>  |  <span data-ttu-id="e9c79-144">Объект, который определяет тип информации, возвращаемый функцией.</span><span class="sxs-lookup"><span data-stu-id="e9c79-144">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="e9c79-145">Дополнительные сведения см. в разделе [result](#result).</span><span class="sxs-lookup"><span data-stu-id="e9c79-145">See [result](#result) for details.</span></span> |

## <a name="options"></a><span data-ttu-id="e9c79-146">options</span><span class="sxs-lookup"><span data-stu-id="e9c79-146">options</span></span>

<span data-ttu-id="e9c79-147">Объект `options` позволяет настроить некоторые аспекты того, как и когда Excel выполняет функцию.</span><span class="sxs-lookup"><span data-stu-id="e9c79-147">The `options` object enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="e9c79-148">В таблице ниже приведены свойства объекта `options`.</span><span class="sxs-lookup"><span data-stu-id="e9c79-148">The following table lists the properties of the `options` object.</span></span>

|  <span data-ttu-id="e9c79-149">Свойство</span><span class="sxs-lookup"><span data-stu-id="e9c79-149">Property</span></span>  |  <span data-ttu-id="e9c79-150">Тип данных</span><span class="sxs-lookup"><span data-stu-id="e9c79-150">Data type</span></span>  |  <span data-ttu-id="e9c79-151">Обязательное</span><span class="sxs-lookup"><span data-stu-id="e9c79-151">Required</span></span>  |  <span data-ttu-id="e9c79-152">Описание</span><span class="sxs-lookup"><span data-stu-id="e9c79-152">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="e9c79-153">boolean</span><span class="sxs-lookup"><span data-stu-id="e9c79-153">boolean</span></span>  |  <span data-ttu-id="e9c79-154">Нет</span><span class="sxs-lookup"><span data-stu-id="e9c79-154">No</span></span><br/><br/><span data-ttu-id="e9c79-155">Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="e9c79-155">Default value is `false`.</span></span>  |  <span data-ttu-id="e9c79-156">Если это свойство имеет значение `true`, Excel будет вызывать обработчик `onCanceled` каждый раз, когда пользователь будет предпринимать действия, которые приводят к отмене функции (например, вручную вызывает пересчет или редактирует ячейку, на которую ссылается функция).</span><span class="sxs-lookup"><span data-stu-id="e9c79-156">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="e9c79-157">Если вы используете это свойство, Excel вызовет функцию JavaScript с дополнительным параметром `caller`.</span><span class="sxs-lookup"><span data-stu-id="e9c79-157">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="e9c79-158">(***Не*** регистрируйте этот параметр в свойстве `parameters`.)</span><span class="sxs-lookup"><span data-stu-id="e9c79-158">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="e9c79-159">В тексте функции обработчик необходимо назначить элементу `caller.onCanceled`.</span><span class="sxs-lookup"><span data-stu-id="e9c79-159">In the body of the function, a handler must be assigned to the `caller.onCanceled` member.</span></span> <span data-ttu-id="e9c79-160">Дополнительные сведения см. в разделе [Отмена функции](custom-functions-overview.md#canceling-a-function).</span><span class="sxs-lookup"><span data-stu-id="e9c79-160">For more information, see [Canceling a function](custom-functions-overview.md#canceling-a-function).</span></span> |
|  `stream`  |  <span data-ttu-id="e9c79-161">boolean</span><span class="sxs-lookup"><span data-stu-id="e9c79-161">boolean</span></span>  |  <span data-ttu-id="e9c79-162">Нет</span><span class="sxs-lookup"><span data-stu-id="e9c79-162">No</span></span><br/><br/><span data-ttu-id="e9c79-163">Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="e9c79-163">Default value is `false`.</span></span>  |  <span data-ttu-id="e9c79-164">Если это свойство имеет значение `true`, функция может выводить значение в ячейку несколько раз, даже если вызвана всего единожды.</span><span class="sxs-lookup"><span data-stu-id="e9c79-164">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="e9c79-165">Этот параметр полезен для быстро изменяющихся источников данных, таких как цена акций.</span><span class="sxs-lookup"><span data-stu-id="e9c79-165">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="e9c79-166">Если вы используете это свойство, Excel вызовет функцию JavaScript с дополнительным параметром `caller`.</span><span class="sxs-lookup"><span data-stu-id="e9c79-166">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="e9c79-167">(***Не*** регистрируйте этот параметр в свойстве `parameters`.)</span><span class="sxs-lookup"><span data-stu-id="e9c79-167">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="e9c79-168">Функция не должна содержать оператор `return`.</span><span class="sxs-lookup"><span data-stu-id="e9c79-168">The function should have no `return` statement.</span></span> <span data-ttu-id="e9c79-169">Вместо этого результирующее значение передается как аргумент метода обратного вызова `caller.setResult`.</span><span class="sxs-lookup"><span data-stu-id="e9c79-169">Instead, the result value is passed as the argument of the `caller.setResult` callback method.</span></span> <span data-ttu-id="e9c79-170">Дополнительные сведения см. в разделе [Потоковые функции](custom-functions-overview.md#streaming-functions).</span><span class="sxs-lookup"><span data-stu-id="e9c79-170">For more information, see [Streaming functions](custom-functions-overview.md#streaming-functions).</span></span> |
|  `volatile`  | <span data-ttu-id="e9c79-171">boolean</span><span class="sxs-lookup"><span data-stu-id="e9c79-171">boolean</span></span> | <span data-ttu-id="e9c79-172">Нет</span><span class="sxs-lookup"><span data-stu-id="e9c79-172">No</span></span> <br/><br/><span data-ttu-id="e9c79-173">Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="e9c79-173">Default value is `false`.</span></span> | <br /><br /> <span data-ttu-id="e9c79-174">Если присвоено значение `true`, функция пересчитывается при каждом выполнении пересчета в Excel, а не только при изменении зависимых значений формулы.</span><span class="sxs-lookup"><span data-stu-id="e9c79-174">If `true`, the function will recalculate each time Excel recalculates, instead of only when the formula's dependent values have changed.</span></span> <span data-ttu-id="e9c79-175">Функция не может быть одновременно потоковой и переменной.</span><span class="sxs-lookup"><span data-stu-id="e9c79-175">A function cannot be both streaming and volatile.</span></span> <span data-ttu-id="e9c79-176">Если обоим свойствам `stream` и `volatile` присвоено значение `true`, параметр переменности будет игнорироваться.</span><span class="sxs-lookup"><span data-stu-id="e9c79-176">If the `stream` and `volatile` properties are both set to `true`, the volatile option will be ignored.</span></span> |

## <a name="parameters"></a><span data-ttu-id="e9c79-177">parameters</span><span class="sxs-lookup"><span data-stu-id="e9c79-177">parameters</span></span>

<span data-ttu-id="e9c79-178">Свойство `parameters` представляет собой массив объектов параметров.</span><span class="sxs-lookup"><span data-stu-id="e9c79-178">The `parameters` property is an array of parameter objects.</span></span> <span data-ttu-id="e9c79-179">В таблице ниже приведены свойства каждого объекта.</span><span class="sxs-lookup"><span data-stu-id="e9c79-179">The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="e9c79-180">Свойство</span><span class="sxs-lookup"><span data-stu-id="e9c79-180">Property</span></span>  |  <span data-ttu-id="e9c79-181">Тип данных</span><span class="sxs-lookup"><span data-stu-id="e9c79-181">Data type</span></span>  |  <span data-ttu-id="e9c79-182">Обязательное</span><span class="sxs-lookup"><span data-stu-id="e9c79-182">Required</span></span>  |  <span data-ttu-id="e9c79-183">Описание</span><span class="sxs-lookup"><span data-stu-id="e9c79-183">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="e9c79-184">string</span><span class="sxs-lookup"><span data-stu-id="e9c79-184">string</span></span>  |  <span data-ttu-id="e9c79-185">Нет</span><span class="sxs-lookup"><span data-stu-id="e9c79-185">No</span></span> |  <span data-ttu-id="e9c79-186">Описание параметра.</span><span class="sxs-lookup"><span data-stu-id="e9c79-186">A description of the parameter.</span></span> <span data-ttu-id="e9c79-187">Отображается в IntelliSense Excel.</span><span class="sxs-lookup"><span data-stu-id="e9c79-187">This is displayed in Excel's intelliSense.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="e9c79-188">string</span><span class="sxs-lookup"><span data-stu-id="e9c79-188">string</span></span>  |  <span data-ttu-id="e9c79-189">Нет</span><span class="sxs-lookup"><span data-stu-id="e9c79-189">No</span></span>  |  <span data-ttu-id="e9c79-190">Должно быть **скалярным** (значение, отличное от массива) или **матричным** (двухмерный массив).</span><span class="sxs-lookup"><span data-stu-id="e9c79-190">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="e9c79-191">string</span><span class="sxs-lookup"><span data-stu-id="e9c79-191">string</span></span>  |  <span data-ttu-id="e9c79-192">Да</span><span class="sxs-lookup"><span data-stu-id="e9c79-192">Yes</span></span>  |  <span data-ttu-id="e9c79-193">Имя параметра.</span><span class="sxs-lookup"><span data-stu-id="e9c79-193">The name of the parameter.</span></span> <span data-ttu-id="e9c79-194">Оно отображается в IntelliSense Excel.</span><span class="sxs-lookup"><span data-stu-id="e9c79-194">This name is displayed in Excel's intelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="e9c79-195">string</span><span class="sxs-lookup"><span data-stu-id="e9c79-195">string</span></span>  |  <span data-ttu-id="e9c79-196">Нет</span><span class="sxs-lookup"><span data-stu-id="e9c79-196">No</span></span>  |  <span data-ttu-id="e9c79-197">Тип данных параметра.</span><span class="sxs-lookup"><span data-stu-id="e9c79-197">The data type of the parameter.</span></span> <span data-ttu-id="e9c79-198">Может иметь значение **boolean**, **number**, **string** или **any**, что позволяет использовать любой из трех предыдущих типов.</span><span class="sxs-lookup"><span data-stu-id="e9c79-198">Can be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> <span data-ttu-id="e9c79-199">Если это свойство не задано, по умолчанию устанавливается тип данных **any**.</span><span class="sxs-lookup"><span data-stu-id="e9c79-199">If this property is not specified, the data type defaults to **any**.</span></span> |
|  `optional`  | <span data-ttu-id="e9c79-200">boolean</span><span class="sxs-lookup"><span data-stu-id="e9c79-200">boolean</span></span> | <span data-ttu-id="e9c79-201">Нет</span><span class="sxs-lookup"><span data-stu-id="e9c79-201">No</span></span> | <span data-ttu-id="e9c79-202">Если присвоено значение `true`, параметр не обязателен.</span><span class="sxs-lookup"><span data-stu-id="e9c79-202">If `true`, the parameter is optional.</span></span> |

>[!NOTE]
> <span data-ttu-id="e9c79-203">Если свойство `type` необязательного параметра не указано или равно `any`, вы можете заметить проблемы, например ошибки линтинга в интегрированной среде разработки (IDE) и отсутствие необязательных параметров при вводе функции в ячейке Excel.</span><span class="sxs-lookup"><span data-stu-id="e9c79-203">If the `type` property of an optional parameter is either not specified or set to `any`, you may notice issues such as linting errors in your IDE and optional parameters not being displayed when the function is being entered into a cell in Excel.</span></span> <span data-ttu-id="e9c79-204">Это планируется изменить в декабре 2018 г.</span><span class="sxs-lookup"><span data-stu-id="e9c79-204">This is projected to change in December of 2018.</span></span>

## <a name="result"></a><span data-ttu-id="e9c79-205">result</span><span class="sxs-lookup"><span data-stu-id="e9c79-205">result</span></span>

<span data-ttu-id="e9c79-206">Объект `result` определяет тип информации, возвращаемый функцией.</span><span class="sxs-lookup"><span data-stu-id="e9c79-206">The `result` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="e9c79-207">В таблице ниже приведены свойства объекта `result`.</span><span class="sxs-lookup"><span data-stu-id="e9c79-207">The following table lists the properties of the `result` object.</span></span>

|  <span data-ttu-id="e9c79-208">Свойство</span><span class="sxs-lookup"><span data-stu-id="e9c79-208">Property</span></span>  |  <span data-ttu-id="e9c79-209">Тип данных</span><span class="sxs-lookup"><span data-stu-id="e9c79-209">Data type</span></span>  |  <span data-ttu-id="e9c79-210">Обязательное</span><span class="sxs-lookup"><span data-stu-id="e9c79-210">Required</span></span>  |  <span data-ttu-id="e9c79-211">Описание</span><span class="sxs-lookup"><span data-stu-id="e9c79-211">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="e9c79-212">string</span><span class="sxs-lookup"><span data-stu-id="e9c79-212">string</span></span>  |  <span data-ttu-id="e9c79-213">Нет</span><span class="sxs-lookup"><span data-stu-id="e9c79-213">No</span></span>  |  <span data-ttu-id="e9c79-214">Должно быть **скалярным** (значение, отличное от массива) или **матричным** (двухмерный массив).</span><span class="sxs-lookup"><span data-stu-id="e9c79-214">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |
|  `type`  |  <span data-ttu-id="e9c79-215">string</span><span class="sxs-lookup"><span data-stu-id="e9c79-215">string</span></span>  |  <span data-ttu-id="e9c79-216">Да</span><span class="sxs-lookup"><span data-stu-id="e9c79-216">Yes</span></span>  |  <span data-ttu-id="e9c79-217">Тип данных параметра.</span><span class="sxs-lookup"><span data-stu-id="e9c79-217">The data type of the parameter.</span></span> <span data-ttu-id="e9c79-218">Должен иметь значение **boolean**, **number**, **string** или **any**, что позволяет использовать любой из трех предыдущих типов.</span><span class="sxs-lookup"><span data-stu-id="e9c79-218">Must be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> |

## <a name="see-also"></a><span data-ttu-id="e9c79-219">См. также</span><span class="sxs-lookup"><span data-stu-id="e9c79-219">See also</span></span>

* [<span data-ttu-id="e9c79-220">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="e9c79-220">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="e9c79-221">Среда выполнения для пользовательских функций Excel</span><span class="sxs-lookup"><span data-stu-id="e9c79-221">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="e9c79-222">Рекомендации по настраиваемым функциям</span><span class="sxs-lookup"><span data-stu-id="e9c79-222">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="e9c79-223">Журнал изменений пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="e9c79-223">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="e9c79-224">Руководство по настраиваемым функциям в Excel</span><span class="sxs-lookup"><span data-stu-id="e9c79-224">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
