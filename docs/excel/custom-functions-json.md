---
ms.date: 10/17/2018
description: Определение метаданных для настраиваемых функций в Excel.
title: Метаданные для настраиваемых функций в Excel
ms.openlocfilehash: 0c77474188a2deefd23a73bb64e87569bb1fa52a
ms.sourcegitcommit: 2ac7d64bb2db75ace516a604866850fce5cb2174
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/14/2018
ms.locfileid: "26298546"
---
# <a name="custom-functions-metadata-preview"></a><span data-ttu-id="ef08b-103">Метаданные для настраиваемых функций (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="ef08b-103">Custom functions metadata (preview)</span></span>

<span data-ttu-id="ef08b-104">При определении [настраиваемых функций](custom-functions-overview.md) в надстройке Excel в проект надстройки необходимо включить JSON-файл метаданных, содержащий информацию, необходимую Excel для регистрации настраиваемых функций и предоставления пользователям доступа к ним.</span><span class="sxs-lookup"><span data-stu-id="ef08b-104">When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project must include a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users.</span></span> <span data-ttu-id="ef08b-105">В этой статье описан формат JSON-файла метаданных.</span><span class="sxs-lookup"><span data-stu-id="ef08b-105">This article describes the format of the JSON metadata file.</span></span>

<span data-ttu-id="ef08b-106">Сведения о других файлах, которые необходимо добавить в проект надстройки для включения настраиваемых функций, см. в статье [Создание пользовательских функций в Excel](custom-functions-overview.md).</span><span class="sxs-lookup"><span data-stu-id="ef08b-106">For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="example-metadata"></a><span data-ttu-id="ef08b-107">Пример метаданных</span><span class="sxs-lookup"><span data-stu-id="ef08b-107">Example metadata</span></span>

<span data-ttu-id="ef08b-108">В примере кода ниже показано содержимое JSON-файла метаданных для надстройки, определяющей настраиваемые функции.</span><span class="sxs-lookup"><span data-stu-id="ef08b-108">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="ef08b-109">В следующих за этим примером разделах приводятся подробные сведения об отдельных свойствах, представленных в этом примере JSON.</span><span class="sxs-lookup"><span data-stu-id="ef08b-109">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
> <span data-ttu-id="ef08b-110">Пример готового JSON-файла приводится в репозитории GitHub [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).</span><span class="sxs-lookup"><span data-stu-id="ef08b-110">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json) GitHub repository.</span></span>

## <a name="functions"></a><span data-ttu-id="ef08b-111">functions</span><span class="sxs-lookup"><span data-stu-id="ef08b-111">functions</span></span> 

<span data-ttu-id="ef08b-112">Свойство `functions` представляет собой массив объектов настраиваемых функций.</span><span class="sxs-lookup"><span data-stu-id="ef08b-112">The `functions` property is an array of custom function objects.</span></span> <span data-ttu-id="ef08b-113">В таблице ниже приведены свойства каждого объекта.</span><span class="sxs-lookup"><span data-stu-id="ef08b-113">The following table lists the properties of the SP.ContentTypeCreationInformation object.</span></span>

|  <span data-ttu-id="ef08b-114">Свойство</span><span class="sxs-lookup"><span data-stu-id="ef08b-114">Property</span></span>  |  <span data-ttu-id="ef08b-115">Тип данных</span><span class="sxs-lookup"><span data-stu-id="ef08b-115">Data type</span></span>  |  <span data-ttu-id="ef08b-116">Обязательное</span><span class="sxs-lookup"><span data-stu-id="ef08b-116">Required</span></span>  |  <span data-ttu-id="ef08b-117">Описание</span><span class="sxs-lookup"><span data-stu-id="ef08b-117">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="ef08b-118">string</span><span class="sxs-lookup"><span data-stu-id="ef08b-118">string</span></span>  |  <span data-ttu-id="ef08b-119">Нет</span><span class="sxs-lookup"><span data-stu-id="ef08b-119">No</span></span>  |  <span data-ttu-id="ef08b-120">Описание функции, которое отображается пользователям в Excel</span><span class="sxs-lookup"><span data-stu-id="ef08b-120">The description of the function that end users see in Excel.</span></span> <span data-ttu-id="ef08b-121">(например, **преобразует значение по шкале Цельсия в температуру по шкале Фаренгейта**).</span><span class="sxs-lookup"><span data-stu-id="ef08b-121">For example, **Converts a Celsius value to Fahrenheit**.</span></span> |
|  `helpUrl`  |  <span data-ttu-id="ef08b-122">string</span><span class="sxs-lookup"><span data-stu-id="ef08b-122">string</span></span>  |   <span data-ttu-id="ef08b-123">Нет</span><span class="sxs-lookup"><span data-stu-id="ef08b-123">No</span></span>  |  <span data-ttu-id="ef08b-124">URL-адрес, по которому можно получить сведения о функции</span><span class="sxs-lookup"><span data-stu-id="ef08b-124">URL that provides information about the function.</span></span> <span data-ttu-id="ef08b-125">(отображается в области задач). Пример: **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span><span class="sxs-lookup"><span data-stu-id="ef08b-125">(It is displayed in a task pane.) For example, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span></span> |
| `id`     | <span data-ttu-id="ef08b-126">string</span><span class="sxs-lookup"><span data-stu-id="ef08b-126">string</span></span> | <span data-ttu-id="ef08b-127">Да</span><span class="sxs-lookup"><span data-stu-id="ef08b-127">Yes</span></span> | <span data-ttu-id="ef08b-128">Уникальный идентификатор для функции.</span><span class="sxs-lookup"><span data-stu-id="ef08b-128">A unique ID for the group.</span></span> <span data-ttu-id="ef08b-129">Этот идентификатор может содержать только буквы, цифры и точки, а его изменение после настройки не допускается.</span><span class="sxs-lookup"><span data-stu-id="ef08b-129">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span> |
|  `name`  |  <span data-ttu-id="ef08b-130">string</span><span class="sxs-lookup"><span data-stu-id="ef08b-130">string</span></span>  |  <span data-ttu-id="ef08b-131">Да</span><span class="sxs-lookup"><span data-stu-id="ef08b-131">Yes</span></span>  |  <span data-ttu-id="ef08b-132">Имя функции, которое отображается пользователям в Excel.</span><span class="sxs-lookup"><span data-stu-id="ef08b-132">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="ef08b-133">В Excel имя этой функции будет присоединено в качестве префикса пространством имен настраиваемой функции, указанным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="ef08b-133">In Excel, this function name will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
|  `options`  |  <span data-ttu-id="ef08b-134">object</span><span class="sxs-lookup"><span data-stu-id="ef08b-134">object</span></span>  |  <span data-ttu-id="ef08b-135">Нет</span><span class="sxs-lookup"><span data-stu-id="ef08b-135">No</span></span>  |  <span data-ttu-id="ef08b-136">Позволяет настроить некоторые аспекты того, как и когда Excel выполняет функцию.</span><span class="sxs-lookup"><span data-stu-id="ef08b-136">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="ef08b-137">Дополнительные сведения см. в [разделе, посвященном объекту options](#options-object).</span><span class="sxs-lookup"><span data-stu-id="ef08b-137">See object load [options](#options-object) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="ef08b-138">array</span><span class="sxs-lookup"><span data-stu-id="ef08b-138">array</span></span>  |  <span data-ttu-id="ef08b-139">Да</span><span class="sxs-lookup"><span data-stu-id="ef08b-139">Yes</span></span>  |  <span data-ttu-id="ef08b-140">Массив, который определяет входные параметры для функции.</span><span class="sxs-lookup"><span data-stu-id="ef08b-140">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="ef08b-141">Дополнительные сведения см. в [разделе, посвященном массиву parameters](#parameters-array).</span><span class="sxs-lookup"><span data-stu-id="ef08b-141">See [parameters array](#parameters-array)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="ef08b-142">object</span><span class="sxs-lookup"><span data-stu-id="ef08b-142">object</span></span>  |  <span data-ttu-id="ef08b-143">Да</span><span class="sxs-lookup"><span data-stu-id="ef08b-143">Yes</span></span>  |  <span data-ttu-id="ef08b-144">Объект, который определяет тип информации, возвращаемый функцией.</span><span class="sxs-lookup"><span data-stu-id="ef08b-144">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="ef08b-145">Подробные сведения см. в [разделе, посвященном объекту result](#result-object).</span><span class="sxs-lookup"><span data-stu-id="ef08b-145">See object load [options](#result-object) for details.</span></span> |

## <a name="options"></a><span data-ttu-id="ef08b-146">options</span><span class="sxs-lookup"><span data-stu-id="ef08b-146">options</span></span>

<span data-ttu-id="ef08b-147">Объект `options` позволяет настроить некоторые аспекты того, как и когда Excel выполняет функцию.</span><span class="sxs-lookup"><span data-stu-id="ef08b-147">The `options` object enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="ef08b-148">В таблице ниже приведены свойства объекта `options`.</span><span class="sxs-lookup"><span data-stu-id="ef08b-148">The following table lists the properties of the SP.FieldRatingScale`options` object.</span></span>

|  <span data-ttu-id="ef08b-149">Свойство</span><span class="sxs-lookup"><span data-stu-id="ef08b-149">Property</span></span>  |  <span data-ttu-id="ef08b-150">Тип данных</span><span class="sxs-lookup"><span data-stu-id="ef08b-150">Data type</span></span>  |  <span data-ttu-id="ef08b-151">Обязательное</span><span class="sxs-lookup"><span data-stu-id="ef08b-151">Required</span></span>  |  <span data-ttu-id="ef08b-152">Описание</span><span class="sxs-lookup"><span data-stu-id="ef08b-152">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="ef08b-153">boolean</span><span class="sxs-lookup"><span data-stu-id="ef08b-153">boolean</span></span>  |  <span data-ttu-id="ef08b-154">Нет</span><span class="sxs-lookup"><span data-stu-id="ef08b-154">No</span></span><br/><br/><span data-ttu-id="ef08b-155">Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="ef08b-155">Default value is `false`.</span></span>  |  <span data-ttu-id="ef08b-156">Если это свойство имеет значение `true`, Excel будет вызывать обработчик `onCanceled` каждый раз, когда пользователь будет предпринимать действия, которые приводят к отмене функции (например, вручную вызывает пересчет или редактирует ячейку, на которую ссылается функция).</span><span class="sxs-lookup"><span data-stu-id="ef08b-156">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="ef08b-157">Если вы используете это свойство, Excel вызовет функцию JavaScript с дополнительным параметром `caller`.</span><span class="sxs-lookup"><span data-stu-id="ef08b-157">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="ef08b-158">(***Не*** регистрируйте этот параметр в свойстве `parameters`.)</span><span class="sxs-lookup"><span data-stu-id="ef08b-158">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="ef08b-159">В тексте функции обработчик необходимо назначить элементу `caller.onCanceled`.</span><span class="sxs-lookup"><span data-stu-id="ef08b-159">In the body of the function, a handler must be assigned to the `caller.onCanceled` member.</span></span> <span data-ttu-id="ef08b-160">Дополнительные сведения см. в разделе [Отмена функции](custom-functions-overview.md#canceling-a-function).</span><span class="sxs-lookup"><span data-stu-id="ef08b-160">For more information, see [Canceling a function](custom-functions-overview.md#canceling-a-function).</span></span> |
|  `stream`  |  <span data-ttu-id="ef08b-161">boolean</span><span class="sxs-lookup"><span data-stu-id="ef08b-161">boolean</span></span>  |  <span data-ttu-id="ef08b-162">Нет</span><span class="sxs-lookup"><span data-stu-id="ef08b-162">No</span></span><br/><br/><span data-ttu-id="ef08b-163">Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="ef08b-163">Default value is `false`.</span></span>  |  <span data-ttu-id="ef08b-164">Если это свойство имеет значение `true`, функция может выводить значение в ячейку несколько раз, даже если вызвана всего единожды.</span><span class="sxs-lookup"><span data-stu-id="ef08b-164">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="ef08b-165">Этот параметр полезен для быстро изменяющихся источников данных, таких как цена акций.</span><span class="sxs-lookup"><span data-stu-id="ef08b-165">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="ef08b-166">Если вы используете это свойство, Excel вызовет функцию JavaScript с дополнительным параметром `caller`.</span><span class="sxs-lookup"><span data-stu-id="ef08b-166">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="ef08b-167">(***Не*** регистрируйте этот параметр в свойстве `parameters`.)</span><span class="sxs-lookup"><span data-stu-id="ef08b-167">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="ef08b-168">Функция не должна содержать оператор `return`.</span><span class="sxs-lookup"><span data-stu-id="ef08b-168">The function should have no `return` statement.</span></span> <span data-ttu-id="ef08b-169">Вместо этого результирующее значение передается как аргумент метода обратного вызова `caller.setResult`.</span><span class="sxs-lookup"><span data-stu-id="ef08b-169">Instead, the result value is passed as the argument of the `caller.setResult` callback method.</span></span> <span data-ttu-id="ef08b-170">Дополнительные сведения см. в разделе [Потоковая передача функций](custom-functions-overview.md#streaming-functions).</span><span class="sxs-lookup"><span data-stu-id="ef08b-170">For more information, see [Streaming functions](custom-functions-overview.md#streaming-functions).</span></span> |

## <a name="parameters"></a><span data-ttu-id="ef08b-171">parameters</span><span class="sxs-lookup"><span data-stu-id="ef08b-171">parameters</span></span>

<span data-ttu-id="ef08b-172">Свойство `parameters` представляет собой массив объектов параметров.</span><span class="sxs-lookup"><span data-stu-id="ef08b-172">The `parameters` property is an array of parameter objects.</span></span> <span data-ttu-id="ef08b-173">В таблице ниже приведены свойства каждого объекта.</span><span class="sxs-lookup"><span data-stu-id="ef08b-173">The following table lists the properties of the SP.ContentTypeCreationInformation object.</span></span>

|  <span data-ttu-id="ef08b-174">Свойство</span><span class="sxs-lookup"><span data-stu-id="ef08b-174">Property</span></span>  |  <span data-ttu-id="ef08b-175">Тип данных</span><span class="sxs-lookup"><span data-stu-id="ef08b-175">Data type</span></span>  |  <span data-ttu-id="ef08b-176">Обязательное</span><span class="sxs-lookup"><span data-stu-id="ef08b-176">Required</span></span>  |  <span data-ttu-id="ef08b-177">Описание</span><span class="sxs-lookup"><span data-stu-id="ef08b-177">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="ef08b-178">string</span><span class="sxs-lookup"><span data-stu-id="ef08b-178">string</span></span>  |  <span data-ttu-id="ef08b-179">Нет</span><span class="sxs-lookup"><span data-stu-id="ef08b-179">No</span></span> |  <span data-ttu-id="ef08b-180">Описание параметра.</span><span class="sxs-lookup"><span data-stu-id="ef08b-180">A description of the error.</span></span> <span data-ttu-id="ef08b-181">Отображается в IntelliSense Excel.</span><span class="sxs-lookup"><span data-stu-id="ef08b-181">This is displayed in Excel's intelliSense.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="ef08b-182">string</span><span class="sxs-lookup"><span data-stu-id="ef08b-182">string</span></span>  |  <span data-ttu-id="ef08b-183">Нет</span><span class="sxs-lookup"><span data-stu-id="ef08b-183">No</span></span>  |  <span data-ttu-id="ef08b-184">Должно быть **скалярным** (значение, отличное от массива) или **матричным** (двухмерный массив).</span><span class="sxs-lookup"><span data-stu-id="ef08b-184">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="ef08b-185">string</span><span class="sxs-lookup"><span data-stu-id="ef08b-185">string</span></span>  |  <span data-ttu-id="ef08b-186">Да</span><span class="sxs-lookup"><span data-stu-id="ef08b-186">Yes</span></span>  |  <span data-ttu-id="ef08b-187">Имя параметра.</span><span class="sxs-lookup"><span data-stu-id="ef08b-187">The name of the parameter.</span></span> <span data-ttu-id="ef08b-188">Оно отображается в IntelliSense Excel.</span><span class="sxs-lookup"><span data-stu-id="ef08b-188">This name is displayed in Excel's intelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="ef08b-189">string</span><span class="sxs-lookup"><span data-stu-id="ef08b-189">string</span></span>  |  <span data-ttu-id="ef08b-190">Нет</span><span class="sxs-lookup"><span data-stu-id="ef08b-190">No</span></span>  |  <span data-ttu-id="ef08b-191">Тип данных параметра.</span><span class="sxs-lookup"><span data-stu-id="ef08b-191">The System data type of the parameter.</span></span> <span data-ttu-id="ef08b-192">Может иметь значение **boolean**, **number**, **string** или **any**, что позволяет использовать любой из трех предыдущих типов.</span><span class="sxs-lookup"><span data-stu-id="ef08b-192">Can be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> <span data-ttu-id="ef08b-193">Если это свойство не задано, по умолчанию устанавливается тип данных **any**.</span><span class="sxs-lookup"><span data-stu-id="ef08b-193">If this property is not specified, the data type defaults to **any**.</span></span> |

## <a name="result"></a><span data-ttu-id="ef08b-194">result</span><span class="sxs-lookup"><span data-stu-id="ef08b-194">result</span></span>

<span data-ttu-id="ef08b-195">Объект `result` определяет тип информации, возвращаемый функцией.</span><span class="sxs-lookup"><span data-stu-id="ef08b-195">The `result` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="ef08b-196">В таблице ниже приведены свойства объекта `result`.</span><span class="sxs-lookup"><span data-stu-id="ef08b-196">The following table lists the properties of the SP.FieldRatingScale`result` object.</span></span>

|  <span data-ttu-id="ef08b-197">Свойство</span><span class="sxs-lookup"><span data-stu-id="ef08b-197">Property</span></span>  |  <span data-ttu-id="ef08b-198">Тип данных</span><span class="sxs-lookup"><span data-stu-id="ef08b-198">Data type</span></span>  |  <span data-ttu-id="ef08b-199">Обязательное</span><span class="sxs-lookup"><span data-stu-id="ef08b-199">Required</span></span>  |  <span data-ttu-id="ef08b-200">Описание</span><span class="sxs-lookup"><span data-stu-id="ef08b-200">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="ef08b-201">string</span><span class="sxs-lookup"><span data-stu-id="ef08b-201">string</span></span>  |  <span data-ttu-id="ef08b-202">Нет</span><span class="sxs-lookup"><span data-stu-id="ef08b-202">No</span></span>  |  <span data-ttu-id="ef08b-203">Должно быть **скалярным** (значение, отличное от массива) или **матричным** (двухмерный массив).</span><span class="sxs-lookup"><span data-stu-id="ef08b-203">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |
|  `type`  |  <span data-ttu-id="ef08b-204">string</span><span class="sxs-lookup"><span data-stu-id="ef08b-204">string</span></span>  |  <span data-ttu-id="ef08b-205">Да</span><span class="sxs-lookup"><span data-stu-id="ef08b-205">Yes</span></span>  |  <span data-ttu-id="ef08b-206">Тип данных параметра.</span><span class="sxs-lookup"><span data-stu-id="ef08b-206">The System data type of the parameter.</span></span> <span data-ttu-id="ef08b-207">Должен иметь значение **boolean**, **number**, **string** или **any**, что позволяет использовать любой из трех предыдущих типов.</span><span class="sxs-lookup"><span data-stu-id="ef08b-207">Must be **boolean**, **number**, **string**, or **any**, which allows you to use of any of the previous three types.</span></span> |

## <a name="see-also"></a><span data-ttu-id="ef08b-208">См. также</span><span class="sxs-lookup"><span data-stu-id="ef08b-208">See also</span></span>

* [<span data-ttu-id="ef08b-209">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="ef08b-209">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="ef08b-210">Среда выполнения для пользовательских функций Excel</span><span class="sxs-lookup"><span data-stu-id="ef08b-210">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="ef08b-211">Рекомендации по настраиваемым функциям</span><span class="sxs-lookup"><span data-stu-id="ef08b-211">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="ef08b-212">Руководство по настраиваемым функциям в Excel</span><span class="sxs-lookup"><span data-stu-id="ef08b-212">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)
