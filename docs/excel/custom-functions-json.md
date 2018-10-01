---
ms.date: 09/27/2018
description: Определение метаданных для настраиваемых функций в Excel.
title: Метаданные для настраиваемых функций в Excel
ms.openlocfilehash: 025be277a5e436a1ce2885815e9b8cbf9b206799
ms.sourcegitcommit: fdf7f4d686700edd6e6b04b2ea1bd43e59d4a03a
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/28/2018
ms.locfileid: "25348137"
---
# <a name="custom-functions-metadata-preview"></a><span data-ttu-id="664e4-103">Метаданные для настраиваемых функций (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="664e4-103">Custom functions metadata</span></span>

<span data-ttu-id="664e4-p101">При определении [настраиваемых функций](custom-functions-overview.md) в надстройке Excel в проект надстройки необходимо включить файл метаданных JSON, содержащий информацию, необходимую Excel для регистрации настраиваемых функций и предоставления пользователям доступа к ним. В этой статье описан формат JSON-файла метаданных.</span><span class="sxs-lookup"><span data-stu-id="664e4-p101">When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project must include a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end-users. This article describes the format of the JSON metadata file.</span></span>

<span data-ttu-id="664e4-106">Сведения о других файлах, которые необходимо добавить в проект надстройки для включения настраиваемых функций, см. в статье [Создание настраиваемых функций в Excel](custom-functions-overview.md).</span><span class="sxs-lookup"><span data-stu-id="664e4-106">For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="example-metadata"></a><span data-ttu-id="664e4-107">Пример метаданных</span><span class="sxs-lookup"><span data-stu-id="664e4-107">Example metadata</span></span>

<span data-ttu-id="664e4-108">В следующем примере показано содержимое файла метаданных JSON для надстройки, определяющей настраиваемые функции.</span><span class="sxs-lookup"><span data-stu-id="664e4-108">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="664e4-109">В следующих за этим примером разделах приводится подробная информация об отдельных свойствах, представленных в данном примере JSON.</span><span class="sxs-lookup"><span data-stu-id="664e4-109">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
> <span data-ttu-id="664e4-110">Пример готового файла JSON приводится в [репозитории OfficeDev/Excel-Custom-Functions GitHub](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).</span><span class="sxs-lookup"><span data-stu-id="664e4-110">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions GitHub repository](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).</span></span>

## <a name="functions"></a><span data-ttu-id="664e4-111">functions</span><span class="sxs-lookup"><span data-stu-id="664e4-111">functions</span></span> 

<span data-ttu-id="664e4-112">Свойство `functions` представляет собой массив объектов настраиваемых функций.</span><span class="sxs-lookup"><span data-stu-id="664e4-112">The `functions` property is an array of objects.</span></span> <span data-ttu-id="664e4-113">В следующей таблице приводятся свойства каждого объекта.</span><span class="sxs-lookup"><span data-stu-id="664e4-113">The following table lists the properties of the SP.ContentTypeCreationInformation object.</span></span>

|  <span data-ttu-id="664e4-114">Свойство</span><span class="sxs-lookup"><span data-stu-id="664e4-114">Property</span></span>  |  <span data-ttu-id="664e4-115">Тип данных</span><span class="sxs-lookup"><span data-stu-id="664e4-115">Data type</span></span>  |  <span data-ttu-id="664e4-116">Обязательное</span><span class="sxs-lookup"><span data-stu-id="664e4-116">Required</span></span>  |  <span data-ttu-id="664e4-117">Описание</span><span class="sxs-lookup"><span data-stu-id="664e4-117">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="664e4-118">string</span><span class="sxs-lookup"><span data-stu-id="664e4-118">string</span></span>  |  <span data-ttu-id="664e4-119">Нет</span><span class="sxs-lookup"><span data-stu-id="664e4-119">No</span></span>  |  <span data-ttu-id="664e4-p104">Описание функции, которое пользователи видят в Excel. Например, **преобразует значение по шкале Цельсия в температуру по шкале Фаренгейта**.</span><span class="sxs-lookup"><span data-stu-id="664e4-p104">A description of the function that appears in the Excel UI. For example, **Converts a Celsius value to Fahrenheit**.</span></span> |
|  `helpUrl`  |  <span data-ttu-id="664e4-122">string</span><span class="sxs-lookup"><span data-stu-id="664e4-122">string</span></span>  |   <span data-ttu-id="664e4-123">Нет</span><span class="sxs-lookup"><span data-stu-id="664e4-123">No</span></span>  |  <span data-ttu-id="664e4-p105">URL-адрес, который предоставляет сведения о функции (отображается в области задач). Например, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span><span class="sxs-lookup"><span data-stu-id="664e4-p105">URL where users can get information about the function. (It is displayed in a task pane.) For example, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span></span> |
| `id`     | <span data-ttu-id="664e4-126">string</span><span class="sxs-lookup"><span data-stu-id="664e4-126">string</span></span> | <span data-ttu-id="664e4-127">Да</span><span class="sxs-lookup"><span data-stu-id="664e4-127">Yes</span></span> | <span data-ttu-id="664e4-128">Уникальный идентификатор функции.</span><span class="sxs-lookup"><span data-stu-id="664e4-128">A unique ID for the group.</span></span> <span data-ttu-id="664e4-129">Изменение этого идентификатора после его настройки не допускается.</span><span class="sxs-lookup"><span data-stu-id="664e4-129">This ID should not be changed after it is set.</span></span> |
|  `name`  |  <span data-ttu-id="664e4-130">string</span><span class="sxs-lookup"><span data-stu-id="664e4-130">string</span></span>  |  <span data-ttu-id="664e4-131">Да</span><span class="sxs-lookup"><span data-stu-id="664e4-131">Yes</span></span>  |  <span data-ttu-id="664e4-132">Название функции, которое пользователи видят в Excel.</span><span class="sxs-lookup"><span data-stu-id="664e4-132">The name of the function that end users see in Excel.</span></span> <span data-ttu-id="664e4-133">В Excel название этой функции будет иметь префикс пространства имен настраиваемых функций, указанного в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="664e4-133">In the autocomplete menu, this value will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
|  `options`  |  <span data-ttu-id="664e4-134">object</span><span class="sxs-lookup"><span data-stu-id="664e4-134">object</span></span>  |  <span data-ttu-id="664e4-135">Нет</span><span class="sxs-lookup"><span data-stu-id="664e4-135">No</span></span>  |  <span data-ttu-id="664e4-136">Это свойство позволяет настраивать некоторые аспекты того, как и когда Excel выполняет эту функцию.</span><span class="sxs-lookup"><span data-stu-id="664e4-136">The  property enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="664e4-137">См. [объект параметров](#options-object) для получения дополнительной информации.</span><span class="sxs-lookup"><span data-stu-id="664e4-137">See [options object](#options-object) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="664e4-138">array</span><span class="sxs-lookup"><span data-stu-id="664e4-138">array</span></span>  |  <span data-ttu-id="664e4-139">Да</span><span class="sxs-lookup"><span data-stu-id="664e4-139">Yes</span></span>  |  <span data-ttu-id="664e4-140">Массив, который определяет входные параметры для функции.</span><span class="sxs-lookup"><span data-stu-id="664e4-140">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="664e4-141">См. [массив параметров](#parameters-array) для получения дополнительной информации.</span><span class="sxs-lookup"><span data-stu-id="664e4-141">See [parameters array](#parameters-array)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="664e4-142">object</span><span class="sxs-lookup"><span data-stu-id="664e4-142">object</span></span>  |  <span data-ttu-id="664e4-143">Да</span><span class="sxs-lookup"><span data-stu-id="664e4-143">Yes</span></span>  |  <span data-ttu-id="664e4-144">Объект, который определяет тип возвращаемой функцией информации.</span><span class="sxs-lookup"><span data-stu-id="664e4-144">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="664e4-145">См. [объект результата](#result-object) для получения дополнительной информации.</span><span class="sxs-lookup"><span data-stu-id="664e4-145">See [result object](#result-object) for details.</span></span> |

## <a name="options"></a><span data-ttu-id="664e4-146">options</span><span class="sxs-lookup"><span data-stu-id="664e4-146">options</span></span>

<span data-ttu-id="664e4-147">Объект `options` позволяет настроить некоторые аспекты того, как и когда Excel выполняет данные функции.</span><span class="sxs-lookup"><span data-stu-id="664e4-147">The `options` property enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="664e4-148">В следующей таблице описываются свойства объекта `options`.</span><span class="sxs-lookup"><span data-stu-id="664e4-148">The following table lists the properties of the SP.FieldRatingScale`options` object.</span></span>

|  <span data-ttu-id="664e4-149">Свойство</span><span class="sxs-lookup"><span data-stu-id="664e4-149">Property</span></span>  |  <span data-ttu-id="664e4-150">Тип данных</span><span class="sxs-lookup"><span data-stu-id="664e4-150">Data type</span></span>  |  <span data-ttu-id="664e4-151">Обязательное</span><span class="sxs-lookup"><span data-stu-id="664e4-151">Required</span></span>  |  <span data-ttu-id="664e4-152">Описание</span><span class="sxs-lookup"><span data-stu-id="664e4-152">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="664e4-153">boolean</span><span class="sxs-lookup"><span data-stu-id="664e4-153">boolean</span></span>  |  <span data-ttu-id="664e4-154">Нет</span><span class="sxs-lookup"><span data-stu-id="664e4-154">No</span></span><br/><br/><span data-ttu-id="664e4-155">Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="664e4-155">Default value is 4.</span></span>  |  <span data-ttu-id="664e4-156">Если значение `true`, Excel будет вызывать обработчик `onCanceled` каждый раз, когда пользователь будет предпринимать действия, которые имеют эффект отмены функции, например, вручную вызывая пересчет или редактирование ячейки, на которую ссылается функция.</span><span class="sxs-lookup"><span data-stu-id="664e4-156">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="664e4-157">Если вы используете эту опцию, Excel вызовет функцию JavaScript с дополнительным `caller` параметром.</span><span class="sxs-lookup"><span data-stu-id="664e4-157">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="664e4-158">(Не \*\*\*\*\*\* регистрируйте свои параметры в свойстве `parameters`).</span><span class="sxs-lookup"><span data-stu-id="664e4-158">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="664e4-159">В теле функции обработчик необходимо назначить члену `caller.onCanceled`.</span><span class="sxs-lookup"><span data-stu-id="664e4-159">In the body of the function, a handler must be assigned to the `caller.onCanceled` member.</span></span> <span data-ttu-id="664e4-160">Для получения дополнительной информации см. [Отмена функции](custom-functions-overview.md#canceling-a-function).</span><span class="sxs-lookup"><span data-stu-id="664e4-160">For more information, see [Canceling a function](custom-functions-overview.md#canceling-a-function).</span></span> |
|  `stream`  |  <span data-ttu-id="664e4-161">boolean</span><span class="sxs-lookup"><span data-stu-id="664e4-161">boolean</span></span>  |  <span data-ttu-id="664e4-162">Нет</span><span class="sxs-lookup"><span data-stu-id="664e4-162">No</span></span><br/><br/><span data-ttu-id="664e4-163">Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="664e4-163">Default value is 4.</span></span>  |  <span data-ttu-id="664e4-164">Если значение `true`, функция может выводить значение в ячейку несколько раз, даже если была вызвана всего один раз.</span><span class="sxs-lookup"><span data-stu-id="664e4-164">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="664e4-165">Этот параметр полезен для быстро изменяющихся источников данных, таких как цена акций.</span><span class="sxs-lookup"><span data-stu-id="664e4-165">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="664e4-166">Если вы используете эту опцию, Excel вызовет функцию JavaScript с дополнительным параметром `caller`.</span><span class="sxs-lookup"><span data-stu-id="664e4-166">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="664e4-167">(Не \*\*\*\*\*\* регистрируйте свои параметры в свойстве `parameters`.)</span><span class="sxs-lookup"><span data-stu-id="664e4-167">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="664e4-168">Функция не должна содержать оператор `return`.</span><span class="sxs-lookup"><span data-stu-id="664e4-168">The function should have no `return` statement.</span></span> <span data-ttu-id="664e4-169">Вместо этого результирующее значение передается как аргумент метода обратного вызова `caller.setResult`.</span><span class="sxs-lookup"><span data-stu-id="664e4-169">Instead, the result value is passed as the argument of the `caller.setResult` callback method.</span></span> <span data-ttu-id="664e4-170">Для получения дополнительной информации см. статью [Потоковые функции](custom-functions-overview.md#streamed-functions).</span><span class="sxs-lookup"><span data-stu-id="664e4-170">For more information, see [Streamed functions](custom-functions-overview.md#streamed-functions).</span></span> |

## <a name="parameters"></a><span data-ttu-id="664e4-171">parameters</span><span class="sxs-lookup"><span data-stu-id="664e4-171">parameters</span></span>

<span data-ttu-id="664e4-172">Свойство `parameters` представляет собой массив параметров объекта.</span><span class="sxs-lookup"><span data-stu-id="664e4-172">The `parameters` property is an array of objects.</span></span> <span data-ttu-id="664e4-173">В следующей таблице приводятся свойства каждого объекта.</span><span class="sxs-lookup"><span data-stu-id="664e4-173">The following table lists the properties of the SP.ContentTypeCreationInformation object.</span></span>

|  <span data-ttu-id="664e4-174">Свойство</span><span class="sxs-lookup"><span data-stu-id="664e4-174">Property</span></span>  |  <span data-ttu-id="664e4-175">Тип данных</span><span class="sxs-lookup"><span data-stu-id="664e4-175">Data type</span></span>  |  <span data-ttu-id="664e4-176">Обязательное</span><span class="sxs-lookup"><span data-stu-id="664e4-176">Required</span></span>  |  <span data-ttu-id="664e4-177">Описание</span><span class="sxs-lookup"><span data-stu-id="664e4-177">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="664e4-178">string</span><span class="sxs-lookup"><span data-stu-id="664e4-178">string</span></span>  |  <span data-ttu-id="664e4-179">Нет</span><span class="sxs-lookup"><span data-stu-id="664e4-179">No</span></span> |  <span data-ttu-id="664e4-180">Описание параметра.</span><span class="sxs-lookup"><span data-stu-id="664e4-180">A description of the parameter.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="664e4-181">string</span><span class="sxs-lookup"><span data-stu-id="664e4-181">string</span></span>  |  <span data-ttu-id="664e4-182">Нет</span><span class="sxs-lookup"><span data-stu-id="664e4-182">No</span></span>  |  <span data-ttu-id="664e4-183">Должно быть **скалярным** (значение, не являющееся массивом) или **матричным** (двухмерный массив).</span><span class="sxs-lookup"><span data-stu-id="664e4-183">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="664e4-184">string</span><span class="sxs-lookup"><span data-stu-id="664e4-184">string</span></span>  |  <span data-ttu-id="664e4-185">Да</span><span class="sxs-lookup"><span data-stu-id="664e4-185">Yes</span></span>  |  <span data-ttu-id="664e4-186">Имя параметра.</span><span class="sxs-lookup"><span data-stu-id="664e4-186">The name of the parameter.</span></span> <span data-ttu-id="664e4-187">Это имя отображается в IntelliSense Excel.</span><span class="sxs-lookup"><span data-stu-id="664e4-187">This name is displayed in Excel's IntelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="664e4-188">string</span><span class="sxs-lookup"><span data-stu-id="664e4-188">string</span></span>  |  <span data-ttu-id="664e4-189">Нет</span><span class="sxs-lookup"><span data-stu-id="664e4-189">No</span></span>  |  <span data-ttu-id="664e4-190">Тип данных параметра.</span><span class="sxs-lookup"><span data-stu-id="664e4-190">The data type of the parameter.</span></span> <span data-ttu-id="664e4-191">Должен представлять собой значение типа **boolean**, **number** или **string**.</span><span class="sxs-lookup"><span data-stu-id="664e4-191">Must be "boolean", "number", or "string".</span></span>  |

## <a name="result"></a><span data-ttu-id="664e4-192">result</span><span class="sxs-lookup"><span data-stu-id="664e4-192">result</span></span>

<span data-ttu-id="664e4-193">Объект `results`, определяющий тип возвращаемой функцией информации.</span><span class="sxs-lookup"><span data-stu-id="664e4-193">The `results` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="664e4-194">В следующей таблице описываются свойства объекта `result`.</span><span class="sxs-lookup"><span data-stu-id="664e4-194">The following table lists the properties of the SP.FieldRatingScale`result` object.</span></span>

|  <span data-ttu-id="664e4-195">Свойство</span><span class="sxs-lookup"><span data-stu-id="664e4-195">Property</span></span>  |  <span data-ttu-id="664e4-196">Тип данных</span><span class="sxs-lookup"><span data-stu-id="664e4-196">Data type</span></span>  |  <span data-ttu-id="664e4-197">Обязательное</span><span class="sxs-lookup"><span data-stu-id="664e4-197">Required</span></span>  |  <span data-ttu-id="664e4-198">Описание</span><span class="sxs-lookup"><span data-stu-id="664e4-198">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="664e4-199">string</span><span class="sxs-lookup"><span data-stu-id="664e4-199">string</span></span>  |  <span data-ttu-id="664e4-200">Нет</span><span class="sxs-lookup"><span data-stu-id="664e4-200">No</span></span>  |  <span data-ttu-id="664e4-201">Должно быть **скалярным** (значение, не являющееся массивом) или **матричным** (двухмерный массив).</span><span class="sxs-lookup"><span data-stu-id="664e4-201">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |
|  `type`  |  <span data-ttu-id="664e4-202">string</span><span class="sxs-lookup"><span data-stu-id="664e4-202">string</span></span>  |  <span data-ttu-id="664e4-203">Да</span><span class="sxs-lookup"><span data-stu-id="664e4-203">Yes</span></span>  |  <span data-ttu-id="664e4-204">Тип данных параметра.</span><span class="sxs-lookup"><span data-stu-id="664e4-204">The data type of the parameter.</span></span> <span data-ttu-id="664e4-205">Должен представлять собой значение типа **boolean**, **number** или **string**.</span><span class="sxs-lookup"><span data-stu-id="664e4-205">Must be "boolean", "number", or "string".</span></span>  |

## <a name="see-also"></a><span data-ttu-id="664e4-206">См. также</span><span class="sxs-lookup"><span data-stu-id="664e4-206">See also</span></span>

* [<span data-ttu-id="664e4-207">Создание настраиваемых функций в Excel</span><span class="sxs-lookup"><span data-stu-id="664e4-207">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="664e4-208">Среда выполнения для настраиваемых функций Excel</span><span class="sxs-lookup"><span data-stu-id="664e4-208">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="664e4-209">Рекомендации по настраиваемым функциям</span><span class="sxs-lookup"><span data-stu-id="664e4-209">Custom functions best practices</span></span>](custom-functions-best-practices.md)