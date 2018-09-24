---
ms.date: 09/20/2018
description: Определение метаданных для настраиваемых функций в Excel.
title: Метаданные для настраиваемых функций в Excel
ms.openlocfilehash: 815b0c6e65966867d9e5d953a40ffc705a63ee63
ms.sourcegitcommit: 470d8212b256275587e651abaa6f28beafebcab4
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/21/2018
ms.locfileid: "24062146"
---
# <a name="custom-functions-metadata"></a><span data-ttu-id="2738f-103">Метаданные настраиваемых функций</span><span class="sxs-lookup"><span data-stu-id="2738f-103">Custom functions metadata</span></span>

<span data-ttu-id="2738f-104">При определении [настраиваемых функций](custom-functions-overview.md) в надстройке Excel в проект надстройки необходимо включить файл метаданных JSON, содержащий информацию о том, что требуется Excel для того, чтобы зарегистрировать настраиваемые функции и сделать их доступными для пользователей.</span><span class="sxs-lookup"><span data-stu-id="2738f-104">When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project must include a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end-users.</span></span> <span data-ttu-id="2738f-105">В этой статье описывается формат файла метаданных JSON.</span><span class="sxs-lookup"><span data-stu-id="2738f-105">This article describes the format of the JSON file with examples.</span></span>

> [!NOTE]
> <span data-ttu-id="2738f-106">Сведения о других файлах, котрые необходимо включить в проект надстройки для включения настраиваемых функций, см. в статье [Создание настраиваемых функций в Excel](custom-functions-overview.md#learn-the-basics).</span><span class="sxs-lookup"><span data-stu-id="2738f-106">For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md#learn-the-basics).</span></span>

## <a name="example-metadata"></a><span data-ttu-id="2738f-107">Пример метаданных</span><span class="sxs-lookup"><span data-stu-id="2738f-107">Example metadata</span></span>

<span data-ttu-id="2738f-108">В следующем примере показано содержимое файла метаданных JSON для надстройки, определяющей настраиваемые функции.</span><span class="sxs-lookup"><span data-stu-id="2738f-108">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions.</span></span> <span data-ttu-id="2738f-109">В следующих за этим примером разделах приводится подробная информация об отдельных свойствах, рассматриваемых в данном примере JSON.</span><span class="sxs-lookup"><span data-stu-id="2738f-109">The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

```json
{
    "functions": [
        {
            "id": "ADD42",
            "name": "ADD42",
            "description":  "Adds 42 to the input number",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "num",
                    "description": "Number",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ]
        },
        {
            "id": "ADD42ASYNC",
            "name": "ADD42ASYNC",
            "description":  "asynchronously wait 250ms, then add 42",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "num",
                    "description": "Number",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ]
        },
        {
            "id": "ISEVEN",
            "name": "ISEVEN", 
            "description":  "Determines whether a number is even",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "boolean",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
                    "name": "num",
                    "description": "the number to be evaluated",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ]
        },
        {
            "id": "GETDAY",
            "name": "GETDAY",
            "description": "Gets the day of the week",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "string"
            },
            "parameters": []
        },
        {
            "id": "INCREMENTVALUE",
            "name": "INCREMENTVALUE", 
            "description":  "Counts up from zero",
            "helpUrl": "http://dev.office.com",
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
            "description":  "gets the second highest number from a range",
            "helpUrl": "http://dev.office.com",
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
> <span data-ttu-id="2738f-110">Пример готового файла JSON приводится в [репозитории OfficeDev/Excel-Custom-Functions GitHub](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).</span><span class="sxs-lookup"><span data-stu-id="2738f-110">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions GitHub repository](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).</span></span>

## <a name="functions"></a><span data-ttu-id="2738f-111">functions</span><span class="sxs-lookup"><span data-stu-id="2738f-111">functions</span></span> 

<span data-ttu-id="2738f-112">Свойство `functions` представляет собой массив объектов настраиваемых функций.</span><span class="sxs-lookup"><span data-stu-id="2738f-112">The `functions` property is an array of objects.</span></span> <span data-ttu-id="2738f-113">В следующей таблице приводятся свойства каждого объекта.</span><span class="sxs-lookup"><span data-stu-id="2738f-113">The following table lists the properties of the SP.ContentTypeCreationInformation object.</span></span>

|  <span data-ttu-id="2738f-114">Свойство</span><span class="sxs-lookup"><span data-stu-id="2738f-114">Property</span></span>  |  <span data-ttu-id="2738f-115">Тип данных</span><span class="sxs-lookup"><span data-stu-id="2738f-115">Data type</span></span>  |  <span data-ttu-id="2738f-116">Обязательное</span><span class="sxs-lookup"><span data-stu-id="2738f-116">Required</span></span>  |  <span data-ttu-id="2738f-117">Описание</span><span class="sxs-lookup"><span data-stu-id="2738f-117">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="2738f-118">string</span><span class="sxs-lookup"><span data-stu-id="2738f-118">string</span></span>  |  <span data-ttu-id="2738f-119">Нет</span><span class="sxs-lookup"><span data-stu-id="2738f-119">No</span></span>  |  <span data-ttu-id="2738f-120">Описание функции, отображаемое в пользовательском интерфейсе Excel.</span><span class="sxs-lookup"><span data-stu-id="2738f-120">A description of the function that appears in the Excel UI.</span></span> <span data-ttu-id="2738f-121">К примеру, **Преобразует градусы Цельсия в градусы Фаренгейта**.</span><span class="sxs-lookup"><span data-stu-id="2738f-121">For example, "Converts a Celsius value to Fahrenheit".</span></span> |
|  `helpUrl`  |  <span data-ttu-id="2738f-122">string</span><span class="sxs-lookup"><span data-stu-id="2738f-122">string</span></span>  |   <span data-ttu-id="2738f-123">Нет</span><span class="sxs-lookup"><span data-stu-id="2738f-123">No</span></span>  |  <span data-ttu-id="2738f-124">URL-адрес, позволяющий пользователю получить информацию о функции.</span><span class="sxs-lookup"><span data-stu-id="2738f-124">URL where your users can get help about the function.</span></span> <span data-ttu-id="2738f-125">(Отображается в области задач). Например, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span><span class="sxs-lookup"><span data-stu-id="2738f-125">(It is displayed in a taskpane.) For example, "http://contoso.com/help/convertcelsiustofahrenheit.html"</span></span> |
| `id`     | <span data-ttu-id="2738f-126">string</span><span class="sxs-lookup"><span data-stu-id="2738f-126">string</span></span> | <span data-ttu-id="2738f-127">Да</span><span class="sxs-lookup"><span data-stu-id="2738f-127">Yes</span></span> | <span data-ttu-id="2738f-128">Уникальный идентификатор функции.</span><span class="sxs-lookup"><span data-stu-id="2738f-128">A unique ID for the group.</span></span> <span data-ttu-id="2738f-129">Изменение этого идентификатора после его настройки не допускается.</span><span class="sxs-lookup"><span data-stu-id="2738f-129">This ID should not be changed after it is set.</span></span> |
|  `name`  |  <span data-ttu-id="2738f-130">string</span><span class="sxs-lookup"><span data-stu-id="2738f-130">string</span></span>  |  <span data-ttu-id="2738f-131">Да</span><span class="sxs-lookup"><span data-stu-id="2738f-131">Yes</span></span>  |  <span data-ttu-id="2738f-132">Имя функции, которая будет отображаться (добавлено в пространстве имен) в пользовательском интерфейсе Excel, когда пользователь выбирает функцию.</span><span class="sxs-lookup"><span data-stu-id="2738f-132">The name of the function as it will appear (prepended with a namespace) in the Excel UI when a user is selecting a function.</span></span> <span data-ttu-id="2738f-133">Его совпадение с именем функции, указанным при ее определении в JavaScript, не обязательно.</span><span class="sxs-lookup"><span data-stu-id="2738f-133">It should be the same as the function's name where it is defined in the JavaScript.</span></span> |
|  `options`  |  <span data-ttu-id="2738f-134">object</span><span class="sxs-lookup"><span data-stu-id="2738f-134">object</span></span>  |  <span data-ttu-id="2738f-135">Нет</span><span class="sxs-lookup"><span data-stu-id="2738f-135">No</span></span>  |  <span data-ttu-id="2738f-136">Это свойство позволяет настраивать некоторые аспекты того, как и когда Excel выполняет эту функцию.</span><span class="sxs-lookup"><span data-stu-id="2738f-136">The  property enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="2738f-137">См. [объект параметров](#options-object) для получения дополнительной информации.</span><span class="sxs-lookup"><span data-stu-id="2738f-137">See [options object](#options-object) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="2738f-138">array</span><span class="sxs-lookup"><span data-stu-id="2738f-138">array</span></span>  |  <span data-ttu-id="2738f-139">Да</span><span class="sxs-lookup"><span data-stu-id="2738f-139">Yes</span></span>  |  <span data-ttu-id="2738f-140">Массив, который определяет входные параметры для функции.</span><span class="sxs-lookup"><span data-stu-id="2738f-140">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="2738f-141">См. [массив параметров](#parameters-array) для получения дополнительной информации.</span><span class="sxs-lookup"><span data-stu-id="2738f-141">See [parameters array](#parameters-array)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="2738f-142">object</span><span class="sxs-lookup"><span data-stu-id="2738f-142">object</span></span>  |  <span data-ttu-id="2738f-143">Да</span><span class="sxs-lookup"><span data-stu-id="2738f-143">Yes</span></span>  |  <span data-ttu-id="2738f-144">Объект, который определяет тип возвращаемой функцией информации.</span><span class="sxs-lookup"><span data-stu-id="2738f-144">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="2738f-145">См. [объект результата](#result-object) для получения дополнительной информации.</span><span class="sxs-lookup"><span data-stu-id="2738f-145">See [result object](#result-object) for details.</span></span> |

## <a name="options"></a><span data-ttu-id="2738f-146">options</span><span class="sxs-lookup"><span data-stu-id="2738f-146">options</span></span>

<span data-ttu-id="2738f-147">Объект `options` позволяет настроить некоторые аспекты того, как и когда Excel выполняет данные функции.</span><span class="sxs-lookup"><span data-stu-id="2738f-147">The `options` property enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="2738f-148">В следующей таблице описываются свойства объекта `options`.</span><span class="sxs-lookup"><span data-stu-id="2738f-148">The following table lists the properties of the SP.FieldRatingScale`options` object.</span></span>

|  <span data-ttu-id="2738f-149">Свойство</span><span class="sxs-lookup"><span data-stu-id="2738f-149">Property</span></span>  |  <span data-ttu-id="2738f-150">Тип данных</span><span class="sxs-lookup"><span data-stu-id="2738f-150">Data type</span></span>  |  <span data-ttu-id="2738f-151">Обязательное</span><span class="sxs-lookup"><span data-stu-id="2738f-151">Required</span></span>  |  <span data-ttu-id="2738f-152">Description</span><span class="sxs-lookup"><span data-stu-id="2738f-152">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="2738f-153">boolean</span><span class="sxs-lookup"><span data-stu-id="2738f-153">boolean</span></span>  |  <span data-ttu-id="2738f-154">Нет, значение по умолчанию — `false`.</span><span class="sxs-lookup"><span data-stu-id="2738f-154">No, default is `false`.</span></span>  |  <span data-ttu-id="2738f-155">Если `true`, Excel вызывает обработчика `onCanceled` всякий раз, когда пользователь предпринимает действие, которое имеет эффект отмены функции, например, вручную вызывая пересчет или редактирование ячейки, на которую ссылается функция.</span><span class="sxs-lookup"><span data-stu-id="2738f-155">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="2738f-156">Если вы используете эту опцию, Excel вызовет функцию JavaScript с дополнительным параметром `caller`.</span><span class="sxs-lookup"><span data-stu-id="2738f-156">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="2738f-157">(***Не***регистрируйте этот параметр в свойстве `parameters`.)</span><span class="sxs-lookup"><span data-stu-id="2738f-157">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="2738f-158">В теле функции обработчик необходимо назначить члену `caller.onCanceled`.</span><span class="sxs-lookup"><span data-stu-id="2738f-158">In the body of the function, a handler must be assigned to the `caller.onCanceled` member.</span></span> <span data-ttu-id="2738f-159">Для получения дополнительной информации см. [Отмена функции](custom-functions-overview.md#canceling-a-function).</span><span class="sxs-lookup"><span data-stu-id="2738f-159">For more information, see [Canceling a function](custom-functions-overview.md#canceling-a-function).</span></span> |
|  `stream`  |  <span data-ttu-id="2738f-160">boolean</span><span class="sxs-lookup"><span data-stu-id="2738f-160">boolean</span></span>  |  <span data-ttu-id="2738f-161">Нет, значение по умолчанию — `false`.</span><span class="sxs-lookup"><span data-stu-id="2738f-161">No, default is `false`.</span></span>  |  <span data-ttu-id="2738f-162">Если `true`, функция может выводить несколько раз в ячейку даже при вызове только один раз.</span><span class="sxs-lookup"><span data-stu-id="2738f-162">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="2738f-163">Этот параметр полезен для быстро изменяющихся источников данных, таких как цена акций.</span><span class="sxs-lookup"><span data-stu-id="2738f-163">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="2738f-164">Если вы используете эту опцию, Excel вызовет функцию JavaScript с дополнительным параметром `caller`.</span><span class="sxs-lookup"><span data-stu-id="2738f-164">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="2738f-165">(***Не***регистрируйте этот параметр в свойстве `parameters`.)</span><span class="sxs-lookup"><span data-stu-id="2738f-165">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="2738f-166">Функция должна иметь выписку `return`.</span><span class="sxs-lookup"><span data-stu-id="2738f-166">The function should have no `return` statement.</span></span> <span data-ttu-id="2738f-167">Вместо этого результирующее значение передается как аргумент метода обратного вызова `caller.setResult`.</span><span class="sxs-lookup"><span data-stu-id="2738f-167">Instead, the result value is passed as the argument of the `caller.setResult` callback method.</span></span> <span data-ttu-id="2738f-168">Для получения дополнительной информации см. статью [Потоковые функции](custom-functions-overview.md#streamed-functions).</span><span class="sxs-lookup"><span data-stu-id="2738f-168">For more information, see [Streamed functions](custom-functions-overview.md#streamed-functions).</span></span> |

## <a name="parameters"></a><span data-ttu-id="2738f-169">parameters</span><span class="sxs-lookup"><span data-stu-id="2738f-169">parameters</span></span>

<span data-ttu-id="2738f-170">Свойство `parameters` представляет собой массив параметров объекта.</span><span class="sxs-lookup"><span data-stu-id="2738f-170">The `parameters` property is an array of objects.</span></span> <span data-ttu-id="2738f-171">В следующей таблице приводятся свойства каждого объекта.</span><span class="sxs-lookup"><span data-stu-id="2738f-171">The following table lists the properties of the SP.ContentTypeCreationInformation object.</span></span>

|  <span data-ttu-id="2738f-172">Свойство</span><span class="sxs-lookup"><span data-stu-id="2738f-172">Property</span></span>  |  <span data-ttu-id="2738f-173">Тип данных</span><span class="sxs-lookup"><span data-stu-id="2738f-173">Data type</span></span>  |  <span data-ttu-id="2738f-174">Обязательное</span><span class="sxs-lookup"><span data-stu-id="2738f-174">Required</span></span>  |  <span data-ttu-id="2738f-175">Описание</span><span class="sxs-lookup"><span data-stu-id="2738f-175">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="2738f-176">string</span><span class="sxs-lookup"><span data-stu-id="2738f-176">string</span></span>  |  <span data-ttu-id="2738f-177">Нет</span><span class="sxs-lookup"><span data-stu-id="2738f-177">No</span></span> |  <span data-ttu-id="2738f-178">Описание параметра.</span><span class="sxs-lookup"><span data-stu-id="2738f-178">A description of the parameter.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="2738f-179">string</span><span class="sxs-lookup"><span data-stu-id="2738f-179">string</span></span>  |  <span data-ttu-id="2738f-180">Нет</span><span class="sxs-lookup"><span data-stu-id="2738f-180">No</span></span>  |  <span data-ttu-id="2738f-181">Должно быть **скалярным** (значение, не являющееся массивом) или **матричным** (двухмерный массив).</span><span class="sxs-lookup"><span data-stu-id="2738f-181">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="2738f-182">string</span><span class="sxs-lookup"><span data-stu-id="2738f-182">string</span></span>  |  <span data-ttu-id="2738f-183">Да</span><span class="sxs-lookup"><span data-stu-id="2738f-183">Yes</span></span>  |  <span data-ttu-id="2738f-184">Имя параметра.</span><span class="sxs-lookup"><span data-stu-id="2738f-184">The name of the parameter.</span></span> <span data-ttu-id="2738f-185">Это имя отображается в IntelliSense Excel.</span><span class="sxs-lookup"><span data-stu-id="2738f-185">This name is displayed in Excel's IntelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="2738f-186">string</span><span class="sxs-lookup"><span data-stu-id="2738f-186">string</span></span>  |  <span data-ttu-id="2738f-187">Нет</span><span class="sxs-lookup"><span data-stu-id="2738f-187">No</span></span>  |  <span data-ttu-id="2738f-188">Тип данных параметра.</span><span class="sxs-lookup"><span data-stu-id="2738f-188">The data type of the parameter.</span></span> <span data-ttu-id="2738f-189">Должен представлять собой значение типа **boolean**, **number** или **string**.</span><span class="sxs-lookup"><span data-stu-id="2738f-189">Must be "boolean", "number", or "string".</span></span>  |

## <a name="result"></a><span data-ttu-id="2738f-190">result</span><span class="sxs-lookup"><span data-stu-id="2738f-190">result</span></span>

<span data-ttu-id="2738f-191">Объект `results`, определяющий тип возвращаемой функцией информации.</span><span class="sxs-lookup"><span data-stu-id="2738f-191">The `results` object defines the type of information that is returned by the function.</span></span> <span data-ttu-id="2738f-192">В следующей таблице описываются свойства объекта `result`.</span><span class="sxs-lookup"><span data-stu-id="2738f-192">The following table lists the properties of the SP.FieldRatingScale`result` object.</span></span>

|  <span data-ttu-id="2738f-193">Свойство</span><span class="sxs-lookup"><span data-stu-id="2738f-193">Property</span></span>  |  <span data-ttu-id="2738f-194">Тип данных</span><span class="sxs-lookup"><span data-stu-id="2738f-194">Data type</span></span>  |  <span data-ttu-id="2738f-195">Обязательное</span><span class="sxs-lookup"><span data-stu-id="2738f-195">Required</span></span>  |  <span data-ttu-id="2738f-196">Описание</span><span class="sxs-lookup"><span data-stu-id="2738f-196">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="2738f-197">string</span><span class="sxs-lookup"><span data-stu-id="2738f-197">string</span></span>  |  <span data-ttu-id="2738f-198">Нет</span><span class="sxs-lookup"><span data-stu-id="2738f-198">No</span></span>  |  <span data-ttu-id="2738f-199">Должно быть **скалярным** (значение, не являющееся массивом) или **матричным** (двухмерный массив).</span><span class="sxs-lookup"><span data-stu-id="2738f-199">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |
|  `type`  |  <span data-ttu-id="2738f-200">string</span><span class="sxs-lookup"><span data-stu-id="2738f-200">string</span></span>  |  <span data-ttu-id="2738f-201">Да</span><span class="sxs-lookup"><span data-stu-id="2738f-201">Yes</span></span>  |  <span data-ttu-id="2738f-202">Тип данных параметра.</span><span class="sxs-lookup"><span data-stu-id="2738f-202">The data type of the parameter.</span></span> <span data-ttu-id="2738f-203">Должен представлять собой значение типа **boolean**, **number** или **string**.</span><span class="sxs-lookup"><span data-stu-id="2738f-203">Must be "boolean", "number", or "string".</span></span>  |

## <a name="see-also"></a><span data-ttu-id="2738f-204">См. также</span><span class="sxs-lookup"><span data-stu-id="2738f-204">See also</span></span>

* [<span data-ttu-id="2738f-205">Создание настраиваемых функций в Excel</span><span class="sxs-lookup"><span data-stu-id="2738f-205">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="2738f-206">Среда выполнения для настраиваемых функций Excel</span><span class="sxs-lookup"><span data-stu-id="2738f-206">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="2738f-207">Рекомендации по настраиваемым функциям</span><span class="sxs-lookup"><span data-stu-id="2738f-207">Custom functions best practices</span></span>](custom-functions-best-practices.md)