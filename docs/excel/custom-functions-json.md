---
ms.date: 09/27/2018
description: Определение метаданных для настраиваемых функций в Excel.
title: Метаданные для настраиваемых функций в Excel
ms.openlocfilehash: e8af13b8855d6c5e1a3b1ce99edb24445e066756
ms.sourcegitcommit: 563c53bac52b31277ab935f30af648f17c5ed1e2
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/10/2018
ms.locfileid: "25459240"
---
# <a name="custom-functions-metadata-preview"></a><span data-ttu-id="699ec-103">Метаданные для настраиваемых функций (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="699ec-103">Custom functions metadata</span></span>

<span data-ttu-id="699ec-104">При определении [настраиваемых функций](custom-functions-overview.md) в надстройке Excel необходимо включить в проект вашей надстройки файл метаданных JSON, содержащий информацию о том, что требуется Excel для того, чтобы зарегистрировать настраиваемые функции и сделать их доступными для пользователей.</span><span class="sxs-lookup"><span data-stu-id="699ec-104">When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project must include a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users.</span></span> <span data-ttu-id="699ec-105">В этой статье описывается формат файла метаданных JSON.</span><span class="sxs-lookup"><span data-stu-id="699ec-105">This article describes the format of the JSON metadata file.</span></span>

<span data-ttu-id="699ec-106">Сведения о других файлах, котрые необходимо включить в проект надстройки для включения настраиваемых функций, см. в статье [Создание настраиваемых функций в Excel](custom-functions-overview.md).</span><span class="sxs-lookup"><span data-stu-id="699ec-106">For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="example-metadata"></a><span data-ttu-id="699ec-107">Пример метаданных</span><span class="sxs-lookup"><span data-stu-id="699ec-107">Example metadata</span></span>

<span data-ttu-id="699ec-p102">В следующем примере кода показано содержимое файла метаданных JSON для надстройки, определяющей настраиваемые функции. В следующих за этим примером разделах приводится подробная информация об отдельных свойствах, представленных в данном примере JSON.</span><span class="sxs-lookup"><span data-stu-id="699ec-p102">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions. The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
> <span data-ttu-id="699ec-110">Пример готового файла JSON приводится в репозитории GitHub [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).</span><span class="sxs-lookup"><span data-stu-id="699ec-110">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions GitHub repository](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).</span></span>

## <a name="functions"></a><span data-ttu-id="699ec-111">functions</span><span class="sxs-lookup"><span data-stu-id="699ec-111">functions</span></span> 

<span data-ttu-id="699ec-p103"> Свойство `functions` представляет собой массив объектов настраиваемых функций. В следующей таблице приведены свойства каждого объекта.</span><span class="sxs-lookup"><span data-stu-id="699ec-p103">The `functions` property is an array of custom function objects. The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="699ec-114">Свойство</span><span class="sxs-lookup"><span data-stu-id="699ec-114">Property</span></span>  |  <span data-ttu-id="699ec-115">Тип данных</span><span class="sxs-lookup"><span data-stu-id="699ec-115">Data type</span></span>  |  <span data-ttu-id="699ec-116">Обязательное</span><span class="sxs-lookup"><span data-stu-id="699ec-116">Required</span></span>  |  <span data-ttu-id="699ec-117">Описание</span><span class="sxs-lookup"><span data-stu-id="699ec-117">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="699ec-118">string</span><span class="sxs-lookup"><span data-stu-id="699ec-118">string</span></span>  |  <span data-ttu-id="699ec-119">Нет</span><span class="sxs-lookup"><span data-stu-id="699ec-119">No</span></span>  |  <span data-ttu-id="699ec-120">Описание функции, которое пользователи видят в Excel.</span><span class="sxs-lookup"><span data-stu-id="699ec-120">The description of the function that end users see in Excel.</span></span> <span data-ttu-id="699ec-121">Например, **преобразует значение по шкале Цельсия в температуру по шкале Фаренгейта.**.</span><span class="sxs-lookup"><span data-stu-id="699ec-121">For example, **Converts a Celsius value to Fahrenheit**.</span></span> |
|  `helpUrl`  |  <span data-ttu-id="699ec-122">string</span><span class="sxs-lookup"><span data-stu-id="699ec-122">string</span></span>  |   <span data-ttu-id="699ec-123">Нет</span><span class="sxs-lookup"><span data-stu-id="699ec-123">No</span></span>  |  <span data-ttu-id="699ec-124">URL-адрес, который предоставляет сведения о функции.</span><span class="sxs-lookup"><span data-stu-id="699ec-124">URL that provides information about the function.</span></span> <span data-ttu-id="699ec-125">(Отображается в области задач). Например, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span><span class="sxs-lookup"><span data-stu-id="699ec-125">(It is displayed in a task pane.) For example, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span></span> |
| `id`     | <span data-ttu-id="699ec-126">string</span><span class="sxs-lookup"><span data-stu-id="699ec-126">string</span></span> | <span data-ttu-id="699ec-127">Да</span><span class="sxs-lookup"><span data-stu-id="699ec-127">Yes</span></span> | <span data-ttu-id="699ec-p106">Уникальный идентификатор функции. Изменение этого идентификатора после его настройки не допускается.</span><span class="sxs-lookup"><span data-stu-id="699ec-p106">A unique ID for the function. This ID should not be changed after it is set.</span></span> |
|  `name`  |  <span data-ttu-id="699ec-130">string</span><span class="sxs-lookup"><span data-stu-id="699ec-130">string</span></span>  |  <span data-ttu-id="699ec-131">Да</span><span class="sxs-lookup"><span data-stu-id="699ec-131">Yes</span></span>  |  <span data-ttu-id="699ec-p107">Название функции, которое пользователи видят в Excel. В Excel название этой функции будет присоединено в качестве приставки пространством имен настраиваемой функции, указанным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="699ec-p107">The name of the function that end users see in Excel. In Excel, this function name will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
|  `options`  |  <span data-ttu-id="699ec-134">object</span><span class="sxs-lookup"><span data-stu-id="699ec-134">object</span></span>  |  <span data-ttu-id="699ec-135">Нет</span><span class="sxs-lookup"><span data-stu-id="699ec-135">No</span></span>  |  <span data-ttu-id="699ec-p108">Это свойство позволяет настраивать некоторые аспекты того, как и когда Excel выполняет эту функцию. См. [Объект параметров](#options-object)  для получения дополнительной информации.</span><span class="sxs-lookup"><span data-stu-id="699ec-p108">Enables you to customize some aspects of how and when Excel executes the function. See [options object](#options-object) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="699ec-138">array</span><span class="sxs-lookup"><span data-stu-id="699ec-138">array</span></span>  |  <span data-ttu-id="699ec-139">Да</span><span class="sxs-lookup"><span data-stu-id="699ec-139">Yes</span></span>  |  <span data-ttu-id="699ec-p109">Массив, который определяет входные параметры для функции. См. [Массив параметров ](#parameters-array) для получения дополнительной информации.</span><span class="sxs-lookup"><span data-stu-id="699ec-p109">Array that defines the input parameters for the function. See [parameters array](#parameters-array)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="699ec-142">object</span><span class="sxs-lookup"><span data-stu-id="699ec-142">object</span></span>  |  <span data-ttu-id="699ec-143">Да</span><span class="sxs-lookup"><span data-stu-id="699ec-143">Yes</span></span>  |  <span data-ttu-id="699ec-p110">Объект, который определяет тип возвращаемой функцией информации. См. [Объект результата](#result-object) для получения дополнительной информации.</span><span class="sxs-lookup"><span data-stu-id="699ec-p110">Object that defines the type of information that is returned by the function. See [result object](#result-object) for details.</span></span> |

## <a name="options"></a><span data-ttu-id="699ec-146">options</span><span class="sxs-lookup"><span data-stu-id="699ec-146">options</span></span>

<span data-ttu-id="699ec-p111">Объект `options` позволяет настроить некоторые аспекты того, как и когда Excel выполняет данные функции. В следующей таблице описываются свойства объекта  `options`.</span><span class="sxs-lookup"><span data-stu-id="699ec-p111">The `options` object enables you to customize some aspects of how and when Excel executes the function. The following table lists the properties of the `options` object.</span></span>

|  <span data-ttu-id="699ec-149">Свойство</span><span class="sxs-lookup"><span data-stu-id="699ec-149">Property</span></span>  |  <span data-ttu-id="699ec-150">Тип данных</span><span class="sxs-lookup"><span data-stu-id="699ec-150">Data type</span></span>  |  <span data-ttu-id="699ec-151">Обязательное</span><span class="sxs-lookup"><span data-stu-id="699ec-151">Required</span></span>  |  <span data-ttu-id="699ec-152">Описание</span><span class="sxs-lookup"><span data-stu-id="699ec-152">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="699ec-153">boolean</span><span class="sxs-lookup"><span data-stu-id="699ec-153">boolean</span></span>  |  <span data-ttu-id="699ec-154">Нет</span><span class="sxs-lookup"><span data-stu-id="699ec-154">No</span></span><br/><br/><span data-ttu-id="699ec-155">Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="699ec-155">Default value is 4.</span></span>  |  <span data-ttu-id="699ec-p112">Если `true`, Excel будет вызывать обработчик `onCanceled` каждый раз, когда пользователь будет предпринимать действия, которые имеют эффект отмены функции, например, вручную вызывая пересчет или редактирование ячейки, на которую ссылается функция. Если вы используете эту опцию, Excel вызовет функцию JavaScript с дополнительным `caller`  параметром. (Не \*\*\* \*\*\* регистрируйте свои параметры в свойстве `parameters`). В теле функции необходимо назначить обработчик члену `caller.onCanceled`. Для получения дополнительной информации см.  [Отмена функции](custom-functions-overview.md#canceling-a-function).</span><span class="sxs-lookup"><span data-stu-id="699ec-p112">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function. If you use this option, Excel will call the JavaScript function with an additional `caller` parameter. (Do ***not*** register this parameter in the `parameters` property). In the body of the function, a handler must be assigned to the `caller.onCanceled` member. For more information, see [Canceling a function](custom-functions-overview.md#canceling-a-function).</span></span> |
|  `stream`  |  <span data-ttu-id="699ec-161">boolean</span><span class="sxs-lookup"><span data-stu-id="699ec-161">boolean</span></span>  |  <span data-ttu-id="699ec-162">Нет</span><span class="sxs-lookup"><span data-stu-id="699ec-162">No</span></span><br/><br/><span data-ttu-id="699ec-163">Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="699ec-163">Default value is 4.</span></span>  |  <span data-ttu-id="699ec-p113">Если `true`, функция может выводить значение в ячейку несколько раз, даже если была вызвана всего единожды. Этот параметр полезен для быстро изменяющихся источников данных, таких как цена акций. Если вы используете эту опцию, Excel вызовет функцию JavaScript с дополнительным параметром `caller`. (Не \*\*\* \*\*\* регистрируйте свои параметры в свойстве `parameters`). Функция должна содержать оператор `return`. Вместо этого результирующее значение передается как аргумент метода обратного вызова `caller.setResult`. Для получения дополнительной информации см. статью [Потоковые функции](custom-functions-overview.md#streaming-functions).</span><span class="sxs-lookup"><span data-stu-id="699ec-p113">If `true`, the function can output repeatedly to the cell even when invoked only once. This option is useful for rapidly-changing data sources, such as a stock price. If you use this option, Excel will call the JavaScript function with an additional `caller` parameter. (Do ***not*** register this parameter in the `parameters` property). The function should have no `return` statement. Instead, the result value is passed as the argument of the `caller.setResult` callback method. For more information, see [Streamed functions](custom-functions-overview.md#streaming-functions).</span></span> |

## <a name="parameters"></a><span data-ttu-id="699ec-171">parameters</span><span class="sxs-lookup"><span data-stu-id="699ec-171">parameters</span></span>

<span data-ttu-id="699ec-p114">Свойство `parameters`  представляет собой массив параметров объекта. В следующей таблице приводятся свойства каждого объекта.</span><span class="sxs-lookup"><span data-stu-id="699ec-p114">The `parameters` property is an array of parameter objects. The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="699ec-174">Свойство</span><span class="sxs-lookup"><span data-stu-id="699ec-174">Property</span></span>  |  <span data-ttu-id="699ec-175">Тип данных</span><span class="sxs-lookup"><span data-stu-id="699ec-175">Data type</span></span>  |  <span data-ttu-id="699ec-176">Обязательное</span><span class="sxs-lookup"><span data-stu-id="699ec-176">Required</span></span>  |  <span data-ttu-id="699ec-177">Описание</span><span class="sxs-lookup"><span data-stu-id="699ec-177">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="699ec-178">string</span><span class="sxs-lookup"><span data-stu-id="699ec-178">string</span></span>  |  <span data-ttu-id="699ec-179">Нет</span><span class="sxs-lookup"><span data-stu-id="699ec-179">No</span></span> |  <span data-ttu-id="699ec-180">Описание параметра.</span><span class="sxs-lookup"><span data-stu-id="699ec-180">A description of the parameter.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="699ec-181">string</span><span class="sxs-lookup"><span data-stu-id="699ec-181">string</span></span>  |  <span data-ttu-id="699ec-182">Нет</span><span class="sxs-lookup"><span data-stu-id="699ec-182">No</span></span>  |  <span data-ttu-id="699ec-183">Должно быть **скалярным** (значение, не являющееся массивом) или **матричным** (двухмерный массив).</span><span class="sxs-lookup"><span data-stu-id="699ec-183">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="699ec-184">string</span><span class="sxs-lookup"><span data-stu-id="699ec-184">string</span></span>  |  <span data-ttu-id="699ec-185">Да</span><span class="sxs-lookup"><span data-stu-id="699ec-185">Yes</span></span>  |  <span data-ttu-id="699ec-p115">Имя параметра. Это имя отображается в IntelliSense Excel.</span><span class="sxs-lookup"><span data-stu-id="699ec-p115">The name of the parameter. This name is displayed in Excel's intelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="699ec-188">string</span><span class="sxs-lookup"><span data-stu-id="699ec-188">string</span></span>  |  <span data-ttu-id="699ec-189">Нет</span><span class="sxs-lookup"><span data-stu-id="699ec-189">No</span></span>  |  <span data-ttu-id="699ec-p116">Тип данных параметра. Должен представлять собой значение типа  **boolean**, **number** или **string**.</span><span class="sxs-lookup"><span data-stu-id="699ec-p116">The data type of the parameter. Must be **boolean**, **number**, or **string**.</span></span>  |

## <a name="result"></a><span data-ttu-id="699ec-192">result</span><span class="sxs-lookup"><span data-stu-id="699ec-192">result</span></span>

<span data-ttu-id="699ec-p117">Объект  `results` определяет тип возвращаемой функцией информации. В следующей таблице описываются свойства объекта `result`.</span><span class="sxs-lookup"><span data-stu-id="699ec-p117">The `results` object defines the type of information that is returned by the function. The following table lists the properties of the `result` object.</span></span>

|  <span data-ttu-id="699ec-195">Свойство</span><span class="sxs-lookup"><span data-stu-id="699ec-195">Property</span></span>  |  <span data-ttu-id="699ec-196">Тип данных</span><span class="sxs-lookup"><span data-stu-id="699ec-196">Data type</span></span>  |  <span data-ttu-id="699ec-197">Обязательное</span><span class="sxs-lookup"><span data-stu-id="699ec-197">Required</span></span>  |  <span data-ttu-id="699ec-198">Описание</span><span class="sxs-lookup"><span data-stu-id="699ec-198">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="699ec-199">string</span><span class="sxs-lookup"><span data-stu-id="699ec-199">string</span></span>  |  <span data-ttu-id="699ec-200">Нет</span><span class="sxs-lookup"><span data-stu-id="699ec-200">No</span></span>  |  <span data-ttu-id="699ec-201">Должно быть **скалярным** (значение, не являющееся массивом) или **матричным** (двухмерный массив).</span><span class="sxs-lookup"><span data-stu-id="699ec-201">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |
|  `type`  |  <span data-ttu-id="699ec-202">string</span><span class="sxs-lookup"><span data-stu-id="699ec-202">string</span></span>  |  <span data-ttu-id="699ec-203">Да</span><span class="sxs-lookup"><span data-stu-id="699ec-203">Yes</span></span>  |  <span data-ttu-id="699ec-p118">Тип данных параметра. Должен представлять собой значение типа  **boolean**, **number** или **string**.</span><span class="sxs-lookup"><span data-stu-id="699ec-p118">The data type of the parameter. Must be **boolean**, **number**, or **string**.</span></span>  |

## <a name="see-also"></a><span data-ttu-id="699ec-206">См. также</span><span class="sxs-lookup"><span data-stu-id="699ec-206">See also</span></span>

* [<span data-ttu-id="699ec-207">Создание настраиваемых функций в Excel</span><span class="sxs-lookup"><span data-stu-id="699ec-207">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="699ec-208">Среда выполнения для настраиваемых функций Excel</span><span class="sxs-lookup"><span data-stu-id="699ec-208">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="699ec-209">Рекомендации по настраиваемым функциям</span><span class="sxs-lookup"><span data-stu-id="699ec-209">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="699ec-210">Руководство по настраиваемым функциям Excel</span><span class="sxs-lookup"><span data-stu-id="699ec-210">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)