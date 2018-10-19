---
ms.date: 10/17/2018
description: Определение метаданных для настраиваемых функций в Excel.
title: Метаданные для настраиваемых функций в Excel
ms.openlocfilehash: cff1cbc22f39c99597d4abe7005d7b8bbce6e185
ms.sourcegitcommit: a6d6348075c1abed76d2146ddfc099b0151fe403
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/19/2018
ms.locfileid: "25640010"
---
# <a name="custom-functions-metadata-preview"></a><span data-ttu-id="5165e-103">Метаданные для настраиваемых функций (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="5165e-103">Custom functions metadata</span></span>

<span data-ttu-id="5165e-p101">При определении [настраиваемых функций](custom-functions-overview.md) в надстройке Excel в проект надстройки необходимо включить файл метаданных JSON, содержащий информацию, необходимую Excel для регистрации настраиваемых функций и предоставления пользователям доступа к ним. В этой статье описан формат JSON-файла метаданных.</span><span class="sxs-lookup"><span data-stu-id="5165e-p101">When you define [custom functions](custom-functions-overview.md) within your Excel add-in, your add-in project must include a JSON metadata file which provides the information that Excel requires to register the custom functions and make them available to end users. This article describes the format of the JSON metadata file.</span></span>

<span data-ttu-id="5165e-106">Сведения о других файлах, которые необходимо добавить в проект надстройки для включения настраиваемых функций, см. в статье [Создание настраиваемых функций в Excel](custom-functions-overview.md).</span><span class="sxs-lookup"><span data-stu-id="5165e-106">For information about the other files that you must include in your add-in project to enable custom functions, see [Create custom functions in Excel](custom-functions-overview.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="example-metadata"></a><span data-ttu-id="5165e-107">Пример метаданных</span><span class="sxs-lookup"><span data-stu-id="5165e-107">Example metadata</span></span>

<span data-ttu-id="5165e-p102">В следующем примере кода показано содержимое файла метаданных JSON для надстройки, определяющей настраиваемые функции. В следующих за этим примером разделах приводится подробная информация об отдельных свойствах, представленных в данном примере JSON.</span><span class="sxs-lookup"><span data-stu-id="5165e-p102">The following example shows the contents of a JSON metadata file for an add-in that defines custom functions. The sections that follow this example provide detailed information about the individual properties within this JSON example.</span></span>

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
> <span data-ttu-id="5165e-110">Пример готового файла JSON приводится в репозитории GitHub [OfficeDev/Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).</span><span class="sxs-lookup"><span data-stu-id="5165e-110">A complete sample JSON file is available in the [OfficeDev/Excel-Custom-Functions GitHub repository](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).</span></span>

## <a name="functions"></a><span data-ttu-id="5165e-111">functions</span><span class="sxs-lookup"><span data-stu-id="5165e-111">functions</span></span> 

<span data-ttu-id="5165e-p103">Свойство `functions` представляет собой массив объектов настраиваемых функций. В следующей таблице приведены свойства каждого объекта.</span><span class="sxs-lookup"><span data-stu-id="5165e-p103">The `functions` property is an array of custom function objects. The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="5165e-114">Свойство</span><span class="sxs-lookup"><span data-stu-id="5165e-114">Property</span></span>  |  <span data-ttu-id="5165e-115">Тип данных</span><span class="sxs-lookup"><span data-stu-id="5165e-115">Data type</span></span>  |  <span data-ttu-id="5165e-116">Обязательное</span><span class="sxs-lookup"><span data-stu-id="5165e-116">Required</span></span>  |  <span data-ttu-id="5165e-117">Описание</span><span class="sxs-lookup"><span data-stu-id="5165e-117">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="5165e-118">string</span><span class="sxs-lookup"><span data-stu-id="5165e-118">string</span></span>  |  <span data-ttu-id="5165e-119">Нет</span><span class="sxs-lookup"><span data-stu-id="5165e-119">No</span></span>  |  <span data-ttu-id="5165e-p104">Описание функции, которое пользователи видят в Excel. Например, **преобразует значение по шкале Цельсия в температуру по шкале Фаренгейта**.</span><span class="sxs-lookup"><span data-stu-id="5165e-p104">The description of the function that end users see in Excel. For example, **Converts a Celsius value to Fahrenheit**.</span></span> |
|  `helpUrl`  |  <span data-ttu-id="5165e-122">string</span><span class="sxs-lookup"><span data-stu-id="5165e-122">string</span></span>  |   <span data-ttu-id="5165e-123">Нет</span><span class="sxs-lookup"><span data-stu-id="5165e-123">No</span></span>  |  <span data-ttu-id="5165e-p105">URL-адрес, который предоставляет сведения о функции (отображается в области задач). Например, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span><span class="sxs-lookup"><span data-stu-id="5165e-p105">URL that provides information about the function. (It is displayed in a task pane.) For example, **http://contoso.com/help/convertcelsiustofahrenheit.html**.</span></span> |
| `id`     | <span data-ttu-id="5165e-126">string</span><span class="sxs-lookup"><span data-stu-id="5165e-126">string</span></span> | <span data-ttu-id="5165e-127">Да</span><span class="sxs-lookup"><span data-stu-id="5165e-127">Yes</span></span> | <span data-ttu-id="5165e-p106">Уникальный идентификатор для функции. Он может содержать только буквенно-цифровые символы и точки, а его изменение после настройки не допускается.</span><span class="sxs-lookup"><span data-stu-id="5165e-p106">A unique ID for the function. This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span> |
|  `name`  |  <span data-ttu-id="5165e-130">string</span><span class="sxs-lookup"><span data-stu-id="5165e-130">string</span></span>  |  <span data-ttu-id="5165e-131">Да</span><span class="sxs-lookup"><span data-stu-id="5165e-131">Yes</span></span>  |  <span data-ttu-id="5165e-p107">Название функции, которое пользователи видят в Excel. В Excel название этой функции будет присоединено в качестве приставки пространством имен настраиваемой функции, указанным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="5165e-p107">The name of the function that end users see in Excel. In Excel, this function name will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
|  `options`  |  <span data-ttu-id="5165e-134">object</span><span class="sxs-lookup"><span data-stu-id="5165e-134">object</span></span>  |  <span data-ttu-id="5165e-135">Нет</span><span class="sxs-lookup"><span data-stu-id="5165e-135">No</span></span>  |  <span data-ttu-id="5165e-p108">Это свойство позволяет настраивать некоторые аспекты того, как и когда Excel выполняет эту функцию. См. [Объект параметров](#options-object)  для получения дополнительной информации.</span><span class="sxs-lookup"><span data-stu-id="5165e-p108">Enables you to customize some aspects of how and when Excel executes the function. See [options object](#options-object) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="5165e-138">array</span><span class="sxs-lookup"><span data-stu-id="5165e-138">array</span></span>  |  <span data-ttu-id="5165e-139">Да</span><span class="sxs-lookup"><span data-stu-id="5165e-139">Yes</span></span>  |  <span data-ttu-id="5165e-p109">Массив, который определяет входные параметры для функции. См. [Массив параметров ](#parameters-array) для получения дополнительной информации.</span><span class="sxs-lookup"><span data-stu-id="5165e-p109">Array that defines the input parameters for the function. See [parameters array](#parameters-array)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="5165e-142">object</span><span class="sxs-lookup"><span data-stu-id="5165e-142">object</span></span>  |  <span data-ttu-id="5165e-143">Да</span><span class="sxs-lookup"><span data-stu-id="5165e-143">Yes</span></span>  |  <span data-ttu-id="5165e-p110">Объект, который определяет тип возвращаемой функцией информации. См. [Объект результата](#result-object) для получения дополнительной информации.</span><span class="sxs-lookup"><span data-stu-id="5165e-p110">Object that defines the type of information that is returned by the function. See [result object](#result-object) for details.</span></span> |

## <a name="options"></a><span data-ttu-id="5165e-146">options</span><span class="sxs-lookup"><span data-stu-id="5165e-146">options</span></span>

<span data-ttu-id="5165e-p111">Объект `options` позволяет настроить некоторые аспекты того, как и когда Excel выполняет данные функции. В следующей таблице описываются свойства объекта  `options`.</span><span class="sxs-lookup"><span data-stu-id="5165e-p111">The `options` object enables you to customize some aspects of how and when Excel executes the function. The following table lists the properties of the `options` object.</span></span>

|  <span data-ttu-id="5165e-149">Свойство</span><span class="sxs-lookup"><span data-stu-id="5165e-149">Property</span></span>  |  <span data-ttu-id="5165e-150">Тип данных</span><span class="sxs-lookup"><span data-stu-id="5165e-150">Data type</span></span>  |  <span data-ttu-id="5165e-151">Обязательное</span><span class="sxs-lookup"><span data-stu-id="5165e-151">Required</span></span>  |  <span data-ttu-id="5165e-152">Описание</span><span class="sxs-lookup"><span data-stu-id="5165e-152">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="5165e-153">boolean</span><span class="sxs-lookup"><span data-stu-id="5165e-153">boolean</span></span>  |  <span data-ttu-id="5165e-154">Нет</span><span class="sxs-lookup"><span data-stu-id="5165e-154">No</span></span><br/><br/><span data-ttu-id="5165e-155">Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="5165e-155">Default value is 4.</span></span>  |  <span data-ttu-id="5165e-p112">Если `true`, Excel будет вызывать обработчик `onCanceled` каждый раз, когда пользователь будет предпринимать действия, которые имеют эффект отмены функции, например, вручную вызывая пересчет или редактирование ячейки, на которую ссылается функция. Если вы используете эту опцию, Excel вызовет функцию JavaScript с дополнительным `caller`  параметром. (Не \*\*\* \*\*\* регистрируйте свои параметры в свойстве `parameters`). В теле функции необходимо назначить обработчик члену `caller.onCanceled`. Для получения дополнительной информации см.  [Отмена функции](custom-functions-overview.md#canceling-a-function).</span><span class="sxs-lookup"><span data-stu-id="5165e-p112">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function. If you use this option, Excel will call the JavaScript function with an additional `caller` parameter. (Do ***not*** register this parameter in the `parameters` property). In the body of the function, a handler must be assigned to the `caller.onCanceled` member. For more information, see [Canceling a function](custom-functions-overview.md#canceling-a-function).</span></span> |
|  `stream`  |  <span data-ttu-id="5165e-161">boolean</span><span class="sxs-lookup"><span data-stu-id="5165e-161">boolean</span></span>  |  <span data-ttu-id="5165e-162">Нет</span><span class="sxs-lookup"><span data-stu-id="5165e-162">No</span></span><br/><br/><span data-ttu-id="5165e-163">Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="5165e-163">Default value is 4.</span></span>  |  <span data-ttu-id="5165e-p113">Если `true`, функция может выводить значение в ячейку несколько раз, даже если была вызвана всего единожды. Этот параметр полезен для быстро изменяющихся источников данных, таких как цена акций. Если вы используете эту опцию, Excel вызовет функцию JavaScript с дополнительным параметром `caller`. (Не \*\*\* \*\*\* регистрируйте свои параметры в свойстве `parameters`). Функция должна содержать оператор `return`. Вместо этого результирующее значение передается как аргумент метода обратного вызова `caller.setResult`. Для получения дополнительной информации см. статью [Потоковые функции](custom-functions-overview.md#streaming-functions).</span><span class="sxs-lookup"><span data-stu-id="5165e-p113">If `true`, the function can output repeatedly to the cell even when invoked only once. This option is useful for rapidly-changing data sources, such as a stock price. If you use this option, Excel will call the JavaScript function with an additional `caller` parameter. (Do ***not*** register this parameter in the `parameters` property). The function should have no `return` statement. Instead, the result value is passed as the argument of the `caller.setResult` callback method. For more information, see [Streaming functions](custom-functions-overview.md#streaming-functions).</span></span> |

## <a name="parameters"></a><span data-ttu-id="5165e-171">parameters</span><span class="sxs-lookup"><span data-stu-id="5165e-171">parameters</span></span>

<span data-ttu-id="5165e-p114">Свойство `parameters`  представляет собой массив параметров объекта. В следующей таблице приводятся свойства каждого объекта.</span><span class="sxs-lookup"><span data-stu-id="5165e-p114">The `parameters` property is an array of parameter objects. The following table lists the properties of each object.</span></span>

|  <span data-ttu-id="5165e-174">Свойство</span><span class="sxs-lookup"><span data-stu-id="5165e-174">Property</span></span>  |  <span data-ttu-id="5165e-175">Тип данных</span><span class="sxs-lookup"><span data-stu-id="5165e-175">Data type</span></span>  |  <span data-ttu-id="5165e-176">Обязательное</span><span class="sxs-lookup"><span data-stu-id="5165e-176">Required</span></span>  |  <span data-ttu-id="5165e-177">Описание</span><span class="sxs-lookup"><span data-stu-id="5165e-177">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="5165e-178">string</span><span class="sxs-lookup"><span data-stu-id="5165e-178">string</span></span>  |  <span data-ttu-id="5165e-179">Нет</span><span class="sxs-lookup"><span data-stu-id="5165e-179">No</span></span> |  <span data-ttu-id="5165e-180">Описание параметра.</span><span class="sxs-lookup"><span data-stu-id="5165e-180">A description of the parameter.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="5165e-181">string</span><span class="sxs-lookup"><span data-stu-id="5165e-181">string</span></span>  |  <span data-ttu-id="5165e-182">Нет</span><span class="sxs-lookup"><span data-stu-id="5165e-182">No</span></span>  |  <span data-ttu-id="5165e-183">Должно быть **скалярным** (значение, не являющееся массивом) или **матричным** (двухмерный массив).</span><span class="sxs-lookup"><span data-stu-id="5165e-183">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span>  |
|  `name`  |  <span data-ttu-id="5165e-184">string</span><span class="sxs-lookup"><span data-stu-id="5165e-184">string</span></span>  |  <span data-ttu-id="5165e-185">Да</span><span class="sxs-lookup"><span data-stu-id="5165e-185">Yes</span></span>  |  <span data-ttu-id="5165e-p115">Имя параметра. Это имя отображается в IntelliSense Excel.</span><span class="sxs-lookup"><span data-stu-id="5165e-p115">The name of the parameter. This name is displayed in Excel's intelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="5165e-188">string</span><span class="sxs-lookup"><span data-stu-id="5165e-188">string</span></span>  |  <span data-ttu-id="5165e-189">Нет</span><span class="sxs-lookup"><span data-stu-id="5165e-189">No</span></span>  |  <span data-ttu-id="5165e-p116">Тип данных параметра. Должен представлять собой значение типа  **boolean**, **number** или **string**.</span><span class="sxs-lookup"><span data-stu-id="5165e-p116">The data type of the parameter. Must be **boolean**, **number**, or **string**.</span></span>  |

## <a name="result"></a><span data-ttu-id="5165e-192">result</span><span class="sxs-lookup"><span data-stu-id="5165e-192">result</span></span>

<span data-ttu-id="5165e-p117">Объект  `results` определяет тип возвращаемой функцией информации. В следующей таблице описываются свойства объекта `result`.</span><span class="sxs-lookup"><span data-stu-id="5165e-p117">The `results` object defines the type of information that is returned by the function. The following table lists the properties of the `result` object.</span></span>

|  <span data-ttu-id="5165e-195">Свойство</span><span class="sxs-lookup"><span data-stu-id="5165e-195">Property</span></span>  |  <span data-ttu-id="5165e-196">Тип данных</span><span class="sxs-lookup"><span data-stu-id="5165e-196">Data type</span></span>  |  <span data-ttu-id="5165e-197">Обязательное</span><span class="sxs-lookup"><span data-stu-id="5165e-197">Required</span></span>  |  <span data-ttu-id="5165e-198">Описание</span><span class="sxs-lookup"><span data-stu-id="5165e-198">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="5165e-199">string</span><span class="sxs-lookup"><span data-stu-id="5165e-199">string</span></span>  |  <span data-ttu-id="5165e-200">Нет</span><span class="sxs-lookup"><span data-stu-id="5165e-200">No</span></span>  |  <span data-ttu-id="5165e-201">Должно быть **скалярным** (значение, не являющееся массивом) или **матричным** (двухмерный массив).</span><span class="sxs-lookup"><span data-stu-id="5165e-201">Must be either **scalar** (a non-array value) or **matrix** (a 2-dimensional array).</span></span> |
|  `type`  |  <span data-ttu-id="5165e-202">string</span><span class="sxs-lookup"><span data-stu-id="5165e-202">string</span></span>  |  <span data-ttu-id="5165e-203">Да</span><span class="sxs-lookup"><span data-stu-id="5165e-203">Yes</span></span>  |  <span data-ttu-id="5165e-p118">Тип данных параметра. Должен представлять собой значение типа  **boolean**, **number** или **string**.</span><span class="sxs-lookup"><span data-stu-id="5165e-p118">The data type of the parameter. Must be **boolean**, **number**, or **string**.</span></span>  |

## <a name="see-also"></a><span data-ttu-id="5165e-206">См. также</span><span class="sxs-lookup"><span data-stu-id="5165e-206">See also</span></span>

* [<span data-ttu-id="5165e-207">Создание настраиваемых функций в Excel</span><span class="sxs-lookup"><span data-stu-id="5165e-207">Create custom functions in Excel (Preview)</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="5165e-208">Среда выполнения для настраиваемых функций Excel</span><span class="sxs-lookup"><span data-stu-id="5165e-208">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="5165e-209">Рекомендации по настраиваемым функциям</span><span class="sxs-lookup"><span data-stu-id="5165e-209">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="5165e-210">Руководство по настраиваемым функциям Excel</span><span class="sxs-lookup"><span data-stu-id="5165e-210">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)
