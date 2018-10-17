---
ms.date: 10/09/2018
description: Создание настраиваемых функций в Excel с помощью JavaScript.
title: Создание настраиваемых функций в Excel (предварительная версия)
ms.openlocfilehash: 75d6bd6db28207323027f6f4470c06e7c0ccb29b
ms.sourcegitcommit: f47654582acbe9f618bec49fb97e1d30f8701b62
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/17/2018
ms.locfileid: "25579914"
---
# <a name="create-custom-functions-in-excel-preview"></a><span data-ttu-id="00755-103">Создание настраиваемых функций в Excel (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="00755-103">Create custom functions in Excel (Preview)</span></span>

<span data-ttu-id="00755-p101">Настраиваемые функции позволяют разработчикам добавлять новые функции в Excel путем определения этих функций в JavaScript как части надстройки. Пользователи в Excel могут получать доступ к настраиваемым функциям так же, как и к любой собственной функции в Excel, например, `SUM()`. В этой статье описывается создание настраиваемых функций в Excel.</span><span class="sxs-lookup"><span data-stu-id="00755-p101">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in. Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`. This article describes how to create custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="00755-p102">Следующий рисунок демонстрирует процесс вставки настраиваемой функции в рабочий лист Excel конечным пользователем. Настраиваемая функция `CONTOSO.ADD42` предназначена для добавления 42 к паре чисел, которые пользователь указывает в качестве входных параметров для функции.</span><span class="sxs-lookup"><span data-stu-id="00755-p102">The following illustration shows an end user inserting a custom function into a cell of an Excel worksheet. The `CONTOSO.ADD42` custom function is designed to add 42 to the pair of numbers that the user specifies as input parameters to the function.</span></span>

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

<span data-ttu-id="00755-109">Следующий код определяет настраиваемую функцию `ADD42`.</span><span class="sxs-lookup"><span data-stu-id="00755-109">The following code defines the `ADD42` custom function.</span></span>

```js
function add42(a, b) {
  return a + b + 42;
}
```

> [!NOTE]
> <span data-ttu-id="00755-110">В разделе [Известные проблемы](#known-issues) далее в этой статье указаны текущие ограничения настраиваемых функций.</span><span class="sxs-lookup"><span data-stu-id="00755-110">The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.</span></span>

## <a name="components-of-a-custom-functions-add-in-project"></a><span data-ttu-id="00755-111">Компоненты проекта надстройки пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="00755-111">Components of a custom functions add-in project</span></span>

<span data-ttu-id="00755-112">Если вы используете [генератор Yo Office](https://github.com/OfficeDev/generator-office) для создания проекта надстройки настраиваемых функций Excel, вы увидите следующие файлы в проекте, который создает генератор:</span><span class="sxs-lookup"><span data-stu-id="00755-112">If you use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create an Excel custom functions add-in project, you'll see the following files in the project that the generator creates:</span></span>

| <span data-ttu-id="00755-113">Файл</span><span class="sxs-lookup"><span data-stu-id="00755-113">File</span></span> | <span data-ttu-id="00755-114">Формат файла</span><span class="sxs-lookup"><span data-stu-id="00755-114">File format</span></span> | <span data-ttu-id="00755-115">Описание</span><span class="sxs-lookup"><span data-stu-id="00755-115">Description</span></span> |
|------|-------------|-------------|
| <span data-ttu-id="00755-116">**./src/customfunctions.js**</span><span class="sxs-lookup"><span data-stu-id="00755-116">**./src/customfunctions.js**</span></span><br/><span data-ttu-id="00755-117">или</span><span class="sxs-lookup"><span data-stu-id="00755-117">or</span></span><br/><span data-ttu-id="00755-118">**./src/customfunctions.ts**</span><span class="sxs-lookup"><span data-stu-id="00755-118">**./src/customfunctions.ts**</span></span> | <span data-ttu-id="00755-119">JavaScript</span><span class="sxs-lookup"><span data-stu-id="00755-119">JavaScript</span></span><br/><span data-ttu-id="00755-120">или</span><span class="sxs-lookup"><span data-stu-id="00755-120">or</span></span><br/><span data-ttu-id="00755-121">TypeScript</span><span class="sxs-lookup"><span data-stu-id="00755-121">TypeScript</span></span> | <span data-ttu-id="00755-122">Содержит код, который определяет настраиваемые функции.</span><span class="sxs-lookup"><span data-stu-id="00755-122">Contains the code that defines custom functions.</span></span> |
| <span data-ttu-id="00755-123">**./config/customfunctions.json**</span><span class="sxs-lookup"><span data-stu-id="00755-123">**./config/customfunctions.json**</span></span> | <span data-ttu-id="00755-124">JSON</span><span class="sxs-lookup"><span data-stu-id="00755-124">JSON</span></span> | <span data-ttu-id="00755-125">Содержит метаданные, которые описывают настраиваемые функции и позволяют Excel регистрировать настраиваемые функции, чтобы сделать их доступными для пользователей.</span><span class="sxs-lookup"><span data-stu-id="00755-125">Contains metadata that describes custom functions and enables Excel to register the custom functions in order to make them available to end-users.</span></span> |
| <span data-ttu-id="00755-126">**./index.html**</span><span class="sxs-lookup"><span data-stu-id="00755-126">**./index.html**</span></span> | <span data-ttu-id="00755-127">HTML</span><span class="sxs-lookup"><span data-stu-id="00755-127">HTML</span></span> | <span data-ttu-id="00755-128">Предоставляет ссылку в тегах &lt;script&gt; на файл JavaScript, который определяет пользовательские функции.</span><span class="sxs-lookup"><span data-stu-id="00755-128">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| <span data-ttu-id="00755-129">**./manifest.xml**</span><span class="sxs-lookup"><span data-stu-id="00755-129">**Manifest.xml**</span></span> | <span data-ttu-id="00755-130">XML</span><span class="sxs-lookup"><span data-stu-id="00755-130">XML</span></span> | <span data-ttu-id="00755-131">Указывает пространство имен для всех настраиваемых функций в пределах надстройки и расположение файлов JavaScript, JSON и HTML, указанных ранее в этой таблице.</span><span class="sxs-lookup"><span data-stu-id="00755-131">Specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files that are listed previously in this table.</span></span> |

<span data-ttu-id="00755-132">Дополнительные сведения об этих файлах можно найти в следующих разделах.</span><span class="sxs-lookup"><span data-stu-id="00755-132">The following sections provide more details about these settings.</span></span>

### <a name="script-file"></a><span data-ttu-id="00755-133">Файл сценария</span><span class="sxs-lookup"><span data-stu-id="00755-133">Script file</span></span> 

<span data-ttu-id="00755-134">Файл сценария (**./src/customfunctions.js** или **./src/customfunctions.ts** в проекте, который создает генератор Yo Office) содержит код, который определяет настраиваемые функции и сопоставляется с объектами в [файле метаданных JSON](#json-metadata-file).</span><span class="sxs-lookup"><span data-stu-id="00755-134">The script file (**./src/customfunctions.js** or **./src/customfunctions.ts** in the project that the Yo Office generator creates) contains the code that defines custom functions and maps the names of the custom functions to objects in the [JSON metadata file](#json-metadata-file).</span></span> 

<span data-ttu-id="00755-p103">Например, следующий код определяет настраиваемые функции `add` и `increment`, а затем определяет информацию о сопоставлении для обеих функций. Функция `add` сопоставляется с объектом в файле метаданных JSON, где значение свойства `id` равно **ADD**, а функция `increment` сопоставляется с объектом в файле метаданных, где значение свойства `id` равно **INCREMENT**. Подробнее о сопоставлении имен функций в файле сценария с объектами в файле метаданных JSON см. [Практические рекомендации по настраиваемым функциям](custom-functions-best-practices.md#mapping-function-names-to-json-metadata).</span><span class="sxs-lookup"><span data-stu-id="00755-p103">For example, the following code defines the custom functions `add` and `increment` and then specifies mapping information for both functions. The `add` function is mapped to the object in the JSON metadata file where the value of the `id` property is **ADD**, and the `increment` function is mapped to the object in the metadata file where the value of the `id` property is **INCREMENT**. See [Custom functions best practices](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) for more information about mapping function names in the script file to objects in the JSON metadata file.</span></span>

```js
function add(first, second){
  return first + second;
}

function increment(incrementBy, callback) {
  var result = 0;
  var timer = setInterval(function() {
    result += incrementBy;
    callback.setResult(result);
  }, 1000);

  callback.onCanceled = function() {
    clearInterval(timer);
  };
}

// map `id` values in the JSON metadata file to the JavaScript function names
CustomFunctionMappings.ADD = add;
CustomFunctionMappings.INCREMENT = increment;
```

### <a name="json-metadata-file"></a><span data-ttu-id="00755-138">Файл метаданных JSON</span><span class="sxs-lookup"><span data-stu-id="00755-138">JSON metadata file</span></span> 

<span data-ttu-id="00755-p104">Файл метаданных настраиваемых функций (**./config/customfunctions.json** в проекте, создаваемом генератором Yo Office) предоставляет информацию, которую Excel требует, чтобы зарегистрировать настраиваемые функции и сделать их доступными для конечных пользователей. Настраиваемые функции регистрируются, когда пользователь запускает надстройку в первый раз. После этого они доступны для того же пользователя во всех книгах (т. е. не только в книге, в которой первоначально выполнялась надстройка).</span><span class="sxs-lookup"><span data-stu-id="00755-p104">The custom functions metadata file (**./config/customfunctions.json** in the project that the Yo Office generator creates) provides the information that Excel requires to register custom functions and make them available to end users. Custom functions are registered when a user runs an add-in for the first time. After that, they are available to that same user in all workbooks (i.e., not only in the workbook where the add-in initially ran.)</span></span>

> [!TIP]
> <span data-ttu-id="00755-142">Чтобы настраиваемые функции правильно работали в Excel Online, в параметры сервера, на котором размещен файл JSON, необходимо включить [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS).</span><span class="sxs-lookup"><span data-stu-id="00755-142">Your server settings for the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>

<span data-ttu-id="00755-p105">Следующий код в **customfunctions.json** определяет метаданные для описанных ранее функций `add` и `increment`. В таблице, следующей за данным примером кода, приведены подробные сведения об отдельных свойствах в этом объекте JSON. См. [Рекомендации по настраиваемым функциям](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) для получения более подробных сведений о задании значений для свойств `id` и `name` в файле метаданных JSON.</span><span class="sxs-lookup"><span data-stu-id="00755-p105">The following code in **customfunctions.json** specifies the metadata for the `add` function and the `increment` function that were described previously. The table that follows this code sample provides detailed information about the individual properties within this JSON object. See [Custom functions best practices](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) for more information about specifying the value of `id` and `name` properties in the JSON metadata file.</span></span>

```json
{
  "$schema": "https://developer.microsoft.com/en-us/json-schemas/office-js/custom-functions.schema.json",
  "functions": [
    {
      "id": "ADD",
      "name": "ADD",
      "description": "Add two numbers",
      "helpUrl": "http://www.contoso.com",
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
      "id": "INCREMENT",
      "name": "INCREMENT",
      "description": "Periodically increment a value",
      "helpUrl": "http://www.contoso.com",
      "result": {
          "type": "number",
          "dimensionality": "scalar"
    },
    "parameters": [
        {
            "name": "increment",
            "description": "Amount to increment",
            "type": "number",
            "dimensionality": "scalar"
        }
    ],
    "options": {
        "cancelable": true,
        "stream": true
      }
    }
  ]
}
```

<span data-ttu-id="00755-p106">В следующей таблице перечислены свойства, которые обычно присутствуют в файле метаданных JSON. Более подробные сведения о файле метаданных JSON см. в статье [Метаданные настраиваемых функций](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="00755-p106">The following table lists the properties that are typically present in the JSON metadata file. For more detailed information about the JSON metadata file, see [Custom functions metadata](custom-functions-json.md).</span></span>

| <span data-ttu-id="00755-148">Свойство</span><span class="sxs-lookup"><span data-stu-id="00755-148">Property</span></span>  | <span data-ttu-id="00755-149">Описание</span><span class="sxs-lookup"><span data-stu-id="00755-149">Description</span></span> |
|---------|---------|
| `id` | <span data-ttu-id="00755-150">Уникальный идентификатор функции.</span><span class="sxs-lookup"><span data-stu-id="00755-150">A unique ID for the function.</span></span> <span data-ttu-id="00755-151">Этот идентификатор может содержать только буквенно-цифровые символы и точки, его изменение после настройки не допускается.</span><span class="sxs-lookup"><span data-stu-id="00755-151">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span> |
| `name` | <span data-ttu-id="00755-p108">Имя функции, которое конечный пользователь видит в Excel. В Excel имя этой функции будет содержать префикс пространства имен настраиваемых функций, который указан в [XML-файле манифеста](#manifest-file).</span><span class="sxs-lookup"><span data-stu-id="00755-p108">Name of the function that the end user sees in Excel. In Excel, this function name will be prefixed by the custom functions namespace that's specified in the [XML manifest file](#manifest-file).</span></span> |
| `helpUrl` | <span data-ttu-id="00755-154">URL-адрес страницы, которая отображается, когда пользователь запрашивает справку.</span><span class="sxs-lookup"><span data-stu-id="00755-154">Url for a page that is shown when a user requests help.</span></span> |
| `description` | <span data-ttu-id="00755-p109">Описывает, что выполняет функция. Это значение появляется как подсказка, когда функция является выбранным элементом в меню автозаполнения в Excel.</span><span class="sxs-lookup"><span data-stu-id="00755-p109">Describes what the function does. This value appears as a tooltip when the function is the selected item in the autocomplete menu within Excel.</span></span> |
| `result`  | <span data-ttu-id="00755-p110">Объект, который определяет тип данных, возвращаемых функцией. Значение дочернего свойства `type` может быть **string**, **number** или **boolean**. Дочернему свойству `dimensionality` может присваиваться значение **scalar** или **matrix** (двумерный массив значений указанного типа `type`).</span><span class="sxs-lookup"><span data-stu-id="00755-p110">Object that defines the type of information that is returned by the function. The value of the `type` child property can be **string**, **number**, or **boolean**. The value of the `dimensionality` child property can be **scalar** or **matrix** (a two-dimensional array of values of the specified `type`).</span></span> |
| `parameters` | <span data-ttu-id="00755-p111">Массив, который определяет входные параметры для функции. Дочерние свойства `name` и `description` отображаются в Excel intelliSense. Значение дочернего свойства `type` может быть **string**, **number** или **boolean**. Дочернему свойству `dimensionality` может присваиваться значение **scalar** или **matrix** (двумерный массив значений указанного типа `type`).</span><span class="sxs-lookup"><span data-stu-id="00755-p111">Array that defines the input parameters for the function. The `name` and `description` child properties appear in the Excel intelliSense. The value of the `type` child property can be **string**, **number**, or **boolean**. The value of the `dimensionality` child property can be **scalar** or **matrix** (a two-dimensional array of values of the specified `type`).</span></span> |
| `options` | <span data-ttu-id="00755-p112">Позволяет настроить некоторые аспекты того, как и когда Excel выполняет эту функцию. Подробнее о том, как это свойство можно использовать, см. в разделах [Потоковые функции](#streaming-functions) и [Отмена функции](#canceling-a-function) ниже в этой статье.</span><span class="sxs-lookup"><span data-stu-id="00755-p112">Enables you to customize some aspects of how and when Excel executes the function. For more information about how this property can be used, see [Streaming functions](#streaming-functions) and [Canceling a function](#canceling-a-function) later in this article.</span></span> |

### <a name="manifest-file"></a><span data-ttu-id="00755-166">Файл манифеста</span><span class="sxs-lookup"><span data-stu-id="00755-166">Manifest file</span></span>

<span data-ttu-id="00755-p113">XML-файл манифеста для надстройки, который определяет настраиваемые функции (**./manifest.xml** в проекте, создаваемом генератором Yo Office), определяет пространство имен для всех настраиваемых функций в пределах надстройки и расположение файлов JavaScript, JSON и HTML. Ниже показан пример использования элементов `<ExtensionPoint>` и `<Resources>` в разметке XML. Эти элементы необходимо добавить в манифест надстройки для включения настраиваемых функций.</span><span class="sxs-lookup"><span data-stu-id="00755-p113">The XML manifest file for an add-in that defines custom functions (**./manifest.xml** in the project that the Yo Office generator creates) specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files. The following XML markup shows an example of the `<ExtensionPoint>` and `<Resources>` elements that you must include in an add-in's manifest to enable custom functions.</span></span>  

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
        <Host xsi:type="Workbook">
            <AllFormFactors>
                <ExtensionPoint xsi:type="CustomFunctions">
                    <Script>
                        <SourceLocation resid="JS-URL" /> <!--resid points to location of JavaScript file-->
                    </Script>
                    <Page>
                        <SourceLocation resid="HTML-URL"/> <!--resid points to location of HTML file-->
                    </Page>
                    <Metadata>
                        <SourceLocation resid="JSON-URL" /> <!--resid points to location of JSON file-->
                    </Metadata>
                    <Namespace resid="namespace" />
                </ExtensionPoint>
            </AllFormFactors>
        </Host>
    </Hosts>
    <Resources>
        <bt:Urls>
            <bt:Url id="JSON-URL" DefaultValue="http://127.0.0.1:8080/customfunctions.json" /> <!--specifies the location of your JSON file-->
            <bt:Url id="JS-URL" DefaultValue="http://127.0.0.1:8080/customfunctions.js" /> <!--specifies the location of your JavaScript file-->
            <bt:Url id="HTML-URL" DefaultValue="http://127.0.0.1:8080/index.html" /> <!--specifies the location of your HTML file-->
        </bt:Urls>
        <bt:ShortStrings>
            <bt:String id="namespace" DefaultValue="CONTOSO" /> <!--specifies the namespace that will be prepended to a function's name when it is called in Excel. Can only contain alphanumeric characters and periods.-->
        </bt:ShortStrings>
    </Resources>
</VersionOverrides>
```

> [!NOTE]
> <span data-ttu-id="00755-p114">К функциям Excel добавляется пространство имен, указанное в файле манифеста XML. Пространство имен функции предшествует имени функции и отделяется от него точкой. Например, чтобы вызвать функцию `ADD42` в ячейке листа Excel, необходимо ввести `=CONTOSO.ADD42`, так как `CONTOSO`  — это пространство имен, а `ADD42`   — это имя функции, указанной в файле JSON. Пространство имен предназначено для использования в качестве идентификатора для вашей компании или надстройки. Пространство имен может состоять только из буквенно-цифровых символов и точек.</span><span class="sxs-lookup"><span data-stu-id="00755-p114">Functions in Excel are prepended by the namespace specified in your XML manifest file. A function's namespace comes before the function name and they are separated by a period. For example, to call the function  in the cell of an Excel worksheet, you would type , because CONTOSO is the namespace and  is the name of the function specified in the JSON file. The namespace is intended to be used as an identifier for your company or the add-in.</span></span>

## <a name="functions-that-return-data-from-external-sources"></a><span data-ttu-id="00755-174">Функции, возвращающие данные из внешних источников</span><span class="sxs-lookup"><span data-stu-id="00755-174">Functions that return data from external sources</span></span>

<span data-ttu-id="00755-175">Если настраиваемая функция получает данные из внешнего источника, например веб-сайта, она должна:</span><span class="sxs-lookup"><span data-stu-id="00755-175">If a custom function retrieves data from an external source such as the web, it must:</span></span>

1. <span data-ttu-id="00755-176">возвращать обещание JavaScript в Excel.</span><span class="sxs-lookup"><span data-stu-id="00755-176">Return a JavaScript Promise to Excel.</span></span>

2. <span data-ttu-id="00755-177">разрешать Promise окончательным значением, используя функцию обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="00755-177">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="00755-p115">Пока Excel ожидает конечный результат, настраиваемые функции отображают в ячейке временный результат `#GETTING_DATA`. Во время ожидания результата пользователи могут нормально взаимодействовать с остальной частью листа.</span><span class="sxs-lookup"><span data-stu-id="00755-p115">Custom functions display a `#GETTING_DATA` temporary result in the cell while Excel waits for the final result. Users can interact normally with the rest of the worksheet while they wait for the result.</span></span>

<span data-ttu-id="00755-p116">В следующем примере кода настраиваемая функция `getTemperature()` получает от термометра текущую температуру. Обратите внимание на то, что функция `sendWebRequest` является гипотетической (не указывается здесь) и использует [XHR](custom-functions-runtime.md#xhr-example) для вызова веб-службы температуры.</span><span class="sxs-lookup"><span data-stu-id="00755-p116">In the following code sample, the `getTemperature()` custom function retrieves the current temperature of a thermometer. Note that `sendWebRequest` is a hypothetical function (not specified here) that uses [XHR](custom-functions-runtime.md#xhr-example) to call a temperature web service.</span></span>

```js
function getTemperature(thermometerID){
    return new Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

## <a name="streaming-functions"></a><span data-ttu-id="00755-182">Потоковые функции</span><span class="sxs-lookup"><span data-stu-id="00755-182">Streaming functions</span></span>

<span data-ttu-id="00755-183">Настраиваемые функции потоковой передачи позволяют вам выводить данные в ячейки многократно с течением времени, не требуя от пользователя явно запрашивать обновление данных.</span><span class="sxs-lookup"><span data-stu-id="00755-183">Streaming custom functions enable you to output data to cells repeatedly over time, without requiring a user to explicitly request data refresh.</span></span> <span data-ttu-id="00755-184">Следующий пример кода представляет собой настраиваемую функцию, которая каждую секунду добавляет число к результату.</span><span class="sxs-lookup"><span data-stu-id="00755-184">The following code sample is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="00755-185">Обратите внимание на следующие особенности этого кода:</span><span class="sxs-lookup"><span data-stu-id="00755-185">Note the following about this code:</span></span>

- <span data-ttu-id="00755-186">Excel автоматически отображает каждое новое значение при помощи обратного вызова `setResult`.</span><span class="sxs-lookup"><span data-stu-id="00755-186">Excel displays each new value automatically using the `setResult` callback.</span></span>

- <span data-ttu-id="00755-187">Второй входной параметр `handler` не отображается для конечных пользователей в Excel при выборе функции из меню автозаполнения.</span><span class="sxs-lookup"><span data-stu-id="00755-187">The second input parameter, `handler`, is not displayed to end users in Excel when they select the function from the autocomplete menu.</span></span>

- <span data-ttu-id="00755-188">Обратный вызов `onCanceled` определяет функцию, которая выполняется при отмене функции.</span><span class="sxs-lookup"><span data-stu-id="00755-188">The `onCanceled` callback defines the function that executes when the function is canceled.</span></span> <span data-ttu-id="00755-189">Необходимо реализовать обработчик отмены следующим образом для любой функции потоковой передачи.</span><span class="sxs-lookup"><span data-stu-id="00755-189">You must implement a cancellation handler like this for any streaming function.</span></span> <span data-ttu-id="00755-190">Для получения дополнительных сведений см. статью [Отмена функции](#canceling-a-function).</span><span class="sxs-lookup"><span data-stu-id="00755-190">For more information, see [Canceling a function](#canceling-a-function).</span></span>

```js
function incrementValue(increment, handler){
  var result = 0;
  setInterval(function(){
    result += increment;
    handler.setResult(result);
  }, 1000);

  handler.onCanceled = function(){
    clearInterval(timer);
  }
}
```

<span data-ttu-id="00755-191">При указании метаданных для функции потоковой передачи в файле метаданных JSON необходимо задать свойства `"cancelable": true` и `"stream": true` для объекта `options`, как показано в следующем примере.</span><span class="sxs-lookup"><span data-stu-id="00755-191">When you specify metadata for a streamed function in the JSON metadata file, you must set the properties `"cancelable": true` and `"stream": true` within the `options` object, as shown in the following example.</span></span>

```json
{
  "id": "INCREMENT",
  "name": "INCREMENT",
  "description": "Periodically increment a value",
  "helpUrl": "http://www.contoso.com",
  "result": {
    "type": "number",
    "dimensionality": "scalar"
  },
  "parameters": [
    {
      "name": "increment",
      "description": "Amount to increment",
      "type": "number",
      "dimensionality": "scalar"
    }
  ],
  "options": {
    "cancelable": true,
    "stream": true
  }
}
```

## <a name="canceling-a-function"></a><span data-ttu-id="00755-192">Отмена функции</span><span class="sxs-lookup"><span data-stu-id="00755-192">Canceling a function</span></span>

<span data-ttu-id="00755-193">В некоторых случаях может потребоваться отменить выполнение настраиваемой функции потоковой передачи, чтобы снизить потребление ею пропускной способности, рабочей памяти, а а также снизить загрузку процессора.</span><span class="sxs-lookup"><span data-stu-id="00755-193">In some situations, you may need to cancel the execution of a streaming custom function to reduce its bandwidth consumption, working memory, and CPU load.</span></span> <span data-ttu-id="00755-194">Excel отменяет выполнение функции в следующих ситуациях.</span><span class="sxs-lookup"><span data-stu-id="00755-194">Excel cancels the execution of a function in the following situations:</span></span>

- <span data-ttu-id="00755-195">Когда пользователь редактирует или удаляет ячейку, содержащую ссылку на функцию.</span><span class="sxs-lookup"><span data-stu-id="00755-195">The user edits or deletes a cell that references the function.</span></span>

- <span data-ttu-id="00755-p120">Когда изменяется один из аргументов (входных параметров) функции. В этом случае после отмены активируется новый вызов функции.</span><span class="sxs-lookup"><span data-stu-id="00755-p120">When one of the arguments (inputs) for the function changes. In this case, a new function call is triggered following the cancellation.</span></span>

- <span data-ttu-id="00755-p121">Когда пользователь запускает пересчет вручную. В этом случае после отмены активируется новый вызов функции.</span><span class="sxs-lookup"><span data-stu-id="00755-p121">When the user triggers recalculation manually. In this case, a new function call is triggered following the cancellation.</span></span>

<span data-ttu-id="00755-p122">Чтобы включить возможность отмены функции, необходимо реализовать обработчик отмены в функции JavaScript и указать свойство `"cancelable": true` в объекте `options` в метаданных JSON, которые описывают функцию. В примерах кода в предыдущем разделе данной статьи приводится пример этой техники.</span><span class="sxs-lookup"><span data-stu-id="00755-p122">To enable the ability to cancel a function, you must implement a cancellation handler within the JavaScript function and specify the property `"cancelable": true` within the `options` object in the JSON metadata that describes the function. The code samples in the previous section of this article provide an example of these techniques.</span></span>

## <a name="saving-and-sharing-state"></a><span data-ttu-id="00755-202">Сохранение и передача состояния</span><span class="sxs-lookup"><span data-stu-id="00755-202">Saving and sharing state</span></span>

<span data-ttu-id="00755-203">Пользовательские функции могут сохранять данные в глобальных переменных JavaScript, которые можно использовать в последующих вызовах.</span><span class="sxs-lookup"><span data-stu-id="00755-203">Custom functions can save data in global JavaScript variables, which can be used in subsequent calls.</span></span> <span data-ttu-id="00755-204">Сохранение состояния может быть полезно при вызове пользователями одной настраиваемой функции из нескольких ячеек, поскольку все экземпляры функции могут совместно использовать ее состояние.</span><span class="sxs-lookup"><span data-stu-id="00755-204">Saved state is useful when users add the same custom function to more than one cell, because all the instances of the function can share the state.</span></span> <span data-ttu-id="00755-205">Например, вы можете сохранить данные, возвращенные при вызове веб-ресурса, чтобы не пришлось делать дополнительные вызовы одного и того же веб-ресурса.</span><span class="sxs-lookup"><span data-stu-id="00755-205">For example, you may save the data returned from a call to a web resource to avoid making additional calls to the same web resource.</span></span>

<span data-ttu-id="00755-p124">В приведенном ниже примере кода показана реализация вышеописанной потоковой функции температуры, осуществляющей глобальное сохранение состояния. Обратите внимание на следующие особенности этого кода:</span><span class="sxs-lookup"><span data-stu-id="00755-p124">The following code sample shows an implementation of a temperature-streaming function that saves state globally. Note the following about this code:</span></span>

- <span data-ttu-id="00755-208">Функция `streamTemperature` обновляет значение температуры, которое отображается в ячейке каждую секунду, и использует в качестве источника данных переменную `savedTemperatures`.</span><span class="sxs-lookup"><span data-stu-id="00755-208">`streamTemperature` updates the temperature values displayed in the cell every second and it uses `savedTemperatures` variable as its data source.</span></span>

- <span data-ttu-id="00755-209">Так как `streamTemperature` — это функция потоковой передачи, она реализует обработчик отмены, который будет выполняться при отмене функции.</span><span class="sxs-lookup"><span data-stu-id="00755-209">Because `streamTemperature` is a streaming function, it implements a cancellation handler that will run when the function is canceled.</span></span>

- <span data-ttu-id="00755-210">Если пользователь вызывает функцию `streamTemperature` из нескольких ячеек в Excel, функция `streamTemperature` считывает данные из одной той же переменной `savedTemperatures` каждый раз, когда она запускается.</span><span class="sxs-lookup"><span data-stu-id="00755-210">If a user calls the `streamTemperature` function from multiple cells in Excel, the `streamTemperature` function reads data from the same `savedTemperatures` variable each time it runs.</span></span> 

- <span data-ttu-id="00755-211">Функция `refreshTemperature` считывает температуру определенного термометра каждую секунду и сохраняет результат в переменной `savedTemperatures`.</span><span class="sxs-lookup"><span data-stu-id="00755-211">The `refreshTemperature` function reads the temperature of a particular thermometer every second and stores the result in the `savedTemperatures` variable.</span></span> <span data-ttu-id="00755-212">Поскольку функция `refreshTemperature` не предоставляется конечным пользователям в Excel, ее не нужно регистрировать в файле JSON.</span><span class="sxs-lookup"><span data-stu-id="00755-212">Because the `refreshTemperature` function is not exposed to end users in Excel, it does not need to be registered in the JSON file.</span></span>

```js
var savedTemperatures;

function streamTemperature(thermometerID, handler){
  if(!savedTemperatures[thermometerID]){
    refreshTemperature(thermometerID); // starts fetching temperatures if the thermometer hasn't been read yet
  }

  function getNextTemperature(){
    handler.setResult(savedTemperatures[thermometerID]); // setResult sends the saved temperature value to Excel.
    var delayTime = 1000; // Amount of milliseconds to delay a request by.
    setTimeout(getNextTemperature, delayTime); // Wait 1 second before updating Excel again.

    handler.onCancelled() = function {
      clearTimeout(delayTime);
    }
  }
  getNextTemperature();
}

function refreshTemperature(thermometerID){
  sendWebRequest(thermometerID, function(data){
    savedTemperatures[thermometerID] = data.temperature;
  });
  setTimeout(function(){
    refreshTemperature(thermometerID);
  }, 1000); // Wait 1 second before reading the thermometer again, and then update the saved temperature of thermometerID.
}
```

## <a name="working-with-ranges-of-data"></a><span data-ttu-id="00755-213">Работа с диапазонами данных</span><span class="sxs-lookup"><span data-stu-id="00755-213">Working with ranges of data</span></span>

<span data-ttu-id="00755-p126">Настраиваемая функция может принимать диапазон данных в качестве входного параметра, или она может возвращать диапазон данных. В JavaScript диапазон данных представляется как двухмерный массив.</span><span class="sxs-lookup"><span data-stu-id="00755-p126">Your custom function may accept a range of data as an input parameter, or it may return a range of data. In JavaScript, a range of data is represented as a 2-dimensional array.</span></span>

<span data-ttu-id="00755-p127">Предположим, к примеру, что ваша функция возвращает второе наибольшее значение из диапазона чисел, хранящихся в Excel. Следующая функция принимает параметр `values`, который имеет тип `Excel.CustomFunctionDimensionality.matrix`. Обратите внимание, что в метаданных JSON для этой функции вы должны для параметра `type` установить значение `matrix`.</span><span class="sxs-lookup"><span data-stu-id="00755-p127">For example, suppose that your function returns the second highest value from a range of numbers stored in Excel. The following function accepts the parameter `values`, which is of type `Excel.CustomFunctionDimensionality.matrix`. Note that in the JSON metadata for this function, you would set the parameter's `type` property to `matrix`.</span></span>

```js
function secondHighest(values){
  let highest = values[0][0], secondHighest = values[0][0];
  for(var i = 0; i < values.length; i++){
    for(var j = 1; j < values[i].length; j++){
      if(values[i][j] >= highest){
        secondHighest = highest;
        highest = values[i][j];
      }
      else if(values[i][j] >= secondHighest){
        secondHighest = values[i][j];
      }
    }
  }
  return secondHighest;
}
```

## <a name="handling-errors"></a><span data-ttu-id="00755-219">Обработка ошибок</span><span class="sxs-lookup"><span data-stu-id="00755-219">handling errors</span></span>

<span data-ttu-id="00755-p128">При построении надстройки, определяющей настраиваемые функции, не забудьте добавить логику для обработки ошибок, возникающих в среде выполнения. Обработка ошибок для настраиваемых функций такая же, как и в случае [обработки ошибок для API JavaScript Excel в целом](excel-add-ins-error-handling.md). В следующем примере кода метод `.catch` будет обрабатывать все ошибки, возникающие ранее в коде.</span><span class="sxs-lookup"><span data-stu-id="00755-p128">When you build an add-in that defines custom functions, be sure to include error handling logic to account for runtime errors. Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md). In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

```js
function getComment(x) {
  let url = "https://www.contoso.com/comments/" + x;

  return fetch(url)
    .then(function (data) {
      return data.json();
    })
    .then((json) => {
      return json.body;
    })
    .catch(function (error) {
      throw error;
    })
}
```

## <a name="known-issues"></a><span data-ttu-id="00755-223">Известные проблемы</span><span class="sxs-lookup"><span data-stu-id="00755-223">Known issues</span></span>

- <span data-ttu-id="00755-224">URL-адреса справки и описания параметров пока не используются в Excel.</span><span class="sxs-lookup"><span data-stu-id="00755-224">Help URLs and parameter descriptions are not yet used by Excel.</span></span>
- <span data-ttu-id="00755-225">Настраиваемые функции в настоящее время недоступны в Excel для мобильных клиентов.</span><span class="sxs-lookup"><span data-stu-id="00755-225">Custom functions are not currently available on Excel for mobile clients.</span></span>
- <span data-ttu-id="00755-226">Изменяемые функции (которые пересчитываются автоматически при изменении несвязанных данных в электронной таблице) еще не поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="00755-226">Volatile functions (those which recalculate automatically whenever unrelated data changes in the spreadsheet) are not yet supported.</span></span>
- <span data-ttu-id="00755-227">Развертывание через портал администрирования Office 365 и AppSource еще не включено.</span><span class="sxs-lookup"><span data-stu-id="00755-227">Deployment via the Office 365 Admin Portal and AppSource are not yet enabled.</span></span>
- <span data-ttu-id="00755-p129">Настраиваемые функции в Excel Online могут перестать работать во время сеанса после периода бездействия. Для восстановления функции обновите страницу веб-обозревателя (F5) и повторно введите настраиваемую функцию.</span><span class="sxs-lookup"><span data-stu-id="00755-p129">Custom functions in Excel Online may stop working during a session after a period of inactivity. Refresh the browser page (F5) and re-enter a custom function to restore the feature.</span></span>
- <span data-ttu-id="00755-p130">Если у вас есть несколько надстроек, работающих на Excel для Windows, внутри ячейки таблицы может отображаться временный результат **#GETTING_DATA**. Закройте все окна Excel и перезапустите Excel.</span><span class="sxs-lookup"><span data-stu-id="00755-p130">You may see the **#GETTING_DATA** temporary result within the cell(s) of a worksheet if you have multiple add-ins running on Excel for Windows. Close all Excel windows and restart Excel.</span></span>
- <span data-ttu-id="00755-p131">Возможно, в будущем появятся специальные средства отладки для настраиваемых функций. Тем временем вы можете выполнить отладку в Excel Online с помощью средств разработчика F12. Подробнее см. в статье [Рекомендации по настраиваемым функциям](custom-functions-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="00755-p131">Debugging tools specifically for custom functions may be available in the future. In the meantime, you can debug on Excel Online using F12 developer tools. See more details in [Custom functions best practices](custom-functions-best-practices.md).</span></span>

## <a name="changelog"></a><span data-ttu-id="00755-235">Журнал изменений</span><span class="sxs-lookup"><span data-stu-id="00755-235">Changelog</span></span>

- <span data-ttu-id="00755-236">**7 ноября 2017 г.**. Доставлена\* предварительная версия настраиваемых функций с примерами</span><span class="sxs-lookup"><span data-stu-id="00755-236">**Nov 7, 2017**: Shipped the custom functions preview and samples</span></span>
- <span data-ttu-id="00755-237">**20 ноября 2017 года** исправлена ошибка совместимости для пользователей, использующих сборки 8801 и более новых версий</span><span class="sxs-lookup"><span data-stu-id="00755-237">**Nov 20, 2017**: Fixed compatibility bug for those using builds 8801 and later</span></span>
- <span data-ttu-id="00755-238">**28 ноября 2017 г.**. Доставлена\* поддержка отмены вызова асинхронных функций (необходимо изменение потоковых функций)</span><span class="sxs-lookup"><span data-stu-id="00755-238">**Nov 28, 2017**: Shipped support for cancellation on asynchronous functions (requires change for streaming functions)</span></span>
- <span data-ttu-id="00755-239">**7 мая 2018 г.** Реализована\*​​поддержка Mac, Excel Online и синхронных функций, выполняемых внутри процесса</span><span class="sxs-lookup"><span data-stu-id="00755-239">**May 7, 2018**: Shipped support for Mac, Excel Online, and synchronous functions running in-process</span></span>
- <span data-ttu-id="00755-p132">**20 сентября 2018 г.** Реализована поддержка среды выполнения JavaScript настраиваемых функций. Подробнее см. статью [Среда выполнения для настраиваемых функций Excel](custom-functions-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="00755-p132">**September 20, 2018**: Shipped support for custom functions JavaScript runtime. For more information, see [Runtime for Excel custom functions](custom-functions-runtime.md).</span></span>

<span data-ttu-id="00755-242">\* на канале участников программы предварительной оценки Office</span><span class="sxs-lookup"><span data-stu-id="00755-242">\* to the Office Insiders Channel</span></span>

## <a name="see-also"></a><span data-ttu-id="00755-243">См. также</span><span class="sxs-lookup"><span data-stu-id="00755-243">See also</span></span>

* [<span data-ttu-id="00755-244">Метаданные настраиваемых функций</span><span class="sxs-lookup"><span data-stu-id="00755-244">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="00755-245">Среда выполнения для настраиваемых функций Excel</span><span class="sxs-lookup"><span data-stu-id="00755-245">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="00755-246">Рекомендации по настраиваемым функциям</span><span class="sxs-lookup"><span data-stu-id="00755-246">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="00755-247">Руководство по настраиваемым функциям Excel</span><span class="sxs-lookup"><span data-stu-id="00755-247">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)