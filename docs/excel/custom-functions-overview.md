---
ms.date: 10/17/2018
description: Создание пользовательских функций в Excel с помощью JavaScript.
title: Создание пользовательских функций в Excel (Ознакомительная версия)
ms.openlocfilehash: 8383b5f6d568a1ce2da036fbacfb90404bbe8297
ms.sourcegitcommit: 2ac7d64bb2db75ace516a604866850fce5cb2174
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/14/2018
ms.locfileid: "26298553"
---
# <a name="create-custom-functions-in-excel-preview"></a><span data-ttu-id="4546d-103">Создание пользовательских функций в Excel (ознакомительная версия)</span><span class="sxs-lookup"><span data-stu-id="4546d-103">Create custom functions in Excel (Preview)</span></span>

<span data-ttu-id="4546d-104">Пользовательские функции позволяют разработчикам добавлять новые функции в Excel, посредством определения этих функций в JavaScript как части надстройки.</span><span class="sxs-lookup"><span data-stu-id="4546d-104">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="4546d-105">Пользователи в Excel могут получить доступ к пользовательским функциям так же, как и к любой встроенной функции в Excel, например `SUM()`.</span><span class="sxs-lookup"><span data-stu-id="4546d-105">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="4546d-106">В этой статье описано создание специальных функций в Excel.</span><span class="sxs-lookup"><span data-stu-id="4546d-106">This article explains how to create custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="4546d-107">Ниже продемонстрировано, как конечный пользователь, вставляет настраиваемую функцию в ячейке на листе Excel.</span><span class="sxs-lookup"><span data-stu-id="4546d-107">The following illustration shows an end user inserting a custom function into a cell of an Excel worksheet.</span></span> <span data-ttu-id="4546d-108">Настраиваемая функция `CONTOSO.ADD42` предназначена для добавления 42 к паре чисел, которые пользователь указывает в качестве входных параметров для функции.</span><span class="sxs-lookup"><span data-stu-id="4546d-108">The `CONTOSO.ADD42` custom function is designed to add 42 to the pair of numbers that the user specifies as input parameters to the function.</span></span>

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

<span data-ttu-id="4546d-109">Приведенный ниже код определяет настраиваемую функцию `ADD42`.</span><span class="sxs-lookup"><span data-stu-id="4546d-109">The following code defines the `ADD42` custom function.</span></span>

```js
function add42(a, b) {
  return a + b + 42;
}
```

> [!NOTE]
> <span data-ttu-id="4546d-110">В разделе [Известные проблемы](#known-issues) далее в этой статье определены текущие ограничения для пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="4546d-110">The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.</span></span>

## <a name="components-of-a-custom-functions-add-in-project"></a><span data-ttu-id="4546d-111">Компоненты пользовательские функции для надстройки проекта.</span><span class="sxs-lookup"><span data-stu-id="4546d-111">Components of a custom functions add-in project</span></span>

<span data-ttu-id="4546d-112">Если вы используете [генератор Yo Office](https://github.com/OfficeDev/generator-office) для создания в Excel проекта с пользовательскими функциями, вы увидите следующие файлы в проекте, созданном генератором:</span><span class="sxs-lookup"><span data-stu-id="4546d-112">If you use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create an Excel custom functions add-in project, you'll see the following files in the project that the generator creates:</span></span>

| <span data-ttu-id="4546d-113">Файл</span><span class="sxs-lookup"><span data-stu-id="4546d-113">File</span></span> | <span data-ttu-id="4546d-114">Формат файла</span><span class="sxs-lookup"><span data-stu-id="4546d-114">File Format</span></span> | <span data-ttu-id="4546d-115">Описание</span><span class="sxs-lookup"><span data-stu-id="4546d-115">Description</span></span> |
|------|-------------|-------------|
| <span data-ttu-id="4546d-116">**./src/customfunctions.js**</span><span class="sxs-lookup"><span data-stu-id="4546d-116">**./src/customfunctions.js**</span></span><br/><span data-ttu-id="4546d-117">или</span><span class="sxs-lookup"><span data-stu-id="4546d-117">or</span></span><br/><span data-ttu-id="4546d-118">**./src/customfunctions.ts**</span><span class="sxs-lookup"><span data-stu-id="4546d-118">**./src/customfunctions.ts**</span></span> | <span data-ttu-id="4546d-119">JavaScript</span><span class="sxs-lookup"><span data-stu-id="4546d-119">JavaScript</span></span><br/><span data-ttu-id="4546d-120">или</span><span class="sxs-lookup"><span data-stu-id="4546d-120">or</span></span><br/><span data-ttu-id="4546d-121">TypeScript</span><span class="sxs-lookup"><span data-stu-id="4546d-121">TypeScript</span></span> | <span data-ttu-id="4546d-122">Содержит код, который определяет пользовательские функции.</span><span class="sxs-lookup"><span data-stu-id="4546d-122">Contains the code that defines custom functions.</span></span> |
| <span data-ttu-id="4546d-123">**./config/customfunctions.json**</span><span class="sxs-lookup"><span data-stu-id="4546d-123">**./config/customfunctions.json**</span></span> | <span data-ttu-id="4546d-124">JSON</span><span class="sxs-lookup"><span data-stu-id="4546d-124">JSON</span></span> | <span data-ttu-id="4546d-125">Содержит метаданные с описанием пользовательских функций и позволяет Excel регистрировать пользовательские функции и сделать их доступными для конечных пользователей.</span><span class="sxs-lookup"><span data-stu-id="4546d-125">Contains metadata that describes custom functions and enables Excel to register the custom functions and make them available to end users.</span></span> |
| <span data-ttu-id="4546d-126">**./index.html**</span><span class="sxs-lookup"><span data-stu-id="4546d-126">**./index.html**</span></span> | <span data-ttu-id="4546d-127">HTML</span><span class="sxs-lookup"><span data-stu-id="4546d-127">HTML</span></span> | <span data-ttu-id="4546d-128">Предоставляет &lt;скрипт&gt; со ссылкой на файл JavaScript, который определяет пользовательские функции.</span><span class="sxs-lookup"><span data-stu-id="4546d-128">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| <span data-ttu-id="4546d-129">**./manifest.xml**</span><span class="sxs-lookup"><span data-stu-id="4546d-129">**Manifest.xml**</span></span> | <span data-ttu-id="4546d-130">XML</span><span class="sxs-lookup"><span data-stu-id="4546d-130">XML</span></span> | <span data-ttu-id="4546d-131">Определяет пространство имен для всех пользовательских функций в надстройку и расположение JavaScript, JSON и HTML-файлов, которые указаны ранее в этой таблице.</span><span class="sxs-lookup"><span data-stu-id="4546d-131">Specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files that are listed previously in this table.</span></span> |

<span data-ttu-id="4546d-132">В разделах ниже приведены дополнительные сведения о данных файлах.</span><span class="sxs-lookup"><span data-stu-id="4546d-132">The following sections provide more information about these files.</span></span>

### <a name="script-file"></a><span data-ttu-id="4546d-133">Файл скрипта</span><span class="sxs-lookup"><span data-stu-id="4546d-133">Script file</span></span> 

<span data-ttu-id="4546d-134">Файл сценария (**./src/customfunctions.js** или **./src/customfunctions.ts** в проекте, созданном генератором Yo Office) содержит код, который определяет пользовательские функции и размещает имена пользовательских функций к объектам в [файле метаданных JSON](#json-metadata-file).</span><span class="sxs-lookup"><span data-stu-id="4546d-134">The script file (**./src/customfunctions.js** or **./src/customfunctions.ts** in the project that the Yo Office generator creates) contains the code that defines custom functions and maps the names of the custom functions to objects in the [JSON metadata file](#json-metadata-file).</span></span> 

<span data-ttu-id="4546d-135">Например, приведенный ниже код определяет пользовательские функции `add` и `increment`, а затем указывают информация о сопоставлении для обоих функций.</span><span class="sxs-lookup"><span data-stu-id="4546d-135">For example, the following code defines the custom functions `add` and `increment` and then specifies mapping information for both functions.</span></span> <span data-ttu-id="4546d-136">Функция `add` будет сопоставлена с объектом в файле метаданных JSON, где значение свойства `id` **ADD**, и функция `increment` будет сопоставлена с объектом в файле метаданных, где значение свойства`id` **INCREMENT**.</span><span class="sxs-lookup"><span data-stu-id="4546d-136">The `add` function is mapped to the object in the JSON metadata file where the value of the `id` property is **ADD**, and the `increment` function is mapped to the object in the metadata file where the value of the `id` property is **INCREMENT**.</span></span> <span data-ttu-id="4546d-137">См. статью [Советы и рекомендации по работе с пользовательскими функциями](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) для получения дополнительных данных о сопоставление имен функций в файле скрипта с объектами в файле метаданных JSON.</span><span class="sxs-lookup"><span data-stu-id="4546d-137">See [Custom functions best practices](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) for more information about mapping function names in the script file to objects in the JSON metadata file.</span></span>

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

### <a name="json-metadata-file"></a><span data-ttu-id="4546d-138">Файл метаданных JSON</span><span class="sxs-lookup"><span data-stu-id="4546d-138">JSON metadata file</span></span> 

<span data-ttu-id="4546d-139">Файл метаданных пользовательских функций (**./config/customfunctions.json** в проекте, созданном во время генератора Yo Office) предоставляет информацию, которая необходима Excel для регистрации пользовательских функций и обеспечения их доступности для конечных пользователей.</span><span class="sxs-lookup"><span data-stu-id="4546d-139">The custom functions metadata file (**./config/customfunctions.json** in the project that the Yo Office generator creates) provides the information that Excel requires to register custom functions and make them available to end users.</span></span> <span data-ttu-id="4546d-140">Пользовательские функции регистрируются, когда пользователь запускает надстройку в первый раз.</span><span class="sxs-lookup"><span data-stu-id="4546d-140">Custom functions are registered when a user runs an add-in for the first time.</span></span> <span data-ttu-id="4546d-141">После этого как они становятся доступны тому самому пользователю во всех рабочих книгах (т.е. не только в рабочей книге, где надстройка первоначально запущена).</span><span class="sxs-lookup"><span data-stu-id="4546d-141">After that, they are available to that same user in all workbooks (i.e., not only in the workbook where the add-in initially ran.)</span></span>

> [!TIP]
> <span data-ttu-id="4546d-142">Настройки сервера на сервере, на котором размещен JSON-файл, должны включать активацию [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS), чтобы пользовательские функции сработали надлежащим образом в Excel Online.</span><span class="sxs-lookup"><span data-stu-id="4546d-142">Server settings on the server that hosts the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>

<span data-ttu-id="4546d-143">Код ниже в **customfunctions.json** определяет метаданные для функции `add` и функции `increment`, описанные ранее.</span><span class="sxs-lookup"><span data-stu-id="4546d-143">The following code in **customfunctions.json** specifies the metadata for the `add` function and the `increment` function that were described previously.</span></span> <span data-ttu-id="4546d-144">Таблица, которая следует за этим примером кода, предоставляет подробные сведения об отдельных свойств для этого объекта JSON.</span><span class="sxs-lookup"><span data-stu-id="4546d-144">The table that follows this code sample provides detailed information about the individual properties within this JSON object.</span></span> <span data-ttu-id="4546d-145">См. статью [Советы и рекомендации по работе с пользовательскими функциями](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) для получения дополнительных данных об указании имен свойств `id` и `name` в файле метаданных JSON.</span><span class="sxs-lookup"><span data-stu-id="4546d-145">See [Custom functions best practices](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) for more information about specifying the value of `id` and `name` properties in the JSON metadata file.</span></span>

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

<span data-ttu-id="4546d-146">В таблице ниже перечислены свойства, которые обычно есть в файле метаданных JSON.</span><span class="sxs-lookup"><span data-stu-id="4546d-146">The following table lists the properties that are typically present in the JSON metadata file.</span></span> <span data-ttu-id="4546d-147">Дополнительные сведения о файле метаданных JSON см. в статье [Пользовательские функции метаданных](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="4546d-147">For more detailed information about the JSON metadata file, see [Custom functions metadata](custom-functions-json.md).</span></span>

| <span data-ttu-id="4546d-148">Свойство</span><span class="sxs-lookup"><span data-stu-id="4546d-148">Property</span></span>  | <span data-ttu-id="4546d-149">Описание</span><span class="sxs-lookup"><span data-stu-id="4546d-149">Description</span></span> |
|---------|---------|
| `id` | <span data-ttu-id="4546d-150">Уникальный идентификатор для функции.</span><span class="sxs-lookup"><span data-stu-id="4546d-150">A unique ID for the group.</span></span> <span data-ttu-id="4546d-151">Этот идентификатор может содержать только буквы, цифры и точки и не может изменяться после настройки.</span><span class="sxs-lookup"><span data-stu-id="4546d-151">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span> |
| `name` | <span data-ttu-id="4546d-152">Имя функции, которая будет отображаться пользователю в Excel.</span><span class="sxs-lookup"><span data-stu-id="4546d-152">Name of the function that the end user sees in Excel.</span></span> <span data-ttu-id="4546d-153">В Excel это имя функции будет включать префикс пространства имен пользовательских функций, который указан в [XML файле манифеста](#manifest-file).</span><span class="sxs-lookup"><span data-stu-id="4546d-153">In Excel, this function name will be prefixed by the custom functions namespace that's specified in the [XML manifest file](#manifest-file).</span></span> |
| `helpUrl` | <span data-ttu-id="4546d-154">URL-адрес страницы, который отображается при запросе пользователем справки.</span><span class="sxs-lookup"><span data-stu-id="4546d-154">URL for the page that is shown when a user requests help.</span></span> |
| `description` | <span data-ttu-id="4546d-155">Описание того, что делает функция.</span><span class="sxs-lookup"><span data-stu-id="4546d-155">Describes what the function does.</span></span> <span data-ttu-id="4546d-156">Это значение отображается в виде подсказки, когда функция представляет собой выделенный элемент в меню автозаполнения в Excel.</span><span class="sxs-lookup"><span data-stu-id="4546d-156">This value appears as a tooltip when the function is the selected item in the autocomplete menu within Excel.</span></span> |
| `result`  | <span data-ttu-id="4546d-157">Объект, который определяет тип информации, возвращаемый функцией.</span><span class="sxs-lookup"><span data-stu-id="4546d-157">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="4546d-158">Для получения более подробной информации об этом объекте см. [результат](custom-functions-json.md#result).</span><span class="sxs-lookup"><span data-stu-id="4546d-158">For detailed information about this object, see [result](custom-functions-json.md#result).</span></span> |
| `parameters` | <span data-ttu-id="4546d-159">Массив, который определяет входные параметры для функции.</span><span class="sxs-lookup"><span data-stu-id="4546d-159">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="4546d-160">Для получения более подробной информации об этом объекте см. [параметры](custom-functions-json.md#parameters).</span><span class="sxs-lookup"><span data-stu-id="4546d-160">For detailed information about this object, see [parameters](custom-functions-json.md#parameters).</span></span> |
| `options` | <span data-ttu-id="4546d-161">Позволяет настроить некоторые аспекты того, как и когда Excel выполняет функцию.</span><span class="sxs-lookup"><span data-stu-id="4546d-161">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="4546d-162">Дополнительные сведения о способах использования этого свойства см. разделы [Потоковая передача функции](#streaming-functions) и [Отмена функция](#canceling-a-function) ниже в этой статье.</span><span class="sxs-lookup"><span data-stu-id="4546d-162">For more information about how this property can be used, see [Streaming functions](#streaming-functions) and [Canceling a function](#canceling-a-function) later in this article.</span></span> |

### <a name="manifest-file"></a><span data-ttu-id="4546d-163">Файл манифеста</span><span class="sxs-lookup"><span data-stu-id="4546d-163">Manifest file</span></span>

<span data-ttu-id="4546d-164">XML-файл манифеста для надстройки, который определяет пользовательские функции (**./manifest.xml** в проекте, который создает генератор Yo Office) и определяет пространство имен для всех пользовательских функций в надстройке, а также расположение файлов JavaScript, JSON и HTML.</span><span class="sxs-lookup"><span data-stu-id="4546d-164">The XML manifest file for an add-in that defines custom functions (**./manifest.xml** in the project that the Yo Office generator creates) specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files.</span></span> <span data-ttu-id="4546d-165">XML-разметка ниже представляет пример элементов `<ExtensionPoint>` и `<Resources>`, которые необходимо включить в манифест надстройки, чтобы активировать пользовательские функции.</span><span class="sxs-lookup"><span data-stu-id="4546d-165">The following XML markup shows an example of the `<ExtensionPoint>` and `<Resources>` elements that you must include in an add-in's manifest to enable custom functions.</span></span>  

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
> <span data-ttu-id="4546d-166">Функции в Excel имеют в начале пространство имен, указанное в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="4546d-166">Functions in Excel are prepended by the namespace specified in your XML manifest file.</span></span> <span data-ttu-id="4546d-167">Пространство имен функции предшествует названию функции, и они будут разделены точкой.</span><span class="sxs-lookup"><span data-stu-id="4546d-167">A function's namespace comes before the function name and they are separated by a period.</span></span> <span data-ttu-id="4546d-168">Например, чтобы вызвать функцию `ADD42` в ячейке на листе Excel, введите `=CONTOSO.ADD42`, так как `CONTOSO` является пространством имен, а `ADD42` — это имя функции, определяемой в JSON-файл.</span><span class="sxs-lookup"><span data-stu-id="4546d-168">For example, to call the function `ADD42` in the cell of an Excel worksheet, you would type `=CONTOSO.ADD42`, because `CONTOSO` is the namespace and `ADD42` is the name of the function specified in the JSON file.</span></span> <span data-ttu-id="4546d-169">Пространство имен служит в качестве идентификатора для вашей компании или надстройки.</span><span class="sxs-lookup"><span data-stu-id="4546d-169">The prefix is intended to be used as an identifier for your add-in.</span></span> <span data-ttu-id="4546d-170">Пространство имен может содержать только буквы, цифры и точки.</span><span class="sxs-lookup"><span data-stu-id="4546d-170">A namespace can only contain alphanumeric characters and periods.</span></span>

## <a name="functions-that-return-data-from-external-sources"></a><span data-ttu-id="4546d-171">Функции, которые возвращают данные из внешних источников</span><span class="sxs-lookup"><span data-stu-id="4546d-171">Functions that return data from external sources</span></span>

<span data-ttu-id="4546d-172">Если пользовательская функция извлекает данные из внешнего источника, например, сайта, она должна:</span><span class="sxs-lookup"><span data-stu-id="4546d-172">If a custom function retrieves data from an external source such as the web, it must:</span></span>

1. <span data-ttu-id="4546d-173">Возвращать обещание JavaScript в Excel;</span><span class="sxs-lookup"><span data-stu-id="4546d-173">Return a JavaScript Promise to Excel.</span></span>

2. <span data-ttu-id="4546d-174">Устранять обещание с итоговым значением с помощью функции обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="4546d-174">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="4546d-175">Пользовательские функции отображают `#GETTING_DATA` временный результат в ячейке, пока Excel ожидает конечный результат.</span><span class="sxs-lookup"><span data-stu-id="4546d-175">Custom functions display a `#GETTING_DATA` temporary result in the cell while Excel waits for the final result.</span></span> <span data-ttu-id="4546d-176">Пользователи могут нормально взаимодействовать с остальным листом, хотя они ожидают результат.</span><span class="sxs-lookup"><span data-stu-id="4546d-176">Users can interact normally with the rest of the worksheet while they wait for the result.</span></span>

<span data-ttu-id="4546d-177">В приведенном ниже примере кода пользовательская функция `getTemperature()` возвращает текущую температуру термометра.</span><span class="sxs-lookup"><span data-stu-id="4546d-177">In the following code sample, the `getTemperature()` custom function retrieves the current temperature of a thermometer.</span></span> <span data-ttu-id="4546d-178">Обратите внимание, что `sendWebRequest` — это гипотетическая функция (не указанная ниже), которая использует [XHR](custom-functions-runtime.md#xhr-example) для вызова веб-службы.</span><span class="sxs-lookup"><span data-stu-id="4546d-178">Note that `sendWebRequest` is a hypothetical function (not specified here) that uses [XHR](custom-functions-runtime.md#xhr-example) to call a temperature web service.</span></span>

```js
function getTemperature(thermometerID){
    return new Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

## <a name="streaming-functions"></a><span data-ttu-id="4546d-179">Потоковая передача функций</span><span class="sxs-lookup"><span data-stu-id="4546d-179">Streaming functions</span></span>

<span data-ttu-id="4546d-180">Потоковая передача пользовательских функций позволяет выводить данные в ячейки несколько раз в течением времени, избавляя пользователя от необходимости явным образом запрашивать обновление данных.</span><span class="sxs-lookup"><span data-stu-id="4546d-180">Streaming custom functions enable you to output data to cells repeatedly over time, without requiring a user to explicitly request data refresh.</span></span> <span data-ttu-id="4546d-181">Приведенный ниже пример кода — это настраиваемая функция, которая добавляет число к результату каждую секунду.</span><span class="sxs-lookup"><span data-stu-id="4546d-181">The following code sample is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="4546d-182">Обратите внимание на следующие особенности этого кода:</span><span class="sxs-lookup"><span data-stu-id="4546d-182">Note the following about this code:</span></span>

- <span data-ttu-id="4546d-183">Excel отображает каждое новое значением автоматически с помощью обратного вызова `setResult`.</span><span class="sxs-lookup"><span data-stu-id="4546d-183">Excel displays each new value automatically using the `setResult` callback.</span></span>

- <span data-ttu-id="4546d-184">Второй параметр ввода, `handler`, не отображается для конечных пользователей в Excel, когда они выбирают функцию в меню "Автозаполнение".</span><span class="sxs-lookup"><span data-stu-id="4546d-184">The second input parameter, `handler`, is not displayed to end users in Excel when they select the function from the autocomplete menu.</span></span>

- <span data-ttu-id="4546d-185">Обратный вызов `onCanceled` определяет функцию, которая выполняется при отмене функции.</span><span class="sxs-lookup"><span data-stu-id="4546d-185">The `onCanceled` callback defines the function that executes when the function is canceled.</span></span> <span data-ttu-id="4546d-186">Вам необходимо реализовать уведомление об отмене следующим образом для любой функции потоковой передачи.</span><span class="sxs-lookup"><span data-stu-id="4546d-186">You must implement a cancellation handler like this for any streaming function.</span></span> <span data-ttu-id="4546d-187">Дополнительные сведения см. в статье [Отмена функции](#canceling-a-function).</span><span class="sxs-lookup"><span data-stu-id="4546d-187">For more information, see [Canceling a function](#canceling-a-function).</span></span>

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

<span data-ttu-id="4546d-188">Когда вы указываете метаданные для функции потоковой передачи в файле метаданных JSON, необходимо задать свойства `"cancelable": true` и `"stream": true` в объекте `options`, как показано в приведенном ниже примере.</span><span class="sxs-lookup"><span data-stu-id="4546d-188">When you specify metadata for a streaming function in the JSON metadata file, you must set the properties `"cancelable": true` and `"stream": true` within the `options` object, as shown in the following example.</span></span>

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

## <a name="canceling-a-function"></a><span data-ttu-id="4546d-189">Отмена функции</span><span class="sxs-lookup"><span data-stu-id="4546d-189">Canceling a function</span></span>

<span data-ttu-id="4546d-190">В некоторых случаях может потребоваться отмена выполнения пользовательских функций потоковой передачи, чтобы уменьшить использования пропускной способности, рабочей памяти и загрузку ЦП.</span><span class="sxs-lookup"><span data-stu-id="4546d-190">In some situations, you may need to cancel the execution of a streaming custom function to reduce its bandwidth consumption, working memory, and CPU load.</span></span> <span data-ttu-id="4546d-191">Excel отменяет выполнение функций в следующих случаях:</span><span class="sxs-lookup"><span data-stu-id="4546d-191">Excel cancels the execution of a function in the following situations:</span></span>

- <span data-ttu-id="4546d-192">Когда пользователь редактирует или удаляет ячейку, ссылающуюся на функцию.</span><span class="sxs-lookup"><span data-stu-id="4546d-192">The user edits or deletes a cell that references the function.</span></span>

- <span data-ttu-id="4546d-193">Когда изменяется один из аргументов (входных параметров) функции.</span><span class="sxs-lookup"><span data-stu-id="4546d-193">One of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="4546d-194">В этом случае после отмены выполняется новый вызов функции.</span><span class="sxs-lookup"><span data-stu-id="4546d-194">In this case, a new function call is triggered in addition to the cancelation.</span></span>

- <span data-ttu-id="4546d-195">Когда пользователь вручную вызывает пересчет.</span><span class="sxs-lookup"><span data-stu-id="4546d-195">When the user triggers recalculation manually.</span></span> <span data-ttu-id="4546d-196">В этом случае после отмены выполняется новый вызов функции.</span><span class="sxs-lookup"><span data-stu-id="4546d-196">In this case, a new function call is triggered in addition to the cancelation.</span></span>

<span data-ttu-id="4546d-197">Чтобы активировать возможность отмены функции, необходимо реализовать обработчик отмены в функции JavaScript, а также указать свойство `"cancelable": true` в объекте `options` в метаданных JSON, который описывает функцию.</span><span class="sxs-lookup"><span data-stu-id="4546d-197">To enable the ability to cancel a function, you must implement a cancellation handler within the JavaScript function and specify the property `"cancelable": true` within the `options` object in the JSON metadata that describes the function.</span></span> <span data-ttu-id="4546d-198">Примеры кода в предыдущем разделе этой статьи предоставляют собой пример использования данных техник.</span><span class="sxs-lookup"><span data-stu-id="4546d-198">The code samples in the previous section of this article provide an example of these techniques.</span></span>

## <a name="saving-and-sharing-state"></a><span data-ttu-id="4546d-199">Состояние сохранения и совместного использования</span><span class="sxs-lookup"><span data-stu-id="4546d-199">Saving and sharing state</span></span>

<span data-ttu-id="4546d-200">Пользовательские функции могут сохранять данные в глобальных переменных JavaScript, которые можно использовать в последующих вызовах.</span><span class="sxs-lookup"><span data-stu-id="4546d-200">Custom functions can save data in global JavaScript variables, which can be used in subsequent calls.</span></span> <span data-ttu-id="4546d-201">Сохраненное состояние полезно, когда пользователи вызывают одни и те же настраиваемые функций из более чем одной ячейки, так как все экземпляры функции могут получить доступ к состоянию.</span><span class="sxs-lookup"><span data-stu-id="4546d-201">Saved state is useful when users call the same custom function from more than one cell, because all instances of the function can access the state.</span></span> <span data-ttu-id="4546d-202">Например, вы можете сохранить данные, возвращенные при вызове веб-ресурса, чтобы не пришлось обеспечивать выполнение дополнительных вызовов.</span><span class="sxs-lookup"><span data-stu-id="4546d-202">For example, you may save the data returned from a call to a web resource to avoid making additional calls to the same web resource.</span></span>

<span data-ttu-id="4546d-203">В приведенном ниже примере кода показана реализация вышеописанной функции передачи температуры, сохраняющей состояние с помощью глобальной переменной.</span><span class="sxs-lookup"><span data-stu-id="4546d-203">The following code shows an implementation of the previous temperature-streaming function that saves state using the  variable.</span></span> <span data-ttu-id="4546d-204">Обратите внимание на следующие особенности этого кода:</span><span class="sxs-lookup"><span data-stu-id="4546d-204">Note the following about this code:</span></span>

- <span data-ttu-id="4546d-205">Функция `streamTemperature` обновляет значение температуры, которое отображается в ячейке, каждую секунду и использует переменную `savedTemperatures` как источник данных.</span><span class="sxs-lookup"><span data-stu-id="4546d-205">The `streamTemperature` function updates the temperature value that's displayed in the cell every second and it uses the `savedTemperatures` variable as its data source.</span></span>

- <span data-ttu-id="4546d-206">Так как `streamTemperature` — это функция потоковой передачи, она реализует обработчик отмены, который будет запускаться, если функция была отменена.</span><span class="sxs-lookup"><span data-stu-id="4546d-206">Because `streamTemperature` is a streaming function, it implements a cancellation handler that will run when the function is canceled.</span></span>

- <span data-ttu-id="4546d-207">Если пользователь вызывает функцию `streamTemperature` из нескольких ячеек в Excel, функция `streamTemperature` считывает данные из той же самой переменной `savedTemperatures` при каждом запуске.</span><span class="sxs-lookup"><span data-stu-id="4546d-207">If a user calls the `streamTemperature` function from multiple cells in Excel, the `streamTemperature` function reads data from the same `savedTemperatures` variable each time it runs.</span></span> 

- <span data-ttu-id="4546d-208">Функция `refreshTemperature` ежесекундно считывает температуру определенного термометра и сохраняет результат в переменной `savedTemperatures`.</span><span class="sxs-lookup"><span data-stu-id="4546d-208">The `refreshTemperature` function reads the temperature of a particular thermometer every second and stores the result in the `savedTemperatures` variable.</span></span> <span data-ttu-id="4546d-209">Так как функция `refreshTemperature` недоступна для конечных пользователей в Excel, ее не нужно регистрировать в JSON-файле.</span><span class="sxs-lookup"><span data-stu-id="4546d-209">Because the `refreshTemperature` function is not exposed to end users in Excel, it does not need to be registered in the JSON file.</span></span>

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

## <a name="working-with-ranges-of-data"></a><span data-ttu-id="4546d-210">Работа с диапазонами данных</span><span class="sxs-lookup"><span data-stu-id="4546d-210">Working with ranges of data</span></span>

<span data-ttu-id="4546d-211">Ваша пользовательская функция может принимать широкий диапазон данных в виде входных параметров или возвращать широкий диапазон данных.</span><span class="sxs-lookup"><span data-stu-id="4546d-211">Your custom function may accept a range of data as an input parameter, or it may return a range of data.</span></span> <span data-ttu-id="4546d-212">В JavaScript диапазон данных будет иметь вид двумерного массива.</span><span class="sxs-lookup"><span data-stu-id="4546d-212">In JavaScript, a range of data is represented as a 2-dimensional array.</span></span>

<span data-ttu-id="4546d-213">Например, предположим, что функция возвращает второе по величине значение из диапазона значений, хранящихся в Excel.</span><span class="sxs-lookup"><span data-stu-id="4546d-213">For example, suppose that your function returns the second highest temperature from a range of temperature values stored in Excel.</span></span> <span data-ttu-id="4546d-214">Приведенная ниже функция принимает параметр `values`, относящийся к типу `Excel.CustomFunctionDimensionality.matrix`.</span><span class="sxs-lookup"><span data-stu-id="4546d-214">The following function takes the parameter `values`, which is an `Excel.CustomFunctionDimensionality.matrix` parameter type.</span></span> <span data-ttu-id="4546d-215">Обратите внимание, что в метаданных JSON для данной функции вам следует задать для параметра свойство `type` в `matrix`.</span><span class="sxs-lookup"><span data-stu-id="4546d-215">Note that in the JSON metadata for this function, you would set the parameter's `type` property to `matrix`.</span></span>

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

## <a name="handling-errors"></a><span data-ttu-id="4546d-216">Обработка ошибок</span><span class="sxs-lookup"><span data-stu-id="4546d-216">handling errors</span></span>

<span data-ttu-id="4546d-217">При создании надстройки, которая определяет пользовательские функции, не забудьте включить логику для обработки ошибок, возникающих в среде выполнения.</span><span class="sxs-lookup"><span data-stu-id="4546d-217">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="4546d-218">Обработка ошибок для пользовательских функций совпадает с [обработкой ошибок для Excel JavaScript API ошибок в значительной степени](excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="4546d-218">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="4546d-219">В следующем примере кода `.catch` будет обрабатывать любые ошибки, возникающие ранее в коде.</span><span class="sxs-lookup"><span data-stu-id="4546d-219">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

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

## <a name="known-issues"></a><span data-ttu-id="4546d-220">Известные проблемы</span><span class="sxs-lookup"><span data-stu-id="4546d-220">Known issues</span></span>

- <span data-ttu-id="4546d-221">URL-адреса справки и описания параметров в настоящее время не используются Excel.</span><span class="sxs-lookup"><span data-stu-id="4546d-221">Help URLs and parameter descriptions are not yet used by Excel.</span></span>
- <span data-ttu-id="4546d-222">Пользовательские функции в настоящее время недоступны в Excel для мобильных клиентов.</span><span class="sxs-lookup"><span data-stu-id="4546d-222">Custom functions are not currently available on Excel for mobile clients.</span></span>
- <span data-ttu-id="4546d-223">Переменные функции (которые пересчитываются автоматически всякий раз при изменениях несвязанных данных на листе) еще не поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="4546d-223">Volatile functions (those that recalculate automatically whenever unrelated data changes in the spreadsheet) are not yet supported.</span></span>
- <span data-ttu-id="4546d-224">Развертывание через портал администрирования Office 365 и AppSource еще не активировано.</span><span class="sxs-lookup"><span data-stu-id="4546d-224">Deployment via the Office 365 Admin Portal and AppSource are not yet enabled.</span></span>
- <span data-ttu-id="4546d-225">Пользовательские функции в Excel Online могут перестать работать во время сеанса после периода бездействия.</span><span class="sxs-lookup"><span data-stu-id="4546d-225">Custom functions in Excel Online may stop working during a session after a period of inactivity.</span></span> <span data-ttu-id="4546d-226">Обновите страницу браузера (F5) и еще раз введите пользовательскую функции для восстановления работоспособности.</span><span class="sxs-lookup"><span data-stu-id="4546d-226">Refresh the browser page (F5) and re-enter a custom function to restore the feature.</span></span>
- <span data-ttu-id="4546d-227">Вы можете увидеть временный результат **#GETTING_DATA** (# ОЖИДАНИЕ_ДАННЫХ) внутри ячейки(-ек), листа, если у вас есть несколько надстроек, запущенных в Excel для Windows.</span><span class="sxs-lookup"><span data-stu-id="4546d-227">You may see the **#GETTING_DATA** temporary result within the cell(s) of a worksheet if you have multiple add-ins running on Excel for Windows.</span></span> <span data-ttu-id="4546d-228">Закройте все окна Excel и перезапустите Excel.</span><span class="sxs-lookup"><span data-stu-id="4546d-228">Close all Excel windows and restart Excel.</span></span>
- <span data-ttu-id="4546d-229">Инструменты для отладки, предназначенные специально для пользовательских функций, могут быть доступны в будущем.</span><span class="sxs-lookup"><span data-stu-id="4546d-229">Debugging tools specifically for custom functions may be available in the future.</span></span> <span data-ttu-id="4546d-230">В настоящее время вы можете выполнить отладку в Excel Online при использовании средств разработчика F12.</span><span class="sxs-lookup"><span data-stu-id="4546d-230">In the meantime, you can debug on Excel Online using F12 developer tools.</span></span> <span data-ttu-id="4546d-231">Дополнительные данные см. [Советы и рекомендации в отношении пользовательских функций](custom-functions-best-practices.md)</span><span class="sxs-lookup"><span data-stu-id="4546d-231">See more details in [Custom functions best practices](custom-functions-best-practices.md).</span></span>

## <a name="changelog"></a><span data-ttu-id="4546d-232">Журнал изменений</span><span class="sxs-lookup"><span data-stu-id="4546d-232">Changelog</span></span>

- <span data-ttu-id="4546d-233">**7 ноября 2017 г.**: Выпущена ознакомительная версия пользовательских функций с примерами.</span><span class="sxs-lookup"><span data-stu-id="4546d-233">**Nov 7, 2017**: Shipped the custom functions preview and samples</span></span>
- <span data-ttu-id="4546d-234">**20 ноября 2017 г.**: Исправлена ошибка совместимости для пользователей, использующих сборки 8801 и выше.</span><span class="sxs-lookup"><span data-stu-id="4546d-234">**Nov 20, 2017**: Fixed compatibility bug for those using builds 8801 and later</span></span>
- <span data-ttu-id="4546d-235">**28 ноября 2017 г.**: Добавлена поддержка отмены вызова асинхронных функций (необходимо изменение для потоковых функций).</span><span class="sxs-lookup"><span data-stu-id="4546d-235">**Nov 28, 2017**: Shipped support for cancellation on asynchronous functions (requires change for streaming functions)</span></span>
- <span data-ttu-id="4546d-236">**7 мая 2018 г.**: Реализована\* поддержка запущенный подпроцессов для Mac, Excel Online и синхронных функций</span><span class="sxs-lookup"><span data-stu-id="4546d-236">**May 7, 2018**: Shipped\* support for Mac, Excel Online, and synchronous functions running in-process</span></span>
- <span data-ttu-id="4546d-237">**20 сентября 2018 г.**: Реализована поддержка пользовательских функций среды выполнения JavaScript.</span><span class="sxs-lookup"><span data-stu-id="4546d-237">**September 20, 2018**: Shipped support for custom functions JavaScript runtime.</span></span> <span data-ttu-id="4546d-238">Дополнительные сведения см. в статье [Среда выполнения для пользовательских функций Excel](custom-functions-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="4546d-238">For more information, see [Runtime for Excel custom functions](custom-functions-runtime.md).</span></span>
- <span data-ttu-id="4546d-239">**20 октября 2018 г.**: После выхода [Сборки October Insiders](https://support.office.com/ru-RU/article/what-s-new-for-office-insiders-c152d1e2-96ff-4ce9-8c14-e74e13847a24), пользовательские функции требуют параметр «идентификатор» в [метаданных пользовательских функций](custom-functions-json.md) для настольных версий Windows и Online.</span><span class="sxs-lookup"><span data-stu-id="4546d-239">**October 20, 2018**: With the [October Insiders build](https://support.office.com/ru-RU/article/what-s-new-for-office-insiders-c152d1e2-96ff-4ce9-8c14-e74e13847a24), Custom Functions now requires the 'id' parameter in your [custom functions metadata](custom-functions-json.md) for Windows Desktop and Online.</span></span> <span data-ttu-id="4546d-240">На компьютерах Mac можно игнорировать этот параметр.</span><span class="sxs-lookup"><span data-stu-id="4546d-240">On Mac, this parameter should be ignored.</span></span>


<span data-ttu-id="4546d-241">\* к каналу [Office Insider ](https://products.office.com/office-insider) (ранее "Предварительная оценка — ранний доступ")</span><span class="sxs-lookup"><span data-stu-id="4546d-241">\* to the [Office Insider](https://products.office.com/office-insider) channel (formerly called "Insider Fast")</span></span>

## <a name="see-also"></a><span data-ttu-id="4546d-242">См. также</span><span class="sxs-lookup"><span data-stu-id="4546d-242">See also</span></span>

* [<span data-ttu-id="4546d-243">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="4546d-243">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="4546d-244">Среда выполнения для пользовательских функций Excel</span><span class="sxs-lookup"><span data-stu-id="4546d-244">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="4546d-245">Советы и рекомендации в отношении пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="4546d-245">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="4546d-246">Руководство по пользовательским функциям в Excel</span><span class="sxs-lookup"><span data-stu-id="4546d-246">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)
