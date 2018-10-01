---
ms.date: 09/27/2018
description: Создание настраиваемой функции в Excel с помощью JavaScript.
title: Создание настраиваемых функций в Excel (предварительная версия)
ms.openlocfilehash: 98e418f843f6f5574088cea9c7393afc4a42060b
ms.sourcegitcommit: 1852ae367de53deb91d03ca55d16eb69709340d3
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/29/2018
ms.locfileid: "25348803"
---
# <a name="create-custom-functions-in-excel-preview"></a><span data-ttu-id="8df15-103">Создание настраиваемых функций в Excel (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="8df15-103">Create custom functions in Excel (Preview)</span></span>

<span data-ttu-id="8df15-p101">Настраиваемые функции позволяют разработчикам добавлять новые функции в Excel путем определения этих функций в JavaScript как части надстройки. Пользователи в Excel могут получать доступ к настраиваемым функциям так же, как и к любой собственной функции в Excel, например, `SUM()`. В этой статье описывается создание настраиваемых функций в Excel.</span><span class="sxs-lookup"><span data-stu-id="8df15-p101">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in. Users within Excel can access custom functions like any other native function in Excel (such as `SUM()`). This article describes how to create custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="8df15-107">На следующем рисунке показан конечный пользователь, вставляющий настраиваемую функцию в ячейку листа Excel.</span><span class="sxs-lookup"><span data-stu-id="8df15-107">The following illustration shows an end user inserting a custom function into a cell of an Excel worksheet.</span></span> <span data-ttu-id="8df15-108">Настраиваемая функция `CONTOSO.ADD42` предназначена для добавления 42 к паре чисел, которую пользователь указывает в качестве входных параметров для функции.</span><span class="sxs-lookup"><span data-stu-id="8df15-108">The `CONTOSO.ADD42` custom function is designed to add 42 to the pair of numbers that the user specifies as input parameters to the function.</span></span>

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

<span data-ttu-id="8df15-109">Следующий код определяет настраиваемую функцию `ADD42`.</span><span class="sxs-lookup"><span data-stu-id="8df15-109">The following code defines the `ADD42` custom function.</span></span>

```js
function add42(a, b) {
  return a + b + 42;
}
```

> [!NOTE]
> <span data-ttu-id="8df15-110">В разделе [Известные проблемы](#known-issues) далее в этой статье указаны текущие ограничения настраиваемых функций.</span><span class="sxs-lookup"><span data-stu-id="8df15-110">The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.</span></span>

## <a name="components-of-a-custom-functions-add-in-project"></a><span data-ttu-id="8df15-111">Компоненты проекта надстройки настраиваемых функций</span><span class="sxs-lookup"><span data-stu-id="8df15-111">Components of a custom functions add-in project</span></span>

<span data-ttu-id="8df15-112">Если вы используете [генератор Yo Office](https://github.com/OfficeDev/generator-office) для создания проекта надстройки настраиваемых функций Excel, вы увидите следующие файлы в проекте, который создает генератор:</span><span class="sxs-lookup"><span data-stu-id="8df15-112">If you use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create an Excel custom functions add-in project, you'll see the following files in the project that the generator creates:</span></span>

| <span data-ttu-id="8df15-113">Файл</span><span class="sxs-lookup"><span data-stu-id="8df15-113">File</span></span> | <span data-ttu-id="8df15-114">Формат файла</span><span class="sxs-lookup"><span data-stu-id="8df15-114">File format</span></span> | <span data-ttu-id="8df15-115">Описание</span><span class="sxs-lookup"><span data-stu-id="8df15-115">Description</span></span> |
|------|-------------|-------------|
| <span data-ttu-id="8df15-116">**./src/customfunctions.js**</span><span class="sxs-lookup"><span data-stu-id="8df15-116">**./src/customfunctions.js**</span></span><br/><span data-ttu-id="8df15-117">или</span><span class="sxs-lookup"><span data-stu-id="8df15-117">or</span></span><br/><span data-ttu-id="8df15-118">**./src/customfunctions.ts**</span><span class="sxs-lookup"><span data-stu-id="8df15-118">**./src/customfunctions.ts**</span></span> | <span data-ttu-id="8df15-119">JavaScript</span><span class="sxs-lookup"><span data-stu-id="8df15-119">JavaScript</span></span><br/><span data-ttu-id="8df15-120">или</span><span class="sxs-lookup"><span data-stu-id="8df15-120">or</span></span><br/><span data-ttu-id="8df15-121">TypeScript</span><span class="sxs-lookup"><span data-stu-id="8df15-121">TypeScript</span></span> | <span data-ttu-id="8df15-122">Содержит код, который определяет настраиваемые функции.</span><span class="sxs-lookup"><span data-stu-id="8df15-122">Contains the code that defines custom functions.</span></span> |
| <span data-ttu-id="8df15-123">**./config/customfunctions.json**</span><span class="sxs-lookup"><span data-stu-id="8df15-123">**./config/customfunctions.json**</span></span> | <span data-ttu-id="8df15-124">JSON</span><span class="sxs-lookup"><span data-stu-id="8df15-124">JSON</span></span> | <span data-ttu-id="8df15-125">Содержит метаданные, которые описывают настраиваемые функции и позволяют Excel регистрировать настраиваемые функции, чтобы сделать их доступными для пользователей.</span><span class="sxs-lookup"><span data-stu-id="8df15-125">Contains metadata that describes custom functions and enables Excel to register the custom functions in order to make them available to end-users.</span></span> |
| <span data-ttu-id="8df15-126">**./index.html**</span><span class="sxs-lookup"><span data-stu-id="8df15-126">**./index.html**</span></span> | <span data-ttu-id="8df15-127">HTML</span><span class="sxs-lookup"><span data-stu-id="8df15-127">HTML</span></span> | <span data-ttu-id="8df15-128">Предоставляет ссылку в тегах &lt;script&gt; на файл JavaScript, который определяет пользовательские функции.</span><span class="sxs-lookup"><span data-stu-id="8df15-128">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| <span data-ttu-id="8df15-129">**./manifest.xml**</span><span class="sxs-lookup"><span data-stu-id="8df15-129">**Manifest.xml**</span></span> | <span data-ttu-id="8df15-130">XML</span><span class="sxs-lookup"><span data-stu-id="8df15-130">XML</span></span> | <span data-ttu-id="8df15-131">Указывает пространство имен для всех настраиваемых функций в пределах надстройки и расположение файлов JavaScript, JSON и HTML, указанных ранее в этой таблице.</span><span class="sxs-lookup"><span data-stu-id="8df15-131">Specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files that are listed previously in this table.</span></span> |

<span data-ttu-id="8df15-132">Дополнительные сведения об этих файлах можно найти в следующих разделах.</span><span class="sxs-lookup"><span data-stu-id="8df15-132">The following sections provide more details about these settings.</span></span>

### <a name="script-file"></a><span data-ttu-id="8df15-133">Файл сценария</span><span class="sxs-lookup"><span data-stu-id="8df15-133">Script file</span></span> 

<span data-ttu-id="8df15-134">Файл сценария (**./src/customfunctions.js** или **./src/customfunctions.ts** в проекте, который создает генератор Yo Office) содержит код, который определяет настраиваемые функции и сопоставляется с объектами в [файле метаданных JSON](#json-metadata-file).</span><span class="sxs-lookup"><span data-stu-id="8df15-134">The script file (**./src/customfunctions.js** or **./src/customfunctions.ts** in the project that the Yo Office generator creates) contains the code that defines custom functions and maps the names of the custom functions to objects in the [JSON metadata file](#json-metadata-file).</span></span> 

<span data-ttu-id="8df15-135">Так, к примеру, в приведенном далее примере кода определяются настраиваемые функции `add` и `increment`, а затем указывается информация о сопоставлении для обеих функций.</span><span class="sxs-lookup"><span data-stu-id="8df15-135">For example, the following code defines the custom functions `add` and `increment` and then specifies mapping information for both functions.</span></span> <span data-ttu-id="8df15-136">Функция `add` сопоставляется с объектом в файле метаданных JSON, где значение свойства `id` – это **ADD**, а функция `increment` сопоставляется с объектом в файле метаданных, где значение свойства `id` – это **INCREMENT**.</span><span class="sxs-lookup"><span data-stu-id="8df15-136">The `add` function is mapped to the object in the JSON metadata file where the value of the `id` property is **ADD**, and the `increment` function is mapped to the object in the metadata file where the value of the `id` property is **INCREMENT**.</span></span> <span data-ttu-id="8df15-137">См. [Рекомендации по настраиваемым функциям](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) для получения более подробных сведений о сопоставлении имен функций в файле сценария с объектами в файле метаданных JSON.</span><span class="sxs-lookup"><span data-stu-id="8df15-137">See [Custom functions best practices](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) for more information about mapping function names in the script file to objects in the JSON metadata file.</span></span>

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

### <a name="json-metadata-file"></a><span data-ttu-id="8df15-138">Файл метаданных JSON</span><span class="sxs-lookup"><span data-stu-id="8df15-138">JSON metadata file</span></span> 

<span data-ttu-id="8df15-139">Файл метаданных настраиваемых функций (**./config/customfunctions.json** в проекте, который создает генератор Yo Office) предоставляет информацию о том, что Excel требуется зарегистрировать настраиваемые функции и сделать их доступными для конечных пользователей.</span><span class="sxs-lookup"><span data-stu-id="8df15-139">The custom functions metadata file (**./config/customfunctions.json** in the project that the Yo Office generator creates) provides the information that Excel requires to register custom functions and make them available to end users.</span></span> <span data-ttu-id="8df15-140">Настраиваемые функции регистрируются при первом запуске надстройки пользователем.</span><span class="sxs-lookup"><span data-stu-id="8df15-140">The custom functions are registered when a user runs the add-in for the first time.</span></span> <span data-ttu-id="8df15-141">После этого пользователь может использовать их во всех книгах (то есть, не только в книге, в которой первоначально выполнялась надстройка).</span><span class="sxs-lookup"><span data-stu-id="8df15-141">After that, they are available, for that same user, in all workbooks (not only the one where the add-in ran initially.)</span></span>

> [!TIP]
> <span data-ttu-id="8df15-142">Чтобы настраиваемые функции правильно работали в Excel Online, в параметры сервера, на котором размещен файл JSON, необходимо включить [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS).</span><span class="sxs-lookup"><span data-stu-id="8df15-142">Your server settings for the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>

<span data-ttu-id="8df15-143">Следующий код в **customfunctions.json** определяет метаданные для функций `add` и `increment`, описанных ранее.</span><span class="sxs-lookup"><span data-stu-id="8df15-143">The following code in **customfunctions.json** specifies the metadata for the `add` function that was described previously in this article.</span></span> <span data-ttu-id="8df15-144">В таблице, следующей за этим примером кода, содержится подробная информация об отдельных свойствах этого объекта JSON.</span><span class="sxs-lookup"><span data-stu-id="8df15-144">The table that follows this code sample provides detailed information about the individual properties within this JSON object.</span></span> <span data-ttu-id="8df15-145">См. [Рекомендации по настраиваемым функциям](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) для получения более подробных сведений о задании значений свойств `id` и `name` в файле метаданных JSON.</span><span class="sxs-lookup"><span data-stu-id="8df15-145">See [Custom functions best practices](custom-functions-best-practices.md#mapping-function-names-to-json-metadata) for more information about specifying the value of `id` and `name` properties in the JSON metadata file.</span></span>

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

<span data-ttu-id="8df15-146">В следующей таблице перечислены свойства, которые обычно присутствуют в файле метаданных JSON.</span><span class="sxs-lookup"><span data-stu-id="8df15-146">The following table lists the properties that are typically present in the JSON metadata file.</span></span> <span data-ttu-id="8df15-147">Более подробные сведения о файле метаданных JSON см. в статье [Метаданные настраиваемых функций](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="8df15-147">For more detailed information about the JSON metadata file, including options not used in the previous example, see [Custom functions metadata](custom-functions-json.md).</span></span>

| <span data-ttu-id="8df15-148">Свойство</span><span class="sxs-lookup"><span data-stu-id="8df15-148">Property</span></span>  | <span data-ttu-id="8df15-149">Описание</span><span class="sxs-lookup"><span data-stu-id="8df15-149">Description</span></span> |
|---------|---------|
| `id` | <span data-ttu-id="8df15-150">Уникальный идентификатор для функции.</span><span class="sxs-lookup"><span data-stu-id="8df15-150">A unique ID for the group.</span></span> <span data-ttu-id="8df15-151">Изменение этого идентификатора после его настройки не допускается.</span><span class="sxs-lookup"><span data-stu-id="8df15-151">This ID should not be changed after it is set.</span></span> |
| `name` | <span data-ttu-id="8df15-152">Имя функции, которое конечный пользователь видит в Excel.</span><span class="sxs-lookup"><span data-stu-id="8df15-152">Name of the function that the end user sees in Excel.</span></span> <span data-ttu-id="8df15-153">В Excel название этой функции будет иметь префикс пространства имен настраиваемых функций, [который указан в XML-файле манифеста](#manifest-file).</span><span class="sxs-lookup"><span data-stu-id="8df15-153">In the autocomplete menu, this value will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
| `helpUrl` | <span data-ttu-id="8df15-154">URL-адрес страницы, которая отображается, когда пользователь запрашивает справку.</span><span class="sxs-lookup"><span data-stu-id="8df15-154">Url for a page that is shown when a user requests help.</span></span> |
| `description` | <span data-ttu-id="8df15-155">Описывает, что делает функция.</span><span class="sxs-lookup"><span data-stu-id="8df15-155">Describes what the function does.</span></span> <span data-ttu-id="8df15-156">Это значение появляется как подсказка, когда функция является выбранным элементом в меню автозаполнения в Excel.</span><span class="sxs-lookup"><span data-stu-id="8df15-156">This value appears as a tooltip when the function is the selected item in the autocomplete menu within Excel.</span></span> |
| `result`  | <span data-ttu-id="8df15-157">Объект, который определяет тип данных, который возвращается функцией.</span><span class="sxs-lookup"><span data-stu-id="8df15-157">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="8df15-158">Значение дочернего свойства `type` может быть **string**, **number**или **boolean**.</span><span class="sxs-lookup"><span data-stu-id="8df15-158">The value of the `type` child property can be **string**, **number**, or **boolean**.</span></span> <span data-ttu-id="8df15-159">Дочернему свойству `dimensionality` может присваиваться значение **scalar** или **matrix** (двухмерный массив значений указанного типа `type`).</span><span class="sxs-lookup"><span data-stu-id="8df15-159">The `dimensionality` property can be \*\*\*\* or \*\*\*\* (a two-dimensional array of values of the specified `type`.)</span></span> |
| `parameters` | <span data-ttu-id="8df15-160">Массив, который определяет входные параметры для функции.</span><span class="sxs-lookup"><span data-stu-id="8df15-160">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="8df15-161">В Excel intelliSense появляются дочерние свойства `name` и `description`.</span><span class="sxs-lookup"><span data-stu-id="8df15-161">The `name` and `description` child properties are used in the Excel intellisense.</span></span> <span data-ttu-id="8df15-162">Значение дочернего свойства `type` может быть **string**, **number**или **boolean**.</span><span class="sxs-lookup"><span data-stu-id="8df15-162">The value of the `type` child property can be **string**, **number**, or **boolean**.</span></span> <span data-ttu-id="8df15-163">Дочернему свойству `dimensionality` может присваиваться значение **scalar** или **matrix** (двухмерный массив значений указанного типа `type`).</span><span class="sxs-lookup"><span data-stu-id="8df15-163">The `dimensionality` property can be \*\*\*\* or \*\*\*\* (a two-dimensional array of values of the specified `type`.)</span></span> |
| `options` | <span data-ttu-id="8df15-164">Это позволяет настраивать некоторые аспекты того, как и когда Excel выполняет эту функцию.</span><span class="sxs-lookup"><span data-stu-id="8df15-164">The  property enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="8df15-165">Подробнее о том, как можно использовать это свойство, см. в разделах [Потоковые функции](#streamed-functions) и [Отмена функции](#canceling-a-function) ниже в этой статье.</span><span class="sxs-lookup"><span data-stu-id="8df15-165">For more information about how this property can be used, see [Streamed functions](#streamed-functions) and [Cancellation](#canceling-a-function) later in this article.</span></span> |

### <a name="manifest-file"></a><span data-ttu-id="8df15-166">Файл манифеста</span><span class="sxs-lookup"><span data-stu-id="8df15-166">Manifest file</span></span>

<span data-ttu-id="8df15-167">XML-файл манифеста для надстройки, который определяет настраиваемые функции (**./manifest.xml** в проекте, создаваемом генератором Yo Office), определяет пространство имен для всех настраиваемых функций в пределах надстройки и расположение файлов JavaScript, JSON и HTML.</span><span class="sxs-lookup"><span data-stu-id="8df15-167">The XML manifest file for an add-in that defines custom functions specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files.</span></span> <span data-ttu-id="8df15-168">Ниже показан пример использования элементов `<ExtensionPoint>` и `<Resources>` в разметке XML. Эти элементы необходимо включить в манифест надстройки, чтобы иметь возможность выполнять настраиваемые функции.</span><span class="sxs-lookup"><span data-stu-id="8df15-168">The following XML markup shows an example of the `<ExtensionPoint>` and `<Resources>` elements that you must include in an add-in's manifest in order to enable Excel to run custom functions.</span></span>  

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
            <bt:String id="namespace" DefaultValue="CONTOSO" /> <!--specifies the namespace that will be prepended to a function's name when it is called in Excel. -->
        </bt:ShortStrings>
    </Resources>
</VersionOverrides>
```

> [!NOTE]
> <span data-ttu-id="8df15-169">Функции Excel добавляются пространством имен, указанным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="8df15-169">Functions in Excel are prepended by the namespace specified in your XML manifest file.</span></span> <span data-ttu-id="8df15-170">Пространство имен функции предшествует имени функции и отделяется от него точкой.</span><span class="sxs-lookup"><span data-stu-id="8df15-170">A function's namespace comes before the function name and they are separated by a period.</span></span> <span data-ttu-id="8df15-171">Например, чтобы вызвать функцию `ADD42` в ячейке листа Excel, следует ввести `=CONTOSO.ADD42`, так как CONTOSO — это пространство имен, а `ADD42` — имя функции, указанной в файле JSON.</span><span class="sxs-lookup"><span data-stu-id="8df15-171">For example, to call the function `ADD42` in the cell of an Excel worksheet, you would type `=CONTOSO.ADD42`, because CONTOSO is the namespace and `ADD42` is the name of the function specified in the JSON file.</span></span> <span data-ttu-id="8df15-172">Данное пространство имен используется в качестве идентификатора для вашей организации или надстройки.</span><span class="sxs-lookup"><span data-stu-id="8df15-172">The prefix is intended to be used as an identifier for your add-in.</span></span> 

## <a name="functions-that-return-data-from-external-sources"></a><span data-ttu-id="8df15-173">Функции, возвращающие данные из внешних источников</span><span class="sxs-lookup"><span data-stu-id="8df15-173">Functions that return data from external sources</span></span>

<span data-ttu-id="8df15-174">Если настраиваемая функция получает данные из внешнего источника, например веб-сайта, она должна:</span><span class="sxs-lookup"><span data-stu-id="8df15-174">If a custom function retrieves data from an external source such as the web, it must:</span></span>

1. <span data-ttu-id="8df15-175">возвращать обещание JavaScript в Excel.</span><span class="sxs-lookup"><span data-stu-id="8df15-175">Return a JavaScript Promise to Excel.</span></span>

2. <span data-ttu-id="8df15-176">Разрешать Promise окончательным значением, используя функцию обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="8df15-176">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="8df15-177">Пока Excel ожидает конечный результат, настраиваемые функции отображают в ячейке временный результат `#GETTING_DATA`.</span><span class="sxs-lookup"><span data-stu-id="8df15-177">Asynchronous functions display a `#GETTING_DATA` temporary error in the cell while Excel waits for the final result.</span></span> <span data-ttu-id="8df15-178">Во время ожидания результата пользователи могут нормально взаимодействовать с остальной частью листа.</span><span class="sxs-lookup"><span data-stu-id="8df15-178">Users can interact normally with the rest of the spreadsheet while they wait for the result.</span></span>

<span data-ttu-id="8df15-179">В следующем примере кода настраиваемая функция `getTemperature()` получает от термометра текущую температуру.</span><span class="sxs-lookup"><span data-stu-id="8df15-179">In the following code sample, the `getTemperature()` custom function retrieves the current temperature of a thermometer.</span></span> <span data-ttu-id="8df15-180">Обратите внимание на то, что функция `sendWebRequest` является гипотетической (не указывается здесь) и использует [XHR](custom-functions-runtime.md#xhr) для вызова веб-службы температуры.</span><span class="sxs-lookup"><span data-stu-id="8df15-180">Note that `sendWebRequest` is a hypothetical function, not specified here, that uses XHR to call a temperature web service.</span></span>

```js
function getTemperature(thermometerID){
    return new Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

## <a name="streamed-functions"></a><span data-ttu-id="8df15-181">Потоковые функции</span><span class="sxs-lookup"><span data-stu-id="8df15-181">Streamed functions</span></span>

<span data-ttu-id="8df15-182">Потоковые настраиваемые функции позволяют вам выводить данные в ячейки многократно с течением времени, не требуя от пользователя явно запрашивать обновление данных.</span><span class="sxs-lookup"><span data-stu-id="8df15-182">Streamed custom functions enable you to output data to cells repeatedly over time, without requiring a user to explicitly request recalculation.</span></span> <span data-ttu-id="8df15-183">Следующий пример кода представляет собой настраиваемую функцию, которая каждую секунду добавляет число к результату.</span><span class="sxs-lookup"><span data-stu-id="8df15-183">The following example is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="8df15-184">Обратите внимание на следующие особенности этого кода:</span><span class="sxs-lookup"><span data-stu-id="8df15-184">Note the following about this code:</span></span>

- <span data-ttu-id="8df15-185">Excel автоматически отображает каждое новое значение при помощи обратного вызова `setResult`.</span><span class="sxs-lookup"><span data-stu-id="8df15-185">Excel displays each new value automatically using the `setResult` callback.</span></span>

- <span data-ttu-id="8df15-186">Второй входной параметр `handler` не отображается для конечных пользователей в Excel при выборе функции из меню автозаполнения.</span><span class="sxs-lookup"><span data-stu-id="8df15-186">The second input parameter, `handler`, is not displayed to end users in Excel when they select the function from the autocomplete menu.</span></span>

- <span data-ttu-id="8df15-187">Обратный вызов `onCanceled` определяет функцию, которая выполняется при отмене функции.</span><span class="sxs-lookup"><span data-stu-id="8df15-187">The `onCanceled` callback defines the function that executes when the function is canceled.</span></span> <span data-ttu-id="8df15-188">Необходимо реализовать обработчик отмены следующим образом для любой потоковой функции.</span><span class="sxs-lookup"><span data-stu-id="8df15-188">You must implement a cancellation handler like this for any streamed function.</span></span> <span data-ttu-id="8df15-189">Для получения дополнительных сведений см. статью [Отмена функции](#canceling-a-function).</span><span class="sxs-lookup"><span data-stu-id="8df15-189">For more information, see [Canceling a function](#canceling-a-function).</span></span> 

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

<span data-ttu-id="8df15-190">При указании метаданных для потоковой функции в файле метаданных JSON необходимо задать свойства `"cancelable": true` и `"stream": true` для объекта `options`, как показано в следующем примере.</span><span class="sxs-lookup"><span data-stu-id="8df15-190">When you specify metadata for a streamed function in the JSON metadata file, you must set the properties `"cancelable": true` and `"stream": true` within the `options` object, as shown in the following example.</span></span>

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

## <a name="canceling-a-function"></a><span data-ttu-id="8df15-191">Отмена функции</span><span class="sxs-lookup"><span data-stu-id="8df15-191">Canceling a function</span></span>

<span data-ttu-id="8df15-192">В некоторых случаях может потребоваться отменить выполнение потоковой настраиваемой функции, чтобы снизить ее потребление пропускной способности, рабочей памяти и загрузку процессора.</span><span class="sxs-lookup"><span data-stu-id="8df15-192">In some situations, you may need to cancel the execution of a streamed custom function to reduce its bandwidth consumption, working memory, and CPU load.</span></span> <span data-ttu-id="8df15-193">Excel отменяет выполнение функции в следующих ситуациях.</span><span class="sxs-lookup"><span data-stu-id="8df15-193">Excel cancels the execution of a function in the following situations:</span></span>

- <span data-ttu-id="8df15-194">Когда пользователь редактирует или удаляет ячейку, содержащую ссылку на функцию.</span><span class="sxs-lookup"><span data-stu-id="8df15-194">The user edits or deletes a cell that references the function.</span></span>

- <span data-ttu-id="8df15-195">Когда изменяется один из аргументов (входных параметров) функции.</span><span class="sxs-lookup"><span data-stu-id="8df15-195">One of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="8df15-196">В этом случае после отмены активируется новый вызов функции.</span><span class="sxs-lookup"><span data-stu-id="8df15-196">In this case, a new function call is triggered in addition to the cancelation.</span></span>

- <span data-ttu-id="8df15-197">Пользователь вызывает пересчет вручную.</span><span class="sxs-lookup"><span data-stu-id="8df15-197">When the user triggers recalculation manually.</span></span> <span data-ttu-id="8df15-198">В этом случае после отмены активируется новый вызов функции.</span><span class="sxs-lookup"><span data-stu-id="8df15-198">In this case, a new function call is triggered in addition to the cancelation.</span></span>

<span data-ttu-id="8df15-199">Чтобы включить возможность отмены функции, необходимо реализовать обработчик отмены в функции JavaScript и указать свойство `"cancelable": true` в объекте `options` в метаданных JSON, которые описывают функцию.</span><span class="sxs-lookup"><span data-stu-id="8df15-199">To enable the ability to cancel a function, you must implement a cancellation handler within the JavaScript function and specify the property `"cancelable": true` within the `options` object in the JSON metadata that describes the function.</span></span> <span data-ttu-id="8df15-200">В примерах кода в предыдущем разделе данной статьи приводится пример из этих методов.</span><span class="sxs-lookup"><span data-stu-id="8df15-200">The code samples in the previous section of this article provide an example of these techniques.</span></span>

## <a name="saving-and-sharing-state"></a><span data-ttu-id="8df15-201">Сохранение и передача состояния</span><span class="sxs-lookup"><span data-stu-id="8df15-201">Saving and sharing state</span></span>

<span data-ttu-id="8df15-202">Настраиваемые функции могут сохранять данные в глобальных переменных JavaScript.</span><span class="sxs-lookup"><span data-stu-id="8df15-202">Custom functions can save data in global JavaScript variables.</span></span> <span data-ttu-id="8df15-203">При последующих вызовах настраиваемая функция может использовать значения, сохраненные в этих переменных.</span><span class="sxs-lookup"><span data-stu-id="8df15-203">In subsequent calls, your custom function may use the values saved in these variables.</span></span> <span data-ttu-id="8df15-204">Сохранение состояния может быть полезно, когда пользователи добавляют одну настраиваемую функцию к нескольким ячейкам, потому что все экземпляры функции могут совместно использовать ее состояние.</span><span class="sxs-lookup"><span data-stu-id="8df15-204">Saved state is useful when users add the same custom function to more than one cell, because all the instances of the function can share the state.</span></span> <span data-ttu-id="8df15-205">Например, вы можете сохранить данные, возвращенные при вызове веб-ресурса, чтобы не пришлось делать дополнительные вызовы одного и того же веб-ресурса.</span><span class="sxs-lookup"><span data-stu-id="8df15-205">For example, you may save the data returned from a call to a web resource to avoid making additional calls to the same web resource.</span></span>

<span data-ttu-id="8df15-206">В приведенном ниже примере кода показана реализация вышеописанной потоковой функции температуры, осуществляющей глобальное сохранение состояния.</span><span class="sxs-lookup"><span data-stu-id="8df15-206">The following code shows an implementation of the previous temperature-streaming function that saves state using the  variable.</span></span> <span data-ttu-id="8df15-207">Обратите внимание на следующие особенности этого кода:</span><span class="sxs-lookup"><span data-stu-id="8df15-207">Note the following about this code:</span></span>

- <span data-ttu-id="8df15-208">`refreshTemperature` это потоковая функция, ежесекундно считывающая температуру определенного термометра.</span><span class="sxs-lookup"><span data-stu-id="8df15-208">`refreshTemperature` is a streamed function that reads the temperature of a particular thermometer every second.</span></span> <span data-ttu-id="8df15-209">Новые температуры сохраняются в переменную `savedTemperatures`, но не обновляют значение ячейки напрямую.</span><span class="sxs-lookup"><span data-stu-id="8df15-209">New temperatures are saved in the `savedTemperatures` variable, but does not directly update the cell value.</span></span> <span data-ttu-id="8df15-210">Она не должен вызываться непосредственно из ячейки листа, *поэтому она не регистрируется в файле JSON*.</span><span class="sxs-lookup"><span data-stu-id="8df15-210">It should not be directly called from a worksheet cell, *so it is not registered in the JSON file*.</span></span>

- <span data-ttu-id="8df15-211">`streamTemperature` обновляет значения температуры, которые отображаются в ячейке каждую секунду, а в качестве источника данных использует переменную `savedTemperatures`.</span><span class="sxs-lookup"><span data-stu-id="8df15-211">`streamTemperature` updates the temperature values displayed in the cell every second and it uses `savedTemperatures` variable as its data source.</span></span> <span data-ttu-id="8df15-212">Она должна быть зарегистрирована в файле JSON и записана прописными буквами: `STREAMTEMPERATURE`.</span><span class="sxs-lookup"><span data-stu-id="8df15-212">It must be registered in the JSON file, and named with all upper-case letters, `STREAMTEMPERATURE`.</span></span>

- <span data-ttu-id="8df15-213">Пользователи могут вызывать функцию `streamTemperature` из нескольких ячеек в пользовательском интерфейсе Excel.</span><span class="sxs-lookup"><span data-stu-id="8df15-213">Users may call `streamTemperature` from several cells in the Excel UI.</span></span> <span data-ttu-id="8df15-214">Каждый вызов считывает данные из той же переменной `savedTemperatures`.</span><span class="sxs-lookup"><span data-stu-id="8df15-214">Each call reads data from the same `savedTemperatures` variable.</span></span>

```js
var savedTemperatures;

function streamTemperature(thermometerID, handler){
  if(!savedTemperatures[thermometerID]){
    refreshTemperatures(thermometerID); // starts fetching temperatures if the thermometer hasn't been read yet
  }

  function getNextTemperature(){
    handler.setResult(savedTemperatures[thermometerID]); // setResult sends the saved temperature value to Excel.
    setTimeout(getNextTemperature, 1000); // Wait 1 second before updating Excel again.
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

## <a name="working-with-ranges-of-data"></a><span data-ttu-id="8df15-215">Работа с диапазонами данных</span><span class="sxs-lookup"><span data-stu-id="8df15-215">Working with ranges of data</span></span>

<span data-ttu-id="8df15-216">Настраиваемая функция может принимать диапазон данных в качестве входного параметра, или она может возвращать диапазон данных.</span><span class="sxs-lookup"><span data-stu-id="8df15-216">Your custom function may accept a range of data as an input parameter, or it may return a range of data.</span></span> <span data-ttu-id="8df15-217">В JavaScript диапазон данных представляется как двухмерный массив.</span><span class="sxs-lookup"><span data-stu-id="8df15-217">In JavaScript, a range of data is represented as a 2-dimensional array.</span></span>

<span data-ttu-id="8df15-218">Предположим, к примеру, что ваша функция возвращает второе наибольшее значение из диапазона чисел, хранящихся в Excel.</span><span class="sxs-lookup"><span data-stu-id="8df15-218">For example, suppose that your function returns the second highest temperature from a range of temperature values stored in Excel.</span></span> <span data-ttu-id="8df15-219">Следующая функция принимает параметр `values`, который имеет тип `Excel.CustomFunctionDimensionality.matrix`.</span><span class="sxs-lookup"><span data-stu-id="8df15-219">The following function takes the parameter `values`, which is an `Excel.CustomFunctionDimensionality.matrix` parameter type.</span></span> <span data-ttu-id="8df15-220">Обратите внимание, что в JSON-метаданных для этой функции вы должны для параметра `type` установить значение `matrix`.</span><span class="sxs-lookup"><span data-stu-id="8df15-220">Note that in the registration JSON for this function, you would set the parameter's `type` property to `matrix`.</span></span>

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

## <a name="handling-errors"></a><span data-ttu-id="8df15-221">Обработка ошибок</span><span class="sxs-lookup"><span data-stu-id="8df15-221">handling errors</span></span>

<span data-ttu-id="8df15-222">При построении надстройки, определяющей настраиваемые функции, не забудьте добавить логику для обработки ошибок, возникающих в среде выполнения.</span><span class="sxs-lookup"><span data-stu-id="8df15-222">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="8df15-223">Обработка ошибок для настраиваемых функций такая же, как и в случае [обработки ошибок для API JavaScript Excel в целом](excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="8df15-223">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="8df15-224">В следующем примере кода метод `.catch` будет обрабатывать все ошибки, возникающие ранее в коде.</span><span class="sxs-lookup"><span data-stu-id="8df15-224">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

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

## <a name="known-issues"></a><span data-ttu-id="8df15-225">Известные проблемы</span><span class="sxs-lookup"><span data-stu-id="8df15-225">Known issues</span></span>

- <span data-ttu-id="8df15-226">URL-адреса справки и описания параметров пока не используются в Excel.</span><span class="sxs-lookup"><span data-stu-id="8df15-226">Help URLs and parameter descriptions are not yet used by Excel.</span></span>
- <span data-ttu-id="8df15-227">Настраиваемые функции в настоящее время недоступны в Excel для мобильных клиентов.</span><span class="sxs-lookup"><span data-stu-id="8df15-227">Custom functions are not currently available on Excel for mobile clients.</span></span>
- <span data-ttu-id="8df15-228">Изменяемые функции (которые пересчитываются автоматически при изменении несвязанных данных в электронной таблице) еще не поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="8df15-228">Volatile functions (those which recalculate automatically whenever unrelated data changes in the spreadsheet) are not yet supported.</span></span>
- <span data-ttu-id="8df15-229">Развертывание через портал администрирования Office 365 и AppSource еще не включено.</span><span class="sxs-lookup"><span data-stu-id="8df15-229">Deployment via the Office 365 Admin Portal and AppSource are not yet enabled.</span></span>
- <span data-ttu-id="8df15-230">Настраиваемые функции в Excel Online могут перестать работать во время сеанса после периода бездействия.</span><span class="sxs-lookup"><span data-stu-id="8df15-230">Custom functions in Excel Online may stop working during a session after a period of inactivity.</span></span> <span data-ttu-id="8df15-231">Для восстановления функции обновите страницу веб-обозревателя (F5) и повторно введите настраиваемую функцию.</span><span class="sxs-lookup"><span data-stu-id="8df15-231">Refresh the browser page (F5) and re-enter a custom function to restore the feature.</span></span>
- <span data-ttu-id="8df15-232">Если у вас есть несколько надстроек, работающих на Excel для Windows, внутри ячейки таблицы может отображаться временный результат **#GETTING_DATA**.</span><span class="sxs-lookup"><span data-stu-id="8df15-232">You may see the **#GETTING_DATA** temporary result within the cell(s) of a worksheet if you have multiple add-ins running on Excel for Windows.</span></span> <span data-ttu-id="8df15-233">Закройте все окна Excel и перезапустите Excel.</span><span class="sxs-lookup"><span data-stu-id="8df15-233">Close all Excel windows and restart Excel.</span></span>
- <span data-ttu-id="8df15-234">Возможно, в будущем появятся специальные средства отладки для настраиваемых функций.</span><span class="sxs-lookup"><span data-stu-id="8df15-234">Debugging tools specifically for custom functions may be available in the future.</span></span> <span data-ttu-id="8df15-235">Тем временем вы можете выполнить отладку в Excel Online с помощью средств разработчика F12.</span><span class="sxs-lookup"><span data-stu-id="8df15-235">In the meantime, you can debug on Excel Online using F12 developer tools.</span></span> <span data-ttu-id="8df15-236">Подробнее см. в статье [Рекомендации по настраиваемым функциям](custom-functions-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="8df15-236">See more details in [Custom functions best practices](custom-functions-best-practices.md).</span></span>

## <a name="changelog"></a><span data-ttu-id="8df15-237">Журнал изменений</span><span class="sxs-lookup"><span data-stu-id="8df15-237">Changelog</span></span>

- <span data-ttu-id="8df15-238">**7 ноября 2017 г.**. Доставлена\* предварительная версия настраиваемых функций с примерами</span><span class="sxs-lookup"><span data-stu-id="8df15-238">**Nov 7, 2017**: Shipped the custom functions preview and samples</span></span>
- <span data-ttu-id="8df15-239">**20 ноября 2017 года** исправлена ошибка совместимости для пользователей, использующих сборки 8801 и более новых версий</span><span class="sxs-lookup"><span data-stu-id="8df15-239">**Nov 20, 2017**: Fixed compatibility bug for those using builds 8801 and later</span></span>
- <span data-ttu-id="8df15-240">**28 ноября 2017 г.**. Доставлена\* поддержка отмены вызова асинхронных функций (необходимо изменение потоковых функций)</span><span class="sxs-lookup"><span data-stu-id="8df15-240">**Nov 28, 2017**: Shipped support for cancellation on asynchronous functions (requires change for streaming functions)</span></span>
- <span data-ttu-id="8df15-241">**7 мая 2018 г.**. Доставлена\*​​поддержка Mac, Excel Online и синхронных функций, выполняемых внутри процесса</span><span class="sxs-lookup"><span data-stu-id="8df15-241">**May 7, 2018**: Shipped support for Mac, Excel Online, and synchronous functions running in-process</span></span>
- <span data-ttu-id="8df15-242">**20 сентября 2018 г.**. Выпущена поддержка среды выполнения JavaScript настраиваемых функций.</span><span class="sxs-lookup"><span data-stu-id="8df15-242">**September 20, 2018**: Shipped support for custom functions JavaScript runtime.</span></span> <span data-ttu-id="8df15-243">Подробнее см. статью [Среда выполнения для настраиваемых функций Excel](custom-functions-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="8df15-243">For more information, see [Runtime for Excel custom functions](custom-functions-runtime.md).</span></span>

<span data-ttu-id="8df15-244">\* на канале участников программы предварительной оценки Office</span><span class="sxs-lookup"><span data-stu-id="8df15-244">\* to the Office Insiders Channel</span></span>

## <a name="see-also"></a><span data-ttu-id="8df15-245">См. также</span><span class="sxs-lookup"><span data-stu-id="8df15-245">See also</span></span>

* [<span data-ttu-id="8df15-246">Метаданные настраиваемых функций</span><span class="sxs-lookup"><span data-stu-id="8df15-246">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="8df15-247">Среда выполнения для настраиваемых функций Excel</span><span class="sxs-lookup"><span data-stu-id="8df15-247">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="8df15-248">Рекомендации по настраиваемым функциям</span><span class="sxs-lookup"><span data-stu-id="8df15-248">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="8df15-249">Руководство по настраиваемым функциям Excel</span><span class="sxs-lookup"><span data-stu-id="8df15-249">Excel custom functions tutorial</span></span>](excel-tutorial-custom-functions.md)