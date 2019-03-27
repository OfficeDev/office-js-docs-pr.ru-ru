---
ms.date: 03/19/2019
description: Создание пользовательских функций в Excel с помощью JavaScript.
title: Создание пользовательских функций в Excel (ознакомительная версия)
localization_priority: Priority
ms.openlocfilehash: ac3410267da415c4d567092da2e653fcffd10b72
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870452"
---
# <a name="create-custom-functions-in-excel-preview"></a><span data-ttu-id="c5e6c-103">Создание пользовательских функций в Excel (ознакомительная версия)</span><span class="sxs-lookup"><span data-stu-id="c5e6c-103">Create custom functions in Excel (preview)</span></span>

<span data-ttu-id="c5e6c-104">Пользовательские функции позволяют разработчикам добавлять новые функции в Excel, посредством определения этих функций в JavaScript как части надстройки.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-104">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="c5e6c-105">Пользователи в Excel могут получить доступ к пользовательским функциям так же, как и к любой встроенной функции в Excel, например `SUM()`.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-105">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="c5e6c-106">В этой статье описано создание специальных функций в Excel.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-106">This article describes how to create custom functions in Excel.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="c5e6c-107">Ниже продемонстрировано, как конечный пользователь, вставляет настраиваемую функцию в ячейке на листе Excel.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-107">The following illustration shows an end user inserting a custom function into a cell of an Excel worksheet.</span></span> <span data-ttu-id="c5e6c-108">Настраиваемая функция `CONTOSO.ADD42` предназначена для добавления 42 к паре чисел, которые пользователь указывает в качестве входных параметров для функции.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-108">The `CONTOSO.ADD42` custom function is designed to add 42 to the pair of numbers that the user specifies as input parameters to the function.</span></span>

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

<span data-ttu-id="c5e6c-109">Приведенный ниже код определяет настраиваемую функцию `ADD42`.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-109">The following code defines the `ADD42` custom function.</span></span>

```js
function add42(a, b) {
  return a + b + 42;
}
```

> [!NOTE]
> <span data-ttu-id="c5e6c-110">В разделе [Известные проблемы](#known-issues) далее в этой статье определены текущие ограничения для пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-110">The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.</span></span>

## <a name="components-of-a-custom-functions-add-in-project"></a><span data-ttu-id="c5e6c-111">Компоненты пользовательские функции для надстройки проекта.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-111">Components of a custom functions add-in project</span></span>

<span data-ttu-id="c5e6c-112">Если вы используете [генератор Yo Office](https://github.com/OfficeDev/generator-office) для создания в Excel проекта с пользовательскими функциями, вы увидите следующие файлы в проекте, созданном генератором:</span><span class="sxs-lookup"><span data-stu-id="c5e6c-112">If you use the [Yo Office generator](https://github.com/OfficeDev/generator-office) to create an Excel custom functions add-in project, you'll see the following files in the project that the generator creates:</span></span>

| <span data-ttu-id="c5e6c-113">Файл</span><span class="sxs-lookup"><span data-stu-id="c5e6c-113">File</span></span> | <span data-ttu-id="c5e6c-114">Формат файла</span><span class="sxs-lookup"><span data-stu-id="c5e6c-114">File format</span></span> | <span data-ttu-id="c5e6c-115">Описание</span><span class="sxs-lookup"><span data-stu-id="c5e6c-115">Description</span></span> |
|------|-------------|-------------|
| <span data-ttu-id="c5e6c-116">**./src/customfunctions.js**</span><span class="sxs-lookup"><span data-stu-id="c5e6c-116">**./src/customfunctions.js**</span></span><br/><span data-ttu-id="c5e6c-117">или</span><span class="sxs-lookup"><span data-stu-id="c5e6c-117">or</span></span><br/><span data-ttu-id="c5e6c-118">**./src/customfunctions.ts**</span><span class="sxs-lookup"><span data-stu-id="c5e6c-118">**./src/customfunctions.ts**</span></span> | <span data-ttu-id="c5e6c-119">JavaScript</span><span class="sxs-lookup"><span data-stu-id="c5e6c-119">JavaScript</span></span><br/><span data-ttu-id="c5e6c-120">или</span><span class="sxs-lookup"><span data-stu-id="c5e6c-120">or</span></span><br/><span data-ttu-id="c5e6c-121">TypeScript</span><span class="sxs-lookup"><span data-stu-id="c5e6c-121">TypeScript</span></span> | <span data-ttu-id="c5e6c-122">Содержит код, который определяет пользовательские функции.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-122">Contains the code that defines custom functions.</span></span> |
| <span data-ttu-id="c5e6c-123">**./config/customfunctions.json**</span><span class="sxs-lookup"><span data-stu-id="c5e6c-123">**./config/customfunctions.json**</span></span> | <span data-ttu-id="c5e6c-124">JSON</span><span class="sxs-lookup"><span data-stu-id="c5e6c-124">JSON</span></span> | <span data-ttu-id="c5e6c-125">Содержит метаданные с описанием пользовательских функций и позволяет Excel регистрировать пользовательские функции и сделать их доступными для конечных пользователей.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-125">Contains metadata that describes custom functions and enables Excel to register the custom functions and make them available to end users.</span></span> |
| <span data-ttu-id="c5e6c-126">**./index.html**</span><span class="sxs-lookup"><span data-stu-id="c5e6c-126">**./index.html**</span></span> | <span data-ttu-id="c5e6c-127">HTML</span><span class="sxs-lookup"><span data-stu-id="c5e6c-127">HTML</span></span> | <span data-ttu-id="c5e6c-128">Предоставляет &lt;скрипт&gt; со ссылкой на файл JavaScript, который определяет пользовательские функции.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-128">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| <span data-ttu-id="c5e6c-129">**./manifest.xml**</span><span class="sxs-lookup"><span data-stu-id="c5e6c-129">**./manifest.xml**</span></span> | <span data-ttu-id="c5e6c-130">XML</span><span class="sxs-lookup"><span data-stu-id="c5e6c-130">XML</span></span> | <span data-ttu-id="c5e6c-131">Определяет пространство имен для всех пользовательских функций в надстройку и расположение JavaScript, JSON и HTML-файлов, которые указаны ранее в этой таблице.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-131">Specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files that are listed previously in this table.</span></span> |

<span data-ttu-id="c5e6c-132">В разделах ниже приведены дополнительные сведения о данных файлах.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-132">The following sections provide more information about these files.</span></span>

### <a name="script-file"></a><span data-ttu-id="c5e6c-133">Файл скрипта</span><span class="sxs-lookup"><span data-stu-id="c5e6c-133">Script file</span></span>

<span data-ttu-id="c5e6c-134">Файл сценария (**./src/customfunctions.js** или **./src/customfunctions.ts** в проекте, созданном генератором Yo Office) содержит код, который определяет пользовательские функции и размещает имена пользовательских функций к объектам в [файле метаданных JSON](#json-metadata-file).</span><span class="sxs-lookup"><span data-stu-id="c5e6c-134">The script file (**./src/customfunctions.js** or **./src/customfunctions.ts** in the project that the Yo Office generator creates) contains the code that defines custom functions and maps the names of the custom functions to objects in the [JSON metadata file](#json-metadata-file).</span></span> 

<span data-ttu-id="c5e6c-135">Например, приведенный ниже код определяет пользовательские функции `add` и `increment`, а затем указывают информацию о сопоставлении для обеих функций.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-135">For example, the following code defines the custom functions `add` and `increment` and then specifies association information for both functions.</span></span> <span data-ttu-id="c5e6c-136">Функция `add` сопоставляется с объектом в файле метаданных JSON, где значение свойства `id` **ADD**, и функция `increment` будет сопоставляться с объектом в файле метаданных, где значение свойства`id` **INCREMENT**.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-136">The `add` function is associated with the object in the JSON metadata file where the value of the `id` property is **ADD**, and the `increment` function is associated with the object in the metadata file where the value of the `id` property is **INCREMENT**.</span></span> <span data-ttu-id="c5e6c-137">См. статью [Советы и рекомендации по работе с пользовательскими функциями](custom-functions-best-practices.md#associating-function-names-with-json-metadata) для получения дополнительных данных о сопоставлении имен функций в файле скрипта с объектами в файле метаданных JSON.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-137">See [Custom functions best practices](custom-functions-best-practices.md#associating-function-names-with-json-metadata) for more information about associating function names in the script file to objects in the JSON metadata file.</span></span>

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

// associate `id` values in the JSON metadata file to the JavaScript function names
 CustomFunctions.associate("ADD", add);
 CustomFunctions.associate("INCREMENT", increment);
```

### <a name="json-metadata-file"></a><span data-ttu-id="c5e6c-138">Файл метаданных JSON</span><span class="sxs-lookup"><span data-stu-id="c5e6c-138">JSON metadata file</span></span>

<span data-ttu-id="c5e6c-139">Файл метаданных пользовательских функций (**./config/customfunctions.json** в проекте, созданном во время генератора Yo Office) предоставляет информацию, которая необходима Excel для регистрации пользовательских функций и обеспечения их доступности для конечных пользователей.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-139">The custom functions metadata file (**./config/customfunctions.json** in the project that the Yo Office generator creates) provides the information that Excel requires to register custom functions and make them available to end users.</span></span> <span data-ttu-id="c5e6c-140">Пользовательские функции регистрируются, когда пользователь запускает надстройку в первый раз.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-140">Custom functions are registered when a user runs an add-in for the first time.</span></span> <span data-ttu-id="c5e6c-141">После этого как они становятся доступны тому самому пользователю во всех рабочих книгах (т.е. не только в рабочей книге, где надстройка первоначально запущена).</span><span class="sxs-lookup"><span data-stu-id="c5e6c-141">After that, they are available to that same user in all workbooks (i.e., not only in the workbook where the add-in initially ran.)</span></span>

> [!TIP]
> <span data-ttu-id="c5e6c-142">Настройки сервера на сервере, на котором размещен JSON-файл, должны включать активацию [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS), чтобы пользовательские функции сработали надлежащим образом в Excel Online.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-142">Server settings on the server that hosts the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>

<span data-ttu-id="c5e6c-143">Код ниже в **customfunctions.json** определяет метаданные для функции `add` и функции `increment`, описанные ранее.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-143">The following code in **customfunctions.json** specifies the metadata for the `add` function and the `increment` function that were described previously.</span></span> <span data-ttu-id="c5e6c-144">Таблица, которая следует за этим примером кода, предоставляет подробные сведения об отдельных свойств для этого объекта JSON.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-144">The table that follows this code sample provides detailed information about the individual properties within this JSON object.</span></span> <span data-ttu-id="c5e6c-145">См. статью [Советы и рекомендации по работе с пользовательскими функциями](custom-functions-best-practices.md#associating-function-names-with-json-metadata) для получения дополнительных данных об указании имен свойств `id` и `name` в файле метаданных JSON.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-145">See [Custom functions best practices](custom-functions-best-practices.md#associating-function-names-with-json-metadata) for more information about specifying the value of `id` and `name` properties in the JSON metadata file.</span></span>

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

<span data-ttu-id="c5e6c-146">В таблице ниже перечислены свойства, которые обычно есть в файле метаданных JSON.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-146">The following table lists the properties that are typically present in the JSON metadata file.</span></span> <span data-ttu-id="c5e6c-147">Дополнительные сведения о файле метаданных JSON см. в статье [Пользовательские функции метаданных](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="c5e6c-147">For more detailed information about the JSON metadata file, see [Custom functions metadata](custom-functions-json.md).</span></span>

| <span data-ttu-id="c5e6c-148">Свойство</span><span class="sxs-lookup"><span data-stu-id="c5e6c-148">Property</span></span>  | <span data-ttu-id="c5e6c-149">Описание</span><span class="sxs-lookup"><span data-stu-id="c5e6c-149">Description</span></span> |
|---------|---------|
| `id` | <span data-ttu-id="c5e6c-150">Уникальный идентификатор для функции.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-150">A unique ID for the function.</span></span> <span data-ttu-id="c5e6c-151">Этот идентификатор может содержать только буквы, цифры и точки и не может изменяться после настройки.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-151">This ID can only contain alphanumeric characters and periods and should not be changed after it is set.</span></span> |
| `name` | <span data-ttu-id="c5e6c-152">Имя функции, которая будет отображаться пользователю в Excel.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-152">Name of the function that the end user sees in Excel.</span></span> <span data-ttu-id="c5e6c-153">В Excel это имя функции будет включать префикс пространства имен пользовательских функций, который указан в [XML файле манифеста](#manifest-file).</span><span class="sxs-lookup"><span data-stu-id="c5e6c-153">In Excel, this function name will be prefixed by the custom functions namespace that's specified in the [XML manifest file](#manifest-file).</span></span> |
| `helpUrl` | <span data-ttu-id="c5e6c-154">URL-адрес страницы, который отображается при запросе пользователем справки.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-154">URL for the page that is shown when a user requests help.</span></span> |
| `description` | <span data-ttu-id="c5e6c-155">Описание того, что делает функция.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-155">Describes what the function does.</span></span> <span data-ttu-id="c5e6c-156">Это значение отображается в виде подсказки, когда функция представляет собой выделенный элемент в меню автозаполнения в Excel.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-156">This value appears as a tooltip when the function is the selected item in the autocomplete menu within Excel.</span></span> |
| `result`  | <span data-ttu-id="c5e6c-157">Объект, который определяет тип информации, возвращаемый функцией.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-157">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="c5e6c-158">Для получения более подробной информации об этом объекте см. [результат](custom-functions-json.md#result).</span><span class="sxs-lookup"><span data-stu-id="c5e6c-158">For detailed information about this object, see [result](custom-functions-json.md#result).</span></span> |
| `parameters` | <span data-ttu-id="c5e6c-159">Массив, который определяет входные параметры для функции.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-159">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="c5e6c-160">Для получения более подробной информации об этом объекте см. [параметры](custom-functions-json.md#parameters).</span><span class="sxs-lookup"><span data-stu-id="c5e6c-160">For detailed information about this object, see [parameters](custom-functions-json.md#parameters).</span></span> |
| `options` | <span data-ttu-id="c5e6c-161">Позволяет настроить некоторые аспекты того, как и когда Excel выполняет функцию.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-161">Enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="c5e6c-162">Дополнительные сведения о способах использования этого свойства см. в разделах [Потоковая передача функций](#streaming-functions) и [Отмена функции](#canceling-a-function).</span><span class="sxs-lookup"><span data-stu-id="c5e6c-162">For more information about how this property can be used, see [Streaming functions](#streaming-functions) and [canceling a function](#canceling-a-function).</span></span> |

### <a name="manifest-file"></a><span data-ttu-id="c5e6c-163">Файл манифеста</span><span class="sxs-lookup"><span data-stu-id="c5e6c-163">Manifest file</span></span>

<span data-ttu-id="c5e6c-164">XML-файл манифеста для надстройки, который определяет пользовательские функции (**./manifest.xml** в проекте, который создает генератор Yo Office) и определяет пространство имен для всех пользовательских функций в надстройке, а также расположение файлов JavaScript, JSON и HTML.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-164">The XML manifest file for an add-in that defines custom functions (**./manifest.xml** in the project that the Yo Office generator creates) specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files.</span></span> <span data-ttu-id="c5e6c-165">XML-разметка ниже представляет пример элементов `<ExtensionPoint>` и `<Resources>`, которые необходимо включить в манифест надстройки, чтобы активировать пользовательские функции.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-165">The following XML markup shows an example of the `<ExtensionPoint>` and `<Resources>` elements that you must include in an add-in's manifest to enable custom functions.</span></span>  

```xml
<?xml version="1.0" encoding="UTF-8" standalone="yes"?>
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0" xmlns:ov="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="TaskPaneApp">
  <Id>6f4e46e8-07a8-4644-b126-547d5b539ece</Id>
  <Version>1.0.0.0</Version>
  <ProviderName>Contoso</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="helloworld"/>
  <Description DefaultValue="Samples to test custom functions"/>
  <Hosts>
    <Host Name="Workbook"/>
  </Hosts>
  <DefaultSettings>
    <SourceLocation DefaultValue="https://localhost:8081/index.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Workbook">
        <AllFormFactors>
          <ExtensionPoint xsi:type="CustomFunctions">
            <Script>
              <SourceLocation resid="JS-URL"/>
            </Script>
            <Page>
              <SourceLocation resid="HTML-URL"/>
            </Page>
            <Metadata>
              <SourceLocation resid="JSON-URL"/>
            </Metadata>
            <Namespace resid="namespace"/>
          </ExtensionPoint>
        </AllFormFactors>
      </Host>
    </Hosts>
    <Resources>
      <bt:Urls>
        <bt:Url id="JSON-URL" DefaultValue="https://localhost:8081/config/customfunctions.json"/>
        <bt:Url id="JS-URL" DefaultValue="https://localhost:8081/dist/win32/ship/index.win32.bundle"/>
        <bt:Url id="HTML-URL" DefaultValue="https://localhost:8081/index.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="namespace" DefaultValue="CONTOSO"/>
      </bt:ShortStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
```

> [!NOTE]
> <span data-ttu-id="c5e6c-166">Функции в Excel имеют в начале пространство имен, указанное в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-166">Functions in Excel are prepended by the namespace specified in your XML manifest file.</span></span> <span data-ttu-id="c5e6c-167">Пространство имен функции предшествует названию функции, и они будут разделены точкой.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-167">A function's namespace comes before the function name and they are separated by a period.</span></span> <span data-ttu-id="c5e6c-168">Например, чтобы вызвать функцию `ADD42` в ячейке на листе Excel, введите `=CONTOSO.ADD42`, так как `CONTOSO` является пространством имен, а `ADD42` — это имя функции, определяемой в JSON-файл.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-168">For example, to call the function `ADD42` in the cell of an Excel worksheet, you would type `=CONTOSO.ADD42`, because `CONTOSO` is the namespace and `ADD42` is the name of the function specified in the JSON file.</span></span> <span data-ttu-id="c5e6c-169">Пространство имен служит в качестве идентификатора для вашей компании или надстройки.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-169">The namespace is intended to be used as an identifier for your company or the add-in.</span></span> <span data-ttu-id="c5e6c-170">Пространство имен может содержать только буквы, цифры и точки.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-170">A namespace can only contain alphanumeric characters and periods.</span></span>

## <a name="declaring-a-volatile-function"></a><span data-ttu-id="c5e6c-171">Объявление переменной функции</span><span class="sxs-lookup"><span data-stu-id="c5e6c-171">Declaring a volatile function</span></span>

<span data-ttu-id="c5e6c-172">[Переменные функции](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions) — это функции, значение которых периодически изменяется, даже если никакой из аргументов функции не меняется.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-172">[Volatile functions](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions) are functions in which the value changes from moment to moment, even if none of the function's arguments have changed.</span></span> <span data-ttu-id="c5e6c-173">Эти функции пересчитываются при каждом пересчете в Excel.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-173">These functions recalculate every time Excel recalculates.</span></span> <span data-ttu-id="c5e6c-174">К примеру, представьте себе ячейку, вызывающую функцию `NOW`.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-174">For example, imagine a cell that calls the function `NOW`.</span></span> <span data-ttu-id="c5e6c-175">При каждом вызове `NOW` она будет автоматически возвращать текущую дату и время.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-175">Every time `NOW` is called, it will automatically return the current date and time.</span></span>

<span data-ttu-id="c5e6c-176">В Excel есть несколько встроенных переменных функций, таких как `RAND` и `TODAY`.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-176">Excel contains several built-in volatile functions, such as `RAND` and `TODAY`.</span></span> <span data-ttu-id="c5e6c-177">Полный список переменных функций Excel см. в статье [Переменные и постоянные функции](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).</span><span class="sxs-lookup"><span data-stu-id="c5e6c-177">For a comprehensive list of Excel's volatile functions, see [Volatile and Non-Volatile Functions](/office/client-developer/excel/excel-recalculation#volatile-and-non-volatile-functions).</span></span>

<span data-ttu-id="c5e6c-178">Пользовательские функции позволяют создавать собственные переменные функции, которые могут быть полезны при обработке дат, времени, случайных чисел и моделировании.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-178">Custom functions allow you to create your own volatile functions, which may be useful when handling dates, times, random numbers, and modelling.</span></span> <span data-ttu-id="c5e6c-179">Например, при моделированиях методом Монте-Карло требуется создание случайных входных данных, чтобы определить оптимальное решение.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-179">For example, Monte Carlo simulations require generation of random inputs to determine an optimal solution.</span></span>

<span data-ttu-id="c5e6c-180">Чтобы объявить функцию переменной, добавьте `"volatile": true` в объект `options` для функции в файле метаданных JSON, как показано в приведенном ниже примере кода.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-180">To declare a function volatile, add `"volatile": true` within the `options` object  for the function in the JSON metadata file, as shown in the following code sample.</span></span> <span data-ttu-id="c5e6c-181">Обратите внимание, что функция не может одновременно иметь значения `"streaming": true` и `"volatile": true`. Если оба параметра помечены как `true`, параметр переменности будет игнорироваться.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-181">Note that a function cannot be marked both `"streaming": true` and `"volatile": true`; in the case where both are marked `true` the volatile option will be ignored.</span></span>

```json
{
 "id": "TOMORROW",
  "name": "TOMORROW",
  "description":  "Returns tomorrow’s date",
  "helpUrl": "http://www.contoso.com",
  "result": {
      "type": "string",
      "dimensionality": "scalar"
  },
  "options": {
      "volatile": true
  }
}
```

## <a name="saving-and-sharing-state"></a><span data-ttu-id="c5e6c-182">Состояние сохранения и совместного использования</span><span class="sxs-lookup"><span data-stu-id="c5e6c-182">Saving and sharing state</span></span>

<span data-ttu-id="c5e6c-183">Пользовательские функции могут сохранять данные в глобальных переменных JavaScript, которые можно использовать в последующих вызовах.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-183">Custom functions can save data in global JavaScript variables, which can be used in subsequent calls.</span></span> <span data-ttu-id="c5e6c-184">Сохраненное состояние полезно, когда пользователи вызывают одни и те же настраиваемые функций из более чем одной ячейки, так как все экземпляры функции могут получить доступ к состоянию.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-184">Saved state is useful when users call the same custom function from more than one cell, because all instances of the function can access the state.</span></span> <span data-ttu-id="c5e6c-185">Например, вы можете сохранить данные, возвращенные при вызове веб-ресурса, чтобы не пришлось обеспечивать выполнение дополнительных вызовов.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-185">For example, you may save the data returned from a call to a web resource to avoid making additional calls to the same web resource.</span></span>

<span data-ttu-id="c5e6c-186">В приведенном ниже примере кода показана реализация вышеописанной функции передачи температуры, сохраняющей состояние с помощью глобальной переменной.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-186">The following code sample shows an implementation of a temperature-streaming function that saves state globally.</span></span> <span data-ttu-id="c5e6c-187">Обратите внимание на следующие особенности этого кода:</span><span class="sxs-lookup"><span data-stu-id="c5e6c-187">Note the following about this code:</span></span>

- <span data-ttu-id="c5e6c-188">Функция `streamTemperature` обновляет значение температуры, которое отображается в ячейке, каждую секунду и использует переменную `savedTemperatures` как источник данных.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-188">The `streamTemperature` function updates the temperature value that's displayed in the cell every second and it uses the `savedTemperatures` variable as its data source.</span></span>

- <span data-ttu-id="c5e6c-189">Так как `streamTemperature` — это функция потоковой передачи, она реализует обработчик отмены, который будет запускаться, если функция была отменена.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-189">Because `streamTemperature` is a streaming function, it implements a cancellation handler that will run when the function is canceled.</span></span>

- <span data-ttu-id="c5e6c-190">Если пользователь вызывает функцию `streamTemperature` из нескольких ячеек в Excel, функция `streamTemperature` считывает данные из той же самой переменной `savedTemperatures` при каждом запуске.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-190">If a user calls the `streamTemperature` function from multiple cells in Excel, the `streamTemperature` function reads data from the same `savedTemperatures` variable each time it runs.</span></span> 

- <span data-ttu-id="c5e6c-191">Функция `refreshTemperature` ежесекундно считывает температуру определенного термометра и сохраняет результат в переменной `savedTemperatures`.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-191">The `refreshTemperature` function reads the temperature of a particular thermometer every second and stores the result in the `savedTemperatures` variable.</span></span> <span data-ttu-id="c5e6c-192">Так как функция `refreshTemperature` недоступна для конечных пользователей в Excel, ее не нужно регистрировать в JSON-файле.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-192">Because the `refreshTemperature` function is not exposed to end users in Excel, it does not need to be registered in the JSON file.</span></span>

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

## <a name="coauthoring"></a><span data-ttu-id="c5e6c-193">Совместное редактирование</span><span class="sxs-lookup"><span data-stu-id="c5e6c-193">Coauthoring</span></span>

<span data-ttu-id="c5e6c-194">Excel Online и Excel для Windows с подпиской на Office 365 позволяют совместно редактировать документы. Эта функция работает с пользовательскими функциями.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-194">Excel Online and Excel for Windows with an Office 365 subscription allow you to coauthor documents and this feature works with custom functions.</span></span> <span data-ttu-id="c5e6c-195">Если в книге используется пользовательская функция, вашему коллеге будет предложено загрузить надстройку пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-195">If your workbook uses a custom function, your colleague will be prompted to load the custom function's add-in.</span></span> <span data-ttu-id="c5e6c-196">Когда вы оба загрузите надстройку, пользовательская функция поделится результатами с помощью совместного редактирования.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-196">Once you both have loaded the add-in, the custom function will share results through coauthoring.</span></span>

<span data-ttu-id="c5e6c-197">Дополнительные сведения о совместном редактировании см. в статье [О совместном редактировании в Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).</span><span class="sxs-lookup"><span data-stu-id="c5e6c-197">For more information on coauthoring, see [About coauthoring in Excel](/office/vba/excel/concepts/about-coauthoring-in-excel).</span></span>

## <a name="working-with-ranges-of-data"></a><span data-ttu-id="c5e6c-198">Работа с диапазонами данных</span><span class="sxs-lookup"><span data-stu-id="c5e6c-198">Working with ranges of data</span></span>

<span data-ttu-id="c5e6c-199">Ваша пользовательская функция может принимать широкий диапазон данных в виде входных параметров или возвращать широкий диапазон данных.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-199">Your custom function may accept a range of data as an input parameter, or it may return a range of data.</span></span> <span data-ttu-id="c5e6c-200">В JavaScript диапазон данных будет иметь вид двумерного массива.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-200">In JavaScript, a range of data is represented as a two-dimensional array.</span></span>

<span data-ttu-id="c5e6c-201">Например, предположим, что функция возвращает второе по величине значение из диапазона значений, хранящихся в Excel.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-201">For example, suppose that your function returns the second highest value from a range of numbers stored in Excel.</span></span> <span data-ttu-id="c5e6c-202">Приведенная ниже функция принимает параметр `values`, относящийся к типу `Excel.CustomFunctionDimensionality.matrix`.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-202">The following function accepts the parameter `values`, which is of type `Excel.CustomFunctionDimensionality.matrix`.</span></span> <span data-ttu-id="c5e6c-203">Обратите внимание, что в метаданных JSON для данной функции вам следует задать для параметра свойство `type` в `matrix`.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-203">Note that in the JSON metadata for this function, you would set the parameter's `type` property to `matrix`.</span></span>

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

## <a name="determine-which-cell-invoked-your-custom-function"></a><span data-ttu-id="c5e6c-204">Определение того, какая ячейка вызывала пользовательскую функцию</span><span class="sxs-lookup"><span data-stu-id="c5e6c-204">Determine which cell invoked your custom function</span></span>

<span data-ttu-id="c5e6c-205">В некоторых случаях вам потребуется получить адрес ячейки, которая вызывала пользовательскую функцию.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-205">In some cases you'll need to get the address of the cell that invoked your custom function.</span></span> <span data-ttu-id="c5e6c-206">Это может быть полезно в следующих типах сценариев:</span><span class="sxs-lookup"><span data-stu-id="c5e6c-206">This may be useful in the following types of scenarios:</span></span>

- <span data-ttu-id="c5e6c-207">Форматирование диапазонов: Используйте адрес ячейки в качестве ключа для хранения сведений в [AsyncStorage](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data).</span><span class="sxs-lookup"><span data-stu-id="c5e6c-207">Formatting ranges: Use the cell's address as the key to store information in [AsyncStorage](/office/dev/add-ins/excel/custom-functions-runtime#storing-and-accessing-data).</span></span> <span data-ttu-id="c5e6c-208">После этого используйте событие [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) в Excel, чтобы загрузить ключ из `AsyncStorage`.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-208">Then, use [onCalculated](/javascript/api/excel/excel.worksheet#oncalculated) in Excel to load the key from `AsyncStorage`.</span></span>
- <span data-ttu-id="c5e6c-209">Отображение кэшированных значений. Если функция используется в автономном режиме, отображайте сохраненные в кэше значения из `AsyncStorage` с помощью `onCalculated`.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-209">Displaying cached values: If your function is used offline, display stored cached values from `AsyncStorage` using `onCalculated`.</span></span>
- <span data-ttu-id="c5e6c-210">Сверка: используйте адрес ячейки, чтобы найти исходную ячейку, чтобы упростить сверку при выполнении обработки.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-210">Reconciliation: Use the cell's address to discover an origin cell to help you reconcile where processing is occurring.</span></span>

<span data-ttu-id="c5e6c-211">Сведения об адресе ячейки предоставляются только в том случае, если параметру `requiresAddress` присвоено значение `true` в файле метаданных JSON функции.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-211">The information about a cell's address is exposed only if `requiresAddress` is marked as `true` in the function's JSON metadata file.</span></span> <span data-ttu-id="c5e6c-212">Ниже приведен пример:</span><span class="sxs-lookup"><span data-stu-id="c5e6c-212">The following sample gives an example of this:</span></span>

```JSON
{
   "id": "ADDTIME",
   "name": "ADDTIME",
   "description": "Display current date and add the amount of hours to it designated by the parameter",
   "helpUrl": "http://www.contoso.com",
   "result": {
      "type": "number",
      "dimensionality": "scalar"
   },
   "parameters": [
      {
         "name": "Additional time",
         "description": "Amount of hours to increase current date by",
         "type": "number",
         "dimensionality": "scalar"
      }
   ],
   "options": {
      "requiresAddress": true
   }
}
```

<span data-ttu-id="c5e6c-213">Чтобы найти адрес ячейки, в файл скрипта (**./src/customfunctions.js** или **./src/customfunctions.ts**) потребуется также добавить функцию `getAddress`.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-213">In the script file (**./src/customfunctions.js** or **./src/customfunctions.ts**), you'll also need to add a `getAddress` function to find a cell's address.</span></span> <span data-ttu-id="c5e6c-214">В этой функции можно использовать параметры, как показано в примере ниже в виде `parameter1`.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-214">This function may take parameters, as shown in the following sample as `parameter1`.</span></span> <span data-ttu-id="c5e6c-215">В качестве последнего параметра всегда будет использоваться `invocationContext` — объект, содержащий расположение ячейки, которое передает приложение Excel, если параметру `requiresAddress` присвоено значение `true` в файле метаданных JSON.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-215">The last parameter will always be `invocationContext`, an object containing the cell's location that Excel passes down when `requiresAddress` is marked as `true` in your JSON metadata file.</span></span>

```js
function getAddress(parameter1, invocationContext) {
    return invocationContext.address;
}
```

<span data-ttu-id="c5e6c-216">По умолчанию значения, возвращаемые из функции `getAddress`, соответствуют следующему формату: `SheetName!CellNumber`.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-216">By default, values returned from a `getAddress` function follow the following format: `SheetName!CellNumber`.</span></span> <span data-ttu-id="c5e6c-217">Например, если функция вызвана с листа с названием Expenses (Расходы) в ячейке B2, возвращаемым значением будет `Expenses!B2`.</span><span class="sxs-lookup"><span data-stu-id="c5e6c-217">For example, if a function was called from a sheet called Expenses in cell B2, the returned value would be `Expenses!B2`.</span></span>

## <a name="known-issues"></a><span data-ttu-id="c5e6c-218">Известные проблемы</span><span class="sxs-lookup"><span data-stu-id="c5e6c-218">Known issues</span></span>

<span data-ttu-id="c5e6c-219">С известными проблемами можно ознакомиться в нашем [репозитории GitHub, посвященном пользовательским функциям в Excel](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span><span class="sxs-lookup"><span data-stu-id="c5e6c-219">See known issues on our [Excel Custom Functions GitHub repo](https://github.com/OfficeDev/Excel-Custom-Functions/issues).</span></span>

## <a name="see-also"></a><span data-ttu-id="c5e6c-220">См. также</span><span class="sxs-lookup"><span data-stu-id="c5e6c-220">See also</span></span>

* [<span data-ttu-id="c5e6c-221">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="c5e6c-221">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="c5e6c-222">Среда выполнения для пользовательских функций Excel</span><span class="sxs-lookup"><span data-stu-id="c5e6c-222">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="c5e6c-223">Рекомендации по пользовательским функциям</span><span class="sxs-lookup"><span data-stu-id="c5e6c-223">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="c5e6c-224">Журнал изменений пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="c5e6c-224">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="c5e6c-225">Руководство по настраиваемым функциям в Excel</span><span class="sxs-lookup"><span data-stu-id="c5e6c-225">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
