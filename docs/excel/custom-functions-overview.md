---
ms.date: 09/20/2018
description: Создание настраиваемой функции в Excel с помощью JavaScript.
title: Создание настраиваемых функций в Excel (предварительная версия)
ms.openlocfilehash: b214329fe50955d0f39d50f674152f475ca24b4d
ms.sourcegitcommit: eb74e94d3e1bc1930a9c6582a0a99355d0da34f2
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/25/2018
ms.locfileid: "25005045"
---
# <a name="create-custom-functions-in-excel-preview"></a><span data-ttu-id="e06d5-103">Создание настраиваемых функций в Excel (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="e06d5-103">Create custom functions in Excel (Preview)</span></span>

<span data-ttu-id="e06d5-104">Настраиваемые функции позволяют разработчикам добавлять новые функции в Excel, определяя эти функции в JavaScript как часть надстройки.</span><span class="sxs-lookup"><span data-stu-id="e06d5-104">Custom functions enable developers to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="e06d5-105">Пользователи в Excel могут получать доступ к настраиваемым функциям, как к любой другой встроенной функции Excel (например, `SUM()`).</span><span class="sxs-lookup"><span data-stu-id="e06d5-105">Users within Excel can access custom functions like any other native function in Excel (such as `SUM()`).</span></span> <span data-ttu-id="e06d5-106">В этой статье описывается порядок создания настраиваемых функций в Excel.</span><span class="sxs-lookup"><span data-stu-id="e06d5-106">This article explains how to create custom functions in Excel.</span></span>

<span data-ttu-id="e06d5-107">На следующем рисунке показан конечный пользователь, вставляющий пользовательскую функцию в ячейку листа Excel.</span><span class="sxs-lookup"><span data-stu-id="e06d5-107">The following illustration shows an end user inserting a custom function into a cell of an Excel worksheet.</span></span> <span data-ttu-id="e06d5-108">Настраиваемая функция `CONTOSO.ADD42` предназначена для добавления 42 к паре чисел, которую пользователь указывает в качестве входных параметров для функции.</span><span class="sxs-lookup"><span data-stu-id="e06d5-108">The `CONTOSO.ADD42` custom function is designed to add 42 to the pair of numbers that the user specifies as input parameters to the function.</span></span>

<img alt="animated image showing an end user inserting the CONTOSO.ADD42 custom function into a cell of an Excel worksheet" src="../images/custom-function.gif" width="579" height="383" />

<span data-ttu-id="e06d5-109">Следующий код определяет настраиваемую функцию `ADD42`.</span><span class="sxs-lookup"><span data-stu-id="e06d5-109">The following code defines the `ADD42` custom function.</span></span>

```js
function ADD42(a, b) {
    return a + b + 42;
}
```

<span data-ttu-id="e06d5-110">Настраиваемые функции теперь доступны для разработчика в форме предварительной версии на Windows, Mac, а также в Excel Online.</span><span class="sxs-lookup"><span data-stu-id="e06d5-110">Custom functions are now available in Developer Preview on Windows, Mac, and Excel Online.</span></span> <span data-ttu-id="e06d5-111">Чтобы попробовать их, выполните следующие действия.</span><span class="sxs-lookup"><span data-stu-id="e06d5-111">To try them, complete these steps:</span></span>

1. <span data-ttu-id="e06d5-112">Установите Office (сборка 10827 на Windows или 13.329 на Mac) и присоединитесь к программе [предварительной оценки Office](https://products.office.com/office-insider) .</span><span class="sxs-lookup"><span data-stu-id="e06d5-112">Install Office (build 9325 on Windows or 13.329 on Mac) and join the [Office Insider](https://products.office.com/office-insider) program.</span></span> <span data-ttu-id="e06d5-113">Вы должны присоединиться к программе предварительной оценки Office, чтобы иметь доступ к настраиваемым функциям. В настоящее время настраиваемые функции отключены во всех сборках Office, если вы не являетесь членом программы предварительной оценки Office.</span><span class="sxs-lookup"><span data-stu-id="e06d5-113">You must join the Office Insider program in order to have access to custom functions; currently, custom functions are disabled across all Office builds unless you are a member of the Office Insider program.</span></span>

2. <span data-ttu-id="e06d5-114">Создайте проект надстройки настраиваемых функций Excel с помощью [Yo Office](https://github.com/OfficeDev/generator-office), а затем следуйте инструкциям в [OfficeDev/Excel-Custom-Functions README](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/README.md) для использования проекта.</span><span class="sxs-lookup"><span data-stu-id="e06d5-114">Use [Yo Office](https://github.com/OfficeDev/generator-office) to create an Excel Custom Functions add-in project, and then follow the instructions in the [OfficeDev/Excel-Custom-Functions README](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/README.md) to use the project.</span></span>

3. <span data-ttu-id="e06d5-115">Введите `=CONTOSO.ADD42(1,2)` в любой ячейке листа Excel, после чего нажмите на клавишу **Enter**, чтобы запустить настраиваемую функцию.</span><span class="sxs-lookup"><span data-stu-id="e06d5-115">Type `=CONTOSO.ADD42(1,2)` into any cell, and press **Enter** to run the custom function.</span></span>

> [!NOTE]
> <span data-ttu-id="e06d5-116">В разделе [Известные проблемы](#known-issues) далее в этой статье указаны текущие ограничения настраиваемых функций.</span><span class="sxs-lookup"><span data-stu-id="e06d5-116">The [Known issues](#known-issues) section later in this article specifies current limitations of custom functions.</span></span>

## <a name="learn-the-basics"></a><span data-ttu-id="e06d5-117">Ознакомьтесь с основными сведениями</span><span class="sxs-lookup"><span data-stu-id="e06d5-117">Learn the basics</span></span>

<span data-ttu-id="e06d5-118">В проекте настраиваемых функций, который вы создали с помощью [Yo Office](https://github.com/OfficeDev/generator-office), вы увидите следующие файлы:</span><span class="sxs-lookup"><span data-stu-id="e06d5-118">In the custom functions project that you've created using [Yo Office](https://github.com/OfficeDev/generator-office), you’ll see the following files:</span></span>

| <span data-ttu-id="e06d5-119">Файл</span><span class="sxs-lookup"><span data-stu-id="e06d5-119">File</span></span> | <span data-ttu-id="e06d5-120">Формат файла</span><span class="sxs-lookup"><span data-stu-id="e06d5-120">File format</span></span> | <span data-ttu-id="e06d5-121">Описание</span><span class="sxs-lookup"><span data-stu-id="e06d5-121">Description</span></span> |
|------|-------------|-------------|
| <span data-ttu-id="e06d5-122">**./src/customfunctions.js**</span><span class="sxs-lookup"><span data-stu-id="e06d5-122">**./src/customfunctions.js**</span></span> | <span data-ttu-id="e06d5-123">JavaScript</span><span class="sxs-lookup"><span data-stu-id="e06d5-123">JavaScript</span></span> | <span data-ttu-id="e06d5-124">Содержит код, который определяет настраиваемые функции.</span><span class="sxs-lookup"><span data-stu-id="e06d5-124">Contains the code that defines custom functions.</span></span> |
| <span data-ttu-id="e06d5-125">**./config/customfunctions.json**</span><span class="sxs-lookup"><span data-stu-id="e06d5-125">**./config/customfunctions.json**</span></span> | <span data-ttu-id="e06d5-126">JSON</span><span class="sxs-lookup"><span data-stu-id="e06d5-126">JSON</span></span> | <span data-ttu-id="e06d5-127">Содержит метаданные, которые описывают настраиваемые функции и позволяют Excel регистрировать настраиваемые функции, чтобы сделать их доступными для пользователей.</span><span class="sxs-lookup"><span data-stu-id="e06d5-127">Contains metadata that describes custom functions and enables Excel to register the custom functions in order to make them available to end-users.</span></span> |
| <span data-ttu-id="e06d5-128">**./index.html**</span><span class="sxs-lookup"><span data-stu-id="e06d5-128">**./index.html**</span></span> | <span data-ttu-id="e06d5-129">HTML</span><span class="sxs-lookup"><span data-stu-id="e06d5-129">HTML</span></span> | <span data-ttu-id="e06d5-130">Предоставляет ссылку в тегах &lt;script&gt; на файл JavaScript, который определяет пользовательские функции.</span><span class="sxs-lookup"><span data-stu-id="e06d5-130">Provides a &lt;script&gt; reference to the JavaScript file that defines custom functions.</span></span> |
| <span data-ttu-id="e06d5-131">**./manifest.xml**</span><span class="sxs-lookup"><span data-stu-id="e06d5-131">**Manifest.xml**</span></span> | <span data-ttu-id="e06d5-132">XML</span><span class="sxs-lookup"><span data-stu-id="e06d5-132">XML</span></span> | <span data-ttu-id="e06d5-133">Указывает пространство имен для всех настраиваемых функций в пределах надстройки и расположение файлов JavaScript, JSON и HTML, указанных ранее в этой таблице.</span><span class="sxs-lookup"><span data-stu-id="e06d5-133">Specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files that are listed previously in this table.</span></span> |

### <a name="manifest-file-manifestxml"></a><span data-ttu-id="e06d5-134">Файл манифеста (./manifest.xml)</span><span class="sxs-lookup"><span data-stu-id="e06d5-134">Manifest file (manifest.xml)</span></span>

<span data-ttu-id="e06d5-135">XML-файл манифеста для надстройки, который определяет настраиваемые функции, определяет пространство имен для всех настраиваемых функций в пределах надстройки и расположение файлов JavaScript, JSON и HTML.</span><span class="sxs-lookup"><span data-stu-id="e06d5-135">The XML manifest file for an add-in that defines custom functions specifies the namespace for all custom functions within the add-in and the location of the JavaScript, JSON, and HTML files.</span></span> <span data-ttu-id="e06d5-136">Ниже показан пример использования элементов `<ExtensionPoint>` и `<Resources>` в разметке XML. Эти элементы необходимо включить в манифест надстройки, чтобы Excel мог выполнять настраиваемые функции.</span><span class="sxs-lookup"><span data-stu-id="e06d5-136">The following XML markup shows an example of the `<ExtensionPoint>` and `<Resources>` elements that you must include in an add-in's manifest in order to enable Excel to run custom functions.</span></span>  

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
            <bt:String id="namespace" DefaultValue="CONTOSO" /> <!--specifies the namespace that will be prepended to a function's name when it is called in Excel. For example, a function named "ADD42" is invoked as `=CONTOSO.ADD42` in Excel.-->
        </bt:ShortStrings>
    </Resources>
</VersionOverrides>
```

> [!NOTE]
> <span data-ttu-id="e06d5-137">Функции Excel добавляются пространством имен, указанным в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="e06d5-137">Functions in Excel are prepended by the namespace specified in your XML manifest file.</span></span> <span data-ttu-id="e06d5-138">Пространство имен функции предшествует имени функции и отделяется от него точкой.</span><span class="sxs-lookup"><span data-stu-id="e06d5-138">A function's namespace comes before the function name and they are separated by a period.</span></span> <span data-ttu-id="e06d5-139">Например, чтобы вызвать функцию `ADD42()` в ячейке листа Excel, следует ввести `=CONTOSO.ADD42`, так как CONTOSO — это пространство имен, а `ADD42` — имя функции, указанной в файле JSON.</span><span class="sxs-lookup"><span data-stu-id="e06d5-139">For example, to call the function `ADD42()` in the cell of an Excel worksheet, you would type `=CONTOSO.ADD42`, because CONTOSO is the namespace and `ADD42` is the name of the function specified in the JSON file.</span></span> <span data-ttu-id="e06d5-140">Данное пространство имен используется в качестве идентификатора для вашей организации или надстройки.</span><span class="sxs-lookup"><span data-stu-id="e06d5-140">The prefix is intended to be used as an identifier for your add-in.</span></span> 

### <a name="json-file-configcustomfunctionsjson"></a><span data-ttu-id="e06d5-141">Файл JSON (./config/customfunctions.json)</span><span class="sxs-lookup"><span data-stu-id="e06d5-141">JSON file (./config/customfunctions.json)</span></span>

<span data-ttu-id="e06d5-142">Файл метаданных настраиваемых функций предоставляет информацию, которую Excel требует для их регистрации, и делает их доступными для конечных пользователей.</span><span class="sxs-lookup"><span data-stu-id="e06d5-142">A custom functions metadata file provides the information that Excel requires to register the custom functions and make them available to end-users.</span></span> <span data-ttu-id="e06d5-143">Настраиваемые функции регистрируются при первом запуске надстройки пользователем.</span><span class="sxs-lookup"><span data-stu-id="e06d5-143">The custom functions are registered when a user runs the add-in for the first time.</span></span> <span data-ttu-id="e06d5-144">После этого пользователь может использовать их во всех книгах (то есть, не только в книге, в которой первоначально выполнялась надстройка).</span><span class="sxs-lookup"><span data-stu-id="e06d5-144">After that, they are available, for that same user, in all workbooks (not only the one where the add-in ran initially.)</span></span>

> [!TIP]
> <span data-ttu-id="e06d5-145">Чтобы настраиваемая функция правильно работала в Excel Online, в параметрах сервера, на котором размещен файл JSON, необходимо включить [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS).</span><span class="sxs-lookup"><span data-stu-id="e06d5-145">Your server settings for the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>

<span data-ttu-id="e06d5-146">Следующий код в файле **customfunctions.json** определяет метаданные для функции `ADD42`, описанной выше в этой статье.</span><span class="sxs-lookup"><span data-stu-id="e06d5-146">The following code in **customfunctions.json** specifies the metadata for the `ADD42` function that was described previously in this article.</span></span> <span data-ttu-id="e06d5-147">Эти метаданные определяют имя функции, ее описание, возвращаемое значение, входные параметры и многое другое.</span><span class="sxs-lookup"><span data-stu-id="e06d5-147">This metadata defines the function's name, description, return value, input parameters, and more.</span></span> <span data-ttu-id="e06d5-148">В таблице, следующей за этим примером кода, содержится подробная информация об отдельных свойствах этого объекта JSON.</span><span class="sxs-lookup"><span data-stu-id="e06d5-148">The table that follows this code sample provides detailed information about the individual properties within this JSON object.</span></span>

```json
{
    "$schema": "https://developer.microsoft.com/json-schemas/office-js/custom-functions.schema.json",
    "functions": [
        {
            "id": "ADD42",
            "name": "ADD42",
            "description":  "adds 42 to the input numbers",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [                {
                    "name": "number 1",
                    "description": "the first number to be added",
                    "type": "number",
                    "dimensionality": "scalar"
                },
                {
                    "name": "number 2",
                    "description": "the second number to be added",
                    "type": "number",
                    "dimensionality": "scalar"
                }
            ],
        }
    ]
}
```

<span data-ttu-id="e06d5-149">В следующей таблице перечислены свойства, которые обычно присутствуют в файле метаданных JSON.</span><span class="sxs-lookup"><span data-stu-id="e06d5-149">The following table lists the properties that are typically present in the JSON metadata file.</span></span> <span data-ttu-id="e06d5-150">Более подробные сведения о файле метаданных JSON, в том числе о параметрах, не использующихся в предыдущем примере, см. в статье [Метаданные настраиваемых функций](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="e06d5-150">For more detailed information about the JSON metadata file, including options not used in the previous example, see [Custom functions metadata](custom-functions-json.md).</span></span>

| <span data-ttu-id="e06d5-151">Свойство</span><span class="sxs-lookup"><span data-stu-id="e06d5-151">Property</span></span>  | <span data-ttu-id="e06d5-152">Описание</span><span class="sxs-lookup"><span data-stu-id="e06d5-152">Description</span></span> |
|---------|---------|
| `id` | <span data-ttu-id="e06d5-153">Уникальный идентификатор для функции.</span><span class="sxs-lookup"><span data-stu-id="e06d5-153">A unique ID for the group.</span></span> <span data-ttu-id="e06d5-154">Этот идентификатор не должен изменяться после его установки.</span><span class="sxs-lookup"><span data-stu-id="e06d5-154">This ID should not be changed after it is set.</span></span> |
| `name` | <span data-ttu-id="e06d5-155">Имя функции, отображаемое в меню автозаполнения, когда пользователь вводит формулу в ячейке.</span><span class="sxs-lookup"><span data-stu-id="e06d5-155">Name of the function that is shown in the autocomplete menu as a user types a formula within a cell.</span></span> <span data-ttu-id="e06d5-156">В меню автозаполнения это значение будет иметь префикс пространства имен настраиваемых функций, указанного в XML-файле манифеста.</span><span class="sxs-lookup"><span data-stu-id="e06d5-156">In the autocomplete menu, this value will be prefixed by the custom functions namespace that's specified in the XML manifest file.</span></span> |
| `helpUrl` | <span data-ttu-id="e06d5-157">URL-адрес страницы, которая отображается, когда пользователь запрашивает справку.</span><span class="sxs-lookup"><span data-stu-id="e06d5-157">Url for a page that is shown when a user requests help.</span></span> |
| `description` | <span data-ttu-id="e06d5-158">Описывает, что делает функция.</span><span class="sxs-lookup"><span data-stu-id="e06d5-158">Describes what the function does.</span></span> <span data-ttu-id="e06d5-159">Это значение появляется как подсказка, когда функция является выбранным элементом в меню автозаполнения в Excel.</span><span class="sxs-lookup"><span data-stu-id="e06d5-159">This value appears as a tooltip when the function is the selected item in the autocomplete menu within Excel.</span></span> |
| `result`  | <span data-ttu-id="e06d5-160">Объект, который определяет тип данных, который возвращается функцией.</span><span class="sxs-lookup"><span data-stu-id="e06d5-160">Object that defines the type of information that is returned by the function.</span></span> <span data-ttu-id="e06d5-161">Значение дочернего свойства `type` может быть **string**, **number**или **boolean**.</span><span class="sxs-lookup"><span data-stu-id="e06d5-161">The value of the `type` child property can be **string**, **number**, or **boolean**.</span></span> <span data-ttu-id="e06d5-162">Дочернему свойству `dimensionality` может присваиваться значение **scalar** или **matrix** (двухмерный массив значений указанного типа `type`).</span><span class="sxs-lookup"><span data-stu-id="e06d5-162">The `dimensionality` property can be \*\*\*\* or \*\*\*\* (a two-dimensional array of values of the specified `type`.)</span></span> |
| `parameters` | <span data-ttu-id="e06d5-163">Массив, который определяет входные параметры для функции.</span><span class="sxs-lookup"><span data-stu-id="e06d5-163">Array that defines the input parameters for the function.</span></span> <span data-ttu-id="e06d5-164">В Excel intelliSense появляются дочерние свойства `name` и `description`.</span><span class="sxs-lookup"><span data-stu-id="e06d5-164">The `name` and `description` child properties are used in the Excel intellisense.</span></span> <span data-ttu-id="e06d5-165">Дочерние свойства `type` и `dimensionality` идентичны дочерним свойствам объекта `result`, описанного выше в этой таблице.</span><span class="sxs-lookup"><span data-stu-id="e06d5-165">The `type` and `dimensionality` child properties are identical to the child properties of the `result` object that is described previously in this table.</span></span> |
| `options` | <span data-ttu-id="e06d5-166">Это позволяет настраивать некоторые аспекты того, как и когда Excel выполняет эту функцию.</span><span class="sxs-lookup"><span data-stu-id="e06d5-166">The  property enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="e06d5-167">Подробнее о том, как можно использовать это свойство, см. в разделах [Потоковые функции](#streamed-functions) и [Отмена](#canceling-a-function) ниже в этой статье.</span><span class="sxs-lookup"><span data-stu-id="e06d5-167">For more information about how this property can be used, see [Streamed functions](#streamed-functions) and [Cancellation](#canceling-a-function) later in this article.</span></span> |

## <a name="functions-that-return-data-from-external-sources"></a><span data-ttu-id="e06d5-168">Функции, возвращающие данные из внешних источников</span><span class="sxs-lookup"><span data-stu-id="e06d5-168">Functions that return data from external sources</span></span>

<span data-ttu-id="e06d5-169">Если настраиваемая функция получает данные из внешнего источника, например веб-сайта, она должна:</span><span class="sxs-lookup"><span data-stu-id="e06d5-169">If a custom function retrieves data from an external source such as the web, it must:</span></span>

1. <span data-ttu-id="e06d5-170">возвращать обещание JavaScript в Excel.</span><span class="sxs-lookup"><span data-stu-id="e06d5-170">Return a JavaScript Promise to Excel.</span></span>

2. <span data-ttu-id="e06d5-171">Разрешите обещание с помощью окончательного значения, воспользовавшись функцией обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e06d5-171">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="e06d5-172">Пока Excel ожидает конечный результат, настраиваемые функции отображают в ячейке временный результат `#GETTING_DATA`.</span><span class="sxs-lookup"><span data-stu-id="e06d5-172">Asynchronous functions display a `#GETTING_DATA` temporary error in the cell while Excel waits for the final result.</span></span> <span data-ttu-id="e06d5-173">Во время ожидания результата пользователи могут нормально взаимодействовать с остальной частью листа.</span><span class="sxs-lookup"><span data-stu-id="e06d5-173">Users can interact normally with the rest of the spreadsheet while they wait for the result.</span></span>

<span data-ttu-id="e06d5-174">В следующем примере кода настраиваемая функция `getTemperature()` получает от термометра текущую температуру.</span><span class="sxs-lookup"><span data-stu-id="e06d5-174">In the following code sample, the `getTemperature()` custom function retrieves the current temperature of a thermometer.</span></span> <span data-ttu-id="e06d5-175">Обратите внимание на то, что функция `sendWebRequest` является гипотетической (не указывается здесь) и использует XHR для вызова веб-службы температуры.</span><span class="sxs-lookup"><span data-stu-id="e06d5-175">Note that `sendWebRequest` is a hypothetical function, not specified here, that uses XHR to call a temperature web service.</span></span>

```js
function getTemperature(thermometerID){
    return new Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

## <a name="streamed-functions"></a><span data-ttu-id="e06d5-176">Потоковые функции</span><span class="sxs-lookup"><span data-stu-id="e06d5-176">Streamed functions</span></span>

<span data-ttu-id="e06d5-177">Потоковые настраиваемые функции позволяют вам выводить данные в ячейки многократно с течением времени, не требуя от пользователя явно запрашивать пересчет.</span><span class="sxs-lookup"><span data-stu-id="e06d5-177">Streamed custom functions enable you to output data to cells repeatedly over time, without requiring a user to explicitly request recalculation.</span></span> <span data-ttu-id="e06d5-178">Следующий пример кода представляет собой настраиваемую функцию, которая каждую секунду добавляет число к результату.</span><span class="sxs-lookup"><span data-stu-id="e06d5-178">The following example is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="e06d5-179">Обратите внимание на следующие особенности этого кода:</span><span class="sxs-lookup"><span data-stu-id="e06d5-179">Note the following about this code:</span></span>

- <span data-ttu-id="e06d5-180">Excel автоматически отображает каждое новое значение при помощи `setResult` обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e06d5-180">Excel displays each new value automatically using the `setResult` callback.</span></span>

- <span data-ttu-id="e06d5-181">Последний параметр, `handler`, никогда не указывается в коде регистрации и не отображается в меню автозаполнения для пользователей Excel, когда они вводят функцию.</span><span class="sxs-lookup"><span data-stu-id="e06d5-181">For streamed functions, the final parameter, `handler`, is never specified in your registration code, and it does not display in the autocomplete menu to Excel users when they enter the function.</span></span> <span data-ttu-id="e06d5-182">Это объект, который содержит функцию обратного вызова `setResult`, используемую для передачи данных из функции в Excel и обновления значения ячейки.</span><span class="sxs-lookup"><span data-stu-id="e06d5-182">It’s an object that contains a `setResult` callback function that’s used to pass data from the function to Excel to update the value of a cell.</span></span>

- <span data-ttu-id="e06d5-183">Чтобы Excel передал функцию `setResult` объекту `handler`, необходимо объявить поддержку потоковой передачи при регистрации функции, установив параметр `"stream": true` в свойстве `options` для настраиваемой функции в JSON-файле метаданных.</span><span class="sxs-lookup"><span data-stu-id="e06d5-183">In order for Excel to pass the `setResult` function in the `handler` object, you must declare support for streaming during your function registration by setting the option `"stream": true` in the `options` property for the custom function in the registration JSON file.</span></span>

```js
function incrementValue(increment, handler){
    var result = 0;
    setInterval(function(){
         result += increment;
         handler.setResult(result);
    }, 1000);
}
```

## <a name="canceling-a-function"></a><span data-ttu-id="e06d5-184">Отмена функции</span><span class="sxs-lookup"><span data-stu-id="e06d5-184">Canceling a function</span></span>

<span data-ttu-id="e06d5-185">В некоторых случаях может потребоваться отменить выполнение потоковой настраиваемой функции, чтобы снизить ее потребление пропускной способности, рабочей памяти и загрузку процессора.</span><span class="sxs-lookup"><span data-stu-id="e06d5-185">In some situations, you may need to cancel the execution of a streamed custom function to reduce its bandwidth consumption, working memory, and CPU load.</span></span> <span data-ttu-id="e06d5-186">Excel отменяет выполнение функции в следующих ситуациях.</span><span class="sxs-lookup"><span data-stu-id="e06d5-186">Excel cancels the execution of a function in the following situations:</span></span>

- <span data-ttu-id="e06d5-187">Когда пользователь редактирует или удаляет ячейку, содержащую ссылку на функцию.</span><span class="sxs-lookup"><span data-stu-id="e06d5-187">The user edits or deletes a cell that references the function.</span></span>

- <span data-ttu-id="e06d5-188">Когда изменяется один из аргументов (входных параметров) функции.</span><span class="sxs-lookup"><span data-stu-id="e06d5-188">One of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="e06d5-189">В этом случае после отмены активируется новый вызов функции.</span><span class="sxs-lookup"><span data-stu-id="e06d5-189">In this case, a new function call is triggered in addition to the cancelation.</span></span>

- <span data-ttu-id="e06d5-190">Пользователь вызывает пересчет вручную.</span><span class="sxs-lookup"><span data-stu-id="e06d5-190">The user triggers recalculation manually.</span></span> <span data-ttu-id="e06d5-191">В этом случае после отмены активируется новый вызов функции.</span><span class="sxs-lookup"><span data-stu-id="e06d5-191">In this case, a new function call is triggered in addition to the cancelation.</span></span>

> [!NOTE]
> <span data-ttu-id="e06d5-192">Обработчик отмены необходимо реализовать для каждой потоковой функции.</span><span class="sxs-lookup"><span data-stu-id="e06d5-192">You must implement a cancellation handler for every streaming function.</span></span>

<span data-ttu-id="e06d5-193">Чтобы сделать функцию отменяемой, установите для настраиваемой функции параметр `"cancelable": true` в свойстве `options` с помощью JSON-файла метаданных.</span><span class="sxs-lookup"><span data-stu-id="e06d5-193">To make a function cancelable, set the option `"cancelable": true` in the `options` property for the custom function in the registration JSON file.</span></span>

<span data-ttu-id="e06d5-194">В следующем коде показана та же функция `incrementValue`, которая была описана выше, но на этот раз с реализованным обработчиком отмены.</span><span class="sxs-lookup"><span data-stu-id="e06d5-194">The following code shows the same `incrementValue` function that was described previously, but this time with a cancellation handler implemented.</span></span> <span data-ttu-id="e06d5-195">В этом примере при отмене функции `incrementValue` будет выполняться метод `clearInterval()`.</span><span class="sxs-lookup"><span data-stu-id="e06d5-195">In this example, `clearInterval()` will run when the `incrementValue` function is canceled.</span></span>

```js
function incrementValue(increment, handler){
    var result = 0;
    var timer = setInterval(function(){
         result += increment;
         handler.setResult(result);
    }, 1000);

    handler.onCanceled = function(){
        clearInterval(timer);
    }
}
```

## <a name="saving-and-sharing-state"></a><span data-ttu-id="e06d5-196">Сохранение и передача состояния</span><span class="sxs-lookup"><span data-stu-id="e06d5-196">Saving and sharing state</span></span>

<span data-ttu-id="e06d5-197">Настраиваемые функции могут сохранять данные в глобальных переменных JavaScript.</span><span class="sxs-lookup"><span data-stu-id="e06d5-197">Custom functions can save data in global JavaScript variables.</span></span> <span data-ttu-id="e06d5-198">При последующих вызовах настраиваемая функция может использовать значения, сохраненные в этих переменных.</span><span class="sxs-lookup"><span data-stu-id="e06d5-198">In subsequent calls, your custom function may use the values saved in these variables.</span></span> <span data-ttu-id="e06d5-199">Сохранение состояния может быть полезно, когда пользователи добавляют одну настраиваемую функцию к нескольким ячейкам, потому что все экземпляры функции могут совместно использовать ее состояние.</span><span class="sxs-lookup"><span data-stu-id="e06d5-199">Saved state is useful when users add the same custom function to more than one cell, because all the instances of the function can share the state.</span></span> <span data-ttu-id="e06d5-200">Например, вы можете сохранить данные, возвращенные при вызове веб-ресурса, чтобы не пришлось обеспечивать выполнение дополнительных вызовов.</span><span class="sxs-lookup"><span data-stu-id="e06d5-200">For example, you may save the data returned from a call to a web resource to avoid making additional calls to the same web resource.</span></span>

<span data-ttu-id="e06d5-201">В приведенном ниже примере кода показана реализация вышеописанной потоковой функции температуры, осуществляющей глобальное сохранение состояния.</span><span class="sxs-lookup"><span data-stu-id="e06d5-201">The following code shows an implementation of the previous temperature-streaming function that saves state using the  variable.</span></span> <span data-ttu-id="e06d5-202">Обратите внимание на следующие особенности этого кода:</span><span class="sxs-lookup"><span data-stu-id="e06d5-202">Note the following about this code:</span></span>

- <span data-ttu-id="e06d5-203">`refreshTemperature` это потоковая функция, ежесекундно считывающая температуру определенного термометра.</span><span class="sxs-lookup"><span data-stu-id="e06d5-203">`refreshTemperature` is a streamed function that reads the temperature of a particular thermometer every second.</span></span> <span data-ttu-id="e06d5-204">Новые температуры сохраняются в переменную `savedTemperatures`, но не обновляют значение ячейки напрямую.</span><span class="sxs-lookup"><span data-stu-id="e06d5-204">New temperatures are saved in the `savedTemperatures` variable, but does not directly update the cell value.</span></span> <span data-ttu-id="e06d5-205">Она не должен вызываться непосредственно из ячейки листа, *поэтому она не регистрируется в файле JSON*.</span><span class="sxs-lookup"><span data-stu-id="e06d5-205">It should not be directly called from a worksheet cell, *so it is not registered in the JSON file*.</span></span>

- <span data-ttu-id="e06d5-206">`streamTemperature` обновляет значения температуры, которые отображаются в ячейке каждую секунду, а в качестве источника данных использует переменную `savedTemperatures`.</span><span class="sxs-lookup"><span data-stu-id="e06d5-206">`streamTemperature` updates the temperature values displayed in the cell every second and it uses `savedTemperatures` variable as its data source.</span></span> <span data-ttu-id="e06d5-207">Она должна быть зарегистрирована в файле JSON и записана прописными буквами: `STREAMTEMPERATURE`.</span><span class="sxs-lookup"><span data-stu-id="e06d5-207">It must be registered in the JSON file, and named with all upper-case letters, `STREAMTEMPERATURE`.</span></span>

- <span data-ttu-id="e06d5-208">Пользователи могут вызывать функцию `streamTemperature` из нескольких ячеек в пользовательском интерфейсе Excel.</span><span class="sxs-lookup"><span data-stu-id="e06d5-208">Users may call `streamTemperature` from several cells in the Excel UI.</span></span> <span data-ttu-id="e06d5-209">Каждый вызов считывает данные из той же переменной `savedTemperatures`.</span><span class="sxs-lookup"><span data-stu-id="e06d5-209">Each call reads data from the same `savedTemperatures` variable.</span></span>

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

## <a name="working-with-ranges-of-data"></a><span data-ttu-id="e06d5-210">Работа с диапазонами данных</span><span class="sxs-lookup"><span data-stu-id="e06d5-210">Working with ranges of data</span></span>

<span data-ttu-id="e06d5-211">Настраиваемая функция может принимать диапазон данных в качестве входного параметра, или она может возвращать диапазон данных.</span><span class="sxs-lookup"><span data-stu-id="e06d5-211">Your custom function may accept a range of data as an input parameter, or it may return a range of data.</span></span> <span data-ttu-id="e06d5-212">В JavaScript диапазон данных представляется как двухмерный массив.</span><span class="sxs-lookup"><span data-stu-id="e06d5-212">In JavaScript, a range of data is represented as a 2-dimensional array.</span></span>

<span data-ttu-id="e06d5-213">Предположим, к примеру, что ваша функция возвращает второе наибольшое значение из диапазона чисел, хранящихся в Excel.</span><span class="sxs-lookup"><span data-stu-id="e06d5-213">For example, suppose that your function returns the second highest temperature from a range of temperature values stored in Excel.</span></span> <span data-ttu-id="e06d5-214">Следующая функция принимает параметр `values`, который имеет тип `Excel.CustomFunctionDimensionality.matrix`.</span><span class="sxs-lookup"><span data-stu-id="e06d5-214">The following function takes the parameter `values`, which is an `Excel.CustomFunctionDimensionality.matrix` parameter type.</span></span> <span data-ttu-id="e06d5-215">Обратите внимание, что в JSON-метаданных для этой функции вы должны для параметра `type` установить значение `matrix`.</span><span class="sxs-lookup"><span data-stu-id="e06d5-215">Note that in the registration JSON for this function, you would set the parameter's `type` property to `matrix`.</span></span>

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

## <a name="handling-errors"></a><span data-ttu-id="e06d5-216">Обработка ошибок</span><span class="sxs-lookup"><span data-stu-id="e06d5-216">handling errors</span></span>

<span data-ttu-id="e06d5-217">При построении надстройки, определяющей настраиваемые функции, не забудьте добавить логику для обработки ошибок, возникающих в среде выполнения.</span><span class="sxs-lookup"><span data-stu-id="e06d5-217">When you build an add-in using the Excel JavaScript API, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="e06d5-218">Обработка ошибок для настраиваемых функций совпадает с [обработкой ошибок для Excel API JavaScript в целом](excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="e06d5-218">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="e06d5-219">В следующем примере кода метод `.catch` будет обрабатывать все ошибки, возникающие ранее в коде.</span><span class="sxs-lookup"><span data-stu-id="e06d5-219">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

```js
function getComment(x) {
    let url = "https://yourhypotheticalapi/comments/" + x;

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

## <a name="known-issues"></a><span data-ttu-id="e06d5-220">Известные проблемы</span><span class="sxs-lookup"><span data-stu-id="e06d5-220">Known issues</span></span>

- <span data-ttu-id="e06d5-221">URL-адреса справки и описания параметров пока не используются в Excel.</span><span class="sxs-lookup"><span data-stu-id="e06d5-221">Help URLs and parameter descriptions are not yet used by Excel.</span></span>
- <span data-ttu-id="e06d5-222">Настраиваемые функции в настоящее время недоступны в Excel для мобильных клиентов.</span><span class="sxs-lookup"><span data-stu-id="e06d5-222">Custom functions are not currently available on Excel for mobile clients.</span></span>
- <span data-ttu-id="e06d5-223">Изменяемые функции (которые пересчитываются автоматически при изменении несвязанных данных в электронной таблице) еще не поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="e06d5-223">Volatile functions (those which recalculate automatically whenever unrelated data changes in the spreadsheet) are not yet supported.</span></span>
- <span data-ttu-id="e06d5-224">Развертывание через портал администрирования Office 365 и AppSource еще не включено.</span><span class="sxs-lookup"><span data-stu-id="e06d5-224">Deployment via the Office 365 Admin Portal and AppSource are not yet enabled.</span></span>
- <span data-ttu-id="e06d5-225">Настраиваемые функции в Excel Online могут перестать работать во время сеанса после периода бездействия.</span><span class="sxs-lookup"><span data-stu-id="e06d5-225">Custom functions in Excel Online may stop working during a session after a period of inactivity.</span></span> <span data-ttu-id="e06d5-226">Для восстановления функции обновите страницу веб-обозревателя (F5) и повторно введите настраиваемую функцию.</span><span class="sxs-lookup"><span data-stu-id="e06d5-226">Refresh the browser page (F5) and re-enter a custom function to restore the feature.</span></span>
- <span data-ttu-id="e06d5-227">Если у вас есть несколько надстроек, работающих на Excel для Windows, внутри ячейки таблицы может отображаться временный результат **#GETTING_DATA**.</span><span class="sxs-lookup"><span data-stu-id="e06d5-227">You may see the **#GETTING_DATA** temporary result within the cell(s) of a worksheet if you have multiple add-ins running on Excel for Windows.</span></span> <span data-ttu-id="e06d5-228">Закройте все окна Excel и перезапустите Excel.</span><span class="sxs-lookup"><span data-stu-id="e06d5-228">Close all Excel windows and restart Excel.</span></span>
- <span data-ttu-id="e06d5-229">Возможно, в будущем появятся специальные средства отладки для настраиваемых функций.</span><span class="sxs-lookup"><span data-stu-id="e06d5-229">Debugging tools specifically for custom functions may be available in the future.</span></span> <span data-ttu-id="e06d5-230">Тем временем вы можете выполнить отладку в Excel Online с помощью средств разработчика F12.</span><span class="sxs-lookup"><span data-stu-id="e06d5-230">In the meantime, you can debug on Excel Online using F12 developer tools.</span></span> <span data-ttu-id="e06d5-231">Подробнее см. в статье [Рекомендации по настраиваемым функциям](custom-functions-best-practices.md).</span><span class="sxs-lookup"><span data-stu-id="e06d5-231">See more details in [Custom functions best practices](custom-functions-best-practices.md).</span></span>

## <a name="changelog"></a><span data-ttu-id="e06d5-232">Журнал изменений</span><span class="sxs-lookup"><span data-stu-id="e06d5-232">Changelog</span></span>

- <span data-ttu-id="e06d5-233">**7 ноября 2017 г.**. Выпущена\* предварительная версия настраиваемых функций с примерами</span><span class="sxs-lookup"><span data-stu-id="e06d5-233">**Nov 7, 2017**: Shipped the custom functions preview and samples</span></span>
- <span data-ttu-id="e06d5-234">**20 ноября 2017 г.** Исправлена ошибка совместимости для пользователей, использующих сборки 8801 и выше.</span><span class="sxs-lookup"><span data-stu-id="e06d5-234">**Nov 20, 2017**: Fixed compatibility bug for those using builds 8801 and later</span></span>
- <span data-ttu-id="e06d5-235">**28 ноября 2017 г.**. Выпущена\* поддержка отмены вызова асинхронных функций (необходимо изменение потоковых функций)</span><span class="sxs-lookup"><span data-stu-id="e06d5-235">**Nov 28, 2017**: Shipped support for cancellation on asynchronous functions (requires change for streaming functions)</span></span>
- <span data-ttu-id="e06d5-236">**7 мая 2018 г.**. Выпущена\*​​поддержка Mac, Excel Online и синхронных функций, выполняемых внутри процесса</span><span class="sxs-lookup"><span data-stu-id="e06d5-236">**May 7, 2018**: Shipped support for Mac, Excel Online, and synchronous functions running in-process</span></span>
- <span data-ttu-id="e06d5-237">**20 сентября 2018 г.**. Выпущена поддержка среды выполнения JavaScript настраиваемых функций.</span><span class="sxs-lookup"><span data-stu-id="e06d5-237">**September 20, 2018**: Shipped support for custom functions JavaScript runtime.</span></span> <span data-ttu-id="e06d5-238">Подробнее см. статью [Среда выполнения для настраиваемых функций Excel](custom-functions-runtime.md).</span><span class="sxs-lookup"><span data-stu-id="e06d5-238">For more information, see [Runtime for Excel custom functions](custom-functions-runtime.md).</span></span>

<span data-ttu-id="e06d5-239">\* на канале участников программы предварительной оценки Office</span><span class="sxs-lookup"><span data-stu-id="e06d5-239">\* to the Office Insiders Channel</span></span>

## <a name="see-also"></a><span data-ttu-id="e06d5-240">См. также</span><span class="sxs-lookup"><span data-stu-id="e06d5-240">See also</span></span>

* [<span data-ttu-id="e06d5-241">Метаданные настраиваемых функций</span><span class="sxs-lookup"><span data-stu-id="e06d5-241">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="e06d5-242">Среда выполнения для настраиваемых функций Excel</span><span class="sxs-lookup"><span data-stu-id="e06d5-242">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="e06d5-243">Рекомендации по настраиваемым функциям</span><span class="sxs-lookup"><span data-stu-id="e06d5-243">Custom functions best practices</span></span>](custom-functions-best-practices.md)
