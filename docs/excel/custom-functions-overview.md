# <a name="create-custom-functions-in-excel-preview"></a><span data-ttu-id="b2eb6-101">Создание настраиваемых функций в Excel (ознакомительная версия)</span><span class="sxs-lookup"><span data-stu-id="b2eb6-101">Create custom functions in Excel (Preview)</span></span>

<span data-ttu-id="b2eb6-102">Настраиваемые функции (как и пользовательские функции, или UDF) позволяют разработчикам добавлять любую функцию JavaScript в Excel с помощью надстройки.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-102">Custom functions (similar to user-defined functions, or UDFs), allow developers to add any JavaScript function to Excel using an add-in.</span></span> <span data-ttu-id="b2eb6-103">После этого пользователи смогут получать доступ к настраиваемым функциям, как к любой другой встроенной функции Excel (например, `=SUM()`).</span><span class="sxs-lookup"><span data-stu-id="b2eb6-103">Users can then access custom functions like any other native function in Excel (like =SUM()).</span></span> <span data-ttu-id="b2eb6-104">В этой статье описано создание настраиваемых функций в Excel.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-104">This article explains how to create custom functions in Excel.</span></span>

<span data-ttu-id="b2eb6-105">Ниже показано, как конечный пользователь вставляет настраиваемую функцию в ячейку.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-105">The following illustration shows you how an end user would insert a custom function into a cell.</span></span> <span data-ttu-id="b2eb6-106">Функция, которая добавляет 42 к паре чисел.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-106">Here’s the code for a sample custom function that adds 42 to a pair of numbers.</span></span>

<img alt="custom functions" src="../images/custom-function.gif" width="579" height="383" />

<span data-ttu-id="b2eb6-107">Ниже представлен пример кода такой настраиваемой функции.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-107">Here’s the code for the same custom function.</span></span>

```js
function ADD42(a, b) {
    return a + b + 42;
}
```

<span data-ttu-id="b2eb6-108">Настраиваемые функции теперь доступны в предварительной версии для разработчиков на Windows, Mac, а также в Excel Online.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-108">Custom functions are now available in Developer Preview on Windows, Mac, and Excel Online.</span></span> <span data-ttu-id="b2eb6-109">Чтобы опробовать их, выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-109">Follow these steps to try them:</span></span>

1.  <span data-ttu-id="b2eb6-110">Установите Office (сборка 9325 на Windows или 13.329 на Mac) и присоединитесь к программе [предварительной оценки Office](https://products.office.com/en-us/office-insider) .</span><span class="sxs-lookup"><span data-stu-id="b2eb6-110">Install Office (build 9325 on Windows or 13.329 on Mac) and join the [Office Insider](https://products.office.com/en-us/office-insider) program.</span></span> <span data-ttu-id="b2eb6-111">(Обратите внимание, что недостаточно просто установить последнюю сборку: функция будет отключена в любой сборке, пока вы не присоединитесь к программе предварительной оценки.)</span><span class="sxs-lookup"><span data-stu-id="b2eb6-111">(Note that it isn't enough just to get the latest build; the feature will be disabled on any build until you join the Insider program)</span></span>
2.  <span data-ttu-id="b2eb6-112">Клонируйте репозиторий [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) и следуйте инструкциям в файле README.md, чтобы запустить надстройку в Excel, внести изменения в код и выполнить отладку.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-112">Clone the [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) repo and follow the instructions in the README.md to start the add-in in Excel, make changes in the code, and debug.</span></span>
3.  <span data-ttu-id="b2eb6-113">Введите `=CONTOSO.ADD42(1,2)` в любую ячейку и нажмите клавишу **Ввод**, чтобы выполнить настраиваемую функцию.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-113">Type `=CONTOSO.ADD42(1,2)` into any cell, and press **Enter** to run the custom function.</span></span>

<span data-ttu-id="b2eb6-114">В разделе **Известные проблемы**  в конце этой статьи указаны текущие ограничения на настраиваемые функции, которые со временем будут обновляться.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-114">See the Known Issues section at the end of this article, which includes current limitations of custom functions and will be updated over time.</span></span>

## <a name="learn-the-basics"></a><span data-ttu-id="b2eb6-115">Основы</span><span class="sxs-lookup"><span data-stu-id="b2eb6-115">Learn the basics</span></span>

<span data-ttu-id="b2eb6-116">В клонированном примере репозитория вы увидите перечисленные ниже файлы.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-116">In the cloned sample repo, you’ll see the following files:</span></span>

- <span data-ttu-id="b2eb6-117">**customfunctions.js**, который содержит код настраиваемой функции (см. пример простого кода выше для функции `ADD42`).</span><span class="sxs-lookup"><span data-stu-id="b2eb6-117">**customfunctions.js**, which contains the custom function code (see the simple code example above for the `ADD42` function).</span></span>
- <span data-ttu-id="b2eb6-118">**customfunctions.json**, который содержит данные о регистрации в JSON, которые сообщают Excel о вашей настраиваемой функции.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-118">**customfunctions.json**, which contains the registration JSON that tells Excel about your custom function.</span></span> <span data-ttu-id="b2eb6-119">После регистрации ваши настраиваемые функции появятся в списке доступных функций, отображаемых при выполнении пользователем ввода в ячейку.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-119">Registration makes your custom functions appear in the list of available functions displayed when users type in cells.</span></span>
- <span data-ttu-id="b2eb6-120">**customfunctions.html**, который предоставляет ссылку &lt;Script&gt; на файл JS.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-120">customfunctions.html, which provides a Script reference to customfunctions.js.</span></span> <span data-ttu-id="b2eb6-121">Этот файл не отображается в пользовательском интерфейсе Excel.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-121">This file does not display UI in Excel.</span></span>
- <span data-ttu-id="b2eb6-122">**customfunctions.xml**, который сообщает Excel о местонахождении файлов HTML, JavaScript и JSON, а также определяет пространство имен для всех настраиваемых функций, которые установлены с надстройкой.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-122">**customfunctions.xml**, which tells Excel the location of the HTML, JavaScript, and JSON files; and also specifies a namespace for all the custom functions that are installed with the add-in.</span></span>

### <a name="json-file-customfunctionsjson"></a><span data-ttu-id="b2eb6-123">Файл JSON (customfunctions.json)</span><span class="sxs-lookup"><span data-stu-id="b2eb6-123">JSON file (customfunctions.json)</span></span>

<span data-ttu-id="b2eb6-124">Следующий код в файле customfunctions.json указывает метаданные для той же функции `ADD42`.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-124">The following code in customfunctions.json specifies the metadata for the same `ADD42` function.</span></span>

> [!NOTE]
> <span data-ttu-id="b2eb6-125">Подробные сведения о файле JSON, включая параметры, которые не использовались в этом примере, находится в статье [JSON-файл регистрации настраиваемой функции](https://dev.office.com/reference/add-ins/custom-functions-json).</span><span class="sxs-lookup"><span data-stu-id="b2eb6-125">Detailed reference information for the JSON file, including options not used in this example, is at [Custom Functions Registration JSON](https://dev.office.com/reference/add-ins/custom-functions-json).</span></span>

<span data-ttu-id="b2eb6-126">Обратите внимание, что для этого примера:</span><span class="sxs-lookup"><span data-stu-id="b2eb6-126">Note that for this example:</span></span>

- <span data-ttu-id="b2eb6-127">В нем присутствует только одна настраиваемая функция, поэтому элемент массива `functions` только один.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-127">There's only one custom function, so there's only one member of the `functions` array.</span></span>
- <span data-ttu-id="b2eb6-128">Свойство  `name` определяет имя функции.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-128">The `name` property defines the function name.</span></span> <span data-ttu-id="b2eb6-129">Как можно видеть на анимационном GIF, расположенном ранее, пространство имен (`CONTOSO`) добавляется к имени функции в меню автозаполнения Excel.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-129">As you see in the animated gif shown previously, a namespace (`CONTOSO`) is prepended to the function name in the Excel autocomplete menu.</span></span> <span data-ttu-id="b2eb6-130">Этот префикс определен в манифесте надстройки, описанном ниже.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-130">This prefix is defined in the add-in manifest, described below.</span></span> <span data-ttu-id="b2eb6-131">Префикс и имя функции разделяются точкой, и, согласно принятой форме записи, префиксы и имена функций указываются прописными буквами.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-131">The prefix and the function name are separated using a period, and by convention prefixes and function names are uppercase.</span></span> <span data-ttu-id="b2eb6-132">Для использования настраиваемой функции пользователь вводит в ячейку пространство имен, за которым следует имя функции (`ADD42`), в нашем случае — `=CONTOSO.ADD42`.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-132">To use your custom function, a user types the namespace followed by the function's name (`ADD42`) into a cell, in this case `=CONTOSO.ADD42`.</span></span> <span data-ttu-id="b2eb6-133">Префикс служит в качестве идентификатора вашей компании или надстройки.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-133">The prefix is intended to be used as an identifier for your add-in.</span></span> 
- <span data-ttu-id="b2eb6-134">`description` отображается в меню автозаполнения в Excel.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-134">`description`: The description appears in the autocomplete menu in Excel.</span></span>
- <span data-ttu-id="b2eb6-135">Когда пользователь запрашивает справку по функции, Excel открывает область задач, в которой отображается веб-страница, расположенная по URL-адресу, который указан в `helpUrl`.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-135">`helpUrl`: When the user requests help for a function, Excel opens a task pane and displays the web page found at this URL.</span></span>
- <span data-ttu-id="b2eb6-136">Свойство  `result` задает тип данных, возвращаемых функцией в Excel.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-136">`result`: Defines the type of information returned by the function to Excel.</span></span> <span data-ttu-id="b2eb6-137">Дочернее свойство `type` может быть типа `"string"`, `"number"` или `"boolean"`.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-137">The `type` child property can `"string"`, `"number"`, or `"boolean"`.</span></span> <span data-ttu-id="b2eb6-138">Свойство `dimensionality` может быть `scalar` или `matrix` (двумерный массив значений указанного типа `type`.)</span><span class="sxs-lookup"><span data-stu-id="b2eb6-138">The `dimensionality` property can be `scalar` or `matrix` (a two-dimensional array of values of the specified `type`.)</span></span>
- <span data-ttu-id="b2eb6-139">Массив `parameters` указывает *по порядку* тип данных в каждом параметре, который передается функции.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-139">The `parameters` array specifies, *in order*, the type of data in each parameter that is passed to the function.</span></span> <span data-ttu-id="b2eb6-140">Дочерние свойства `name` и `description` используются в Excel intellisense.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-140">The `name` and `description` child properties are used in the Excel intellisense.</span></span> <span data-ttu-id="b2eb6-141">Дочерние свойства `type` и `dimensionality`  идентичны дочерним свойствам родительского `result`, описанного выше.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-141">The `type` and `dimensionality` child properties are identical to the child properties of the `result` property described above.</span></span>
- <span data-ttu-id="b2eb6-142">Свойство  `options` позволяет настраивать некоторые аспекты того, как и когда Excel выполняет эту функцию.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-142">The `options` property enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="b2eb6-143">Более подробные сведения об этих параметрах приведены далее в этой статье.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-143">There is more information about these options later in this article.</span></span>

 ```js
{
    "$schema": "https://developer.microsoft.com/en-us/json-schemas/office-js/custom-functions.schema.json",
    "functions": [
        {
            "name": "ADD42", 
            "description":  "adds 42 to the input numbers",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "number",
                "dimensionality": "scalar"
            },
            "parameters": [
                {
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
            "options": {
                "sync": true
            }
        }
    ]
}
```

> [!NOTE]
> <span data-ttu-id="b2eb6-144">Настраиваемые функции регистрируются тогда, когда пользователь запускает надстройку впервые.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-144">The custom functions are registered when a user runs the add-in for the first time.</span></span> <span data-ttu-id="b2eb6-145">После этого они доступны для этого же пользователя во всех книгах (а не только в той, в которой надстройка запускалась первоначально).</span><span class="sxs-lookup"><span data-stu-id="b2eb6-145">After that, they are available, for that same user, in all workbooks (not only the one where the add-in ran initially.)</span></span>

<span data-ttu-id="b2eb6-146">Чтобы настраиваемая функция работала корректно в Excel Online, в параметрах сервера для файла JSON должен быть включен [CORS](https://developer.mozilla.org/en-US/docs/Web/HTTP/CORS).</span><span class="sxs-lookup"><span data-stu-id="b2eb6-146">Your server settings for the JSON file must have [CORS](https://developer.mozilla.org/en-US/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>


### <a name="manifest-file-customfunctionsxml"></a><span data-ttu-id="b2eb6-147">Файл манифеста (customfunctions.xml)</span><span class="sxs-lookup"><span data-stu-id="b2eb6-147">Manifest file (customfunctions.xml)</span></span>


<span data-ttu-id="b2eb6-148">Ниже приведен пример элементов разметки `<ExtensionPoint>` и `<Resources>`, которые необходимо включить в манифест надстройки, чтобы позволить Excel выполнять функции.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-148">The following is an example of the `<ExtensionPoint>` and `<Resources>` markup that you include in the add-in's manifest to enable Excel to run your functions.</span></span> <span data-ttu-id="b2eb6-149">Обратите внимание на следующие особенности этой разметки:</span><span class="sxs-lookup"><span data-stu-id="b2eb6-149">Note the following facts about this markup:</span></span>

- <span data-ttu-id="b2eb6-150">Элемент `<Script>` и его соответствующий идентификатор ресурса определяет расположение файла JavaScript с вашими функциями.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-150">The `<Script>` element and its corresponding resource ID specifies the location of the JavaScript file with your functions.</span></span>
- <span data-ttu-id="b2eb6-151">Элемент`<Page>` и его соответствующий ИД ресурса определяет расположение HTML-страницы вашей надстройки.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-151">The `<Page>` element and its corresponding resource ID specifies the location of the HTML page of your add-in.</span></span> <span data-ttu-id="b2eb6-152">HTML-страница содержит тег `<Script>`, который загружает файл JavaScript (customfunctions.js).</span><span class="sxs-lookup"><span data-stu-id="b2eb6-152">The HTML page includes a `<Script>` tag that loads the JavaScript file (customfunctions.js).</span></span> <span data-ttu-id="b2eb6-153">HTML-страница скрыта и никогда не отображается в пользовательском интерфейсе.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-153">The HTML page is a hidden page and is never displayed in the UI.</span></span>
- <span data-ttu-id="b2eb6-154">Элемент `<Metadata>` и его соответствующий идентификатор ресурса определяет расположение JSON-файла.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-154">The `<Metadata>` element and its corresponding resource ID specifies the location of the JSON file.</span></span>
- <span data-ttu-id="b2eb6-155">Элемент`<Namespace>` и его соответствующий ИД ресурса определяет префикс для настраиваемой функции в надстройке.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-155">A `<Namespace>` element and its corresponding resource ID specifies the prefix for all custom functions in the add-in.</span></span>


```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1\_0">
    <Hosts>
        <Host xsi:type="Workbook">
            <AllFormFactors>
                <ExtensionPoint xsi:type="CustomFunctions">
                    <Script>
                        <SourceLocation resid="residjs" />
                    </Script>
                    <Page>
                        <SourceLocation resid="residhtml"/>
                    </Page>
                    <Metadata>
                        <SourceLocation resid="residjson" />
                    </Metadata>
                    <Namespace resid="residNS" />
                </ExtensionPoint>
            </AllFormFactors>
        </Host>
    </Hosts>
    <Resources>
        <bt:Urls>
            <bt:Url id="residjson" DefaultValue="http://127.0.0.1:8080/customfunctions.json" />
            <bt:Url id="residjs" DefaultValue="http://127.0.0.1:8080/customfunctions.js" />
            <bt:Url id="residhtml" DefaultValue="http://127.0.0.1:8080/customfunctions.html" />
        </bt:Urls>
        <bt:ShortStrings>
            <bt:String id="residNS" DefaultValue="CONTOSO" />
        </bt:ShortStrings>
    </Resources>
</VersionOverrides>

```

## <a name="initializing-custom-functions"></a><span data-ttu-id="b2eb6-156">Инициализация настраиваемых функций</span><span class="sxs-lookup"><span data-stu-id="b2eb6-156">Initializing custom functions</span></span>

<span data-ttu-id="b2eb6-157">Ваш код должен инициализировать настраиваемые функции перед их использованием.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-157">Your code must initialize the custom functions feature before using it.</span></span> <span data-ttu-id="b2eb6-158">Вы можете сделать это либо в теге &lt;Script&gt; в файле HTML (customfunctions.html), либо в начале файла JavaScript (customfunctions.js).</span><span class="sxs-lookup"><span data-stu-id="b2eb6-158">You can do this either in a &lt;Script&gt; tag in the HTML file (customfunctions.html) or at the top of the JavaScript file (customfunctions.js).</span></span> <span data-ttu-id="b2eb6-159">При использовании предварительной версии настраиваемых функций у вас есть выбор из двух вариантов синтаксиса инициализации.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-159">During the preview of custom functions, you have your choice of two syntaxes for intializing.</span></span> <span data-ttu-id="b2eb6-160">HTML-файл в репозитории использует следующий синтаксис:</span><span class="sxs-lookup"><span data-stu-id="b2eb6-160">The HTML file in the repo uses the following syntax:</span></span>

```js
Office.initialize = function (reason) {
    return Excel.CustomFunctions.initialize();
};
```

<span data-ttu-id="b2eb6-161">Также можно использовать следующий синтаксис:</span><span class="sxs-lookup"><span data-stu-id="b2eb6-161">You can also use the following syntax:</span></span>

```js
Office.Preview.StartCustomFunctions();
```

## <a name="synchronous-and-asynchronous-functions"></a><span data-ttu-id="b2eb6-162">Синхронные и асинхронные функции</span><span class="sxs-lookup"><span data-stu-id="b2eb6-162">Synchronous and asynchronous functions</span></span>

<span data-ttu-id="b2eb6-163">Функция `ADD42`, представленная выше, является синхронной относительно Excel (обозначается установкой параметра `"sync": true` в JSON-файле).</span><span class="sxs-lookup"><span data-stu-id="b2eb6-163">The function `ADD42` above is synchronous with respect to Excel (designated by setting the option `"sync": true` in the JSON file).</span></span> <span data-ttu-id="b2eb6-164">Синхронные функции обеспечивают высокую производительность, поскольку запускаются в том же процессе, что и Excel, и работают параллельно при многопоточном вычислении.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-164">Synchronous functions offer fast performance because they run in the same process as Excel and they run in parallel during multithreaded calculation.</span></span>   

<span data-ttu-id="b2eb6-165">С другой стороны, если ваша настраиваемая функция извлекает данные из Интернета, она должна быть асинхронной относительно Excel.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-165">On the other hand, if your custom function retrieves data from the web, it must be asynchronous with respect to Excel.</span></span> <span data-ttu-id="b2eb6-166">Асинхронные функции должны:</span><span class="sxs-lookup"><span data-stu-id="b2eb6-166">Asynchronous functions must:</span></span>

1. <span data-ttu-id="b2eb6-167">Возвращать JavaScript Promise в Excel.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-167">Return a JavaScript Promise to Excel.</span></span>
3. <span data-ttu-id="b2eb6-168">Разрешать Promise окончательным значением, используя функцию обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-168">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="b2eb6-169">В приведенном ниже коде показан пример асинхронной настраиваемой функции, получающей температуру термометра.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-169">The following code shows an example of a custom function that retrieves the temperature of a thermometer.</span></span> <span data-ttu-id="b2eb6-170">Обратите внимание, что функция `sendWebRequest` является гипотетической, не указанной здесь, и использует XHR для вызова веб-службы температуры.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-170">Note that `sendWebRequest` is a hypothetical function, not specified here, that uses XHR to call a temperature web service.</span></span>

```js
function getTemperature(thermometerID){
    return new OfficeExtension.Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

<span data-ttu-id="b2eb6-171">Асинхронные функции отображают временную ошибку `GETTING_DATA`  в ячейке, пока Excel ждет окончательный результат.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-171">Asynchronous functions display a `GETTING_DATA` temporary error in the cell while Excel waits for the final result.</span></span> <span data-ttu-id="b2eb6-172">Во время ожидания результата пользователи могут нормально взаимодействовать с остальной частью электронной таблицы.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-172">Users can interact normally with the rest of the spreadsheet while they wait for the result.</span></span>

> [!NOTE]
> <span data-ttu-id="b2eb6-173">По умолчанию настраиваемые функции асинхронны.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-173">Custom functions are asynchronous by default.</span></span> <span data-ttu-id="b2eb6-174">Чтобы сделать функции синхронными, установите параметр `"sync": true` для свойства `options` для настраиваемой функции в JSON-файле регистрации.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-174">To designate functions as synchronous set the option `"sync": true` in the `options` property for the custom function in the registration JSON file.</span></span>

## <a name="streamed-functions"></a><span data-ttu-id="b2eb6-175">Потоковые функции</span><span class="sxs-lookup"><span data-stu-id="b2eb6-175">Streamed functions</span></span>

<span data-ttu-id="b2eb6-176">Асинхронная функция может быть потоковой.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-176">An asynchronous function can be streamed.</span></span> <span data-ttu-id="b2eb6-177">С помощью потоковых настраиваемых функций вы можете многократно выводить данные в ячейки, не дожидаясь, пока Excel или пользователь запросят повторное вычисление.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-177">Streamed custom functions let you output data to cells repeatedly over time, without waiting for Excel or users to request recalculations.</span></span> <span data-ttu-id="b2eb6-178">Следующий пример - это настраиваемая функция, которая добавляет число к результату каждую секунду.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-178">The following example is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="b2eb6-179">Обратите внимание на особенности этого кода:</span><span class="sxs-lookup"><span data-stu-id="b2eb6-179">Note the following about this code:</span></span>

- <span data-ttu-id="b2eb6-180">Excel автоматически отображает каждое новое значение при помощи `setResult` обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-180">Excel displays each new value automatically using the `setResult` callback.</span></span>
- <span data-ttu-id="b2eb6-181">Последний параметр, `caller`, никогда не указывается в коде регистрации и не отображается в меню автозаполнения, когда пользователи Excel вводят функцию.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-181">For streamed functions, the final parameter, `caller`, is never specified in your registration code, and it does not display in the autocomplete menu to Excel users when they enter the function.</span></span> <span data-ttu-id="b2eb6-182">Это объект, который содержит функцию обратного вызова `setResult`, используемую для передачи данных из функции в Excel и обновления значения ячейки.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-182">It’s an object that contains a `setResult` callback function that’s used to pass data from the function to Excel to update the value of a cell.</span></span>
- <span data-ttu-id="b2eb6-183">Чтобы Excel передал функцию `setResult` объекту `caller`, необходимо объявить поддержку потоковой передачи при регистрации функции, установив параметр `"stream": true` для свойства `options` для настраиваемой функции в JSON-файле регистрации.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-183">In order for Excel to pass the `setResult` function in the `caller` object, you must declare support for streaming during your function registration by setting the option `"stream": true` in the `options` property for the custom function in the registration JSON file.</span></span>

```js
function incrementValue(increment, caller){
    var result = 0;
    setInterval(function(){
         result += increment;
         caller.setResult(result);
    }, 1000);
}
```

## <a name="cancellation"></a><span data-ttu-id="b2eb6-184">Отмена</span><span class="sxs-lookup"><span data-stu-id="b2eb6-184">Cancellation</span></span>

<span data-ttu-id="b2eb6-185">Вы можете отменять вызовы потоковых и асинхронных функций.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-185">You can cancel streamed functions and asynchronous functions.</span></span> <span data-ttu-id="b2eb6-186">Отмена вызова функций позволяет снизить потребление пропускной способности, использование рабочей памяти и нагрузку на ЦП.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-186">Canceling your function calls is important to reduce their bandwith consumption, working memory, and CPU load.</span></span> <span data-ttu-id="b2eb6-187">Excel отменяет вызовы функций в следующих случаях:</span><span class="sxs-lookup"><span data-stu-id="b2eb6-187">Excel cancels function calls in the following situations:</span></span>

- <span data-ttu-id="b2eb6-188">Пользователь редактирует или удаляет ячейку, ссылающуюся на функцию.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-188">The user edits or deletes a cell that references the function.</span></span>
- <span data-ttu-id="b2eb6-189">Изменился один из аргументов (входных параметров) функции.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-189">One of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="b2eb6-190">В этом случае помимо отмены выполняется новый вызов функции.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-190">In this case, a new function call is triggered in addition to the cancelation.</span></span>
- <span data-ttu-id="b2eb6-p124">Пользователь активирует пересчет вручную. Как и в вышеописанном случае, помимо отмены выполняется новый вызов функции.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-p124">The user triggers recalculation manually. As with the above case, a new function call is triggered in addition to the cancelation.</span></span>

<span data-ttu-id="b2eb6-193">Вы *должны* реализовать обработчик отмены для каждой функции потоковой передачи.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-193">You *must* implement a cancellation handler for every streaming function.</span></span> <span data-ttu-id="b2eb6-194">Асинхронные, непотоковые функции могут подлежать или не подлежать отмене по вашему усмотрению.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-194">Asynchronous, non-streaming functions may or may not be cancelable; it's up to you.</span></span> <span data-ttu-id="b2eb6-195">Синхронные функции отмене не подлежат.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-195">Synchronous functions cannot be canceled.</span></span>

<span data-ttu-id="b2eb6-196">Чтобы сделать функцию отменяемой, установите параметр `"cancelable": true` для свойства `options` для настраиваемой функции в JSON-файле регистрации.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-196">To make a function cancelable, set the option `"cancelable": true` in the `options` property for the custom function in the registration JSON file.</span></span>

<span data-ttu-id="b2eb6-197">Ниже показан код из предыдущего примера с реализованной отменой.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-197">The following code shows the previous example with cancellation implemented.</span></span> <span data-ttu-id="b2eb6-198">Объект `caller` в коде содержит функцию `onCanceled`, которую необходимо определить для каждой отменяемой настраиваемой функции.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-198">In the code, the `caller` object contains an `onCanceled` function which should be defined for each custom function.</span></span>

```js
function incrementValue(increment, caller){ 
    var result = 0;
    var timer = setInterval(function(){
         result += increment;
         caller.setResult(result);
    }, 1000);

    caller.onCanceled = function(){
        clearInterval(timer);
    }
}
```

## <a name="saving-and-sharing-state"></a><span data-ttu-id="b2eb6-199">Сохранение и передача состояния</span><span class="sxs-lookup"><span data-stu-id="b2eb6-199">Saving and sharing state</span></span>

<span data-ttu-id="b2eb6-200">Асинхронные функции могут сохранять данные в глобальных переменных JavaScript.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-200">Custom functions can save data in global JavaScript variables.</span></span> <span data-ttu-id="b2eb6-201">В последующих вызовах настраиваемая функция может использовать значения, сохраненные в этих переменных.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-201">In subsequent calls, your custom function may use the values saved in these variables.</span></span> <span data-ttu-id="b2eb6-202">Сохранение состояния может быть полезно, когда пользователи добавляют одну настраиваемую функцию к нескольким ячейкам, потому что все экземпляры функции могут совместно использовать ее состояние.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-202">Saved state is useful when users add the same custom function to more than one cell, because all the instances of the function can share the state.</span></span> <span data-ttu-id="b2eb6-203">Например, вы можете сохранить данные, возвращенные при вызове веб-ресурса, чтобы не пришлось обеспечивать выполнение дополнительных вызовов.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-203">For example, you may save the data returned from a call to a web resource to avoid making additional calls to the same web resource.</span></span>

<span data-ttu-id="b2eb6-204">В приведенном ниже коде показана реализация вышеописанной функции передачи температуры, глобально сохраняющей состояние.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-204">The following code shows an implementation of the previous temperature-streaming function that saves state using the  variable.</span></span> <span data-ttu-id="b2eb6-205">Обратите внимание на особенности этого кода:</span><span class="sxs-lookup"><span data-stu-id="b2eb6-205">Note the following about this code:</span></span>

- <span data-ttu-id="b2eb6-206">`refreshTemperature` это потоковая функция, ежесекундно считывающая температуру определенного термометра.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-206">`refreshTemperature` is a streamed function that reads the temperature of a particular thermometer every second.</span></span> <span data-ttu-id="b2eb6-207">Новые температуры сохраняются в переменную `savedTemperatures`, но не обновляют значение ячейки напрямую.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-207">New temperatures are saved in the `savedTemperatures` variable, but does not directly update the cell value.</span></span> <span data-ttu-id="b2eb6-208">Она не должен вызываться непосредственно из ячейки листа, *поэтому она не регистрируется в файле JSON*.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-208">It should not be directly called from a worksheet cell, *so it is not registered in the JSON file*.</span></span>
- <span data-ttu-id="b2eb6-209">`streamTemperature` обновляет значения температуры, которые отображаются в ячейке каждую секунду, а в качестве источника данных использует переменную `savedTemperatures`.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-209">`streamTemperature` updates the temperature values displayed in the cell every second and it uses `savedTemperatures` variable as its data source.</span></span> <span data-ttu-id="b2eb6-210">Она должна быть зарегистрирована в файле JSON и записана прописными буквами: `STREAMTEMPERATURE`.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-210">It must be registered in the JSON file, and named with all upper-case letters, `STREAMTEMPERATURE`.</span></span>
- <span data-ttu-id="b2eb6-211">Пользователи могут вызывать функцию `streamTemperature` из нескольких ячеек в пользовательском интерфейсе Excel.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-211">Users may call `streamTemperature` from several cells in the Excel UI.</span></span> <span data-ttu-id="b2eb6-212">Каждый вызов считывает данные из той же переменной `savedTemperatures`.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-212">Each call reads data from the same `savedTemperatures` variable.</span></span>

```js
var savedTemperatures;

function streamTemperature(thermometerID, caller){ 
     if(!savedTemperatures[thermometerID]){
         refreshTemperatures(thermometerID); // starts fetching temperatures if the thermometer hasn't been read yet
     }

     function getNextTemperature(){
         caller.setResult(savedTemperatures[thermometerID]); // setResult sends the saved temperature value to Excel.
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

> [!NOTE]
> <span data-ttu-id="b2eb6-213">Синхронные функции (назначаются путем установки параметра `"sync": true` в файле JSON) не могут передавать состояние, потому что Excel распараллеливает их во время многопоточного вычисления.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-213">Synchronous functions (designated by setting the option `"sync": true` in the JSON file) cannot share state because Excel parallelizes them during multithreaded calculation.</span></span> <span data-ttu-id="b2eb6-214">Только асинхронные функции могут передавать состояние, поскольку синхронные функции надстройки используют один контекст JavaScript в каждом сеансе.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-214">Only asynchronous functions may share state because an add-in's synchronous functions share the same JavaScript context in each session.</span></span>

## <a name="working-with-ranges-of-data"></a><span data-ttu-id="b2eb6-215">Работа с диапазонами данных</span><span class="sxs-lookup"><span data-stu-id="b2eb6-215">Working with ranges of data</span></span>

<span data-ttu-id="b2eb6-216">Настраиваемая функция может принимать диапазон данных в качестве параметра или возвращать диапазон данных.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-216">Your custom function can take a range of data as a parameter, or you can return a range of data from a custom function.</span></span>

<span data-ttu-id="b2eb6-217">Например, предположим, что ваша функция возвращает второе наивысшее значение из диапазона чисел, хранящихся в Excel.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-217">For example, suppose that your function returns the second highest temperature from a range of temperature values stored in Excel.</span></span> <span data-ttu-id="b2eb6-218">Приведенная ниже функция принимает параметр `values`, относящийся к типу `Excel.CustomFunctionDimensionality.matrix`.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-218">The following function takes the parameter `values`, which is an `Excel.CustomFunctionDimensionality.matrix` parameter type.</span></span> <span data-ttu-id="b2eb6-219">Обратите внимание, что в JSON-регистрации для этой функции необходимо для параметра `type` установить значение `matrix`.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-219">Note that in the registration JSON for this function, you would set the parameter's `type` property to `matrix`.</span></span>

```js
function secondHighest(values){ 
     var highest = values[0][0], secondHighest = values[0][0];
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

<span data-ttu-id="b2eb6-220">Можно заметить, что диапазоны обрабатываются в JavaScript как массивы строк (двумерный массив).</span><span class="sxs-lookup"><span data-stu-id="b2eb6-220">As you can see, ranges are handled in JavaScript as arrays of row arrays (like a 2-dimensional array).</span></span>

## <a name="known-issues"></a><span data-ttu-id="b2eb6-221">Известные проблемы</span><span class="sxs-lookup"><span data-stu-id="b2eb6-221">Known issues</span></span>

- <span data-ttu-id="b2eb6-222">URL-адреса справки и описания параметров пока не используются в Excel.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-222">Help URLs and parameter descriptions are not yet used by Excel.</span></span>
- <span data-ttu-id="b2eb6-223">Настраиваемые функции в настоящее время недоступны в Excel для мобильных клиентов.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-223">Custom functions are not currently available on Excel for mobile clients.</span></span>
- <span data-ttu-id="b2eb6-224">В настоящее время надстройки используют скрытый процесс веб-обозревателя для выполнения асинхронных настраиваемых функций.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-224">Currently, add-ins rely on a hidden browser process to run custom functions.</span></span> <span data-ttu-id="b2eb6-225">В будущем JavaScript будет работать на некоторых платформах напрямую, чтобы настраиваемые функции выполнялись быстрее и использовали меньше памяти.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-225">In the future, JavaScript will run directly on some platforms to ensure custom functions are faster and use less memory.</span></span> <span data-ttu-id="b2eb6-226">Кроме того, HTML-страница, на которую ссылается элемент `<Page>` манифеста, не будет необходима для большинства платформ, так как Excel будет выполнять код JavaScript напрямую.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-226">Additionally, the HTML page referenced by the `<Page>`Page element in the manifest won’t be needed for most platforms because Excel will run the JavaScript directly.</span></span> <span data-ttu-id="b2eb6-227">Чтобы подготовиться к этому изменению, убедитесь, что в ваших настраиваемых функциях не используется модель DOM для веб-страниц.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-227">To prepare for this change, ensure your custom functions do not use the webpage DOM.</span></span> <span data-ttu-id="b2eb6-228">Поддерживаемые основные API приложения для доступа к Интернету будут [WebSocket](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) и [XHR](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest) с использованием GET или POST.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-228">The supported host APIs for accessing the web will be [WebSocket](https://developer.mozilla.org/en-US/docs/Web/API/WebSockets_API) and [XHR](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest) using GET or POST.</span></span>
- <span data-ttu-id="b2eb6-229">Изменяемые функции (которые пересчитываются автоматически, когда в электронной таблице изменяются несвязанные данных) еще не поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-229">Volatile functions (those which recalculate automatically whenever unrelated data changes in the spreadsheet) are not yet supported.</span></span>
- <span data-ttu-id="b2eb6-230">Отладка включена только для асинхронных функций в Excel для Windows.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-230">Debugging is only enabled for asynchronous functions on Excel for Windows.</span></span>
- <span data-ttu-id="b2eb6-231">Развертывание через Портал администрирования Office 365 и AppSource еще не включены.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-231">Deployment via the Office 365 Admin Portal and AppSource are not yet enabled.</span></span>
- <span data-ttu-id="b2eb6-232">Настраиваемые функции в Excel Online могут перестать работать во время сеанса после периода бездействия.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-232">Custom functions in Excel Online may stop working during a session after a period of inactivity.</span></span> <span data-ttu-id="b2eb6-233">Для восстановления работы обновите страницу браузера (F5) и повторно введите настраиваемую функцию.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-233">Refresh the browser page (F5) and re-enter a custom function to restore the feature.</span></span>

## <a name="changelog"></a><span data-ttu-id="b2eb6-234">Журнал изменений</span><span class="sxs-lookup"><span data-stu-id="b2eb6-234">Changelog</span></span>

- <span data-ttu-id="b2eb6-235">**7 ноября 2017 г.** Выпущена ознакомительная версия пользовательских функций с примерами.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-235">**Nov 7, 2017**: Shipped the custom functions preview and samples</span></span>
- <span data-ttu-id="b2eb6-236">**20 ноября 2017 г.** Исправлена ошибка совместимости для пользователей, использующих сборки 8801 и выше.</span><span class="sxs-lookup"><span data-stu-id="b2eb6-236">**Nov 20, 2017**: Fixed compatibility bug for those using builds 8801 and later</span></span>
- <span data-ttu-id="b2eb6-237">**28 ноября 2017 г.** Добавлена поддержка отмены вызова асинхронных функций (необходимо изменение для потоковых функций).</span><span class="sxs-lookup"><span data-stu-id="b2eb6-237">**Nov 28, 2017**: Shipped support for cancellation on asynchronous functions (requires change for streaming functions)</span></span>
- <span data-ttu-id="b2eb6-238">**7 мая 2018 г.** Добавлена ​​поддержка Mac, Excel Online и синхронных функций, выполняемых в процессе</span><span class="sxs-lookup"><span data-stu-id="b2eb6-238">**May 7, 2018**: Shipped support for Mac, Excel Online, and synchronous functions running in-process</span></span>
