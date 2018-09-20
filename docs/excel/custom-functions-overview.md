# <a name="create-custom-functions-in-excel-preview"></a><span data-ttu-id="2d1be-101">Создание настраиваемых функций в Excel (ознакомительная версия)</span><span class="sxs-lookup"><span data-stu-id="2d1be-101">Create custom functions in Excel (Preview)</span></span>

<span data-ttu-id="2d1be-102">Настраиваемые функции (подобные пользовательским функциям или UDF) позволяют разработчикам добавлять любую функцию JavaScript в Excel с помощью надстройки.</span><span class="sxs-lookup"><span data-stu-id="2d1be-102">Custom functions (similar to user-defined functions, or UDFs), allow developers to add any JavaScript function to Excel using an add-in.</span></span> <span data-ttu-id="2d1be-103">После этого пользователи смогут получать доступ к настраиваемым функциям, как к любой другой встроенной функции Excel (например, `=SUM()`).</span><span class="sxs-lookup"><span data-stu-id="2d1be-103">Users can then access custom functions like any other native function in Excel (like =SUM()).</span></span> <span data-ttu-id="2d1be-104">В этой статье описано создание специальных функций в Excel.</span><span class="sxs-lookup"><span data-stu-id="2d1be-104">This article explains how to create custom functions in Excel.</span></span>

<span data-ttu-id="2d1be-105">Ниже показано, как конечный пользователь вставляет настраиваемую функцию в ячейку.</span><span class="sxs-lookup"><span data-stu-id="2d1be-105">The following illustration shows you how an end user would insert a custom function into a cell.</span></span> <span data-ttu-id="2d1be-106">Функция, которая добавляет 42 к паре чисел.</span><span class="sxs-lookup"><span data-stu-id="2d1be-106">Here’s the code for a sample custom function that adds 42 to a pair of numbers.</span></span>

<img alt="custom functions" src="../images/custom-function.gif" width="579" height="383" />

<span data-ttu-id="2d1be-107">Ниже представлен пример кода такой настраиваемой функции.</span><span class="sxs-lookup"><span data-stu-id="2d1be-107">Here’s the code for the same custom function.</span></span>

```js
function ADD42(a, b) {
    return a + b + 42;
}
```

<span data-ttu-id="2d1be-108">Настраиваемые функции теперь доступны в предварительной версии для разработчиков на Windows, Mac, а также в Excel Online.</span><span class="sxs-lookup"><span data-stu-id="2d1be-108">Custom functions are now available in Developer Preview on Windows, Mac, and Excel Online.</span></span> <span data-ttu-id="2d1be-109">Чтобы опробовать их, выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="2d1be-109">Follow these steps to try them:</span></span>

1. <span data-ttu-id="2d1be-110">Установите Office (сборка 9325 на Windows или 13.329 на Mac) и присоединитесь к программе [предварительной оценки Office](https://products.office.com/office-insider) .</span><span class="sxs-lookup"><span data-stu-id="2d1be-110">Install Office (build 9325 on Windows or 13.329 on Mac) and join the [Office Insider](https://products.office.com/office-insider) program.</span></span> <span data-ttu-id="2d1be-111">(Обратите внимание, что недостаточно просто установить последнюю сборку: функция будет отключена в любой сборке, пока вы не присоединитесь к программе предварительной оценки.)</span><span class="sxs-lookup"><span data-stu-id="2d1be-111">(Note that it isn't enough just to get the latest build; the feature will be disabled on any build until you join the Insider program)</span></span>
2. <span data-ttu-id="2d1be-112">Создайте проект надстройки настраиваемой функции Excel с помощью [Yo Office](https://github.com/OfficeDev/generator-office) и следуйте инструкциям в [проекте README.md](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/README.md) для запуска надстройки в Excel, внесите изменения в код и выполните отладку.</span><span class="sxs-lookup"><span data-stu-id="2d1be-112">Create an Excel Custom Functions Add-in project using [Yo Office](https://github.com/OfficeDev/generator-office), and follow the instructions in the [project README.md](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/README.md) to start the add-in in Excel, make changes in the code, and debug.</span></span>
3. <span data-ttu-id="2d1be-113">Введите `=CONTOSO.ADD42(1,2)` в любой ячейке и нажмите клавишу **ВВОД**, чтобы выполнить специальную функцию.</span><span class="sxs-lookup"><span data-stu-id="2d1be-113">Type `=CONTOSO.ADD42(1,2)` into any cell, and press **Enter** to run the custom function.</span></span>

<span data-ttu-id="2d1be-114">В конце статьи раздела **Известные проблемы** указаны текущие ограничения на настраиваемые функции, которые со временем будут обновляться.</span><span class="sxs-lookup"><span data-stu-id="2d1be-114">See the Known Issues section at the end of this article, which includes current limitations of custom functions and will be updated over time.</span></span>

## <a name="learn-the-basics"></a><span data-ttu-id="2d1be-115">Основы</span><span class="sxs-lookup"><span data-stu-id="2d1be-115">Learn the basics</span></span>

<span data-ttu-id="2d1be-116">В клонированном примере репозитория вы увидите перечисленные ниже файлы.</span><span class="sxs-lookup"><span data-stu-id="2d1be-116">In the cloned sample repo, you’ll see the following files:</span></span>

- <span data-ttu-id="2d1be-117">**./src/customfunctions.js**, который содержит код настраиваемой функции (см. пример простого кода выше для функции `ADD42`).</span><span class="sxs-lookup"><span data-stu-id="2d1be-117">**customfunctions.js**, which contains the custom function code (see the simple code example above for the `ADD42` function).</span></span>
- <span data-ttu-id="2d1be-118">**./config/customfunctions.json**, который содержит данные о регистрации в JSON, которые сообщают Excel о вашей настраиваемой функции.</span><span class="sxs-lookup"><span data-stu-id="2d1be-118">**customfunctions.json**, which contains the registration JSON that tells Excel about your custom function.</span></span> <span data-ttu-id="2d1be-119">После регистрации ваши настраиваемые функции появятся в списке доступных функций, отображаемых при выполнении пользователем ввода в ячейку.</span><span class="sxs-lookup"><span data-stu-id="2d1be-119">Registration makes your custom functions appear in the list of available functions displayed when users type in cells.</span></span>
- <span data-ttu-id="2d1be-120">**./index.html**, который предоставляет ссылку &lt;Script&gt; на файл JS.</span><span class="sxs-lookup"><span data-stu-id="2d1be-120">**./index.html**, which provides a &lt;Script&gt; reference to the JS file.</span></span> <span data-ttu-id="2d1be-121">Этот файл не отображается в пользовательском интерфейсе Excel.</span><span class="sxs-lookup"><span data-stu-id="2d1be-121">This file does not display UI in Excel.</span></span>
- <span data-ttu-id="2d1be-122">**./manifest.xml**, который сообщает Excel расположение файлов HTML, JavaScript и JSON, а также определяет пространство имен для всех настраиваемых функций, установленных с помощью надстройки.</span><span class="sxs-lookup"><span data-stu-id="2d1be-122">**customfunctions.xml**, which tells Excel the location of the HTML, JavaScript, and JSON files; and also specifies a namespace for all the custom functions that are installed with the add-in.</span></span>

### <a name="json-file-configcustomfunctionsjson"></a><span data-ttu-id="2d1be-123">Файл JSON (. / config/customfunctions.json)</span><span class="sxs-lookup"><span data-stu-id="2d1be-123">JSON file (./config/customfunctions.json)</span></span>

<span data-ttu-id="2d1be-124">Следующий код в файле customfunctions.json указывает метаданные для той же функции `ADD42`.</span><span class="sxs-lookup"><span data-stu-id="2d1be-124">The following code in customfunctions.json specifies the metadata for the same `ADD42` function.</span></span>

> [!NOTE]
> <span data-ttu-id="2d1be-125">Подробные сведения о файле JSON, включая параметры, которые не использовались в этом примере, находится в статье [JSON-файл регистрации настраиваемой функции](custom-functions-json.md).</span><span class="sxs-lookup"><span data-stu-id="2d1be-125">Detailed reference information for the JSON file, including options not used in this example, is at [Custom Functions Registration JSON](custom-functions-json.md).</span></span>

<span data-ttu-id="2d1be-126">Обратите внимание, что для этого примера:</span><span class="sxs-lookup"><span data-stu-id="2d1be-126">Note that for this example:</span></span>

- <span data-ttu-id="2d1be-127">В нем присутствует только одна настраиваемая функция, поэтому элемент массива `functions` только один.</span><span class="sxs-lookup"><span data-stu-id="2d1be-127">There's only one custom function, so there's only one member of the `functions` array.</span></span>
- <span data-ttu-id="2d1be-128">Свойство  `name` определяет имя функции.</span><span class="sxs-lookup"><span data-stu-id="2d1be-128">The `name` property defines the function name.</span></span> <span data-ttu-id="2d1be-129">Как можно видеть на анимационном GIF, расположенном ранее, пространство имен (`CONTOSO`) добавляется к имени функции в меню автозаполнения Excel.</span><span class="sxs-lookup"><span data-stu-id="2d1be-129">As you see in the animated gif shown previously, a namespace (`CONTOSO`) is prepended to the function name in the Excel autocomplete menu.</span></span> <span data-ttu-id="2d1be-130">Этот префикс определен в манифесте надстройки, описанном ниже.</span><span class="sxs-lookup"><span data-stu-id="2d1be-130">This prefix is defined in the add-in manifest, described below.</span></span> <span data-ttu-id="2d1be-131">Префикс и имя функции разделяются точкой, и, согласно принятой форме записи, префиксы и имена функций указываются прописными буквами.</span><span class="sxs-lookup"><span data-stu-id="2d1be-131">The prefix and the function name are separated using a period, and by convention prefixes and function names are uppercase.</span></span> <span data-ttu-id="2d1be-132">Для использования настраиваемой функции пользователь вводит в ячейку пространство имен, за которым следует имя функции (`ADD42`), в нашем случае — `=CONTOSO.ADD42`.</span><span class="sxs-lookup"><span data-stu-id="2d1be-132">To use your custom function, a user types the namespace followed by the function's name (`ADD42`) into a cell, in this case `=CONTOSO.ADD42`.</span></span> <span data-ttu-id="2d1be-133">Префикс служит в качестве идентификатора вашей компании или надстройки.</span><span class="sxs-lookup"><span data-stu-id="2d1be-133">The prefix is intended to be used as an identifier for your add-in.</span></span> 
- <span data-ttu-id="2d1be-134">`description` отображается в меню автозаполнения в Excel.</span><span class="sxs-lookup"><span data-stu-id="2d1be-134">`description`: The description appears in the autocomplete menu in Excel.</span></span>
- <span data-ttu-id="2d1be-135">Когда пользователь запрашивает справку по функции, Excel открывает область задач, в которой отображается веб-страница, расположенная по URL-адресу, который указан в `helpUrl`.</span><span class="sxs-lookup"><span data-stu-id="2d1be-135">`helpUrl`: When the user requests help for a function, Excel opens a task pane and displays the web page found at this URL.</span></span>
- <span data-ttu-id="2d1be-136">Свойство `result` задает тип данных, возвращаемых функцией в Excel.</span><span class="sxs-lookup"><span data-stu-id="2d1be-136">`result`: Defines the type of information returned by the function to Excel.</span></span> <span data-ttu-id="2d1be-137">Дочернее свойство `type` может быть типа `"string"`, `"number"` или `"boolean"`.</span><span class="sxs-lookup"><span data-stu-id="2d1be-137">The `type` child property can `"string"`, `"number"`, or `"boolean"`.</span></span> <span data-ttu-id="2d1be-138">Свойство `dimensionality` может быть `scalar` или `matrix` (двумерный массив значений указанного типа `type`.)</span><span class="sxs-lookup"><span data-stu-id="2d1be-138">The `dimensionality` property can be `scalar` or `matrix` (a two-dimensional array of values of the specified `type`.)</span></span>
- <span data-ttu-id="2d1be-139">Массив `parameters` указывает *по порядку* тип данных в каждом параметре, который передается функции.</span><span class="sxs-lookup"><span data-stu-id="2d1be-139">The `parameters` array specifies, *in order*, the type of data in each parameter that is passed to the function.</span></span> <span data-ttu-id="2d1be-140">Дочерние свойства `name` и `description` используются в Excel intellisense.</span><span class="sxs-lookup"><span data-stu-id="2d1be-140">The `name` and `description` child properties are used in the Excel intellisense.</span></span> <span data-ttu-id="2d1be-141">Дочерние свойства `type` и `dimensionality`  идентичны дочерним свойствам родительского `result`, описанного выше.</span><span class="sxs-lookup"><span data-stu-id="2d1be-141">The `type` and `dimensionality` child properties are identical to the child properties of the `result` property described above.</span></span>
- <span data-ttu-id="2d1be-142">Свойство  `options` позволяет настраивать некоторые аспекты того, как и когда Excel выполняет эту функцию.</span><span class="sxs-lookup"><span data-stu-id="2d1be-142">The `options` property enables you to customize some aspects of how and when Excel executes the function.</span></span> <span data-ttu-id="2d1be-143">Более подробные сведения об этих параметрах приведены далее в этой статье.</span><span class="sxs-lookup"><span data-stu-id="2d1be-143">There is more information about these options later in this article.</span></span>

```js
    {
        "$schema": "https://developer.microsoft.com/json-schemas/office-js/custom-functions.schema.json",
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
> <span data-ttu-id="2d1be-144">Настраиваемые функции регистрируются тогда, когда пользователь запускает надстройку впервые.</span><span class="sxs-lookup"><span data-stu-id="2d1be-144">The custom functions are registered when a user runs the add-in for the first time.</span></span> <span data-ttu-id="2d1be-145">После этого они доступны для этого же пользователя во всех книгах (а не только в той, в которой надстройка запускалась первоначально).</span><span class="sxs-lookup"><span data-stu-id="2d1be-145">After that, they are available, for that same user, in all workbooks (not only the one where the add-in ran initially.)</span></span>

<span data-ttu-id="2d1be-146">Чтобы настраиваемая функция работала корректно в Excel Online, в параметрах сервера для файла JSON должен быть включен [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS).</span><span class="sxs-lookup"><span data-stu-id="2d1be-146">Your server settings for the JSON file must have [CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS) enabled in order for custom functions to work correctly in Excel Online.</span></span>


### <a name="manifest-file-manifestxml"></a><span data-ttu-id="2d1be-147">Файл манифеста (./manifest.xml)</span><span class="sxs-lookup"><span data-stu-id="2d1be-147">Manifest file (manifest.xml)</span></span>


<span data-ttu-id="2d1be-148">Ниже приведен пример элементов разметки `<ExtensionPoint>` и `<Resources>`, которые необходимо включить в манифест надстройки, чтобы позволить Excel выполнять функции.</span><span class="sxs-lookup"><span data-stu-id="2d1be-148">The following is an example of the `<ExtensionPoint>` and `<Resources>` markup that you include in the add-in's manifest to enable Excel to run your functions.</span></span> <span data-ttu-id="2d1be-149">Обратите внимание на следующие особенности этой разметки:</span><span class="sxs-lookup"><span data-stu-id="2d1be-149">Note the following facts about this markup:</span></span>

- <span data-ttu-id="2d1be-150">Элемент `<Script>` и его соответствующий идентификатор ресурса определяет расположение файла JavaScript с вашими функциями.</span><span class="sxs-lookup"><span data-stu-id="2d1be-150">The `<Script>` element and its corresponding resource ID specifies the location of the JavaScript file with your functions.</span></span>
- <span data-ttu-id="2d1be-151">Элемент`<Page>` и его соответствующий ИД ресурса определяет расположение HTML-страницы вашей надстройки.</span><span class="sxs-lookup"><span data-stu-id="2d1be-151">The `<Page>` element and its corresponding resource ID specifies the location of the HTML page of your add-in.</span></span> <span data-ttu-id="2d1be-152">HTML-страница содержит тег `<Script>`, который загружает файл JavaScript (customfunctions.js).</span><span class="sxs-lookup"><span data-stu-id="2d1be-152">The HTML page includes a `<Script>` tag that loads the JavaScript file (customfunctions.js).</span></span> <span data-ttu-id="2d1be-153">HTML-страница скрыта и никогда не отображается в пользовательском интерфейсе.</span><span class="sxs-lookup"><span data-stu-id="2d1be-153">The HTML page is a hidden page and is never displayed in the UI.</span></span>
- <span data-ttu-id="2d1be-154">Элемент `<Metadata>` и его соответствующий идентификатор ресурса определяет расположение JSON-файла.</span><span class="sxs-lookup"><span data-stu-id="2d1be-154">The `<Metadata>` element and its corresponding resource ID specifies the location of the JSON file.</span></span>
- <span data-ttu-id="2d1be-155">Элемент`<Namespace>` и его соответствующий ИД ресурса определяет префикс для настраиваемой функции в надстройке.</span><span class="sxs-lookup"><span data-stu-id="2d1be-155">A `<Namespace>` element and its corresponding resource ID specifies the prefix for all custom functions in the add-in.</span></span>


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

## <a name="initializing-custom-functions"></a><span data-ttu-id="2d1be-156">Инициализация настраиваемых функций</span><span class="sxs-lookup"><span data-stu-id="2d1be-156">Initializing custom functions</span></span>

<span data-ttu-id="2d1be-157">Ваш код должен инициализировать настраиваемые функции перед их использованием.</span><span class="sxs-lookup"><span data-stu-id="2d1be-157">Your code must initialize the custom functions feature before using it.</span></span> <span data-ttu-id="2d1be-158">Вы можете сделать это либо в теге &lt;Script&gt; в файле HTML (customfunctions.html), либо в начале файла JavaScript (customfunctions.js).</span><span class="sxs-lookup"><span data-stu-id="2d1be-158">You can do this either in a &lt;Script&gt; tag in the HTML file (customfunctions.html) or at the top of the JavaScript file (customfunctions.js).</span></span> <span data-ttu-id="2d1be-159">При использовании предварительной версии настраиваемых функций у вас есть выбор из двух вариантов синтаксиса инициализации.</span><span class="sxs-lookup"><span data-stu-id="2d1be-159">During the preview of custom functions, you have your choice of two syntaxes for intializing.</span></span> <span data-ttu-id="2d1be-160">HTML-файл в репозитории использует следующий синтаксис:</span><span class="sxs-lookup"><span data-stu-id="2d1be-160">The HTML file in the repo uses the following syntax:</span></span>

```js
Office.initialize = function (reason) {
    return Excel.CustomFunctions.initialize();
};
```

<span data-ttu-id="2d1be-161">Также можно использовать следующий синтаксис.</span><span class="sxs-lookup"><span data-stu-id="2d1be-161">You can also use the following syntax:</span></span>

```js
Office.Preview.StartCustomFunctions();
```

## <a name="handling-errors"></a><span data-ttu-id="2d1be-162">Обработка ошибок</span><span class="sxs-lookup"><span data-stu-id="2d1be-162">handling errors</span></span>
<span data-ttu-id="2d1be-163">Обработка ошибок для настраиваемых функций совпадает с [обработкой ошибок для Excel API JavaScript в целом](./excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="2d1be-163">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](./excel-add-ins-error-handling.md).</span></span> <span data-ttu-id="2d1be-164">Как правило, вы будете использовать `.catch` для обработки ошибок.</span><span class="sxs-lookup"><span data-stu-id="2d1be-164">Generally, you will use `.catch` to handle errors.</span></span> <span data-ttu-id="2d1be-165">В следующем примере кода приведен пример `.catch`.</span><span class="sxs-lookup"><span data-stu-id="2d1be-165">The code below gives an example of `.catch`.</span></span> 

```js
function getComment(x) {
    var url = "https://jsonplaceholder.typicode.com/comments/" + x; //this delivers a section of lorem ipsum from the jsonplaceholder API
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

## <a name="synchronous-and-asynchronous-functions"></a><span data-ttu-id="2d1be-166">Синхронные и асинхронные функции</span><span class="sxs-lookup"><span data-stu-id="2d1be-166">Synchronous and asynchronous functions</span></span>

<span data-ttu-id="2d1be-167">Функция `ADD42`, представленная выше, является синхронной относительно Excel (обозначается установкой параметра `"sync": true` в JSON-файле).</span><span class="sxs-lookup"><span data-stu-id="2d1be-167">The function `ADD42` above is synchronous with respect to Excel (designated by setting the option `"sync": true` in the JSON file).</span></span> <span data-ttu-id="2d1be-168">Синхронные функции обеспечивают высокую производительность, поскольку запускаются в том же процессе, что и Excel, и работают параллельно при многопоточном вычислении.</span><span class="sxs-lookup"><span data-stu-id="2d1be-168">Synchronous functions offer fast performance because they run in the same process as Excel and they run in parallel during multithreaded calculation.</span></span>   

<span data-ttu-id="2d1be-169">С другой стороны, если ваша настраиваемая функция извлекает данные из Интернета, она должна быть асинхронной относительно Excel.</span><span class="sxs-lookup"><span data-stu-id="2d1be-169">On the other hand, if your custom function retrieves data from the web, it must be asynchronous with respect to Excel.</span></span> <span data-ttu-id="2d1be-170">Асинхронные функции должны:</span><span class="sxs-lookup"><span data-stu-id="2d1be-170">Asynchronous functions must:</span></span>

1. <span data-ttu-id="2d1be-171">Возвращение обещания JavaScript в Excel.</span><span class="sxs-lookup"><span data-stu-id="2d1be-171">Return a JavaScript Promise to Excel.</span></span>
3. <span data-ttu-id="2d1be-172">Разрешать Promise окончательным значением, используя функцию обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="2d1be-172">Resolve the Promise with the final value using the callback function.</span></span>

<span data-ttu-id="2d1be-173">В приведенном ниже коде показан пример асинхронной настраиваемой функции, возвращающей температуру термометра.</span><span class="sxs-lookup"><span data-stu-id="2d1be-173">The following code shows an example of a custom function that retrieves the temperature of a thermometer.</span></span> <span data-ttu-id="2d1be-174">Обратите внимание, что функция `sendWebRequest` является гипотетической, не указанной здесь, и использует XHR для вызова веб-службы температуры.</span><span class="sxs-lookup"><span data-stu-id="2d1be-174">Note that `sendWebRequest` is a hypothetical function, not specified here, that uses XHR to call a temperature web service.</span></span>

```js
function getTemperature(thermometerID){
    return new OfficeExtension.Promise(function(setResult){
        sendWebRequest(thermometerID, function(data){
            setResult(data.temperature);
        });
    });
}
```

<span data-ttu-id="2d1be-175">Асинхронные функции отображают временную ошибку `GETTING_DATA`  в ячейке, пока Excel ждет окончательный результат.</span><span class="sxs-lookup"><span data-stu-id="2d1be-175">Asynchronous functions display a `GETTING_DATA` temporary error in the cell while Excel waits for the final result.</span></span> <span data-ttu-id="2d1be-176">Во время ожидания результата пользователи могут нормально взаимодействовать с остальной частью электронной таблицы.</span><span class="sxs-lookup"><span data-stu-id="2d1be-176">Users can interact normally with the rest of the spreadsheet while they wait for the result.</span></span>

> [!NOTE]
> <span data-ttu-id="2d1be-177">По умолчанию настраиваемые функции асинхронны.</span><span class="sxs-lookup"><span data-stu-id="2d1be-177">Custom functions are asynchronous by default.</span></span> <span data-ttu-id="2d1be-178">Чтобы сделать функции синхронными, установите параметр `"sync": true` для свойства `options` для настраиваемой функции в JSON-файле регистрации.</span><span class="sxs-lookup"><span data-stu-id="2d1be-178">To designate functions as synchronous set the option `"sync": true` in the `options` property for the custom function in the registration JSON file.</span></span>

## <a name="streamed-functions"></a><span data-ttu-id="2d1be-179">Потоковые функции</span><span class="sxs-lookup"><span data-stu-id="2d1be-179">Streamed functions</span></span>

<span data-ttu-id="2d1be-180">Асинхронная функция может быть потоковой.</span><span class="sxs-lookup"><span data-stu-id="2d1be-180">An asynchronous function can be streamed.</span></span> <span data-ttu-id="2d1be-181">С помощью потоковых настраиваемых функций вы можете многократно выводить данные в ячейки, не дожидаясь, пока Excel или пользователь запросит повторное вычисление.</span><span class="sxs-lookup"><span data-stu-id="2d1be-181">Streamed custom functions let you output data to cells repeatedly over time, without waiting for Excel or users to request recalculations.</span></span> <span data-ttu-id="2d1be-182">Следующий пример - это настраиваемая функция, которая добавляет число к результату каждую секунду.</span><span class="sxs-lookup"><span data-stu-id="2d1be-182">The following example is a custom function that adds a number to the result every second.</span></span> <span data-ttu-id="2d1be-183">Обратите внимание на следующие особенности этого кода:</span><span class="sxs-lookup"><span data-stu-id="2d1be-183">Note the following about this code:</span></span>

- <span data-ttu-id="2d1be-184">Excel автоматически отображает каждое новое значение при помощи `setResult` обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="2d1be-184">Excel displays each new value automatically using the `setResult` callback.</span></span>
- <span data-ttu-id="2d1be-185">Последний параметр, `handler`, никогда не указывается в коде регистрации и не отображается в меню автозаполнения, когда пользователи Excel вводят функцию.</span><span class="sxs-lookup"><span data-stu-id="2d1be-185">For streamed functions, the final parameter, `handler`, is never specified in your registration code, and it does not display in the autocomplete menu to Excel users when they enter the function.</span></span> <span data-ttu-id="2d1be-186">Это объект, который содержит функцию обратного вызова `setResult`, используемую для передачи данных из функции в Excel и обновления значения ячейки.</span><span class="sxs-lookup"><span data-stu-id="2d1be-186">It’s an object that contains a `setResult` callback function that’s used to pass data from the function to Excel to update the value of a cell.</span></span>
- <span data-ttu-id="2d1be-187">Чтобы Excel передал функцию `setResult` объекту `handler`, необходимо объявить поддержку потоковой передачи при регистрации функции, установив параметр `"stream": true` для свойства `options` для настраиваемой функции в JSON-файле регистрации.</span><span class="sxs-lookup"><span data-stu-id="2d1be-187">In order for Excel to pass the `setResult` function in the `handler` object, you must declare support for streaming during your function registration by setting the option `"stream": true` in the `options` property for the custom function in the registration JSON file.</span></span>

```js
function incrementValue(increment, handler){
    var result = 0;
    setInterval(function(){
         result += increment;
         handler.setResult(result);
    }, 1000);
}
```

## <a name="cancellation"></a><span data-ttu-id="2d1be-188">Отмена</span><span class="sxs-lookup"><span data-stu-id="2d1be-188">Cancellation</span></span>

<span data-ttu-id="2d1be-189">Вы можете отменять вызовы потоковых и асинхронных функций.</span><span class="sxs-lookup"><span data-stu-id="2d1be-189">You can cancel streamed functions and asynchronous functions.</span></span> <span data-ttu-id="2d1be-190">Отмена вызова функций позволяет снизить потребление пропускной способности, использование рабочей памяти и нагрузку на ЦП.</span><span class="sxs-lookup"><span data-stu-id="2d1be-190">Canceling your function calls is important to reduce their bandwith consumption, working memory, and CPU load.</span></span> <span data-ttu-id="2d1be-191">Excel отменяет вызовы функций в следующих случаях:</span><span class="sxs-lookup"><span data-stu-id="2d1be-191">Excel cancels function calls in the following situations:</span></span>

- <span data-ttu-id="2d1be-192">Пользователь редактирует или удаляет ячейку, ссылающуюся на функцию.</span><span class="sxs-lookup"><span data-stu-id="2d1be-192">The user edits or deletes a cell that references the function.</span></span>
- <span data-ttu-id="2d1be-193">Изменился один из аргументов (входных параметров) функции.</span><span class="sxs-lookup"><span data-stu-id="2d1be-193">One of the arguments (inputs) for the function changes.</span></span> <span data-ttu-id="2d1be-194">В этом случае помимо отмены выполняется новый вызов функции.</span><span class="sxs-lookup"><span data-stu-id="2d1be-194">In this case, a new function call is triggered in addition to the cancelation.</span></span>
- <span data-ttu-id="2d1be-p125">Пользователь активирует пересчет вручную. Как и в вышеописанном случае, помимо отмены выполняется новый вызов функции.</span><span class="sxs-lookup"><span data-stu-id="2d1be-p125">The user triggers recalculation manually. As with the above case, a new function call is triggered in addition to the cancelation.</span></span>

<span data-ttu-id="2d1be-197">Вы *должны* реализовать обработчик отмены для каждой функции потоковой передачи.</span><span class="sxs-lookup"><span data-stu-id="2d1be-197">You *must* implement a cancellation handler for every streaming function.</span></span> <span data-ttu-id="2d1be-198">Асинхронные, непотоковые функции могут подлежать или не подлежать отмене по вашему усмотрению.</span><span class="sxs-lookup"><span data-stu-id="2d1be-198">Asynchronous, non-streaming functions may or may not be cancelable; it's up to you.</span></span> <span data-ttu-id="2d1be-199">Синхронные функции отмене не подлежат.</span><span class="sxs-lookup"><span data-stu-id="2d1be-199">Synchronous functions cannot be canceled.</span></span>

<span data-ttu-id="2d1be-200">Чтобы сделать функцию отменяемой, установите параметр `"cancelable": true` для свойства `options` для настраиваемой функции в JSON-файле регистрации.</span><span class="sxs-lookup"><span data-stu-id="2d1be-200">To make a function cancelable, set the option `"cancelable": true` in the `options` property for the custom function in the registration JSON file.</span></span>

<span data-ttu-id="2d1be-201">Ниже показан код из предыдущего примера с реализованной отменой.</span><span class="sxs-lookup"><span data-stu-id="2d1be-201">The following code shows the previous example with cancellation implemented.</span></span> <span data-ttu-id="2d1be-202">Объект `handler` в коде содержит функцию`onCanceled`, которую необходимо определить для каждой отменяемой настраиваемой функции.</span><span class="sxs-lookup"><span data-stu-id="2d1be-202">In the code, the `handler` object contains an `onCanceled` function which should be defined for each custom function.</span></span>

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

## <a name="saving-and-sharing-state"></a><span data-ttu-id="2d1be-203">Сохранение и передача состояния</span><span class="sxs-lookup"><span data-stu-id="2d1be-203">Saving and sharing state</span></span>

<span data-ttu-id="2d1be-204">Асинхронные настраиваемые функции могут сохранять данные в глобальных переменных JavaScript.</span><span class="sxs-lookup"><span data-stu-id="2d1be-204">Custom functions can save data in global JavaScript variables.</span></span> <span data-ttu-id="2d1be-205">В последующих вызовах настраиваемая функция может использовать значения, сохраненные в этих переменных.</span><span class="sxs-lookup"><span data-stu-id="2d1be-205">In subsequent calls, your custom function may use the values saved in these variables.</span></span> <span data-ttu-id="2d1be-206">Сохранение состояния может быть полезно, когда пользователи добавляют одну настраиваемую функцию к нескольким ячейкам, потому что все экземпляры функции могут совместно использовать ее состояние.</span><span class="sxs-lookup"><span data-stu-id="2d1be-206">Saved state is useful when users add the same custom function to more than one cell, because all the instances of the function can share the state.</span></span> <span data-ttu-id="2d1be-207">Например, вы можете сохранить данные, возвращенные при вызове веб-ресурса, чтобы не пришлось обеспечивать выполнение дополнительных вызовов.</span><span class="sxs-lookup"><span data-stu-id="2d1be-207">For example, you may save the data returned from a call to a web resource to avoid making additional calls to the same web resource.</span></span>

<span data-ttu-id="2d1be-208">В приведенном ниже коде показана реализация вышеописанной функции передачи температуры, глобально сохраняющей состояние.</span><span class="sxs-lookup"><span data-stu-id="2d1be-208">The following code shows an implementation of the previous temperature-streaming function that saves state using the  variable.</span></span> <span data-ttu-id="2d1be-209">Обратите внимание на следующие особенности этого кода:</span><span class="sxs-lookup"><span data-stu-id="2d1be-209">Note the following about this code:</span></span>

- <span data-ttu-id="2d1be-210">`refreshTemperature` — это потоковая функция, ежесекундно считывающая температуру определенного термометра.</span><span class="sxs-lookup"><span data-stu-id="2d1be-210">`refreshTemperature` is a streamed function that reads the temperature of a particular thermometer every second.</span></span> <span data-ttu-id="2d1be-211">Новые температуры сохраняются в переменную `savedTemperatures`, но не обновляют значение ячейки напрямую.</span><span class="sxs-lookup"><span data-stu-id="2d1be-211">New temperatures are saved in the `savedTemperatures` variable, but does not directly update the cell value.</span></span> <span data-ttu-id="2d1be-212">Она не должен вызываться непосредственно из ячейки листа, *поэтому она не регистрируется в файле JSON*.</span><span class="sxs-lookup"><span data-stu-id="2d1be-212">It should not be directly called from a worksheet cell, *so it is not registered in the JSON file*.</span></span>
- <span data-ttu-id="2d1be-213">`streamTemperature` обновляет значения температуры, которые отображаются в ячейке каждую секунду, а в качестве источника данных использует переменную `savedTemperatures`.</span><span class="sxs-lookup"><span data-stu-id="2d1be-213">`streamTemperature` updates the temperature values displayed in the cell every second and it uses `savedTemperatures` variable as its data source.</span></span> <span data-ttu-id="2d1be-214">Она должна быть зарегистрирована в файле JSON и записана прописными буквами: `STREAMTEMPERATURE`.</span><span class="sxs-lookup"><span data-stu-id="2d1be-214">It must be registered in the JSON file, and named with all upper-case letters, `STREAMTEMPERATURE`.</span></span>
- <span data-ttu-id="2d1be-215">Пользователи могут вызывать функцию `streamTemperature` из нескольких ячеек в пользовательском интерфейсе Excel.</span><span class="sxs-lookup"><span data-stu-id="2d1be-215">Users may call `streamTemperature` from several cells in the Excel UI.</span></span> <span data-ttu-id="2d1be-216">Каждый вызов считывает данные из той же переменной `savedTemperatures`.</span><span class="sxs-lookup"><span data-stu-id="2d1be-216">Each call reads data from the same `savedTemperatures` variable.</span></span>

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

> [!NOTE]
> <span data-ttu-id="2d1be-217">Синхронные функции (назначаются путем установки параметра `"sync": true` в файле JSON) не могут передавать состояние, потому что Excel распараллеливает их во время многопоточного вычисления.</span><span class="sxs-lookup"><span data-stu-id="2d1be-217">Synchronous functions (designated by setting the option `"sync": true` in the JSON file) cannot share state because Excel parallelizes them during multithreaded calculation.</span></span> <span data-ttu-id="2d1be-218">Только асинхронные функции могут передавать состояние, поскольку синхронные функции надстройки используют один контекст JavaScript в каждом сеансе.</span><span class="sxs-lookup"><span data-stu-id="2d1be-218">Only asynchronous functions may share state because an add-in's synchronous functions share the same JavaScript context in each session.</span></span>

## <a name="working-with-ranges-of-data"></a><span data-ttu-id="2d1be-219">Работа с диапазонами данных</span><span class="sxs-lookup"><span data-stu-id="2d1be-219">Working with ranges of data</span></span>

<span data-ttu-id="2d1be-220">Настраиваемая функция может принимать диапазон данных в качестве параметра или возвращать диапазон данных.</span><span class="sxs-lookup"><span data-stu-id="2d1be-220">Your custom function can take a range of data as a parameter, or you can return a range of data from a custom function.</span></span>

<span data-ttu-id="2d1be-221">Например, предположим, что ваша функция возвращает второе наивысшее значение из диапазона чисел, хранящихся в Excel.</span><span class="sxs-lookup"><span data-stu-id="2d1be-221">For example, suppose that your function returns the second highest temperature from a range of temperature values stored in Excel.</span></span> <span data-ttu-id="2d1be-222">Приведенная ниже функция принимает параметр `values`, относящийся к типу `Excel.CustomFunctionDimensionality.matrix`.</span><span class="sxs-lookup"><span data-stu-id="2d1be-222">The following function takes the parameter `values`, which is an `Excel.CustomFunctionDimensionality.matrix` parameter type.</span></span> <span data-ttu-id="2d1be-223">Обратите внимание, что в JSON-регистрации для этой функции необходимо для параметра `type` установить значение `matrix`.</span><span class="sxs-lookup"><span data-stu-id="2d1be-223">Note that in the registration JSON for this function, you would set the parameter's `type` property to `matrix`.</span></span>

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

<span data-ttu-id="2d1be-224">Можно заметить, что диапазоны обрабатываются в JavaScript как массивы строк (двумерный массив).</span><span class="sxs-lookup"><span data-stu-id="2d1be-224">As you can see, ranges are handled in JavaScript as arrays of row arrays (like a 2-dimensional array).</span></span>

## <a name="known-issues"></a><span data-ttu-id="2d1be-225">Известные проблемы</span><span class="sxs-lookup"><span data-stu-id="2d1be-225">Known issues</span></span>

- <span data-ttu-id="2d1be-226">URL-адреса справки и описания параметров пока не используются в Excel.</span><span class="sxs-lookup"><span data-stu-id="2d1be-226">Help URLs and parameter descriptions are not yet used by Excel.</span></span>
- <span data-ttu-id="2d1be-227">Настраиваемые функции в настоящее время недоступны в Excel для мобильных клиентов.</span><span class="sxs-lookup"><span data-stu-id="2d1be-227">Custom functions are not currently available on Excel for mobile clients.</span></span>
- <span data-ttu-id="2d1be-228">В настоящее время надстройки используют скрытый процесс браузера для выполнения асинхронных настраиваемых функций.</span><span class="sxs-lookup"><span data-stu-id="2d1be-228">Currently, add-ins rely on a hidden browser process to run custom functions.</span></span> <span data-ttu-id="2d1be-229">В будущем JavaScript будет работать на некоторых платформах напрямую, чтобы настраиваемые функции выполнялись быстрее и использовали меньше памяти.</span><span class="sxs-lookup"><span data-stu-id="2d1be-229">In the future, JavaScript will run directly on some platforms to ensure custom functions are faster and use less memory.</span></span> <span data-ttu-id="2d1be-230">Кроме того, HTML-страница, на которую ссылается элемент `<Page>` манифеста, не будет необходима для большинства платформ, так как Excel будет выполнять код JavaScript напрямую.</span><span class="sxs-lookup"><span data-stu-id="2d1be-230">Additionally, the HTML page referenced by the `<Page>`Page element in the manifest won’t be needed for most platforms because Excel will run the JavaScript directly.</span></span> <span data-ttu-id="2d1be-231">Чтобы подготовиться к этому изменению, убедитесь, что в ваших настраиваемых функциях не используется модель DOM для веб-страниц.</span><span class="sxs-lookup"><span data-stu-id="2d1be-231">To prepare for this change, ensure your custom functions do not use the webpage DOM.</span></span> <span data-ttu-id="2d1be-232">Поддерживаемые основные API приложения для доступа к Интернету будут [WebSocket](https://developer.mozilla.org/docs/Web/API/WebSockets_API) и [XHR](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest) с использованием GET или POST.</span><span class="sxs-lookup"><span data-stu-id="2d1be-232">The supported host APIs for accessing the web will be [WebSocket](https://developer.mozilla.org/docs/Web/API/WebSockets_API) and [XHR](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest) using GET or POST.</span></span>
- <span data-ttu-id="2d1be-233">Изменяемые функции (которые пересчитываются автоматически, когда в электронной таблице изменяются несвязанные данных) еще не поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="2d1be-233">Volatile functions (those which recalculate automatically whenever unrelated data changes in the spreadsheet) are not yet supported.</span></span>
- <span data-ttu-id="2d1be-234">Отладка включена только для асинхронных функций в Excel для Windows.</span><span class="sxs-lookup"><span data-stu-id="2d1be-234">Debugging is only enabled for asynchronous functions on Excel for Windows.</span></span>
- <span data-ttu-id="2d1be-235">Развертывание через Портал администрирования Office 365 и AppSource еще не включены.</span><span class="sxs-lookup"><span data-stu-id="2d1be-235">Deployment via the Office 365 Admin Portal and AppSource are not yet enabled.</span></span>
- <span data-ttu-id="2d1be-236">Настраиваемые функции в Excel Online могут перестать работать во время сеанса после периода бездействия.</span><span class="sxs-lookup"><span data-stu-id="2d1be-236">Custom functions in Excel Online may stop working during a session after a period of inactivity.</span></span> <span data-ttu-id="2d1be-237">Для восстановления работы обновите страницу браузера (F5) и повторно введите настраиваемую функцию.</span><span class="sxs-lookup"><span data-stu-id="2d1be-237">Refresh the browser page (F5) and re-enter a custom function to restore the feature.</span></span>

## <a name="changelog"></a><span data-ttu-id="2d1be-238">Журнал изменений</span><span class="sxs-lookup"><span data-stu-id="2d1be-238">Changelog</span></span>

- <span data-ttu-id="2d1be-239">**7 ноября 2017 г.** Доставлена\*  предварительная версия настраиваемых функций с примерами</span><span class="sxs-lookup"><span data-stu-id="2d1be-239">**Nov 7, 2017**: Shipped the custom functions preview and samples</span></span>
- <span data-ttu-id="2d1be-240">**20 ноября 2017 г.** Исправлена ошибка совместимости для пользователей, использующих сборки 8801 и выше.</span><span class="sxs-lookup"><span data-stu-id="2d1be-240">**Nov 20, 2017**: Fixed compatibility bug for those using builds 8801 and later</span></span>
- <span data-ttu-id="2d1be-241">**28 ноября 2017 г.** Доставлена\* поддержка отмены вызова асинхронных функций (необходимо изменение для функций потоковой передачи)</span><span class="sxs-lookup"><span data-stu-id="2d1be-241">**Nov 28, 2017**: Shipped support for cancellation on asynchronous functions (requires change for streaming functions)</span></span>
- <span data-ttu-id="2d1be-242">**7 мая 2018 г.** Доставлена\*​​поддержка Mac, Excel Online и синхронных функций, выполняемых внутри процесса</span><span class="sxs-lookup"><span data-stu-id="2d1be-242">**May 7, 2018**: Shipped support for Mac, Excel Online, and synchronous functions running in-process</span></span>

<span data-ttu-id="2d1be-243">\* канал участников программы предварительной оценки Office</span><span class="sxs-lookup"><span data-stu-id="2d1be-243">\* to the Office Insiders Channel</span></span>
