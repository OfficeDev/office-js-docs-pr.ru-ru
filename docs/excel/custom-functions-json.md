# <a name="custom-function-metadata"></a><span data-ttu-id="7237b-101">Метаданные настраиваемой функции</span><span class="sxs-lookup"><span data-stu-id="7237b-101">Custom function metadata</span></span>

<span data-ttu-id="7237b-102">Когда вы включаете [настраиваемые функции](custom-functions-overview.md) в надстройке Excel, вы должны разместить файл JSON, содержащий метаданные о функциях (в дополнение к размещению файла JavaScript с функциями и HTML-файлом без пользовательского интерфейса, который будет служить родителем файла JavaScript).</span><span class="sxs-lookup"><span data-stu-id="7237b-102">When you include [custom functions](custom-functions-overview.md) in an Excel add-in, you must host a JSON file that contains metadata about the functions (in addition to hosting a JavaScript file with the functions and a UI-less HTML file to serve as the parent of the JavaScript file).</span></span> <span data-ttu-id="7237b-103">В этой статье описывается формат файла JSON с примерами.</span><span class="sxs-lookup"><span data-stu-id="7237b-103">This article describes the format of the JSON file with examples.</span></span>

<span data-ttu-id="7237b-104">Полная выборка файла JSON доступна [здесь](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/customfunctions.json).</span><span class="sxs-lookup"><span data-stu-id="7237b-104">A complete sample JSON file is available [here](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/customfunctions.json).</span></span>

## <a name="functions-array"></a><span data-ttu-id="7237b-105">Массив функций</span><span class="sxs-lookup"><span data-stu-id="7237b-105">Functions array</span></span>

<span data-ttu-id="7237b-106">Метаданные — это объект JSON, содержащий одно свойство `functions`, значение которого представляет собой массив объектов.</span><span class="sxs-lookup"><span data-stu-id="7237b-106">The metadata is a JSON object that contains a single `functions` property whose value is an array of objects.</span></span> <span data-ttu-id="7237b-107">Каждый из этих объектов представляет собой одну настраиваемую функцию.</span><span class="sxs-lookup"><span data-stu-id="7237b-107">Each of these objects represents one custom function.</span></span> <span data-ttu-id="7237b-108">Следующая таблица содержит ее свойства:</span><span class="sxs-lookup"><span data-stu-id="7237b-108">The following table contains its properties:</span></span>

|  <span data-ttu-id="7237b-109">Свойство</span><span class="sxs-lookup"><span data-stu-id="7237b-109">Property</span></span>  |  <span data-ttu-id="7237b-110">Тип данных</span><span class="sxs-lookup"><span data-stu-id="7237b-110">Data Type</span></span>  |  <span data-ttu-id="7237b-111">Обязательность</span><span class="sxs-lookup"><span data-stu-id="7237b-111">Required?</span></span>  |  <span data-ttu-id="7237b-112">Описание</span><span class="sxs-lookup"><span data-stu-id="7237b-112">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="7237b-113">строка</span><span class="sxs-lookup"><span data-stu-id="7237b-113">string</span></span>  |  <span data-ttu-id="7237b-114">No (Нет)</span><span class="sxs-lookup"><span data-stu-id="7237b-114">No</span></span>  |  <span data-ttu-id="7237b-115">Описание функции, которая появляется в пользовательском интерфейсе Excel.</span><span class="sxs-lookup"><span data-stu-id="7237b-115">A description of the function that appears in the Excel UI.</span></span> <span data-ttu-id="7237b-116">Например, «Преобразует значение Цельсия в Фаренгейт».</span><span class="sxs-lookup"><span data-stu-id="7237b-116">For example, "Converts a Celsius value to Fahrenheit".</span></span> |
|  `helpUrl`  |  <span data-ttu-id="7237b-117">строка</span><span class="sxs-lookup"><span data-stu-id="7237b-117">string</span></span>  |   <span data-ttu-id="7237b-118">No (Нет)</span><span class="sxs-lookup"><span data-stu-id="7237b-118">No</span></span>  |  <span data-ttu-id="7237b-119">URL-адрес, где ваши пользователи могут получить помощь по функции.</span><span class="sxs-lookup"><span data-stu-id="7237b-119">URL where your users can get help about the function.</span></span> <span data-ttu-id="7237b-120">(Он отображается в панели задач.) Например, «http://contoso.com/help/convertcelsiustofahrenheit.html»</span><span class="sxs-lookup"><span data-stu-id="7237b-120">(It is displayed in a taskpane.) For example, "http://contoso.com/help/convertcelsiustofahrenheit.html"</span></span>  |
|  `name`  |  <span data-ttu-id="7237b-121">строка</span><span class="sxs-lookup"><span data-stu-id="7237b-121">string</span></span>  |  <span data-ttu-id="7237b-122">Да</span><span class="sxs-lookup"><span data-stu-id="7237b-122">Yes</span></span>  |  <span data-ttu-id="7237b-123">Имя функции, которая будет отображаться (добавлено в пространстве имен) в пользовательском интерфейсе Excel, когда пользователь выбирает функцию.</span><span class="sxs-lookup"><span data-stu-id="7237b-123">The name of the function as it will appear (prepended with a namespace) in the Excel UI when a user is selecting a function.</span></span> <span data-ttu-id="7237b-124">Оно должно совпадать с именем функции, указанном при ее определении в JavaScript.</span><span class="sxs-lookup"><span data-stu-id="7237b-124">It should be the same as the function's name where it is defined in the JavaScript.</span></span> |
|  `options`  |  <span data-ttu-id="7237b-125">объект</span><span class="sxs-lookup"><span data-stu-id="7237b-125">object</span></span>  |  <span data-ttu-id="7237b-126">No (Нет)</span><span class="sxs-lookup"><span data-stu-id="7237b-126">No</span></span>  |  <span data-ttu-id="7237b-127">Настройте, как Excel будет обрабатывать эту функцию.</span><span class="sxs-lookup"><span data-stu-id="7237b-127">Configure how Excel processes the function.</span></span> <span data-ttu-id="7237b-128">См. [объект опций](#options-object) для получения сведений.</span><span class="sxs-lookup"><span data-stu-id="7237b-128">See [options object](#options-object) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="7237b-129">массив</span><span class="sxs-lookup"><span data-stu-id="7237b-129">array</span></span>  |  <span data-ttu-id="7237b-130">Да</span><span class="sxs-lookup"><span data-stu-id="7237b-130">Yes</span></span>  |  <span data-ttu-id="7237b-131">Метаданные о параметрах функции.</span><span class="sxs-lookup"><span data-stu-id="7237b-131">Metadata about the parameters to the function.</span></span> <span data-ttu-id="7237b-132">См. [массив параметров](#parameters-array) для получения сведений.</span><span class="sxs-lookup"><span data-stu-id="7237b-132">See [parameters array](#parameters-array)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="7237b-133">объект</span><span class="sxs-lookup"><span data-stu-id="7237b-133">object</span></span>  |  <span data-ttu-id="7237b-134">Да</span><span class="sxs-lookup"><span data-stu-id="7237b-134">Yes</span></span>  |  <span data-ttu-id="7237b-135">Метаданные о значении, возвращаемом функцией.</span><span class="sxs-lookup"><span data-stu-id="7237b-135">Metadata about the value returned by the function.</span></span> <span data-ttu-id="7237b-136">См. [объект результата](#result-object) для получения сведений.</span><span class="sxs-lookup"><span data-stu-id="7237b-136">See [result object](#result-object) for details.</span></span> |

## <a name="options-object"></a><span data-ttu-id="7237b-137">Объект Options</span><span class="sxs-lookup"><span data-stu-id="7237b-137">Options object</span></span>

<span data-ttu-id="7237b-138">Объект `options` настраивает, как Excel обрабатывает эту функцию.</span><span class="sxs-lookup"><span data-stu-id="7237b-138">The `options` object configures how Excel processes the function.</span></span> <span data-ttu-id="7237b-139">Следующая таблица содержит ее свойства:</span><span class="sxs-lookup"><span data-stu-id="7237b-139">The following table contains its properties:</span></span>

|  <span data-ttu-id="7237b-140">Свойство</span><span class="sxs-lookup"><span data-stu-id="7237b-140">Property</span></span>  |  <span data-ttu-id="7237b-141">Тип данных</span><span class="sxs-lookup"><span data-stu-id="7237b-141">Data Type</span></span>  |  <span data-ttu-id="7237b-142">Обязательность</span><span class="sxs-lookup"><span data-stu-id="7237b-142">Required?</span></span>  |  <span data-ttu-id="7237b-143">Описание</span><span class="sxs-lookup"><span data-stu-id="7237b-143">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="7237b-144">логический</span><span class="sxs-lookup"><span data-stu-id="7237b-144">boolean</span></span>  |  <span data-ttu-id="7237b-145">Нет, значение по умолчанию — `false`.</span><span class="sxs-lookup"><span data-stu-id="7237b-145">No, default is `false`.</span></span>  |  <span data-ttu-id="7237b-146">Если `true`, Excel вызывает обработчика `onCanceled` всякий раз, когда пользователь предпринимает действие, которое имеет эффект отмены функции, например, вручную вызывая пересчет или редактирование ячейки, на которую ссылается функция.</span><span class="sxs-lookup"><span data-stu-id="7237b-146">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function.</span></span> <span data-ttu-id="7237b-147">Если вы используете эту опцию, Excel вызовет функцию JavaScript с дополнительным параметром `caller`.</span><span class="sxs-lookup"><span data-stu-id="7237b-147">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="7237b-148">(***Не***регистрируйте этот параметр в свойстве `parameters`.)</span><span class="sxs-lookup"><span data-stu-id="7237b-148">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="7237b-149">В тексте функции обработчик должен быть назначен члену `caller.onCanceled`.</span><span class="sxs-lookup"><span data-stu-id="7237b-149">In the body of the function, a handler must be assigned to the `caller.onCanceled` member.</span></span> <span data-ttu-id="7237b-150">Обратите внимание, что `cancelable` и `sync` не могут оба быть `true`.</span><span class="sxs-lookup"><span data-stu-id="7237b-150">Note, `cancelable` and `sync` cannot both be `true`.</span></span>  |
|  `stream`  |  <span data-ttu-id="7237b-151">логический</span><span class="sxs-lookup"><span data-stu-id="7237b-151">boolean</span></span>  |  <span data-ttu-id="7237b-152">Нет, значение по умолчанию — `false`.</span><span class="sxs-lookup"><span data-stu-id="7237b-152">No, default is `false`.</span></span>  |  <span data-ttu-id="7237b-153">Если `true`, функция может выводить несколько раз в ячейку даже при вызове только один раз.</span><span class="sxs-lookup"><span data-stu-id="7237b-153">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="7237b-154">Этот параметр полезен для быстро изменяющихся источников данных, таких как цена акций.</span><span class="sxs-lookup"><span data-stu-id="7237b-154">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="7237b-155">Если вы используете эту опцию, Excel вызовет функцию JavaScript с дополнительным параметром `caller`.</span><span class="sxs-lookup"><span data-stu-id="7237b-155">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="7237b-156">(***Не***регистрируйте этот параметр в свойстве `parameters`.)</span><span class="sxs-lookup"><span data-stu-id="7237b-156">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="7237b-157">Функция должна иметь выписку `return`.</span><span class="sxs-lookup"><span data-stu-id="7237b-157">The function should have no `return` statement.</span></span> <span data-ttu-id="7237b-158">Вместо этого значение результата передается как аргумент метода `caller.setResult` обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="7237b-158">Instead, the result value is passed as the argument of the `caller.setResult` callback method.</span></span> <span data-ttu-id="7237b-159">Обратите внимание, что `stream` и `sync` не могут быть оба `true`.</span><span class="sxs-lookup"><span data-stu-id="7237b-159">Note, `stream` and `sync` may not both be `true`.</span></span>|
|  `sync`  |  <span data-ttu-id="7237b-160">логический</span><span class="sxs-lookup"><span data-stu-id="7237b-160">boolean</span></span>  |  <span data-ttu-id="7237b-161">Нет, значение по умолчанию `false`</span><span class="sxs-lookup"><span data-stu-id="7237b-161">No, default is `false`</span></span>  |  <span data-ttu-id="7237b-162">Если `true`, функция запускается синхронно и должна возвращать значение.</span><span class="sxs-lookup"><span data-stu-id="7237b-162">If `true`, the function runs synchronously and it must return a value.</span></span> <span data-ttu-id="7237b-163">Если `false`, функция выполняется асинхронно, и она должна возвращать объект`OfficeExtension.Promise`.</span><span class="sxs-lookup"><span data-stu-id="7237b-163">If `false`, the function runs asynchronously and it must return a `OfficeExtension.Promise` object.</span></span> <span data-ttu-id="7237b-164">Примечание. `sync` может не являться`true`, если `cancelable` или `stream` являются `true`.</span><span class="sxs-lookup"><span data-stu-id="7237b-164">Note, `sync`  may not be `true` if either `cancelable` or `stream` are `true`.</span></span>  |

## <a name="parameters-array"></a><span data-ttu-id="7237b-165">Массив параметров</span><span class="sxs-lookup"><span data-stu-id="7237b-165">Parameters array</span></span>

<span data-ttu-id="7237b-166">Свойство `parameters` находится в массиве параметров.</span><span class="sxs-lookup"><span data-stu-id="7237b-166">The `parameters` property is an array of objects.</span></span> <span data-ttu-id="7237b-167">Каждый из этих объектов представляет собой параметр.</span><span class="sxs-lookup"><span data-stu-id="7237b-167">Each of these objects represents a parameter.</span></span> <span data-ttu-id="7237b-168">Следующая таблица содержит ее свойства:</span><span class="sxs-lookup"><span data-stu-id="7237b-168">The following table contains its properties:</span></span>

|  <span data-ttu-id="7237b-169">Свойство</span><span class="sxs-lookup"><span data-stu-id="7237b-169">Property</span></span>  |  <span data-ttu-id="7237b-170">Тип данных</span><span class="sxs-lookup"><span data-stu-id="7237b-170">Data Type</span></span>  |  <span data-ttu-id="7237b-171">Обязательность</span><span class="sxs-lookup"><span data-stu-id="7237b-171">Required?</span></span>  |  <span data-ttu-id="7237b-172">Описание</span><span class="sxs-lookup"><span data-stu-id="7237b-172">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="7237b-173">строка</span><span class="sxs-lookup"><span data-stu-id="7237b-173">string</span></span>  |  <span data-ttu-id="7237b-174">No (Нет)</span><span class="sxs-lookup"><span data-stu-id="7237b-174">No</span></span> |  <span data-ttu-id="7237b-175">Описание параметра.</span><span class="sxs-lookup"><span data-stu-id="7237b-175">A description of the parameter.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="7237b-176">строка</span><span class="sxs-lookup"><span data-stu-id="7237b-176">string</span></span>  |  <span data-ttu-id="7237b-177">Да</span><span class="sxs-lookup"><span data-stu-id="7237b-177">Yes</span></span>  |  <span data-ttu-id="7237b-178">Должно быть либо «скалярным», то есть значением без массива, либо «матрицей», то есть массивом массивов строк.</span><span class="sxs-lookup"><span data-stu-id="7237b-178">Must be either "scalar", meaning a non-array value, or "matrix", meaning an array of row arrays.</span></span>  |
|  `name`  |  <span data-ttu-id="7237b-179">строка</span><span class="sxs-lookup"><span data-stu-id="7237b-179">string</span></span>  |  <span data-ttu-id="7237b-180">Да</span><span class="sxs-lookup"><span data-stu-id="7237b-180">Yes</span></span>  |  <span data-ttu-id="7237b-181">Имя параметра.</span><span class="sxs-lookup"><span data-stu-id="7237b-181">The name of the parameter.</span></span> <span data-ttu-id="7237b-182">Это имя отображается в Excel IntelliSense.</span><span class="sxs-lookup"><span data-stu-id="7237b-182">This name is displayed in Excel's IntelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="7237b-183">строка</span><span class="sxs-lookup"><span data-stu-id="7237b-183">string</span></span>  |  <span data-ttu-id="7237b-184">Да</span><span class="sxs-lookup"><span data-stu-id="7237b-184">Yes</span></span>  |  <span data-ttu-id="7237b-185">Тип данных параметра.</span><span class="sxs-lookup"><span data-stu-id="7237b-185">The data type of the parameter.</span></span> <span data-ttu-id="7237b-186">Должно быть «логический», «числовой» или «строка».</span><span class="sxs-lookup"><span data-stu-id="7237b-186">Must be "boolean", "number", or "string".</span></span>  |

## <a name="result-object"></a><span data-ttu-id="7237b-187">Результирующий объект</span><span class="sxs-lookup"><span data-stu-id="7237b-187">Result object</span></span>

<span data-ttu-id="7237b-188">Свойство `results` предоставляет метаданные о значении, возвращаемом функцией.</span><span class="sxs-lookup"><span data-stu-id="7237b-188">The `results` property provides metadata about the value returned from the function.</span></span> <span data-ttu-id="7237b-189">Следующая таблица содержит ее свойства:</span><span class="sxs-lookup"><span data-stu-id="7237b-189">The following table contains its properties:</span></span>

|  <span data-ttu-id="7237b-190">Свойство</span><span class="sxs-lookup"><span data-stu-id="7237b-190">Property</span></span>  |  <span data-ttu-id="7237b-191">Тип данных</span><span class="sxs-lookup"><span data-stu-id="7237b-191">Data Type</span></span>  |  <span data-ttu-id="7237b-192">Обязательность</span><span class="sxs-lookup"><span data-stu-id="7237b-192">Required?</span></span>  |  <span data-ttu-id="7237b-193">Описание</span><span class="sxs-lookup"><span data-stu-id="7237b-193">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="7237b-194">строка</span><span class="sxs-lookup"><span data-stu-id="7237b-194">string</span></span>  |  <span data-ttu-id="7237b-195">No (Нет)</span><span class="sxs-lookup"><span data-stu-id="7237b-195">No</span></span>  |  <span data-ttu-id="7237b-196">Должно быть либо «скалярным», то есть значением без массива, либо «матрицей», то есть массивом массивов строк.</span><span class="sxs-lookup"><span data-stu-id="7237b-196">Must be either "scalar", meaning a non-array value, or "matrix", meaning an array of row arrays.</span></span>  |
|  `type`  |  <span data-ttu-id="7237b-197">строка</span><span class="sxs-lookup"><span data-stu-id="7237b-197">string</span></span>  |  <span data-ttu-id="7237b-198">Да</span><span class="sxs-lookup"><span data-stu-id="7237b-198">Yes</span></span>  |  <span data-ttu-id="7237b-199">Тип данных параметра.</span><span class="sxs-lookup"><span data-stu-id="7237b-199">The data type of the parameter.</span></span> <span data-ttu-id="7237b-200">Должно быть «логический», «числовой» или «строка».</span><span class="sxs-lookup"><span data-stu-id="7237b-200">Must be "boolean", "number", or "string".</span></span>  |

## <a name="example"></a><span data-ttu-id="7237b-201">Пример</span><span class="sxs-lookup"><span data-stu-id="7237b-201">Example</span></span>

<span data-ttu-id="7237b-202">Следующий код JSON является примером файла метаданных для пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="7237b-202">The following JSON code is an example of a metadata file for custom functions.</span></span>

```json
{
    "functions": [
        {
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
            ],
            "options": {
                "sync": true
            }
        },
        {
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
            ],
            "options": {
                "sync": false
            }
        },
        {
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
            ],
            "options": {
                "sync": true
            }
        },
        {
            "name": "GETDAY",
            "description": "Gets the day of the week",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "string"
            },
            "parameters": [],
            "options": {
                "sync": true
            }
        },
        {
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
                "sync": false,
                "stream": true,
                "cancelable": true
            }
        },
        {
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
            ],
            "options": {
                "sync": true
            }
        }
    ]
}

```

## <a name="see-also"></a><span data-ttu-id="7237b-203">См. также</span><span class="sxs-lookup"><span data-stu-id="7237b-203">See also</span></span>
[<span data-ttu-id="7237b-204">Настраиваемые функции</span><span class="sxs-lookup"><span data-stu-id="7237b-204">Custom functions</span></span>](custom-functions-overview.md)<br>
[<span data-ttu-id="7237b-205">Руководства и примеры формул массива</span><span class="sxs-lookup"><span data-stu-id="7237b-205">Guidelines and examples of array formulas</span></span>](https://support.office.com/article/Guidelines-and-examples-of-array-formulas-7d94a64e-3ff3-4686-9372-ecfd5caa57c7)
