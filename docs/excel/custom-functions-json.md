# <a name="custom-function-metadata"></a><span data-ttu-id="e8ef5-101">Метаданные настраиваемой функции</span><span class="sxs-lookup"><span data-stu-id="e8ef5-101">Custom function metadata</span></span>

<span data-ttu-id="e8ef5-102">Когда вы включаете [настраиваемые функции](custom-functions-overview.md) в надстройке Excel, вы должны разместить файл JSON, содержащий метаданные о функциях (в дополнение к размещению файла JavaScript с функциями и HTML-файлом без пользовательского интерфейса, который будет служить родителем файла JavaScript).</span><span class="sxs-lookup"><span data-stu-id="e8ef5-102">When you include [custom functions](custom-functions-overview.md) in an Excel add-in, you must host a JSON file that contains metadata about the functions (in addition to hosting a JavaScript file with the functions and a UI-less HTML file to serve as the parent of the JavaScript file).</span></span> <span data-ttu-id="e8ef5-103">В этой статье описывается формат файла JSON с примерами.</span><span class="sxs-lookup"><span data-stu-id="e8ef5-103">This article describes the format of the JSON file with examples.</span></span>

<span data-ttu-id="e8ef5-104">Полная выборка файла JSON доступна [здесь](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).</span><span class="sxs-lookup"><span data-stu-id="e8ef5-104">A complete sample JSON file is available [here](https://github.com/OfficeDev/Excel-Custom-Functions/blob/master/config/customfunctions.json).</span></span>

## <a name="functions-array"></a><span data-ttu-id="e8ef5-105">Массив функций</span><span class="sxs-lookup"><span data-stu-id="e8ef5-105">Functions array</span></span>

<span data-ttu-id="e8ef5-106">Метаданные — это объект JSON, содержащий одно свойство `functions`, значение которого представляет собой массив объектов.</span><span class="sxs-lookup"><span data-stu-id="e8ef5-106">The metadata is a JSON object that contains a single `functions` property whose value is an array of objects.</span></span> <span data-ttu-id="e8ef5-107">Каждый из этих объектов представляет собой одну настраиваемую функцию.</span><span class="sxs-lookup"><span data-stu-id="e8ef5-107">Each of these objects represents one custom function.</span></span> <span data-ttu-id="e8ef5-108">Следующая таблица содержит ее свойства:</span><span class="sxs-lookup"><span data-stu-id="e8ef5-108">The following table contains its properties:</span></span>

|  <span data-ttu-id="e8ef5-109">Свойство</span><span class="sxs-lookup"><span data-stu-id="e8ef5-109">Property</span></span>  |  <span data-ttu-id="e8ef5-110">Тип данных</span><span class="sxs-lookup"><span data-stu-id="e8ef5-110">Data Type</span></span>  |  <span data-ttu-id="e8ef5-111">Обязательность</span><span class="sxs-lookup"><span data-stu-id="e8ef5-111">Required?</span></span>  |  <span data-ttu-id="e8ef5-112">Описание</span><span class="sxs-lookup"><span data-stu-id="e8ef5-112">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="e8ef5-113">строка</span><span class="sxs-lookup"><span data-stu-id="e8ef5-113">string</span></span>  |  <span data-ttu-id="e8ef5-114">Нет</span><span class="sxs-lookup"><span data-stu-id="e8ef5-114">No</span></span>  |  <span data-ttu-id="e8ef5-115">Описание функции, которая появляется в пользовательском интерфейсе Excel.</span><span class="sxs-lookup"><span data-stu-id="e8ef5-115">A description of the function that appears in the Excel UI.</span></span> <span data-ttu-id="e8ef5-116">Например, «Преобразует значение Цельсия в Фаренгейт».</span><span class="sxs-lookup"><span data-stu-id="e8ef5-116">For example, "Converts a Celsius value to Fahrenheit".</span></span> |
|  `helpUrl`  |  <span data-ttu-id="e8ef5-117">строка</span><span class="sxs-lookup"><span data-stu-id="e8ef5-117">string</span></span>  |   <span data-ttu-id="e8ef5-118">Нет</span><span class="sxs-lookup"><span data-stu-id="e8ef5-118">No</span></span>  |  <span data-ttu-id="e8ef5-119">URL-адрес, где ваши пользователи могут получить помощь по функции.</span><span class="sxs-lookup"><span data-stu-id="e8ef5-119">URL where your users can get help about the function.</span></span> <span data-ttu-id="e8ef5-120">(Он отображается в панели задач.) Например, «http://contoso.com/help/convertcelsiustofahrenheit.html»</span><span class="sxs-lookup"><span data-stu-id="e8ef5-120">(It is displayed in a taskpane.) For example, "http://contoso.com/help/convertcelsiustofahrenheit.html"</span></span>  |
|  `name`  |  <span data-ttu-id="e8ef5-121">строка</span><span class="sxs-lookup"><span data-stu-id="e8ef5-121">string</span></span>  |  <span data-ttu-id="e8ef5-122">Да</span><span class="sxs-lookup"><span data-stu-id="e8ef5-122">Yes</span></span>  |  <span data-ttu-id="e8ef5-123">Имя функции, которая будет отображаться (добавлено в пространстве имен) в пользовательском интерфейсе Excel, когда пользователь выбирает функцию.</span><span class="sxs-lookup"><span data-stu-id="e8ef5-123">The name of the function as it will appear (prepended with a namespace) in the Excel UI when a user is selecting a function.</span></span> <span data-ttu-id="e8ef5-124">Оно должно совпадать с именем функции, указанном при ее определении в JavaScript.</span><span class="sxs-lookup"><span data-stu-id="e8ef5-124">It should be the same as the function's name where it is defined in the JavaScript.</span></span> |
|  `options`  |  <span data-ttu-id="e8ef5-125">объект</span><span class="sxs-lookup"><span data-stu-id="e8ef5-125">object</span></span>  |  <span data-ttu-id="e8ef5-126">Нет</span><span class="sxs-lookup"><span data-stu-id="e8ef5-126">No</span></span>  |  <span data-ttu-id="e8ef5-127">Настройте, как Excel будет обрабатывать эту функцию.</span><span class="sxs-lookup"><span data-stu-id="e8ef5-127">Configure how Excel processes the function.</span></span> <span data-ttu-id="e8ef5-128">См. [объект опций](#options-object) для получения сведений.</span><span class="sxs-lookup"><span data-stu-id="e8ef5-128">See [options object](#options-object) for details.</span></span> |
|  `parameters`  |  <span data-ttu-id="e8ef5-129">array</span><span class="sxs-lookup"><span data-stu-id="e8ef5-129">array</span></span>  |  <span data-ttu-id="e8ef5-130">Да</span><span class="sxs-lookup"><span data-stu-id="e8ef5-130">Yes</span></span>  |  <span data-ttu-id="e8ef5-131">Метаданные о параметрах функции.</span><span class="sxs-lookup"><span data-stu-id="e8ef5-131">Metadata about the parameters to the function.</span></span> <span data-ttu-id="e8ef5-132">См. [массив параметров](#parameters-array) для получения сведений.</span><span class="sxs-lookup"><span data-stu-id="e8ef5-132">See [parameters array](#parameters-array)  for details.</span></span> |
|  `result`  |  <span data-ttu-id="e8ef5-133">объект</span><span class="sxs-lookup"><span data-stu-id="e8ef5-133">object</span></span>  |  <span data-ttu-id="e8ef5-134">Да</span><span class="sxs-lookup"><span data-stu-id="e8ef5-134">Yes</span></span>  |  <span data-ttu-id="e8ef5-135">Метаданные о значении, возвращаемом функцией.</span><span class="sxs-lookup"><span data-stu-id="e8ef5-135">Metadata about the value returned by the function.</span></span> <span data-ttu-id="e8ef5-136">См. [объект результата](#result-object) для получения сведений.</span><span class="sxs-lookup"><span data-stu-id="e8ef5-136">See [result object](#result-object) for details.</span></span> |

## <a name="options-object"></a><span data-ttu-id="e8ef5-137">Объект Options</span><span class="sxs-lookup"><span data-stu-id="e8ef5-137">Options object</span></span>

<span data-ttu-id="e8ef5-138">Объект `options` настраивает, как Excel обрабатывает эту функцию.</span><span class="sxs-lookup"><span data-stu-id="e8ef5-138">The `options` object configures how Excel processes the function.</span></span> <span data-ttu-id="e8ef5-139">Следующая таблица содержит ее свойства:</span><span class="sxs-lookup"><span data-stu-id="e8ef5-139">The following table contains its properties:</span></span>

|  <span data-ttu-id="e8ef5-140">Свойство</span><span class="sxs-lookup"><span data-stu-id="e8ef5-140">Property</span></span>  |  <span data-ttu-id="e8ef5-141">Тип данных</span><span class="sxs-lookup"><span data-stu-id="e8ef5-141">Data Type</span></span>  |  <span data-ttu-id="e8ef5-142">Обязательность</span><span class="sxs-lookup"><span data-stu-id="e8ef5-142">Required?</span></span>  |  <span data-ttu-id="e8ef5-143">Описание</span><span class="sxs-lookup"><span data-stu-id="e8ef5-143">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `cancelable`  |  <span data-ttu-id="e8ef5-144">boolean</span><span class="sxs-lookup"><span data-stu-id="e8ef5-144">boolean</span></span>  |  <span data-ttu-id="e8ef5-145">Нет, значение по умолчанию — `false`.</span><span class="sxs-lookup"><span data-stu-id="e8ef5-145">No, default is `false`.</span></span>  |  <span data-ttu-id="e8ef5-p110">Если `true` Excel вызывает `onCanceled`обработчик каждый раз, когда пользователь предпринимает действие, которое имеет тот же эффект, что и отмена функции; например, вручную запуск пересчета или редактирования ячейки, на который ссылается функция. Если используется этот параметр, Excel вызовет функцию JavaScript с дополнительным `caller` параметром. (Не ***регистрируйте***этот параметр в `parameters`свойстве). В теле функции, должен быть назначен обработчик `caller.onCanceled`члена.</span><span class="sxs-lookup"><span data-stu-id="e8ef5-p110">If `true`, Excel calls the `onCanceled` handler whenever the user takes an action that has the effect of canceling the function; for example, manually triggering recalculation or editing a cell that is referenced by the function. If you use this option, Excel will call the JavaScript function with an additional `caller` parameter. (Do ***not*** register this parameter in the `parameters` property). In the body of the function, a handler must be assigned to the `caller.onCanceled` member. Note,  and  cannot both be .</span></span>|
|  `stream`  |  <span data-ttu-id="e8ef5-150">boolean</span><span class="sxs-lookup"><span data-stu-id="e8ef5-150">boolean</span></span>  |  <span data-ttu-id="e8ef5-151">Нет, значение по умолчанию — `false`.</span><span class="sxs-lookup"><span data-stu-id="e8ef5-151">No, default is `false`.</span></span>  |  <span data-ttu-id="e8ef5-152">Если `true`, функция может выводить несколько раз в ячейку даже при вызове только один раз.</span><span class="sxs-lookup"><span data-stu-id="e8ef5-152">If `true`, the function can output repeatedly to the cell even when invoked only once.</span></span> <span data-ttu-id="e8ef5-153">Этот параметр полезен для быстро изменяющихся источников данных, таких как цена акций.</span><span class="sxs-lookup"><span data-stu-id="e8ef5-153">This option is useful for rapidly-changing data sources, such as a stock price.</span></span> <span data-ttu-id="e8ef5-154">Если вы используете эту опцию, Excel вызовет функцию JavaScript с дополнительным параметром `caller`.</span><span class="sxs-lookup"><span data-stu-id="e8ef5-154">If you use this option, Excel will call the JavaScript function with an additional `caller` parameter.</span></span> <span data-ttu-id="e8ef5-155">(***Не***регистрируйте этот параметр в свойстве `parameters`.)</span><span class="sxs-lookup"><span data-stu-id="e8ef5-155">(Do ***not*** register this parameter in the `parameters` property).</span></span> <span data-ttu-id="e8ef5-156">Функция должна иметь выписку `return`.</span><span class="sxs-lookup"><span data-stu-id="e8ef5-156">The function should have no `return` statement.</span></span> <span data-ttu-id="e8ef5-157">Вместо этого значение результата передается как аргумент метода `caller.setResult` обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="e8ef5-157">Instead, the result value is passed as the argument of the `caller.setResult` callback method.</span></span>|

## <a name="parameters-array"></a><span data-ttu-id="e8ef5-158">Массив параметров</span><span class="sxs-lookup"><span data-stu-id="e8ef5-158">Parameters array</span></span>

<span data-ttu-id="e8ef5-159">Свойство `parameters` находится в массиве параметров.</span><span class="sxs-lookup"><span data-stu-id="e8ef5-159">The `parameters` property is an array of objects.</span></span> <span data-ttu-id="e8ef5-160">Каждый из этих объектов представляет собой параметр.</span><span class="sxs-lookup"><span data-stu-id="e8ef5-160">Each of these objects represents a parameter.</span></span> <span data-ttu-id="e8ef5-161">Следующая таблица содержит ее свойства:</span><span class="sxs-lookup"><span data-stu-id="e8ef5-161">The following table contains its properties:</span></span>

|  <span data-ttu-id="e8ef5-162">Свойство</span><span class="sxs-lookup"><span data-stu-id="e8ef5-162">Property</span></span>  |  <span data-ttu-id="e8ef5-163">Тип данных</span><span class="sxs-lookup"><span data-stu-id="e8ef5-163">Data Type</span></span>  |  <span data-ttu-id="e8ef5-164">Обязательность</span><span class="sxs-lookup"><span data-stu-id="e8ef5-164">Required?</span></span>  |  <span data-ttu-id="e8ef5-165">Описание</span><span class="sxs-lookup"><span data-stu-id="e8ef5-165">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `description`  |  <span data-ttu-id="e8ef5-166">строка</span><span class="sxs-lookup"><span data-stu-id="e8ef5-166">string</span></span>  |  <span data-ttu-id="e8ef5-167">Нет</span><span class="sxs-lookup"><span data-stu-id="e8ef5-167">No</span></span> |  <span data-ttu-id="e8ef5-168">Описание параметра.</span><span class="sxs-lookup"><span data-stu-id="e8ef5-168">A description of the parameter.</span></span>  |
|  `dimensionality`  |  <span data-ttu-id="e8ef5-169">строка</span><span class="sxs-lookup"><span data-stu-id="e8ef5-169">string</span></span>  |  <span data-ttu-id="e8ef5-170">Да</span><span class="sxs-lookup"><span data-stu-id="e8ef5-170">Yes</span></span>  |  <span data-ttu-id="e8ef5-171">Должно быть либо «скалярным», то есть значением без массива, либо «матрицей», то есть массивом массивов строк.</span><span class="sxs-lookup"><span data-stu-id="e8ef5-171">Must be either "scalar", meaning a non-array value, or "matrix", meaning an array of row arrays.</span></span>  |
|  `name`  |  <span data-ttu-id="e8ef5-172">строка</span><span class="sxs-lookup"><span data-stu-id="e8ef5-172">string</span></span>  |  <span data-ttu-id="e8ef5-173">Да</span><span class="sxs-lookup"><span data-stu-id="e8ef5-173">Yes</span></span>  |  <span data-ttu-id="e8ef5-174">Имя параметра.</span><span class="sxs-lookup"><span data-stu-id="e8ef5-174">The name of the parameter.</span></span> <span data-ttu-id="e8ef5-175">Это имя отображается в Excel IntelliSense.</span><span class="sxs-lookup"><span data-stu-id="e8ef5-175">This name is displayed in Excel's IntelliSense.</span></span>  |
|  `type`  |  <span data-ttu-id="e8ef5-176">строка</span><span class="sxs-lookup"><span data-stu-id="e8ef5-176">string</span></span>  |  <span data-ttu-id="e8ef5-177">Да</span><span class="sxs-lookup"><span data-stu-id="e8ef5-177">Yes</span></span>  |  <span data-ttu-id="e8ef5-178">Тип данных параметра.</span><span class="sxs-lookup"><span data-stu-id="e8ef5-178">The data type of the parameter.</span></span> <span data-ttu-id="e8ef5-179">Должно быть «логический», «числовой» или «строка».</span><span class="sxs-lookup"><span data-stu-id="e8ef5-179">Must be "boolean", "number", or "string".</span></span>  |

## <a name="result-object"></a><span data-ttu-id="e8ef5-180">Результирующий объект</span><span class="sxs-lookup"><span data-stu-id="e8ef5-180">Result object</span></span>

<span data-ttu-id="e8ef5-181">Свойство `results` предоставляет метаданные о значении, возвращаемом функцией.</span><span class="sxs-lookup"><span data-stu-id="e8ef5-181">The `results` property provides metadata about the value returned from the function.</span></span> <span data-ttu-id="e8ef5-182">Следующая таблица содержит ее свойства:</span><span class="sxs-lookup"><span data-stu-id="e8ef5-182">The following table contains its properties:</span></span>

|  <span data-ttu-id="e8ef5-183">Свойство</span><span class="sxs-lookup"><span data-stu-id="e8ef5-183">Property</span></span>  |  <span data-ttu-id="e8ef5-184">Тип данных</span><span class="sxs-lookup"><span data-stu-id="e8ef5-184">Data Type</span></span>  |  <span data-ttu-id="e8ef5-185">Обязательность</span><span class="sxs-lookup"><span data-stu-id="e8ef5-185">Required?</span></span>  |  <span data-ttu-id="e8ef5-186">Описание</span><span class="sxs-lookup"><span data-stu-id="e8ef5-186">Description</span></span>  |
|:-----|:-----|:-----|:-----|
|  `dimensionality`  |  <span data-ttu-id="e8ef5-187">строка</span><span class="sxs-lookup"><span data-stu-id="e8ef5-187">string</span></span>  |  <span data-ttu-id="e8ef5-188">Нет</span><span class="sxs-lookup"><span data-stu-id="e8ef5-188">No</span></span>  |  <span data-ttu-id="e8ef5-189">Должно быть либо «скалярным», то есть значением без массива, либо «матрицей», то есть массивом массивов строк.</span><span class="sxs-lookup"><span data-stu-id="e8ef5-189">Must be either "scalar", meaning a non-array value, or "matrix", meaning an array of row arrays.</span></span>  |
|  `type`  |  <span data-ttu-id="e8ef5-190">строка</span><span class="sxs-lookup"><span data-stu-id="e8ef5-190">string</span></span>  |  <span data-ttu-id="e8ef5-191">Да</span><span class="sxs-lookup"><span data-stu-id="e8ef5-191">Yes</span></span>  |  <span data-ttu-id="e8ef5-192">Тип данных параметра.</span><span class="sxs-lookup"><span data-stu-id="e8ef5-192">The data type of the parameter.</span></span> <span data-ttu-id="e8ef5-193">Должно быть «логический», «числовой» или «строка».</span><span class="sxs-lookup"><span data-stu-id="e8ef5-193">Must be "boolean", "number", or "string".</span></span>  |

## <a name="example"></a><span data-ttu-id="e8ef5-194">Пример</span><span class="sxs-lookup"><span data-stu-id="e8ef5-194">Example</span></span>

<span data-ttu-id="e8ef5-195">Следующий код JSON является примером файла метаданных для пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="e8ef5-195">The following JSON code is an example of a metadata file for custom functions.</span></span>

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
            ]
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
            ]
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
            ]
        },
        {
            "name": "GETDAY",
            "description": "Gets the day of the week",
            "helpUrl": "http://dev.office.com",
            "result": {
                "type": "string"
            },
            "parameters": []
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
            ]
        }
    ]
}

```

## <a name="see-also"></a><span data-ttu-id="e8ef5-196">См. также</span><span class="sxs-lookup"><span data-stu-id="e8ef5-196">See also</span></span>
[<span data-ttu-id="e8ef5-197">Настраиваемые функции</span><span class="sxs-lookup"><span data-stu-id="e8ef5-197">Custom functions</span></span>](custom-functions-overview.md)<br>
[<span data-ttu-id="e8ef5-198">Руководства и примеры формул массива</span><span class="sxs-lookup"><span data-stu-id="e8ef5-198">Guidelines and examples of array formulas</span></span>](https://support.office.com/article/Guidelines-and-examples-of-array-formulas-7d94a64e-3ff3-4686-9372-ecfd5caa57c7)
