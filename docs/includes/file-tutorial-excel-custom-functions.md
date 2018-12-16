# <a name="tutorial-create-custom-functions-in-excel"></a><span data-ttu-id="9d9de-101">Руководство: создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="9d9de-101">Tutorial: Create custom functions in Excel</span></span>

## <a name="introduction"></a><span data-ttu-id="9d9de-102">Введение</span><span class="sxs-lookup"><span data-stu-id="9d9de-102">Introduction</span></span>

<span data-ttu-id="9d9de-103">Пользовательские функции позволяют добавлять новые функции в Excel путем определения этих функций в JavaScript как части надстройки.</span><span class="sxs-lookup"><span data-stu-id="9d9de-103">Custom functions enable you to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="9d9de-104">Пользователи в Excel могут получить доступ к пользовательским функциям так же, как и к любой встроенной функции в Excel, например `SUM()`.</span><span class="sxs-lookup"><span data-stu-id="9d9de-104">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="9d9de-105">Вы можете создавать пользовательские функции, которые будут выполнять простые задачи, такие как настраиваемые вычисления, или более сложные задачи, такие как потоковая передача данных в режиме реального времени из Интернета на лист.</span><span class="sxs-lookup"><span data-stu-id="9d9de-105">You can create custom functions that perform simple tasks such as custom calculations or more complex tasks such as streaming real-time data from the web into a worksheet.</span></span>

<span data-ttu-id="9d9de-106">В этом руководстве описан порядок выполнения перечисленных ниже задач.</span><span class="sxs-lookup"><span data-stu-id="9d9de-106">In this tutorial, you will:</span></span>
> [!div class="checklist"]
> * <span data-ttu-id="9d9de-107">Создание проекта пользовательских функций с помощью генератора Yo Office</span><span class="sxs-lookup"><span data-stu-id="9d9de-107">Create a custom functions project by using the Yo Office generator</span></span>
> * <span data-ttu-id="9d9de-108">Использование готовой пользовательской функции для выполнения простых вычислений</span><span class="sxs-lookup"><span data-stu-id="9d9de-108">Use a prebuilt custom function to perform a simple calculation</span></span>
> * <span data-ttu-id="9d9de-109">Создание пользовательской функции, которая запрашивает данные из Интернета</span><span class="sxs-lookup"><span data-stu-id="9d9de-109">Create a custom function that requests data from the web</span></span>
> * <span data-ttu-id="9d9de-110">Создание пользовательской функции, которая осуществляет потоковую передачу данных в реальном времени из Интернета</span><span class="sxs-lookup"><span data-stu-id="9d9de-110">Create a custom function that streams real-time data from the web</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="prerequisites"></a><span data-ttu-id="9d9de-111">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="9d9de-111">Prerequisites</span></span>

* [<span data-ttu-id="9d9de-112">Node.js и npm</span><span class="sxs-lookup"><span data-stu-id="9d9de-112">Node.js and npm</span></span>](https://nodejs.org/en/)

* <span data-ttu-id="9d9de-113">[Git Bash](https://git-scm.com/downloads) (или другой клиент Git)</span><span class="sxs-lookup"><span data-stu-id="9d9de-113">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

* <span data-ttu-id="9d9de-114">Последняя версия [Yeoman](https://yeoman.io/) и [генератора Yo Office](https://www.npmjs.com/package/generator-office).</span><span class="sxs-lookup"><span data-stu-id="9d9de-114">The latest version of [Yeoman](https://yeoman.io/) and the [Yo Office generator](https://www.npmjs.com/package/generator-office).</span></span> <span data-ttu-id="9d9de-115">Выполните в командной строке указанную ниже команду, чтобы установить эти инструменты глобально.</span><span class="sxs-lookup"><span data-stu-id="9d9de-115">To install these tools globally, run the following command via the command prompt:</span></span>

    ```bash
    npm install -g yo generator-office
    ```

* <span data-ttu-id="9d9de-116">Excel для Windows (версии 1810 или более поздней) или Excel Online</span><span class="sxs-lookup"><span data-stu-id="9d9de-116">Excel for Windows (version 1810 or later) or Excel Online</span></span>

* <span data-ttu-id="9d9de-117">Присоединитесь к [Программе предварительной оценки Office](https://products.office.com/office-insider) (уровень **Участник**; ранее "Предварительная оценка — ранний доступ")</span><span class="sxs-lookup"><span data-stu-id="9d9de-117">Join the [Office Insider program](https://products.office.com/office-insider) (**Insider** level -- formerly called "Insider Fast")</span></span>

## <a name="create-a-custom-functions-project"></a><span data-ttu-id="9d9de-118">Создание проекта пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="9d9de-118">Create a custom functions project</span></span>

<span data-ttu-id="9d9de-119">Вы начнете работу с этим руководством с использования генератора Yo Office для создания файлов, которые необходимы для проекта пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="9d9de-119">You’ll begin this tutorial by using the Yo Office generator to create the files that you need for your custom functions project.</span></span>

1. <span data-ttu-id="9d9de-120">Выполните указанную ниже команду и ответьте на вопросы, как показано ниже.</span><span class="sxs-lookup"><span data-stu-id="9d9de-120">Run the following command and then answer the prompts as follows.</span></span>

    ```bash
    yo office
    ```

    * <span data-ttu-id="9d9de-121">Выберите тип проекта: `Excel Custom Functions Add-in project (...)`</span><span class="sxs-lookup"><span data-stu-id="9d9de-121">Choose a project type: `Excel Custom Functions Add-in project (...)`</span></span>
    * <span data-ttu-id="9d9de-122">Выберите тип сценария: `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="9d9de-122">Choose a script type: `JavaScript`</span></span>
    * <span data-ttu-id="9d9de-123">Как вы хотите назвать свою надстройку?</span><span class="sxs-lookup"><span data-stu-id="9d9de-123">What do you want to name your add-in?</span></span> `stock-ticker`

    ![Подсказки Bash Yo Office для пользовательских функций](../images/yo-office-cfs-stock-ticker-3.png)

    <span data-ttu-id="9d9de-125">После завершения работы мастера генератор создает файлы проекта и устанавливает вспомогательные компоненты Node.</span><span class="sxs-lookup"><span data-stu-id="9d9de-125">After you complete the wizard, the generator will create the project files and install supporting Node components.</span></span> <span data-ttu-id="9d9de-126">Файлы проекта взяты из репозитория [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) GitHub.</span><span class="sxs-lookup"><span data-stu-id="9d9de-126">The project files come from the [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) GitHub repository.</span></span>

2. <span data-ttu-id="9d9de-127">Перейдите в папку проекта.</span><span class="sxs-lookup"><span data-stu-id="9d9de-127">Navigate to the project folder.</span></span>

    ```bash
    cd stock-ticker
    ```

3. <span data-ttu-id="9d9de-128">Запустите локальный веб-сервер.</span><span class="sxs-lookup"><span data-stu-id="9d9de-128">Start the local web server.</span></span>

    * <span data-ttu-id="9d9de-129">Если вы будете использовать Excel для Windows для тестирования ваших пользовательских функций, выполните следующую команду, чтобы запустить локальный веб-сервер, запустить программу Excel и загрузить неопубликованную надстройку:</span><span class="sxs-lookup"><span data-stu-id="9d9de-129">If you'll be using Excel for Windows to test your custom functions, run the following command to start the local web server, launch Excel, and sideload the add-in:</span></span>

        ```bash
        npm run start-desktop
        ```

    * <span data-ttu-id="9d9de-130">Если вы будете использовать Excel Online для тестирования ваших пользовательских функций, выполните следующую команду, чтобы запустить локальный веб-сервер:</span><span class="sxs-lookup"><span data-stu-id="9d9de-130">If you'll be using Excel Online to test your custom functions, run the following command to start the local web server:</span></span> 

        ```bash
        npm run start-web
        ```

## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="9d9de-131">Проверка работы готовой пользовательской функции</span><span class="sxs-lookup"><span data-stu-id="9d9de-131">Try out a prebuilt custom function</span></span>

<span data-ttu-id="9d9de-132">Проект пользовательских функций, созданный с помощью генератора Yo Office, содержит некоторые готовые пользовательские функции, определенные в файле **src/functions/functions.js**.</span><span class="sxs-lookup"><span data-stu-id="9d9de-132">The custom functions project that you created by using the Yo Office generator contains some prebuilt custom functions, defined within the **src/customfunction.js** file.</span></span> <span data-ttu-id="9d9de-133">Файл **manifest.xml** в корневом каталоге проекта указывает, что все пользовательские функции принадлежат пространству имен `CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="9d9de-133">The **manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.</span></span>

<span data-ttu-id="9d9de-134">Прежде чем вы сможете использовать любую из готовых пользовательских функций, необходимо зарегистрировать надстройку пользовательских функций в Excel.</span><span class="sxs-lookup"><span data-stu-id="9d9de-134">Before you can use any of the prebuilt custom functions, you must register the custom functions add-in in Excel.</span></span> <span data-ttu-id="9d9de-135">Сделайте это, выполнив действия для платформы, которую будете использовать в этом руководстве.</span><span class="sxs-lookup"><span data-stu-id="9d9de-135">Do so by completing steps for the platform that you'll be using in this tutorial.</span></span>

* <span data-ttu-id="9d9de-136">Если будет использоваться Excel for Windows для тестирования пользовательских функций, выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="9d9de-136">If you'll be using Excel for Windows to test your custom functions:</span></span>

    1. <span data-ttu-id="9d9de-137">В Excel выберите вкладку **Вставка**, а затем нажмите стрелку вниз, которая находится справа от пункта **Мои надстройки**. ![Вставьте ленту в Excel для Windows с выделенной стрелкой "Мои надстройки"](../images/excel-cf-register-add-in-1b.png).</span><span class="sxs-lookup"><span data-stu-id="9d9de-137">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

    2. <span data-ttu-id="9d9de-138">В списке доступных надстроек найдите раздел **Надстройки разработчика** и выберите надстройку **Пользовательские функции Excel**, чтобы зарегистрировать ее.</span><span class="sxs-lookup"><span data-stu-id="9d9de-138">In the list of available add-ins, find the **Developer Add-ins** section and select the **Excel Custom Functions** add-in to register it.</span></span>
        <span data-ttu-id="9d9de-139">![Вставьте ленту в Excel для Windows с выделенной надстройкой "Пользовательские функции Excel" в списке "Мои надстройки"](../images/excel-cf-register-add-in-2.png).</span><span class="sxs-lookup"><span data-stu-id="9d9de-139">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

* <span data-ttu-id="9d9de-140">Если вы будете использовать Excel Online для тестирования своих настраиваемых функций, выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="9d9de-140">If you'll be using Excel Online to test your custom functions:</span></span> 

    1. <span data-ttu-id="9d9de-141">В Excel Online выберите вкладку **Вставка**, а затем выберите **Надстройки**. ![Вставьте ленту в Excel Online с выделенным значком "Мои надстройки"](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="9d9de-141">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

    2. <span data-ttu-id="9d9de-142">Выберите пункт **Управление моими надстройками**, а затем выберите **Отправить мою надстройку**.</span><span class="sxs-lookup"><span data-stu-id="9d9de-142">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

    3. <span data-ttu-id="9d9de-143">Выберите \*\*Обзор... \*\* и откройте корневой каталог проекта, созданный генератором Yo Office.</span><span class="sxs-lookup"><span data-stu-id="9d9de-143">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span> 

    4. <span data-ttu-id="9d9de-144">Выберите файл **manifest.xml** и нажмите кнопку **Открыть**, затем нажмите кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="9d9de-144">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

<span data-ttu-id="9d9de-145">На этом этапе готовые пользовательские функции в вашем проекте уже загружены и доступны в Excel.</span><span class="sxs-lookup"><span data-stu-id="9d9de-145">At this point, the prebuilt custom functions in your project are loaded and available within Excel.</span></span> <span data-ttu-id="9d9de-146">Попробуйте, как работает пользовательская функция `ADD`, выполнив в Excel описанные далее действия.</span><span class="sxs-lookup"><span data-stu-id="9d9de-146">Try out the `ADD` custom function by completing the following steps in Excel:</span></span>

1. <span data-ttu-id="9d9de-147">Введите в ячейке **=CONTOSO**.</span><span class="sxs-lookup"><span data-stu-id="9d9de-147">Within a cell, type **=CONTOSO**.</span></span> <span data-ttu-id="9d9de-148">Обратите внимание на то, что в меню автозаполнения содержится список всех функций в пространстве имен `CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="9d9de-148">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="9d9de-149">Выполните запуск функции `CONTOSO.ADD` с числами `10` и `200` в качестве входных параметров, введя значение `=CONTOSO.ADD(10,200)` в ячейке и нажав клавишу ВВОД.</span><span class="sxs-lookup"><span data-stu-id="9d9de-149">Run the `CONTOSO.ADD` function, with numbers `10` and `200` as input parameters, by specifying the following value in the cell and pressing enter:</span></span>

<span data-ttu-id="9d9de-150">Пользовательская функция `ADD` вычисляет сумму двух чисел, которые вы указываете в качестве входных параметров.</span><span class="sxs-lookup"><span data-stu-id="9d9de-150">The `ADD` custom function computes the sum of the two numbers that you specify as input parameters.</span></span> <span data-ttu-id="9d9de-151">При вводе `=CONTOSO.ADD(10,200)` в ячейке должен отобразиться результат **210** после нажатия клавиши ВВОД.</span><span class="sxs-lookup"><span data-stu-id="9d9de-151">Typing `=CONTOSO.ADD(10,200)` should produce the result **210** in the cell after you press enter.</span></span>

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a><span data-ttu-id="9d9de-152">Создание пользовательской функции, которая запрашивает данные из Интернета</span><span class="sxs-lookup"><span data-stu-id="9d9de-152">Create a custom function that requests data from the web</span></span>

<span data-ttu-id="9d9de-153">Что делать, если требуется функция, которая сможет запросить цену на акцию из API и отобразить результат в ячейке на листе?</span><span class="sxs-lookup"><span data-stu-id="9d9de-153">What if you needed a function that could request the price of a stock from an API and display the result in the cell of a worksheet?</span></span> <span data-ttu-id="9d9de-154">Пользовательские функции разрабатываются таким образом, что вы можете легко асинхронно запросить данные из Интернета.</span><span class="sxs-lookup"><span data-stu-id="9d9de-154">Custom functions are designed so that you can easily request data from the web asynchronously.</span></span>

<span data-ttu-id="9d9de-155">Выполните указанные ниже действия, чтобы создать пользовательскую функцию с именем `stockPrice`, которая принимает код акции (например, **MSFT**) и возвращает цену этой акции.</span><span class="sxs-lookup"><span data-stu-id="9d9de-155">Complete the following steps to create a custom function named `stockPrice` that accepts a stock ticker (e.g., **MSFT**) and returns the price of that stock.</span></span> <span data-ttu-id="9d9de-156">Такая пользовательская функция использует API IEX Trading, который предоставляется бесплатно и не требует проверки подлинности.</span><span class="sxs-lookup"><span data-stu-id="9d9de-156">This custom function uses the IEX Trading API, which is free and does not require authentication.</span></span>

1. <span data-ttu-id="9d9de-157">В проекте **stock-ticker**, созданном генератором Yo Office, найдите файл **src/functions/functions.js** и откройте его в редакторе кода.</span><span class="sxs-lookup"><span data-stu-id="9d9de-157">In the **stock-ticker** project that the Yo Office generator created, find the file **src/customfunctions.js** and open it in your code editor.</span></span>

2. <span data-ttu-id="9d9de-158">Добавьте указанный ниже код в **customfunctions.js** и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="9d9de-158">Add the following code to **customfunctions.js** and save the file.</span></span>

    ```js
    function stockPrice(ticker) {
        var url = "https://api.iextrading.com/1.0/stock/" + ticker + "/price";
        return fetch(url)
            .then(function(response) {
                return response.text();
            })
            .then(function(text) {
                return parseFloat(text);
            });

        // Note: in case of an error, the returned rejected Promise
        //    will be bubbled up to Excel to indicate an error.
    }

    CustomFunctionMappings.STOCKPRICE = stockPrice;
    ```

3. <span data-ttu-id="9d9de-159">Прежде чем можно будет сделать в Excel такую функцию доступной конечным пользователям, необходимо указать метаданные, описывающие эту функцию.</span><span class="sxs-lookup"><span data-stu-id="9d9de-159">Before Excel can make this new function available to end-users, you must specify metadata that describes this function.</span></span> <span data-ttu-id="9d9de-160">В проекте **stock-ticker**, созданном генератором Yo Office, найдите файл **src/functions/functions.json** и откройте его в редакторе кода.</span><span class="sxs-lookup"><span data-stu-id="9d9de-160">In the **stock-ticker** project that the Yo Office generator created, find the file **config/customfunctions.json** and open it in your code editor.</span></span> <span data-ttu-id="9d9de-161">Добавьте указанный ниже объект в массив `functions` в файле **src/functions/functions.json** и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="9d9de-161">Add the following object to the `functions` array within the **config/customfunctions.json** file and save the file.</span></span>

    <span data-ttu-id="9d9de-162">Формат JSON описывает функцию `stockPrice`.</span><span class="sxs-lookup"><span data-stu-id="9d9de-162">This JSON describes the `stockPrice` function.</span></span>

    ```json
    {
        "id": "STOCKPRICE",
        "name": "STOCKPRICE",
        "description": "Fetches current stock price",
        "helpUrl": "http://www.contoso.com/help",
        "result": {
            "type": "number",
            "dimensionality": "scalar"
        },  
        "parameters": [
            {
                "name": "ticker",
                "description": "stock ticker name",
                "type": "string",
                "dimensionality": "scalar"
            }
        ]
    }
    ```

4. <span data-ttu-id="9d9de-163">Необходимо зарегистрировать надстройку в Excel, чтобы новая функция стала доступной конечным пользователям.</span><span class="sxs-lookup"><span data-stu-id="9d9de-163">You must reregister the add-in in Excel in order for the new function to be available to end-users.</span></span> <span data-ttu-id="9d9de-164">Выполните указанные ниже действия для платформы, которую вы используете в этом руководстве.</span><span class="sxs-lookup"><span data-stu-id="9d9de-164">Complete the following steps for the platform that you're using in this tutorial.</span></span>

    * <span data-ttu-id="9d9de-165">Если вы используете Excel для Windows, выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="9d9de-165">If you're using Excel for Windows:</span></span>

        1. <span data-ttu-id="9d9de-166">Закройте Excel, а затем откройте Excel повторно.</span><span class="sxs-lookup"><span data-stu-id="9d9de-166">Close Excel and then reopen Excel.</span></span>

        2. <span data-ttu-id="9d9de-167">В Excel выберите вкладку **Вставка**, а затем нажмите стрелку вниз, которая находится справа от пункта **Мои надстройки**. ![Вставьте ленту в Excel для Windows с выделенной стрелкой "Мои надстройки"](../images/excel-cf-register-add-in-1b.png).</span><span class="sxs-lookup"><span data-stu-id="9d9de-167">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

        1. <span data-ttu-id="9d9de-168">В списке доступных надстроек найдите раздел **Надстройки разработчика** и выберите надстройку **Пользовательские функции Excel**, чтобы зарегистрировать ее.</span><span class="sxs-lookup"><span data-stu-id="9d9de-168">In the list of available add-ins, find the **Developer Add-ins** section and select the **Excel Custom Functions** add-in to register it.</span></span>
            <span data-ttu-id="9d9de-169">![Вставьте ленту в Excel для Windows с выделенной надстройкой "Пользовательские функции Excel" в списке "Мои надстройки"](../images/excel-cf-register-add-in-2.png).</span><span class="sxs-lookup"><span data-stu-id="9d9de-169">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

    * <span data-ttu-id="9d9de-170">Если вы используете Excel Online, выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="9d9de-170">If you're using Excel Online:</span></span> 

        1. <span data-ttu-id="9d9de-171">В Excel Online выберите вкладку **Вставка**, а затем выберите **Надстройки**. ![Вставьте ленту в Excel Online с выделенным значком "Мои надстройки"](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="9d9de-171">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

        2. <span data-ttu-id="9d9de-172">Выберите пункт **Управление моими надстройками**, а затем выберите **Отправить мою надстройку**.</span><span class="sxs-lookup"><span data-stu-id="9d9de-172">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

        3. <span data-ttu-id="9d9de-173">Выберите \*\*Обзор... \*\* и откройте корневой каталог проекта, созданный генератором Yo Office.</span><span class="sxs-lookup"><span data-stu-id="9d9de-173">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span> 

        4. <span data-ttu-id="9d9de-174">Выберите файл **manifest.xml** и нажмите кнопку **Открыть**, затем нажмите кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="9d9de-174">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

5. <span data-ttu-id="9d9de-175">Теперь давайте попробуем, как работает новая функция.</span><span class="sxs-lookup"><span data-stu-id="9d9de-175">Now, let's try out the new function.</span></span> <span data-ttu-id="9d9de-176">В ячейке **B1** введите текст `=CONTOSO.STOCKPRICE("MSFT")` и нажмите клавишу ВВОД.</span><span class="sxs-lookup"><span data-stu-id="9d9de-176">In cell **B1**, type the text `=CONTOSO.STOCKPRICE("MSFT")` and press enter.</span></span> <span data-ttu-id="9d9de-177">Вы увидите, что результат в ячейке **B1** является текущей ценой одной акции корпорации Майкрософт.</span><span class="sxs-lookup"><span data-stu-id="9d9de-177">You should see that the result in cell **B1** is the current stock price for one share of Microsoft stock.</span></span>

## <a name="create-a-streaming-asynchronous-custom-function"></a><span data-ttu-id="9d9de-178">Создание потоковой асинхронной пользовательской функции</span><span class="sxs-lookup"><span data-stu-id="9d9de-178">Create a streaming asynchronous custom function</span></span>

<span data-ttu-id="9d9de-179">Функция `stockPrice`, которую вы только что создали, возвращает цену акции в конкретный момент времени, однако цены на акции всегда меняются.</span><span class="sxs-lookup"><span data-stu-id="9d9de-179">The `stockPrice` function that you just created returns the price of a stock at a specific moment in time, but stock prices are always changing.</span></span> <span data-ttu-id="9d9de-180">Давайте создадим пользовательскую функцию, которая осуществляет потоковую передачу данных из API, чтобы получать обновления цен на акции в реальном времени.</span><span class="sxs-lookup"><span data-stu-id="9d9de-180">Let's create a custom function that streams data from an API to get real-time updates on a stock price.</span></span>

<span data-ttu-id="9d9de-181">Выполните указанные ниже действия, чтобы создать функцию с именем `stockPriceStream`, которая будет запрашивать цену указанной акции каждые 1000 миллисекунд (при условии, что предыдущий запрос был выполнен).</span><span class="sxs-lookup"><span data-stu-id="9d9de-181">Complete the following steps to create a custom function named `stockPriceStream` that requests the price of the specified stock every 1000 milliseconds (provided that the previous request has completed).</span></span> <span data-ttu-id="9d9de-182">Во время выполнения первоначального запроса в ячейке, где вызывается функция, может появиться значение-заполнитель **#GETTING_DATA**.</span><span class="sxs-lookup"><span data-stu-id="9d9de-182">While the initial request is in-progress, you may see the placeholder value **#GETTING_DATA** the cell where the function is being called.</span></span> <span data-ttu-id="9d9de-183">Когда значение будет возвращено функцией, оно заменит значение-заполнитель **#GETTING_DATA** в ячейке.</span><span class="sxs-lookup"><span data-stu-id="9d9de-183">When a value is returned by the function, **#GETTING_DATA** will be replaced by that value in the cell.</span></span>

1. <span data-ttu-id="9d9de-184">В проекте **stock-ticker**, созданном генератором Yo Office, добавьте указанный ниже код в файл **src/functions/functions.js** и сохраните его.</span><span class="sxs-lookup"><span data-stu-id="9d9de-184">In the **stock-ticker** project that the Yo Office generator created, add the following code to **src/customfunctions.js** and save the file.</span></span>

    ```js
    function stockPriceStream(ticker, handler) {
        var updateFrequency = 1000 /* milliseconds*/;
        var isPending = false;

        var timer = setInterval(function() {
            // If there is already a pending request, skip this iteration:
            if (isPending) {
                return;
            }

            var url = "https://api.iextrading.com/1.0/stock/" + ticker + "/price";
            isPending = true;

            fetch(url)
                .then(function(response) {
                    return response.text();
                })
                .then(function(text) {
                    handler.setResult(parseFloat(text));
                })
                .catch(function(error) {
                    handler.setResult(error);
                })
                .then(function() {
                    isPending = false;
                });
        }, updateFrequency);

        handler.onCanceled = () => {
            clearInterval(timer);
        };
    }

    CustomFunctionMappings.STOCKPRICESTREAM = stockPriceStream;
    ```

2. <span data-ttu-id="9d9de-185">Прежде чем можно будет сделать в Excel такую функцию доступной конечным пользователям, необходимо указать метаданные, описывающие эту функцию.</span><span class="sxs-lookup"><span data-stu-id="9d9de-185">Before Excel can make this new function available to end-users, you must specify metadata that describes this function.</span></span> <span data-ttu-id="9d9de-186">В проекте **stock-ticker**, созданном генератором Yo Office, добавьте указанный ниже объект в массив `functions` в файле **src/functions/functions.json** и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="9d9de-186">In the **stock-ticker** project that the Yo Office generator created, add the following object to the `functions` array within the **config/customfunctions.json** file and save the file.</span></span>

    <span data-ttu-id="9d9de-187">Формат JSON описывает функцию `stockPriceStream`.</span><span class="sxs-lookup"><span data-stu-id="9d9de-187">This JSON describes the `stockPriceStream` function.</span></span> <span data-ttu-id="9d9de-188">Для любой функции потоковой передачи свойство `stream` и свойство `cancelable` должны быть заданы как `true` в объекте `options`, как показано в этом примере кода.</span><span class="sxs-lookup"><span data-stu-id="9d9de-188">For any streaming function, the `stream` property and the `cancelable` property must be set to `true` within the `options` object, as shown in this code sample.</span></span>

    ```json
    { 
        "id": "STOCKPRICESTREAM",
        "name": "STOCKPRICESTREAM",
        "description": "Streams real time stock price",
        "helpUrl": "http://www.contoso.com/help",
        "result": {
            "type": "number",
            "dimensionality": "scalar"
        },  
        "parameters": [
            {
                "name": "ticker",
                "description": "stock ticker name",
                "type": "string",
                "dimensionality": "scalar"
            }
        ],
        "options": {
            "stream": true,
            "cancelable": true
        }
    }
    ```

3. <span data-ttu-id="9d9de-189">Необходимо зарегистрировать надстройку в Excel, чтобы новая функция стала доступной конечным пользователям.</span><span class="sxs-lookup"><span data-stu-id="9d9de-189">You must reregister the add-in in Excel in order for the new function to be available to end-users.</span></span> <span data-ttu-id="9d9de-190">Выполните указанные ниже действия для платформы, которую вы используете в этом руководстве.</span><span class="sxs-lookup"><span data-stu-id="9d9de-190">Complete the following steps for the platform that you're using in this tutorial.</span></span>

    * <span data-ttu-id="9d9de-191">Если вы используете Excel для Windows, выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="9d9de-191">If you're using Excel for Windows:</span></span>

        1. <span data-ttu-id="9d9de-192">Закройте Excel, а затем откройте Excel повторно.</span><span class="sxs-lookup"><span data-stu-id="9d9de-192">Close Excel and then reopen Excel.</span></span>
        
        2. <span data-ttu-id="9d9de-193">В Excel выберите вкладку **Вставка**, а затем нажмите стрелку вниз, которая находится справа от пункта **Мои надстройки**. ![Вставьте ленту в Excel для Windows с выделенной стрелкой "Мои надстройки"](../images/excel-cf-register-add-in-1b.png).</span><span class="sxs-lookup"><span data-stu-id="9d9de-193">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

        3. <span data-ttu-id="9d9de-194">В списке доступных надстроек найдите раздел **Надстройки разработчика** и выберите надстройку **Пользовательские функции Excel**, чтобы зарегистрировать ее.</span><span class="sxs-lookup"><span data-stu-id="9d9de-194">In the list of available add-ins, find the **Developer Add-ins** section and select the **Excel Custom Functions** add-in to register it.</span></span>
            <span data-ttu-id="9d9de-195">![Вставьте ленту в Excel для Windows с выделенной надстройкой "Пользовательские функции Excel" в списке "Мои надстройки"](../images/excel-cf-register-add-in-2.png).</span><span class="sxs-lookup"><span data-stu-id="9d9de-195">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

    * <span data-ttu-id="9d9de-196">Если вы используете Excel Online, выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="9d9de-196">If you're using Excel Online:</span></span> 

        1. <span data-ttu-id="9d9de-197">В Excel Online выберите вкладку **Вставка**, а затем выберите **Надстройки**. ![Вставьте ленту в Excel Online с выделенным значком "Мои надстройки"](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="9d9de-197">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

        2. <span data-ttu-id="9d9de-198">Выберите пункт **Управление моими надстройками**, а затем выберите **Отправить мою надстройку**.</span><span class="sxs-lookup"><span data-stu-id="9d9de-198">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

        3. <span data-ttu-id="9d9de-199">Выберите \*\*Обзор... \*\* и откройте корневой каталог проекта, созданный генератором Yo Office.</span><span class="sxs-lookup"><span data-stu-id="9d9de-199">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span> 

        4. <span data-ttu-id="9d9de-200">Выберите файл **manifest.xml** и нажмите кнопку **Открыть**, затем нажмите кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="9d9de-200">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

4. <span data-ttu-id="9d9de-201">Теперь давайте попробуем, как работает новая функция.</span><span class="sxs-lookup"><span data-stu-id="9d9de-201">Now, let's try out the new function.</span></span> <span data-ttu-id="9d9de-202">В ячейке **C1** введите текст `=CONTOSO.STOCKPRICESTREAM("MSFT")` и нажмите клавишу ВВОД.</span><span class="sxs-lookup"><span data-stu-id="9d9de-202">In cell **C1**, type the text `=CONTOSO.STOCKPRICESTREAM("MSFT")` and press enter.</span></span> <span data-ttu-id="9d9de-203">Если рынок ценных бумаг открыт, вы увидите, что результат в ячейке **C1** постоянно обновляется, отражая в режиме реального времени цену одной акции корпорации Майкрософт.</span><span class="sxs-lookup"><span data-stu-id="9d9de-203">Provided that the stock market is open, you should see that the result in cell **C1** is constantly updated to reflect the real-time price for one share of Microsoft stock.</span></span>

## <a name="next-steps"></a><span data-ttu-id="9d9de-204">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="9d9de-204">Next steps</span></span>

<span data-ttu-id="9d9de-205">В ходе работы с данным руководством вы создали новый проект пользовательских функций, попробовали, как работает готовая функция, создали пользовательскую функцию, которая запрашивает данные из Интернета, а также создали пользовательскую функцию, которая осуществляет потоковую передачу данных в реальном времени из Интернета.</span><span class="sxs-lookup"><span data-stu-id="9d9de-205">In this tutorial, you've created a new custom functions project, tried out a prebuilt function, created a custom function that requests data from the web, and created a custom function that streams real-time data from the web.</span></span> <span data-ttu-id="9d9de-206">Чтобы узнать больше о пользовательских функции в Excel, перейдите к следующей статье:</span><span class="sxs-lookup"><span data-stu-id="9d9de-206">To learn more about custom functions in Excel, continue to the following article:</span></span> 

> [!div class="nextstepaction"]
> [<span data-ttu-id="9d9de-207">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="9d9de-207">Create custom functions in Excel</span></span>](../excel/custom-functions-overview.md)

## <a name="legal-information"></a><span data-ttu-id="9d9de-208">Юридические сведения</span><span class="sxs-lookup"><span data-stu-id="9d9de-208">Legal information</span></span>

<span data-ttu-id="9d9de-209">Данные предоставлены бесплатно компанией [IEX](https://iextrading.com/developer/).</span><span class="sxs-lookup"><span data-stu-id="9d9de-209">Data provided free by [IEX](https://iextrading.com/developer/).</span></span> <span data-ttu-id="9d9de-210">Ознакомьтесь с [Условиями использования IEX](https://iextrading.com/api-exhibit-a/).</span><span class="sxs-lookup"><span data-stu-id="9d9de-210">View [IEX's Terms of Use](https://iextrading.com/api-exhibit-a/).</span></span> <span data-ttu-id="9d9de-211">Корпорация Майкрософт использует API компании IEX в этом руководстве исключительно в ознакомительных целях.</span><span class="sxs-lookup"><span data-stu-id="9d9de-211">Microsoft's use of the IEX API in this tutorial is for educational purposes only.</span></span>
