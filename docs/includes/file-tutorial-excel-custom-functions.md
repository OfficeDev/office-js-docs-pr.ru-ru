# <a name="tutorial-create-custom-functions-in-excel"></a><span data-ttu-id="0dbdb-101">Руководство: создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="0dbdb-101">Tutorial: Create custom functions in Excel</span></span>

## <a name="introduction"></a><span data-ttu-id="0dbdb-102">Введение</span><span class="sxs-lookup"><span data-stu-id="0dbdb-102">Introduction</span></span>

<span data-ttu-id="0dbdb-103">Пользовательские функции позволяют добавлять новые функции в Excel путем определения этих функций в JavaScript как части надстройки.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-103">Custom functions enable you to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="0dbdb-104">Пользователи в Excel могут получить доступ к пользовательским функциям так же, как и к любой встроенной функции в Excel, например `SUM()`.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-104">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="0dbdb-105">Вы можете создавать пользовательские функции, которые будут выполнять простые задачи, такие как настраиваемые вычисления, или более сложные задачи, такие как потоковая передача данных в режиме реального времени из Интернета на лист.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-105">You can create custom functions that perform simple tasks such as custom calculations or more complex tasks such as streaming real-time data from the web into a worksheet.</span></span>

<span data-ttu-id="0dbdb-106">В этом руководстве описан порядок выполнения перечисленных ниже задач.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-106">In this tutorial, you will:</span></span>
> [!div class="checklist"]
> * <span data-ttu-id="0dbdb-107">Создание проекта пользовательских функций с помощью генератора Yo Office</span><span class="sxs-lookup"><span data-stu-id="0dbdb-107">Create a custom functions project by using the Yo Office generator</span></span>
> * <span data-ttu-id="0dbdb-108">Использование готовой пользовательской функции для выполнения простых вычислений</span><span class="sxs-lookup"><span data-stu-id="0dbdb-108">Use a prebuilt custom function to perform a simple calculation</span></span>
> * <span data-ttu-id="0dbdb-109">Создание пользовательской функции, которая запрашивает данные из Интернета</span><span class="sxs-lookup"><span data-stu-id="0dbdb-109">Create a custom function that requests data from the web</span></span>
> * <span data-ttu-id="0dbdb-110">Создание пользовательской функции, которая осуществляет потоковую передачу данных в реальном времени из Интернета</span><span class="sxs-lookup"><span data-stu-id="0dbdb-110">Create a custom function that streams real-time data from the web</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="prerequisites"></a><span data-ttu-id="0dbdb-111">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="0dbdb-111">Prerequisites</span></span>

* <span data-ttu-id="0dbdb-112">[Node.js](https://nodejs.org/en/) (версия 8.0.0 или более поздняя)</span><span class="sxs-lookup"><span data-stu-id="0dbdb-112">[Node.js](https://nodejs.org/en/) (version 8.0.0 or later)</span></span>

* <span data-ttu-id="0dbdb-113">[Git Bash](https://git-scm.com/downloads) (или другой клиент Git)</span><span class="sxs-lookup"><span data-stu-id="0dbdb-113">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

* <span data-ttu-id="0dbdb-114">Последняя версия [Yeoman](https://yeoman.io/) и [генератора Yeoman для надстроек Office](https://www.npmjs.com/package/generator-office). Выполните в командной строке указанную ниже команду, чтобы установить эти инструменты глобально.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-114">The latest version of [Yeoman](https://yeoman.io/) and the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office). To install these tools globally, run the following command from the command prompt:</span></span>

    ```
    npm install -g yo generator-office
    ```

    > [!NOTE]
    > <span data-ttu-id="0dbdb-115">Даже если у вас установлен генератор Yeoman, рекомендуется обновить пакет до последней версии из npm.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-115">Even if you have previously installed the Yeoman generator, we recommend updating your package to the latest version from npm.</span></span>

* <span data-ttu-id="0dbdb-116">Excel для Windows (64-разрядная версия 1810 или более поздняя) или Excel Online</span><span class="sxs-lookup"><span data-stu-id="0dbdb-116">Excel for Windows (version 1810 or later) or Excel Online</span></span>

* <span data-ttu-id="0dbdb-117">Присоединитесь к [Программе предварительной оценки Office](https://products.office.com/office-insider) (уровень **Участник**; ранее "Предварительная оценка — ранний доступ")</span><span class="sxs-lookup"><span data-stu-id="0dbdb-117">Join the [Office Insider program](https://products.office.com/office-insider) (**Insider** level -- formerly called "Insider Fast")</span></span>

## <a name="create-a-custom-functions-project"></a><span data-ttu-id="0dbdb-118">Создание проекта пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="0dbdb-118">Create a custom functions project</span></span>

 <span data-ttu-id="0dbdb-119">Чтобы начать работу, создайте проект пользовательских функций с помощью генератора Yeoman.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-119">To start, you'll use the Yeoman generator to create the custom functions project.</span></span> <span data-ttu-id="0dbdb-120">Это позволит настроить для проекта правильную структуру папок, исходные файлы и зависимости, чтобы начать написание кода пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-120">This will set up your project with the correct folder structure, source files, and dependencies to begin coding your custom functions.</span></span>

1. <span data-ttu-id="0dbdb-121">Выполните указанную ниже команду и ответьте на вопросы, как показано ниже.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-121">Run the following command and then answer the prompts as follows.</span></span>

    ```
    yo office
    ```

    * <span data-ttu-id="0dbdb-122">Выберите тип проекта: `Excel Custom Functions Add-in project (...)`</span><span class="sxs-lookup"><span data-stu-id="0dbdb-122">Choose a project type: `Excel Custom Functions Add-in project (...)`</span></span>

    * <span data-ttu-id="0dbdb-123">Выберите тип сценария: `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="0dbdb-123">Choose a script type: `JavaScript`</span></span>

    * <span data-ttu-id="0dbdb-124">Как вы хотите назвать свою надстройку?</span><span class="sxs-lookup"><span data-stu-id="0dbdb-124">What do you want to name your add-in?</span></span> `stock-ticker`

    ![Генератор Yeoman для надстройки Office, приглашающий к созданию пользовательских функций](../images/12-10-fork-cf-pic.jpg)

    <span data-ttu-id="0dbdb-126">Генератор Yeoman создаст файлы проекта и установит вспомогательные компоненты Node.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-126">The generator will create the project and install supporting Node components.</span></span> <span data-ttu-id="0dbdb-127">Файлы проекта взяты из репозитория [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) GitHub.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-127">The project files come from the [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) GitHub repository.</span></span>

2. <span data-ttu-id="0dbdb-128">Перейдите в папку проекта.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-128">Go to the project folder.</span></span>

    ```
    cd stock-ticker
    ```

3. <span data-ttu-id="0dbdb-129">Сделайте доверенным самозаверяющий сертификат, необходимый для выполнения этого проекта.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-129">Trust the self-signed certificate that is needed to run this project.</span></span> <span data-ttu-id="0dbdb-130">Подробные инструкции для Windows или Mac см. в статье [Добавление самозаверяющих сертификатов в качестве доверенных корневых сертификатов](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span><span class="sxs-lookup"><span data-stu-id="0dbdb-130">For detailed instructions for either Windows or Mac, see [Adding Self Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span></span>  

4. <span data-ttu-id="0dbdb-131">Выполните сборку проекта.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-131">Build the project.</span></span>

    ```
    npm run build
    ```

5. <span data-ttu-id="0dbdb-132">Запустите локальный веб-сервер, работающий на Node.js.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-132">Start the local web server, which runs in Node.js.</span></span>

    * <span data-ttu-id="0dbdb-133">Если вы будете использовать Excel для Windows для тестирования ваших пользовательских функций, выполните следующую команду, чтобы запустить локальный веб-сервер, запустить программу Excel и загрузить неопубликованную надстройку:</span><span class="sxs-lookup"><span data-stu-id="0dbdb-133">If you'll be using Excel for Windows to test your custom functions, run the following command to start the local web server, launch Excel, and sideload the add-in:</span></span>

        ```
         npm run start
        ```
        <span data-ttu-id="0dbdb-134">После выполнения этой команды, в командной строке отобразятся сведения о выполненных действиях, откроется другое окно npm со сведениями о сборке, и запустится Excel с загруженной надстройкой.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-134">After running this command, your command prompt will show details about what has been done, another npm window will open showing the details of the build, and Excel will start with your add-in loaded.</span></span> <span data-ttu-id="0dbdb-135">Если надстройка не загружается, проверьте правильность выполнения шага 3.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-135">If you add-in does not load, check that you have completed step 3 properly.</span></span>

    * <span data-ttu-id="0dbdb-136">Если вы будете использовать Excel Online для тестирования ваших пользовательских функций, выполните следующую команду, чтобы запустить локальный веб-сервер:</span><span class="sxs-lookup"><span data-stu-id="0dbdb-136">If you'll be using Excel Online to test your custom functions, run the following command to start the local web server:</span></span>

        ```
        npm run start-web
        ```

         <span data-ttu-id="0dbdb-137">После выполнения этой команды откроется другое окно со сведениями о сборке.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-137">After running this command, another window will open showing you the details of the build.</span></span> <span data-ttu-id="0dbdb-138">Чтобы использовать свои функции, откройте новую книгу в Office Online.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-138">To use your functions, open a new workbook in Office Online.</span></span>

## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="0dbdb-139">Проверка работы готовой пользовательской функции</span><span class="sxs-lookup"><span data-stu-id="0dbdb-139">Try out a prebuilt custom function</span></span>

<span data-ttu-id="0dbdb-140">Проект пользовательских функций, созданный с помощью генератора Yeoman, содержит некоторые готовые пользовательские функции, определенные в файле **src/customfunctions.js**.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-140">The custom functions project that you created by using the Yo Office generator contains some prebuilt custom functions, defined within the **src/functions/functions.js** file.</span></span> <span data-ttu-id="0dbdb-141">Файл **manifest.xml** в корневом каталоге проекта указывает, что все пользовательские функции принадлежат пространству имен `CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-141">The **manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.</span></span>

<span data-ttu-id="0dbdb-142">В книге Excel попробуйте, как работает пользовательская функция `ADD`, выполнив описанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-142">In your Excel workbook, try out the `ADD` custom function by completing the following steps in Excel:</span></span>

1. <span data-ttu-id="0dbdb-143">Введите в ячейке **=CONTOSO**.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-143">Within a cell, type **=CONTOSO**.</span></span> <span data-ttu-id="0dbdb-144">Обратите внимание на то, что в меню автозаполнения содержится список всех функций в пространстве имен `CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-144">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="0dbdb-145">Выполните запуск функции `CONTOSO.ADD` с числами `10` и `200` в качестве входных параметров, введя значение `=CONTOSO.ADD(10,200)` в ячейке и нажав клавишу ВВОД.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-145">Run the `CONTOSO.ADD` function, with numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.</span></span>

<span data-ttu-id="0dbdb-146">Пользовательская функция `ADD` вычисляет сумму двух чисел, которые вы указываете в качестве входных параметров.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-146">The `ADD` custom function computes the sum of the two numbers that you specify as input parameters.</span></span> <span data-ttu-id="0dbdb-147">При вводе `=CONTOSO.ADD(10,200)` в ячейке должен отобразиться результат **210** после нажатия клавиши ВВОД.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-147">Typing `=CONTOSO.ADD(10,200)` should produce the result **210** in the cell after you press enter.</span></span>

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a><span data-ttu-id="0dbdb-148">Создание пользовательской функции, которая запрашивает данные из Интернета</span><span class="sxs-lookup"><span data-stu-id="0dbdb-148">Create a custom function that requests data from the web</span></span>

<span data-ttu-id="0dbdb-149">Что делать, если требуется функция, которая сможет запросить цену на акцию из API и отобразить результат в ячейке на листе?</span><span class="sxs-lookup"><span data-stu-id="0dbdb-149">What if you needed a function that could request the price of a stock from an API and display the result in the cell of a worksheet?</span></span> <span data-ttu-id="0dbdb-150">Пользовательские функции разрабатываются таким образом, что вы можете легко асинхронно запросить данные из Интернета.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-150">Custom functions are designed so that you can easily request data from the web asynchronously.</span></span>

<span data-ttu-id="0dbdb-151">Выполните указанные ниже действия, чтобы создать пользовательскую функцию с именем `stockPrice`, которая принимает код акции (например, **MSFT**) и возвращает цену этой акции.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-151">Complete the following steps to create a custom function named `stockPrice` that accepts a stock ticker (e.g., **MSFT**) and returns the price of that stock.</span></span> <span data-ttu-id="0dbdb-152">Такая пользовательская функция использует API IEX Trading, который предоставляется бесплатно и не требует проверки подлинности.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-152">This custom function uses the IEX Trading API, which is free and does not require authentication.</span></span>

1. <span data-ttu-id="0dbdb-153">В проекте **stock-ticker**, созданном генератором Yeoman, найдите файл **src/customfunctions.js** и откройте его в редакторе кода.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-153">In the **stock-ticker** project that the Yo Office generator created, find the file **src/customfunctions.js** and open it in your code editor.</span></span>

2. <span data-ttu-id="0dbdb-154">В файле **customfunctions.js** найдите функцию `increment` и добавьте приведенный ниже код сразу после этой функции.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-154">In **customfunctions.js**, locate the `increment` function and add the following code immediately after that function.</span></span>

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

3. In **customfunctions.js**, locate the line`CustomFunctionMappings.INCREMENT = increment;`, add the following line of code immediately after that line, and save the file.

    ```js
    CustomFunctionMappings.STOCKPRICE = stockPrice;
    ```

4. <span data-ttu-id="0dbdb-155">Прежде чем можно будет сделать в Excel такую функцию доступной, необходимо указать метаданные, чтобы описать функцию для Excel.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-155">Before Excel can make this new function available to end-users, you must specify metadata that describes this function.</span></span> <span data-ttu-id="0dbdb-156">Откройте файл **config/customfunctions.json**.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-156">Open the **config/customfunctions.json** file.</span></span> <span data-ttu-id="0dbdb-157">Добавьте указанный ниже объект JSON в массив 'functions' и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-157">Add the following object to the  array within the src/functions/functions.json file and save the file.</span></span>

    <span data-ttu-id="0dbdb-158">Объект JSON описывает функцию `stockPrice`.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-158">This JSON describes the `stockPrice` function.</span></span>

    ```JSON
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

5. <span data-ttu-id="0dbdb-159">Необходимо повторно зарегистрировать надстройку в Excel, чтобы новая функция стала доступной конечным пользователям.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-159">You must reregister the add-in in Excel in order for the new function to be available to end-users.</span></span> <span data-ttu-id="0dbdb-160">Выполните указанные ниже действия для платформы, которую вы используете в этом руководстве.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-160">Complete the following steps for the platform that you're using in this tutorial.</span></span>

    * <span data-ttu-id="0dbdb-161">Если вы используете Excel для Windows, выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-161">If you're using Excel for Windows:</span></span>

        1. <span data-ttu-id="0dbdb-162">Закройте Excel, а затем откройте Excel повторно.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-162">Close Excel and then reopen Excel.</span></span>

        2. <span data-ttu-id="0dbdb-163">В Excel выберите вкладку **Вставка**, а затем нажмите стрелку вниз, которая находится справа от пункта **Мои надстройки**. ![Вставьте ленту в Excel для Windows с выделенной стрелкой "Мои надстройки"](../images/excel-cf-register-add-in-1b.png).</span><span class="sxs-lookup"><span data-stu-id="0dbdb-163">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

        3. <span data-ttu-id="0dbdb-164">В списке доступных надстроек найдите раздел **Надстройки разработчика** и выберите надстройку **stock-ticker**, чтобы зарегистрировать ее.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-164">In the list of available add-ins, find the Developer Add-ins section and select the your add-in to register it.</span></span>
            <span data-ttu-id="0dbdb-165">![Вставьте ленту в Excel для Windows с выделенной надстройкой "Пользовательские функции Excel" в списке "Мои надстройки"](../images/excel-cf-register-add-in-2.png).</span><span class="sxs-lookup"><span data-stu-id="0dbdb-165">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

    * <span data-ttu-id="0dbdb-166">Если вы используете Excel Online, выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-166">If you're using Excel Online:</span></span>

        1. <span data-ttu-id="0dbdb-167">В Excel Online выберите вкладку **Вставка**, а затем выберите **Надстройки**. ![Вставьте ленту в Excel Online с выделенным значком "Мои надстройки"](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="0dbdb-167">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

        2. <span data-ttu-id="0dbdb-168">Выберите пункт **Управление моими надстройками**, а затем выберите **Отправить мою надстройку**.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-168">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

        3. <span data-ttu-id="0dbdb-169">Выберите \*\*Обзор... \*\* и откройте корневой каталог проекта, созданный генератором Yeoman.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-169">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span> 

        4. <span data-ttu-id="0dbdb-170">Выберите файл **manifest.xml** и нажмите кнопку **Открыть**, затем нажмите кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-170">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

6. <span data-ttu-id="0dbdb-171">Теперь давайте попробуем, как работает новая функция.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-171">Now, let's try out the new function.</span></span> <span data-ttu-id="0dbdb-172">В ячейке **B1** введите текст `=CONTOSO.STOCKPRICE("MSFT")` и нажмите клавишу ВВОД.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-172">In cell **B1**, type the text `=CONTOSO.STOCKPRICE("MSFT")` and press enter.</span></span> <span data-ttu-id="0dbdb-173">Вы увидите, что результат в ячейке **B1** является текущей ценой одной акции корпорации Майкрософт.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-173">You should see that the result in cell **B1** is the current stock price for one share of Microsoft stock.</span></span>

## <a name="create-a-streaming-asynchronous-custom-function"></a><span data-ttu-id="0dbdb-174">Создание потоковой асинхронной пользовательской функции</span><span class="sxs-lookup"><span data-stu-id="0dbdb-174">Create a streaming asynchronous custom function</span></span>

<span data-ttu-id="0dbdb-175">Функция `stockPrice`, которую вы только что создали, возвращает цену акции в конкретный момент времени, однако цены на акции всегда меняются.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-175">The `stockPrice` function that you just created returns the price of a stock at a specific moment in time, but stock prices are always changing.</span></span> <span data-ttu-id="0dbdb-176">Давайте создадим пользовательскую функцию, которая осуществляет потоковую передачу данных из API, чтобы получать обновления цен на акции в реальном времени.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-176">Let's create a custom function that streams data from an API to get real-time updates on a stock price.</span></span>

<span data-ttu-id="0dbdb-177">Выполните указанные ниже действия, чтобы создать функцию с именем `stockPriceStream`, которая будет запрашивать цену указанной акции каждые 1000 миллисекунд (при условии, что предыдущий запрос был выполнен).</span><span class="sxs-lookup"><span data-stu-id="0dbdb-177">Complete the following steps to create a custom function named `stockPriceStream` that requests the price of the specified stock every 1000 milliseconds (provided that the previous request has completed).</span></span> <span data-ttu-id="0dbdb-178">Во время выполнения первоначального запроса в ячейке, где вызывается функция, может появиться значение-заполнитель **#GETTING_DATA**.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-178">While the initial request is in-progress, you may see the placeholder value **#GETTING_DATA** the cell where the function is being called.</span></span> <span data-ttu-id="0dbdb-179">Когда значение будет возвращено функцией, оно заменит значение-заполнитель **#GETTING_DATA** в ячейке.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-179">When a value is returned by the function, **#GETTING_DATA** will be replaced by that value in the cell.</span></span>

1. <span data-ttu-id="0dbdb-180">В проекте **stock-ticker**, созданном генератором Yeoman, добавьте указанный ниже код в файл **src/customfunctions.js** и сохраните его.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-180">In the **stock-ticker** project that the Yo Office generator created, add the following code to **src/customfunctions.js** and save the file.</span></span>

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

2. <span data-ttu-id="0dbdb-181">Прежде чем можно будет сделать в Excel такую функцию доступной пользователям, укажите метаданные, описывающие эту функцию.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-181">Before Excel can make this new function available to end-users, you must specify metadata that describes this function.</span></span> <span data-ttu-id="0dbdb-182">В проекте **stock-ticker**, созданном генератором Yeoman, добавьте указанный ниже объект в массив `functions` в файле **config/customfunctions.json** и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-182">In the **stock-ticker** project that the Yo Office generator created, add the following object to the `functions` array within the **config/customfunctions.json** file and save the file.</span></span>

    <span data-ttu-id="0dbdb-183">Объект JSON описывает функцию `stockPriceStream`.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-183">This JSON describes the `stockPriceStream` function.</span></span> <span data-ttu-id="0dbdb-184">Для любой функции потоковой передачи свойство `stream` и свойство `cancelable` должны быть заданы как `true` в объекте `options`, как показано в этом примере кода.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-184">For any streaming function, the `stream` property and the `cancelable` property must be set to `true` within the `options` object, as shown in this code sample.</span></span>

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

3. <span data-ttu-id="0dbdb-185">Необходимо повторно зарегистрировать надстройку в Excel, чтобы новая функция стала доступной конечным пользователям.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-185">You must reregister the add-in in Excel in order for the new function to be available to end-users.</span></span> <span data-ttu-id="0dbdb-186">Выполните указанные ниже действия для платформы, которую вы используете в этом руководстве.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-186">Complete the following steps for the platform that you're using in this tutorial.</span></span>

    * <span data-ttu-id="0dbdb-187">Если вы используете Excel для Windows, выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-187">If you're using Excel for Windows:</span></span>

        1. <span data-ttu-id="0dbdb-188">Закройте Excel, а затем откройте Excel повторно.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-188">Close Excel and then reopen Excel.</span></span>
        
        2. <span data-ttu-id="0dbdb-189">В Excel выберите вкладку **Вставка**, а затем нажмите стрелку вниз, которая находится справа от пункта **Мои надстройки**. ![Вставьте ленту в Excel для Windows с выделенной стрелкой "Мои надстройки"](../images/excel-cf-register-add-in-1b.png).</span><span class="sxs-lookup"><span data-stu-id="0dbdb-189">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

        3. <span data-ttu-id="0dbdb-190">В списке доступных надстроек найдите раздел **Надстройки разработчика** и выберите надстройку **stock-ticker**, чтобы зарегистрировать ее.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-190">In the list of available add-ins, find the Developer Add-ins section and select the your add-in to register it.</span></span>
            <span data-ttu-id="0dbdb-191">![Вставьте ленту в Excel для Windows с выделенной надстройкой "Пользовательские функции Excel" в списке "Мои надстройки"](../images/excel-cf-register-add-in-2.png).</span><span class="sxs-lookup"><span data-stu-id="0dbdb-191">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

    * <span data-ttu-id="0dbdb-192">Если вы используете Excel Online, выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-192">If you're using Excel Online:</span></span>

        1. <span data-ttu-id="0dbdb-193">В Excel Online выберите вкладку **Вставка**, а затем выберите **Надстройки**. ![Вставьте ленту в Excel Online с выделенным значком "Мои надстройки"](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="0dbdb-193">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

        2. <span data-ttu-id="0dbdb-194">Выберите пункт **Управление моими надстройками**, а затем выберите **Отправить мою надстройку**.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-194">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

        3. <span data-ttu-id="0dbdb-195">Выберите \*\*Обзор... \*\* и откройте корневой каталог проекта, созданный генератором Yeoman.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-195">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span>

        4. <span data-ttu-id="0dbdb-196">Выберите файл **manifest.xml** и нажмите кнопку **Открыть**, затем нажмите кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-196">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

4. <span data-ttu-id="0dbdb-197">Теперь давайте попробуем, как работает новая функция.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-197">Now, let's try out the new function.</span></span> <span data-ttu-id="0dbdb-198">В ячейке **C1** введите текст `=CONTOSO.STOCKPRICESTREAM("MSFT")` и нажмите клавишу ВВОД.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-198">In cell **C1**, type the text `=CONTOSO.STOCKPRICESTREAM("MSFT")` and press enter.</span></span> <span data-ttu-id="0dbdb-199">Если рынок ценных бумаг открыт, вы увидите, что результат в ячейке **C1** постоянно обновляется, отражая в режиме реального времени цену одной акции корпорации Майкрософт.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-199">Provided that the stock market is open, you should see that the result in cell **C1** is constantly updated to reflect the real-time price for one share of Microsoft stock.</span></span>

## <a name="next-steps"></a><span data-ttu-id="0dbdb-200">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="0dbdb-200">Next steps</span></span>

<span data-ttu-id="0dbdb-201">В ходе работы с данным руководством вы создали новый проект пользовательских функций, попробовали, как работает готовая функция, создали пользовательскую функцию, которая запрашивает данные из Интернета, а также создали пользовательскую функцию, которая осуществляет потоковую передачу данных в реальном времени из Интернета.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-201">In this tutorial, you've created a new custom functions project, tried out a prebuilt function, created a custom function that requests data from the web, and created a custom function that streams real-time data from the web.</span></span> <span data-ttu-id="0dbdb-202">Чтобы узнать больше о пользовательских функции в Excel, перейдите к следующей статье:</span><span class="sxs-lookup"><span data-stu-id="0dbdb-202">To learn more about custom functions in Excel, continue to the following article:</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="0dbdb-203">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="0dbdb-203">Create custom functions in Excel</span></span>](../excel/custom-functions-overview.md)

## <a name="legal-information"></a><span data-ttu-id="0dbdb-204">Юридические сведения</span><span class="sxs-lookup"><span data-stu-id="0dbdb-204">Legal information</span></span>

<span data-ttu-id="0dbdb-205">Данные предоставлены бесплатно компанией [IEX](https://iextrading.com/developer/).</span><span class="sxs-lookup"><span data-stu-id="0dbdb-205">Data provided free by [IEX](https://iextrading.com/developer/).</span></span> <span data-ttu-id="0dbdb-206">Ознакомьтесь с [Условиями использования IEX](https://iextrading.com/api-exhibit-a/).</span><span class="sxs-lookup"><span data-stu-id="0dbdb-206">View [IEX's Terms of Use](https://iextrading.com/api-exhibit-a/).</span></span> <span data-ttu-id="0dbdb-207">Корпорация Майкрософт использует API компании IEX в этом руководстве исключительно в ознакомительных целях.</span><span class="sxs-lookup"><span data-stu-id="0dbdb-207">Microsoft's use of the IEX API in this tutorial is for educational purposes only.</span></span>
