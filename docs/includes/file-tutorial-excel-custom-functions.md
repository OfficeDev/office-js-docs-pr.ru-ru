# <a name="tutorial-create-custom-functions-in-excel"></a><span data-ttu-id="249ca-101">Урок: создание настраиваемых функций в Excel</span><span class="sxs-lookup"><span data-stu-id="249ca-101">Create custom functions in Excel (Preview)</span></span>

## <a name="introduction"></a><span data-ttu-id="249ca-102">Введение</span><span class="sxs-lookup"><span data-stu-id="249ca-102">Introduction</span></span>

<span data-ttu-id="249ca-p101">Настраиваемые функции позволяют добавлять новые функции в Excel, определяя эти функции в JavaScript как часть надстройки.  Пользователи в Excel могут получать доступ к настраиваемым функциям так же, как к любой собственной функции в Excel, например  `SUM()`. Можно создавать настраиваемые функции, которые будут выполнять простые задачи, такие как настраиваемые вычисления, или более сложные задачи, например потоковая передача данных в режиме реального времени из Интернета в лист таблицы.</span><span class="sxs-lookup"><span data-stu-id="249ca-p101">Custom functions enable you to add new functions to Excel by defining those functions in JavaScript as part of an add-in. Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`. You can create custom functions that perform simple tasks such as custom calculations or more complex tasks such as streaming real-time data from the web into a worksheet.</span></span>

<span data-ttu-id="249ca-106">В этом руководстве вы:</span><span class="sxs-lookup"><span data-stu-id="249ca-106">In this tutorial, you will use Visual Studio Code.</span></span>
> [!div class="checklist"]
> * <span data-ttu-id="249ca-107">Создадите проект с настраиваемыми функциями, используя генератор Yo Office.</span><span class="sxs-lookup"><span data-stu-id="249ca-107">Create a custom functions project by using the Yo Office generator</span></span>
> * <span data-ttu-id="249ca-108">Используете готовую настраиваемую функцию для выполнения простых вычислений.</span><span class="sxs-lookup"><span data-stu-id="249ca-108">Use a prebuilt custom function to perform a simple calculation</span></span>
> * <span data-ttu-id="249ca-109">Создадите настраиваемую функцию, которая будет запрашивать данные с веб-сайта.</span><span class="sxs-lookup"><span data-stu-id="249ca-109">Create a custom function that requests data from the web</span></span>
> * <span data-ttu-id="249ca-110">Создадите настраиваемую функцию, которая будет передавать данные в реальном времени с веб-сайта.</span><span class="sxs-lookup"><span data-stu-id="249ca-110">Create a custom function that streams real-time data from the web</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="prerequisites"></a><span data-ttu-id="249ca-111">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="249ca-111">Prerequisites</span></span>

* [<span data-ttu-id="249ca-112">Node.js и npm.</span><span class="sxs-lookup"><span data-stu-id="249ca-112">Node and npm</span></span>](https://nodejs.org/en/)

* <span data-ttu-id="249ca-113">[Git Bash](https://git-scm.com/downloads) (или другой клиент Git)</span><span class="sxs-lookup"><span data-stu-id="249ca-113">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

* <span data-ttu-id="249ca-p102">Последняя версия [Yeoman](http://yeoman.io/) и [Генератор Yo Office](https://www.npmjs.com/package/generator-office). Чтобы установить эти средства глобально, выполните следующую команду из командной строки:</span><span class="sxs-lookup"><span data-stu-id="249ca-p102">The latest version of [Yeoman](http://yeoman.io/) and the [Yo Office generator](https://www.npmjs.com/package/generator-office). To install these tools globally, run the following command via the command prompt:</span></span>

    ```bash
    npm install -g yo generator-office
    ```

* <span data-ttu-id="249ca-116">Excel для Windows (сборка 10827 или более поздняя версия) или Excel Online.</span><span class="sxs-lookup"><span data-stu-id="249ca-116">Excel for Windows (build number 10827 or later) or Excel Online</span></span>

* [<span data-ttu-id="249ca-117">Примите участие в программе предварительной оценки Office.</span><span class="sxs-lookup"><span data-stu-id="249ca-117">Join the Office Insider program</span></span>](https://products.office.com/office-insider)

## <a name="create-a-custom-functions-project"></a><span data-ttu-id="249ca-118">Создание проекта настраиваемых функций</span><span class="sxs-lookup"><span data-stu-id="249ca-118">Create a custom functions project by using the Yo Office generator</span></span>

<span data-ttu-id="249ca-119">Вы начнете этот урок с использования генератора Yo Office для создания файлов, необходимых для проекта настраиваемых функций.</span><span class="sxs-lookup"><span data-stu-id="249ca-119">You’ll begin this tutorial by using the Yo Office generator to create the files that you need for your custom functions project.</span></span>

1. <span data-ttu-id="249ca-120">Выполните приведенную ниже команду и ответьте на запросы, как показано ниже.</span><span class="sxs-lookup"><span data-stu-id="249ca-120">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    * <span data-ttu-id="249ca-121">Выберите тип проекта: `Excel Custom Functions Add-in project (...)`</span><span class="sxs-lookup"><span data-stu-id="249ca-121">Choose a project type  </span></span>
    * <span data-ttu-id="249ca-122">Выберите тип сценария: `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="249ca-122">Choose a script type  </span></span>
    * <span data-ttu-id="249ca-123">Как вы хотите назвать надстройку?</span><span class="sxs-lookup"><span data-stu-id="249ca-123">What do you want to name your add-in?</span></span> `stock-ticker`

    ![Yo Office bash запросит настраиваемые функции.](../images/yo-office-cfs-stock-ticker-3.png)

    <span data-ttu-id="249ca-125">После завершения работы мастера генератор создаст файлы проекта и установит вспомогательные компоненты Node.</span><span class="sxs-lookup"><span data-stu-id="249ca-125">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span> <span data-ttu-id="249ca-126">Файлы проекта поступают из репозитория GitHub [настраиваемых функций Excel](https://github.com/OfficeDev/Excel-Custom-Functions) .</span><span class="sxs-lookup"><span data-stu-id="249ca-126">The project files come from the [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) GitHub repository.</span></span>

2. <span data-ttu-id="249ca-127">Перейдите в папку проекта.</span><span class="sxs-lookup"><span data-stu-id="249ca-127">Navigate to the project folder:</span></span>

    ```bash
    cd stock-ticker
    ```

3. <span data-ttu-id="249ca-128">Запустите локальный веб-сервер.</span><span class="sxs-lookup"><span data-stu-id="249ca-128">Start the local web server.</span></span>

    * <span data-ttu-id="249ca-129">При использовании Excel для Windows для тестирования настраиваемых функций выполните следующую команду для запуска локального веб-сервера, запустите Excel и загрузите надстройку:</span><span class="sxs-lookup"><span data-stu-id="249ca-129">If you'll be using Excel for Windows to test your custom functions, run the following command to start the local web server, launch Excel, and sideload the add-in:</span></span>

        ```bash
        npm start
        ```

    * <span data-ttu-id="249ca-130">При использовании Excel Online для тестирования настраиваемых функций выполните следующую команду для запуска локального веб-сервера:</span><span class="sxs-lookup"><span data-stu-id="249ca-130">If you'll be using Excel Online to test your custom functions, run the following command to start the local web server:</span></span> 

        ```bash
        npm run start-web
        ```

## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="249ca-131">Испытание готовой настраиваемой функции</span><span class="sxs-lookup"><span data-stu-id="249ca-131">Try out a prebuilt custom function</span></span>

<span data-ttu-id="249ca-132">Проект настраиваемых функций, созданный с помощью генератора Yo Office, содержит некоторые готовые настраиваемые функции, определенные в файле **src/customfunction.js** .</span><span class="sxs-lookup"><span data-stu-id="249ca-132">The custom functions project that you created by using the Yo Office generator contains some prebuilt custom functions, defined within the **src/customfunction.js** file.</span></span> <span data-ttu-id="249ca-133">Файл **manifest.xml** в корневом каталоге проекта указывает, что все настраиваемые функции принадлежат пространству имен `CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="249ca-133">The **manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.</span></span>

<span data-ttu-id="249ca-134">Прежде чем использовать любые готовые настраиваемые функции, необходимо зарегистрировать надстройку настраиваемых функций в Excel.</span><span class="sxs-lookup"><span data-stu-id="249ca-134">Before you can use any of the prebuilt custom functions, you must register the custom functions add-in in Excel.</span></span> <span data-ttu-id="249ca-135">Сделайте это, выполнив нужные шаги для той платформы, которая будет использоваться в этом руководстве.</span><span class="sxs-lookup"><span data-stu-id="249ca-135">Do so by completing steps for the platform that you'll be using in this tutorial.</span></span>

* <span data-ttu-id="249ca-136">При использовании Excel для Windows для тестирования настраиваемых функций:</span><span class="sxs-lookup"><span data-stu-id="249ca-136">If you'll be using Excel for Windows to test your custom functions:</span></span>

    1. <span data-ttu-id="249ca-137">В Excel перейдите на вкладку **Вставка** и нажмите стрелку вниз, расположенную справа от раздела **Мои надстройки**.  ![Вставьте ленту в Excel для Windows, используя выделенную стрелку "Мои надстройки".](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="249ca-137">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

    2. <span data-ttu-id="249ca-138">В списке доступных надстроек найдите раздел **Надстройки разработчика** и выберите надстройку **Настраиваемые функции Excel** для ее регистрации.</span><span class="sxs-lookup"><span data-stu-id="249ca-138">In the list of available add-ins, find the **Developer Add-ins** section and select the **Excel Custom Functions** add-in to register it.</span></span>
        <span data-ttu-id="249ca-139">![Вставьте ленту в Excel для Windows с помощью надстройки настраиваемых функций Excel, выделенной в списке "Мои надстройки".](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="249ca-139">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

* <span data-ttu-id="249ca-140">При использовании Excel Online для тестирования настраиваемых функций:</span><span class="sxs-lookup"><span data-stu-id="249ca-140">If you'll be using Excel Online to test your custom functions:</span></span> 

    1. <span data-ttu-id="249ca-141">В Excel Online перейдите на вкладку **Вставка** и выберите **Надстройки**.  ![Вставьте ленту в Excel Online, используя выделенный значок "Мои надстройки".](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="249ca-141">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

    2. <span data-ttu-id="249ca-142">Выберите **Управление моими надстройками** и **Отправить мою надстройку**.</span><span class="sxs-lookup"><span data-stu-id="249ca-142">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

    3. <span data-ttu-id="249ca-143">Нажмите кнопку **Обзор...** и перейдите в корневой каталог проекта, созданный генератором Yo Office.</span><span class="sxs-lookup"><span data-stu-id="249ca-143">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span> 

    4. <span data-ttu-id="249ca-144">Выберите файл **manifest.xml** и нажмите кнопки **Открыть**, а затем **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="249ca-144">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

<span data-ttu-id="249ca-145">На этом этапе готовые настраиваемые функции в вашем проекте уже загружены и доступны в Excel.</span><span class="sxs-lookup"><span data-stu-id="249ca-145">At this point, the prebuilt custom functions in your project are loaded and available within Excel.</span></span> <span data-ttu-id="249ca-146">Испытайте `ADD` настраиваемую функцию, выполнив следующие действия в Excel:</span><span class="sxs-lookup"><span data-stu-id="249ca-146">Try out the `ADD` custom function by completing the following steps in Excel:</span></span>

1. <span data-ttu-id="249ca-147">В ячейке введите **= CONTOSO**.</span><span class="sxs-lookup"><span data-stu-id="249ca-147">Within a cell, type **=CONTOSO**.</span></span> <span data-ttu-id="249ca-148">Обратите внимание на то, что в меню автозаполнения отображается список всех функций в пространстве имен `CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="249ca-148">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="249ca-149">Выполните функцию `CONTOSO.ADD` с числами `10` и `200` в качестве входных параметров, указав следующее значение в ячейке и нажав клавишу ВВОД:</span><span class="sxs-lookup"><span data-stu-id="249ca-149">Run the `CONTOSO.ADD` function, with numbers `10` and `200` as input parameters, by specifying the following value in the cell and pressing enter:</span></span>

    ```
    =CONTOSO.ADD(10,200)
    ```

<span data-ttu-id="249ca-150">`ADD` Настраиваемая функция вычисляет сумму двух чисел, которые указаны в качестве входных параметров.</span><span class="sxs-lookup"><span data-stu-id="249ca-150">The `ADD` custom function computes the sum of the two numbers that you specify as input parameters.</span></span> <span data-ttu-id="249ca-151">После набора `=CONTOSO.ADD(10,200)` в ячейке должен отобразиться результат **210** после нажатия клавиши ВВОД.</span><span class="sxs-lookup"><span data-stu-id="249ca-151">Typing `=CONTOSO.ADD(10,200)` should produce the result **210** in the cell after you press enter.</span></span>

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a><span data-ttu-id="249ca-152">Создание настраиваемой функции, которая будет запрашивать данные с веб-сайта</span><span class="sxs-lookup"><span data-stu-id="249ca-152">Create a custom function that requests data from the web</span></span>

<span data-ttu-id="249ca-153">Что делать, если требуется функция, которая может запросить цену акции из интерфейса API и отобразить результат в ячейке таблицы?</span><span class="sxs-lookup"><span data-stu-id="249ca-153">What if you needed a function that could request the price of a stock from an API and display the result in the cell of a worksheet?</span></span> <span data-ttu-id="249ca-154">Настраиваемые функции построены таким образом, чтобы вы могли легко запрашивать данные из Интернета асинхронным образом.</span><span class="sxs-lookup"><span data-stu-id="249ca-154">Custom functions are designed so that you can easily request data from the web asynchronously.</span></span>

<span data-ttu-id="249ca-155">Выполните следующие действия, чтобы создать настраиваемую функцию с именем `stockPrice` , которая будет принимать биржевой символ (например, **MSFT**) и возвращать цену акции.</span><span class="sxs-lookup"><span data-stu-id="249ca-155">Complete the following steps to create a custom function named `stockPrice` that accepts a stock ticker (e.g., **MSFT**) and returns the price of that stock.</span></span> <span data-ttu-id="249ca-156">Настраиваемая функция использует API для трейдинга IEX, который является бесплатным и не требует проверки подлинности.</span><span class="sxs-lookup"><span data-stu-id="249ca-156">This custom function uses the IEX Trading API, which is free and does not require authentication.</span></span>

1. <span data-ttu-id="249ca-157">В проекте **stock-ticker** , созданном генератором Yo Office, найдите файл **src/customfunctions.js** и откройте его в редакторе кода.</span><span class="sxs-lookup"><span data-stu-id="249ca-157">In the **stock-ticker** project that the Yo Office generator created, find the file **src/customfunctions.js** and open it in your code editor.</span></span>

2. <span data-ttu-id="249ca-158">Добавьте приведенный ниже код в файл **customfunctions.js** и сохраните его.</span><span class="sxs-lookup"><span data-stu-id="249ca-158">Add the following code to **home.js** and save the file.</span></span>

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

3. <span data-ttu-id="249ca-159">Прежде чем эта новая функция станет доступной в Excel  для конечных пользователей, необходимо указать метаданные, которые описывают эту функцию.</span><span class="sxs-lookup"><span data-stu-id="249ca-159">Before Excel can make this new function available to end-users, you must specify metadata that describes this function.</span></span> <span data-ttu-id="249ca-160">В проекте **stock-ticker**, созданном генератором Yo Office, найдите файл **config/customfunctions.json** и откройте его в редакторе кода.</span><span class="sxs-lookup"><span data-stu-id="249ca-160">In the **stock-ticker** project that the Yo Office generator created, find the file **config/customfunctions.json** and open it in your code editor.</span></span> <span data-ttu-id="249ca-161">Добавьте следующий объект в массив `functions` в файле **config/customfunctions.json** и сохраните его.</span><span class="sxs-lookup"><span data-stu-id="249ca-161">Add the following object to the `functions` array within the **config/customfunctions.json** file and save the file.</span></span>

    <span data-ttu-id="249ca-162">Этот JSON описывает функцию `stockPrice`.</span><span class="sxs-lookup"><span data-stu-id="249ca-162">This JSON describes the `stockPrice` function.</span></span>

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

4. <span data-ttu-id="249ca-163">Необходимо перерегистрировать надстройку в Excel, чтобы новая функция стала доступной для конечных пользователей.</span><span class="sxs-lookup"><span data-stu-id="249ca-163">You must reregister the add-in in Excel in order for the new function to be available to end-users.</span></span> <span data-ttu-id="249ca-164">Выполните следующие действия для той платформы, которая используется в данном уроке.</span><span class="sxs-lookup"><span data-stu-id="249ca-164">Complete the following steps for the platform that you're using in this tutorial.</span></span>

    * <span data-ttu-id="249ca-165">Если вы используете Excel для Windows:</span><span class="sxs-lookup"><span data-stu-id="249ca-165">If you're using Excel for Windows:</span></span>

        1. <span data-ttu-id="249ca-166">Закройте Excel и снова его откройте.</span><span class="sxs-lookup"><span data-stu-id="249ca-166">Close Excel and then reopen Excel.</span></span>

        2. <span data-ttu-id="249ca-167">В Excel перейдите на вкладку **Вставка** и нажмите стрелку вниз, расположенную справа от раздела **Мои надстройки**.  ![Вставьте ленту в Excel для Windows, используя выделенную стрелку "Мои надстройки".](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="249ca-167">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

        1. <span data-ttu-id="249ca-168">В списке доступных надстроек найдите раздел **Надстройки разработчика** и выберите надстройку **Настраиваемые функции Excel** для ее регистрации.</span><span class="sxs-lookup"><span data-stu-id="249ca-168">In the list of available add-ins, find the **Developer Add-ins** section and select the **Excel Custom Functions** add-in to register it.</span></span>
            <span data-ttu-id="249ca-169">![Вставьте ленту в Excel для Windows с помощью надстройки настраиваемых функций Excel, выделенной в списке "Мои надстройки".](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="249ca-169">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

    * <span data-ttu-id="249ca-170">Если вы используете Excel Online:</span><span class="sxs-lookup"><span data-stu-id="249ca-170">If you're using Excel Online:</span></span> 

        1. <span data-ttu-id="249ca-171">В Excel Online перейдите на вкладку **Вставка** и выберите **Надстройки**.  ![Вставьте ленту в Excel Online, используя выделенный значок "Мои надстройки".](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="249ca-171">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

        2. <span data-ttu-id="249ca-172">Выберите **Управление моими надстройками** и **Отправить мою надстройку**.</span><span class="sxs-lookup"><span data-stu-id="249ca-172">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

        3. <span data-ttu-id="249ca-173">Нажмите кнопку **Обзор...** и перейдите в корневой каталог проекта, созданный генератором Yo Office.</span><span class="sxs-lookup"><span data-stu-id="249ca-173">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span> 

        4. <span data-ttu-id="249ca-174">Выберите файл **manifest.xml** и нажмите кнопки **Открыть**, а затем **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="249ca-174">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

5. <span data-ttu-id="249ca-175">Теперь давайте испытаем новую функцию.</span><span class="sxs-lookup"><span data-stu-id="249ca-175">Now, let's try out the new function.</span></span> <span data-ttu-id="249ca-176">В ячейке **B1** введите текст `=CONTOSO.STOCKPRICE("MSFT")` и нажмите клавишу ВВОД.</span><span class="sxs-lookup"><span data-stu-id="249ca-176">In cell **B1**, type the text `=CONTOSO.STOCKPRICE("MSFT")` and press enter.</span></span> <span data-ttu-id="249ca-177">Вы должны увидеть, что результат в ячейке **B1** является текущей биржевой ценой одной акции корпорации Майкрософт.</span><span class="sxs-lookup"><span data-stu-id="249ca-177">You should see that the result in cell **B1** is the current stock price for one share of Microsoft stock.</span></span>

## <a name="create-a-streaming-asynchronous-custom-function"></a><span data-ttu-id="249ca-178">Создание асинхронной настраиваемой функции для потоковой передачи</span><span class="sxs-lookup"><span data-stu-id="249ca-178">Create a streaming asynchronous custom function</span></span>

<span data-ttu-id="249ca-179">Только что созданная функция  `stockPrice` возвращает цену акции в определенный момент времени, но биржевые котировки постоянно меняются.</span><span class="sxs-lookup"><span data-stu-id="249ca-179">The `stockPrice` function that you just created returns the price of a stock at a specific moment in time, but stock prices are always changing.</span></span> <span data-ttu-id="249ca-180">Давайте создадим настраиваемую функцию, которая в потоковом режиме будет передавать данные из интерфейса API для получения обновлений цен акций в реальном времени.</span><span class="sxs-lookup"><span data-stu-id="249ca-180">Let's create a custom function that streams data from an API to get real-time updates on a stock price.</span></span>

<span data-ttu-id="249ca-181">Выполните следующие действия, чтобы создать настраиваемую функцию с именем `stockPriceStream` , которая будет запрашивать цену указанной акции каждые 1000 миллисекунд (при условии, что предыдущий запрос был выполнен).</span><span class="sxs-lookup"><span data-stu-id="249ca-181">Complete the following steps to create a custom function named `stockPriceStream` that requests the price of the specified stock every 1000 milliseconds (provided that the previous request has completed).</span></span> <span data-ttu-id="249ca-182">В ходе выполнения начального запроса вы можете видеть значение заполнителя **#GETTING_DATA** в ячейке, в которой вызывается функция.</span><span class="sxs-lookup"><span data-stu-id="249ca-182">While the initial request is in-progress, you may see the placeholder value **#GETTING_DATA** the cell where the function is being called.</span></span> <span data-ttu-id="249ca-183">Когда значение возвращается функцией, **#GETTING_DATA** заменяется этим значением в ячейке.</span><span class="sxs-lookup"><span data-stu-id="249ca-183">When a value is returned by the function, **#GETTING_DATA** will be replaced by that value in the cell.</span></span>

1. <span data-ttu-id="249ca-184">В проекте **stock-ticker**, созданном генератором Yo Office, добавьте следующий код в **src/customfunctions.js** и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="249ca-184">In the **stock-ticker** project that the Yo Office generator created, add the following code to **src/customfunctions.js** and save the file.</span></span>

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

2. <span data-ttu-id="249ca-185">Прежде чем эта новая функция станет доступной в Excel  для конечных пользователей, необходимо указать метаданные, которые описывают эту функцию.</span><span class="sxs-lookup"><span data-stu-id="249ca-185">Before Excel can make this new function available to end-users, you must specify metadata that describes this function.</span></span> <span data-ttu-id="249ca-186">В проект **stock-ticker**, созданный генератором Yo Office, добавьте следующий объект в массив `functions` в файле **config/customfunctions.json** и сохраните его.</span><span class="sxs-lookup"><span data-stu-id="249ca-186">In the **stock-ticker** project that the Yo Office generator created, add the following object to the `functions` array within the **config/customfunctions.json** file and save the file.</span></span>

    <span data-ttu-id="249ca-187">Этот JSON описывает функцию `stockPriceStream`.</span><span class="sxs-lookup"><span data-stu-id="249ca-187">This JSON describes the `stockPriceStream` function.</span></span> <span data-ttu-id="249ca-188">Для любой функции потоковой передачи свойства `stream` и `cancelable` должны иметь значение `true` в объекте `options` , как показано в этом примере кода.</span><span class="sxs-lookup"><span data-stu-id="249ca-188">For any streaming function, the `stream` property and the `cancelable` property must be set to `true` within the `options` object, as shown in this code sample.</span></span>

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

3. <span data-ttu-id="249ca-189">Необходимо перерегистрировать надстройку в Excel, чтобы новая функция стала доступной для конечных пользователей.</span><span class="sxs-lookup"><span data-stu-id="249ca-189">You must reregister the add-in in Excel in order for the new function to be available to end-users.</span></span> <span data-ttu-id="249ca-190">Выполните следующие действия для той платформы, которая используется в данном уроке.</span><span class="sxs-lookup"><span data-stu-id="249ca-190">Complete the following steps for the platform that you're using in this tutorial.</span></span>

    * <span data-ttu-id="249ca-191">Если вы используете Excel для Windows:</span><span class="sxs-lookup"><span data-stu-id="249ca-191">If you're using Excel for Windows:</span></span>

        1. <span data-ttu-id="249ca-192">Закройте Excel и снова его откройте.</span><span class="sxs-lookup"><span data-stu-id="249ca-192">Close Excel and then reopen Excel.</span></span>
        
        2. <span data-ttu-id="249ca-193">В Excel перейдите на вкладку **Вставка** и нажмите стрелку вниз, расположенную справа от раздела **Мои надстройки**.  ![Вставьте ленту в Excel для Windows, используя выделенную стрелку "Мои надстройки".](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="249ca-193">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

        3. <span data-ttu-id="249ca-194">В списке доступных надстроек найдите раздел **Надстройки разработчика** и выберите надстройку **Настраиваемые функции Excel** для ее регистрации.</span><span class="sxs-lookup"><span data-stu-id="249ca-194">In the list of available add-ins, find the **Developer Add-ins** section and select the **Excel Custom Functions** add-in to register it.</span></span>
            <span data-ttu-id="249ca-195">![Вставьте ленту в Excel для Windows с помощью надстройки настраиваемых функций Excel, выделенной в списке "Мои надстройки".](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="249ca-195">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

    * <span data-ttu-id="249ca-196">Если вы используете Excel Online:</span><span class="sxs-lookup"><span data-stu-id="249ca-196">If you're using Excel Online:</span></span> 

        1. <span data-ttu-id="249ca-197">В Excel Online перейдите на вкладку **Вставка** и выберите **Надстройки**.  ![Вставьте ленту в Excel Online, используя выделенный значок "Мои надстройки".](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="249ca-197">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

        2. <span data-ttu-id="249ca-198">Выберите **Управление моими надстройками** и **Отправить мою надстройку**.</span><span class="sxs-lookup"><span data-stu-id="249ca-198">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

        3. <span data-ttu-id="249ca-199">Нажмите кнопку **Обзор...** и перейдите в корневой каталог проекта, созданный генератором Yo Office.</span><span class="sxs-lookup"><span data-stu-id="249ca-199">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span> 

        4. <span data-ttu-id="249ca-200">Выберите файл **manifest.xml** и нажмите кнопки **Открыть**, а затем **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="249ca-200">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

4. <span data-ttu-id="249ca-201">Теперь давайте испытаем новую функцию.</span><span class="sxs-lookup"><span data-stu-id="249ca-201">Now, let's try out the new function.</span></span> <span data-ttu-id="249ca-202">В ячейке **C1** введите текст `=CONTOSO.STOCKPRICESTREAM("MSFT")` и нажмите клавишу ВВОД.</span><span class="sxs-lookup"><span data-stu-id="249ca-202">In cell **C1**, type the text `=CONTOSO.STOCKPRICESTREAM("MSFT")` and press enter.</span></span> <span data-ttu-id="249ca-203">При условии что фондовая биржа открыта, вы должны видеть, что результат в ячейке **C1** постоянно обновляется в режиме реального времени и показывает цену одной акции корпорации Майкрософт.</span><span class="sxs-lookup"><span data-stu-id="249ca-203">Provided that the stock market is open, you should see that the result in cell **C1** is constantly updated to reflect the real-time price for one share of Microsoft stock.</span></span>

## <a name="next-steps"></a><span data-ttu-id="249ca-204">Дальнейшие шаги</span><span class="sxs-lookup"><span data-stu-id="249ca-204">Next steps</span></span>

<span data-ttu-id="249ca-205">В этом уроке вы создали новый проект настраиваемых функций, испытали готовую функцию, создали настраиваемую функцию, которая запрашивает данные с веб-сайта, а также создали настраиваемую функцию для  потоковой передачи данных в режиме реального времени из Интернета.</span><span class="sxs-lookup"><span data-stu-id="249ca-205">In this tutorial, you've created a new custom functions project, tried out a prebuilt function, created a custom function that requests data from the web, and created a custom function that streams real-time data from the web.</span></span> <span data-ttu-id="249ca-206">Для получения дополнительных сведений о настраиваемых функциях в Excel см. следующую статью:</span><span class="sxs-lookup"><span data-stu-id="249ca-206">To learn more about custom functions in Excel, continue to the following article:</span></span> 

> [!div class="nextstepaction"]
> [<span data-ttu-id="249ca-207">Создание настраиваемых функций в Excel</span><span class="sxs-lookup"><span data-stu-id="249ca-207">Create custom functions in Excel (Preview)</span></span>](../excel/custom-functions-overview.md)

## <a name="legal-information"></a><span data-ttu-id="249ca-208">Юридическая информация</span><span class="sxs-lookup"><span data-stu-id="249ca-208">Legal Information</span></span>

<span data-ttu-id="249ca-209">Данные бесплатно предоставлены компанией [IEX](https://iextrading.com/developer/).</span><span class="sxs-lookup"><span data-stu-id="249ca-209">Data provided free by [IEX](https://iextrading.com/developer/).</span></span> <span data-ttu-id="249ca-210">См. [Условия использования IEX](https://iextrading.com/api-exhibit-a/).</span><span class="sxs-lookup"><span data-stu-id="249ca-210">View [IEX's Terms of Use](https://iextrading.com/api-exhibit-a/).</span></span> <span data-ttu-id="249ca-211">Интерфейс API IEX используется в этом уроке корпорацией Майкрософт только для обучения.</span><span class="sxs-lookup"><span data-stu-id="249ca-211">Microsoft's use of the IEX API in this tutorial is for educational purposes only.</span></span>
