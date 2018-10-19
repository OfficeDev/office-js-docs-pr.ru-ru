# <a name="tutorial-create-custom-functions-in-excel"></a><span data-ttu-id="90123-101">Урок: создание настраиваемых функций в Excel</span><span class="sxs-lookup"><span data-stu-id="90123-101">Create custom functions in Excel (Preview)</span></span>

## <a name="introduction"></a><span data-ttu-id="90123-102">Введение</span><span class="sxs-lookup"><span data-stu-id="90123-102">Introduction</span></span>

<span data-ttu-id="90123-p101">Настраиваемые функции позволяют добавлять новые функции в Excel, определяя эти функции в JavaScript как часть надстройки.  Пользователи в Excel могут получать доступ к настраиваемым функциям так же, как к любой собственной функции в Excel, например  `SUM()`. Можно создавать настраиваемые функции, которые будут выполнять простые задачи, такие как настраиваемые вычисления, или более сложные задачи, например потоковая передача данных в режиме реального времени из Интернета в лист таблицы.</span><span class="sxs-lookup"><span data-stu-id="90123-p101">Custom functions enable you to add new functions to Excel by defining those functions in JavaScript as part of an add-in. Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`. You can create custom functions that perform simple tasks such as custom calculations or more complex tasks such as streaming real-time data from the web into a worksheet.</span></span>

<span data-ttu-id="90123-106">В этом руководстве вы:</span><span class="sxs-lookup"><span data-stu-id="90123-106">In this tutorial, you will use Visual Studio Code.</span></span>
> [!div class="checklist"]
> * <span data-ttu-id="90123-107">Создадите проект с настраиваемыми функциями, используя генератор Yo Office.</span><span class="sxs-lookup"><span data-stu-id="90123-107">Create a custom functions project by using the Yo Office generator</span></span>
> * <span data-ttu-id="90123-108">Используете готовую настраиваемую функцию для выполнения простых вычислений.</span><span class="sxs-lookup"><span data-stu-id="90123-108">Use a prebuilt custom function to perform a simple calculation</span></span>
> * <span data-ttu-id="90123-109">Создадите настраиваемую функцию, которая будет запрашивать данные с веб-сайта.</span><span class="sxs-lookup"><span data-stu-id="90123-109">Create a custom function that requests data from the web</span></span>
> * <span data-ttu-id="90123-110">Создадите настраиваемую функцию, которая будет передавать данные в реальном времени с веб-сайта.</span><span class="sxs-lookup"><span data-stu-id="90123-110">Create a custom function that streams real-time data from the web</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="prerequisites"></a><span data-ttu-id="90123-111">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="90123-111">Prerequisites</span></span>

* [<span data-ttu-id="90123-112">Node.js и npm.</span><span class="sxs-lookup"><span data-stu-id="90123-112">Node and npm</span></span>](https://nodejs.org/en/)

* <span data-ttu-id="90123-113">[Git Bash](https://git-scm.com/downloads) (или другой клиент Git)</span><span class="sxs-lookup"><span data-stu-id="90123-113">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

* <span data-ttu-id="90123-p102">Последняя версия [Yeoman](http://yeoman.io/) и [Генератор Yo Office](https://www.npmjs.com/package/generator-office). Чтобы установить эти средства глобально, выполните следующую команду из командной строки:</span><span class="sxs-lookup"><span data-stu-id="90123-p102">The latest version of [Yeoman](http://yeoman.io/) and the [Yo Office generator](https://www.npmjs.com/package/generator-office). To install these tools globally, run the following command via the command prompt:</span></span>

    ```bash
    npm install -g yo generator-office
    ```

* <span data-ttu-id="90123-116">Excel для Windows (сборка 10827 или более поздняя версия) или Excel Online.</span><span class="sxs-lookup"><span data-stu-id="90123-116">Excel for Windows (build number 10827 or later) or Excel Online</span></span>

* <span data-ttu-id="90123-117">Присоединяйтесь к [программе предварительной оценки Office](https://products.office.com/office-insider) (уровень**Insider** ранее именовался «Insider Fast»)</span><span class="sxs-lookup"><span data-stu-id="90123-117">Join the [Office Insider program](https://products.office.com/office-insider) (**Insider** level -- formerly called "Insider Fast")</span></span>

## <a name="create-a-custom-functions-project"></a><span data-ttu-id="90123-118">Создание проекта настраиваемых функций</span><span class="sxs-lookup"><span data-stu-id="90123-118">Create a custom functions project by using the Yo Office generator</span></span>

<span data-ttu-id="90123-119">Вы начнете этот урок с использования генератора Yo Office для создания файлов, необходимых для проекта настраиваемых функций.</span><span class="sxs-lookup"><span data-stu-id="90123-119">You’ll begin this tutorial by using the Yo Office generator to create the files that you need for your custom functions project.</span></span>

1. <span data-ttu-id="90123-120">Выполните приведенную ниже команду и ответьте на запросы, как показано ниже.</span><span class="sxs-lookup"><span data-stu-id="90123-120">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    * <span data-ttu-id="90123-121">Выберите тип проекта: `Excel Custom Functions Add-in project (...)`</span><span class="sxs-lookup"><span data-stu-id="90123-121">Choose a project type:`Excel Custom Functions Add-in project (...)`</span></span>
    * <span data-ttu-id="90123-122">Выберите тип сценария: `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="90123-122">Choose a script type:`JavaScript`</span></span>
    * <span data-ttu-id="90123-123">Как вы хотите назвать надстройку?</span><span class="sxs-lookup"><span data-stu-id="90123-123">What do you want to name your add-in?:</span></span> `stock-ticker`

    ![Yo Office выводит запросы, касающиеся настраиваемых функций.](../images/yo-office-cfs-stock-ticker-3.png)

    <span data-ttu-id="90123-p103">После завершения работы мастера генератор создает файлы проекта и устанавливает вспомогательные компоненты Node. Файлы проекта поступают из репозитория GitHub [Настраиваемые функции Excel](https://github.com/OfficeDev/Excel-Custom-Functions).</span><span class="sxs-lookup"><span data-stu-id="90123-p103">After you complete the wizard, the generator will create the project files and install supporting Node components. The project files come from the [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) GitHub repository.</span></span>

2. <span data-ttu-id="90123-127">Перейдите в папку проекта.</span><span class="sxs-lookup"><span data-stu-id="90123-127">Navigate to the project folder:</span></span>

    ```bash
    cd stock-ticker
    ```

3. <span data-ttu-id="90123-128">Запустите локальный веб-сервер.</span><span class="sxs-lookup"><span data-stu-id="90123-128">Start the local web server.</span></span>

    * <span data-ttu-id="90123-129">При использовании Excel для Windows для тестирования настраиваемых функций выполните следующую команду для запуска локального веб-сервера, запустите Excel и загрузите надстройку:</span><span class="sxs-lookup"><span data-stu-id="90123-129">If you'll be using Excel for Windows to test your custom functions, run the following command to start the local web server, launch Excel, and sideload the add-in:</span></span>

        ```bash
        npm start
        ```

    * <span data-ttu-id="90123-130">При использовании Excel Online для тестирования настраиваемых функций выполните следующую команду для запуска локального веб-сервера:</span><span class="sxs-lookup"><span data-stu-id="90123-130">If you'll be using Excel Online to test your custom functions, run the following command to start the local web server:</span></span> 

        ```bash
        npm run start-web
        ```

## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="90123-131">Тестирование стандартной настраиваемой функции</span><span class="sxs-lookup"><span data-stu-id="90123-131">Try out a prebuilt custom function</span></span>

<span data-ttu-id="90123-p104">Проект настраиваемых функций, созданный с помощью генератора Yo Office, содержит несколько стандартных настраиваемых функций, которые определены в файле **src/customfunction.js**. Файл **manifest.xml** в корневом каталоге проекта указывает на то, что все настраиваемые функции принадлежат к пространству имен `CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="90123-p104">The custom functions project that you created by using the Yo Office generator contains some prebuilt custom functions, defined within the **src/customfunction.js** file. The **manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.</span></span>

<span data-ttu-id="90123-p105">Перед началом использования любых стандартных настраиваемых функций надстройки настраиваемых функций необходимо зарегистрировать в Excel. Для этого следует выполнить действия, указанные в настоящем руководстве для используемой вами платформы.</span><span class="sxs-lookup"><span data-stu-id="90123-p105">Before you can use any of the prebuilt custom functions, you must register the custom functions add-in in Excel. Do so by completing steps for the platform that you'll be using in this tutorial.</span></span>

* <span data-ttu-id="90123-136">Если для тестирования настраиваемых функций используется Excel для Windows:</span><span class="sxs-lookup"><span data-stu-id="90123-136">If you'll be using Excel for Windows to test your custom functions:</span></span>

    1. <span data-ttu-id="90123-137">В Excel перейдите на вкладку **Вставка** и нажмите на стрелку вниз, расположенную справа от секции **Мои надстройки**.  ![Вставьте ленту в Excel для Windows с помощью выделенной стрелки «Мои надстройки».](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="90123-137">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

    2. <span data-ttu-id="90123-p106">В списке доступных надстроек найдите секцию \*\* Надстройки разработчика\*\*  и выберите надстройку **Настраиваемые функции Excel**, чтобы зарегистрировать ее. ![Вставьте ленту в Excel для Windows с помощью выделенной в списке «Мои надстройки» надстройки «Настраиваемые функции Excel»](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="90123-p106">In the list of available add-ins, find the **Developer Add-ins** section and select the **Excel Custom Functions** add-in to register it.  ![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

* <span data-ttu-id="90123-140">Если для тестирования настраиваемых функций используется Excel Online:</span><span class="sxs-lookup"><span data-stu-id="90123-140">If you'll be using Excel Online to test your custom functions:</span></span> 

    1. <span data-ttu-id="90123-141">В Excel Online перейдите на вкладку **Вставка** и выберите **Надстройки**.  ![Вставьте ленту в Excel Online, используя выделенный значок "Мои надстройки".](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="90123-141">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

    2. <span data-ttu-id="90123-142">Выберите **Управление моими надстройками** и **Отправить мою надстройку**.</span><span class="sxs-lookup"><span data-stu-id="90123-142">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

    3. <span data-ttu-id="90123-143">Нажмите кнопку **Обзор...** и перейдите в корневой каталог проекта, созданный генератором Yo Office.</span><span class="sxs-lookup"><span data-stu-id="90123-143">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span> 

    4. <span data-ttu-id="90123-144">Выберите файл **manifest.xml** и нажмите на кнопку **Открыть**, а затем — на кнопку **Загрузить**.</span><span class="sxs-lookup"><span data-stu-id="90123-144">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

<span data-ttu-id="90123-p107">На этом этапе стандартные настраиваемые функции, имеющиеся в вашем проекте, оказываются загруженными и доступными для использования в Excel. Проверьте `ADD` настраиваемую функцию, выполнив в Excel следующие действия:</span><span class="sxs-lookup"><span data-stu-id="90123-p107">At this point, the prebuilt custom functions in your project are loaded and available within Excel. Try out the `ADD` custom function by completing the following steps in Excel:</span></span>

1. <span data-ttu-id="90123-p108">В ячейке введите **= CONTOSO**. Обратите внимание на то, что в меню автозаполнения отображается список всех функций, принадлежащих пространству имен `CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="90123-p108">Within a cell, type **=CONTOSO**. Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="90123-149">Выполните функцию `CONTOSO.ADD` с числами `10` и `200` в качестве входных параметров, указав в ячейке следующее значение и нажав на клавишу ВВОД:</span><span class="sxs-lookup"><span data-stu-id="90123-149">Run the `CONTOSO.ADD` function, with numbers `10` and `200` as input parameters, by specifying the following value in the cell and pressing enter:</span></span>

    ```
    =CONTOSO.ADD(10,200)
    ```

<span data-ttu-id="90123-p109">Настраиваемая функция `ADD` вычисляет сумму двух чисел, указанных в качестве входных параметров. При вводе `=CONTOSO.ADD(10,200)` результирующим значением в ячейке после нажатия на клавишу ВВОД должно стать **210**.</span><span class="sxs-lookup"><span data-stu-id="90123-p109">The `ADD` custom function computes the sum of the two numbers that you specify as input parameters. Typing `=CONTOSO.ADD(10,200)` should produce the result **210** in the cell after you press enter.</span></span>

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a><span data-ttu-id="90123-152">Создание настраиваемой функции, запрашивающей данные с веб-сервера</span><span class="sxs-lookup"><span data-stu-id="90123-152">Create a custom function that requests data from the web</span></span>

<span data-ttu-id="90123-p110">Что делать, если требуется создать функцию, которая может запросить у API цену акции и отобразить результат в ячейке листа? Настраиваемые функции позволяют пользователю без труда запрашивать данные с веб-сервера при работе в асинхронном режиме.</span><span class="sxs-lookup"><span data-stu-id="90123-p110">What if you needed a function that could request the price of a stock from an API and display the result in the cell of a worksheet? Custom functions are designed so that you can easily request data from the web asynchronously.</span></span>

<span data-ttu-id="90123-p111">Чтобы создать настраиваемую функцию с именем `stockPrice`, которая принимает код акции (например, **MSFT**) и возвращает цену этой акции, выполните следующие действия. Эта настраиваемая функция использует API для трейдинга IEX, который является бесплатным и не требует проверки подлинности.</span><span class="sxs-lookup"><span data-stu-id="90123-p111">Complete the following steps to create a custom function named `stockPrice` that accepts a stock ticker (e.g., **MSFT**) and returns the price of that stock. This custom function uses the IEX Trading API, which is free and does not require authentication.</span></span>

1. <span data-ttu-id="90123-157">В созданном генератором Yo Office проекте **stock-ticker** найдите файл **src/customfunctions.js** и откройте его в редакторе кода.</span><span class="sxs-lookup"><span data-stu-id="90123-157">In the **stock-ticker** project that the Yo Office generator created, find the file **src/customfunctions.js** and open it in your code editor.</span></span>

2. <span data-ttu-id="90123-158">Добавьте приведенный ниже код в файл **customfunctions.js** и выполните сохранение файла.</span><span class="sxs-lookup"><span data-stu-id="90123-158">Add the following code to **home.js** and save the file.</span></span>

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

3. <span data-ttu-id="90123-p112">Перед тем, как Excel сможет сделать эту функцию доступной для конечных пользователей, необходимо указать описывающие ее метаданные. В созданном генератором Yo Office проекте **stock-ticker** найдите файл **config/customfunctions.json** и откройте его в редакторе кода. Добавьте следующий объект в массив `functions`, имеющийся в файле **config/customfunctions.json**, и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="90123-p112">Before Excel can make this new function available to end-users, you must specify metadata that describes this function. In the **stock-ticker** project that the Yo Office generator created, find the file **config/customfunctions.json** and open it in your code editor. Add the following object to the `functions` array within the **config/customfunctions.json** file and save the file.</span></span>

    <span data-ttu-id="90123-162">Этот файл JSON описывает функцию `stockPrice`.</span><span class="sxs-lookup"><span data-stu-id="90123-162">This JSON describes the `stockPrice` function.</span></span>

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

4. <span data-ttu-id="90123-p113">Чтобы новая функция стала доступна для конечных пользователей, надстройку необходимо повторно зарегистрировать в Excel. Выполните дальнейшие действия, указанные в настоящем руководстве для используемой вами платформы.</span><span class="sxs-lookup"><span data-stu-id="90123-p113">You must reregister the add-in in Excel in order for the new function to be available to end-users. Complete the following steps for the platform that you're using in this tutorial.</span></span>

    * <span data-ttu-id="90123-165">В случае использования Excel для Windows:</span><span class="sxs-lookup"><span data-stu-id="90123-165">If you're using Excel for Windows:</span></span>

        1. <span data-ttu-id="90123-166">Закройте Excel и снова его откройте.</span><span class="sxs-lookup"><span data-stu-id="90123-166">Close Excel and then reopen Excel.</span></span>

        2. <span data-ttu-id="90123-167">В Excel перейдите на вкладку **Вставка** и нажмите на стрелку вниз, расположенную справа от секции **Мои надстройки**.  ![Вставьте ленту в Excel для Windows с помощью выделенной стрелки «Мои надстройки».](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="90123-167">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

        1. <span data-ttu-id="90123-p114">В списке доступных надстроек найдите секцию \*\* Надстройки разработчика\*\*  и выберите надстройку **Настраиваемые функции Excel**, чтобы зарегистрировать ее. ![Вставьте ленту в Excel для Windows с помощью выделенной в списке «Мои надстройки» надстройки «Настраиваемые функции Excel»](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="90123-p114">In the list of available add-ins, find the **Developer Add-ins** section and select the **Excel Custom Functions** add-in to register it.  ![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

    * <span data-ttu-id="90123-170">В случае использования Excel Online:</span><span class="sxs-lookup"><span data-stu-id="90123-170">If you're using Excel Online:</span></span> 

        1. <span data-ttu-id="90123-171">В Excel Online перейдите на вкладку **Вставка** и выберите **Надстройки**.  ![Вставьте ленту в Excel Online, используя выделенный значок "Мои надстройки".](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="90123-171">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

        2. <span data-ttu-id="90123-172">Выберите **Управление моими надстройками** и **Отправить мою надстройку**.</span><span class="sxs-lookup"><span data-stu-id="90123-172">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

        3. <span data-ttu-id="90123-173">Нажмите кнопку **Обзор...** и перейдите в корневой каталог проекта, созданный генератором Yo Office.</span><span class="sxs-lookup"><span data-stu-id="90123-173">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span> 

        4. <span data-ttu-id="90123-174">Выберите файл **manifest.xml** и нажмите на кнопку **Открыть**, а затем — на кнопку **Загрузить**.</span><span class="sxs-lookup"><span data-stu-id="90123-174">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

5. <span data-ttu-id="90123-p115">Теперь можно опробовать новую функцию. В ячейке **B1** введите текст `=CONTOSO.STOCKPRICE("MSFT")` и нажмите на клавишу ВВОД. В ячейке **B1** должен отобразиться результат, представляющий собой текущую стоимость одной акции корпорации Майкрософт.</span><span class="sxs-lookup"><span data-stu-id="90123-p115">Now, let's try out the new function. In cell **B1**, type the text `=CONTOSO.STOCKPRICE("MSFT")` and press enter. You should see that the result in cell **B1** is the current stock price for one share of Microsoft stock.</span></span>

## <a name="create-a-streaming-asynchronous-custom-function"></a><span data-ttu-id="90123-178">Создание асинхронной настраиваемой функции для потоковой передачи</span><span class="sxs-lookup"><span data-stu-id="90123-178">Create a streaming asynchronous custom function</span></span>

<span data-ttu-id="90123-p116">Только что созданная функция `stockPrice` возвращает цену акции в определенный момент времени, но стоимость акций постоянно меняется. Попробуем создать настраиваемую функцию для поточной передачи данных от API, чтобы обновлять сведения о цене акции в реальном времени.</span><span class="sxs-lookup"><span data-stu-id="90123-p116">The `stockPrice` function that you just created returns the price of a stock at a specific moment in time, but stock prices are always changing. Let's create a custom function that streams data from an API to get real-time updates on a stock price.</span></span>

<span data-ttu-id="90123-p117">Чтобы создать настраиваемую функцию с именем `stockPriceStream`, которая запрашивает цену указанной акции через каждые 1000 миллисекунд (при условии завершения предыдущего запроса), выполните следующие действия. Во время выполнения начального запроса в ячейке, из которой вызывается функция, можно видеть значение заполнителя **#GETTING_DATA**. При возврате функцией значения оно будет отображаться в ячейке вместо **#GETTING_DATA**.</span><span class="sxs-lookup"><span data-stu-id="90123-p117">Complete the following steps to create a custom function named `stockPriceStream` that requests the price of the specified stock every 1000 milliseconds (provided that the previous request has completed). While the initial request is in-progress, you may see the placeholder value **#GETTING_DATA** the cell where the function is being called. When a value is returned by the function, **#GETTING_DATA** will be replaced by that value in the cell.</span></span>

1. <span data-ttu-id="90123-184">В созданном генератором Yo Office проекте **stock-ticker** добавьте следующий код в файл **src/customfunctions.js** и сохраните его.</span><span class="sxs-lookup"><span data-stu-id="90123-184">In the **stock-ticker** project that the Yo Office generator created, add the following code to **src/customfunctions.js** and save the file.</span></span>

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

2. <span data-ttu-id="90123-p118">Перед тем, как Excel сможет сделать эту функцию доступной для конечных пользователей, необходимо указать описывающие ее метаданные. В созданном генератором Yo Office проекте **stock-ticker** добавьте следующий объект в массив `functions`, имеющийся в файле **config/customfunctions.json**, после чего произведите сохранение файла.</span><span class="sxs-lookup"><span data-stu-id="90123-p118">Before Excel can make this new function available to end-users, you must specify metadata that describes this function. In the **stock-ticker** project that the Yo Office generator created, add the following object to the `functions` array within the **config/customfunctions.json** file and save the file.</span></span>

    <span data-ttu-id="90123-p119">Этот файл JSON описывает функцию `stockPriceStream`. Для любой функции потоковой передачи данных свойству `stream` и свойству `cancelable` следует присвоить значение `true` внутри объекта `options`, как показано в приведенном далее примере кода.</span><span class="sxs-lookup"><span data-stu-id="90123-p119">This JSON describes the `stockPriceStream` function. For any streaming function, the `stream` property and the `cancelable` property must be set to `true` within the `options` object, as shown in this code sample.</span></span>

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

3. <span data-ttu-id="90123-p120">Чтобы новая функция стала доступна для конечных пользователей, надстройку необходимо повторно зарегистрировать в Excel. Выполните дальнейшие действия, указанные в настоящем руководстве для используемой вами платформы.</span><span class="sxs-lookup"><span data-stu-id="90123-p120">You must reregister the add-in in Excel in order for the new function to be available to end-users. Complete the following steps for the platform that you're using in this tutorial.</span></span>

    * <span data-ttu-id="90123-191">В случае использования Excel для Windows:</span><span class="sxs-lookup"><span data-stu-id="90123-191">If you're using Excel for Windows:</span></span>

        1. <span data-ttu-id="90123-192">Закройте Excel и снова его откройте.</span><span class="sxs-lookup"><span data-stu-id="90123-192">Close Excel and then reopen Excel.</span></span>
        
        2. <span data-ttu-id="90123-193">В Excel перейдите на вкладку **Вставка** и нажмите на стрелку вниз, расположенную справа от секции **Мои надстройки**.  ![Вставьте ленту в Excel для Windows с помощью выделенной стрелки «Мои надстройки».](../images/excel-cf-register-add-in-1b.png)</span><span class="sxs-lookup"><span data-stu-id="90123-193">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

        3. <span data-ttu-id="90123-p121">В списке доступных надстроек найдите секцию \*\* Надстройки разработчика\*\*  и выберите надстройку **Настраиваемые функции Excel**, чтобы зарегистрировать ее. ![Вставьте ленту в Excel для Windows с помощью выделенной в списке «Мои надстройки» надстройки «Настраиваемые функции Excel»](../images/excel-cf-register-add-in-2.png)</span><span class="sxs-lookup"><span data-stu-id="90123-p121">In the list of available add-ins, find the **Developer Add-ins** section and select the **Excel Custom Functions** add-in to register it.  ![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

    * <span data-ttu-id="90123-196">В случае использования Excel Online:</span><span class="sxs-lookup"><span data-stu-id="90123-196">If you're using Excel Online:</span></span> 

        1. <span data-ttu-id="90123-197">В Excel Online перейдите на вкладку **Вставка** и выберите **Надстройки**.  ![Вставьте ленту в Excel Online, используя выделенный значок "Мои надстройки".](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="90123-197">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

        2. <span data-ttu-id="90123-198">Выберите **Управление моими надстройками** и **Отправить мою надстройку**.</span><span class="sxs-lookup"><span data-stu-id="90123-198">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

        3. <span data-ttu-id="90123-199">Нажмите кнопку **Обзор...** и перейдите в корневой каталог проекта, созданный генератором Yo Office.</span><span class="sxs-lookup"><span data-stu-id="90123-199">Choose **Browse...** and navigate to the root directory of the project that the Yo Office generator created.</span></span> 

        4. <span data-ttu-id="90123-200">Выберите файл **manifest.xml** и нажмите на кнопку **Открыть**, а затем — на кнопку **Загрузить**.</span><span class="sxs-lookup"><span data-stu-id="90123-200">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

4. <span data-ttu-id="90123-p122">Теперь можно опробовать новую функцию. В ячейке **C1** введите текст `=CONTOSO.STOCKPRICESTREAM("MSFT")` и нажмите на клавишу ВВОД. Если фондовая биржа открыта, то результирующее значение в ячейке **C1** будет постоянно обновляться для отображения в реальном времени стоимости одной акции корпорации Майкрософт.</span><span class="sxs-lookup"><span data-stu-id="90123-p122">Now, let's try out the new function. In cell **C1**, type the text `=CONTOSO.STOCKPRICESTREAM("MSFT")` and press enter. Provided that the stock market is open, you should see that the result in cell **C1** is constantly updated to reflect the real-time price for one share of Microsoft stock.</span></span>

## <a name="next-steps"></a><span data-ttu-id="90123-204">Дальнейшие шаги</span><span class="sxs-lookup"><span data-stu-id="90123-204">Next steps</span></span>

<span data-ttu-id="90123-p123">В этом руководстве рассматривается создание нового проекта настраиваемых функций, создание настраиваемой функции, которая запрашивает данные от веб-сервера, и создание настраиваемой функции, осуществляющей поточную передачу данных с веб-сервера в режиме реального времени. Чтобы больше узнать о настраиваемых функциях в Excel, ознакомьтесь со следующей статьей:</span><span class="sxs-lookup"><span data-stu-id="90123-p123">In this tutorial, you've created a new custom functions project, tried out a prebuilt function, created a custom function that requests data from the web, and created a custom function that streams real-time data from the web. To learn more about custom functions in Excel, continue to the following article:</span></span> 

> [!div class="nextstepaction"]
> [<span data-ttu-id="90123-207">Создание настраиваемых функций в Excel</span><span class="sxs-lookup"><span data-stu-id="90123-207">Create custom functions in Excel (Preview)</span></span>](../excel/custom-functions-overview.md)

## <a name="legal-information"></a><span data-ttu-id="90123-208">Юридическая информация</span><span class="sxs-lookup"><span data-stu-id="90123-208">Legal Information</span></span>

<span data-ttu-id="90123-p124">Данные бесплатно предоставляются [IEX](https://iextrading.com/developer/). Ознакомьтесь с [Условиями использования IEX](https://iextrading.com/api-exhibit-a/). Содержащееся в данном руководстве описание использования API IEX корпорацией Майкрософт приводится только для обучения.</span><span class="sxs-lookup"><span data-stu-id="90123-p124">Data provided free by [IEX](https://iextrading.com/developer/). View [IEX's Terms of Use](https://iextrading.com/api-exhibit-a/). Microsoft's use of the IEX API in this tutorial is for educational purposes only.</span></span>
