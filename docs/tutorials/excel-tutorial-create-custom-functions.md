---
title: Руководство по пользовательским функциям в Excel
description: Из этого руководства вы узнаете, как создать надстройку, Excel, содержащую пользовательские функции, которые могут выполнять вычисления, запрашивать или передавать веб-данные.
ms.date: 01/02/2019
ms.topic: tutorial
ms.openlocfilehash: 2a06bbff8fff23f9cb41f914a486c9cf58bea33b
ms.sourcegitcommit: 3007bf57515b0811ff98a7e1518ecc6fc9462276
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/04/2019
ms.locfileid: "27724881"
---
# <a name="tutorial-create-custom-functions-in-excel"></a><span data-ttu-id="6b8b2-103">Руководство: создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="6b8b2-103">Tutorial: Create custom functions in Excel</span></span>

<span data-ttu-id="6b8b2-104">Пользовательские функции позволяют добавлять новые функции в Excel путем определения этих функций в JavaScript как части надстройки.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-104">Custom functions enable you to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="6b8b2-105">Пользователи в Excel могут получить доступ к пользовательским функциям так же, как и к любой встроенной функции в Excel, например `SUM()`.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-105">Users within Excel can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="6b8b2-106">Вы можете создавать пользовательские функции, которые будут выполнять простые задачи, такие как вычисления, или более сложные задачи, такие как потоковая передача данных в режиме реального времени из Интернета на лист.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-106">You can create custom functions that perform simple tasks such as custom calculations or more complex tasks such as streaming real-time data from the web into a worksheet.</span></span>

<span data-ttu-id="6b8b2-107">В этом руководстве описан порядок выполнения перечисленных ниже задач.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-107">In this tutorial, you will:</span></span>
> [!div class="checklist"]
> * <span data-ttu-id="6b8b2-108">Создание проекта пользовательских функций с помощью генератора Yo Office</span><span class="sxs-lookup"><span data-stu-id="6b8b2-108">Create a custom functions project by using the Yo Office generator</span></span>
> * <span data-ttu-id="6b8b2-109">Использование готовой пользовательской функции для выполнения простых вычислений</span><span class="sxs-lookup"><span data-stu-id="6b8b2-109">Use a prebuilt custom function to perform a simple calculation</span></span>
> * <span data-ttu-id="6b8b2-110">Создание пользовательской функции, которая запрашивает данные из Интернета</span><span class="sxs-lookup"><span data-stu-id="6b8b2-110">Create a custom function that requests data from the web</span></span>
> * <span data-ttu-id="6b8b2-111">Создание пользовательской функции, которая осуществляет потоковую передачу данных в реальном времени из Интернета</span><span class="sxs-lookup"><span data-stu-id="6b8b2-111">Create a custom function that streams real-time data from the web</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

## <a name="prerequisites"></a><span data-ttu-id="6b8b2-112">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="6b8b2-112">Prerequisites</span></span>

* <span data-ttu-id="6b8b2-113">[Node.js](https://nodejs.org/en/) (версия 8.0.0 или более поздняя)</span><span class="sxs-lookup"><span data-stu-id="6b8b2-113">[Node.js](https://nodejs.org/en/) (version 8.0.0 or later)</span></span>

* <span data-ttu-id="6b8b2-114">[Git Bash](https://git-scm.com/downloads) (или другой клиент Git)</span><span class="sxs-lookup"><span data-stu-id="6b8b2-114">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

* <span data-ttu-id="6b8b2-115">Последняя версия [Yeoman](https://yeoman.io/) и [генератора Yeoman для надстроек Office](https://www.npmjs.com/package/generator-office). Выполните в командной строке указанную ниже команду, чтобы установить эти инструменты глобально.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-115">The latest version of [Yeoman](https://yeoman.io/) and the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office). To install these tools globally, run the following command via the command prompt:</span></span>

    ```
    npm install -g yo generator-office
    ```

    > [!NOTE]
    > <span data-ttu-id="6b8b2-116">Даже если у вас установлен генератор Yeoman, рекомендуется обновить пакет до последней версии из npm.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-116">Even if you have previously installed the Yeoman generator, we recommend updating your package to the latest version from npm.</span></span>

* <span data-ttu-id="6b8b2-117">Excel для Windows (64-разрядная версия 1810 или более поздняя) или Excel Online</span><span class="sxs-lookup"><span data-stu-id="6b8b2-117">Excel for Windows (64-bit version 1810 or later) or Excel Online</span></span>

* <span data-ttu-id="6b8b2-118">Присоединитесь к [Программе предварительной оценки Office](https://products.office.com/office-insider) (уровень **Участник**; ранее "Предварительная оценка — ранний доступ")</span><span class="sxs-lookup"><span data-stu-id="6b8b2-118">Join the [Office Insider program](https://products.office.com/office-insider) (**Insider** level -- formerly called "Insider Fast")</span></span>

## <a name="create-a-custom-functions-project"></a><span data-ttu-id="6b8b2-119">Создание проекта пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="6b8b2-119">Create a custom functions project</span></span>

 <span data-ttu-id="6b8b2-120">Чтобы начать работу, создайте проект пользовательских функций с помощью генератора Yeoman.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-120">To start, you'll use the Yeoman generator to create the custom functions project.</span></span> <span data-ttu-id="6b8b2-121">Это позволит настроить для проекта правильную структуру папок, исходные файлы и зависимости, чтобы начать написание кода пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-121">This will set up your project with the correct folder structure, source files, and dependencies to begin coding your custom functions.</span></span>

1. <span data-ttu-id="6b8b2-122">Выполните указанную ниже команду и ответьте на вопросы, как показано ниже.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-122">Run the following command and then answer the prompts as follows.</span></span>

    ```
    yo office
    ```

    * <span data-ttu-id="6b8b2-123">Выберите тип проекта: `Excel Custom Functions Add-in project (...)`</span><span class="sxs-lookup"><span data-stu-id="6b8b2-123">Choose a project type: `Excel Custom Functions Add-in project (...)`</span></span>

    * <span data-ttu-id="6b8b2-124">Выберите тип сценария: `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="6b8b2-124">Choose a script type: `JavaScript`</span></span>

    * <span data-ttu-id="6b8b2-125">Как вы хотите назвать свою надстройку?</span><span class="sxs-lookup"><span data-stu-id="6b8b2-125">What do you want to name your add-in?</span></span> `stock-ticker`

    ![Генератор Yeoman для надстройки Office, приглашающий к созданию пользовательских функций](../images/12-10-fork-cf-pic.jpg)

    <span data-ttu-id="6b8b2-127">Генератор Yeoman создаст файлы проекта и установит вспомогательные компоненты Node.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-127">The Yeoman generator will create the project files and install supporting Node components.</span></span> <span data-ttu-id="6b8b2-128">Файлы проекта взяты из репозитория [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) GitHub.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-128">The project files come from the [Excel-Custom-Functions](https://github.com/OfficeDev/Excel-Custom-Functions) GitHub repository.</span></span>

2. <span data-ttu-id="6b8b2-129">Перейдите в папку проекта.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-129">Go to the project folder.</span></span>

    ```
    cd stock-ticker
    ```

3. <span data-ttu-id="6b8b2-130">Сделайте доверенным самозаверяющий сертификат, необходимый для выполнения этого проекта.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-130">Trust the self-signed certificate that is needed to run this project.</span></span> <span data-ttu-id="6b8b2-131">Подробные инструкции для Windows или Mac см. в статье [Добавление самозаверяющих сертификатов в качестве доверенных корневых сертификатов](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span><span class="sxs-lookup"><span data-stu-id="6b8b2-131">For detailed instructions for either Windows or Mac, see [Adding Self Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span></span>  

4. <span data-ttu-id="6b8b2-132">Выполните сборку проекта.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-132">Build the project.</span></span>

    ```
    npm run build
    ```

5. <span data-ttu-id="6b8b2-133">Запустите локальный веб-сервер, работающий на Node.js.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-133">Start the local web server, which runs in Node.js.</span></span>

    * <span data-ttu-id="6b8b2-134">Если вы будете использовать Excel для Windows для тестирования ваших пользовательских функций, выполните следующую команду, чтобы запустить локальный веб-сервер, запустить программу Excel и загрузить неопубликованную надстройку:</span><span class="sxs-lookup"><span data-stu-id="6b8b2-134">If you'll be using Excel for Windows to test your custom functions, run the following command to start the local web server, launch Excel, and sideload the add-in:</span></span>

        ```
         npm run start
        ```
        <span data-ttu-id="6b8b2-135">После выполнения этой команды, в командной строке отобразятся сведения о выполненных действиях, откроется другое окно npm со сведениями о сборке, и запустится Excel с загруженной надстройкой.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-135">After running this command, your command prompt will show details about what has been done, another npm window will open showing the details of the build, and Excel will start with your add-in loaded.</span></span> <span data-ttu-id="6b8b2-136">Если надстройка не загружается, проверьте правильность выполнения шага 3.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-136">If you add-in does not load, check that you have completed step 3 properly.</span></span>

    * <span data-ttu-id="6b8b2-137">Если вы будете использовать Excel Online для тестирования ваших пользовательских функций, выполните следующую команду, чтобы запустить локальный веб-сервер:</span><span class="sxs-lookup"><span data-stu-id="6b8b2-137">If you'll be using Excel Online to test your custom functions, run the following command to start the local web server:</span></span>

        ```
        npm run start-web
        ```

         <span data-ttu-id="6b8b2-138">После выполнения этой команды откроется другое окно со сведениями о сборке.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-138">After running this command, another window will open showing you the details of the build.</span></span> <span data-ttu-id="6b8b2-139">Чтобы использовать свои функции, откройте новую книгу в Office Online.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-139">To use your functions, open a new workbook in Office Online.</span></span>

## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="6b8b2-140">Проверка работы готовой пользовательской функции</span><span class="sxs-lookup"><span data-stu-id="6b8b2-140">Try out a prebuilt custom function</span></span>

<span data-ttu-id="6b8b2-141">Проект пользовательских функций, созданный с помощью генератора Yeoman, содержит некоторые готовые пользовательские функции, определенные в файле **src/customfunctions.js**.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-141">The custom functions project that you created by using the Yeoman generator contains some prebuilt custom functions, defined within the **src/customfunctions.js** file.</span></span> <span data-ttu-id="6b8b2-142">Файл **manifest.xml** в корневом каталоге проекта указывает, что все пользовательские функции принадлежат пространству имен `CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-142">The **manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.</span></span>

<span data-ttu-id="6b8b2-143">В книге Excel попробуйте, как работает пользовательская функция `ADD`, выполнив описанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-143">In your Excel workbook, try out the `ADD` custom function by completing the following steps in Excel:</span></span>

1. <span data-ttu-id="6b8b2-144">Введите в ячейке `=CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-144">Within a cell, type `=CONTOSO`.</span></span> <span data-ttu-id="6b8b2-145">Обратите внимание на то, что в меню автозаполнения содержится список всех функций в пространстве имен `CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-145">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="6b8b2-146">Выполните запуск функции `CONTOSO.ADD` с числами `10` и `200` в качестве входных параметров, введя значение `=CONTOSO.ADD(10,200)` в ячейке и нажав клавишу ВВОД.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-146">Run the `CONTOSO.ADD` function, with numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.</span></span>

<span data-ttu-id="6b8b2-147">Пользовательская функция `ADD` вычисляет сумму двух чисел, которые вы указываете в качестве входных параметров.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-147">The `ADD` custom function computes the sum of the two numbers that you specify as input parameters.</span></span> <span data-ttu-id="6b8b2-148">При вводе `=CONTOSO.ADD(10,200)` в ячейке должен отобразиться результат **210** после нажатия клавиши ВВОД.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-148">Typing `=CONTOSO.ADD(10,200)` should produce the result **210** in the cell after you press enter.</span></span>

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a><span data-ttu-id="6b8b2-149">Создание пользовательской функции, которая запрашивает данные из Интернета</span><span class="sxs-lookup"><span data-stu-id="6b8b2-149">Create a custom function that requests data from the web</span></span>

<span data-ttu-id="6b8b2-150">Что делать, если требуется функция, которая сможет запросить цену на акцию из API и отобразить результат в ячейке на листе?</span><span class="sxs-lookup"><span data-stu-id="6b8b2-150">What if you needed a function that could request the price of a stock from an API and display the result in the cell of a worksheet?</span></span> <span data-ttu-id="6b8b2-151">Пользовательские функции разрабатываются таким образом, что вы можете легко асинхронно запросить данные из Интернета.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-151">Custom functions are designed so that you can easily request data from the web asynchronously.</span></span>

<span data-ttu-id="6b8b2-152">Выполните указанные ниже действия, чтобы создать пользовательскую функцию с именем `stockPrice`, которая принимает код акции (например, **MSFT**) и возвращает цену этой акции.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-152">Complete the following steps to create a custom function named `stockPrice` that accepts a stock ticker symbol (e.g., **MSFT**) and returns the price of that stock.</span></span> <span data-ttu-id="6b8b2-153">Такая пользовательская функция использует API IEX Trading, который предоставляется бесплатно и не требует проверки подлинности.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-153">This custom function uses the IEX Trading API, which is free and does not require authentication.</span></span>

1. <span data-ttu-id="6b8b2-154">В проекте **stock-ticker**, созданном генератором Yeoman, найдите файл **src/customfunctions.js** и откройте его в редакторе кода.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-154">In the **stock-ticker** project that the Yeoman generator created, find the file **src/customfunctions.js** and open it in your code editor.</span></span>

2. <span data-ttu-id="6b8b2-155">В файле **customfunctions.js** найдите функцию `increment` и добавьте приведенный ниже код сразу после этой функции.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-155">In **customfunctions.js**, locate the `increment` function and add the following code immediately after that function.</span></span>

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

4. <span data-ttu-id="6b8b2-156">Прежде чем можно будет сделать в Excel такую функцию доступной, необходимо указать метаданные, чтобы описать функцию для Excel.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-156">Before Excel can make this new function available, you must specify metadata to describe the function to Excel.</span></span> <span data-ttu-id="6b8b2-157">Откройте файл **config/customfunctions.json**.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-157">Open the **config/customfunctions.json** file.</span></span> <span data-ttu-id="6b8b2-158">Добавьте указанный ниже объект JSON в массив 'functions' и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-158">Add the following JSON object to the 'functions' array and save the file.</span></span>

    <span data-ttu-id="6b8b2-159">Объект JSON описывает функцию `stockPrice`.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-159">This JSON describes the `stockPrice` function.</span></span>

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
                "description": "stock symbol",
                "type": "string",
                "dimensionality": "scalar"
            }
        ]
    }
    ```

5. <span data-ttu-id="6b8b2-160">Необходимо повторно зарегистрировать надстройку в Excel, чтобы новая функция стала доступной конечным пользователям.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-160">You must re-register the add-in in Excel in order for the new function to be available to end-users.</span></span> <span data-ttu-id="6b8b2-161">Выполните указанные ниже действия для платформы, которую вы используете в этом руководстве.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-161">Complete the following steps for the platform that you're using in this tutorial.</span></span>

    * <span data-ttu-id="6b8b2-162">Если вы используете Excel для Windows, выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-162">If you're using Excel for Windows:</span></span>

        1. <span data-ttu-id="6b8b2-163">Закройте Excel, а затем откройте Excel повторно.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-163">Close Excel and then reopen Excel.</span></span>

        2. <span data-ttu-id="6b8b2-164">В Excel выберите вкладку **Вставка**, а затем нажмите стрелку вниз, которая находится справа от пункта **Мои надстройки**. ![Вставьте ленту в Excel для Windows с выделенной стрелкой "Мои надстройки"](../images/excel-cf-register-add-in-1b.png).</span><span class="sxs-lookup"><span data-stu-id="6b8b2-164">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

        3. <span data-ttu-id="6b8b2-165">В списке доступных надстроек найдите раздел **Надстройки разработчика** и выберите надстройку **stock-ticker**, чтобы зарегистрировать ее.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-165">In the list of available add-ins, find the **Developer Add-ins** section and select the **stock-ticker** add-in to register it.</span></span>
            <span data-ttu-id="6b8b2-166">![Вставьте ленту в Excel для Windows с выделенной надстройкой "Пользовательские функции Excel" в списке "Мои надстройки"](../images/excel-cf-register-add-in-2.png).</span><span class="sxs-lookup"><span data-stu-id="6b8b2-166">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

    * <span data-ttu-id="6b8b2-167">Если вы используете Excel Online, выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-167">If you're using Excel Online:</span></span>

        1. <span data-ttu-id="6b8b2-168">В Excel Online выберите вкладку **Вставка**, а затем выберите **Надстройки**. ![Вставьте ленту в Excel Online с выделенным значком "Мои надстройки"](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="6b8b2-168">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

        2. <span data-ttu-id="6b8b2-169">Выберите пункт **Управление моими надстройками**, а затем выберите **Отправить мою надстройку**.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-169">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span> 

        3. <span data-ttu-id="6b8b2-170">Выберите \*\*Обзор... \*\* и откройте корневой каталог проекта, созданный генератором Yeoman.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-170">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span> 

        4. <span data-ttu-id="6b8b2-171">Выберите файл **manifest.xml** и нажмите кнопку **Открыть**, затем нажмите кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-171">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

6. <span data-ttu-id="6b8b2-172">Теперь давайте попробуем, как работает новая функция.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-172">Now, let's try out the new function.</span></span> <span data-ttu-id="6b8b2-173">В ячейке **B1** введите текст `=CONTOSO.STOCKPRICE("MSFT")` и нажмите клавишу ВВОД.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-173">In cell **B1**, type the text `=CONTOSO.STOCKPRICE("MSFT")` and press enter.</span></span> <span data-ttu-id="6b8b2-174">Вы увидите, что результат в ячейке **B1** является текущей ценой одной акции корпорации Майкрософт.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-174">You should see that the result in cell **B1** is the current stock price for one share of Microsoft stock.</span></span>

## <a name="create-a-streaming-asynchronous-custom-function"></a><span data-ttu-id="6b8b2-175">Создание потоковой асинхронной пользовательской функции</span><span class="sxs-lookup"><span data-stu-id="6b8b2-175">Create a streaming asynchronous custom function</span></span>

<span data-ttu-id="6b8b2-176">Функция `stockPrice`, которую вы только что создали, возвращает цену акции в конкретный момент времени, однако цены на акции всегда меняются.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-176">The `stockPrice` function that you just created returns the price of a stock at a specific moment in time, but stock prices are always changing.</span></span> <span data-ttu-id="6b8b2-177">Давайте создадим пользовательскую функцию, которая осуществляет потоковую передачу данных из API, чтобы получать обновления цен на акции в реальном времени.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-177">Let's create a custom function that streams data from an API to get real-time updates on a stock price.</span></span>

<span data-ttu-id="6b8b2-178">Выполните указанные ниже действия, чтобы создать функцию с именем `stockPriceStream`, которая будет запрашивать цену указанной акции каждые 1000 миллисекунд (при условии, что предыдущий запрос был выполнен).</span><span class="sxs-lookup"><span data-stu-id="6b8b2-178">Complete the following steps to create a custom function named `stockPriceStream` that requests the price of the specified stock every 1000 milliseconds (provided that the previous request has completed).</span></span> <span data-ttu-id="6b8b2-179">Во время выполнения первоначального запроса в ячейке, где вызывается функция, может появиться значение-заполнитель **#GETTING_DATA**.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-179">While the initial request is in-progress, you may see the placeholder value **#GETTING_DATA** the cell where the function is being called.</span></span> <span data-ttu-id="6b8b2-180">Когда значение будет возвращено функцией, оно заменит значение-заполнитель **#GETTING_DATA** в ячейке.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-180">When a value is returned by the function, **#GETTING_DATA** will be replaced by that value in the cell.</span></span>

1. <span data-ttu-id="6b8b2-181">В проекте **stock-ticker**, созданном генератором Yeoman, добавьте указанный ниже код в файл **src/customfunctions.js** и сохраните его.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-181">In the **stock-ticker** project that the Yeoman generator created, add the following code to **src/customfunctions.js** and save the file.</span></span>

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

2. <span data-ttu-id="6b8b2-182">Прежде чем можно будет сделать в Excel такую функцию доступной пользователям, укажите метаданные, описывающие эту функцию.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-182">Before Excel can make this new function available to users, specify metadata that describes this function.</span></span> <span data-ttu-id="6b8b2-183">В проекте **stock-ticker**, созданном генератором Yeoman, добавьте указанный ниже объект в массив `functions` в файле **config/customfunctions.json** и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-183">In the **stock-ticker** project that the Yeoman generator created, add the following object to the `functions` array within the **config/customfunctions.json** file and save the file.</span></span>

    <span data-ttu-id="6b8b2-184">Объект JSON описывает функцию `stockPriceStream`.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-184">This JSON describes the `stockPriceStream` function.</span></span> <span data-ttu-id="6b8b2-185">Для любой функции потоковой передачи свойство `stream` и свойство `cancelable` должны быть заданы как `true` в объекте `options`, как показано в этом примере кода.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-185">For any streaming function, the `stream` property and the `cancelable` property must be set to `true` within the `options` object, as shown in this code sample.</span></span>

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
                "description": "stock symbol",
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

3. <span data-ttu-id="6b8b2-186">Необходимо повторно зарегистрировать надстройку в Excel, чтобы новая функция стала доступной конечным пользователям.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-186">You must re-register the add-in in Excel in order for the new function to be available to end-users.</span></span> <span data-ttu-id="6b8b2-187">Выполните указанные ниже действия для платформы, которую вы используете в этом руководстве.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-187">Complete the following steps for the platform that you're using in this tutorial.</span></span>

    * <span data-ttu-id="6b8b2-188">Если вы используете Excel для Windows, выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-188">If you're using Excel for Windows:</span></span>

        1. <span data-ttu-id="6b8b2-189">Закройте Excel, а затем откройте Excel повторно.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-189">Close Excel and then reopen Excel.</span></span>
        
        2. <span data-ttu-id="6b8b2-190">В Excel выберите вкладку **Вставка**, а затем нажмите стрелку вниз, которая находится справа от пункта **Мои надстройки**. ![Вставьте ленту в Excel для Windows с выделенной стрелкой "Мои надстройки"](../images/excel-cf-register-add-in-1b.png).</span><span class="sxs-lookup"><span data-stu-id="6b8b2-190">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel for Windows with the My Add-ins arrow highlighted](../images/excel-cf-register-add-in-1b.png)</span></span>

        3. <span data-ttu-id="6b8b2-191">В списке доступных надстроек найдите раздел **Надстройки разработчика** и выберите надстройку **stock-ticker**, чтобы зарегистрировать ее.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-191">In the list of available add-ins, find the **Developer Add-ins** section and select the **stock-ticker** add-in to register it.</span></span>
            <span data-ttu-id="6b8b2-192">![Вставьте ленту в Excel для Windows с выделенной надстройкой "Пользовательские функции Excel" в списке "Мои надстройки"](../images/excel-cf-register-add-in-2.png).</span><span class="sxs-lookup"><span data-stu-id="6b8b2-192">![Insert ribbon in Excel for Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/excel-cf-register-add-in-2.png)</span></span>

    * <span data-ttu-id="6b8b2-193">Если вы используете Excel Online, выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-193">If you're using Excel Online:</span></span>

        1. <span data-ttu-id="6b8b2-194">В Excel Online выберите вкладку **Вставка**, а затем выберите **Надстройки**. ![Вставьте ленту в Excel Online с выделенным значком "Мои надстройки"](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="6b8b2-194">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel Online with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

        2. <span data-ttu-id="6b8b2-195">Выберите пункт **Управление моими надстройками**, а затем выберите **Отправить мою надстройку**.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-195">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

        3. <span data-ttu-id="6b8b2-196">Выберите \*\*Обзор... \*\* и откройте корневой каталог проекта, созданный генератором Yeoman.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-196">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

        4. <span data-ttu-id="6b8b2-197">Выберите файл **manifest.xml** и нажмите кнопку **Открыть**, затем нажмите кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-197">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

4. <span data-ttu-id="6b8b2-198">Теперь давайте попробуем, как работает новая функция.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-198">Now, let's try out the new function.</span></span> <span data-ttu-id="6b8b2-199">В ячейке **C1** введите текст `=CONTOSO.STOCKPRICESTREAM("MSFT")` и нажмите клавишу ВВОД.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-199">In cell **C1**, type the text `=CONTOSO.STOCKPRICESTREAM("MSFT")` and press enter.</span></span> <span data-ttu-id="6b8b2-200">Если рынок ценных бумаг открыт, вы увидите, что результат в ячейке **C1** постоянно обновляется, отражая в режиме реального времени цену одной акции корпорации Майкрософт.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-200">Provided that the stock market is open, you should see that the result in cell **C1** is constantly updated to reflect the real-time price for one share of Microsoft stock.</span></span>

## <a name="next-steps"></a><span data-ttu-id="6b8b2-201">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="6b8b2-201">Next steps</span></span>

<span data-ttu-id="6b8b2-202">В ходе работы с данным руководством вы создали новый проект пользовательских функций, попробовали, как работает готовая функция, создали пользовательскую функцию, которая запрашивает данные из Интернета, а также создали пользовательскую функцию, которая осуществляет потоковую передачу данных в реальном времени из Интернета.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-202">In this tutorial, you've created a new custom functions project, tried out a prebuilt function, created a custom function that requests data from the web, and created a custom function that streams real-time data from the web.</span></span> <span data-ttu-id="6b8b2-203">Чтобы узнать больше о пользовательских функции в Excel, перейдите к следующей статье:</span><span class="sxs-lookup"><span data-stu-id="6b8b2-203">To learn more about custom functions in Excel, continue to the following article:</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="6b8b2-204">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="6b8b2-204">Create custom functions in Excel</span></span>](../excel/custom-functions-overview.md)

### <a name="legal-information"></a><span data-ttu-id="6b8b2-205">Юридические сведения</span><span class="sxs-lookup"><span data-stu-id="6b8b2-205">Legal information</span></span>

<span data-ttu-id="6b8b2-206">Данные предоставлены бесплатно компанией [IEX](https://iextrading.com/developer/).</span><span class="sxs-lookup"><span data-stu-id="6b8b2-206">Data provided free by [IEX](https://iextrading.com/developer/).</span></span> <span data-ttu-id="6b8b2-207">Ознакомьтесь с [Условиями использования IEX](https://iextrading.com/api-exhibit-a/).</span><span class="sxs-lookup"><span data-stu-id="6b8b2-207">View [IEX's Terms of Use](https://iextrading.com/api-exhibit-a/).</span></span> <span data-ttu-id="6b8b2-208">Корпорация Майкрософт использует API компании IEX в этом руководстве исключительно в ознакомительных целях.</span><span class="sxs-lookup"><span data-stu-id="6b8b2-208">Microsoft's use of the IEX API in this tutorial is for educational purposes only.</span></span>


