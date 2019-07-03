---
title: Руководство по пользовательским функциям в Excel
description: Из этого руководства вы узнаете, как создать надстройку, Excel, содержащую пользовательские функции, которые могут выполнять вычисления, запрашивать или передавать веб-данные.
ms.date: 06/27/2019
ms.prod: excel
ms.topic: tutorial
localization_priority: Normal
ms.openlocfilehash: 1aa05581d1b0dfb1f5affa019e51b84126c8d199
ms.sourcegitcommit: 90c2d8236c6b30d80ac2b13950028a208ef60973
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/02/2019
ms.locfileid: "35454734"
---
# <a name="tutorial-create-custom-functions-in-excel"></a><span data-ttu-id="9ed03-103">Руководство: создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="9ed03-103">Tutorial: Create custom functions in Excel</span></span>

<span data-ttu-id="9ed03-104">Пользовательские функции позволяют добавлять новые функции в Excel путем определения этих функций в JavaScript как части надстройки.</span><span class="sxs-lookup"><span data-stu-id="9ed03-104">Custom functions enable you to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="9ed03-105">Пользователи в Excel могут получить доступ к пользовательским функциям так же, как и к любой встроенной функции в Excel, например `SUM()`.</span><span class="sxs-lookup"><span data-stu-id="9ed03-105">Users within Excel can access custom functions as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="9ed03-106">Вы можете создавать пользовательские функции, которые будут выполнять простые задачи, такие как вычисления, или более сложные задачи, такие как потоковая передача данных в режиме реального времени из Интернета на лист.</span><span class="sxs-lookup"><span data-stu-id="9ed03-106">You can create custom functions that perform simple tasks like calculations or more complex tasks such as streaming real-time data from the web into a worksheet.</span></span>

<span data-ttu-id="9ed03-107">В этом руководстве описан порядок выполнения перечисленных ниже задач.</span><span class="sxs-lookup"><span data-stu-id="9ed03-107">In this tutorial, you will:</span></span>
> [!div class="checklist"]
> * <span data-ttu-id="9ed03-108">Создание надстройки пользовательской функции с помощью [генератора Yeoman для надстроек Office](https://www.npmjs.com/package/generator-office).</span><span class="sxs-lookup"><span data-stu-id="9ed03-108">Create a custom function add-in using the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office).</span></span> 
> * <span data-ttu-id="9ed03-109">Использование готовой пользовательской функции для выполнения простых вычислений</span><span class="sxs-lookup"><span data-stu-id="9ed03-109">Use a prebuilt custom function to perform a simple calculation.</span></span>
> * <span data-ttu-id="9ed03-110">Создание пользовательской функции, которая получает данные из сети Интернет.</span><span class="sxs-lookup"><span data-stu-id="9ed03-110">Create a custom function that gets data from the web.</span></span>
> * <span data-ttu-id="9ed03-111">Создание пользовательской функции, которая осуществляет потоковую передачу данных в реальном времени из сети Интернет</span><span class="sxs-lookup"><span data-stu-id="9ed03-111">Create a custom function that streams real-time data from the web.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="9ed03-112">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="9ed03-112">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

* <span data-ttu-id="9ed03-113">Excel в Windows (версия 1904 или более поздняя версия, подключенная к подписке на Office 365) или в Интернете</span><span class="sxs-lookup"><span data-stu-id="9ed03-113">Excel on Windows (version 1904 or later, connected to Office 365 subscription) or on the web</span></span>

## <a name="create-a-custom-functions-project"></a><span data-ttu-id="9ed03-114">Создание проекта пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="9ed03-114">Create a custom functions project</span></span>

 <span data-ttu-id="9ed03-115">Чтобы начать, вам необходимо создать проект кода для разработки надстройки пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="9ed03-115">To start, you'll create the code project to build your custom function add-in.</span></span> <span data-ttu-id="9ed03-116">[Генератор Yeoman для надстроек Office](https://www.npmjs.com/package/generator-office) настроит проект с помощью некоторых предварительно созданных настраиваемых функций, которые можно испытать. Если вы уже запустили функцию быстрого запуска пользовательских функций и создали проект, продолжайте использовать этот проект и переходите к [этому шагу](#create-a-custom-function-that-requests-data-from-the-web) .</span><span class="sxs-lookup"><span data-stu-id="9ed03-116">The [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) will set up your project with some prebuilt custom functions that you can try out. If you have already run the custom functions quick start and generated a project, continue to use that project and skip to [this step](#create-a-custom-function-that-requests-data-from-the-web) instead.</span></span>

1. <span data-ttu-id="9ed03-117">Выполните указанную ниже команду и ответьте на вопросы, как показано ниже.</span><span class="sxs-lookup"><span data-stu-id="9ed03-117">Run the following command and then answer the prompts as follows.</span></span>
    
    ```command&nbsp;line
    yo office
    ```
    
    * <span data-ttu-id="9ed03-118">**Выберите тип проекта:** `Excel Custom Functions Add-in project`</span><span class="sxs-lookup"><span data-stu-id="9ed03-118">**Choose a project type:** `Excel Custom Functions Add-in project`</span></span>
    * <span data-ttu-id="9ed03-119">**Выберите тип сценария:** `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="9ed03-119">**Choose a script type:** `JavaScript`</span></span>
    * <span data-ttu-id="9ed03-120">**Как вы хотите назвать надстройку?**</span><span class="sxs-lookup"><span data-stu-id="9ed03-120">**What do you want to name your add-in?**</span></span> `starcount`

    ![Генератор Yeoman для надстройки Office, приглашающий к созданию пользовательских функций](../images/starcountPrompt.png)
    
    <span data-ttu-id="9ed03-122">Генератор Yeoman создаст файлы проекта и установит вспомогательные компоненты Node.</span><span class="sxs-lookup"><span data-stu-id="9ed03-122">The Yeoman generator will create the project files and install supporting Node components.</span></span>

2. <span data-ttu-id="9ed03-123">Перейдите к корневой папке проекта.</span><span class="sxs-lookup"><span data-stu-id="9ed03-123">Navigate to the root folder of the project.</span></span>
    
    ```command&nbsp;line
    cd starcount
    ```

3. <span data-ttu-id="9ed03-124">Выполните построение проекта.</span><span class="sxs-lookup"><span data-stu-id="9ed03-124">Build the project.</span></span>
    
    ```command&nbsp;line
    npm run build
    ```

    > [!NOTE]
    > <span data-ttu-id="9ed03-125">Надстройки Office должны использовать HTTPS, а не HTTP, даже в случае разработки.</span><span class="sxs-lookup"><span data-stu-id="9ed03-125">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="9ed03-126">Если вам будет предложено установить сертификат после того, как вы запустите `npm run build`, примите предложение установить сертификат от генератора Yeoman.</span><span class="sxs-lookup"><span data-stu-id="9ed03-126">If you are prompted to install a certificate after you run `npm run build`, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

4. <span data-ttu-id="9ed03-127">Запустите локальный веб-сервер, работающий на Node.js.</span><span class="sxs-lookup"><span data-stu-id="9ed03-127">Start the local web server, which runs in Node.js.</span></span> <span data-ttu-id="9ed03-128">Вы можете испытать надстройку настраиваемой функции в Excel в Интернете или в Windows.</span><span class="sxs-lookup"><span data-stu-id="9ed03-128">You can try out the custom function add-in in Excel on the web or Windows.</span></span>

# <a name="excel-on-windowstabexcel-windows"></a>[<span data-ttu-id="9ed03-129">Excel в Windows</span><span class="sxs-lookup"><span data-stu-id="9ed03-129">Excel on Windows</span></span>](#tab/excel-windows)

<span data-ttu-id="9ed03-130">Чтобы протестировать надстройку в Excel в Windows, выполните следующую команду.</span><span class="sxs-lookup"><span data-stu-id="9ed03-130">To test your add-in in Excel on Windows, run the following command.</span></span> <span data-ttu-id="9ed03-131">При выполнении этой команды запустится локальный веб-сервер, и откроется приложение Excel с загруженной надстройкой.</span><span class="sxs-lookup"><span data-stu-id="9ed03-131">When you run this command, the local web server will start and Excel will open with your add-in loaded.</span></span>

```command&nbsp;line
npm run start:desktop
```

# <a name="excel-on-the-webtabexcel-online"></a>[<span data-ttu-id="9ed03-132">Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="9ed03-132">Excel on the web</span></span>](#tab/excel-online)

<span data-ttu-id="9ed03-133">Чтобы протестировать надстройку в Excel в браузере, выполните следующую команду.</span><span class="sxs-lookup"><span data-stu-id="9ed03-133">To test your add-in in Excel on a browser, run the following command.</span></span> <span data-ttu-id="9ed03-134">После выполнения этой команды запустится локальный веб-сервер.</span><span class="sxs-lookup"><span data-stu-id="9ed03-134">When you run this command, the local web server will start.</span></span>

```command&nbsp;line
npm run start:web
```

<span data-ttu-id="9ed03-135">Чтобы использовать надстройку с пользовательскими функциями, откройте новую книгу в Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="9ed03-135">To use your custom functions add-in, open a new workbook in Excel on the web.</span></span> <span data-ttu-id="9ed03-136">В этой книге выполните следующие действия, чтобы Загрузка неопубликованных надстройку.</span><span class="sxs-lookup"><span data-stu-id="9ed03-136">In this workbook, complete the following steps to sideload your add-in.</span></span>

1. <span data-ttu-id="9ed03-137">В Excel перейдите на вкладку **Вставка** , а затем выберите \*\*\*\* пункт надстройки.</span><span class="sxs-lookup"><span data-stu-id="9ed03-137">In Excel, choose the **Insert** tab and then choose **Add-ins**.</span></span>

   ![Вставка ленты в Excel в Интернете с выделенным значком "Мои надстройки"](../images/excel-cf-online-register-add-in-1.png)
   
2. <span data-ttu-id="9ed03-139">Выберите пункт **Управление моими надстройками**, а затем выберите **Отправить мою надстройку**.</span><span class="sxs-lookup"><span data-stu-id="9ed03-139">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="9ed03-140">Выберите \*\*Обзор... \*\* и откройте корневой каталог проекта, созданный генератором Yeoman.</span><span class="sxs-lookup"><span data-stu-id="9ed03-140">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="9ed03-141">Выберите файл **manifest.xml** и нажмите кнопку **Открыть**, затем нажмите кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="9ed03-141">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

--- 
    
## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="9ed03-142">Проверка работы готовой пользовательской функции</span><span class="sxs-lookup"><span data-stu-id="9ed03-142">Try out a prebuilt custom function</span></span>

<span data-ttu-id="9ed03-143">Созданный проект пользовательских функций содержит некоторые предварительно созданные пользовательские функции, определенные в файле **./СРК/функтионс/функтионс.ЖС** .</span><span class="sxs-lookup"><span data-stu-id="9ed03-143">The custom functions project that you created contains some prebuilt custom functions, defined within the **./src/functions/functions.js** file.</span></span> <span data-ttu-id="9ed03-144">Файл **./manifest.xml** указывает, что все пользовательские функции принадлежат пространству имен `CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="9ed03-144">The **./manifest.xml** file specifies that all custom functions belong to the `CONTOSO` namespace.</span></span> <span data-ttu-id="9ed03-145">Вы будете использовать пространство имен CONTOSO для доступа к пользовательским функциям в Excel.</span><span class="sxs-lookup"><span data-stu-id="9ed03-145">You'll use the CONTOSO namespace to access the custom functions in Excel.</span></span>

<span data-ttu-id="9ed03-146">Затем вы проверите пользовательскую функцию `ADD`, выполнив описанные ниже действия:</span><span class="sxs-lookup"><span data-stu-id="9ed03-146">Next you'll try out the `ADD` custom function by completing the following steps:</span></span>

1. <span data-ttu-id="9ed03-147">В Excel перейдите в любую ячейку и введите `=CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="9ed03-147">In Excel, go to any cell and enter `=CONTOSO`.</span></span> <span data-ttu-id="9ed03-148">Обратите внимание на то, что в меню автозаполнения содержится список всех функций в пространстве имен `CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="9ed03-148">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="9ed03-149">Выполните запуск функции `CONTOSO.ADD` с числами `10` и `200` в качестве входных параметров, введя значение `=CONTOSO.ADD(10,200)` в ячейке и нажав клавишу ВВОД.</span><span class="sxs-lookup"><span data-stu-id="9ed03-149">Run the `CONTOSO.ADD` function, with numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.</span></span>

<span data-ttu-id="9ed03-150">Пользовательская функция `ADD` вычисляет сумму двух чисел, которые вы указываете и возвращает результат **210**.</span><span class="sxs-lookup"><span data-stu-id="9ed03-150">The `ADD` custom function computes the sum of the two numbers that you provided and returns the result of **210**.</span></span>

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a><span data-ttu-id="9ed03-151">Создание пользовательской функции, которая запрашивает данные из сети Интернет</span><span class="sxs-lookup"><span data-stu-id="9ed03-151">Create a custom function that requests data from the web</span></span>

<span data-ttu-id="9ed03-152">Интеграция данных из Интернета — отличный способ расширения функционала Excel через пользовательские функции.</span><span class="sxs-lookup"><span data-stu-id="9ed03-152">Integrating data from the Web is a great way to extend Excel through custom functions.</span></span> <span data-ttu-id="9ed03-153">Далее вы создадите пользовательскую функцию с именем `getStarCount` , которая показывает количество звезд, которыми обладает данный репозиторий GitHub.</span><span class="sxs-lookup"><span data-stu-id="9ed03-153">Next you’ll create a custom function named `getStarCount` that shows how many stars a given Github repository possesses.</span></span>

1. <span data-ttu-id="9ed03-154">В проекте **старкаунт** найдите файл **./СРК/функтионс/функтионс.ЖС** и откройте его в редакторе кода.</span><span class="sxs-lookup"><span data-stu-id="9ed03-154">In the **starcount** project, find the file **./src/functions/functions.js** and open it in your code editor.</span></span> 

2. <span data-ttu-id="9ed03-155">В файле **Function. js**добавьте следующий код:</span><span class="sxs-lookup"><span data-stu-id="9ed03-155">In **function.js**, add the following code:</span></span> 

```JS
 /**
   * Gets the star count for a given Github repository.
   * @customfunction 
   * @param {string} userName string name of Github user or organization.
   * @param {string} repoName string name of the Github repository.
   * @return {number} number of stars given to a Github repository.
   */
    async function getStarCount(userName, repoName) {
      try {
        //You can change this URL to any web request you want to work with.
        const url = "https://api.github.com/repos/" + userName + "/" + repoName;
        const response = await fetch(url);
        //Expect that status code is in 200-299 range
        if (!response.ok) {
          throw new Error(response.statusText)
        }
          const jsonResponse = await response.json();
          return jsonResponse.watchers_count;
      }
      catch (error) {
        return error;
      }
      }
    CustomFunctions.associate("GETSTARCOUNT", getStarCount);
```

<span data-ttu-id="9ed03-156">Код `CustomFunctions.associate` сопоставляет `id` функции с адресом функции `getStarCount` в JavaScript, чтобы Excel мог вызвать вашу функцию.</span><span class="sxs-lookup"><span data-stu-id="9ed03-156">The `CustomFunctions.associate` code associates the `id` of the function with the function address of `getStarCount` in JavaScript so that Excel can call your function.</span></span>

3. <span data-ttu-id="9ed03-157">Выполните указанную ниже команду, чтобы повторно собрать проект.</span><span class="sxs-lookup"><span data-stu-id="9ed03-157">Run the following command to rebuild the project.</span></span>

    ```command&nbsp;line
    npm run build
    ```

4. <span data-ttu-id="9ed03-158">Выполните следующие действия (для Excel в Интернете или Windows), чтобы повторно зарегистрировать надстройку в Excel.</span><span class="sxs-lookup"><span data-stu-id="9ed03-158">Complete the following steps (for either Excel on the web or Windows) to re-register the add-in in Excel.</span></span> <span data-ttu-id="9ed03-159">Прежде чем новая функция станет доступна, необходимо выполнить указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="9ed03-159">You must complete these steps before the new function will be available.</span></span>

### <a name="excel-on-windowstabexcel-windows"></a>[<span data-ttu-id="9ed03-160">Excel в Windows</span><span class="sxs-lookup"><span data-stu-id="9ed03-160">Excel on Windows</span></span>](#tab/excel-windows)

1. <span data-ttu-id="9ed03-161">Закройте Excel, а затем откройте Excel повторно.</span><span class="sxs-lookup"><span data-stu-id="9ed03-161">Close Excel and then reopen Excel.</span></span>

2. <span data-ttu-id="9ed03-162">В Excel перейдите на вкладку **Вставка** , а затем щелкните стрелку вниз, расположенную справа от **моих надстроек**.  ![Вставка ленты в Excel в Windows с выделенной стрелкой "Мои надстройки"](../images/select-insert.png)</span><span class="sxs-lookup"><span data-stu-id="9ed03-162">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel on Windows with the My Add-ins arrow highlighted](../images/select-insert.png)</span></span>

3. <span data-ttu-id="9ed03-163">В списке доступных надстроек найдите раздел надстройки для **разработчиков** и выберите надстройку **старкаунт** , чтобы зарегистрировать ее.</span><span class="sxs-lookup"><span data-stu-id="9ed03-163">In the list of available add-ins, find the **Developer Add-ins** section and select the **starcount** add-in to register it.</span></span>
    <span data-ttu-id="9ed03-164">![Вставка ленты в Excel в Windows с выделенной надстройкой "пользовательские функции Excel" в списке "Мои надстройки"](../images/list-starcount.png)</span><span class="sxs-lookup"><span data-stu-id="9ed03-164">![Insert ribbon in Excel on Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/list-starcount.png)</span></span>


# <a name="excel-on-the-webtabexcel-online"></a>[<span data-ttu-id="9ed03-165">Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="9ed03-165">Excel on the web</span></span>](#tab/excel-online)

1. <span data-ttu-id="9ed03-166">В Excel перейдите на вкладку **Вставка** , а затем выберите \*\*\*\* пункт надстройки.  ![Вставка ленты в Excel в Интернете с выделенным значком "Мои надстройки"](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="9ed03-166">In Excel, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel on the web with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

2. <span data-ttu-id="9ed03-167">Выберите пункт **Управление моими надстройками**, а затем выберите **Отправить мою надстройку**.</span><span class="sxs-lookup"><span data-stu-id="9ed03-167">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="9ed03-168">Выберите \*\*Обзор... \*\* и откройте корневой каталог проекта, созданный генератором Yeoman.</span><span class="sxs-lookup"><span data-stu-id="9ed03-168">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="9ed03-169">Выберите файл **manifest.xml** и нажмите **Открыть**, затем нажмите кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="9ed03-169">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

---

<ol start="5">
<li> <span data-ttu-id="9ed03-170">Теперь давайте оценим, как работает новая функция.</span><span class="sxs-lookup"><span data-stu-id="9ed03-170">Try out the new function.</span></span> <span data-ttu-id="9ed03-171">В ячейке <strong>B1</strong>введите текст <strong>= contoso. ЖЕТСТАРКАУНТ ("OfficeDev", "Excel-Custom-functions")</strong> и нажмите клавишу ВВОД.</span><span class="sxs-lookup"><span data-stu-id="9ed03-171">In cell <strong>B1</strong>, type the text <strong>=CONTOSO.GETSTARCOUNT("OfficeDev", "Excel-Custom-Functions")</strong> and press enter.</span></span> <span data-ttu-id="9ed03-172">Вы увидите, что в ячейке <strong>B1</strong> получено текущее число звезд, заданное репозиторием [GitHub Excel-Custom-functions](https://github.com/OfficeDev/Excel-Custom-Functions).</span><span class="sxs-lookup"><span data-stu-id="9ed03-172">You should see that the result in cell <strong>B1</strong> is the current number of stars given to the [Excel-Custom-Functions Github repository](https://github.com/OfficeDev/Excel-Custom-Functions).</span></span></li>
</ol>

## <a name="create-a-streaming-asynchronous-custom-function"></a><span data-ttu-id="9ed03-173">Создание потоковой асинхронной пользовательской функции</span><span class="sxs-lookup"><span data-stu-id="9ed03-173">Create a streaming asynchronous custom function</span></span>

<span data-ttu-id="9ed03-174">`getStarCount` Функция возвращает число звезд, которые репозиторий содержит в определенный момент времени.</span><span class="sxs-lookup"><span data-stu-id="9ed03-174">The `getStarCount` function returns the number of stars a repository has at a specific moment in time.</span></span> <span data-ttu-id="9ed03-175">Пользовательские функции также могут возвращать непрерывно изменяемые данные.</span><span class="sxs-lookup"><span data-stu-id="9ed03-175">Custom functions can also return data that is continuously changing.</span></span> <span data-ttu-id="9ed03-176">Эти функции называются потоковыми функциями.</span><span class="sxs-lookup"><span data-stu-id="9ed03-176">These functions are called streaming functions.</span></span> <span data-ttu-id="9ed03-177">Они должны включать `invocation` параметр, который ссылается на ячейку, в которой была вызвана функция.</span><span class="sxs-lookup"><span data-stu-id="9ed03-177">They must include an `invocation` parameter which refers to the cell where the function was called from.</span></span> <span data-ttu-id="9ed03-178">`invocation` Параметр используется для обновления содержимого ячейки в любое время.</span><span class="sxs-lookup"><span data-stu-id="9ed03-178">The `invocation` parameter is used to update the contents of the cell at any time.</span></span>  

<span data-ttu-id="9ed03-179">В приведенном ниже примере кода обратите внимание, что существуют две функции `currentTime` и. `clock`</span><span class="sxs-lookup"><span data-stu-id="9ed03-179">In the following code sample, you'll notice that there are two functions, `currentTime` and `clock`.</span></span> <span data-ttu-id="9ed03-180">`currentTime` Функция — это статическая функция, которая не использует потоковую передачу.</span><span class="sxs-lookup"><span data-stu-id="9ed03-180">The `currentTime` function is a static function that does not use streaming.</span></span> <span data-ttu-id="9ed03-181">Он возвращает дату в виде строки.</span><span class="sxs-lookup"><span data-stu-id="9ed03-181">It returns the date as a string.</span></span> <span data-ttu-id="9ed03-182">`clock` Функция использует `currentTime` функцию, чтобы указать новое время каждую секунду для ячейки в Excel.</span><span class="sxs-lookup"><span data-stu-id="9ed03-182">The `clock` function uses the `currentTime` function to provide the new time every second to a cell in Excel.</span></span> <span data-ttu-id="9ed03-183">Он используется `invocation.setResult` для доставки времени в ячейку Excel и `invocation.onCanceled` обработки действий, выполняемых при отмене функции.</span><span class="sxs-lookup"><span data-stu-id="9ed03-183">It uses `invocation.setResult` to deliver the time to the Excel cell and `invocation.onCanceled` to handle what occurs when the function is canceled.</span></span>

1. <span data-ttu-id="9ed03-184">В проекте **старкаунт** добавьте следующий код в файл **./СРК/функтионс/функтионс.ЖС** и сохраните его.</span><span class="sxs-lookup"><span data-stu-id="9ed03-184">In the **starcount** project, add the following code to **./src/functions/functions.js** and save the file.</span></span>

```JS
/**
 * Returns the current time
 * @returns {string} String with the current time formatted for the current locale.
 */
function currentTime() {
  return new Date().toLocaleTimeString();
}

CustomFunctions.associate("CURRENTTIME", currentTime); 

 /**
 * Displays the current time once a second
 * @customfunction
 * @param {CustomFunctions.StreamingInvocation<string>} invocation Custom function invocation
 */
function clock(invocation) {
  const timer = setInterval(() => {
    const time = currentTime();
    invocation.setResult(time);
  }, 1000);

  invocation.onCanceled = () => {
    clearInterval(timer);
  };
}
CustomFunctions.associate("CLOCK", clock);
```

<span data-ttu-id="9ed03-185">Код `CustomFunctions.associate` сопоставляет `id` функции с адресом функции `CLOCK` в JavaScript, чтобы Excel мог вызвать вашу функцию.</span><span class="sxs-lookup"><span data-stu-id="9ed03-185">The `CustomFunctions.associate` code associates the `id` of the function with the function address of `CLOCK` in JavaScript so that Excel can call your function.</span></span>

2. <span data-ttu-id="9ed03-186">Выполните указанную ниже команду, чтобы повторно собрать проект.</span><span class="sxs-lookup"><span data-stu-id="9ed03-186">Run the following command to rebuild the project.</span></span>

    ```command&nbsp;line
    npm run build
    ```

3. <span data-ttu-id="9ed03-187">Выполните следующие действия (для Excel в Интернете или Windows), чтобы повторно зарегистрировать надстройку в Excel.</span><span class="sxs-lookup"><span data-stu-id="9ed03-187">Complete the following steps (for either Excel on the web or Windows) to re-register the add-in in Excel.</span></span> <span data-ttu-id="9ed03-188">Прежде чем новая функция станет доступна, необходимо выполнить указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="9ed03-188">You must complete these steps before the new function will be available.</span></span> 

# <a name="excel-on-windowstabexcel-windows"></a>[<span data-ttu-id="9ed03-189">Excel в Windows</span><span class="sxs-lookup"><span data-stu-id="9ed03-189">Excel on Windows</span></span>](#tab/excel-windows)

1. <span data-ttu-id="9ed03-190">Закройте Excel, а затем откройте Excel повторно.</span><span class="sxs-lookup"><span data-stu-id="9ed03-190">Close Excel and then reopen Excel.</span></span>

2. <span data-ttu-id="9ed03-191">В Excel перейдите на вкладку **Вставка** , а затем щелкните стрелку вниз, расположенную справа от **моих надстроек**.  ![Вставка ленты в Excel в Windows с выделенной стрелкой "Мои надстройки"](../images/select-insert.png)</span><span class="sxs-lookup"><span data-stu-id="9ed03-191">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Insert ribbon in Excel on Windows with the My Add-ins arrow highlighted](../images/select-insert.png)</span></span>

3. <span data-ttu-id="9ed03-192">В списке доступных надстроек найдите раздел надстройки для **разработчиков** и выберите надстройку **старкаунт** , чтобы зарегистрировать ее.</span><span class="sxs-lookup"><span data-stu-id="9ed03-192">In the list of available add-ins, find the **Developer Add-ins** section and select the **starcount** add-in to register it.</span></span>
    <span data-ttu-id="9ed03-193">![Вставка ленты в Excel в Windows с выделенной надстройкой "пользовательские функции Excel" в списке "Мои надстройки"](../images/list-starcount.png)</span><span class="sxs-lookup"><span data-stu-id="9ed03-193">![Insert ribbon in Excel on Windows with the Excel Custom Functions add-in highlighted in the My Add-ins list](../images/list-starcount.png)</span></span>

# <a name="excel-on-the-webtabexcel-online"></a>[<span data-ttu-id="9ed03-194">Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="9ed03-194">Excel on the web</span></span>](#tab/excel-online)

1. <span data-ttu-id="9ed03-195">В Excel перейдите на вкладку **Вставка** , а затем выберите \*\*\*\* пункт надстройки.  ![Вставка ленты в Excel в Интернете с выделенным значком "Мои надстройки"](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="9ed03-195">In Excel, choose the **Insert** tab and then choose **Add-ins**.  ![Insert ribbon in Excel on the web with the My Add-ins icon highlighted](../images/excel-cf-online-register-add-in-1.png)</span></span>

2. <span data-ttu-id="9ed03-196">Выберите пункт **Управление моими надстройками**, а затем выберите **Отправить мою надстройку**.</span><span class="sxs-lookup"><span data-stu-id="9ed03-196">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="9ed03-197">Выберите \*\*Обзор... \*\* и откройте корневой каталог проекта, созданный генератором Yeoman.</span><span class="sxs-lookup"><span data-stu-id="9ed03-197">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="9ed03-198">Выберите файл **manifest.xml** и нажмите **Открыть**, затем нажмите кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="9ed03-198">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

--- 

<ol start="4">
<li><span data-ttu-id="9ed03-199">Теперь давайте оценим, как работает новая функция.</span><span class="sxs-lookup"><span data-stu-id="9ed03-199">Try out the new function.</span></span> <span data-ttu-id="9ed03-200">В ячейке <strong>C1</strong>введите текст <strong>= contoso. CLOCK ())</strong> и нажмите клавишу ВВОД.</span><span class="sxs-lookup"><span data-stu-id="9ed03-200">In cell <strong>C1</strong>, type the text <strong>=CONTOSO.CLOCK())</strong> and press enter.</span></span> <span data-ttu-id="9ed03-201">Должна отобразиться текущая дата, которая пересылает обновление каждую секунду.</span><span class="sxs-lookup"><span data-stu-id="9ed03-201">You should see the current date, which streams an update every second.</span></span> <span data-ttu-id="9ed03-202">Несмотря на то, что часы находятся только в цикле, вы можете использовать ту же идею задания таймера для более сложных функций, которые делают веб-запросы для данных в режиме реального времени.</span><span class="sxs-lookup"><span data-stu-id="9ed03-202">While this clock is just a timer on a loop, you can use the same idea of setting a timer on more complex functions that make web requests for real-time data.</span></span></li>
</ol>

## <a name="next-steps"></a><span data-ttu-id="9ed03-203">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="9ed03-203">Next steps</span></span>

<span data-ttu-id="9ed03-204">Поздравляем!</span><span class="sxs-lookup"><span data-stu-id="9ed03-204">Congratulations!</span></span> <span data-ttu-id="9ed03-205">Вы создали новый проект пользовательских функций, выполнили предварительно составленную функцию, создал пользовательскую функцию, которая запрашивает данные из веб-сайта, и создала пользовательскую функцию, которая пересылает данные.</span><span class="sxs-lookup"><span data-stu-id="9ed03-205">You've created a new custom functions project, tried out a prebuilt function, created a custom function that requests data from the web, and created a custom function that streams data.</span></span> <span data-ttu-id="9ed03-206">Вы также можете попробовать выполнить отладку этой функции [, используя инструкции по отладке пользовательских функций](../excel/custom-functions-debugging.md).</span><span class="sxs-lookup"><span data-stu-id="9ed03-206">You can also try out debugging this function using [the custom function debugging instructions](../excel/custom-functions-debugging.md).</span></span> <span data-ttu-id="9ed03-207">Чтобы узнать больше о пользовательских функции в Excel, перейдите к следующей статье:</span><span class="sxs-lookup"><span data-stu-id="9ed03-207">To learn more about custom functions in Excel, continue to the following article:</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="9ed03-208">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="9ed03-208">Create custom functions in Excel</span></span>](../excel/custom-functions-overview.md)