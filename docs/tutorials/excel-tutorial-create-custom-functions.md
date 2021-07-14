---
title: Руководство по пользовательским функциям в Excel
description: В этом руководстве вы создадите надстройку Excel, содержащую пользовательскую функцию, которая может выполнять вычисления, запрашивать веб-данные или потоковые веб-данные.
ms.date: 07/07/2021
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 56e8a31f8d33756ca1668e2fa1468d10d1ad4821
ms.sourcegitcommit: 95fc1fc8a0dbe8fc94f0ea647836b51cc7f8601d
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/14/2021
ms.locfileid: "53418715"
---
# <a name="tutorial-create-custom-functions-in-excel"></a><span data-ttu-id="7bc5d-103">Руководство: создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="7bc5d-103">Tutorial: Create custom functions in Excel</span></span>

<span data-ttu-id="7bc5d-104">Пользовательские функции позволяют добавлять новые функции в Excel путем определения этих функций в JavaScript как части надстройки.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-104">Custom functions enable you to add new functions to Excel by defining those functions in JavaScript as part of an add-in.</span></span> <span data-ttu-id="7bc5d-105">Пользователи в Excel могут получить доступ к пользовательским функциям так же, как и к любой встроенной функции в Excel, например `SUM()`.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-105">Users within Excel can access custom functions as they would any native function in Excel, such as `SUM()`.</span></span> <span data-ttu-id="7bc5d-106">Вы можете создавать пользовательские функции, которые будут выполнять простые задачи, такие как вычисления, или более сложные задачи, такие как потоковая передача данных в режиме реального времени из Интернета на лист.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-106">You can create custom functions that perform simple tasks like calculations or more complex tasks such as streaming real-time data from the web into a worksheet.</span></span>

<span data-ttu-id="7bc5d-107">В этом руководстве описан порядок выполнения перечисленных ниже задач.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-107">In this tutorial, you will:</span></span>
> [!div class="checklist"]
> - <span data-ttu-id="7bc5d-108">Создание надстройки пользовательской функции с помощью [генератора Yeoman для надстроек Office](https://www.npmjs.com/package/generator-office).</span><span class="sxs-lookup"><span data-stu-id="7bc5d-108">Create a custom function add-in using the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office).</span></span> 
> - <span data-ttu-id="7bc5d-109">Использование готовой пользовательской функции для выполнения простых вычислений</span><span class="sxs-lookup"><span data-stu-id="7bc5d-109">Use a prebuilt custom function to perform a simple calculation.</span></span>
> - <span data-ttu-id="7bc5d-110">Создание пользовательской функции, которая получает данные из сети Интернет.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-110">Create a custom function that gets data from the web.</span></span>
> - <span data-ttu-id="7bc5d-111">Создание пользовательской функции, которая осуществляет потоковую передачу данных в реальном времени из сети Интернет</span><span class="sxs-lookup"><span data-stu-id="7bc5d-111">Create a custom function that streams real-time data from the web.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="7bc5d-112">Предварительные требования</span><span class="sxs-lookup"><span data-stu-id="7bc5d-112">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- <span data-ttu-id="7bc5d-113">Excel для Windows (версия 1904 или более поздняя, подключенная к подписке на Microsoft 365) или в Интернете</span><span class="sxs-lookup"><span data-stu-id="7bc5d-113">Excel on Windows (version 1904 or later, connected to a Microsoft 365 subscription) or on the web</span></span>

## <a name="create-a-custom-functions-project"></a><span data-ttu-id="7bc5d-114">Создание проекта пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="7bc5d-114">Create a custom functions project</span></span>

 <span data-ttu-id="7bc5d-115">Чтобы начать, вам необходимо создать проект кода для разработки надстройки пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-115">To start, you'll create the code project to build your custom function add-in.</span></span> <span data-ttu-id="7bc5d-116">[Генератор Yeoman для надстроек Office](https://www.npmjs.com/package/generator-office) настроит в вашем проекте некоторые готовые пользовательские функции, которые можно использовать. Если вы уже запустили «Краткое руководство по пользовательским функциям» и создали проект, то продолжайте этот проект и [пропустите эту процедуру](#create-a-custom-function-that-requests-data-from-the-web).</span><span class="sxs-lookup"><span data-stu-id="7bc5d-116">The [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office) will set up your project with some prebuilt custom functions that you can try out. If you have already run the custom functions quick start and generated a project, continue to use that project and skip to [this step](#create-a-custom-function-that-requests-data-from-the-web) instead.</span></span>

1. [!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

    - <span data-ttu-id="7bc5d-117">**Выберите тип проекта:** `Excel Custom Functions Add-in project`</span><span class="sxs-lookup"><span data-stu-id="7bc5d-117">**Choose a project type:** `Excel Custom Functions Add-in project`</span></span>
    - <span data-ttu-id="7bc5d-118">**Выберите тип сценария:** `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="7bc5d-118">**Choose a script type:** `JavaScript`</span></span>
    - <span data-ttu-id="7bc5d-119">**Как вы хотите назвать надстройку?**</span><span class="sxs-lookup"><span data-stu-id="7bc5d-119">**What do you want to name your add-in?**</span></span> `starcount`

    ![Снимок экрана: интерфейс командной строки генератора Yeoman надстроек Office, запрашивающий проекты пользовательских функций.](../images/starcountPrompt.png)

    <span data-ttu-id="7bc5d-121">Генератор Yeoman создаст файлы проекта и установит вспомогательные компоненты Node.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-121">The Yeoman generator will create the project files and install supporting Node components.</span></span>

    [!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

1. <span data-ttu-id="7bc5d-122">Перейдите к корневой папке проекта.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-122">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd starcount
    ```

1. <span data-ttu-id="7bc5d-123">Выполните построение проекта.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-123">Build the project.</span></span>

    ```command&nbsp;line
    npm run build
    ```

    > [!NOTE]
    > <span data-ttu-id="7bc5d-124">Надстройки Office должны использовать HTTPS, а не HTTP, даже в случае разработки.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-124">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="7bc5d-125">Если вам будет предложено установить сертификат после того, как вы запустите `npm run build`, примите предложение установить сертификат от генератора Yeoman.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-125">If you are prompted to install a certificate after you run `npm run build`, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

1. <span data-ttu-id="7bc5d-126">Запустите локальный веб-сервер, работающий на Node.js.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-126">Start the local web server, which runs in Node.js.</span></span> <span data-ttu-id="7bc5d-127">Вы можете попробовать использовать надстройку пользовательской функции в Excel в Интернете или для Windows.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-127">You can try out the custom function add-in in Excel on the web or Windows.</span></span>

# <a name="excel-on-windows-or-mac"></a>[<span data-ttu-id="7bc5d-128">Excel для Windows или Mac</span><span class="sxs-lookup"><span data-stu-id="7bc5d-128">Excel on Windows or Mac</span></span>](#tab/excel-windows)

<span data-ttu-id="7bc5d-129">Чтобы проверить надстройку в Excel для Windows или Mac, выполните следующую команду.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-129">To test your add-in in Excel on Windows or Mac, run the following command.</span></span> <span data-ttu-id="7bc5d-130">Когда вы выполните эту команду, запустится локальный веб-сервер и откроется приложение Excel, в котором будет загружена ваша надстройка.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-130">When you run this command, the local web server will start and Excel will open with your add-in loaded.</span></span>

```command&nbsp;line
npm run start:desktop
```

# <a name="excel-on-the-web"></a>[<span data-ttu-id="7bc5d-131">Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="7bc5d-131">Excel on the web</span></span>](#tab/excel-online)

<span data-ttu-id="7bc5d-132">Чтобы проверить надстройку в Excel в браузере, выполните следующую команду.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-132">To test your add-in in Excel on a browser, run the following command.</span></span> <span data-ttu-id="7bc5d-133">После выполнения этой команды запустится локальный веб-сервер.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-133">When you run this command, the local web server will start.</span></span>

```command&nbsp;line
npm run start:web
```

<span data-ttu-id="7bc5d-134">Чтобы использовать надстройку пользовательских функций, откройте новую книгу в Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-134">To use your custom functions add-in, open a new workbook in Excel on the web.</span></span> <span data-ttu-id="7bc5d-135">В этой книге выполните шаги, описанные ниже, чтобы загрузить неопубликованную надстройку.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-135">In this workbook, complete the following steps to sideload your add-in.</span></span>

1. <span data-ttu-id="7bc5d-136">В Excel на вкладке **Вставка** выберите пункт **Надстройки**.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-136">In Excel, choose the **Insert** tab and then choose **Add-ins**.</span></span>

   ![Снимок экрана: лента "Вставка" в Excel в Интернете с выделенной кнопкой "Мои надстройки".](../images/excel-cf-online-register-add-in-1.png)

1. <span data-ttu-id="7bc5d-138">Выберите пункт **Управление моими надстройками**, а затем выберите **Отправить мою надстройку**.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-138">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

1. <span data-ttu-id="7bc5d-139">Выберите **Обзор...** и откройте корневой каталог проекта, созданный генератором Yeoman.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-139">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

1. <span data-ttu-id="7bc5d-140">Выберите файл **manifest.xml** и нажмите **Открыть**, затем нажмите кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-140">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

---

## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="7bc5d-141">Проверка работы готовой пользовательской функции</span><span class="sxs-lookup"><span data-stu-id="7bc5d-141">Try out a prebuilt custom function</span></span>

<span data-ttu-id="7bc5d-142">Созданный проект пользовательских функций содержит некоторые готовые пользовательские функции, определенные в файле **src/functions/functions.js**.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-142">The custom functions project that you created contains some prebuilt custom functions, defined within the **./src/functions/functions.js** file.</span></span> <span data-ttu-id="7bc5d-143">Файл **./manifest.xml** указывает, что все пользовательские функции принадлежат пространству имен `CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-143">The **./manifest.xml** file specifies that all custom functions belong to the `CONTOSO` namespace.</span></span> <span data-ttu-id="7bc5d-144">Вы будете использовать пространство имен CONTOSO для доступа к пользовательским функциям в Excel.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-144">You'll use the CONTOSO namespace to access the custom functions in Excel.</span></span>

<span data-ttu-id="7bc5d-145">Затем вы проверите пользовательскую функцию `ADD`, выполнив действия ниже.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-145">Next you'll try out the `ADD` custom function by completing the following steps.</span></span>

1. <span data-ttu-id="7bc5d-146">В Excel перейдите в любую ячейку и введите `=CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-146">In Excel, go to any cell and enter `=CONTOSO`.</span></span> <span data-ttu-id="7bc5d-147">Обратите внимание на то, что в меню автозаполнения содержится список всех функций в пространстве имен `CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-147">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

1. <span data-ttu-id="7bc5d-148">Выполните запуск функции `CONTOSO.ADD` с числами `10` и `200` в качестве входных параметров, введя значение `=CONTOSO.ADD(10,200)` в ячейке и нажав клавишу ВВОД.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-148">Run the `CONTOSO.ADD` function, with numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.</span></span>

<span data-ttu-id="7bc5d-149">Пользовательская функция `ADD` вычисляет сумму двух чисел, которые вы указываете и возвращает результат **210**.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-149">The `ADD` custom function computes the sum of the two numbers that you provided and returns the result of **210**.</span></span>

## <a name="create-a-custom-function-that-requests-data-from-the-web"></a><span data-ttu-id="7bc5d-150">Создание пользовательской функции, которая запрашивает данные из сети Интернет</span><span class="sxs-lookup"><span data-stu-id="7bc5d-150">Create a custom function that requests data from the web</span></span>

<span data-ttu-id="7bc5d-151">Интеграция данных из Интернета — отличный способ расширения функционала Excel через пользовательские функции.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-151">Integrating data from the Web is a great way to extend Excel through custom functions.</span></span> <span data-ttu-id="7bc5d-152">Затем вы создадите пользовательскую функцию с именем `getStarCount`, показывающую, сколько звезд имеет данный репозиторий Github.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-152">Next you'll create a custom function named `getStarCount` that shows how many stars a given Github repository possesses.</span></span>

1. <span data-ttu-id="7bc5d-153">В проекте **starcount** найдите файл **./src/functions/functions.js** и откройте его в редакторе кода.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-153">In the **starcount** project, find the file **./src/functions/functions.js** and open it in your code editor.</span></span>

1. <span data-ttu-id="7bc5d-154">В **function.js** добавьте следующий код.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-154">In **function.js**, add the following code.</span></span>

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
    ```

1. <span data-ttu-id="7bc5d-155">Выполните указанную ниже команду, чтобы повторно собрать проект.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-155">Run the following command to rebuild the project.</span></span>

    ```command&nbsp;line
    npm run build
    ```

1. <span data-ttu-id="7bc5d-156">Чтобы повторно зарегистрировать надстройку в Excel, выполните указанные ниже действия (для Excel в Интернете, для Windows или для Mac).</span><span class="sxs-lookup"><span data-stu-id="7bc5d-156">Complete the following steps (for Excel on the web, Windows, or Mac) to re-register the add-in in Excel.</span></span> <span data-ttu-id="7bc5d-157">Выполните описанные ниже действия, чтобы новая функция стала доступной.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-157">You must complete these steps before the new function will be available.</span></span>

### <a name="excel-on-windows-or-mac"></a>[<span data-ttu-id="7bc5d-158">Excel для Windows или Mac</span><span class="sxs-lookup"><span data-stu-id="7bc5d-158">Excel on Windows or Mac</span></span>](#tab/excel-windows)

1. <span data-ttu-id="7bc5d-159">Закройте Excel, а затем откройте Excel повторно.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-159">Close Excel and then reopen Excel.</span></span>

1. <span data-ttu-id="7bc5d-160">В Excel выберите вкладку **Вставка**, а затем нажмите стрелку вниз, которая находится справа от пункта **Мои надстройки**. ![Снимок экрана: лента "Вставка" Excel для Windows с выделенной стрелкой вниз "Мои надстройки".](../images/select-insert.png)</span><span class="sxs-lookup"><span data-stu-id="7bc5d-160">In Excel, choose the **Insert** tab and then choose the down-arrow located to the right of **My Add-ins**.  ![Screenshot of the Insert ribbon in Excel on Windows, with the My Add-ins down-arrow highlighted.](../images/select-insert.png)</span></span>

1. <span data-ttu-id="7bc5d-161">В списке доступных надстроек найдите раздел **Надстройки разработчика** и выберите вашу надстройку **starcount**, чтобы ее зарегистрировать.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-161">In the list of available add-ins, find the **Developer Add-ins** section and select the **starcount** add-in to register it.</span></span>
    <span data-ttu-id="7bc5d-162">![Снимок экрана: лента "Вставка" в Excel для Windows с выделенной надстройкой "Пользовательские функции Excel" в списке "Мои надстройки".](../images/list-starcount.png)</span><span class="sxs-lookup"><span data-stu-id="7bc5d-162">![Screenshot of the Insert ribbon in Excel on Windows, with the Excel Custom Functions add-in highlighted in the My Add-ins list.](../images/list-starcount.png)</span></span>


# <a name="excel-on-the-web"></a>[<span data-ttu-id="7bc5d-163">Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="7bc5d-163">Excel on the web</span></span>](#tab/excel-online)

1. <span data-ttu-id="7bc5d-164">В Excel выберите вкладку **Вставка**, а затем выберите **Надстройки**. ![Снимок экрана: лента "Вставка" Excel для Интернета с выделенной кнопкой "Мои надстройки".](../images/excel-cf-online-register-add-in-1.png)</span><span class="sxs-lookup"><span data-stu-id="7bc5d-164">In Excel, choose the **Insert** tab and then choose **Add-ins**.  ![Screenshot of the Insert ribbon in Excel on the web, with the My Add-ins button highlighted.](../images/excel-cf-online-register-add-in-1.png)</span></span>

1. <span data-ttu-id="7bc5d-165&quot;>Выберите пункт **Управление моими надстройками**, а затем выберите **Отправить мою надстройку**.</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;7bc5d-165&quot;>Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

1. <span data-ttu-id=&quot;7bc5d-166&quot;>Выберите **Обзор...** и откройте корневой каталог проекта, созданный генератором Yeoman.</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;7bc5d-166&quot;>Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

1. <span data-ttu-id=&quot;7bc5d-167&quot;>Выберите файл **manifest.xml** и нажмите **Открыть**, затем нажмите кнопку **Отправить**.</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;7bc5d-167&quot;>Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

---

<ol start=&quot;5&quot;>
<li> <span data-ttu-id=&quot;7bc5d-168&quot;>Теперь давайте оценим, как работает новая функция.</span><span class=&quot;sxs-lookup&quot;><span data-stu-id=&quot;7bc5d-168&quot;>Try out the new function.</span></span> <span data-ttu-id=&quot;7bc5d-169&quot;>В ячейке <strong>B1</strong>введите текст <strong>=CONTOSO.GETSTARCOUNT(&quot;OfficeDev&quot;, &quot;Excel-Custom-Functions")</strong> и нажмите клавишу ВВОД.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-169">In cell <strong>B1</strong>, type the text <strong>=CONTOSO.GETSTARCOUNT("OfficeDev", "Excel-Custom-Functions")</strong> and press enter.</span></span> <span data-ttu-id="7bc5d-170">Результат в ячейке <strong>B1</strong> — это текущее количество звезд, отданных репозиторию [Excel-Custom-Functions Github](https://github.com/OfficeDev/Excel-Custom-Functions).</span><span class="sxs-lookup"><span data-stu-id="7bc5d-170">You should see that the result in cell <strong>B1</strong> is the current number of stars given to the [Excel-Custom-Functions Github repository](https://github.com/OfficeDev/Excel-Custom-Functions).</span></span></li>
</ol>

## <a name="create-a-streaming-asynchronous-custom-function"></a><span data-ttu-id="7bc5d-171">Создание потоковой асинхронной пользовательской функции</span><span class="sxs-lookup"><span data-stu-id="7bc5d-171">Create a streaming asynchronous custom function</span></span>

<span data-ttu-id="7bc5d-172">Функция `getStarCount` возвращает количество звезд, которые есть у репозитория в определенный момент времени.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-172">The `getStarCount` function returns the number of stars a repository has at a specific moment in time.</span></span> <span data-ttu-id="7bc5d-173">Пользовательские функции также возвращают непрерывно изменяемые данные.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-173">Custom functions also return data that is continuously changing.</span></span> <span data-ttu-id="7bc5d-174">Эти функции называются потоковыми передачами функций.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-174">These functions are called streaming functions.</span></span> <span data-ttu-id="7bc5d-175">Они должны содержать параметр `invocation`, ссылающийся на ячейку, из которой была вызвана функция.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-175">They must include an `invocation` parameter which refers to the cell that called the function.</span></span> <span data-ttu-id="7bc5d-176">Параметр `invocation` используется для обновления содержимого ячейки в любое время.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-176">The `invocation` parameter is used to update the contents of the cell at any time.</span></span>  

<span data-ttu-id="7bc5d-177">В примере кода ниже вы заметите наличие двух функций, `currentTime` и `clock`.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-177">In the following code sample, notice that there are two functions, `currentTime` and `clock`.</span></span> <span data-ttu-id="7bc5d-178">Функция `currentTime` — это статическая функция, которая не использует потоковую передачу функций.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-178">The `currentTime` function is a static function that doesn't use streaming.</span></span> <span data-ttu-id="7bc5d-179">Она возвращает дату в виде строки.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-179">It returns the date as a string.</span></span> <span data-ttu-id="7bc5d-180">Функция `clock` использует функцию `currentTime` для обеспечения нового времени каждую секунду для ячейки в Excel.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-180">The `clock` function uses the `currentTime` function to provide the new time every second to a cell in Excel.</span></span> <span data-ttu-id="7bc5d-181">В ней используется `invocation.setResult` для передачи времени в ячейку Excel и `invocation.onCanceled` для обработки отмены функции.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-181">It uses `invocation.setResult` to deliver the time to the Excel cell and `invocation.onCanceled` to handle function cancellation.</span></span> 

<span data-ttu-id="7bc5d-182">Проект **starcount** уже содержит две следующие функции в файле **./src/functions/functions.js**.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-182">The **starcount** project already contains the following two functions in the **./src/functions/functions.js** file.</span></span>

```JS
/**
 * Returns the current time
 * @returns {string} String with the current time formatted for the current locale.
 */
function currentTime() {
  return new Date().toLocaleTimeString();
}
    
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
```

<span data-ttu-id="7bc5d-183">Чтобы опробовать функции, введите текст **=CONTOSO.CLOCK()** в ячейку **C1** и нажмите ВВОД.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-183">To try out the functions, type the text **=CONTOSO.CLOCK()** in cell **C1** and press enter.</span></span> <span data-ttu-id="7bc5d-184">Должна отобразиться текущая дата, которая потоком обновляется каждую секунду.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-184">You should see the current date, which streams an update every second.</span></span> <span data-ttu-id="7bc5d-185">Хотя эти часы являются просто таймером в цикле, однако можно использовать аналогичную идею настройки таймера для более сложных функций, которые выполняют веб-запросы в режиме реального времени.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-185">While this clock is just a timer on a loop, you can use the same idea of setting a timer on more complex functions that make web requests for real-time data.</span></span>

## <a name="next-steps"></a><span data-ttu-id="7bc5d-186">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="7bc5d-186">Next steps</span></span>

<span data-ttu-id="7bc5d-187">Поздравляем!</span><span class="sxs-lookup"><span data-stu-id="7bc5d-187">Congratulations!</span></span> <span data-ttu-id="7bc5d-188">Вы создали новый проект пользовательских функций, попробовали, как работает готовая функция, создали пользовательскую функцию, которая запрашивает данные из Интернета, а также создали пользовательскую функцию, которая осуществляет потоковую передачу данных.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-188">You've created a new custom functions project, tried out a prebuilt function, created a custom function that requests data from the web, and created a custom function that streams data.</span></span> <span data-ttu-id="7bc5d-189">Затем вы можете изменить свой проект, чтобы использовать общую среду выполнения, упрощая взаимодействие с панелью задач.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-189">Next, you can modify your project to use a shared runtime, making it easier for your function to interact with the task pane.</span></span> <span data-ttu-id="7bc5d-190">Выполните инструкции из следующей статьи.</span><span class="sxs-lookup"><span data-stu-id="7bc5d-190">Follow the steps in the following article.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="7bc5d-191">Настройка надстройки для использования общей среды выполнения</span><span class="sxs-lookup"><span data-stu-id="7bc5d-191">Configure your add-in to use a shared runtime</span></span>](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
