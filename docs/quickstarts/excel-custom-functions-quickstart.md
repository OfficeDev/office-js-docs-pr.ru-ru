---
ms.date: 11/09/2020
description: Краткое руководство по разработке пользовательских функций в Excel.
title: Краткое руководство по пользовательским функциям
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 0b0e42149e771978026db3eb84594bd172d09459
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076632"
---
# <a name="get-started-developing-excel-custom-functions"></a><span data-ttu-id="ebe5e-103">Начало разработки пользовательских функций Excel</span><span class="sxs-lookup"><span data-stu-id="ebe5e-103">Get started developing Excel custom functions</span></span>

<span data-ttu-id="ebe5e-104">С помощью пользовательских функций разработчики могут добавлять новые функции в Excel путем их определения в JavaScript или Typescript как части надстройки.</span><span class="sxs-lookup"><span data-stu-id="ebe5e-104">With custom functions, developers can now add new functions to Excel by defining them in JavaScript or Typescript as part of an add-in.</span></span> <span data-ttu-id="ebe5e-105">Пользователи Excel могут получить доступ к пользовательским функциям так же, как и к любой встроенной функции в Excel, например `SUM()`.</span><span class="sxs-lookup"><span data-stu-id="ebe5e-105">Excel users can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="ebe5e-106">Предварительные требования</span><span class="sxs-lookup"><span data-stu-id="ebe5e-106">Prerequisites</span></span>

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

* <span data-ttu-id="ebe5e-107">Excel для Windows (версия 1904 или более поздняя, подключенная к подписке на Microsoft 365) или Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="ebe5e-107">Excel on Windows (version 1904 or later, connected to a Microsoft 365 subscription) or Excel on the web</span></span>
* <span data-ttu-id="ebe5e-108">Пользовательские функции Excel поддерживаются в Office для Mac (подключенном к подписке на Microsoft 365), и ожидается обновление данного руководства.</span><span class="sxs-lookup"><span data-stu-id="ebe5e-108">Excel custom functions are supported in Office on Mac (connected to a Microsoft 365 subscription) and an update to this tutorial is forthcoming.</span></span>

>[!NOTE]
><span data-ttu-id="ebe5e-109">Пользовательские функции Excel не поддерживаются в Office 2019 (единовременная покупка).</span><span class="sxs-lookup"><span data-stu-id="ebe5e-109">Excel custom functions are not supported in Office 2019 (one-time purchase).</span></span>

## <a name="build-your-first-custom-functions-project"></a><span data-ttu-id="ebe5e-110">Создание первого проекта пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="ebe5e-110">Build your first custom functions project</span></span>

<span data-ttu-id="ebe5e-111">Чтобы начать работу, создайте проект пользовательских функций с помощью генератора Yeoman.</span><span class="sxs-lookup"><span data-stu-id="ebe5e-111">To start, you'll use the Yeoman generator to create the custom functions project.</span></span> <span data-ttu-id="ebe5e-112">Это позволит настроить для проекта правильную структуру папок, исходные файлы и зависимости, чтобы начать написание кода пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="ebe5e-112">This will set up your project with the correct folder structure, source files, and dependencies to begin coding your custom functions.</span></span>

1. [!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

    - <span data-ttu-id="ebe5e-113">**Выберите тип проекта:** `Excel Custom Functions Add-in project`</span><span class="sxs-lookup"><span data-stu-id="ebe5e-113">**Choose a project type:** `Excel Custom Functions Add-in project`</span></span>
    - <span data-ttu-id="ebe5e-114">**Выберите тип сценария:** `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="ebe5e-114">**Choose a script type:** `JavaScript`</span></span>
    - <span data-ttu-id="ebe5e-115">**Как вы хотите назвать надстройку?**</span><span class="sxs-lookup"><span data-stu-id="ebe5e-115">**What do you want to name your add-in?**</span></span> `starcount`

    ![Снимок экрана: интерфейс командной строки генератора Yeoman надстроек Office, запрашивающий проекты пользовательских функций.](../images/starcountPrompt.png)

    <span data-ttu-id="ebe5e-117">Генератор Yeoman создаст файлы проекта и установит вспомогательные компоненты Node.</span><span class="sxs-lookup"><span data-stu-id="ebe5e-117">The Yeoman generator will create the project files and install supporting Node components.</span></span>

2. <span data-ttu-id="ebe5e-p103">Генератор Yeoman предоставит вам инструкции в командной строке по действиям с проектом, но вам нужно их проигнорировать и продолжить выполнять наши инструкции. Перейдите к корневой папке проекта.</span><span class="sxs-lookup"><span data-stu-id="ebe5e-p103">The Yeoman generator will give you some instructions in your command line about what to do with the project, but ignore them and continue to follow our instructions. Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd starcount
    ```

3. <span data-ttu-id="ebe5e-120">Выполните построение проекта.</span><span class="sxs-lookup"><span data-stu-id="ebe5e-120">Build the project.</span></span> 

    ```command&nbsp;line
    npm run build
    ```

    > [!NOTE]
    > <span data-ttu-id="ebe5e-121">Надстройки Office должны использовать HTTPS, а не HTTP, даже в случае разработки.</span><span class="sxs-lookup"><span data-stu-id="ebe5e-121">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="ebe5e-122">Если вам будет предложено установить сертификат после того, как вы запустите `npm run build`, примите предложение установить сертификат от генератора Yeoman.</span><span class="sxs-lookup"><span data-stu-id="ebe5e-122">If you are prompted to install a certificate after you run `npm run build`, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

4. <span data-ttu-id="ebe5e-123">Запустите локальный веб-сервер, работающий на Node.js.</span><span class="sxs-lookup"><span data-stu-id="ebe5e-123">Start the local web server, which runs in Node.js.</span></span> <span data-ttu-id="ebe5e-124">Вы можете попробовать использовать надстройку пользовательской функции в Excel в Интернете или для Windows.</span><span class="sxs-lookup"><span data-stu-id="ebe5e-124">You can try out the custom function add-in in Excel on the web or Windows.</span></span> <span data-ttu-id="ebe5e-125">Вам может быть предложено открыть область задач надстройки, но это необязательно.</span><span class="sxs-lookup"><span data-stu-id="ebe5e-125">You may be prompted to open the add-in's task pane, although this is optional.</span></span> <span data-ttu-id="ebe5e-126">Вы по-прежнему можете запускать свои пользовательские функции, не открывая область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="ebe5e-126">You can still run your custom functions without opening your add-in's task pane.</span></span>

# <a name="excel-on-windows"></a>[<span data-ttu-id="ebe5e-127">Excel для Windows</span><span class="sxs-lookup"><span data-stu-id="ebe5e-127">Excel on Windows</span></span>](#tab/excel-windows)

<span data-ttu-id="ebe5e-128">Чтобы проверить надстройку в Excel для Windows, выполните следующую команду.</span><span class="sxs-lookup"><span data-stu-id="ebe5e-128">To test your add-in in Excel on Windows, run the following command.</span></span> <span data-ttu-id="ebe5e-129">Когда вы выполните эту команду, запустится локальный веб-сервер и откроется приложение Excel, в котором будет загружена ваша надстройка.</span><span class="sxs-lookup"><span data-stu-id="ebe5e-129">When you run this command, the local web server will start and Excel will open with your add-in loaded.</span></span>

```command&nbsp;line
npm run start:desktop
```

# <a name="excel-on-the-web"></a>[<span data-ttu-id="ebe5e-130">Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="ebe5e-130">Excel on the web</span></span>](#tab/excel-online)

<span data-ttu-id="ebe5e-131">Чтобы проверить надстройку в Excel в Интернете, выполните следующую команду.</span><span class="sxs-lookup"><span data-stu-id="ebe5e-131">To test your add-in in Excel on the web, run the following command.</span></span> <span data-ttu-id="ebe5e-132">После выполнения этой команды запустится локальный веб-сервер.</span><span class="sxs-lookup"><span data-stu-id="ebe5e-132">When you run this command, the local web server will start.</span></span>

```command&nbsp;line
npm run start:web
```

<span data-ttu-id="ebe5e-133">Чтобы использовать надстройку пользовательских функций, откройте новую книгу Excel в браузере.</span><span class="sxs-lookup"><span data-stu-id="ebe5e-133">To use your custom functions add-in, open a new workbook in Excel on a browser.</span></span> <span data-ttu-id="ebe5e-134">В этой книге выполните шаги, описанные ниже, чтобы загрузить неопубликованную надстройку.</span><span class="sxs-lookup"><span data-stu-id="ebe5e-134">In this workbook, complete the following steps to sideload your add-in.</span></span>

1. <span data-ttu-id="ebe5e-135">В Excel на вкладке **Вставка** выберите пункт **Надстройки**.</span><span class="sxs-lookup"><span data-stu-id="ebe5e-135">In Excel, choose the **Insert** tab and then choose **Add-ins**.</span></span>

   ![Снимок экрана: лента "Вставка" в Excel в Интернете с выделенной кнопкой "Мои надстройки".](../images/excel-cf-online-register-add-in-1.png)
   
2. <span data-ttu-id="ebe5e-137">Выберите пункт **Управление моими надстройками**, а затем выберите **Отправить мою надстройку**.</span><span class="sxs-lookup"><span data-stu-id="ebe5e-137">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="ebe5e-138">Выберите **Обзор...** и откройте корневой каталог проекта, созданный генератором Yeoman.</span><span class="sxs-lookup"><span data-stu-id="ebe5e-138">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="ebe5e-139">Выберите файл **manifest.xml** и нажмите **Открыть**, затем нажмите кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="ebe5e-139">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

---

## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="ebe5e-140">Проверка работы готовой пользовательской функции</span><span class="sxs-lookup"><span data-stu-id="ebe5e-140">Try out a prebuilt custom function</span></span>

<span data-ttu-id="ebe5e-141">Проект пользовательских функций, созданный с помощью генератора Yeoman, содержит некоторые готовые пользовательские функции, определенные в файле **./src/functions/functions.js**.</span><span class="sxs-lookup"><span data-stu-id="ebe5e-141">The custom functions project that you created by using the Yeoman generator contains some prebuilt custom functions, defined within the **./src/functions/functions.js** file.</span></span> <span data-ttu-id="ebe5e-142">Файл **./manifest.xml** в корневом каталоге проекта указывает, что все пользовательские функции принадлежат пространству имен `CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="ebe5e-142">The **./manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.</span></span>

<span data-ttu-id="ebe5e-143">В книге Excel проверьте, как работает пользовательская функция `ADD`, выполнив описанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="ebe5e-143">In your Excel workbook, try out the `ADD` custom function by completing the following steps:</span></span>

1. <span data-ttu-id="ebe5e-144">Выделите ячейку и введите `=CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="ebe5e-144">Select a cell and type `=CONTOSO`.</span></span> <span data-ttu-id="ebe5e-145">Обратите внимание, что в меню автозаполнения содержится список всех функций в пространстве имен `CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="ebe5e-145">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="ebe5e-146">Запустите функцию `CONTOSO.ADD` с числами `10` и `200` в качестве входных параметров, введя значение `=CONTOSO.ADD(10,200)` в ячейке и нажав клавишу ВВОД.</span><span class="sxs-lookup"><span data-stu-id="ebe5e-146">Run the `CONTOSO.ADD` function, using numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.</span></span>

<span data-ttu-id="ebe5e-147">Пользовательская функция `ADD` вычисляет сумму двух чисел, которые вы указываете в качестве входных параметров.</span><span class="sxs-lookup"><span data-stu-id="ebe5e-147">The `ADD` custom function computes the sum of the two numbers that you specify as input parameters.</span></span> <span data-ttu-id="ebe5e-148">При вводе `=CONTOSO.ADD(10,200)` в ячейке должен отобразиться результат **210** после нажатия клавиши ВВОД.</span><span class="sxs-lookup"><span data-stu-id="ebe5e-148">Typing `=CONTOSO.ADD(10,200)` should produce the result **210** in the cell after you press enter.</span></span>

## <a name="next-steps"></a><span data-ttu-id="ebe5e-149">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="ebe5e-149">Next steps</span></span>

<span data-ttu-id="ebe5e-150">Поздравляем, вы успешно создали пользовательскую функцию в надстройке Excel!</span><span class="sxs-lookup"><span data-stu-id="ebe5e-150">Congratulations, you've successfully created a custom function in an Excel add-in!</span></span> <span data-ttu-id="ebe5e-151">Затем создайте более сложную надстройку с возможностью потоковой передачи данных.</span><span class="sxs-lookup"><span data-stu-id="ebe5e-151">Next, build a more complex add-in with streaming data capability.</span></span> <span data-ttu-id="ebe5e-152">Следующая ссылка поможет вам выполнить дальнейшие действия в руководстве по надстройке Excel с пользовательскими функциями.</span><span class="sxs-lookup"><span data-stu-id="ebe5e-152">The following link takes you through the next steps in the Excel add-in with custom functions tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="ebe5e-153">Руководство по надстройке Excel для пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="ebe5e-153">Excel custom functions add-in tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md#create-a-custom-function-that-requests-data-from-the-web
)

## <a name="see-also"></a><span data-ttu-id="ebe5e-154">См. также</span><span class="sxs-lookup"><span data-stu-id="ebe5e-154">See also</span></span>

* [<span data-ttu-id="ebe5e-155">Обзор пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="ebe5e-155">Custom functions overview</span></span>](../excel/custom-functions-overview.md)
* [<span data-ttu-id="ebe5e-156">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="ebe5e-156">Custom functions metadata</span></span>](../excel/custom-functions-json.md)
* [<span data-ttu-id="ebe5e-157">Среда выполнения для пользовательских функций Excel</span><span class="sxs-lookup"><span data-stu-id="ebe5e-157">Runtime for Excel custom functions</span></span>](../excel/custom-functions-runtime.md)