---
ms.date: 09/18/2019
description: Разработка настраиваемых функций в кратком руководстве по Excel.
title: Краткое руководство по настраиваемым функциям
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: f34a8817a7c8ef2679fc8ce0a6ad17cec600531b
ms.sourcegitcommit: a0257feabcfe665061c14b8bdb70cf82f7aca414
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/18/2019
ms.locfileid: "37035331"
---
# <a name="get-started-developing-excel-custom-functions"></a><span data-ttu-id="e1720-103">Приступая к разработке пользовательских функций Excel</span><span class="sxs-lookup"><span data-stu-id="e1720-103">Get started developing Excel custom functions</span></span>

<span data-ttu-id="e1720-104">С помощью пользовательских функций разработчики могут добавлять новые функции в Excel, определяя их в JavaScript или typescript как часть надстройки.</span><span class="sxs-lookup"><span data-stu-id="e1720-104">With custom functions, developers can now add new functions to Excel by defining them in JavaScript or Typescript as part of an add-in.</span></span> <span data-ttu-id="e1720-105">Пользователи Excel могут получить доступ к пользовательским функциям так же, как и к любой собственной `SUM()`функции в Excel, например.</span><span class="sxs-lookup"><span data-stu-id="e1720-105">Excel users can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="e1720-106">Необходимые условия</span><span class="sxs-lookup"><span data-stu-id="e1720-106">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

* <span data-ttu-id="e1720-107">Excel в Windows (версия 1904 или более поздняя версия, подключенная к подписке на Office 365) или Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="e1720-107">Excel on Windows (version 1904 or later, connected to Office 365 subscription) or Excel on the web</span></span>
* <span data-ttu-id="e1720-108">Пользовательские функции Excel поддерживаются в Office в Mac (подключены к подписке Office 365), а обновление для этого учебника выполняется.</span><span class="sxs-lookup"><span data-stu-id="e1720-108">Excel custom functions are supported in Office on Mac (connected to Office 365 subscription) and an update to this tutorial is forthcoming.</span></span>

>[!NOTE]
><span data-ttu-id="e1720-109">Пользовательские функции Excel не поддерживаются в Office 2019 (одноразовая покупка).</span><span class="sxs-lookup"><span data-stu-id="e1720-109">Excel custom functions are not supported in Office 2019 (one-time purchase).</span></span>

## <a name="build-your-first-custom-functions-project"></a><span data-ttu-id="e1720-110">Создание проекта пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="e1720-110">Build your first custom functions project</span></span>

<span data-ttu-id="e1720-111">Чтобы начать работу, создайте проект пользовательских функций с помощью генератора Yeoman.</span><span class="sxs-lookup"><span data-stu-id="e1720-111">To start, you'll use the Yeoman generator to create the custom functions project.</span></span> <span data-ttu-id="e1720-112">Это позволит настроить для проекта правильную структуру папок, исходные файлы и зависимости, чтобы начать написание кода пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="e1720-112">This will set up your project with the correct folder structure, source files, and dependencies to begin coding your custom functions.</span></span>

1. [!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

    - <span data-ttu-id="e1720-113">**Выберите тип проекта:** `Excel Custom Functions Add-in project`</span><span class="sxs-lookup"><span data-stu-id="e1720-113">**Choose a project type:** `Excel Custom Functions Add-in project`</span></span>
    - <span data-ttu-id="e1720-114">**Выберите тип сценария:** `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="e1720-114">**Choose a script type:** `JavaScript`</span></span>
    - <span data-ttu-id="e1720-115">**Как вы хотите назвать надстройку?**</span><span class="sxs-lookup"><span data-stu-id="e1720-115">**What do you want to name your add-in?**</span></span> `starcount`

    ![Генератор Yeoman для надстройки Office, приглашающий к созданию пользовательских функций](../images/starcountPrompt.png)

    <span data-ttu-id="e1720-117">Генератор Yeoman создаст файлы проекта и установит вспомогательные компоненты Node.</span><span class="sxs-lookup"><span data-stu-id="e1720-117">The Yeoman generator will create the project files and install supporting Node components.</span></span>

2. <span data-ttu-id="e1720-118">Генератор Yeoman предоставит вам некоторые инструкции в командной строке о том, что делать с проектом, но проигнорируя их и продолжите выполнение дальнейших действий.</span><span class="sxs-lookup"><span data-stu-id="e1720-118">The Yeoman generator will give you some instructions in your command line about what to do with the project, but ignore them and continue to follow our instructions.</span></span> <span data-ttu-id="e1720-119">Перейдите к корневой папке проекта.</span><span class="sxs-lookup"><span data-stu-id="e1720-119">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd starcount
    ```

3. <span data-ttu-id="e1720-120">Выполните построение проекта.</span><span class="sxs-lookup"><span data-stu-id="e1720-120">Build the project.</span></span> 

    ```command&nbsp;line
    npm run build
    ```

    > [!NOTE]
    > <span data-ttu-id="e1720-121">Надстройки Office должны использовать HTTPS, а не HTTP, даже в случае разработки.</span><span class="sxs-lookup"><span data-stu-id="e1720-121">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="e1720-122">Если вам будет предложено установить сертификат после того, как вы запустите `npm run build`, примите предложение установить сертификат от генератора Yeoman.</span><span class="sxs-lookup"><span data-stu-id="e1720-122">If you are prompted to install a certificate after you run `npm run build`, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

4. <span data-ttu-id="e1720-123">Запустите локальный веб-сервер, работающий на Node.js.</span><span class="sxs-lookup"><span data-stu-id="e1720-123">Start the local web server, which runs in Node.js.</span></span> <span data-ttu-id="e1720-124">Вы можете испытать надстройку настраиваемой функции в Excel в Интернете или в Windows.</span><span class="sxs-lookup"><span data-stu-id="e1720-124">You can try out the custom function add-in in Excel on the web or Windows.</span></span> <span data-ttu-id="e1720-125">Вам может быть предложено открыть область задач надстройки, хотя это необязательный параметр.</span><span class="sxs-lookup"><span data-stu-id="e1720-125">You may be prompted to open the add-in's task pane, although this is optional.</span></span> <span data-ttu-id="e1720-126">Вы можете по-прежнему запускать пользовательские функции, не открывая область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="e1720-126">You can still run your custom functions without opening your add-in's task pane.</span></span>

# <a name="excel-on-windowstabexcel-windows"></a>[<span data-ttu-id="e1720-127">Excel в Windows</span><span class="sxs-lookup"><span data-stu-id="e1720-127">Excel on Windows</span></span>](#tab/excel-windows)

<span data-ttu-id="e1720-128">Чтобы протестировать надстройку в Excel в Windows, выполните следующую команду.</span><span class="sxs-lookup"><span data-stu-id="e1720-128">To test your add-in in Excel on Windows, run the following command.</span></span> <span data-ttu-id="e1720-129">При выполнении этой команды запустится локальный веб-сервер, и откроется приложение Excel с загруженной надстройкой.</span><span class="sxs-lookup"><span data-stu-id="e1720-129">When you run this command, the local web server will start and Excel will open with your add-in loaded.</span></span>

```command&nbsp;line
npm run start:desktop
```

# <a name="excel-on-the-webtabexcel-online"></a>[<span data-ttu-id="e1720-130">Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="e1720-130">Excel on the web</span></span>](#tab/excel-online)

<span data-ttu-id="e1720-131">Чтобы протестировать надстройку в Excel в Интернете, выполните следующую команду.</span><span class="sxs-lookup"><span data-stu-id="e1720-131">To test your add-in in Excel on the web, run the following command.</span></span> <span data-ttu-id="e1720-132">После выполнения этой команды запустится локальный веб-сервер.</span><span class="sxs-lookup"><span data-stu-id="e1720-132">When you run this command, the local web server will start.</span></span>

```command&nbsp;line
npm run start:web
```

<span data-ttu-id="e1720-133">Чтобы использовать надстройку с пользовательскими функциями, откройте новую книгу в Excel в браузере.</span><span class="sxs-lookup"><span data-stu-id="e1720-133">To use your custom functions add-in, open a new workbook in Excel on a browser.</span></span> <span data-ttu-id="e1720-134">В этой книге выполните следующие действия, чтобы Загрузка неопубликованных надстройку.</span><span class="sxs-lookup"><span data-stu-id="e1720-134">In this workbook, complete the following steps to sideload your add-in.</span></span>

1. <span data-ttu-id="e1720-135">В Excel перейдите на вкладку **Вставка** , а затем выберите **пункт**надстройки.</span><span class="sxs-lookup"><span data-stu-id="e1720-135">In Excel, choose the **Insert** tab and then choose **Add-ins**.</span></span>

   ![Вставка ленты в Excel в Интернете с выделенным значком "Мои надстройки"](../images/excel-cf-online-register-add-in-1.png)
   
2. <span data-ttu-id="e1720-137">Выберите пункт **Управление моими надстройками**, а затем выберите **Отправить мою надстройку**.</span><span class="sxs-lookup"><span data-stu-id="e1720-137">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="e1720-138">Выберите \*\*Обзор... \*\* и откройте корневой каталог проекта, созданный генератором Yeoman.</span><span class="sxs-lookup"><span data-stu-id="e1720-138">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="e1720-139">Выберите файл **manifest.xml** и нажмите кнопку **Открыть**, затем нажмите кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="e1720-139">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

---

## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="e1720-140">Проверка работы готовой пользовательской функции</span><span class="sxs-lookup"><span data-stu-id="e1720-140">Try out a prebuilt custom function</span></span>

<span data-ttu-id="e1720-141">Проект пользовательских функций, созданный с помощью генератора Yeoman, содержит некоторые предварительно созданные пользовательские функции, определенные в файле **./СРК/функтионс/функтионс.ЖС** .</span><span class="sxs-lookup"><span data-stu-id="e1720-141">The custom functions project that you created by using the Yeoman generator contains some prebuilt custom functions, defined within the **./src/functions/functions.js** file.</span></span> <span data-ttu-id="e1720-142">Файл **./манифест.ксмл** в корневом каталоге проекта указывает, что все пользовательские функции принадлежат `CONTOSO` пространству имен.</span><span class="sxs-lookup"><span data-stu-id="e1720-142">The **./manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.</span></span>

<span data-ttu-id="e1720-143">В книге Excel проверьте `ADD` пользовательскую функцию, выполнив следующие действия:</span><span class="sxs-lookup"><span data-stu-id="e1720-143">In your Excel workbook, try out the `ADD` custom function by completing the following steps:</span></span>

1. <span data-ttu-id="e1720-144">Выберите ячейку и введите текст `=CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="e1720-144">Select a cell and type `=CONTOSO`.</span></span> <span data-ttu-id="e1720-145">Обратите внимание на то, что в меню автозаполнения содержится список всех функций в пространстве имен `CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="e1720-145">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="e1720-146">Выполните `CONTOSO.ADD` функцию, используя числа `10` и `200` входные параметры, введя значение `=CONTOSO.ADD(10,200)` в ячейке и нажав клавишу ВВОД.</span><span class="sxs-lookup"><span data-stu-id="e1720-146">Run the `CONTOSO.ADD` function, using numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.</span></span>

<span data-ttu-id="e1720-147">Пользовательская функция `ADD` вычисляет сумму двух чисел, которые вы указываете в качестве входных параметров.</span><span class="sxs-lookup"><span data-stu-id="e1720-147">The `ADD` custom function computes the sum of the two numbers that you specify as input parameters.</span></span> <span data-ttu-id="e1720-148">При вводе `=CONTOSO.ADD(10,200)` в ячейке должен отобразиться результат **210** после нажатия клавиши ВВОД.</span><span class="sxs-lookup"><span data-stu-id="e1720-148">Typing `=CONTOSO.ADD(10,200)` should produce the result **210** in the cell after you press enter.</span></span>

## <a name="next-steps"></a><span data-ttu-id="e1720-149">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="e1720-149">Next steps</span></span>

<span data-ttu-id="e1720-150">Поздравляем, вы успешно создали пользовательскую функцию в надстройке Excel!</span><span class="sxs-lookup"><span data-stu-id="e1720-150">Congratulations, you've successfully created a custom function in an Excel add-in!</span></span> <span data-ttu-id="e1720-151">Затем создайте более сложную надстройку с возможностью потоковой передачи данных.</span><span class="sxs-lookup"><span data-stu-id="e1720-151">Next, build a more complex add-in with streaming data capability.</span></span> <span data-ttu-id="e1720-152">Следующая ссылка поможет вам выполнить следующие действия в руководстве надстройки Excel с пользовательскими функциями.</span><span class="sxs-lookup"><span data-stu-id="e1720-152">The following link takes you through the next steps in the Excel add-in with custom functions tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="e1720-153">Руководство по надстройкам Excel для пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="e1720-153">Excel custom functions add-in tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md#create-a-custom-function-that-requests-data-from-the-web
)

## <a name="see-also"></a><span data-ttu-id="e1720-154">См. также</span><span class="sxs-lookup"><span data-stu-id="e1720-154">See also</span></span>

* [<span data-ttu-id="e1720-155">Обзор настраиваемых функций</span><span class="sxs-lookup"><span data-stu-id="e1720-155">Custom functions overview</span></span>](../excel/custom-functions-overview.md)
* [<span data-ttu-id="e1720-156">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="e1720-156">Custom functions metadata</span></span>](../excel/custom-functions-json.md)
* [<span data-ttu-id="e1720-157">Среда выполнения для пользовательских функций Excel</span><span class="sxs-lookup"><span data-stu-id="e1720-157">Runtime for Excel custom functions</span></span>](../excel/custom-functions-runtime.md)