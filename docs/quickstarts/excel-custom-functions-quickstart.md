---
ms.date: 09/06/2019
description: Разработка настраиваемых функций в кратком руководстве по Excel.
title: Краткое руководство по настраиваемым функциям
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: b3344b19ec29b24407c83bb043dcbaa613d5e8ec
ms.sourcegitcommit: ce7e7087a4550b9c090dc565fee5eac08a2985a2
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/06/2019
ms.locfileid: "36782235"
---
# <a name="get-started-developing-excel-custom-functions"></a><span data-ttu-id="0915c-103">Приступая к разработке пользовательских функций Excel</span><span class="sxs-lookup"><span data-stu-id="0915c-103">Get started developing Excel custom functions</span></span>

<span data-ttu-id="0915c-104">С помощью пользовательских функций разработчики могут добавлять новые функции в Excel, определяя их в JavaScript или typescript как часть надстройки.</span><span class="sxs-lookup"><span data-stu-id="0915c-104">With custom functions, developers can now add new functions to Excel by defining them in JavaScript or Typescript as part of an add-in.</span></span> <span data-ttu-id="0915c-105">Пользователи Excel могут получить доступ к пользовательским функциям так же, как и к любой собственной `SUM()`функции в Excel, например.</span><span class="sxs-lookup"><span data-stu-id="0915c-105">Excel users can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="0915c-106">Необходимые условия</span><span class="sxs-lookup"><span data-stu-id="0915c-106">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

* <span data-ttu-id="0915c-107">Excel в Windows (версия 1904 или более поздняя версия, подключенная к подписке на Office 365) или Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="0915c-107">Excel on Windows (version 1904 or later, connected to Office 365 subscription) or Excel on the web</span></span>
* <span data-ttu-id="0915c-108">Пользовательские функции Excel поддерживаются в Office в Mac (подключены к подписке Office 365), а обновление для этого учебника выполняется.</span><span class="sxs-lookup"><span data-stu-id="0915c-108">Excel custom functions are supported in Office on Mac (connected to Office 365 subscription) and an update to this tutorial is forthcoming.</span></span>

>[!NOTE]
><span data-ttu-id="0915c-109">Пользовательские функции Excel не поддерживаются в Office 2019 (одноразовая покупка).</span><span class="sxs-lookup"><span data-stu-id="0915c-109">Excel custom functions are not supported in Office 2019 (one-time purchase).</span></span>

## <a name="build-your-first-custom-functions-project"></a><span data-ttu-id="0915c-110">Создание проекта пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="0915c-110">Build your first custom functions project</span></span>

<span data-ttu-id="0915c-111">Чтобы начать работу, создайте проект пользовательских функций с помощью генератора Yeoman.</span><span class="sxs-lookup"><span data-stu-id="0915c-111">To start, you'll use the Yeoman generator to create the custom functions project.</span></span> <span data-ttu-id="0915c-112">Это позволит настроить для проекта правильную структуру папок, исходные файлы и зависимости, чтобы начать написание кода пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="0915c-112">This will set up your project with the correct folder structure, source files, and dependencies to begin coding your custom functions.</span></span>

1. <span data-ttu-id="0915c-113">В выбранной папке выполните следующую команду, а затем ответьте на приглашение следующим образом.</span><span class="sxs-lookup"><span data-stu-id="0915c-113">In a folder of your choice, run the following command and then answer the prompts as follows.</span></span>

    ```command&nbsp;line
    yo office
    ```

    - <span data-ttu-id="0915c-114">**Выберите тип проекта:** `Excel Custom Functions Add-in project`</span><span class="sxs-lookup"><span data-stu-id="0915c-114">**Choose a project type:** `Excel Custom Functions Add-in project`</span></span>
    - <span data-ttu-id="0915c-115">**Выберите тип сценария:** `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="0915c-115">**Choose a script type:** `JavaScript`</span></span>
    - <span data-ttu-id="0915c-116">**Как вы хотите назвать надстройку?**</span><span class="sxs-lookup"><span data-stu-id="0915c-116">**What do you want to name your add-in?**</span></span> `starcount`

    ![Генератор Yeoman для надстройки Office, приглашающий к созданию пользовательских функций](../images/starcountPrompt.png)

    <span data-ttu-id="0915c-118">Генератор Yeoman создаст файлы проекта и установит вспомогательные компоненты Node.</span><span class="sxs-lookup"><span data-stu-id="0915c-118">The Yeoman generator will create the project files and install supporting Node components.</span></span>

2. <span data-ttu-id="0915c-119">Генератор Yeoman предоставит вам некоторые инструкции в командной строке о том, что делать с проектом, но проигнорируя их и продолжите выполнение дальнейших действий.</span><span class="sxs-lookup"><span data-stu-id="0915c-119">The Yeoman generator will give you some instructions in your command line about what to do with the project, but ignore them and continue to follow our instructions.</span></span> <span data-ttu-id="0915c-120">Перейдите к корневой папке проекта.</span><span class="sxs-lookup"><span data-stu-id="0915c-120">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd starcount
    ```

3. <span data-ttu-id="0915c-121">Выполните построение проекта.</span><span class="sxs-lookup"><span data-stu-id="0915c-121">Build the project.</span></span> 

    ```command&nbsp;line
    npm run build
    ```

    > [!NOTE]
    > <span data-ttu-id="0915c-122">Надстройки Office должны использовать HTTPS, а не HTTP, даже в случае разработки.</span><span class="sxs-lookup"><span data-stu-id="0915c-122">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="0915c-123">Если вам будет предложено установить сертификат после того, как вы запустите `npm run build`, примите предложение установить сертификат от генератора Yeoman.</span><span class="sxs-lookup"><span data-stu-id="0915c-123">If you are prompted to install a certificate after you run `npm run build`, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

4. <span data-ttu-id="0915c-124">Запустите локальный веб-сервер, работающий на Node.js.</span><span class="sxs-lookup"><span data-stu-id="0915c-124">Start the local web server, which runs in Node.js.</span></span> <span data-ttu-id="0915c-125">Вы можете испытать надстройку настраиваемой функции в Excel в Интернете или в Windows.</span><span class="sxs-lookup"><span data-stu-id="0915c-125">You can try out the custom function add-in in Excel on the web or Windows.</span></span> <span data-ttu-id="0915c-126">Вам может быть предложено открыть область задач надстройки, хотя это необязательный параметр.</span><span class="sxs-lookup"><span data-stu-id="0915c-126">You may be prompted to open the add-in's task pane, although this is optional.</span></span> <span data-ttu-id="0915c-127">Вы можете по-прежнему запускать пользовательские функции, не открывая область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="0915c-127">You can still run your custom functions without opening your add-in's task pane.</span></span>

# <a name="excel-on-windowstabexcel-windows"></a>[<span data-ttu-id="0915c-128">Excel в Windows</span><span class="sxs-lookup"><span data-stu-id="0915c-128">Excel on Windows</span></span>](#tab/excel-windows)

<span data-ttu-id="0915c-129">Чтобы протестировать надстройку в Excel в Windows, выполните следующую команду.</span><span class="sxs-lookup"><span data-stu-id="0915c-129">To test your add-in in Excel on Windows, run the following command.</span></span> <span data-ttu-id="0915c-130">При выполнении этой команды запустится локальный веб-сервер, и откроется приложение Excel с загруженной надстройкой.</span><span class="sxs-lookup"><span data-stu-id="0915c-130">When you run this command, the local web server will start and Excel will open with your add-in loaded.</span></span>

```command&nbsp;line
npm run start:desktop
```

# <a name="excel-on-the-webtabexcel-online"></a>[<span data-ttu-id="0915c-131">Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="0915c-131">Excel on the web</span></span>](#tab/excel-online)

<span data-ttu-id="0915c-132">Чтобы протестировать надстройку в Excel в Интернете, выполните следующую команду.</span><span class="sxs-lookup"><span data-stu-id="0915c-132">To test your add-in in Excel on the web, run the following command.</span></span> <span data-ttu-id="0915c-133">После выполнения этой команды запустится локальный веб-сервер.</span><span class="sxs-lookup"><span data-stu-id="0915c-133">When you run this command, the local web server will start.</span></span>

```command&nbsp;line
npm run start:web
```

<span data-ttu-id="0915c-134">Чтобы использовать надстройку с пользовательскими функциями, откройте новую книгу в Excel в браузере.</span><span class="sxs-lookup"><span data-stu-id="0915c-134">To use your custom functions add-in, open a new workbook in Excel on a browser.</span></span> <span data-ttu-id="0915c-135">В этой книге выполните следующие действия, чтобы Загрузка неопубликованных надстройку.</span><span class="sxs-lookup"><span data-stu-id="0915c-135">In this workbook, complete the following steps to sideload your add-in.</span></span>

1. <span data-ttu-id="0915c-136">В Excel перейдите на вкладку **Вставка** , а затем выберите **пункт**надстройки.</span><span class="sxs-lookup"><span data-stu-id="0915c-136">In Excel, choose the **Insert** tab and then choose **Add-ins**.</span></span>

   ![Вставка ленты в Excel в Интернете с выделенным значком "Мои надстройки"](../images/excel-cf-online-register-add-in-1.png)
   
2. <span data-ttu-id="0915c-138">Выберите пункт **Управление моими надстройками**, а затем выберите **Отправить мою надстройку**.</span><span class="sxs-lookup"><span data-stu-id="0915c-138">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="0915c-139">Выберите \*\*Обзор... \*\* и откройте корневой каталог проекта, созданный генератором Yeoman.</span><span class="sxs-lookup"><span data-stu-id="0915c-139">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="0915c-140">Выберите файл **manifest.xml** и нажмите кнопку **Открыть**, затем нажмите кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="0915c-140">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

---

## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="0915c-141">Проверка работы готовой пользовательской функции</span><span class="sxs-lookup"><span data-stu-id="0915c-141">Try out a prebuilt custom function</span></span>

<span data-ttu-id="0915c-142">Проект пользовательских функций, созданный с помощью генератора Yeoman, содержит некоторые предварительно созданные пользовательские функции, определенные в файле **./СРК/функтионс/функтионс.ЖС** .</span><span class="sxs-lookup"><span data-stu-id="0915c-142">The custom functions project that you created by using the Yeoman generator contains some prebuilt custom functions, defined within the **./src/functions/functions.js** file.</span></span> <span data-ttu-id="0915c-143">Файл **./манифест.ксмл** в корневом каталоге проекта указывает, что все пользовательские функции принадлежат `CONTOSO` пространству имен.</span><span class="sxs-lookup"><span data-stu-id="0915c-143">The **./manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.</span></span>

<span data-ttu-id="0915c-144">В книге Excel проверьте `ADD` пользовательскую функцию, выполнив следующие действия:</span><span class="sxs-lookup"><span data-stu-id="0915c-144">In your Excel workbook, try out the `ADD` custom function by completing the following steps:</span></span>

1. <span data-ttu-id="0915c-145">Выберите ячейку и введите текст `=CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="0915c-145">Select a cell and type `=CONTOSO`.</span></span> <span data-ttu-id="0915c-146">Обратите внимание на то, что в меню автозаполнения содержится список всех функций в пространстве имен `CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="0915c-146">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="0915c-147">Выполните `CONTOSO.ADD` функцию, используя числа `10` и `200` входные параметры, введя значение `=CONTOSO.ADD(10,200)` в ячейке и нажав клавишу ВВОД.</span><span class="sxs-lookup"><span data-stu-id="0915c-147">Run the `CONTOSO.ADD` function, using numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.</span></span>

<span data-ttu-id="0915c-148">Пользовательская функция `ADD` вычисляет сумму двух чисел, которые вы указываете в качестве входных параметров.</span><span class="sxs-lookup"><span data-stu-id="0915c-148">The `ADD` custom function computes the sum of the two numbers that you specify as input parameters.</span></span> <span data-ttu-id="0915c-149">При вводе `=CONTOSO.ADD(10,200)` в ячейке должен отобразиться результат **210** после нажатия клавиши ВВОД.</span><span class="sxs-lookup"><span data-stu-id="0915c-149">Typing `=CONTOSO.ADD(10,200)` should produce the result **210** in the cell after you press enter.</span></span>

## <a name="next-steps"></a><span data-ttu-id="0915c-150">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="0915c-150">Next steps</span></span>

<span data-ttu-id="0915c-151">Поздравляем, вы успешно создали пользовательскую функцию в надстройке Excel!</span><span class="sxs-lookup"><span data-stu-id="0915c-151">Congratulations, you've successfully created a custom function in an Excel add-in!</span></span> <span data-ttu-id="0915c-152">Затем создайте более сложную надстройку с возможностью потоковой передачи данных.</span><span class="sxs-lookup"><span data-stu-id="0915c-152">Next, build a more complex add-in with streaming data capability.</span></span> <span data-ttu-id="0915c-153">Следующая ссылка поможет вам выполнить следующие действия в руководстве надстройки Excel с пользовательскими функциями.</span><span class="sxs-lookup"><span data-stu-id="0915c-153">The following link takes you through the next steps in the Excel add-in with custom functions tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="0915c-154">Руководство по надстройкам Excel для пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="0915c-154">Excel custom functions add-in tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md#create-a-custom-function-that-requests-data-from-the-web
)

## <a name="see-also"></a><span data-ttu-id="0915c-155">См. также</span><span class="sxs-lookup"><span data-stu-id="0915c-155">See also</span></span>

* [<span data-ttu-id="0915c-156">Обзор настраиваемых функций</span><span class="sxs-lookup"><span data-stu-id="0915c-156">Custom functions overview</span></span>](../excel/custom-functions-overview.md)
* [<span data-ttu-id="0915c-157">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="0915c-157">Custom functions metadata</span></span>](../excel/custom-functions-json.md)
* [<span data-ttu-id="0915c-158">Среда выполнения для пользовательских функций Excel</span><span class="sxs-lookup"><span data-stu-id="0915c-158">Runtime for Excel custom functions</span></span>](../excel/custom-functions-runtime.md)