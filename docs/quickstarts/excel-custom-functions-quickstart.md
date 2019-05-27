---
ms.date: 05/15/2019
description: Разработка настраиваемых функций в кратком руководстве по Excel.
title: Краткое руководство по настраиваемым функциям
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 372e493d85add0a942a8f18ad67f65d08c92f6f2
ms.sourcegitcommit: adaee1329ae9bb69e49bde7f54a4c0444c9ba642
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/24/2019
ms.locfileid: "34432257"
---
# <a name="get-started-developing-excel-custom-functions"></a><span data-ttu-id="f1063-103">Приступая к разработке пользовательских функций Excel</span><span class="sxs-lookup"><span data-stu-id="f1063-103">Get started developing Excel custom functions</span></span>

<span data-ttu-id="f1063-104">С помощью пользовательских функций разработчики могут добавлять новые функции в Excel, определяя их в JavaScript или typescript как часть надстройки.</span><span class="sxs-lookup"><span data-stu-id="f1063-104">With custom functions, developers can now add new functions to Excel by defining them in JavaScript or Typescript as part of an add-in.</span></span> <span data-ttu-id="f1063-105">Пользователи Excel могут получить доступ к пользовательским функциям так же, как и к любой собственной `SUM()`функции в Excel, например.</span><span class="sxs-lookup"><span data-stu-id="f1063-105">Excel users can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="f1063-106">Необходимые условия</span><span class="sxs-lookup"><span data-stu-id="f1063-106">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

* <span data-ttu-id="f1063-107">Excel в Windows (64-разрядная версия 1810 или более поздняя) или Excel Online</span><span class="sxs-lookup"><span data-stu-id="f1063-107">Excel on Windows (64-bit version 1810 or later) or Excel Online</span></span>

## <a name="build-your-first-custom-functions-project"></a><span data-ttu-id="f1063-108">Создание проекта пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="f1063-108">Build your first custom functions project</span></span>

<span data-ttu-id="f1063-109">Чтобы начать работу, создайте проект пользовательских функций с помощью генератора Yeoman.</span><span class="sxs-lookup"><span data-stu-id="f1063-109">To start, you'll use the Yeoman generator to create the custom functions project.</span></span> <span data-ttu-id="f1063-110">Это позволит настроить для проекта правильную структуру папок, исходные файлы и зависимости, чтобы начать написание кода пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="f1063-110">This will set up your project with the correct folder structure, source files, and dependencies to begin coding your custom functions.</span></span>

1. <span data-ttu-id="f1063-111">В выбранной папке выполните следующую команду, а затем ответьте на приглашение следующим образом.</span><span class="sxs-lookup"><span data-stu-id="f1063-111">In a folder of your choice, run the following command and then answer the prompts as follows.</span></span>

    ```command&nbsp;line
    yo office
    ```

    - <span data-ttu-id="f1063-112">**Выберите тип проекта:** `Excel Custom Functions Add-in project`</span><span class="sxs-lookup"><span data-stu-id="f1063-112">**Choose a project type:** `Excel Custom Functions Add-in project`</span></span>
    - <span data-ttu-id="f1063-113">**Выберите тип сценария:** `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="f1063-113">**Choose a script type:** `JavaScript`</span></span>
    - <span data-ttu-id="f1063-114">**Как вы хотите назвать надстройку?**</span><span class="sxs-lookup"><span data-stu-id="f1063-114">**What do you want to name your add-in?**</span></span> `stock-ticker`

    ![Генератор Yeoman для надстройки Office, приглашающий к созданию пользовательских функций](../images/UpdatedYoOfficePrompt.png)

    <span data-ttu-id="f1063-116">Генератор Yeoman создаст файлы проекта и установит вспомогательные компоненты Node.</span><span class="sxs-lookup"><span data-stu-id="f1063-116">The Yeoman generator will create the project files and install supporting Node components.</span></span>

2. <span data-ttu-id="f1063-117">Генератор Yeoman предоставит вам некоторые инструкции в командной строке о том, что делать с проектом, но проигнорируя их и продолжите выполнение дальнейших действий.</span><span class="sxs-lookup"><span data-stu-id="f1063-117">The Yeoman generator will give you some instructions in your command line about what to do with the project, but ignore them and continue to follow our instructions.</span></span> <span data-ttu-id="f1063-118">Перейдите к корневой папке проекта.</span><span class="sxs-lookup"><span data-stu-id="f1063-118">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd stock-ticker
    ```

3. <span data-ttu-id="f1063-119">Выполните построение проекта.</span><span class="sxs-lookup"><span data-stu-id="f1063-119">Build the project.</span></span> <span data-ttu-id="f1063-120">При этом также будут установлены сертификаты, необходимые вашему проекту для правильной работы.</span><span class="sxs-lookup"><span data-stu-id="f1063-120">This will also install certificates that your project needs in order to function properly.</span></span> 

    ```command&nbsp;line
    npm run build
    ```

4. <span data-ttu-id="f1063-121">Запустите локальный веб-сервер, работающий на Node.js.</span><span class="sxs-lookup"><span data-stu-id="f1063-121">Start the local web server, which runs in Node.js.</span></span> <span data-ttu-id="f1063-122">Вы можете испытать надстройку настраиваемой функции в Excel для Windows или Excel Online.</span><span class="sxs-lookup"><span data-stu-id="f1063-122">You can try out the custom function add-in in Excel on Windows or Excel Online.</span></span> <span data-ttu-id="f1063-123">Вам может быть предложено открыть область задач надстройки, хотя это необязательный параметр.</span><span class="sxs-lookup"><span data-stu-id="f1063-123">You may be prompted to open the add-in's task pane, although this is optional.</span></span> <span data-ttu-id="f1063-124">Вы можете по-прежнему запускать пользовательские функции, не открывая область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="f1063-124">You can still run your custom functions without opening your add-in's task pane.</span></span>

> [!NOTE]
> <span data-ttu-id="f1063-125">Надстройки Office должны использовать HTTPS, а не HTTP, даже в случае разработки.</span><span class="sxs-lookup"><span data-stu-id="f1063-125">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="f1063-126">Если вам будет предложено установить сертификат после того, как вы запустите `npm run start:desktop`, примите предложение установить сертификат от генератора Yeoman.</span><span class="sxs-lookup"><span data-stu-id="f1063-126">If you are prompted to install a certificate after you run `npm run start:desktop`, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

# <a name="excel-on-windowstabexcel-windows"></a>[<span data-ttu-id="f1063-127">Excel в Windows</span><span class="sxs-lookup"><span data-stu-id="f1063-127">Excel on Windows</span></span>](#tab/excel-windows)

<span data-ttu-id="f1063-128">Чтобы протестировать надстройку в Excel в Windows, выполните следующую команду.</span><span class="sxs-lookup"><span data-stu-id="f1063-128">To test your add-in in Excel on Windows, run the following command.</span></span> <span data-ttu-id="f1063-129">При выполнении этой команды запустится локальный веб-сервер, и откроется приложение Excel с загруженной надстройкой.</span><span class="sxs-lookup"><span data-stu-id="f1063-129">When you run this command, the local web server will start and Excel will open with your add-in loaded.</span></span>

```command&nbsp;line
npm run start:desktop
```

# <a name="excel-onlinetabexcel-online"></a>[<span data-ttu-id="f1063-130">Excel Online</span><span class="sxs-lookup"><span data-stu-id="f1063-130">Excel Online</span></span>](#tab/excel-online)

<span data-ttu-id="f1063-131">Чтобы протестировать надстройку в Excel Online, выполните следующую команду.</span><span class="sxs-lookup"><span data-stu-id="f1063-131">To test your add-in in Excel Online, run the following command.</span></span> <span data-ttu-id="f1063-132">После выполнения этой команды запустится локальный веб-сервер.</span><span class="sxs-lookup"><span data-stu-id="f1063-132">When you run this command, the local web server will start.</span></span>

```command&nbsp;line
npm run start:web
```

> [!NOTE]
> <span data-ttu-id="f1063-133">Надстройки Office должны использовать HTTPS, а не HTTP, даже в случае разработки.</span><span class="sxs-lookup"><span data-stu-id="f1063-133">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="f1063-134">Если вам будет предложено установить сертификат после того, как вы запустите `npm run start:web`, примите предложение установить сертификат от генератора Yeoman.</span><span class="sxs-lookup"><span data-stu-id="f1063-134">If you are prompted to install a certificate after you run `npm run start:web`, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

<span data-ttu-id="f1063-135">Чтобы использовать надстройку с пользовательскими функциями, откройте новую книгу в Excel Online.</span><span class="sxs-lookup"><span data-stu-id="f1063-135">To use your custom functions add-in, open a new workbook in Excel Online.</span></span> <span data-ttu-id="f1063-136">В этой книге выполните следующие действия, чтобы Загрузка неопубликованных надстройку.</span><span class="sxs-lookup"><span data-stu-id="f1063-136">In this workbook, complete the following steps to sideload your add-in.</span></span>

1. <span data-ttu-id="f1063-137">В Excel Online на вкладке **Вставка** выберите пункт **Надстройки**.</span><span class="sxs-lookup"><span data-stu-id="f1063-137">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.</span></span>

   ![Вставка ленты в Excel Online с выделенным значком "Мои надстройки"](../images/excel-cf-online-register-add-in-1.png)
   
2. <span data-ttu-id="f1063-139">Выберите пункт **Управление моими надстройками**, а затем выберите **Отправить мою надстройку**.</span><span class="sxs-lookup"><span data-stu-id="f1063-139">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="f1063-140">Выберите \*\*Обзор... \*\* и откройте корневой каталог проекта, созданный генератором Yeoman.</span><span class="sxs-lookup"><span data-stu-id="f1063-140">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="f1063-141">Выберите файл **manifest.xml** и нажмите кнопку **Открыть**, затем нажмите кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="f1063-141">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

---

## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="f1063-142">Проверка работы готовой пользовательской функции</span><span class="sxs-lookup"><span data-stu-id="f1063-142">Try out a prebuilt custom function</span></span>

<span data-ttu-id="f1063-143">Проект пользовательских функций, созданный с помощью генератора Yeoman, содержит некоторые предварительно созданные пользовательские функции, определенные в файле **./СРК/функтионс/функтионс.ЖС** .</span><span class="sxs-lookup"><span data-stu-id="f1063-143">The custom functions project that you created by using the Yeoman generator contains some prebuilt custom functions, defined within the **./src/functions/functions.js** file.</span></span> <span data-ttu-id="f1063-144">Файл **./манифест.ксмл** в корневом каталоге проекта указывает, что все пользовательские функции принадлежат `CONTOSO` пространству имен.</span><span class="sxs-lookup"><span data-stu-id="f1063-144">The **./manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.</span></span>

<span data-ttu-id="f1063-145">В книге Excel проверьте `ADD` пользовательскую функцию, выполнив следующие действия:</span><span class="sxs-lookup"><span data-stu-id="f1063-145">In your Excel workbook, try out the `ADD` custom function by completing the following steps:</span></span>

1. <span data-ttu-id="f1063-146">Выберите ячейку и введите текст `=CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="f1063-146">Select a cell and type `=CONTOSO`.</span></span> <span data-ttu-id="f1063-147">Обратите внимание на то, что в меню автозаполнения содержится список всех функций в пространстве имен `CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="f1063-147">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="f1063-148">Выполните `CONTOSO.ADD` функцию, используя числа `10` и `200` входные параметры, введя значение `=CONTOSO.ADD(10,200)` в ячейке и нажав клавишу ВВОД.</span><span class="sxs-lookup"><span data-stu-id="f1063-148">Run the `CONTOSO.ADD` function, using numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.</span></span>

<span data-ttu-id="f1063-149">Пользовательская функция `ADD` вычисляет сумму двух чисел, которые вы указываете в качестве входных параметров.</span><span class="sxs-lookup"><span data-stu-id="f1063-149">The `ADD` custom function computes the sum of the two numbers that you specify as input parameters.</span></span> <span data-ttu-id="f1063-150">При вводе `=CONTOSO.ADD(10,200)` в ячейке должен отобразиться результат **210** после нажатия клавиши ВВОД.</span><span class="sxs-lookup"><span data-stu-id="f1063-150">Typing `=CONTOSO.ADD(10,200)` should produce the result **210** in the cell after you press enter.</span></span>

## <a name="next-steps"></a><span data-ttu-id="f1063-151">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="f1063-151">Next steps</span></span>

<span data-ttu-id="f1063-152">Поздравляем, вы успешно создали пользовательскую функцию в надстройке Excel!</span><span class="sxs-lookup"><span data-stu-id="f1063-152">Congratulations, you've successfully created a custom function in an Excel add-in!</span></span> <span data-ttu-id="f1063-153">Затем создайте более сложную надстройку с возможностью потоковой передачи данных.</span><span class="sxs-lookup"><span data-stu-id="f1063-153">Next, build a more complex add-in with streaming data capability.</span></span> <span data-ttu-id="f1063-154">Следующая ссылка поможет вам выполнить следующие действия в руководстве надстройки Excel с пользовательскими функциями.</span><span class="sxs-lookup"><span data-stu-id="f1063-154">The following link takes you through the next steps in the Excel add-in with custom functions tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="f1063-155">Руководство по надстройкам Excel для пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="f1063-155">Excel custom functions add-in tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md#create-a-custom-function-that-requests-data-from-the-web
)

## <a name="see-also"></a><span data-ttu-id="f1063-156">См. также</span><span class="sxs-lookup"><span data-stu-id="f1063-156">See also</span></span>

* [<span data-ttu-id="f1063-157">Обзор настраиваемых функций</span><span class="sxs-lookup"><span data-stu-id="f1063-157">Custom functions overview</span></span>](../excel/custom-functions-overview.md)
* [<span data-ttu-id="f1063-158">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="f1063-158">Custom functions metadata</span></span>](../excel/custom-functions-json.md)
* [<span data-ttu-id="f1063-159">Среда выполнения для пользовательских функций Excel</span><span class="sxs-lookup"><span data-stu-id="f1063-159">Runtime for Excel custom functions</span></span>](../excel/custom-functions-runtime.md)
* [<span data-ttu-id="f1063-160">Рекомендации по пользовательским функциям</span><span class="sxs-lookup"><span data-stu-id="f1063-160">Custom functions best practices</span></span>](../excel/custom-functions-best-practices.md)
