---
ms.date: 05/15/2019
description: Разработка настраиваемых функций в кратком руководстве по Excel.
title: Краткое руководство по настраиваемым функциям
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 2ef7e05390750173c1679cdb53ed96ff19020e67
ms.sourcegitcommit: b299b8a5dfffb6102cb14b431bdde4861abfb47f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/30/2019
ms.locfileid: "34589113"
---
# <a name="get-started-developing-excel-custom-functions"></a><span data-ttu-id="a6c4d-103">Приступая к разработке пользовательских функций Excel</span><span class="sxs-lookup"><span data-stu-id="a6c4d-103">Get started developing Excel custom functions</span></span>

<span data-ttu-id="a6c4d-104">С помощью пользовательских функций разработчики могут добавлять новые функции в Excel, определяя их в JavaScript или typescript как часть надстройки.</span><span class="sxs-lookup"><span data-stu-id="a6c4d-104">With custom functions, developers can now add new functions to Excel by defining them in JavaScript or Typescript as part of an add-in.</span></span> <span data-ttu-id="a6c4d-105">Пользователи Excel могут получить доступ к пользовательским функциям так же, как и к любой собственной `SUM()`функции в Excel, например.</span><span class="sxs-lookup"><span data-stu-id="a6c4d-105">Excel users can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="a6c4d-106">Необходимые условия</span><span class="sxs-lookup"><span data-stu-id="a6c4d-106">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

* <span data-ttu-id="a6c4d-107">Excel в Windows (64-разрядная версия 1810 или более поздняя) или Excel Online</span><span class="sxs-lookup"><span data-stu-id="a6c4d-107">Excel on Windows (64-bit version 1810 or later) or Excel Online</span></span>

## <a name="build-your-first-custom-functions-project"></a><span data-ttu-id="a6c4d-108">Создание проекта пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="a6c4d-108">Build your first custom functions project</span></span>

<span data-ttu-id="a6c4d-109">Чтобы начать работу, создайте проект пользовательских функций с помощью генератора Yeoman.</span><span class="sxs-lookup"><span data-stu-id="a6c4d-109">To start, you'll use the Yeoman generator to create the custom functions project.</span></span> <span data-ttu-id="a6c4d-110">Это позволит настроить для проекта правильную структуру папок, исходные файлы и зависимости, чтобы начать написание кода пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="a6c4d-110">This will set up your project with the correct folder structure, source files, and dependencies to begin coding your custom functions.</span></span>

1. <span data-ttu-id="a6c4d-111">В выбранной папке выполните следующую команду, а затем ответьте на приглашение следующим образом.</span><span class="sxs-lookup"><span data-stu-id="a6c4d-111">In a folder of your choice, run the following command and then answer the prompts as follows.</span></span>

    ```command&nbsp;line
    yo office
    ```

    - <span data-ttu-id="a6c4d-112">**Выберите тип проекта:** `Excel Custom Functions Add-in project`</span><span class="sxs-lookup"><span data-stu-id="a6c4d-112">**Choose a project type:** `Excel Custom Functions Add-in project`</span></span>
    - <span data-ttu-id="a6c4d-113">**Выберите тип сценария:** `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="a6c4d-113">**Choose a script type:** `JavaScript`</span></span>
    - <span data-ttu-id="a6c4d-114">**Как вы хотите назвать надстройку?**</span><span class="sxs-lookup"><span data-stu-id="a6c4d-114">**What do you want to name your add-in?**</span></span> `stock-ticker`

    ![Генератор Yeoman для надстройки Office, приглашающий к созданию пользовательских функций](../images/UpdatedYoOfficePrompt.png)

    <span data-ttu-id="a6c4d-116">Генератор Yeoman создаст файлы проекта и установит вспомогательные компоненты Node.</span><span class="sxs-lookup"><span data-stu-id="a6c4d-116">The Yeoman generator will create the project files and install supporting Node components.</span></span>

2. <span data-ttu-id="a6c4d-117">Генератор Yeoman предоставит вам некоторые инструкции в командной строке о том, что делать с проектом, но проигнорируя их и продолжите выполнение дальнейших действий.</span><span class="sxs-lookup"><span data-stu-id="a6c4d-117">The Yeoman generator will give you some instructions in your command line about what to do with the project, but ignore them and continue to follow our instructions.</span></span> <span data-ttu-id="a6c4d-118">Перейдите к корневой папке проекта.</span><span class="sxs-lookup"><span data-stu-id="a6c4d-118">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd stock-ticker
    ```

3. <span data-ttu-id="a6c4d-119">Выполните построение проекта.</span><span class="sxs-lookup"><span data-stu-id="a6c4d-119">Build the project.</span></span> 

    ```command&nbsp;line
    npm run build
    ```

    > [!NOTE]
    > <span data-ttu-id="a6c4d-120">Надстройки Office должны использовать HTTPS, а не HTTP, даже в случае разработки.</span><span class="sxs-lookup"><span data-stu-id="a6c4d-120">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="a6c4d-121">Если вам будет предложено установить сертификат после того, как вы запустите `npm run build`, примите предложение установить сертификат от генератора Yeoman.</span><span class="sxs-lookup"><span data-stu-id="a6c4d-121">If you are prompted to install a certificate after you run `npm run build`, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

4. <span data-ttu-id="a6c4d-122">Запустите локальный веб-сервер, работающий на Node.js.</span><span class="sxs-lookup"><span data-stu-id="a6c4d-122">Start the local web server, which runs in Node.js.</span></span> <span data-ttu-id="a6c4d-123">Вы можете испытать надстройку настраиваемой функции в Excel для Windows или Excel Online.</span><span class="sxs-lookup"><span data-stu-id="a6c4d-123">You can try out the custom function add-in in Excel on Windows or Excel Online.</span></span> <span data-ttu-id="a6c4d-124">Вам может быть предложено открыть область задач надстройки, хотя это необязательный параметр.</span><span class="sxs-lookup"><span data-stu-id="a6c4d-124">You may be prompted to open the add-in's task pane, although this is optional.</span></span> <span data-ttu-id="a6c4d-125">Вы можете по-прежнему запускать пользовательские функции, не открывая область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="a6c4d-125">You can still run your custom functions without opening your add-in's task pane.</span></span>

# <a name="excel-on-windowstabexcel-windows"></a>[<span data-ttu-id="a6c4d-126">Excel в Windows</span><span class="sxs-lookup"><span data-stu-id="a6c4d-126">Excel on Windows</span></span>](#tab/excel-windows)

<span data-ttu-id="a6c4d-127">Чтобы протестировать надстройку в Excel в Windows, выполните следующую команду.</span><span class="sxs-lookup"><span data-stu-id="a6c4d-127">To test your add-in in Excel on Windows, run the following command.</span></span> <span data-ttu-id="a6c4d-128">При выполнении этой команды запустится локальный веб-сервер, и откроется приложение Excel с загруженной надстройкой.</span><span class="sxs-lookup"><span data-stu-id="a6c4d-128">When you run this command, the local web server will start and Excel will open with your add-in loaded.</span></span>

```command&nbsp;line
npm run start:desktop
```

# <a name="excel-onlinetabexcel-online"></a>[<span data-ttu-id="a6c4d-129">Excel Online</span><span class="sxs-lookup"><span data-stu-id="a6c4d-129">Excel Online</span></span>](#tab/excel-online)

<span data-ttu-id="a6c4d-130">Чтобы протестировать надстройку в Excel Online, выполните следующую команду.</span><span class="sxs-lookup"><span data-stu-id="a6c4d-130">To test your add-in in Excel Online, run the following command.</span></span> <span data-ttu-id="a6c4d-131">После выполнения этой команды запустится локальный веб-сервер.</span><span class="sxs-lookup"><span data-stu-id="a6c4d-131">When you run this command, the local web server will start.</span></span>

```command&nbsp;line
npm run start:web
```

<span data-ttu-id="a6c4d-132">Чтобы использовать надстройку с пользовательскими функциями, откройте новую книгу в Excel Online.</span><span class="sxs-lookup"><span data-stu-id="a6c4d-132">To use your custom functions add-in, open a new workbook in Excel Online.</span></span> <span data-ttu-id="a6c4d-133">В этой книге выполните следующие действия, чтобы Загрузка неопубликованных надстройку.</span><span class="sxs-lookup"><span data-stu-id="a6c4d-133">In this workbook, complete the following steps to sideload your add-in.</span></span>

1. <span data-ttu-id="a6c4d-134">В Excel Online на вкладке **Вставка** выберите пункт **Надстройки**.</span><span class="sxs-lookup"><span data-stu-id="a6c4d-134">In Excel Online, choose the **Insert** tab and then choose **Add-ins**.</span></span>

   ![Вставка ленты в Excel Online с выделенным значком "Мои надстройки"](../images/excel-cf-online-register-add-in-1.png)
   
2. <span data-ttu-id="a6c4d-136">Выберите пункт **Управление моими надстройками**, а затем выберите **Отправить мою надстройку**.</span><span class="sxs-lookup"><span data-stu-id="a6c4d-136">Choose **Manage My Add-ins** and select **Upload My Add-in**.</span></span>

3. <span data-ttu-id="a6c4d-137">Выберите \*\*Обзор... \*\* и откройте корневой каталог проекта, созданный генератором Yeoman.</span><span class="sxs-lookup"><span data-stu-id="a6c4d-137">Choose **Browse...** and navigate to the root directory of the project that the Yeoman generator created.</span></span>

4. <span data-ttu-id="a6c4d-138">Выберите файл **manifest.xml** и нажмите кнопку **Открыть**, затем нажмите кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="a6c4d-138">Select the file **manifest.xml** and choose **Open**, then choose **Upload**.</span></span>

---

## <a name="try-out-a-prebuilt-custom-function"></a><span data-ttu-id="a6c4d-139">Проверка работы готовой пользовательской функции</span><span class="sxs-lookup"><span data-stu-id="a6c4d-139">Try out a prebuilt custom function</span></span>

<span data-ttu-id="a6c4d-140">Проект пользовательских функций, созданный с помощью генератора Yeoman, содержит некоторые предварительно созданные пользовательские функции, определенные в файле **./СРК/функтионс/функтионс.ЖС** .</span><span class="sxs-lookup"><span data-stu-id="a6c4d-140">The custom functions project that you created by using the Yeoman generator contains some prebuilt custom functions, defined within the **./src/functions/functions.js** file.</span></span> <span data-ttu-id="a6c4d-141">Файл **./манифест.ксмл** в корневом каталоге проекта указывает, что все пользовательские функции принадлежат `CONTOSO` пространству имен.</span><span class="sxs-lookup"><span data-stu-id="a6c4d-141">The **./manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.</span></span>

<span data-ttu-id="a6c4d-142">В книге Excel проверьте `ADD` пользовательскую функцию, выполнив следующие действия:</span><span class="sxs-lookup"><span data-stu-id="a6c4d-142">In your Excel workbook, try out the `ADD` custom function by completing the following steps:</span></span>

1. <span data-ttu-id="a6c4d-143">Выберите ячейку и введите текст `=CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="a6c4d-143">Select a cell and type `=CONTOSO`.</span></span> <span data-ttu-id="a6c4d-144">Обратите внимание на то, что в меню автозаполнения содержится список всех функций в пространстве имен `CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="a6c4d-144">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="a6c4d-145">Выполните `CONTOSO.ADD` функцию, используя числа `10` и `200` входные параметры, введя значение `=CONTOSO.ADD(10,200)` в ячейке и нажав клавишу ВВОД.</span><span class="sxs-lookup"><span data-stu-id="a6c4d-145">Run the `CONTOSO.ADD` function, using numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.</span></span>

<span data-ttu-id="a6c4d-146">Пользовательская функция `ADD` вычисляет сумму двух чисел, которые вы указываете в качестве входных параметров.</span><span class="sxs-lookup"><span data-stu-id="a6c4d-146">The `ADD` custom function computes the sum of the two numbers that you specify as input parameters.</span></span> <span data-ttu-id="a6c4d-147">При вводе `=CONTOSO.ADD(10,200)` в ячейке должен отобразиться результат **210** после нажатия клавиши ВВОД.</span><span class="sxs-lookup"><span data-stu-id="a6c4d-147">Typing `=CONTOSO.ADD(10,200)` should produce the result **210** in the cell after you press enter.</span></span>

## <a name="next-steps"></a><span data-ttu-id="a6c4d-148">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="a6c4d-148">Next steps</span></span>

<span data-ttu-id="a6c4d-149">Поздравляем, вы успешно создали пользовательскую функцию в надстройке Excel!</span><span class="sxs-lookup"><span data-stu-id="a6c4d-149">Congratulations, you've successfully created a custom function in an Excel add-in!</span></span> <span data-ttu-id="a6c4d-150">Затем создайте более сложную надстройку с возможностью потоковой передачи данных.</span><span class="sxs-lookup"><span data-stu-id="a6c4d-150">Next, build a more complex add-in with streaming data capability.</span></span> <span data-ttu-id="a6c4d-151">Следующая ссылка поможет вам выполнить следующие действия в руководстве надстройки Excel с пользовательскими функциями.</span><span class="sxs-lookup"><span data-stu-id="a6c4d-151">The following link takes you through the next steps in the Excel add-in with custom functions tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="a6c4d-152">Руководство по надстройкам Excel для пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="a6c4d-152">Excel custom functions add-in tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md#create-a-custom-function-that-requests-data-from-the-web
)

## <a name="see-also"></a><span data-ttu-id="a6c4d-153">См. также</span><span class="sxs-lookup"><span data-stu-id="a6c4d-153">See also</span></span>

* [<span data-ttu-id="a6c4d-154">Обзор настраиваемых функций</span><span class="sxs-lookup"><span data-stu-id="a6c4d-154">Custom functions overview</span></span>](../excel/custom-functions-overview.md)
* [<span data-ttu-id="a6c4d-155">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="a6c4d-155">Custom functions metadata</span></span>](../excel/custom-functions-json.md)
* [<span data-ttu-id="a6c4d-156">Среда выполнения для пользовательских функций Excel</span><span class="sxs-lookup"><span data-stu-id="a6c4d-156">Runtime for Excel custom functions</span></span>](../excel/custom-functions-runtime.md)
* [<span data-ttu-id="a6c4d-157">Рекомендации по пользовательским функциям</span><span class="sxs-lookup"><span data-stu-id="a6c4d-157">Custom functions best practices</span></span>](../excel/custom-functions-best-practices.md)
