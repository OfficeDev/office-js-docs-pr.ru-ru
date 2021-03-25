---
ms.date: 07/10/2020
description: Узнайте, как отлагировать настраиваемые функции Excel, которые не используют области задач.
title: Отладка пользовательских функций без пользовательского интерфейса
localization_priority: Normal
ms.openlocfilehash: 00065a465a22f83891dfb207943102b079e96a0f
ms.sourcegitcommit: 7482ab6bc258d98acb9ba9b35c7dd3b5cc5bed21
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/24/2021
ms.locfileid: "51178078"
---
# <a name="ui-less-custom-functions-debugging"></a><span data-ttu-id="d684d-103">Отладка пользовательских функций без пользовательского интерфейса</span><span class="sxs-lookup"><span data-stu-id="d684d-103">UI-less custom functions debugging</span></span>

<span data-ttu-id="d684d-104">Отладка настраиваемой функции, не использующей области задач или другие элементы пользовательского интерфейса (пользовательские функции без пользовательского интерфейса), может быть выполнена несколькими средствами в зависимости от используемой платформы.</span><span class="sxs-lookup"><span data-stu-id="d684d-104">Debugging for custom functions that don't use a task pane or other user interface elements (UI-less custom functions) can be accomplished by multiple means, depending on what platform you're using.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

<span data-ttu-id="d684d-105">В Windows:</span><span class="sxs-lookup"><span data-stu-id="d684d-105">On Windows:</span></span>
- [<span data-ttu-id="d684d-106">Отладка Visual Studio и кода Excel</span><span class="sxs-lookup"><span data-stu-id="d684d-106">Excel Desktop and Visual Studio Code (VS Code) debugger</span></span>](#use-the-vs-code-debugger-for-excel-desktop)
- [<span data-ttu-id="d684d-107">Excel в Интернете и отладка кода VS</span><span class="sxs-lookup"><span data-stu-id="d684d-107">Excel on the web and VS Code debugger</span></span>](#use-the-vs-code-debugger-for-excel-in-microsoft-edge)
- [<span data-ttu-id="d684d-108">Excel в веб-средствах и средствах браузера</span><span class="sxs-lookup"><span data-stu-id="d684d-108">Excel on the web and browser tools</span></span>](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [<span data-ttu-id="d684d-109">Командная строка</span><span class="sxs-lookup"><span data-stu-id="d684d-109">Command line</span></span>](#use-the-command-line-tools-to-debug)

<span data-ttu-id="d684d-110">На Mac:</span><span class="sxs-lookup"><span data-stu-id="d684d-110">On Mac:</span></span>
- [<span data-ttu-id="d684d-111">Excel в веб-средствах и средствах браузера</span><span class="sxs-lookup"><span data-stu-id="d684d-111">Excel on the web and browser tools</span></span>](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [<span data-ttu-id="d684d-112">Командная строка</span><span class="sxs-lookup"><span data-stu-id="d684d-112">Command line</span></span>](#use-the-command-line-tools-to-debug)

> [!NOTE]
> <span data-ttu-id="d684d-113">Для простоты в этой статье показана отладка в контексте использования Visual Studio кода для редактирования, выполнения задач и в некоторых случаях использования представления отладки.</span><span class="sxs-lookup"><span data-stu-id="d684d-113">For simplicity, this article shows debugging in the context of using Visual Studio Code to edit, run tasks, and in some cases use the debug view.</span></span> <span data-ttu-id="d684d-114">При использовании другого средства редактора или [](#commands-for-building-and-running-your-add-in) командной строки см. инструкции по командной строке в конце этой статьи.</span><span class="sxs-lookup"><span data-stu-id="d684d-114">If you are using a different editor or command line tool, see the [command line instructions](#commands-for-building-and-running-your-add-in) at the end of this article.</span></span>

## <a name="requirements"></a><span data-ttu-id="d684d-115">Требования</span><span class="sxs-lookup"><span data-stu-id="d684d-115">Requirements</span></span>

<span data-ttu-id="d684d-116">Перед отлагиванием необходимо использовать генератор [Yeoman](https://github.com/OfficeDev/generator-office) для надстроек Office для создания настраиваемого проекта функций.</span><span class="sxs-lookup"><span data-stu-id="d684d-116">Before starting to debug, you should use the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) to create a custom functions project.</span></span> <span data-ttu-id="d684d-117">Руководство по созданию настраиваемой функции см. в руководстве [по пользовательским функциям.](../tutorials/excel-tutorial-create-custom-functions.md)</span><span class="sxs-lookup"><span data-stu-id="d684d-117">For guidance about how to create a custom functions project, see the [custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md).</span></span>

## <a name="use-the-vs-code-debugger-for-excel-desktop"></a><span data-ttu-id="d684d-118">Использование отладки кода VS для настольных компьютеров Excel</span><span class="sxs-lookup"><span data-stu-id="d684d-118">Use the VS Code debugger for Excel Desktop</span></span>

<span data-ttu-id="d684d-119">Вы можете использовать VS Code для отлаговки пользовательских функций без пользовательского интерфейса в Office Excel на рабочем столе.</span><span class="sxs-lookup"><span data-stu-id="d684d-119">You can use VS Code to debug UI-less custom functions in Office Excel on the desktop.</span></span>

> [!NOTE]
> <span data-ttu-id="d684d-120">Отладка рабочего стола для Mac недоступна, но может быть достигнута с помощью средств браузера и командной строки для отладки [Excel в Интернете).](#use-the-command-line-tools-to-debug)</span><span class="sxs-lookup"><span data-stu-id="d684d-120">Desktop debugging for the Mac is not available but can be achieved [using the browser tools and command line to debug Excel on the web](#use-the-command-line-tools-to-debug)).</span></span>

### <a name="run-your-add-in-from-vs-code"></a><span data-ttu-id="d684d-121">Запуск надстройки из VS Code</span><span class="sxs-lookup"><span data-stu-id="d684d-121">Run your add-in from VS Code</span></span>

1. <span data-ttu-id="d684d-122">Откройте настраиваемую папку корневого проекта функций [в VS Code.](https://code.visualstudio.com/)</span><span class="sxs-lookup"><span data-stu-id="d684d-122">Open your custom functions root project folder in [VS Code](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="d684d-123">Выберите **терминал > выполнить задачу и** введите или выберите **Часы**.</span><span class="sxs-lookup"><span data-stu-id="d684d-123">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="d684d-124">Это позволит отслеживать и восстанавливать любые изменения файлов.</span><span class="sxs-lookup"><span data-stu-id="d684d-124">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="d684d-125">Выберите **терминальный > выполнить задачу и** введите или выберите **Сервер разработчиков**.</span><span class="sxs-lookup"><span data-stu-id="d684d-125">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span>

### <a name="start-the-vs-code-debugger"></a><span data-ttu-id="d684d-126">Запуск отладки кода VS</span><span class="sxs-lookup"><span data-stu-id="d684d-126">Start the VS Code debugger</span></span>

4. <span data-ttu-id="d684d-127">Выберите **просмотр > выполнить** или ввести **Ctrl+Shift+D,** чтобы перейти на отлагивание представления.</span><span class="sxs-lookup"><span data-stu-id="d684d-127">Choose **View > Run** or enter **Ctrl+Shift+D** to switch to debug view.</span></span>
5. <span data-ttu-id="d684d-128">Из выпадаемого меню Run выберите **Excel Desktop (Edge Chromium).**</span><span class="sxs-lookup"><span data-stu-id="d684d-128">From the Run drop-down menu, choose **Excel Desktop (Edge Chromium)**.</span></span>
6. <span data-ttu-id="d684d-129">Чтобы начать отладку, выберите **F5** **(или > запустить** отладку из меню).</span><span class="sxs-lookup"><span data-stu-id="d684d-129">Select **F5** (or select **Run -> Start Debugging** from the menu) to begin debugging.</span></span> <span data-ttu-id="d684d-130">Новая книга Excel откроется с уже загруженной и готовой к использованию надстройке.</span><span class="sxs-lookup"><span data-stu-id="d684d-130">A new Excel workbook will open with your add-in already sideloaded and ready to use.</span></span>

### <a name="start-debugging"></a><span data-ttu-id="d684d-131">Начало отладки</span><span class="sxs-lookup"><span data-stu-id="d684d-131">Start debugging</span></span>

1. <span data-ttu-id="d684d-132">В vs Code откройте исходный файл скрипта кода **(functions.js** **или functions.ts).**</span><span class="sxs-lookup"><span data-stu-id="d684d-132">In VS Code, open your source code script file (**functions.js** or **functions.ts**).</span></span>
2. <span data-ttu-id="d684d-133">[Установите точку разрыва в](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) пользовательском коде источника функции.</span><span class="sxs-lookup"><span data-stu-id="d684d-133">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span>
3. <span data-ttu-id="d684d-134">В книге Excel введите формулу, использующую настраиваемую функцию.</span><span class="sxs-lookup"><span data-stu-id="d684d-134">In the Excel workbook, enter a formula that uses your custom function.</span></span>

<span data-ttu-id="d684d-135">На этом этапе выполнение остановится на строке кода, где установлена точка разрыва.</span><span class="sxs-lookup"><span data-stu-id="d684d-135">At this point execution will stop on the line of code where you set the breakpoint.</span></span> <span data-ttu-id="d684d-136">Теперь вы можете пройти через код, установить часы и использовать все необходимые функции отладки кода VS.</span><span class="sxs-lookup"><span data-stu-id="d684d-136">Now you can step through your code, set watches, and use any VS Code debugging features you need.</span></span>

## <a name="use-the-vs-code-debugger-for-excel-in-microsoft-edge"></a><span data-ttu-id="d684d-137">Использование отладки кода VS для Excel в Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="d684d-137">Use the VS Code debugger for Excel in Microsoft Edge</span></span>

<span data-ttu-id="d684d-138">Вы можете использовать VS Code для отлаговки пользовательских функций в Excel в браузере Microsoft Edge.</span><span class="sxs-lookup"><span data-stu-id="d684d-138">You can use VS Code to debug UI-less custom functions in Excel on the Microsoft Edge browser.</span></span> <span data-ttu-id="d684d-139">Чтобы использовать vs Code с Microsoft Edge, необходимо установить расширение [Debugger для Microsoft Edge.](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge)</span><span class="sxs-lookup"><span data-stu-id="d684d-139">To use VS Code with Microsoft Edge, you must install the [Debugger for Microsoft Edge](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge) extension.</span></span>

### <a name="run-your-add-in-from-vs-code"></a><span data-ttu-id="d684d-140">Запуск надстройки из VS Code</span><span class="sxs-lookup"><span data-stu-id="d684d-140">Run your add-in from VS Code</span></span>

1. <span data-ttu-id="d684d-141">Откройте настраиваемую папку корневого проекта функций [в VS Code.](https://code.visualstudio.com/)</span><span class="sxs-lookup"><span data-stu-id="d684d-141">Open your custom functions root project folder in [VS Code](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="d684d-142">Выберите **терминал > выполнить задачу и** введите или выберите **Часы**.</span><span class="sxs-lookup"><span data-stu-id="d684d-142">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="d684d-143">Это позволит отслеживать и восстанавливать любые изменения файлов.</span><span class="sxs-lookup"><span data-stu-id="d684d-143">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="d684d-144">Выберите **терминальный > выполнить задачу и** введите или выберите **Сервер разработчиков**.</span><span class="sxs-lookup"><span data-stu-id="d684d-144">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span>

### <a name="start-the-vs-code-debugger"></a><span data-ttu-id="d684d-145">Запуск отладки кода VS</span><span class="sxs-lookup"><span data-stu-id="d684d-145">Start the VS Code debugger</span></span>

4. <span data-ttu-id="d684d-146">Выберите **просмотр > выполнить** или ввести **Ctrl+Shift+D,** чтобы перейти на отлагивание представления.</span><span class="sxs-lookup"><span data-stu-id="d684d-146">Choose **View > Run** or enter **Ctrl+Shift+D** to switch to debug view.</span></span>
5. <span data-ttu-id="d684d-147">Из параметров отладки выберите **Office Online (Edge Chromium).**</span><span class="sxs-lookup"><span data-stu-id="d684d-147">From the Debug options, choose **Office Online (Edge Chromium)**.</span></span>
6. <span data-ttu-id="d684d-148">Откройте Excel в браузере Microsoft Edge и создайте новую книгу.</span><span class="sxs-lookup"><span data-stu-id="d684d-148">Open Excel in the Microsoft Edge browser and create a new workbook.</span></span>
7. <span data-ttu-id="d684d-149">Выберите **Share** в ленте и скопируйте ссылку на URL-адрес этой новой книги.</span><span class="sxs-lookup"><span data-stu-id="d684d-149">Choose **Share** in the ribbon and copy the link for the URL for this new workbook.</span></span>
8. <span data-ttu-id="d684d-150">Чтобы начать отладку, выберите **F5** (> **запустить** отладку из меню).</span><span class="sxs-lookup"><span data-stu-id="d684d-150">Select **F5** (or select **Run > Start Debugging** from the menu) to begin debugging.</span></span> <span data-ttu-id="d684d-151">Появится запрос, в котором будет указан URL-адрес документа.</span><span class="sxs-lookup"><span data-stu-id="d684d-151">A prompt will appear, which asks for the URL of your document.</span></span>
9. <span data-ttu-id="d684d-152">Введите URL-адрес книги и нажмите кнопку Ввод.</span><span class="sxs-lookup"><span data-stu-id="d684d-152">Paste in the URL for your workbook and press Enter.</span></span>

### <a name="sideload-your-add-in"></a><span data-ttu-id="d684d-153">Загрузка неопубликованной надстройки</span><span class="sxs-lookup"><span data-stu-id="d684d-153">Sideload your add-in</span></span>

1. <span data-ttu-id="d684d-154">Выберите **вкладку Insert** на ленте и в разделе Надстройки, выберите  **надстройки Office.**</span><span class="sxs-lookup"><span data-stu-id="d684d-154">Select the **Insert** tab on the ribbon and in the **Add-ins** section, choose **Office Add-ins**.</span></span>
2. <span data-ttu-id="d684d-155">В **диалоговом** окантовке Надстройки Office выберите вкладку **MY ADD-INS,** выберите **Управление** надстройками, а затем загрузите мои **надстройки.**</span><span class="sxs-lookup"><span data-stu-id="d684d-155">On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.</span></span>
    
    ![Диалоговое окно "Надстройки Office" с раскрывающимся меню в правом верхнем углу, в котором выделен пункт "Управление моими надстройками", а под ним — раскрывающийся список с пунктом "Отправить надстройку"](../images/office-add-ins-my-account.png)

3. <span data-ttu-id="d684d-157">**Просмотрите** файл манифеста надстройки и выберите **Upload**.</span><span class="sxs-lookup"><span data-stu-id="d684d-157">**Browse** to the add-in manifest file and then select **Upload**.</span></span>
    
    ![Диалоговое окно отправки надстройки с кнопками "Обзор", "Отправить" и "Отмена"](../images/upload-add-in.png)


### <a name="set-breakpoints"></a><span data-ttu-id="d684d-159">Настройка точек разрыва</span><span class="sxs-lookup"><span data-stu-id="d684d-159">Set breakpoints</span></span>
1. <span data-ttu-id="d684d-160">В vs Code откройте исходный файл скрипта кода **(functions.js** **или functions.ts).**</span><span class="sxs-lookup"><span data-stu-id="d684d-160">In VS Code, open your source code script file (**functions.js** or **functions.ts**).</span></span>
2. <span data-ttu-id="d684d-161">[Установите точку разрыва в](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) пользовательском коде источника функции.</span><span class="sxs-lookup"><span data-stu-id="d684d-161">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span>
3. <span data-ttu-id="d684d-162">В книге Excel введите формулу, использующую настраиваемую функцию.</span><span class="sxs-lookup"><span data-stu-id="d684d-162">In the Excel workbook, enter a formula that uses your custom function.</span></span>

## <a name="use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web"></a><span data-ttu-id="d684d-163">Использование средств разработчика браузера для отлаговки пользовательских функций в Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="d684d-163">Use the browser developer tools to debug custom functions in Excel on the web</span></span>

<span data-ttu-id="d684d-164">Средства разработчика браузера можно использовать для отлаговки пользовательских функций в Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="d684d-164">You can use the browser developer tools to debug UI-less custom functions in Excel on the web.</span></span> <span data-ttu-id="d684d-165">Следующие действия работают как для Windows, так и для macOS.</span><span class="sxs-lookup"><span data-stu-id="d684d-165">The following steps work for both Windows and macOS.</span></span>

### <a name="run-your-add-in-from-visual-studio-code"></a><span data-ttu-id="d684d-166">Запустите надстройку из Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="d684d-166">Run your add-in from Visual Studio Code</span></span>

1. <span data-ttu-id="d684d-167">Откройте настраиваемую папку корневого проекта [функций в Visual Studio Code (VS Code).](https://code.visualstudio.com/)</span><span class="sxs-lookup"><span data-stu-id="d684d-167">Open your custom functions root project folder in [Visual Studio Code (VS Code)](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="d684d-168">Выберите **терминал > выполнить задачу и** введите или выберите **Часы**.</span><span class="sxs-lookup"><span data-stu-id="d684d-168">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="d684d-169">Это позволит отслеживать и восстанавливать любые изменения файлов.</span><span class="sxs-lookup"><span data-stu-id="d684d-169">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="d684d-170">Выберите **терминальный > выполнить задачу и** введите или выберите **Сервер разработчиков**.</span><span class="sxs-lookup"><span data-stu-id="d684d-170">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span>

### <a name="sideload-your-add-in"></a><span data-ttu-id="d684d-171">Загрузка неопубликованной надстройки</span><span class="sxs-lookup"><span data-stu-id="d684d-171">Sideload your add-in</span></span>

1. <span data-ttu-id="d684d-172">Откройте [Office в Интернете.](https://office.live.com/)</span><span class="sxs-lookup"><span data-stu-id="d684d-172">Open [Office on the web](https://office.live.com/).</span></span>
2. <span data-ttu-id="d684d-173">Откройте новую книгу Excel.</span><span class="sxs-lookup"><span data-stu-id="d684d-173">Open a new Excel workbook.</span></span>
3. <span data-ttu-id="d684d-174">Откройте **вкладку Insert** на ленте и в разделе **Надстройки** выберите **надстройки Office.**</span><span class="sxs-lookup"><span data-stu-id="d684d-174">Open the **Insert** tab on the ribbon and, in the **Add-ins** section, choose **Office Add-ins**.</span></span>
4. <span data-ttu-id="d684d-175">В **диалоговом** окантовке Надстройки Office выберите вкладку **MY ADD-INS,** выберите **Управление** надстройками, а затем загрузите мои **надстройки.**</span><span class="sxs-lookup"><span data-stu-id="d684d-175">On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.</span></span>
    
    ![Диалоговое окно "Надстройки Office" с раскрывающимся меню в правом верхнем углу, в котором выделен пункт "Управление моими надстройками", а под ним — раскрывающийся список с пунктом "Отправить надстройку"](../images/office-add-ins-my-account.png)

5. <span data-ttu-id="d684d-177">**Найдите** файл манифеста надстройки и выберите **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="d684d-177">**Browse** to the add-in manifest file, and then select **Upload**.</span></span>
    
    ![Диалоговое окно отправки надстройки с кнопками "Обзор", "Отправить" и "Отмена"](../images/upload-add-in.png)

> [!NOTE]
> <span data-ttu-id="d684d-179">После загрузки в документ он будет оставаться в стороне при каждом открываемом документе.</span><span class="sxs-lookup"><span data-stu-id="d684d-179">Once you've sideloaded to the document, it will remain sideloaded each time you open the document.</span></span>

### <a name="start-debugging"></a><span data-ttu-id="d684d-180">Начало отладки</span><span class="sxs-lookup"><span data-stu-id="d684d-180">Start debugging</span></span>

1. <span data-ttu-id="d684d-181">Откройте средства разработчика в браузере.</span><span class="sxs-lookup"><span data-stu-id="d684d-181">Open developer tools in the browser.</span></span> <span data-ttu-id="d684d-182">Для Chrome и большинства браузеров F12 откроет средства разработчика.</span><span class="sxs-lookup"><span data-stu-id="d684d-182">For Chrome and most browsers F12 will open the developer tools.</span></span>
2. <span data-ttu-id="d684d-183">В средствах разработчика откройте исходный файл скрипта кода с помощью **Cmd+P** или **Ctrl+P** **(functions.js** **или functions.ts).**</span><span class="sxs-lookup"><span data-stu-id="d684d-183">In developer tools, open your source code script file using **Cmd+P** or **Ctrl+P** (**functions.js** or **functions.ts**).</span></span>
3. <span data-ttu-id="d684d-184">[Установите точку разрыва в](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) пользовательском коде источника функции.</span><span class="sxs-lookup"><span data-stu-id="d684d-184">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span> 

<span data-ttu-id="d684d-185">Если вам нужно изменить код, вы можете внести изменения в VS Code и сохранить изменения.</span><span class="sxs-lookup"><span data-stu-id="d684d-185">If you need to change the code you can make edits in VS Code and save the changes.</span></span> <span data-ttu-id="d684d-186">Обновите браузер, чтобы увидеть загруженные изменения.</span><span class="sxs-lookup"><span data-stu-id="d684d-186">Refresh the browser to see the changes loaded.</span></span>

## <a name="use-the-command-line-tools-to-debug"></a><span data-ttu-id="d684d-187">Отламывка с помощью средств командной строки</span><span class="sxs-lookup"><span data-stu-id="d684d-187">Use the command line tools to debug</span></span>

<span data-ttu-id="d684d-188">Если вы не используете VS Code, для запуска надстройки можно использовать командную строку (например, bash или PowerShell).</span><span class="sxs-lookup"><span data-stu-id="d684d-188">If you are not using VS Code, you can use the command line (such as bash, or PowerShell) to run your add-in.</span></span> <span data-ttu-id="d684d-189">Для отлаговки кода в Excel в Интернете необходимо использовать средства разработчика браузера.</span><span class="sxs-lookup"><span data-stu-id="d684d-189">You'll need to use the browser developer tools to debug your code in Excel on the web.</span></span> <span data-ttu-id="d684d-190">Отламывка настольной версии Excel с помощью командной строки не удается.</span><span class="sxs-lookup"><span data-stu-id="d684d-190">You cannot debug the desktop version of Excel using the command line.</span></span>

1. <span data-ttu-id="d684d-191">Из командной строки запустите, чтобы следить за изменениями кода и восстанавливать `npm run watch` их.</span><span class="sxs-lookup"><span data-stu-id="d684d-191">From the command line run `npm run watch` to watch for and rebuild when code changes occur.</span></span>
2. <span data-ttu-id="d684d-192">Откройте второе окно командной строки (первое будет заблокировано во время запуска часов.)</span><span class="sxs-lookup"><span data-stu-id="d684d-192">Open a second command line window (the first one will be blocked while running the watch.)</span></span>

3. <span data-ttu-id="d684d-193">Если вы хотите запустить надстройку в настольной версии Excel, запустите следующую команду</span><span class="sxs-lookup"><span data-stu-id="d684d-193">If you want to start your add-in in the desktop version of Excel, run the following command</span></span>
    
    `npm run start:desktop`
    
    <span data-ttu-id="d684d-194">Или если вы предпочитаете запускать надстройку в Excel в Интернете, запустите следующую команду</span><span class="sxs-lookup"><span data-stu-id="d684d-194">Or if you prefer to start your add-in in Excel on the web run the following command</span></span>
    
    `npm run start:web`
    
    <span data-ttu-id="d684d-195">Для Excel в Интернете также необходимо побокзагружать надстройку.</span><span class="sxs-lookup"><span data-stu-id="d684d-195">For Excel on the web you also need to sideload your add-in.</span></span> <span data-ttu-id="d684d-196">Выполните действия [в Sideload надстройки,](#sideload-your-add-in) чтобы побокзагрузить надстройку.</span><span class="sxs-lookup"><span data-stu-id="d684d-196">Follow the steps in [Sideload your add-in](#sideload-your-add-in) to sideload your add-in.</span></span> <span data-ttu-id="d684d-197">Затем продолжайте отладку в следующем разделе.</span><span class="sxs-lookup"><span data-stu-id="d684d-197">Then continue to the next section to start debugging.</span></span>
    
4. <span data-ttu-id="d684d-198">Откройте средства разработчика в браузере.</span><span class="sxs-lookup"><span data-stu-id="d684d-198">Open developer tools in the browser.</span></span> <span data-ttu-id="d684d-199">Для Chrome и большинства браузеров F12 откроет средства разработчика.</span><span class="sxs-lookup"><span data-stu-id="d684d-199">For Chrome and most browsers F12 will open the developer tools.</span></span>
5. <span data-ttu-id="d684d-200">В средствах разработчика откройте исходный файл скрипта кода **(functions.js** **или functions.ts).**</span><span class="sxs-lookup"><span data-stu-id="d684d-200">In developer tools, open your source code script file (**functions.js** or **functions.ts**).</span></span> <span data-ttu-id="d684d-201">Пользовательский код функций может быть расположен в конце файла.</span><span class="sxs-lookup"><span data-stu-id="d684d-201">Your custom functions code may be located near the end of the file.</span></span>
6. <span data-ttu-id="d684d-202">В пользовательском коде источника функции нанесите точку разрыва, выбрав строку кода.</span><span class="sxs-lookup"><span data-stu-id="d684d-202">In the custom function source code, apply a breakpoint by selecting a line of code.</span></span>

<span data-ttu-id="d684d-203">Если вам нужно изменить код, вы можете внести изменения в Visual Studio и сохранить изменения.</span><span class="sxs-lookup"><span data-stu-id="d684d-203">If you need to change the code you can make edits in Visual Studio and save the changes.</span></span> <span data-ttu-id="d684d-204">Обновите браузер, чтобы увидеть загруженные изменения.</span><span class="sxs-lookup"><span data-stu-id="d684d-204">Refresh the browser to see the changes loaded.</span></span>

### <a name="commands-for-building-and-running-your-add-in"></a><span data-ttu-id="d684d-205">Команды для создания и запуска надстройки</span><span class="sxs-lookup"><span data-stu-id="d684d-205">Commands for building and running your add-in</span></span>

<span data-ttu-id="d684d-206">Существует несколько задач сборки:</span><span class="sxs-lookup"><span data-stu-id="d684d-206">There are several build tasks available:</span></span>
- <span data-ttu-id="d684d-207">`npm run watch`: сборки для разработки и автоматическое восстановление при сэкономлении исходных файлов</span><span class="sxs-lookup"><span data-stu-id="d684d-207">`npm run watch`: builds for development and automatically rebuilds when a source file is saved</span></span>
- <span data-ttu-id="d684d-208">`npm run build-dev`: сборки для разработки один раз</span><span class="sxs-lookup"><span data-stu-id="d684d-208">`npm run build-dev`: builds for development once</span></span>
- <span data-ttu-id="d684d-209">`npm run build`: сборки для производства</span><span class="sxs-lookup"><span data-stu-id="d684d-209">`npm run build`: builds for production</span></span>
- <span data-ttu-id="d684d-210">`npm run dev-server`: запускает веб-сервер, используемый для разработки</span><span class="sxs-lookup"><span data-stu-id="d684d-210">`npm run dev-server`: runs the web server used for development</span></span>

<span data-ttu-id="d684d-211">Вы можете использовать следующие задачи для начала отладки на рабочем столе или в Интернете.</span><span class="sxs-lookup"><span data-stu-id="d684d-211">You can use the following tasks to start debugging on desktop or online.</span></span>
- <span data-ttu-id="d684d-212">`npm run start:desktop`: Запускает Excel на рабочем столе и заряжает надстройку.</span><span class="sxs-lookup"><span data-stu-id="d684d-212">`npm run start:desktop`: Starts Excel on desktop and sideloads your add-in.</span></span>
- <span data-ttu-id="d684d-213">`npm run start:web`: Запускает Excel в Интернете и перегружает надстройку.</span><span class="sxs-lookup"><span data-stu-id="d684d-213">`npm run start:web`: Starts Excel on the web and sideloads your add-in.</span></span>
- <span data-ttu-id="d684d-214">`npm run stop`: Останавливает Excel и отладку.</span><span class="sxs-lookup"><span data-stu-id="d684d-214">`npm run stop`: Stops Excel and debugging.</span></span>

## <a name="next-steps"></a><span data-ttu-id="d684d-215">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="d684d-215">Next steps</span></span>
<span data-ttu-id="d684d-216">Узнайте о [практике проверки подлинности для пользовательских функций без пользовательского интерфейса.](custom-functions-authentication.md)</span><span class="sxs-lookup"><span data-stu-id="d684d-216">Learn about [authentication practices for UI-less custom functions](custom-functions-authentication.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="d684d-217">См. также</span><span class="sxs-lookup"><span data-stu-id="d684d-217">See also</span></span>

* [<span data-ttu-id="d684d-218">Устранение неполадок пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="d684d-218">Custom functions troubleshooting</span></span>](custom-functions-troubleshooting.md)
* [<span data-ttu-id="d684d-219">Обработка ошибок в пользовательских функциях Excel</span><span class="sxs-lookup"><span data-stu-id="d684d-219">Error handling for custom functions in Excel</span></span>](custom-functions-errors.md)
* [<span data-ttu-id="d684d-220">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="d684d-220">Create custom functions in Excel</span></span>](custom-functions-overview.md)
