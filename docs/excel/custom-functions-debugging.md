---
ms.date: 04/09/2021
description: Узнайте, как отлагировать настраиваемые функции Excel, которые не используют области задач.
title: Отладка пользовательских функций без пользовательского интерфейса
localization_priority: Normal
ms.openlocfilehash: 5b27ca44dbb891c2e1f4ae86175595dc902b74ba
ms.sourcegitcommit: 094caf086c2696e78fbdfdc6030cb0c89d32b585
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/16/2021
ms.locfileid: "51862339"
---
# <a name="ui-less-custom-functions-debugging"></a><span data-ttu-id="c08b7-103">Отладка пользовательских функций без пользовательского интерфейса</span><span class="sxs-lookup"><span data-stu-id="c08b7-103">UI-less custom functions debugging</span></span>

<span data-ttu-id="c08b7-104">В этой статье обсуждается  отладка только для настраиваемой функции, которая не использует области задач или другие элементы пользовательского интерфейса (пользовательские функции без пользовательского интерфейса).</span><span class="sxs-lookup"><span data-stu-id="c08b7-104">This article discusses debugging *only* for custom functions that don't use a task pane or other user interface elements (UI-less custom functions).</span></span> 

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

<span data-ttu-id="c08b7-105">В Windows:</span><span class="sxs-lookup"><span data-stu-id="c08b7-105">On Windows:</span></span>
- [<span data-ttu-id="c08b7-106">Отладка Visual Studio и кода Excel</span><span class="sxs-lookup"><span data-stu-id="c08b7-106">Excel Desktop and Visual Studio Code (VS Code) debugger</span></span>](#use-the-vs-code-debugger-for-excel-desktop)
- [<span data-ttu-id="c08b7-107">Excel в Интернете и отладка кода VS</span><span class="sxs-lookup"><span data-stu-id="c08b7-107">Excel on the web and VS Code debugger</span></span>](#use-the-vs-code-debugger-for-excel-in-microsoft-edge)
- [<span data-ttu-id="c08b7-108">Excel в веб-средствах и средствах браузера</span><span class="sxs-lookup"><span data-stu-id="c08b7-108">Excel on the web and browser tools</span></span>](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [<span data-ttu-id="c08b7-109">Командная строка</span><span class="sxs-lookup"><span data-stu-id="c08b7-109">Command line</span></span>](#use-the-command-line-tools-to-debug)

<span data-ttu-id="c08b7-110">На Mac:</span><span class="sxs-lookup"><span data-stu-id="c08b7-110">On Mac:</span></span>
- [<span data-ttu-id="c08b7-111">Excel в веб-средствах и средствах браузера</span><span class="sxs-lookup"><span data-stu-id="c08b7-111">Excel on the web and browser tools</span></span>](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [<span data-ttu-id="c08b7-112">Командная строка</span><span class="sxs-lookup"><span data-stu-id="c08b7-112">Command line</span></span>](#use-the-command-line-tools-to-debug)

> [!NOTE]
> <span data-ttu-id="c08b7-113">Для простоты в этой статье показана отладка в контексте использования Visual Studio кода для редактирования, выполнения задач и в некоторых случаях использования представления отладки.</span><span class="sxs-lookup"><span data-stu-id="c08b7-113">For simplicity, this article shows debugging in the context of using Visual Studio Code to edit, run tasks, and in some cases use the debug view.</span></span> <span data-ttu-id="c08b7-114">При использовании другого средства редактора или [](#commands-for-building-and-running-your-add-in) командной строки см. инструкции по командной строке в конце этой статьи.</span><span class="sxs-lookup"><span data-stu-id="c08b7-114">If you are using a different editor or command line tool, see the [command line instructions](#commands-for-building-and-running-your-add-in) at the end of this article.</span></span>

## <a name="requirements"></a><span data-ttu-id="c08b7-115">Требования</span><span class="sxs-lookup"><span data-stu-id="c08b7-115">Requirements</span></span>

<span data-ttu-id="c08b7-116">Этот процесс отладки работает **только** для пользовательских функций без пользовательского интерфейса, которые не используют области задач или другие элементы пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="c08b7-116">This debugging process works **only** for UI-less custom functions, which don't use a task pane or other UI elements.</span></span> <span data-ttu-id="c08b7-117">Настраиваемая функция без пользовательского интерфейса может быть создана, следуя шагам в руководстве Create custom [functions in Excel,](../tutorials/excel-tutorial-create-custom-functions.md) а затем удалив все элементы области задач и пользовательского интерфейса, установленные генератором [Yeoman](https://www.npmjs.com/package/generator-office)для надстроек Office.</span><span class="sxs-lookup"><span data-stu-id="c08b7-117">A UI-less custom function can be created by following the steps in the [Create custom functions in Excel](../tutorials/excel-tutorial-create-custom-functions.md) tutorial, and then removing all of the task pane and UI elements that are installed by the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office).</span></span>

<span data-ttu-id="c08b7-118">Обратите внимание, что этот процесс отладки не совместим с пользовательскими проектами функций с помощью общего [времени запуска.](../develop/configure-your-add-in-to-use-a-shared-runtime.md)</span><span class="sxs-lookup"><span data-stu-id="c08b7-118">Note that this debugging process is not compatible with custom functions projects using a [shared runtime](../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

## <a name="use-the-vs-code-debugger-for-excel-desktop"></a><span data-ttu-id="c08b7-119">Использование отладки кода VS для настольных компьютеров Excel</span><span class="sxs-lookup"><span data-stu-id="c08b7-119">Use the VS Code debugger for Excel Desktop</span></span>

<span data-ttu-id="c08b7-120">Вы можете использовать VS Code для отлаговки пользовательских функций без пользовательского интерфейса в Office Excel на рабочем столе.</span><span class="sxs-lookup"><span data-stu-id="c08b7-120">You can use VS Code to debug UI-less custom functions in Office Excel on the desktop.</span></span>

> [!NOTE]
> <span data-ttu-id="c08b7-121">Отладка рабочего стола для Mac недоступна, но может быть достигнута с помощью средств браузера и командной строки для отладки [Excel в Интернете).](#use-the-command-line-tools-to-debug)</span><span class="sxs-lookup"><span data-stu-id="c08b7-121">Desktop debugging for the Mac is not available but can be achieved [using the browser tools and command line to debug Excel on the web](#use-the-command-line-tools-to-debug)).</span></span>

### <a name="run-your-add-in-from-vs-code"></a><span data-ttu-id="c08b7-122">Запуск надстройки из VS Code</span><span class="sxs-lookup"><span data-stu-id="c08b7-122">Run your add-in from VS Code</span></span>

1. <span data-ttu-id="c08b7-123">Откройте настраиваемую папку корневого проекта функций [в VS Code.](https://code.visualstudio.com/)</span><span class="sxs-lookup"><span data-stu-id="c08b7-123">Open your custom functions root project folder in [VS Code](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="c08b7-124">Выберите **терминал > выполнить задачу и** введите или выберите **Часы**.</span><span class="sxs-lookup"><span data-stu-id="c08b7-124">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="c08b7-125">Это позволит отслеживать и восстанавливать любые изменения файлов.</span><span class="sxs-lookup"><span data-stu-id="c08b7-125">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="c08b7-126">Выберите **терминальный > выполнить задачу и** введите или выберите **Сервер разработчиков**.</span><span class="sxs-lookup"><span data-stu-id="c08b7-126">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span>

### <a name="start-the-vs-code-debugger"></a><span data-ttu-id="c08b7-127">Запуск отладки кода VS</span><span class="sxs-lookup"><span data-stu-id="c08b7-127">Start the VS Code debugger</span></span>

4. <span data-ttu-id="c08b7-128">Выберите **просмотр > выполнить** или ввести **Ctrl+Shift+D,** чтобы перейти на отлагивание представления.</span><span class="sxs-lookup"><span data-stu-id="c08b7-128">Choose **View > Run** or enter **Ctrl+Shift+D** to switch to debug view.</span></span>
5. <span data-ttu-id="c08b7-129">Из выпадаемого меню Run выберите **Excel Desktop (Edge Chromium).**</span><span class="sxs-lookup"><span data-stu-id="c08b7-129">From the Run drop-down menu, choose **Excel Desktop (Edge Chromium)**.</span></span>
6. <span data-ttu-id="c08b7-130">Чтобы начать отладку, выберите **F5** **(или > запустить** отладку из меню).</span><span class="sxs-lookup"><span data-stu-id="c08b7-130">Select **F5** (or select **Run -> Start Debugging** from the menu) to begin debugging.</span></span> <span data-ttu-id="c08b7-131">Новая книга Excel откроется с уже загруженной и готовой к использованию надстройке.</span><span class="sxs-lookup"><span data-stu-id="c08b7-131">A new Excel workbook will open with your add-in already sideloaded and ready to use.</span></span>

### <a name="start-debugging"></a><span data-ttu-id="c08b7-132">Начало отладки</span><span class="sxs-lookup"><span data-stu-id="c08b7-132">Start debugging</span></span>

1. <span data-ttu-id="c08b7-133">В vs Code откройте исходный файл скрипта кода **(functions.js** **или functions.ts).**</span><span class="sxs-lookup"><span data-stu-id="c08b7-133">In VS Code, open your source code script file (**functions.js** or **functions.ts**).</span></span>
2. <span data-ttu-id="c08b7-134">[Установите точку разрыва в](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) пользовательском коде источника функции.</span><span class="sxs-lookup"><span data-stu-id="c08b7-134">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span>
3. <span data-ttu-id="c08b7-135">В книге Excel введите формулу, использующую настраиваемую функцию.</span><span class="sxs-lookup"><span data-stu-id="c08b7-135">In the Excel workbook, enter a formula that uses your custom function.</span></span>

<span data-ttu-id="c08b7-136">На этом этапе выполнение остановится на строке кода, где установлена точка разрыва.</span><span class="sxs-lookup"><span data-stu-id="c08b7-136">At this point execution will stop on the line of code where you set the breakpoint.</span></span> <span data-ttu-id="c08b7-137">Теперь вы можете пройти через код, установить часы и использовать все необходимые функции отладки кода VS.</span><span class="sxs-lookup"><span data-stu-id="c08b7-137">Now you can step through your code, set watches, and use any VS Code debugging features you need.</span></span>

## <a name="use-the-vs-code-debugger-for-excel-in-microsoft-edge"></a><span data-ttu-id="c08b7-138">Использование отладки кода VS для Excel в Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="c08b7-138">Use the VS Code debugger for Excel in Microsoft Edge</span></span>

<span data-ttu-id="c08b7-139">Вы можете использовать VS Code для отлаговки пользовательских функций в Excel в браузере Microsoft Edge.</span><span class="sxs-lookup"><span data-stu-id="c08b7-139">You can use VS Code to debug UI-less custom functions in Excel on the Microsoft Edge browser.</span></span> <span data-ttu-id="c08b7-140">Чтобы использовать vs Code с Microsoft Edge, необходимо установить расширение [Debugger для Microsoft Edge.](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge)</span><span class="sxs-lookup"><span data-stu-id="c08b7-140">To use VS Code with Microsoft Edge, you must install the [Debugger for Microsoft Edge](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge) extension.</span></span>

### <a name="run-your-add-in-from-vs-code"></a><span data-ttu-id="c08b7-141">Запуск надстройки из VS Code</span><span class="sxs-lookup"><span data-stu-id="c08b7-141">Run your add-in from VS Code</span></span>

1. <span data-ttu-id="c08b7-142">Откройте настраиваемую папку корневого проекта функций [в VS Code.](https://code.visualstudio.com/)</span><span class="sxs-lookup"><span data-stu-id="c08b7-142">Open your custom functions root project folder in [VS Code](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="c08b7-143">Выберите **терминал > выполнить задачу и** введите или выберите **Часы**.</span><span class="sxs-lookup"><span data-stu-id="c08b7-143">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="c08b7-144">Это позволит отслеживать и восстанавливать любые изменения файлов.</span><span class="sxs-lookup"><span data-stu-id="c08b7-144">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="c08b7-145">Выберите **терминальный > выполнить задачу и** введите или выберите **Сервер разработчиков**.</span><span class="sxs-lookup"><span data-stu-id="c08b7-145">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span>

### <a name="start-the-vs-code-debugger"></a><span data-ttu-id="c08b7-146">Запуск отладки кода VS</span><span class="sxs-lookup"><span data-stu-id="c08b7-146">Start the VS Code debugger</span></span>

4. <span data-ttu-id="c08b7-147">Выберите **просмотр > выполнить** или ввести **Ctrl+Shift+D,** чтобы перейти на отлагивание представления.</span><span class="sxs-lookup"><span data-stu-id="c08b7-147">Choose **View > Run** or enter **Ctrl+Shift+D** to switch to debug view.</span></span>
5. <span data-ttu-id="c08b7-148">Из параметров отладки выберите **Office Online (Edge Chromium).**</span><span class="sxs-lookup"><span data-stu-id="c08b7-148">From the Debug options, choose **Office Online (Edge Chromium)**.</span></span>
6. <span data-ttu-id="c08b7-149">Откройте Excel в браузере Microsoft Edge и создайте новую книгу.</span><span class="sxs-lookup"><span data-stu-id="c08b7-149">Open Excel in the Microsoft Edge browser and create a new workbook.</span></span>
7. <span data-ttu-id="c08b7-150">Выберите **Share** в ленте и скопируйте ссылку на URL-адрес этой новой книги.</span><span class="sxs-lookup"><span data-stu-id="c08b7-150">Choose **Share** in the ribbon and copy the link for the URL for this new workbook.</span></span>
8. <span data-ttu-id="c08b7-151">Чтобы начать отладку, выберите **F5** (> **запустить** отладку из меню).</span><span class="sxs-lookup"><span data-stu-id="c08b7-151">Select **F5** (or select **Run > Start Debugging** from the menu) to begin debugging.</span></span> <span data-ttu-id="c08b7-152">Появится запрос, в котором будет указан URL-адрес документа.</span><span class="sxs-lookup"><span data-stu-id="c08b7-152">A prompt will appear, which asks for the URL of your document.</span></span>
9. <span data-ttu-id="c08b7-153">Введите URL-адрес книги и нажмите кнопку Ввод.</span><span class="sxs-lookup"><span data-stu-id="c08b7-153">Paste in the URL for your workbook and press Enter.</span></span>

### <a name="sideload-your-add-in"></a><span data-ttu-id="c08b7-154">Загрузка неопубликованной надстройки</span><span class="sxs-lookup"><span data-stu-id="c08b7-154">Sideload your add-in</span></span>

1. <span data-ttu-id="c08b7-155">Выберите **вкладку Insert** на ленте и в разделе Надстройки, выберите  **надстройки Office.**</span><span class="sxs-lookup"><span data-stu-id="c08b7-155">Select the **Insert** tab on the ribbon and in the **Add-ins** section, choose **Office Add-ins**.</span></span>
2. <span data-ttu-id="c08b7-156">В **диалоговом** окантовке Надстройки Office выберите вкладку **MY ADD-INS,** выберите **Управление** надстройками, а затем загрузите мои **надстройки.**</span><span class="sxs-lookup"><span data-stu-id="c08b7-156">On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.</span></span>
    
    ![Диалоговое окно "Надстройки Office" с раскрывающимся меню в правом верхнем углу, в котором выделен пункт "Управление моими надстройками", а под ним — раскрывающийся список с пунктом "Отправить надстройку"](../images/office-add-ins-my-account.png)

3. <span data-ttu-id="c08b7-158">**Просмотрите** файл манифеста надстройки и выберите **Upload**.</span><span class="sxs-lookup"><span data-stu-id="c08b7-158">**Browse** to the add-in manifest file and then select **Upload**.</span></span>
    
    ![Диалоговое окно отправки надстройки с кнопками "Обзор", "Отправить" и "Отмена"](../images/upload-add-in.png)


### <a name="set-breakpoints"></a><span data-ttu-id="c08b7-160">Настройка точек разрыва</span><span class="sxs-lookup"><span data-stu-id="c08b7-160">Set breakpoints</span></span>
1. <span data-ttu-id="c08b7-161">В vs Code откройте исходный файл скрипта кода **(functions.js** **или functions.ts).**</span><span class="sxs-lookup"><span data-stu-id="c08b7-161">In VS Code, open your source code script file (**functions.js** or **functions.ts**).</span></span>
2. <span data-ttu-id="c08b7-162">[Установите точку разрыва в](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) пользовательском коде источника функции.</span><span class="sxs-lookup"><span data-stu-id="c08b7-162">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span>
3. <span data-ttu-id="c08b7-163">В книге Excel введите формулу, использующую настраиваемую функцию.</span><span class="sxs-lookup"><span data-stu-id="c08b7-163">In the Excel workbook, enter a formula that uses your custom function.</span></span>

## <a name="use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web"></a><span data-ttu-id="c08b7-164">Использование средств разработчика браузера для отлаговки пользовательских функций в Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="c08b7-164">Use the browser developer tools to debug custom functions in Excel on the web</span></span>

<span data-ttu-id="c08b7-165">Средства разработчика браузера можно использовать для отлаговки пользовательских функций в Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="c08b7-165">You can use the browser developer tools to debug UI-less custom functions in Excel on the web.</span></span> <span data-ttu-id="c08b7-166">Следующие действия работают как для Windows, так и для macOS.</span><span class="sxs-lookup"><span data-stu-id="c08b7-166">The following steps work for both Windows and macOS.</span></span>

### <a name="run-your-add-in-from-visual-studio-code"></a><span data-ttu-id="c08b7-167">Запустите надстройку из Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="c08b7-167">Run your add-in from Visual Studio Code</span></span>

1. <span data-ttu-id="c08b7-168">Откройте настраиваемую папку корневого проекта [функций в Visual Studio Code (VS Code).](https://code.visualstudio.com/)</span><span class="sxs-lookup"><span data-stu-id="c08b7-168">Open your custom functions root project folder in [Visual Studio Code (VS Code)](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="c08b7-169">Выберите **терминал > выполнить задачу и** введите или выберите **Часы**.</span><span class="sxs-lookup"><span data-stu-id="c08b7-169">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="c08b7-170">Это позволит отслеживать и восстанавливать любые изменения файлов.</span><span class="sxs-lookup"><span data-stu-id="c08b7-170">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="c08b7-171">Выберите **терминальный > выполнить задачу и** введите или выберите **Сервер разработчиков**.</span><span class="sxs-lookup"><span data-stu-id="c08b7-171">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span>

### <a name="sideload-your-add-in"></a><span data-ttu-id="c08b7-172">Загрузка неопубликованной надстройки</span><span class="sxs-lookup"><span data-stu-id="c08b7-172">Sideload your add-in</span></span>

1. <span data-ttu-id="c08b7-173">Откройте [Office в Интернете.](https://office.live.com/)</span><span class="sxs-lookup"><span data-stu-id="c08b7-173">Open [Office on the web](https://office.live.com/).</span></span>
2. <span data-ttu-id="c08b7-174">Откройте новую книгу Excel.</span><span class="sxs-lookup"><span data-stu-id="c08b7-174">Open a new Excel workbook.</span></span>
3. <span data-ttu-id="c08b7-175">Откройте **вкладку Insert** на ленте и в разделе **Надстройки** выберите **надстройки Office.**</span><span class="sxs-lookup"><span data-stu-id="c08b7-175">Open the **Insert** tab on the ribbon and, in the **Add-ins** section, choose **Office Add-ins**.</span></span>
4. <span data-ttu-id="c08b7-176">В **диалоговом** окантовке Надстройки Office выберите вкладку **MY ADD-INS,** выберите **Управление** надстройками, а затем загрузите мои **надстройки.**</span><span class="sxs-lookup"><span data-stu-id="c08b7-176">On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.</span></span>
    
    ![Диалоговое окно "Надстройки Office" с раскрывающимся меню в правом верхнем углу, в котором выделен пункт "Управление моими надстройками", а под ним — раскрывающийся список с пунктом "Отправить надстройку"](../images/office-add-ins-my-account.png)

5. <span data-ttu-id="c08b7-178">**Найдите** файл манифеста надстройки и выберите **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="c08b7-178">**Browse** to the add-in manifest file, and then select **Upload**.</span></span>
    
    ![Диалоговое окно отправки надстройки с кнопками "Обзор", "Отправить" и "Отмена"](../images/upload-add-in.png)

> [!NOTE]
> <span data-ttu-id="c08b7-180">После загрузки в документ он будет оставаться в стороне при каждом открываемом документе.</span><span class="sxs-lookup"><span data-stu-id="c08b7-180">Once you've sideloaded to the document, it will remain sideloaded each time you open the document.</span></span>

### <a name="start-debugging"></a><span data-ttu-id="c08b7-181">Начало отладки</span><span class="sxs-lookup"><span data-stu-id="c08b7-181">Start debugging</span></span>

1. <span data-ttu-id="c08b7-182">Откройте средства разработчика в браузере.</span><span class="sxs-lookup"><span data-stu-id="c08b7-182">Open developer tools in the browser.</span></span> <span data-ttu-id="c08b7-183">Для Chrome и большинства браузеров F12 откроет средства разработчика.</span><span class="sxs-lookup"><span data-stu-id="c08b7-183">For Chrome and most browsers F12 will open the developer tools.</span></span>
2. <span data-ttu-id="c08b7-184">В средствах разработчика откройте исходный файл скрипта кода с помощью **Cmd+P** или **Ctrl+P** **(functions.js** **или functions.ts).**</span><span class="sxs-lookup"><span data-stu-id="c08b7-184">In developer tools, open your source code script file using **Cmd+P** or **Ctrl+P** (**functions.js** or **functions.ts**).</span></span>
3. <span data-ttu-id="c08b7-185">[Установите точку разрыва в](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) пользовательском коде источника функции.</span><span class="sxs-lookup"><span data-stu-id="c08b7-185">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span> 

<span data-ttu-id="c08b7-186">Если вам нужно изменить код, вы можете внести изменения в VS Code и сохранить изменения.</span><span class="sxs-lookup"><span data-stu-id="c08b7-186">If you need to change the code you can make edits in VS Code and save the changes.</span></span> <span data-ttu-id="c08b7-187">Обновите браузер, чтобы увидеть загруженные изменения.</span><span class="sxs-lookup"><span data-stu-id="c08b7-187">Refresh the browser to see the changes loaded.</span></span>

## <a name="use-the-command-line-tools-to-debug"></a><span data-ttu-id="c08b7-188">Отламывка с помощью средств командной строки</span><span class="sxs-lookup"><span data-stu-id="c08b7-188">Use the command line tools to debug</span></span>

<span data-ttu-id="c08b7-189">Если вы не используете VS Code, для запуска надстройки можно использовать командную строку (например, bash или PowerShell).</span><span class="sxs-lookup"><span data-stu-id="c08b7-189">If you are not using VS Code, you can use the command line (such as bash, or PowerShell) to run your add-in.</span></span> <span data-ttu-id="c08b7-190">Для отлаговки кода в Excel в Интернете необходимо использовать средства разработчика браузера.</span><span class="sxs-lookup"><span data-stu-id="c08b7-190">You'll need to use the browser developer tools to debug your code in Excel on the web.</span></span> <span data-ttu-id="c08b7-191">Отламывка настольной версии Excel с помощью командной строки не удается.</span><span class="sxs-lookup"><span data-stu-id="c08b7-191">You cannot debug the desktop version of Excel using the command line.</span></span>

1. <span data-ttu-id="c08b7-192">Из командной строки запустите, чтобы следить за изменениями кода и восстанавливать `npm run watch` их.</span><span class="sxs-lookup"><span data-stu-id="c08b7-192">From the command line run `npm run watch` to watch for and rebuild when code changes occur.</span></span>
2. <span data-ttu-id="c08b7-193">Откройте второе окно командной строки (первое будет заблокировано во время запуска часов.)</span><span class="sxs-lookup"><span data-stu-id="c08b7-193">Open a second command line window (the first one will be blocked while running the watch.)</span></span>

3. <span data-ttu-id="c08b7-194">Если вы хотите запустить надстройку в настольной версии Excel, запустите следующую команду</span><span class="sxs-lookup"><span data-stu-id="c08b7-194">If you want to start your add-in in the desktop version of Excel, run the following command</span></span>
    
    `npm run start:desktop`
    
    <span data-ttu-id="c08b7-195">Или если вы предпочитаете запускать надстройку в Excel в Интернете, запустите следующую команду</span><span class="sxs-lookup"><span data-stu-id="c08b7-195">Or if you prefer to start your add-in in Excel on the web run the following command</span></span>
    
    `npm run start:web`
    
    <span data-ttu-id="c08b7-196">Для Excel в Интернете также необходимо побокзагружать надстройку.</span><span class="sxs-lookup"><span data-stu-id="c08b7-196">For Excel on the web you also need to sideload your add-in.</span></span> <span data-ttu-id="c08b7-197">Выполните действия [в Sideload надстройки,](#sideload-your-add-in) чтобы побокзагрузить надстройку.</span><span class="sxs-lookup"><span data-stu-id="c08b7-197">Follow the steps in [Sideload your add-in](#sideload-your-add-in) to sideload your add-in.</span></span> <span data-ttu-id="c08b7-198">Затем продолжайте отладку в следующем разделе.</span><span class="sxs-lookup"><span data-stu-id="c08b7-198">Then continue to the next section to start debugging.</span></span>
    
4. <span data-ttu-id="c08b7-199">Откройте средства разработчика в браузере.</span><span class="sxs-lookup"><span data-stu-id="c08b7-199">Open developer tools in the browser.</span></span> <span data-ttu-id="c08b7-200">Для Chrome и большинства браузеров F12 откроет средства разработчика.</span><span class="sxs-lookup"><span data-stu-id="c08b7-200">For Chrome and most browsers F12 will open the developer tools.</span></span>
5. <span data-ttu-id="c08b7-201">В средствах разработчика откройте исходный файл скрипта кода **(functions.js** **или functions.ts).**</span><span class="sxs-lookup"><span data-stu-id="c08b7-201">In developer tools, open your source code script file (**functions.js** or **functions.ts**).</span></span> <span data-ttu-id="c08b7-202">Пользовательский код функций может быть расположен в конце файла.</span><span class="sxs-lookup"><span data-stu-id="c08b7-202">Your custom functions code may be located near the end of the file.</span></span>
6. <span data-ttu-id="c08b7-203">В пользовательском коде источника функции нанесите точку разрыва, выбрав строку кода.</span><span class="sxs-lookup"><span data-stu-id="c08b7-203">In the custom function source code, apply a breakpoint by selecting a line of code.</span></span>

<span data-ttu-id="c08b7-204">Если вам нужно изменить код, вы можете внести изменения в Visual Studio и сохранить изменения.</span><span class="sxs-lookup"><span data-stu-id="c08b7-204">If you need to change the code you can make edits in Visual Studio and save the changes.</span></span> <span data-ttu-id="c08b7-205">Обновите браузер, чтобы увидеть загруженные изменения.</span><span class="sxs-lookup"><span data-stu-id="c08b7-205">Refresh the browser to see the changes loaded.</span></span>

### <a name="commands-for-building-and-running-your-add-in"></a><span data-ttu-id="c08b7-206">Команды для создания и запуска надстройки</span><span class="sxs-lookup"><span data-stu-id="c08b7-206">Commands for building and running your add-in</span></span>

<span data-ttu-id="c08b7-207">Существует несколько задач сборки:</span><span class="sxs-lookup"><span data-stu-id="c08b7-207">There are several build tasks available:</span></span>
- <span data-ttu-id="c08b7-208">`npm run watch`: сборки для разработки и автоматическое восстановление при сэкономлении исходных файлов</span><span class="sxs-lookup"><span data-stu-id="c08b7-208">`npm run watch`: builds for development and automatically rebuilds when a source file is saved</span></span>
- <span data-ttu-id="c08b7-209">`npm run build-dev`: сборки для разработки один раз</span><span class="sxs-lookup"><span data-stu-id="c08b7-209">`npm run build-dev`: builds for development once</span></span>
- <span data-ttu-id="c08b7-210">`npm run build`: сборки для производства</span><span class="sxs-lookup"><span data-stu-id="c08b7-210">`npm run build`: builds for production</span></span>
- <span data-ttu-id="c08b7-211">`npm run dev-server`: запускает веб-сервер, используемый для разработки</span><span class="sxs-lookup"><span data-stu-id="c08b7-211">`npm run dev-server`: runs the web server used for development</span></span>

<span data-ttu-id="c08b7-212">Вы можете использовать следующие задачи для начала отладки на рабочем столе или в Интернете.</span><span class="sxs-lookup"><span data-stu-id="c08b7-212">You can use the following tasks to start debugging on desktop or online.</span></span>
- <span data-ttu-id="c08b7-213">`npm run start:desktop`: Запускает Excel на рабочем столе и заряжает надстройку.</span><span class="sxs-lookup"><span data-stu-id="c08b7-213">`npm run start:desktop`: Starts Excel on desktop and sideloads your add-in.</span></span>
- <span data-ttu-id="c08b7-214">`npm run start:web`: Запускает Excel в Интернете и перегружает надстройку.</span><span class="sxs-lookup"><span data-stu-id="c08b7-214">`npm run start:web`: Starts Excel on the web and sideloads your add-in.</span></span>
- <span data-ttu-id="c08b7-215">`npm run stop`: Останавливает Excel и отладку.</span><span class="sxs-lookup"><span data-stu-id="c08b7-215">`npm run stop`: Stops Excel and debugging.</span></span>

## <a name="next-steps"></a><span data-ttu-id="c08b7-216">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="c08b7-216">Next steps</span></span>
<span data-ttu-id="c08b7-217">Узнайте о [практике проверки подлинности для пользовательских функций без пользовательского интерфейса.](custom-functions-authentication.md)</span><span class="sxs-lookup"><span data-stu-id="c08b7-217">Learn about [authentication practices for UI-less custom functions](custom-functions-authentication.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="c08b7-218">См. также</span><span class="sxs-lookup"><span data-stu-id="c08b7-218">See also</span></span>

* [<span data-ttu-id="c08b7-219">Устранение неполадок пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="c08b7-219">Custom functions troubleshooting</span></span>](custom-functions-troubleshooting.md)
* [<span data-ttu-id="c08b7-220">Обработка ошибок в пользовательских функциях Excel</span><span class="sxs-lookup"><span data-stu-id="c08b7-220">Error handling for custom functions in Excel</span></span>](custom-functions-errors.md)
* [<span data-ttu-id="c08b7-221">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="c08b7-221">Create custom functions in Excel</span></span>](custom-functions-overview.md)
