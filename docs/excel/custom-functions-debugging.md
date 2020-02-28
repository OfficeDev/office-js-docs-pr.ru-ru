---
ms.date: 07/10/2019
description: Отладка пользовательских функций в Excel.
title: Отладка пользовательских функций
localization_priority: Normal
ms.openlocfilehash: dc620d8bab50c5efb3b9d9ec4f79f6532605f48b
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324626"
---
# <a name="custom-functions-debugging"></a><span data-ttu-id="dbf49-103">Отладка пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="dbf49-103">Custom functions debugging</span></span>

<span data-ttu-id="dbf49-104">Отладка настраиваемых функций может осуществляться несколькими способами, в зависимости от используемой платформы.</span><span class="sxs-lookup"><span data-stu-id="dbf49-104">Debugging for custom functions can be accomplished by multiple means, depending on what platform you're using.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="dbf49-105">В Windows:</span><span class="sxs-lookup"><span data-stu-id="dbf49-105">On Windows:</span></span>
- [<span data-ttu-id="dbf49-106">Отладчик Excel для настольных ПК и Visual Studio Code (VS Code)</span><span class="sxs-lookup"><span data-stu-id="dbf49-106">Excel Desktop and Visual Studio Code (VS Code) debugger</span></span>](#use-the-vs-code-debugger-for-excel-desktop)
- [<span data-ttu-id="dbf49-107">Приложение Excel в отладчике кода для Интернета и VS</span><span class="sxs-lookup"><span data-stu-id="dbf49-107">Excel on the web and VS Code debugger</span></span>](#use-the-vs-code-debugger-for-excel-in-microsoft-edge)
- [<span data-ttu-id="dbf49-108">Excel в веб-средствах и веб-браузерах</span><span class="sxs-lookup"><span data-stu-id="dbf49-108">Excel on the web and browser tools</span></span>](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [<span data-ttu-id="dbf49-109">Командная строка</span><span class="sxs-lookup"><span data-stu-id="dbf49-109">Command line</span></span>](#use-the-command-line-tools-to-debug)

<span data-ttu-id="dbf49-110">На компьютерах Mac:</span><span class="sxs-lookup"><span data-stu-id="dbf49-110">On Mac:</span></span>
- [<span data-ttu-id="dbf49-111">Excel в веб-средствах и веб-браузерах</span><span class="sxs-lookup"><span data-stu-id="dbf49-111">Excel on the web and browser tools</span></span>](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web)
- [<span data-ttu-id="dbf49-112">Командная строка</span><span class="sxs-lookup"><span data-stu-id="dbf49-112">Command line</span></span>](#use-the-command-line-tools-to-debug)

> [!NOTE]
> <span data-ttu-id="dbf49-113">Для простоты в этой статье показана Отладка в контексте использования Visual Studio Code для редактирования, запуска задач и в некоторых случаях использования представления отладки.</span><span class="sxs-lookup"><span data-stu-id="dbf49-113">For simplicity, this article shows debugging in the context of using Visual Studio Code to edit, run tasks, and in some cases use the debug view.</span></span> <span data-ttu-id="dbf49-114">Если вы используете другой редактор или средство командной строки, ознакомьтесь с [инструкциями по использованию командной строки](#commands-for-building-and-running-your-add-in) в конце этой статьи.</span><span class="sxs-lookup"><span data-stu-id="dbf49-114">If you are using a different editor or command line tool, see the [command line instructions](#commands-for-building-and-running-your-add-in) at the end of this article.</span></span>

## <a name="requirements"></a><span data-ttu-id="dbf49-115">Requirements</span><span class="sxs-lookup"><span data-stu-id="dbf49-115">Requirements</span></span>

<span data-ttu-id="dbf49-116">Перед началом отладки следует использовать [генератор Yeoman для надстроек Office](https://github.com/OfficeDev/generator-office) , чтобы создать проект пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="dbf49-116">Before starting to debug, you should use the [Yeoman generator for Office add-ins](https://github.com/OfficeDev/generator-office) to create a custom functions project.</span></span> <span data-ttu-id="dbf49-117">Руководство по созданию проекта пользовательских функций представлено в [руководстве Custom functions](../tutorials/excel-tutorial-create-custom-functions.md).</span><span class="sxs-lookup"><span data-stu-id="dbf49-117">For guidance about how to create a custom functions project, see the [custom functions tutorial](../tutorials/excel-tutorial-create-custom-functions.md).</span></span>

## <a name="use-the-vs-code-debugger-for-excel-desktop"></a><span data-ttu-id="dbf49-118">Использование отладчика кода VS для классической версии Excel</span><span class="sxs-lookup"><span data-stu-id="dbf49-118">Use the VS Code debugger for Excel Desktop</span></span>

<span data-ttu-id="dbf49-119">Вы можете использовать код VS для отладки настраиваемых функций в Office Excel на настольном компьютере.</span><span class="sxs-lookup"><span data-stu-id="dbf49-119">You can use VS Code to debug custom functions in Office Excel on the desktop.</span></span>

> [!NOTE]
> <span data-ttu-id="dbf49-120">Отладка на рабочем столе для Mac недоступна, но ее можно получить [с помощью средств браузера и командной строки для отладки Excel в Интернете](#use-the-command-line-tools-to-debug)).</span><span class="sxs-lookup"><span data-stu-id="dbf49-120">Desktop debugging for the Mac is not available but can be achieved [using the browser tools and command line to debug Excel on the web](#use-the-command-line-tools-to-debug)).</span></span>

### <a name="run-your-add-in-from-vs-code"></a><span data-ttu-id="dbf49-121">Запуск надстройки из кода VS</span><span class="sxs-lookup"><span data-stu-id="dbf49-121">Run your add-in from VS Code</span></span>

1. <span data-ttu-id="dbf49-122">Откройте корневую папку проекта "пользовательские функции" в [VS Code](https://code.visualstudio.com/).</span><span class="sxs-lookup"><span data-stu-id="dbf49-122">Open your custom functions root project folder in [VS Code](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="dbf49-123">Выберите пункт **терминал > выполнить задачу** и введите или выберите **Контрольное значение**.</span><span class="sxs-lookup"><span data-stu-id="dbf49-123">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="dbf49-124">В этом случае будут отслеживаться и перестраиваться все изменения файлов.</span><span class="sxs-lookup"><span data-stu-id="dbf49-124">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="dbf49-125">Выберите пункт **терминал > выполнить задачу** и введите или выберите **сервер разработки**.</span><span class="sxs-lookup"><span data-stu-id="dbf49-125">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span>

### <a name="start-the-vs-code-debugger"></a><span data-ttu-id="dbf49-126">Запуск отладчика кода VS</span><span class="sxs-lookup"><span data-stu-id="dbf49-126">Start the VS Code debugger</span></span>

4. <span data-ttu-id="dbf49-127">Нажмите кнопку **просмотр > Отладка** или введите **CTRL + SHIFT + D** , чтобы переключиться в представление отладки.</span><span class="sxs-lookup"><span data-stu-id="dbf49-127">Choose **View > Debug** or enter **Ctrl+Shift+D** to switch to debug view.</span></span>
5. <span data-ttu-id="dbf49-128">В разделе Параметры отладки выберите пункт **Рабочий стол Excel**.</span><span class="sxs-lookup"><span data-stu-id="dbf49-128">From the Debug options, choose **Excel Desktop**.</span></span>
6. <span data-ttu-id="dbf49-129">Нажмите **клавишу F5** (или выберите **Debug-> начать отладку** в меню), чтобы начать отладку.</span><span class="sxs-lookup"><span data-stu-id="dbf49-129">Select **F5** (or choose **Debug -> Start Debugging** from the menu) to begin debugging.</span></span> <span data-ttu-id="dbf49-130">Откроется новая книга Excel с уже неопубликованные и готовым к использованию надстройкой.</span><span class="sxs-lookup"><span data-stu-id="dbf49-130">A new Excel workbook will open with your add-in already sideloaded and ready to use.</span></span>

### <a name="start-debugging"></a><span data-ttu-id="dbf49-131">Начало отладки</span><span class="sxs-lookup"><span data-stu-id="dbf49-131">Start debugging</span></span>

1. <span data-ttu-id="dbf49-132">В коде VS откройте файл сценария исходного кода (**functions. js** или **functions. TS**).</span><span class="sxs-lookup"><span data-stu-id="dbf49-132">In VS Code, open your source code script file (**functions.js** or **functions.ts**).</span></span>
2. <span data-ttu-id="dbf49-133">[Задайте точку останова](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) в исходном коде пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="dbf49-133">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span>
3. <span data-ttu-id="dbf49-134">В книге Excel введите формулу, использующую пользовательскую функцию.</span><span class="sxs-lookup"><span data-stu-id="dbf49-134">In the Excel workbook, enter a formula that uses your custom function.</span></span>

<span data-ttu-id="dbf49-135">При этом выполнение будет остановлено в строке кода, в которой вы задаете точку останова.</span><span class="sxs-lookup"><span data-stu-id="dbf49-135">At this point execution will stop on the line of code where you set the breakpoint.</span></span> <span data-ttu-id="dbf49-136">Теперь вы можете выполнить отладку кода, задать контрольные значения и использовать любые необходимые возможности отладки кода VS.</span><span class="sxs-lookup"><span data-stu-id="dbf49-136">Now you can step through your code, set watches, and use any VS Code debugging features you need.</span></span>

## <a name="use-the-vs-code-debugger-for-excel-in-microsoft-edge"></a><span data-ttu-id="dbf49-137">Использование отладчика кода VS для Excel в Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="dbf49-137">Use the VS Code debugger for Excel in Microsoft Edge</span></span>

<span data-ttu-id="dbf49-138">Вы можете использовать код VS для отладки настраиваемых функций в Excel в браузере Microsoft Edge.</span><span class="sxs-lookup"><span data-stu-id="dbf49-138">You can use VS Code to debug custom functions in Excel on the Microsoft Edge browser.</span></span> <span data-ttu-id="dbf49-139">Чтобы использовать код VS с Microsoft EDGE, необходимо установить [отладчик для расширения Microsoft Edge](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge) .</span><span class="sxs-lookup"><span data-stu-id="dbf49-139">To use VS Code with Microsoft Edge, you must install the [Debugger for Microsoft Edge](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge) extension.</span></span>

### <a name="run-your-add-in-from-vs-code"></a><span data-ttu-id="dbf49-140">Запуск надстройки из кода VS</span><span class="sxs-lookup"><span data-stu-id="dbf49-140">Run your add-in from VS Code</span></span>

1. <span data-ttu-id="dbf49-141">Откройте корневую папку проекта "пользовательские функции" в [VS Code](https://code.visualstudio.com/).</span><span class="sxs-lookup"><span data-stu-id="dbf49-141">Open your custom functions root project folder in [VS Code](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="dbf49-142">Выберите пункт **терминал > выполнить задачу** и введите или выберите **Контрольное значение**.</span><span class="sxs-lookup"><span data-stu-id="dbf49-142">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="dbf49-143">В этом случае будут отслеживаться и перестраиваться все изменения файлов.</span><span class="sxs-lookup"><span data-stu-id="dbf49-143">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="dbf49-144">Выберите пункт **терминал > выполнить задачу** и введите или выберите **сервер разработки**.</span><span class="sxs-lookup"><span data-stu-id="dbf49-144">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span>

### <a name="start-the-vs-code-debugger"></a><span data-ttu-id="dbf49-145">Запуск отладчика кода VS</span><span class="sxs-lookup"><span data-stu-id="dbf49-145">Start the VS Code debugger</span></span>

4. <span data-ttu-id="dbf49-146">Нажмите кнопку **просмотр > Отладка** или введите **CTRL + SHIFT + D** , чтобы переключиться в представление отладки.</span><span class="sxs-lookup"><span data-stu-id="dbf49-146">Choose **View > Debug** or enter **Ctrl+Shift+D** to switch to debug view.</span></span>
5. <span data-ttu-id="dbf49-147">В разделе Параметры отладки выберите **Office Online (Microsoft EDGE)**.</span><span class="sxs-lookup"><span data-stu-id="dbf49-147">From the Debug options, choose **Office Online (Microsoft Edge)**.</span></span>
6. <span data-ttu-id="dbf49-148">Откройте Excel в браузере Microsoft EDGE и создайте новую книгу.</span><span class="sxs-lookup"><span data-stu-id="dbf49-148">Open Excel in the Microsoft Edge browser and create a new workbook.</span></span>
7. <span data-ttu-id="dbf49-149">Выберите **общий доступ** на ленте и скопируйте ссылку на URL-адрес этой новой книги.</span><span class="sxs-lookup"><span data-stu-id="dbf49-149">Choose **Share** in the ribbon and copy the link for the URL for this new workbook.</span></span>
8. <span data-ttu-id="dbf49-150">Нажмите **клавишу F5** (или выберите **Отладка > начать отладку** из меню), чтобы начать отладку.</span><span class="sxs-lookup"><span data-stu-id="dbf49-150">Select **F5** (or choose **Debug > Start Debugging** from the menu) to begin debugging.</span></span> <span data-ttu-id="dbf49-151">Появится запрос, в котором будет предложен URL-адрес документа.</span><span class="sxs-lookup"><span data-stu-id="dbf49-151">A prompt will appear, which asks for the URL of your document.</span></span>
9. <span data-ttu-id="dbf49-152">Вставьте URL-адрес книги и нажмите клавишу ВВОД.</span><span class="sxs-lookup"><span data-stu-id="dbf49-152">Paste in the URL for your workbook and press Enter.</span></span>

### <a name="sideload-your-add-in"></a><span data-ttu-id="dbf49-153">Загрузка неопубликованной надстройки</span><span class="sxs-lookup"><span data-stu-id="dbf49-153">Sideload your add-in</span></span>

1. <span data-ttu-id="dbf49-154">Перейдите на вкладку **Вставка** на ленте и **в разделе надстройки выберите надстройки** **Office**.</span><span class="sxs-lookup"><span data-stu-id="dbf49-154">Select the **Insert** tab on the ribbon and in the **Add-ins** section, choose **Office Add-ins**.</span></span>
2. <span data-ttu-id="dbf49-155">В диалоговом окне **надстройки Office** откройте вкладку **Мои** надстройки, выберите **Управление моими**надстройками, а затем **отправьте надстройку**.</span><span class="sxs-lookup"><span data-stu-id="dbf49-155">On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.</span></span>
    
    ![Диалоговое окно "Надстройки Office" с раскрывающимся меню в правом верхнем углу, в котором выделен пункт "Управление моими надстройками", а под ним — раскрывающийся список с пунктом "Отправить надстройку"](../images/office-add-ins-my-account.png)

3. <span data-ttu-id="dbf49-157">**Найдите** файл манифеста надстройки и нажмите кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="dbf49-157">**Browse** to the add-in manifest file and then select **Upload**.</span></span>
    
    ![Диалоговое окно отправки надстройки с кнопками "Обзор", "Отправить" и "Отмена"](../images/upload-add-in.png)


### <a name="set-breakpoints"></a><span data-ttu-id="dbf49-159">Задание точек останова</span><span class="sxs-lookup"><span data-stu-id="dbf49-159">Set breakpoints</span></span>
1. <span data-ttu-id="dbf49-160">В коде VS откройте файл сценария исходного кода (**functions. js** или **functions. TS**).</span><span class="sxs-lookup"><span data-stu-id="dbf49-160">In VS Code, open your source code script file (**functions.js** or **functions.ts**).</span></span>
2. <span data-ttu-id="dbf49-161">[Задайте точку останова](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) в исходном коде пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="dbf49-161">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span>
3. <span data-ttu-id="dbf49-162">В книге Excel введите формулу, использующую пользовательскую функцию.</span><span class="sxs-lookup"><span data-stu-id="dbf49-162">In the Excel workbook, enter a formula that uses your custom function.</span></span>

## <a name="use-the-browser-developer-tools-to-debug-custom-functions-in-excel-on-the-web"></a><span data-ttu-id="dbf49-163">Использование средств разработчика браузера для отладки настраиваемых функций в Excel в Интернете</span><span class="sxs-lookup"><span data-stu-id="dbf49-163">Use the browser developer tools to debug custom functions in Excel on the web</span></span>

<span data-ttu-id="dbf49-164">Средства разработчика браузера можно использовать для отладки настраиваемых функций в Excel в Интернете.</span><span class="sxs-lookup"><span data-stu-id="dbf49-164">You can use the browser developer tools to debug custom functions in Excel on the web.</span></span> <span data-ttu-id="dbf49-165">Следующие действия работают как для Windows, так и для macOS.</span><span class="sxs-lookup"><span data-stu-id="dbf49-165">The following steps work for both Windows and macOS.</span></span>

### <a name="run-your-add-in-from-visual-studio-code"></a><span data-ttu-id="dbf49-166">Запуск надстройки из Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="dbf49-166">Run your add-in from Visual Studio Code</span></span>

1. <span data-ttu-id="dbf49-167">Откройте корневую папку проекта пользовательских функций в [Visual Studio Code (VS Code)](https://code.visualstudio.com/).</span><span class="sxs-lookup"><span data-stu-id="dbf49-167">Open your custom functions root project folder in [Visual Studio Code (VS Code)](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="dbf49-168">Выберите пункт **терминал > выполнить задачу** и введите или выберите **Контрольное значение**.</span><span class="sxs-lookup"><span data-stu-id="dbf49-168">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="dbf49-169">В этом случае будут отслеживаться и перестраиваться все изменения файлов.</span><span class="sxs-lookup"><span data-stu-id="dbf49-169">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="dbf49-170">Выберите пункт **терминал > выполнить задачу** и введите или выберите **сервер разработки**.</span><span class="sxs-lookup"><span data-stu-id="dbf49-170">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span>

### <a name="sideload-your-add-in"></a><span data-ttu-id="dbf49-171">Загрузка неопубликованной надстройки</span><span class="sxs-lookup"><span data-stu-id="dbf49-171">Sideload your add-in</span></span>

1. <span data-ttu-id="dbf49-172">Откройте [Microsoft Office в Интернете](https://office.live.com/).</span><span class="sxs-lookup"><span data-stu-id="dbf49-172">Open [Microsoft Office on the web](https://office.live.com/).</span></span>
2. <span data-ttu-id="dbf49-173">Откройте новую книгу Excel.</span><span class="sxs-lookup"><span data-stu-id="dbf49-173">Open a new Excel workbook.</span></span>
3. <span data-ttu-id="dbf49-174">Откройте вкладку **Вставка** на ленте и в разделе **надстройки** выберите надстройки **Office**.</span><span class="sxs-lookup"><span data-stu-id="dbf49-174">Open the **Insert** tab on the ribbon and, in the **Add-ins** section, choose **Office Add-ins**.</span></span>
4. <span data-ttu-id="dbf49-175">В диалоговом окне **надстройки Office** откройте вкладку **Мои** надстройки, выберите **Управление моими**надстройками, а затем **отправьте надстройку**.</span><span class="sxs-lookup"><span data-stu-id="dbf49-175">On the **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then **Upload My Add-in**.</span></span>
    
    ![Диалоговое окно "Надстройки Office" с раскрывающимся меню в правом верхнем углу, в котором выделен пункт "Управление моими надстройками", а под ним — раскрывающийся список с пунктом "Отправить надстройку"](../images/office-add-ins-my-account.png)

5. <span data-ttu-id="dbf49-177">**Найдите** файл манифеста надстройки и выберите **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="dbf49-177">**Browse** to the add-in manifest file, and then select **Upload**.</span></span>
    
    ![Диалоговое окно отправки надстройки с кнопками "Обзор", "Отправить" и "Отмена"](../images/upload-add-in.png)

> [!NOTE]
> <span data-ttu-id="dbf49-179">После неопубликованные документа оно остается неопубликованные при каждом открытии документа.</span><span class="sxs-lookup"><span data-stu-id="dbf49-179">Once you've sideloaded to the document, it will remain sideloaded each time you open the document.</span></span>

### <a name="start-debugging"></a><span data-ttu-id="dbf49-180">Начало отладки</span><span class="sxs-lookup"><span data-stu-id="dbf49-180">Start debugging</span></span>

1. <span data-ttu-id="dbf49-181">Откройте Инструменты разработчика в браузере.</span><span class="sxs-lookup"><span data-stu-id="dbf49-181">Open developer tools in the browser.</span></span> <span data-ttu-id="dbf49-182">Для Chrome и большинства браузеров F12 откроет средства разработчика.</span><span class="sxs-lookup"><span data-stu-id="dbf49-182">For Chrome and most browsers F12 will open the developer tools.</span></span>
2. <span data-ttu-id="dbf49-183">В средствах разработчика откройте файл скрипта исходного кода с помощью **команд Cmd + P** или **CTRL + p** (**functions. js** или **functions. TS**).</span><span class="sxs-lookup"><span data-stu-id="dbf49-183">In developer tools, open your source code script file using **Cmd+P** or **Ctrl+P** (**functions.js** or **functions.ts**).</span></span>
3. <span data-ttu-id="dbf49-184">[Задайте точку останова](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) в исходном коде пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="dbf49-184">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span> 

<span data-ttu-id="dbf49-185">Если вам нужно изменить код, вы можете внести изменения в код VS и сохранить изменения.</span><span class="sxs-lookup"><span data-stu-id="dbf49-185">If you need to change the code you can make edits in VS Code and save the changes.</span></span> <span data-ttu-id="dbf49-186">Обновите браузер, чтобы увидеть загруженные изменения.</span><span class="sxs-lookup"><span data-stu-id="dbf49-186">Refresh the browser to see the changes loaded.</span></span>

## <a name="use-the-command-line-tools-to-debug"></a><span data-ttu-id="dbf49-187">Использование средств командной строки для отладки</span><span class="sxs-lookup"><span data-stu-id="dbf49-187">Use the command line tools to debug</span></span>

<span data-ttu-id="dbf49-188">Если вы не используете код VS, для запуска надстройки можно использовать командную строку (например, bash или PowerShell).</span><span class="sxs-lookup"><span data-stu-id="dbf49-188">If you are not using VS Code, you can use the command line (such as bash, or PowerShell) to run your add-in.</span></span> <span data-ttu-id="dbf49-189">Для отладки кода в Excel в Интернете необходимо использовать инструменты разработчика браузера.</span><span class="sxs-lookup"><span data-stu-id="dbf49-189">You'll need to use the browser developer tools to debug your code in Excel on the web.</span></span> <span data-ttu-id="dbf49-190">Вы не можете выполнить отладку классической версии Excel с помощью командной строки.</span><span class="sxs-lookup"><span data-stu-id="dbf49-190">You cannot debug the desktop version of Excel using the command line.</span></span>

1. <span data-ttu-id="dbf49-191">В командной строке выполняется `npm run watch` Поиск и перестроение при возникновении изменений кода.</span><span class="sxs-lookup"><span data-stu-id="dbf49-191">From the command line run `npm run watch` to watch for and rebuild when code changes occur.</span></span>
2. <span data-ttu-id="dbf49-192">Откройте второе окно командной строки (первый из них будет заблокирован при запуске контрольного значения).</span><span class="sxs-lookup"><span data-stu-id="dbf49-192">Open a second command line window (the first one will be blocked while running the watch.)</span></span>

3. <span data-ttu-id="dbf49-193">Если вы хотите запустить надстройку в классической версии Excel, выполните следующую команду:</span><span class="sxs-lookup"><span data-stu-id="dbf49-193">If you want to start your add-in in the desktop version of Excel, run the following command</span></span>
    
    `npm run start:desktop`
    
    <span data-ttu-id="dbf49-194">Если вы предпочитаете запустить надстройку в Excel в Интернете, выполните следующую команду:</span><span class="sxs-lookup"><span data-stu-id="dbf49-194">Or if you prefer to start your add-in in Excel on the web run the following command</span></span>
    
    `npm run start:web`
    
    <span data-ttu-id="dbf49-195">Для Excel в Интернете вам также потребуется Загрузка неопубликованных надстройку.</span><span class="sxs-lookup"><span data-stu-id="dbf49-195">For Excel on the web you also need to sideload your add-in.</span></span> <span data-ttu-id="dbf49-196">Выполните действия, описанные в [Загрузка неопубликованных надстройки](#sideload-your-add-in) , чтобы Загрузка неопубликованных надстройку.</span><span class="sxs-lookup"><span data-stu-id="dbf49-196">Follow the steps in [Sideload your add-in](#sideload-your-add-in) to sideload your add-in.</span></span> <span data-ttu-id="dbf49-197">Затем перейдите к следующему разделу, чтобы начать отладку.</span><span class="sxs-lookup"><span data-stu-id="dbf49-197">Then continue to the next section to start debugging.</span></span>
    
4. <span data-ttu-id="dbf49-198">Откройте Инструменты разработчика в браузере.</span><span class="sxs-lookup"><span data-stu-id="dbf49-198">Open developer tools in the browser.</span></span> <span data-ttu-id="dbf49-199">Для Chrome и большинства браузеров F12 откроет средства разработчика.</span><span class="sxs-lookup"><span data-stu-id="dbf49-199">For Chrome and most browsers F12 will open the developer tools.</span></span>
5. <span data-ttu-id="dbf49-200">В средствах разработчика откройте файл сценария исходного кода (**functions. js** или **functions. TS**).</span><span class="sxs-lookup"><span data-stu-id="dbf49-200">In developer tools, open your source code script file (**functions.js** or **functions.ts**).</span></span> <span data-ttu-id="dbf49-201">Код настраиваемых функций может располагаться около конца файла.</span><span class="sxs-lookup"><span data-stu-id="dbf49-201">Your custom functions code may be located near the end of the file.</span></span>
6. <span data-ttu-id="dbf49-202">В исходном коде пользовательской функции примените точку останова, выбрав строку кода.</span><span class="sxs-lookup"><span data-stu-id="dbf49-202">In the custom function source code, apply a breakpoint by selecting a line of code.</span></span>

<span data-ttu-id="dbf49-203">Если необходимо изменить код, который можно внести в Visual Studio, и сохранить изменения.</span><span class="sxs-lookup"><span data-stu-id="dbf49-203">If you need to change the code you can make edits in Visual Studio and save the changes.</span></span> <span data-ttu-id="dbf49-204">Обновите браузер, чтобы увидеть загруженные изменения.</span><span class="sxs-lookup"><span data-stu-id="dbf49-204">Refresh the browser to see the changes loaded.</span></span>

### <a name="commands-for-building-and-running-your-add-in"></a><span data-ttu-id="dbf49-205">Команды для построения и запуска надстройки</span><span class="sxs-lookup"><span data-stu-id="dbf49-205">Commands for building and running your add-in</span></span>

<span data-ttu-id="dbf49-206">Доступно несколько задач сборки:</span><span class="sxs-lookup"><span data-stu-id="dbf49-206">There are several build tasks available:</span></span>
- <span data-ttu-id="dbf49-207">`npm run watch`: сборки для разработки и автоматически перестраивается при сохранении исходного файла</span><span class="sxs-lookup"><span data-stu-id="dbf49-207">`npm run watch`: builds for development and automatically rebuilds when a source file is saved</span></span>
- <span data-ttu-id="dbf49-208">`npm run build-dev`: сборки для разработки один раз</span><span class="sxs-lookup"><span data-stu-id="dbf49-208">`npm run build-dev`: builds for development once</span></span>
- <span data-ttu-id="dbf49-209">`npm run build`: сборки для производства</span><span class="sxs-lookup"><span data-stu-id="dbf49-209">`npm run build`: builds for production</span></span>
- <span data-ttu-id="dbf49-210">`npm run dev-server`: запускает веб-сервер, используемый для разработки</span><span class="sxs-lookup"><span data-stu-id="dbf49-210">`npm run dev-server`: runs the web server used for development</span></span>

<span data-ttu-id="dbf49-211">Для запуска отладки на рабочем столе или в сети можно использовать следующие задачи.</span><span class="sxs-lookup"><span data-stu-id="dbf49-211">You can use the following tasks to start debugging on desktop or online.</span></span>
- <span data-ttu-id="dbf49-212">`npm run start:desktop`: Запускает Excel на настольном компьютере и сиделоадс надстройку.</span><span class="sxs-lookup"><span data-stu-id="dbf49-212">`npm run start:desktop`: Starts Excel on desktop and sideloads your add-in.</span></span>
- <span data-ttu-id="dbf49-213">`npm run start:web`: Запуск Excel в Интернете и сиделоадс надстройки.</span><span class="sxs-lookup"><span data-stu-id="dbf49-213">`npm run start:web`: Starts Excel on the web and sideloads your add-in.</span></span>
- <span data-ttu-id="dbf49-214">`npm run stop`: Остановка Excel и отладка.</span><span class="sxs-lookup"><span data-stu-id="dbf49-214">`npm run stop`: Stops Excel and debugging.</span></span>

## <a name="next-steps"></a><span data-ttu-id="dbf49-215">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="dbf49-215">Next steps</span></span>
<span data-ttu-id="dbf49-216">Узнайте о [методиках проверки подлинности в пользовательских функциях](custom-functions-authentication.md).</span><span class="sxs-lookup"><span data-stu-id="dbf49-216">Learn about [authentication practices in custom functions](custom-functions-authentication.md).</span></span> <span data-ttu-id="dbf49-217">Или просмотрите [уникальную архитектуру пользовательской функции](custom-functions-architecture.md).</span><span class="sxs-lookup"><span data-stu-id="dbf49-217">Or, review [custom function's unique architecture](custom-functions-architecture.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="dbf49-218">См. также</span><span class="sxs-lookup"><span data-stu-id="dbf49-218">See also</span></span>

* [<span data-ttu-id="dbf49-219">Устранение неполадок пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="dbf49-219">Custom functions troubleshooting</span></span>](custom-functions-troubleshooting.md)
* [<span data-ttu-id="dbf49-220">Обработка ошибок в пользовательских функциях Excel</span><span class="sxs-lookup"><span data-stu-id="dbf49-220">Error handling for custom functions in Excel</span></span>](custom-functions-errors.md)
* [<span data-ttu-id="dbf49-221">Создание пользовательских функций, совместимых с функциями XLL, определенными пользователями</span><span class="sxs-lookup"><span data-stu-id="dbf49-221">Make your custom functions compatible with XLL user-defined functions</span></span>](make-custom-functions-compatible-with-xll-udf.md)
* [<span data-ttu-id="dbf49-222">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="dbf49-222">Create custom functions in Excel</span></span>](custom-functions-overview.md)
