---
ms.date: 03/13/2019
description: Отладка пользовательских функций в Excel.
title: Отладка пользовательских функций (Предварительная версия)
localization_priority: Normal
ms.openlocfilehash: 66b55855fdbdc3b3cfc7a316cb8fd7e06f073213
ms.sourcegitcommit: 14ceac067e0e130869b861d289edb438b5e3eff9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/04/2019
ms.locfileid: "31478968"
---
# <a name="custom-functions-debugging-preview"></a><span data-ttu-id="7d50e-103">Отладка пользовательских функций (Предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="7d50e-103">Custom functions debugging (preview)</span></span>

<span data-ttu-id="7d50e-104">Отладка настраиваемых функций может осуществляться несколькими способами, в зависимости от используемой платформы.</span><span class="sxs-lookup"><span data-stu-id="7d50e-104">Debugging for custom functions can be accomplished by multiple means, depending on what platform you're using.</span></span>

<span data-ttu-id="7d50e-105">В Windows:</span><span class="sxs-lookup"><span data-stu-id="7d50e-105">On Windows:</span></span>
- [<span data-ttu-id="7d50e-106">Отладчик Excel для наСтольных ПК и Visual Studio Code (VS Code)</span><span class="sxs-lookup"><span data-stu-id="7d50e-106">Excel Desktop and Visual Studio Code (VS Code) debugger</span></span>](#use-the-vs-code-debugger-for-excel-desktop)
- [<span data-ttu-id="7d50e-107">Microsoft Excel Online и отладчик кода VS</span><span class="sxs-lookup"><span data-stu-id="7d50e-107">Excel Online and VS Code debugger</span></span>](#use-the-vs-code-debugger-for-excel-online-in-microsoft-edge)
- [<span data-ttu-id="7d50e-108">Средства Excel Online и браузера</span><span class="sxs-lookup"><span data-stu-id="7d50e-108">Excel Online and browser tools</span></span>](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-online)
- [<span data-ttu-id="7d50e-109">Командная строка</span><span class="sxs-lookup"><span data-stu-id="7d50e-109">Command line</span></span>](#use-the-command-line-tools-to-debug)

<span data-ttu-id="7d50e-110">На компьютерах Mac:</span><span class="sxs-lookup"><span data-stu-id="7d50e-110">On Mac:</span></span>
- [<span data-ttu-id="7d50e-111">Средства Excel Online и браузера</span><span class="sxs-lookup"><span data-stu-id="7d50e-111">Excel Online and browser tools</span></span>](#use-the-browser-developer-tools-to-debug-custom-functions-in-excel-online)
- [<span data-ttu-id="7d50e-112">Командная строка</span><span class="sxs-lookup"><span data-stu-id="7d50e-112">Command line</span></span>](#use-the-command-line-tools-to-debug)

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

> [!NOTE]
> <span data-ttu-id="7d50e-113">Для простоты в этой статье показана Отладка в контексте использования Visual Studio Code для редактирования, запуска задач и в некоторых случаях использования представления отладки.</span><span class="sxs-lookup"><span data-stu-id="7d50e-113">For simplicity, this article shows debugging in the context of using Visual Studio Code to edit, run tasks, and in some cases use the debug view.</span></span> <span data-ttu-id="7d50e-114">Если вы используете другой редактор или средство командной строки, ознакомьтесь с инструкциями по использованию [командной строки](#Use-the-command-line-tools-to-debug) в конце этой статьи.</span><span class="sxs-lookup"><span data-stu-id="7d50e-114">If you are using a different editor or command line tool, see the [command line instructions](#Use-the-command-line-tools-to-debug) at the end of this article.</span></span>

## <a name="requirements"></a><span data-ttu-id="7d50e-115">Требования</span><span class="sxs-lookup"><span data-stu-id="7d50e-115">Requirements</span></span>

<span data-ttu-id="7d50e-116">Перед началом отладки необходимо создать проект надстройки настраиваемых функций с помощью генератора Yo Office и убедиться, что у вас есть доверенные самозаверяющие сертификаты для вашего проекта.</span><span class="sxs-lookup"><span data-stu-id="7d50e-116">Before starting to debug, you should create a custom functions add-in project using the Yo Office generator and ensured that you have trusted self-signed certificates for your project.</span></span> <span data-ttu-id="7d50e-117">Инструкции по созданию проекта представлены в [руководстве Custom functions](https://review.docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions).</span><span class="sxs-lookup"><span data-stu-id="7d50e-117">For instructions to create a project, see the [custom functions tutorial](https://review.docs.microsoft.com/office/dev/add-ins/tutorials/excel-tutorial-create-custom-functions).</span></span> <span data-ttu-id="7d50e-118">Инструкции по доверенным сертификатам можно узнать в статье [Добавление самозаверяющих сертификатов в качестве доверенных корневых сертификатов](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span><span class="sxs-lookup"><span data-stu-id="7d50e-118">For instructions on trusting certificates, see [Adding self-signed certificates as trusted root certificates](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span></span>

## <a name="use-the-vs-code-debugger-for-excel-desktop"></a><span data-ttu-id="7d50e-119">Использование отладчика кода VS для классической версии Excel</span><span class="sxs-lookup"><span data-stu-id="7d50e-119">Use the VS Code debugger for Excel Desktop</span></span>

<span data-ttu-id="7d50e-120">Вы можете использовать код VS для отладки настраиваемых функций в Office Excel на настольном компьютере.</span><span class="sxs-lookup"><span data-stu-id="7d50e-120">You can use VS Code to debug custom functions in Office Excel on the desktop.</span></span>

> [!NOTE]
> <span data-ttu-id="7d50e-121">Отладка на рабочем столе для Mac недоступна, но ее можно [использовать с помощью средств браузера для отладки Excel Online](#debug-in-excel-online-by-using-the-browser-developer-tools).</span><span class="sxs-lookup"><span data-stu-id="7d50e-121">Desktop debugging for the Mac is not available but can be achieved [using the browser tools to debug Excel Online](#debug-in-excel-online-by-using-the-browser-developer-tools).</span></span>

### <a name="run-your-add-in-from-vs-code"></a><span data-ttu-id="7d50e-122">Запуск надстройки из кода VS</span><span class="sxs-lookup"><span data-stu-id="7d50e-122">Run your add-in from VS Code</span></span>

1. <span data-ttu-id="7d50e-123">Откройте корневую папку проекта "пользовательские функции" в [VS Code](https://code.visualstudio.com/).</span><span class="sxs-lookup"><span data-stu-id="7d50e-123">Open your custom functions root project folder in [VS Code](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="7d50e-124">Выберите пункт **терминал _Гт_ запуск задачи** и введите или выберите **Контрольное значение**.</span><span class="sxs-lookup"><span data-stu-id="7d50e-124">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="7d50e-125">В этом случае будут отслеживаться и перестраиваться все изменения файлов.</span><span class="sxs-lookup"><span data-stu-id="7d50e-125">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="7d50e-126">Выберите пункт **терминал _Гт_ запуск задачи** и введите или выберите **сервер разработки**.</span><span class="sxs-lookup"><span data-stu-id="7d50e-126">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span> 

### <a name="start-the-vs-code-debugger"></a><span data-ttu-id="7d50e-127">Запуск отладчика кода VS</span><span class="sxs-lookup"><span data-stu-id="7d50e-127">Start the VS Code debugger</span></span>

4. <span data-ttu-id="7d50e-128">Выберите **Просмотр _Гт_ отладки** или введите **CTRL + SHIFT + D** , чтобы переключиться в представление отладки.</span><span class="sxs-lookup"><span data-stu-id="7d50e-128">Choose **View > Debug** or enter **Ctrl+Shift+D** to switch to debug view.</span></span>
5. <span data-ttu-id="7d50e-129">В разделе Параметры отладки выберите пункт **Рабочий стол Excel**.</span><span class="sxs-lookup"><span data-stu-id="7d50e-129">From the Debug options, choose **Excel Desktop**.</span></span>
6. <span data-ttu-id="7d50e-130">Чтобы начать отладку, нажмите **клавишу F5** (или выберите **Отладка — _гт_ начать отладку** в меню).</span><span class="sxs-lookup"><span data-stu-id="7d50e-130">Select **F5** (or choose **Debug -> Start Debugging** from the menu) to begin debugging.</span></span> <span data-ttu-id="7d50e-131">Откроется новая книга Excel с уже неопубликованные и готовым к использованию надстройкой.</span><span class="sxs-lookup"><span data-stu-id="7d50e-131">A new Excel workbook will open with your add-in already sideloaded and ready to use.</span></span>

### <a name="start-debugging"></a><span data-ttu-id="7d50e-132">Начало отладки</span><span class="sxs-lookup"><span data-stu-id="7d50e-132">Start debugging</span></span>

1. <span data-ttu-id="7d50e-133">В коде VS откройте файл сценария исходного кода (functions. js или functions. TS).</span><span class="sxs-lookup"><span data-stu-id="7d50e-133">In VS Code, open your source code script file (functions.js or functions.ts).</span></span>
2. <span data-ttu-id="7d50e-134">[Задайте точку останова](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) в исходном коде пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="7d50e-134">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span>
3. <span data-ttu-id="7d50e-135">В книге Excel введите формулу, использующую пользовательскую функцию.</span><span class="sxs-lookup"><span data-stu-id="7d50e-135">In the Excel workbook, enter a formula that uses your custom function.</span></span>

<span data-ttu-id="7d50e-136">При этом выполнение будет остановлено в строке кода, в которой вы задаете точку останова.</span><span class="sxs-lookup"><span data-stu-id="7d50e-136">At this point execution will stop on the line of code where you set the breakpoint.</span></span> <span data-ttu-id="7d50e-137">Теперь вы можете выполнить отладку кода, задать контрольные значения и использовать любые необходимые возможности отладки кода VS.</span><span class="sxs-lookup"><span data-stu-id="7d50e-137">Now you can step through your code, set watches, and use any VS Code debugging features you need.</span></span>

## <a name="use-the-vs-code-debugger-for-excel-online-in-microsoft-edge"></a><span data-ttu-id="7d50e-138">Использование отладчика кода VS для Excel Online в Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="7d50e-138">Use the VS Code debugger for Excel Online in Microsoft Edge</span></span>

<span data-ttu-id="7d50e-139">Вы можете использовать код VS для отладки настраиваемых функций в Excel Online в браузере Microsoft Edge.</span><span class="sxs-lookup"><span data-stu-id="7d50e-139">You can use VS Code to debug custom functions in Excel Online in the Microsoft Edge browser.</span></span> <span data-ttu-id="7d50e-140">Чтобы использовать код VS с Microsoft EDGE, необходимо установить [отладчик для расширения Microsoft Edge](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge) .</span><span class="sxs-lookup"><span data-stu-id="7d50e-140">To use VS Code with Microsoft Edge, you must install the [Debugger for Microsoft Edge](https://marketplace.visualstudio.com/items?itemName=msjsdiag.debugger-for-edge) extension.</span></span>

### <a name="run-your-add-in-from-vs-code"></a><span data-ttu-id="7d50e-141">Запуск надстройки из кода VS</span><span class="sxs-lookup"><span data-stu-id="7d50e-141">Run your add-in from VS Code</span></span>

1. <span data-ttu-id="7d50e-142">Откройте корневую папку проекта "пользовательские функции" в [VS Code](https://code.visualstudio.com/).</span><span class="sxs-lookup"><span data-stu-id="7d50e-142">Open your custom functions root project folder in [VS Code](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="7d50e-143">Выберите пункт **терминал _Гт_ запуск задачи** и введите или выберите **Контрольное значение**.</span><span class="sxs-lookup"><span data-stu-id="7d50e-143">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="7d50e-144">В этом случае будут отслеживаться и перестраиваться все изменения файлов.</span><span class="sxs-lookup"><span data-stu-id="7d50e-144">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="7d50e-145">Выберите пункт **терминал _Гт_ запуск задачи** и введите или выберите **сервер разработки**.</span><span class="sxs-lookup"><span data-stu-id="7d50e-145">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span> 

### <a name="start-the-vs-code-debugger"></a><span data-ttu-id="7d50e-146">Запуск отладчика кода VS</span><span class="sxs-lookup"><span data-stu-id="7d50e-146">Start the VS Code debugger</span></span>

4. <span data-ttu-id="7d50e-147">Выберите **Просмотр _Гт_ отладки** или введите **CTRL + SHIFT + D** , чтобы переключиться в представление отладки.</span><span class="sxs-lookup"><span data-stu-id="7d50e-147">Choose **View > Debug** or enter **Ctrl+Shift+D** to switch to debug view.</span></span>
5. <span data-ttu-id="7d50e-148">В разделе Параметры отладки выберите **Office Online (EDGE)**.</span><span class="sxs-lookup"><span data-stu-id="7d50e-148">From the Debug options, choose **Office Online (Edge)**.</span></span>
6. <span data-ttu-id="7d50e-149">Откройте Excel Online с помощью браузера Microsoft EDGE, откройте Excel Online и создайте новую книгу.</span><span class="sxs-lookup"><span data-stu-id="7d50e-149">Open Excel Online using the Microsoft Edge browser, open Excel Online, and create a new workbook.</span></span>
7. <span data-ttu-id="7d50e-150">Выберите **общий доступ** на ленте и скопируйте ссылку на URL-адрес этой новой книги.</span><span class="sxs-lookup"><span data-stu-id="7d50e-150">Choose **Share** in the ribbon and copy the link for the URL for this new workbook.</span></span>
8. <span data-ttu-id="7d50e-151">Нажмите **клавишу F5** (или выберите **Отладка _гт_ начать отладку** в меню), чтобы начать отладку.</span><span class="sxs-lookup"><span data-stu-id="7d50e-151">Select **F5** (or choose **Debug > Start Debugging** from the menu) to begin debugging.</span></span> <span data-ttu-id="7d50e-152">Появится запрос, в котором будет предложен URL-адрес документа.</span><span class="sxs-lookup"><span data-stu-id="7d50e-152">A prompt will appear, which asks for the URL of your document.</span></span>
9. <span data-ttu-id="7d50e-153">Вставьте URL-адрес книги и нажмите клавишу ВВОД.</span><span class="sxs-lookup"><span data-stu-id="7d50e-153">Paste in the URL for your workbook and press Enter.</span></span>

### <a name="sideload-your-add-in"></a><span data-ttu-id="7d50e-154">Загрузка неопубликованной надстройки</span><span class="sxs-lookup"><span data-stu-id="7d50e-154">Sideload your add-in</span></span>   

1. <span data-ttu-id="7d50e-155">Перейдите на вкладку **Вставка** на ленте и в разделе надстройки выберите надстройки **Office**. \*\*\*\*</span><span class="sxs-lookup"><span data-stu-id="7d50e-155">Select the  **Insert** tab on the ribbon and in the **Add-ins** section, choose **Office Add-ins**.</span></span>
2. <span data-ttu-id="7d50e-156">В диалоговом окне **Надстройки Office** откройте вкладку **МОИ НАДСТРОЙКИ** и выберите **Управление моими надстройками** > **Отправить надстройку**.</span><span class="sxs-lookup"><span data-stu-id="7d50e-156">On the  **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then  **Upload My Add-in**.</span></span>
    
    ![Диалоговое окно "Надстройки Office" с раскрывающимся меню в правом верхнем углу, в котором выделен пункт "Управление моими надстройками", а под ним — раскрывающийся список с пунктом "Отправить надстройку"](../images/office-add-ins-my-account.png)

3.  <span data-ttu-id="7d50e-158">**Найдите** файл манифеста надстройки и нажмите кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="7d50e-158">**Browse** to the add-in manifest file and then select **Upload**.</span></span>
    
    ![Диалоговое окно отправки надстройки с кнопками "Обзор", "Отправить" и "Отмена"](../images/upload-add-in.png)


### <a name="set-breakpoints"></a><span data-ttu-id="7d50e-160">Задание точек останова</span><span class="sxs-lookup"><span data-stu-id="7d50e-160">Set breakpoints</span></span>
1. <span data-ttu-id="7d50e-161">В коде VS откройте файл сценария исходного кода (functions. js или functions. TS).</span><span class="sxs-lookup"><span data-stu-id="7d50e-161">In VS Code, open your source code script file (functions.js or functions.ts).</span></span>
2. <span data-ttu-id="7d50e-162">[Задайте точку останова](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) в исходном коде пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="7d50e-162">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span>
3. <span data-ttu-id="7d50e-163">В книге Excel введите формулу, использующую пользовательскую функцию.</span><span class="sxs-lookup"><span data-stu-id="7d50e-163">In the Excel workbook, enter a formula that uses your custom function.</span></span>

## <a name="use-the-browser-developer-tools-to-debug-custom-functions-in-excel-online"></a><span data-ttu-id="7d50e-164">Использование средств разработчика браузера для отладки настраиваемых функций в Excel Online</span><span class="sxs-lookup"><span data-stu-id="7d50e-164">Use the browser developer tools to debug custom functions in Excel Online</span></span>

<span data-ttu-id="7d50e-165">Средства разработчика браузера можно использовать для отладки настраиваемых функций в Excel Online.</span><span class="sxs-lookup"><span data-stu-id="7d50e-165">You can use the browser developer tools to debug custom functions in Excel Online.</span></span> <span data-ttu-id="7d50e-166">Следующие действия работают как для Windows, так и для macOS.</span><span class="sxs-lookup"><span data-stu-id="7d50e-166">The following steps work for both Windows and macOS.</span></span>

### <a name="run-your-add-in-from-visual-studio-code"></a><span data-ttu-id="7d50e-167">Запуск надстройки из Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="7d50e-167">Run your add-in from Visual Studio Code</span></span>

1. <span data-ttu-id="7d50e-168">Откройте корневую папку проекта пользовательских функций в [Visual Studio Code (VS Code)](https://code.visualstudio.com/).</span><span class="sxs-lookup"><span data-stu-id="7d50e-168">Open your custom functions root project folder in [Visual Studio Code (VS Code)](https://code.visualstudio.com/).</span></span>
2. <span data-ttu-id="7d50e-169">Выберите пункт **терминал _Гт_ запуск задачи** и введите или выберите **Контрольное значение**.</span><span class="sxs-lookup"><span data-stu-id="7d50e-169">Choose **Terminal > Run Task** and type or select **Watch**.</span></span> <span data-ttu-id="7d50e-170">В этом случае будут отслеживаться и перестраиваться все изменения файлов.</span><span class="sxs-lookup"><span data-stu-id="7d50e-170">This will monitor and rebuild for any file changes.</span></span>
3. <span data-ttu-id="7d50e-171">Выберите пункт **терминал _Гт_ запуск задачи** и введите или выберите **сервер разработки**.</span><span class="sxs-lookup"><span data-stu-id="7d50e-171">Choose **Terminal > Run Task** and type or select **Dev Server**.</span></span> 

### <a name="sideload-your-add-in"></a><span data-ttu-id="7d50e-172">Загрузка неопубликованной надстройки</span><span class="sxs-lookup"><span data-stu-id="7d50e-172">Sideload your add-in</span></span>   

1. <span data-ttu-id="7d50e-173">Откройте [Microsoft Office Online](https://office.live.com/).</span><span class="sxs-lookup"><span data-stu-id="7d50e-173">Open [Microsoft Office Online](https://office.live.com/).</span></span>
2. <span data-ttu-id="7d50e-174">Откройте новую книгу Excel.</span><span class="sxs-lookup"><span data-stu-id="7d50e-174">Open a new Excel workbook.</span></span>
3. <span data-ttu-id="7d50e-175">Откройте вкладку  **Вставка** на ленте и в разделе **Надстройки** выберите **Надстройки Office**.</span><span class="sxs-lookup"><span data-stu-id="7d50e-175">Open the  **Insert** tab on the ribbon and, in the **Add-ins** section, choose **Office Add-ins**.</span></span>
4. <span data-ttu-id="7d50e-176">В диалоговом окне **Надстройки Office** откройте вкладку **МОИ НАДСТРОЙКИ** и выберите **Управление моими надстройками** > **Отправить надстройку**.</span><span class="sxs-lookup"><span data-stu-id="7d50e-176">On the  **Office Add-ins** dialog, select the **MY ADD-INS** tab, choose **Manage My Add-ins**, and then  **Upload My Add-in**.</span></span>
    
    ![Диалоговое окно "Надстройки Office" с раскрывающимся меню в правом верхнем углу, в котором выделен пункт "Управление моими надстройками", а под ним — раскрывающийся список с пунктом "Отправить надстройку"](../images/office-add-ins-my-account.png)

5.  <span data-ttu-id="7d50e-178">**Найдите** файл манифеста надстройки и выберите **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="7d50e-178">**Browse** to the add-in manifest file, and then select **Upload**.</span></span>
    
    ![Диалоговое окно отправки надстройки с кнопками "Обзор", "Отправить" и "Отмена"](../images/upload-add-in.png)

> [!NOTE]
> <span data-ttu-id="7d50e-180">После неопубликованные документа оно остается неопубликованные при каждом открытии документа.</span><span class="sxs-lookup"><span data-stu-id="7d50e-180">Once you've sideloaded to the document, it will remain sideloaded each time you open the document.</span></span>

### <a name="start-debugging"></a><span data-ttu-id="7d50e-181">Начало отладки</span><span class="sxs-lookup"><span data-stu-id="7d50e-181">Start debugging</span></span>

1. <span data-ttu-id="7d50e-182">Откройте Инструменты разработчика в браузере.</span><span class="sxs-lookup"><span data-stu-id="7d50e-182">Open developer tools in the browser.</span></span> <span data-ttu-id="7d50e-183">Для Chrome и большинства браузеров F12 откроет средства разработчика.</span><span class="sxs-lookup"><span data-stu-id="7d50e-183">For Chrome and most browsers F12 will open the developer tools.</span></span>
2. <span data-ttu-id="7d50e-184">В средствах разработчика откройте файл скрипта исходного кода с помощью **команд Cmd + P** или **CTRL + p** (functions. js или functions. TS).</span><span class="sxs-lookup"><span data-stu-id="7d50e-184">In developer tools, open your source code script file using **Cmd+P** or **Ctrl+P** (functions.js or functions.ts).</span></span>
3. <span data-ttu-id="7d50e-185">[Задайте точку останова](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) в исходном коде пользовательской функции.</span><span class="sxs-lookup"><span data-stu-id="7d50e-185">[Set a breakpoint](https://code.visualstudio.com/Docs/editor/debugging#_breakpoints) in the custom function source code.</span></span> 

<span data-ttu-id="7d50e-186">Если вам нужно изменить код, вы можете внести изменения в код VS и сохранить изменения.</span><span class="sxs-lookup"><span data-stu-id="7d50e-186">If you need to change the code you can make edits in VS Code and save the changes.</span></span> <span data-ttu-id="7d50e-187">Обновите браузер, чтобы увидеть загруженные изменения.</span><span class="sxs-lookup"><span data-stu-id="7d50e-187">Refresh the browser to see the changes loaded.</span></span>

## <a name="use-the-command-line-tools-to-debug"></a><span data-ttu-id="7d50e-188">Использование средств командной строки для отладки</span><span class="sxs-lookup"><span data-stu-id="7d50e-188">Use the command line tools to debug</span></span>

<span data-ttu-id="7d50e-189">Если вы не используете код VS, для запуска надстройки можно использовать командную строку (например, bash или PowerShell).</span><span class="sxs-lookup"><span data-stu-id="7d50e-189">If you are not using VS Code, you can use the command line (such as bash, or PowerShell) to run your add-in.</span></span> <span data-ttu-id="7d50e-190">Для отладки кода в Excel Online необходимо использовать инструменты разработчика браузера.</span><span class="sxs-lookup"><span data-stu-id="7d50e-190">You'll need to use the browser developer tools to debug your code in Excel Online.</span></span> <span data-ttu-id="7d50e-191">Вы не можете выполнить отладку классической версии Excel с помощью командной строки.</span><span class="sxs-lookup"><span data-stu-id="7d50e-191">You cannot debug the desktop version of Excel using the command line.</span></span>

1. <span data-ttu-id="7d50e-192">В командной строке выполняется `npm run watch` Поиск и перестроение при возникновении изменений кода.</span><span class="sxs-lookup"><span data-stu-id="7d50e-192">From the command line run `npm run watch` to watch for and rebuild when code changes occur.</span></span>
2. <span data-ttu-id="7d50e-193">Откройте второе окно командной строки (первый из них будет заблокирован при запуске контрольного значения).</span><span class="sxs-lookup"><span data-stu-id="7d50e-193">Open a second command line window (the first one will be blocked while running the watch.)</span></span>

3. <span data-ttu-id="7d50e-194">Если вы хотите запустить надстройку в классической версии Excel, выполните следующую команду:</span><span class="sxs-lookup"><span data-stu-id="7d50e-194">If you want to start your add-in in the desktop version of Excel, run the following command</span></span>
    
    `npm run start desktop`
    
    <span data-ttu-id="7d50e-195">Если вы предпочитаете запустить надстройку в Excel Online, выполните следующую команду</span><span class="sxs-lookup"><span data-stu-id="7d50e-195">Or if you prefer to start your add-in in Excel Online run the following command</span></span>
    
    `npm run start web`
    
    <span data-ttu-id="7d50e-196">Для Excel Online также потребуется Загрузка неопубликованных надстройку.</span><span class="sxs-lookup"><span data-stu-id="7d50e-196">For Excel Online you also need to sideload your add-in.</span></span> <span data-ttu-id="7d50e-197">Выполните действия, описанные в [Загрузка неопубликованных надстройки](#Sideload-your-add-in) , чтобы Загрузка неопубликованных надстройку.</span><span class="sxs-lookup"><span data-stu-id="7d50e-197">Follow the steps in [Sideload your add-in](#Sideload-your-add-in) to sideload your add-in.</span></span> <span data-ttu-id="7d50e-198">Затем перейдите к следующему разделу, чтобы начать отладку.</span><span class="sxs-lookup"><span data-stu-id="7d50e-198">Then continue to the next section to start debugging.</span></span>
    
4. <span data-ttu-id="7d50e-199">Откройте Инструменты разработчика в браузере.</span><span class="sxs-lookup"><span data-stu-id="7d50e-199">Open developer tools in the browser.</span></span> <span data-ttu-id="7d50e-200">Для Chrome и большинства браузеров F12 откроет средства разработчика.</span><span class="sxs-lookup"><span data-stu-id="7d50e-200">For Chrome and most browsers F12 will open the developer tools.</span></span>
5. <span data-ttu-id="7d50e-201">В средствах разработчика откройте файл сценария исходного кода (functions. js или functions. TS).</span><span class="sxs-lookup"><span data-stu-id="7d50e-201">In developer tools, open your source code script file (functions.js or functions.ts).</span></span> <span data-ttu-id="7d50e-202">Код настраиваемых функций может располагаться около конца файла.</span><span class="sxs-lookup"><span data-stu-id="7d50e-202">Your custom functions code may be located near the end of the file.</span></span>
6. <span data-ttu-id="7d50e-203">В исходном коде пользовательской функции примените точку останова, выбрав строку кода.</span><span class="sxs-lookup"><span data-stu-id="7d50e-203">In the custom function source code, apply a breakpoint by selecting a line of code.</span></span>

<span data-ttu-id="7d50e-204">Если необходимо изменить код, который можно внести в Visual Studio, и сохранить изменения.</span><span class="sxs-lookup"><span data-stu-id="7d50e-204">If you need to change the code you can make edits in Visual Studio and save the changes.</span></span> <span data-ttu-id="7d50e-205">Обновите браузер, чтобы увидеть загруженные изменения.</span><span class="sxs-lookup"><span data-stu-id="7d50e-205">Refresh the browser to see the changes loaded.</span></span>

### <a name="commands-for-building-and-running-your-add-in"></a><span data-ttu-id="7d50e-206">Команды для построения и запуска надстройки</span><span class="sxs-lookup"><span data-stu-id="7d50e-206">Commands for building and running your add-in</span></span>

<span data-ttu-id="7d50e-207">Доступно несколько задач сборки:</span><span class="sxs-lookup"><span data-stu-id="7d50e-207">There are several build tasks available:</span></span>
- `npm run watch`<span data-ttu-id="7d50e-208">: сборки для разработки и автоматически перестраивается при сохранении исходного файла</span><span class="sxs-lookup"><span data-stu-id="7d50e-208">: builds for development and automatically rebuilds when a source file is saved</span></span>
- `npm run build-dev`<span data-ttu-id="7d50e-209">: сборки для разработки один раз</span><span class="sxs-lookup"><span data-stu-id="7d50e-209">: builds for development once</span></span>
- `npm run build`<span data-ttu-id="7d50e-210">: сборки для производства</span><span class="sxs-lookup"><span data-stu-id="7d50e-210">: builds for production</span></span>
- `npm run dev-server`<span data-ttu-id="7d50e-211">: запускает веб-сервер, используемый для разработки</span><span class="sxs-lookup"><span data-stu-id="7d50e-211">: runs the web server used for development</span></span>

<span data-ttu-id="7d50e-212">Для запуска отладки на рабочем столе или в сети можно использовать следующие задачи.</span><span class="sxs-lookup"><span data-stu-id="7d50e-212">You can use the following tasks to start debugging on desktop or online.</span></span>
- `npm run start desktop`<span data-ttu-id="7d50e-213">: Запускает Excel на настольном компьютере и сиделоадс надстройку.</span><span class="sxs-lookup"><span data-stu-id="7d50e-213">: Starts Excel on desktop and sideloads your add-in.</span></span>
- `npm run start web`<span data-ttu-id="7d50e-214">: Запуск Excel Online и сиделоадс надстройки.</span><span class="sxs-lookup"><span data-stu-id="7d50e-214">: Starts Excel Online and sideloads your add-in.</span></span>
- `npm run stop`<span data-ttu-id="7d50e-215">: Остановка Excel и отладка.</span><span class="sxs-lookup"><span data-stu-id="7d50e-215">: Stops Excel and debugging.</span></span>

## <a name="see-also"></a><span data-ttu-id="7d50e-216">См. также</span><span class="sxs-lookup"><span data-stu-id="7d50e-216">See also</span></span>

* [<span data-ttu-id="7d50e-217">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="7d50e-217">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="7d50e-218">Среда выполнения для пользовательских функций Excel</span><span class="sxs-lookup"><span data-stu-id="7d50e-218">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="7d50e-219">Рекомендации по пользовательским функциям</span><span class="sxs-lookup"><span data-stu-id="7d50e-219">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="7d50e-220">Журнал изменений пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="7d50e-220">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="7d50e-221">Руководство по пользовательским функциям в Excel</span><span class="sxs-lookup"><span data-stu-id="7d50e-221">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
