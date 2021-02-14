---
title: Надстройка Microsoft Office "Расширение отладчика для Visual Studio Code"
description: Используйте расширение Visual Studio кода Microsoft Office отладить надстройку Office.
ms.date: 02/01/2021
localization_priority: Normal
ms.openlocfilehash: 60f7e6646cc0bfa2740e3bac0cab5f603b32dd84
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237933"
---
# <a name="microsoft-office-add-in-debugger-extension-for-visual-studio-code"></a><span data-ttu-id="3c1ad-103">Надстройка Microsoft Office "Расширение отладчика для Visual Studio Code"</span><span class="sxs-lookup"><span data-stu-id="3c1ad-103">Microsoft Office Add-in Debugger Extension for Visual Studio Code</span></span>

<span data-ttu-id="3c1ad-104">Расширение Microsoft Office отладка надстройки для Visual Studio Code позволяет выполнить отладку надстройки Office в Microsoft Edge с помощью исходной времени работы WebView (EdgeHTML).</span><span class="sxs-lookup"><span data-stu-id="3c1ad-104">The Microsoft Office Add-in Debugger Extension for Visual Studio Code allows you to debug your Office Add-in against the Microsoft Edge with the original webView (EdgeHTML) runtime.</span></span> <span data-ttu-id="3c1ad-105">Инструкции по отладке в Microsoft Edge WebView2 (на основе Chromium) см. [в этой статье](./debug-desktop-using-edge-chromium.md)</span><span class="sxs-lookup"><span data-stu-id="3c1ad-105">For instructions about debugging against Microsoft Edge WebView2 (Chromium-based), [see this article](./debug-desktop-using-edge-chromium.md)</span></span>

<span data-ttu-id="3c1ad-106">Этот режим отладки является динамическим, что позволяет устанавливать точки останова во время работы кода.</span><span class="sxs-lookup"><span data-stu-id="3c1ad-106">This debugging mode is dynamic, allowing you to set breakpoints while code is running.</span></span> <span data-ttu-id="3c1ad-107">Вы можете сразу увидеть изменения в коде, когда отладка подключена, без потери сеанса отладки.</span><span class="sxs-lookup"><span data-stu-id="3c1ad-107">You can see changes in your code immediately while the debugger is attached, all without losing your debugging session.</span></span> <span data-ttu-id="3c1ad-108">Изменения в коде также сохраняются, поэтому вы можете увидеть результаты нескольких изменений в коде.</span><span class="sxs-lookup"><span data-stu-id="3c1ad-108">Your code changes also persist, so you can see the results of multiple changes to your code.</span></span> <span data-ttu-id="3c1ad-109">На следующем рисунке показано это расширение в действии.</span><span class="sxs-lookup"><span data-stu-id="3c1ad-109">The following image shows this extension in action.</span></span>

![Расширение надстройки Office Addin Debugger Extension отладка раздела надстроек Excel](../images/vs-debugger-extension-for-office-addins.jpg)

## <a name="prerequisites"></a><span data-ttu-id="3c1ad-111">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="3c1ad-111">Prerequisites</span></span>

- <span data-ttu-id="3c1ad-112">[Visual Studio кода](https://code.visualstudio.com/) (должен запускаться от учетной записи администратора)</span><span class="sxs-lookup"><span data-stu-id="3c1ad-112">[Visual Studio Code](https://code.visualstudio.com/) (must be run as an administrator)</span></span>
- [<span data-ttu-id="3c1ad-113">Node.js (версия 10+)</span><span class="sxs-lookup"><span data-stu-id="3c1ad-113">Node.js (version 10+)</span></span>](https://nodejs.org/)
- <span data-ttu-id="3c1ad-114">Windows 10</span><span class="sxs-lookup"><span data-stu-id="3c1ad-114">Windows 10</span></span>
- [<span data-ttu-id="3c1ad-115">Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="3c1ad-115">Microsoft Edge</span></span>](https://www.microsoft.com/edge)

<span data-ttu-id="3c1ad-116">В этих инструкциях предполагается, что у вас есть опыт работы с командной строкой, вы понимаете базовый javaScript и создали проект надстройки Office перед использованием генератора Yo Office.</span><span class="sxs-lookup"><span data-stu-id="3c1ad-116">These instructions assume you have experience using the command line, understand basic JavaScript, and have created an Office Add-in project before using the Yo Office generator.</span></span> <span data-ttu-id="3c1ad-117">Если вы еще не сделали этого, рассмотрите возможность посетить одно из наших учебников, например это руководство по [надстройки Excel Для Office.](../tutorials/excel-tutorial.md)</span><span class="sxs-lookup"><span data-stu-id="3c1ad-117">If you haven't done this before, consider visiting one of our tutorials, like this [Excel Office Add-in tutorial](../tutorials/excel-tutorial.md).</span></span>

## <a name="install-and-use-the-debugger"></a><span data-ttu-id="3c1ad-118">Установка и использование отладщика</span><span class="sxs-lookup"><span data-stu-id="3c1ad-118">Install and use the debugger</span></span>

1. <span data-ttu-id="3c1ad-119">Если вам нужно создать проект надстройки, создайте его с помощью генератора [Yo Office.](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator)</span><span class="sxs-lookup"><span data-stu-id="3c1ad-119">If you need to create an add-in project, [use the Yo Office generator to create one](../quickstarts/excel-quickstart-jquery.md?tabs=yeomangenerator).</span></span> <span data-ttu-id="3c1ad-120">Следуйте подсказкам в командной строке, чтобы настроить проект.</span><span class="sxs-lookup"><span data-stu-id="3c1ad-120">Follow the prompts within the command line to set up your project.</span></span> <span data-ttu-id="3c1ad-121">Вы можете выбрать любой язык или тип проекта в соответствии со своими потребностями.</span><span class="sxs-lookup"><span data-stu-id="3c1ad-121">You can choose any language or type of project to suit your needs.</span></span>

> [!NOTE]
> <span data-ttu-id="3c1ad-122">Если у вас уже есть проект, пропустите шаг 1 и переходить к шагу 2.</span><span class="sxs-lookup"><span data-stu-id="3c1ad-122">If you already have a project, skip step 1 and move to step 2.</span></span>

2. <span data-ttu-id="3c1ad-123">Откройте командную подсказку от администратора.</span><span class="sxs-lookup"><span data-stu-id="3c1ad-123">Open a command prompt as administrator.</span></span>
   <span data-ttu-id="3c1ad-124">![Параметры командной подсказки, в том числе "Запуск от администратора" в Windows 10](../images/run-as-administrator-vs-code.jpg)</span><span class="sxs-lookup"><span data-stu-id="3c1ad-124">![Command prompt options, including "run as administrator" in Windows 10](../images/run-as-administrator-vs-code.jpg)</span></span>

3. <span data-ttu-id="3c1ad-125">Перейдите в каталог проекта.</span><span class="sxs-lookup"><span data-stu-id="3c1ad-125">Navigate to your project directory.</span></span>

4. <span data-ttu-id="3c1ad-126">Чтобы открыть проект в Visual Studio code от администратора, Visual Studio следующую команду.</span><span class="sxs-lookup"><span data-stu-id="3c1ad-126">Run the following command to open your project in Visual Studio Code as an administrator.</span></span>

```command&nbsp;line
code .
```

<span data-ttu-id="3c1ad-127">После Visual Studio кода перейдите в папку проекта вручную.</span><span class="sxs-lookup"><span data-stu-id="3c1ad-127">Once Visual Studio Code is open, navigate manually to the project folder.</span></span>

> [!TIP]
> <span data-ttu-id="3c1ad-128">Чтобы открыть Visual Studio Code от имени администратора, выберите параметр "Запуск от имени администратора" при открытии Visual Studio Code после его поиска в Windows. </span><span class="sxs-lookup"><span data-stu-id="3c1ad-128">To open Visual Studio Code as an administrator, select the **run as administrator** option when opening Visual Studio Code after searching for it in Windows.</span></span>

5. <span data-ttu-id="3c1ad-129">В VS Code выберите **CTRL + SHIFT + X,** чтобы открыть план "Расширения".</span><span class="sxs-lookup"><span data-stu-id="3c1ad-129">Within VS Code, select **CTRL + SHIFT + X** to open the Extensions bar.</span></span> <span data-ttu-id="3c1ad-130">Найщите расширение Microsoft Office надстройки и установите его.</span><span class="sxs-lookup"><span data-stu-id="3c1ad-130">Search for the "Microsoft Office Add-in Debugger" extension and install it.</span></span>

6. <span data-ttu-id="3c1ad-131">В папке VSCODE проекта откройте файлlaunch.js **файла.**</span><span class="sxs-lookup"><span data-stu-id="3c1ad-131">In the .vscode folder of your project, open the **launch.json** file.</span></span> <span data-ttu-id="3c1ad-132">Добавьте в раздел следующий `configurations` код:</span><span class="sxs-lookup"><span data-stu-id="3c1ad-132">Add the following code to the `configurations` section:</span></span>

```JSON
{
  "type": "office-addin",
  "request": "attach",
  "name": "Attach to Office Add-ins",
  "port": 9222,
  "trace": "verbose",
  "url": "https://localhost:3000/taskpane.html?_host_Info=HOST$Win32$16.01$en-US$$$$0",
  "webRoot": "${workspaceFolder}",
  "timeout": 45000
}
```

7. <span data-ttu-id="3c1ad-133">В разделе JSON, который вы только что скопировали, найдите раздел "URL".</span><span class="sxs-lookup"><span data-stu-id="3c1ad-133">In the section of JSON you just copied, find the "url" section.</span></span> <span data-ttu-id="3c1ad-134">В этом URL-адресе необходимо заменить верхний регистр текста HOST на приложение, в которое размещена надстройка Office.</span><span class="sxs-lookup"><span data-stu-id="3c1ad-134">In this URL, you will need to replace the uppercase HOST text with the application that is hosting your Office Add-in.</span></span> <span data-ttu-id="3c1ad-135">Например, если ваша надстройка Office для Excel, url-адрес будет иметь значение https://localhost:3000/taskpane.html?_host_Info= <strong>"Excel</strong>$Win 32$16.01$en-US$ \$ \$ \$ 0".</span><span class="sxs-lookup"><span data-stu-id="3c1ad-135">For example, if your Office Add-in is for Excel, your URL value would be "https://localhost:3000/taskpane.html?_host_Info=<strong>Excel</strong>$Win32$16.01$en-US$\$\$\$0".</span></span>

8. <span data-ttu-id="3c1ad-136">Откройте командную подсказку и убедитесь, что находитесь в корневой папке проекта.</span><span class="sxs-lookup"><span data-stu-id="3c1ad-136">Open the command prompt and ensure you are at the root folder of your project.</span></span> <span data-ttu-id="3c1ad-137">Запустите `npm start` команду, чтобы запустить сервер разработчиков.</span><span class="sxs-lookup"><span data-stu-id="3c1ad-137">Run the command `npm start` to start the dev server.</span></span> <span data-ttu-id="3c1ad-138">Когда надстройка загружается в клиенте Office, откройте области задач.</span><span class="sxs-lookup"><span data-stu-id="3c1ad-138">When your add-in loads in the Office client, open the task pane.</span></span>

9. <span data-ttu-id="3c1ad-139">Вернись к Visual Studio Code и выберите "Просмотр > **Отлаки"** или введите **CTRL + SHIFT + D,** чтобы переключиться на представление отлаки.</span><span class="sxs-lookup"><span data-stu-id="3c1ad-139">Return to Visual Studio Code and choose **View > Debug** or enter **CTRL + SHIFT + D** to switch to debug view.</span></span>

10. <span data-ttu-id="3c1ad-140">В параметрах отлаки выберите **"Присоединение к надстройкам Office".** Select **F5** or choose **Debug -> Start Debugging** from the menu to begin debugging.</span><span class="sxs-lookup"><span data-stu-id="3c1ad-140">From the Debug options, choose **Attach to Office Add-ins**. Select **F5** or choose **Debug -> Start Debugging** from the menu to begin debugging.</span></span>

11. <span data-ttu-id="3c1ad-141">Установите точку останова в файле области задач проекта.</span><span class="sxs-lookup"><span data-stu-id="3c1ad-141">Set a breakpoint in your project's task pane file.</span></span> <span data-ttu-id="3c1ad-142">Вы можете установить точки останова в VS Code, наведите курсор на строку кода и выберите красный круг.</span><span class="sxs-lookup"><span data-stu-id="3c1ad-142">You can set breakpoints in VS Code by hovering next to a line of code and selecting the red circle which appears.</span></span>

![Красный круг отображается на строке кода в VS Code](../images/set-breakpoint.jpg)

12. <span data-ttu-id="3c1ad-144">Запустите надстройку.</span><span class="sxs-lookup"><span data-stu-id="3c1ad-144">Run your add-in.</span></span> <span data-ttu-id="3c1ad-145">Вы увидите, что были сбиты точки останова, и можете проверить локальные переменные.</span><span class="sxs-lookup"><span data-stu-id="3c1ad-145">You will see that breakpoints have been hit and you can inspect local variables.</span></span>

## <a name="see-also"></a><span data-ttu-id="3c1ad-146">См. также</span><span class="sxs-lookup"><span data-stu-id="3c1ad-146">See also</span></span>

* [<span data-ttu-id="3c1ad-147">Тестирование и отладка надстроек Office</span><span class="sxs-lookup"><span data-stu-id="3c1ad-147">Test and debug Office Add-ins</span></span>](test-debug-office-add-ins.md)

* [<span data-ttu-id="3c1ad-148">Отладка надстроек с помощью средств разработчика в Windows 10</span><span class="sxs-lookup"><span data-stu-id="3c1ad-148">Debug add-ins using developer tools on Windows 10</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

* [<span data-ttu-id="3c1ad-149">Отладка надстроек в Windows с помощью Microsoft Edge WebView2 (на основе Chromium)</span><span class="sxs-lookup"><span data-stu-id="3c1ad-149">Debug add-ins on Windows using Microsoft Edge WebView2 (Chromium-based)</span></span>](debug-desktop-using-edge-chromium.md)
