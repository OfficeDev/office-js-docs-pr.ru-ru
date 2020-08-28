---
title: Надстройка Microsoft Office "Расширение отладчика для Visual Studio Code"
description: Используйте отладчик надстроек Microsoft Office с расширением кода Visual Studio, чтобы отладить надстройку Office.
ms.date: 05/14/2020
localization_priority: Normal
ms.openlocfilehash: 1343014fa875509fd6f0c615c3504cc9ae50dc0d
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293445"
---
# <a name="microsoft-office-add-in-debugger-extension-for-visual-studio-code"></a><span data-ttu-id="9b200-103">Надстройка Microsoft Office "Расширение отладчика для Visual Studio Code"</span><span class="sxs-lookup"><span data-stu-id="9b200-103">Microsoft Office Add-in Debugger Extension for Visual Studio Code</span></span>

<span data-ttu-id="9b200-104">Расширение отладчика надстроек Microsoft Office для Visual Studio Code позволяет отлаживать надстройку Office в пограничной среде выполнения.</span><span class="sxs-lookup"><span data-stu-id="9b200-104">The Microsoft Office Add-in Debugger Extension for Visual Studio Code allows you to debug your Office Add-in against the Edge runtime.</span></span>

<span data-ttu-id="9b200-105">Этот режим отладки динамический, позволяющий задавать точки останова во время выполнения кода.</span><span class="sxs-lookup"><span data-stu-id="9b200-105">This debugging mode is dynamic, allowing you to set breakpoints while code is running.</span></span> <span data-ttu-id="9b200-106">Вы можете видеть изменения в коде сразу же после присоединения отладчика, все без потери сеанса отладки.</span><span class="sxs-lookup"><span data-stu-id="9b200-106">You can see changes in your code immediately while the debugger is attached, all without losing your debugging session.</span></span> <span data-ttu-id="9b200-107">Изменения в коде также остаются неизменными, поэтому вы можете увидеть результаты нескольких изменений в коде.</span><span class="sxs-lookup"><span data-stu-id="9b200-107">Your code changes also persist, so you can see the results of multiple changes to your code.</span></span> <span data-ttu-id="9b200-108">На следующем изображении показано это расширение в действии.</span><span class="sxs-lookup"><span data-stu-id="9b200-108">The following image shows this extension in action.</span></span>

![Отладка расширения отладчика надстроек Office раздел надстройки Excel](../images/vs-debugger-extension-for-office-addins.jpg)

## <a name="prerequisites"></a><span data-ttu-id="9b200-110">Предварительные условия</span><span class="sxs-lookup"><span data-stu-id="9b200-110">Prerequisites</span></span>

- <span data-ttu-id="9b200-111">[Visual Studio Code](https://code.visualstudio.com/) (необходимо запускать от имени администратора)</span><span class="sxs-lookup"><span data-stu-id="9b200-111">[Visual Studio Code](https://code.visualstudio.com/) (must be run as an administrator)</span></span>
- [<span data-ttu-id="9b200-112">Node.js (версия 10 +)</span><span class="sxs-lookup"><span data-stu-id="9b200-112">Node.js (version 10+)</span></span>](https://nodejs.org/)
- <span data-ttu-id="9b200-113">Windows 10</span><span class="sxs-lookup"><span data-stu-id="9b200-113">Windows 10</span></span>
- [<span data-ttu-id="9b200-114">Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="9b200-114">Microsoft Edge</span></span>](https://www.microsoft.com/edge)

<span data-ttu-id="9b200-115">В этих инструкциях предполагается, что у вас есть опыт работы с помощью командной строки, общие сведения об основном коде JavaScript и создание проекта надстройки Office перед использованием генератора Yo Office.</span><span class="sxs-lookup"><span data-stu-id="9b200-115">These instructions assume you have experience using the command line, understand basic JavaScript, and have created an Office add-in project before using the Yo Office generator.</span></span> <span data-ttu-id="9b200-116">Если вы еще этого не сделали, ознакомьтесь с одним из наших руководств, как в этом [руководстве по надстройкам Office для Excel](../tutorials/excel-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="9b200-116">If you haven't done this before, consider visiting one of our tutorials, like this [Excel Office Add-in tutorial](../tutorials/excel-tutorial.md).</span></span>

## <a name="install-and-use-the-debugger"></a><span data-ttu-id="9b200-117">Установка и использование отладчика</span><span class="sxs-lookup"><span data-stu-id="9b200-117">Install and use the debugger</span></span>

1. <span data-ttu-id="9b200-118">Если вам нужно создать проект надстройки, [создайте его с помощью генератора Yo Office](https://docs.microsoft.com/office/dev/add-ins/quickstarts/excel-quickstart-jquery?tabs=yeomangenerator).</span><span class="sxs-lookup"><span data-stu-id="9b200-118">If you need to create an add-in project, [use the Yo Office generator to create one](https://docs.microsoft.com/office/dev/add-ins/quickstarts/excel-quickstart-jquery?tabs=yeomangenerator).</span></span> <span data-ttu-id="9b200-119">Чтобы настроить проект, следуйте инструкциям в командной строке.</span><span class="sxs-lookup"><span data-stu-id="9b200-119">Follow the prompts within the command line to set up your project.</span></span> <span data-ttu-id="9b200-120">Вы можете выбрать любой язык или тип проекта в соответствии со своими потребностями.</span><span class="sxs-lookup"><span data-stu-id="9b200-120">You can choose any language or type of project to suit your needs.</span></span>

> [!NOTE]
> <span data-ttu-id="9b200-121">Если у вас уже есть проект, пропустите шаг 1 и перейдите к шагу 2.</span><span class="sxs-lookup"><span data-stu-id="9b200-121">If you already have a project, skip step 1 and move to step 2.</span></span>

2. <span data-ttu-id="9b200-122">Откройте командную строку от имени администратора.</span><span class="sxs-lookup"><span data-stu-id="9b200-122">Open a command prompt as administrator.</span></span>
   <span data-ttu-id="9b200-123">![Параметры командной строки, в том числе "Запуск от имени администратора" в Windows 10](../images/run-as-administrator-vs-code.jpg)</span><span class="sxs-lookup"><span data-stu-id="9b200-123">![Command prompt options, including "run as administrator" in Windows 10](../images/run-as-administrator-vs-code.jpg)</span></span>

3. <span data-ttu-id="9b200-124">Перейдите к каталогу проекта.</span><span class="sxs-lookup"><span data-stu-id="9b200-124">Navigate to your project directory.</span></span>

4. <span data-ttu-id="9b200-125">Выполните следующую команду, чтобы открыть проект в Visual Studio Code от имени администратора.</span><span class="sxs-lookup"><span data-stu-id="9b200-125">Run the following command to open your project in Visual Studio Code as an administrator.</span></span>

```command&nbsp;line
code .
```

<span data-ttu-id="9b200-126">После открытия кода Visual Studio перейдите в папку проекта вручную.</span><span class="sxs-lookup"><span data-stu-id="9b200-126">Once Visual Studio Code is open, navigate manually to the project folder.</span></span>

> [!TIP]
> <span data-ttu-id="9b200-127">Чтобы открыть Visual Studio Code от имени администратора, установите флажок **Запуск от имени администратора** при открытии кода Visual Studio после его поиска в Windows.</span><span class="sxs-lookup"><span data-stu-id="9b200-127">To open Visual Studio Code as an administrator, select the **run as administrator** option when opening Visual Studio Code after searching for it in Windows.</span></span>

5. <span data-ttu-id="9b200-128">В рамках кода VS нажмите **клавиши CTRL + SHIFT + X** , чтобы открыть панель расширений.</span><span class="sxs-lookup"><span data-stu-id="9b200-128">Within VS Code, select **CTRL + SHIFT + X** to open the Extensions bar.</span></span> <span data-ttu-id="9b200-129">Выполните поиск расширения "надстройка Microsoft Office Debugger Debugger" и установите его.</span><span class="sxs-lookup"><span data-stu-id="9b200-129">Search for the "Microsoft Office Add-in Debugger" extension and install it.</span></span>

6. <span data-ttu-id="9b200-130">В папке. вскоде проекта откройте **launch.jsв** файле.</span><span class="sxs-lookup"><span data-stu-id="9b200-130">In the .vscode folder of your project, open the **launch.json** file.</span></span> <span data-ttu-id="9b200-131">Добавьте в раздел следующий код `configurations` :</span><span class="sxs-lookup"><span data-stu-id="9b200-131">Add the following code to the `configurations` section:</span></span>

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

7. <span data-ttu-id="9b200-132">В разделе только что скопированный JSON найдите раздел "URL".</span><span class="sxs-lookup"><span data-stu-id="9b200-132">In the section of JSON you just copied, find the "url" section.</span></span> <span data-ttu-id="9b200-133">В этом URL-адресе необходимо заменить текст узла в верхнем регистре на приложение, в котором размещается надстройка Office.</span><span class="sxs-lookup"><span data-stu-id="9b200-133">In this URL, you will need to replace the uppercase HOST text with the application that is hosting your Office add-in.</span></span> <span data-ttu-id="9b200-134">Например, если надстройка Office предназначена для Excel, URL-адрес будет иметь значение " https://localhost:3000/taskpane.html?_host_Info= <strong>Excel</strong>$Win 32 $16.01 $ en-US $ \$ \$ \$ 0".</span><span class="sxs-lookup"><span data-stu-id="9b200-134">For example, if your Office add-in is for Excel, your URL value would be "https://localhost:3000/taskpane.html?_host_Info=<strong>Excel</strong>$Win32$16.01$en-US$\$\$\$0".</span></span>

8. <span data-ttu-id="9b200-135">Откройте командную строку и убедитесь в наличии корневой папки проекта.</span><span class="sxs-lookup"><span data-stu-id="9b200-135">Open the command prompt and ensure you are at the root folder of your project.</span></span> <span data-ttu-id="9b200-136">Выполните команду, `npm start` чтобы запустить сервер разработки.</span><span class="sxs-lookup"><span data-stu-id="9b200-136">Run the command `npm start` to start the dev server.</span></span> <span data-ttu-id="9b200-137">Когда надстройка загружается в клиенте Office, откройте область задач.</span><span class="sxs-lookup"><span data-stu-id="9b200-137">When your add-in loads in the Office client, open the task pane.</span></span>

9. <span data-ttu-id="9b200-138">Вернитесь к Visual Studio Code и выберите **просмотр > Отладка** или ввод **CTRL + SHIFT + D** , чтобы перейти в представление отладки.</span><span class="sxs-lookup"><span data-stu-id="9b200-138">Return to Visual Studio Code and choose **View > Debug** or enter **CTRL + SHIFT + D** to switch to debug view.</span></span>

10. <span data-ttu-id="9b200-139">В разделе Параметры отладки выберите команду **присоединиться к**надстройкам Office. Нажмите **клавишу F5** или выберите **Debug — > начать отладку** в меню, чтобы начать отладку.</span><span class="sxs-lookup"><span data-stu-id="9b200-139">From the Debug options, choose **Attach to Office Add-ins**. Select **F5** or choose **Debug -> Start Debugging** from the menu to begin debugging.</span></span>

11. <span data-ttu-id="9b200-140">Задайте точку останова в файле области задач проекта.</span><span class="sxs-lookup"><span data-stu-id="9b200-140">Set a breakpoint in your project's task pane file.</span></span> <span data-ttu-id="9b200-141">Вы можете задать точки останова в коде VS, наведя курсор рядом с строкой кода и выбрав красный круг.</span><span class="sxs-lookup"><span data-stu-id="9b200-141">You can set breakpoints in VS Code by hovering next to a line of code and selecting the red circle which appears.</span></span>

![В строке кода в VS отображается красный кружок](../images/set-breakpoint.jpg)

12. <span data-ttu-id="9b200-143">Запустите надстройку.</span><span class="sxs-lookup"><span data-stu-id="9b200-143">Run your add-in.</span></span> <span data-ttu-id="9b200-144">Вы увидите, что достигнуты точки останова, и можете проверить локальные переменные.</span><span class="sxs-lookup"><span data-stu-id="9b200-144">You will see that breakpoints have been hit and you can inspect local variables.</span></span>

## <a name="see-also"></a><span data-ttu-id="9b200-145">См. также</span><span class="sxs-lookup"><span data-stu-id="9b200-145">See also</span></span>

* [<span data-ttu-id="9b200-146">Тестирование и отладка надстроек Office</span><span class="sxs-lookup"><span data-stu-id="9b200-146">Test and debug Office Add-ins</span></span>](test-debug-office-add-ins.md)

* [<span data-ttu-id="9b200-147">Отладка надстроек с помощью средств разработчика в Windows 10</span><span class="sxs-lookup"><span data-stu-id="9b200-147">Debug add-ins using developer tools on Windows 10</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

* [<span data-ttu-id="9b200-148">Подключение отладчика из области задач</span><span class="sxs-lookup"><span data-stu-id="9b200-148">Attach a debugger from the task pane</span></span>](attach-debugger-from-task-pane.md)
