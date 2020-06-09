---
title: Расширение отладчика надстроек Microsoft Office для Visual Studio Code
description: Используйте отладчик надстроек Microsoft Office с расширением кода Visual Studio, чтобы отладить надстройку Office.
ms.date: 05/14/2020
localization_priority: Normal
ms.openlocfilehash: 1bd3814eba6da2339e7865d720b8a4c792b9310e
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611213"
---
# <a name="microsoft-office-add-in-debugger-extension-for-visual-studio-code"></a><span data-ttu-id="93b5f-103">Расширение отладчика надстроек Microsoft Office для Visual Studio Code</span><span class="sxs-lookup"><span data-stu-id="93b5f-103">Microsoft Office Add-in Debugger Extension for Visual Studio Code</span></span>

<span data-ttu-id="93b5f-104">Расширение отладчика надстроек Microsoft Office для Visual Studio Code позволяет отлаживать надстройку Office в пограничной среде выполнения.</span><span class="sxs-lookup"><span data-stu-id="93b5f-104">The Microsoft Office Add-in Debugger Extension for Visual Studio Code allows you to debug your Office Add-in against the Edge runtime.</span></span>

<span data-ttu-id="93b5f-105">Этот режим отладки динамический, позволяющий задавать точки останова во время выполнения кода.</span><span class="sxs-lookup"><span data-stu-id="93b5f-105">This debugging mode is dynamic, allowing you to set breakpoints while code is running.</span></span> <span data-ttu-id="93b5f-106">Вы можете видеть изменения в коде сразу же после присоединения отладчика, все без потери сеанса отладки.</span><span class="sxs-lookup"><span data-stu-id="93b5f-106">You can see changes in your code immediately while the debugger is attached, all without losing your debugging session.</span></span> <span data-ttu-id="93b5f-107">Изменения в коде также остаются неизменными, поэтому вы можете увидеть результаты нескольких изменений в коде.</span><span class="sxs-lookup"><span data-stu-id="93b5f-107">Your code changes also persist, so you can see the results of multiple changes to your code.</span></span> <span data-ttu-id="93b5f-108">На следующем изображении показано это расширение в действии.</span><span class="sxs-lookup"><span data-stu-id="93b5f-108">The following image shows this extension in action.</span></span>

![Отладка расширения отладчика надстроек Office раздел надстройки Excel](../images/vs-debugger-extension-for-office-addins.jpg)

## <a name="prerequisites"></a><span data-ttu-id="93b5f-110">Предварительные требования</span><span class="sxs-lookup"><span data-stu-id="93b5f-110">Prerequisites</span></span>

- <span data-ttu-id="93b5f-111">[Visual Studio Code](https://code.visualstudio.com/) (необходимо запускать от имени администратора)</span><span class="sxs-lookup"><span data-stu-id="93b5f-111">[Visual Studio Code](https://code.visualstudio.com/) (must be run as an administrator)</span></span>
- [<span data-ttu-id="93b5f-112">Node. js (версия 10 +)</span><span class="sxs-lookup"><span data-stu-id="93b5f-112">Node.js (version 10+)</span></span>](https://nodejs.org/)
- <span data-ttu-id="93b5f-113">Windows 10</span><span class="sxs-lookup"><span data-stu-id="93b5f-113">Windows 10</span></span>
- [<span data-ttu-id="93b5f-114">Microsoft Edge</span><span class="sxs-lookup"><span data-stu-id="93b5f-114">Microsoft Edge</span></span>](https://www.microsoft.com/edge)

<span data-ttu-id="93b5f-115">В этих инструкциях предполагается, что у вас есть опыт работы с помощью командной строки, общие сведения об основном коде JavaScript и создание проекта надстройки Office перед использованием генератора Yo Office.</span><span class="sxs-lookup"><span data-stu-id="93b5f-115">These instructions assume you have experience using the command line, understand basic JavaScript, and have created an Office add-in project before using the Yo Office generator.</span></span> <span data-ttu-id="93b5f-116">Если вы еще этого не сделали, ознакомьтесь с одним из наших руководств, как в этом [руководстве по надстройкам Office для Excel](../tutorials/excel-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="93b5f-116">If you haven't done this before, consider visiting one of our tutorials, like this [Excel Office Add-in tutorial](../tutorials/excel-tutorial.md).</span></span>

## <a name="install-and-use-the-debugger"></a><span data-ttu-id="93b5f-117">Установка и использование отладчика</span><span class="sxs-lookup"><span data-stu-id="93b5f-117">Install and use the debugger</span></span>

1. <span data-ttu-id="93b5f-118">Если вам нужно создать проект надстройки, [создайте его с помощью генератора Yo Office](https://docs.microsoft.com/office/dev/add-ins/quickstarts/excel-quickstart-jquery?tabs=yeomangenerator).</span><span class="sxs-lookup"><span data-stu-id="93b5f-118">If you need to create an add-in project, [use the Yo Office generator to create one](https://docs.microsoft.com/office/dev/add-ins/quickstarts/excel-quickstart-jquery?tabs=yeomangenerator).</span></span> <span data-ttu-id="93b5f-119">Чтобы настроить проект, следуйте инструкциям в командной строке.</span><span class="sxs-lookup"><span data-stu-id="93b5f-119">Follow the prompts within the command line to set up your project.</span></span> <span data-ttu-id="93b5f-120">Вы можете выбрать любой язык или тип проекта в соответствии со своими потребностями.</span><span class="sxs-lookup"><span data-stu-id="93b5f-120">You can choose any language or type of project to suit your needs.</span></span>

> <span data-ttu-id="93b5f-121">! НОТЕ Если у вас уже есть проект, пропустите шаг 1 и перейдите к шагу 2.</span><span class="sxs-lookup"><span data-stu-id="93b5f-121">![NOTE] If you already have a project, skip step 1 and move to step 2.</span></span>

2. <span data-ttu-id="93b5f-122">Откройте командную строку от имени администратора.</span><span class="sxs-lookup"><span data-stu-id="93b5f-122">Open a command prompt as administrator.</span></span>
   <span data-ttu-id="93b5f-123">![Параметры командной строки, в том числе "Запуск от имени администратора" в Windows 10](../images/run-as-administrator-vs-code.jpg)</span><span class="sxs-lookup"><span data-stu-id="93b5f-123">![Command prompt options, including "run as administrator" in Windows 10](../images/run-as-administrator-vs-code.jpg)</span></span>

3. <span data-ttu-id="93b5f-124">Перейдите к каталогу проекта.</span><span class="sxs-lookup"><span data-stu-id="93b5f-124">Navigate to your project directory.</span></span>

4. <span data-ttu-id="93b5f-125">Выполните следующую команду, чтобы открыть проект в Visual Studio Code от имени администратора.</span><span class="sxs-lookup"><span data-stu-id="93b5f-125">Run the following command to open your project in Visual Studio Code as an administrator.</span></span>

```command&nbsp;line
code .
```

<span data-ttu-id="93b5f-126">После открытия кода Visual Studio перейдите в папку проекта вручную.</span><span class="sxs-lookup"><span data-stu-id="93b5f-126">Once Visual Studio Code is open, navigate manually to the project folder.</span></span>

> [!TIP]
> <span data-ttu-id="93b5f-127">Чтобы открыть Visual Studio Code от имени администратора, установите флажок **Запуск от имени администратора** при открытии кода Visual Studio после его поиска в Windows.</span><span class="sxs-lookup"><span data-stu-id="93b5f-127">To open Visual Studio Code as an administrator, select the **run as administrator** option when opening Visual Studio Code after searching for it in Windows.</span></span>

5. <span data-ttu-id="93b5f-128">В рамках кода VS нажмите **клавиши CTRL + SHIFT + X** , чтобы открыть панель расширений.</span><span class="sxs-lookup"><span data-stu-id="93b5f-128">Within VS Code, select **CTRL + SHIFT + X** to open the Extensions bar.</span></span> <span data-ttu-id="93b5f-129">Выполните поиск расширения "надстройка Microsoft Office Debugger Debugger" и установите его.</span><span class="sxs-lookup"><span data-stu-id="93b5f-129">Search for the "Microsoft Office Add-in Debugger" extension and install it.</span></span>

6. <span data-ttu-id="93b5f-130">В папке. вскоде проекта откройте файл **Launch. JSON** .</span><span class="sxs-lookup"><span data-stu-id="93b5f-130">In the .vscode folder of your project, open the **launch.json** file.</span></span> <span data-ttu-id="93b5f-131">Добавьте в раздел следующий код `configurations` :</span><span class="sxs-lookup"><span data-stu-id="93b5f-131">Add the following code to the `configurations` section:</span></span>

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

7. <span data-ttu-id="93b5f-132">В разделе только что скопированный JSON найдите раздел "URL".</span><span class="sxs-lookup"><span data-stu-id="93b5f-132">In the section of JSON you just copied, find the "url" section.</span></span> <span data-ttu-id="93b5f-133">В этом URL-адресе необходимо заменить текст узла в верхнем регистре на ведущее приложение для надстройки Office.</span><span class="sxs-lookup"><span data-stu-id="93b5f-133">In this URL, you will need to replace the uppercase HOST text with the host application for your Office add-in.</span></span> <span data-ttu-id="93b5f-134">Например, если надстройка Office предназначена для Excel, URL-адрес будет иметь значение " https://localhost:3000/taskpane.html?_host_Info= <strong>Excel</strong>$Win 32 $16.01 $ en-US $ \$ \$ \$ 0".</span><span class="sxs-lookup"><span data-stu-id="93b5f-134">For example, if your Office add-in is for Excel, your URL value would be "https://localhost:3000/taskpane.html?_host_Info=<strong>Excel</strong>$Win32$16.01$en-US$\$\$\$0".</span></span>

8. <span data-ttu-id="93b5f-135">Откройте командную строку и убедитесь в наличии корневой папки проекта.</span><span class="sxs-lookup"><span data-stu-id="93b5f-135">Open the command prompt and ensure you are at the root folder of your project.</span></span> <span data-ttu-id="93b5f-136">Выполните команду, `npm start` чтобы запустить сервер разработки.</span><span class="sxs-lookup"><span data-stu-id="93b5f-136">Run the command `npm start` to start the dev server.</span></span> <span data-ttu-id="93b5f-137">Когда надстройка загружается в клиенте Office, откройте область задач.</span><span class="sxs-lookup"><span data-stu-id="93b5f-137">When your add-in loads in the Office client, open the task pane.</span></span>

9. <span data-ttu-id="93b5f-138">Вернитесь к Visual Studio Code и выберите **просмотр > Отладка** или ввод **CTRL + SHIFT + D** , чтобы перейти в представление отладки.</span><span class="sxs-lookup"><span data-stu-id="93b5f-138">Return to Visual Studio Code and choose **View > Debug** or enter **CTRL + SHIFT + D** to switch to debug view.</span></span>

10. <span data-ttu-id="93b5f-139">В разделе Параметры отладки выберите команду **присоединиться к**надстройкам Office. Нажмите **клавишу F5** или выберите **Debug — > начать отладку** в меню, чтобы начать отладку.</span><span class="sxs-lookup"><span data-stu-id="93b5f-139">From the Debug options, choose **Attach to Office Add-ins**. Select **F5** or choose **Debug -> Start Debugging** from the menu to begin debugging.</span></span>

11. <span data-ttu-id="93b5f-140">Задайте точку останова в файле области задач проекта.</span><span class="sxs-lookup"><span data-stu-id="93b5f-140">Set a breakpoint in your project's task pane file.</span></span> <span data-ttu-id="93b5f-141">Вы можете задать точки останова в коде VS, наведя курсор рядом с строкой кода и выбрав красный круг.</span><span class="sxs-lookup"><span data-stu-id="93b5f-141">You can set breakpoints in VS Code by hovering next to a line of code and selecting the red circle which appears.</span></span>

![В строке кода в VS отображается красный кружок](../images/set-breakpoint.jpg)

12. <span data-ttu-id="93b5f-143">Запустите надстройку.</span><span class="sxs-lookup"><span data-stu-id="93b5f-143">Run your add-in.</span></span> <span data-ttu-id="93b5f-144">Вы увидите, что достигнуты точки останова, и можете проверить локальные переменные.</span><span class="sxs-lookup"><span data-stu-id="93b5f-144">You will see that breakpoints have been hit and you can inspect local variables.</span></span>

## <a name="see-also"></a><span data-ttu-id="93b5f-145">См. также</span><span class="sxs-lookup"><span data-stu-id="93b5f-145">See also</span></span>

* [<span data-ttu-id="93b5f-146">Тестирование и отладка надстроек Office</span><span class="sxs-lookup"><span data-stu-id="93b5f-146">Test and debug Office Add-ins</span></span>](test-debug-office-add-ins.md)

* [<span data-ttu-id="93b5f-147">Отладка надстроек с помощью средств разработчика в Windows 10</span><span class="sxs-lookup"><span data-stu-id="93b5f-147">Debug add-ins using developer tools on Windows 10</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)

* [<span data-ttu-id="93b5f-148">Подключение отладчика из области задач</span><span class="sxs-lookup"><span data-stu-id="93b5f-148">Attach a debugger from the task pane</span></span>](attach-debugger-from-task-pane.md)
