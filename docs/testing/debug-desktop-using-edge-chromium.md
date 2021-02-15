---
title: Отладка настроек в Windows с использованием Microsoft Edge WebView2 (на основе Chromium)
description: Узнайте, как осуществлять отладку надстроек Office, в которых используется Microsoft Edge WebView2 (на основе Chromium) с помощью отладчика для расширения Microsoft Edge в коде VS.
ms.date: 01/29/2021
localization_priority: Priority
ms.openlocfilehash: 0908bb5040b49568006324600acacb5e36dbd1a5
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/14/2021
ms.locfileid: "50238118"
---
# <a name="debug-add-ins-on-windows-using-edge-chromium-webview2"></a><span data-ttu-id="e5d28-103">Отладка надстроек в Windows с помощью Edge Chromium WebView2</span><span class="sxs-lookup"><span data-stu-id="e5d28-103">Debug add-ins on Windows using Edge Chromium WebView2</span></span>

<span data-ttu-id="e5d28-104">Надстройки Office, работающие в Windows, могут использовать отладчик для расширения Microsoft Edge в коде VS для отладки среды Edge Chromium WebView2.</span><span class="sxs-lookup"><span data-stu-id="e5d28-104">Office Add-ins running on Windows can use the Debugger for Microsoft Edge extension in VS Code to debug against the Edge Chromium WebView2 runtime.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="e5d28-105">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="e5d28-105">Prerequisites</span></span>

- <span data-ttu-id="e5d28-106">[Код Visual Studio](https://code.visualstudio.com/) (необходимо запускать от имени администратора)</span><span class="sxs-lookup"><span data-stu-id="e5d28-106">[Visual Studio Code](https://code.visualstudio.com/) (must be run as an administrator)</span></span>
- [<span data-ttu-id="e5d28-107">Node.js (версия 10. или более поздняя)</span><span class="sxs-lookup"><span data-stu-id="e5d28-107">Node.js (version 10+)</span></span>](https://nodejs.org/)
- <span data-ttu-id="e5d28-108">Windows 10</span><span class="sxs-lookup"><span data-stu-id="e5d28-108">Windows 10</span></span>
- [<span data-ttu-id="e5d28-109">Microsoft Edge Chromium доступна участникам программы предварительной оценки Windows</span><span class="sxs-lookup"><span data-stu-id="e5d28-109">Microsoft Edge Chromium available to Windows Insiders</span></span>](https://www.microsoftedgeinsider.com/)

## <a name="install-and-use-the-debugger"></a><span data-ttu-id="e5d28-110">Установка и использование отладчика</span><span class="sxs-lookup"><span data-stu-id="e5d28-110">Install and use the debugger</span></span>

1. <span data-ttu-id="e5d28-111">Создайте проект с помощью [генератора Yeoman для надстроек Office](https://github.com/OfficeDev/generator-office). Для этого можно использовать любые краткие руководства по началу работы, например [Краткое руководство по надстройкам Outlook](../quickstarts/outlook-quickstart.md).</span><span class="sxs-lookup"><span data-stu-id="e5d28-111">Create a project using the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office). You can use any one of our quick start guides, such as the [Outlook add-in quickstart](../quickstarts/outlook-quickstart.md), in order to do this.</span></span>

> [!TIP]
> <span data-ttu-id="e5d28-112">Если вы не используете надстройку, основанную на генераторе Yeoman, необходимо настроить ключ реестра.</span><span class="sxs-lookup"><span data-stu-id="e5d28-112">If you aren't using a Yeoman generator based add-in, you need to adjust a registry key.</span></span> <span data-ttu-id="e5d28-113">В корневой папке проекта выполните указанные ниже действия в командной строке: `office-add-in-debugging start <your manifest path>`.</span><span class="sxs-lookup"><span data-stu-id="e5d28-113">While in the root folder of your project, run the following in the command line: `office-add-in-debugging start <your manifest path>`.</span></span>

2. <span data-ttu-id="e5d28-114">Откройте проект в VS Code.</span><span class="sxs-lookup"><span data-stu-id="e5d28-114">Open your project in VS Code.</span></span> <span data-ttu-id="e5d28-115">Находясь в коде VS, нажмите **CTRL + SHIFT + X**, чтобы открыть меню расширений.</span><span class="sxs-lookup"><span data-stu-id="e5d28-115">Within VS Code, select **CTRL + SHIFT + X** to open the Extensions bar.</span></span> <span data-ttu-id="e5d28-116">Выполните поиск расширения "Debugger для Microsoft Edge" и установите его.</span><span class="sxs-lookup"><span data-stu-id="e5d28-116">Search for the "Debugger for Microsoft Edge" extension and install it.</span></span>

3. <span data-ttu-id="e5d28-117">В папке проекта **. vscode** проекта откройте файл **launch.json**.</span><span class="sxs-lookup"><span data-stu-id="e5d28-117">In the **.vscode** folder of your project, open the **launch.json** file.</span></span> <span data-ttu-id="e5d28-118">Добавьте указанный ниже код в раздел конфигураций.</span><span class="sxs-lookup"><span data-stu-id="e5d28-118">Add the following code to the configurations section:</span></span>

```JSON
  {
      "name": "Debug Office Add-in (Edge Chromium)",
      "type": "edge",
      "request": "attach",
      "useWebView": "advanced",
      "port": 9229,
      "timeout": 600000,
      "webRoot": "${workspaceRoot}",
    },
```

4. <span data-ttu-id="e5d28-119">Чтобы перейти к представлению отладки, нажмите **Просмотр> Отладка** или введите **CTRL + SHIFT + D**.</span><span class="sxs-lookup"><span data-stu-id="e5d28-119">Next, choose  **View > Debug** or enter **CTRL + SHIFT + D** to switch to debug view.</span></span>

5. <span data-ttu-id="e5d28-120">В разделе параметров отладки выберите параметр Edge Chromium для ведущего приложения, например **классического приложения Excel (Edge Chromium)**.</span><span class="sxs-lookup"><span data-stu-id="e5d28-120">From the Debug options, choose the Edge Chromium option for your host application, such as **Excel Desktop (Edge Chromium)**.</span></span> <span data-ttu-id="e5d28-121">Чтобы начать отладку, нажмите **F5** или выберите **Отладка > Начать отладку** в меню.</span><span class="sxs-lookup"><span data-stu-id="e5d28-121">Select **F5** or choose **Debug > Start Debugging** from the menu to begin debugging.</span></span>

6. <span data-ttu-id="e5d28-122">Теперь надстройка готова к использованию в ведущем приложении, таком как Excel.</span><span class="sxs-lookup"><span data-stu-id="e5d28-122">In the host application, such as Excel, your add-in is now ready to use.</span></span> <span data-ttu-id="e5d28-123">Нажмите кнопку **Показать область задач** или выполнить другие дополнительные команды надстройки.</span><span class="sxs-lookup"><span data-stu-id="e5d28-123">Select **Show Taskpane** or run any other add-in command.</span></span> <span data-ttu-id="e5d28-124">Появится диалоговое окно подтверждения действия с надписью</span><span class="sxs-lookup"><span data-stu-id="e5d28-124">A dialog box will appear, reading:</span></span>

> <span data-ttu-id="e5d28-125">WebView Stop On Load.</span><span class="sxs-lookup"><span data-stu-id="e5d28-125">WebView Stop On Load.</span></span> 
> <span data-ttu-id="e5d28-126">Чтобы выполнить отладку WebView, вложите код VS в экземпляр WebView с помощью отладчика Microsoft для Edge и нажмите кнопку ОК.</span><span class="sxs-lookup"><span data-stu-id="e5d28-126">To debug the webview, attach VS Code to the webview instance using the Microsoft Debugger for Edge extension, and click OK to continue.</span></span> <span data-ttu-id="e5d28-127">Чтобы предотвратить появление диалогового окна в дальнейшем, нажмите кнопку"Отмена".</span><span class="sxs-lookup"><span data-stu-id="e5d28-127">To prevent this dialog from appearing in the future, click Cancel."</span></span>

<span data-ttu-id="e5d28-128">Нажмите **ОК**.</span><span class="sxs-lookup"><span data-stu-id="e5d28-128">Select **OK**.</span></span>

> [!NOTE]
> <span data-ttu-id="e5d28-129">После нажатия кнопки **Отмена** диалоговое окно не будет отображаться в процессе работы с этим экземпляром надстройки.</span><span class="sxs-lookup"><span data-stu-id="e5d28-129">If you select **Cancel**, the dialog won't be shown again while this instance of the add-in is running.</span></span> <span data-ttu-id="e5d28-130">Однако при перезапуске надстройки диалоговое окно снова появится.</span><span class="sxs-lookup"><span data-stu-id="e5d28-130">However, if you restart your add-in, you'll see the dialog again.</span></span>

7. <span data-ttu-id="e5d28-131">Теперь можно задать точки останова в коде проекта и выполнить отладку.</span><span class="sxs-lookup"><span data-stu-id="e5d28-131">You're now able to set breakpoints in your project's code and debug.</span></span>

## <a name="see-also"></a><span data-ttu-id="e5d28-132">См. также</span><span class="sxs-lookup"><span data-stu-id="e5d28-132">See also</span></span>

* [<span data-ttu-id="e5d28-133">Тестирование и отладка надстроек Office</span><span class="sxs-lookup"><span data-stu-id="e5d28-133">Test and debug Office Add-ins</span></span>](test-debug-office-add-ins.md)
* [<span data-ttu-id="e5d28-134">Надстройка Microsoft Office "Расширение отладчика для Visual Studio Code"</span><span class="sxs-lookup"><span data-stu-id="e5d28-134">Microsoft Office Add-in Debugger Extension for Visual Studio Code</span></span>](debug-with-vs-extension.md)
* [<span data-ttu-id="e5d28-135">Подключение отладчика из области задач</span><span class="sxs-lookup"><span data-stu-id="e5d28-135">Attach a debugger from the task pane</span></span>](attach-debugger-from-task-pane.md)