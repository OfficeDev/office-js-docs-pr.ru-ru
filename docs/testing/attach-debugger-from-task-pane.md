---
title: Подключение отладчика из области задач
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 941a67695f573cbf0955038916cb1e7ada9cc08f
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/27/2019
ms.locfileid: "30870431"
---
# <a name="attach-a-debugger-from-the-task-pane"></a><span data-ttu-id="6e827-102">Подключение отладчика из области задач</span><span class="sxs-lookup"><span data-stu-id="6e827-102">Attach a debugger from the task pane</span></span>

<span data-ttu-id="6e827-p101">В Office 2016 для Windows (сборка 77xx.xxxx или более поздней версии) можно подключать отладчик из области задач. Функция "Подключить отладчик" подключит отладчик непосредственно к нужному процессу Internet Explorer. Вы можете подключить отладчик независимо от того, какой инструмент используете: генератор Yeoman, Visual Studio Code, node.js, Angular или другой.</span><span class="sxs-lookup"><span data-stu-id="6e827-p101">In Office 2016 for Windows, Build 77xx.xxxx or later, you can attach the debugger from the task pane. The attach debugger feature will directly attach the debugger to the correct Internet Explorer process for you. You can attach a debugger regardless of whether you are using Yeoman Generator, Visual Studio Code, node.js, Angular, or another tool.</span></span> 

<span data-ttu-id="6e827-106">Для запуска средства **подключения отладчика** откройте меню **Личные данные** в правом верхнем углу области задач (выделено красным на рисунке ниже).</span><span class="sxs-lookup"><span data-stu-id="6e827-106">To launch the **Attach Debugger** tool, choose the top right corner of the task pane to activate the **Personality** menu (as shown in the red circle in the following image).</span></span>   

> [!NOTE]
> - <span data-ttu-id="6e827-p102">В настоящее время поддерживается только отладчик [Visual Studio 2015](https://www.visualstudio.com/downloads/) с [обновлением 3](https://msdn.microsoft.com/library/mt752379.aspx) или более поздней версии. Если у вас нет Visual Studio, выбор параметра **Подключить отладчик** не даст результата.</span><span class="sxs-lookup"><span data-stu-id="6e827-p102">Currently the only supported debugger tool is [Visual Studio 2015](https://www.visualstudio.com/downloads/) with [Update 3](https://msdn.microsoft.com/library/mt752379.aspx) or later. If you don't have Visual Studio installed, selecting the **Attach Debugger** option doesn’t result in any action.</span></span>   
> - <span data-ttu-id="6e827-p103">Для отладки клиентского кода JavaScript можно использовать только средство **Подключить отладчик**. Для отладки серверного кода, например на сервере Node.js, существует множество вариантов. Сведения о том, как выполнять отладку в Visual Studio Code, см. в статье [Отладка Node.js в VS Code](https://code.visualstudio.com/docs/nodejs/nodejs-debugging). Если вы не используете Visual Studio Code, выполните поиск по запросу "отладка Node.js" или "отладка {имя_сервера}".</span><span class="sxs-lookup"><span data-stu-id="6e827-p103">You can only debug client-side JavaScript with the **Attach Debugger** tool. To debug server-side code, such as with a Node.js server, you have many options. For information on how to debug with Visual Studio Code, see [Node.js Debugging in VS Code](https://code.visualstudio.com/docs/nodejs/nodejs-debugging). If you are not using Visual Studio Code, search for "debug Node.js" or "debug {name-of-server}".</span></span>

![Снимок экрана: меню подключения отладчика](../images/attach-debugger.png)

<span data-ttu-id="6e827-p104">Выберите элемент **Подключить отладчик**. Откроется диалоговое окно **JIT-отладчик Visual Studio** (см. рисунок ниже).</span><span class="sxs-lookup"><span data-stu-id="6e827-p104">Select **Attach Debugger**. This launches the **Visual Studio Just-in-Time Debugger** dialog box, as shown in the following image.</span></span> 

![Снимок экрана: JIT-отладчик Visual Studio](../images/visual-studio-debugger.png)

<span data-ttu-id="6e827-p105">В **обозревателе решений** Visual Studio вы увидите файлы кода.   Вы можете задать точки останова для отлаживаемой строки кода в Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="6e827-p105">In Visual Studio, you will see the code files in **Solution Explorer**.   You can set breakpoints to the line of code you want to debug in Visual Studio.</span></span>

> [!NOTE]
> <span data-ttu-id="6e827-119">Если меню "Личные данные" не отображается, отладить надстройку можно с помощью Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="6e827-119">If you don't see the Personality menu, you can debug your add-in using Visual Studio.</span></span> <span data-ttu-id="6e827-120">Убедитесь, что надстройка области задач открыта в Office, и выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="6e827-120">Ensure your task pane add-in is open in Office, and then follow these steps:</span></span>

> 1. <span data-ttu-id="6e827-121">В Visual Studio выберите **ОТЛАДКА** > **Присоединиться к процессу**.</span><span class="sxs-lookup"><span data-stu-id="6e827-121">In Visual Studio, choose **DEBUG** > **Attach to Process**.</span></span>
> 2. <span data-ttu-id="6e827-122">В окне **Присоединение к процессу** выберите все доступные процессы Iexplore.exe, а затем нажмите кнопку **Присоединиться**.</span><span class="sxs-lookup"><span data-stu-id="6e827-122">In **Attach to Process**, choose all of the available Iexplore.exe processes, and then choose the **Attach** button.</span></span>

<span data-ttu-id="6e827-123">Дополнительные сведения об отладке в Visual Studio см. в следующих статьях:</span><span class="sxs-lookup"><span data-stu-id="6e827-123">For more information about debugging in Visual Studio, see the following:</span></span>

-   <span data-ttu-id="6e827-124">Дополнительные сведения о запуске и использовании Проводника DOM в Visual Studio приведены в совете № 4 в разделе [Советы и рекомендации](https://blogs.msdn.microsoft.com/officeapps/2013/04/16/building-great-looking-apps-for-office-using-the-new-project-templates/#tips_tricks) записи в блоге [Создание отличных приложений для Office с помощью новых шаблонов проекта](https://blogs.msdn.microsoft.com/officeapps/2013/04/16/building-great-looking-apps-for-office-using-the-new-project-templates).</span><span class="sxs-lookup"><span data-stu-id="6e827-124">To launch and use the DOM Explorer in Visual Studio, see Tip 4 in the [Tips and Tricks](https://blogs.msdn.microsoft.com/officeapps/2013/04/16/building-great-looking-apps-for-office-using-the-new-project-templates/#tips_tricks) section of the [Building great-looking apps for Office using the new project templates](https://blogs.msdn.microsoft.com/officeapps/2013/04/16/building-great-looking-apps-for-office-using-the-new-project-templates) blog post.</span></span>
-   <span data-ttu-id="6e827-125">Как задать точки останова, можно узнать в статье [Использование точек останова](/visualstudio/debugger/using-breakpoints?view=vs-2015).</span><span class="sxs-lookup"><span data-stu-id="6e827-125">To set breakpoints, see [Using Breakpoints](/visualstudio/debugger/using-breakpoints?view=vs-2015).</span></span>
-   <span data-ttu-id="6e827-126">Сведения об использовании F12 см. в статье [Использование средств разработчика F12](/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85)).</span><span class="sxs-lookup"><span data-stu-id="6e827-126">To use F12, see [Using the F12 developer tools](/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85)).</span></span>

## <a name="see-also"></a><span data-ttu-id="6e827-127">См. также</span><span class="sxs-lookup"><span data-stu-id="6e827-127">See also</span></span>

- [<span data-ttu-id="6e827-128">Создание и отладка надстроек Office в Visual Studio</span><span class="sxs-lookup"><span data-stu-id="6e827-128">Create and debug Office Add-ins in Visual Studio</span></span>](../develop/create-and-debug-office-add-ins-in-visual-studio.md)
- [<span data-ttu-id="6e827-129">Публикация надстройки Office</span><span class="sxs-lookup"><span data-stu-id="6e827-129">Publish your Office Add-in</span></span>](../publish/publish.md)
