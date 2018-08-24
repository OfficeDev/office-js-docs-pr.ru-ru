---
title: Подключение отладчика из области задач
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 28f7741a6d04f8f1492fec45649cb55a9447bdd7
ms.sourcegitcommit: 4de2a1b62ccaa8e51982e95537fc9f52c0c5e687
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/10/2018
ms.locfileid: "22925152"
---
# <a name="attach-a-debugger-from-the-task-pane"></a><span data-ttu-id="5b45d-102">Подключение отладчика из области задач</span><span class="sxs-lookup"><span data-stu-id="5b45d-102">Attach a debugger from the task pane</span></span>

<span data-ttu-id="5b45d-p101">В Office 2016 для Windows (сборка 77xx.xxxx или более поздней версии) можно подключать отладчик из области задач. Функция "Подключить отладчик" подключит отладчик непосредственно к нужному процессу Internet Explorer. Вы можете подключить отладчик независимо от того, какой инструмент используете: генератор Yeoman, Visual Studio Code, node.js, Angular или другой.</span><span class="sxs-lookup"><span data-stu-id="5b45d-p101">In Office 2016 for Windows, Build 77xx.xxxx or later, you can attach the debugger from the task pane. The attach debugger feature will directly attach the debugger to the correct Internet Explorer process for you. You can attach a debugger regardless of whether you are using Yeoman Generator, Visual Studio Code, node.js, Angular, or another tool.</span></span> 

<span data-ttu-id="5b45d-106">Для запуска средства **подключения отладчика** откройте меню **Личные данные** в правом верхнем углу области задач (выделено красным на рисунке ниже).</span><span class="sxs-lookup"><span data-stu-id="5b45d-106">To launch the **Attach Debugger** tool, choose the top right corner of the task pane to activate the **Personality** menu (as shown in the red circle in the following image).</span></span>   

> [!NOTE]
> - <span data-ttu-id="5b45d-p102">В настоящее время поддерживается только отладчик [Visual Studio 2015](https://www.visualstudio.com/downloads/) с [обновлением 3](https://msdn.microsoft.com/library/mt752379.aspx) или более поздней версии. Если у вас нет Visual Studio, выбор параметра **Подключить отладчик** не даст результата.</span><span class="sxs-lookup"><span data-stu-id="5b45d-p102">Currently the only supported debugger tool is [Visual Studio 2015](https://www.visualstudio.com/downloads/) with [Update 3](https://msdn.microsoft.com/library/mt752379.aspx) or later. If you don't have Visual Studio installed, selecting the **Attach Debugger** option doesn’t result in any action.</span></span>   
> - <span data-ttu-id="5b45d-109">Для отладки клиентского кода JavaScript можно использовать только средство **Подключить отладчик**.</span><span class="sxs-lookup"><span data-stu-id="5b45d-109">You can only debug client-side JavaScript with the **Attach Debugger** tool.</span></span> <span data-ttu-id="5b45d-110">Для отладки серверного кода, например на сервере Node.js, существует множество вариантов.</span><span class="sxs-lookup"><span data-stu-id="5b45d-110">To debug server-side code, such as with a Node.js server, you have many options.</span></span> <span data-ttu-id="5b45d-111">Сведения о том, как выполнять отладку в Visual Studio Code, см. в статье [Отладка Node.js в VS Code](https://code.visualstudio.com/docs/nodejs/nodejs-debugging).</span><span class="sxs-lookup"><span data-stu-id="5b45d-111">For information on how to debug with Visual Studio Code, see [Node.js Debugging in VS Code](https://code.visualstudio.com/docs/nodejs/nodejs-debugging).</span></span> <span data-ttu-id="5b45d-112">Если вы не используете Visual Studio Code, выполните поиск по запросу "отладка Node.js" или "отладка {имя_сервера}".</span><span class="sxs-lookup"><span data-stu-id="5b45d-112">If you are not using Visual Studio Code, search for "debug Node.js" or "debug {name-of-server}".</span></span>

![Снимок экрана: меню подключения отладчика](../images/attach-debugger.png)

<span data-ttu-id="5b45d-p104">Выберите элемент **Подключить отладчик**. Откроется диалоговое окно **JIT-отладчик Visual Studio** (см. рисунок ниже).</span><span class="sxs-lookup"><span data-stu-id="5b45d-p104">Select **Attach Debugger**. This launches the **Visual Studio Just-in-Time Debugger** dialog box, as shown in the following image.</span></span> 

![Снимок экрана: JIT-отладчик Visual Studio](../images/visual-studio-debugger.png)

<span data-ttu-id="5b45d-117">В **обозревателе решений** Visual Studio вы увидите файлы кода.</span><span class="sxs-lookup"><span data-stu-id="5b45d-117">In Visual Studio, you will see the code files in **Solution Explorer**.</span></span>   <span data-ttu-id="5b45d-118">Вы можете задать точки останова для отлаживаемой строки кода в Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="5b45d-118">You can set breakpoints to the line of code you want to debug in Visual Studio.</span></span>

<span data-ttu-id="5b45d-119">Дополнительные сведения об отладке в Visual Studio см. в следующих статьях:</span><span class="sxs-lookup"><span data-stu-id="5b45d-119">For more information about debugging in Visual Studio, see the following:</span></span>

-   <span data-ttu-id="5b45d-120">Дополнительные сведения о запуске и использовании Проводника DOM в Visual Studio приведены в совете № 4 в разделе [Советы и рекомендации](https://blogs.msdn.microsoft.com/officeapps/2013/04/16/building-great-looking-apps-for-office-using-the-new-project-templates/#tips_tricks) записи в блоге [Создание отличных приложений для Office с помощью новых шаблонов проекта](https://blogs.msdn.microsoft.com/officeapps/2013/04/16/building-great-looking-apps-for-office-using-the-new-project-templates).</span><span class="sxs-lookup"><span data-stu-id="5b45d-120">To launch and use the DOM Explorer in Visual Studio, see Tip 4 in the [Tips and Tricks](https://blogs.msdn.microsoft.com/officeapps/2013/04/16/building-great-looking-apps-for-office-using-the-new-project-templates/#tips_tricks) section of the [Building great-looking apps for Office using the new project templates](https://blogs.msdn.microsoft.com/officeapps/2013/04/16/building-great-looking-apps-for-office-using-the-new-project-templates) blog post.</span></span>
-   <span data-ttu-id="5b45d-121">Как задать точки останова, можно узнать в статье [Использование точек останова](https://msdn.microsoft.com/library/5557y8b4.aspx).</span><span class="sxs-lookup"><span data-stu-id="5b45d-121">To set breakpoints, see [Using Breakpoints](https://msdn.microsoft.com/library/5557y8b4.aspx).</span></span>
-   <span data-ttu-id="5b45d-122">Сведения об использовании F12 см. в статье [Использование средств разработчика F12](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85)).</span><span class="sxs-lookup"><span data-stu-id="5b45d-122">To use F12, see [Using the F12 developer tools](https://docs.microsoft.com/previous-versions/windows/internet-explorer/ie-developer/samples/bg182326(v=vs.85)).</span></span>

## <a name="see-also"></a><span data-ttu-id="5b45d-123">См. также</span><span class="sxs-lookup"><span data-stu-id="5b45d-123">See also</span></span>

- [<span data-ttu-id="5b45d-124">Создание и отладка надстроек Office в Visual Studio</span><span class="sxs-lookup"><span data-stu-id="5b45d-124">Create and debug Office Add-ins in Visual Studio</span></span>](../develop/create-and-debug-office-add-ins-in-visual-studio.md)
- [<span data-ttu-id="5b45d-125">Публикация надстройки Office</span><span class="sxs-lookup"><span data-stu-id="5b45d-125">Publish your Office Add-in</span></span>](../publish/publish.md)
