---
title: Тестирование Internet Explorer 11
description: Проверьте Office надстройки в Internet Explorer 11.
ms.date: 05/19/2021
localization_priority: Normal
ms.openlocfilehash: de256ee8b0633f18d3188c5bbfae52cb24ff2c35
ms.sourcegitcommit: 0d3bf72f8ddd1b287bf95f832b7ecb9d9fa62a24
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/02/2021
ms.locfileid: "52727936"
---
# <a name="test-your-office-add-in-on-internet-explorer-11"></a><span data-ttu-id="6b6ed-103">Проверьте Office надстройки в Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="6b6ed-103">Test your Office Add-in on Internet Explorer 11</span></span>

<span data-ttu-id="6b6ed-104">Если вы планируете выставлять надстройку на рынок через AppSource или планируете поддерживать более старые версии Windows и Office, надстройка должна работать в встраиваемом контроле браузера, основанном на Internet Explorer 11 (IE11).</span><span class="sxs-lookup"><span data-stu-id="6b6ed-104">If you plan to market your add-in through AppSource or you plan to support older versions of Windows and Office, your add-in must work in the embeddable browser control that is based on Internet Explorer 11 (IE11).</span></span> <span data-ttu-id="6b6ed-105">Вы можете использовать командную строку для перехода от более современных времен работы, используемых надстройки, к времени запуска Internet Explorer 11 для этого тестирования.</span><span class="sxs-lookup"><span data-stu-id="6b6ed-105">You can use a command line to switch from more modern runtimes used by add-ins to the Internet Explorer 11 runtime for this testing.</span></span> <span data-ttu-id="6b6ed-106">Сведения о том, какие версии Windows и Office используют управление веб-представлением Internet Explorer 11, см. в браузерах, используемых Office [надстройки.](../concepts/browsers-used-by-office-web-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="6b6ed-106">For information about which versions of Windows and Office use the Internet Explorer 11 web view control, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).</span></span>

> [!IMPORTANT]
> <span data-ttu-id="6b6ed-107">Internet Explorer 11 не поддерживает версии JavaScript, более поздние, чем ES5.</span><span class="sxs-lookup"><span data-stu-id="6b6ed-107">Internet Explorer 11 does not support JavaScript versions later than ES5.</span></span> <span data-ttu-id="6b6ed-108">Если вы хотите использовать синтаксис и функции ECMAScript 2015 или более поздней части, у вас есть два варианта:</span><span class="sxs-lookup"><span data-stu-id="6b6ed-108">If you want to use the syntax and features of ECMAScript 2015 or later, you have two options:</span></span>
>
> - <span data-ttu-id="6b6ed-109">Напишите код в ECMAScript 2015 (также называемый ES6) или позже JavaScript, или в TypeScript, а затем скомпилировать код в ES5 JavaScript с помощью компиляторов, таких как [babel](https://babeljs.io/) или [tsc](https://www.typescriptlang.org/index.html).</span><span class="sxs-lookup"><span data-stu-id="6b6ed-109">Write your code in ECMAScript 2015 (also called ES6) or later JavaScript, or in TypeScript, and then compile your code to ES5 JavaScript using a compiler such as [babel](https://babeljs.io/) or [tsc](https://www.typescriptlang.org/index.html).</span></span>
> - <span data-ttu-id="6b6ed-110">Напишите в ECMAScript 2015 или более [](https://en.wikipedia.org/wiki/Polyfill_(programming)) поздний JavaScript, а также загрузите библиотеку полифильмов, например [core-js,](https://github.com/zloirock/core-js) которая позволяет IE запускать код.</span><span class="sxs-lookup"><span data-stu-id="6b6ed-110">Write in ECMAScript 2015 or later JavaScript, but also load a [polyfill](https://en.wikipedia.org/wiki/Polyfill_(programming)) library such as [core-js](https://github.com/zloirock/core-js) that enables IE to run your code.</span></span>
>
> <span data-ttu-id="6b6ed-111">Дополнительные сведения об этих параметрах см. в [меню Support Internet Explorer 11.](../develop/support-ie-11.md)</span><span class="sxs-lookup"><span data-stu-id="6b6ed-111">For more information about these options, see [Support Internet Explorer 11](../develop/support-ie-11.md).</span></span>
>
> <span data-ttu-id="6b6ed-112">Кроме того, Internet Explorer 11 не поддерживает некоторые элементы HTML5, в частности медиа, запись и местоположение.</span><span class="sxs-lookup"><span data-stu-id="6b6ed-112">Also, Internet Explorer 11 does not support some HTML5 features such as media, recording, and location.</span></span>

> [!NOTE]
> <span data-ttu-id="6b6ed-113">Чтобы протестировать надстройку в браузере Internet Explorer 11, откройте Office в Интернете в Internet Explorer и разгрузите [надстройку.](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="6b6ed-113">To test your add-in on the Internet Explorer 11 browser, open Office on the web in Internet Explorer and [sideload the add-in](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="6b6ed-114">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="6b6ed-114">Prerequisites</span></span>

- <span data-ttu-id="6b6ed-115">[Node.js](https://nodejs.org/) (последняя версия [LTS](https://nodejs.org/about/releases))</span><span class="sxs-lookup"><span data-stu-id="6b6ed-115">[Node.js](https://nodejs.org/) (the latest [LTS](https://nodejs.org/about/releases) version)</span></span>

<span data-ttu-id="6b6ed-116">Эти инструкции предполагают, что вы создали проект генератора Yo Office ранее.</span><span class="sxs-lookup"><span data-stu-id="6b6ed-116">These instructions assume you have set up a Yo Office generator project before.</span></span> <span data-ttu-id="6b6ed-117">Если вы еще не сделали этого раньше, рассмотрите возможность быстрого начала чтения, например для Excel [надстройки.](../quickstarts/excel-quickstart-jquery.md)</span><span class="sxs-lookup"><span data-stu-id="6b6ed-117">If you haven't done this before, consider reading a quick start, such as [this one for Excel add-ins](../quickstarts/excel-quickstart-jquery.md).</span></span>

## <a name="switching-to-the-internet-explorer-11-webview"></a><span data-ttu-id="6b6ed-118">Переход на веб-просмотр Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="6b6ed-118">Switching to the Internet Explorer 11 webview</span></span>

1. <span data-ttu-id="6b6ed-119">Создайте проект Office Yo.</span><span class="sxs-lookup"><span data-stu-id="6b6ed-119">Create a Yo Office generator project.</span></span> <span data-ttu-id="6b6ed-120">Неважно, какой проект вы выберете, этот инструментарий будет работать со всеми типами проектов.</span><span class="sxs-lookup"><span data-stu-id="6b6ed-120">It doesn't matter what kind of project you select, this tooling will work with all project types.</span></span>

    > [!NOTE]
    > <span data-ttu-id="6b6ed-121">Если у вас есть существующий проект и вы хотите добавить этот инструмент без создания нового проекта, пропустите этот шаг и перейдйте к следующему шагу.</span><span class="sxs-lookup"><span data-stu-id="6b6ed-121">If you have an existing project and want to add this tooling without creating a new project, skip this step and move to the next step.</span></span> 

1. <span data-ttu-id="6b6ed-122">В корневой папке проекта запустите следующую строку в командной строке.</span><span class="sxs-lookup"><span data-stu-id="6b6ed-122">In the root folder of your project, run the following in the command line.</span></span> <span data-ttu-id="6b6ed-123">В этом примере предполагается, что файл манифеста проекта находится в корне.</span><span class="sxs-lookup"><span data-stu-id="6b6ed-123">This example assumes that your project's manifest file is in the root.</span></span> <span data-ttu-id="6b6ed-124">Если это не так, укажите относительный путь к файлу манифеста.</span><span class="sxs-lookup"><span data-stu-id="6b6ed-124">If it isn't, specify the relative path to the manifest file.</span></span> <span data-ttu-id="6b6ed-125">В командной строке должно быть видно сообщение о том, что тип веб-представления теперь настроен на IE.</span><span class="sxs-lookup"><span data-stu-id="6b6ed-125">You should see a message in the command line that the web view type is now set to IE.</span></span>

    ```command&nbsp;line
    npx office-addin-dev-settings webview manifest.xml ie
    ```

> [!TIP]
> <span data-ttu-id="6b6ed-126">Эта команда не требуется, но она должна помочь отламеть большинство проблем, связанных с запуском Internet Explorer 11.</span><span class="sxs-lookup"><span data-stu-id="6b6ed-126">It isn't necessary to use this command, but it should help debug the majority of issues related to the Internet Explorer 11 runtime.</span></span> <span data-ttu-id="6b6ed-127">Для полной надежности необходимо проверить использование компьютеров с различными комбинациями Windows 7, 8.1 и 10 и различных Office.</span><span class="sxs-lookup"><span data-stu-id="6b6ed-127">For complete robustness, you should test using computers with various combinations of Windows 7, 8.1, and 10 and various versions of Office.</span></span> <span data-ttu-id="6b6ed-128">Дополнительные сведения [](../concepts/browsers-used-by-office-web-add-ins.md) см. в Office надстройки и сведения о том, как вернуться к более ранней версии [Office.](https://support.microsoft.com/topic/how-to-revert-to-an-earlier-version-of-office-2bd5c457-a917-d57e-35a1-f709e3dda841)</span><span class="sxs-lookup"><span data-stu-id="6b6ed-128">For more information, see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md) and [How to revert to an earlier version of Office](https://support.microsoft.com/topic/how-to-revert-to-an-earlier-version-of-office-2bd5c457-a917-d57e-35a1-f709e3dda841).</span></span>

### <a name="command-options"></a><span data-ttu-id="6b6ed-129">Параметры команды</span><span class="sxs-lookup"><span data-stu-id="6b6ed-129">Command options</span></span>

<span data-ttu-id="6b6ed-130">В качестве аргументов команда может также использовать несколько времен `office-addin-dev-settings webview` работы:</span><span class="sxs-lookup"><span data-stu-id="6b6ed-130">The `office-addin-dev-settings webview` command can also take a number of runtimes as arguments:</span></span>

- <span data-ttu-id="6b6ed-131">ie</span><span class="sxs-lookup"><span data-stu-id="6b6ed-131">ie</span></span>
- <span data-ttu-id="6b6ed-132">edge</span><span class="sxs-lookup"><span data-stu-id="6b6ed-132">edge</span></span>
- <span data-ttu-id="6b6ed-133">default</span><span class="sxs-lookup"><span data-stu-id="6b6ed-133">default</span></span>

## <a name="see-also"></a><span data-ttu-id="6b6ed-134">См. также</span><span class="sxs-lookup"><span data-stu-id="6b6ed-134">See also</span></span>

* [<span data-ttu-id="6b6ed-135">Тестирование и отладка надстроек Office</span><span class="sxs-lookup"><span data-stu-id="6b6ed-135">Test and debug Office Add-ins</span></span>](test-debug-office-add-ins.md)
* [<span data-ttu-id="6b6ed-136">Загрузка неопубликованных надстроек Office для тестирования</span><span class="sxs-lookup"><span data-stu-id="6b6ed-136">Sideload Office Add-ins for testing</span></span>](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
* [<span data-ttu-id="6b6ed-137">Отладка надстроек с помощью средств разработчика в Windows 10</span><span class="sxs-lookup"><span data-stu-id="6b6ed-137">Debug add-ins using developer tools on Windows 10</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
* [<span data-ttu-id="6b6ed-138">Подключение отладчика из области задач</span><span class="sxs-lookup"><span data-stu-id="6b6ed-138">Attach a debugger from the task pane</span></span>](attach-debugger-from-task-pane.md)
