---
ms.date: 05/16/2020
description: Протестируйте надстройку Office с помощью Internet Explorer 11.
title: Тестирование Internet Explorer 11
localization_priority: Normal
ms.openlocfilehash: 4ea2b4da153e2908f928086cd4997502c194e578
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611206"
---
# <a name="test-your-office-add-in-using-internet-explorer-11"></a><span data-ttu-id="7214b-103">Тестирование надстройки Office с помощью Internet Explorer 11</span><span class="sxs-lookup"><span data-stu-id="7214b-103">Test your Office Add-in using Internet Explorer 11</span></span>

<span data-ttu-id="7214b-104">В зависимости от спецификаций надстройки вы можете запланировать поддержку более ранних версий Windows и Office, которые требуют тестирования в Internet Explorer 11.</span><span class="sxs-lookup"><span data-stu-id="7214b-104">Depending on the specifications of your add-in, you may plan to support older versions of Windows and Office, which require testing on Internet Explorer 11.</span></span> <span data-ttu-id="7214b-105">Это часто требуется при отправке надстройки в AppSource.</span><span class="sxs-lookup"><span data-stu-id="7214b-105">This is often necessary as part of submitting your add-in to AppSource.</span></span> <span data-ttu-id="7214b-106">С помощью средства командной строки можно переключиться с более современных сред выполнения, используемых надстройками, в среду выполнения Internet Explorer 11 для этого тестирования.</span><span class="sxs-lookup"><span data-stu-id="7214b-106">You can use the following command line tooling to switch from more modern runtimes used by add-ins to the Internet Explorer 11 runtime for this testing.</span></span>

## <a name="pre-requisites"></a><span data-ttu-id="7214b-107">Необходимые условия</span><span class="sxs-lookup"><span data-stu-id="7214b-107">Pre-requisites</span></span>

- <span data-ttu-id="7214b-108">[Node.js](https://nodejs.org/) (последняя версия [LTS](https://nodejs.org/about/releases))</span><span class="sxs-lookup"><span data-stu-id="7214b-108">[Node.js](https://nodejs.org/) (the latest [LTS](https://nodejs.org/about/releases) version)</span></span>
- <span data-ttu-id="7214b-109">Редактор кода.</span><span class="sxs-lookup"><span data-stu-id="7214b-109">A code editor.</span></span> <span data-ttu-id="7214b-110">Мы рекомендуем [Visual Studio Code](https://code.visualstudio.com/)</span><span class="sxs-lookup"><span data-stu-id="7214b-110">We recommend [Visual Studio Code](https://code.visualstudio.com/)</span></span>
- [<span data-ttu-id="7214b-111">Участие в программе предварительной оценки Office</span><span class="sxs-lookup"><span data-stu-id="7214b-111">Be part of the Office Insider program</span></span>](https://insider.office.com)

<span data-ttu-id="7214b-112">В этих инструкциях предполагается, что ранее был настроен проект генератора Yo Office.</span><span class="sxs-lookup"><span data-stu-id="7214b-112">These instructions assume you have set up a Yo Office generator project before.</span></span> <span data-ttu-id="7214b-113">Если вы еще этого не сделали, рекомендуем ознакомиться со кратким руководством, например: [для надстроек Excel](../quickstarts/excel-quickstart-jquery.md).</span><span class="sxs-lookup"><span data-stu-id="7214b-113">If you haven't done this before, consider reading a quick start, such as [this one for Excel add-ins](../quickstarts/excel-quickstart-jquery.md).</span></span>

## <a name="using-ie11-tooling"></a><span data-ttu-id="7214b-114">Использование средства IE11</span><span class="sxs-lookup"><span data-stu-id="7214b-114">Using IE11 tooling</span></span>

1. <span data-ttu-id="7214b-115">Создайте проект генератора Yo Office.</span><span class="sxs-lookup"><span data-stu-id="7214b-115">Create a Yo Office generator project.</span></span> <span data-ttu-id="7214b-116">В этом случае не имеет значения, какой тип проекта будет выбран, это средство будет работать со всеми типами проектов.</span><span class="sxs-lookup"><span data-stu-id="7214b-116">It doesn't matter what kind of project you select, this tooling will work with all project types.</span></span>

> <span data-ttu-id="7214b-117">! НОТЕ Если у вас есть проект и вы хотите добавить этот инструмент без создания нового проекта, пропустите этот шаг и перейдите к следующему шагу.</span><span class="sxs-lookup"><span data-stu-id="7214b-117">![NOTE] If you have an existing project and want to add this tooling without creating a new project, skip this step and move to the next step.</span></span> 

2. <span data-ttu-id="7214b-118">В корневой папке нового проекта выполните в командной строке следующую команду:</span><span class="sxs-lookup"><span data-stu-id="7214b-118">In the root folder of your new project, run the following in the command line:</span></span>

```command&nbsp;line
office-add-dev-settings webview manifest.xml ie
```
<span data-ttu-id="7214b-119">В командной строке должно появиться примечание о том, что в качестве типа представления веб-сайта теперь задано значение IE.</span><span class="sxs-lookup"><span data-stu-id="7214b-119">You should see a note in the command line that the web view type is now set to IE.</span></span>

> <span data-ttu-id="7214b-120">! Последняя Это средство не обязательно использовать, но оно должно помочь отладить большинство проблем, связанных со средой выполнения Internet Explorer 11.</span><span class="sxs-lookup"><span data-stu-id="7214b-120">![TIP] It isn't necessary to use this tooling, but it should help debug the majority of issues related to the Internet Explorer 11 runtime.</span></span> <span data-ttu-id="7214b-121">Для полной надежности необходимо протестировать использование компьютера с установленной копией Windows 7 и Office 2013.</span><span class="sxs-lookup"><span data-stu-id="7214b-121">For complete robustness, you should test using a computer with a copy of Windows 7 and Office 2013 installed.</span></span>

## <a name="command-settings"></a><span data-ttu-id="7214b-122">Параметры команды</span><span class="sxs-lookup"><span data-stu-id="7214b-122">Command settings</span></span>

<span data-ttu-id="7214b-123">Если у вас есть другой путь манифеста, укажите его в команде, как показано в следующем примере:</span><span class="sxs-lookup"><span data-stu-id="7214b-123">Should you have a different manifest path, specify this in the command, as shown in the following:</span></span>

`office-add-dev-settings webview [path to your manifest] ie`

<span data-ttu-id="7214b-124">`office-addin-dev-settings webview`Кроме того, в качестве аргументов команды можно использовать ряд сред выполнения:</span><span class="sxs-lookup"><span data-stu-id="7214b-124">The `office-addin-dev-settings webview` command can also take a number of runtimes as arguments:</span></span>

- <span data-ttu-id="7214b-125">Explorer</span><span class="sxs-lookup"><span data-stu-id="7214b-125">ie</span></span>
- <span data-ttu-id="7214b-126">кромки</span><span class="sxs-lookup"><span data-stu-id="7214b-126">edge</span></span>
- <span data-ttu-id="7214b-127">Значение  по умолчанию</span><span class="sxs-lookup"><span data-stu-id="7214b-127">default</span></span>

## <a name="see-also"></a><span data-ttu-id="7214b-128">См. также</span><span class="sxs-lookup"><span data-stu-id="7214b-128">See also</span></span>
* [<span data-ttu-id="7214b-129">Тестирование и отладка надстроек Office</span><span class="sxs-lookup"><span data-stu-id="7214b-129">Test and debug Office Add-ins</span></span>](test-debug-office-add-ins.md)
* [<span data-ttu-id="7214b-130">Загрузка неопубликованных надстроек Office для тестирования</span><span class="sxs-lookup"><span data-stu-id="7214b-130">Sideload Office Add-ins for testing</span></span>](create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
* [<span data-ttu-id="7214b-131">Отладка надстроек с помощью средств разработчика в Windows 10</span><span class="sxs-lookup"><span data-stu-id="7214b-131">Debug add-ins using developer tools on Windows 10</span></span>](debug-add-ins-using-f12-developer-tools-on-windows-10.md)
* [<span data-ttu-id="7214b-132">Подключение отладчика из области задач</span><span class="sxs-lookup"><span data-stu-id="7214b-132">Attach a debugger from the task pane</span></span>](attach-debugger-from-task-pane.md)