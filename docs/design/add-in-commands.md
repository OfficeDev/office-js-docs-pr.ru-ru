---
title: Команды надстроек для Excel, Word и PowerPoint
description: ''
ms.date: 05/08/2019
localization_priority: Priority
ms.openlocfilehash: 7b85d3016b195b353b1e7f314aceb761cf4e31b3
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952182"
---
# <a name="add-in-commands-for-excel-word-and-powerpoint"></a><span data-ttu-id="ea4c4-102">Команды надстроек для Excel, Word и PowerPoint</span><span class="sxs-lookup"><span data-stu-id="ea4c4-102">Add-in commands for Excel, Word, and PowerPoint</span></span>

<span data-ttu-id="ea4c4-p101">Команды надстроек — это элементы, которые расширяют пользовательский интерфейс Office и запускают действия в надстройке. Команды надстроек можно использовать для добавления кнопки на ленту или элемента в контекстное меню. Когда пользователи выбирают команду надстройки, они инициируют действия, такие как запуск кода JavaScript или отображение страницы надстройки в области задач. Команды надстройки помогают пользователям находить и использовать вашу надстройку, что может повысить показатель внедрения надстройки и коэффициент удержания клиентов.</span><span class="sxs-lookup"><span data-stu-id="ea4c4-p101">Add-in commands are UI elements that extend the Office UI and start actions in your add-in. You can use add-in commands to add a button on the ribbon or an item to a context menu. When users select an add-in command, they initiate actions such as running JavaScript code, or showing a page of the add-in in a task pane. Add-in commands help users find and use your add-in, which can help increase your add-in's adoption and reuse, and improve customer retention.</span></span>

<span data-ttu-id="ea4c4-107">Обзор этой функции приведен в видео, посвященном [командам надстроек на ленте Office](https://channel9.msdn.com/events/Build/2016/P551).</span><span class="sxs-lookup"><span data-stu-id="ea4c4-107">For an overview of the feature, see the video [Add-in Commands in the Office Ribbon](https://channel9.msdn.com/events/Build/2016/P551).</span></span>

> [!NOTE]
> <span data-ttu-id="ea4c4-p102">В каталогах SharePoint не поддерживаются команды надстроек. Последние можно развернуть с помощью компонента [централизованного развертывания](../publish/centralized-deployment.md) или [AppSource](/office/dev/store/submit-to-the-office-store). Чтобы развернуть команду надстройки для тестирования, выполните [загрузку неопубликованного приложения](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="ea4c4-p102">SharePoint catalogs do not support add-in commands. You can deploy add-in commands via [Centralized Deployment](../publish/centralized-deployment.md) or [AppSource](/office/dev/store/submit-to-the-office-store), or use [sideloading](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) to deploy your add-in command for testing.</span></span> 

<span data-ttu-id="ea4c4-110">*Рисунок 1. Надстройка с командами, работающая в классическом приложении Excel*</span><span class="sxs-lookup"><span data-stu-id="ea4c4-110">*Figure 1. Add-in with commands running in Excel Desktop*</span></span>

![Снимок экрана с командой надстройки в приложении Excel](../images/add-in-commands-1.png)

<span data-ttu-id="ea4c4-112">*Рисунок 2. Надстройка с командами, работающая в Excel Online*</span><span class="sxs-lookup"><span data-stu-id="ea4c4-112">*Figure 2. Add-in with commands running in Excel Online*</span></span>

![Снимок экрана с командой надстройки в Excel Online](../images/add-in-commands-2.png)

## <a name="command-capabilities"></a><span data-ttu-id="ea4c4-114">Возможности команд</span><span class="sxs-lookup"><span data-stu-id="ea4c4-114">Command capabilities</span></span>

<span data-ttu-id="ea4c4-115">В настоящее время поддерживаются указанные ниже возможности команд.</span><span class="sxs-lookup"><span data-stu-id="ea4c4-115">The following command capabilities are currently supported.</span></span>

> [!NOTE]
> <span data-ttu-id="ea4c4-116">Контентные надстройки на данный момент не поддерживают команды.</span><span class="sxs-lookup"><span data-stu-id="ea4c4-116">Content add-ins do not currently support add-in commands.</span></span>

<span data-ttu-id="ea4c4-117">**Точки расширения**</span><span class="sxs-lookup"><span data-stu-id="ea4c4-117">**Extension points**</span></span>

- <span data-ttu-id="ea4c4-118">Вкладки ленты: расширение возможностей встроенных вкладок или создание пользовательской вкладки.</span><span class="sxs-lookup"><span data-stu-id="ea4c4-118">Ribbon tabs - Extend built-in tabs or create a new custom tab.</span></span>
- <span data-ttu-id="ea4c4-119">Контекстные меню: расширение возможностей выбранных контекстных меню.</span><span class="sxs-lookup"><span data-stu-id="ea4c4-119">Context menus - Extend selected context menus.</span></span>

<span data-ttu-id="ea4c4-120">**Типы элементов управления**</span><span class="sxs-lookup"><span data-stu-id="ea4c4-120">**Control types**</span></span>

- <span data-ttu-id="ea4c4-121">Простые кнопки, запускающие определенные действия.</span><span class="sxs-lookup"><span data-stu-id="ea4c4-121">Simple buttons - trigger specific actions.</span></span>
- <span data-ttu-id="ea4c4-122">Простые раскрывающиеся меню с кнопками, которые запускают действия.</span><span class="sxs-lookup"><span data-stu-id="ea4c4-122">Menus - simple menu dropdown with buttons that trigger actions.</span></span>

<span data-ttu-id="ea4c4-123">**Действия**</span><span class="sxs-lookup"><span data-stu-id="ea4c4-123">**Actions**</span></span>

- <span data-ttu-id="ea4c4-124">ShowTaskpane: отображает одну или несколько областей, в которые можно загрузить пользовательские HTML-страницы.</span><span class="sxs-lookup"><span data-stu-id="ea4c4-124">ShowTaskpane - Displays one or multiple panes that load custom HTML pages inside them.</span></span>
- <span data-ttu-id="ea4c4-p103">ExecuteFunction загружает невидимую HTML-страницу, а затем выполняет содержащуюся в ней функцию JavaScript. Для показа ошибок, хода выполнения или дополнительных данных функции можно использовать API [displayDialog](/javascript/api/office/office.ui).</span><span class="sxs-lookup"><span data-stu-id="ea4c4-p103">ExecuteFunction - Loads an invisible HTML page and then execute a JavaScript function within it. To show UI within your function (such as errors, progress, or additional input) you can use the [displayDialog](/javascript/api/office/office.ui) API.</span></span>  

## <a name="supported-platforms"></a><span data-ttu-id="ea4c4-127">Поддерживаемые платформы</span><span class="sxs-lookup"><span data-stu-id="ea4c4-127">Supported platforms</span></span>

<span data-ttu-id="ea4c4-128">В настоящее время команды надстроек поддерживаются на следующих платформах:</span><span class="sxs-lookup"><span data-stu-id="ea4c4-128">Add-in commands are currently supported on the following platforms:</span></span>

- <span data-ttu-id="ea4c4-129">Outlook 2016 для Windows (сборка 16.0.4678.1000 или более поздняя)</span><span class="sxs-lookup"><span data-stu-id="ea4c4-129">Outlook 2016 on Windows (build 16.0.4678.1000+)</span></span>
- <span data-ttu-id="ea4c4-130">Office для Windows, подключенный к Office 365 (сборка 16.0.6769 или более поздняя)</span><span class="sxs-lookup"><span data-stu-id="ea4c4-130">Office on Windows connected to Office 365 (build 16.0.6769+)</span></span>
- <span data-ttu-id="ea4c4-131">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="ea4c4-131">Office 2019 for Windows</span></span>
- <span data-ttu-id="ea4c4-132">Office для Mac, подключенный к Office 365 (сборка 15.33 или более поздняя)</span><span class="sxs-lookup"><span data-stu-id="ea4c4-132">Office for Mac connected to Office 365 (build 15.33+)</span></span>
- <span data-ttu-id="ea4c4-133">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="ea4c4-133">Office 2019 for Mac</span></span>
- <span data-ttu-id="ea4c4-134">Office Online</span><span class="sxs-lookup"><span data-stu-id="ea4c4-134">Office Online</span></span>

<span data-ttu-id="ea4c4-135">Скоро можно будет использовать другие платформы.</span><span class="sxs-lookup"><span data-stu-id="ea4c4-135">More platforms are coming soon.</span></span>

## <a name="debugging"></a><span data-ttu-id="ea4c4-136">Отладка</span><span class="sxs-lookup"><span data-stu-id="ea4c4-136">Debugging</span></span>

<span data-ttu-id="ea4c4-137">Чтобы выполнить отладку команды надстройки, необходимо запустить ее в Office Online.</span><span class="sxs-lookup"><span data-stu-id="ea4c4-137">To debug an Add-in Command, you must run it in Office Online.</span></span> <span data-ttu-id="ea4c4-138">Дополнительные сведения см. в статье [Отладка надстроек в Office Online](../testing/debug-add-ins-in-office-online.md)</span><span class="sxs-lookup"><span data-stu-id="ea4c4-138">For details, see [Debug add-ins in Office Online](../testing/debug-add-ins-in-office-online.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="ea4c4-139">Рекомендации</span><span class="sxs-lookup"><span data-stu-id="ea4c4-139">Best practices</span></span>

<span data-ttu-id="ea4c4-140">При разработке надстроек придерживайтесь следующих рекомендаций:</span><span class="sxs-lookup"><span data-stu-id="ea4c4-140">Apply the following best practices when you develop add-in commands:</span></span>

- <span data-ttu-id="ea4c4-p105">Каждая команда должна представлять определенное действие с очевидным и конкретным исходом для пользователей. Не совмещайте несколько действий в одной кнопке.</span><span class="sxs-lookup"><span data-stu-id="ea4c4-p105">Use commands to represent a specific action with a clear and specific outcome for users. Do not combine multiple actions in a single button.</span></span>
- <span data-ttu-id="ea4c4-p106">Предоставляйте точные действия, которые делают выполнение распространенных задач в надстройке более эффективным. Максимально сократите количество шагов, необходимых для выполнения действия.</span><span class="sxs-lookup"><span data-stu-id="ea4c4-p106">Provide granular actions that make common tasks within your add-in more efficient to perform. Minimize the number of steps an action takes to complete.</span></span>
- <span data-ttu-id="ea4c4-145">Расположение команд на ленте Office:</span><span class="sxs-lookup"><span data-stu-id="ea4c4-145">For the placement of your commands in the Office ribbon:</span></span>
    - <span data-ttu-id="ea4c4-p107">Помещайте команды на имеющиеся вкладки ("Вставка", "Рецензирование" и т. д.), если соответствующая функция подходит для них. Например, если надстройка позволяет вставлять файлы мультимедиа, добавьте группу на вкладку "Вставка". Обратите внимание, что некоторые вкладки доступны не во всех версиях Office. Дополнительные сведения см. в статье [XML-манифест надстроек Office](../develop/add-in-manifests.md).</span><span class="sxs-lookup"><span data-stu-id="ea4c4-p107">Place commands on an existing tab (Insert, Review, and so on) if the functionality provided fits there. For example, if your add-in enables users to insert media, add a group to the Insert tab. Note that not all tabs are available across all Office versions. For more information, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span>
    - <span data-ttu-id="ea4c4-p108">Добавляйте команды на вкладку "Главная", если соответствующие функции не относятся к другим вкладкам и надстройка содержит менее шести команд верхнего уровня. Вы также можете добавлять команды на вкладку "Главная", если надстройка должна работать в разных версиях Office (например, классических приложениях Office и Office Online), а нужная вкладка доступна не во всех версиях (например, вкладка "Конструктор" отсутствует в Office Online).</span><span class="sxs-lookup"><span data-stu-id="ea4c4-p108">Place commands on the Home tab if the functionality doesn't fit on another tab, and you have fewer than six top-level commands. You can also add commands to the Home tab if your add-in needs to work across Office versions (such as Office Desktop and Office Online) and a tab is not available in all versions (for example, the Design tab doesn't exist in Office Online).</span></span>  
    - <span data-ttu-id="ea4c4-151">Добавляйте команды на пользовательскую вкладку, если надстройка содержит более шести команд верхнего уровня.</span><span class="sxs-lookup"><span data-stu-id="ea4c4-151">Place commands on a custom tab if you have more than six top-level commands.</span></span>
    - <span data-ttu-id="ea4c4-p109">Название группы должно соответствовать названию надстройки. Если у вас есть несколько групп, их имена должны быть связаны с функциями, которые выполняют команды из этих групп.</span><span class="sxs-lookup"><span data-stu-id="ea4c4-p109">Name your group to match the name of your add-in. If you have multiple groups, name each group based on the functionality that the commands in that group provide.</span></span>
    - <span data-ttu-id="ea4c4-154">Не добавляйте избыточные кнопки, чтобы надстройка занимала больше места на экране.</span><span class="sxs-lookup"><span data-stu-id="ea4c4-154">Do not add superfluous buttons to increase the real estate of your add-in.</span></span>

     > [!NOTE]
     > <span data-ttu-id="ea4c4-155">Надстройки, которые занимают слишком много места, могут не пройти [проверку в AppSource](/office/dev/store/validation-policies).</span><span class="sxs-lookup"><span data-stu-id="ea4c4-155">Add-ins that take up too much space might not pass [AppSource validation](/office/dev/store/validation-policies).</span></span>

- <span data-ttu-id="ea4c4-156">[Руководство по оформлению значков](add-in-icons.md) подходит для всех значков.</span><span class="sxs-lookup"><span data-stu-id="ea4c4-156">For all icons, follow the [icon design guidelines](add-in-icons.md).</span></span>
- <span data-ttu-id="ea4c4-157">Предоставьте версию надстройки, которая работает в ведущих приложениях, не поддерживающих команды.</span><span class="sxs-lookup"><span data-stu-id="ea4c4-157">Provide a version of your add-in that also works on hosts that do not support commands.</span></span> <span data-ttu-id="ea4c4-158">Один манифест надстройки может работать в ведущих приложениях независимо от того, поддерживают ли они команды.</span><span class="sxs-lookup"><span data-stu-id="ea4c4-158">A single add-in manifest can work in both command-aware (with commands) and non-command-aware (as a task pane) hosts.</span></span>

   <span data-ttu-id="ea4c4-159">*Рис. 3. Надстройка области задач в Office 2013 и эта же надстройка, использующая команды надстройки в Office 2016*</span><span class="sxs-lookup"><span data-stu-id="ea4c4-159">*Figure 3. Task pane add-in in Office 2013 and the same add-in using add-in commands in Office 2016*</span></span>

   ![Снимок экрана: надстройка области задач в Office 2013 и эта же надстройка, использующая команды надстройки в Office 2016](../images/office-task-pane-add-ins.png)


## <a name="next-steps"></a><span data-ttu-id="ea4c4-161">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="ea4c4-161">Next steps</span></span>

<span data-ttu-id="ea4c4-162">Лучший способ начать работу с командами надстроек Office — ознакомиться с [примерами](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/) на сайте GitHub.</span><span class="sxs-lookup"><span data-stu-id="ea4c4-162">The best way to get started using add-in commands is to take a look at the [Office Add-in commands samples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/) on GitHub.</span></span>

<span data-ttu-id="ea4c4-163">Дополнительные сведения об указании команд надстройки в манифесте см. в статье [Создание команд надстроек в манифесте](../develop/create-addin-commands.md) и справочных материалах по [VersionOverrides](/office/dev/add-ins/reference/manifest/versionoverrides).</span><span class="sxs-lookup"><span data-stu-id="ea4c4-163">For more information about specifying add-in commands in your manifest, see [Create add-in commands in your manifest](../develop/create-addin-commands.md) and the [VersionOverrides](/office/dev/add-ins/reference/manifest/versionoverrides) reference content.</span></span>
