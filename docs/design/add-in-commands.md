---
title: Основные концепции команд надстроек
description: Как добавить настраиваемые кнопки ленты и элементы меню в Office в составе надстройки Office
ms.date: 07/10/2020
localization_priority: Priority
ms.openlocfilehash: 13db2191d9691a699c5976b812e1ca6d8f3bf1ae
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293361"
---
# <a name="add-in-commands-for-excel-powerpoint-and-word"></a><span data-ttu-id="32a92-103">Команды надстроек для Excel, PowerPoint и Word</span><span class="sxs-lookup"><span data-stu-id="32a92-103">Add-in commands for Excel, PowerPoint, and Word</span></span>

<span data-ttu-id="32a92-p101">Команды надстроек — это элементы, которые расширяют пользовательский интерфейс Office и запускают действия в надстройке. Команды надстроек можно использовать для добавления кнопки на ленту или элемента в контекстное меню. Когда пользователи выбирают команду надстройки, они инициируют действия, такие как запуск кода JavaScript или отображение страницы надстройки в области задач. Команды надстройки помогают пользователям находить и использовать вашу надстройку, что может повысить показатель внедрения надстройки и коэффициент удержания клиентов.</span><span class="sxs-lookup"><span data-stu-id="32a92-p101">Add-in commands are UI elements that extend the Office UI and start actions in your add-in. You can use add-in commands to add a button on the ribbon or an item to a context menu. When users select an add-in command, they initiate actions such as running JavaScript code, or showing a page of the add-in in a task pane. Add-in commands help users find and use your add-in, which can help increase your add-in's adoption and reuse, and improve customer retention.</span></span>

<span data-ttu-id="32a92-108">Обзор этой функции приведен в видео, посвященном [командам надстроек на ленте приложения Office](https://channel9.msdn.com/events/Build/2016/P551).</span><span class="sxs-lookup"><span data-stu-id="32a92-108">For an overview of the feature, see the video [Add-in Commands in the Office app ribbon](https://channel9.msdn.com/events/Build/2016/P551).</span></span>

> [!NOTE]
> <span data-ttu-id="32a92-p102">В каталогах SharePoint не поддерживаются команды надстроек. Последние можно развернуть с помощью компонента [централизованного развертывания](../publish/centralized-deployment.md) или [AppSource](/office/dev/store/submit-to-appsource-via-partner-center). Чтобы развернуть команду надстройки для тестирования, выполните [загрузку неопубликованного приложения](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="32a92-p102">SharePoint catalogs do not support add-in commands. You can deploy add-in commands via [Centralized Deployment](../publish/centralized-deployment.md) or [AppSource](/office/dev/store/submit-to-appsource-via-partner-center), or use [sideloading](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) to deploy your add-in command for testing.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="32a92-111">В Outlook также поддерживаются команды надстроек.</span><span class="sxs-lookup"><span data-stu-id="32a92-111">Add-in commands are also supported in Outlook.</span></span> <span data-ttu-id="32a92-112">Дополнительные сведения см. в статье [Команды надстроек для Outlook](../outlook/add-in-commands-for-outlook.md).</span><span class="sxs-lookup"><span data-stu-id="32a92-112">For more information, see [Add-in commands for Outlook](../outlook/add-in-commands-for-outlook.md).</span></span>

<span data-ttu-id="32a92-113">*Рисунок 1. Надстройка с командами, работающая в классическом приложении Excel*</span><span class="sxs-lookup"><span data-stu-id="32a92-113">*Figure 1. Add-in with commands running in Excel Desktop*</span></span>

![Снимок экрана с командой надстройки в приложении Excel](../images/add-in-commands-1.png)

<span data-ttu-id="32a92-115">*Рисунок 2. Надстройка с командами, работающая в Excel в Интернете*</span><span class="sxs-lookup"><span data-stu-id="32a92-115">*Figure 2. Add-in with commands running in Excel on the web*</span></span>

![Снимок экрана с командой надстройки в Excel в Интернете](../images/add-in-commands-2.png)

## <a name="command-capabilities"></a><span data-ttu-id="32a92-117">Возможности команд</span><span class="sxs-lookup"><span data-stu-id="32a92-117">Command capabilities</span></span>

<span data-ttu-id="32a92-118">В настоящее время поддерживаются указанные ниже возможности команд.</span><span class="sxs-lookup"><span data-stu-id="32a92-118">The following command capabilities are currently supported.</span></span>

> [!NOTE]
> <span data-ttu-id="32a92-119">Контентные надстройки на данный момент не поддерживают команды.</span><span class="sxs-lookup"><span data-stu-id="32a92-119">Content add-ins do not currently support add-in commands.</span></span>

### <a name="extension-points"></a><span data-ttu-id="32a92-120">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="32a92-120">Extension points</span></span>

- <span data-ttu-id="32a92-121">Вкладки ленты: расширение возможностей встроенных вкладок или создание пользовательской вкладки.</span><span class="sxs-lookup"><span data-stu-id="32a92-121">Ribbon tabs - Extend built-in tabs or create a new custom tab.</span></span>
- <span data-ttu-id="32a92-122">Контекстные меню: расширение возможностей выбранных контекстных меню.</span><span class="sxs-lookup"><span data-stu-id="32a92-122">Context menus - Extend selected context menus.</span></span>

### <a name="control-types"></a><span data-ttu-id="32a92-123">Типы элементов управления</span><span class="sxs-lookup"><span data-stu-id="32a92-123">Control types</span></span>

- <span data-ttu-id="32a92-124">Простые кнопки, запускающие определенные действия.</span><span class="sxs-lookup"><span data-stu-id="32a92-124">Simple buttons - trigger specific actions.</span></span>
- <span data-ttu-id="32a92-125">Простые раскрывающиеся меню с кнопками, которые запускают действия.</span><span class="sxs-lookup"><span data-stu-id="32a92-125">Menus - simple menu dropdown with buttons that trigger actions.</span></span>

### <a name="actions"></a><span data-ttu-id="32a92-126">Действия</span><span class="sxs-lookup"><span data-stu-id="32a92-126">Actions</span></span>

- <span data-ttu-id="32a92-127">ShowTaskpane: отображает одну или несколько областей, в которые можно загрузить пользовательские HTML-страницы.</span><span class="sxs-lookup"><span data-stu-id="32a92-127">ShowTaskpane - Displays one or multiple panes that load custom HTML pages inside them.</span></span>
- <span data-ttu-id="32a92-p104">ExecuteFunction загружает невидимую HTML-страницу, а затем выполняет содержащуюся в ней функцию JavaScript. Для показа ошибок, хода выполнения или дополнительных данных функции можно использовать API [displayDialog](/javascript/api/office/office.ui).</span><span class="sxs-lookup"><span data-stu-id="32a92-p104">ExecuteFunction - Loads an invisible HTML page and then execute a JavaScript function within it. To show UI within your function (such as errors, progress, or additional input) you can use the [displayDialog](/javascript/api/office/office.ui) API.</span></span>  

### <a name="default-enabled-or-disabled-status-preview"></a><span data-ttu-id="32a92-130">Состояние по умолчанию: "Включено" или "Отключено" (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="32a92-130">Default Enabled or Disabled Status (preview)</span></span>

<span data-ttu-id="32a92-131">Вы можете указать, включена или отключена команда при запуске надстройки, а также изменять параметр программными средствами.</span><span class="sxs-lookup"><span data-stu-id="32a92-131">You can specify whether the command is enabled or disabled when your add-in launches, and programmatically change the setting.</span></span>

> [!NOTE]
> <span data-ttu-id="32a92-132">Эта функция доступна в предварительной версии и поддерживается не всеми приложениями Office и сценариями.</span><span class="sxs-lookup"><span data-stu-id="32a92-132">This feature is in preview and is not supported in all Office applications or scenarios.</span></span> <span data-ttu-id="32a92-133">Дополнительные сведения см. в статье о [Включение и отключение команд надстроек](disable-add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="32a92-133">For more information, see [Enable and Disable Add-in Commands](disable-add-in-commands.md).</span></span>

## <a name="supported-platforms"></a><span data-ttu-id="32a92-134">Поддерживаемые платформы</span><span class="sxs-lookup"><span data-stu-id="32a92-134">Supported platforms</span></span>

<span data-ttu-id="32a92-135">В настоящее время команды надстроек поддерживаются на перечисленных ниже платформах.</span><span class="sxs-lookup"><span data-stu-id="32a92-135">Add-in commands are currently supported on the following platforms.</span></span>

- <span data-ttu-id="32a92-136">Office для Windows (сборка 16.0.6769+, подключенная к подписке на Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="32a92-136">Office on Windows (build 16.0.6769+, connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="32a92-137">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="32a92-137">Office 2019 on Windows</span></span>
- <span data-ttu-id="32a92-138">Office для Mac (сборка 15.33+, подключенная к подписке на Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="32a92-138">Office on Mac (build 15.33+, connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="32a92-139">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="32a92-139">Office 2019 on Mac</span></span>
- <span data-ttu-id="32a92-140">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="32a92-140">Office on the web</span></span>

> [!NOTE]
> <span data-ttu-id="32a92-141">Сведения о поддержке Outlook см. в[Команды надстройки для Outlook](../outlook/add-in-commands-for-outlook.md).</span><span class="sxs-lookup"><span data-stu-id="32a92-141">For information about support in Outlook, see [Add-in commands for Outlook](../outlook/add-in-commands-for-outlook.md).</span></span>

## <a name="debugging"></a><span data-ttu-id="32a92-142">Отладка</span><span class="sxs-lookup"><span data-stu-id="32a92-142">Debugging</span></span>

<span data-ttu-id="32a92-143">Чтобы отлаживать команду надстройки, необходимо запустить ее в Office в Интернете.</span><span class="sxs-lookup"><span data-stu-id="32a92-143">To debug an Add-in Command, you must run it in Office on the web.</span></span> <span data-ttu-id="32a92-144">Дополнительные сведения см. в статье [Отладка надстроек в Office в Интернете](../testing/debug-add-ins-in-office-online.md)</span><span class="sxs-lookup"><span data-stu-id="32a92-144">For details, see [Debug add-ins in Office on the web](../testing/debug-add-ins-in-office-online.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="32a92-145">Рекомендации</span><span class="sxs-lookup"><span data-stu-id="32a92-145">Best practices</span></span>

<span data-ttu-id="32a92-146">При разработке надстроек придерживайтесь следующих рекомендаций:</span><span class="sxs-lookup"><span data-stu-id="32a92-146">Apply the following best practices when you develop add-in commands:</span></span>

- <span data-ttu-id="32a92-p107">Каждая команда должна представлять определенное действие с очевидным и конкретным исходом для пользователей. Не совмещайте несколько действий в одной кнопке.</span><span class="sxs-lookup"><span data-stu-id="32a92-p107">Use commands to represent a specific action with a clear and specific outcome for users. Do not combine multiple actions in a single button.</span></span>
- <span data-ttu-id="32a92-p108">Предоставляйте точные действия, которые делают выполнение распространенных задач в надстройке более эффективным. Максимально сократите количество шагов, необходимых для выполнения действия.</span><span class="sxs-lookup"><span data-stu-id="32a92-p108">Provide granular actions that make common tasks within your add-in more efficient to perform. Minimize the number of steps an action takes to complete.</span></span>
- <span data-ttu-id="32a92-151">Расположение команд на ленте приложения Office:</span><span class="sxs-lookup"><span data-stu-id="32a92-151">For the placement of your commands in the Office app ribbon:</span></span>
    - <span data-ttu-id="32a92-p109">Помещайте команды на имеющиеся вкладки ("Вставка", "Рецензирование" и т. д.), если соответствующая функция подходит для них. Например, если надстройка позволяет вставлять файлы мультимедиа, добавьте группу на вкладку "Вставка". Обратите внимание, что некоторые вкладки доступны не во всех версиях Office. Дополнительные сведения см. в статье [XML-манифест надстроек Office](../develop/add-in-manifests.md).</span><span class="sxs-lookup"><span data-stu-id="32a92-p109">Place commands on an existing tab (Insert, Review, and so on) if the functionality provided fits there. For example, if your add-in enables users to insert media, add a group to the Insert tab. Note that not all tabs are available across all Office versions. For more information, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span>
    - <span data-ttu-id="32a92-p110">Добавляйте команды на вкладку "Главная", если соответствующие функции не относятся к другим вкладкам, а надстройка содержит менее шести команд верхнего уровня. Вы также можете добавлять команды на вкладку "Главная", если надстройка должна работать в разных версиях Office (например, Office в Интернете и классических приложениях Office), а нужная вкладка доступна не во всех версиях (например, вкладка "Конструктор" отсутствует в Office в Интернете).</span><span class="sxs-lookup"><span data-stu-id="32a92-p110">Place commands on the Home tab if the functionality doesn't fit on another tab, and you have fewer than six top-level commands. You can also add commands to the Home tab if your add-in needs to work across Office versions (such as Office on the web or desktop) and a tab is not available in all versions (for example, the Design tab doesn't exist in Office on the web).</span></span>  
    - <span data-ttu-id="32a92-157">Добавляйте команды на пользовательскую вкладку, если надстройка содержит более шести команд верхнего уровня.</span><span class="sxs-lookup"><span data-stu-id="32a92-157">Place commands on a custom tab if you have more than six top-level commands.</span></span>
    - <span data-ttu-id="32a92-p111">Название группы должно соответствовать названию надстройки. Если у вас есть несколько групп, их имена должны быть связаны с функциями, которые выполняют команды из этих групп.</span><span class="sxs-lookup"><span data-stu-id="32a92-p111">Name your group to match the name of your add-in. If you have multiple groups, name each group based on the functionality that the commands in that group provide.</span></span>
    - <span data-ttu-id="32a92-160">Не добавляйте избыточные кнопки, чтобы надстройка занимала больше места на экране.</span><span class="sxs-lookup"><span data-stu-id="32a92-160">Do not add superfluous buttons to increase the real estate of your add-in.</span></span>

     > [!NOTE]
     > <span data-ttu-id="32a92-161">Надстройки, которые занимают слишком много места, могут не пройти [проверку в AppSource](/legal/marketplace/certification-policies).</span><span class="sxs-lookup"><span data-stu-id="32a92-161">Add-ins that take up too much space might not pass [AppSource validation](/legal/marketplace/certification-policies).</span></span>

- <span data-ttu-id="32a92-162">[Руководство по оформлению значков](add-in-icons.md) подходит для всех значков.</span><span class="sxs-lookup"><span data-stu-id="32a92-162">For all icons, follow the [icon design guidelines](add-in-icons.md).</span></span>
- <span data-ttu-id="32a92-163">Предоставьте версию надстройки, которая работает в приложениях Office, не поддерживающих команды.</span><span class="sxs-lookup"><span data-stu-id="32a92-163">Provide a version of your add-in that also works on Office applications that do not support commands.</span></span> <span data-ttu-id="32a92-164">Один манифест надстройки может работать в приложениях независимо от того, поддерживают ли они команды.</span><span class="sxs-lookup"><span data-stu-id="32a92-164">A single add-in manifest can work in both command-aware (with commands) and non-command-aware (as a task pane) applications.</span></span>

   <span data-ttu-id="32a92-165">*Рис. 3. Надстройка области задач в Office 2013 и эта же надстройка, использующая команды надстройки в Office 2016*</span><span class="sxs-lookup"><span data-stu-id="32a92-165">*Figure 3. Task pane add-in in Office 2013 and the same add-in using add-in commands in Office 2016*</span></span>

   ![Снимок экрана: надстройка области задач в Office 2013 и эта же надстройка, использующая команды надстройки в Office 2016](../images/office-task-pane-add-ins.png)


## <a name="next-steps"></a><span data-ttu-id="32a92-167">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="32a92-167">Next steps</span></span>

<span data-ttu-id="32a92-168">Лучший способ начать работу с командами надстроек Office — ознакомиться с [примерами](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/) на сайте GitHub.</span><span class="sxs-lookup"><span data-stu-id="32a92-168">The best way to get started using add-in commands is to take a look at the [Office Add-in commands samples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/) on GitHub.</span></span>

<span data-ttu-id="32a92-169">Дополнительные сведения об указании команд надстройки в манифесте см. в статье [Создание команд надстроек в манифесте](../develop/create-addin-commands.md) и справочных материалах по [VersionOverrides](../reference/manifest/versionoverrides.md).</span><span class="sxs-lookup"><span data-stu-id="32a92-169">For more information about specifying add-in commands in your manifest, see [Create add-in commands in your manifest](../develop/create-addin-commands.md) and the [VersionOverrides](../reference/manifest/versionoverrides.md) reference content.</span></span>
