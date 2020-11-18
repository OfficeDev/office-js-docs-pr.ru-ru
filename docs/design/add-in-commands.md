---
title: Основные концепции команд надстроек
description: Как добавить настраиваемые кнопки ленты и элементы меню в Office в составе надстройки Office
ms.date: 11/01/2020
localization_priority: Priority
ms.openlocfilehash: 3d7d99f05e9b02712a4f416b891d3be38875525b
ms.sourcegitcommit: 3189c4bd62dbe5950b19f28ac2c1314b6d304dca
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/17/2020
ms.locfileid: "49087968"
---
# <a name="add-in-commands-for-excel-powerpoint-and-word"></a><span data-ttu-id="b183a-103">Команды надстроек для Excel, PowerPoint и Word</span><span class="sxs-lookup"><span data-stu-id="b183a-103">Add-in commands for Excel, PowerPoint, and Word</span></span>

<span data-ttu-id="b183a-p101">Команды надстроек — это элементы, которые расширяют пользовательский интерфейс Office и запускают действия в надстройке. Команды надстроек можно использовать для добавления кнопки на ленту или элемента в контекстное меню. Когда пользователи выбирают команду надстройки, они инициируют действия, такие как запуск кода JavaScript или отображение страницы надстройки в области задач. Команды надстройки помогают пользователям находить и использовать вашу надстройку, что может повысить показатель внедрения надстройки и коэффициент удержания клиентов.</span><span class="sxs-lookup"><span data-stu-id="b183a-p101">Add-in commands are UI elements that extend the Office UI and start actions in your add-in. You can use add-in commands to add a button on the ribbon or an item to a context menu. When users select an add-in command, they initiate actions such as running JavaScript code, or showing a page of the add-in in a task pane. Add-in commands help users find and use your add-in, which can help increase your add-in's adoption and reuse, and improve customer retention.</span></span>

<span data-ttu-id="b183a-108">Обзор этой функции приведен в видео, посвященном [командам надстроек на ленте приложения Office](https://channel9.msdn.com/events/Build/2016/P551).</span><span class="sxs-lookup"><span data-stu-id="b183a-108">For an overview of the feature, see the video [Add-in Commands in the Office app ribbon](https://channel9.msdn.com/events/Build/2016/P551).</span></span>

> [!NOTE]
> <span data-ttu-id="b183a-p102">В каталогах SharePoint не поддерживаются команды надстроек. Последние можно развернуть с помощью компонента [централизованного развертывания](../publish/centralized-deployment.md) или [AppSource](/office/dev/store/submit-to-appsource-via-partner-center). Чтобы развернуть команду надстройки для тестирования, выполните [загрузку неопубликованного приложения](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="b183a-p102">SharePoint catalogs do not support add-in commands. You can deploy add-in commands via [Centralized Deployment](../publish/centralized-deployment.md) or [AppSource](/office/dev/store/submit-to-appsource-via-partner-center), or use [sideloading](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) to deploy your add-in command for testing.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="b183a-111">В Outlook также поддерживаются команды надстроек.</span><span class="sxs-lookup"><span data-stu-id="b183a-111">Add-in commands are also supported in Outlook.</span></span> <span data-ttu-id="b183a-112">Дополнительные сведения см. в статье [Команды надстроек для Outlook](../outlook/add-in-commands-for-outlook.md).</span><span class="sxs-lookup"><span data-stu-id="b183a-112">For more information, see [Add-in commands for Outlook](../outlook/add-in-commands-for-outlook.md).</span></span>

<span data-ttu-id="b183a-113">*Рисунок 1. Надстройка с командами, работающая в классическом приложении Excel*</span><span class="sxs-lookup"><span data-stu-id="b183a-113">*Figure 1. Add-in with commands running in Excel Desktop*</span></span>

![Снимок экрана с командой надстройки в приложении Excel](../images/add-in-commands-1.png)

<span data-ttu-id="b183a-115">*Рисунок 2. Надстройка с командами, работающая в Excel в Интернете*</span><span class="sxs-lookup"><span data-stu-id="b183a-115">*Figure 2. Add-in with commands running in Excel on the web*</span></span>

![Снимок экрана с командой надстройки в Excel в Интернете](../images/add-in-commands-2.png)

## <a name="command-capabilities"></a><span data-ttu-id="b183a-117">Возможности команд</span><span class="sxs-lookup"><span data-stu-id="b183a-117">Command capabilities</span></span>

<span data-ttu-id="b183a-118">В настоящее время поддерживаются указанные ниже возможности команд.</span><span class="sxs-lookup"><span data-stu-id="b183a-118">The following command capabilities are currently supported.</span></span>

> [!NOTE]
> <span data-ttu-id="b183a-119">Контентные надстройки на данный момент не поддерживают команды.</span><span class="sxs-lookup"><span data-stu-id="b183a-119">Content add-ins do not currently support add-in commands.</span></span>

### <a name="extension-points"></a><span data-ttu-id="b183a-120">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="b183a-120">Extension points</span></span>

- <span data-ttu-id="b183a-121">Вкладки ленты: расширение возможностей встроенных вкладок или создание пользовательской вкладки.</span><span class="sxs-lookup"><span data-stu-id="b183a-121">Ribbon tabs - Extend built-in tabs or create a new custom tab.</span></span>
- <span data-ttu-id="b183a-122">Контекстные меню: расширение возможностей выбранных контекстных меню.</span><span class="sxs-lookup"><span data-stu-id="b183a-122">Context menus - Extend selected context menus.</span></span>

### <a name="control-types"></a><span data-ttu-id="b183a-123">Типы элементов управления</span><span class="sxs-lookup"><span data-stu-id="b183a-123">Control types</span></span>

- <span data-ttu-id="b183a-124">Простые кнопки, запускающие определенные действия.</span><span class="sxs-lookup"><span data-stu-id="b183a-124">Simple buttons - trigger specific actions.</span></span>
- <span data-ttu-id="b183a-125">Простые раскрывающиеся меню с кнопками, которые запускают действия.</span><span class="sxs-lookup"><span data-stu-id="b183a-125">Menus - simple menu dropdown with buttons that trigger actions.</span></span>

### <a name="actions"></a><span data-ttu-id="b183a-126">Действия</span><span class="sxs-lookup"><span data-stu-id="b183a-126">Actions</span></span>

- <span data-ttu-id="b183a-127">ShowTaskpane: отображает одну или несколько областей, в которые можно загрузить пользовательские HTML-страницы.</span><span class="sxs-lookup"><span data-stu-id="b183a-127">ShowTaskpane - Displays one or multiple panes that load custom HTML pages inside them.</span></span>
- <span data-ttu-id="b183a-p104">ExecuteFunction загружает невидимую HTML-страницу, а затем выполняет содержащуюся в ней функцию JavaScript. Для показа ошибок, хода выполнения или дополнительных данных функции можно использовать API [displayDialog](/javascript/api/office/office.ui).</span><span class="sxs-lookup"><span data-stu-id="b183a-p104">ExecuteFunction - Loads an invisible HTML page and then execute a JavaScript function within it. To show UI within your function (such as errors, progress, or additional input) you can use the [displayDialog](/javascript/api/office/office.ui) API.</span></span>  

### <a name="default-enabled-or-disabled-status"></a><span data-ttu-id="b183a-130">Состояние по умолчанию: "Включено" или "Отключено"</span><span class="sxs-lookup"><span data-stu-id="b183a-130">Default Enabled or Disabled Status</span></span>

<span data-ttu-id="b183a-131">Вы можете указать, включена или отключена команда при запуске надстройки, а также изменять параметр программными средствами.</span><span class="sxs-lookup"><span data-stu-id="b183a-131">You can specify whether the command is enabled or disabled when your add-in launches, and programmatically change the setting.</span></span>

> [!NOTE]
> <span data-ttu-id="b183a-132">Эта функция поддерживается не всеми приложениями Office и сценариями.</span><span class="sxs-lookup"><span data-stu-id="b183a-132">This feature is not supported in all Office applications or scenarios.</span></span> <span data-ttu-id="b183a-133">Дополнительные сведения см. в статье [Включение и отключение команд надстроек](disable-add-in-commands.md).</span><span class="sxs-lookup"><span data-stu-id="b183a-133">For more information, see [Enable and Disable Add-in Commands](disable-add-in-commands.md).</span></span>

### <a name="position-on-the-ribbon-preview"></a><span data-ttu-id="b183a-134">Расположение на ленте (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="b183a-134">Position on the ribbon (preview)</span></span>

<span data-ttu-id="b183a-135">Вы можете указать, где настраиваемая вкладка будет отображаться на ленте приложения Office, например "справа от вкладки «Главная»".</span><span class="sxs-lookup"><span data-stu-id="b183a-135">You can specify where a custom tab appears on the Office application's ribbon, such as "just to the right of the Home tab".</span></span>

> [!NOTE]
> <span data-ttu-id="b183a-136">Эта функция поддерживается не всеми приложениями Office и сценариями.</span><span class="sxs-lookup"><span data-stu-id="b183a-136">This feature is not supported in all Office applications or scenarios.</span></span> <span data-ttu-id="b183a-137">Дополнительные сведения см. в статье [Расположение настраиваемой вкладки на ленте](custom-tab-placement.md).</span><span class="sxs-lookup"><span data-stu-id="b183a-137">For more information, see [Position a custom tab on the ribbon](custom-tab-placement.md).</span></span>

### <a name="integration-of-built-in-office-buttons-preview"></a><span data-ttu-id="b183a-138">Интеграция встроенных кнопок Office (предварительная версия)</span><span class="sxs-lookup"><span data-stu-id="b183a-138">Integration of built-in Office buttons (preview)</span></span>

<span data-ttu-id="b183a-139">Вы можете вставлять встроенные кнопки ленты Office в свои группы настраиваемых команд и настраиваемые вкладки ленты.</span><span class="sxs-lookup"><span data-stu-id="b183a-139">You can insert the built-in Office ribbon buttons into your custom command groups and custom ribbon tabs.</span></span>

> [!NOTE]
> <span data-ttu-id="b183a-140">Эта функция поддерживается не всеми приложениями Office и сценариями.</span><span class="sxs-lookup"><span data-stu-id="b183a-140">This feature is not supported in all Office applications or scenarios.</span></span> <span data-ttu-id="b183a-141">Дополнительные сведения см. в статье [Интеграция встроенных кнопок Office в настраиваемые вкладки](built-in-button-integration.md).</span><span class="sxs-lookup"><span data-stu-id="b183a-141">For more information, see [Integrate built-in Office buttons into custom tabs](built-in-button-integration.md).</span></span>


## <a name="supported-platforms"></a><span data-ttu-id="b183a-142">Поддерживаемые платформы</span><span class="sxs-lookup"><span data-stu-id="b183a-142">Supported platforms</span></span>

<span data-ttu-id="b183a-143">В настоящее время команды надстроек поддерживаются на следующих платформах (за исключением ограничений, указанных в подразделах [Возможности команд](#command-capabilities) ранее).</span><span class="sxs-lookup"><span data-stu-id="b183a-143">Add-in commands are currently supported on the following platforms, except for limitations specified in the subsections of [Command capabilities](#command-capabilities) earlier.</span></span>

- <span data-ttu-id="b183a-144">Office для Windows (сборка 16.0.6769+, подключенная к подписке на Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="b183a-144">Office on Windows (build 16.0.6769+, connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="b183a-145">Office 2019 для Windows</span><span class="sxs-lookup"><span data-stu-id="b183a-145">Office 2019 on Windows</span></span>
- <span data-ttu-id="b183a-146">Office для Mac (сборка 15.33+, подключенная к подписке на Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="b183a-146">Office on Mac (build 15.33+, connected to a Microsoft 365 subscription)</span></span>
- <span data-ttu-id="b183a-147">Office 2019 для Mac</span><span class="sxs-lookup"><span data-stu-id="b183a-147">Office 2019 on Mac</span></span>
- <span data-ttu-id="b183a-148">Office в Интернете</span><span class="sxs-lookup"><span data-stu-id="b183a-148">Office on the web</span></span>

> [!NOTE]
> <span data-ttu-id="b183a-149">Сведения о поддержке Outlook см. в[Команды надстройки для Outlook](../outlook/add-in-commands-for-outlook.md).</span><span class="sxs-lookup"><span data-stu-id="b183a-149">For information about support in Outlook, see [Add-in commands for Outlook](../outlook/add-in-commands-for-outlook.md).</span></span>

## <a name="debugging"></a><span data-ttu-id="b183a-150">Отладка</span><span class="sxs-lookup"><span data-stu-id="b183a-150">Debugging</span></span>

<span data-ttu-id="b183a-151">Чтобы отлаживать команду надстройки, необходимо запустить ее в Office в Интернете.</span><span class="sxs-lookup"><span data-stu-id="b183a-151">To debug an Add-in Command, you must run it in Office on the web.</span></span> <span data-ttu-id="b183a-152">Дополнительные сведения см. в статье [Отладка надстроек в Office в Интернете](../testing/debug-add-ins-in-office-online.md)</span><span class="sxs-lookup"><span data-stu-id="b183a-152">For details, see [Debug add-ins in Office on the web](../testing/debug-add-ins-in-office-online.md).</span></span>

## <a name="best-practices"></a><span data-ttu-id="b183a-153">Рекомендации</span><span class="sxs-lookup"><span data-stu-id="b183a-153">Best practices</span></span>

<span data-ttu-id="b183a-154">При разработке надстроек придерживайтесь следующих рекомендаций:</span><span class="sxs-lookup"><span data-stu-id="b183a-154">Apply the following best practices when you develop add-in commands:</span></span>

- <span data-ttu-id="b183a-p109">Каждая команда должна представлять определенное действие с очевидным и конкретным исходом для пользователей. Не совмещайте несколько действий в одной кнопке.</span><span class="sxs-lookup"><span data-stu-id="b183a-p109">Use commands to represent a specific action with a clear and specific outcome for users. Do not combine multiple actions in a single button.</span></span>
- <span data-ttu-id="b183a-p110">Предоставляйте точные действия, которые делают выполнение распространенных задач в надстройке более эффективным. Максимально сократите количество шагов, необходимых для выполнения действия.</span><span class="sxs-lookup"><span data-stu-id="b183a-p110">Provide granular actions that make common tasks within your add-in more efficient to perform. Minimize the number of steps an action takes to complete.</span></span>
- <span data-ttu-id="b183a-159">Расположение команд на ленте приложения Office:</span><span class="sxs-lookup"><span data-stu-id="b183a-159">For the placement of your commands in the Office app ribbon:</span></span>
    - <span data-ttu-id="b183a-p111">Помещайте команды на имеющиеся вкладки ("Вставка", "Рецензирование" и т. д.), если соответствующая функция подходит для них. Например, если надстройка позволяет вставлять файлы мультимедиа, добавьте группу на вкладку "Вставка". Обратите внимание, что некоторые вкладки доступны не во всех версиях Office. Дополнительные сведения см. в статье [XML-манифест надстроек Office](../develop/add-in-manifests.md).</span><span class="sxs-lookup"><span data-stu-id="b183a-p111">Place commands on an existing tab (Insert, Review, and so on) if the functionality provided fits there. For example, if your add-in enables users to insert media, add a group to the Insert tab. Note that not all tabs are available across all Office versions. For more information, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span>
    - <span data-ttu-id="b183a-p112">Добавляйте команды на вкладку "Главная", если соответствующие функции не относятся к другим вкладкам, а надстройка содержит менее шести команд верхнего уровня. Вы также можете добавлять команды на вкладку "Главная", если надстройка должна работать в разных версиях Office (например, Office в Интернете и классических приложениях Office), а нужная вкладка доступна не во всех версиях (например, вкладка "Конструктор" отсутствует в Office в Интернете).</span><span class="sxs-lookup"><span data-stu-id="b183a-p112">Place commands on the Home tab if the functionality doesn't fit on another tab, and you have fewer than six top-level commands. You can also add commands to the Home tab if your add-in needs to work across Office versions (such as Office on the web or desktop) and a tab is not available in all versions (for example, the Design tab doesn't exist in Office on the web).</span></span>  
    - <span data-ttu-id="b183a-165">Добавляйте команды на пользовательскую вкладку, если надстройка содержит более шести команд верхнего уровня.</span><span class="sxs-lookup"><span data-stu-id="b183a-165">Place commands on a custom tab if you have more than six top-level commands.</span></span>
    - <span data-ttu-id="b183a-p113">Название группы должно соответствовать названию надстройки. Если у вас есть несколько групп, их имена должны быть связаны с функциями, которые выполняют команды из этих групп.</span><span class="sxs-lookup"><span data-stu-id="b183a-p113">Name your group to match the name of your add-in. If you have multiple groups, name each group based on the functionality that the commands in that group provide.</span></span>
    - <span data-ttu-id="b183a-168">Не добавляйте избыточные кнопки, чтобы надстройка занимала больше места на экране.</span><span class="sxs-lookup"><span data-stu-id="b183a-168">Do not add superfluous buttons to increase the real estate of your add-in.</span></span>
    - <span data-ttu-id="b183a-169">Не размещайте настраиваемую вкладку слева от вкладки "Главная" или переводите на нее фокус по умолчанию при открытии документа, если ваша надстройка не является основным способом взаимодействия с документом.</span><span class="sxs-lookup"><span data-stu-id="b183a-169">Do not position a custom tab to the left of the Home tab, or give it focus by default when the document opens, unless your add-in is the primary way users will interact with the document.</span></span> <span data-ttu-id="b183a-170">Чрезмерное выделение вашей надстройки создает неудобства и раздражает пользователей и администраторов.</span><span class="sxs-lookup"><span data-stu-id="b183a-170">Giving excessive prominence to your add-in inconveniences and annoys users and administrators.</span></span>
    - <span data-ttu-id="b183a-171">Если надстройка является основным способом взаимодействия пользователей с документом и у вас есть настраиваемая вкладка ленты, рассмотрите возможность интеграции кнопок во вкладку для применения функций Office, которые часто требуются пользователям.</span><span class="sxs-lookup"><span data-stu-id="b183a-171">If your add-in is the primary way users interact with the document and you have a custom ribbon tab, consider integrating into the tab the buttons for the Office functions that users will frequently need.</span></span>

     > [!NOTE]
     > <span data-ttu-id="b183a-172">Надстройки, которые занимают слишком много места, могут не пройти [проверку в AppSource](/legal/marketplace/certification-policies).</span><span class="sxs-lookup"><span data-stu-id="b183a-172">Add-ins that take up too much space might not pass [AppSource validation](/legal/marketplace/certification-policies).</span></span>

- <span data-ttu-id="b183a-173">[Руководство по оформлению значков](add-in-icons.md) подходит для всех значков.</span><span class="sxs-lookup"><span data-stu-id="b183a-173">For all icons, follow the [icon design guidelines](add-in-icons.md).</span></span>
- <span data-ttu-id="b183a-174">Предоставьте версию надстройки, которая работает в приложениях Office, не поддерживающих команды.</span><span class="sxs-lookup"><span data-stu-id="b183a-174">Provide a version of your add-in that also works on Office applications that do not support commands.</span></span> <span data-ttu-id="b183a-175">Один манифест надстройки может работать в приложениях независимо от того, поддерживают ли они команды.</span><span class="sxs-lookup"><span data-stu-id="b183a-175">A single add-in manifest can work in both command-aware (with commands) and non-command-aware (as a task pane) applications.</span></span>

   <span data-ttu-id="b183a-176">*Рис. 3. Надстройка области задач в Office 2013 и эта же надстройка, использующая команды надстройки в Office 2016*</span><span class="sxs-lookup"><span data-stu-id="b183a-176">*Figure 3. Task pane add-in in Office 2013 and the same add-in using add-in commands in Office 2016*</span></span>

   ![Снимок экрана: надстройка области задач в Office 2013 и эта же надстройка, использующая команды надстройки в Office 2016](../images/office-task-pane-add-ins.png)


## <a name="next-steps"></a><span data-ttu-id="b183a-178">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="b183a-178">Next steps</span></span>

<span data-ttu-id="b183a-179">Лучший способ начать работу с командами надстроек Office — ознакомиться с [примерами](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/) на сайте GitHub.</span><span class="sxs-lookup"><span data-stu-id="b183a-179">The best way to get started using add-in commands is to take a look at the [Office Add-in commands samples](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/) on GitHub.</span></span>

<span data-ttu-id="b183a-180">Дополнительные сведения об указании команд надстройки в манифесте см. в статье [Создание команд надстроек в манифесте](../develop/create-addin-commands.md) и справочных материалах по [VersionOverrides](../reference/manifest/versionoverrides.md).</span><span class="sxs-lookup"><span data-stu-id="b183a-180">For more information about specifying add-in commands in your manifest, see [Create add-in commands in your manifest](../develop/create-addin-commands.md) and the [VersionOverrides](../reference/manifest/versionoverrides.md) reference content.</span></span>
