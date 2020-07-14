---
title: Основные концепции команд надстроек
description: Как добавить настраиваемые кнопки ленты и элементы меню в Office в составе надстройки Office
ms.date: 05/12/2020
localization_priority: Priority
ms.openlocfilehash: 2fe14a41c93b53164ab0fa3a7d25f5b9810b9c6a
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093877"
---
# <a name="add-in-commands-for-excel-powerpoint-and-word"></a>Команды надстроек для Excel, PowerPoint и Word

Add-in commands are UI elements that extend the Office UI and start actions in your add-in. You can use add-in commands to add a button on the ribbon or an item to a context menu. When users select an add-in command, they initiate actions such as running JavaScript code, or showing a page of the add-in in a task pane. Add-in commands help users find and use your add-in, which can help increase your add-in's adoption and reuse, and improve customer retention.

Обзор этой функции приведен в видео, посвященном [командам надстроек на ленте приложения Office](https://channel9.msdn.com/events/Build/2016/P551).

> [!NOTE]
> SharePoint catalogs do not support add-in commands. You can deploy add-in commands via [Centralized Deployment](../publish/centralized-deployment.md) or [AppSource](/office/dev/store/submit-to-appsource-via-partner-center), or use [sideloading](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md) to deploy your add-in command for testing.

> [!IMPORTANT]
> В Outlook также поддерживаются команды надстроек. Дополнительные сведения см. в статье [Команды надстроек для Outlook](../outlook/add-in-commands-for-outlook.md).

*Рисунок 1. Надстройка с командами, работающая в классическом приложении Excel*

![Снимок экрана с командой надстройки в приложении Excel](../images/add-in-commands-1.png)

*Рисунок 2. Надстройка с командами, работающая в Excel в Интернете*

![Снимок экрана с командой надстройки в Excel в Интернете](../images/add-in-commands-2.png)

## <a name="command-capabilities"></a>Возможности команд

В настоящее время поддерживаются указанные ниже возможности команд.

> [!NOTE]
> Контентные надстройки на данный момент не поддерживают команды.

### <a name="extension-points"></a>Точки расширения

- Вкладки ленты: расширение возможностей встроенных вкладок или создание пользовательской вкладки.
- Контекстные меню: расширение возможностей выбранных контекстных меню.

### <a name="control-types"></a>Типы элементов управления

- Простые кнопки, запускающие определенные действия.
- Простые раскрывающиеся меню с кнопками, которые запускают действия.

### <a name="actions"></a>Действия

- ShowTaskpane: отображает одну или несколько областей, в которые можно загрузить пользовательские HTML-страницы.
- ExecuteFunction - Loads an invisible HTML page and then execute a JavaScript function within it. To show UI within your function (such as errors, progress, or additional input) you can use the [displayDialog](/javascript/api/office/office.ui) API.  

### <a name="default-enabled-or-disabled-status-preview"></a>Состояние по умолчанию: "Включено" или "Отключено" (предварительная версия)

Вы можете указать, включена или отключена команда при запуске надстройки, а также изменять параметр программными средствами.

> [!NOTE]
> Эта возможность доступна в предварительной версии и поддерживается не всеми узлами и сценариями. Дополнительные сведения см. в статье о [Включение и отключение команд надстроек](disable-add-in-commands.md).

## <a name="supported-platforms"></a>Поддерживаемые платформы

В настоящее время команды надстроек поддерживаются на перечисленных ниже платформах.

- Office для Windows (сборка 16.0.6769+, подключенная к подписке на Microsoft 365)
- Office 2019 для Windows
- Office для Mac (сборка 15.33+, подключенная к подписке на Microsoft 365)
- Office 2019 для Mac
- Office в Интернете

> [!NOTE]
> Сведения о поддержке Outlook см. в[Команды надстройки для Outlook](../outlook/add-in-commands-for-outlook.md).

## <a name="debugging"></a>Отладка

Чтобы отлаживать команду надстройки, необходимо запустить ее в Office в Интернете. Дополнительные сведения см. в статье [Отладка надстроек в Office в Интернете](../testing/debug-add-ins-in-office-online.md)

## <a name="best-practices"></a>Рекомендации

При разработке надстроек придерживайтесь следующих рекомендаций:

- Use commands to represent a specific action with a clear and specific outcome for users. Do not combine multiple actions in a single button.
- Provide granular actions that make common tasks within your add-in more efficient to perform. Minimize the number of steps an action takes to complete.
- Расположение команд на ленте приложения Office:
    - Place commands on an existing tab (Insert, Review, and so on) if the functionality provided fits there. For example, if your add-in enables users to insert media, add a group to the Insert tab. Note that not all tabs are available across all Office versions. For more information, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).
    - Place commands on the Home tab if the functionality doesn't fit on another tab, and you have fewer than six top-level commands. You can also add commands to the Home tab if your add-in needs to work across Office versions (such as Office on the web or desktop) and a tab is not available in all versions (for example, the Design tab doesn't exist in Office on the web).  
    - Добавляйте команды на пользовательскую вкладку, если надстройка содержит более шести команд верхнего уровня.
    - Name your group to match the name of your add-in. If you have multiple groups, name each group based on the functionality that the commands in that group provide.
    - Не добавляйте избыточные кнопки, чтобы надстройка занимала больше места на экране.

     > [!NOTE]
     > Надстройки, которые занимают слишком много места, могут не пройти [проверку в AppSource](/legal/marketplace/certification-policies).

- [Руководство по оформлению значков](add-in-icons.md) подходит для всех значков.
- Предоставьте версию надстройки, которая работает в ведущих приложениях, не поддерживающих команды. Один манифест надстройки может работать в ведущих приложениях независимо от того, поддерживают ли они команды.

   *Рис. 3. Надстройка области задач в Office 2013 и эта же надстройка, использующая команды надстройки в Office 2016*

   ![Снимок экрана: надстройка области задач в Office 2013 и эта же надстройка, использующая команды надстройки в Office 2016](../images/office-task-pane-add-ins.png)


## <a name="next-steps"></a>Дальнейшие действия

Лучший способ начать работу с командами надстроек Office — ознакомиться с [примерами](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/) на сайте GitHub.

Дополнительные сведения об указании команд надстройки в манифесте см. в статье [Создание команд надстроек в манифесте](../develop/create-addin-commands.md) и справочных материалах по [VersionOverrides](../reference/manifest/versionoverrides.md).
