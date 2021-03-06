---
title: Основные концепции команд надстроек
description: Как добавить настраиваемые кнопки ленты и элементы меню в Office в составе надстройки Office
ms.date: 01/29/2021
localization_priority: Priority
ms.openlocfilehash: 1f34a6335949a4cbd2a0f58cdefa12426414770e
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349184"
---
# <a name="add-in-commands-for-excel-powerpoint-and-word"></a>Команды надстроек для Excel, PowerPoint и Word

Команды надстроек — это элементы, которые расширяют пользовательский интерфейс Office и запускают действия в надстройке. Команды надстроек можно использовать для добавления кнопки на ленту или элемента в контекстное меню. Когда пользователи выбирают команду надстройки, они инициируют действия, такие как запуск кода JavaScript или отображение страницы надстройки в области задач. Команды надстройки помогают пользователям находить и использовать вашу надстройку, что может повысить показатель внедрения надстройки и коэффициент удержания клиентов.

Обзор этой функции приведен в видео, посвященном [командам надстроек на ленте приложения Office](https://channel9.msdn.com/events/Build/2016/P551).

> [!NOTE]
> В каталогах SharePoint не поддерживаются команды надстроек. Последние можно развернуть с помощью компонента [централизованного развертывания](../publish/centralized-deployment.md) или [AppSource](/office/dev/store/submit-to-appsource-via-partner-center). Чтобы развернуть команду надстройки для тестирования, выполните [загрузку неопубликованного приложения](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).

> [!IMPORTANT]
> В Outlook также поддерживаются команды надстроек. Дополнительные сведения см. в статье [Команды надстроек для Outlook](../outlook/add-in-commands-for-outlook.md).

*Рисунок 1. Надстройка с командами, работающая в классическом приложении Excel*

![Снимок экрана, на котором выделены команды надстройки на ленте Excel.](../images/add-in-commands-1.png)

*Рисунок 2. Надстройка с командами, работающая в Excel в Интернете*

![Снимок экрана с командами надстроек в Excel в Интернете.](../images/add-in-commands-2.png)

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
- ExecuteFunction загружает невидимую HTML-страницу, а затем выполняет содержащуюся в ней функцию JavaScript. Для показа ошибок, хода выполнения или дополнительных данных функции можно использовать API [displayDialog](/javascript/api/office/office.ui).  

### <a name="default-enabled-or-disabled-status"></a>Состояние по умолчанию: "Включено" или "Отключено"

Вы можете указать, включена или отключена команда при запуске надстройки, а также изменять параметр программными средствами.

> [!NOTE]
> Эта функция поддерживается не всеми приложениями Office и сценариями. Дополнительные сведения см. в статье о [Включение и отключение команд надстроек](disable-add-in-commands.md).

### <a name="position-on-the-ribbon-preview"></a>Расположение на ленте (предварительная версия)

Вы можете указать, где настраиваемая вкладка будет отображаться на ленте приложения Office, например "справа от вкладки «Главная»".

> [!NOTE]
> Эта функция поддерживается не всеми приложениями Office и сценариями. Дополнительные сведения см. в статье [Расположение настраиваемой вкладки на ленте](custom-tab-placement.md).

### <a name="integration-of-built-in-office-buttons-preview"></a>Интеграция встроенных кнопок Office (предварительная версия)

Вы можете вставлять встроенные кнопки ленты Office в свои группы настраиваемых команд и настраиваемые вкладки ленты.

> [!NOTE]
> Эта функция поддерживается не всеми приложениями Office и сценариями. Дополнительные сведения см. в статье [Интеграция встроенных кнопок Office в настраиваемые вкладки](built-in-button-integration.md).

### <a name="contextual-tabs-preview"></a>Контекстные вкладки (предварительная версия)

Вы можете настроить отображение вкладки на ленте только в определенных контекстах, например при выборе диаграммы в Excel.

> [!NOTE]
> Эта функция поддерживается не всеми приложениями Office и сценариями. Дополнительные сведения см. в статье [Создание пользовательских контекстных вкладок в надстройках Office](contextual-tabs.md).

## <a name="supported-platforms"></a>Поддерживаемые платформы

В настоящее время команды надстроек поддерживаются на следующих платформах (за исключением ограничений, указанных в подразделах [Возможности команд](#command-capabilities) ранее).

- Office для Windows (сборка 16.0.6769+, подключенная к подписке на Microsoft 365)
- Office 2019 для Windows
- Office для Mac (сборка 15.33+, подключенная к подписке на Microsoft 365)
- Office 2019 для Mac
- Office в Интернете

> [!NOTE]
> Сведения о поддержке Outlook см. в[Команды надстройки для Outlook](../outlook/add-in-commands-for-outlook.md).

## <a name="debugging"></a>Отладка

Чтобы отлаживать команду надстройки, необходимо запустить ее в Office в Интернете. Дополнительные сведения см. в статье [Отладка надстроек в Office в Интернете](../testing/debug-add-ins-in-office-online.md)

## <a name="best-practices"></a>Рекомендации

При разработке надстроек придерживайтесь следующих рекомендаций.

- Каждая команда должна представлять определенное действие с очевидным и конкретным исходом для пользователей. Не совмещайте несколько действий в одной кнопке.
- Предоставляйте точные действия, которые делают выполнение распространенных задач в надстройке более эффективным. Максимально сократите количество шагов, необходимых для выполнения действия.
- Расположение команд на ленте приложения Office:
  - Помещайте команды на имеющиеся вкладки ("Вставка", "Рецензирование" и т. д.), если соответствующая функция подходит для них. Например, если надстройка позволяет вставлять файлы мультимедиа, добавьте группу на вкладку "Вставка". Обратите внимание, что некоторые вкладки доступны не во всех версиях Office. Дополнительные сведения см. в статье [XML-манифест надстроек Office](../develop/add-in-manifests.md).
  - Добавляйте команды на вкладку "Главная", если соответствующие функции не относятся к другим вкладкам, а надстройка содержит менее шести команд верхнего уровня. Вы также можете добавлять команды на вкладку "Главная", если надстройка должна работать в разных версиях Office (например, Office в Интернете и классических приложениях Office), а нужная вкладка доступна не во всех версиях (например, вкладка "Конструктор" отсутствует в Office в Интернете).  
  - Добавляйте команды на пользовательскую вкладку, если надстройка содержит более шести команд верхнего уровня.
  - Название группы должно соответствовать названию надстройки. Если у вас есть несколько групп, их имена должны быть связаны с функциями, которые выполняют команды из этих групп.
  - Не добавляйте избыточные кнопки, чтобы надстройка занимала больше места на экране.
  - Не размещайте настраиваемую вкладку слева от вкладки "Главная" или переводите на нее фокус по умолчанию при открытии документа, если ваша надстройка не является основным способом взаимодействия с документом. Чрезмерное выделение вашей надстройки создает неудобства и раздражает пользователей и администраторов.
  - Если надстройка является основным способом взаимодействия пользователей с документом и у вас есть настраиваемая вкладка ленты, рассмотрите возможность интеграции кнопок во вкладку для применения функций Office, которые часто требуются пользователям.
  - Если функции, предоставляемые в настраиваемой вкладке, должны быть доступны только в определенных контекстах, используйте [настраиваемые контекстные вкладки](contextual-tabs.md). Если вы используете настраиваемые контекстные вкладки, реализуйте [резервный интерфейс, когда ваша надстройка запускается на платформах, не поддерживающих настраиваемые контекстные вкладки](contextual-tabs.md#implement-an-alternate-ui-experience-when-custom-contextual-tabs-are-not-supported).

  > [!NOTE]
  > Надстройки, которые занимают слишком много места, могут не пройти [проверку в AppSource](/legal/marketplace/certification-policies).

- [Руководство по оформлению значков](add-in-icons.md) подходит для всех значков.
- Предоставьте версию надстройки, которая работает в приложениях Office, не поддерживающих команды. Один манифест надстройки может работать в приложениях независимо от того, поддерживают ли они команды.

   *Рис. 3. Надстройка области задач в Office 2013 и эта же надстройка, использующая команды надстройки в Office 2016*

   ![Снимок экрана со сравнением надстройки области задач в Office 2013 и этой же надстройки, использующей команды надстройки в Office 2016. В версии 2013 в области задач должны содержаться все команды, в то время как в версии 2016 эти команды могут быть на ленте.](../images/office-task-pane-add-ins.png)

## <a name="next-steps"></a>Дальнейшие действия

Лучший способ начать работу с командами надстроек Office — ознакомиться с [примерами](https://github.com/OfficeDev/Office-Add-in-Commands-Samples/) на сайте GitHub.

Дополнительные сведения об указании команд надстройки в манифесте см. в статье [Создание команд надстроек в манифесте](../develop/create-addin-commands.md) и справочных материалах по [VersionOverrides](../reference/manifest/versionoverrides.md).
