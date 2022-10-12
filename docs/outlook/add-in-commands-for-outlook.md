---
title: Команды надстроек Outlook
description: Команды надстроек Outlook предоставляют доступ к определенным действиям надстройки с ленты, добавляя на нее кнопки или раскрывающиеся меню.
ms.date: 10/11/2022
ms.localizationpriority: high
ms.openlocfilehash: d029fd4acc1a32c912c73d6e5f468b9c217b9262
ms.sourcegitcommit: 787fbe4d4a5462ff6679ad7fd00748bf07391610
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/12/2022
ms.locfileid: "68546461"
---
# <a name="add-in-commands-for-outlook"></a>Команды надстроек Outlook

Outlook add-in commands provide ways to initiate specific add-in actions from the ribbon by adding buttons or drop-down menus. This lets users access add-ins in a simple, intuitive, and unobtrusive way. Because they offer increased functionality in a seamless manner, you can use add-in commands to create more engaging solutions.

> [!NOTE]
> Команды надстроек доступны только в Outlook 2013 или более поздней версии для Windows, Outlook 2016 или более поздней версии для Mac, Outlook для iOS, Outlook для Android, Outlook в Интернете для Exchange 2016 или более поздней версии, Outlook в Интернете для Microsoft 365 и Outlook.com.
>
> Для поддержки команд надстроек в Outlook 2013 необходимы три обновления:
> - [Обновление для системы безопасности для Outlook от 8 марта 2016 г.](https://support.microsoft.com/kb/3114829)
> - [Обновление для системы безопасности для Office (KB3114816) от 8 марта 2016 г.](https://support.microsoft.com/topic/3d3eb171-78c2-0e61-62a2-85723bc4bcc0)
> - [Обновление для системы безопасности для Office (KB3114828) от 8 марта 2016 г.](https://support.microsoft.com/topic/54437016-d1e0-7aac-dbb7-4ecfbd57f5f0)
>
> Для поддержки команд надстроек в Exchange 2016 требуется [накопительный пакет обновления 5](https://support.microsoft.com/topic/d67d7693-96a4-fb6e-b60b-e64984e267bd).

> [!TIP]
> Если надстройка использует XML-манифест, команды надстроек доступны только для надстроек, которые не используют [правила ItemHasAttachment, ItemHasKnownEntity или ItemHasRegularExpressionMatch](activation-rules.md) , чтобы ограничить типы элементов, в которых они активируются. Однако [контекстные](contextual-outlook-add-ins.md) надстройки могут представлять различные команды в зависимости от того, является ли текущий выбранный элемент сообщением или встречей, и могут отображаться в сценариях чтения или создания. [Рекомендуем](../concepts/add-in-development-best-practices.md) использовать команды надстроек по мере возможности.

## <a name="create-the-ui-for-the-add-in-command"></a>Создание пользовательского интерфейса для команды надстройки

Команды надстройки объявляются в манифесте надстройки. Разметка зависит от типа манифеста.

# <a name="xml-manifest"></a>[XML-манифест](#tab/xmlmanifest)

Команды надстройки объявляются в [элементе VersionOverrides](/javascript/api/manifest/versionoverrides). Этот элемент является дополнением к схеме XML-манифеста версии 1.1, которая обеспечивает обратную совместимость. В клиенте, который не поддерживает узел **\<VersionOverrides\>**, имеющиеся надстройки продолжат работать так же, как и без команд надстроек.

В записях манифеста **\<VersionOverrides\>** указывается множество свойств надстройки, например приложение, типы элементов управления, добавляемых на ленту, текст, значки и соответствующие функции.

When an add-in needs to provide status updates, such as progress indicators or error messages, it must do so through the [notification APIs](/javascript/api/outlook/office.notificationmessages). The processing for the notifications must also be defined in a separate HTML file that is specified in the `FunctionFile` node of the manifest.

# <a name="teams-manifest-developer-preview"></a>[Манифест Teams (предварительная версия для разработчиков)](#tab/jsonmanifest)

Команды надстроек объявляются со свойствами extensions.runtimes и extensions.ribbons. Эти свойства задают для надстройки множество вещей, таких как приложение, типы элементов управления, добавляемого на ленту, текст, значки и все связанные функции.

Если надстройка должна предоставлять сведения о состоянии, например индикаторы хода выполнения или сообщения об ошибках, это необходимо сделать с помощью [API-интерфейсов уведомлений](/javascript/api/outlook/office.notificationmessages). Обработка уведомлений также должна быть определена в отдельном HTML-файле, указанном в свойстве runtimes.code.page манифеста.

---
### <a name="icons"></a>Значки

Developers should define icons for all required sizes so that the add-in commands will adjust smoothly along with the ribbon. The required icon sizes are 80 x 80 pixels, 32 x 32 pixels, and 16 x 16 pixels for desktop, and 48 x 48 pixels, 32 x 32 pixels, and 25 x 25 pixels for mobile.

## <a name="how-do-add-in-commands-appear"></a>Отображение команд надстройки

Команда надстройки отображается на ленте в виде кнопки или элемента в раскрывающемся меню. Когда пользователь устанавливает надстройку, ее команды отображаются в пользовательском интерфейсе как группа кнопок. Она может располагаться на вкладке по умолчанию или пользовательской вкладке ленты. Для сообщений вкладкой по умолчанию является вкладка **Главная** или **Сообщение**. Для календаря вкладкой по умолчанию является вкладка **Собрание**, **Экземпляр собрания**, **Ряд собраний** или **Встреча**. Для расширений модуля вкладкой по умолчанию является пользовательская вкладка. На вкладке по умолчанию у каждой надстройки может быть одна группа ленты, содержащая до 6 команд. Пользовательские вкладки могут включать до 10 групп, по 6 команд в каждой. У надстройки может быть только одна пользовательская вкладка.

Команды надстройки отображаются в меню переполнения по мере заполнения ленты элементами. Команды надстройки обычно группируются вместе.

![Кнопки команд надстройки на ленте.](../images/commands-normal.png)

![Кнопки команд надстройки на ленте и в меню переполнения.](../images/commands-collapsed.png)

When an add-in command is added to an add-in, the add-in name is removed from the app bar. Only the add-in command button on the ribbon remains.

### <a name="modern-outlook-on-the-web"></a>Современная версия Outlook в Интернете

В Outlook в Интернете имя надстройки отображается в меню переполнения. Если у надстройки есть несколько команд, вы можете развернуть меню надстройки, чтобы просмотреть группу кнопок с именем надстройки.

![Меню переполнения, в котором находятся кнопки команд надстройки.](../images/commands-overflow-menu-web.png)

![Меню переполнения, отображающее кнопки команд надстройки.](../images/commands-overflow-menu-expand-web.png)

## <a name="what-are-the-types-of-add-in-commands"></a>Какие существуют типы команд надстройки?

Пользовательский интерфейс для команды надстройки состоит из кнопки ленты или элемента в раскрывающемся меню. Существует два типа команд надстроек, основанных на типе действия, активируемого командой.

- **Команды области задач**: кнопка или элемент меню открывает панель задач надстроек. Этот тип команды надстройки добавляется с разметкой в манифесте. "Код программной части" команды предоставляется Office.
- **Команды функций**: кнопка или элемент меню запускает любой произвольный JavaScript. Код почти всегда вызывает API в библиотеке JavaScript для Office, но это не обязательно. Этот тип надстройки обычно не отображает пользовательский интерфейс, кроме самой кнопки или элемента меню. Обратите внимание на следующие сведения о командах функций:

   - Запущенная функция может вызвать метод [displayDialogAsync](/javascript/api/office/office.ui#office-office-ui-displaydialogasync-member(1)) для отображения диалогового окна, что является хорошим способом отобразить ошибку, показать прогресс или запросить ввод от пользователя.
   - Среда выполнения, в которой выполняется команда функции, является полной средой выполнения [на основе браузера](../testing/runtimes.md#browser-runtime). Она может отображать HTML и обращаться к Интернету для отправки или получения данных.

### <a name="run-a-function-command"></a>Выполнение команды функции

Use an add-in command button that executes a JavaScript function for scenarios where the user doesn't need to make any additional selections to initiate the action. This can be for actions such as track, remind me, or print, or scenarios when the user wants more in-depth information from a service.

В расширениях модуля кнопка команды надстройки может выполнять функции JavaScript, взаимодействующие с содержимым в основном пользовательском интерфейсе.

![Кнопка, которая выполняет функцию, на ленте Outlook.](../images/commands-uiless-button-1.png)

### <a name="launch-a-task-pane"></a>Запуск области задач

Use an add-in command button to launch a task pane for scenarios where a user needs to interact with an add-in for a longer period of time. For example, the add-in requires changes to settings or the completion of many fields.

The default width of the vertical task pane is 320 px. The vertical task pane can be resized in both the Outlook Explorer and inspector. The pane can be resized in the same way the to-do pane and list view resize.

![Кнопка, открывающая область задач, на ленте Outlook.](../images/commands-task-pane-button-1.png)

<br/>

This screenshot shows an example of a vertical task pane. The pane opens with the name of the add-in command in the top left corner. Users can use the **X** button in the upper-right corner of the pane to close the add-in when they are finished using it. By default, this pane will not persist across messages. Add-ins can [support pinning](pinnable-taskpane.md) for the task pane and receive events when a new message is selected. All UI elements rendered in the task pane, aside from the add-in name and the close button, are provided by the add-in.

If a user chooses another add-in command that opens a task pane, the task pane is replaced with the recently used command. If a user chooses an add-in command button that executes a function, or drop-down menu while the task pane is open, the action will be completed and the task pane will remain open.

### <a name="drop-down-menu"></a>Раскрывающееся меню

Команда надстройки с раскрывающимся меню определяет статический список элементов. Меню может содержать любой набор элементов, которые выполняют функцию или открывают область задач. Подменю не поддерживаются.

![Кнопка раскрывающегося меню на ленте Outlook.](../images/commands-menu-button-1.png)

## <a name="where-do-add-in-commands-appear-in-the-ui"></a>Отображение команд надстройки в пользовательском интерфейсе

Команды надстроек поддерживаются для четырех сценариев:

### <a name="reading-a-message"></a>Просмотр сообщения

Когда пользователь просматривает сообщение в области чтения или на вкладке **Сообщение** для всплывающей формы чтения, команды надстройки, добавленные на вкладку по умолчанию, отображаются на вкладке **Главная**.

### <a name="composing-a-message"></a>Создание сообщения

Когда пользователь создает сообщение, команды надстройки, добавленные на вкладку по умолчанию, отображаются на вкладке **Сообщение**.

### <a name="creating-or-viewing-an-appointment-or-meeting-as-the-organizer"></a>Создание или просмотр встречи или собрания организатором

When creating or viewing an appointment or meeting as the organizer, add-in commands added to the default tab appear on the **Meeting**, **Meeting Occurrence**, **Meeting Series**, or **Appointment** tabs on pop-out forms. However, if the user selects an item in the calendar but doesn't open the pop-out, the add-in's ribbon group won't be visible in the ribbon.

### <a name="viewing-a-meeting-as-an-attendee"></a>Просмотр собрания участником

When viewing a meeting as an attendee, add-in commands added to the default tab appear on the **Meeting**, **Meeting Occurrence**, or **Meeting Series** tabs on pop-out forms. However, if a user selects an item in the calendar but doesn't open the pop-out, the add-in's ribbon group won't be visible in the ribbon

### <a name="using-a-module-extension"></a>Использование расширения модуля

Если используется расширение модуля, команды надстройки отображаются на пользовательской вкладке расширения.

## <a name="see-also"></a>См. также

- [Надстройка Outlook "Демонстрация команд надстройки"](https://github.com/officedev/outlook-add-in-command-demo)
- [Создание команд надстроек в манифесте для Excel, PowerPoint и Word](../develop/create-addin-commands.md)
- [Отладка команд функций в надстройках Outlook](debug-ui-less.md)
- [Руководство. Сборка надстройки Outlook для создания сообщения](../tutorials/outlook-tutorial.md)
