---
title: Команды надстроек Outlook
description: Команды надстроек Outlook предоставляют доступ к определенным действиям надстройки с ленты, добавляя на нее кнопки или раскрывающиеся меню.
ms.date: 12/05/2019
localization_priority: Priority
ms.openlocfilehash: 8d0fce06277de6ef699dbde640f13c657ca4919c
ms.sourcegitcommit: 9da68c00ecc00a2f307757e0f5a903a8e31b7769
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/22/2020
ms.locfileid: "43785790"
---
# <a name="add-in-commands-for-outlook"></a><span data-ttu-id="728b9-103">Команды надстроек Outlook</span><span class="sxs-lookup"><span data-stu-id="728b9-103">Add-in commands for Outlook</span></span>

<span data-ttu-id="728b9-p101">Команды надстроек Outlook позволяют вызывать определенные действия надстроек с ленты путем добавления кнопок или раскрывающихся меню. Благодаря этому пользователи могут легко получать доступ к надстройкам простым и интуитивно понятным способом. Так как команды надстроек предлагают удобные расширенные функции, с их помощью вы можете создавать более интересные решения.</span><span class="sxs-lookup"><span data-stu-id="728b9-p101">Outlook add-in commands provide ways to initiate specific add-in actions from the ribbon by adding buttons or drop-down menus. This lets users access add-ins in a simple, intuitive, and unobtrusive way. Because they offer increased functionality in a seamless manner, you can use add-in commands to create more engaging solutions.</span></span>

> [!NOTE]
> <span data-ttu-id="728b9-107">Команды надстроек доступны только в Outlook 2013 или более поздней версии для Windows, Outlook 2016 для Mac, Outlook для iPhone, Outlook для Android, Outlook в Интернете для Exchange 2016 или более поздней версии, Outlook в Интернете для Office 365 и Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="728b9-107">Add-in commands are available only in Outlook 2013 or later on Windows, Outlook 2016 or later on Mac, Outlook on iPhone, Outlook on Android, Outlook on the web for Exchange 2016 or later, and Outlook on the web for Office 365 and Outlook.com.</span></span>
>
> <span data-ttu-id="728b9-108">Для поддержки команд надстроек в Outlook 2013 необходимы три обновления:</span><span class="sxs-lookup"><span data-stu-id="728b9-108">Support for add-in commands in Outlook 2013 requires three updates:</span></span>
> - [<span data-ttu-id="728b9-109">Обновление для системы безопасности для Outlook от 8 марта 2016 г.</span><span class="sxs-lookup"><span data-stu-id="728b9-109">March 8, 2016 security update for Outlook</span></span>](https://support.microsoft.com/kb/3114829)
> - [<span data-ttu-id="728b9-110">Обновление для системы безопасности для Office (KB3114816) от 8 марта 2016 г.</span><span class="sxs-lookup"><span data-stu-id="728b9-110">March 8, 2016 security update for Office (KB3114816)</span></span>](https://support.microsoft.com/help/3114816/march-8,-2016,-update-for-office-2013-kb3114816)
> - [<span data-ttu-id="728b9-111">Обновление для системы безопасности для Office (KB3114828) от 8 марта 2016 г.</span><span class="sxs-lookup"><span data-stu-id="728b9-111">March 8, 2016 security update for Office (KB3114828)</span></span>](https://support.microsoft.com/help/3114828/march-8,-2016,-update-for-office-2013-kb3114828)
>
> <span data-ttu-id="728b9-112">Для поддержки команд надстроек в Exchange 2016 требуется [накопительный пакет обновления 5](https://support.microsoft.com/help/4012106/cumulative-update-5-for-exchange-server-2016).</span><span class="sxs-lookup"><span data-stu-id="728b9-112">Support for add-in commands in Exchange 2016 requires [Cumulative Update 5](https://support.microsoft.com/help/4012106/cumulative-update-5-for-exchange-server-2016).</span></span>

<span data-ttu-id="728b9-p102">Команды надстроек доступны в тех надстройках, где не используются [правила ItemHasAttachment, ItemHasKnownEntity и ItemHasRegularExpressionMatch](activation-rules.md), для ограничения типов элементов, активирующих надстройки. В [контекстных надстройках](contextual-outlook-add-ins.md) могут отображаться разные команды в зависимости от того, какой элемент выбран (сообщение или встреча), и они могут работать в режимах просмотра и создания. [Рекомендуем](../concepts/add-in-development-best-practices.md) использовать команды надстроек по мере возможности.</span><span class="sxs-lookup"><span data-stu-id="728b9-p102">Add-in commands are only available for add-ins that do not use [ItemHasAttachment, ItemHasKnownEntity, or ItemHasRegularExpressionMatch rules](activation-rules.md) to limit the types of items they activate on. However, [contextual add-ins](contextual-outlook-add-ins.md) can present different commands depending on whether the currently selected item is a message or appointment, and can choose to appear in read or compose scenarios. Using add-in commands if possible is a [best practice](../concepts/add-in-development-best-practices.md).</span></span>

## <a name="creating-the-add-in-command"></a><span data-ttu-id="728b9-116">Создание команды надстройки</span><span class="sxs-lookup"><span data-stu-id="728b9-116">Creating the add-in command</span></span>

<span data-ttu-id="728b9-p103">Команды надстроек объявляются в манифесте в элементе [VersionOverrides](../reference/manifest/versionoverrides.md). Этот элемент является дополнением к схеме манифестов версии 1.1, которая обеспечивает обратную совместимость. В клиенте, который не поддерживает узел `VersionOverrides`, имеющиеся надстройки продолжат работать так же, как и без команд надстроек.</span><span class="sxs-lookup"><span data-stu-id="728b9-p103">Add-in commands are declared in the add-in manifest in the [VersionOverrides element](../reference/manifest/versionoverrides.md). This element is an addition to the manifest schema v1.1 that ensures backward compatibility. In a client that doesn't support `VersionOverrides`, existing add-ins will continue to function as they did without add-in commands.</span></span>

<span data-ttu-id="728b9-120">В записях манифеста `VersionOverrides` указывается множество свойств надстройки, например ведущее приложение, типы элементов управления, добавляемых на ленту, текст, значки и соответствующие функции.</span><span class="sxs-lookup"><span data-stu-id="728b9-120">The `VersionOverrides` manifest entries specify many things for the add-in, such as the host, types of controls to add to the ribbon, the text, the icons, and any associated functions.</span></span>

<span data-ttu-id="728b9-p104">Если надстройка должна показывать индикаторы хода выполнения или сообщения об ошибках, она должна делать это через [API уведомления](/javascript/api/outlook/office.notificationmessages). Обработка уведомлений также должна быть определена в отдельном HTML-файле, указанном в узле `FunctionFile` манифеста.</span><span class="sxs-lookup"><span data-stu-id="728b9-p104">When an add-in needs to provide status updates, such as progress indicators or error messages, it must do so through the [notification APIs](/javascript/api/outlook/office.notificationmessages). The processing for the notifications must also be defined in a separate HTML file that is specified in the `FunctionFile` node of the manifest.</span></span>

<span data-ttu-id="728b9-p105">Разработчикам следует определить значки всех требуемых размеров, чтобы команды надстройки плавно изменялись при изменении размера ленты. Количество пикселей, составляющее требуемый размер значков: 80 x 80, 32 x 32 и 16 x 16 для компьютера, а также 48 x 48, 32 x 32 и 25 x 25 для мобильного устройства.</span><span class="sxs-lookup"><span data-stu-id="728b9-p105">Developers should define icons for all required sizes so that the add-in commands will adjust smoothly along with the ribbon. The required icon sizes are 80 x 80 pixels, 32 x 32 pixels, and 16 x 16 pixels for desktop, and 48 x 48 pixels, 32 x 32 pixels, and 25 x 25 pixels for mobile.</span></span>

<span data-ttu-id="728b9-125">Дополнительные сведения о создании команд надстроек см. в статье [Создание команд надстроек в манифесте](../develop/create-addin-commands.md).</span><span class="sxs-lookup"><span data-stu-id="728b9-125">For more information about creating add-in commands, see [Create add-in commands in your manifest](../develop/create-addin-commands.md).</span></span>

## <a name="how-do-add-in-commands-appear"></a><span data-ttu-id="728b9-126">Отображение команд надстройки</span><span class="sxs-lookup"><span data-stu-id="728b9-126">How do add-in commands appear?</span></span>

<span data-ttu-id="728b9-p106">Команда надстройки отображается на ленте в виде кнопки. Когда пользователь устанавливает надстройку, ее команды отображаются в пользовательском интерфейсе как группа кнопок. Они могут отображаться либо на вкладке ленты по умолчанию, либо на пользовательской вкладке. Для сообщений по умолчанию используется вкладка **Главная** или **Сообщение**. Для календаря стандартной является вкладка **Сообщение**, **Экземпляр собрания**, **Ряд собраний** или **Встреча**. Для расширений модуля стандартной является пользовательская вкладка. На вкладке по умолчанию у каждой надстройки может быть одна группа ленты, включающая до 6 команд. Пользовательские вкладки могут включать до 10 групп, по 6 команд в каждой. У надстройки может быть только одна пользовательская вкладка.</span><span class="sxs-lookup"><span data-stu-id="728b9-p106">An add-in command appears on the ribbon as a button. When a user installs an add-in, its commands appear in the UI as a group of buttons. This can either be on the ribbon's default tab or on a custom tab. For messages, the default is either the **Home** or **Message** tab. For the calendar, the default is the **Meeting**, **Meeting Occurrence**, **Meeting Series**, or **Appointment** tab. For module extensions, the default is a custom tab. On the default tab, each add-in can have one ribbon group with up to 6 commands. On custom tabs, the add-in can have up to 10 groups, each with 6 commands. Add-ins are limited to only one custom tab.</span></span>

<span data-ttu-id="728b9-132">Команды надстройки отображаются в меню переполнения по мере заполнения ленты элементами.</span><span class="sxs-lookup"><span data-stu-id="728b9-132">As the ribbon gets more crowded, add-in commands will be displayed in the overflow menu.</span></span> <span data-ttu-id="728b9-133">Команды надстройки обычно группируются вместе.</span><span class="sxs-lookup"><span data-stu-id="728b9-133">The add-in commands for an add-in are usually grouped together.</span></span>

![Кнопки команд надстройки на ленте](../images/commands-normal.png)

![Кнопки команд надстройки на ленте и в меню переполнения](../images/commands-collapsed.png)

<span data-ttu-id="728b9-p108">При добавлении команды в надстройку имя надстройки удаляется с панели приложения. На ленте остается только кнопка надстройки.</span><span class="sxs-lookup"><span data-stu-id="728b9-p108">When an add-in command is added to an add-in, the add-in name is removed from the app bar. Only the add-in command button on the ribbon remains.</span></span>

### <a name="modern-outlook-on-the-web"></a><span data-ttu-id="728b9-138">Современная версия Outlook в Интернете</span><span class="sxs-lookup"><span data-stu-id="728b9-138">Modern Outlook on the web</span></span>

<span data-ttu-id="728b9-139">В Outlook в Интернете имя надстройки отображается в меню переполнения.</span><span class="sxs-lookup"><span data-stu-id="728b9-139">In Outlook on the web, the add-in name is displayed in an overflow menu.</span></span> <span data-ttu-id="728b9-140">Если у надстройки есть несколько команд, вы можете развернуть меню надстройки, чтобы просмотреть группу кнопок с именем надстройки.</span><span class="sxs-lookup"><span data-stu-id="728b9-140">If the add-in has multiple add-in commands, you can expand the add-in menu to see the group of buttons labeled with the add-in name.</span></span>

![Меню переполнения, в котором находятся кнопки команд надстройки](../images/commands-overflow-menu-web.png)

![Меню переполнения, отображающее кнопки команд надстройки](../images/commands-overflow-menu-expand-web.png)

## <a name="what-ux-shapes-exist-for-add-in-commands"></a><span data-ttu-id="728b9-143">Оформление команд надстройки в пользовательском интерфейсе</span><span class="sxs-lookup"><span data-stu-id="728b9-143">What UX shapes exist for add-in commands?</span></span>

<span data-ttu-id="728b9-p110">Оформление пользовательского интерфейса для команды надстройки представляет собой вкладку ленты с кнопками, которые могут выполнять различные функции. В настоящий момент поддерживаются три варианта оформления пользовательского интерфейса:</span><span class="sxs-lookup"><span data-stu-id="728b9-p110">The UX shape for an add-in command consists of a ribbon tab in the host application that contains buttons that can perform various functions. Currently, three UI shapes are supported:</span></span>

- <span data-ttu-id="728b9-146">кнопка, выполняющая функцию JavaScript;</span><span class="sxs-lookup"><span data-stu-id="728b9-146">A button that executes a JavaScript function</span></span>
- <span data-ttu-id="728b9-147">кнопка, запускающая область задач.</span><span class="sxs-lookup"><span data-stu-id="728b9-147">A button that launches a task pane</span></span>
- <span data-ttu-id="728b9-148">кнопка, которая выводит раскрывающееся меню с одной или несколькими кнопками остальных двух типов;</span><span class="sxs-lookup"><span data-stu-id="728b9-148">A button that shows a drop-down menu with one or more buttons of the other two types</span></span>

### <a name="executing-a-javascript-function"></a><span data-ttu-id="728b9-149">Выполнение функции JavaScript</span><span class="sxs-lookup"><span data-stu-id="728b9-149">Executing a JavaScript function</span></span>

<span data-ttu-id="728b9-p111">Командные кнопки, выполняющие функцию JavaScript, используются для тех случаев, когда пользователю не требуется выбирать дополнительные параметры действия. Это могут быть такие действия, как отслеживание, напоминание или печать. Их также можно использовать для получения более подробных сведений от службы.</span><span class="sxs-lookup"><span data-stu-id="728b9-p111">Use an add-in command button that executes a JavaScript function for scenarios where the user doesn't need to make any additional selections to initiate the action. This can be for actions such as track, remind me, or print, or scenarios when the user wants more in-depth information from a service.</span></span>

<span data-ttu-id="728b9-152">В расширениях модуля кнопка команды надстройки может выполнять функции JavaScript, взаимодействующие с содержимым в основном пользовательском интерфейсе.</span><span class="sxs-lookup"><span data-stu-id="728b9-152">In module extensions, the add-in command button can execute JavaScript functions that interact with the content in the main user interface.</span></span>

![Кнопка, которая выполняет функцию, на ленте Outlook.](../images/commands-uiless-button-1.png)

### <a name="launching-a-task-pane"></a><span data-ttu-id="728b9-154">Запуск области задач</span><span class="sxs-lookup"><span data-stu-id="728b9-154">Launching a task pane</span></span>

<span data-ttu-id="728b9-p112">Командные кнопки надстроек, запускающие область задач, используются в тех случаях, когда пользователю нужно работать с надстройкой длительное время. Например, надстройке может потребоваться изменить параметры или заполнить множество полей.</span><span class="sxs-lookup"><span data-stu-id="728b9-p112">Use an add-in command button to launch a task pane for scenarios where a user needs to interact with an add-in for a longer period of time. For example, the add-in requires changes to settings or the completion of many fields.</span></span>

<span data-ttu-id="728b9-p113">По умолчанию ширина вертикальной области задач составляет 320 точек. Размер области задач можно изменить как в проводнике Outlook, так и в инспекторе. Изменить размеры области можно будет точно так же, как и размеры панели списка дел и представления списка.</span><span class="sxs-lookup"><span data-stu-id="728b9-p113">The default width of the vertical task pane is 320 px. The vertical task pane can be resized in both the Outlook Explorer and inspector. The pane can be resized in the same way the to-do pane and list view resize.</span></span>

![Кнопка, открывающая область задач, на ленте Outlook.](../images/commands-task-pane-button-1.png)

<br/>

<span data-ttu-id="728b9-p114">На этом снимке экрана показан пример вертикальной области задач. В левом верхнем углу области указано имя команды надстройки. Пользователи могут нажать кнопку **X** в правом верхнем углу области, чтобы закрыть надстройку после завершения работы с ней. По умолчанию эта область не сохраняется в сообщениях. Надстройки могут [поддерживать закрепление](pinnable-taskpane.md) в области задач и получать события при выборе нового сообщения. Все элементы пользовательского интерфейса в области задач, кроме имени надстройки и кнопки "Закрыть", предоставляются надстройкой.</span><span class="sxs-lookup"><span data-stu-id="728b9-p114">This screenshot shows an example of a vertical task pane. The pane opens with the name of the add-in command in the top left corner. Users can use the **X** button in the upper-right corner of the pane to close the add-in when they are finished using it. By default, this pane will not persist across messages. Add-ins can [support pinning](pinnable-taskpane.md) for the task pane and receive events when a new message is selected. All UI elements rendered in the task pane, aside from the add-in name and the close button, are provided by the add-in.</span></span>

<span data-ttu-id="728b9-p115">Если пользователь выбирает другую команду надстройки, которая открывает область задач, то область задач заменяется выбранной командой. Когда пользователь нажимает командную кнопку, которая выполняет функцию, или раскрывающееся меню при открытой области задач, действие будет выполнено, а область задач останется открытой.</span><span class="sxs-lookup"><span data-stu-id="728b9-p115">If a user chooses another add-in command that opens a task pane, the task pane is replaced with the recently used command. If a user chooses an add-in command button that executes a function, or drop-down menu while the task pane is open, the action will be completed and the task pane will remain open.</span></span>

### <a name="drop-down-menu"></a><span data-ttu-id="728b9-169">Раскрывающееся меню</span><span class="sxs-lookup"><span data-stu-id="728b9-169">Drop-down menu</span></span>

<span data-ttu-id="728b9-p116">Команда надстройки с раскрывающимся меню определяет статический список кнопок. Меню может содержать любой набор кнопок, которые выполняют функцию или открывают область задач. Подменю не поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="728b9-p116">A drop-down menu add-in command defines a static list of buttons. The buttons within the menu can be any mix of buttons that execute a function or buttons that open a task pane. Submenus are not supported.</span></span>

![Кнопка раскрывающегося меню на ленте Outlook.](../images/commands-menu-button-1.png)

## <a name="where-do-add-in-commands-appear-in-the-ui"></a><span data-ttu-id="728b9-174">Отображение команд надстройки в пользовательском интерфейсе</span><span class="sxs-lookup"><span data-stu-id="728b9-174">Where do add-in commands appear in the UI?</span></span>

<span data-ttu-id="728b9-175">Команды надстроек поддерживаются для четырех сценариев:</span><span class="sxs-lookup"><span data-stu-id="728b9-175">Add-in commands are supported for four scenarios:</span></span>

### <a name="reading-a-message"></a><span data-ttu-id="728b9-176">Просмотр сообщения</span><span class="sxs-lookup"><span data-stu-id="728b9-176">Reading a message</span></span>

<span data-ttu-id="728b9-177">Когда пользователь просматривает сообщение в области чтения или на вкладке **Сообщение** для всплывающей формы чтения, команды надстройки, добавленные на вкладку по умолчанию, отображаются на вкладке **Главная**.</span><span class="sxs-lookup"><span data-stu-id="728b9-177">When the user is reading a message in the reading pane or in the **Message** tab for a pop-out read form, add-in commands added to the default tab appear on the **Home** tab.</span></span>

### <a name="composing-a-message"></a><span data-ttu-id="728b9-178">Создание сообщения</span><span class="sxs-lookup"><span data-stu-id="728b9-178">Composing a message</span></span>

<span data-ttu-id="728b9-179">Когда пользователь создает сообщение, команды надстройки, добавленные на вкладку по умолчанию, отображаются на вкладке **Сообщение**.</span><span class="sxs-lookup"><span data-stu-id="728b9-179">When the user is composing a message, add-in commands added to the default tab appear on the **Message** tab.</span></span>

### <a name="creating-or-viewing-an-appointment-or-meeting-as-the-organizer"></a><span data-ttu-id="728b9-180">Создание или просмотр встречи или собрания организатором</span><span class="sxs-lookup"><span data-stu-id="728b9-180">Creating or viewing an appointment or meeting as the organizer</span></span>

<span data-ttu-id="728b9-p117">Когда организатор создает или просматривает встречу или собрание, команды надстройки, добавленные на вкладку по умолчанию, отображаются на вкладках **Собрание**, **Экземпляр собрания**, **Ряд собраний** или **Встреча** во всплывающих формах. Тем не менее, если пользователь выберет элемент календаря, но не откроет всплывающее окно, группа надстройки не будет отображаться на ленте.</span><span class="sxs-lookup"><span data-stu-id="728b9-p117">When creating or viewing an appointment or meeting as the organizer, add-in commands added to the default tab appear on the **Meeting**, **Meeting Occurrence**, **Meeting Series**, or **Appointment** tabs on pop-out forms. However, if the user selects an item in the calendar but doesn't open the pop-out, the add-in's ribbon group won't be visible in the ribbon.</span></span>

### <a name="viewing-a-meeting-as-an-attendee"></a><span data-ttu-id="728b9-183">Просмотр собрания участником</span><span class="sxs-lookup"><span data-stu-id="728b9-183">Viewing a meeting as an attendee</span></span>

<span data-ttu-id="728b9-p118">Когда участник просматривает собрание, команды надстройки, добавленные на вкладку по умолчанию, отображаются на вкладках **Собрание**, **Экземпляр собрания** или **Ряд собраний** всплывающих форм. Тем не менее, если пользователь выберет элемент календаря, не открывая всплывающее окно, группа надстройки не будет отображаться на ленте.</span><span class="sxs-lookup"><span data-stu-id="728b9-p118">When viewing a meeting as an attendee, add-in commands added to the default tab appear on the **Meeting**, **Meeting Occurrence**, or **Meeting Series** tabs on pop-out forms. However, if a user selects an item in the calendar but doesn't open the pop-out, the add-in's ribbon group won't be visible in the ribbon</span></span>

### <a name="using-a-module-extension"></a><span data-ttu-id="728b9-186">Использование расширения модуля</span><span class="sxs-lookup"><span data-stu-id="728b9-186">Using a module extension</span></span>

<span data-ttu-id="728b9-187">Если используется расширение модуля, команды надстройки отображаются на пользовательской вкладке расширения.</span><span class="sxs-lookup"><span data-stu-id="728b9-187">When using a module extension, add-in commands appear on the extension's custom tab.</span></span>

## <a name="see-also"></a><span data-ttu-id="728b9-188">См. также</span><span class="sxs-lookup"><span data-stu-id="728b9-188">See also</span></span>

- [<span data-ttu-id="728b9-189">Определение команд надстройки в манифесте</span><span class="sxs-lookup"><span data-stu-id="728b9-189">Define add-in commands in your manifest</span></span>](../develop/create-addin-commands.md)
- [<span data-ttu-id="728b9-190">Надстройка Outlook "Add-in Command Demo"</span><span class="sxs-lookup"><span data-stu-id="728b9-190">Add-in command demo Outlook add-in</span></span>](https://github.com/officedev/outlook-add-in-command-demo)
