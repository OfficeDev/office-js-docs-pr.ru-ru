---
title: Реализация закрепляемой области задач в надстройке Outlook.
description: Фигура пользовательского интерфейса области задач для команд надстройки открывает вертикальную область задач справа от открытого сообщения или приглашения на собрание, предоставляя интерфейс для дополнительных действий.
ms.date: 11/18/2019
localization_priority: Normal
ms.openlocfilehash: 94c136a74dfddac1af663aea06c3c6ca27f22dcd
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166651"
---
# <a name="implement-a-pinnable-task-pane-in-outlook"></a><span data-ttu-id="c8f85-103">Реализация закрепляемой области задач в Outlook</span><span class="sxs-lookup"><span data-stu-id="c8f85-103">Implement a pinnable task pane in Outlook</span></span>

<span data-ttu-id="c8f85-p101">Фигура пользовательского интерфейса [области задач](add-in-commands-for-outlook.md#launching-a-task-pane) для команд надстройки открывает вертикальную область задач справа от открытого сообщения или приглашения на собрание, предоставляя интерфейс для дополнительных действий (заполнение нескольких полей и т. д.). Эта область задач может отображаться в области чтения при просмотре списка сообщений для быстрой обработки сообщения.</span><span class="sxs-lookup"><span data-stu-id="c8f85-p101">The [task pane](add-in-commands-for-outlook.md#launching-a-task-pane) UX shape for add-in commands opens a vertical task pane to the right of an open message or meeting request, allowing the add-in to provide UI for more detailed interactions (filling in multiple fields, etc.). This task pane can be shown in the Reading Pane when viewing a list of messages, allowing for quick processing of a message.</span></span>

<span data-ttu-id="c8f85-p102">Но по умолчанию, если пользователь выбирает новое сообщение, область задач надстройки для сообщения в области чтения автоматически закрывается. Если надстройка используется часто, пользователь может закрепить эту область, чтобы не активировать ее повторно для каждого сообщения. Для этого необходимо добавить в надстройку закрепляемые области задач.</span><span class="sxs-lookup"><span data-stu-id="c8f85-p102">However, by default, if a user has an add-in task pane open for a message in the Reading Pane, and then selects a new message, the task pane is automatically closed. For a heavily-used add-in, the user may prefer to keep that pane open, eliminating the need to reactivate the add-in on each message. With pinnable task panes, your add-in can give the user that option.</span></span>

> [!NOTE]
> <span data-ttu-id="c8f85-109">В настоящее время закрепляемые области задач доступны для подписчиков Office 365, использующих Outlook 2016 или более позднюю версию для Windows (сборка 7668.2000 или более поздняя для пользователей быстрого канала или канала предварительной оценки Office; сборка 7900.xxxx или более поздняя для пользователей отложенных каналов), Outlook 2016 или более позднюю версию для Mac (версия 16.13.503 или более поздняя), а также Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="c8f85-109">Pinnable task panes are currently available to Office 365 subscribers using Outlook 2016 or later on Windows (build 7668.2000 or later for users in the Current or Office Insider Channels, build 7900.xxxx or later for users in Deferred channels), Outlook 2016 or later on Mac (version 16.13.503 or later), and Outlook on the web.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="c8f85-110">Места, где недоступны закрепляемые области задач:</span><span class="sxs-lookup"><span data-stu-id="c8f85-110">Pinnable task panes are not available for the following.</span></span>
> - <span data-ttu-id="c8f85-111">Встречи и собрания</span><span class="sxs-lookup"><span data-stu-id="c8f85-111">Appointments/Meetings</span></span>
> - <span data-ttu-id="c8f85-112">Outlook.com</span><span class="sxs-lookup"><span data-stu-id="c8f85-112">Outlook.com</span></span>

## <a name="support-task-pane-pinning"></a><span data-ttu-id="c8f85-113">Поддержка закрепления области задач</span><span class="sxs-lookup"><span data-stu-id="c8f85-113">Support task pane pinning</span></span>

<span data-ttu-id="c8f85-p103">Для начала нужно добавить поддержку закрепления в [манифест](manifests.md) надстройки. Для этого добавьте элемент [SupportsPinning](../reference/manifest/action.md#supportspinning) в элемент `Action`, который описывает кнопку области задач.</span><span class="sxs-lookup"><span data-stu-id="c8f85-p103">The first step is to add pinning support, which is done in the add-in [manifest](manifests.md). This is done by adding the [SupportsPinning](../reference/manifest/action.md#supportspinning) element to the `Action` element that describes the task pane button.</span></span>

<span data-ttu-id="c8f85-116">Элемент `SupportsPinning` определяется в схеме VersionOverrides 1.1, поэтому элемент [VersionOverrides](../reference/manifest/versionoverrides.md) необходимо включить как для версии 1.0, так и для версии 1.1.</span><span class="sxs-lookup"><span data-stu-id="c8f85-116">The `SupportsPinning` element is defined in the VersionOverrides v1.1 schema, so you will need to include a [VersionOverrides](../reference/manifest/versionoverrides.md) element both for v1.0 and v1.1.</span></span>

> [!NOTE]
> <span data-ttu-id="c8f85-117">Если вы планируете [публиковать](../publish/publish.md) надстройку Outlook в [AppSource](https://appsource.microsoft.com) и используете элемент **SupportsPinning** для прохождения [проверки AppSource](/office/dev/store/validation-policies), контент надстройки не должен быть статическим. Необходимо, чтобы он четко отображал данные, которые относятся к сообщению, открытому или выбранному в почтовом ящике.</span><span class="sxs-lookup"><span data-stu-id="c8f85-117">If you plan to [publish](../publish/publish.md) your Outlook add-in to [AppSource](https://appsource.microsoft.com), when you use the **SupportsPinning** element, in order to pass [AppSource validation](/office/dev/store/validation-policies), your add-in content must not be static and it must clearly display data related to the message that is open or selected in the mailbox.</span></span>

```xml
<!-- Task pane button -->
<Control xsi:type="Button" id="msgReadOpenPaneButton">
  <Label resid="paneReadButtonLabel" />
  <Supertip>
    <Title resid="paneReadSuperTipTitle" />
    <Description resid="paneReadSuperTipDescription" />
  </Supertip>
  <Icon>
    <bt:Image size="16" resid="green-icon-16" />
    <bt:Image size="32" resid="green-icon-32" />
    <bt:Image size="80" resid="green-icon-80" />
  </Icon>
  <Action xsi:type="ShowTaskpane">
    <SourceLocation resid="readTaskPaneUrl" />
    <SupportsPinning>true</SupportsPinning>
  </Action>
</Control>
```

<span data-ttu-id="c8f85-118">Полный пример: элемент управления `msgReadOpenPaneButton` в [примере манифеста command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo/blob/master/command-demo-manifest.xml).</span><span class="sxs-lookup"><span data-stu-id="c8f85-118">For a full example, see the `msgReadOpenPaneButton` control in the [command-demo sample manifest](https://github.com/OfficeDev/outlook-add-in-command-demo/blob/master/command-demo-manifest.xml).</span></span>

## <a name="handling-ui-updates-based-on-currently-selected-message"></a><span data-ttu-id="c8f85-119">Обновление пользовательского интерфейса на основе выбранного сообщения</span><span class="sxs-lookup"><span data-stu-id="c8f85-119">Handling UI updates based on currently selected message</span></span>

<span data-ttu-id="c8f85-120">Чтобы обновлять пользовательский интерфейс или внутренние переменные области задач на основе текущего элемента, необходимо зарегистрировать обработчик событий, чтобы получать уведомления об изменении.</span><span class="sxs-lookup"><span data-stu-id="c8f85-120">To update your task pane's UI or internal variables based on the current item, you'll need to register an event handler to get notified of the change.</span></span>

### <a name="implement-the-event-handler"></a><span data-ttu-id="c8f85-121">Реализация обработчика событий</span><span class="sxs-lookup"><span data-stu-id="c8f85-121">Implement the event handler</span></span>

<span data-ttu-id="c8f85-p104">Обработчик событий должен принимать один параметр, а именно — объектный литерал. Для свойства `type` этого объекта будет установлено значение `Office.EventType.ItemChanged`. При вызове события объект `Office.context.mailbox.item` уже обновлен с учетом выбранного элемента.</span><span class="sxs-lookup"><span data-stu-id="c8f85-p104">The event handler should accept a single parameter, which is an object literal. The `type` property of this object will be set to `Office.EventType.ItemChanged`. When the event is called, the `Office.context.mailbox.item` object is already updated to reflect the currently selected item.</span></span>

```js
function itemChanged(eventArgs) {
  // Update UI based on the new current item
  UpdateTaskPaneUI(Office.context.mailbox.item);
}
```

> [!IMPORTANT]
> <span data-ttu-id="c8f85-125">При реализации обработчиков событий для события ItemChanged необходимо проверять, задано ли для элемента Office.content.mailbox.item значение NULL.</span><span class="sxs-lookup"><span data-stu-id="c8f85-125">The implementation of event handlers for an ItemChanged event should check whether or not the Office.content.mailbox.item is null.</span></span>
>
> ```js
> // Example implementation
> function UpdateTaskPaneUI(item)
> {
>   // Assuming that item is always a read item (instead of a compose item).
>   if (item != null) console.log(item.subject);
> }
> ```

### <a name="register-the-event-handler"></a><span data-ttu-id="c8f85-126">Регистрация обработчика событий</span><span class="sxs-lookup"><span data-stu-id="c8f85-126">Register the event handler</span></span>

<span data-ttu-id="c8f85-p105">Используйте метод [Office.context.mailbox.addHandlerAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods), чтобы зарегистрировать обработчик событий для события `Office.EventType.ItemChanged`. Это следует сделать в функции `Office.initialize` для области задач.</span><span class="sxs-lookup"><span data-stu-id="c8f85-p105">Use the [Office.context.mailbox.addHandlerAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method to register your event handler for the `Office.EventType.ItemChanged` event. This should be done in the `Office.initialize` function for your task pane.</span></span>

```js
Office.initialize = function (reason) {
  $(document).ready(function () {

    // Set up ItemChanged event
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, itemChanged);

    UpdateTaskPaneUI(Office.context.mailbox.item);
  });
};
```

## <a name="see-also"></a><span data-ttu-id="c8f85-129">См. также</span><span class="sxs-lookup"><span data-stu-id="c8f85-129">See also</span></span>

<span data-ttu-id="c8f85-130">Пример надстройки, в которой реализована закрепляемая область задач, на сайте GitHub: [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo).</span><span class="sxs-lookup"><span data-stu-id="c8f85-130">For an example add-in that implements a pinnable task pane, see [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo) on GitHub.</span></span>
