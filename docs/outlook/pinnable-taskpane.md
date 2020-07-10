---
title: Реализация закрепляемой области задач в надстройке Outlook.
description: Фигура пользовательского интерфейса области задач для команд надстройки открывает вертикальную область задач справа от открытого сообщения или приглашения на собрание, предоставляя интерфейс для дополнительных действий.
ms.date: 07/07/2020
localization_priority: Normal
ms.openlocfilehash: 39af3a532d553835b02709301c998a78dc9958bb
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093870"
---
# <a name="implement-a-pinnable-task-pane-in-outlook"></a><span data-ttu-id="eb3a8-103">Реализация закрепляемой области задач в Outlook</span><span class="sxs-lookup"><span data-stu-id="eb3a8-103">Implement a pinnable task pane in Outlook</span></span>

<span data-ttu-id="eb3a8-104">The [task pane](add-in-commands-for-outlook.md#launching-a-task-pane) UX shape for add-in commands opens a vertical task pane to the right of an open message or meeting request, allowing the add-in to provide UI for more detailed interactions (filling in multiple fields, etc.).</span><span class="sxs-lookup"><span data-stu-id="eb3a8-104">The [task pane](add-in-commands-for-outlook.md#launching-a-task-pane) UX shape for add-in commands opens a vertical task pane to the right of an open message or meeting request, allowing the add-in to provide UI for more detailed interactions (filling in multiple fields, etc.).</span></span> <span data-ttu-id="eb3a8-105">This task pane can be shown in the Reading Pane when viewing a list of messages, allowing for quick processing of a message.</span><span class="sxs-lookup"><span data-stu-id="eb3a8-105">This task pane can be shown in the Reading Pane when viewing a list of messages, allowing for quick processing of a message.</span></span>

<span data-ttu-id="eb3a8-106">However, by default, if a user has an add-in task pane open for a message in the Reading Pane, and then selects a new message, the task pane is automatically closed.</span><span class="sxs-lookup"><span data-stu-id="eb3a8-106">However, by default, if a user has an add-in task pane open for a message in the Reading Pane, and then selects a new message, the task pane is automatically closed.</span></span> <span data-ttu-id="eb3a8-107">For a heavily-used add-in, the user may prefer to keep that pane open, eliminating the need to reactivate the add-in on each message.</span><span class="sxs-lookup"><span data-stu-id="eb3a8-107">For a heavily-used add-in, the user may prefer to keep that pane open, eliminating the need to reactivate the add-in on each message.</span></span> <span data-ttu-id="eb3a8-108">With pinnable task panes, your add-in can give the user that option.</span><span class="sxs-lookup"><span data-stu-id="eb3a8-108">With pinnable task panes, your add-in can give the user that option.</span></span>

> [!NOTE]
> <span data-ttu-id="eb3a8-109">Несмотря на то, что функция областей задач закрепляемая была введена в [наборе](../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)обязательных элементов 1,5, в настоящее время она доступна только подписчикам Microsoft 365 с помощью следующих компонентов.</span><span class="sxs-lookup"><span data-stu-id="eb3a8-109">Although the pinnable task panes feature was introduced in [requirement set 1.5](../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md), it's currently only available to Microsoft 365 subscribers using the following.</span></span>
> - <span data-ttu-id="eb3a8-110">Outlook 2016 или более поздняя версия в Windows (сборка 7668,2000 или более поздняя версия для пользователей в текущем канале или программе предварительной оценки Office) создайте 7900. xxxx или более поздней версии для пользователей в отложенных каналах.</span><span class="sxs-lookup"><span data-stu-id="eb3a8-110">Outlook 2016 or later on Windows (build 7668.2000 or later for users in the Current or Office Insider Channels, build 7900.xxxx or later for users in Deferred channels)</span></span>
> - <span data-ttu-id="eb3a8-111">Outlook 2016 или более поздняя версия в Mac (версия 16.13.503 или более поздняя)</span><span class="sxs-lookup"><span data-stu-id="eb3a8-111">Outlook 2016 or later on Mac (version 16.13.503 or later)</span></span>
> - <span data-ttu-id="eb3a8-112">Современная версия Outlook в Интернете</span><span class="sxs-lookup"><span data-stu-id="eb3a8-112">Modern Outlook on the web</span></span>

> [!IMPORTANT]
> <span data-ttu-id="eb3a8-113">Места, где недоступны закрепляемые области задач:</span><span class="sxs-lookup"><span data-stu-id="eb3a8-113">Pinnable task panes are not available for the following.</span></span>
> - <span data-ttu-id="eb3a8-114">Встречи и собрания</span><span class="sxs-lookup"><span data-stu-id="eb3a8-114">Appointments/Meetings</span></span>
> - <span data-ttu-id="eb3a8-115">Outlook.com</span><span class="sxs-lookup"><span data-stu-id="eb3a8-115">Outlook.com</span></span>

## <a name="support-task-pane-pinning"></a><span data-ttu-id="eb3a8-116">Поддержка закрепления области задач</span><span class="sxs-lookup"><span data-stu-id="eb3a8-116">Support task pane pinning</span></span>

<span data-ttu-id="eb3a8-117">The first step is to add pinning support, which is done in the add-in [manifest](manifests.md).</span><span class="sxs-lookup"><span data-stu-id="eb3a8-117">The first step is to add pinning support, which is done in the add-in [manifest](manifests.md).</span></span> <span data-ttu-id="eb3a8-118">This is done by adding the [SupportsPinning](../reference/manifest/action.md#supportspinning) element to the `Action` element that describes the task pane button.</span><span class="sxs-lookup"><span data-stu-id="eb3a8-118">This is done by adding the [SupportsPinning](../reference/manifest/action.md#supportspinning) element to the `Action` element that describes the task pane button.</span></span>

<span data-ttu-id="eb3a8-119">Элемент `SupportsPinning` определяется в схеме VersionOverrides 1.1, поэтому элемент [VersionOverrides](../reference/manifest/versionoverrides.md) необходимо включить как для версии 1.0, так и для версии 1.1.</span><span class="sxs-lookup"><span data-stu-id="eb3a8-119">The `SupportsPinning` element is defined in the VersionOverrides v1.1 schema, so you will need to include a [VersionOverrides](../reference/manifest/versionoverrides.md) element both for v1.0 and v1.1.</span></span>

> [!NOTE]
> <span data-ttu-id="eb3a8-120">Если вы планируете [публиковать](../publish/publish.md) надстройку Outlook в [AppSource](https://appsource.microsoft.com) и используете элемент **SupportsPinning** для прохождения [проверки AppSource](/legal/marketplace/certification-policies), контент надстройки не должен быть статическим. Необходимо, чтобы он четко отображал данные, которые относятся к сообщению, открытому или выбранному в почтовом ящике.</span><span class="sxs-lookup"><span data-stu-id="eb3a8-120">If you plan to [publish](../publish/publish.md) your Outlook add-in to [AppSource](https://appsource.microsoft.com), when you use the **SupportsPinning** element, in order to pass [AppSource validation](/legal/marketplace/certification-policies), your add-in content must not be static and it must clearly display data related to the message that is open or selected in the mailbox.</span></span>

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

<span data-ttu-id="eb3a8-121">Полный пример: элемент управления `msgReadOpenPaneButton` в [примере манифеста command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo/blob/master/command-demo-manifest.xml).</span><span class="sxs-lookup"><span data-stu-id="eb3a8-121">For a full example, see the `msgReadOpenPaneButton` control in the [command-demo sample manifest](https://github.com/OfficeDev/outlook-add-in-command-demo/blob/master/command-demo-manifest.xml).</span></span>

## <a name="handling-ui-updates-based-on-currently-selected-message"></a><span data-ttu-id="eb3a8-122">Обновление пользовательского интерфейса на основе выбранного сообщения</span><span class="sxs-lookup"><span data-stu-id="eb3a8-122">Handling UI updates based on currently selected message</span></span>

<span data-ttu-id="eb3a8-123">Чтобы обновлять пользовательский интерфейс или внутренние переменные области задач на основе текущего элемента, необходимо зарегистрировать обработчик событий, чтобы получать уведомления об изменении.</span><span class="sxs-lookup"><span data-stu-id="eb3a8-123">To update your task pane's UI or internal variables based on the current item, you'll need to register an event handler to get notified of the change.</span></span>

### <a name="implement-the-event-handler"></a><span data-ttu-id="eb3a8-124">Реализация обработчика событий</span><span class="sxs-lookup"><span data-stu-id="eb3a8-124">Implement the event handler</span></span>

<span data-ttu-id="eb3a8-125">The event handler should accept a single parameter, which is an object literal.</span><span class="sxs-lookup"><span data-stu-id="eb3a8-125">The event handler should accept a single parameter, which is an object literal.</span></span> <span data-ttu-id="eb3a8-126">The `type` property of this object will be set to `Office.EventType.ItemChanged`.</span><span class="sxs-lookup"><span data-stu-id="eb3a8-126">The `type` property of this object will be set to `Office.EventType.ItemChanged`.</span></span> <span data-ttu-id="eb3a8-127">When the event is called, the `Office.context.mailbox.item` object is already updated to reflect the currently selected item.</span><span class="sxs-lookup"><span data-stu-id="eb3a8-127">When the event is called, the `Office.context.mailbox.item` object is already updated to reflect the currently selected item.</span></span>

```js
function itemChanged(eventArgs) {
  // Update UI based on the new current item
  UpdateTaskPaneUI(Office.context.mailbox.item);
}
```

> [!IMPORTANT]
> <span data-ttu-id="eb3a8-128">При реализации обработчиков событий для события ItemChanged необходимо проверять, задано ли для элемента Office.content.mailbox.item значение NULL.</span><span class="sxs-lookup"><span data-stu-id="eb3a8-128">The implementation of event handlers for an ItemChanged event should check whether or not the Office.content.mailbox.item is null.</span></span>
>
> ```js
> // Example implementation
> function UpdateTaskPaneUI(item)
> {
>   // Assuming that item is always a read item (instead of a compose item).
>   if (item != null) console.log(item.subject);
> }
> ```

### <a name="register-the-event-handler"></a><span data-ttu-id="eb3a8-129">Регистрация обработчика событий</span><span class="sxs-lookup"><span data-stu-id="eb3a8-129">Register the event handler</span></span>

<span data-ttu-id="eb3a8-130">Use the [Office.context.mailbox.addHandlerAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method to register your event handler for the `Office.EventType.ItemChanged` event.</span><span class="sxs-lookup"><span data-stu-id="eb3a8-130">Use the [Office.context.mailbox.addHandlerAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method to register your event handler for the `Office.EventType.ItemChanged` event.</span></span> <span data-ttu-id="eb3a8-131">This should be done in the `Office.initialize` function for your task pane.</span><span class="sxs-lookup"><span data-stu-id="eb3a8-131">This should be done in the `Office.initialize` function for your task pane.</span></span>

```js
Office.initialize = function (reason) {
  $(document).ready(function () {

    // Set up ItemChanged event
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, itemChanged);

    UpdateTaskPaneUI(Office.context.mailbox.item);
  });
};
```

## <a name="see-also"></a><span data-ttu-id="eb3a8-132">См. также</span><span class="sxs-lookup"><span data-stu-id="eb3a8-132">See also</span></span>

<span data-ttu-id="eb3a8-133">Пример надстройки, в которой реализована закрепляемая область задач, на сайте GitHub: [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo).</span><span class="sxs-lookup"><span data-stu-id="eb3a8-133">For an example add-in that implements a pinnable task pane, see [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo) on GitHub.</span></span>
