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
# <a name="implement-a-pinnable-task-pane-in-outlook"></a>Реализация закрепляемой области задач в Outlook

The [task pane](add-in-commands-for-outlook.md#launching-a-task-pane) UX shape for add-in commands opens a vertical task pane to the right of an open message or meeting request, allowing the add-in to provide UI for more detailed interactions (filling in multiple fields, etc.). This task pane can be shown in the Reading Pane when viewing a list of messages, allowing for quick processing of a message.

However, by default, if a user has an add-in task pane open for a message in the Reading Pane, and then selects a new message, the task pane is automatically closed. For a heavily-used add-in, the user may prefer to keep that pane open, eliminating the need to reactivate the add-in on each message. With pinnable task panes, your add-in can give the user that option.

> [!NOTE]
> Несмотря на то, что функция областей задач закрепляемая была введена в [наборе](../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)обязательных элементов 1,5, в настоящее время она доступна только подписчикам Microsoft 365 с помощью следующих компонентов.
> - Outlook 2016 или более поздняя версия в Windows (сборка 7668,2000 или более поздняя версия для пользователей в текущем канале или программе предварительной оценки Office) создайте 7900. xxxx или более поздней версии для пользователей в отложенных каналах.
> - Outlook 2016 или более поздняя версия в Mac (версия 16.13.503 или более поздняя)
> - Современная версия Outlook в Интернете

> [!IMPORTANT]
> Места, где недоступны закрепляемые области задач:
> - Встречи и собрания
> - Outlook.com

## <a name="support-task-pane-pinning"></a>Поддержка закрепления области задач

The first step is to add pinning support, which is done in the add-in [manifest](manifests.md). This is done by adding the [SupportsPinning](../reference/manifest/action.md#supportspinning) element to the `Action` element that describes the task pane button.

Элемент `SupportsPinning` определяется в схеме VersionOverrides 1.1, поэтому элемент [VersionOverrides](../reference/manifest/versionoverrides.md) необходимо включить как для версии 1.0, так и для версии 1.1.

> [!NOTE]
> Если вы планируете [публиковать](../publish/publish.md) надстройку Outlook в [AppSource](https://appsource.microsoft.com) и используете элемент **SupportsPinning** для прохождения [проверки AppSource](/legal/marketplace/certification-policies), контент надстройки не должен быть статическим. Необходимо, чтобы он четко отображал данные, которые относятся к сообщению, открытому или выбранному в почтовом ящике.

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

Полный пример: элемент управления `msgReadOpenPaneButton` в [примере манифеста command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo/blob/master/command-demo-manifest.xml).

## <a name="handling-ui-updates-based-on-currently-selected-message"></a>Обновление пользовательского интерфейса на основе выбранного сообщения

Чтобы обновлять пользовательский интерфейс или внутренние переменные области задач на основе текущего элемента, необходимо зарегистрировать обработчик событий, чтобы получать уведомления об изменении.

### <a name="implement-the-event-handler"></a>Реализация обработчика событий

The event handler should accept a single parameter, which is an object literal. The `type` property of this object will be set to `Office.EventType.ItemChanged`. When the event is called, the `Office.context.mailbox.item` object is already updated to reflect the currently selected item.

```js
function itemChanged(eventArgs) {
  // Update UI based on the new current item
  UpdateTaskPaneUI(Office.context.mailbox.item);
}
```

> [!IMPORTANT]
> При реализации обработчиков событий для события ItemChanged необходимо проверять, задано ли для элемента Office.content.mailbox.item значение NULL.
>
> ```js
> // Example implementation
> function UpdateTaskPaneUI(item)
> {
>   // Assuming that item is always a read item (instead of a compose item).
>   if (item != null) console.log(item.subject);
> }
> ```

### <a name="register-the-event-handler"></a>Регистрация обработчика событий

Use the [Office.context.mailbox.addHandlerAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method to register your event handler for the `Office.EventType.ItemChanged` event. This should be done in the `Office.initialize` function for your task pane.

```js
Office.initialize = function (reason) {
  $(document).ready(function () {

    // Set up ItemChanged event
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, itemChanged);

    UpdateTaskPaneUI(Office.context.mailbox.item);
  });
};
```

## <a name="see-also"></a>См. также

Пример надстройки, в которой реализована закрепляемая область задач, на сайте GitHub: [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo).
