---
title: Реализация закрепляемой области задач в надстройке Outlook.
description: Фигура пользовательского интерфейса области задач для команд надстройки открывает вертикальную область задач справа от открытого сообщения или приглашения на собрание, предоставляя интерфейс для дополнительных действий.
ms.date: 10/13/2022
ms.localizationpriority: medium
ms.openlocfilehash: 834d43a6046ddaa63a7c8899cfd5b07d0ea80ef6
ms.sourcegitcommit: a2df9538b3deb32ae3060ecb09da15f5a3d6cb8d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/12/2022
ms.locfileid: "68541124"
---
# <a name="implement-a-pinnable-task-pane-in-outlook"></a>Реализация закрепляемой области задач в Outlook

The [task pane](add-in-commands-for-outlook.md#launch-a-task-pane) UX shape for add-in commands opens a vertical task pane to the right of an open message or meeting request, allowing the add-in to provide UI for more detailed interactions (filling in multiple fields, etc.). This task pane can be shown in the Reading Pane when viewing a list of messages, allowing for quick processing of a message.

However, by default, if a user has an add-in task pane open for a message in the Reading Pane, and then selects a new message, the task pane is automatically closed. For a heavily-used add-in, the user may prefer to keep that pane open, eliminating the need to reactivate the add-in on each message. With pinnable task panes, your add-in can give the user that option.

> [!NOTE]
> Хотя функция закрепленных областей задач была представлена в наборе обязательных элементов [1.5](/javascript/api/requirement-sets/outlook/requirement-set-1.5/outlook-requirement-set-1.5), в настоящее время она доступна только подписчикам Microsoft 365 с помощью следующих средств:
>
> - Outlook версии не ниже 2016 для Windows (сборки начиная с 7668.2000 для пользователей актуального канала и канала программы предварительной оценки Office; сборки начиная с 7900.xxxx для пользователей отложенных каналов)
> - Outlook версии не ниже 2016 для Mac (версия не ниже 16.13.503)
> - Современная версия Outlook в Интернете

> [!IMPORTANT]
> Закрепленные области задач недоступны для следующих элементов:
>
> - Встречи и собрания
> - Outlook.com

> [!TIP]
> Если вы планируете опубликовать [](../publish/publish.md) надстройку Outlook в [AppSource](https://appsource.microsoft.com) и она настроена для закрепленной области задач, для проверки [AppSource](/legal/marketplace/certification-policies) содержимое надстройки не должно быть статическим и должно четко отображать данные, связанные с сообщением, которое открыто или выбрано в почтовом ящике.

## <a name="support-task-pane-pinning"></a>Поддержка закрепления области задач

Для начала нужно добавить поддержку закрепления в манифест надстройки. Разметка зависит от типа манифеста.

# <a name="xml-manifest"></a>[XML-манифест](#tab/xmlmanifest)

Добавьте элемент [SupportsPinning](/javascript/api/manifest/action#supportspinning) в элемент **\<Action\>** , описывающий кнопку области задач. Ниже приведен пример.

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

Элемент **\<SupportsPinning\>** определен в схеме VersionOverrides версии 1.1, поэтому необходимо включить элемент [VersionOverrides](/javascript/api/manifest/versionoverrides) как для версий 1.0, так и для версии 1.1.

# <a name="teams-manifest-developer-preview"></a>[Манифест Teams (предварительная версия для разработчиков)](#tab/jsonmanifest)

Добавьте свойство "pinnable" в объект в массиве actions, `true`который определяет кнопку или пункт меню, открывающее область задач. Ниже приведен пример.

```json
"actions": [
    {
        "id": "OpenTaskPane",
        "type": "openPage",
        "view": "TaskPaneView",
        "displayName": "OpenTaskPane",
        "pinnable": true
    }
]
```

---

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

Use the [Office.context.mailbox.addHandlerAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) method to register your event handler for the `Office.EventType.ItemChanged` event. This should be done in the `Office.initialize` function for your task pane.

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
