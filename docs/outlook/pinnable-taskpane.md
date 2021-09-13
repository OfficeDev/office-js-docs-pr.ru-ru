---
title: Реализация закрепляемой области задач в надстройке Outlook.
description: Фигура пользовательского интерфейса области задач для команд надстройки открывает вертикальную область задач справа от открытого сообщения или приглашения на собрание, предоставляя интерфейс для дополнительных действий.
ms.date: 07/07/2020
ms.localizationpriority: medium
ms.openlocfilehash: e418ba10fa5c0b35406b5b105fd1e97599323bc1
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59151395"
---
# <a name="implement-a-pinnable-task-pane-in-outlook"></a>Реализация закрепляемой области задач в Outlook

Фигура пользовательского интерфейса [области задач](add-in-commands-for-outlook.md#launching-a-task-pane) для команд надстройки открывает вертикальную область задач справа от открытого сообщения или приглашения на собрание, предоставляя интерфейс для дополнительных действий (заполнение нескольких полей и т. д.). Эта область задач может отображаться в области чтения при просмотре списка сообщений для быстрой обработки сообщения.

Но по умолчанию, если пользователь выбирает новое сообщение, область задач надстройки для сообщения в области чтения автоматически закрывается. Если надстройка используется часто, пользователь может закрепить эту область, чтобы не активировать ее повторно для каждого сообщения. Для этого необходимо добавить в надстройку закрепляемые области задач.

> [!NOTE]
> Несмотря на то, что функция pinnable task panes была представлена в наборе [требований 1.5,](../reference/objectmodel/requirement-set-1.5/outlook-requirement-set-1.5.md)в настоящее время она доступна только для Microsoft 365 абонентов с помощью следующих ниже:
>
> - Outlook 2016 или более поздней Windows (сборка 7668.2000 или более поздней части для пользователей в каналах Current или Office Insider, сборка 7900.xxxx или более поздней части для пользователей в отложенных каналах)
> - Outlook 2016 или позже на Mac (версия 16.13.503 или более поздней версии)
> - Современная версия Outlook в Интернете

> [!IMPORTANT]
> Области задач pinnable недоступны для следующих задач:
>
> - Встречи и собрания
> - Outlook.com

## <a name="support-task-pane-pinning"></a>Поддержка закрепления области задач

Для начала нужно добавить поддержку закрепления в [манифест](manifests.md) надстройки. Для этого добавьте элемент [SupportsPinning](../reference/manifest/action.md#supportspinning) в элемент `Action`, который описывает кнопку области задач.

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

Обработчик событий должен принимать один параметр, а именно — объектный литерал. Для свойства `type` этого объекта будет установлено значение `Office.EventType.ItemChanged`. При вызове события объект `Office.context.mailbox.item` уже обновлен с учетом выбранного элемента.

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

Используйте метод [Office.context.mailbox.addHandlerAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods), чтобы зарегистрировать обработчик событий для события `Office.EventType.ItemChanged`. Это следует сделать в функции `Office.initialize` для области задач.

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
