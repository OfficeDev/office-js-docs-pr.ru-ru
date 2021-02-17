---
title: Создание мобильной надстройки Outlook для поставщика собраний по сети
description: В этой теме обсуждается настройка мобильной надстройки Outlook для поставщика услуг собраний по сети.
ms.topic: article
ms.date: 02/12/2021
localization_priority: Normal
ms.openlocfilehash: fb98ddeeef8615476659a0abb798ea7901d81248
ms.sourcegitcommit: 1cdf5728102424a46998e1527508b4e7f9f74a4c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/17/2021
ms.locfileid: "50270744"
---
# <a name="create-an-outlook-mobile-add-in-for-an-online-meeting-provider"></a>Создание мобильной надстройки Outlook для поставщика собраний по сети

Настройка собрания по сети — это основная задача пользователя Outlook, и ее легко создать с помощью [Outlook](/microsoftteams/teams-add-in-for-outlook) Mobile. Однако создание собрания по сети в Outlook с помощью службы, не относякой к Майкрософт, может быть очень важным. Реализуя эту функцию, поставщики услуг могут упростить создание собраний по сети для пользователей надстройки Outlook.

> [!IMPORTANT]
> Эта функция поддерживается только на Android и iOS с подпиской на Microsoft 365.

В этой статье вы узнаете, как настроить мобильную надстройка Outlook, чтобы пользователи могли организовывать собрания и присоединяться к ним с помощью службы собраний по сети. В этой статье мы будем использовать вымышленного поставщика услуг онлайн-собраний Contoso.

## <a name="set-up-your-environment"></a>Настройка среды

Завершите [краткое начало работы с Outlook,](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) которое создает проект надстройки с помощью генератора Yeoman для надстройки Office.

## <a name="configure-the-manifest"></a>Настройка манифеста

Чтобы позволить пользователям создавать собрания по сети с помощью надстройки, необходимо настроить точку расширения [MobileOnlineMeetingCommandSurface](../reference/manifest/extensionpoint.md#mobileonlinemeetingcommandsurface) в манифесте в родительском `MobileFormFactor` элементе. Другие форм-факторы не поддерживаются.

1. В редакторе кода откройте проект быстрого запуска.

1. Откройте файл **manifest.xml,** расположенный в корневой папке проекта.

1. Выберите весь узел (включая открытые и закрываемые `<VersionOverrides>` теги) и замените его на следующий XML-

```xml
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    <Description resid="residDescription"></Description>
    <Requirements>
      <bt:Sets>
        <bt:Set Name="Mailbox" MinVersion="1.3"/>
      </bt:Sets>
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <FunctionFile resid="residFunctionFile"/>
          <ExtensionPoint xsi:type="AppointmentOrganizerCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="apptComposeGroup">
                <Label resid="residDescription"/>
                <Control xsi:type="Button" id="insertMeetingButton">
                  <Label resid="residLabel"/>
                  <Supertip>
                    <Title resid="residLabel"/>
                    <Description resid="residTooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="icon-16"/>
                    <bt:Image size="32" resid="icon-32"/>
                    <bt:Image size="64" resid="icon-64"/>
                    <bt:Image size="80" resid="icon-80"/>
                  </Icon>
                  <Action xsi:type="ExecuteFunction">
                    <FunctionName>insertContosoMeeting</FunctionName>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>

        <MobileFormFactor>
          <FunctionFile resid="residFunctionFile"/>
          <ExtensionPoint xsi:type="MobileOnlineMeetingCommandSurface">
            <Control xsi:type="MobileButton" id="insertMeetingButton">
              <Label resid="residLabel"/>
              <Icon>
                <bt:Image size="25" scale="1" resid="icon-16"/>
                <bt:Image size="25" scale="2" resid="icon-16"/>
                <bt:Image size="25" scale="3" resid="icon-16"/>

                <bt:Image size="32" scale="1" resid="icon-32"/>
                <bt:Image size="32" scale="2" resid="icon-32"/>
                <bt:Image size="32" scale="3" resid="icon-32"/>

                <bt:Image size="48" scale="1" resid="icon-48"/>
                <bt:Image size="48" scale="2" resid="icon-48"/>
                <bt:Image size="48" scale="3" resid="icon-48"/>
              </Icon>
              <Action xsi:type="ExecuteFunction">
                <FunctionName>insertContosoMeeting</FunctionName>
              </Action>
            </Control>
          </ExtensionPoint>
        </MobileFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="icon-16" DefaultValue="https://contoso.com/assets/icon-16.png"/>
        <bt:Image id="icon-32" DefaultValue="https://contoso.com/assets/icon-32.png"/>
        <bt:Image id="icon-48" DefaultValue="https://contoso.com/assets/icon-48.png"/>
        <bt:Image id="icon-64" DefaultValue="https://contoso.com/assets/icon-64.png"/>
        <bt:Image id="icon-80" DefaultValue="https://contoso.com/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="residFunctionFile" DefaultValue="https://contoso.com/commands.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="residDescription" DefaultValue="Contoso meeting"/>
        <bt:String id="residLabel" DefaultValue="Add a contoso meeting"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="residTooltip" DefaultValue="Add a contoso meeting to this appointment."/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</VersionOverrides>
```

> [!TIP]
> Дополнительные информацию о манифестах надстройки Outlook см. в манифестах надстройки [Outlook](manifests.md) и добавлении поддержки команд надстройки [для Outlook Mobile.](add-mobile-support.md)

## <a name="implement-adding-online-meeting-details"></a>Реализация добавления сведений о собрании по сети

В этом разделе вы узнаете, как скрипт надстройки может обновить собрание пользователя, включив сведения о собрании по сети.

1. В том же проекте быстрого запуска откройте файл **./src/commands/commands.js** в редакторе кода.

1. Замените все содержимое файла **commands.js** следующим javaScript.

    ```js
    // 1. How to construct online meeting details.
    // Not shown: How to get the meeting organizer's ID and other details from your service.
    const newBody = '<br>' +
        '<a href="https://contoso.com/meeting?id=123456789" target="_blank">Join Contoso meeting</a>' +
        '<br><br>' +
        'Phone Dial-in: +1(123)456-7890' +
        '<br><br>' +
        'Meeting ID: 123 456 789' +
        '<br><br>' +
        'Want to test your video connection?' +
        '<br><br>' +
        '<a href="https://contoso.com/testmeeting" target="_blank">Join test meeting</a>' +
        '<br><br>';

    var mailboxItem;

    // Office is ready.
    Office.onReady(function () {
            mailboxItem = Office.context.mailbox.item;
        }
    );

    // 2. How to define a UI-less function named `insertContosoMeeting` (referenced in the manifest)
    //    to update the meeting body with the online meeting details.
    function insertContosoMeeting(event) {
        // Get HTML body from the client.
        mailboxItem.body.getAsync("html",
            { asyncContext: event },
            function (getBodyResult) {
                if (getBodyResult.status === Office.AsyncResultStatus.Succeeded) {
                    updateBody(getBodyResult.asyncContext, getBodyResult.value);
                } else {
                    console.error("Failed to get HTML body.");
                    getBodyResult.asyncContext.completed({ allowEvent: false });
                }
            }
        );
    }

    // 3. How to implement a supporting function `updateBody`
    //    that appends the online meeting details to the current body of the meeting.
    function updateBody(event, existingBody) {
        // Append new body to the existing body.
        mailboxItem.body.setAsync(existingBody + newBody,
            { asyncContext: event, coercionType: "html" },
            function (setBodyResult) {
                if (setBodyResult.status === Office.AsyncResultStatus.Succeeded) {
                    setBodyResult.asyncContext.completed({ allowEvent: true });
                } else {
                    console.error("Failed to set HTML body.");
                    setBodyResult.asyncContext.completed({ allowEvent: false });
                }
            }
        );
    }

    function getGlobal() {
      return typeof self !== "undefined"
        ? self
        : typeof window !== "undefined"
        ? window
        : typeof global !== "undefined"
        ? global
        : undefined;
    }

    const g = getGlobal();

    // The add-in command functions need to be available in global scope.
    g.insertContosoMeeting = insertContosoMeeting;
    ```

## <a name="testing-and-validation"></a>Тестирование и проверка

Следуйте обычным рекомендациям [по проверке и проверке надстройки.](testing-and-tips.md) После [загрузки](sideload-outlook-add-ins-for-testing.md) неогрузки в Outlook в Интернете, Windows или Mac перезапустите Outlook на мобильном устройстве с Android. (На данный момент единственным поддерживаемым клиентом является Android.) Затем на новом экране собрания убедитесь, что толль Microsoft Teams или Skype заменен вашим.

### <a name="create-meeting-ui"></a>Создание пользовательского интерфейса собрания

В качестве организатора собрания при создании собрания должны появиться экраны, аналогичные следующим трем изображениям.

[ ![ Screenshot of create meeting screen on Android - Contoso toggle off](../images/outlook-android-create-online-meeting-off.png)](../images/outlook-android-create-online-meeting-off-expanded.png#lightbox) [ ![ screenshot of create meeting screen on Android - loading Contoso toggle](../images/outlook-android-create-online-meeting-load.png)](../images/outlook-android-create-online-meeting-load-expanded.png#lightbox) [ ![ screenshot of create meeting screen on Android - Contoso toggle on](../images/outlook-android-create-online-meeting-on.png)](../images/outlook-android-create-online-meeting-on-expanded.png#lightbox)

### <a name="join-meeting-ui"></a>Присоединяйтесь к пользовательскому интерфейсу собрания

В качестве участника собрания при просмотре собрания должен отобраться экран, подобный следующему.

[![снимок экрана присоединиться к собранию на Android](../images/outlook-android-join-online-meeting-view-1.png)](../images/outlook-android-join-online-meeting-view-1-expanded.png#lightbox)

> [!IMPORTANT]
> Если вы не видите  ссылку "Присоединиться", возможно, шаблон собрания по сети для вашей службы не зарегистрирован на наших серверах. Подробные сведения см. в разделе "Регистрация [шаблона собрания](#register-your-online-meeting-template) по сети".

## <a name="register-your-online-meeting-template"></a>Регистрация шаблона собрания по сети

Если вы хотите зарегистрировать шаблон собрания по сети для своей службы, вы можете создать проблему с GitHub с подробными сведениями. После этого мы свяемся с вами, чтобы скоординировать временную шкалу регистрации.

1. Перейдите в раздел **"Отзывы"** в конце этой статьи.
1. Нажмите **ссылку "Эта страница".**
1. **Задайте для новой** проблемы заголовок "Регистрация шаблона собрания по сети для моей службы", заменив ее `my-service` именем службы.
1. В тексте проблемы замените строку "[Введите здесь отзыв]" строкой, заданной в переменной или аналогичной из раздела "Реализация добавления сведений о собрании по сети" ранее `newBody` в этой статье. [](#implement-adding-online-meeting-details)
1. Нажмите **кнопку "Отправить новую проблему"**.

![снимок экрана с новым экраном проблемы GitHub с образцом контента Contoso](../images/outlook-request-to-register-online-meeting-template.png)

## <a name="available-apis"></a>Доступные API

Для этой функции доступны следующие API.

- API организатора встреч
  - [Office.context.mailbox.item.subject](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#subject) ([Subject)](/javascript/api/outlook/office.subject?view=outlook-js-preview&preserve-view=true)
  - [Office.context.mailbox.item.start](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#start) ([Time)](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true)
  - [Office.context.mailbox.item.end](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#end) ([Time](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.location](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#location) ([Location)](/javascript/api/outlook/office.location?view=outlook-js-preview&preserve-view=true)
  - [Office.context.mailbox.item.optionalAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#optionalattendees) ([Recipients)](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true)
  - [Office.context.mailbox.item.requiredAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#requiredattendees) ([Recipients)](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true)
  - [Office.context.mailbox.item.body](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#body) ([Body.getAsync,](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#getasync-coerciontype--options--callback-) [Body.setAsync)](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#setasync-data--options--callback-)
  - [Office.context.mailbox.item.loadCustomPropertiesAsync](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#loadcustompropertiesasync-callback--usercontext-) ([CustomProperties)](/javascript/api/outlook/office.customproperties?view=outlook-js-preview&preserve-view=true)
  - [Office.context.roamingSettings](../reference/objectmodel/preview-requirement-set/office.context.md?view=outlook-js-preview&preserve-view=true#roamingsettings-roamingsettings) ([RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true))
- Обработка потока auth
  - [API диалоговых окон](../develop/dialog-api-in-office-add-ins.md)

## <a name="restrictions"></a>Ограничения

Применяется несколько ограничений.

- Применимо только к поставщикам услуг собраний по сети.
- На экране составить собрание будут отображаться только установленные администратором надстройки, заменяющие параметр Teams или Skype по умолчанию. Установленные пользователем надстройки не активируются.
- Значок надстройки должен быть в серой области с использованием hex-кода или его эквивалента `#919191` в [других форматах цвета.](https://convertingcolors.com/hex-color-919191.html)
- В режиме организатора встреч (составить) поддерживается только одна команда без пользовательского интерфейса.

## <a name="see-also"></a>См. также

- [Надстройки для Outlook Mobile](outlook-mobile-addins.md)
- [Добавление поддержки команд надстройки для Outlook Mobile](add-mobile-support.md)
