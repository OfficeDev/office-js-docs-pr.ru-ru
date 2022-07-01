---
title: Создание надстройки Outlook для поставщика онлайн-собраний
description: Описывается, как настроить надстройку Outlook для поставщика услуг для собраний по сети.
ms.topic: article
ms.date: 06/28/2022
ms.localizationpriority: medium
ms.openlocfilehash: 884e27b75f3fc44a645021f8211d7aaf748f3a1d
ms.sourcegitcommit: e8ce48605f7f33bc5c9af8bfd75d54d4b6b15039
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/01/2022
ms.locfileid: "66574427"
---
# <a name="create-an-outlook-add-in-for-an-online-meeting-provider"></a>Создание надстройки Outlook для поставщика онлайн-собраний

Настройка собрания по сети является основным интерфейсом для пользователя Outlook, и его легко создать в [Outlook](/microsoftteams/teams-add-in-for-outlook). Однако создание собрания по сети в Outlook с помощью службы, отличной от Майкрософт, может быть сложной задачей. Внедрив эту функцию, поставщики услуг могут упростить создание собраний по сети и присоединение пользователей надстроек Outlook.

> [!IMPORTANT]
> Эта функция поддерживается в Outlook в Интернете, Windows, Mac, Android и iOS с подпиской на Microsoft 365.

Из этой статьи вы узнаете, как настроить надстройку Outlook, чтобы пользователи могли упорядочивать собрания и присоединяться к собранию с помощью службы онлайн-собраний. В этой статье мы будем использовать вымышленного поставщика услуг онлайн-собраний Contoso.

## <a name="set-up-your-environment"></a>Настройка среды

Выполните [краткое руководство outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) , которое создает проект надстройки с помощью генератора Yeoman для надстроек Office.

## <a name="configure-the-manifest"></a>Настройка манифеста

Чтобы разрешить пользователям создавать собрания по сети с помощью надстройки, необходимо настроить узел **VersionOverrides** в манифесте. Если вы создаете надстройку, которая будет поддерживаться только в Outlook в Интернете, Windows и Mac, выберите вкладку **Windows, Mac, веб-вкладку** для справки. Однако если ваша надстройка также будет поддерживаться в Outlook для Android и iOS, перейдите на **вкладку "Мобильные** устройства".

# <a name="windows-mac-web"></a>[Windows, Mac, Web](#tab/non-mobile)

1. В редакторе кода откройте созданный проект быстрого запуска Outlook.

1. Откройте файл **manifest.xml** , расположенный в корневом каталоге проекта.

1. Выберите весь узел **VersionOverrides** (включая открытые и закрывающее теги) и замените его следующим XML-кодом.

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

# <a name="mobile"></a>[Мобильные устройства](#tab/mobile)

Чтобы разрешить пользователям создавать собрания по сети с мобильного устройства, точка расширения [MobileOnlineMeetingCommandSurface](/javascript/api/manifest/extensionpoint#mobileonlinemeetingcommandsurface) настраивается в манифесте в родительском элементе **MobileFormFactor**. Эта точка расширения не поддерживается в других форм-факторах.

1. В редакторе кода откройте созданный проект быстрого запуска Outlook.

1. Откройте файл **manifest.xml** , расположенный в корневом каталоге проекта.

1. Выберите весь узел **VersionOverrides** (включая открытые и закрывающее теги) и замените его следующим XML-кодом.

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

---

> [!TIP]
> Дополнительные сведения о манифестах для надстроек Outlook см. в манифестах надстроек [Outlook](manifests.md) и добавлении поддержки команд надстроек [для Outlook Mobile](add-mobile-support.md).

## <a name="implement-adding-online-meeting-details"></a>Реализация добавления сведений о собрании по сети

В этом разделе описано, как скрипт надстройки может обновить собрание пользователя, чтобы включить сведения о собрании по сети. Следующее относится ко всем поддерживаемым платформам.

1. В том же проекте быстрого запуска откройте файл **./src/commands/commands.js** в редакторе кода.

1. Замените все содержимое файла **commands.js** следующим кодом JavaScript.

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

    // 2. How to define and register a UI-less function named `insertContosoMeeting` (referenced in the manifest)
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
    // Register the function.
    Office.actions.associate("insertContosoMeeting", insertContosoMeeting);

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
    ```

## <a name="testing-and-validation"></a>Тестирование и проверка

Следуйте обычным рекомендациям, чтобы протестировать [и](testing-and-tips.md) проверить надстройку, а затем загрузить неопубликованный манифест в Outlook в Интернете, Windows или Mac.[](sideload-outlook-add-ins-for-testing.md) Если ваша надстройка также поддерживает мобильные устройства, перезапустите Outlook на устройстве Android или iOS после загрузки неопубликованных приложений. После загрузки неопубликоваемой надстройки создайте новое собрание и убедитесь, что переключатель Microsoft Teams или Skype заменен вашим собственным.

### <a name="create-meeting-ui"></a>Создание пользовательского интерфейса собрания

Организатор собрания должен видеть экраны, аналогичные следующим трем изображениям при создании собрания.

[![Экран создания собрания на Android с выключенным переключательом Contoso.](../images/outlook-android-create-online-meeting-off.png)](../images/outlook-android-create-online-meeting-off-expanded.png#lightbox) [![Экран создания собрания на Android с переключательом загрузки Contoso.](../images/outlook-android-create-online-meeting-load.png)](../images/outlook-android-create-online-meeting-load-expanded.png#lightbox) [![Экран создания собрания на Android с включенной кнопкой "Contoso".](../images/outlook-android-create-online-meeting-on.png)](../images/outlook-android-create-online-meeting-on-expanded.png#lightbox)

### <a name="join-meeting-ui"></a>Присоединение к пользовательскому интерфейсу собрания

Как участник собрания вы должны увидеть экран, аналогичный следующему изображению, при просмотре собрания.

[![Экран присоединения к собранию на Android.](../images/outlook-android-join-online-meeting-view-1.png)](../images/outlook-android-join-online-meeting-view-1-expanded.png#lightbox)

> [!IMPORTANT]
> **Кнопка "** Присоединиться" поддерживается только в Outlook в Интернете, Mac, Android и iOS. Если вы видите только ссылку на собрание, но не видите кнопку  "Присоединиться" в поддерживаемом клиенте, возможно, шаблон собрания по сети для вашей службы не зарегистрирован на наших серверах. Дополнительные сведения см. в разделе ["Регистрация шаблона](#register-your-online-meeting-template) собрания по сети".

## <a name="register-your-online-meeting-template"></a>Регистрация шаблона собрания по сети

Регистрация надстройки собрания по сети не является обязательной. Это применимо только в том случае, если вы  хотите добавить кнопку "Присоединиться" на собраниях в дополнение к ссылке на собрание. После разработки надстройки для собраний по сети и ее регистрации создайте проблему на GitHub, следуя приведенным ниже рекомендациям. Мы свяжитесь с вами для координации временной шкалы регистрации.

> [!IMPORTANT]
> **Кнопка "** Присоединиться" поддерживается только в Outlook в Интернете, Mac, Android и iOS.

1. Создайте [новую проблему на GitHub](https://github.com/OfficeDev/office-js/issues/new).
1. **Задайте для** заголовка новой проблемы значение "Outlook: Регистрация шаблона собрания по сети для моей службы", `my-service` заменив ее именем службы.
1. В тексте проблемы замените существующий текст строкой, `newBody` заданной в переменной или аналогичной переменной из раздела "Реализация добавления сведений о собрании по сети" ранее в этой статье.[](#implement-adding-online-meeting-details)
1. Нажмите **кнопку "Отправить новую проблему"**.

![Новый экран проблем GitHub с примером содержимого Contoso.](../images/outlook-request-to-register-online-meeting-template.png)

## <a name="available-apis"></a>Доступные API

Для этой функции доступны следующие API.

- API организатора встреч
  - [Office.context.mailbox.item.body](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-body-member) ([Body.getAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#outlook-office-body-getasync-member(1)), [Body.setAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#outlook-office-body-setasync-member(1)))
  - [Office.context.mailbox.item.end](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-end-member) ([Time](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.loadCustomPropertiesAsync](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-loadcustompropertiesasync-member(1)) ([CustomProperties](/javascript/api/outlook/office.customproperties?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.location](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-location-member) ([расположение](/javascript/api/outlook/office.location?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.optionalAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-optionalattendees-member) ([recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.requiredAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-requiredattendees-member) ([получатели](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.start](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-start-member) ([Time](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.subject](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-subject-member) ([тема](/javascript/api/outlook/office.subject?view=outlook-js-preview&preserve-view=true))
  - [Office.context.roamingSettings](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context?view=outlook-js-preview&preserve-view=true#roamingsettings-roamingsettings) ([RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true))
- Обработка потока проверки подлинности
  - [API диалоговых окон](../develop/dialog-api-in-office-add-ins.md)

## <a name="restrictions"></a>Ограничения

Применяется несколько ограничений.

- Применимо только к поставщикам услуг онлайн-собраний.
- На экране создания собрания будут отображаться только надстройки, установленные администратором, заменив параметр Teams или Skype по умолчанию. Установленные пользователем надстройки не будут активироваться.
- Значок надстройки должен быть в оттенках серого с использованием шестнадцатеричных `#919191` кодов или его эквивалента в [других форматах цвета](https://convertingcolors.com/hex-color-919191.html).
- В режиме организатора встреч (создания) поддерживается только одна команда без пользовательского интерфейса.
- Надстройка должна обновить сведения о собрании в форме встречи в течение минутного времени ожидания. Однако любое время, затраченное в диалоговом окне, которое надстройка открывает для проверки подлинности и т. д., исключается из периода времени ожидания.

## <a name="see-also"></a>См. также

- [Надстройки для Outlook Mobile](outlook-mobile-addins.md)
- [Добавлена поддержка команд надстройки для Outlook Mobile](add-mobile-support.md)
