---
title: Создание надстройки Outlook Mobile для поставщика собраний в Интернете
description: Сведения о том, как настроить надстройку Outlook Mobile для поставщика услуг по подключению к интерактивному собранию.
ms.topic: article
ms.date: 06/25/2020
localization_priority: Normal
ms.openlocfilehash: 052ab4e71f8bc90e655a6ba780eacc18d43069e1
ms.sourcegitcommit: 065bf4f8e0d26194cee9689f7126702b391340cc
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/01/2020
ms.locfileid: "45006427"
---
# <a name="create-an-outlook-mobile-add-in-for-an-online-meeting-provider"></a>Создание надстройки Outlook Mobile для поставщика собраний в Интернете

Настройка собрания по сети — это основной интерфейс для пользователя Outlook, который позволяет легко [создать собрание Teams с помощью Outlook](/microsoftteams/teams-add-in-for-outlook) Mobile. Однако создание собрания по сети в Outlook со службой, отличной от Майкрософт, может быть утомительным. Реализуя эту функцию, поставщики услуг могут упростить процесс создания собраний по сети для пользователей надстроек Outlook.

> [!IMPORTANT]
> Эта функция поддерживается только в Android с подпиской на Office 365.

В этой статье вы узнаете, как настроить надстройку Outlook Mobile, чтобы позволить пользователям упорядочивать и присоединяться к собранию с помощью службы собраний по сети. В этой статье мы будем использовать фиктивный поставщик услуг по подключению к собраниям, "contoso".

## <a name="set-up-your-environment"></a>Настройка среды

Завершите работу с [быстрым запуском Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) , который создает проект надстройки с помощью генератора Yeoman для надстроек Office.

## <a name="configure-the-manifest"></a>Настройка манифеста

Чтобы позволить пользователям создавать собрания по сети с надстройкой, необходимо настроить `MobileOnlineMeetingCommandSurface` точку расширения в манифесте под родительским элементом `MobileFormFactor` . Другие конструктивные параметры не поддерживаются.

1. В редакторе кода откройте Быстрый запуск проекта.

1. Откройте файл **manifest.xml** , расположенный в корневом каталоге проекта.

1. Выберите весь `<VersionOverrides>` узел (включая открывающие и закрывающие теги) и замените его следующим XML-документом.

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
> Чтобы узнать больше о манифестах для надстроек Outlook, ознакомьтесь с разделом [манифесты надстроек Outlook](manifests.md) и [добавьте поддержку команд надстроек для Outlook Mobile](add-mobile-support.md).

## <a name="implement-adding-online-meeting-details"></a>Реализация добавления сведений о собрании по сети

В этом разделе описывается, как сценарий надстройки может обновить собрание пользователя, включив сведения о собрании по сети.

1. В проекте быстрого запуска откройте **commands.jsфайл./СРК/коммандс/** в редакторе кода.

1. Замените весь контент файла **commands.js** на следующий код JavaScript.

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

Следуйте обычным рекомендациям по [тестированию и проверке надстройки](testing-and-tips.md). После [загрузки неопубликованных приложений](sideload-outlook-add-ins-for-testing.md) в Outlook в Интернете, Windows или Mac перезапустите Outlook на мобильном устройстве с Android. (Android это единственный поддерживаемый клиент для сейчас.) Затем на новом экране собрания убедитесь, что переключатель Microsoft Teams или Skype заменяется вашим собственным.

### <a name="create-meeting-ui"></a>Создание пользовательского интерфейса собрания

Как организатор собрания, при создании собрания должны отображаться растры, аналогичные следующим трем изображениям.

снимок экрана: " [ ![ Создание экрана собрания" на странице "переключить](../images/outlook-android-create-online-meeting-off.png)](../images/outlook-android-create-online-meeting-off-expanded.png#lightbox) снимок экрана" на [ ![ странице "Создание собрания" на странице "Создание собрания" на](../images/outlook-android-create-online-meeting-load.png)](../images/outlook-android-create-online-meeting-load-expanded.png#lightbox) странице "Создание собрания" на странице " [ ![ Создание собрания](../images/outlook-android-create-online-meeting-on.png)](../images/outlook-android-create-online-meeting-on-expanded.png#lightbox) " на панели Android/contoso

### <a name="join-meeting-ui"></a>Пользовательский интерфейс присоединения к собранию

При просмотре собрания в качестве участника собрания должен отображаться экран, аналогичный следующему изображению.

[![снимок экрана с экраном "присоединение к собранию" на Android](../images/outlook-android-join-online-meeting-view-1.png)](../images/outlook-android-join-online-meeting-view-1-expanded.png#lightbox)

> [!IMPORTANT]
> Если вы не видите ссылку **присоединиться** , возможно, на наших серверах не зарегистрирован шаблон собрания в Интернете для вашей службы. Подробные сведения можно найти в разделе [Register The Online Template (шаблон собрания)](#register-your-online-meeting-template) .

## <a name="register-your-online-meeting-template"></a>Регистрация шаблона собрания в Интернете

Если вы хотите зарегистрировать шаблон собрания в Интернете для своей службы, вы можете создать ошибку GitHub с подробными сведениями. После этого мы свяжемся с вами, чтобы координировать временную шкалу регистрации.

1. Перейдите к разделу **Отзывы** в конце этой статьи.
1. Нажмите ссылку на **эту страницу** .
1. Задайте **название** новой неисправности "зарегистрировать шаблон собрания в сети для My-Service", заменив его на `my-service` имя службы.
1. В тексте вопроса замените строку "[Введите здесь обратную связь]" на строку, указанную в `newBody` переменной или аналогичной переменной в разделе [Реализация Добавление сведений о собрании по сети](#implement-adding-online-meeting-details) ранее в этой статье.
1. Нажмите кнопку **Добавить новую ошибку**.

![снимок экрана с новым экраном о проблемах GitHub с образцом контента contoso](../images/outlook-request-to-register-online-meeting-template.png)

## <a name="available-apis"></a>Доступные API

Для этой функции доступны следующие API.

- API организатора встречи
  - [Office. Context. Mailbox. Item. subject](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#subject) ([subject](/javascript/api/outlook/office.subject?view=outlook-js-preview))
  - [Office. Context. Mailbox. Item. Start](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#start) ([время](/javascript/api/outlook/office.time?view=outlook-js-preview))
  - [Office. Context. Mailbox. Item. end](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#end) ([время](/javascript/api/outlook/office.time?view=outlook-js-preview))
  - [Office. Context. Mailbox. Item. Location](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#location) ([Location](/javascript/api/outlook/office.location?view=outlook-js-preview))
  - [Office. Context. Mailbox. Item. optionalAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#optionalattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview))
  - [Office. Context. Mailbox. Item. requiredAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#requiredattendees) ([Recipients](/javascript/api/outlook/office.recipients?view=outlook-js-preview))
  - [Office. Context. Mailbox. Item. Body](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#body) ([Body. onasync](/javascript/api/outlook/office.body?view=outlook-js-preview#getasync-coerciontype--options--callback-), [Body. setAsync](/javascript/api/outlook/office.body?view=outlook-js-preview#setasync-data--options--callback-))
  - [Office. Context. Mailbox. Item. loadCustomPropertiesAsync](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview#loadcustompropertiesasync-callback--usercontext-) ([CustomProperties](/javascript/api/outlook/office.customproperties?view=outlook-js-preview))
  - [Office. Context. roamingSettings](../reference/objectmodel/preview-requirement-set/office.context.md?view=outlook-js-preview#roamingsettings-roamingsettings) ([roamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview))
- Обработка процесса проверки подлинности
  - [API диалоговых окон](../develop/dialog-api-in-office-add-ins.md)

## <a name="restrictions"></a>Наложен

Применяются некоторые ограничения.

- Применяется только к поставщикам услуг для собраний по сети.
- В настоящее время Android является единственным поддерживаемым клиентом. Поддержка iOS скоро будет доступна.
- Только надстройки, установленные администратором, будут отображаться на экране создания собрания, заменив параметры группы по умолчанию или Skype. Надстройки, установленные пользователем, не будут активированы.
- Значок надстройки должен быть в оттенках серого с использованием шестнадцатеричного кода `#919191` или его эквивалента в [других цветовых форматах](https://convertingcolors.com/hex-color-919191.html).
- В режиме организатора встречи (создания) поддерживается только одна команда без пользовательского интерфейса.

## <a name="see-also"></a>См. также

- [Надстройки для Outlook Mobile](outlook-mobile-addins.md)
- [Добавление поддержки команд надстроек для Outlook Mobile](add-mobile-support.md)
