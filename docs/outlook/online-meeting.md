---
title: Создание надстройки Outlook Mobile для поставщика собраний в Интернете (Предварительная версия)
description: Сведения о том, как настроить надстройку Outlook Mobile для поставщика услуг по подключению к интерактивному собранию.
ms.topic: article
ms.date: 04/15/2020
localization_priority: Normal
ms.openlocfilehash: ed89205962bf4662096167eb78388b475fffdf91
ms.sourcegitcommit: 90c5830a5f2973a9ccd5c803b055e1b98d83f099
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/16/2020
ms.locfileid: "43529115"
---
# <a name="create-an-outlook-mobile-add-in-for-an-online-meeting-provider-preview"></a>Создание надстройки Outlook Mobile для поставщика собраний в Интернете (Предварительная версия)

Настройка собрания по сети — это основной интерфейс для пользователя Outlook, который позволяет легко [создать собрание Teams с помощью Outlook](/microsoftteams/teams-add-in-for-outlook) Mobile. Однако создание собрания по сети в Outlook со службой, отличной от Майкрософт, может быть утомительным. Реализуя эту функцию, поставщики услуг могут упростить процесс создания собраний по сети для пользователей надстроек Outlook.

> [!NOTE]
> Эта функция поддерживается только в [предварительной версии](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) на Android с подпиской на Office 365.

В этой статье вы узнаете, как настроить надстройку Outlook Mobile, чтобы позволить пользователям упорядочивать и присоединяться к собранию с помощью службы собраний по сети. В этой статье мы будем использовать фиктивный поставщик услуг по подключению к собраниям, "contoso".

## <a name="configure-the-manifest"></a>Настройка манифеста

Чтобы позволить пользователям создавать собрания по сети с надстройкой, необходимо настроить точку `MobileOnlineMeetingCommandSurface` расширения в манифесте под родительским элементом. `MobileFormFactor` Другие конструктивные параметры не поддерживаются.

В приведенном ниже примере показан пример манифеста, включающего `MobileFormFactor` элемент и `MobileOnlineMeetingCommandSurface` точку расширения.

```xml
...
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    ...
    <Hosts>
      <Host xsi:type="MailHost">
        <MobileFormFactor>
          <FunctionFile resid="residMobileFuncUrl" />
          <ExtensionPoint xsi:type="MobileOnlineMeetingCommandSurface">
            <!-- Configure selected extension point. -->
            <Control xsi:type="MobileButton" id="onlineMeetingFunctionButton">
              <Label resid="residUILessButton0Name" />
              <Icon>
                <bt:Image resid="UiLessIcon" size="25" scale="1" />
                <bt:Image resid="UiLessIcon" size="25" scale="2" />
                <bt:Image resid="UiLessIcon" size="25" scale="3" />
                <bt:Image resid="UiLessIcon" size="32" scale="1" />
                <bt:Image resid="UiLessIcon" size="32" scale="2" />
                <bt:Image resid="UiLessIcon" size="32" scale="2" />
                <bt:Image resid="UiLessIcon" size="48" scale="1" />
                <bt:Image resid="UiLessIcon" size="48" scale="2" />
                <bt:Image resid="UiLessIcon" size="48" scale="3" />
              </Icon>
              <Action xsi:type="ExecuteFunction">
                <FunctionName>insertContosoMeeting</FunctionName>
              </Action>
            </Control>
          </ExtensionPoint>
        </MobileFormFactor>
      </Host>
    </Hosts>
    ...
  </VersionOverrides>
</VersionOverrides>
...
```

## <a name="implement-adding-online-meeting-details"></a>Реализация добавления сведений о собрании по сети

В этом разделе описывается, как сценарий надстройки может обновить собрание пользователя, включив сведения о собрании по сети.

В приведенном ниже примере показано, как создать сведения о собрании по сети. Не отображается — как получить идентификатор организатора собрания и другие сведения из службы.

```js
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
```

В приведенном ниже примере показано, как определить функцию без пользовательского интерфейса `insertContosoMeeting` , именуемую ссылкой в манифесте, чтобы обновить текст собрания, используя сведения о собрании по сети.

```js
var mailboxItem;

// Office is ready.
Office.onReady(function () {
        mailboxItem = Office.context.mailbox.item;
    }
);

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
```

В следующем примере показана реализация вспомогательной функции `updateBody` , используемой в предыдущем примере, которая добавляет сведения о собрании по сети в текущий текст собрания.

```js
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

Следуйте обычным рекомендациям по [тестированию и проверке надстройки](testing-and-tips.md). После [загрузки неопубликованных приложений](sideload-outlook-add-ins-for-testing.md) в Outlook в Интернете, Windows или Mac перезапустите Outlook на мобильном устройстве с Android (Android — единственный поддерживаемый клиент). Затем на новом экране собрания убедитесь, что переключатель Microsoft Teams или Skype заменяется вашим собственным.

### <a name="create-meeting-ui"></a>Создание пользовательского интерфейса собрания

Как организатор собрания, при создании собрания должны отображаться растры, аналогичные следующим трем изображениям.

снимок экрана: " [Создание экрана собрания" на странице "переключить снимок экрана" на странице "Создание собрания" на странице "Создание собрания" на странице "Создание собрания" на странице "Создание собрания" на панели Android/Contoso ![](../images/outlook-android-create-online-meeting-off.png)](../images/outlook-android-create-online-meeting-off-expanded.png#lightbox) [ ![](../images/outlook-android-create-online-meeting-load.png)](../images/outlook-android-create-online-meeting-load-expanded.png#lightbox) [ ![](../images/outlook-android-create-online-meeting-on.png)](../images/outlook-android-create-online-meeting-on-expanded.png#lightbox)

### <a name="join-meeting-ui"></a>Пользовательский интерфейс присоединения к собранию

При просмотре собрания в качестве участника собрания должен отображаться экран, аналогичный следующему изображению.

[![снимок экрана с экраном "присоединение к собранию" на Android](../images/outlook-android-join-online-meeting-view-1.png)](../images/outlook-android-join-online-meeting-view-1-expanded.png#lightbox)

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
- В настоящее время эта функция не должна использоваться в производственных надстройках.
- В настоящее время Android является единственным поддерживаемым клиентом. Поддержка iOS скоро будет доступна.
- Только надстройки, установленные администратором, будут отображаться на экране создания собрания, заменив параметры группы по умолчанию или Skype. Надстройки, установленные пользователем, не будут активированы.
- Значок надстройки должен быть в оттенках серого с использованием `#919191` шестнадцатеричного кода или его эквивалента в [других цветовых форматах](https://convertingcolors.com/hex-color-919191.html).
- В режиме организатора встречи (создания) поддерживается только одна команда без пользовательского интерфейса.

## <a name="see-also"></a>См. также

- [Надстройки для Outlook Mobile](outlook-mobile-addins.md)
- [Добавление поддержки команд надстроек для Outlook Mobile](add-mobile-support.md)
