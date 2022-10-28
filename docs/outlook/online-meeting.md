---
title: Создание надстройки Outlook для поставщика собраний по сети
description: Описывается настройка надстройки Outlook для поставщика услуг собраний по сети.
ms.topic: article
ms.date: 10/24/2022
ms.localizationpriority: medium
ms.openlocfilehash: 7c2cdb9f6369fd851a13fe45df132482b0ccdc0e
ms.sourcegitcommit: 693e9a9b24bb81288d41508cb89c02b7285c4b08
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/28/2022
ms.locfileid: "68767184"
---
# <a name="create-an-outlook-add-in-for-an-online-meeting-provider"></a>Создание надстройки Outlook для поставщика собраний по сети

Настройка собрания по сети — это основной интерфейс для пользователя Outlook, и вы можете [легко создать собрание Teams с помощью Outlook](/microsoftteams/teams-add-in-for-outlook). Однако создание собрания по сети в Outlook с помощью службы сторонних поставщиков может быть громоздким. Реализуя эту функцию, поставщики услуг могут упростить процесс создания собраний по сети и присоединения к ней для пользователей надстройки Outlook.

> [!IMPORTANT]
> Эта функция поддерживается в Outlook в Интернете, Windows, Mac, Android и iOS с подпиской Microsoft 365.

Из этой статьи вы узнаете, как настроить надстройку Outlook, чтобы пользователи могли организовать собрание и присоединиться к нему с помощью службы собраний по сети. В этой статье мы будем использовать вымышленного поставщика услуг онлайн-собраний Contoso.

## <a name="set-up-your-environment"></a>Настройка среды

Завершите [краткое руководство По созданию](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) проекта надстройки с помощью генератора Yeoman для надстроек Office.

## <a name="configure-the-manifest"></a>Настройка манифеста

Чтобы разрешить пользователям создавать собрания по сети с помощью надстройки, необходимо настроить манифест. Разметка отличается в зависимости от двух переменных:

- Тип целевой платформы; мобильный или не мобильный.
- Тип манифеста; xml или [манифест Teams для надстроек Office (предварительная версия).](../develop/json-manifest-overview.md)

Если надстройка использует XML-манифест, а надстройка будет поддерживаться только в Outlook в Интернете, Windows и Mac, выберите вкладку **Windows, Mac, веб-вкладку**. Однако если ваша надстройка также будет поддерживаться в Outlook для Android и iOS, перейдите на вкладку **Мобильные** устройства.

Если надстройка использует манифест Teams (предварительная версия), перейдите на вкладку **Манифест Teams (предварительная версия для разработчиков).**

> [!IMPORTANT]
> Поставщики онлайн-собраний пока не поддерживают манифест Teams (предварительная версия). Мы работаем над предоставлением этой поддержки в ближайшее время.

# <a name="windows-mac-web"></a>[Windows, Mac, Web](#tab/non-mobile)

1. В редакторе кода откройте созданный проект быстрого запуска Outlook.

1. Откройте **файлmanifest.xml** , расположенный в корне проекта.

1. Выберите весь **\<VersionOverrides\>** узел (включая открытые и закрытые теги) и замените его следующим XML-кодом.

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

Чтобы разрешить пользователям создавать собрания по сети со своего мобильного устройства, [точка расширения MobileOnlineMeetingCommandSurface](/javascript/api/manifest/extensionpoint#mobileonlinemeetingcommandsurface) настраивается в манифесте родительского элемента **\<MobileFormFactor\>**. Эта точка расширения не поддерживается в других форм-факторах.

1. В редакторе кода откройте созданный проект быстрого запуска Outlook.

1. Откройте **файлmanifest.xml** , расположенный в корне проекта.

1. Выберите весь **\<VersionOverrides\>** узел (включая открытые и закрытые теги) и замените его следующим XML-кодом.

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

# <a name="teams-manifest-developer-preview"></a>[Манифест Teams (предварительная версия для разработчиков)](#tab/jsonmanifest)

> [!IMPORTANT]
> Поставщики онлайн-собраний пока не поддерживают [манифест Teams для надстроек Office (предварительная версия).](../develop/json-manifest-overview.md) Эта вкладка предназначена для использования в будущем.

1. Откройте файл **manifest.json** .

1. Найдите *первый* объект в массиве authorization.permissions.resourceSpecific и задайте для его свойства name значение MailboxItem.ReadWrite.User. Это должно выглядеть так, когда вы закончите.

    ```json
    {
        "name": "MailboxItem.ReadWrite.User",
        "type": "Delegated"
    }
    ```

1. В массиве validDomains измените URL-адрес на "https://contoso.com", который является URL-адресом вымышленного поставщика онлайн-собраний. По завершении массив должен выглядеть следующим образом.

    ```json
    "validDomains": [
        "https://contoso.com"
    ],
    ```

1. Добавьте следующий объект в массив extensions.runtimes. Обратите внимание на указанные ниже аспекты этого кода.

   - Параметр minVersion набора обязательных элементов почтового ящика имеет значение 1.3, поэтому среда выполнения не будет запускаться на платформах и в версиях Office, где эта функция не поддерживается.
   - Для идентификатора среды выполнения задается описательное имя "online_meeting_runtime".
   - Для свойства "code.page" задается URL-адрес HTML-файла без пользовательского интерфейса, который загрузит команду функции.
   - Свойство "время существования" имеет значение "short", что означает, что среда выполнения запускается при выборе кнопки команды функции и завершает работу после завершения функции. (В некоторых редких случаях среда выполнения завершает работу до завершения обработчика. См [. раздел Среды выполнения в надстройках Office](../testing/runtimes.md).)
   - Существует действие для запуска функции с именем insertContosoMeeting. Вы создадите эту функцию на следующем шаге.

    ```json
    {
        "requirements": {
            "capabilities": [
                {
                    "name": "Mailbox",
                    "minVersion": "1.3"
                }
            ],
            "formFactors": [
                "desktop"
            ]
        },
        "id": "online_meeting_runtime",
        "type": "general",
        "code": {
            "page": "https://contoso.com/commands.html"
        },
        "lifetime": "short",
        "actions": [
            {
                "id": "insertContosoMeeting",
                "type": "executeFunction",
                "displayName": "insertContosoMeeting"
            }
        ]
    }
    ```

1. Замените массив extensions.ribbons следующим. Обратите внимание на следующие особенности этой разметки.

   - Параметр minVersion набора обязательных элементов почтового ящика имеет значение "1.3", поэтому настройки ленты не будут отображаться на платформах и в версиях Office, где эта функция не поддерживается.
   - Массив contexts указывает, что лента доступна только в окне организатора сведений о собрании.
   - На вкладке ленты по умолчанию (в окне организатора сведений о собрании) будет находиться пользовательская группа управления, помеченная **как собрание Contoso**.
   - Группа будет иметь кнопку с меткой **Добавить собрание Contoso**.
   - Для параметра actionId кнопки задано значение insertContosoMeeting, которое соответствует идентификатору действия, созданного на предыдущем шаге.

    ```json
    "ribbons": [
      {
        "requirements": {
            "capabilities": [
                {
                    "name": "Mailbox",
                    "minVersion": "1.3"
                }
            ],
            "scopes": [
                "mail"
            ],
            "formFactors": [
                "desktop"
            ]
        },
        "contexts": [
            "meetingDetailsOrganizer"
        ],
        "tabs": [
            {
                "builtInTabId": "TabDefault",
                "groups": [
                    {
                        "id": "apptComposeGroup",
                        "label": "Contoso meeting",
                        "controls": [
                            {
                                "id": "insertMeetingButton",
                                "type": "button",
                                "label": "Add a Contoso meeting",
                                "icons": [
                                    {
                                        "size": 16,
                                        "file": "icon-16.png"
                                    },
                                    {
                                        "size": 32,
                                        "file": "icon-32.png"
                                    },
                                    {
                                        "size": 64,
                                        "file": "icon-64_02.png"
                                    },
                                    {
                                        "size": 80,
                                        "file": "icon-80.png"
                                    }
                                ],
                                "supertip": {
                                    "title": "Add a Contoso meeting",
                                    "description": "Add a Contoso meeting to this appointment."
                                },
                                "actionId": "insertContosoMeeting",
                            }
                        ]
                    }
                ]
            }
        ]
      }
    ]
    ```

---

> [!TIP]
> Дополнительные сведения о манифестах надстроек Outlook см. [в разделах Манифесты надстроек Outlook](manifests.md) и [Добавление поддержки команд надстроек для Outlook Mobile](add-mobile-support.md).

## <a name="implement-adding-online-meeting-details"></a>Реализация добавления сведений о собрании по сети

В этом разделе вы узнаете, как скрипт надстройки может обновить собрание пользователя, чтобы включить сведения о собрании по сети. Следующее относится ко всем поддерживаемым платформам.

1. В том же проекте быстрого запуска откройте файл **./src/commands/commands.js** в редакторе кода.

1. Замените все содержимое **файлаcommands.js** следующим кодом JavaScript.

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

    let mailboxItem;

    // Office is ready.
    Office.onReady(function () {
            mailboxItem = Office.context.mailbox.item;
        }
    );

    // 2. How to define and register a function command named `insertContosoMeeting` (referenced in the manifest)
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

Следуйте обычному руководству, чтобы [протестировать и проверить надстройку](testing-and-tips.md), а затем [загрузить манифест неопубликованного](sideload-outlook-add-ins-for-testing.md) приложения в Outlook в Интернете, Windows или Mac. Если надстройка также поддерживает мобильные устройства, перезапустите Outlook на устройстве Android или iOS после загрузки неопубликованного приложения. После загрузки надстройки создайте собрание и убедитесь, что переключатель Microsoft Teams или Skype заменен собственным.

### <a name="create-meeting-ui"></a>Создание пользовательского интерфейса собрания

Как организатор собрания, при создании собрания должны отображаться экраны, аналогичные приведенным ниже трем изображениям.

[![Экран создания собрания в Android с выключенным переключателем Contoso.](../images/outlook-android-create-online-meeting-off.png)](../images/outlook-android-create-online-meeting-off-expanded.png#lightbox) [![Экран создания собрания в Android с переключателем загрузки Contoso.](../images/outlook-android-create-online-meeting-load.png)](../images/outlook-android-create-online-meeting-load-expanded.png#lightbox) [![Экран создания собрания в Android с переключателем Contoso.](../images/outlook-android-create-online-meeting-on.png)](../images/outlook-android-create-online-meeting-on-expanded.png#lightbox)

### <a name="join-meeting-ui"></a>Присоединение к пользовательскому интерфейсу собрания

При просмотре собрания в качестве участника собрания должен отобразиться экран, аналогичный следующему изображению.

[![Экран присоединения к собранию в Android.](../images/outlook-android-join-online-meeting-view-1.png)](../images/outlook-android-join-online-meeting-view-1-expanded.png#lightbox)

> [!IMPORTANT]
> Кнопка **Присоединиться** поддерживается только в Outlook в Интернете, Mac, Android и iOS. Если вы видите только ссылку на собрание, но не видите кнопку **Присоединиться** в поддерживаемом клиенте, возможно, шаблон собрания по сети для вашей службы не зарегистрирован на наших серверах. Дополнительные сведения см. в разделе [Регистрация шаблона онлайн-собрания](#register-your-online-meeting-template) .

## <a name="register-your-online-meeting-template"></a>Регистрация шаблона собрания по сети

Регистрация надстройки для собраний по сети не является обязательной. Она применяется только в том случае, если вы хотите отображать кнопку **Присоединиться** на собраниях в дополнение к ссылке на собрание. После разработки надстройки для собраний по сети и ее регистрации создайте проблему в GitHub, используя следующие рекомендации. Мы свяжемся с вами, чтобы согласовать временную шкалу регистрации.

> [!IMPORTANT]
> Кнопка **Присоединиться** поддерживается только в Outlook в Интернете, Mac, Android и iOS.

1. Создайте [новую проблему GitHub](https://github.com/OfficeDev/office-js/issues/new).
1. Задайте **для новой проблемы заголовок** "Outlook: регистрация шаблона собрания по сети для my-service", заменив `my-service` именем своей службы.
1. В тексте проблемы замените существующий текст строкой, заданной в `newBody` или аналогичной переменной из раздела [Реализация добавления сведений о собрании по сети](#implement-adding-online-meeting-details) ранее в этой статье.
1. Нажмите **кнопку Отправить новую проблему**.

![Новый экран проблемы GitHub с примером содержимого Contoso.](../images/outlook-request-to-register-online-meeting-template.png)

## <a name="available-apis"></a>Доступные API

Для этой функции доступны следующие API.

- API организатора встреч
  - [Office.context.mailbox.item.body](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-body-member) ([Body.getAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#outlook-office-body-getasync-member(1)), [Body.setAsync](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#outlook-office-body-setasync-member(1)))
  - [Office.context.mailbox.item.end](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-end-member) ([Time](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.loadCustomPropertiesAsync](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-loadcustompropertiesasync-member(1)) ([CustomProperties](/javascript/api/outlook/office.customproperties?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.location](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-location-member) ([расположение](/javascript/api/outlook/office.location?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.optionalAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-optionalattendees-member) ([Получатели](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.requiredAttendees](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-requiredattendees-member) ([Получатели](/javascript/api/outlook/office.recipients?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.start](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-start-member) ([Time](/javascript/api/outlook/office.time?view=outlook-js-preview&preserve-view=true))
  - [Office.context.mailbox.item.subject](/javascript/api/outlook/office.appointmentcompose?view=outlook-js-preview&preserve-view=true#outlook-office-appointmentcompose-subject-member) ([Subject](/javascript/api/outlook/office.subject?view=outlook-js-preview&preserve-view=true))
  - [Office.context.roamingSettings](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context?view=outlook-js-preview&preserve-view=true#roamingsettings-roamingsettings) ([RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true))
- Обработка потока проверки подлинности
  - [API диалоговых окон](../develop/dialog-api-in-office-add-ins.md)

## <a name="restrictions"></a>Ограничения

Применяется несколько ограничений.

- Применимо только к поставщикам услуг собраний по сети.
- На экране создания собрания будут отображаться только установленные администратором надстройки, заменяющие параметры Teams или Skype по умолчанию. Установленные пользователем надстройки не активируются.
- Значок надстройки должен быть в оттенках серого с использованием шестнадцатеричного кода `#919191` или его эквивалента в [других цветовых форматах](https://convertingcolors.com/hex-color-919191.html).
- В режиме организатора встреч (создание) поддерживается только одна команда функции.
- Надстройка должна обновить сведения о собрании в форме встречи в течение одной минуты времени ожидания. Однако любое время, затраченное в диалоговом окне надстройки, открытой для проверки подлинности, например, исключается из периода ожидания.

## <a name="see-also"></a>См. также

- [Надстройки для Outlook Mobile](outlook-mobile-addins.md)
- [Добавлена поддержка команд надстроек для Outlook Mobile](add-mobile-support.md)
