---
title: Настройка надстройки Outlook для активации на основе событий
description: Узнайте, как настроить надстройку Outlook для активации на основе событий.
ms.topic: article
ms.date: 10/24/2022
ms.localizationpriority: medium
ms.openlocfilehash: b5ae744350389ed222284808a67a9b7c30211136
ms.sourcegitcommit: 693e9a9b24bb81288d41508cb89c02b7285c4b08
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/28/2022
ms.locfileid: "68767177"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation"></a>Настройка надстройки Outlook для активации на основе событий

Без функции активации на основе событий пользователь должен явно запустить надстройку для выполнения своих задач. Эта функция позволяет надстройке выполнять задачи на основе определенных событий, особенно для операций, которые применяются к каждому элементу. Также можно интегрировать с командами области задач и функций.

К концу этого пошагового руководства у вас будет надстройка, которая запускается при каждом создании нового элемента и задает тему.

> [!NOTE]
> Поддержка этой функции появилась в [наборе требований 1.10](/javascript/api/requirement-sets/outlook/requirement-set-1.10/outlook-requirement-set-1.10), а дополнительные события теперь доступны в последующих наборах требований. Дополнительные сведения о минимальном наборе требований события, а также о клиентах и платформах, которые его поддерживают, см. в [разделах Поддерживаемые события](#supported-events) и [Наборы требований, поддерживаемые серверами Exchange и клиентами Outlook](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients).
>
> Активация на основе событий не поддерживается в Outlook для iOS или Android.

> [!IMPORTANT]
> Активация на основе событий пока не поддерживается для [манифеста Teams для надстроек Office (предварительная версия).](../develop/json-manifest-overview.md) Мы работаем над предоставлением этой поддержки в ближайшее время.

## <a name="supported-events"></a>Поддерживаемые события

В следующей таблице перечислены доступные в настоящее время события и поддерживаемые клиенты для каждого события. При возникновении события обработчик получает `event` объект, который может содержать сведения, относящиеся к типу события. Столбец **Описание** содержит ссылку на связанный объект, если применимо.

|Каноническое имя события</br>и имя манифеста XML|Имя манифеста Teams|Описание|Минимальный набор требований и поддерживаемые клиенты|
|---|---|---|---|
|`OnNewMessageCompose`| newMessageComposeCreated |При создании нового сообщения (включая ответ, ответить всем и пересылать), но не при редактировании, например, черновика.|[1.10](/javascript/api/requirement-sets/outlook/requirement-set-1.10/outlook-requirement-set-1.10)<br><br>— Windows<sup>1</sup><br>— Веб-браузер<br>— Новый пользовательский интерфейс Mac |
|`OnNewAppointmentOrganizer`|newAppointmentOrganizerCreated|При создании новой встречи, но не при редактировании существующей.|[1.10](/javascript/api/requirement-sets/outlook/requirement-set-1.10/outlook-requirement-set-1.10)<br><br>— Windows<sup>1</sup><br>— Веб-браузер<br>— Новый пользовательский интерфейс Mac |
|`OnMessageAttachmentsChanged`|messageAttachmentsChanged|При добавлении или удалении вложений при создании сообщения.<br><br>Объект данных, зависящий от события: [AttachmentsChangedEventArgs](/javascript/api/outlook/office.attachmentschangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>— Windows<sup>1</sup><br>— Веб-браузер<br>— Новый пользовательский интерфейс Mac|
|`OnAppointmentAttachmentsChanged`|appointmentAttachmentsChanged|При добавлении или удалении вложений при составлении встречи.<br><br>Объект данных, зависящий от события: [AttachmentsChangedEventArgs](/javascript/api/outlook/office.attachmentschangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>— Windows<sup>1</sup><br>— Веб-браузер<br>— Новый пользовательский интерфейс Mac|
|`OnMessageRecipientsChanged`|messageRecipientsChanged|При добавлении или удалении получателей при создании сообщения.<br><br>Объект данных, зависящий от события: [RecipientsChangedEventArgs](/javascript/api/outlook/office.recipientschangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>— Windows<sup>1</sup><br>— Веб-браузер<br>— Новый пользовательский интерфейс Mac|
|`OnAppointmentAttendeesChanged`|appointmentAttendeesChanged|При добавлении или удалении участников во время создания встречи.<br><br>Объект данных, зависящий от события: [RecipientsChangedEventArgs](/javascript/api/outlook/office.recipientschangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>— Windows<sup>1</sup><br>— Веб-браузер<br>— Новый пользовательский интерфейс Mac|
|`OnAppointmentTimeChanged`|appointmentTimeChanged|При изменении даты и времени при составлении встречи.<br><br>Объект данных, зависящий от события: [AppointmentTimeChangedEventArgs](/javascript/api/outlook/office.appointmenttimechangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>— Windows<sup>1</sup><br>— Веб-браузер<br>— Новый пользовательский интерфейс Mac|
|`OnAppointmentRecurrenceChanged`|appointmentRecurrenceChanged|При добавлении, изменении или удалении сведений о повторении при составлении встречи. Если дата и время изменены, `OnAppointmentTimeChanged` событие также будет запущено.<br><br>Объект данных, зависящий от события: [RecurrenceChangedEventArgs](/javascript/api/outlook/office.recurrencechangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>— Windows<sup>1</sup><br>— Веб-браузер<br>— Новый пользовательский интерфейс Mac|
|`OnInfoBarDismissClicked`|infoBarDismissClicked|При отклонении уведомления при создании сообщения или элемента встречи. Будет уведомлена только надстройка, которая добавила уведомление.<br><br>Объект данных, зависящий от события: [InfobarClickedEventArgs](/javascript/api/outlook/office.infobarclickedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>— Windows<sup>1</sup><br>— Веб-браузер<br>— Новый пользовательский интерфейс Mac|
|`OnMessageSend`|messageSending|При отправке элемента сообщения. Дополнительные сведения см. в [пошаговом руководстве по интеллектуальным оповещениям](smart-alerts-onmessagesend-walkthrough.md).|[1.12](/javascript/api/requirement-sets/outlook/requirement-set-1.12/outlook-requirement-set-1.12)<br><br>— Windows<sup>1</sup><br>— Веб-браузер<br>— Новый пользовательский интерфейс Mac|
|`OnAppointmentSend`|appointmentSending|При отправке элемента встречи. Дополнительные сведения см. в [пошаговом руководстве по интеллектуальным оповещениям](smart-alerts-onmessagesend-walkthrough.md).|[1.12](/javascript/api/requirement-sets/outlook/requirement-set-1.12/outlook-requirement-set-1.12)<br><br>— Windows<sup>1</sup><br>— Веб-браузер<br>— Новый пользовательский интерфейс Mac|
|`OnMessageCompose`|messageComposeOpened|При создании нового сообщения (включая ответить, ответить всем и пересылать) или редактировании черновика.|[1.12](/javascript/api/requirement-sets/outlook/requirement-set-1.12/outlook-requirement-set-1.12)<br><br>— Windows<sup>1</sup><br>— Веб-браузер<br>— Новый пользовательский интерфейс Mac|
|`OnAppointmentOrganizer`|appointmentOrganizerOpened|При создании новой встречи или редактировании существующей.|[1.12](/javascript/api/requirement-sets/outlook/requirement-set-1.12/outlook-requirement-set-1.12)<br><br>— Windows<sup>1</sup><br>— Веб-браузер<br>— Новый пользовательский интерфейс Mac|

> [!NOTE]
> <sup>1</sup> Для работы надстроек на основе событий в Outlook в Windows требуется как минимум Windows 10 версии 1903 (сборка 18362) или Windows Server 2019 версии 1903.

## <a name="set-up-your-environment"></a>Настройка среды

Завершите [краткое руководство По созданию](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) проекта надстройки с [помощью генератора Yeoman для надстроек Office](../develop/yeoman-generator-overview.md).

## <a name="configure-the-manifest"></a>Настройка манифеста

Чтобы настроить манифест, выберите вкладку для используемого типа манифеста.

# <a name="xml-manifest"></a>[XML-манифест](#tab/xmlmanifest)

Чтобы включить активацию надстройки на основе событий, необходимо настроить элемент [Runtimes](/javascript/api/manifest/runtimes) и точку `VersionOverridesV1_1` расширения [LaunchEvent](/javascript/api/manifest/extensionpoint#launchevent) в узле манифеста. На данный момент `DesktopFormFactor` является единственным поддерживаемым форм-фактором.

1. В редакторе кода откройте проект быстрого запуска.

1. Откройте **файлmanifest.xml** , расположенный в корне проекта.

1. Выберите весь **\<VersionOverrides\>** узел (включая открытые и закрытые теги) и замените его следующим XML-кодом, а затем сохраните изменения.

```XML
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    <Requirements>
      <bt:Sets DefaultMinVersion="1.10">
        <bt:Set Name="Mailbox" />
      </bt:Sets>
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- Event-based activation happens in a lightweight runtime.-->
        <Runtimes>
          <!-- HTML file including reference to or inline JavaScript event handlers.
               This is used by Outlook on the web and Outlook on the new Mac UI. -->
          <Runtime resid="WebViewRuntime.Url">
            <!-- JavaScript file containing event handlers. This is used by Outlook on Windows. -->
            <Override type="javascript" resid="JSRuntime.Url"/>
          </Runtime>
        </Runtimes>
        <DesktopFormFactor>
          <FunctionFile resid="Commands.Url" />
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <OfficeTab id="TabDefault">
              <Group id="msgReadGroup">
                <Label resid="GroupLabel" />
                <Control xsi:type="Button" id="msgReadOpenPaneButton">
                  <Label resid="TaskpaneButton.Label" />
                  <Supertip>
                    <Title resid="TaskpaneButton.Label" />
                    <Description resid="TaskpaneButton.Tooltip" />
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16" />
                    <bt:Image size="32" resid="Icon.32x32" />
                    <bt:Image size="80" resid="Icon.80x80" />
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <SourceLocation resid="Taskpane.Url" />
                  </Action>
                </Control>
                <Control xsi:type="Button" id="ActionButton">
                  <Label resid="ActionButton.Label"/>
                  <Supertip>
                    <Title resid="ActionButton.Label"/>
                    <Description resid="ActionButton.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16x16"/>
                    <bt:Image size="32" resid="Icon.32x32"/>
                    <bt:Image size="80" resid="Icon.80x80"/>
                  </Icon>
                  <Action xsi:type="ExecuteFunction">
                    <FunctionName>action</FunctionName>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>

          <!-- Can configure other command surface extension points for add-in command support. -->

          <!-- Enable launching the add-in on the included events. -->
          <ExtensionPoint xsi:type="LaunchEvent">
            <LaunchEvents>
              <LaunchEvent Type="OnNewMessageCompose" FunctionName="onNewMessageComposeHandler"/>
              <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="onNewAppointmentComposeHandler"/>
              
              <!-- Other available events -->
              <!--
              <LaunchEvent Type="OnMessageAttachmentsChanged" FunctionName="onMessageAttachmentsChangedHandler" />
              <LaunchEvent Type="OnAppointmentAttachmentsChanged" FunctionName="onAppointmentAttachmentsChangedHandler" />
              <LaunchEvent Type="OnMessageRecipientsChanged" FunctionName="onMessageRecipientsChangedHandler" />
              <LaunchEvent Type="OnAppointmentAttendeesChanged" FunctionName="onAppointmentAttendeesChangedHandler" />
              <LaunchEvent Type="OnAppointmentTimeChanged" FunctionName="onAppointmentTimeChangedHandler" />
              <LaunchEvent Type="OnAppointmentRecurrenceChanged" FunctionName="onAppointmentRecurrenceChangedHandler" />
              <LaunchEvent Type="OnInfoBarDismissClicked" FunctionName="onInfobarDismissClickedHandler" />
              <LaunchEvent Type="OnMessageSend" FunctionName="onMessageSendHandler" SendMode="PromptUser" />
              <LaunchEvent Type="OnAppointmentSend" FunctionName="onAppointmentSendHandler" SendMode="PromptUser" />
              <LaunchEvent Type="OnMessageCompose" FunctionName="onMessageComposeHandler" />
              <LaunchEvent Type="OnAppointmentOrganizer" FunctionName="onAppointmentOrganizerHandler" />
              -->
            </LaunchEvents>
            <!-- Identifies the runtime to be used (also referenced by the Runtime element). -->
            <SourceLocation resid="WebViewRuntime.Url"/>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16" DefaultValue="https://localhost:3000/assets/icon-16.png"/>
        <bt:Image id="Icon.32x32" DefaultValue="https://localhost:3000/assets/icon-32.png"/>
        <bt:Image id="Icon.80x80" DefaultValue="https://localhost:3000/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Commands.Url" DefaultValue="https://localhost:3000/commands.html" />
        <bt:Url id="Taskpane.Url" DefaultValue="https://localhost:3000/taskpane.html" />
        <bt:Url id="WebViewRuntime.Url" DefaultValue="https://localhost:3000/commands.html" />
        <!-- Entry needed for Outlook on Windows. -->
        <bt:Url id="JSRuntime.Url" DefaultValue="https://localhost:3000/launchevent.js" />
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GroupLabel" DefaultValue="Contoso Add-in"/>
        <bt:String id="TaskpaneButton.Label" DefaultValue="Show Taskpane"/>
        <bt:String id="ActionButton.Label" DefaultValue="Perform an action"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Opens a pane displaying all available properties."/>
        <bt:String id="ActionButton.Tooltip" DefaultValue="Perform an action when clicked."/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</VersionOverrides>
```

Outlook в Windows использует файл JavaScript, а Outlook в Интернете и в новом пользовательском интерфейсе Mac — HTML-файл, который может ссылаться на тот же файл JavaScript. Необходимо предоставить ссылки на оба этих файла в `Resources` узле манифеста, так как платформа Outlook в конечном итоге определяет, следует ли использовать HTML или JavaScript на основе клиента Outlook. Таким образом, чтобы настроить обработку событий, укажите расположение HTML-кода в элементе **\<Runtime\>** , а затем в его `Override` дочернем элементе укажите расположение файла JavaScript, встроенного или на который ссылается HTML-код.

# <a name="teams-manifest-developer-preview"></a>[Манифест Teams (предварительная версия для разработчиков)](#tab/jsonmanifest)

> [!IMPORTANT]
> Активация на основе событий пока не поддерживается для [манифеста Teams для надстроек Office (предварительная версия).](../develop/json-manifest-overview.md) Эта вкладка предназначена для использования в будущем.

1. Откройте файл **manifest.json** .

1. Добавьте следующий объект в массив extensions.runtimes. Обратите внимание на указанные ниже особенности этой разметки.

   - Параметр minVersion набора обязательных элементов почтового ящика имеет значение "1.10", так как в таблице, приведенной выше в этой статье, указано, что это самая низкая версия набора требований, поддерживающая `OnNewMessageCompose` события и `OnNewAppointmentCompose` .
   - Для идентификатора среды выполнения задается описательное имя "autorun_runtime".
   - Свойство "code" имеет дочернее свойство page, для которого задано значение HTML-файла, и дочернее свойство script, для которого задается файл JavaScript. Вы создадите или измените эти файлы на последующих шагах. Office использует одно из этих значений в зависимости от платформы.
       - Office в Windows выполняет обработчики событий в среде выполнения, доступной только для JavaScript, которая загружает файл JavaScript напрямую.
       - Office на Mac и в Интернете выполняют обработчики в среде выполнения браузера, которая загружает HTML-файл. Этот файл, в свою очередь, содержит `<script>` тег, который загружает файл JavaScript.
     Дополнительные сведения см. [в разделе Среды выполнения в надстройках Office](../testing/runtimes.md).
   - Свойство "время существования" имеет значение "short", что означает, что среда выполнения запускается при активации одного из событий и завершает работу по завершении обработчика. (В некоторых редких случаях среда выполнения завершает работу до завершения обработчика. См [. раздел Среды выполнения в надстройках Office](../testing/runtimes.md).)
   - Существует два типа "действий", которые могут выполняться в среде выполнения. Вы создадите функции, соответствующие этим действиям, на следующем шаге.

    ```json
     {
        "requirements": {
            "capabilities": [
                {
                    "name": "Mailbox",
                    "minVersion": "1.10"
                }
            ]
        },
        "id": "autorun_runtime",
        "type": "general",
        "code": {
            "page": "https://localhost:3000/commands.html",
            "script": "https://localhost:3000/launchevent.js"
        },
        "lifetime": "short",
        "actions": [
            {
                "id": "onNewMessageComposeHandler",
                "type": "executeFunction",
                "displayName": "onNewMessageComposeHandler"
            },
            {
                "id": "onNewAppointmentComposeHandler",
                "type": "executeFunction",
                "displayName": "onNewAppointmentComposeHandler"
            }
        ]
    }
    ```

1. Добавьте следующий массив autoRunEvents в качестве свойства объекта в массиве extensions.

    ```json
    "autoRunEvents": [
    
    ]
    ```

1. Добавьте следующий объект в массив autoRunEvents. Свойство events сопоставляет обработчики с событиями, как описано в таблице ранее в этой статье. Имена обработчиков должны совпадать с именами, используемыми в свойствах id объектов в массиве actions на предыдущем шаге.

    ```json
      {
          "requirements": {
              "capabilities": [
                  {
                      "name": "Mailbox",
                      "minVersion": "1.10"
                  }
              ],
              "scopes": [
                  "mail"
              ]
          },
          "events": [
              {
                  "type": "newMessageComposeCreated",
                  "actionId": "onNewMessageComposeHandler"
              },
              {
                  "type": "newAppointmentOrganizerCreated",
                  "actionId": "onNewAppointmentComposeHandler"
              }
          ]
      }
    ```

---

> [!TIP]
>
> - Сведения о средах выполнения в надстройках см. в статье [Среды выполнения в надстройках Office](../testing/runtimes.md).
> - Дополнительные сведения о манифестах надстроек Outlook см. [в статье Манифесты надстроек Outlook](manifests.md).

## <a name="implement-event-handling"></a>Реализация обработки событий

Необходимо реализовать обработку выбранных событий.

В этом сценарии вы добавите обработку создания новых элементов.

1. В том же проекте быстрого запуска создайте папку **launchevent** в каталоге **./src** .

1. В папке **./src/launchevent** создайте файл **с именемlaunchevent.js**.

1. Откройте файл **./src/launchevent/launchevent.js** в редакторе кода и добавьте следующий код JavaScript.

    ```js
    /*
    * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
    * See LICENSE in the project root for license information.
    */

    function onNewMessageComposeHandler(event) {
      setSubject(event);
    }
    function onNewAppointmentComposeHandler(event) {
      setSubject(event);
    }
    function setSubject(event) {
      Office.context.mailbox.item.subject.setAsync(
        "Set by an event-based add-in!",
        {
          "asyncContext": event
        },
        function (asyncResult) {
          // Handle success or error.
          if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
            console.error("Failed to set subject: " + JSON.stringify(asyncResult.error));
          }

          // Call event.completed() after all work is done.
          asyncResult.asyncContext.completed();
        });
    }

    // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
    Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
    Office.actions.associate("onNewAppointmentComposeHandler", onNewAppointmentComposeHandler);
    ```

1. Сохраните изменения.

> [!IMPORTANT]
> Windows. В настоящее время импорт не поддерживается в файле JavaScript, в котором реализуется обработка активации на основе событий.

## <a name="update-the-commands-html-file"></a>Обновление HTML-файла команд

1. В папке **./src/commands** откройте **commands.html**.

1. Непосредственно перед закрывающим **тегом head** (`</head>`) добавьте запись скрипта, чтобы включить код JavaScript для обработки событий.

    ```html
    <script type="text/javascript" src="../launchevent/launchevent.js"></script>
    ```

1. Сохраните изменения.

## <a name="update-webpack-config-settings"></a>Обновление настроек конфигурации webpack

1. Откройте **файлwebpack.config.js** , который находится в корневом каталоге проекта, и выполните следующие действия.

1. `plugins` Найдите массив в объекте `config` и добавьте этот новый объект в начале массива.

    ```js
    new CopyWebpackPlugin({
      patterns: [
        {
          from: "./src/launchevent/launchevent.js",
          to: "launchevent.js",
        },
      ],
    }),
    ```

1. Сохраните изменения.

## <a name="try-it-out"></a>Проверка

1. Выполните следующие команды в корневом каталоге проекта. При запуске `npm start`запустится локальный веб-сервер (если он еще не запущен), и надстройка будет загружена неопубликованно.

    ```command&nbsp;line
    npm run build
    ```

    ```command&nbsp;line
    npm start
    ```

    > [!NOTE]
    > Если надстройка не была автоматически загружена неопубликованным приложением, следуйте инструкциям в разделе [Загрузка неопубликованных надстроек Outlook для тестирования,](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-manually) чтобы вручную загрузить надстройку неопубликованного приложения в Outlook.

1. Создайте новое сообщение в веб-версии Outlook.

    ![Окно сообщения в Outlook в Интернете с темой, заданной в compose.](../images/outlook-web-autolaunch-1.png)

1. В Outlook в новом пользовательском интерфейсе Mac создайте новое сообщение.

    ![Окно сообщения в Outlook в новом пользовательском интерфейсе Mac с темой, заданной в compose.](../images/outlook-mac-autolaunch.png)

1. В Outlook в Windows создайте новое сообщение.

    ![Окно сообщения в Outlook в Windows с заданной темой при создании.](../images/outlook-win-autolaunch.png)

## <a name="debug"></a>Отладка

При внесении изменений в обработку событий запуска в надстройке следует учитывать следующее:

- Если вы обновили манифест, [удалите надстройку](sideload-outlook-add-ins-for-testing.md#remove-a-sideloaded-add-in), а затем снова загрузите ее неопубликованное приложение. Если вы используете Outlook в Windows, закройте и снова откройте Outlook.
- Если вы внесли изменения в файлы, отличные от манифеста, закройте и снова откройте Outlook в Windows или обновите вкладку браузера с Outlook в Интернете.

При реализации собственных функциональных возможностей может потребоваться выполнить отладку кода. Инструкции по отладке активации надстройки на основе событий см. [в разделе Отладка надстройки Outlook на основе событий](debug-autolaunch.md).

Ведение журнала среды выполнения также доступно для этой функции в Windows. Дополнительные сведения см [. в статье Отладка надстройки с помощью ведения журнала среды выполнения](../testing/runtime-logging.md#runtime-logging-on-windows).

[!INCLUDE [Loopback exemption note](../includes/outlook-loopback-exemption.md)]

## <a name="deploy-to-users"></a>Развертывание для пользователей

Надстройки на основе событий можно развернуть, отправив манифест через Центр администрирования Microsoft 365. На портале администрирования разверните раздел **Параметры** в области навигации и выберите **Интегрированные приложения**. На странице **Интегрированные приложения** выберите действие **Отправить пользовательские приложения** .

![Страница Интегрированные приложения на Центр администрирования Microsoft 365 с выделенным действием Отправить пользовательские приложения.](../images/outlook-deploy-event-based-add-ins.png)

> [!IMPORTANT]
> Надстройки на основе событий ограничены только развертываниями, управляемыми администратором. Пользователи не могут активировать надстройки на основе событий из AppSource или магазина Office в приложении. Дополнительные сведения см. [в статье Параметры описания AppSource для надстройки Outlook на основе событий](autolaunch-store-options.md).

[!INCLUDE [outlook-smart-alerts-deployment](../includes/outlook-smart-alerts-deployment.md)]

## <a name="event-based-activation-behavior-and-limitations"></a>Поведение и ограничения активации на основе событий

Ожидается, что обработчики событий запуска надстроек будут короткими, упрощенными и максимально неактивными. После активации время ожидания надстройки будет истекать в течение примерно 300 секунд, максимально допустимого времени для запуска надстроек на основе событий. Чтобы сообщить о том, что надстройка завершила обработку события запуска, связанный обработчик событий должен вызвать `event.completed` метод . (Обратите внимание, что код, включенный после инструкции `event.completed` , не гарантируется для выполнения.) При каждом срабатывании события, которое обрабатывает надстройка, надстройка активируется повторно и запускает связанный обработчик событий, а время ожидания сбрасывается. Надстройка заканчивается после того, как истекает время ожидания, или пользователь закрывает окно создания или отправляет элемент.

Если у пользователя есть несколько надстроек, которые подписаны на одно событие, платформа Outlook запускает надстройки в не определенном порядке. В настоящее время активно может выполняться только пять надстроек на основе событий.

Во всех поддерживаемых клиентах Outlook пользователь должен оставаться в текущем почтовом элементе, где была активирована надстройка, чтобы завершить выполнение. Переход от текущего элемента (например, переключение на другое окно создания или вкладку) завершает операцию надстройки. Надстройка также прекращает работу, когда пользователь отправляет сообщение или встречу, которую он создает.

Импорт не поддерживается в файле JavaScript, в котором реализуется обработка активации на основе событий в клиенте Windows.

Некоторые api Office.js, которые изменяют или изменяют пользовательский интерфейс, не допускаются из надстроек на основе событий. Ниже приведены заблокированные API.

- В разделе `Office.context.auth`:
  - `getAccessToken`
  - `getAccessTokenAsync`
    > [!NOTE]
    > [OfficeRuntime.auth](/javascript/api/office-runtime/officeruntime.auth) поддерживается во всех версиях Outlook, поддерживающих активацию на основе событий и единый вход, а [Office.auth](/javascript/api/office/office.auth) поддерживается только в некоторых сборках Outlook. Дополнительные сведения см. [в статье Включение единого входа в надстройках Outlook, использующих активацию на основе событий](use-sso-in-event-based-activation.md).
- В разделе `Office.context.mailbox`:
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- В разделе `Office.context.mailbox.item`:
  - `close`
- В разделе `Office.context.ui`:
  - `displayDialogAsync`
  - `messageParent`

### <a name="requesting-external-data"></a>Запрос внешних данных

Внешние данные можно запрашивать с помощью API [, например Fetch](https://developer.mozilla.org/docs/Web/API/Fetch_API) , или с помощью [XMLHttpRequest (XHR) —](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest) стандартного веб-API, который отправляет HTTP-запросы для взаимодействия с серверами.

Имейте в виду, что при использовании объектов XMLHttpRequest необходимо использовать дополнительные меры безопасности, требуя [одну и ту же политику источника](https://developer.mozilla.org/docs/Web/Security/Same-origin_policy) и [простой CORS (общий доступ к ресурсам между источниками).](https://developer.mozilla.org/docs/Web/HTTP/CORS)

[Простая реализация CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS#simple_requests):

- Не удается использовать файлы cookie.
- Поддерживает только простые методы, такие как `GET`, `HEAD`и `POST`.
- Принимает простые заголовки с именами `Accept`полей , `Accept-Language`или `Content-Language`.
- Можно использовать `Content-Type`, при условии, что тип контента : `application/x-www-form-urlencoded`, `text/plain`или `multipart/form-data`.
- Не удается зарегистрировать прослушиватели событий в объекте, возвращаемом `XMLHttpRequest.upload`.
- Не удается использовать `ReadableStream` объекты в запросах.

> [!NOTE]
> Полная поддержка CORS доступна в Outlook в Интернете, Mac и Windows (начиная с версии 2201, сборка 16.0.14813.10000).

## <a name="see-also"></a>См. также

- [Манифесты надстроек Outlook](manifests.md)
- [Отладка надстроек на основе событий](debug-autolaunch.md)
- [Параметры описания AppSource для надстройки Outlook на основе событий](autolaunch-store-options.md)
- [Пошаговое руководство по интеллектуальным оповещениям и OnMessageSend](smart-alerts-onmessagesend-walkthrough.md)
- Примеры кода надстроек Office:
  - [Использование активации на основе событий Outlook для шифрования вложений, обработки участников приглашения на собрание и реагирования на изменения даты и времени встречи](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-encrypt-attachments)
  - [Использование активации Outlook на основе событий для задания подписи](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-set-signature)
  - [Использование активации Outlook на основе событий для пометки внешних получателей](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-tag-external)
  - [Использование интеллектуальных оповещений Outlook](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-check-item-categories)
