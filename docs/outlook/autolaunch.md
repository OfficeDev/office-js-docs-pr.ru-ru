---
title: Настройка надстройки Outlook для активации на основе событий
description: Узнайте, как настроить надстройку Outlook для активации на основе событий.
ms.topic: article
ms.date: 10/13/2022
ms.localizationpriority: medium
ms.openlocfilehash: ce2821ed5d226ff2c6a2b3c718d5711689523ac6
ms.sourcegitcommit: d402c37fc3388bd38761fedf203a7d10fce4e899
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/21/2022
ms.locfileid: "68664681"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation"></a>Настройка надстройки Outlook для активации на основе событий

Без функции активации на основе событий пользователь должен явным образом запустить надстройку для выполнения своих задач. Эта функция позволяет надстройке выполнять задачи на основе определенных событий, особенно для операций, применяемых к каждому элементу. Вы также можете выполнить интеграцию с областью задач и командами функций.

К концу этого пошагового руководства у вас будет надстройка, которая запускается каждый раз, когда создается новый элемент и задает тему.

> [!NOTE]
> Поддержка этой функции была реализована в наборе обязательных [элементов 1.10](/javascript/api/requirement-sets/outlook/requirement-set-1.10/outlook-requirement-set-1.10), а дополнительные события теперь доступны в последующих наборах обязательных элементов. Дополнительные сведения о минимальном наборе обязательных элементов события, а также о клиентах и платформах, которые его поддерживают, см. в разделе "Поддерживаемые события и наборы обязательных элементов, поддерживаемые серверами [Exchange и клиентами Outlook"](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients).[](#supported-events)
>
> Активация на основе событий не поддерживается в Outlook для iOS или Android.

## <a name="supported-events"></a>Поддерживаемые события

В следующей таблице перечислены доступные в настоящее время события и поддерживаемые клиенты для каждого события. При возникновении события `event` обработчик получает объект, который может содержать сведения, относящиеся к типу события. **Столбец Description** содержит ссылку на связанный объект, если это применимо.

|Каноническое имя события</br>и имя манифеста XML|Имя манифеста Teams|Описание|Минимальный набор требований и поддерживаемые клиенты|
|---|---|---|---|
|`OnNewMessageCompose`| newMessageComposeCreated |При создании нового сообщения (включает ответ, ответ всем и пересылка), но не при редактировании, например черновика.|[1.10](/javascript/api/requirement-sets/outlook/requirement-set-1.10/outlook-requirement-set-1.10)<br><br>— Windows<sup>1</sup><br>— Веб-браузер<br>— Новый пользовательский интерфейс Mac |
|`OnNewAppointmentOrganizer`|newAppointmentOrganizerCreated|При создании новой встречи, но не при редактировании существующей.|[1.10](/javascript/api/requirement-sets/outlook/requirement-set-1.10/outlook-requirement-set-1.10)<br><br>— Windows<sup>1</sup><br>— Веб-браузер<br>— Новый пользовательский интерфейс Mac |
|`OnMessageAttachmentsChanged`|messageAttachmentsChanged|При добавлении или удалении вложений при создании сообщения.<br><br>Объект данных для конкретного события: [AttachmentsChangedEventArgs](/javascript/api/outlook/office.attachmentschangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>— Windows<sup>1</sup><br>— Веб-браузер<br>— Новый пользовательский интерфейс Mac|
|`OnAppointmentAttachmentsChanged`|appointmentAttachmentsChanged|При добавлении или удалении вложений при создании встречи.<br><br>Объект данных для конкретного события: [AttachmentsChangedEventArgs](/javascript/api/outlook/office.attachmentschangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>— Windows<sup>1</sup><br>— Веб-браузер<br>— Новый пользовательский интерфейс Mac|
|`OnMessageRecipientsChanged`|messageRecipientsChanged|При добавлении или удалении получателей при создании сообщения.<br><br>Объект данных для конкретного события: [RecipientsChangedEventArgs](/javascript/api/outlook/office.recipientschangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>— Windows<sup>1</sup><br>— Веб-браузер<br>— Новый пользовательский интерфейс Mac|
|`OnAppointmentAttendeesChanged`|appointmentAttendeesChanged|При добавлении или удалении участников при создании встречи.<br><br>Объект данных для конкретного события: [RecipientsChangedEventArgs](/javascript/api/outlook/office.recipientschangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>— Windows<sup>1</sup><br>— Веб-браузер<br>— Новый пользовательский интерфейс Mac|
|`OnAppointmentTimeChanged`|appointmentTimeChanged|При изменении даты и времени при создании встречи.<br><br>Объект данных для конкретного события: [AppointmentTimeChangedEventArgs](/javascript/api/outlook/office.appointmenttimechangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>— Windows<sup>1</sup><br>— Веб-браузер<br>— Новый пользовательский интерфейс Mac|
|`OnAppointmentRecurrenceChanged`|appointmentRecurrenceChanged|При добавлении, изменении или удалении сведений о повторе при создании встречи. При изменении даты и времени `OnAppointmentTimeChanged` событие также будет срабатывать.<br><br>Объект данных для конкретного события: [RecurrenceChangedEventArgs](/javascript/api/outlook/office.recurrencechangedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>— Windows<sup>1</sup><br>— Веб-браузер<br>— Новый пользовательский интерфейс Mac|
|`OnInfoBarDismissClicked`|infoBarDismissClicked|При закрытии уведомления при создании сообщения или элемента встречи. Уведомление будет получать только надстройка, которая добавила уведомление.<br><br>Объект данных для конкретного события: [InfobarClickedEventArgs](/javascript/api/outlook/office.infobarclickedeventargs?view=outlook-js-1.11&preserve-view=true)|[1.11](/javascript/api/requirement-sets/outlook/requirement-set-1.11/outlook-requirement-set-1.11)<br><br>— Windows<sup>1</sup><br>— Веб-браузер<br>— Новый пользовательский интерфейс Mac|
|`OnMessageSend`|messageSending|При отправке элемента сообщения. Дополнительные сведения см. в [пошаговом руководстве по интеллектуальным оповещениям](smart-alerts-onmessagesend-walkthrough.md).|[1.12](/javascript/api/requirement-sets/outlook/requirement-set-1.12/outlook-requirement-set-1.12)<br><br>— Windows<sup>1</sup><br>— Веб-браузер|
|`OnAppointmentSend`|appointmentSending|При отправке элемента встречи. Дополнительные сведения см. в [пошаговом руководстве по интеллектуальным оповещениям](smart-alerts-onmessagesend-walkthrough.md).|[1.12](/javascript/api/requirement-sets/outlook/requirement-set-1.12/outlook-requirement-set-1.12)<br><br>— Windows<sup>1</sup><br>— Веб-браузер|
|`OnMessageCompose`|messageComposeOpened|При создании нового сообщения (включает ответ, ответ всем и пересылка) или редактировании черновика.|[1.12](/javascript/api/requirement-sets/outlook/requirement-set-1.12/outlook-requirement-set-1.12)<br><br>— Windows<sup>1</sup><br>— Веб-браузер|
|`OnAppointmentOrganizer`|appointmentOrganizerOpened|При создании новой встречи или изменении существующей.|[1.12](/javascript/api/requirement-sets/outlook/requirement-set-1.12/outlook-requirement-set-1.12)<br><br>— Windows<sup>1</sup><br>— Веб-браузер|

> [!NOTE]
> <sup>Для</sup> работы надстроек на основе событий в Outlook для Windows требуется как минимум Windows 10 версии 1903 (сборка 18362) или Windows Server 2019 версии 1903.

## <a name="set-up-your-environment"></a>Настройка среды

Выполните [краткое руководство outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) , которое создает проект надстройки с помощью генератора [Yeoman для надстроек Office](../develop/yeoman-generator-overview.md).

> [!NOTE]
> Если вы хотите использовать манифест Teams для надстроек [Office (](../develop/json-manifest-overview.md)предварительная версия), выполните альтернативное краткое руководство в Outlook с помощью манифеста [Teams (](../quickstarts/outlook-quickstart-json-manifest.md)предварительная версия), но пропустите все разделы после раздела **"** Попробовать".

## <a name="configure-the-manifest"></a>Настройка манифеста

Чтобы настроить манифест, выберите вкладку для типа манифеста, который вы используете.

# <a name="xml-manifest"></a>[XML-манифест](#tab/xmlmanifest)

Чтобы включить активацию надстройки на основе событий, необходимо настроить элемент [Runtimes](/javascript/api/manifest/runtimes) и точку расширения [LaunchEvent](/javascript/api/manifest/extensionpoint#launchevent) `VersionOverridesV1_1` в узле манифеста. Пока это `DesktopFormFactor` единственный поддерживаемый форм-фактор.

1. В редакторе кода откройте проект быстрого запуска.

1. Откройте файл **manifest.xml** , расположенный в корневом каталоге проекта.

1. Выберите весь узел **\<VersionOverrides\>** (включая открытый и закрывающий теги) и замените его следующим XML-кодом, а затем сохраните изменения.

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

В Outlook для Windows используется файл JavaScript, а в Outlook в Интернете и в новом пользовательском интерфейсе Mac — HTML-файл, который может ссылаться на тот же файл JavaScript. Необходимо предоставить ссылки `Resources` на оба этих файла в узле манифеста, так как платформа Outlook в конечном итоге определяет, следует ли использовать HTML или JavaScript на основе клиента Outlook. Таким образом, чтобы настроить обработку событий, укажите расположение HTML-кода **\<Runtime\>** в элементе, `Override` а затем в его дочернем элементе укажите расположение файла JavaScript, встроенного или на который ссылаются HTML.

# <a name="teams-manifest-developer-preview"></a>[Манифест Teams (предварительная версия для разработчиков)](#tab/jsonmanifest)

1. Откройте файл **manifest.json** .

1. Добавьте следующий объект в массив extensions.runtimes. Обратите внимание на указанные ниже особенности этой разметки.

   - Для параметра minVersion набора обязательных элементов почтового ящика задано значение "1.10 `OnNewMessageCompose` `OnNewAppointmentCompose` ", так как в таблице выше в этой статье указано, что это самая раннюю версию набора обязательных элементов, поддерживающих события и события.
   - Идентификатор среды выполнения имеет описательное имя "autorun_runtime".
   - Свойство code имеет дочернее свойство page, для которого задано значение HTML-файла, и дочернее свойство script, задающее файл JavaScript. Вы создадите или измените эти файлы на следующих шагах. Office использует одно из этих значений в зависимости от платформы.
       - Office в Windows выполняет обработчики событий в среде выполнения, доступной только для JavaScript, которая загружает файл JavaScript напрямую.
       - Office на Mac и в Интернете выполняют обработчики в среде выполнения браузера, которая загружает HTML-файл. Этот файл, в свою очередь `<script>` , содержит тег, который загружает файл JavaScript.
     Дополнительные сведения см. в [разделе "Среды выполнения" надстроек Office](../testing/runtimes.md).
   - Свойство lifetime имеет значение "short". Это означает, что среда выполнения запускается при активации одного из событий и завершает работу после завершения работы обработчика. (В некоторых редких случаях среда выполнения завершает работу до завершения работы обработчика. См [. раздел "Среды выполнения" в надстройки Office](../testing/runtimes.md).)
   - Существует два типа действий, которые могут выполняться в среде выполнения. На следующем шаге вы создадите функции, соответствующие этим действиям.

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

1. Добавьте следующий массив autoRunEvents в качестве свойства объекта в массив "extensions".

    ```json
    "autoRunEvents": [
    
    ]
    ```

1. Добавьте следующий объект в массив autoRunEvents. Свойство events сопоставляет обработчики с событиями, как описано в таблице выше в этой статье. Имена обработчиков должны совпадать с именами, используемыми в свойствах id объектов в массиве actions на предыдущем шаге.

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
> - Дополнительные сведения о средах выполнения в надстройке см. в разделе ["Среды выполнения" надстроек Office](../testing/runtimes.md).
> - Дополнительные сведения о манифестах для надстроек Outlook см. в манифестах [надстроек Outlook](manifests.md).

## <a name="implement-event-handling"></a>Реализация обработки событий

Необходимо реализовать обработку выбранных событий.

В этом сценарии вы добавите обработку для создания новых элементов.

1. В том же проекте быстрого запуска создайте папку с именем **launchevent** в **каталоге ./src** .

1. В **папке ./src/launchevent** создайте файл с именем **launchevent.js**.

1. Откройте файл **./src/launchevent/launchevent.js** редакторе кода и добавьте следующий код JavaScript.

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
> Windows: в настоящее время импорт не поддерживается в файле JavaScript, где реализована обработка активации на основе событий.

## <a name="update-the-commands-html-file"></a>Обновление HTML-файла команд

1. В **папке ./src/commands** **откройтеcommands.html.**

1. Непосредственно перед **закрывающего головного** тега (`</head>`) добавьте запись скрипта, чтобы включить код JavaScript для обработки событий.

    ```html
    <script type="text/javascript" src="../launchevent/launchevent.js"></script>
    ```

1. Сохраните изменения.

## <a name="update-webpack-config-settings"></a>Обновление настроек конфигурации webpack

1. Откройте **webpack.config.jsфайл** , найденный в корневом каталоге проекта, и выполните следующие действия.

1. Найдите `plugins` массив в объекте `config` и добавьте этот новый объект в начале массива.

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

1. Выполните следующие команды в корневом каталоге проекта. При запуске `npm start`запустится локальный веб-сервер (если он еще не запущен), а надстройка будет загружена неопубликованным приложением.

    ```command&nbsp;line
    npm run build
    ```

    ```command&nbsp;line
    npm start
    ```

    > [!NOTE]
    > Если ваша надстройка не была загружена неопубликованным автоматически, следуйте инструкциям в статье "Загрузка неопубликованных надстроек [Outlook](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-manually) " для тестирования, чтобы вручную загрузить неопубликованную надстройку в Outlook.

1. Создайте новое сообщение в веб-версии Outlook.

    ![Окно сообщения в Outlook в Интернете с темой, задаемой при составлении.](../images/outlook-web-autolaunch-1.png)

1. В Outlook в новом пользовательском интерфейсе Mac создайте новое сообщение.

    ![Окно сообщения в Outlook в новом пользовательском интерфейсе Mac с темой, задаемой при составлении.](../images/outlook-mac-autolaunch.png)

1. В Outlook для Windows создайте новое сообщение.

    ![Окно сообщения в Outlook для Windows с темой, задаемой при составлении.](../images/outlook-win-autolaunch.png)

## <a name="debug"></a>Отладка

При внесении изменений в обработку событий запуска в надстройке следует учитывать следующее:

- Если вы обновили манифест, [удалите надстройку](sideload-outlook-add-ins-for-testing.md#remove-a-sideloaded-add-in), а затем снова перезагрузите ее. Если вы используете Outlook в Windows, закройте и снова откройте Outlook.
- Если вы внесли изменения в файлы, отличные от манифеста, закройте и снова откройте Outlook в Windows или обновите вкладку браузера, на Outlook в Интернете.

При реализации собственных функций может потребоваться отладка кода. Инструкции по отладке активации надстройки на основе событий см. в разделе "Отладка надстройки Outlook на основе [событий"](debug-autolaunch.md).

Ведение журнала среды выполнения также доступно для этой функции в Windows. Дополнительные сведения см. в разделе ["Отладка надстройки с помощью ведения журнала среды выполнения"](../testing/runtime-logging.md#runtime-logging-on-windows).

[!INCLUDE [Loopback exemption note](../includes/outlook-loopback-exemption.md)]

## <a name="deploy-to-users"></a>Развертывание для пользователей

Вы можете развернуть надстройки на основе событий, передав манифест через Центр администрирования Microsoft 365. На портале администрирования разверните раздел **"Параметры** " в области навигации и выберите **"Интегрированные приложения"**. На странице **"Интегрированные приложения** " выберите действие **"Отправить пользовательские приложения** ".

![Страница "Интегрированные приложения" на Центр администрирования Microsoft 365 с выделенным действием "Отправить пользовательские приложения".](../images/outlook-deploy-event-based-add-ins.png)

> [!IMPORTANT]
> Надстройки на основе событий ограничены только развертываниями, управляемыми администратором. Пользователи не могут активировать надстройки на основе событий из AppSource или магазина Office в приложении. Дополнительные сведения см. в [описании параметров appSource для надстройки Outlook](autolaunch-store-options.md) на основе событий.

[!INCLUDE [outlook-smart-alerts-deployment](../includes/outlook-smart-alerts-deployment.md)]

## <a name="event-based-activation-behavior-and-limitations"></a>Поведение и ограничения активации на основе событий

Обработчики событий запуска надстройки должны быть короткими, упрощенными и неинициативными. После активации время ожидания надстройки будет истекает примерно в течение 300 секунд — максимального времени, допустимого для запуска надстроек на основе событий. Чтобы указать, что надстройка завершила обработку события запуска, связанный обработчик событий должен вызвать `event.completed` метод. (Обратите внимание, что код, включенный `event.completed` после выполнения инструкции, не гарантируется.) При каждом запуске события, срабатывающего с помощью дескрипторов надстройки, надстройка повторно активируется и запускает связанный обработчик событий, а время ожидания сбрасывается. Надстройка заканчивается по истечении времени ожидания, или пользователь закрывает окно создания или отправляет элемент.

Если у пользователя есть несколько надстроек, которые подписаны на одно и то же событие, платформа Outlook запускает надстройки в определенном порядке. В настоящее время можно активно запускать только пять надстроек на основе событий.

Во всех поддерживаемых клиентах Outlook пользователь должен оставаться в текущем почтовом элементе, где была активирована надстройка, чтобы завершить выполнение. Переход от текущего элемента (например, переход к другому окну создания или вкладке) завершает операцию надстройки. Надстройка также прекращает работу, когда пользователь отправляет сообщение или встречу, которые он создает.

Импорт не поддерживается в файле JavaScript, где реализована обработка активации на основе событий в клиенте Windows.

Некоторые Office.js API, которые изменяют или изменяют пользовательский интерфейс, не допускаются из надстроек на основе событий. Ниже приведены заблокированные API.

- В разделе `Office.context.auth`:
  - `getAccessToken`
  - `getAccessTokenAsync`
    > [!NOTE]
    > [OfficeRuntime.auth](/javascript/api/office-runtime/officeruntime.auth) поддерживается во всех версиях Outlook, поддерживающих активацию на основе событий и единый вход (SSO), в то время как [Office.auth](/javascript/api/office/office.auth) поддерживается только в некоторых сборках Outlook. Дополнительные сведения см. в разделе "Включение единого входа" в надстройки Outlook, использующих активацию [на основе событий](use-sso-in-event-based-activation.md).
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

Внешние данные можно запрашивать с помощью API, например [Fetch](https://developer.mozilla.org/docs/Web/API/Fetch_API) , или [С помощью XMLHttpRequest (XHR)](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest) — стандартного веб-API, который выполняет HTTP-запросы для взаимодействия с серверами.

Имейте в виду, что при использовании объектов XMLHttpRequest необходимо использовать дополнительные меры безопасности[](https://developer.mozilla.org/docs/Web/Security/Same-origin_policy), для которых требуется политика одинакового источника и простая [CORS (](https://developer.mozilla.org/docs/Web/HTTP/CORS)общий доступ к ресурсам независимо от источника).

[Простая реализация CORS](https://developer.mozilla.org/docs/Web/HTTP/CORS#simple_requests):

- Не может использовать файлы cookie.
- Поддерживаются только простые методы, такие `GET`как , и `HEAD``POST`.
- Принимает простые заголовки с именами полей `Accept`или `Accept-Language``Content-Language`.
- Может использовать , `Content-Type`при условии, что тип контента — `application/x-www-form-urlencoded`, `text/plain`или `multipart/form-data`.
- Не удается зарегистрировать прослушиватели событий для объекта, возвращаемого .`XMLHttpRequest.upload`
- Не может использовать объекты `ReadableStream` в запросах.

> [!NOTE]
> Полная поддержка CORS доступна в Outlook в Интернете, Mac и Windows (начиная с версии 2201, сборка 16.0.14813.10000).

## <a name="see-also"></a>См. также

- [Манифесты надстроек Outlook](manifests.md)
- [Отладка надстроек на основе событий](debug-autolaunch.md)
- [Параметры описания AppSource для надстройки Outlook на основе событий](autolaunch-store-options.md)
- [Пошаговое руководство по интеллектуальным оповещениям и OnMessageSend](smart-alerts-onmessagesend-walkthrough.md)
- Примеры кода надстроек Office:
  - [Использование активации на основе событий Outlook для шифрования вложений, обработки приглашений на собрание и реагирования на изменения даты и времени встречи](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-encrypt-attachments)
  - [Использование активации Outlook на основе событий для задания подписи](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-set-signature)
  - [Использование активации Outlook на основе событий для пометки внешних получателей](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-tag-external)
  - [Использование интеллектуальных оповещений Outlook](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/outlook-check-item-categories)
