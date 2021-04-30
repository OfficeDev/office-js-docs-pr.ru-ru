---
title: Настройка надстройки Outlook для активации на основе событий (предварительный просмотр)
description: Узнайте, как настроить Outlook надстройку для активации на основе событий.
ms.topic: article
ms.date: 04/29/2021
localization_priority: Normal
ms.openlocfilehash: 45f9ff16b3aca0a1fb8f3a8ee3d9ffa8e0f33ea2
ms.sourcegitcommit: 6057afc1776e1667b231d2e9809d261d372151f6
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/30/2021
ms.locfileid: "52100301"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation-preview"></a>Настройка надстройки Outlook для активации на основе событий (предварительный просмотр)

Без функции активации на основе событий пользователю необходимо явно запустить надстройки для выполнения задач. Эта функция позволяет надстройки выполнять задачи на основе определенных событий, особенно для операций, применимых к каждому элементу. Вы также можете интегрироваться с области задач и функциональными возможностями без пользовательского интерфейса. В настоящее время поддерживаются следующие события.

|Событие|Описание|
|---|---|
|`OnNewMessageCompose`|При составлении нового сообщения (включает ответ, ответ все и вперед), но не при редактировании, например, черновика.|
|`OnNewAppointmentOrganizer`|О создании новой встречи, но не о редактировании существующего.|
|`OnMessageAttachmentsChanged`|При добавлении или удалении вложений при сочинении сообщения.|
|`OnAppointmentAttachmentsChanged`|При добавлении или удалении вложений при записи на прием.|
|`OnMessageRecipientsChanged`|При добавлении или удалении получателей при сочинении сообщения.|
|`OnAppointmentAttendeesChanged`|При добавлении или удалении участников при записи на прием.|
|`OnAppointmentTimeChanged`|При изменении даты и времени при записи на прием.|
|`OnAppointmentRecurrenceChanged`|При добавлении, изменении или удалении сведений о повторении при записи на прием. Если дата и время изменены, `OnAppointmentTimeChanged` событие также будет уволено.|
|`OnInfoBarDismissClicked`|При отклонении уведомления при записи сообщения или элемента встречи. Уведомления будут получать только надстройка, которая добавила уведомление.|

К концу этого погона у вас будет надстройка, которая запускается всякий раз, когда создается новый элемент и задает объект.

> [!IMPORTANT]
> Эта функция поддерживается [](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) только для предварительного просмотра Outlook в Интернете и Windows с Microsoft 365 подпиской. Дополнительные [сведения см. в](#how-to-preview-the-event-based-activation-feature) статье Как просмотреть функцию активации на основе событий.
>
> Так как функции предварительного просмотра могут изменяться без предварительного уведомления, их не следует использовать в надстройки производства.

## <a name="how-to-preview-the-event-based-activation-feature"></a>Просмотр функции активации на основе событий

Мы приглашаем вас попробовать функцию активации на основе событий! Дайте нам знать о ваших сценариях и о том, как мы можем улучшить ситуацию, GitHub с помощью GitHub (см. раздел **Обратная** связь в конце этой страницы).

Чтобы просмотреть эту функцию:

- Для Outlook в Интернете:
  - [Настройка целевого выпуска для](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)Microsoft 365 клиента.
  - Ссылка  на бета-библиотеку на CDN ( https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) . Файл [определения типа](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) для компиляции и IntelliSense typeScript CDN и [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts). Эти типы можно установить с `npm install --save-dev @types/office-js-preview` помощью .
- Для Outlook на Windows: минимальная требуемая сборка — 16.0.13729.20000. Присоединяйтесь [к Office программы insider](https://insider.office.com) для доступа к Office бета-сборки.

## <a name="set-up-your-environment"></a>Настройка среды

Выполните [Outlook,](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) который создает проект надстройки с генератором Yeoman для Office надстройки.

## <a name="configure-the-manifest"></a>Настройка манифеста

Чтобы включить активацию надстройки на основе событий, необходимо настроить элемент [Runtimes](../reference/manifest/runtimes.md) и точку расширения [LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) в узле `VersionOverridesV1_1` манифеста. Пока это `DesktopFormFactor` единственный поддерживаемый форм-фактор.

1. В редакторе кода откройте проект быстрого запуска.

1. Откройте **manifest.xml** файл, расположенный в корне проекта.

1. Выберите весь узел (включая открытые и закрываемые теги) и замените его на следующий XML, а затем `<VersionOverrides>` сохраните изменения.

```XML
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    <Requirements>
      <bt:Sets DefaultMinVersion="1.3">
        <bt:Set Name="Mailbox" />
      </bt:Sets>
    </Requirements>
    <Hosts>
      <Host xsi:type="MailHost">
        <!-- Event-based activation happens in a lightweight runtime.-->
        <Runtimes>
          <!-- HTML file including reference to or inline JavaScript event handlers.
               This is used by Outlook on the web. -->
          <Runtime resid="WebViewRuntime.Url">
            <!-- JavaScript file containing event handlers. This is used by Outlook Desktop. -->
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
              <LaunchEvent Type="OnNewMessageCompose" FunctionName="onMessageComposeHandler"/>
              <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="onAppointmentComposeHandler"/>
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
        <!-- Entry needed for Outlook Desktop. -->
        <bt:Url id="JSRuntime.Url" DefaultValue="https://localhost:3000/src/commands/commands.js" />
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

Outlook на Windows использует файл JavaScript, а Outlook в Интернете использует HTML-файл, который может ссылаться на один и тот же файл JavaScript. Необходимо предоставить ссылки на оба этих файла в узле манифеста, так как платформа Outlook в конечном счете определяет, следует ли использовать HTML или JavaScript на основе Outlook `Resources` клиента. Таким образом, чтобы настроить обработку событий, укадь расположение HTML в элементе, а затем в его детском элементе укаймляй расположение файла JavaScript, вписаного или ссылаемого `Runtime` `Override` HTML.

> [!TIP]
> Дополнительные информацию о манифестах для Outlook надстройки см. в Outlook [манифестах надстройки.](manifests.md)

## <a name="implement-event-handling"></a>Реализация обработки событий

Для выбранных событий необходимо реализовать обработку.

В этом сценарии вы добавим обработку для составления новых элементов.

1. В том же проекте быстрого запуска откройте **файл ./src/commands/commands.js** в редакторе кода.

1. После `action` функции вставьте следующие функции JavaScript.

    ```js
    function onMessageComposeHandler(event) {
      setSubject(event);
    }
    function onAppointmentComposeHandler(event) {
      setSubject(event);
    }
    function setSubject(event) {
      Office.context.mailbox.item.subject.setAsync(
        "Set by an event-based add-in!",
        {
          "asyncContext" : event
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
    ```

1. Чтобы функции работали  Outlook в Интернете с этим проектом, созданным генератором Yeoman для Office надстройки, добавьте следующие утверждения в конце файла.

    ```js
    g.onMessageComposeHandler = onMessageComposeHandler;
    g.onAppointmentComposeHandler = onAppointmentComposeHandler;
    ```

1. Чтобы функции работали **в Outlook Windows, добавьте** следующий код JavaScript в конце файла.

    ```js
    if (Office.actions) {
      // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
      Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
      Office.actions.associate("onAppointmentComposeHandler", onAppointmentComposeHandler);
    }
    ```

    **Примечание.** Проверка `Office.actions` на то, что Outlook в Интернете игнорирует эти утверждения.

1. Сохраните изменения.

## <a name="try-it-out"></a>Проверка

1. Выполните следующую команду в корневом каталоге своего проекта. После выполнения этой команды запустится локальный веб-сервер (если он еще не запущен) и будет загружена ваша неопубликованная надстройка.

    ```command&nbsp;line
    npm start
    ```

1. Создайте новое сообщение в веб-версии Outlook.

    ![Снимок экрана окна сообщения в Outlook веб-страницы с набором субъекта на композит](../images/outlook-web-autolaunch-1.png)

1. В Outlook Windows создайте новое сообщение.

    ![Снимок экрана окна сообщения в Outlook Windows с набором субъекта на композицию](../images/outlook-win-autolaunch.png)

    > [!NOTE]
    > Если вы видите ошибку "Мы не можем открыть эту надстройку из localhost", необходимо включить освобождение от циклов.
    >
    > 1. Закройте Outlook.
    > 2. Откройте диспетчер **задач** и убедитесь, что **msoadfs.exe** процесс не запущен.
    > 3. Выполните следующую команду.
    >
    >     ```command&nbsp;line
    >     call %SystemRoot%\System32\CheckNetIsolation.exe LoopbackExempt -a -n=1_http___localhost_300004ACA5EC-D79A-43EA-AB47-E50E47DD96FC
    >     ```
    >
    > 4. Перезапустите Outlook.

## <a name="debug"></a>Debug

При реализации собственных функций может потребоваться отламывка кода. Инструкции по отламывить активацию надстройки на основе событий см. в Outlook [событиях.](debug-autolaunch.md)

## <a name="event-based-activation-behavior-and-limitations"></a>Поведение и ограничения активации на основе событий

Надстройки, которые активируются на основе событий, как ожидается, будут короткими, легкими и максимально неинвазивными. Чтобы сигнализировать, что надстройка завершила обработку события запуска, рекомендуется использовать метод вызова `event.completed` надстройки. Если этот вызов не будет выполнен, надстройка будет работать в течение примерно 300 секунд, что является максимальным сроком, разрешенным для запуска надстроек на основе событий. Надстройка также заканчивается, когда пользователь закрывает окно записи.

Если у пользователя есть несколько надстройок, которые подписаны на одно и то же событие, Outlook платформа запускает надстройки без определенного порядка. В настоящее время можно активно запускать только пять надстройок на основе событий. Все дополнительные надстройки отодвигаются в очередь, а затем запускаются по мере завершения или отключения ранее активных надстроек.

Пользователь может переключаться или перемещаться от текущего элемента почты, где надстройка начала работать. Запущенная надстройка завершит свою работу в фоновом режиме.

Некоторые Office.js API, которые изменяют или изменяют пользовательский интерфейс, не допускаются из надстройок на основе событий. Следующие API заблокированы:

- В `Office.context.auth` статье:
  - `getAccessToken`
  - `getAccessTokenAsync`
- В `Office.context.mailbox` статье:
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- В `Office.context.mailbox.item` статье:
  - `close`
- В `Office.context.ui` статье:
  - `displayDialogAsync`
  - `messageParent`

## <a name="see-also"></a>См. также

[Outlook манифесты надстройки](manifests.md) 
 [Отламывка](debug-autolaunch.md) надстроек на основе событий
