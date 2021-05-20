---
title: Настройте Outlook для активации на основе событий (предварительный просмотр)
description: Узнайте, как настроить Outlook для активации на основе событий.
ms.topic: article
ms.date: 05/18/2021
localization_priority: Normal
ms.openlocfilehash: 721f05e1c835e066744598ecb2bd416c6a6b0526
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555242"
---
# <a name="configure-your-outlook-add-in-for-event-based-activation-preview"></a>Настройте Outlook для активации на основе событий (предварительный просмотр)

Без функции активации на основе событий пользователь должен явно запустить надстройу для выполнения своих задач. Эта функция позволяет надстройки выполнять задачи на основе определенных событий, особенно для операций, которые применяются к каждому элементу. Вы также можете интегрироваться с функцией панели задач и пользовательского интерфейса.

К концу этого пошагового руководства, вы будете иметь надстройку, которая работает всякий раз, когда новый элемент создается и устанавливает тему.

> [!IMPORTANT]
> Эта функция поддерживается только для [предварительного](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) просмотра Outlook веб-сайтах и Windows с Microsoft 365 подпиской. Для получения более подробной информации [в этой статье узнайте, как просмотреть функцию активации на](#how-to-preview-the-event-based-activation-feature) основе событий.
>
> Поскольку функции предварительного просмотра могут быть изменения без предварительного уведомления, они не должны использоваться в производственных дополнениях.

## <a name="supported-events"></a>Поддерживаемые события

В настоящее время поддерживаются следующие мероприятия.

|Событие|Описание|Клиенты|
|---|---|---|
|`OnNewMessageCompose`|О составлении нового сообщения (включает ответ, ответьте все и вперед), но не о редактировании, например, проекта.|Windows, веб|
|`OnNewAppointmentOrganizer`|О создании новой встречи, но не о редактировании существующей.|Windows, веб|
|`OnMessageAttachmentsChanged`|При добавлении или удалении вложений при составлении сообщения.|Windows|
|`OnAppointmentAttachmentsChanged`|При добавлении или удалении вложений при составлении встречи.|Windows|
|`OnMessageRecipientsChanged`|При добавлении или удалении получателей при составлении сообщения.|Windows|
|`OnAppointmentAttendeesChanged`|При добавлении или удалении участников при составлении встречи.|Windows|
|`OnAppointmentTimeChanged`|При изменении даты/времени при составлении встречи.|Windows|
|`OnAppointmentRecurrenceChanged`|При добавлении, изменении или удалении деталей повторения при составлении записи на прием. Если дата/время изменены, `OnAppointmentTimeChanged` событие также будет уволено.|Windows|
|`OnInfoBarDismissClicked`|При увольнении уведомления при составлении сообщения или пункта назначения. Только надстройкое, добавляемое уведомление, будет уведомлено.|Windows|

## <a name="how-to-preview-the-event-based-activation-feature"></a>Как просмотреть функцию активации на основе событий

Мы приглашаем Вас опробовать функцию активации на основе событий! Сообщите нам о ваших сценариях и о том, как мы можем улучшить их, дав нам обратную связь GitHub **(см.** раздел Обратная связь в конце этой страницы).

Для просмотра этой функции:

- Для Outlook в Интернете:
  - [Настройте целевой релиз на Microsoft 365 арендатора.](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)
  - Ссылка  на бета-библиотеку на CDN ( https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) . Файл [определения типа для](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) компиляции TypeScript и IntelliSense найден на CDN [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts). Вы можете установить эти типы с `npm install --save-dev @types/office-js-preview` .
- Для Outlook на Windows:
  - Минимальная требуемая сборка составляет 16.0.14026.20000. Присоединяйтесь [к Office Insider для](https://insider.office.com) доступа к бета Office котейной версии.
  - Настройте реестр. Outlook включает в себя локательную копию производственной и бета-версии Office.js вместо загрузки с CDN. По умолчанию ссылается локал-производственная копия API. Чтобы перейти на локаную бета-копию api Outlook JavaScript, необходимо добавить эту запись реестра, в противном случае бета-API могут не быть найдены.
    1. Создание ключа `HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\Outlook\Options\WebExt\Developer` реестра.
    1. Добавьте `EnableBetaAPIsInJavaScript` именуемую запись и установите `1` значение. На приведенном ниже изображении показано, как должен выглядеть реестр.

        ![Скриншот редактора реестра с ключевым значением реестра EnableBetaAPIsInJavaScript](../images/outlook-beta-registry-key.png)

## <a name="set-up-your-environment"></a>Настройка среды

Завершите [Outlook,](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) который создает надстройки проекта с генератором Yeoman для Office дополнительных висел.

## <a name="configure-the-manifest"></a>Настройка манифеста

Для активации надстройок на основе событий необходимо настроить элемент [Runtimes и](../reference/manifest/runtimes.md) [точку расширения LaunchEvent](../reference/manifest/extensionpoint.md#launchevent-preview) `VersionOverridesV1_1` в узел манифеста. На данный `DesktopFormFactor` момент, является единственным поддерживаемым форм-фактором.

1. В редакторе кода откройте проект быстрого запуска.

1. Откройте **manifest.xml** файл, расположенный в корне вашего проекта.

1. Выберите весь `<VersionOverrides>` узел (включая открытые и близкие теги) и замените его следующим XML, а затем сохраните изменения.

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
              <!-- Events supported on the web and on Windows. -->
              <LaunchEvent Type="OnNewMessageCompose" FunctionName="onMessageComposeHandler"/>
              <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="onAppointmentComposeHandler"/>
              <!-- Events supported only on Windows. -->
              <LaunchEvent Type="OnMessageAttachmentsChanged" FunctionName="onMessageAttachmentsChangedHandler" />
              <LaunchEvent Type="OnAppointmentAttachmentsChanged" FunctionName="onAppointmentAttachmentsChangedHandler" />
              <LaunchEvent Type="OnMessageRecipientsChanged" FunctionName="onMessageRecipientsChangedHandler" />
              <LaunchEvent Type="OnAppointmentAttendeesChanged" FunctionName="onAppointmentAttendeesChangedHandler" />
              <LaunchEvent Type="OnAppointmentTimeChanged" FunctionName="onAppointmentTimeChangedHandler" />
              <LaunchEvent Type="OnAppointmentRecurrenceChanged" FunctionName="onAppointmentRecurrenceChangedHandler" />
              <LaunchEvent Type="OnInfoBarDismissClicked" FunctionName="onInfobarDismissClickedHandler" />
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

Outlook на Windows файл JavaScript, в то время как Outlook в Интернете использует HTML файл, который может ссылаться на тот же файл JavaScript. Вы должны предоставить ссылки на оба этих файла `Resources` в узел манифеста, как платформа Outlook в конечном итоге определяет, следует ли использовать HTML или JavaScript на основе Outlook клиента. Таким образом, чтобы настроить обработку событий, уведите расположение HTML в `Runtime` элементе, затем в `Override` его элементе ребенка предоставьте расположение файла JavaScript, вписанный или на который ссылается HTML.

> [!TIP]
> Чтобы узнать больше о манифестах Outlook дополнительных надстройок, [Outlook дополнительные дополнения.](manifests.md)

## <a name="implement-event-handling"></a>Реализация обработки событий

Вы должны реализовать обработку выбранных событий.

В этом сценарии вы добавите обработку для составления новых элементов.

1. С того же проекта быстрого запуска откройте файл **./src/commands/commands.js** в редакторе кода.

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

1. Добавьте следующий код JavaScript в конце файла.

    ```js
    // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
    Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
    Office.actions.associate("onAppointmentComposeHandler", onAppointmentComposeHandler);
    ```

1. Сохраните изменения.

## <a name="try-it-out"></a>Проверка

1. Выполните следующую команду в корневом каталоге своего проекта. После выполнения этой команды запустится локальный веб-сервер (если он еще не запущен) и будет загружена ваша неопубликованная надстройка.

    ```command&nbsp;line
    npm start
    ```

    > [!NOTE]
    > Если надстройка не была автоматически загружена, следуйте инструкциям [в Sideload Outlook надстройки для](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-manually) тестирования, чтобы вручную перезагрузить надстройку в Outlook.

1. Создайте новое сообщение в веб-версии Outlook.

    ![Скриншот окна сообщения в Outlook в Интернете с темой, установленной на compose](../images/outlook-web-autolaunch-1.png)

1. В Outlook на Windows, создать новое сообщение.

    ![Скриншот окна сообщения в Outlook на Windows с темой, установленной на compose](../images/outlook-win-autolaunch.png)

    > [!NOTE]
    > Если вы работаете надстройок от localhost и видите ошибку "Простите, мы не могли получить доступ *к "ваш-добавить в имя-здесь"*. Убедитесь, что у вас есть сетевое соединение. Если проблема продолжается, пожалуйста, повторите попытку позже.", Возможно, потребуется включить исключение из цикла.
    >
    > 1. Закройте Outlook.
    > 1. Откройте менеджера **задач и** убедитесь, **чтоmsoadfsb.exe** процесс не работает.
    > 1. Выполните следующую команду.
    >
    >    ```command&nbsp;line
    >    call %SystemRoot%\System32\CheckNetIsolation.exe LoopbackExempt -a -n=1_http___localhost_300004ACA5EC-D79A-43EA-AB47-E50E47DD96FC
    >    ```
    >
    > 1. Перезапустите Outlook.

## <a name="debug"></a>Debug

При внесении изменений в обработку событий запуска в надстройку следует знать, что:

- Если вы обновили манифест, [удалите надстройку, а](sideload-outlook-add-ins-for-testing.md#remove-a-sideloaded-add-in) затем загрузите ее снова.
- Если вы внесли изменения в файлы, не в которые был манифест, закройте и Outlook на Windows или обновите вкладку браузера, Outlook в Интернете.

При реализации собственной функциональности может потребоваться отладить свой код. Для получения рекомендаций о том, как отладить активацию надстройки на [Outlook](debug-autolaunch.md)основе событий, см.

Запись времени выполнения также доступна для этой функции на Windows. Для получения дополнительной информации [см.](../testing/runtime-logging.md#runtime-logging-on-windows)

## <a name="deploy-to-users"></a>Развертывание для пользователей

Вы можете развернуть дополнения на основе событий, загрузив манифест через Microsoft 365 администратора. На портале администратора расширьте раздел **Параметры навигационном** стекле, а затем выберите **интегрированные приложения.** На странице **Интегрированные приложения** выберите Upload **пользовательских приложений.**

![Скриншот страницы интегрированных приложений в центре Microsoft 365, включая Upload пользовательских приложений](../images/outlook-deploy-event-based-add-ins.png)

AppSource и магазины inclient: Возможность развертывания надстройок на основе событий или обновления существующих надстройок, включая функцию активации на основе событий, должна быть доступна в ближайшее время.

> [!IMPORTANT]
> Надстройки на основе событий ограничиваются только развертыванием, управляемым администратором. На данный момент пользователи не могут получить дополнения на основе событий из AppSource или inclient магазинов.

## <a name="event-based-activation-behavior-and-limitations"></a>Поведение активации на основе событий и ограничения

Ожидается, что обработчики дополнительных событий будут короткими, легкими и неинвазивными. После активации надстройки будут тайм-аут в течение примерно 300 секунд, максимальное время, разрешенное для запуска надстройок на основе событий. Чтобы сигнализировать о том, что надстройки завершили обработку события запуска, мы рекомендуем вам позвонить в метод связанному `event.completed` обработчику. (Обратите внимание, что код, `event.completed` включенный после выписки, не гарантируется запуск.) Каждый раз, когда срабатывает событие, срабатываемое с ручками надстройок, надстройка активируется и запускается связанный обработчик событий, а окно тайм-аута сбрасывается. Надстройка заканчивается после того, как она раз, или пользователь закрывает окно compose или отправляет элемент.

Если пользователь имеет несколько надстройок, подписавшихся на одно и то же событие, Outlook запускает надстройки в определенном порядке. В настоящее время только пять надстройок на основе событий могут активно работать.

Пользователь может переключиться или перейти от текущего элемента почты, где началось запуск надстройка. Запущенная надстройа завершит свою работу в фоновом режиме.

Некоторые Office.js API, которые изменяют или изменяют пользовательский интерфейс, не допускаются из надстройок на основе событий. Ниже приведены заблокированные API:

- В соответствии `OfficeRuntime.auth` с :
  - `getAccessToken`(Windows только)
- В соответствии `Office.context.auth` с :
  - `getAccessToken`
  - `getAccessTokenAsync`
- В соответствии `Office.context.mailbox` с :
  - `displayAppointmentForm`
  - `displayMessageForm`
  - `displayNewAppointmentForm`
  - `displayNewMessageForm`
- В соответствии `Office.context.mailbox.item` с :
  - `close`
- В соответствии `Office.context.ui` с :
  - `displayDialogAsync`
  - `messageParent`

## <a name="see-also"></a>См. также

- [Манифесты надстроек Outlook](manifests.md)
- [Как отладить дополнения на основе событий](debug-autolaunch.md)
