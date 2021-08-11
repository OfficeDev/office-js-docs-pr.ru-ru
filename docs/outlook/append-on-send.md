---
title: Реализация приложения-на-отправки в Outlook надстройки
description: Узнайте, как реализовать функцию приложения по отправке в Outlook надстройки.
ms.topic: article
ms.date: 02/01/2021
localization_priority: Normal
ms.openlocfilehash: c590c26f9c7f092702efe91dd4debea911af38fc165a0806435f10ed1192180f
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57095575"
---
# <a name="implement-append-on-send-in-your-outlook-add-in"></a>Реализация приложения-на-отправки в Outlook надстройки

К концу этого погона у вас будет надстройка Outlook, которая может вставить отказ при отправлении сообщения.

> [!NOTE]
> Поддержка этой функции была представлена в наборе требований 1.9. См [клиенты и платформы](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients), поддерживающие этот набор обязательных требований.

## <a name="set-up-your-environment"></a>Настройка среды

Выполните [Outlook,](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) который создает проект надстройки с генератором Yeoman для Office надстройки.

## <a name="configure-the-manifest"></a>Настройка манифеста

Чтобы включить функцию приложения по отправке в надстройке, необходимо включить разрешение в коллекцию `AppendOnSend` [ExtendedPermissions.](../reference/manifest/extendedpermissions.md)

В этом сценарии вместо выполнения функции при выборе кнопки Выполнить действие будет `action` запущена  `appendOnSend` функция.

1. В редакторе кода откройте проект быстрого запуска.

1. Откройте **manifest.xml** файл, расположенный в корне проекта.

1. Выберите весь узел (включая открытые и закрываемые теги) и замените его `<VersionOverrides>` на следующий XML.

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
            <DesktopFormFactor>
              <FunctionFile resid="Commands.Url" />
              <ExtensionPoint xsi:type="MessageComposeCommandSurface">
                <OfficeTab id="TabDefault">
                  <Group id="msgComposeGroup">
                    <Label resid="GroupLabel" />
                    <Control xsi:type="Button" id="msgComposeOpenPaneButton">
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
                        <FunctionName>appendDisclaimerOnSend</FunctionName>
                      </Action>
                    </Control>
                  </Group>
                </OfficeTab>
              </ExtensionPoint>

              <!-- Configure AppointmentOrganizerCommandSurface extension point to support
              append on sending a new appointment. -->

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
            <bt:Url id="JSRuntime.Url" DefaultValue="https://localhost:3000/runtime.js" />
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
        <ExtendedPermissions>
          <ExtendedPermission>AppendOnSend</ExtendedPermission>
        </ExtendedPermissions>
      </VersionOverrides>
    </VersionOverrides>
    ```

> [!TIP]
> Дополнительные информацию о манифестах для Outlook надстройки см. в Outlook [манифестах надстройки.](manifests.md)

## <a name="implement-append-on-send-handling"></a>Реализация обработки приложения по отправке

Далее реализуйте приложение в событии отправки.

> [!IMPORTANT]
> Если надстройка [ `ItemSend` ](outlook-on-send-addins.md)также реализует обработку событий при отправке с помощью, вызов обработчиком отправки возвращает ошибку, так как этот сценарий `AppendOnSendAsync` не поддерживается.

В этом случае при отправке пользователя будет реализовано приложение к элементу об отказе.

1. В том же проекте быстрого запуска откройте **файл ./src/commands/commands.js** в редакторе кода.

1. После `action` функции вставьте следующую функцию JavaScript.

    ```js
    function appendDisclaimerOnSend(event) {
      var appendText =
        '<p style = "color:blue"> <i>This and subsequent emails on the same topic are for discussion and information purposes only. Only those matters set out in a fully executed agreement are legally binding. This email may contain confidential information and should not be shared with any third party without the prior written agreement of Contoso. If you are not the intended recipient, take no action and contact the sender immediately.<br><br>Contoso Limited (company number 01624297) is a company registered in England and Wales whose registered office is at Contoso Campus, Thames Valley Park, Reading RG6 1WG</i></p>';  
      /**
        *************************************************************
         Ideal Usage - Call the getBodyType API. Use the coercionType
         it returns as the parameter value below.
        *************************************************************
      */
      Office.context.mailbox.item.body.appendOnSendAsync(
        appendText,
        {
          coercionType: Office.CoercionType.Html
        },
        function(asyncResult) {
          console.log(asyncResult);
        }
      );

      event.completed();
    }
    ```

1. В конце файла добавьте следующее утверждение.

    ```js
    g.appendDisclaimerOnSend = appendDisclaimerOnSend;
    ```

## <a name="try-it-out"></a>Проверка

1. Выполните указанную ниже команду в корневом каталоге своего проекта. При запуске этой команды начнется локальный веб-сервер, если он еще не запущен и ваша надстройка будет перегружена. 

    ```command&nbsp;line
    npm start
    ```

1. Создайте новое сообщение и добавьте себя в **строку To.**

1. Из меню ленты или переполнения выберите **Выполнить действие.**

1. Отправьте сообщение, а затем откройте его из папки **"Входящие"** или **"Отправленные** элементы", чтобы просмотреть приложенный отказ.

    ![Снимок экрана примера сообщения с приложением об отказе при отправке Outlook в Интернете.](../images/outlook-web-append-disclaimer.png)

## <a name="see-also"></a>См. также

[Манифесты надстроек Outlook](manifests.md)
