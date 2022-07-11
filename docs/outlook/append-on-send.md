---
title: Реализация добавления при отправке в надстройке Outlook
description: Узнайте, как реализовать функцию добавления при отправке в надстройке Outlook.
ms.topic: article
ms.date: 07/07/2022
ms.localizationpriority: medium
ms.openlocfilehash: 762d8d14bb09d50c836b9a097534d1d23c493e66
ms.sourcegitcommit: d8ea4b761f44d3227b7f2c73e52f0d2233bf22e2
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/11/2022
ms.locfileid: "66712974"
---
# <a name="implement-append-on-send-in-your-outlook-add-in"></a>Реализация добавления при отправке в надстройке Outlook

К концу этого пошагового руководства у вас будет надстройка Outlook, которая может вставить заявление об отказе при отправке сообщения.

> [!NOTE]
> Поддержка этой функции реализована в наборе обязательных элементов 1.9. См [клиенты и платформы](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients), поддерживающие этот набор обязательных требований.

## <a name="set-up-your-environment"></a>Настройка среды

Выполните [краткое руководство outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) , которое создает проект надстройки с помощью генератора Yeoman для надстроек Office.

## <a name="configure-the-manifest"></a>Настройка манифеста

Чтобы включить функцию добавления при отправке в надстройке, `AppendOnSend` необходимо включить разрешение в коллекцию [ExtendedPermissions](/javascript/api/manifest/extendedpermissions).

В этом сценарии вместо того`action`, чтобы запускать функцию при  нажатии кнопки "Выполнить действие", вы будете запускать эту функцию`appendOnSend`.

1. В редакторе кода откройте проект быстрого запуска.

1. Откройте файл **manifest.xml** , расположенный в корневом каталоге проекта.

1. Выберите весь узел **\<VersionOverrides\>** (включая открытый и закрывающий теги) и замените его следующим XML-кодом.

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
> Дополнительные сведения о манифестах для надстроек Outlook см. в манифестах [надстроек Outlook](manifests.md).

## <a name="implement-append-on-send-handling"></a>Реализация обработки добавления при отправке

Затем реализуйте добавление к событию отправки.

> [!IMPORTANT]
> Если надстройка также реализует обработку событий при отправке с помощью, `AppendOnSendAsync` вызов обработчика при отправке возвращает ошибку, так как этот сценарий не поддерживается.[`ItemSend`](outlook-on-send-addins.md)

В этом сценарии вы реализуете добавление заявления об отказе от ответственности к элементу при отправке пользователем.

1. В том же проекте быстрого запуска откройте файл **./src/commands/commands.js** в редакторе кода.

1. После функции `action` вставьте следующую функцию JavaScript.

    ```js
    function appendDisclaimerOnSend(event) {
      const appendText =
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

1. Непосредственно под функцией добавьте следующую строку, чтобы зарегистрировать функцию.

    ```js
    Office.actions.associate("appendDisclaimerOnSend", appendDisclaimerOnSend);
    ```

## <a name="try-it-out"></a>Проверка

1. Выполните указанную ниже команду в корневом каталоге своего проекта. При выполнении этой команды локальный веб-сервер запустится, если он еще не запущен и надстройка будет загружена неопубликованным образом.

    ```command&nbsp;line
    npm start
    ```

1. Создайте новое сообщение и добавьте себя в **строку "К** ".

1. На ленте или в меню переполнения выберите " **Выполнить действие"**.

1. Отправьте сообщение, а затем откройте его из папки **"** Входящие" или "Отправленные", чтобы просмотреть добавленное заявление об отказе.

    ![Пример сообщения с заявлением об отказе, добавленным при отправке Outlook в Интернете.](../images/outlook-web-append-disclaimer.png)

## <a name="see-also"></a>См. также

[Манифесты надстроек Outlook](manifests.md)
