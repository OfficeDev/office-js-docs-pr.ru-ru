---
title: Реализация добавления при отправке в надстройке Outlook
description: Узнайте, как реализовать функцию добавления при отправке в надстройке Outlook.
ms.topic: article
ms.date: 10/24/2022
ms.localizationpriority: medium
ms.openlocfilehash: c8239634b6c9ca281255caf89276fb1b454efc84
ms.sourcegitcommit: 693e9a9b24bb81288d41508cb89c02b7285c4b08
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/28/2022
ms.locfileid: "68767163"
---
# <a name="implement-append-on-send-in-your-outlook-add-in"></a>Реализация добавления при отправке в надстройке Outlook

К концу этого пошагового руководства у вас будет надстройка Outlook, которая может вставить заявление об отказе от ответственности при отправке сообщения.

> [!NOTE]
> Поддержка этой функции появилась в наборе требований 1.9. См [клиенты и платформы](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients), поддерживающие этот набор обязательных требований.

## <a name="set-up-your-environment"></a>Настройка среды

Завершите [краткое руководство По созданию](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) проекта надстройки с помощью генератора Yeoman для надстроек Office.

## <a name="configure-the-manifest"></a>Настройка манифеста

Чтобы настроить манифест, откройте вкладку с типом манифеста, который вы используете.

# <a name="xml-manifest"></a>[XML-манифест](#tab/xmlmanifest)

Чтобы включить функцию добавления при отправке в надстройке, необходимо включить `AppendOnSend` разрешение в коллекцию [ExtendedPermissions](/javascript/api/manifest/extendedpermissions).

В этом сценарии вместо того, чтобы запускать функцию `action` при нажатии кнопки **Выполнить действие** , вы будете запускать функцию `appendOnSend` .

1. В редакторе кода откройте проект быстрого запуска.

1. Откройте **файлmanifest.xml** , расположенный в корне проекта.

1. Выберите весь **\<VersionOverrides\>** узел (включая открытые и закрытые теги) и замените его следующим XML-кодом.

    ```XML
    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
      <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
        <Requirements>
          <bt:Sets DefaultMinVersion="1.9">
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

# <a name="teams-manifest-developer-preview"></a>[Манифест Teams (предварительная версия для разработчиков)](#tab/jsonmanifest)

> [!IMPORTANT]
> Добавление при отправке пока не [поддерживается для манифеста Teams для надстроек Office (предварительная версия).](../develop/json-manifest-overview.md) Эта вкладка предназначена для использования в будущем.

1. Откройте файл manifest.json.

1. Добавьте следующий объект в массив extensions.runtimes. Обратите внимание на указанные ниже аспекты этого кода.

   - Параметр minVersion набора требований почтового ящика имеет значение "1.9", поэтому надстройку нельзя установить на платформах и версиях Office, где эта функция не поддерживается. 
   - Для идентификатора среды выполнения задается описательное имя "function_command_runtime".
   - Для свойства "code.page" задается URL-адрес HTML-файла без пользовательского интерфейса, который загрузит команду функции.
   - Свойство "время существования" имеет значение "short", что означает, что среда выполнения запускается при выборе кнопки команды функции и завершает работу после завершения функции. (В некоторых редких случаях среда выполнения завершает работу до завершения обработчика. См [. раздел Среды выполнения в надстройках Office](../testing/runtimes.md).)
   - Существует действие для запуска функции с именем "appendDisclaimerOnSend". Вы создадите эту функцию на следующем шаге.

    ```json
    {
        "requirements": {
            "capabilities": [
                {
                    "name": "Mailbox",
                    "minVersion": "1.9"
                }
            ],
            "formFactors": [
                "desktop"
            ]
        },
        "id": "function_command_runtime",
        "type": "general",
        "code": {
            "page": "https://localhost:3000/commands.html"
        },
        "lifetime": "short",
        "actions": [
            {
                "id": "appendDisclaimerOnSend",
                "type": "executeFunction",
                "displayName": "appendDisclaimerOnSend"
            }
        ]
    }
    ```

1. В массиве authorization.permissions.resourceSpecific добавьте следующий объект. Убедитесь, что он отделен от других объектов в массиве запятой.

    ```json
    {
      "name": "Mailbox.AppendOnSend.User",
      "type": "Delegated"
    }
    ```

---

> [!TIP]
> Дополнительные сведения о манифестах надстроек Outlook см. [в статье Манифесты надстроек Outlook](manifests.md).

## <a name="implement-append-on-send-handling"></a>Реализация обработки добавления при отправке

Затем реализуйте добавление в событие отправки.

> [!IMPORTANT]
> Если надстройка также реализует [обработку событий при отправке с помощью `ItemSend`](outlook-on-send-addins.md), вызов `AppendOnSendAsync` в обработчике при отправке возвращает ошибку, так как этот сценарий не поддерживается.

В этом сценарии вы реализуете добавление заявления об ответственности к элементу при отправке пользователем.

1. В том же проекте быстрого запуска откройте файл **./src/commands/commands.js** в редакторе кода.

1. После функции вставьте `action` следующую функцию JavaScript.

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

1. Выполните указанную ниже команду в корневом каталоге своего проекта. При выполнении этой команды локальный веб-сервер запустится, если он еще не запущен, и надстройка будет загружена неопубликованно.

    ```command&nbsp;line
    npm start
    ```

1. Создайте новое сообщение и добавьте себя в строку **Кому** .

1. В меню ленты или переполнения выберите **Выполнить действие**.

1. Отправьте сообщение, а затем откройте его из папки **"Входящие"** или **"Отправленные",** чтобы просмотреть добавленное заявление об отказе от ответственности.

    ![Пример сообщения с заявлением об отказе от ответственности, добавленным при отправке в Outlook в Интернете.](../images/outlook-web-append-disclaimer.png)

## <a name="see-also"></a>См. также

[Манифесты надстроек Outlook](manifests.md)
