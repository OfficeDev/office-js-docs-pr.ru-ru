---
title: Используйте смарт-оповещения и событие OnMessageSend в Outlook надстройки (предварительный просмотр)
description: Узнайте, как обрабатывать событие отправки сообщений в Outlook надстройки с помощью активации на основе событий.
ms.topic: article
ms.date: 03/03/2022
ms.localizationpriority: medium
ms.openlocfilehash: dba12ba6ae667f3f5db740495a58ffc425d3aef3
ms.sourcegitcommit: 7b6ee73fa70b8e0ff45c68675dd26dd7a7b8c3e9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/08/2022
ms.locfileid: "63340850"
---
# <a name="use-smart-alerts-and-the-onmessagesend-event-in-your-outlook-add-in-preview"></a>Используйте смарт-оповещения и событие OnMessageSend в Outlook надстройки (предварительный просмотр)

В `OnMessageSend` этом событии вы можете использовать смарт-оповещения, которые позволяют запускать логику после выбора пользователем отправки Outlook сообщения. Обработник событий позволяет предоставить пользователям возможность улучшить свои электронные сообщения перед отправкой. Событие `OnAppointmentSend` аналогично, но применяется к встрече.

К концу этого погона у вас будет надстройка, которая запускается при отправке сообщения и проверяет, забыл ли пользователь добавить документ или фотографию, упомянутые в электронной почте.

> [!IMPORTANT]
> События `OnMessageSend` и `OnAppointmentSend` события доступны только в предварительной версии с подпиской Microsoft 365 в Outlook на Windows. Дополнительные сведения см. в [материале How to preview](autolaunch.md#how-to-preview). События предварительного просмотра не следует использовать в производственных надстройках.

## <a name="prerequisites"></a>Предварительные требования

Событие `OnMessageSend` доступно с помощью функции активации на основе событий. Чтобы понять, как настроить надстройку для использования этой функции, доступных событий, предварительного просмотра этого события, отладки, ограничений функций и других, обратитесь к настройке надстройки [Outlook](autolaunch.md) для активации на основе событий.

## <a name="set-up-your-environment"></a>Настройка среды

Выполните [Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator), который создает проект надстройки с генератором Yeoman для Office надстройки.

## <a name="configure-the-manifest"></a>Настройка манифеста

1. В редакторе кода откройте проект быстрого запуска.

1. Откройте файл **manifest.xml** , расположенный в корне проекта.

1. Выберите весь **узел VersionOverrides** (включая открытые и закрываемые теги) и замените его на следующий XML, а затем сохраните изменения.

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
               This is used by Outlook on the web and Outlook on the new Mac UI preview. -->
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

          <!-- Enable launching the add-in on the included event. -->
          <ExtensionPoint xsi:type="LaunchEvent">
            <LaunchEvents>
              <LaunchEvent Type="OnMessageSend" FunctionName="onMessageSendHandler" SendMode="PromptUser" />
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

> [!TIP]
>
> - Для **параметров SendMode** , доступных с событием `OnMessageSend` , обратитесь к [опциям Available SendMode](../reference/manifest/launchevent.md#available-sendmode-options-preview).
> - Дополнительные новости об манифестах для Outlook надстройки см. в Outlook [манифестах надстройки](manifests.md).

## <a name="implement-event-handling"></a>Реализация обработки событий

Для выбранного события необходимо реализовать обработку.

В этом сценарии будет добавлена обработка для отправки сообщения. Ваша надстройка проверит определенные ключевые слова в сообщении. Если какие-либо из этих ключевых слов найдены, он будет проверять, есть ли какие-либо вложения. Если вложений нет, надстройка будет рекомендовать пользователю добавить возможно отсутствующие вложения.

1. В этом же проекте быстрого запуска создайте новую папку с именем **launchevent** в **каталоге /src/** .

1. В **папке ./src/launchevent** создайте новый файл с **именемlaunchevent.js**.

1. Откройте файл **./src/launchevent/launchevent.js** в редакторе кода и добавьте следующий код JavaScript.

    ```js
    /*
    * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
    * See LICENSE in the project root for license information.
    */

    function onMessageSendHandler(event) {
      Office.context.mailbox.item.body.getAsync(
        "text",
        { "asyncContext": event },
        function (asyncResult) {
          let event = asyncResult.asyncContext;
          let body = "";
          let matches;
          if (asyncResult.status !== Office.AsyncResultStatus.Failed && asyncResult.value !== undefined) {
            body = asyncResult.value;
          }

          const arrayOfTerms = ["send", "picture", "document", "attachment"];
          for (let index = 0; index < arrayOfTerms.length; index++) {
            let term = arrayOfTerms[index].trim();
            const regex = RegExp(term, 'i');
            if (regex.test(body)) {
              matches.push(term);
            }
          }

          if (matches.length > 0) {
            // Let's verify if there's an attachment!
            Office.context.mailbox.item.getAttachmentsAsync(
              { "asyncContext": event },
              function(result) {
                let event = result.asyncContext;
                if (result.value.length <= 0) {
                  const message = "Looks like you're forgetting to include an attachment?";
                  event.completed({ allowEvent: false, errorMessage: message });
                } else {
                  for (let i = 0; i < result.value.length; i++) {
                    if (result.value[i].isInline == false) {
                      event.completed({ allowEvent: true });
                      return;
                    }
                  }
      
                  const message = "Looks like you forgot to include an attachment?";
                  event.completed({ allowEvent: false, errorMessage: message });
                }
              });
            } else {
              event.completed({ allowEvent: true });
            }
          }
        );
    }

    // 1st parameter: FunctionName of LaunchEvent in the manifest; 2nd parameter: Its implementation in this .js file.
    Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
    ```

1. Сохраните изменения.

> [!IMPORTANT]
> Windows. В настоящее время импорт не поддерживается в файле JavaScript, где выполняется обработка активации на основе событий.

## <a name="update-webpack-config-settings"></a>Обновление настроек конфигурации webpack

Откройте файл **webpack.config.js** , найденный в корневом каталоге проекта, и выполните следующие действия.

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

1. Запустите следующие команды в корневом каталоге проекта. При запуске `npm start`запустится локальный веб-сервер (если он еще не запущен) и надстройка будет перегружена.

    ```command&nbsp;line
    npm run build
    ```
    ```command&nbsp;line
    npm start
    ```

    > [!NOTE]
    > Если надстройка не была автоматически перегружена, следуйте инструкциям в [Sideload Outlook](../outlook/sideload-outlook-add-ins-for-testing.md#sideload-manually) надстройки для тестирования, чтобы вручную разгрузить надстройку в Outlook.

1. В Outlook на Windows создайте новое сообщение и установите тему. В теле добавьте текст типа "Эй, проверьте эту фотографию моей собаки!".
1. Отправка сообщения. Диалоговое окно должно всплывающее окно с рекомендацией для вас добавить вложение.
1. Добавьте вложение, а затем снова отправьте сообщение. В этот раз оповещения не должно быть.

[!INCLUDE [Loopback exemption note](../includes/outlook-loopback-exemption.md)]

## <a name="see-also"></a>См. также

- [Манифесты надстроек Outlook](manifests.md)
- [Настройка надстройки Outlook для активации на основе событий](autolaunch.md)
- [Отламывка надстроек на основе событий](debug-autolaunch.md)
- [Параметры списка AppSource для надстройки на Outlook событий](autolaunch-store-options.md)
