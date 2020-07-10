---
title: Реализация добавления при отправке в надстройке Outlook (Предварительная версия)
description: Узнайте, как реализовать функцию "присоединение к передаче" в надстройке Outlook.
ms.topic: article
ms.date: 05/26/2020
localization_priority: Normal
ms.openlocfilehash: b9c834778d68e50806da908732cd0c8663ec6680
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093989"
---
# <a name="implement-append-on-send-in-your-outlook-add-in-preview"></a>Реализация добавления при отправке в надстройке Outlook (Предварительная версия)

По завершении этого пошагового руководства у вас будет надстройка Outlook, которая может вставить заявление об отказе при отправке сообщения.

> [!IMPORTANT]
> Эта функция в настоящее время поддерживается для [предварительной версии](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) в Outlook в Интернете и Windows с подпиской на Microsoft 365. Узнайте [, как выполнить предварительный просмотр функции присоединения при отправке](#how-to-preview-the-append-on-send-feature) в этой статье для получения дополнительных сведений.
>
> Так как функции предварительного просмотра могут быть изменены без предварительного уведомления, они не должны использоваться в производственных надстройках.

## <a name="how-to-preview-the-append-on-send-feature"></a>Предварительный просмотр функции присоединения при отправке

Мы приглашаем вас испытать функцию "дописывать от отправки"! Сообщите нам о своих сценариях и способах их усовершенствования, предоставив отзыв на сайте GitHub (обратитесь к разделу **Отзывы** в конце этой страницы).

Чтобы просмотреть эту функцию, выполните указанные ниже действия.

- Ссылка на **бета-** библиотеку в сети CDN ( https://appsforoffice.microsoft.com/lib/beta/hosted/office.js) . [Файл определения типа](https://appsforoffice.microsoft.com/lib/beta/hosted/office.d.ts) для компиляции TypeScript и IntelliSense находится в сети CDN и [DefinitelyTyped](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js-preview/index.d.ts). Вы можете установить эти типы с помощью `npm install --save-dev @types/office-js-preview` .
- Для Windows вы можете присоединиться к [программе предварительной оценки Office](https://insider.office.com) , чтобы получить доступ к последним сборкам Office.
- Для Outlook в Интернете [Настройте целевой выпуск на клиенте Microsoft 365](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center).

## <a name="set-up-your-environment"></a>Настройка среды

Завершите работу с [быстрым запуском Outlook](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) , который создает проект надстройки с помощью генератора Yeoman для надстроек Office.

## <a name="configure-the-manifest"></a>Настройка манифеста

Чтобы включить функцию Append-on-Send в надстройке, необходимо включить `AppendOnSend` разрешение в коллекцию [екстендедпермиссионс](../reference/manifest/extendedpermissions.md).

В этом сценарии вместо того, чтобы запускать `action` функцию при нажатии кнопки **выполнить действие** , вы заработаете `appendOnSend` функцию.

1. В редакторе кода откройте Быстрый запуск проекта.

1. Откройте файл **manifest.xml** , расположенный в корневом каталоге проекта.

1. Выберите весь `<VersionOverrides>` узел (включая открывающие и закрывающие теги) и замените его следующим XML-документом.

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
> Чтобы узнать больше о манифестах для надстроек Outlook, ознакомьтесь с разделом [манифесты надстроек Outlook](manifests.md).

## <a name="implement-append-on-send-handling"></a>Реализация обработки при отправке по требованию

Затем реализуйте Добавление в событие Send.

В этом сценарии вы реализуете Добавление заявления об отказе для элемента при отправке пользователя.

1. В проекте быстрого запуска откройте **commands.jsфайл./СРК/коммандс/** в редакторе кода.

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

1. В конце файла добавьте следующий оператор:

    ```js
    g.appendDisclaimerOnSend = appendDisclaimerOnSend;
    ```

## <a name="try-it-out"></a>Проверка

1. Выполните следующую команду в корневом каталоге своего проекта. При выполнении этой команды локальный веб-сервер запустится, если он еще не запущен.

    ```command&nbsp;line
    npm run dev-server
    ```

1. Следуйте инструкциям в статье [Загрузка неопубликованных надстройки Outlook для тестирования](sideload-outlook-add-ins-for-testing.md).

1. Создайте новое сообщение и добавьте себя в строку " **Кому** ".

1. В меню лента или переполнение выберите команду **выполнить действие**.

1. Отправьте сообщение, а затем откройте его в папке **"Входящие" или "** **Отправленные** ", чтобы просмотреть добавленное заявление об отказе.

    ![Снимок экрана с примером сообщения с сообщением об отказе, добавленном при отправке в Outlook в Интернете.](../images/outlook-web-append-disclaimer.png)

## <a name="see-also"></a>См. также

[Манифесты надстроек Outlook](manifests.md)
