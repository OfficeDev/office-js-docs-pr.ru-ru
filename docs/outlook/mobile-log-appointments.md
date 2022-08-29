---
title: Запись заметок о встрече во внешнее приложение в мобильных надстройки Outlook
description: Узнайте, как настроить надстройку Outlook Mobile для записи заметок о встречах и других сведений во внешнее приложение.
ms.topic: article
ms.date: 08/11/2022
ms.localizationpriority: medium
ms.openlocfilehash: 310585d821f12bfd400b7b1eaf780ab756bf5a3f
ms.sourcegitcommit: 57258dd38507f791bbb39cbb01d6bbd5a9d226b9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/12/2022
ms.locfileid: "67320690"
---
# <a name="log-appointment-notes-to-an-external-application-in-outlook-mobile-add-ins"></a>Запись заметок о встрече во внешнее приложение в мобильных надстройки Outlook

Сохранение заметок о встречах и других сведений в приложении для управления отношениями с клиентами (CRM) или в приложении для создания заметок поможет вам отслеживать собрания, на которых вы присутствовали.

Из этой статьи вы узнаете, как настроить мобильную надстройку Outlook, чтобы пользователи могли фиксировать заметки и другие сведения о встречах в приложении CRM или приложении для создания заметок. В этой статье мы будем использовать вымышленного поставщика служб CRM с именем Contoso.

> [!IMPORTANT]
> Эта функция поддерживается только в Android с подпиской на Microsoft 365.

## <a name="set-up-your-environment"></a>Настройка среды

Краткое [руководство по созданию](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) проекта надстройки Outlook с помощью генератора Yeoman для надстроек Office.

## <a name="capture-and-view-appointment-notes"></a>Запись и просмотр заметок о встрече

Вы можете реализовать команду функции или область задач. Чтобы обновить надстройку, выберите вкладку для команды функции или области задач, а затем следуйте инструкциям.

# <a name="function-command"></a>[Команда функции](#tab/noui)

Этот параметр позволяет пользователю в журнале просматривать заметки и другие сведения о встречах при выборе команды функции на ленте.

### <a name="configure-the-manifest"></a>Настройка манифеста

Чтобы пользователи могли занося заметки о встрече в надстройку, необходимо настроить точку расширения [MobileLogEventAppointmentAttendee](/javascript/api/manifest/extensionpoint#mobilelogeventappointmentattendee) в манифесте в родительском элементе `MobileFormFactor`. Другие форм-факторы не поддерживаются.

1. В редакторе кода откройте проект быстрого запуска.

1. Откройте файл **manifest.xml** , расположенный в корневом каталоге проекта.

1. Выберите весь узел `<VersionOverrides>` (включая открытый и закрывающий теги) и замените его следующим XML-кодом. Обязательно замените все ссылки **на Contoso** сведениями своей компании.

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
              <ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
                <OfficeTab id="TabDefault">
                  <Group id="apptReadGroup">
                    <Label resid="residDescription"/>
                    <Control xsi:type="Button" id="apptReadOpenPaneButton">
                      <Label resid="residLabel"/>
                      <Supertip>
                        <Title resid="residLabel"/>
                        <Description resid="residTooltip"/>
                      </Supertip>
                      <Icon>
                        <bt:Image size="16" resid="icon-16"/>
                        <bt:Image size="32" resid="icon-32"/>
                        <bt:Image size="80" resid="icon-80"/>
                      </Icon>
                      <Action xsi:type="ExecuteFunction">
                        <FunctionName>logCRMEvent</FunctionName>
                      </Action>
                    </Control>
                  </Group>
                </OfficeTab>
              </ExtensionPoint>
            </DesktopFormFactor>
            <MobileFormFactor>
              <FunctionFile resid="residFunctionFile"/>
              <ExtensionPoint xsi:type="MobileLogEventAppointmentAttendee">
                <Control xsi:type="MobileButton" id="appointmentReadFunctionButton">
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
                    <FunctionName>logCRMEvent</FunctionName>
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
            <bt:Image id="icon-80" DefaultValue="https://contoso.com/assets/icon-80.png"/>
          </bt:Images>
          <bt:Urls>
            <bt:Url id="residFunctionFile" DefaultValue="https://contoso.com/commands.html"/>
          </bt:Urls>
          <bt:ShortStrings>
            <bt:String id="residDescription" DefaultValue="Log appointment notes and other details to Contoso CRM."/>
            <bt:String id="residLabel" DefaultValue="Log to Contoso CRM"/>
          </bt:ShortStrings>
          <bt:LongStrings>
            <bt:String id="residTooltip" DefaultValue="Log notes to Contoso CRM for this appointment."/>
          </bt:LongStrings>
        </Resources>
      </VersionOverrides>
    </VersionOverrides>
    ```

> [!TIP]
> Дополнительные сведения о манифестах для надстроек Outlook см. в манифестах надстроек [Outlook](manifests.md) и добавлении поддержки команд надстроек [для Outlook Mobile](add-mobile-support.md).

### <a name="capture-appointment-notes"></a>Запись заметок о встрече

В этом разделе описано, как надстройка может извлекать сведения о встрече, когда пользователь нажмет **кнопку "Журнал** ".

1. В том же проекте быстрого запуска откройте файл **./src/commands/commands.js** в редакторе кода.

1. Замените все содержимое файла **commands.js** следующим кодом JavaScript.

    ```js
    var event;

    Office.initialize = function (reason) {
      // Add any initialization code here.
    };

    function logCRMEvent(appointmentEvent) {
      event = appointmentEvent;
      console.log(`Subject: ${Office.context.mailbox.item.subject}`);
      Office.context.mailbox.item.body.getAsync(
        "html",
        { asyncContext: "This is passed to the callback" },
        function callback(result) {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            event.completed({ allowEvent: true });
          } else {
            console.error("Failed to get body.");
            event.completed({ allowEvent: false });
          }
        }
      );
    }

    // Register the function.
    Office.actions.associate("logCRMEvent", logCRMEvent);
    ```

Затем обновите файл **commands.html** для ссылки **наcommands.js**.

1. В том же проекте быстрого запуска откройте файл **./src/commands/commands.html** в редакторе кода.

1. Найдите и замените `<script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>` следующим кодом:

    ```html
    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>
    <script type="text/javascript" src="commands.js"></script>
    ```

### <a name="view-appointment-notes"></a>Просмотр заметок о встрече

**Метку кнопки** журнала можно переключить для отображения представления,  заданное настраиваемое свойство **EventLogged**, зарезервированное для этой цели. Когда пользователь нажмет кнопку **"** Вид", он сможет просмотреть записи в журнал для этой встречи.

Надстройка определяет интерфейс просмотра журналов. Например, заметки о встрече, зарегистрированные в журнале, можно отобразить в диалоговом окне, когда пользователь нажмет кнопку " **Вид** ". Дополнительные сведения об использовании диалогов см. в статье ["Использование API диалогов Office в надстройке Office"](../develop/dialog-api-in-office-add-ins.md).

Добавьте следующую функцию в **файл ./src/commands/commands.js**. Эта функция задает **настраиваемое свойство EventLogged** для текущего элемента встречи.

```js
function updateCustomProperties() {
  Office.context.mailbox.item.loadCustomPropertiesAsync(
    function callback(customPropertiesResult) {
      if (customPropertiesResult.status === Office.AsyncResultStatus.Succeeded) {
        let customProperties = customPropertiesResult.value;
        customProperties.set("EventLogged", true);
        customProperties.saveAsync(
          function callback(setSaveAsyncResult) {
            if (setSaveAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log("EventLogged custom property saved successfully.");
              event.completed({ allowEvent: true });
              event = undefined;
            }
          }
        );
      }
    }
  );
}
```

Затем вызовите его после того, как надстройка успешно занося в журнал заметки о встрече. Например, его можно вызвать из **logCRMEvent** , как показано в следующей функции.

```js
function logCRMEvent(appointmentEvent) {
  event = appointmentEvent;
  console.log(`Subject: ${Office.context.mailbox.item.subject}`);
  Office.context.mailbox.item.body.getAsync(
    "html",
    { asyncContext: "This is passed to the callback" },
    function callback(result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        // Replace `event.completed({ allowEvent: true });` with the following statement.
        updateCustomProperties();
      } else {
        console.error("Failed to get body.");
        event.completed({ allowEvent: false });
      }
    }
  );
}
```

### <a name="delete-the-appointment-log"></a>Удаление журнала встреч

Если вы хотите разрешить пользователям отменить ведение журнала или удалить записи о встречах, чтобы можно было сохранить журнал замены, у вас есть два варианта.

1. Используйте Microsoft Graph [, чтобы очистить объект настраиваемых свойств](/graph/api/resources/extended-properties-overview?view=graph-rest-1.0&preserve-view=true) , когда пользователь нажмет соответствующую кнопку на ленте.
1. Добавьте следующую функцию в **файл ./src/commands/commands.js** , чтобы очистить пользовательское свойство **EventLogged** для текущего элемента встречи.

    ```js
    function clearCustomProperties() {
      Office.context.mailbox.item.loadCustomPropertiesAsync(
        function callback(customPropertiesResult) {
          if (customPropertiesResult.status === Office.AsyncResultStatus.Succeeded) {
            var customProperties = customPropertiesResult.value;
            customProperties.remove("EventLogged");
            customProperties.saveAsync(
              function callback(removeSaveAsyncResult) {
                if (removeSaveAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
                  console.log("Custom properties cleared");
                  event.completed({ allowEvent: true });
                  event = undefined;
                }
              }
            );
          }
        }
      );
    }
    ```

Затем вызовите его, когда нужно очистить пользовательское свойство. Например, его можно вызвать из **logCRMEvent** , если не удалось установить журнал, как показано в следующей функции.

  ```js
  function logCRMEvent(appointmentEvent) {
    event = appointmentEvent;
    console.log(`Subject: ${Office.context.mailbox.item.subject}`);
    Office.context.mailbox.item.body.getAsync(
      "html",
      { asyncContext: "This is passed to the callback" },
      function callback(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          updateCustomProperties();
        } else {
          console.error("Failed to get body.");
          // Replace `event.completed({ allowEvent: false });` with the following statement.
          clearCustomProperties();
        }
      }
    );
  }
  ```

# <a name="task-pane"></a>[Области задач](#tab/taskpane)

Этот параметр позволяет пользователю в журнале просматривать свои заметки и другие сведения о встречах из области задач.

### <a name="configure-the-manifest"></a>Настройка манифеста

Чтобы пользователи могли занося заметки о встрече в надстройку, необходимо настроить точку расширения [MobileLogEventAppointmentAttendee](/javascript/api/manifest/extensionpoint#mobilelogeventappointmentattendee) в манифесте в родительском элементе `MobileFormFactor`. Другие форм-факторы не поддерживаются.

1. В редакторе кода откройте проект быстрого запуска.

1. Откройте файл **manifest.xml** , расположенный в корневом каталоге проекта.

1. Выберите весь узел `<VersionOverrides>` (включая открытый и закрывающий теги) и замените его следующим XML-кодом. Обязательно замените все ссылки **на Contoso** сведениями своей компании.

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
                <ExtensionPoint xsi:type="AppointmentAttendeeCommandSurface">
                  <OfficeTab id="TabDefault">
                    <Group id="apptReadGroup">
                      <Label resid="residDescription"/>
                      <Control xsi:type="Button" id="apptReadOpenPaneButton">
                        <Label resid="residLabel"/>
                        <Supertip>
                          <Title resid="residLabel"/>
                          <Description resid="residTooltip"/>
                        </Supertip>
                        <Icon>
                          <bt:Image size="16" resid="icon-16"/>
                          <bt:Image size="32" resid="icon-32"/>
                          <bt:Image size="80" resid="icon-80"/>
                        </Icon>
                        <Action xsi:type="ShowTaskpane">
                          <SourceLocation resid="Taskpane.Url"/>
                        </Action>
                      </Control>
                    </Group>
                  </OfficeTab>
                </ExtensionPoint>
              </DesktopFormFactor>
              <MobileFormFactor>
                <ExtensionPoint xsi:type="MobileLogEventAppointmentAttendee">
                  <Control xsi:type="MobileButton" id="appointmentReadFunctionButton">
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
                    <Action xsi:type="ShowTaskpane">
                      <SourceLocation resid="Taskpane.Url"/>
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
              <bt:Image id="icon-80" DefaultValue="https://contoso.com/assets/icon-80.png"/>
            </bt:Images>
            <bt:Urls>
              <bt:Url id="residFunctionFile" DefaultValue="https://contoso.com/commands.html"/>
              <bt:Url id="Taskpane.Url" DefaultValue="https://contoso.com/taskpane.html"/>
            </bt:Urls>
            <bt:ShortStrings>
              <bt:String id="residDescription" DefaultValue="Log appointment notes and other details to Contoso CRM."/>
              <bt:String id="residLabel" DefaultValue="Log to Contoso CRM"/>
            </bt:ShortStrings>
            <bt:LongStrings>
              <bt:String id="residTooltip" DefaultValue="Log notes to Contoso CRM for this appointment."/>
            </bt:LongStrings>
          </Resources>
        </VersionOverrides>
    </VersionOverrides>
    ```

> [!TIP]
> Дополнительные сведения о манифестах для надстроек Outlook см. в манифестах надстроек [Outlook](manifests.md) и добавлении поддержки команд надстроек [для Outlook Mobile](add-mobile-support.md).

### <a name="capture-appointment-notes"></a>Запись заметок о встрече

Из этого раздела вы узнаете, как отображать записи о встречах и другие сведения в области задач, когда пользователь нажмет кнопку **"Журнал** ".

1. В том же проекте быстрого запуска откройте файл **./src/taskpane/taskpane.js** редакторе кода.

1. Замените все содержимое файла **taskpane.js** следующим кодом JavaScript.

    ```js
    // Office is ready.
    Office.onReady(function () {
        getEventData();
      }
    );

    function getEventData() {
      console.log(`Subject: ${Office.context.mailbox.item.subject}`);
      Office.context.mailbox.item.body.getAsync(
        "html",
        function callback(result) {
          if (result.status === Office.AsyncResultStatus.Succeeded) {
            console.log("event logged successfully");
          } else {
            console.error("Failed to get body.");
          }
        }
      );
    }
    ```

Затем обновите файл **taskpane.html** для ссылки **наtaskpane.js**.

1. В том же проекте быстрого запуска откройте файл **./src/taskpane/taskpane.html** в редакторе кода.

1. Найдите и замените `<script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>` следующим кодом:

    ```html
    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>
    <script type="text/javascript" src="taskpane.js"></script>
    ```

### <a name="view-appointment-notes"></a>Просмотр заметок о встрече

**Метку кнопки** журнала можно переключить для отображения представления,  заданное настраиваемое свойство **EventLogged**, зарезервированное для этой цели. Когда пользователь нажмет кнопку **"** Вид", он сможет просмотреть записи в журнал для этой встречи. Надстройка определяет интерфейс просмотра журналов.

Добавьте следующую функцию в **./src/taskpane/taskpane.js**. Эта функция задает **настраиваемое свойство EventLogged** для текущего элемента встречи.

```js
function updateCustomProperties() {
  Office.context.mailbox.item.loadCustomPropertiesAsync(
    function callback(customPropertiesResult) {
      if (customPropertiesResult.status === Office.AsyncResultStatus.Succeeded) {
        let customProperties = customPropertiesResult.value;
        customProperties.set("EventLogged", true);
        customProperties.saveAsync(
          function callback(setSaveAsyncResult) {
            if (setSaveAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
              console.log("EventLogged custom property saved successfully.");
            }
          }
        );
      }
    }
  );
}
```

Затем вызовите его после того, как надстройка успешно занося в журнал заметки о встрече. Например, его можно вызвать из **getEventData** , как показано в следующей функции.

```js
function getEventData() {
  console.log(`Subject: ${Office.context.mailbox.item.subject}`);
  Office.context.mailbox.item.body.getAsync(
    "html",
    function callback(result) {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        console.log("event logged successfully");
        updateCustomProperties();
      } else {
        console.error("Failed to get body.");
      }
    }
  );
}
```

### <a name="delete-the-appointment-log"></a>Удаление журнала встреч

Если вы хотите разрешить пользователям отменить ведение журнала или удалить записи о встречах, чтобы можно было сохранить журнал замены, у вас есть два варианта.

1. Используйте Microsoft Graph [, чтобы очистить объект настраиваемых свойств](/graph/api/resources/extended-properties-overview?view=graph-rest-1.0&preserve-view=true) , когда пользователь нажмет соответствующую кнопку в области задач.
1. Добавьте следующую функцию в **файл ./src/taskpane/taskpane.js** , чтобы очистить пользовательское свойство **EventLogged** для текущего элемента встречи.

    ```js
    function clearCustomProperties() {
      Office.context.mailbox.item.loadCustomPropertiesAsync(
        function callback(customPropertiesResult) {
          if (customPropertiesResult.status === Office.AsyncResultStatus.Succeeded) {
            var customProperties = customPropertiesResult.value;
            customProperties.remove("EventLogged");
            customProperties.saveAsync(
              function callback(removeSaveAsyncResult) {
                if (removeSaveAsyncResult.status === Office.AsyncResultStatus.Succeeded) {
                  console.log("Custom properties cleared");
                }
              }
            );
          }
        }
      );
    }
    ```

Затем вызовите его, когда нужно очистить пользовательское свойство. Например, его можно вызвать из **getEventData** , если не удалось задать журнал, как показано в следующей функции.

  ```js
  function getEventData() {
    console.log(`Subject: ${Office.context.mailbox.item.subject}`);
    Office.context.mailbox.item.body.getAsync(
      "html",
      function callback(result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          console.log("event logged successfully");
          updateCustomProperties();
        } else {
          console.error("Failed to get body.");
          clearCustomProperties();
        }
      }
    );
  }
  ```

---

## <a name="test-and-validate"></a>Тестирование и проверка

1. Следуйте обычным рекомендациям [по проверке и проверке надстройки](testing-and-tips.md).
1. После [загрузки неопубликоваемой](sideload-outlook-add-ins-for-testing.md) надстройки в Outlook в Интернете Windows или Mac перезапустите Outlook на мобильном устройстве Android.
1. Откройте встречу в качестве участника, а затем убедитесь, что в карточке **"** Аналитика собраний" есть новая карточка с именем надстройки рядом с кнопкой **"Журнал** ".

### <a name="ui-log-the-appointment-notes"></a>Пользовательский интерфейс: запись заметок о встрече

Как участник собрания вы должны увидеть экран, аналогичный следующему изображению, при открытии собрания.

![Снимок экрана: кнопка "Журнал" на экране встречи в Android.](../images/outlook-android-log-appointment-details.jpg)

### <a name="ui-view-the-appointment-log"></a>Пользовательский интерфейс: просмотр журнала встреч

После успешного ведения журнала заметок о встрече кнопка должна иметь метку **"Вид** ", а не **"Журнал"**. Отобразится экран, аналогичный приведенному ниже.

![Снимок экрана: кнопка "Вид" на экране встречи в Android.](../images/outlook-android-view-appointment-log.jpg)

## <a name="available-apis"></a>Доступные API

Для этой функции доступны следующие API.

- [API диалоговых окон](../develop/dialog-api-in-office-add-ins.md)
- [Office.AddinCommands.Event](/javascript/api/office/office.addincommands.event?view=outlook-js-preview&preserve-view=true)
- [Office.CustomProperties](/javascript/api/outlook/office.customproperties?view=outlook-js-preview&preserve-view=true)
- [Office.RoamingSettings](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true)
- [API чтения встреч (участников),](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true) **за исключением** следующих:
  - [Office.context.mailbox.item.categories](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#categories)
  - [Office.context.mailbox.item.enhancedLocation](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#enhancedLocation)
  - [Office.context.mailbox.item.isAllDayEvent](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#isAllDayEvent)
  - [Office.context.mailbox.item.recurrence](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#recurrence)
  - [Office.context.mailbox.item.sensitivity](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#sensitivity)
  - [Office.context.mailbox.item.seriesId](/javascript/api/outlook/office.appointmentread?view=outlook-js-preview&preserve-view=true#seriesId)

## <a name="restrictions"></a>Ограничения

Применяется несколько ограничений.

- Имя **кнопки** журнала изменить нельзя. Однако существует способ отображения другой метки путем установки настраиваемого свойства для элемента встречи. Дополнительные сведения см. в разделе  "Просмотр заметок [](?tabs=noui#view-appointment-notes) к встрече" для соответствующей команды функции или [области](?tabs=taskpane#view-appointment-notes-1) задач.
- **Настраиваемое свойство EventLogged** необходимо использовать, если вы хотите переключить метку кнопки "Журнал"  в режим **просмотра и обратно**.
- Значок надстройки должен быть в оттенках серого с использованием шестнадцатеричных `#919191` кодов или его эквивалента в [других форматах цвета](https://convertingcolors.com/hex-color-919191.html).
- Надстройка должна извлечь сведения о собрании из формы встречи в течение минутного времени ожидания. Однако любое время, затраченное на диалоговое окно, которое надстройка, например, открыта для проверки подлинности, исключается из периода ожидания.

## <a name="see-also"></a>См. также

- [Надстройки для Outlook Mobile](outlook-mobile-addins.md)
- [Добавлена поддержка команд надстройки для Outlook Mobile](add-mobile-support.md)
