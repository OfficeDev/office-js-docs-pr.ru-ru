---
title: Включить общие папки и сценарии общих почтовых ящиков в Outlook надстройке
description: Обсуждается настройка поддержки надстройки для общих папок (ака). делегирования доступа) и общих почтовых ящиков.
ms.date: 10/05/2021
ms.localizationpriority: medium
---

# <a name="enable-shared-folders-and-shared-mailbox-scenarios-in-an-outlook-add-in"></a>Включить общие папки и сценарии общих почтовых ящиков в Outlook надстройке

В этой статье описывается, как включить в надстройке Outlook общие папки (также известные как доступ к делегатам[) и](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md#shared-mailboxes) общие почтовые ящики (в настоящее время в предварительном просмотре), в том числе разрешения, поддерживаемые API Office JavaScript.

## <a name="supported-clients-and-platforms"></a>Поддерживаемые клиенты и платформы

В следующей таблице показаны поддерживаемые клиенто-серверные комбинации для этой функции, включая минимальное необходимое накопительное обновление, если это применимо. Исключенные комбинации не поддерживаются.

| Клиент | Exchange Online | Exchange 2019 на месте<br>(Накопительное обновление 1 или более позднее) | Exchange 2016<br>(Накопительное обновление 6 или более позднее) | Exchange 2013 |
|---|:---:|:---:|:---:|:---:|
|Windows:<br>версия 1910 (сборка 12130.20272) или более поздней версии|Да|Нет|Нет|Нет|
|Mac:<br>сборка 16.47 или более поздней|Да|Да|Да|Да|
|Веб-браузер:<br>современный Outlook пользовательского интерфейса|Да|Неприменимо|Неприменимо|Неприменимо|
|Веб-браузер:<br>классический Outlook пользовательского интерфейса|Неприменимо|Нет|Нет|Нет|

> [!IMPORTANT]
> Поддержка этой функции была представлена в [наборе требований 1.8](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md) (подробные сведения см. [в отношении клиентов и платформ](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)). Однако обратите внимание, что матрица поддержки функции является суперсетью набора требований.

## <a name="supported-setups"></a>Поддерживаемые установки

В следующих разделах описываются поддерживаемые конфигурации для общих почтовых ящиков (теперь в предварительном просмотре) и общих папок. API-функции могут работать не так, как ожидалось в других конфигурациях. Выберите платформу, на которой необходимо научиться настраивать.

### <a name="windows"></a>[Windows](#tab/windows)

#### <a name="shared-folders"></a>Общие папки

Сначала владелец почтового ящика [должен предоставить доступ к делегату](https://support.microsoft.com/office/41c40c04-3bd1-4d22-963a-28eafec25926). Затем делегат должен следовать инструкциям, изложенным в разделе "Добавление почтового ящика другого человека в свой профиль" статьи Управление почтовыми и календарями другого [пользователя](https://support.microsoft.com/office/afb79d6b-2967-43b9-a944-a6b953190af5).

#### <a name="shared-mailboxes-preview"></a>Общие почтовые ящики (предварительный просмотр)

Exchange серверов администраторы могут создавать и управлять общими почтовыми ящиками для наборов пользователей для доступа. В настоящее [время Exchange Online](/exchange/collaboration-exo/shared-mailboxes) является единственной поддерживаемой серверной версией для этой функции.

Функция Exchange Server, известная как "автомаппирование", по умолчанию включается, что означает, что [](/microsoft-365/admin/email/create-a-shared-mailbox?view=o365-worldwide&preserve-view=true#add-the-shared-mailbox-to-outlook) после закрытия и открытия Outlook Outlook общего почтового ящика должен автоматически отображаться общий почтовый ящик. Однако если администратор отключил автомаппирование, пользователь должен следовать инструкциям, описанным в разделе "Добавление общего почтового ящика в Outlook" статьи Открыть и использовать общий почтовый [ящик в Outlook](https://support.microsoft.com/office/d94a8e9e-21f1-4240-808b-de9c9c088afd).

> [!WARNING]
> Не **входящие** в общий почтовый ящик с паролем. В этом случае API-функции не будут работать.

### <a name="web-browser---modern-outlook"></a>[Веб-браузер — современная версия Outlook](#tab/modern)

#### <a name="shared-folders"></a>Общие папки

Сначала владелец почтового ящика [должен предоставить доступ к делегату](https://www.microsoft.com/microsoft-365/blog/2013/09/04/configuring-delegate-access-in-outlook-web-app/) , обновив разрешения папок почтовых ящиков. Затем делегат должен следовать инструкциям, изложенным в разделе "Добавление почтового ящика другого человека в список папки Outlook Web App" статьи Доступ к почтовому ящику другого [человека](https://support.microsoft.com/office/a909ad30-e413-40b5-a487-0ea70b763081).

#### <a name="shared-mailboxes-preview"></a>Общие почтовые ящики (предварительный просмотр)

Exchange серверов администраторы могут создавать и управлять общими почтовыми ящиками для наборов пользователей для доступа. В настоящее [время Exchange Online](/exchange/collaboration-exo/shared-mailboxes) является единственной поддерживаемой серверной версией для этой функции.

После получения доступа общий пользователь почтового ящика должен следовать шагам, описанным в разделе "Добавьте общий почтовый ящик, чтобы он отображался в основном почтовом ящике" в статье [Open](https://support.microsoft.com/office/98b5a90d-4e38-415d-a030-f09a4cd28207) и использовать общий почтовый ящик в Outlook в Интернете.

> [!WARNING]
> Не **используйте** другие параметры, такие как "Откройте другой почтовый ящик". API-функции могут работать неправильно.

### <a name="mac"></a>[Mac](#tab/unix)

#### <a name="shared-mailboxes-preview"></a>Общие почтовые ящики (предварительный просмотр)

Почта и календарь делятся с делегатом или общим пользователем почтовых ящиков. Надстройки доступны делегату или пользователю в режимах чтения и записи сообщений и встреч.

#### <a name="shared-folders"></a>Общие папки

Если **папка "Входящие** " совместно с делегатом, надстройки доступны делегату в режиме чтения сообщений.

Если **папка Drafts** также совместно с делегатом, надстройки доступны в режиме составить.

#### <a name="local-shared-calendar-new-model"></a>Локальный общий календарь (новая модель)

Если владелец календаря явно поделился своим календарем с делегатом (весь почтовый ящик может не быть общим), надстройки доступны делегату в режимах чтения и записи записи.

#### <a name="remote-shared-calendar-previous-model"></a>Удаленный общий календарь (предыдущая модель)

Если владелец календаря предоставил широкий доступ к календарю (например, сделал его редактируемым для определенного DL или всей организации), пользователи могут иметь косвенное или неявное разрешение и надстройки доступны для этих пользователей в режимах чтения и записи записи.

---

Чтобы узнать больше о том, где надстройки делают и не активируются в целом, обратитесь к пунктам почтовых ящиков, доступным в разделе надстройки на странице Outlook обзор надстройки.[](outlook-add-ins-overview.md#mailbox-items-available-to-add-ins)

## <a name="supported-permissions"></a>Поддерживаемые разрешения

В следующей таблице описываются разрешения, Office API JavaScript для делегатов и общих пользователей почтовых ящиков.

|Разрешение|Значение|Описание|
|---|---:|---|
|Чтение|1 (000001)|Может читать элементы.|
|Запись|2 (000010)|Можно создавать элементы.|
|DeleteOwn|4 (000100)|Можно удалить только созданные элементы.|
|DeleteAll|8 (001000)|Может удалять любые элементы.|
|EditOwn|16 (010000)|Можно редактировать только созданные элементы.|
|EditAll|32 (100000)|Может изменять любые элементы.|

> [!NOTE]
> В настоящее время API поддерживает получение существующих разрешений, но не установку разрешений.

Объект [DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) реализуется с помощью битмаски для указать разрешения. Каждая позиция в битмаске `1` представляет определенное разрешение, и если оно заданной, у пользователя есть соответствующее разрешение. Например, если справа `1`находится второй бит, у пользователя есть разрешение **Напишите** . Пример проверки определенного разрешения в разделе [Выполнение](#perform-an-operation-as-delegate-or-shared-mailbox-user) операции в качестве делегата или общего пользователя почтового ящика см. в этой статье.

## <a name="sync-across-shared-folder-clients"></a>Синхронизация между общими клиентами папок

Обновления делегата в почтовом ящике владельца обычно синхронизируются между почтовыми ящиками немедленно.

Однако если операции REST или Exchange Web Services (EWS) использовались для набора расширенного свойства элемента, синхронизация таких изменений может занять несколько часов. Мы рекомендуем вместо этого использовать [объект CustomProperties](/javascript/api/outlook/office.customproperties) и связанные API, чтобы избежать такой задержки. Дополнительные статьи см. в [](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) разделе настраиваемые свойства в статье "Получить и установить метаданные в Outlook надстройки".

> [!IMPORTANT]
> В сценарии делегирования нельзя использовать EWS с маркерами, которые в настоящее время office.js API.

## <a name="configure-the-manifest"></a>Настройка манифеста

Чтобы включить общие папки и сценарии общих почтовых ящиков в надстройке, необходимо настроить элемент [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) `true` в манифесте под родительским элементом `DesktopFormFactor`. В настоящее время другие форм-факторы не поддерживаются.

Чтобы поддерживать вызовы REST от делегата, установите узел [Разрешений](../reference/manifest/permissions.md) в манифесте `ReadWriteMailbox`.

В следующем примере показан `SupportsSharedFolders` элемент, `true` установленный в разделе манифеста.

```XML
...
<VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
    ...
    <Hosts>
      <Host xsi:type="MailHost">
        <DesktopFormFactor>
          <SupportsSharedFolders>true</SupportsSharedFolders>
          <FunctionFile resid="residDesktopFuncUrl" />
          <ExtensionPoint xsi:type="MessageReadCommandSurface">
            <!-- configure selected extension point -->
          </ExtensionPoint>

          <!-- You can define more than one ExtensionPoint element as needed -->

        </DesktopFormFactor>
      </Host>
    </Hosts>
    ...
  </VersionOverrides>
</VersionOverrides>
...
```

## <a name="perform-an-operation-as-delegate-or-shared-mailbox-user"></a>Выполните операцию в качестве пользователя делегирования или общего почтового ящика

Общие свойства элемента можно получить в режиме Compose или Read, позвонив по методу [item.getSharedPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) . Это возвращает объект [SharedProperties](/javascript/api/outlook/office.sharedproperties) , который в настоящее время предоставляет разрешения пользователя, адрес электронной почты владельца, базовый URL-адрес API REST и целевой почтовый ящик.

В следующем примере показано, как получить общие свойства сообщения или встречи, проверить, есть ли у делегата или общего пользователя почтового  ящика разрешение на запись, и сделать вызов REST.

```js
function performOperation() {
  Office.context.mailbox.getCallbackTokenAsync({
      isRest: true
    },
    function (asyncResult) {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded && asyncResult.value !== "") {
        Office.context.mailbox.item.getSharedPropertiesAsync({
            // Pass auth token along.
            asyncContext: asyncResult.value
          },
          function (asyncResult1) {
            let sharedProperties = asyncResult1.value;
            let delegatePermissions = sharedProperties.delegatePermissions;

            // Determine if user can do the expected operation.
            // E.g., do they have Write permission?
            if ((delegatePermissions & Office.MailboxEnums.DelegatePermissions.Write) != 0) {
              // Construct REST URL for your operation.
              // Update <version> placeholder with actual Outlook REST API version e.g. "v2.0".
              // Update <operation> placeholder with actual operation.
              let rest_url = sharedProperties.targetRestUrl + "/<version>/users/" + sharedProperties.targetMailbox + "/<operation>";
  
              $.ajax({
                  url: rest_url,
                  dataType: 'json',
                  headers:
                  {
                    "Authorization": "Bearer " + asyncResult1.asyncContext
                  }
                }
              ).done(
                function (response) {
                  console.log("success");
                }
              ).fail(
                function (error) {
                  console.log("error message");
                }
              );
            }
          }
        );
      }
    }
  );
}
```

> [!TIP]
> В качестве делегата можно использовать REST для получения содержимого сообщения Outlook, прикрепленного к элементу Outlook [или групповой публикации](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post).

## <a name="handle-calling-rest-on-shared-and-non-shared-items"></a>Обработка вызовов REST для общих и не общих элементов

Если вы хотите вызвать операцию REST для элемента, является ли этот элемент общим, `getSharedPropertiesAsync` вы можете использовать API, чтобы определить, является ли элемент общим. После этого можно создать URL-адрес REST для операции с помощью соответствующего объекта.

```js
if (item.getSharedPropertiesAsync) {
  // In Windows, Mac, and the web client, this indicates a shared item so use SharedProperties properties to construct the REST URL.
  // Add-ins don't activate on shared items in mobile so no need to handle.

  // Perform operation for shared item.
} else {
  // In general, this is not a shared item, so construct the REST URL using info from the Call REST APIs article:
  // https://docs.microsoft.com/office/dev/add-ins/outlook/use-rest-api

  // Perform operation for non-shared item.
}
```

## <a name="limitations"></a>Ограничения

В зависимости от сценариев надстройки существует несколько ограничений, которые следует учитывать при обработке общих папок или общих ситуаций почтовых ящиков.

### <a name="message-compose-mode"></a>Режим композитации сообщений

В режиме композитации сообщений [getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#outlook-office-messagecompose-getsharedpropertiesasync-member(1)) не поддерживается в Outlook в Интернете или Windows, если не выполнены следующие условия.

а. **Делегирование доступа и общих папок**

1. Владелец почтового ящика запускает сообщение. Это может быть новое сообщение, ответ или форвард.
1. Затем сообщение сохраняется, а затем перемещается из собственной папки **Drafts** в папку, доступную делегату.
1. Делегат открывает черновик из общей папки, а затем продолжает сочинять.

б. **Общий почтовый ящик**

1. Пользователь общего почтового ящика запускает сообщение. Это может быть новое сообщение, ответ или форвард.
1. Затем они сэкономят сообщение из собственной папки **Drafts** в папку в общем почтовом ящике.
1. Другой пользователь общего почтового ящика открывает черновик из общего почтового ящика, а затем продолжает сочинять.

Теперь сообщение находится в общем контексте, и надстройки, поддерживаюные эти общие сценарии, могут получать общие свойства элемента. После отправки сообщения оно обычно находится в папке отправленных элементов отправители.

### <a name="rest-and-ews"></a>REST и EWS

Ваша надстройка может использовать REST `ReadWriteMailbox` , и необходимо установить разрешение надстройки, чтобы включить доступ REST к почтовому ящику владельца или к общему почтовому ящику, как это применимо. EWS не поддерживается.

### <a name="user-or-shared-mailbox-hidden-from-an-address-list"></a>Пользовательский или общий почтовый ящик, скрытый из списка адресов

Если администратор спрятал пользовательский или общий адрес почтового ящика из списка адресов, таких как глобальный список адресов (GAL), `Office.context.mailbox.item` затронутые почтовые элементы, открытые в отчете почтовых ящиков, как null. Например, если пользователь открывает почтовый элемент в общем почтовом ящике, скрытом от GAL, `Office.context.mailbox.item` то этот элемент почты является null.

## <a name="see-also"></a>См. также

- [Разрешить другим пользователям управлять почтой и календарем](https://support.microsoft.com/office/41c40c04-3bd1-4d22-963a-28eafec25926)
- [Общий доступ к календарю в Microsoft 365](https://support.microsoft.com/office/b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [Добавьте общий почтовый ящик в Outlook](/microsoft-365/admin/email/create-a-shared-mailbox?view=o365-worldwide&preserve-view=true#add-the-shared-mailbox-to-outlook)
- [Как заказать элементы манифеста](../develop/manifest-element-ordering.md)
- [Маска (вычисления)](https://en.wikipedia.org/wiki/Mask_(computing))
- [Операторы bitwise JavaScript](https://www.w3schools.com/js/js_bitwise.asp)