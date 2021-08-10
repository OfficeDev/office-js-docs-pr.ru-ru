---
title: Включить общие папки и сценарии общих почтовых ящиков в Outlook надстройке
description: Обсуждается настройка поддержки надстройки для общих папок (ака). делегирования доступа) и общих почтовых ящиков.
ms.date: 07/02/2021
localization_priority: Normal
ms.openlocfilehash: 9bcfaf77ecd837a39c9743d9194aa5e4ef30ba69a32c6caed41a38b8ab0ddb03
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57092352"
---
# <a name="enable-shared-folders-and-shared-mailbox-scenarios-in-an-outlook-add-in"></a>Включить общие папки и сценарии общих почтовых ящиков в Outlook надстройке

В этой статье описывается, как включить общие папки (также известные как доступ к делегатам) и общие почтовые ящики (в настоящее время в предварительном [просмотре)](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md#shared-mailboxes)сценарии в надстройке Outlook, в том числе разрешения, которые поддерживает API Office JavaScript.

> [!IMPORTANT]
> Поддержка этой функции была представлена в [наборе требований 1.8](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md). См [клиенты и платформы](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients), поддерживающие этот набор обязательных требований.

## <a name="supported-setups"></a>Поддерживаемые установки

В следующих разделах описываются поддерживаемые конфигурации для общих почтовых ящиков (теперь в предварительном просмотре) и общих папок. API-функции могут работать не так, как ожидалось в других конфигурациях. Выберите платформу, на которой необходимо научиться настраивать.

### <a name="windows"></a>[Windows](#tab/windows)

#### <a name="shared-folders"></a>Общие папки

Сначала владелец почтового ящика [должен предоставить доступ к делегату.](https://support.microsoft.com/office/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926) Затем делегат должен следовать инструкциям, изложенным в разделе "Добавление почтового ящика другого человека в свой профиль" статьи Управление почтовыми и календарями другого [пользователя.](https://support.microsoft.com/office/manage-another-person-s-mail-and-calendar-items-afb79d6b-2967-43b9-a944-a6b953190af5)

#### <a name="shared-mailboxes-preview"></a>Общие почтовые ящики (предварительный просмотр)

Exchange серверов администраторы могут создавать и управлять общими почтовыми ящиками для наборов пользователей для доступа. В настоящее [время Exchange Online](/exchange/collaboration-exo/shared-mailboxes) является единственной поддерживаемой серверной версией для этой функции.

Функция Exchange Server, известная как "автомаппирование", по умолчанию [](/microsoft-365/admin/email/create-a-shared-mailbox?view=o365-worldwide&preserve-view=true#add-the-shared-mailbox-to-outlook) включается, что означает, что после закрытия и открытия Outlook Outlook общего почтового ящика должен автоматически отображаться общий почтовый ящик. Однако если администратор отключил автомаппирование, пользователь должен следовать инструкциям, описанным в разделе "Добавление общего почтового ящика в Outlook" статьи Open и использовать общий почтовый ящик в [Outlook](https://support.microsoft.com/office/open-and-use-a-shared-mailbox-in-outlook-d94a8e9e-21f1-4240-808b-de9c9c088afd).

> [!WARNING]
> Не **входящие** в общий почтовый ящик с паролем. В этом случае API-функции не будут работать.

### <a name="web-browser---modern-outlook"></a>[Веб-браузер — современная версия Outlook](#tab/modern)

#### <a name="shared-folders"></a>Общие папки

Сначала владелец почтового ящика должен предоставить доступ к [делегату,](https://www.microsoft.com/microsoft-365/blog/2013/09/04/configuring-delegate-access-in-outlook-web-app/) обновив разрешения папок почтовых ящиков. Затем делегат должен следовать инструкциям, изложенным в разделе "Добавление почтового ящика другого человека в список папки в Outlook Web App" раздела статьи Доступ к почтовому ящику другого [человека.](https://support.microsoft.com/office/access-another-person-s-mailbox-a909ad30-e413-40b5-a487-0ea70b763081)

#### <a name="shared-mailboxes-preview"></a>Общие почтовые ящики (предварительный просмотр)

Exchange серверов администраторы могут создавать и управлять общими почтовыми ящиками для наборов пользователей для доступа. В настоящее [время Exchange Online](/exchange/collaboration-exo/shared-mailboxes) является единственной поддерживаемой серверной версией для этой функции.

После получения доступа общий пользователь почтового ящика должен следовать шагам, описанным в разделе "Добавьте общий почтовый ящик, чтобы он отображался в основном почтовом ящике" в статье Open и использовать общий почтовый ящик в [Outlook в Интернете](https://support.microsoft.com/office/open-and-use-a-shared-mailbox-in-outlook-on-the-web-98b5a90d-4e38-415d-a030-f09a4cd28207).

> [!WARNING]
> Не **используйте** другие параметры, такие как "Откройте другой почтовый ящик". API-функции могут работать неправильно.

---

Дополнительные сведения о том, где надстройки делают и [](outlook-add-ins-overview.md#mailbox-items-available-to-add-ins) не активируются в целом, обратитесь к пунктам почтовых ящиков, доступным в разделе надстройки на странице Outlook обзор надстройки.

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

Объект [DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) реализуется с помощью битмаски для указать разрешения. Каждая позиция в битмаске представляет определенное разрешение, и если оно заданной, у пользователя `1` есть соответствующее разрешение. Например, если справа находится второй `1` бит, у пользователя есть разрешение **Напишите.** Пример проверки определенного разрешения в разделе [Выполнение](#perform-an-operation-as-delegate-or-shared-mailbox-user) операции в качестве делегата или общего пользователя почтового ящика см. в этой статье.

## <a name="sync-across-shared-folder-clients"></a>Синхронизация между общими клиентами папок

Обновления делегата в почтовом ящике владельца обычно синхронизируются между почтовыми ящиками немедленно.

Однако если операции REST или Exchange Web Services (EWS) использовались для набора расширенного свойства элемента, синхронизация таких изменений может занять несколько часов. Мы рекомендуем вместо этого использовать [объект CustomProperties](/javascript/api/outlook/office.customproperties) и связанные API, чтобы избежать такой задержки. Дополнительные дополнительные [](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) статьи см. в разделе настраиваемые свойства в статье "Получить и установить метаданные в Outlook надстройки".

> [!IMPORTANT]
> В сценарии делегирования нельзя использовать EWS с маркерами, которые в настоящее время office.js API.

## <a name="configure-the-manifest"></a>Настройка манифеста

Чтобы включить общие папки и сценарии общих почтовых ящиков в надстройке, необходимо настроить элемент [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) в манифесте под `true` родительским элементом. `DesktopFormFactor` В настоящее время другие форм-факторы не поддерживаются.

Чтобы поддерживать вызовы REST от делегата, установите узел [Разрешений](../reference/manifest/permissions.md) в `ReadWriteMailbox` манифесте.

В следующем примере показан элемент, установленный `SupportsSharedFolders` `true` в разделе манифеста.

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

Общие свойства элемента можно получить в режиме Compose или Read, позвонив по методу [item.getSharedPropertiesAsync.](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) Это возвращает объект [SharedProperties,](/javascript/api/outlook/office.sharedproperties) который в настоящее время предоставляет разрешения пользователя, адрес электронной почты владельца, базовый URL-адрес API REST и целевой почтовый ящик.

В следующем примере показано, как получить общие свойства сообщения или встречи, проверить, есть ли у делегата или общего пользователя почтового ящика разрешение на запись, и сделать вызов REST. 

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
> В качестве делегата можно использовать REST для получения содержимого сообщения Outlook, прикрепленного к элементу Outlook [или групповой публикации.](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post)

## <a name="handle-calling-rest-on-shared-and-non-shared-items"></a>Обработка вызовов REST для общих и не общих элементов

Если вы хотите вызвать операцию REST для элемента, является ли этот элемент общим, вы можете использовать API, чтобы определить, является ли элемент `getSharedPropertiesAsync` общим. После этого можно создать URL-адрес REST для операции с помощью соответствующего объекта.

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

В режиме композитации сообщений [getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#getSharedPropertiesAsync_options__callback_) не поддерживается в Outlook в Интернете или Windows, если не выполнены следующие условия.

А. **Делегирование доступа и общих папок**

1. Владелец почтового ящика запускает сообщение. Это может быть новое сообщение, ответ или форвард.
1. Затем сообщение сохраняется, а затем перемещается из собственной папки **Drafts** в папку, доступную делегату.
1. Делегат открывает черновик из общей папки, а затем продолжает сочинять.

Б. **Общий почтовый ящик**

1. Пользователь общего почтового ящика запускает сообщение. Это может быть новое сообщение, ответ или форвард.
1. Затем они сэкономят сообщение из собственной папки **Drafts** в папку в общем почтовом ящике.
1. Другой пользователь общего почтового ящика открывает черновик из общего почтового ящика, а затем продолжает сочинять.

Теперь сообщение находится в общем контексте, и надстройки, поддерживаюные эти общие сценарии, могут получать общие свойства элемента. После отправки сообщения оно обычно находится в папке  отправленных элементов отправители.

### <a name="rest-and-ews"></a>REST и EWS

Ваша надстройка может использовать REST, и необходимо установить разрешение надстройки, чтобы включить доступ REST к почтовому ящику владельца или к общему почтовому ящику, как `ReadWriteMailbox` это применимо. EWS не поддерживается.

### <a name="user-or-shared-mailbox-hidden-from-an-address-list"></a>Пользовательский или общий почтовый ящик, скрытый из списка адресов

Если администратор спрятал пользовательский или общий адрес почтового ящика из списка адресов, таких как глобальный список адресов (GAL), затронутые почтовые элементы, открытые в отчете почтовых ящиков, как `Office.context.mailbox.item` null. Например, если пользователь открывает почтовый элемент в общем почтовом ящике, скрытом от GAL, то этот элемент почты является `Office.context.mailbox.item` null.

## <a name="see-also"></a>См. также

- [Разрешить другим пользователям управлять почтой и календарем](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)
- [Общий доступ к календарю в Microsoft 365](https://support.office.com/article/calendar-sharing-in-office-365-b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [Добавьте общий почтовый ящик в Outlook](/microsoft-365/admin/email/create-a-shared-mailbox?view=o365-worldwide&preserve-view=true#add-the-shared-mailbox-to-outlook)
- [Как заказать элементы манифеста](../develop/manifest-element-ordering.md)
- [Маска (вычисления)](https://en.wikipedia.org/wiki/Mask_(computing))
- [Операторы bitwise JavaScript](https://www.w3schools.com/js/js_bitwise.asp)