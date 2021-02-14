---
title: Включить сценарии делегирования доступа в надстройки Outlook
description: Кратко описывает делегирование доступа и описывает настройку поддержки надстройки.
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: 598f931dbf3a4be8adf029838084ec0767bf6518
ms.sourcegitcommit: fefc279b85e37463413b6b0e84c880d9ed5d7ac3
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/12/2021
ms.locfileid: "50234242"
---
# <a name="enable-delegate-access-scenarios-in-an-outlook-add-in"></a>Включить сценарии делегирования доступа в надстройки Outlook

Владелец почтового ящика может использовать функцию делегирования доступа, чтобы разрешить другим пользователям управлять своей почтой [и календарем.](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926) В этой статье указывается, какие разрешения делегирования поддерживает API JavaScript для Office, и описывается, как включить сценарии делегирования доступа в надстройки Outlook.

> [!IMPORTANT]
> Делегирование доступа в настоящее время не доступно в Outlook для Android и iOS. Кроме того, эта функция в настоящее время недоступна для общих почтовых ящиков групп [в](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide&preserve-view=true#shared-mailboxes) Outlook в Интернете. Эта функция может быть доступна в будущем.
>
> Поддержка этой функции была представлена в наборе требований 1.8. См [клиенты и платформы](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients), поддерживающие этот набор обязательных требований.

## <a name="supported-permissions-for-delegate-access"></a>Поддерживаемые разрешения для делегирования доступа

В следующей таблице описываются разрешения делегатов, поддерживаемые API JavaScript для Office.

|Разрешение|Значение|Описание|
|---|---:|---|
|Чтение|1 (000001)|Может читать элементы.|
|Запись|2 (000010)|Можно создавать элементы.|
|DeleteOwn|4 (000100)|Можно удалять только созданные элементы.|
|DeleteAll|8 (001000)|Может удалять любые элементы.|
|EditOwn|16 (010000)|Можно редактировать только созданные элементы.|
|EditAll|32 (100000)|Можно редактировать любые элементы.|

> [!NOTE]
> В настоящее время API поддерживает получение существующих разрешений делегата, но не настройку разрешений делегата.

Объект [DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) реализуется с помощью битовойmask, чтобы указать разрешения делегата. Каждая позиция в битовойmask представляет определенное разрешение, и если для него установлено соответствующее разрешение, делегат имеет `1` соответствующее разрешение. Например, если второй бит справа , делегат `1` имеет разрешение **на написание.** Пример проверки определенного разрешения см. в [](#perform-an-operation-as-delegate) разделе "Выполнение операции в качестве делегата" далее в этой статье.

## <a name="sync-across-mailbox-clients"></a>Синхронизация между клиентами почтовых ящиков

Обновления почтового ящика владельца делегата обычно синхронизируются между почтовыми ящиками немедленно.

Однако если операции REST или веб-служб Exchange (EWS) использовались для изменения расширенного свойства элемента, синхронизация таких изменений может занять несколько часов. Мы рекомендуем вместо этого использовать объект [CustomProperties](/javascript/api/outlook/office.customproperties) и связанные API, чтобы избежать такой задержки. Дополнительные см. [](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) в разделе пользовательских свойств статьи "Get and set metadata in an Outlook add-in".

> [!IMPORTANT]
> В сценарии делегирования нельзя использовать EWS с маркерами, которые в настоящее время предоставляются office.js API.

## <a name="configure-the-manifest"></a>Настройка манифеста

Чтобы включить сценарии делегирования доступа в надстройку, необходимо установить элемент [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) в манифесте `true` родительского `DesktopFormFactor` элемента. В настоящее время другие форм-факторы не поддерживаются.

Чтобы поддерживать вызовы REST от делегата, установите для узла ["Разрешения"](../reference/manifest/permissions.md) в манифесте разрешение `ReadWriteMailbox` .

В следующем примере показан `SupportsSharedFolders` элемент, установленный `true` в разделе манифеста.

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

## <a name="perform-an-operation-as-delegate"></a>Выполнение операции в качестве делегата

Общие свойства элемента можно получить в режиме compose или Read, вызывая метод [item.getSharedPropertiesAsync.](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) Возвращает объект [SharedProperties,](/javascript/api/outlook/office.sharedproperties) который в настоящее время предоставляет разрешения делегата, электронный адрес владельца, базовый URL-адрес API REST и целевой почтовый ящик.

В следующем примере показано, как получить общие свойства сообщения или встречи, проверить, есть ли у делегата разрешение **на** запись, и вызвать REST.

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
> В качестве делегата вы можете использовать REST для получения содержимого сообщения Outlook, вложенного в элемент [Outlook или публикацию в группе.](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post)

## <a name="handle-calling-rest-on-shared-and-non-shared-items"></a>Обработка вызова REST для общих и не общих элементов

Если вы хотите вызвать операцию REST для элемента, независимо от того, является ли элемент общим, вы можете использовать API, чтобы определить, является ли элемент `getSharedPropertiesAsync` общим. После этого можно создать URL-адрес REST для операции с помощью соответствующего объекта.

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

В зависимости от сценариев надстройки существует несколько ограничений, которые следует учитывать при обработке ситуаций делегатов.

### <a name="rest-and-ews"></a>REST и EWS

Надстройка может использовать REST, но не EWS, а разрешение надстройки должно быть настроено, чтобы разрешить доступ REST к почтовому ящику `ReadWriteMailbox` владельца.

### <a name="message-compose-mode"></a>Режим составить сообщение

В режиме составить сообщение [getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#getsharedpropertiesasync-options--callback-) не поддерживается в Outlook в Интернете или Windows, если не выполнены следующие условия.

1. Владелец делится хотя бы одной папкой почтового ящика с делегатом.
1. Делегат черновики сообщения в общей папке.

    Примеры:

    - Делегат отвечает на сообщение электронной почты в общей папке или переададает его.
    - Делегат сохраняет черновик сообщения, а затем  перемещает его из собственной папки "Черновики" в общую папку. Делегат открывает черновик из общей папки, а затем продолжает составление.

После того как сообщение было отправлено, оно обычно  находится в папке "Отправленные" представителя.

## <a name="see-also"></a>См. также

- [Разрешить другим пользователям управлять вашей почтой и календарем](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)
- [Общий доступ к календарю в Microsoft 365](https://support.office.com/article/calendar-sharing-in-office-365-b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [Порядок элементов манифеста](../develop/manifest-element-ordering.md)
- [Маска (вычисления)](https://en.wikipedia.org/wiki/Mask_(computing))
- [Операторы JavaScript по битовой стрелке](https://www.w3schools.com/js/js_bitwise.asp)