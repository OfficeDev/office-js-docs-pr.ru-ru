---
title: Включить сценарии делегирования доступа в Outlook надстройки
description: Кратко описывает делегатский доступ и рассказывает о настройке поддержки надстройки.
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: 256c37087b10eaf9c8025e19a4990852f9550458
ms.sourcegitcommit: 17b5a076375bc5dc3f91d3602daeb7535d67745d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/06/2021
ms.locfileid: "52783493"
---
# <a name="enable-delegate-access-scenarios-in-an-outlook-add-in"></a>Включить сценарии делегирования доступа в Outlook надстройки

Владелец почтового ящика может использовать функцию доступа к делегатам, чтобы позволить другому человеку управлять [своей почтой и календарем.](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926) В этой статье указывается, какие разрешения делегировать Office API JavaScript поддерживает, и описывается, как включить сценарии делегирования доступа Outlook надстройки.

> [!IMPORTANT]
> В настоящее время доступ к делегированию Outlook на Android и iOS. Кроме того, эта функция [](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide&preserve-view=true#shared-mailboxes) в настоящее время недоступна для групповых общих почтовых ящиков Outlook в Интернете. Эта функция может быть доступна в будущем.
>
> Поддержка этой функции была представлена в наборе требований 1.8. См [клиенты и платформы](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients), поддерживающие этот набор обязательных требований.

## <a name="supported-permissions-for-delegate-access"></a>Поддерживаемые разрешения для доступа к делегированию

В следующей таблице описываются разрешения делегатов, которые Office API JavaScript.

|Разрешение|Значение|Описание|
|---|---:|---|
|Чтение|1 (000001)|Может читать элементы.|
|Запись|2 (000010)|Можно создавать элементы.|
|DeleteOwn|4 (000100)|Можно удалить только созданные элементы.|
|DeleteAll|8 (001000)|Может удалять любые элементы.|
|EditOwn|16 (010000)|Можно редактировать только созданные элементы.|
|EditAll|32 (100000)|Может изменять любые элементы.|

> [!NOTE]
> В настоящее время API поддерживает получение существующих разрешений делегирования, но не установку разрешений делегирования.

Объект [DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) реализуется с помощью битмаски для указать разрешения делегата. Каждая позиция в битмаске представляет определенное разрешение, и если оно заданной, то делегат `1` имеет соответствующее разрешение. Например, если второй бит справа , то у делегата `1` есть разрешение **на записи.** Пример проверки определенного разрешения в разделе [Выполнение](#perform-an-operation-as-delegate) операции в качестве делегата см. в этой статье.

## <a name="sync-across-mailbox-clients"></a>Синхронизация между клиентами почтовых ящиков

Обновления делегата в почтовом ящике владельца обычно синхронизируются между почтовыми ящиками немедленно.

Однако если операции REST или Exchange Web Services (EWS) использовались для набора расширенного свойства элемента, синхронизация таких изменений может занять несколько часов. Мы рекомендуем вместо этого использовать [объект CustomProperties](/javascript/api/outlook/office.customproperties) и связанные API, чтобы избежать такой задержки. Дополнительные дополнительные [](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) статьи см. в разделе настраиваемые свойства в статье "Получить и установить метаданные в Outlook надстройки".

> [!IMPORTANT]
> В сценарии делегирования нельзя использовать EWS с маркерами, которые в настоящее время office.js API.

## <a name="configure-the-manifest"></a>Настройка манифеста

Чтобы включить сценарии делегирования доступа в надстройку, необходимо настроить элемент [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) в манифесте под `true` родительским элементом. `DesktopFormFactor` В настоящее время другие форм-факторы не поддерживаются.

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

## <a name="perform-an-operation-as-delegate"></a>Выполнение операции в качестве делегата

Общие свойства элемента можно получить в режиме Compose или Read, позвонив по методу [item.getSharedPropertiesAsync.](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) Это возвращает объект [SharedProperties,](/javascript/api/outlook/office.sharedproperties) который в настоящее время предоставляет разрешения делегата, адрес электронной почты владельца, базовый URL-адрес API REST и целевой почтовый ящик.

В следующем примере показано, как получить общие свойства сообщения или встречи, проверить, есть ли у делегата разрешение **на** запись, и сделать вызов REST.

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

В зависимости от сценариев надстройки существует несколько ограничений, которые следует учитывать при работе с ситуациями делегатов.

### <a name="rest-and-ews"></a>REST и EWS

Ваша надстройка может использовать REST, но не EWS, и необходимо установить разрешение надстройки, чтобы включить доступ REST к почтовому `ReadWriteMailbox` ящику владельца.

### <a name="message-compose-mode"></a>Режим композитации сообщений

В режиме композитации сообщений [getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#getSharedPropertiesAsync_options__callback_) не поддерживается Outlook в Интернете или Windows, если не выполнены следующие условия.

1. Владелец делит с делегатом по крайней мере одну папку почтовых ящиков.
1. Делегат проектирует сообщение в общей папке.

    Примеры:

    - Делегат отвечает на сообщения электронной почты в общей папке или переададирует их.
    - Делегат сохраняет черновик сообщения, а затем перемещает его из собственной папки **Drafts** в общую папку. Делегат открывает черновик из общей папки, а затем продолжает сочинять.

После того как сообщение отправлено, оно обычно находится в папке **отправленных** элементов делегата.

## <a name="see-also"></a>См. также

- [Разрешить другим пользователям управлять почтой и календарем](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)
- [Общий доступ к календарю в Microsoft 365](https://support.office.com/article/calendar-sharing-in-office-365-b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [Как заказать элементы манифеста](../develop/manifest-element-ordering.md)
- [Маска (вычисления)](https://en.wikipedia.org/wiki/Mask_(computing))
- [Операторы bitwise JavaScript](https://www.w3schools.com/js/js_bitwise.asp)