---
title: Включение сценариев делегирования доступа в надстройке Outlook
description: В кратко описывается доступ представителя и описывается настройка поддержки надстройки.
ms.date: 01/14/2020
localization_priority: Normal
ms.openlocfilehash: 0941e4f0b5e1082b8a762acfa013d4e58be03469
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42721018"
---
# <a name="enable-delegate-access-scenarios-in-an-outlook-add-in"></a>Включение сценариев делегирования доступа в надстройке Outlook

Владелец почтового ящика может использовать функцию делегированного доступа, чтобы [Разрешить другому пользователю управлять своей почтой и календарем](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926). В этой статье указывается, какие разрешения представителей поддерживает API JavaScript для Office, а также описывается включение сценариев делегированного доступа в надстройке Outlook.

> [!IMPORTANT]
> Доступ к представителю в настоящее время недоступен в Outlook для Mac, Android и iOS. Эта функция может быть доступна в будущем.
>
> Поддержка этой функции появилась в наборе требований 1,8. См [клиенты и платформы](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients), поддерживающие этот набор обязательных требований.

## <a name="supported-permissions-for-delegate-access"></a>Поддерживаемые разрешения для делегированного доступа

В следующей таблице описаны разрешения представителей, поддерживаемые API JavaScript для Office.

|Разрешение|Значение|Описание|
|---|---:|---|
|Чтение|1 (000001)|Возможность чтения элементов.|
|Запись|2 (000010)|Может создавать элементы.|
|делетеовн|4 (000100)|Можно удалять только созданные ими элементы.|
|DeleteAll|8 (001000)|Может удалять все элементы.|
|едитовн|16 (010000)|Возможность изменения только созданных ими элементов.|
|едиталл|32 (100000)|Можно изменять любые элементы.|

> [!NOTE]
> В настоящее время API поддерживает доступ к существующим делегированным разрешениям, но не настраивает разрешения делегата.

Объект [делегатепермиссионс](/javascript/api/outlook/office.mailboxenums.delegatepermissions) реализуется с помощью битовой маски для указания разрешений делегата. Каждое положение в битовой маске представляет конкретное разрешение и, если ему `1` присвоено значение, у делегата есть соответствующее разрешение. Например, если второй бит справа `1`, то делегат имеет разрешение на **запись** . Вы можете увидеть пример того, как проверить наличие определенного разрешения в разделе [выполнение операции как делегата](#perform-an-operation-as-delegate) далее в этой статье.

## <a name="sync-across-mailbox-clients"></a>Синхронизация между клиентами почтовых ящиков

Обновление делегата почтового ящика владельца обычно синхронизируется в почтовых ящиках немедленно.

Тем не менее, если надстройка использует операции REST или EWS для задания расширенного свойства элемента, такие изменения могут занять несколько часов. Мы рекомендуем вместо этого использовать объект [CustomProperties](/javascript/api/outlook/office.customproperties) и связанные с ним API, чтобы избежать такой задержки. Чтобы узнать больше, ознакомьтесь с [разделом Настраиваемые свойства](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) статьи "получение и Настройка метаданных в надстройке Outlook".

## <a name="configure-the-manifest"></a>Настройка манифеста

Чтобы включить сценарии делегирования доступа в надстройке, необходимо задать элемент [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) `true` в манифесте под родительским элементом `DesktopFormFactor`. В настоящее время другие конструктивные параметры не поддерживаются.

В приведенном ниже примере `SupportsSharedFolders` показано, как `true` задать элемент в разделе манифеста.

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

Вы можете получить общие свойства элемента в режиме создания или чтения, вызвав метод [Item. жетшаредпропертиесасинк](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) . Возвращает объект [шаредпропертиес](/javascript/api/outlook/office.sharedproperties) , который в настоящее время предоставляет разрешения делегата, адрес электронной почты владельца, базовый URL-адрес REST API и целевой почтовый ящик.

В приведенном ниже примере показано, как получить общие свойства сообщения или встречи, проверить, есть ли у делегата разрешение на **запись** , и СОВЕРШИТЬ вызов REST.

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

## <a name="see-also"></a>См. также

- [Предоставление другим пользователям возможности управлять почтой и календарем](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)
- [Общий доступ к календарю в Office 365](https://support.office.com/article/calendar-sharing-in-office-365-b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [Порядок элементов манифеста](../develop/manifest-element-ordering.md)
- [Mask (вычисления)](https://en.wikipedia.org/wiki/Mask_(computing))
- [Битовые операторы JavaScript](https://www.w3schools.com/js/js_bitwise.asp)