---
title: Включение сценариев делегирования доступа в надстройке Outlook
description: В кратко описывается доступ представителя и описывается настройка поддержки надстройки.
ms.date: 09/30/2020
localization_priority: Normal
ms.openlocfilehash: 68e9c8003f8d223a591283fd1a73f0a38bd3c8a4
ms.sourcegitcommit: 6c3a04acde57832feeaaa599148f93af7e3e36ea
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/02/2020
ms.locfileid: "48336421"
---
# <a name="enable-delegate-access-scenarios-in-an-outlook-add-in"></a>Включение сценариев делегирования доступа в надстройке Outlook

Владелец почтового ящика может использовать функцию делегированного доступа, чтобы [Разрешить другому пользователю управлять своей почтой и календарем](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926). В этой статье указывается, какие разрешения представителей поддерживает API JavaScript для Office, а также описывается включение сценариев делегированного доступа в надстройке Outlook.

> [!IMPORTANT]
> Доступ к представителю в настоящее время недоступен в Outlook на Android и iOS. Кроме того, эта функция в настоящее время недоступна для [групп общих почтовых ящиков](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide&preserve-view=true#shared-mailboxes) в Outlook в Интернете. Эта функция может быть доступна в будущем.
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

Объект [делегатепермиссионс](/javascript/api/outlook/office.mailboxenums.delegatepermissions) реализуется с помощью битовой маски для указания разрешений делегата. Каждое положение в битовой маске представляет конкретное разрешение и, если ему присвоено значение, `1` у делегата есть соответствующее разрешение. Например, если второй бит справа `1` , то делегат имеет разрешение на **запись** . Вы можете увидеть пример того, как проверить наличие определенного разрешения в разделе [выполнение операции как делегата](#perform-an-operation-as-delegate) далее в этой статье.

## <a name="sync-across-mailbox-clients"></a>Синхронизация между клиентами почтовых ящиков

Обновление делегата почтового ящика владельца обычно синхронизируется в почтовых ящиках немедленно.

Тем не менее, если для задания расширенного свойства элемента использовались операции REST или Exchange Web Services (EWS), такие изменения могут занять несколько часов. Мы рекомендуем вместо этого использовать объект [CustomProperties](/javascript/api/outlook/office.customproperties) и связанные с ним API, чтобы избежать такой задержки. Чтобы узнать больше, ознакомьтесь с [разделом Настраиваемые свойства](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) статьи "получение и Настройка метаданных в надстройке Outlook".

> [!IMPORTANT]
> В сценарии делегата EWS невозможно использовать с маркерами, которые в настоящее время предоставляются office.js API.

## <a name="configure-the-manifest"></a>Настройка манифеста

Чтобы включить сценарии делегирования доступа в надстройке, необходимо задать элемент [SupportsSharedFolders](../reference/manifest/supportssharedfolders.md) `true` в манифесте под родительским элементом `DesktopFormFactor` . В настоящее время другие конструктивные параметры не поддерживаются.

Чтобы обеспечить поддержку вызовов REST от делегата, задайте для узла [Permissions](../reference/manifest/permissions.md) в манифесте значение `ReadWriteMailbox` .

В приведенном ниже примере показано, как `SupportsSharedFolders` задать элемент `true` в разделе манифеста.

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

> [!TIP]
> Как представитель вы можете использовать REST для [получения содержимого сообщения Outlook, присоединенного к элементу Outlook или записи группы](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post).

## <a name="handle-calling-rest-on-shared-and-non-shared-items"></a>Обработка вызовов REST для общих и необщих элементов

Если вы хотите вызвать операцию REST для элемента, независимо от того, является ли элемент общим, вы можете использовать `getSharedPropertiesAsync` API, чтобы определить, является ли элемент общим. После этого вы можете создать URL-адрес REST для операции, используя соответствующий объект.

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

В зависимости от сценариев надстройки существует ряд ограничений, которые необходимо учитывать при обработке ситуаций делегата.

### <a name="rest-and-ews"></a>REST и EWS

Надстройка может использовать REST, но не EWS, и разрешение надстройки должно быть настроено на разрешение `ReadWriteMailbox` REST доступа к почтовому ящику владельца.

### <a name="message-compose-mode"></a>Режим создания сообщения

В режиме создания сообщений [жетшаредпропертиесасинк](/javascript/api/outlook/office.messagecompose#getsharedpropertiesasync-options--callback-) не поддерживается в Outlook в Интернете или Windows, если не выполняются следующие условия.

1. Владелец предоставляет по крайней мере одну папку почтового ящика с представителем.
1. Делегирование черновика сообщения в общей папке.

    Примеры:

    - Делегат отправляет сообщение электронной почты в общую папку или пересылает его.
    - Делегат сохраняет черновик сообщения и перемещает его из своей папки **"Черновики** " в общую папку. После этого представитель открывает черновик из общей папки, а затем продолжает сохранится.

После отправки сообщения оно обычно находится в папке " **Отправленные** " делегата.

## <a name="see-also"></a>См. также

- [Предоставление другим пользователям возможности управлять почтой и календарем](https://support.office.com/article/allow-someone-else-to-manage-your-mail-and-calendar-41c40c04-3bd1-4d22-963a-28eafec25926)
- [Общий доступ к календарю в Office 365](https://support.office.com/article/calendar-sharing-in-office-365-b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [Порядок элементов манифеста](../develop/manifest-element-ordering.md)
- [Mask (вычисления)](https://en.wikipedia.org/wiki/Mask_(computing))
- [Битовые операторы JavaScript](https://www.w3schools.com/js/js_bitwise.asp)