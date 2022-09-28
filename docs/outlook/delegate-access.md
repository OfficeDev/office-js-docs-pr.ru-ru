---
title: Включение общих папок и сценариев общих почтовых ящиков в надстройке Outlook
description: Описывается настройка поддержки надстроек для общих папок (например, делегировать доступ) и общим почтовым ящикам.
ms.date: 09/12/2022
ms.localizationpriority: medium
ms.openlocfilehash: 70efecda863e26f085b6f93cf26091fe0b9a9ea6
ms.sourcegitcommit: 05be1086deb2527c6c6ff3eafcef9d7ed90922ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/28/2022
ms.locfileid: "68092926"
---
# <a name="enable-shared-folders-and-shared-mailbox-scenarios-in-an-outlook-add-in"></a>Включение общих папок и сценариев общих почтовых ящиков в надстройке Outlook

В этой статье описывается, как включить в надстройке Outlook сценарии общих папок (также называемых делегированным доступом [) и](/javascript/api/requirement-sets/outlook/preview-requirement-set/outlook-requirement-set-preview#shared-mailboxes) общих почтовых ящиков (в настоящее время в предварительной версии), включая разрешения, поддерживаемые API JavaScript для Office.

## <a name="supported-clients-and-platforms"></a>Поддерживаемые клиенты и платформы

В следующей таблице показаны поддерживаемые сочетания клиента и сервера для этой функции, включая минимальное необходимое накопительное обновление, если это применимо. Исключенные сочетания не поддерживаются.

| Client | Exchange Online | Локальная среда Exchange 2019<br>(накопительное обновление 1 или более поздней версии) | Локальная версия Exchange 2016<br>(накопительный пакет обновления 6 или более поздней версии) | Локальная версия Exchange 2013 |
|---|:---:|:---:|:---:|:---:|
|Windows:<br>Версия 1910 (сборка 12130.20272) или более поздняя|Да|Да\*|Да\*|Да\*|
|Mac:<br>сборка 16.47 или более поздняя|Да|Да|Да|Да|
|Веб-браузер:<br>Современный пользовательский интерфейс Outlook|Да|Неприменимо|Неприменимо|Неприменимо|
|Веб-браузер:<br>классический пользовательский интерфейс Outlook|Неприменимо|НЕТ|Нет|Нет|

> [!NOTE]
> \* Поддержка этой функции в локальной среде Exchange доступна начиная с версии 2206 (сборка 15330.20000) для Канала Current Channel и версии 2207 (сборка 15427.20000) для канала Monthly Enterprise.

> [!IMPORTANT]
> Поддержка этой функции была представлена в наборе обязательных [элементов 1.8](/javascript/api/requirement-sets/outlook/requirement-set-1.8/outlook-requirement-set-1.8) (дополнительные сведения см. на клиентах [и платформах](/javascript/api/requirement-sets/outlook/outlook-api-requirement-sets#requirement-sets-supported-by-exchange-servers-and-outlook-clients)). Однако обратите внимание, что матрица поддержки функции является надмножеством наборов обязательных элементов.

## <a name="supported-setups"></a>Поддерживаемые настройки

В следующих разделах описаны поддерживаемые конфигурации для общих почтовых ящиков (сейчас в предварительной версии) и общих папок. API-интерфейсы функций могут работать не так, как ожидалось в других конфигурациях. Выберите платформу, которую вы хотите настроить.

### <a name="windows"></a>[Windows](#tab/windows)

#### <a name="shared-folders"></a>Общие папки

Владелец почтового ящика должен сначала [предоставить доступ к делегату](https://support.microsoft.com/office/41c40c04-3bd1-4d22-963a-28eafec25926). Затем делегат должен выполнить инструкции, описанные в разделе "Добавление почтового ящика другого пользователя в профиль" статьи "Управление элементами почты и календаря другого [пользователя"](https://support.microsoft.com/office/afb79d6b-2967-43b9-a944-a6b953190af5).

#### <a name="shared-mailboxes-preview"></a>Общие почтовые ящики (предварительная версия)

Администраторы exchange Server могут создавать общие почтовые ящики и управлять ими для доступа к наборам пользователей. [Exchange Online](/exchange/collaboration-exo/shared-mailboxes) [и локальные среды Exchange](/exchange/collaboration/shared-mailboxes/create-shared-mailboxes) поддерживаются.

Функция Exchange Server, известная как "автосопоставка", включена по умолчанию. Это означает, [](/microsoft-365/admin/email/create-a-shared-mailbox?view=o365-worldwide&preserve-view=true#add-the-shared-mailbox-to-outlook) что впоследствии общий почтовый ящик должен автоматически отображаться в приложении Outlook пользователя после закрытия и повторного открытия Outlook. Однако если администратор отключил автоматическое сопоставление, пользователь должен выполнить действия, описанные в разделе "Добавление общего почтового ящика в Outlook" статьи "Открытие и использование общего почтового ящика в [Outlook"](https://support.microsoft.com/office/d94a8e9e-21f1-4240-808b-de9c9c088afd).

> [!WARNING]
> Не **входить** в общий почтовый ящик с помощью пароля. В этом случае API-интерфейсы функций не будут работать.

### <a name="web-browser---modern-outlook"></a>[Веб-браузер — современная версия Outlook](#tab/modern)

#### <a name="shared-folders"></a>Общие папки

Владелец почтового ящика должен сначала [предоставить доступ делегату,](https://www.microsoft.com/microsoft-365/blog/2013/09/04/configuring-delegate-access-in-outlook-web-app/) обновив разрешения папки почтового ящика. Затем делегат должен следовать инструкциям, приведенным в разделе "Добавление почтового ящика другого пользователя в список папок в Outlook Web App" статьи "Доступ к почтовому ящику другого [пользователя"](https://support.microsoft.com/office/a909ad30-e413-40b5-a487-0ea70b763081).

#### <a name="shared-mailboxes"></a>Общие почтовые ящики

Сценарии общих почтовых ящиков в надстройки Outlook в настоящее время не поддерживаются в современных Outlook в Интернете.

### <a name="mac"></a>[Mac](#tab/unix)

#### <a name="shared-mailboxes-preview"></a>Общие почтовые ящики (предварительная версия)

Почта и календарь совместно используются делегатом или пользователем общего почтового ящика. Надстройки доступны делегату или пользователю в режимах чтения и создания сообщений и встреч.

#### <a name="shared-folders"></a>Общие папки

Если **папка "** Входящие" совместно используется делегатом, надстройки доступны делегату в режиме чтения сообщений.

Если **папка Drafts** также предоставляется делегату, надстройки доступны в режиме создания.

#### <a name="local-shared-calendar-new-model"></a>Локальный общий календарь (новая модель)

Если владелец календаря явным образом предоставил общий доступ к календарю делегату (возможно, общий доступ ко всему почтовому ящику отсутствует), надстройки будут доступны делегату в режимах чтения встречи и создания.

#### <a name="remote-shared-calendar-previous-model"></a>Удаленный общий календарь (предыдущая модель)

Если владелец календаря предоставил общий доступ к календарю (например, предоставил возможность редактирования определенному DL или всей организации), пользователи могут иметь косвенное или неявное разрешение, а надстройки будут доступны этим пользователям в режиме чтения и создания встречи.

---

Дополнительные сведения о том, где надстройки обычно и не активируются, см. в [](outlook-add-ins-overview.md#mailbox-items-available-to-add-ins) разделе "Почтовые ящики", доступных для надстроек на странице обзора надстроек Outlook.

## <a name="supported-permissions"></a>Поддерживаемые разрешения

В следующей таблице описаны разрешения, поддерживаемые API JavaScript для Office для делегатов и пользователей общих почтовых ящиков.

|Разрешение|Значение|Описание|
|---|---:|---|
|Чтение|1 (000001)|Может считывать элементы.|
|Запись|2 (000010)|Может создавать элементы.|
|DeleteOwn|4 (000100)|Может удалять только созданные элементы.|
|DeleteAll|8 (001000)|Может удалять любые элементы.|
|EditOwn|16 (010000)|Может изменять только созданные элементы.|
|EditAll|32 (100000)|Может изменять любые элементы.|

> [!NOTE]
> В настоящее время API поддерживает получение существующих разрешений, но не настройку разрешений.

Объект [DelegatePermissions](/javascript/api/outlook/office.mailboxenums.delegatepermissions) реализуется с помощью битовой маски, указывав разрешения. Каждая позиция в битовой маске представляет определенное `1` разрешение, и если оно задано, пользователь имеет соответствующее разрешение. Например, если второй бит справа, `1`пользователь имеет разрешение **на запись** . См. пример проверки на наличие определенного разрешения в разделе "Выполнение операции [](#perform-an-operation-as-delegate-or-shared-mailbox-user) в качестве делегата или общего почтового ящика пользователя" далее в этой статье.

## <a name="sync-across-shared-folder-clients"></a>Синхронизация между клиентами общих папок

Обновления делегата для почтового ящика владельца обычно синхронизируются между почтовыми ящиками немедленно.

Однако если операции REST или веб-служб Exchange (EWS) использовались для задания расширенного свойства элемента, синхронизация таких изменений может занять несколько часов. Вместо этого рекомендуется использовать объект [CustomProperties](/javascript/api/outlook/office.customproperties) и связанные API, чтобы избежать такой задержки. Дополнительные сведения см. в [](metadata-for-an-outlook-add-in.md#custom-data-per-item-in-a-mailbox-custom-properties) разделе пользовательских свойств статьи "Получение и установка метаданных в надстройке Outlook".

> [!IMPORTANT]
> В сценарии делегата нельзя использовать EWS с маркерами, которые в настоящее время предоставляются office.js API.

## <a name="configure-the-manifest"></a>Настройка манифеста

Чтобы включить общие папки и сценарии общих почтовых ящиков в надстройке, необходимо задать элемент [SupportsSharedFolders](/javascript/api/manifest/supportssharedfolders) `true` в манифесте в родительском элементе `DesktopFormFactor`. В настоящее время другие форм-факторы не поддерживаются.

Чтобы поддерживать вызовы REST от делегата, задайте для узла [разрешений](/javascript/api/manifest/permissions) в манифесте значение `ReadWriteMailbox`.

В следующем примере показан элемент `SupportsSharedFolders` , задав `true` значение в разделе манифеста.

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

## <a name="perform-an-operation-as-delegate-or-shared-mailbox-user"></a>Выполнение операции в качестве пользователя делегата или общего почтового ящика

Общие свойства элемента можно получить в режиме создания или чтения, вызвав метод [item.getSharedPropertiesAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#methods) . Возвращает объект [SharedProperties](/javascript/api/outlook/office.sharedproperties) , который в настоящее время предоставляет разрешения пользователя, адрес электронной почты владельца, базовый URL-адрес REST API и целевой почтовый ящик.

В следующем примере показано, как получить общие свойства сообщения или встречи, проверить, имеет ли пользователь делегата или общего почтового ящика разрешение на запись, и выполнить вызов REST.

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
> В качестве делегата можно использовать REST для получения содержимого сообщения Outlook, присоединенного к [элементу Outlook или записи группы](/graph/outlook-get-mime-message#get-mime-content-of-an-outlook-message-attached-to-an-outlook-item-or-group-post).

## <a name="handle-calling-rest-on-shared-and-non-shared-items"></a>Обработка вызова REST для общих и не общих элементов

Если вы хотите вызвать операцию REST для элемента, независимо от того, является ли элемент общим, `getSharedPropertiesAsync` можно использовать API, чтобы определить, является ли элемент общим. После этого можно создать URL-адрес REST для операции с помощью соответствующего объекта.

```js
if (item.getSharedPropertiesAsync) {
  // In Windows, Mac, and the web client, this indicates a shared item so use SharedProperties properties to construct the REST URL.
  // Add-ins don't activate on shared items in mobile so no need to handle.

  // Perform operation for shared item.
} else {
  // In general, this is not a shared item, so construct the REST URL using info from the Call REST APIs article:
  // https://learn.microsoft.com/office/dev/add-ins/outlook/use-rest-api

  // Perform operation for non-shared item.
}
```

## <a name="limitations"></a>Ограничения

В зависимости от сценариев надстройки существует несколько ограничений, которые следует учитывать при обработке общих папок или общих почтовых ящиков.

### <a name="message-compose-mode"></a>Режим создания сообщения

В режиме создания сообщения [getSharedPropertiesAsync](/javascript/api/outlook/office.messagecompose#outlook-office-messagecompose-getsharedpropertiesasync-member(1)) не поддерживается в Outlook в Интернете или Windows, если не выполняются следующие условия.

А. **Делегирование доступа и общих папок**

1. Владелец почтового ящика запускает сообщение. Это может быть новое сообщение, ответ или пересылка.
1. Они сохраняют сообщение, а затем перемещают его из собственной папки **"Черновики** " в папку, предоставленную делегату.
1. Делегат открывает черновик из общей папки, а затем продолжает составление.

Б. **Общий почтовый ящик (применяется только к Outlook в Windows)**

1. Пользователь общего почтового ящика запускает сообщение. Это может быть новое сообщение, ответ или пересылка.
1. Они сохраняют сообщение, а затем перемещают его из собственной папки **Drafts** в папку в общем почтовом ящике.
1. Другой пользователь общего почтового ящика открывает черновик из общего почтового ящика, а затем продолжает составление.

Теперь сообщение находится в общем контексте, и надстройки, поддерживающие эти общие сценарии, могут получить общие свойства элемента. После отправки сообщение обычно находится в папке "Отправленные" **отправителя.**

### <a name="rest-and-ews"></a>REST и EWS

Ваша надстройка может использовать REST `ReadWriteMailbox` , и необходимо задать разрешение надстройки, чтобы разрешить доступ REST к почтовому ящику владельца или к общему почтовому ящику, если это применимо. EWS не поддерживается.

### <a name="user-or-shared-mailbox-hidden-from-an-address-list"></a>Пользователь или общий почтовый ящик, скрытый из списка адресов

Если администратор скрывал адрес пользователя или общего почтового ящика из списка адресов, например глобального списка адресов (GAL), `Office.context.mailbox.item` затронутые элементы электронной почты, открытые в отчете почтового ящика, как NULL. Например, если пользователь открывает почтовый элемент в общем почтовом ящике, который скрыт от gal, `Office.context.mailbox.item` то этот элемент электронной почты имеет значение NULL.

## <a name="see-also"></a>См. также

- [Предоставление другому пользователю разрешения на управление вашей почтой и календарем](https://support.microsoft.com/office/41c40c04-3bd1-4d22-963a-28eafec25926)
- [Общий доступ к календарю в Microsoft 365](https://support.microsoft.com/office/b576ecc3-0945-4d75-85f1-5efafb8a37b4)
- [Добавление общего почтового ящика в Outlook](/microsoft-365/admin/email/create-a-shared-mailbox?view=o365-worldwide&preserve-view=true#add-the-shared-mailbox-to-outlook)
- [Упорядочение элементов манифеста](../develop/manifest-element-ordering.md)
- [Маска (вычисления)](https://en.wikipedia.org/wiki/Mask_(computing))
- [Побитовые операторы JavaScript](https://www.w3schools.com/js/js_bitwise.asp)
