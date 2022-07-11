---
title: Использование REST API Outlook из надстройки Outlook
description: Узнайте, как использовать REST API Outlook из надстройки Outlook, чтобы получить маркер доступа
ms.date: 07/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: 7c02b878b6636e6736ada4a29d123dd8ff772393
ms.sourcegitcommit: d8ea4b761f44d3227b7f2c73e52f0d2233bf22e2
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/11/2022
ms.locfileid: "66712967"
---
# <a name="use-the-outlook-rest-apis-from-an-outlook-add-in"></a>Использование REST API Outlook из надстройки Outlook

Пространство имен [Office.context.mailbox.item](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item) предоставляет доступ ко множеству общих полей сообщений и встреч. Однако в некоторых случаях надстройке может потребоваться доступ к данным, недоступным из этого пространства имен. Например, надстройка может использовать настраиваемые свойства, заданные внешним приложением, или искать в почтовом ящике пользователя сообщения от одного отправителя. В таких случаях для получения сведений рекомендуется использовать [интерфейсы REST API Outlook](/outlook/rest).

> [!IMPORTANT]
> **Интерфейсы REST API Outlook устарели**
>
> Конечные точки REST Outlook будут полностью списываются в ноябре 2022 г. (дополнительные сведения см. в объявлении за [ноябрь 2020 г.](https://developer.microsoft.com/graph/blogs/outlook-rest-api-v2-0-deprecation-notice/)). Для использования Microsoft Graph необходимо перенести существующие [надстройки](/outlook/rest#outlook-rest-api-via-microsoft-graph). Кроме того, [сравните конечные точки REST API Microsoft Graph и Outlook](/outlook/rest/compare-graph).

## <a name="get-an-access-token"></a>Получение токена доступа

Интерфейсам REST API для Outlook необходим маркер носителя в заголовке `Authorization`. Как правило, приложения используют потоки OAuth2 для получения маркера. Однако надстройка может получить маркер без реализации OAuth2, используя новый метод [Office.context.mailbox.getCallbackTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods), который появился в наборе требований 1.5 для почтовых ящиков.

Задав для параметра `isRest` значение `true`, вы можете запросить маркер, совместимый с интерфейсами REST API.

### <a name="add-in-permissions-and-token-scope"></a>Разрешения надстроек и область маркера

Важно учитывать уровень доступа через интерфейсы REST API, который потребуется надстройке. В большинстве случаев маркер, возвращенный методом `getCallbackTokenAsync`, предоставляет доступ только для чтения и только для текущего элемента. Это верно, даже если в манифесте надстройки указан уровень разрешений `ReadWriteItem`.

Если надстройке требуется доступ на запись к текущему элементу или другим элементам в почтовом ящике пользователя, в манифесте надстройки необходимо указать уровень разрешений `ReadWriteMailbox`. В этом случае возвращаемый маркер предоставляет доступ на чтение и запись к сообщениям, событиям и контактам пользователя.

### <a name="example"></a>Пример

```js
Office.context.mailbox.getCallbackTokenAsync({isRest: true}, function(result){
  if (result.status === "succeeded") {
    const accessToken = result.value;

    // Use the access token.
    getCurrentItem(accessToken);
  } else {
    // Handle the error.
  }
});
```

## <a name="get-the-item-id"></a>Получение идентификатора элемента

Чтобы получить текущий элемент с помощью REST, надстройке потребуется его идентификатор, правильно отформатированный для службы REST. Его можно получить из свойства [Office.context.mailbox.item.itemId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item#properties), но необходимо выполнить некоторые проверки, чтобы убедиться, что идентификатор отформатирован для REST.

- В Outlook Mobile свойство `Office.context.mailbox.item.itemId` возвращает идентификатор в формате REST, который можно использовать без изменений.
- В других клиентах Outlook свойство `Office.context.mailbox.item.itemId` возвращает идентификатор в формате EWS, который необходимо преобразовать с помощью метода [Office.context.mailbox.convertToRestId](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods).
- Обратите внимание: чтобы использовать идентификатор вложения, его нужно преобразовать в идентификатор в формате REST. Это преобразование необходимо, потому что идентификаторы EWS могут содержать небезопасные в отношении URL-адресов значения, которые вызывают проблемы с REST.

Надстройка может определить, в каком клиенте Outlook она загружена, с помощью свойства [Office.context.mailbox.diagnostics.hostName](/javascript/api/outlook/office.diagnostics#outlook-office-diagnostics-hostname-member).

### <a name="example"></a>Пример

```js
function getItemRestId() {
  if (Office.context.mailbox.diagnostics.hostName === 'OutlookIOS') {
    // itemId is already REST-formatted.
    return Office.context.mailbox.item.itemId;
  } else {
    // Convert to an item ID for API v2.0.
    return Office.context.mailbox.convertToRestId(
      Office.context.mailbox.item.itemId,
      Office.MailboxEnums.RestVersion.v2_0
    );
  }
}
```

## <a name="get-the-rest-api-url"></a>Использование URL-адреса REST API

Последнее значение, необходимое надстройке для вызова REST API, — это имя узла, используемое для отправки запросов API. Оно содержится в свойстве [Office.context.mailbox.restUrl](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#properties).

### <a name="example"></a>Пример

```js
// Example: https://outlook.office.com
const restHost = Office.context.mailbox.restUrl;
```

## <a name="call-the-api"></a>Вызов API

Когда надстройка получит маркер доступа, идентификатор элемента и URL-адрес REST API, она может передать эти сведения внутренней службе, которая вызовет REST API, или вызвать его напрямую с помощью AJAX. В приведенном ниже примере вызывается REST API почты Outlook для получения текущего сообщения.

> [!IMPORTANT]
> Для локальных развертываний Exchange клиентские запросы, использующие AJAX или аналогичные библиотеки, завершались сбоем, так как CORS не поддерживается в этой настройке сервера.

```js
function getCurrentItem(accessToken) {
  // Get the item's REST ID.
  const itemId = getItemRestId();

  // Construct the REST URL to the current item.
  // Details for formatting the URL can be found at
  // https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations#get-messages.
  const getMessageUrl = Office.context.mailbox.restUrl +
    '/v2.0/me/messages/' + itemId;

  $.ajax({
    url: getMessageUrl,
    dataType: 'json',
    headers: { 'Authorization': 'Bearer ' + accessToken }
  }).done(function(item){
    // Message is passed in `item`.
    const subject = item.Subject;
    ...
  }).fail(function(error){
    // Handle error.
  });
}
```

## <a name="see-also"></a>См. также

- Пример вызова REST API из надстроек Outlook: [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo) на сайте GitHub.
- REST API Outlook также доступны через конечную точку Microsoft Graph, но с некоторыми важными отличиями, включая способ получения надстройкой маркера доступа. Дополнительные сведения см. в разделе [REST API Outlook через Microsoft Graph](/outlook/rest/index#outlook-rest-api-via-microsoft-graph).
