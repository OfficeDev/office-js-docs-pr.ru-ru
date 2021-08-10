---
title: Использование REST API Outlook из надстройки Outlook
description: Узнайте, как использовать REST API Outlook из надстройки Outlook, чтобы получить маркер доступа
ms.date: 07/06/2021
localization_priority: Normal
ms.openlocfilehash: 60a3dab3794dc3c379ae3e6ddf47324a72726e0388839e884ec272c868bdec35
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57093312"
---
# <a name="use-the-outlook-rest-apis-from-an-outlook-add-in"></a>Использование REST API Outlook из надстройки Outlook

Пространство имен [Office.context.mailbox.item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md) предоставляет доступ ко множеству общих полей сообщений и встреч. Однако в некоторых случаях надстройке может потребоваться доступ к данным, недоступным из этого пространства имен. Например, надстройка может использовать настраиваемые свойства, заданные внешним приложением, или искать в почтовом ящике пользователя сообщения от одного отправителя. В таких случаях для получения сведений рекомендуется использовать [интерфейсы REST API Outlook](/outlook/rest).

> [!IMPORTANT]
> **API Outlook REST обесценились**
>
> Конечные точки rest Outlook полностью списаются в ноябре 2022 г. (дополнительные сведения см. в сообщении за ноябрь [2020 г.).](https://developer.microsoft.com/graph/blogs/outlook-rest-api-v2-0-deprecation-notice/) Чтобы использовать Microsoft Graph, следует перенести [существующие надстройки.](/outlook/rest#outlook-rest-api-via-microsoft-graph) Кроме того, [сравните конечные](/outlook/rest/compare-graph)точки Graph и Outlook API REST.

## <a name="get-an-access-token"></a>Получение токена доступа

Интерфейсам REST API для Outlook необходим маркер носителя в заголовке `Authorization`. Как правило, приложения используют потоки OAuth2 для получения маркера. Однако надстройка может получить маркер без реализации OAuth2, используя новый метод [Office.context.mailbox.getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods), который появился в наборе требований 1.5 для почтовых ящиков.

Задав для параметра `isRest` значение `true`, вы можете запросить маркер, совместимый с интерфейсами REST API.

### <a name="add-in-permissions-and-token-scope"></a>Разрешения надстроек и область маркера

Важно учитывать уровень доступа через интерфейсы REST API, который потребуется надстройке. В большинстве случаев маркер, возвращенный методом `getCallbackTokenAsync`, предоставляет доступ только для чтения и только для текущего элемента. Это верно, даже если в манифесте надстройки указан уровень разрешений `ReadWriteItem`.

Если надстройке требуется доступ на запись к текущему элементу или другим элементам в почтовом ящике пользователя, в манифесте надстройки необходимо указать уровень разрешений `ReadWriteMailbox`. В этом случае возвращаемый маркер предоставляет доступ на чтение и запись к сообщениям, событиям и контактам пользователя.

### <a name="example"></a>Пример

```js
Office.context.mailbox.getCallbackTokenAsync({isRest: true}, function(result){
  if (result.status === "succeeded") {
    var accessToken = result.value;

    // Use the access token.
    getCurrentItem(accessToken);
  } else {
    // Handle the error.
  }
});
```

## <a name="get-the-item-id"></a>Получение идентификатора элемента

Чтобы получить текущий элемент с помощью REST, надстройке потребуется его идентификатор, правильно отформатированный для службы REST. Его можно получить из свойства [Office.context.mailbox.item.itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties), но необходимо выполнить некоторые проверки, чтобы убедиться, что идентификатор отформатирован для REST.

- В Outlook Mobile свойство `Office.context.mailbox.item.itemId` возвращает идентификатор в формате REST, который можно использовать без изменений.
- В других клиентах Outlook свойство `Office.context.mailbox.item.itemId` возвращает идентификатор в формате EWS, который необходимо преобразовать с помощью метода [Office.context.mailbox.convertToRestId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods).
- Обратите внимание: чтобы использовать идентификатор вложения, его нужно преобразовать в идентификатор в формате REST. Это преобразование необходимо, потому что идентификаторы EWS могут содержать небезопасные в отношении URL-адресов значения, которые вызывают проблемы с REST.

Надстройка может определить, в каком клиенте Outlook она загружена, с помощью свойства [Office.context.mailbox.diagnostics.hostName](/javascript/api/outlook/office.diagnostics#hostName).

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

Последнее значение, необходимое надстройке для вызова REST API, — это имя узла, используемое для отправки запросов API. Оно содержится в свойстве [Office.context.mailbox.restUrl](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties).

### <a name="example"></a>Пример

```js
// Example: https://outlook.office.com
var restHost = Office.context.mailbox.restUrl;
```

## <a name="call-the-api"></a>Вызов API

Когда надстройка получит маркер доступа, идентификатор элемента и URL-адрес REST API, она может передать эти сведения внутренней службе, которая вызовет REST API, или вызвать его напрямую с помощью AJAX. В приведенном ниже примере вызывается REST API почты Outlook для получения текущего сообщения.

> [!IMPORTANT]
> Для локального Exchange развертывание клиентские запросы с помощью AJAX или аналогичных библиотек сбои, так как CORS не поддерживается в настройке сервера.

```js
function getCurrentItem(accessToken) {
  // Get the item's REST ID.
  var itemId = getItemRestId();

  // Construct the REST URL to the current item.
  // Details for formatting the URL can be found at
  // https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations#get-messages.
  var getMessageUrl = Office.context.mailbox.restUrl +
    '/v2.0/me/messages/' + itemId;

  $.ajax({
    url: getMessageUrl,
    dataType: 'json',
    headers: { 'Authorization': 'Bearer ' + accessToken }
  }).done(function(item){
    // Message is passed in `item`.
    var subject = item.Subject;
    ...
  }).fail(function(error){
    // Handle error.
  });
}
```

## <a name="see-also"></a>См. также

- Пример вызова REST API из надстроек Outlook: [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo) на сайте GitHub.
- REST API Outlook также доступны через конечную точку Microsoft Graph, но с некоторыми важными отличиями, включая способ получения надстройкой маркера доступа. Дополнительные сведения см. в разделе [REST API Outlook через Microsoft Graph](/outlook/rest/index#outlook-rest-api-via-microsoft-graph).