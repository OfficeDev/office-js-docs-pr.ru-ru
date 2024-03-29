---
title: Использование REST API Outlook из надстройки Outlook
description: Узнайте, как использовать REST API Outlook из надстройки Outlook, чтобы получить маркер доступа
ms.date: 10/03/2022
ms.localizationpriority: medium
ms.openlocfilehash: 9f62b2514f05341531a826c29e18c593a590fca0
ms.sourcegitcommit: 005783ddd43cf6582233be1be6e3463d7ab9b0e5
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/05/2022
ms.locfileid: "68467218"
---
# <a name="use-the-outlook-rest-apis-from-an-outlook-add-in"></a>Использование REST API Outlook из надстройки Outlook

The [Office.context.mailbox.item](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox.item) namespace provides access to many of the common fields of messages and appointments. However, in some scenarios an add-in may need to access data that is not exposed by the namespace. For example, the add-in may rely on custom properties set by an outside app, or it needs to search the user's mailbox for messages from the same sender. In these scenarios, the [Outlook REST APIs](/outlook/rest) is the recommended method to retrieve the information.

> [!IMPORTANT]
> **Интерфейсы REST API Outlook устарели**
>
> Конечные точки REST Outlook будут полностью списываются 30 ноября 2022 г. (дополнительные сведения см. в объявлении за [ноябрь 2020 г.](https://developer.microsoft.com/graph/blogs/outlook-rest-api-v2-0-deprecation-notice/)). Для использования Microsoft Graph необходимо перенести существующие [надстройки](/outlook/rest#outlook-rest-api-via-microsoft-graph). Инструкции см. в статье ["Сравнение конечных точек REST API Microsoft Graph и Outlook"](/outlook/rest/compare-graph).
>
> Чтобы помочь вам в миграции, активные надстройки, использующие службу REST, имеют право на освобождение от использования службы до окончания расширенной поддержки [Outlook 2019 14 октября 2025 г](/lifecycle/end-of-support/end-of-support-2025). К ним относятся новые надстройки, разработанные после 30 ноября 2022 г. Исключение основано на идентификаторе манифеста надстройки и применяется к частным выпускам и надстройки, размещаемые в AppSource.
>
> Автоматическая идентификация трафика надстроек Outlook, использующих службу REST, в настоящее время тестируются для проверки исключений. Если вы хотите принять участие в этом этапе тестирования, заполните форму проверки надстройки [REST API](https://aka.ms/RESTCheck) до ноября 2022 г. Дополнительные сведения см. в записи блога о звонках в сообществе надстроек [Office за август 2022 г](https://pnp.github.io/blog/office-add-ins-community-call/2022-08-10/).

## <a name="get-an-access-token"></a>Получение токена доступа

The Outlook REST APIs require a bearer token in the `Authorization` header. Typically apps use OAuth2 flows to retrieve a token. However, add-ins can retrieve a token without implementing OAuth2 by using the new [Office.context.mailbox.getCallbackTokenAsync](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#methods) method introduced in the Mailbox requirement set 1.5.

Задав для параметра `isRest` значение `true`, вы можете запросить маркер, совместимый с интерфейсами REST API.

### <a name="add-in-permissions-and-token-scope"></a>Разрешения надстроек и область маркера

Важно учитывать уровень доступа через интерфейсы REST API, который потребуется надстройке. В большинстве случаев маркер, возвращенный методом `getCallbackTokenAsync`, предоставляет доступ только для чтения и только для текущего элемента. Это верно, даже если надстройка указывает уровень разрешений на чтение [и](understanding-outlook-add-in-permissions.md#readwrite-item-permission) запись элемента в манифесте.

Если надстройке потребуется доступ на запись к текущему элементу или другим элементам в почтовом ящике пользователя, надстройка должна указать разрешение на чтение и запись почтового [ящика](understanding-outlook-add-in-permissions.md#readwrite-mailbox-permission).
в манифесте. В этом случае возвращаемый маркер предоставляет доступ на чтение и запись к сообщениям, событиям и контактам пользователя.

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

The final piece of information your add-in needs to call the REST API is the hostname it should use to send API requests. This information is in the [Office.context.mailbox.restUrl](/javascript/api/requirement-sets/outlook/preview-requirement-set/office.context.mailbox#properties) property.

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
  // https://learn.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations#get-messages.
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
