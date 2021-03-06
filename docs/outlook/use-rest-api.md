---
title: Использование REST API Outlook из надстройки Outlook
description: Узнайте, как использовать REST API Outlook из надстройки Outlook, чтобы получить маркер доступа
ms.date: 02/26/2021
localization_priority: Normal
ms.openlocfilehash: c0df1df4fdbda22768562892874e09bbeb760473
ms.sourcegitcommit: e7009c565b18c607fe0868db2e26e250ad308dce
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/05/2021
ms.locfileid: "50505488"
---
# <a name="use-the-outlook-rest-apis-from-an-outlook-add-in"></a><span data-ttu-id="d6fc1-103">Использование REST API Outlook из надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="d6fc1-103">Use the Outlook REST APIs from an Outlook add-in</span></span>

<span data-ttu-id="d6fc1-p101">Пространство имен [Office.context.mailbox.item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md) предоставляет доступ ко множеству общих полей сообщений и встреч. Однако в некоторых случаях надстройке может потребоваться доступ к данным, недоступным из этого пространства имен. Например, надстройка может использовать настраиваемые свойства, заданные внешним приложением, или искать в почтовом ящике пользователя сообщения от одного отправителя. В таких случаях для получения сведений рекомендуется использовать [интерфейсы REST API Outlook](/outlook/rest).</span><span class="sxs-lookup"><span data-stu-id="d6fc1-p101">The [Office.context.mailbox.item](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md) namespace provides access to many of the common fields of messages and appointments. However, in some scenarios an add-in may need to access data that is not exposed by the namespace. For example, the add-in may rely on custom properties set by an outside app, or it needs to search the user's mailbox for messages from the same sender. In these scenarios, the [Outlook REST APIs](/outlook/rest) is the recommended method to retrieve the information.</span></span>

> [!NOTE]
> <span data-ttu-id="d6fc1-108">Вы также можете получать доступ к [REST API Outlook через Microsoft Graph](/outlook/rest#outlook-rest-api-via-microsoft-graph), но при этом следует учитывать некоторые важные отличия.</span><span class="sxs-lookup"><span data-stu-id="d6fc1-108">You can also access [Outlook REST APIs via Microsoft Graph](/outlook/rest#outlook-rest-api-via-microsoft-graph) but there are some key differences.</span></span> <span data-ttu-id="d6fc1-109">Чтобы узнать больше, см. [сравнение Microsoft Graph и Outlook](/outlook/rest/compare-graph).</span><span class="sxs-lookup"><span data-stu-id="d6fc1-109">For more details, please [Compare Microsoft Graph and Outlook](/outlook/rest/compare-graph).</span></span>

## <a name="get-an-access-token"></a><span data-ttu-id="d6fc1-110">Получение токена доступа</span><span class="sxs-lookup"><span data-stu-id="d6fc1-110">Get an access token</span></span>

<span data-ttu-id="d6fc1-p103">Интерфейсам REST API для Outlook необходим маркер носителя в заголовке `Authorization`. Как правило, приложения используют потоки OAuth2 для получения маркера. Однако надстройка может получить маркер без реализации OAuth2, используя новый метод [Office.context.mailbox.getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods), который появился в наборе требований 1.5 для почтовых ящиков.</span><span class="sxs-lookup"><span data-stu-id="d6fc1-p103">The Outlook REST APIs require a bearer token in the `Authorization` header. Typically apps use OAuth2 flows to retrieve a token. However, add-ins can retrieve a token without implementing OAuth2 by using the new [Office.context.mailbox.getCallbackTokenAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method introduced in the Mailbox requirement set 1.5.</span></span>

<span data-ttu-id="d6fc1-114">Задав для параметра `isRest` значение `true`, вы можете запросить маркер, совместимый с интерфейсами REST API.</span><span class="sxs-lookup"><span data-stu-id="d6fc1-114">By setting the `isRest` option to `true`, you can request a token compatible with the REST APIs.</span></span>

### <a name="add-in-permissions-and-token-scope"></a><span data-ttu-id="d6fc1-115">Разрешения надстроек и область маркера</span><span class="sxs-lookup"><span data-stu-id="d6fc1-115">Add-in permissions and token scope</span></span>

<span data-ttu-id="d6fc1-p104">Важно учитывать уровень доступа через интерфейсы REST API, который потребуется надстройке. В большинстве случаев маркер, возвращенный методом `getCallbackTokenAsync`, предоставляет доступ только для чтения и только для текущего элемента. Это верно, даже если в манифесте надстройки указан уровень разрешений `ReadWriteItem`.</span><span class="sxs-lookup"><span data-stu-id="d6fc1-p104">It is important to consider what level of access your add-in will need via the REST APIs. In most cases, the token returned by `getCallbackTokenAsync` will provide read-only access to the current item only. This is true even if your add-in specifies the `ReadWriteItem` permission level in its manifest.</span></span>

<span data-ttu-id="d6fc1-p105">Если надстройке требуется доступ на запись к текущему элементу или другим элементам в почтовом ящике пользователя, в манифесте надстройки необходимо указать уровень разрешений `ReadWriteMailbox`. В этом случае возвращаемый маркер предоставляет доступ на чтение и запись к сообщениям, событиям и контактам пользователя.</span><span class="sxs-lookup"><span data-stu-id="d6fc1-p105">If your add-in will require write access to the current item or other items in the user's mailbox, your add-in must specify the `ReadWriteMailbox` permission level in its manifest. In this case, the token returned will contain read/write access to the user's messages, events, and contacts.</span></span>

### <a name="example"></a><span data-ttu-id="d6fc1-121">Пример</span><span class="sxs-lookup"><span data-stu-id="d6fc1-121">Example</span></span>

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

## <a name="get-the-item-id"></a><span data-ttu-id="d6fc1-122">Получение идентификатора элемента</span><span class="sxs-lookup"><span data-stu-id="d6fc1-122">Get the item ID</span></span>

<span data-ttu-id="d6fc1-123">Чтобы получить текущий элемент с помощью REST, надстройке потребуется его идентификатор, правильно отформатированный для службы REST.</span><span class="sxs-lookup"><span data-stu-id="d6fc1-123">To retrieve the current item via REST, your add-in will need the item's ID, properly formatted for REST.</span></span> <span data-ttu-id="d6fc1-124">Его можно получить из свойства [Office.context.mailbox.item.itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties), но необходимо выполнить некоторые проверки, чтобы убедиться, что идентификатор отформатирован для REST.</span><span class="sxs-lookup"><span data-stu-id="d6fc1-124">This is obtained from the [Office.context.mailbox.item.itemId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#properties) property, but some checks should be made to ensure that it is a REST-formatted ID.</span></span>

- <span data-ttu-id="d6fc1-125">В Outlook Mobile свойство `Office.context.mailbox.item.itemId` возвращает идентификатор в формате REST, который можно использовать без изменений.</span><span class="sxs-lookup"><span data-stu-id="d6fc1-125">In Outlook Mobile, the value returned by `Office.context.mailbox.item.itemId` is a REST-formatted ID and can be used as-is.</span></span>
- <span data-ttu-id="d6fc1-126">В других клиентах Outlook свойство `Office.context.mailbox.item.itemId` возвращает идентификатор в формате EWS, который необходимо преобразовать с помощью метода [Office.context.mailbox.convertToRestId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods).</span><span class="sxs-lookup"><span data-stu-id="d6fc1-126">In other Outlook clients, the value returned by `Office.context.mailbox.item.itemId` is an EWS-formatted ID, and must be converted using the [Office.context.mailbox.convertToRestId](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method.</span></span>
- <span data-ttu-id="d6fc1-127">Обратите внимание: чтобы использовать идентификатор вложения, его нужно преобразовать в идентификатор в формате REST.</span><span class="sxs-lookup"><span data-stu-id="d6fc1-127">Note you must also convert Attachment ID to a REST-formatted ID in order to use it.</span></span> <span data-ttu-id="d6fc1-128">Это преобразование необходимо, потому что идентификаторы EWS могут содержать небезопасные в отношении URL-адресов значения, которые вызывают проблемы с REST.</span><span class="sxs-lookup"><span data-stu-id="d6fc1-128">The reason the IDs must be converted is that EWS IDs can contain non-URL safe values which will cause problems for REST.</span></span>

<span data-ttu-id="d6fc1-129">Надстройка может определить, в каком клиенте Outlook она загружена, с помощью свойства [Office.context.mailbox.diagnostics.hostName](/javascript/api/outlook/office.diagnostics#hostname).</span><span class="sxs-lookup"><span data-stu-id="d6fc1-129">Your add-in can determine which Outlook client it is loaded in by checking the [Office.context.mailbox.diagnostics.hostName](/javascript/api/outlook/office.diagnostics#hostname) property.</span></span>

### <a name="example"></a><span data-ttu-id="d6fc1-130">Пример</span><span class="sxs-lookup"><span data-stu-id="d6fc1-130">Example</span></span>

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

## <a name="get-the-rest-api-url"></a><span data-ttu-id="d6fc1-131">Использование URL-адреса REST API</span><span class="sxs-lookup"><span data-stu-id="d6fc1-131">Get the REST API URL</span></span>

<span data-ttu-id="d6fc1-p108">Последнее значение, необходимое надстройке для вызова REST API, — это имя узла, используемое для отправки запросов API. Оно содержится в свойстве [Office.context.mailbox.restUrl](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties).</span><span class="sxs-lookup"><span data-stu-id="d6fc1-p108">The final piece of information your add-in needs to call the REST API is the hostname it should use to send API requests. This information is in the [Office.context.mailbox.restUrl](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#properties) property.</span></span>

### <a name="example"></a><span data-ttu-id="d6fc1-134">Пример</span><span class="sxs-lookup"><span data-stu-id="d6fc1-134">Example</span></span>

```js
// Example: https://outlook.office.com
var restHost = Office.context.mailbox.restUrl;
```

## <a name="call-the-api"></a><span data-ttu-id="d6fc1-135">Вызов API</span><span class="sxs-lookup"><span data-stu-id="d6fc1-135">Call the API</span></span>

<span data-ttu-id="d6fc1-136">Когда надстройка получит маркер доступа, идентификатор элемента и URL-адрес REST API, она может передать эти сведения внутренней службе, которая вызовет REST API, или вызвать его напрямую с помощью AJAX.</span><span class="sxs-lookup"><span data-stu-id="d6fc1-136">After your add-in has the access token, item ID, and REST API URL, it can either pass that information to a back-end service which calls the REST API, or it can call it directly using AJAX.</span></span> <span data-ttu-id="d6fc1-137">В приведенном ниже примере вызывается REST API почты Outlook для получения текущего сообщения.</span><span class="sxs-lookup"><span data-stu-id="d6fc1-137">The following example calls the Outlook Mail REST API to get the current message.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="d6fc1-138">В локальном развертывании Exchange клиентские запросы с помощью AJAX или аналогичных библиотек сбой, так как CORS не поддерживается в настройке сервера.</span><span class="sxs-lookup"><span data-stu-id="d6fc1-138">For on-premises Exchange deployments, client-side requests using AJAX or similar libraries fail because CORS isn't supported in that server setup.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="d6fc1-139">См. также</span><span class="sxs-lookup"><span data-stu-id="d6fc1-139">See also</span></span>

- <span data-ttu-id="d6fc1-140">Пример вызова REST API из надстроек Outlook: [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo) на сайте GitHub.</span><span class="sxs-lookup"><span data-stu-id="d6fc1-140">For an example that calls the REST APIs from an Outlook add-in, see [command-demo](https://github.com/OfficeDev/outlook-add-in-command-demo) on GitHub.</span></span>
- <span data-ttu-id="d6fc1-141">REST API Outlook также доступны через конечную точку Microsoft Graph, но с некоторыми важными отличиями, включая способ получения надстройкой маркера доступа.</span><span class="sxs-lookup"><span data-stu-id="d6fc1-141">Outlook REST APIs are also available through the Microsoft Graph endpoint but there are some key differences, including how your add-in gets an access token.</span></span> <span data-ttu-id="d6fc1-142">Дополнительные сведения см. в разделе [REST API Outlook через Microsoft Graph](/outlook/rest/index#outlook-rest-api-via-microsoft-graph).</span><span class="sxs-lookup"><span data-stu-id="d6fc1-142">For more information, see [Outlook REST API via Microsoft Graph](/outlook/rest/index#outlook-rest-api-via-microsoft-graph).</span></span>