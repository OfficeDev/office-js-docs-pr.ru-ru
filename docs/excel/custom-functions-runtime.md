---
ms.date: 05/17/2020
description: Общие сведения о пользовательских функциях Excel, не использующих область задач и определенную среду выполнения JavaScript.
title: Среда выполнения для пользовательских функций Excel без пользовательского интерфейса
localization_priority: Normal
ms.openlocfilehash: 5cb9aa480d6923d31434d58a9683e9a9f5d48458
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609645"
---
# <a name="runtime-for-ui-less-excel-custom-functions"></a><span data-ttu-id="18897-103">Среда выполнения для пользовательских функций Excel без пользовательского интерфейса</span><span class="sxs-lookup"><span data-stu-id="18897-103">Runtime for UI-less Excel custom functions</span></span>

<span data-ttu-id="18897-104">Пользовательские функции, которые не используют область задач (пользовательские функции без пользовательского интерфейса), используют среду выполнения JavaScript, предназначенную для оптимизации производительности вычислений.</span><span class="sxs-lookup"><span data-stu-id="18897-104">Custom functions that don't use a task pane (UI-less custom functions) use a JavaScript runtime that is designed to optimize performance of calculations.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

<span data-ttu-id="18897-105">Эта среда выполнения JavaScript предоставляет доступ к API в `OfficeRuntime` пространстве имен, которые можно использовать в пользовательских функциях без пользовательского интерфейса, и область задач для хранения данных.</span><span class="sxs-lookup"><span data-stu-id="18897-105">This JavaScript runtime provides access to APIs in the `OfficeRuntime` namespace that can be used by UI-less custom functions and the task pane to store data.</span></span>

## <a name="requesting-external-data"></a><span data-ttu-id="18897-106">Запрос внешних данных</span><span class="sxs-lookup"><span data-stu-id="18897-106">Requesting external data</span></span>

<span data-ttu-id="18897-107">В пользовательской функции без пользовательского интерфейса можно запрашивать внешние данные с помощью API, например [Извлечение](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) или с помощью [XMLHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), стандартного веб-API, который отправляет HTTP-запросы для взаимодействия с серверами.</span><span class="sxs-lookup"><span data-stu-id="18897-107">Within a UI-less custom function, you can request external data by using an API like [Fetch](https://developer.mozilla.org/en-US/docs/Web/API/Fetch_API) or by using [XmlHttpRequest (XHR)](https://developer.mozilla.org/en-US/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span>

<span data-ttu-id="18897-108">Имейте в виду, что функции без пользовательского интерфейса должны использовать дополнительные меры безопасности при создании XmlHttpRequest, для чего требуется [одна и та же политика начала](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) и простая [CORS](https://www.w3.org/TR/cors/).</span><span class="sxs-lookup"><span data-stu-id="18897-108">Be aware that UI-less functions must use additional security measures when making XmlHttpRequests, requiring [Same Origin Policy](https://developer.mozilla.org/en-US/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>

<span data-ttu-id="18897-109">Простая реализация CORS не может использовать файлы cookie и поддерживает только простые методы (GET, HEAD, POST).</span><span class="sxs-lookup"><span data-stu-id="18897-109">A simple CORS implementation cannot use cookies and only supports simple methods (GET, HEAD, POST).</span></span> <span data-ttu-id="18897-110">Простые запросы CORS принимают простые заголовки с именами полей `Accept`, `Accept-Language`, `Content-Language`.</span><span class="sxs-lookup"><span data-stu-id="18897-110">Simple CORS accepts simple headers with field names `Accept`, `Accept-Language`, `Content-Language`.</span></span> <span data-ttu-id="18897-111">Вы также можете использовать `Content-Type` заголовок в простой CORS, при условии, что тип контента: `application/x-www-form-urlencoded` , `text/plain` или `multipart/form-data` .</span><span class="sxs-lookup"><span data-stu-id="18897-111">You can also use a `Content-Type` header in simple CORS, provided that the content type is `application/x-www-form-urlencoded`, `text/plain`, or `multipart/form-data`.</span></span>

## <a name="storing-and-accessing-data"></a><span data-ttu-id="18897-112">Хранения данных и доступ к ним</span><span class="sxs-lookup"><span data-stu-id="18897-112">Storing and accessing data</span></span>

<span data-ttu-id="18897-113">В пользовательской функции без пользовательского интерфейса можно хранить и получать доступ к данным с помощью `OfficeRuntime.storage` объекта.</span><span class="sxs-lookup"><span data-stu-id="18897-113">Within a UI-less custom function, you can store and access data by using the `OfficeRuntime.storage` object.</span></span> <span data-ttu-id="18897-114">`Storage`— Это постоянная, незашифрованная система хранения с ключом, которая предоставляет альтернативу [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage), который не может использоваться пользовательскими функциями без пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="18897-114">`Storage` is a persistent, unencrypted, key-value storage system that provides an alternative to [localStorage](https://developer.mozilla.org/en-US/docs/Web/API/Window/localStorage), which cannot be used by UI-less custom functions.</span></span> <span data-ttu-id="18897-115">`Storage`предоставляет 10 МБ данных для каждого домена.</span><span class="sxs-lookup"><span data-stu-id="18897-115">`Storage` offers 10 MB of data per domain.</span></span> <span data-ttu-id="18897-116">Домены могут совместно использоваться несколькими надстройками.</span><span class="sxs-lookup"><span data-stu-id="18897-116">Domains can be shared by more than one add-in.</span></span>

<span data-ttu-id="18897-117">`Storage` предназначается для использования в качестве решения-хранилища с общим доступом. Это означает, что несколько частей надстройки могут выполнять доступ к одним и тем же данным.</span><span class="sxs-lookup"><span data-stu-id="18897-117">`Storage` is intended as a shared storage solution, meaning multiple parts of an add-in are able to access the same data.</span></span> <span data-ttu-id="18897-118">Например, маркеры для проверки подлинности пользователей могут храниться в `storage` связи с тем, что они доступны как в пользовательских функциях, так и в пользовательских элементах пользовательского интерфейса, таких как область задач.</span><span class="sxs-lookup"><span data-stu-id="18897-118">For example, tokens for user authentication may be stored in `storage` because it can be accessed by both a UI-less custom function and add-in UI elements such as a task pane.</span></span> <span data-ttu-id="18897-119">Аналогично, если две надстройки совместно используют один и тот же домен (например, `www.contoso.com/addin1` `www.contoso.com/addin2` ), они также могут обмениваться данными с помощью `storage` .</span><span class="sxs-lookup"><span data-stu-id="18897-119">Similarly, if two add-ins share the same domain (for example, `www.contoso.com/addin1`, `www.contoso.com/addin2`), they are also permitted to share information back and forth through `storage`.</span></span> <span data-ttu-id="18897-120">Обратите внимание, что надстройки, у которых есть разные поддомены, будут иметь разные экземпляры `storage` (например, `subdomain.contoso.com/addin1` `differentsubdomain.contoso.com/addin2` ).</span><span class="sxs-lookup"><span data-stu-id="18897-120">Note that add-ins which have different subdomains will have different instances of `storage` (for example, `subdomain.contoso.com/addin1`, `differentsubdomain.contoso.com/addin2`).</span></span>

<span data-ttu-id="18897-121">Так как `storage` может быть расположением с общим доступом, важно понимать, что можно переопределить пары "ключ-значение".</span><span class="sxs-lookup"><span data-stu-id="18897-121">Because `storage` can be a shared location, it is important to realize that it is possible to override key-value pairs.</span></span>

<span data-ttu-id="18897-122">Ниже указаны методы, доступные в объекте `storage`.</span><span class="sxs-lookup"><span data-stu-id="18897-122">The following methods are available on the `storage` object:</span></span>

 - `getItem`
 - `getItems`
 - `setItem`
 - `setItems`
 - `removeItem`
 - `removeItems`
 - `getKeys`

<span data-ttu-id="18897-123">.</span><span class="sxs-lookup"><span data-stu-id="18897-123">.</span></span>[!NOTE]
> <span data-ttu-id="18897-124">Нет метода для очистки всей информации (например, `clear` ).</span><span class="sxs-lookup"><span data-stu-id="18897-124">There's no method for clearing all information (such as `clear`).</span></span> <span data-ttu-id="18897-125">Вместо этого вам следует использовать `removeItems` для одновременного удаления нескольких записей.</span><span class="sxs-lookup"><span data-stu-id="18897-125">Instead, you should instead use `removeItems` to remove multiple entries at a time.</span></span>

### <a name="officeruntimestorage-example"></a><span data-ttu-id="18897-126">Пример Оффицерунтиме. Storage</span><span class="sxs-lookup"><span data-stu-id="18897-126">OfficeRuntime.storage example</span></span>

<span data-ttu-id="18897-127">В следующем примере кода вызывается `OfficeRuntime.storage.setItem` функция для установки ключа и значения `storage` .</span><span class="sxs-lookup"><span data-stu-id="18897-127">The following code sample calls the `OfficeRuntime.storage.setItem` function to set a key and value into `storage`.</span></span>

```js
function StoreValue(key, value) {

  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}
```

## <a name="additional-considerations"></a><span data-ttu-id="18897-128">Дополнительные рекомендации</span><span class="sxs-lookup"><span data-stu-id="18897-128">Additional considerations</span></span>

<span data-ttu-id="18897-129">Если надстройка использует только пользовательские функции без пользовательского интерфейса, обратите внимание на то, что вы не можете получить доступ к объектной модели документов (DOM) с пользовательскими функциями без пользовательского интерфейса или использовать библиотеки, такие как jQuery, которые используют модель DOM.</span><span class="sxs-lookup"><span data-stu-id="18897-129">If your add-in only uses UI-less custom functions, note that you can't access the Document Object Model (DOM) with UI-less custom functions or use libraries like jQuery that rely on the DOM.</span></span>

## <a name="next-steps"></a><span data-ttu-id="18897-130">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="18897-130">Next steps</span></span>
<span data-ttu-id="18897-131">Узнайте, как [отлаживать пользовательские функции без пользовательского интерфейса](custom-functions-debugging.md).</span><span class="sxs-lookup"><span data-stu-id="18897-131">Learn how to [debug UI-less custom functions](custom-functions-debugging.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="18897-132">См. также</span><span class="sxs-lookup"><span data-stu-id="18897-132">See also</span></span>

* [<span data-ttu-id="18897-133">Проверка подлинности пользовательских функций без пользовательского интерфейса</span><span class="sxs-lookup"><span data-stu-id="18897-133">Authenticate UI-less custom functions</span></span>](custom-functions-authentication.md)
* [<span data-ttu-id="18897-134">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="18897-134">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="18897-135">Руководство по пользовательским функциям</span><span class="sxs-lookup"><span data-stu-id="18897-135">Custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
