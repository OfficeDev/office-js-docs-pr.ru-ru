---
ms.date: 09/25/2020
description: Понимание Excel пользовательских функций, которые не используют области задач и их определенное время запуска JavaScript.
title: Время запуска для пользовательских Excel пользовательских функций
localization_priority: Normal
ms.openlocfilehash: aa2cf2632ddf9eb1ad1eb202b031ee2ca686af01
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/09/2021
ms.locfileid: "53349625"
---
# <a name="runtime-for-ui-less-excel-custom-functions"></a><span data-ttu-id="6692c-103">Время запуска для пользовательских Excel пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="6692c-103">Runtime for UI-less Excel custom functions</span></span>

<span data-ttu-id="6692c-104">Настраиваемые функции, которые не используют области задач (не пользовательские функции без пользовательского интерфейса), используют время выполнения JavaScript, предназначенное для оптимизации производительности вычислений.</span><span class="sxs-lookup"><span data-stu-id="6692c-104">Custom functions that don't use a task pane (UI-less custom functions) use a JavaScript runtime that is designed to optimize performance of calculations.</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

<span data-ttu-id="6692c-105">Это время запуска JavaScript предоставляет доступ к API в пространстве имен, которые могут использоваться пользовательскими функциями без пользовательского интерфейса и областью задач для `OfficeRuntime` хранения данных.</span><span class="sxs-lookup"><span data-stu-id="6692c-105">This JavaScript runtime provides access to APIs in the `OfficeRuntime` namespace that can be used by UI-less custom functions and the task pane to store data.</span></span>

## <a name="requesting-external-data"></a><span data-ttu-id="6692c-106">Запрос внешних данных</span><span class="sxs-lookup"><span data-stu-id="6692c-106">Requesting external data</span></span>

<span data-ttu-id="6692c-107">В настраиваемой функции без пользовательского интерфейса можно запрашивать внешние данные с помощью API типа [Fetch](https://developer.mozilla.org/docs/Web/API/Fetch_API) или С помощью [XmlHttpRequest (XHR)](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest)— стандартного веб-API, который выдает HTTP-запросы для взаимодействия с серверами.</span><span class="sxs-lookup"><span data-stu-id="6692c-107">Within a UI-less custom function, you can request external data by using an API like [Fetch](https://developer.mozilla.org/docs/Web/API/Fetch_API) or by using [XmlHttpRequest (XHR)](https://developer.mozilla.org/docs/Web/API/XMLHttpRequest), a standard web API that issues HTTP requests to interact with servers.</span></span>

<span data-ttu-id="6692c-108">Следует помнить, что функции, не требующие пользовательского интерфейса, должны [](https://developer.mozilla.org/docs/Web/Security/Same-origin_policy) применять дополнительные меры безопасности при создании XmlHttpRequests, для которых требуется та же политика происхождения и [простая CORS.](https://www.w3.org/TR/cors/)</span><span class="sxs-lookup"><span data-stu-id="6692c-108">Be aware that UI-less functions must use additional security measures when making XmlHttpRequests, requiring [Same Origin Policy](https://developer.mozilla.org/docs/Web/Security/Same-origin_policy) and simple [CORS](https://www.w3.org/TR/cors/).</span></span>

<span data-ttu-id="6692c-109">Простая реализация CORS не может использовать файлы cookie и поддерживает только простые методы (GET, HEAD, POST).</span><span class="sxs-lookup"><span data-stu-id="6692c-109">A simple CORS implementation cannot use cookies and only supports simple methods (GET, HEAD, POST).</span></span> <span data-ttu-id="6692c-110">Простые запросы CORS принимают простые заголовки с именами полей `Accept`, `Accept-Language`, `Content-Language`.</span><span class="sxs-lookup"><span data-stu-id="6692c-110">Simple CORS accepts simple headers with field names `Accept`, `Accept-Language`, `Content-Language`.</span></span> <span data-ttu-id="6692c-111">Вы также можете использовать `Content-Type` заготку в простой CORS, при условии, что тип контента `application/x-www-form-urlencoded` , `text/plain` или `multipart/form-data` .</span><span class="sxs-lookup"><span data-stu-id="6692c-111">You can also use a `Content-Type` header in simple CORS, provided that the content type is `application/x-www-form-urlencoded`, `text/plain`, or `multipart/form-data`.</span></span>

## <a name="storing-and-accessing-data"></a><span data-ttu-id="6692c-112">Хранения данных и доступ к ним</span><span class="sxs-lookup"><span data-stu-id="6692c-112">Storing and accessing data</span></span>

<span data-ttu-id="6692c-113">В настраиваемой функции без пользовательского интерфейса можно хранить и получать доступ к данным с помощью `OfficeRuntime.storage` объекта.</span><span class="sxs-lookup"><span data-stu-id="6692c-113">Within a UI-less custom function, you can store and access data by using the `OfficeRuntime.storage` object.</span></span> <span data-ttu-id="6692c-114">`Storage` — это система хранения сохраняемой, незашифрованной и ключевой ценности, которая предоставляет альтернативу [localStorage,](https://developer.mozilla.org/docs/Web/API/Window/localStorage)которая не может использоваться пользовательскими функциями без пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="6692c-114">`Storage` is a persistent, unencrypted, key-value storage system that provides an alternative to [localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage), which cannot be used by UI-less custom functions.</span></span> <span data-ttu-id="6692c-115">`Storage` предоставляет 10 МБ данных на домен.</span><span class="sxs-lookup"><span data-stu-id="6692c-115">`Storage` offers 10 MB of data per domain.</span></span> <span data-ttu-id="6692c-116">Домены могут быть общими для более чем одной надстройки.</span><span class="sxs-lookup"><span data-stu-id="6692c-116">Domains can be shared by more than one add-in.</span></span>

<span data-ttu-id="6692c-117">`Storage` предназначается для использования в качестве решения-хранилища с общим доступом. Это означает, что несколько частей надстройки могут выполнять доступ к одним и тем же данным.</span><span class="sxs-lookup"><span data-stu-id="6692c-117">`Storage` is intended as a shared storage solution, meaning multiple parts of an add-in are able to access the same data.</span></span> <span data-ttu-id="6692c-118">Например, маркеры для проверки подлинности пользователей могут храниться, так как к нему можно получить доступ как с помощью настраиваемой функции без пользовательского интерфейса, так и элементов пользовательского интерфейса надстройки, таких как области `storage` задач.</span><span class="sxs-lookup"><span data-stu-id="6692c-118">For example, tokens for user authentication may be stored in `storage` because it can be accessed by both a UI-less custom function and add-in UI elements such as a task pane.</span></span> <span data-ttu-id="6692c-119">Аналогичным образом, если два надстройки имеют один и тот же домен (например, ), им также разрешено обмениваться информацией взад и `www.contoso.com/addin1` `www.contoso.com/addin2` вперед через `storage` .</span><span class="sxs-lookup"><span data-stu-id="6692c-119">Similarly, if two add-ins share the same domain (for example, `www.contoso.com/addin1`, `www.contoso.com/addin2`), they are also permitted to share information back and forth through `storage`.</span></span> <span data-ttu-id="6692c-120">Обратите внимание, что надстройки с разными поддоменами будут иметь различные экземпляры `storage` `subdomain.contoso.com/addin1` (например, , `differentsubdomain.contoso.com/addin2` ).</span><span class="sxs-lookup"><span data-stu-id="6692c-120">Note that add-ins which have different subdomains will have different instances of `storage` (for example, `subdomain.contoso.com/addin1`, `differentsubdomain.contoso.com/addin2`).</span></span>

<span data-ttu-id="6692c-121">Так как `storage` может быть расположением с общим доступом, важно понимать, что можно переопределить пары "ключ-значение".</span><span class="sxs-lookup"><span data-stu-id="6692c-121">Because `storage` can be a shared location, it is important to realize that it is possible to override key-value pairs.</span></span>

<span data-ttu-id="6692c-122">На объекте доступны следующие `storage` методы.</span><span class="sxs-lookup"><span data-stu-id="6692c-122">The following methods are available on the `storage` object.</span></span>

- `getItem`
- `getItems`
- `setItem`
- `setItems`
- `removeItem`
- `removeItems`
- `getKeys`

> [!NOTE]
> <span data-ttu-id="6692c-123">Нет способа очистки всех сведений `clear` (например).</span><span class="sxs-lookup"><span data-stu-id="6692c-123">There's no method for clearing all information (such as `clear`).</span></span> <span data-ttu-id="6692c-124">Вместо этого вам следует использовать `removeItems` для одновременного удаления нескольких записей.</span><span class="sxs-lookup"><span data-stu-id="6692c-124">Instead, you should instead use `removeItems` to remove multiple entries at a time.</span></span>

### <a name="officeruntimestorage-example"></a><span data-ttu-id="6692c-125">Пример OfficeRuntime.storage</span><span class="sxs-lookup"><span data-stu-id="6692c-125">OfficeRuntime.storage example</span></span>

<span data-ttu-id="6692c-126">В следующем примере кода функция вызывает функцию `OfficeRuntime.storage.setItem` для набора ключа и значения `storage` в .</span><span class="sxs-lookup"><span data-stu-id="6692c-126">The following code sample calls the `OfficeRuntime.storage.setItem` function to set a key and value into `storage`.</span></span>

```js
function StoreValue(key, value) {

  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}
```

## <a name="additional-considerations"></a><span data-ttu-id="6692c-127">Дополнительные рекомендации</span><span class="sxs-lookup"><span data-stu-id="6692c-127">Additional considerations</span></span>

<span data-ttu-id="6692c-128">Если надстройка использует только настраиваемые функции без пользовательского интерфейса, обратите внимание, что вы не можете получить доступ к объектной модели документа (DOM) с пользовательскими функциями без пользовательского интерфейса или использовать библиотеки, такие как jQuery, которые полагаются на DOM.</span><span class="sxs-lookup"><span data-stu-id="6692c-128">If your add-in only uses UI-less custom functions, note that you can't access the Document Object Model (DOM) with UI-less custom functions or use libraries like jQuery that rely on the DOM.</span></span>

## <a name="next-steps"></a><span data-ttu-id="6692c-129">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="6692c-129">Next steps</span></span>
<span data-ttu-id="6692c-130">Узнайте, как [отламыть пользовательские функции без пользовательского интерфейса.](custom-functions-debugging.md)</span><span class="sxs-lookup"><span data-stu-id="6692c-130">Learn how to [debug UI-less custom functions](custom-functions-debugging.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="6692c-131">См. также</span><span class="sxs-lookup"><span data-stu-id="6692c-131">See also</span></span>

* [<span data-ttu-id="6692c-132">Проверка подлинности пользовательских функций без пользовательского интерфейса</span><span class="sxs-lookup"><span data-stu-id="6692c-132">Authenticate UI-less custom functions</span></span>](custom-functions-authentication.md)
* [<span data-ttu-id="6692c-133">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="6692c-133">Create custom functions in Excel</span></span>](custom-functions-overview.md)
* [<span data-ttu-id="6692c-134">Руководство по пользовательским функциям</span><span class="sxs-lookup"><span data-stu-id="6692c-134">Custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
