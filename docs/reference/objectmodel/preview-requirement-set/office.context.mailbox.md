---
title: Office. Context. Mailbox — Предварительная версия набора обязательных элементов
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: f2383ea2d2e097b4e2f786bfb1aa8c06ab9eed0e
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127599"
---
# <a name="mailbox"></a><span data-ttu-id="586d9-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="586d9-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="586d9-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="586d9-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="586d9-104">Предоставляет доступ к объектной модели надстройки Outlook для Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="586d9-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="586d9-105">Требования</span><span class="sxs-lookup"><span data-stu-id="586d9-105">Requirements</span></span>

|<span data-ttu-id="586d9-106">Требование</span><span class="sxs-lookup"><span data-stu-id="586d9-106">Requirement</span></span>| <span data-ttu-id="586d9-107">Значение</span><span class="sxs-lookup"><span data-stu-id="586d9-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="586d9-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="586d9-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="586d9-109">1.0</span><span class="sxs-lookup"><span data-stu-id="586d9-109">1.0</span></span>|
|[<span data-ttu-id="586d9-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="586d9-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="586d9-111">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="586d9-111">Restricted</span></span>|
|[<span data-ttu-id="586d9-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="586d9-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="586d9-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="586d9-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="586d9-114">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="586d9-114">Members and methods</span></span>

| <span data-ttu-id="586d9-115">Элемент</span><span class="sxs-lookup"><span data-stu-id="586d9-115">Member</span></span> | <span data-ttu-id="586d9-116">Тип</span><span class="sxs-lookup"><span data-stu-id="586d9-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="586d9-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="586d9-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="586d9-118">Элемент</span><span class="sxs-lookup"><span data-stu-id="586d9-118">Member</span></span> |
| [<span data-ttu-id="586d9-119">Мастеркатегориес</span><span class="sxs-lookup"><span data-stu-id="586d9-119">masterCategories</span></span>](#mastercategories-mastercategories) | <span data-ttu-id="586d9-120">Элемент</span><span class="sxs-lookup"><span data-stu-id="586d9-120">Member</span></span> |
| [<span data-ttu-id="586d9-121">restUrl</span><span class="sxs-lookup"><span data-stu-id="586d9-121">restUrl</span></span>](#resturl-string) | <span data-ttu-id="586d9-122">Элемент</span><span class="sxs-lookup"><span data-stu-id="586d9-122">Member</span></span> |
| [<span data-ttu-id="586d9-123">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="586d9-123">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="586d9-124">Метод</span><span class="sxs-lookup"><span data-stu-id="586d9-124">Method</span></span> |
| [<span data-ttu-id="586d9-125">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="586d9-125">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="586d9-126">Метод</span><span class="sxs-lookup"><span data-stu-id="586d9-126">Method</span></span> |
| [<span data-ttu-id="586d9-127">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="586d9-127">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="586d9-128">Метод</span><span class="sxs-lookup"><span data-stu-id="586d9-128">Method</span></span> |
| [<span data-ttu-id="586d9-129">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="586d9-129">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="586d9-130">Метод</span><span class="sxs-lookup"><span data-stu-id="586d9-130">Method</span></span> |
| [<span data-ttu-id="586d9-131">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="586d9-131">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="586d9-132">Метод</span><span class="sxs-lookup"><span data-stu-id="586d9-132">Method</span></span> |
| [<span data-ttu-id="586d9-133">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="586d9-133">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="586d9-134">Метод</span><span class="sxs-lookup"><span data-stu-id="586d9-134">Method</span></span> |
| [<span data-ttu-id="586d9-135">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="586d9-135">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="586d9-136">Метод</span><span class="sxs-lookup"><span data-stu-id="586d9-136">Method</span></span> |
| [<span data-ttu-id="586d9-137">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="586d9-137">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="586d9-138">Метод</span><span class="sxs-lookup"><span data-stu-id="586d9-138">Method</span></span> |
| [<span data-ttu-id="586d9-139">Дисплайневмессажеформ</span><span class="sxs-lookup"><span data-stu-id="586d9-139">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="586d9-140">Метод</span><span class="sxs-lookup"><span data-stu-id="586d9-140">Method</span></span> |
| [<span data-ttu-id="586d9-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="586d9-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="586d9-142">Метод</span><span class="sxs-lookup"><span data-stu-id="586d9-142">Method</span></span> |
| [<span data-ttu-id="586d9-143">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="586d9-143">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="586d9-144">Метод</span><span class="sxs-lookup"><span data-stu-id="586d9-144">Method</span></span> |
| [<span data-ttu-id="586d9-145">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="586d9-145">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="586d9-146">Метод</span><span class="sxs-lookup"><span data-stu-id="586d9-146">Method</span></span> |
| [<span data-ttu-id="586d9-147">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="586d9-147">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="586d9-148">Метод</span><span class="sxs-lookup"><span data-stu-id="586d9-148">Method</span></span> |
| [<span data-ttu-id="586d9-149">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="586d9-149">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="586d9-150">Метод</span><span class="sxs-lookup"><span data-stu-id="586d9-150">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="586d9-151">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="586d9-151">Namespaces</span></span>

<span data-ttu-id="586d9-152">[diagnostics](Office.context.mailbox.diagnostics.md). Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="586d9-152">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="586d9-153">[item](Office.context.mailbox.item.md). Предоставляет методы и свойства для доступа к сообщению или встрече в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="586d9-153">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="586d9-154">[userProfile](Office.context.mailbox.userProfile.md). Предоставляет сведения о пользователе в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="586d9-154">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="586d9-155">Элементы</span><span class="sxs-lookup"><span data-stu-id="586d9-155">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="586d9-156">ewsUrl: строка</span><span class="sxs-lookup"><span data-stu-id="586d9-156">ewsUrl: String</span></span>

<span data-ttu-id="586d9-157">Получает URL-адрес конечной точки веб-служб Exchange (EWS) для этой учетной записи электронной почты.</span><span class="sxs-lookup"><span data-stu-id="586d9-157">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account.</span></span> <span data-ttu-id="586d9-158">Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="586d9-158">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="586d9-159">Этот элемент не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="586d9-159">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="586d9-p102">Удаленная служба может использовать значение `ewsUrl`, чтобы выполнять вызовы EWS для почтового ящика пользователя. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="586d9-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="586d9-162">Чтобы вызвать элемент `ewsUrl` в режиме чтения, в манифесте приложения должно быть указано разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="586d9-162">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="586d9-p103">Перед использованием элемента `ewsUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="586d9-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="586d9-165">Тип</span><span class="sxs-lookup"><span data-stu-id="586d9-165">Type</span></span>

*   <span data-ttu-id="586d9-166">String</span><span class="sxs-lookup"><span data-stu-id="586d9-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="586d9-167">Требования</span><span class="sxs-lookup"><span data-stu-id="586d9-167">Requirements</span></span>

|<span data-ttu-id="586d9-168">Требование</span><span class="sxs-lookup"><span data-stu-id="586d9-168">Requirement</span></span>| <span data-ttu-id="586d9-169">Значение</span><span class="sxs-lookup"><span data-stu-id="586d9-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="586d9-170">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="586d9-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="586d9-171">1.0</span><span class="sxs-lookup"><span data-stu-id="586d9-171">1.0</span></span>|
|[<span data-ttu-id="586d9-172">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="586d9-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="586d9-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="586d9-173">ReadItem</span></span>|
|[<span data-ttu-id="586d9-174">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="586d9-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="586d9-175">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="586d9-175">Compose or Read</span></span>|

---
---

#### <a name="mastercategories-mastercategoriesjavascriptapioutlookofficemastercategories"></a><span data-ttu-id="586d9-176">Мастеркатегориес: [мастеркатегориес](/javascript/api/outlook/office.mastercategories)</span><span class="sxs-lookup"><span data-stu-id="586d9-176">masterCategories: [MasterCategories](/javascript/api/outlook/office.mastercategories)</span></span>

<span data-ttu-id="586d9-177">Получает объект, предоставляющий методы для управления главным списком категорий в этом почтовом ящике.</span><span class="sxs-lookup"><span data-stu-id="586d9-177">Gets an object that provides methods to manage the categories master list on this mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="586d9-178">Этот элемент не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="586d9-178">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="586d9-179">Тип</span><span class="sxs-lookup"><span data-stu-id="586d9-179">Type</span></span>

*   [<span data-ttu-id="586d9-180">MasterCategories</span><span class="sxs-lookup"><span data-stu-id="586d9-180">MasterCategories</span></span>](/javascript/api/outlook/office.mastercategories)

##### <a name="requirements"></a><span data-ttu-id="586d9-181">Требования</span><span class="sxs-lookup"><span data-stu-id="586d9-181">Requirements</span></span>

|<span data-ttu-id="586d9-182">Требование</span><span class="sxs-lookup"><span data-stu-id="586d9-182">Requirement</span></span>| <span data-ttu-id="586d9-183">Значение</span><span class="sxs-lookup"><span data-stu-id="586d9-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="586d9-184">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="586d9-184">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="586d9-185">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="586d9-185">Preview</span></span> |
|[<span data-ttu-id="586d9-186">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="586d9-186">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="586d9-187">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="586d9-187">ReadWriteMailbox</span></span> |
|[<span data-ttu-id="586d9-188">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="586d9-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="586d9-189">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="586d9-189">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="586d9-190">Пример</span><span class="sxs-lookup"><span data-stu-id="586d9-190">Example</span></span>

<span data-ttu-id="586d9-191">В этом примере показано получение сводного списка категорий для этого почтового ящика.</span><span class="sxs-lookup"><span data-stu-id="586d9-191">This example gets the categories master list for this mailbox.</span></span>

```javascript
Office.context.mailbox.masterCategories.getAsync(function (asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.log("Action failed with error: " + asyncResult.error.message);
  } else {
    console.log("Master categories: " + JSON.stringify(asyncResult.value));
  }
});
```

---
---

#### <a name="resturl-string"></a><span data-ttu-id="586d9-192">Рестурл: строка</span><span class="sxs-lookup"><span data-stu-id="586d9-192">restUrl: String</span></span>

<span data-ttu-id="586d9-193">Возвращает URL-адрес конечной точки REST для этой учетной записи электронной почты.</span><span class="sxs-lookup"><span data-stu-id="586d9-193">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="586d9-194">С помощью значения `restUrl` можно выполнять вызовы [REST API](/outlook/rest/) для почтового ящика пользователя.</span><span class="sxs-lookup"><span data-stu-id="586d9-194">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="586d9-195">Чтобы вызвать элемент `restUrl` в режиме чтения, в манифесте приложения необходимо указать разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="586d9-195">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="586d9-p104">Перед использованием элемента `restUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="586d9-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="586d9-198">Тип</span><span class="sxs-lookup"><span data-stu-id="586d9-198">Type</span></span>

*   <span data-ttu-id="586d9-199">String</span><span class="sxs-lookup"><span data-stu-id="586d9-199">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="586d9-200">Требования</span><span class="sxs-lookup"><span data-stu-id="586d9-200">Requirements</span></span>

|<span data-ttu-id="586d9-201">Требование</span><span class="sxs-lookup"><span data-stu-id="586d9-201">Requirement</span></span>| <span data-ttu-id="586d9-202">Значение</span><span class="sxs-lookup"><span data-stu-id="586d9-202">Value</span></span>|
|---|---|
|[<span data-ttu-id="586d9-203">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="586d9-203">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="586d9-204">1.5</span><span class="sxs-lookup"><span data-stu-id="586d9-204">1.5</span></span> |
|[<span data-ttu-id="586d9-205">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="586d9-205">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="586d9-206">ReadItem</span><span class="sxs-lookup"><span data-stu-id="586d9-206">ReadItem</span></span>|
|[<span data-ttu-id="586d9-207">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="586d9-207">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="586d9-208">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="586d9-208">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="586d9-209">Методы</span><span class="sxs-lookup"><span data-stu-id="586d9-209">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="586d9-210">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="586d9-210">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="586d9-211">Добавляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="586d9-211">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="586d9-212">В настоящее время поддерживаются типы событий `Office.EventType.ItemChanged` и `Office.EventType.OfficeThemeChanged`.</span><span class="sxs-lookup"><span data-stu-id="586d9-212">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="586d9-213">Параметры</span><span class="sxs-lookup"><span data-stu-id="586d9-213">Parameters</span></span>

| <span data-ttu-id="586d9-214">Имя</span><span class="sxs-lookup"><span data-stu-id="586d9-214">Name</span></span> | <span data-ttu-id="586d9-215">Тип</span><span class="sxs-lookup"><span data-stu-id="586d9-215">Type</span></span> | <span data-ttu-id="586d9-216">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="586d9-216">Attributes</span></span> | <span data-ttu-id="586d9-217">Описание</span><span class="sxs-lookup"><span data-stu-id="586d9-217">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="586d9-218">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="586d9-218">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="586d9-219">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="586d9-219">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="586d9-220">Function</span><span class="sxs-lookup"><span data-stu-id="586d9-220">Function</span></span> || <span data-ttu-id="586d9-p105">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="586d9-p105">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="586d9-224">Объект</span><span class="sxs-lookup"><span data-stu-id="586d9-224">Object</span></span> | <span data-ttu-id="586d9-225">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="586d9-225">&lt;optional&gt;</span></span> | <span data-ttu-id="586d9-226">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="586d9-226">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="586d9-227">Object</span><span class="sxs-lookup"><span data-stu-id="586d9-227">Object</span></span> | <span data-ttu-id="586d9-228">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="586d9-228">&lt;optional&gt;</span></span> | <span data-ttu-id="586d9-229">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="586d9-229">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="586d9-230">функция</span><span class="sxs-lookup"><span data-stu-id="586d9-230">function</span></span>| <span data-ttu-id="586d9-231">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="586d9-231">&lt;optional&gt;</span></span>|<span data-ttu-id="586d9-232">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="586d9-232">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="586d9-233">Требования</span><span class="sxs-lookup"><span data-stu-id="586d9-233">Requirements</span></span>

|<span data-ttu-id="586d9-234">Требование</span><span class="sxs-lookup"><span data-stu-id="586d9-234">Requirement</span></span>| <span data-ttu-id="586d9-235">Значение</span><span class="sxs-lookup"><span data-stu-id="586d9-235">Value</span></span>|
|---|---|
|[<span data-ttu-id="586d9-236">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="586d9-236">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="586d9-237">1.5</span><span class="sxs-lookup"><span data-stu-id="586d9-237">1.5</span></span> |
|[<span data-ttu-id="586d9-238">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="586d9-238">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="586d9-239">ReadItem</span><span class="sxs-lookup"><span data-stu-id="586d9-239">ReadItem</span></span> |
|[<span data-ttu-id="586d9-240">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="586d9-240">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="586d9-241">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="586d9-241">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="586d9-242">Пример</span><span class="sxs-lookup"><span data-stu-id="586d9-242">Example</span></span>

```javascript
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error.
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item.
  loadProps(Office.context.mailbox.item);
}
```

---
---

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="586d9-243">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="586d9-243">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="586d9-244">Преобразовывает идентификатор элемента из формата REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="586d9-244">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="586d9-245">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="586d9-245">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="586d9-p106">Формат идентификаторов, извлекаемых через API REST (например, [API Почты Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)), отличается от формата веб-служб Exchange (EWS). Метод `convertToEwsId` преобразовывает идентификатор в формате REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="586d9-p106">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="586d9-248">Параметры</span><span class="sxs-lookup"><span data-stu-id="586d9-248">Parameters</span></span>

|<span data-ttu-id="586d9-249">Имя</span><span class="sxs-lookup"><span data-stu-id="586d9-249">Name</span></span>| <span data-ttu-id="586d9-250">Тип</span><span class="sxs-lookup"><span data-stu-id="586d9-250">Type</span></span>| <span data-ttu-id="586d9-251">Описание</span><span class="sxs-lookup"><span data-stu-id="586d9-251">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="586d9-252">String</span><span class="sxs-lookup"><span data-stu-id="586d9-252">String</span></span>|<span data-ttu-id="586d9-253">Идентификатор элемента в формате REST API для Outlook</span><span class="sxs-lookup"><span data-stu-id="586d9-253">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="586d9-254">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="586d9-254">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="586d9-255">Значение, определяющее версию REST API для Outlook, которая используется для извлечения идентификатора элемента.</span><span class="sxs-lookup"><span data-stu-id="586d9-255">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="586d9-256">Требования</span><span class="sxs-lookup"><span data-stu-id="586d9-256">Requirements</span></span>

|<span data-ttu-id="586d9-257">Требование</span><span class="sxs-lookup"><span data-stu-id="586d9-257">Requirement</span></span>| <span data-ttu-id="586d9-258">Значение</span><span class="sxs-lookup"><span data-stu-id="586d9-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="586d9-259">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="586d9-259">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="586d9-260">1.3</span><span class="sxs-lookup"><span data-stu-id="586d9-260">1.3</span></span>|
|[<span data-ttu-id="586d9-261">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="586d9-261">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="586d9-262">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="586d9-262">Restricted</span></span>|
|[<span data-ttu-id="586d9-263">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="586d9-263">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="586d9-264">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="586d9-264">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="586d9-265">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="586d9-265">Returns:</span></span>

<span data-ttu-id="586d9-266">Тип: String</span><span class="sxs-lookup"><span data-stu-id="586d9-266">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="586d9-267">Пример</span><span class="sxs-lookup"><span data-stu-id="586d9-267">Example</span></span>

```javascript
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime"></a><span data-ttu-id="586d9-268">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="586d9-268">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span></span>

<span data-ttu-id="586d9-269">Получает словарь, содержащий сведения о локальном времени клиента.</span><span class="sxs-lookup"><span data-stu-id="586d9-269">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="586d9-270">Почтовое приложение для Outlook на настольном компьютере или в Интернете может использовать разные часовые пояса для дат и времени.</span><span class="sxs-lookup"><span data-stu-id="586d9-270">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times.</span></span> <span data-ttu-id="586d9-271">Outlook на рабочем столе использует часовой пояс клиентского компьютера; В Outlook в Интернете используется часовой пояс, установленный в центре администрирования Exchange.</span><span class="sxs-lookup"><span data-stu-id="586d9-271">Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC).</span></span> <span data-ttu-id="586d9-272">Значения даты и времени должны обрабатываться таким образом, чтобы значения, отображаемые в интерфейсе пользователя, всегда согласовывались с часовым поясом, ожидаемым пользователем.</span><span class="sxs-lookup"><span data-stu-id="586d9-272">You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="586d9-273">Если почтовое приложение запущено в Outlook на настольном клиенте `convertToLocalClientTime` , метод возвратит объект Dictionary со значениями, заданными для часового пояса клиентского компьютера.</span><span class="sxs-lookup"><span data-stu-id="586d9-273">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone.</span></span> <span data-ttu-id="586d9-274">Если почтовое приложение запущено в Outlook в Интернете, `convertToLocalClientTime` метод возвратит объект Dictionary со значениями, заданными в часовом поясе, заданном в центре администрирования Exchange.</span><span class="sxs-lookup"><span data-stu-id="586d9-274">If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="586d9-275">Параметры</span><span class="sxs-lookup"><span data-stu-id="586d9-275">Parameters</span></span>

|<span data-ttu-id="586d9-276">Имя</span><span class="sxs-lookup"><span data-stu-id="586d9-276">Name</span></span>| <span data-ttu-id="586d9-277">Тип</span><span class="sxs-lookup"><span data-stu-id="586d9-277">Type</span></span>| <span data-ttu-id="586d9-278">Описание</span><span class="sxs-lookup"><span data-stu-id="586d9-278">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="586d9-279">Дата</span><span class="sxs-lookup"><span data-stu-id="586d9-279">Date</span></span>|<span data-ttu-id="586d9-280">Объект Date</span><span class="sxs-lookup"><span data-stu-id="586d9-280">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="586d9-281">Требования</span><span class="sxs-lookup"><span data-stu-id="586d9-281">Requirements</span></span>

|<span data-ttu-id="586d9-282">Требование</span><span class="sxs-lookup"><span data-stu-id="586d9-282">Requirement</span></span>| <span data-ttu-id="586d9-283">Значение</span><span class="sxs-lookup"><span data-stu-id="586d9-283">Value</span></span>|
|---|---|
|[<span data-ttu-id="586d9-284">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="586d9-284">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="586d9-285">1.0</span><span class="sxs-lookup"><span data-stu-id="586d9-285">1.0</span></span>|
|[<span data-ttu-id="586d9-286">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="586d9-286">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="586d9-287">ReadItem</span><span class="sxs-lookup"><span data-stu-id="586d9-287">ReadItem</span></span>|
|[<span data-ttu-id="586d9-288">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="586d9-288">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="586d9-289">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="586d9-289">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="586d9-290">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="586d9-290">Returns:</span></span>

<span data-ttu-id="586d9-291">Тип: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="586d9-291">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span></span>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="586d9-292">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="586d9-292">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="586d9-293">Преобразовывает идентификатор элемента в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="586d9-293">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="586d9-294">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="586d9-294">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="586d9-p109">Формат идентификаторов, извлекаемых через EWS или свойство `itemId`, отличается от формата API REST (таких как [API Почты Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)). Метод `convertToRestId` преобразовывает идентификатор в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="586d9-p109">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="586d9-297">Параметры</span><span class="sxs-lookup"><span data-stu-id="586d9-297">Parameters</span></span>

|<span data-ttu-id="586d9-298">Имя</span><span class="sxs-lookup"><span data-stu-id="586d9-298">Name</span></span>| <span data-ttu-id="586d9-299">Тип</span><span class="sxs-lookup"><span data-stu-id="586d9-299">Type</span></span>| <span data-ttu-id="586d9-300">Описание</span><span class="sxs-lookup"><span data-stu-id="586d9-300">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="586d9-301">String</span><span class="sxs-lookup"><span data-stu-id="586d9-301">String</span></span>|<span data-ttu-id="586d9-302">Идентификатор элемента в формате EWS</span><span class="sxs-lookup"><span data-stu-id="586d9-302">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="586d9-303">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="586d9-303">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="586d9-304">Значение, определяющее версию REST API для Outlook, с которой будет использоваться преобразованный идентификатор.</span><span class="sxs-lookup"><span data-stu-id="586d9-304">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="586d9-305">Требования</span><span class="sxs-lookup"><span data-stu-id="586d9-305">Requirements</span></span>

|<span data-ttu-id="586d9-306">Требование</span><span class="sxs-lookup"><span data-stu-id="586d9-306">Requirement</span></span>| <span data-ttu-id="586d9-307">Значение</span><span class="sxs-lookup"><span data-stu-id="586d9-307">Value</span></span>|
|---|---|
|[<span data-ttu-id="586d9-308">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="586d9-308">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="586d9-309">1.3</span><span class="sxs-lookup"><span data-stu-id="586d9-309">1.3</span></span>|
|[<span data-ttu-id="586d9-310">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="586d9-310">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="586d9-311">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="586d9-311">Restricted</span></span>|
|[<span data-ttu-id="586d9-312">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="586d9-312">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="586d9-313">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="586d9-313">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="586d9-314">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="586d9-314">Returns:</span></span>

<span data-ttu-id="586d9-315">Тип: String</span><span class="sxs-lookup"><span data-stu-id="586d9-315">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="586d9-316">Пример</span><span class="sxs-lookup"><span data-stu-id="586d9-316">Example</span></span>

```javascript
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="586d9-317">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="586d9-317">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="586d9-318">Получает объект Date из словаря, содержащего сведения о времени.</span><span class="sxs-lookup"><span data-stu-id="586d9-318">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="586d9-319">Метод `convertToUtcClientTime` преобразует словарь, содержащий локальную дату и время, в объект Date с правильными значениями локальной даты и времени.</span><span class="sxs-lookup"><span data-stu-id="586d9-319">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="586d9-320">Параметры</span><span class="sxs-lookup"><span data-stu-id="586d9-320">Parameters</span></span>

|<span data-ttu-id="586d9-321">Имя</span><span class="sxs-lookup"><span data-stu-id="586d9-321">Name</span></span>| <span data-ttu-id="586d9-322">Тип</span><span class="sxs-lookup"><span data-stu-id="586d9-322">Type</span></span>| <span data-ttu-id="586d9-323">Описание</span><span class="sxs-lookup"><span data-stu-id="586d9-323">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="586d9-324">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="586d9-324">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime)|<span data-ttu-id="586d9-325">Значение локального времени для преобразования.</span><span class="sxs-lookup"><span data-stu-id="586d9-325">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="586d9-326">Требования</span><span class="sxs-lookup"><span data-stu-id="586d9-326">Requirements</span></span>

|<span data-ttu-id="586d9-327">Требование</span><span class="sxs-lookup"><span data-stu-id="586d9-327">Requirement</span></span>| <span data-ttu-id="586d9-328">Значение</span><span class="sxs-lookup"><span data-stu-id="586d9-328">Value</span></span>|
|---|---|
|[<span data-ttu-id="586d9-329">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="586d9-329">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="586d9-330">1.0</span><span class="sxs-lookup"><span data-stu-id="586d9-330">1.0</span></span>|
|[<span data-ttu-id="586d9-331">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="586d9-331">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="586d9-332">ReadItem</span><span class="sxs-lookup"><span data-stu-id="586d9-332">ReadItem</span></span>|
|[<span data-ttu-id="586d9-333">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="586d9-333">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="586d9-334">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="586d9-334">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="586d9-335">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="586d9-335">Returns:</span></span>

<span data-ttu-id="586d9-336">Объект Date со временем в формате UTC.</span><span class="sxs-lookup"><span data-stu-id="586d9-336">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="586d9-337">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="586d9-337">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="586d9-338">Date</span><span class="sxs-lookup"><span data-stu-id="586d9-338">Date</span></span></dd>

</dl>

---
---

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="586d9-339">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="586d9-339">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="586d9-340">Отображает имеющуюся встречу из календаря.</span><span class="sxs-lookup"><span data-stu-id="586d9-340">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="586d9-341">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="586d9-341">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="586d9-342">Метод `displayAppointmentForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее сведения календаря о существующей встрече.</span><span class="sxs-lookup"><span data-stu-id="586d9-342">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="586d9-343">В Outlook на Mac Этот метод можно использовать для отображения одной встречи, которая не является частью повторяющегося ряда, или главной встречи повторяющейся серии, но невозможно отобразить экземпляр ряда.</span><span class="sxs-lookup"><span data-stu-id="586d9-343">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series.</span></span> <span data-ttu-id="586d9-344">Это связано с тем, что в Outlook на Mac-адресе невозможно получить доступ к свойствам (включая идентификатор элемента) повторяющихся рядов.</span><span class="sxs-lookup"><span data-stu-id="586d9-344">This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="586d9-345">В Outlook в Интернете этот метод открывает указанную форму, только если текст формы меньше или равен 32 КБ числу символов.</span><span class="sxs-lookup"><span data-stu-id="586d9-345">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="586d9-346">Если указанный идентификатор элемента не определяет существующую встречу, на клиентском компьютере или устройстве открывается пустая страница, и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="586d9-346">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="586d9-347">Параметры</span><span class="sxs-lookup"><span data-stu-id="586d9-347">Parameters</span></span>

|<span data-ttu-id="586d9-348">Имя</span><span class="sxs-lookup"><span data-stu-id="586d9-348">Name</span></span>| <span data-ttu-id="586d9-349">Тип</span><span class="sxs-lookup"><span data-stu-id="586d9-349">Type</span></span>| <span data-ttu-id="586d9-350">Описание</span><span class="sxs-lookup"><span data-stu-id="586d9-350">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="586d9-351">String</span><span class="sxs-lookup"><span data-stu-id="586d9-351">String</span></span>|<span data-ttu-id="586d9-352">Идентификатор веб-служб Exchange для существующей встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="586d9-352">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="586d9-353">Требования</span><span class="sxs-lookup"><span data-stu-id="586d9-353">Requirements</span></span>

|<span data-ttu-id="586d9-354">Требование</span><span class="sxs-lookup"><span data-stu-id="586d9-354">Requirement</span></span>| <span data-ttu-id="586d9-355">Значение</span><span class="sxs-lookup"><span data-stu-id="586d9-355">Value</span></span>|
|---|---|
|[<span data-ttu-id="586d9-356">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="586d9-356">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="586d9-357">1.0</span><span class="sxs-lookup"><span data-stu-id="586d9-357">1.0</span></span>|
|[<span data-ttu-id="586d9-358">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="586d9-358">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="586d9-359">ReadItem</span><span class="sxs-lookup"><span data-stu-id="586d9-359">ReadItem</span></span>|
|[<span data-ttu-id="586d9-360">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="586d9-360">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="586d9-361">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="586d9-361">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="586d9-362">Пример</span><span class="sxs-lookup"><span data-stu-id="586d9-362">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="586d9-363">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="586d9-363">displayMessageForm(itemId)</span></span>

<span data-ttu-id="586d9-364">Отображает имеющееся сообщение.</span><span class="sxs-lookup"><span data-stu-id="586d9-364">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="586d9-365">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="586d9-365">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="586d9-366">Метод `displayMessageForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее существующее сообщение.</span><span class="sxs-lookup"><span data-stu-id="586d9-366">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="586d9-367">В Outlook в Интернете этот метод открывает указанную форму только в том случае, если размер текста формы меньше или равен 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="586d9-367">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="586d9-368">Если указанный идентификатор элемента не определяет существующее сообщение, окно на клиентском компьютере не открывается и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="586d9-368">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="586d9-p111">Не используйте `displayMessageForm` с параметром `itemId`, который представляет собой встречу. Используйте метод `displayAppointmentForm`, чтобы отобразить сведения о существующей встрече, а метод `displayNewAppointmentForm` — для отображения формы создания встречи.</span><span class="sxs-lookup"><span data-stu-id="586d9-p111">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="586d9-371">Параметры</span><span class="sxs-lookup"><span data-stu-id="586d9-371">Parameters</span></span>

|<span data-ttu-id="586d9-372">Имя</span><span class="sxs-lookup"><span data-stu-id="586d9-372">Name</span></span>| <span data-ttu-id="586d9-373">Тип</span><span class="sxs-lookup"><span data-stu-id="586d9-373">Type</span></span>| <span data-ttu-id="586d9-374">Описание</span><span class="sxs-lookup"><span data-stu-id="586d9-374">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="586d9-375">String</span><span class="sxs-lookup"><span data-stu-id="586d9-375">String</span></span>|<span data-ttu-id="586d9-376">Идентификатор веб-служб Exchange для существующего сообщения.</span><span class="sxs-lookup"><span data-stu-id="586d9-376">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="586d9-377">Требования</span><span class="sxs-lookup"><span data-stu-id="586d9-377">Requirements</span></span>

|<span data-ttu-id="586d9-378">Требование</span><span class="sxs-lookup"><span data-stu-id="586d9-378">Requirement</span></span>| <span data-ttu-id="586d9-379">Значение</span><span class="sxs-lookup"><span data-stu-id="586d9-379">Value</span></span>|
|---|---|
|[<span data-ttu-id="586d9-380">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="586d9-380">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="586d9-381">1.0</span><span class="sxs-lookup"><span data-stu-id="586d9-381">1.0</span></span>|
|[<span data-ttu-id="586d9-382">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="586d9-382">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="586d9-383">ReadItem</span><span class="sxs-lookup"><span data-stu-id="586d9-383">ReadItem</span></span>|
|[<span data-ttu-id="586d9-384">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="586d9-384">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="586d9-385">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="586d9-385">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="586d9-386">Пример</span><span class="sxs-lookup"><span data-stu-id="586d9-386">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="586d9-387">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="586d9-387">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="586d9-388">Отображает форму для создания новой встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="586d9-388">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="586d9-389">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="586d9-389">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="586d9-p112">Метод `displayNewAppointmentForm` открывает форму, в которой пользователь может создать встречу или собрание. Если параметры заданы, поля формы встречи автоматически заполняются их содержимым.</span><span class="sxs-lookup"><span data-stu-id="586d9-p112">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="586d9-392">В Outlook в Интернете и на мобильных устройствах этот метод всегда отображает форму с полем участников.</span><span class="sxs-lookup"><span data-stu-id="586d9-392">In Outlook on the web and mobile devices, this method always displays a form with an attendees field.</span></span> <span data-ttu-id="586d9-393">Если вы не укажете участников в качестве входных аргументов, метод отображает форму с кнопкой **Сохранить**.</span><span class="sxs-lookup"><span data-stu-id="586d9-393">If you do not specify any attendees as input arguments, the method displays a form with a **Save** button.</span></span> <span data-ttu-id="586d9-394">Если вы укажете участников, форма будет включать участников и кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="586d9-394">If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="586d9-p114">Если вы укажете участников или ресурсы с помощью параметра `requiredAttendees`, `optionalAttendees` или `resources` в клиенте Outlook с расширенными возможностями и Outlook RT, этот метод отобразит форму собрания с кнопкой **Отправить**. Если не указать получателей, этот метод отобразит форму встречи с кнопкой **Сохранить и закрыть**.</span><span class="sxs-lookup"><span data-stu-id="586d9-p114">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="586d9-397">Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.</span><span class="sxs-lookup"><span data-stu-id="586d9-397">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="586d9-398">Параметры</span><span class="sxs-lookup"><span data-stu-id="586d9-398">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="586d9-399">Все параметры являются необязательными.</span><span class="sxs-lookup"><span data-stu-id="586d9-399">All parameters are optional.</span></span>

|<span data-ttu-id="586d9-400">Имя</span><span class="sxs-lookup"><span data-stu-id="586d9-400">Name</span></span>| <span data-ttu-id="586d9-401">Тип</span><span class="sxs-lookup"><span data-stu-id="586d9-401">Type</span></span>| <span data-ttu-id="586d9-402">Описание</span><span class="sxs-lookup"><span data-stu-id="586d9-402">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="586d9-403">Object</span><span class="sxs-lookup"><span data-stu-id="586d9-403">Object</span></span> | <span data-ttu-id="586d9-404">Словарь параметров, описывающий новую встречу.</span><span class="sxs-lookup"><span data-stu-id="586d9-404">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="586d9-405">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="586d9-405">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="586d9-p115">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из обязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="586d9-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="586d9-408">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="586d9-408">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="586d9-p116">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из необязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="586d9-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="586d9-411">Date</span><span class="sxs-lookup"><span data-stu-id="586d9-411">Date</span></span> | <span data-ttu-id="586d9-412">Объект `Date`, указывающий дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="586d9-412">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="586d9-413">Date</span><span class="sxs-lookup"><span data-stu-id="586d9-413">Date</span></span> | <span data-ttu-id="586d9-414">Объект `Date`, указывающий дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="586d9-414">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="586d9-415">String</span><span class="sxs-lookup"><span data-stu-id="586d9-415">String</span></span> | <span data-ttu-id="586d9-p117">Строка со сведениями о месте встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="586d9-p117">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="586d9-418">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="586d9-418">Array.&lt;String&gt;</span></span> | <span data-ttu-id="586d9-p118">Массив строк, содержащий необходимые для встречи ресурсы. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="586d9-p118">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="586d9-421">String</span><span class="sxs-lookup"><span data-stu-id="586d9-421">String</span></span> | <span data-ttu-id="586d9-p119">Строка с темой встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="586d9-p119">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="586d9-424">String</span><span class="sxs-lookup"><span data-stu-id="586d9-424">String</span></span> | <span data-ttu-id="586d9-p120">Текст сообщения о встрече. Максимальный размер содержимого сообщения — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="586d9-p120">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="586d9-427">Требования</span><span class="sxs-lookup"><span data-stu-id="586d9-427">Requirements</span></span>

|<span data-ttu-id="586d9-428">Требование</span><span class="sxs-lookup"><span data-stu-id="586d9-428">Requirement</span></span>| <span data-ttu-id="586d9-429">Значение</span><span class="sxs-lookup"><span data-stu-id="586d9-429">Value</span></span>|
|---|---|
|[<span data-ttu-id="586d9-430">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="586d9-430">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="586d9-431">1.0</span><span class="sxs-lookup"><span data-stu-id="586d9-431">1.0</span></span>|
|[<span data-ttu-id="586d9-432">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="586d9-432">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="586d9-433">ReadItem</span><span class="sxs-lookup"><span data-stu-id="586d9-433">ReadItem</span></span>|
|[<span data-ttu-id="586d9-434">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="586d9-434">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="586d9-435">Чтение</span><span class="sxs-lookup"><span data-stu-id="586d9-435">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="586d9-436">Пример</span><span class="sxs-lookup"><span data-stu-id="586d9-436">Example</span></span>

```javascript
var start = new Date();
var end = new Date();
end.setHours(start.getHours() + 1);

Office.context.mailbox.displayNewAppointmentForm(
  {
    requiredAttendees: ['bob@contoso.com'],
    optionalAttendees: ['sam@contoso.com'],
    start: start,
    end: end,
    location: 'Home',
    resources: ['projector@contoso.com'],
    subject: 'meeting',
    body: 'Hello World!'
  });
```

---
---

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="586d9-437">Дисплайневмессажеформ (Parameters)</span><span class="sxs-lookup"><span data-stu-id="586d9-437">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="586d9-438">Отображает форму для создания нового сообщения.</span><span class="sxs-lookup"><span data-stu-id="586d9-438">Displays a form for creating a new message.</span></span>

<span data-ttu-id="586d9-439">`displayNewMessageForm` Метод открывает форму, которая позволяет пользователю создать новое сообщение.</span><span class="sxs-lookup"><span data-stu-id="586d9-439">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="586d9-440">Если указаны параметры, поля формы сообщения автоматически заполняются содержимым параметров.</span><span class="sxs-lookup"><span data-stu-id="586d9-440">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="586d9-441">Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.</span><span class="sxs-lookup"><span data-stu-id="586d9-441">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="586d9-442">Параметры</span><span class="sxs-lookup"><span data-stu-id="586d9-442">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="586d9-443">Все параметры являются необязательными.</span><span class="sxs-lookup"><span data-stu-id="586d9-443">All parameters are optional.</span></span>

|<span data-ttu-id="586d9-444">Имя</span><span class="sxs-lookup"><span data-stu-id="586d9-444">Name</span></span>| <span data-ttu-id="586d9-445">Тип</span><span class="sxs-lookup"><span data-stu-id="586d9-445">Type</span></span>| <span data-ttu-id="586d9-446">Описание</span><span class="sxs-lookup"><span data-stu-id="586d9-446">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="586d9-447">Object</span><span class="sxs-lookup"><span data-stu-id="586d9-447">Object</span></span> | <span data-ttu-id="586d9-448">Словарь параметров, описывающих новое сообщение.</span><span class="sxs-lookup"><span data-stu-id="586d9-448">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="586d9-449">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="586d9-449">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="586d9-450">Массив строк, содержащий адреса электронной почты или массив, содержащий `EmailAddressDetails` объект для каждого из получателей в строке "Кому".</span><span class="sxs-lookup"><span data-stu-id="586d9-450">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="586d9-451">Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="586d9-451">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="586d9-452">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="586d9-452">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="586d9-453">Массив строк, содержащий адреса электронной почты или массив, содержащий `EmailAddressDetails` объект для каждого получателя в строке "копия".</span><span class="sxs-lookup"><span data-stu-id="586d9-453">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="586d9-454">Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="586d9-454">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="586d9-455">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="586d9-455">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="586d9-456">Массив строк, содержащий адреса электронной почты или массив, содержащий `EmailAddressDetails` объект для каждого из получателей, указанных в строке "СК".</span><span class="sxs-lookup"><span data-stu-id="586d9-456">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="586d9-457">Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="586d9-457">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="586d9-458">String</span><span class="sxs-lookup"><span data-stu-id="586d9-458">String</span></span> | <span data-ttu-id="586d9-459">Строка, содержащая тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="586d9-459">A string containing the subject of the message.</span></span> <span data-ttu-id="586d9-460">Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="586d9-460">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="586d9-461">String</span><span class="sxs-lookup"><span data-stu-id="586d9-461">String</span></span> | <span data-ttu-id="586d9-462">Текст сообщения в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="586d9-462">The HTML body of the message.</span></span> <span data-ttu-id="586d9-463">Максимальный размер содержимого сообщения — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="586d9-463">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="586d9-464">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="586d9-464">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="586d9-465">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="586d9-465">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="586d9-466">String</span><span class="sxs-lookup"><span data-stu-id="586d9-466">String</span></span> | <span data-ttu-id="586d9-p127">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="586d9-p127">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="586d9-469">Строка</span><span class="sxs-lookup"><span data-stu-id="586d9-469">String</span></span> | <span data-ttu-id="586d9-470">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="586d9-470">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="586d9-471">String</span><span class="sxs-lookup"><span data-stu-id="586d9-471">String</span></span> | <span data-ttu-id="586d9-p128">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="586d9-p128">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="586d9-474">Логический</span><span class="sxs-lookup"><span data-stu-id="586d9-474">Boolean</span></span> | <span data-ttu-id="586d9-p129">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="586d9-p129">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="586d9-477">String</span><span class="sxs-lookup"><span data-stu-id="586d9-477">String</span></span> | <span data-ttu-id="586d9-478">Используется, только если свойству `type` присвоено значение `item`.</span><span class="sxs-lookup"><span data-stu-id="586d9-478">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="586d9-479">Идентификатор элемента EWS существующего сообщения электронной почты, которое необходимо присоединить к новому сообщению.</span><span class="sxs-lookup"><span data-stu-id="586d9-479">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="586d9-480">Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="586d9-480">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="586d9-481">Требования</span><span class="sxs-lookup"><span data-stu-id="586d9-481">Requirements</span></span>

|<span data-ttu-id="586d9-482">Требование</span><span class="sxs-lookup"><span data-stu-id="586d9-482">Requirement</span></span>| <span data-ttu-id="586d9-483">Значение</span><span class="sxs-lookup"><span data-stu-id="586d9-483">Value</span></span>|
|---|---|
|[<span data-ttu-id="586d9-484">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="586d9-484">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="586d9-485">1.6</span><span class="sxs-lookup"><span data-stu-id="586d9-485">1.6</span></span> |
|[<span data-ttu-id="586d9-486">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="586d9-486">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="586d9-487">ReadItem</span><span class="sxs-lookup"><span data-stu-id="586d9-487">ReadItem</span></span>|
|[<span data-ttu-id="586d9-488">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="586d9-488">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="586d9-489">Чтение</span><span class="sxs-lookup"><span data-stu-id="586d9-489">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="586d9-490">Пример</span><span class="sxs-lookup"><span data-stu-id="586d9-490">Example</span></span>

```javascript
Office.context.mailbox.displayNewMessageForm(
  {
    // Copy the To line from current item.
    toRecipients: Office.context.mailbox.item.to,
    ccRecipients: ['sam@contoso.com'],
    subject: 'Outlook add-ins are cool!',
    htmlBody: 'Hello <b>World</b>!<br/><img src="cid:image.png"></i>',
    attachments: [
      {
        type: 'file',
        name: 'image.png',
        url: 'http://contoso.com/image.png',
        isInline: true
      }
    ]
  });
```

---
---

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="586d9-491">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="586d9-491">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="586d9-492">Возвращает строку, содержащую маркер, который используется для вызова интерфейсов REST API или веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="586d9-492">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="586d9-p131">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="586d9-p131">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="586d9-495">Рекомендуем сделать так, чтобы по мере возможности надстройки использовали интерфейсы REST API, а не веб-службы Exchange.</span><span class="sxs-lookup"><span data-stu-id="586d9-495">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="586d9-496">**Маркеры REST**</span><span class="sxs-lookup"><span data-stu-id="586d9-496">**REST Tokens**</span></span>

<span data-ttu-id="586d9-p132">Если запрашивается маркер REST (`options.isRest = true`), полученный маркер не подойдет для проверки подлинности при вызовах веб-служб Exchange. Область действия маркера будет ограничена доступом только для чтения к текущему элементу и его вложениям, если в манифесте надстройки не указано разрешение [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission). Если указано разрешение `ReadWriteMailbox`, полученный маркер предоставит доступ на чтение и запись к почте, календарю и контактам, включая возможность отправки почты.</span><span class="sxs-lookup"><span data-stu-id="586d9-p132">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="586d9-500">С помощью свойства `restUrl` надстройка должна определить правильный URL-адрес для вызовов REST API.</span><span class="sxs-lookup"><span data-stu-id="586d9-500">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="586d9-501">**Маркеры EWS**</span><span class="sxs-lookup"><span data-stu-id="586d9-501">**EWS Tokens**</span></span>

<span data-ttu-id="586d9-p133">Если запрашивается маркер EWS (`options.isRest = false`), полученный маркер не подойдет для проверки подлинности при вызовах REST API. Область действия маркера будет ограничена доступом к текущему элементу.</span><span class="sxs-lookup"><span data-stu-id="586d9-p133">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="586d9-504">С помощью свойства `ewsUrl` надстройка должна определить правильный URL-адрес для вызовов EWS.</span><span class="sxs-lookup"><span data-stu-id="586d9-504">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="586d9-505">Параметры</span><span class="sxs-lookup"><span data-stu-id="586d9-505">Parameters</span></span>

|<span data-ttu-id="586d9-506">Имя</span><span class="sxs-lookup"><span data-stu-id="586d9-506">Name</span></span>| <span data-ttu-id="586d9-507">Тип</span><span class="sxs-lookup"><span data-stu-id="586d9-507">Type</span></span>| <span data-ttu-id="586d9-508">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="586d9-508">Attributes</span></span>| <span data-ttu-id="586d9-509">Описание</span><span class="sxs-lookup"><span data-stu-id="586d9-509">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="586d9-510">Объект</span><span class="sxs-lookup"><span data-stu-id="586d9-510">Object</span></span> | <span data-ttu-id="586d9-511">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="586d9-511">&lt;optional&gt;</span></span> | <span data-ttu-id="586d9-512">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="586d9-512">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="586d9-513">Boolean</span><span class="sxs-lookup"><span data-stu-id="586d9-513">Boolean</span></span> |  <span data-ttu-id="586d9-514">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="586d9-514">&lt;optional&gt;</span></span> | <span data-ttu-id="586d9-p134">Определяет, будет ли предоставленный маркер использоваться для интерфейсов REST API Outlook или веб-служб Exchange. Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="586d9-p134">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="586d9-517">Объект</span><span class="sxs-lookup"><span data-stu-id="586d9-517">Object</span></span> |  <span data-ttu-id="586d9-518">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="586d9-518">&lt;optional&gt;</span></span> | <span data-ttu-id="586d9-519">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="586d9-519">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="586d9-520">function</span><span class="sxs-lookup"><span data-stu-id="586d9-520">function</span></span>||<span data-ttu-id="586d9-p135">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult). Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="586d9-p135">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="586d9-523">Требования</span><span class="sxs-lookup"><span data-stu-id="586d9-523">Requirements</span></span>

|<span data-ttu-id="586d9-524">Требование</span><span class="sxs-lookup"><span data-stu-id="586d9-524">Requirement</span></span>| <span data-ttu-id="586d9-525">Значение</span><span class="sxs-lookup"><span data-stu-id="586d9-525">Value</span></span>|
|---|---|
|[<span data-ttu-id="586d9-526">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="586d9-526">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="586d9-527">1.5</span><span class="sxs-lookup"><span data-stu-id="586d9-527">1.5</span></span> |
|[<span data-ttu-id="586d9-528">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="586d9-528">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="586d9-529">ReadItem</span><span class="sxs-lookup"><span data-stu-id="586d9-529">ReadItem</span></span>|
|[<span data-ttu-id="586d9-530">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="586d9-530">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="586d9-531">Создание и чтение</span><span class="sxs-lookup"><span data-stu-id="586d9-531">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="586d9-532">Пример</span><span class="sxs-lookup"><span data-stu-id="586d9-532">Example</span></span>

```javascript
function getCallbackToken() {
  var options = {
    isRest: true,
    asyncContext: { message: 'Hello World!' }
  };

  Office.context.mailbox.getCallbackTokenAsync(options, cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

---
---

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="586d9-533">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="586d9-533">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="586d9-534">Получает строку, содержащую маркер, используемый для получения вложения или элемента с Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="586d9-534">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="586d9-p136">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="586d9-p136">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="586d9-p137">Вы можете передать сторонней системе маркер и идентификатор вложения или элемента. Сторонняя система использует этот маркер как маркер авторизации, чтобы вызвать операцию [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) или [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) веб-служб Exchange для возврата вложения или элемента. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="586d9-p137">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="586d9-540">Для вызова метода `getCallbackTokenAsync` в режиме чтения манифесте приложения должно быть указано разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="586d9-540">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="586d9-p138">Чтобы получить идентификатор элемента для передачи в метод `getCallbackTokenAsync`, в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="586d9-p138">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="586d9-543">Параметры</span><span class="sxs-lookup"><span data-stu-id="586d9-543">Parameters</span></span>

|<span data-ttu-id="586d9-544">Имя</span><span class="sxs-lookup"><span data-stu-id="586d9-544">Name</span></span>| <span data-ttu-id="586d9-545">Тип</span><span class="sxs-lookup"><span data-stu-id="586d9-545">Type</span></span>| <span data-ttu-id="586d9-546">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="586d9-546">Attributes</span></span>| <span data-ttu-id="586d9-547">Описание</span><span class="sxs-lookup"><span data-stu-id="586d9-547">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="586d9-548">function</span><span class="sxs-lookup"><span data-stu-id="586d9-548">function</span></span>||<span data-ttu-id="586d9-p139">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult). Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="586d9-p139">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="586d9-551">Объект</span><span class="sxs-lookup"><span data-stu-id="586d9-551">Object</span></span>| <span data-ttu-id="586d9-552">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="586d9-552">&lt;optional&gt;</span></span>|<span data-ttu-id="586d9-553">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="586d9-553">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="586d9-554">Требования</span><span class="sxs-lookup"><span data-stu-id="586d9-554">Requirements</span></span>

|<span data-ttu-id="586d9-555">Требование</span><span class="sxs-lookup"><span data-stu-id="586d9-555">Requirement</span></span>| <span data-ttu-id="586d9-556">Значение</span><span class="sxs-lookup"><span data-stu-id="586d9-556">Value</span></span>|
|---|---|
|[<span data-ttu-id="586d9-557">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="586d9-557">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="586d9-558">1.3</span><span class="sxs-lookup"><span data-stu-id="586d9-558">1.3</span></span>|
|[<span data-ttu-id="586d9-559">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="586d9-559">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="586d9-560">ReadItem</span><span class="sxs-lookup"><span data-stu-id="586d9-560">ReadItem</span></span>|
|[<span data-ttu-id="586d9-561">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="586d9-561">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="586d9-562">Создание и чтение</span><span class="sxs-lookup"><span data-stu-id="586d9-562">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="586d9-563">Пример</span><span class="sxs-lookup"><span data-stu-id="586d9-563">Example</span></span>

```javascript
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

---
---

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="586d9-564">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="586d9-564">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="586d9-565">Получает маркер, идентифицирующий пользователя и надстройку Office.</span><span class="sxs-lookup"><span data-stu-id="586d9-565">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="586d9-566">Метод `getUserIdentityTokenAsync` возвращает маркер, который можно использовать для идентификации, а также [проверки подлинности надстройки и пользователя в сторонней системе](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="586d9-566">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="586d9-567">Параметры</span><span class="sxs-lookup"><span data-stu-id="586d9-567">Parameters</span></span>

|<span data-ttu-id="586d9-568">Имя</span><span class="sxs-lookup"><span data-stu-id="586d9-568">Name</span></span>| <span data-ttu-id="586d9-569">Тип</span><span class="sxs-lookup"><span data-stu-id="586d9-569">Type</span></span>| <span data-ttu-id="586d9-570">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="586d9-570">Attributes</span></span>| <span data-ttu-id="586d9-571">Описание</span><span class="sxs-lookup"><span data-stu-id="586d9-571">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="586d9-572">функция</span><span class="sxs-lookup"><span data-stu-id="586d9-572">function</span></span>||<span data-ttu-id="586d9-573">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="586d9-573">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="586d9-574">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="586d9-574">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="586d9-575">Object</span><span class="sxs-lookup"><span data-stu-id="586d9-575">Object</span></span>| <span data-ttu-id="586d9-576">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="586d9-576">&lt;optional&gt;</span></span>|<span data-ttu-id="586d9-577">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="586d9-577">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="586d9-578">Требования</span><span class="sxs-lookup"><span data-stu-id="586d9-578">Requirements</span></span>

|<span data-ttu-id="586d9-579">Требование</span><span class="sxs-lookup"><span data-stu-id="586d9-579">Requirement</span></span>| <span data-ttu-id="586d9-580">Значение</span><span class="sxs-lookup"><span data-stu-id="586d9-580">Value</span></span>|
|---|---|
|[<span data-ttu-id="586d9-581">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="586d9-581">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="586d9-582">1.0</span><span class="sxs-lookup"><span data-stu-id="586d9-582">1.0</span></span>|
|[<span data-ttu-id="586d9-583">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="586d9-583">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="586d9-584">ReadItem</span><span class="sxs-lookup"><span data-stu-id="586d9-584">ReadItem</span></span>|
|[<span data-ttu-id="586d9-585">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="586d9-585">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="586d9-586">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="586d9-586">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="586d9-587">Пример</span><span class="sxs-lookup"><span data-stu-id="586d9-587">Example</span></span>

```javascript
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

---
---

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="586d9-588">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="586d9-588">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="586d9-589">Выполняет асинхронный запрос для веб-служб Exchange (EWS) на сервере Exchange Server, на котором размещен почтовый ящик пользователя.</span><span class="sxs-lookup"><span data-stu-id="586d9-589">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="586d9-590">Этот метод не поддерживается в следующих сценариях:</span><span class="sxs-lookup"><span data-stu-id="586d9-590">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="586d9-591">В Outlook на iOS или Android</span><span class="sxs-lookup"><span data-stu-id="586d9-591">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="586d9-592">Если надстройка загружается в почтовый ящик Gmail.</span><span class="sxs-lookup"><span data-stu-id="586d9-592">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="586d9-593">В таких случаях надстройка должна [использовать REST API](/outlook/add-ins/use-rest-api) для доступа к почтовому ящику пользователя.</span><span class="sxs-lookup"><span data-stu-id="586d9-593">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="586d9-594">Метод `makeEwsRequestAsync` отправляет запрос EWS от имени надстройки в Exchange.</span><span class="sxs-lookup"><span data-stu-id="586d9-594">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="586d9-595">Список поддерживаемых операций EWS см. в статье [Вызов веб-служб из надстройки Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support).</span><span class="sxs-lookup"><span data-stu-id="586d9-595">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="586d9-596">С помощью метода `makeEwsRequestAsync` невозможно запрашивать элементы, связанные с папкой.</span><span class="sxs-lookup"><span data-stu-id="586d9-596">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="586d9-597">В запросе XML должна быть указана кодировка UTF-8.</span><span class="sxs-lookup"><span data-stu-id="586d9-597">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="586d9-p141">У вашей надстройки должно быть разрешение **ReadWriteMailbox** для использования метода `makeEwsRequestAsync`. Сведения об использовании разрешения **ReadWriteMailbox** и операций EWS, которые можно вызывать с помощью метода `makeEwsRequestAsync`, см. в статье [Указание разрешений для доступа почтовой надстройки к почтовому ящику пользователя](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="586d9-p141">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="586d9-600">Администратор сервера должен установить значение true для параметра `OAuthAuthentication` в каталоге сервера клиентского доступа EWS, чтобы метод `makeEwsRequestAsync` мог выполнять запросы EWS.</span><span class="sxs-lookup"><span data-stu-id="586d9-600">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="586d9-601">Различия версий</span><span class="sxs-lookup"><span data-stu-id="586d9-601">Version differences</span></span>

<span data-ttu-id="586d9-602">Если вы используете метод `makeEwsRequestAsync` в почтовых приложениях, которые выполняются в Outlook версии более ранней, чем 15.0.4535.1004, указывайте кодировку `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="586d9-602">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="586d9-603">Значение кодировки не нужно указывать, если почтовое приложение выполняется в Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="586d9-603">You do not need to set the encoding value when your mail app is running in Outlook on the web.</span></span> <span data-ttu-id="586d9-604">Вы можете определить, работает ли почтовое приложение в Outlook в Интернете или на настольном клиенте с помощью свойства Mailbox. Diagnostics. hostName.</span><span class="sxs-lookup"><span data-stu-id="586d9-604">You can determine whether your mail app is running in Outlook on the web or a desktop client by using the mailbox.diagnostics.hostName property.</span></span> <span data-ttu-id="586d9-605">Используемую версию Outlook можно определить с помощью свойства mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="586d9-605">You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="586d9-606">Параметры</span><span class="sxs-lookup"><span data-stu-id="586d9-606">Parameters</span></span>

|<span data-ttu-id="586d9-607">Имя</span><span class="sxs-lookup"><span data-stu-id="586d9-607">Name</span></span>| <span data-ttu-id="586d9-608">Тип</span><span class="sxs-lookup"><span data-stu-id="586d9-608">Type</span></span>| <span data-ttu-id="586d9-609">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="586d9-609">Attributes</span></span>| <span data-ttu-id="586d9-610">Описание</span><span class="sxs-lookup"><span data-stu-id="586d9-610">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="586d9-611">String</span><span class="sxs-lookup"><span data-stu-id="586d9-611">String</span></span>||<span data-ttu-id="586d9-612">Запрос EWS.</span><span class="sxs-lookup"><span data-stu-id="586d9-612">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="586d9-613">function</span><span class="sxs-lookup"><span data-stu-id="586d9-613">function</span></span>||<span data-ttu-id="586d9-614">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="586d9-614">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="586d9-615">Результат XML вызова EWS указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="586d9-615">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="586d9-616">Если размер результата превышает 1 МБ, возвращается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="586d9-616">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="586d9-617">Объект</span><span class="sxs-lookup"><span data-stu-id="586d9-617">Object</span></span>| <span data-ttu-id="586d9-618">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="586d9-618">&lt;optional&gt;</span></span>|<span data-ttu-id="586d9-619">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="586d9-619">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="586d9-620">Требования</span><span class="sxs-lookup"><span data-stu-id="586d9-620">Requirements</span></span>

|<span data-ttu-id="586d9-621">Требование</span><span class="sxs-lookup"><span data-stu-id="586d9-621">Requirement</span></span>| <span data-ttu-id="586d9-622">Значение</span><span class="sxs-lookup"><span data-stu-id="586d9-622">Value</span></span>|
|---|---|
|[<span data-ttu-id="586d9-623">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="586d9-623">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="586d9-624">1.0</span><span class="sxs-lookup"><span data-stu-id="586d9-624">1.0</span></span>|
|[<span data-ttu-id="586d9-625">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="586d9-625">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="586d9-626">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="586d9-626">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="586d9-627">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="586d9-627">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="586d9-628">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="586d9-628">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="586d9-629">Пример</span><span class="sxs-lookup"><span data-stu-id="586d9-629">Example</span></span>

<span data-ttu-id="586d9-630">В приведенном ниже примере вызывается `makeEwsRequestAsync` для получения темы элемента с помощью операции `GetItem`.</span><span class="sxs-lookup"><span data-stu-id="586d9-630">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

```javascript
function getSubjectRequest(id) {
  // Return a GetItem operation request for the subject of the specified item.
  var request =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
    '               xmlns:xsd="http://www.w3.org/2001/XMLSchema"' +
    '               xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/"' +
    '               xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types">' +
    '  <soap:Header>' +
    '    <RequestServerVersion Version="Exchange2013" xmlns="http://schemas.microsoft.com/exchange/services/2006/types" soap:mustUnderstand="0" />' +
    '  </soap:Header>' +
    '  <soap:Body>' +
    '    <GetItem xmlns="http://schemas.microsoft.com/exchange/services/2006/messages">' +
    '      <ItemShape>' +
    '        <t:BaseShape>IdOnly</t:BaseShape>' +
    '        <t:AdditionalProperties>' +
    '            <t:FieldURI FieldURI="item:Subject"/>' +
    '        </t:AdditionalProperties>' +
    '      </ItemShape>' +
    '      <ItemIds><t:ItemId Id="' + id + '"/></ItemIds>' +
    '    </GetItem>' +
    '  </soap:Body>' +
    '</soap:Envelope>';

  return request;
}

function sendRequest() {
  // Create a local variable that contains the mailbox.
  Office.context.mailbox.makeEwsRequestAsync(
    getSubjectRequest(mailbox.item.itemId), callback);
}

function callback(asyncResult)  {
  var result = asyncResult.value;
  var context = asyncResult.asyncContext;

  // Process the returned response here.
}
```

---
---

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="586d9-631">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="586d9-631">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="586d9-632">Удаляет обработчиков для поддерживаемого типа события.</span><span class="sxs-lookup"><span data-stu-id="586d9-632">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="586d9-633">В настоящее время поддерживаются типы событий `Office.EventType.ItemChanged` и `Office.EventType.OfficeThemeChanged`.</span><span class="sxs-lookup"><span data-stu-id="586d9-633">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="586d9-634">Параметры</span><span class="sxs-lookup"><span data-stu-id="586d9-634">Parameters</span></span>

| <span data-ttu-id="586d9-635">Имя</span><span class="sxs-lookup"><span data-stu-id="586d9-635">Name</span></span> | <span data-ttu-id="586d9-636">Тип</span><span class="sxs-lookup"><span data-stu-id="586d9-636">Type</span></span> | <span data-ttu-id="586d9-637">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="586d9-637">Attributes</span></span> | <span data-ttu-id="586d9-638">Описание</span><span class="sxs-lookup"><span data-stu-id="586d9-638">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="586d9-639">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="586d9-639">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="586d9-640">Событие, которое должно отменить обработчик.</span><span class="sxs-lookup"><span data-stu-id="586d9-640">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="586d9-641">Объект</span><span class="sxs-lookup"><span data-stu-id="586d9-641">Object</span></span> | <span data-ttu-id="586d9-642">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="586d9-642">&lt;optional&gt;</span></span> | <span data-ttu-id="586d9-643">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="586d9-643">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="586d9-644">Object</span><span class="sxs-lookup"><span data-stu-id="586d9-644">Object</span></span> | <span data-ttu-id="586d9-645">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="586d9-645">&lt;optional&gt;</span></span> | <span data-ttu-id="586d9-646">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="586d9-646">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="586d9-647">функция</span><span class="sxs-lookup"><span data-stu-id="586d9-647">function</span></span>| <span data-ttu-id="586d9-648">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="586d9-648">&lt;optional&gt;</span></span>|<span data-ttu-id="586d9-649">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="586d9-649">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="586d9-650">Требования</span><span class="sxs-lookup"><span data-stu-id="586d9-650">Requirements</span></span>

|<span data-ttu-id="586d9-651">Требование</span><span class="sxs-lookup"><span data-stu-id="586d9-651">Requirement</span></span>| <span data-ttu-id="586d9-652">Значение</span><span class="sxs-lookup"><span data-stu-id="586d9-652">Value</span></span>|
|---|---|
|[<span data-ttu-id="586d9-653">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="586d9-653">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="586d9-654">1.5</span><span class="sxs-lookup"><span data-stu-id="586d9-654">1.5</span></span> |
|[<span data-ttu-id="586d9-655">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="586d9-655">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="586d9-656">ReadItem</span><span class="sxs-lookup"><span data-stu-id="586d9-656">ReadItem</span></span> |
|[<span data-ttu-id="586d9-657">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="586d9-657">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="586d9-658">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="586d9-658">Compose or Read</span></span>|
