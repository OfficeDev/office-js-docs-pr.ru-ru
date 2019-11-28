---
title: Office. Context. Mailbox — набор обязательных элементов 1,8
description: ''
ms.date: 11/27/2019
localization_priority: Normal
ms.openlocfilehash: 908eff7b34e63b62fbe250f1a6f810be69b17627
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629218"
---
# <a name="mailbox"></a><span data-ttu-id="f5323-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="f5323-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="f5323-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="f5323-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="f5323-104">Предоставляет для Microsoft Outlook доступ к объектной модели надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="f5323-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="f5323-105">Требования</span><span class="sxs-lookup"><span data-stu-id="f5323-105">Requirements</span></span>

|<span data-ttu-id="f5323-106">Требование</span><span class="sxs-lookup"><span data-stu-id="f5323-106">Requirement</span></span>| <span data-ttu-id="f5323-107">Значение</span><span class="sxs-lookup"><span data-stu-id="f5323-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="f5323-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f5323-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f5323-109">1.0</span><span class="sxs-lookup"><span data-stu-id="f5323-109">1.0</span></span>|
|[<span data-ttu-id="f5323-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f5323-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f5323-111">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="f5323-111">Restricted</span></span>|
|[<span data-ttu-id="f5323-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f5323-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f5323-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f5323-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="f5323-114">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="f5323-114">Members and methods</span></span>

| <span data-ttu-id="f5323-115">Элемент</span><span class="sxs-lookup"><span data-stu-id="f5323-115">Member</span></span> | <span data-ttu-id="f5323-116">Тип</span><span class="sxs-lookup"><span data-stu-id="f5323-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="f5323-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="f5323-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="f5323-118">Элемент</span><span class="sxs-lookup"><span data-stu-id="f5323-118">Member</span></span> |
| [<span data-ttu-id="f5323-119">мастеркатегориес</span><span class="sxs-lookup"><span data-stu-id="f5323-119">masterCategories</span></span>](#mastercategories-mastercategories) | <span data-ttu-id="f5323-120">Элемент</span><span class="sxs-lookup"><span data-stu-id="f5323-120">Member</span></span> |
| [<span data-ttu-id="f5323-121">restUrl</span><span class="sxs-lookup"><span data-stu-id="f5323-121">restUrl</span></span>](#resturl-string) | <span data-ttu-id="f5323-122">Элемент</span><span class="sxs-lookup"><span data-stu-id="f5323-122">Member</span></span> |
| [<span data-ttu-id="f5323-123">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="f5323-123">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="f5323-124">Метод</span><span class="sxs-lookup"><span data-stu-id="f5323-124">Method</span></span> |
| [<span data-ttu-id="f5323-125">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="f5323-125">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="f5323-126">Метод</span><span class="sxs-lookup"><span data-stu-id="f5323-126">Method</span></span> |
| [<span data-ttu-id="f5323-127">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="f5323-127">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="f5323-128">Метод</span><span class="sxs-lookup"><span data-stu-id="f5323-128">Method</span></span> |
| [<span data-ttu-id="f5323-129">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="f5323-129">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="f5323-130">Метод</span><span class="sxs-lookup"><span data-stu-id="f5323-130">Method</span></span> |
| [<span data-ttu-id="f5323-131">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="f5323-131">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="f5323-132">Метод</span><span class="sxs-lookup"><span data-stu-id="f5323-132">Method</span></span> |
| [<span data-ttu-id="f5323-133">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="f5323-133">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="f5323-134">Метод</span><span class="sxs-lookup"><span data-stu-id="f5323-134">Method</span></span> |
| [<span data-ttu-id="f5323-135">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="f5323-135">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="f5323-136">Метод</span><span class="sxs-lookup"><span data-stu-id="f5323-136">Method</span></span> |
| [<span data-ttu-id="f5323-137">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="f5323-137">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="f5323-138">Метод</span><span class="sxs-lookup"><span data-stu-id="f5323-138">Method</span></span> |
| [<span data-ttu-id="f5323-139">дисплайневмессажеформ</span><span class="sxs-lookup"><span data-stu-id="f5323-139">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="f5323-140">Метод</span><span class="sxs-lookup"><span data-stu-id="f5323-140">Method</span></span> |
| [<span data-ttu-id="f5323-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="f5323-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="f5323-142">Метод</span><span class="sxs-lookup"><span data-stu-id="f5323-142">Method</span></span> |
| [<span data-ttu-id="f5323-143">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="f5323-143">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="f5323-144">Метод</span><span class="sxs-lookup"><span data-stu-id="f5323-144">Method</span></span> |
| [<span data-ttu-id="f5323-145">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="f5323-145">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="f5323-146">Метод</span><span class="sxs-lookup"><span data-stu-id="f5323-146">Method</span></span> |
| [<span data-ttu-id="f5323-147">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="f5323-147">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="f5323-148">Метод</span><span class="sxs-lookup"><span data-stu-id="f5323-148">Method</span></span> |
| [<span data-ttu-id="f5323-149">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="f5323-149">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="f5323-150">Метод</span><span class="sxs-lookup"><span data-stu-id="f5323-150">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="f5323-151">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="f5323-151">Namespaces</span></span>

<span data-ttu-id="f5323-152">[diagnostics](Office.context.mailbox.diagnostics.md). Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="f5323-152">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="f5323-153">[item](Office.context.mailbox.item.md). Предоставляет методы и свойства для доступа к сообщению или встрече в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="f5323-153">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="f5323-154">[userProfile](Office.context.mailbox.userProfile.md). Предоставляет сведения о пользователе в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="f5323-154">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="f5323-155">Members</span><span class="sxs-lookup"><span data-stu-id="f5323-155">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="f5323-156">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="f5323-156">ewsUrl: String</span></span>

<span data-ttu-id="f5323-p101">Получает URL-адрес конечной точки веб-служб Exchange (EWS) для этой учетной записи электронной почты. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="f5323-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="f5323-159">Этот элемент не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="f5323-159">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f5323-p102">Удаленная служба может использовать значение `ewsUrl`, чтобы выполнять вызовы EWS для почтового ящика пользователя. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="f5323-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="f5323-162">Чтобы вызвать элемент `ewsUrl` в режиме чтения, в манифесте приложения должно быть указано разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="f5323-162">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="f5323-p103">Перед использованием элемента `ewsUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="f5323-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="f5323-165">Тип</span><span class="sxs-lookup"><span data-stu-id="f5323-165">Type</span></span>

*   <span data-ttu-id="f5323-166">String</span><span class="sxs-lookup"><span data-stu-id="f5323-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f5323-167">Requirements</span><span class="sxs-lookup"><span data-stu-id="f5323-167">Requirements</span></span>

|<span data-ttu-id="f5323-168">Требование</span><span class="sxs-lookup"><span data-stu-id="f5323-168">Requirement</span></span>| <span data-ttu-id="f5323-169">Значение</span><span class="sxs-lookup"><span data-stu-id="f5323-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="f5323-170">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f5323-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f5323-171">1.0</span><span class="sxs-lookup"><span data-stu-id="f5323-171">1.0</span></span>|
|[<span data-ttu-id="f5323-172">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f5323-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f5323-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f5323-173">ReadItem</span></span>|
|[<span data-ttu-id="f5323-174">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f5323-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f5323-175">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f5323-175">Compose or Read</span></span>|

<br>

---
---

#### <a name="mastercategories-mastercategoriesjavascriptapioutlookofficemastercategoriesviewoutlook-js-18"></a><span data-ttu-id="f5323-176">Мастеркатегориес: [мастеркатегориес](/javascript/api/outlook/office.mastercategories?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="f5323-176">masterCategories: [MasterCategories](/javascript/api/outlook/office.mastercategories?view=outlook-js-1.8)</span></span>

<span data-ttu-id="f5323-177">Получает объект, предоставляющий методы для управления главным списком категорий в этом почтовом ящике.</span><span class="sxs-lookup"><span data-stu-id="f5323-177">Gets an object that provides methods to manage the categories master list on this mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="f5323-178">Этот элемент не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="f5323-178">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="f5323-179">Тип</span><span class="sxs-lookup"><span data-stu-id="f5323-179">Type</span></span>

*   [<span data-ttu-id="f5323-180">MasterCategories</span><span class="sxs-lookup"><span data-stu-id="f5323-180">MasterCategories</span></span>](/javascript/api/outlook/office.mastercategories?view=outlook-js-1.8)

##### <a name="requirements"></a><span data-ttu-id="f5323-181">Requirements</span><span class="sxs-lookup"><span data-stu-id="f5323-181">Requirements</span></span>

|<span data-ttu-id="f5323-182">Требование</span><span class="sxs-lookup"><span data-stu-id="f5323-182">Requirement</span></span>| <span data-ttu-id="f5323-183">Значение</span><span class="sxs-lookup"><span data-stu-id="f5323-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="f5323-184">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="f5323-184">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f5323-185">1.8</span><span class="sxs-lookup"><span data-stu-id="f5323-185">1.8</span></span> |
|[<span data-ttu-id="f5323-186">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f5323-186">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f5323-187">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="f5323-187">ReadWriteMailbox</span></span> |
|[<span data-ttu-id="f5323-188">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f5323-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f5323-189">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f5323-189">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="f5323-190">Пример</span><span class="sxs-lookup"><span data-stu-id="f5323-190">Example</span></span>

<span data-ttu-id="f5323-191">В этом примере показано получение сводного списка категорий для этого почтового ящика.</span><span class="sxs-lookup"><span data-stu-id="f5323-191">This example gets the categories master list for this mailbox.</span></span>

```js
Office.context.mailbox.masterCategories.getAsync(function (asyncResult) {
  if (asyncResult.status === Office.AsyncResultStatus.Failed) {
    console.log("Action failed with error: " + asyncResult.error.message);
  } else {
    console.log("Master categories: " + JSON.stringify(asyncResult.value));
  }
});
```

<br>

---
---

#### <a name="resturl-string"></a><span data-ttu-id="f5323-192">restUrl: String</span><span class="sxs-lookup"><span data-stu-id="f5323-192">restUrl: String</span></span>

<span data-ttu-id="f5323-193">Возвращает URL-адрес конечной точки REST для этой учетной записи электронной почты.</span><span class="sxs-lookup"><span data-stu-id="f5323-193">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="f5323-194">С помощью значения `restUrl` можно выполнять вызовы [REST API](/outlook/rest/) для почтового ящика пользователя.</span><span class="sxs-lookup"><span data-stu-id="f5323-194">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="f5323-195">Тип</span><span class="sxs-lookup"><span data-stu-id="f5323-195">Type</span></span>

*   <span data-ttu-id="f5323-196">String</span><span class="sxs-lookup"><span data-stu-id="f5323-196">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f5323-197">Requirements</span><span class="sxs-lookup"><span data-stu-id="f5323-197">Requirements</span></span>

|<span data-ttu-id="f5323-198">Требование</span><span class="sxs-lookup"><span data-stu-id="f5323-198">Requirement</span></span>| <span data-ttu-id="f5323-199">Значение</span><span class="sxs-lookup"><span data-stu-id="f5323-199">Value</span></span>|
|---|---|
|[<span data-ttu-id="f5323-200">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="f5323-200">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f5323-201">1.5</span><span class="sxs-lookup"><span data-stu-id="f5323-201">1.5</span></span> |
|[<span data-ttu-id="f5323-202">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f5323-202">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f5323-203">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f5323-203">ReadItem</span></span>|
|[<span data-ttu-id="f5323-204">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f5323-204">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f5323-205">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f5323-205">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="f5323-206">Методы</span><span class="sxs-lookup"><span data-stu-id="f5323-206">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="f5323-207">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f5323-207">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="f5323-208">Добавляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="f5323-208">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="f5323-209">В настоящее время поддерживаются типы событий `Office.EventType.ItemChanged` и `Office.EventType.OfficeThemeChanged`.</span><span class="sxs-lookup"><span data-stu-id="f5323-209">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f5323-210">Параметры</span><span class="sxs-lookup"><span data-stu-id="f5323-210">Parameters</span></span>

| <span data-ttu-id="f5323-211">Имя</span><span class="sxs-lookup"><span data-stu-id="f5323-211">Name</span></span> | <span data-ttu-id="f5323-212">Тип</span><span class="sxs-lookup"><span data-stu-id="f5323-212">Type</span></span> | <span data-ttu-id="f5323-213">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="f5323-213">Attributes</span></span> | <span data-ttu-id="f5323-214">Описание</span><span class="sxs-lookup"><span data-stu-id="f5323-214">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="f5323-215">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="f5323-215">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="f5323-216">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="f5323-216">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="f5323-217">Function</span><span class="sxs-lookup"><span data-stu-id="f5323-217">Function</span></span> || <span data-ttu-id="f5323-p104">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="f5323-p104">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="f5323-221">Объект</span><span class="sxs-lookup"><span data-stu-id="f5323-221">Object</span></span> | <span data-ttu-id="f5323-222">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="f5323-222">&lt;optional&gt;</span></span> | <span data-ttu-id="f5323-223">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="f5323-223">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="f5323-224">Object</span><span class="sxs-lookup"><span data-stu-id="f5323-224">Object</span></span> | <span data-ttu-id="f5323-225">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="f5323-225">&lt;optional&gt;</span></span> | <span data-ttu-id="f5323-226">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="f5323-226">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="f5323-227">функция</span><span class="sxs-lookup"><span data-stu-id="f5323-227">function</span></span>| <span data-ttu-id="f5323-228">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="f5323-228">&lt;optional&gt;</span></span>|<span data-ttu-id="f5323-229">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f5323-229">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f5323-230">Требования</span><span class="sxs-lookup"><span data-stu-id="f5323-230">Requirements</span></span>

|<span data-ttu-id="f5323-231">Требование</span><span class="sxs-lookup"><span data-stu-id="f5323-231">Requirement</span></span>| <span data-ttu-id="f5323-232">Значение</span><span class="sxs-lookup"><span data-stu-id="f5323-232">Value</span></span>|
|---|---|
|[<span data-ttu-id="f5323-233">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="f5323-233">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f5323-234">1.5</span><span class="sxs-lookup"><span data-stu-id="f5323-234">1.5</span></span> |
|[<span data-ttu-id="f5323-235">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f5323-235">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f5323-236">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f5323-236">ReadItem</span></span> |
|[<span data-ttu-id="f5323-237">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f5323-237">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f5323-238">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f5323-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f5323-239">Пример</span><span class="sxs-lookup"><span data-stu-id="f5323-239">Example</span></span>

```js
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

<br>

---
---

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="f5323-240">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="f5323-240">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="f5323-241">Преобразовывает идентификатор элемента из формата REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="f5323-241">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="f5323-242">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="f5323-242">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f5323-p105">Формат идентификаторов, извлекаемых через API REST (например, [API Почты Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)), отличается от формата веб-служб Exchange (EWS). Метод `convertToEwsId` преобразовывает идентификатор в формате REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="f5323-p105">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f5323-245">Параметры</span><span class="sxs-lookup"><span data-stu-id="f5323-245">Parameters</span></span>

|<span data-ttu-id="f5323-246">Имя</span><span class="sxs-lookup"><span data-stu-id="f5323-246">Name</span></span>| <span data-ttu-id="f5323-247">Тип</span><span class="sxs-lookup"><span data-stu-id="f5323-247">Type</span></span>| <span data-ttu-id="f5323-248">Описание</span><span class="sxs-lookup"><span data-stu-id="f5323-248">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="f5323-249">String</span><span class="sxs-lookup"><span data-stu-id="f5323-249">String</span></span>|<span data-ttu-id="f5323-250">Идентификатор элемента в формате REST API для Outlook</span><span class="sxs-lookup"><span data-stu-id="f5323-250">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="f5323-251">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="f5323-251">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.8)|<span data-ttu-id="f5323-252">Значение, определяющее версию REST API для Outlook, которая используется для извлечения идентификатора элемента.</span><span class="sxs-lookup"><span data-stu-id="f5323-252">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f5323-253">Requirements</span><span class="sxs-lookup"><span data-stu-id="f5323-253">Requirements</span></span>

|<span data-ttu-id="f5323-254">Требование</span><span class="sxs-lookup"><span data-stu-id="f5323-254">Requirement</span></span>| <span data-ttu-id="f5323-255">Значение</span><span class="sxs-lookup"><span data-stu-id="f5323-255">Value</span></span>|
|---|---|
|[<span data-ttu-id="f5323-256">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="f5323-256">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f5323-257">1.3</span><span class="sxs-lookup"><span data-stu-id="f5323-257">1.3</span></span>|
|[<span data-ttu-id="f5323-258">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f5323-258">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f5323-259">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="f5323-259">Restricted</span></span>|
|[<span data-ttu-id="f5323-260">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f5323-260">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f5323-261">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f5323-261">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f5323-262">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="f5323-262">Returns:</span></span>

<span data-ttu-id="f5323-263">Тип: String</span><span class="sxs-lookup"><span data-stu-id="f5323-263">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="f5323-264">Пример</span><span class="sxs-lookup"><span data-stu-id="f5323-264">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-18"></a><span data-ttu-id="f5323-265">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.8)}</span><span class="sxs-lookup"><span data-stu-id="f5323-265">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.8)}</span></span>

<span data-ttu-id="f5323-266">Получает словарь, содержащий сведения о локальном времени клиента.</span><span class="sxs-lookup"><span data-stu-id="f5323-266">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="f5323-p106">Почтовое приложение для классической версии Outlook или версии в Интернете может использовать разные часовые пояса для дат и времени. Классическое приложение Outlook использует часовой пояс клиентского компьютера. Outlook в Интернете использует часовой пояс, заданный в Центре администрирования Exchange (EAC). Значения даты и времени должны обрабатываться так, чтобы значения в пользовательском интерфейсе всегда согласовывались с часовым поясом, ожидаемым пользователем.</span><span class="sxs-lookup"><span data-stu-id="f5323-p106">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="f5323-p107">Если почтовое приложение работает в классическом клиенте Outlook, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса клиентского компьютера. Если почтовое приложение работает в Outlook в Интернете, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса, заданного в Центре администрирования Exchange.</span><span class="sxs-lookup"><span data-stu-id="f5323-p107">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f5323-272">Параметры</span><span class="sxs-lookup"><span data-stu-id="f5323-272">Parameters</span></span>

|<span data-ttu-id="f5323-273">Имя</span><span class="sxs-lookup"><span data-stu-id="f5323-273">Name</span></span>| <span data-ttu-id="f5323-274">Тип</span><span class="sxs-lookup"><span data-stu-id="f5323-274">Type</span></span>| <span data-ttu-id="f5323-275">Описание</span><span class="sxs-lookup"><span data-stu-id="f5323-275">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="f5323-276">Date</span><span class="sxs-lookup"><span data-stu-id="f5323-276">Date</span></span>|<span data-ttu-id="f5323-277">Объект Date</span><span class="sxs-lookup"><span data-stu-id="f5323-277">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f5323-278">Requirements</span><span class="sxs-lookup"><span data-stu-id="f5323-278">Requirements</span></span>

|<span data-ttu-id="f5323-279">Требование</span><span class="sxs-lookup"><span data-stu-id="f5323-279">Requirement</span></span>| <span data-ttu-id="f5323-280">Значение</span><span class="sxs-lookup"><span data-stu-id="f5323-280">Value</span></span>|
|---|---|
|[<span data-ttu-id="f5323-281">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f5323-281">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f5323-282">1.0</span><span class="sxs-lookup"><span data-stu-id="f5323-282">1.0</span></span>|
|[<span data-ttu-id="f5323-283">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f5323-283">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f5323-284">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f5323-284">ReadItem</span></span>|
|[<span data-ttu-id="f5323-285">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f5323-285">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f5323-286">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f5323-286">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f5323-287">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="f5323-287">Returns:</span></span>

<span data-ttu-id="f5323-288">Тип: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.8)</span><span class="sxs-lookup"><span data-stu-id="f5323-288">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.8)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="f5323-289">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="f5323-289">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="f5323-290">Преобразовывает идентификатор элемента в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="f5323-290">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="f5323-291">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="f5323-291">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f5323-p108">Формат идентификаторов, извлекаемых через EWS или свойство `itemId`, отличается от формата API REST (таких как [API Почты Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)). Метод `convertToRestId` преобразовывает идентификатор в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="f5323-p108">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f5323-294">Параметры</span><span class="sxs-lookup"><span data-stu-id="f5323-294">Parameters</span></span>

|<span data-ttu-id="f5323-295">Имя</span><span class="sxs-lookup"><span data-stu-id="f5323-295">Name</span></span>| <span data-ttu-id="f5323-296">Тип</span><span class="sxs-lookup"><span data-stu-id="f5323-296">Type</span></span>| <span data-ttu-id="f5323-297">Описание</span><span class="sxs-lookup"><span data-stu-id="f5323-297">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="f5323-298">String</span><span class="sxs-lookup"><span data-stu-id="f5323-298">String</span></span>|<span data-ttu-id="f5323-299">Идентификатор элемента в формате EWS</span><span class="sxs-lookup"><span data-stu-id="f5323-299">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="f5323-300">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="f5323-300">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.8)|<span data-ttu-id="f5323-301">Значение, определяющее версию REST API для Outlook, с которой будет использоваться преобразованный идентификатор.</span><span class="sxs-lookup"><span data-stu-id="f5323-301">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f5323-302">Requirements</span><span class="sxs-lookup"><span data-stu-id="f5323-302">Requirements</span></span>

|<span data-ttu-id="f5323-303">Требование</span><span class="sxs-lookup"><span data-stu-id="f5323-303">Requirement</span></span>| <span data-ttu-id="f5323-304">Значение</span><span class="sxs-lookup"><span data-stu-id="f5323-304">Value</span></span>|
|---|---|
|[<span data-ttu-id="f5323-305">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="f5323-305">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f5323-306">1.3</span><span class="sxs-lookup"><span data-stu-id="f5323-306">1.3</span></span>|
|[<span data-ttu-id="f5323-307">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f5323-307">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f5323-308">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="f5323-308">Restricted</span></span>|
|[<span data-ttu-id="f5323-309">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f5323-309">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f5323-310">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f5323-310">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f5323-311">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="f5323-311">Returns:</span></span>

<span data-ttu-id="f5323-312">Тип: String</span><span class="sxs-lookup"><span data-stu-id="f5323-312">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="f5323-313">Пример</span><span class="sxs-lookup"><span data-stu-id="f5323-313">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="f5323-314">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="f5323-314">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="f5323-315">Получает объект Date из словаря, содержащего сведения о времени.</span><span class="sxs-lookup"><span data-stu-id="f5323-315">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="f5323-316">Метод `convertToUtcClientTime` преобразует словарь, содержащий локальную дату и время, в объект Date с правильными значениями локальной даты и времени.</span><span class="sxs-lookup"><span data-stu-id="f5323-316">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f5323-317">Параметры</span><span class="sxs-lookup"><span data-stu-id="f5323-317">Parameters</span></span>

|<span data-ttu-id="f5323-318">Имя</span><span class="sxs-lookup"><span data-stu-id="f5323-318">Name</span></span>| <span data-ttu-id="f5323-319">Тип</span><span class="sxs-lookup"><span data-stu-id="f5323-319">Type</span></span>| <span data-ttu-id="f5323-320">Описание</span><span class="sxs-lookup"><span data-stu-id="f5323-320">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="f5323-321">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="f5323-321">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.8)|<span data-ttu-id="f5323-322">Значение локального времени для преобразования.</span><span class="sxs-lookup"><span data-stu-id="f5323-322">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f5323-323">Requirements</span><span class="sxs-lookup"><span data-stu-id="f5323-323">Requirements</span></span>

|<span data-ttu-id="f5323-324">Требование</span><span class="sxs-lookup"><span data-stu-id="f5323-324">Requirement</span></span>| <span data-ttu-id="f5323-325">Значение</span><span class="sxs-lookup"><span data-stu-id="f5323-325">Value</span></span>|
|---|---|
|[<span data-ttu-id="f5323-326">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f5323-326">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f5323-327">1.0</span><span class="sxs-lookup"><span data-stu-id="f5323-327">1.0</span></span>|
|[<span data-ttu-id="f5323-328">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f5323-328">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f5323-329">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f5323-329">ReadItem</span></span>|
|[<span data-ttu-id="f5323-330">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f5323-330">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f5323-331">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f5323-331">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="f5323-332">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="f5323-332">Returns:</span></span>

<span data-ttu-id="f5323-333">Объект Date со временем в формате UTC.</span><span class="sxs-lookup"><span data-stu-id="f5323-333">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="f5323-334">Тип: Date</span><span class="sxs-lookup"><span data-stu-id="f5323-334">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="f5323-335">Пример</span><span class="sxs-lookup"><span data-stu-id="f5323-335">Example</span></span>

```js
// Represents 3:37 PM PDT on Monday, August 26, 2019.
var input = {
  date: 26,
  hours: 15,
  milliseconds: 2,
  minutes: 37,
  month: 7,
  seconds: 2,
  timezoneOffset: -420,
  year: 2019
};

// result should be a Date object.
var result = Office.context.mailbox.convertToUtcClientTime(input);

// Output should be "2019-08-26T22:37:02.002Z".
console.log(result.toISOString());
```

<br>

---
---

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="f5323-336">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="f5323-336">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="f5323-337">Отображает имеющуюся встречу из календаря.</span><span class="sxs-lookup"><span data-stu-id="f5323-337">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="f5323-338">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="f5323-338">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f5323-339">Метод `displayAppointmentForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее сведения календаря о существующей встрече.</span><span class="sxs-lookup"><span data-stu-id="f5323-339">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="f5323-p109">В Outlook для Mac с помощью этого метода можно отобразить одну встречу, которая не является частью повторяющегося ряда, или основную встречу такого ряда, но не экземпляр из него, так как в Outlook для Mac невозможно получить доступ к свойствам экземпляра повторяющегося ряда (в том числе к идентификатору элемента).</span><span class="sxs-lookup"><span data-stu-id="f5323-p109">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="f5323-342">В Outlook в Интернете этот метод открывает указанную форму, только если текст формы содержит символы размером не более 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="f5323-342">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="f5323-343">Если указанный идентификатор элемента не определяет существующую встречу, на клиентском компьютере или устройстве открывается пустая страница, и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="f5323-343">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f5323-344">Параметры</span><span class="sxs-lookup"><span data-stu-id="f5323-344">Parameters</span></span>

|<span data-ttu-id="f5323-345">Имя</span><span class="sxs-lookup"><span data-stu-id="f5323-345">Name</span></span>| <span data-ttu-id="f5323-346">Тип</span><span class="sxs-lookup"><span data-stu-id="f5323-346">Type</span></span>| <span data-ttu-id="f5323-347">Описание</span><span class="sxs-lookup"><span data-stu-id="f5323-347">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="f5323-348">String</span><span class="sxs-lookup"><span data-stu-id="f5323-348">String</span></span>|<span data-ttu-id="f5323-349">Идентификатор веб-служб Exchange для существующей встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="f5323-349">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f5323-350">Требования</span><span class="sxs-lookup"><span data-stu-id="f5323-350">Requirements</span></span>

|<span data-ttu-id="f5323-351">Требование</span><span class="sxs-lookup"><span data-stu-id="f5323-351">Requirement</span></span>| <span data-ttu-id="f5323-352">Значение</span><span class="sxs-lookup"><span data-stu-id="f5323-352">Value</span></span>|
|---|---|
|[<span data-ttu-id="f5323-353">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f5323-353">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f5323-354">1.0</span><span class="sxs-lookup"><span data-stu-id="f5323-354">1.0</span></span>|
|[<span data-ttu-id="f5323-355">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f5323-355">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f5323-356">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f5323-356">ReadItem</span></span>|
|[<span data-ttu-id="f5323-357">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f5323-357">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f5323-358">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f5323-358">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f5323-359">Пример</span><span class="sxs-lookup"><span data-stu-id="f5323-359">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="f5323-360">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="f5323-360">displayMessageForm(itemId)</span></span>

<span data-ttu-id="f5323-361">Отображает имеющееся сообщение.</span><span class="sxs-lookup"><span data-stu-id="f5323-361">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="f5323-362">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="f5323-362">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f5323-363">Метод `displayMessageForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее существующее сообщение.</span><span class="sxs-lookup"><span data-stu-id="f5323-363">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="f5323-364">В Outlook в Интернете этот метод открывает указанную форму, только если текст формы содержит символы размером не более 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="f5323-364">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="f5323-365">Если указанный идентификатор элемента не определяет существующее сообщение, окно на клиентском компьютере не открывается и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="f5323-365">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="f5323-p110">Не используйте `displayMessageForm` с параметром `itemId`, который представляет собой встречу. Используйте метод `displayAppointmentForm`, чтобы отобразить сведения о существующей встрече, а метод `displayNewAppointmentForm` — для отображения формы создания встречи.</span><span class="sxs-lookup"><span data-stu-id="f5323-p110">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f5323-368">Параметры</span><span class="sxs-lookup"><span data-stu-id="f5323-368">Parameters</span></span>

|<span data-ttu-id="f5323-369">Имя</span><span class="sxs-lookup"><span data-stu-id="f5323-369">Name</span></span>| <span data-ttu-id="f5323-370">Тип</span><span class="sxs-lookup"><span data-stu-id="f5323-370">Type</span></span>| <span data-ttu-id="f5323-371">Описание</span><span class="sxs-lookup"><span data-stu-id="f5323-371">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="f5323-372">String</span><span class="sxs-lookup"><span data-stu-id="f5323-372">String</span></span>|<span data-ttu-id="f5323-373">Идентификатор веб-служб Exchange для существующего сообщения.</span><span class="sxs-lookup"><span data-stu-id="f5323-373">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f5323-374">Requirements</span><span class="sxs-lookup"><span data-stu-id="f5323-374">Requirements</span></span>

|<span data-ttu-id="f5323-375">Требование</span><span class="sxs-lookup"><span data-stu-id="f5323-375">Requirement</span></span>| <span data-ttu-id="f5323-376">Значение</span><span class="sxs-lookup"><span data-stu-id="f5323-376">Value</span></span>|
|---|---|
|[<span data-ttu-id="f5323-377">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f5323-377">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f5323-378">1.0</span><span class="sxs-lookup"><span data-stu-id="f5323-378">1.0</span></span>|
|[<span data-ttu-id="f5323-379">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f5323-379">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f5323-380">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f5323-380">ReadItem</span></span>|
|[<span data-ttu-id="f5323-381">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f5323-381">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f5323-382">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f5323-382">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f5323-383">Пример</span><span class="sxs-lookup"><span data-stu-id="f5323-383">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="f5323-384">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="f5323-384">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="f5323-385">Отображает форму для создания новой встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="f5323-385">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="f5323-386">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="f5323-386">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="f5323-p111">Метод `displayNewAppointmentForm` открывает форму, в которой пользователь может создать встречу или собрание. Если параметры заданы, поля формы встречи автоматически заполняются их содержимым.</span><span class="sxs-lookup"><span data-stu-id="f5323-p111">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="f5323-p112">В Outlook в Интернете и на мобильных устройствах этот метод всегда отображает форму с полем участников. Если вы не укажете участников в качестве входных аргументов, метод отображает форму с кнопкой **Сохранить**. Если вы укажете участников, форма будет включать участников и кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="f5323-p112">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="f5323-p113">Если вы укажете участников или ресурсы с помощью параметра `requiredAttendees`, `optionalAttendees` или `resources` в клиенте Outlook с расширенными возможностями и Outlook RT, этот метод отобразит форму собрания с кнопкой **Отправить**. Если не указать получателей, этот метод отобразит форму встречи с кнопкой **Сохранить и закрыть**.</span><span class="sxs-lookup"><span data-stu-id="f5323-p113">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="f5323-394">Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.</span><span class="sxs-lookup"><span data-stu-id="f5323-394">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f5323-395">Параметры</span><span class="sxs-lookup"><span data-stu-id="f5323-395">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="f5323-396">Все параметры являются необязательными.</span><span class="sxs-lookup"><span data-stu-id="f5323-396">All parameters are optional.</span></span>

|<span data-ttu-id="f5323-397">Имя</span><span class="sxs-lookup"><span data-stu-id="f5323-397">Name</span></span>| <span data-ttu-id="f5323-398">Тип</span><span class="sxs-lookup"><span data-stu-id="f5323-398">Type</span></span>| <span data-ttu-id="f5323-399">Описание</span><span class="sxs-lookup"><span data-stu-id="f5323-399">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="f5323-400">Object</span><span class="sxs-lookup"><span data-stu-id="f5323-400">Object</span></span> | <span data-ttu-id="f5323-401">Словарь параметров, описывающий новую встречу.</span><span class="sxs-lookup"><span data-stu-id="f5323-401">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="f5323-402">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span><span class="sxs-lookup"><span data-stu-id="f5323-402">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span></span> | <span data-ttu-id="f5323-p114">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из обязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="f5323-p114">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="f5323-405">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span><span class="sxs-lookup"><span data-stu-id="f5323-405">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span></span> | <span data-ttu-id="f5323-p115">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из необязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="f5323-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="f5323-408">Date</span><span class="sxs-lookup"><span data-stu-id="f5323-408">Date</span></span> | <span data-ttu-id="f5323-409">Объект `Date`, указывающий дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="f5323-409">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="f5323-410">Date</span><span class="sxs-lookup"><span data-stu-id="f5323-410">Date</span></span> | <span data-ttu-id="f5323-411">Объект `Date`, указывающий дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="f5323-411">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="f5323-412">String</span><span class="sxs-lookup"><span data-stu-id="f5323-412">String</span></span> | <span data-ttu-id="f5323-p116">Строка со сведениями о месте встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="f5323-p116">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="f5323-415">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="f5323-415">Array.&lt;String&gt;</span></span> | <span data-ttu-id="f5323-p117">Массив строк, содержащий необходимые для встречи ресурсы. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="f5323-p117">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="f5323-418">String</span><span class="sxs-lookup"><span data-stu-id="f5323-418">String</span></span> | <span data-ttu-id="f5323-p118">Строка с темой встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="f5323-p118">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="f5323-421">String</span><span class="sxs-lookup"><span data-stu-id="f5323-421">String</span></span> | <span data-ttu-id="f5323-p119">Текст сообщения о встрече. Максимальный размер содержимого сообщения — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="f5323-p119">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="f5323-424">Требования</span><span class="sxs-lookup"><span data-stu-id="f5323-424">Requirements</span></span>

|<span data-ttu-id="f5323-425">Требование</span><span class="sxs-lookup"><span data-stu-id="f5323-425">Requirement</span></span>| <span data-ttu-id="f5323-426">Значение</span><span class="sxs-lookup"><span data-stu-id="f5323-426">Value</span></span>|
|---|---|
|[<span data-ttu-id="f5323-427">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f5323-427">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f5323-428">1.0</span><span class="sxs-lookup"><span data-stu-id="f5323-428">1.0</span></span>|
|[<span data-ttu-id="f5323-429">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f5323-429">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f5323-430">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f5323-430">ReadItem</span></span>|
|[<span data-ttu-id="f5323-431">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f5323-431">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f5323-432">Чтение</span><span class="sxs-lookup"><span data-stu-id="f5323-432">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f5323-433">Пример</span><span class="sxs-lookup"><span data-stu-id="f5323-433">Example</span></span>

```js
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

<br>

---
---

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="f5323-434">Дисплайневмессажеформ (Parameters)</span><span class="sxs-lookup"><span data-stu-id="f5323-434">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="f5323-435">Отображает форму для создания нового сообщения.</span><span class="sxs-lookup"><span data-stu-id="f5323-435">Displays a form for creating a new message.</span></span>

<span data-ttu-id="f5323-436">`displayNewMessageForm` Метод открывает форму, которая позволяет пользователю создать новое сообщение.</span><span class="sxs-lookup"><span data-stu-id="f5323-436">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="f5323-437">Если указаны параметры, поля формы сообщения автоматически заполняются содержимым параметров.</span><span class="sxs-lookup"><span data-stu-id="f5323-437">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="f5323-438">Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.</span><span class="sxs-lookup"><span data-stu-id="f5323-438">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f5323-439">Параметры</span><span class="sxs-lookup"><span data-stu-id="f5323-439">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="f5323-440">Все параметры являются необязательными.</span><span class="sxs-lookup"><span data-stu-id="f5323-440">All parameters are optional.</span></span>

|<span data-ttu-id="f5323-441">Имя</span><span class="sxs-lookup"><span data-stu-id="f5323-441">Name</span></span>| <span data-ttu-id="f5323-442">Тип</span><span class="sxs-lookup"><span data-stu-id="f5323-442">Type</span></span>| <span data-ttu-id="f5323-443">Описание</span><span class="sxs-lookup"><span data-stu-id="f5323-443">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="f5323-444">Object</span><span class="sxs-lookup"><span data-stu-id="f5323-444">Object</span></span> | <span data-ttu-id="f5323-445">Словарь параметров, описывающих новое сообщение.</span><span class="sxs-lookup"><span data-stu-id="f5323-445">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="f5323-446">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span><span class="sxs-lookup"><span data-stu-id="f5323-446">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span></span> | <span data-ttu-id="f5323-447">Массив строк, содержащий адреса электронной почты или массив, содержащий `EmailAddressDetails` объект для каждого из получателей в строке "Кому".</span><span class="sxs-lookup"><span data-stu-id="f5323-447">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="f5323-448">Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="f5323-448">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="f5323-449">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span><span class="sxs-lookup"><span data-stu-id="f5323-449">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span></span> | <span data-ttu-id="f5323-450">Массив строк, содержащий адреса электронной почты или массив, содержащий `EmailAddressDetails` объект для каждого получателя в строке "копия".</span><span class="sxs-lookup"><span data-stu-id="f5323-450">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="f5323-451">Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="f5323-451">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="f5323-452">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span><span class="sxs-lookup"><span data-stu-id="f5323-452">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.8)&gt;</span></span> | <span data-ttu-id="f5323-453">Массив строк, содержащий адреса электронной почты или массив, содержащий `EmailAddressDetails` объект для каждого из получателей, указанных в строке "СК".</span><span class="sxs-lookup"><span data-stu-id="f5323-453">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="f5323-454">Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="f5323-454">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="f5323-455">String</span><span class="sxs-lookup"><span data-stu-id="f5323-455">String</span></span> | <span data-ttu-id="f5323-456">Строка, содержащая тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="f5323-456">A string containing the subject of the message.</span></span> <span data-ttu-id="f5323-457">Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="f5323-457">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="f5323-458">String</span><span class="sxs-lookup"><span data-stu-id="f5323-458">String</span></span> | <span data-ttu-id="f5323-459">Текст сообщения в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="f5323-459">The HTML body of the message.</span></span> <span data-ttu-id="f5323-460">Максимальный размер содержимого сообщения — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="f5323-460">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="f5323-461">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="f5323-461">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="f5323-462">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="f5323-462">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="f5323-463">String</span><span class="sxs-lookup"><span data-stu-id="f5323-463">String</span></span> | <span data-ttu-id="f5323-p126">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="f5323-p126">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="f5323-466">Строка</span><span class="sxs-lookup"><span data-stu-id="f5323-466">String</span></span> | <span data-ttu-id="f5323-467">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="f5323-467">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="f5323-468">String</span><span class="sxs-lookup"><span data-stu-id="f5323-468">String</span></span> | <span data-ttu-id="f5323-p127">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="f5323-p127">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="f5323-471">Логический</span><span class="sxs-lookup"><span data-stu-id="f5323-471">Boolean</span></span> | <span data-ttu-id="f5323-p128">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="f5323-p128">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="f5323-474">Строка</span><span class="sxs-lookup"><span data-stu-id="f5323-474">String</span></span> | <span data-ttu-id="f5323-475">Используется, только если свойству `type` присвоено значение `item`.</span><span class="sxs-lookup"><span data-stu-id="f5323-475">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="f5323-476">Идентификатор элемента EWS существующего сообщения электронной почты, которое необходимо присоединить к новому сообщению.</span><span class="sxs-lookup"><span data-stu-id="f5323-476">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="f5323-477">Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="f5323-477">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="f5323-478">Requirements</span><span class="sxs-lookup"><span data-stu-id="f5323-478">Requirements</span></span>

|<span data-ttu-id="f5323-479">Требование</span><span class="sxs-lookup"><span data-stu-id="f5323-479">Requirement</span></span>| <span data-ttu-id="f5323-480">Значение</span><span class="sxs-lookup"><span data-stu-id="f5323-480">Value</span></span>|
|---|---|
|[<span data-ttu-id="f5323-481">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="f5323-481">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f5323-482">1.6</span><span class="sxs-lookup"><span data-stu-id="f5323-482">1.6</span></span> |
|[<span data-ttu-id="f5323-483">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f5323-483">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f5323-484">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f5323-484">ReadItem</span></span>|
|[<span data-ttu-id="f5323-485">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f5323-485">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f5323-486">Чтение</span><span class="sxs-lookup"><span data-stu-id="f5323-486">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f5323-487">Пример</span><span class="sxs-lookup"><span data-stu-id="f5323-487">Example</span></span>

```js
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

<br>

---
---

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="f5323-488">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="f5323-488">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="f5323-489">Возвращает строку, содержащую маркер, который используется для вызова интерфейсов REST API или веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="f5323-489">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="f5323-p130">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="f5323-p130">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="f5323-492">Рекомендуем сделать так, чтобы по мере возможности надстройки использовали интерфейсы REST API, а не веб-службы Exchange.</span><span class="sxs-lookup"><span data-stu-id="f5323-492">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="f5323-493">Для вызова метода `getCallbackTokenAsync` в режиме чтения требуется минимальный уровень разрешения **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="f5323-493">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="f5323-494">Для вызова `getCallbackTokenAsync` в режиме создания сообщения требуется сохранить элемент.</span><span class="sxs-lookup"><span data-stu-id="f5323-494">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="f5323-495">Для метода [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) требуется минимальный уровень разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="f5323-495">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

<span data-ttu-id="f5323-496">**Маркеры REST**</span><span class="sxs-lookup"><span data-stu-id="f5323-496">**REST Tokens**</span></span>

<span data-ttu-id="f5323-p132">Если запрашивается маркер REST (`options.isRest = true`), полученный маркер не подойдет для проверки подлинности при вызовах веб-служб Exchange. Область действия маркера будет ограничена доступом только для чтения к текущему элементу и его вложениям, если в манифесте надстройки не указано разрешение [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission). Если указано разрешение `ReadWriteMailbox`, полученный маркер предоставит доступ на чтение и запись к почте, календарю и контактам, включая возможность отправки почты.</span><span class="sxs-lookup"><span data-stu-id="f5323-p132">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="f5323-500">С помощью свойства `restUrl` надстройка должна определить правильный URL-адрес для вызовов REST API.</span><span class="sxs-lookup"><span data-stu-id="f5323-500">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="f5323-501">**Маркеры EWS**</span><span class="sxs-lookup"><span data-stu-id="f5323-501">**EWS Tokens**</span></span>

<span data-ttu-id="f5323-p133">Если запрашивается маркер EWS (`options.isRest = false`), полученный маркер не подойдет для проверки подлинности при вызовах REST API. Область действия маркера будет ограничена доступом к текущему элементу.</span><span class="sxs-lookup"><span data-stu-id="f5323-p133">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="f5323-504">С помощью свойства `ewsUrl` надстройка должна определить правильный URL-адрес для вызовов EWS.</span><span class="sxs-lookup"><span data-stu-id="f5323-504">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

<span data-ttu-id="f5323-505">Вы можете передать сторонней системе маркер и идентификатор вложения или элемента.</span><span class="sxs-lookup"><span data-stu-id="f5323-505">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="f5323-506">Третья система использует маркер в качестве маркера авторизации носителя, чтобы вызвать операцию [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) [или GetItem](/exchange/client-developer/web-service-reference/getitem-operation) веб-служб Exchange (EWS) для получения вложения или элемента.</span><span class="sxs-lookup"><span data-stu-id="f5323-506">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to retrieve an attachment or item.</span></span> <span data-ttu-id="f5323-507">Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="f5323-507">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="parameters"></a><span data-ttu-id="f5323-508">Параметры</span><span class="sxs-lookup"><span data-stu-id="f5323-508">Parameters</span></span>

|<span data-ttu-id="f5323-509">Имя</span><span class="sxs-lookup"><span data-stu-id="f5323-509">Name</span></span>| <span data-ttu-id="f5323-510">Тип</span><span class="sxs-lookup"><span data-stu-id="f5323-510">Type</span></span>| <span data-ttu-id="f5323-511">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="f5323-511">Attributes</span></span>| <span data-ttu-id="f5323-512">Описание</span><span class="sxs-lookup"><span data-stu-id="f5323-512">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="f5323-513">Object</span><span class="sxs-lookup"><span data-stu-id="f5323-513">Object</span></span> | <span data-ttu-id="f5323-514">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="f5323-514">&lt;optional&gt;</span></span> | <span data-ttu-id="f5323-515">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="f5323-515">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="f5323-516">Boolean</span><span class="sxs-lookup"><span data-stu-id="f5323-516">Boolean</span></span> |  <span data-ttu-id="f5323-517">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="f5323-517">&lt;optional&gt;</span></span> | <span data-ttu-id="f5323-p135">Определяет, будет ли предоставленный маркер использоваться для интерфейсов REST API Outlook или веб-служб Exchange. Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="f5323-p135">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="f5323-520">Объект</span><span class="sxs-lookup"><span data-stu-id="f5323-520">Object</span></span> |  <span data-ttu-id="f5323-521">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="f5323-521">&lt;optional&gt;</span></span> | <span data-ttu-id="f5323-522">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="f5323-522">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="f5323-523">функция</span><span class="sxs-lookup"><span data-stu-id="f5323-523">function</span></span>||<span data-ttu-id="f5323-524">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f5323-524">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f5323-525">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="f5323-525">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="f5323-526">При наличии ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут предоставлять дополнительные сведения.</span><span class="sxs-lookup"><span data-stu-id="f5323-526">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f5323-527">Ошибки</span><span class="sxs-lookup"><span data-stu-id="f5323-527">Errors</span></span>

|<span data-ttu-id="f5323-528">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="f5323-528">Error code</span></span>|<span data-ttu-id="f5323-529">Описание</span><span class="sxs-lookup"><span data-stu-id="f5323-529">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="f5323-530">Не удалось выполнить запрос.</span><span class="sxs-lookup"><span data-stu-id="f5323-530">The request has failed.</span></span> <span data-ttu-id="f5323-531">Просмотрите объект диагностики для кода ошибки HTTP.</span><span class="sxs-lookup"><span data-stu-id="f5323-531">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="f5323-532">Сервер Exchange Server вернул ошибку.</span><span class="sxs-lookup"><span data-stu-id="f5323-532">The Exchange server returned an error.</span></span> <span data-ttu-id="f5323-533">Для получения дополнительных сведений просмотрите объект диагностики.</span><span class="sxs-lookup"><span data-stu-id="f5323-533">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="f5323-534">Пользователь отключен от сети.</span><span class="sxs-lookup"><span data-stu-id="f5323-534">The user is no longer connected to the network.</span></span> <span data-ttu-id="f5323-535">Проверьте сетевое подключение и повторите попытку.</span><span class="sxs-lookup"><span data-stu-id="f5323-535">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f5323-536">Требования</span><span class="sxs-lookup"><span data-stu-id="f5323-536">Requirements</span></span>

|<span data-ttu-id="f5323-537">Требование</span><span class="sxs-lookup"><span data-stu-id="f5323-537">Requirement</span></span>| <span data-ttu-id="f5323-538">Значение</span><span class="sxs-lookup"><span data-stu-id="f5323-538">Value</span></span>|
|---|---|
|[<span data-ttu-id="f5323-539">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="f5323-539">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f5323-540">1.5</span><span class="sxs-lookup"><span data-stu-id="f5323-540">1.5</span></span> |
|[<span data-ttu-id="f5323-541">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f5323-541">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f5323-542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f5323-542">ReadItem</span></span>|
|[<span data-ttu-id="f5323-543">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f5323-543">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f5323-544">Создание и чтение</span><span class="sxs-lookup"><span data-stu-id="f5323-544">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="f5323-545">Пример</span><span class="sxs-lookup"><span data-stu-id="f5323-545">Example</span></span>

```js
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

<br>

---
---

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="f5323-546">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="f5323-546">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="f5323-547">Получает строку, содержащую маркер, используемый для получения вложения или элемента с Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="f5323-547">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="f5323-p139">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="f5323-p139">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="f5323-550">Вы можете передать сторонней системе маркер и идентификатор вложения или элемента.</span><span class="sxs-lookup"><span data-stu-id="f5323-550">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="f5323-551">Сторонняя система использует этот маркер как маркер авторизации, чтобы вызвать операцию [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) или [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) веб-служб Exchange для возврата вложения или элемента.</span><span class="sxs-lookup"><span data-stu-id="f5323-551">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="f5323-552">Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="f5323-552">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="f5323-553">Для вызова метода `getCallbackTokenAsync` в режиме чтения требуется минимальный уровень разрешения **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="f5323-553">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="f5323-554">Для вызова `getCallbackTokenAsync` в режиме создания сообщения требуется сохранить элемент.</span><span class="sxs-lookup"><span data-stu-id="f5323-554">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="f5323-555">Для метода [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) требуется минимальный уровень разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="f5323-555">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f5323-556">Параметры</span><span class="sxs-lookup"><span data-stu-id="f5323-556">Parameters</span></span>

|<span data-ttu-id="f5323-557">Имя</span><span class="sxs-lookup"><span data-stu-id="f5323-557">Name</span></span>| <span data-ttu-id="f5323-558">Тип</span><span class="sxs-lookup"><span data-stu-id="f5323-558">Type</span></span>| <span data-ttu-id="f5323-559">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="f5323-559">Attributes</span></span>| <span data-ttu-id="f5323-560">Описание</span><span class="sxs-lookup"><span data-stu-id="f5323-560">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="f5323-561">функция</span><span class="sxs-lookup"><span data-stu-id="f5323-561">function</span></span>||<span data-ttu-id="f5323-562">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f5323-562">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f5323-563">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="f5323-563">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="f5323-564">При наличии ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут предоставлять дополнительные сведения.</span><span class="sxs-lookup"><span data-stu-id="f5323-564">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="f5323-565">Объект</span><span class="sxs-lookup"><span data-stu-id="f5323-565">Object</span></span>| <span data-ttu-id="f5323-566">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="f5323-566">&lt;optional&gt;</span></span>|<span data-ttu-id="f5323-567">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="f5323-567">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f5323-568">Ошибки</span><span class="sxs-lookup"><span data-stu-id="f5323-568">Errors</span></span>

|<span data-ttu-id="f5323-569">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="f5323-569">Error code</span></span>|<span data-ttu-id="f5323-570">Описание</span><span class="sxs-lookup"><span data-stu-id="f5323-570">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="f5323-571">Не удалось выполнить запрос.</span><span class="sxs-lookup"><span data-stu-id="f5323-571">The request has failed.</span></span> <span data-ttu-id="f5323-572">Просмотрите объект диагностики для кода ошибки HTTP.</span><span class="sxs-lookup"><span data-stu-id="f5323-572">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="f5323-573">Сервер Exchange Server вернул ошибку.</span><span class="sxs-lookup"><span data-stu-id="f5323-573">The Exchange server returned an error.</span></span> <span data-ttu-id="f5323-574">Для получения дополнительных сведений просмотрите объект диагностики.</span><span class="sxs-lookup"><span data-stu-id="f5323-574">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="f5323-575">Пользователь отключен от сети.</span><span class="sxs-lookup"><span data-stu-id="f5323-575">The user is no longer connected to the network.</span></span> <span data-ttu-id="f5323-576">Проверьте сетевое подключение и повторите попытку.</span><span class="sxs-lookup"><span data-stu-id="f5323-576">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f5323-577">Требования</span><span class="sxs-lookup"><span data-stu-id="f5323-577">Requirements</span></span>

|<span data-ttu-id="f5323-578">Требование</span><span class="sxs-lookup"><span data-stu-id="f5323-578">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="f5323-579">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f5323-579">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f5323-580">1.0</span><span class="sxs-lookup"><span data-stu-id="f5323-580">1.0</span></span> | <span data-ttu-id="f5323-581">1.3</span><span class="sxs-lookup"><span data-stu-id="f5323-581">1.3</span></span> |
|[<span data-ttu-id="f5323-582">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f5323-582">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f5323-583">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f5323-583">ReadItem</span></span> | <span data-ttu-id="f5323-584">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f5323-584">ReadItem</span></span> |
|[<span data-ttu-id="f5323-585">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f5323-585">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f5323-586">Чтение</span><span class="sxs-lookup"><span data-stu-id="f5323-586">Read</span></span> | <span data-ttu-id="f5323-587">Создание</span><span class="sxs-lookup"><span data-stu-id="f5323-587">Compose</span></span> |

##### <a name="example"></a><span data-ttu-id="f5323-588">Пример</span><span class="sxs-lookup"><span data-stu-id="f5323-588">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

<br>

---
---

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="f5323-589">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="f5323-589">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="f5323-590">Получает маркер, идентифицирующий пользователя и надстройку Office.</span><span class="sxs-lookup"><span data-stu-id="f5323-590">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="f5323-591">Метод `getUserIdentityTokenAsync` возвращает маркер, который можно использовать для идентификации, а также [проверки подлинности надстройки и пользователя в сторонней системе](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="f5323-591">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="f5323-592">Параметры</span><span class="sxs-lookup"><span data-stu-id="f5323-592">Parameters</span></span>

|<span data-ttu-id="f5323-593">Имя</span><span class="sxs-lookup"><span data-stu-id="f5323-593">Name</span></span>| <span data-ttu-id="f5323-594">Тип</span><span class="sxs-lookup"><span data-stu-id="f5323-594">Type</span></span>| <span data-ttu-id="f5323-595">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="f5323-595">Attributes</span></span>| <span data-ttu-id="f5323-596">Описание</span><span class="sxs-lookup"><span data-stu-id="f5323-596">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="f5323-597">функция</span><span class="sxs-lookup"><span data-stu-id="f5323-597">function</span></span>||<span data-ttu-id="f5323-598">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f5323-598">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f5323-599">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="f5323-599">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="f5323-600">При наличии ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут предоставлять дополнительные сведения.</span><span class="sxs-lookup"><span data-stu-id="f5323-600">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="f5323-601">Объект</span><span class="sxs-lookup"><span data-stu-id="f5323-601">Object</span></span>| <span data-ttu-id="f5323-602">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="f5323-602">&lt;optional&gt;</span></span>|<span data-ttu-id="f5323-603">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="f5323-603">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="f5323-604">Ошибки</span><span class="sxs-lookup"><span data-stu-id="f5323-604">Errors</span></span>

|<span data-ttu-id="f5323-605">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="f5323-605">Error code</span></span>|<span data-ttu-id="f5323-606">Описание</span><span class="sxs-lookup"><span data-stu-id="f5323-606">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="f5323-607">Не удалось выполнить запрос.</span><span class="sxs-lookup"><span data-stu-id="f5323-607">The request has failed.</span></span> <span data-ttu-id="f5323-608">Просмотрите объект диагностики для кода ошибки HTTP.</span><span class="sxs-lookup"><span data-stu-id="f5323-608">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="f5323-609">Сервер Exchange Server вернул ошибку.</span><span class="sxs-lookup"><span data-stu-id="f5323-609">The Exchange server returned an error.</span></span> <span data-ttu-id="f5323-610">Для получения дополнительных сведений просмотрите объект диагностики.</span><span class="sxs-lookup"><span data-stu-id="f5323-610">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="f5323-611">Пользователь отключен от сети.</span><span class="sxs-lookup"><span data-stu-id="f5323-611">The user is no longer connected to the network.</span></span> <span data-ttu-id="f5323-612">Проверьте сетевое подключение и повторите попытку.</span><span class="sxs-lookup"><span data-stu-id="f5323-612">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f5323-613">Требования</span><span class="sxs-lookup"><span data-stu-id="f5323-613">Requirements</span></span>

|<span data-ttu-id="f5323-614">Требование</span><span class="sxs-lookup"><span data-stu-id="f5323-614">Requirement</span></span>| <span data-ttu-id="f5323-615">Значение</span><span class="sxs-lookup"><span data-stu-id="f5323-615">Value</span></span>|
|---|---|
|[<span data-ttu-id="f5323-616">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f5323-616">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f5323-617">1.0</span><span class="sxs-lookup"><span data-stu-id="f5323-617">1.0</span></span>|
|[<span data-ttu-id="f5323-618">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f5323-618">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f5323-619">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f5323-619">ReadItem</span></span>|
|[<span data-ttu-id="f5323-620">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f5323-620">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f5323-621">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f5323-621">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f5323-622">Пример</span><span class="sxs-lookup"><span data-stu-id="f5323-622">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

<br>

---
---

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="f5323-623">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="f5323-623">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="f5323-624">Выполняет асинхронный запрос для веб-служб Exchange (EWS) на сервере Exchange Server, на котором размещен почтовый ящик пользователя.</span><span class="sxs-lookup"><span data-stu-id="f5323-624">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="f5323-625">Этот метод не поддерживается в следующих сценариях:</span><span class="sxs-lookup"><span data-stu-id="f5323-625">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="f5323-626">В Outlook для iOS и Android</span><span class="sxs-lookup"><span data-stu-id="f5323-626">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="f5323-627">Если надстройка загружается в почтовый ящик Gmail.</span><span class="sxs-lookup"><span data-stu-id="f5323-627">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="f5323-628">В таких случаях надстройка должна [использовать REST API](/outlook/add-ins/use-rest-api) для доступа к почтовому ящику пользователя.</span><span class="sxs-lookup"><span data-stu-id="f5323-628">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="f5323-629">Метод `makeEwsRequestAsync` отправляет запрос EWS от имени надстройки в Exchange.</span><span class="sxs-lookup"><span data-stu-id="f5323-629">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="f5323-630">Список поддерживаемых операций EWS см. в статье [Вызов веб-служб из надстройки Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support).</span><span class="sxs-lookup"><span data-stu-id="f5323-630">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="f5323-631">С помощью метода `makeEwsRequestAsync` невозможно запрашивать элементы, связанные с папкой.</span><span class="sxs-lookup"><span data-stu-id="f5323-631">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="f5323-632">В запросе XML должна быть указана кодировка UTF-8.</span><span class="sxs-lookup"><span data-stu-id="f5323-632">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="f5323-p149">У вашей надстройки должно быть разрешение **ReadWriteMailbox** для использования метода `makeEwsRequestAsync`. Сведения об использовании разрешения **ReadWriteMailbox** и операций EWS, которые можно вызывать с помощью метода `makeEwsRequestAsync`, см. в статье [Указание разрешений для доступа почтовой надстройки к почтовому ящику пользователя](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="f5323-p149">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="f5323-635">Администратор сервера должен установить значение true для параметра `OAuthAuthentication` в каталоге сервера клиентского доступа EWS, чтобы метод `makeEwsRequestAsync` мог выполнять запросы EWS.</span><span class="sxs-lookup"><span data-stu-id="f5323-635">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="f5323-636">Различия версий</span><span class="sxs-lookup"><span data-stu-id="f5323-636">Version differences</span></span>

<span data-ttu-id="f5323-637">Если вы используете метод `makeEwsRequestAsync` в почтовых приложениях, которые выполняются в Outlook версии более ранней, чем 15.0.4535.1004, указывайте кодировку `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="f5323-637">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="f5323-638">Значение кодировки не нужно указывать, если почтовое приложение выполняется в Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="f5323-638">You do not need to set the encoding value when your mail app is running in Outlook on the web.</span></span> <span data-ttu-id="f5323-639">Вы можете определить, работает ли почтовое приложение в Outlook в Интернете или на настольном клиенте с помощью свойства Mailbox. Diagnostics. hostName.</span><span class="sxs-lookup"><span data-stu-id="f5323-639">You can determine whether your mail app is running in Outlook on the web or a desktop client by using the mailbox.diagnostics.hostName property.</span></span> <span data-ttu-id="f5323-640">Используемую версию Outlook можно определить с помощью свойства mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="f5323-640">You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f5323-641">Параметры</span><span class="sxs-lookup"><span data-stu-id="f5323-641">Parameters</span></span>

|<span data-ttu-id="f5323-642">Имя</span><span class="sxs-lookup"><span data-stu-id="f5323-642">Name</span></span>| <span data-ttu-id="f5323-643">Тип</span><span class="sxs-lookup"><span data-stu-id="f5323-643">Type</span></span>| <span data-ttu-id="f5323-644">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="f5323-644">Attributes</span></span>| <span data-ttu-id="f5323-645">Описание</span><span class="sxs-lookup"><span data-stu-id="f5323-645">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="f5323-646">String</span><span class="sxs-lookup"><span data-stu-id="f5323-646">String</span></span>||<span data-ttu-id="f5323-647">Запрос EWS.</span><span class="sxs-lookup"><span data-stu-id="f5323-647">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="f5323-648">function</span><span class="sxs-lookup"><span data-stu-id="f5323-648">function</span></span>||<span data-ttu-id="f5323-649">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f5323-649">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="f5323-650">Результат XML вызова EWS указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="f5323-650">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="f5323-651">Если размер результата превышает 1 МБ, возвращается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="f5323-651">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="f5323-652">Объект</span><span class="sxs-lookup"><span data-stu-id="f5323-652">Object</span></span>| <span data-ttu-id="f5323-653">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="f5323-653">&lt;optional&gt;</span></span>|<span data-ttu-id="f5323-654">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="f5323-654">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f5323-655">Requirements</span><span class="sxs-lookup"><span data-stu-id="f5323-655">Requirements</span></span>

|<span data-ttu-id="f5323-656">Требование</span><span class="sxs-lookup"><span data-stu-id="f5323-656">Requirement</span></span>| <span data-ttu-id="f5323-657">Значение</span><span class="sxs-lookup"><span data-stu-id="f5323-657">Value</span></span>|
|---|---|
|[<span data-ttu-id="f5323-658">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f5323-658">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f5323-659">1.0</span><span class="sxs-lookup"><span data-stu-id="f5323-659">1.0</span></span>|
|[<span data-ttu-id="f5323-660">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f5323-660">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f5323-661">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="f5323-661">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="f5323-662">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f5323-662">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f5323-663">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f5323-663">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f5323-664">Пример</span><span class="sxs-lookup"><span data-stu-id="f5323-664">Example</span></span>

<span data-ttu-id="f5323-665">В приведенном ниже примере вызывается `makeEwsRequestAsync` для получения темы элемента с помощью операции `GetItem`.</span><span class="sxs-lookup"><span data-stu-id="f5323-665">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

```js
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

<br>

---
---

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="f5323-666">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="f5323-666">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="f5323-667">Удаляет обработчиков для поддерживаемого типа события.</span><span class="sxs-lookup"><span data-stu-id="f5323-667">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="f5323-668">В настоящее время поддерживаются типы событий `Office.EventType.ItemChanged` и `Office.EventType.OfficeThemeChanged`.</span><span class="sxs-lookup"><span data-stu-id="f5323-668">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="f5323-669">Параметры</span><span class="sxs-lookup"><span data-stu-id="f5323-669">Parameters</span></span>

| <span data-ttu-id="f5323-670">Имя</span><span class="sxs-lookup"><span data-stu-id="f5323-670">Name</span></span> | <span data-ttu-id="f5323-671">Тип</span><span class="sxs-lookup"><span data-stu-id="f5323-671">Type</span></span> | <span data-ttu-id="f5323-672">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="f5323-672">Attributes</span></span> | <span data-ttu-id="f5323-673">Описание</span><span class="sxs-lookup"><span data-stu-id="f5323-673">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="f5323-674">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="f5323-674">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="f5323-675">Событие, которое должно отменить обработчик.</span><span class="sxs-lookup"><span data-stu-id="f5323-675">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="f5323-676">Объект</span><span class="sxs-lookup"><span data-stu-id="f5323-676">Object</span></span> | <span data-ttu-id="f5323-677">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="f5323-677">&lt;optional&gt;</span></span> | <span data-ttu-id="f5323-678">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="f5323-678">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="f5323-679">Object</span><span class="sxs-lookup"><span data-stu-id="f5323-679">Object</span></span> | <span data-ttu-id="f5323-680">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="f5323-680">&lt;optional&gt;</span></span> | <span data-ttu-id="f5323-681">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="f5323-681">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="f5323-682">функция</span><span class="sxs-lookup"><span data-stu-id="f5323-682">function</span></span>| <span data-ttu-id="f5323-683">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="f5323-683">&lt;optional&gt;</span></span>|<span data-ttu-id="f5323-684">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="f5323-684">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f5323-685">Requirements</span><span class="sxs-lookup"><span data-stu-id="f5323-685">Requirements</span></span>

|<span data-ttu-id="f5323-686">Требование</span><span class="sxs-lookup"><span data-stu-id="f5323-686">Requirement</span></span>| <span data-ttu-id="f5323-687">Значение</span><span class="sxs-lookup"><span data-stu-id="f5323-687">Value</span></span>|
|---|---|
|[<span data-ttu-id="f5323-688">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="f5323-688">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f5323-689">1.5</span><span class="sxs-lookup"><span data-stu-id="f5323-689">1.5</span></span> |
|[<span data-ttu-id="f5323-690">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f5323-690">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f5323-691">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f5323-691">ReadItem</span></span> |
|[<span data-ttu-id="f5323-692">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f5323-692">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f5323-693">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f5323-693">Compose or Read</span></span>|
