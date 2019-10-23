---
title: Office. Context. Mailbox — Предварительная версия набора обязательных элементов
description: ''
ms.date: 10/21/2019
localization_priority: Normal
ms.openlocfilehash: 29922c9e05cc0380f1e54a16f3350c578d9e4cee
ms.sourcegitcommit: 499bf49b41205f8034c501d4db5fe4b02dab205e
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/22/2019
ms.locfileid: "37627071"
---
# <a name="mailbox"></a><span data-ttu-id="d5c6e-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="d5c6e-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="d5c6e-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="d5c6e-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="d5c6e-104">Предоставляет для Microsoft Outlook доступ к объектной модели надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d5c6e-105">Требования</span><span class="sxs-lookup"><span data-stu-id="d5c6e-105">Requirements</span></span>

|<span data-ttu-id="d5c6e-106">Требование</span><span class="sxs-lookup"><span data-stu-id="d5c6e-106">Requirement</span></span>| <span data-ttu-id="d5c6e-107">Значение</span><span class="sxs-lookup"><span data-stu-id="d5c6e-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="d5c6e-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d5c6e-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d5c6e-109">1.0</span><span class="sxs-lookup"><span data-stu-id="d5c6e-109">1.0</span></span>|
|[<span data-ttu-id="d5c6e-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d5c6e-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d5c6e-111">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="d5c6e-111">Restricted</span></span>|
|[<span data-ttu-id="d5c6e-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d5c6e-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d5c6e-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d5c6e-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="d5c6e-114">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="d5c6e-114">Members and methods</span></span>

| <span data-ttu-id="d5c6e-115">Элемент</span><span class="sxs-lookup"><span data-stu-id="d5c6e-115">Member</span></span> | <span data-ttu-id="d5c6e-116">Тип</span><span class="sxs-lookup"><span data-stu-id="d5c6e-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="d5c6e-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="d5c6e-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="d5c6e-118">Элемент</span><span class="sxs-lookup"><span data-stu-id="d5c6e-118">Member</span></span> |
| [<span data-ttu-id="d5c6e-119">мастеркатегориес</span><span class="sxs-lookup"><span data-stu-id="d5c6e-119">masterCategories</span></span>](#mastercategories-mastercategories) | <span data-ttu-id="d5c6e-120">Элемент</span><span class="sxs-lookup"><span data-stu-id="d5c6e-120">Member</span></span> |
| [<span data-ttu-id="d5c6e-121">restUrl</span><span class="sxs-lookup"><span data-stu-id="d5c6e-121">restUrl</span></span>](#resturl-string) | <span data-ttu-id="d5c6e-122">Элемент</span><span class="sxs-lookup"><span data-stu-id="d5c6e-122">Member</span></span> |
| [<span data-ttu-id="d5c6e-123">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="d5c6e-123">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="d5c6e-124">Метод</span><span class="sxs-lookup"><span data-stu-id="d5c6e-124">Method</span></span> |
| [<span data-ttu-id="d5c6e-125">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="d5c6e-125">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="d5c6e-126">Метод</span><span class="sxs-lookup"><span data-stu-id="d5c6e-126">Method</span></span> |
| [<span data-ttu-id="d5c6e-127">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="d5c6e-127">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="d5c6e-128">Метод</span><span class="sxs-lookup"><span data-stu-id="d5c6e-128">Method</span></span> |
| [<span data-ttu-id="d5c6e-129">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="d5c6e-129">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="d5c6e-130">Метод</span><span class="sxs-lookup"><span data-stu-id="d5c6e-130">Method</span></span> |
| [<span data-ttu-id="d5c6e-131">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="d5c6e-131">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="d5c6e-132">Метод</span><span class="sxs-lookup"><span data-stu-id="d5c6e-132">Method</span></span> |
| [<span data-ttu-id="d5c6e-133">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="d5c6e-133">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="d5c6e-134">Метод</span><span class="sxs-lookup"><span data-stu-id="d5c6e-134">Method</span></span> |
| [<span data-ttu-id="d5c6e-135">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="d5c6e-135">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="d5c6e-136">Метод</span><span class="sxs-lookup"><span data-stu-id="d5c6e-136">Method</span></span> |
| [<span data-ttu-id="d5c6e-137">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="d5c6e-137">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="d5c6e-138">Метод</span><span class="sxs-lookup"><span data-stu-id="d5c6e-138">Method</span></span> |
| [<span data-ttu-id="d5c6e-139">дисплайневмессажеформ</span><span class="sxs-lookup"><span data-stu-id="d5c6e-139">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="d5c6e-140">Метод</span><span class="sxs-lookup"><span data-stu-id="d5c6e-140">Method</span></span> |
| [<span data-ttu-id="d5c6e-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="d5c6e-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="d5c6e-142">Метод</span><span class="sxs-lookup"><span data-stu-id="d5c6e-142">Method</span></span> |
| [<span data-ttu-id="d5c6e-143">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="d5c6e-143">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="d5c6e-144">Метод</span><span class="sxs-lookup"><span data-stu-id="d5c6e-144">Method</span></span> |
| [<span data-ttu-id="d5c6e-145">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="d5c6e-145">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="d5c6e-146">Метод</span><span class="sxs-lookup"><span data-stu-id="d5c6e-146">Method</span></span> |
| [<span data-ttu-id="d5c6e-147">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="d5c6e-147">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="d5c6e-148">Метод</span><span class="sxs-lookup"><span data-stu-id="d5c6e-148">Method</span></span> |
| [<span data-ttu-id="d5c6e-149">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="d5c6e-149">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="d5c6e-150">Метод</span><span class="sxs-lookup"><span data-stu-id="d5c6e-150">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="d5c6e-151">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="d5c6e-151">Namespaces</span></span>

<span data-ttu-id="d5c6e-152">[diagnostics](Office.context.mailbox.diagnostics.md). Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-152">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="d5c6e-153">[item](Office.context.mailbox.item.md). Предоставляет методы и свойства для доступа к сообщению или встрече в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-153">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="d5c6e-154">[userProfile](Office.context.mailbox.userProfile.md). Предоставляет сведения о пользователе в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-154">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="d5c6e-155">Members</span><span class="sxs-lookup"><span data-stu-id="d5c6e-155">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="d5c6e-156">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="d5c6e-156">ewsUrl: String</span></span>

<span data-ttu-id="d5c6e-p101">Получает URL-адрес конечной точки веб-служб Exchange (EWS) для этой учетной записи электронной почты. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d5c6e-159">Этот элемент не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-159">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d5c6e-p102">Удаленная служба может использовать значение `ewsUrl`, чтобы выполнять вызовы EWS для почтового ящика пользователя. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="d5c6e-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="d5c6e-162">Чтобы вызвать элемент `ewsUrl` в режиме чтения, в манифесте приложения должно быть указано разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-162">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="d5c6e-p103">Перед использованием элемента `ewsUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="d5c6e-165">Тип</span><span class="sxs-lookup"><span data-stu-id="d5c6e-165">Type</span></span>

*   <span data-ttu-id="d5c6e-166">String</span><span class="sxs-lookup"><span data-stu-id="d5c6e-166">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d5c6e-167">Требования</span><span class="sxs-lookup"><span data-stu-id="d5c6e-167">Requirements</span></span>

|<span data-ttu-id="d5c6e-168">Требование</span><span class="sxs-lookup"><span data-stu-id="d5c6e-168">Requirement</span></span>| <span data-ttu-id="d5c6e-169">Значение</span><span class="sxs-lookup"><span data-stu-id="d5c6e-169">Value</span></span>|
|---|---|
|[<span data-ttu-id="d5c6e-170">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d5c6e-170">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d5c6e-171">1.0</span><span class="sxs-lookup"><span data-stu-id="d5c6e-171">1.0</span></span>|
|[<span data-ttu-id="d5c6e-172">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d5c6e-172">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d5c6e-173">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d5c6e-173">ReadItem</span></span>|
|[<span data-ttu-id="d5c6e-174">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d5c6e-174">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d5c6e-175">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d5c6e-175">Compose or Read</span></span>|

<br>

---
---

#### <a name="mastercategories-mastercategoriesjavascriptapioutlookofficemastercategories"></a><span data-ttu-id="d5c6e-176">Мастеркатегориес: [мастеркатегориес](/javascript/api/outlook/office.mastercategories)</span><span class="sxs-lookup"><span data-stu-id="d5c6e-176">masterCategories: [MasterCategories](/javascript/api/outlook/office.mastercategories)</span></span>

<span data-ttu-id="d5c6e-177">Получает объект, предоставляющий методы для управления главным списком категорий в этом почтовом ящике.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-177">Gets an object that provides methods to manage the categories master list on this mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="d5c6e-178">Этот элемент не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-178">This member is not supported in Outlook on iOS or Android.</span></span>

##### <a name="type"></a><span data-ttu-id="d5c6e-179">Тип</span><span class="sxs-lookup"><span data-stu-id="d5c6e-179">Type</span></span>

*   [<span data-ttu-id="d5c6e-180">MasterCategories</span><span class="sxs-lookup"><span data-stu-id="d5c6e-180">MasterCategories</span></span>](/javascript/api/outlook/office.mastercategories)

##### <a name="requirements"></a><span data-ttu-id="d5c6e-181">Требования</span><span class="sxs-lookup"><span data-stu-id="d5c6e-181">Requirements</span></span>

|<span data-ttu-id="d5c6e-182">Требование</span><span class="sxs-lookup"><span data-stu-id="d5c6e-182">Requirement</span></span>| <span data-ttu-id="d5c6e-183">Значение</span><span class="sxs-lookup"><span data-stu-id="d5c6e-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="d5c6e-184">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="d5c6e-184">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d5c6e-185">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="d5c6e-185">Preview</span></span> |
|[<span data-ttu-id="d5c6e-186">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d5c6e-186">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d5c6e-187">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="d5c6e-187">ReadWriteMailbox</span></span> |
|[<span data-ttu-id="d5c6e-188">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d5c6e-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d5c6e-189">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d5c6e-189">Compose or Read</span></span> |

##### <a name="example"></a><span data-ttu-id="d5c6e-190">Пример</span><span class="sxs-lookup"><span data-stu-id="d5c6e-190">Example</span></span>

<span data-ttu-id="d5c6e-191">В этом примере показано получение сводного списка категорий для этого почтового ящика.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-191">This example gets the categories master list for this mailbox.</span></span>

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

#### <a name="resturl-string"></a><span data-ttu-id="d5c6e-192">restUrl: String</span><span class="sxs-lookup"><span data-stu-id="d5c6e-192">restUrl: String</span></span>

<span data-ttu-id="d5c6e-193">Возвращает URL-адрес конечной точки REST для этой учетной записи электронной почты.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-193">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="d5c6e-194">С помощью значения `restUrl` можно выполнять вызовы [REST API](/outlook/rest/) для почтового ящика пользователя.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-194">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="d5c6e-195">Чтобы вызвать элемент `restUrl` в режиме чтения, в манифесте приложения необходимо указать разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-195">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="d5c6e-p104">Перед использованием элемента `restUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="d5c6e-198">Тип</span><span class="sxs-lookup"><span data-stu-id="d5c6e-198">Type</span></span>

*   <span data-ttu-id="d5c6e-199">String</span><span class="sxs-lookup"><span data-stu-id="d5c6e-199">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d5c6e-200">Требования</span><span class="sxs-lookup"><span data-stu-id="d5c6e-200">Requirements</span></span>

|<span data-ttu-id="d5c6e-201">Требование</span><span class="sxs-lookup"><span data-stu-id="d5c6e-201">Requirement</span></span>| <span data-ttu-id="d5c6e-202">Значение</span><span class="sxs-lookup"><span data-stu-id="d5c6e-202">Value</span></span>|
|---|---|
|[<span data-ttu-id="d5c6e-203">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="d5c6e-203">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d5c6e-204">1.5</span><span class="sxs-lookup"><span data-stu-id="d5c6e-204">1.5</span></span> |
|[<span data-ttu-id="d5c6e-205">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d5c6e-205">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d5c6e-206">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d5c6e-206">ReadItem</span></span>|
|[<span data-ttu-id="d5c6e-207">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d5c6e-207">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d5c6e-208">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d5c6e-208">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="d5c6e-209">Методы</span><span class="sxs-lookup"><span data-stu-id="d5c6e-209">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="d5c6e-210">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d5c6e-210">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="d5c6e-211">Добавляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-211">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="d5c6e-212">В настоящее время поддерживаются типы событий `Office.EventType.ItemChanged` и `Office.EventType.OfficeThemeChanged`.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-212">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d5c6e-213">Параметры</span><span class="sxs-lookup"><span data-stu-id="d5c6e-213">Parameters</span></span>

| <span data-ttu-id="d5c6e-214">Имя</span><span class="sxs-lookup"><span data-stu-id="d5c6e-214">Name</span></span> | <span data-ttu-id="d5c6e-215">Тип</span><span class="sxs-lookup"><span data-stu-id="d5c6e-215">Type</span></span> | <span data-ttu-id="d5c6e-216">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="d5c6e-216">Attributes</span></span> | <span data-ttu-id="d5c6e-217">Описание</span><span class="sxs-lookup"><span data-stu-id="d5c6e-217">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="d5c6e-218">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="d5c6e-218">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="d5c6e-219">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-219">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="d5c6e-220">Function</span><span class="sxs-lookup"><span data-stu-id="d5c6e-220">Function</span></span> || <span data-ttu-id="d5c6e-p105">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-p105">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="d5c6e-224">Объект</span><span class="sxs-lookup"><span data-stu-id="d5c6e-224">Object</span></span> | <span data-ttu-id="d5c6e-225">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d5c6e-225">&lt;optional&gt;</span></span> | <span data-ttu-id="d5c6e-226">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-226">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="d5c6e-227">Object</span><span class="sxs-lookup"><span data-stu-id="d5c6e-227">Object</span></span> | <span data-ttu-id="d5c6e-228">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d5c6e-228">&lt;optional&gt;</span></span> | <span data-ttu-id="d5c6e-229">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-229">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="d5c6e-230">функция</span><span class="sxs-lookup"><span data-stu-id="d5c6e-230">function</span></span>| <span data-ttu-id="d5c6e-231">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d5c6e-231">&lt;optional&gt;</span></span>|<span data-ttu-id="d5c6e-232">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d5c6e-232">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d5c6e-233">Требования</span><span class="sxs-lookup"><span data-stu-id="d5c6e-233">Requirements</span></span>

|<span data-ttu-id="d5c6e-234">Требование</span><span class="sxs-lookup"><span data-stu-id="d5c6e-234">Requirement</span></span>| <span data-ttu-id="d5c6e-235">Значение</span><span class="sxs-lookup"><span data-stu-id="d5c6e-235">Value</span></span>|
|---|---|
|[<span data-ttu-id="d5c6e-236">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="d5c6e-236">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d5c6e-237">1.5</span><span class="sxs-lookup"><span data-stu-id="d5c6e-237">1.5</span></span> |
|[<span data-ttu-id="d5c6e-238">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d5c6e-238">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d5c6e-239">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d5c6e-239">ReadItem</span></span> |
|[<span data-ttu-id="d5c6e-240">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d5c6e-240">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d5c6e-241">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d5c6e-241">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d5c6e-242">Пример</span><span class="sxs-lookup"><span data-stu-id="d5c6e-242">Example</span></span>

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

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="d5c6e-243">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="d5c6e-243">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="d5c6e-244">Преобразовывает идентификатор элемента из формата REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-244">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="d5c6e-245">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-245">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d5c6e-p106">Формат идентификаторов, извлекаемых через API REST (например, [API Почты Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)), отличается от формата веб-служб Exchange (EWS). Метод `convertToEwsId` преобразовывает идентификатор в формате REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-p106">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d5c6e-248">Параметры</span><span class="sxs-lookup"><span data-stu-id="d5c6e-248">Parameters</span></span>

|<span data-ttu-id="d5c6e-249">Имя</span><span class="sxs-lookup"><span data-stu-id="d5c6e-249">Name</span></span>| <span data-ttu-id="d5c6e-250">Тип</span><span class="sxs-lookup"><span data-stu-id="d5c6e-250">Type</span></span>| <span data-ttu-id="d5c6e-251">Описание</span><span class="sxs-lookup"><span data-stu-id="d5c6e-251">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="d5c6e-252">String</span><span class="sxs-lookup"><span data-stu-id="d5c6e-252">String</span></span>|<span data-ttu-id="d5c6e-253">Идентификатор элемента в формате REST API для Outlook</span><span class="sxs-lookup"><span data-stu-id="d5c6e-253">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="d5c6e-254">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="d5c6e-254">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="d5c6e-255">Значение, определяющее версию REST API для Outlook, которая используется для извлечения идентификатора элемента.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-255">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d5c6e-256">Требования</span><span class="sxs-lookup"><span data-stu-id="d5c6e-256">Requirements</span></span>

|<span data-ttu-id="d5c6e-257">Требование</span><span class="sxs-lookup"><span data-stu-id="d5c6e-257">Requirement</span></span>| <span data-ttu-id="d5c6e-258">Значение</span><span class="sxs-lookup"><span data-stu-id="d5c6e-258">Value</span></span>|
|---|---|
|[<span data-ttu-id="d5c6e-259">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="d5c6e-259">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d5c6e-260">1.3</span><span class="sxs-lookup"><span data-stu-id="d5c6e-260">1.3</span></span>|
|[<span data-ttu-id="d5c6e-261">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d5c6e-261">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d5c6e-262">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="d5c6e-262">Restricted</span></span>|
|[<span data-ttu-id="d5c6e-263">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d5c6e-263">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d5c6e-264">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d5c6e-264">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d5c6e-265">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="d5c6e-265">Returns:</span></span>

<span data-ttu-id="d5c6e-266">Тип: String</span><span class="sxs-lookup"><span data-stu-id="d5c6e-266">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="d5c6e-267">Пример</span><span class="sxs-lookup"><span data-stu-id="d5c6e-267">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime"></a><span data-ttu-id="d5c6e-268">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="d5c6e-268">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span></span>

<span data-ttu-id="d5c6e-269">Получает словарь, содержащий сведения о локальном времени клиента.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-269">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="d5c6e-p107">Почтовое приложение для классической версии Outlook или версии в Интернете может использовать разные часовые пояса для дат и времени. Классическое приложение Outlook использует часовой пояс клиентского компьютера. Outlook в Интернете использует часовой пояс, заданный в Центре администрирования Exchange (EAC). Значения даты и времени должны обрабатываться так, чтобы значения в пользовательском интерфейсе всегда согласовывались с часовым поясом, ожидаемым пользователем.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-p107">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="d5c6e-p108">Если почтовое приложение работает в классическом клиенте Outlook, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса клиентского компьютера. Если почтовое приложение работает в Outlook в Интернете, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса, заданного в Центре администрирования Exchange.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-p108">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d5c6e-275">Параметры</span><span class="sxs-lookup"><span data-stu-id="d5c6e-275">Parameters</span></span>

|<span data-ttu-id="d5c6e-276">Имя</span><span class="sxs-lookup"><span data-stu-id="d5c6e-276">Name</span></span>| <span data-ttu-id="d5c6e-277">Тип</span><span class="sxs-lookup"><span data-stu-id="d5c6e-277">Type</span></span>| <span data-ttu-id="d5c6e-278">Описание</span><span class="sxs-lookup"><span data-stu-id="d5c6e-278">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="d5c6e-279">Date</span><span class="sxs-lookup"><span data-stu-id="d5c6e-279">Date</span></span>|<span data-ttu-id="d5c6e-280">Объект Date</span><span class="sxs-lookup"><span data-stu-id="d5c6e-280">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d5c6e-281">Требования</span><span class="sxs-lookup"><span data-stu-id="d5c6e-281">Requirements</span></span>

|<span data-ttu-id="d5c6e-282">Требование</span><span class="sxs-lookup"><span data-stu-id="d5c6e-282">Requirement</span></span>| <span data-ttu-id="d5c6e-283">Значение</span><span class="sxs-lookup"><span data-stu-id="d5c6e-283">Value</span></span>|
|---|---|
|[<span data-ttu-id="d5c6e-284">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d5c6e-284">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d5c6e-285">1.0</span><span class="sxs-lookup"><span data-stu-id="d5c6e-285">1.0</span></span>|
|[<span data-ttu-id="d5c6e-286">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d5c6e-286">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d5c6e-287">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d5c6e-287">ReadItem</span></span>|
|[<span data-ttu-id="d5c6e-288">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d5c6e-288">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d5c6e-289">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d5c6e-289">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d5c6e-290">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="d5c6e-290">Returns:</span></span>

<span data-ttu-id="d5c6e-291">Тип: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="d5c6e-291">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="d5c6e-292">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="d5c6e-292">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="d5c6e-293">Преобразовывает идентификатор элемента в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-293">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="d5c6e-294">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-294">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d5c6e-p109">Формат идентификаторов, извлекаемых через EWS или свойство `itemId`, отличается от формата API REST (таких как [API Почты Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)). Метод `convertToRestId` преобразовывает идентификатор в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-p109">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d5c6e-297">Параметры</span><span class="sxs-lookup"><span data-stu-id="d5c6e-297">Parameters</span></span>

|<span data-ttu-id="d5c6e-298">Имя</span><span class="sxs-lookup"><span data-stu-id="d5c6e-298">Name</span></span>| <span data-ttu-id="d5c6e-299">Тип</span><span class="sxs-lookup"><span data-stu-id="d5c6e-299">Type</span></span>| <span data-ttu-id="d5c6e-300">Описание</span><span class="sxs-lookup"><span data-stu-id="d5c6e-300">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="d5c6e-301">String</span><span class="sxs-lookup"><span data-stu-id="d5c6e-301">String</span></span>|<span data-ttu-id="d5c6e-302">Идентификатор элемента в формате EWS</span><span class="sxs-lookup"><span data-stu-id="d5c6e-302">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="d5c6e-303">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="d5c6e-303">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="d5c6e-304">Значение, определяющее версию REST API для Outlook, с которой будет использоваться преобразованный идентификатор.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-304">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d5c6e-305">Требования</span><span class="sxs-lookup"><span data-stu-id="d5c6e-305">Requirements</span></span>

|<span data-ttu-id="d5c6e-306">Требование</span><span class="sxs-lookup"><span data-stu-id="d5c6e-306">Requirement</span></span>| <span data-ttu-id="d5c6e-307">Значение</span><span class="sxs-lookup"><span data-stu-id="d5c6e-307">Value</span></span>|
|---|---|
|[<span data-ttu-id="d5c6e-308">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="d5c6e-308">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d5c6e-309">1.3</span><span class="sxs-lookup"><span data-stu-id="d5c6e-309">1.3</span></span>|
|[<span data-ttu-id="d5c6e-310">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d5c6e-310">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d5c6e-311">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="d5c6e-311">Restricted</span></span>|
|[<span data-ttu-id="d5c6e-312">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d5c6e-312">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d5c6e-313">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d5c6e-313">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d5c6e-314">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="d5c6e-314">Returns:</span></span>

<span data-ttu-id="d5c6e-315">Тип: String</span><span class="sxs-lookup"><span data-stu-id="d5c6e-315">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="d5c6e-316">Пример</span><span class="sxs-lookup"><span data-stu-id="d5c6e-316">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="d5c6e-317">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="d5c6e-317">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="d5c6e-318">Получает объект Date из словаря, содержащего сведения о времени.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-318">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="d5c6e-319">Метод `convertToUtcClientTime` преобразует словарь, содержащий локальную дату и время, в объект Date с правильными значениями локальной даты и времени.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-319">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d5c6e-320">Параметры</span><span class="sxs-lookup"><span data-stu-id="d5c6e-320">Parameters</span></span>

|<span data-ttu-id="d5c6e-321">Имя</span><span class="sxs-lookup"><span data-stu-id="d5c6e-321">Name</span></span>| <span data-ttu-id="d5c6e-322">Тип</span><span class="sxs-lookup"><span data-stu-id="d5c6e-322">Type</span></span>| <span data-ttu-id="d5c6e-323">Описание</span><span class="sxs-lookup"><span data-stu-id="d5c6e-323">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="d5c6e-324">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="d5c6e-324">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime)|<span data-ttu-id="d5c6e-325">Значение локального времени для преобразования.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-325">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d5c6e-326">Требования</span><span class="sxs-lookup"><span data-stu-id="d5c6e-326">Requirements</span></span>

|<span data-ttu-id="d5c6e-327">Требование</span><span class="sxs-lookup"><span data-stu-id="d5c6e-327">Requirement</span></span>| <span data-ttu-id="d5c6e-328">Значение</span><span class="sxs-lookup"><span data-stu-id="d5c6e-328">Value</span></span>|
|---|---|
|[<span data-ttu-id="d5c6e-329">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d5c6e-329">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d5c6e-330">1.0</span><span class="sxs-lookup"><span data-stu-id="d5c6e-330">1.0</span></span>|
|[<span data-ttu-id="d5c6e-331">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d5c6e-331">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d5c6e-332">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d5c6e-332">ReadItem</span></span>|
|[<span data-ttu-id="d5c6e-333">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d5c6e-333">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d5c6e-334">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d5c6e-334">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d5c6e-335">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="d5c6e-335">Returns:</span></span>

<span data-ttu-id="d5c6e-336">Объект Date со временем в формате UTC.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-336">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="d5c6e-337">Тип: Date</span><span class="sxs-lookup"><span data-stu-id="d5c6e-337">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="d5c6e-338">Пример</span><span class="sxs-lookup"><span data-stu-id="d5c6e-338">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="d5c6e-339">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="d5c6e-339">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="d5c6e-340">Отображает имеющуюся встречу из календаря.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-340">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="d5c6e-341">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-341">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d5c6e-342">Метод `displayAppointmentForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее сведения календаря о существующей встрече.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-342">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="d5c6e-p110">В Outlook для Mac с помощью этого метода можно отобразить одну встречу, которая не является частью повторяющегося ряда, или основную встречу такого ряда, но не экземпляр из него, так как в Outlook для Mac невозможно получить доступ к свойствам экземпляра повторяющегося ряда (в том числе к идентификатору элемента).</span><span class="sxs-lookup"><span data-stu-id="d5c6e-p110">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="d5c6e-345">В Outlook в Интернете этот метод открывает указанную форму, только если текст формы содержит символы размером не более 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-345">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="d5c6e-346">Если указанный идентификатор элемента не определяет существующую встречу, на клиентском компьютере или устройстве открывается пустая страница, и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-346">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d5c6e-347">Параметры</span><span class="sxs-lookup"><span data-stu-id="d5c6e-347">Parameters</span></span>

|<span data-ttu-id="d5c6e-348">Имя</span><span class="sxs-lookup"><span data-stu-id="d5c6e-348">Name</span></span>| <span data-ttu-id="d5c6e-349">Тип</span><span class="sxs-lookup"><span data-stu-id="d5c6e-349">Type</span></span>| <span data-ttu-id="d5c6e-350">Описание</span><span class="sxs-lookup"><span data-stu-id="d5c6e-350">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="d5c6e-351">String</span><span class="sxs-lookup"><span data-stu-id="d5c6e-351">String</span></span>|<span data-ttu-id="d5c6e-352">Идентификатор веб-служб Exchange для существующей встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-352">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d5c6e-353">Требования</span><span class="sxs-lookup"><span data-stu-id="d5c6e-353">Requirements</span></span>

|<span data-ttu-id="d5c6e-354">Требование</span><span class="sxs-lookup"><span data-stu-id="d5c6e-354">Requirement</span></span>| <span data-ttu-id="d5c6e-355">Значение</span><span class="sxs-lookup"><span data-stu-id="d5c6e-355">Value</span></span>|
|---|---|
|[<span data-ttu-id="d5c6e-356">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d5c6e-356">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d5c6e-357">1.0</span><span class="sxs-lookup"><span data-stu-id="d5c6e-357">1.0</span></span>|
|[<span data-ttu-id="d5c6e-358">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d5c6e-358">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d5c6e-359">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d5c6e-359">ReadItem</span></span>|
|[<span data-ttu-id="d5c6e-360">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d5c6e-360">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d5c6e-361">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d5c6e-361">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d5c6e-362">Пример</span><span class="sxs-lookup"><span data-stu-id="d5c6e-362">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="d5c6e-363">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="d5c6e-363">displayMessageForm(itemId)</span></span>

<span data-ttu-id="d5c6e-364">Отображает имеющееся сообщение.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-364">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="d5c6e-365">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-365">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d5c6e-366">Метод `displayMessageForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее существующее сообщение.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-366">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="d5c6e-367">В Outlook в Интернете этот метод открывает указанную форму, только если текст формы содержит символы размером не более 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-367">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="d5c6e-368">Если указанный идентификатор элемента не определяет существующее сообщение, окно на клиентском компьютере не открывается и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-368">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="d5c6e-p111">Не используйте `displayMessageForm` с параметром `itemId`, который представляет собой встречу. Используйте метод `displayAppointmentForm`, чтобы отобразить сведения о существующей встрече, а метод `displayNewAppointmentForm` — для отображения формы создания встречи.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-p111">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d5c6e-371">Параметры</span><span class="sxs-lookup"><span data-stu-id="d5c6e-371">Parameters</span></span>

|<span data-ttu-id="d5c6e-372">Имя</span><span class="sxs-lookup"><span data-stu-id="d5c6e-372">Name</span></span>| <span data-ttu-id="d5c6e-373">Тип</span><span class="sxs-lookup"><span data-stu-id="d5c6e-373">Type</span></span>| <span data-ttu-id="d5c6e-374">Описание</span><span class="sxs-lookup"><span data-stu-id="d5c6e-374">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="d5c6e-375">String</span><span class="sxs-lookup"><span data-stu-id="d5c6e-375">String</span></span>|<span data-ttu-id="d5c6e-376">Идентификатор веб-служб Exchange для существующего сообщения.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-376">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d5c6e-377">Требования</span><span class="sxs-lookup"><span data-stu-id="d5c6e-377">Requirements</span></span>

|<span data-ttu-id="d5c6e-378">Требование</span><span class="sxs-lookup"><span data-stu-id="d5c6e-378">Requirement</span></span>| <span data-ttu-id="d5c6e-379">Значение</span><span class="sxs-lookup"><span data-stu-id="d5c6e-379">Value</span></span>|
|---|---|
|[<span data-ttu-id="d5c6e-380">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d5c6e-380">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d5c6e-381">1.0</span><span class="sxs-lookup"><span data-stu-id="d5c6e-381">1.0</span></span>|
|[<span data-ttu-id="d5c6e-382">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d5c6e-382">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d5c6e-383">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d5c6e-383">ReadItem</span></span>|
|[<span data-ttu-id="d5c6e-384">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d5c6e-384">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d5c6e-385">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d5c6e-385">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d5c6e-386">Пример</span><span class="sxs-lookup"><span data-stu-id="d5c6e-386">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="d5c6e-387">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="d5c6e-387">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="d5c6e-388">Отображает форму для создания новой встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-388">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="d5c6e-389">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-389">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="d5c6e-p112">Метод `displayNewAppointmentForm` открывает форму, в которой пользователь может создать встречу или собрание. Если параметры заданы, поля формы встречи автоматически заполняются их содержимым.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-p112">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="d5c6e-p113">В Outlook в Интернете и на мобильных устройствах этот метод всегда отображает форму с полем участников. Если вы не укажете участников в качестве входных аргументов, метод отображает форму с кнопкой **Сохранить**. Если вы укажете участников, форма будет включать участников и кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-p113">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="d5c6e-p114">Если вы укажете участников или ресурсы с помощью параметра `requiredAttendees`, `optionalAttendees` или `resources` в клиенте Outlook с расширенными возможностями и Outlook RT, этот метод отобразит форму собрания с кнопкой **Отправить**. Если не указать получателей, этот метод отобразит форму встречи с кнопкой **Сохранить и закрыть**.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-p114">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="d5c6e-397">Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-397">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d5c6e-398">Параметры</span><span class="sxs-lookup"><span data-stu-id="d5c6e-398">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="d5c6e-399">Все параметры являются необязательными.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-399">All parameters are optional.</span></span>

|<span data-ttu-id="d5c6e-400">Имя</span><span class="sxs-lookup"><span data-stu-id="d5c6e-400">Name</span></span>| <span data-ttu-id="d5c6e-401">Тип</span><span class="sxs-lookup"><span data-stu-id="d5c6e-401">Type</span></span>| <span data-ttu-id="d5c6e-402">Описание</span><span class="sxs-lookup"><span data-stu-id="d5c6e-402">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="d5c6e-403">Object</span><span class="sxs-lookup"><span data-stu-id="d5c6e-403">Object</span></span> | <span data-ttu-id="d5c6e-404">Словарь параметров, описывающий новую встречу.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-404">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="d5c6e-405">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="d5c6e-405">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="d5c6e-p115">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из обязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="d5c6e-408">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="d5c6e-408">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="d5c6e-p116">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из необязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="d5c6e-411">Date</span><span class="sxs-lookup"><span data-stu-id="d5c6e-411">Date</span></span> | <span data-ttu-id="d5c6e-412">Объект `Date`, указывающий дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-412">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="d5c6e-413">Date</span><span class="sxs-lookup"><span data-stu-id="d5c6e-413">Date</span></span> | <span data-ttu-id="d5c6e-414">Объект `Date`, указывающий дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-414">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="d5c6e-415">String</span><span class="sxs-lookup"><span data-stu-id="d5c6e-415">String</span></span> | <span data-ttu-id="d5c6e-p117">Строка со сведениями о месте встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-p117">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="d5c6e-418">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="d5c6e-418">Array.&lt;String&gt;</span></span> | <span data-ttu-id="d5c6e-p118">Массив строк, содержащий необходимые для встречи ресурсы. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-p118">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="d5c6e-421">String</span><span class="sxs-lookup"><span data-stu-id="d5c6e-421">String</span></span> | <span data-ttu-id="d5c6e-p119">Строка с темой встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-p119">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="d5c6e-424">String</span><span class="sxs-lookup"><span data-stu-id="d5c6e-424">String</span></span> | <span data-ttu-id="d5c6e-p120">Текст сообщения о встрече. Максимальный размер содержимого сообщения — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-p120">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d5c6e-427">Требования</span><span class="sxs-lookup"><span data-stu-id="d5c6e-427">Requirements</span></span>

|<span data-ttu-id="d5c6e-428">Требование</span><span class="sxs-lookup"><span data-stu-id="d5c6e-428">Requirement</span></span>| <span data-ttu-id="d5c6e-429">Значение</span><span class="sxs-lookup"><span data-stu-id="d5c6e-429">Value</span></span>|
|---|---|
|[<span data-ttu-id="d5c6e-430">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d5c6e-430">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d5c6e-431">1.0</span><span class="sxs-lookup"><span data-stu-id="d5c6e-431">1.0</span></span>|
|[<span data-ttu-id="d5c6e-432">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d5c6e-432">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d5c6e-433">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d5c6e-433">ReadItem</span></span>|
|[<span data-ttu-id="d5c6e-434">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d5c6e-434">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d5c6e-435">Чтение</span><span class="sxs-lookup"><span data-stu-id="d5c6e-435">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d5c6e-436">Пример</span><span class="sxs-lookup"><span data-stu-id="d5c6e-436">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="d5c6e-437">Дисплайневмессажеформ (Parameters)</span><span class="sxs-lookup"><span data-stu-id="d5c6e-437">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="d5c6e-438">Отображает форму для создания нового сообщения.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-438">Displays a form for creating a new message.</span></span>

<span data-ttu-id="d5c6e-439">`displayNewMessageForm` Метод открывает форму, которая позволяет пользователю создать новое сообщение.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-439">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="d5c6e-440">Если указаны параметры, поля формы сообщения автоматически заполняются содержимым параметров.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-440">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="d5c6e-441">Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-441">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d5c6e-442">Параметры</span><span class="sxs-lookup"><span data-stu-id="d5c6e-442">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="d5c6e-443">Все параметры являются необязательными.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-443">All parameters are optional.</span></span>

|<span data-ttu-id="d5c6e-444">Имя</span><span class="sxs-lookup"><span data-stu-id="d5c6e-444">Name</span></span>| <span data-ttu-id="d5c6e-445">Тип</span><span class="sxs-lookup"><span data-stu-id="d5c6e-445">Type</span></span>| <span data-ttu-id="d5c6e-446">Описание</span><span class="sxs-lookup"><span data-stu-id="d5c6e-446">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="d5c6e-447">Object</span><span class="sxs-lookup"><span data-stu-id="d5c6e-447">Object</span></span> | <span data-ttu-id="d5c6e-448">Словарь параметров, описывающих новое сообщение.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-448">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="d5c6e-449">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="d5c6e-449">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="d5c6e-450">Массив строк, содержащий адреса электронной почты или массив, содержащий `EmailAddressDetails` объект для каждого из получателей в строке "Кому".</span><span class="sxs-lookup"><span data-stu-id="d5c6e-450">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="d5c6e-451">Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-451">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="d5c6e-452">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="d5c6e-452">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="d5c6e-453">Массив строк, содержащий адреса электронной почты или массив, содержащий `EmailAddressDetails` объект для каждого получателя в строке "копия".</span><span class="sxs-lookup"><span data-stu-id="d5c6e-453">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="d5c6e-454">Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-454">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="d5c6e-455">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="d5c6e-455">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="d5c6e-456">Массив строк, содержащий адреса электронной почты или массив, содержащий `EmailAddressDetails` объект для каждого из получателей, указанных в строке "СК".</span><span class="sxs-lookup"><span data-stu-id="d5c6e-456">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="d5c6e-457">Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-457">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="d5c6e-458">String</span><span class="sxs-lookup"><span data-stu-id="d5c6e-458">String</span></span> | <span data-ttu-id="d5c6e-459">Строка, содержащая тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-459">A string containing the subject of the message.</span></span> <span data-ttu-id="d5c6e-460">Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-460">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="d5c6e-461">String</span><span class="sxs-lookup"><span data-stu-id="d5c6e-461">String</span></span> | <span data-ttu-id="d5c6e-462">Текст сообщения в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-462">The HTML body of the message.</span></span> <span data-ttu-id="d5c6e-463">Максимальный размер содержимого сообщения — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-463">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="d5c6e-464">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="d5c6e-464">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="d5c6e-465">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-465">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="d5c6e-466">String</span><span class="sxs-lookup"><span data-stu-id="d5c6e-466">String</span></span> | <span data-ttu-id="d5c6e-p127">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-p127">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="d5c6e-469">Строка</span><span class="sxs-lookup"><span data-stu-id="d5c6e-469">String</span></span> | <span data-ttu-id="d5c6e-470">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-470">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="d5c6e-471">String</span><span class="sxs-lookup"><span data-stu-id="d5c6e-471">String</span></span> | <span data-ttu-id="d5c6e-p128">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-p128">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="d5c6e-474">Логический</span><span class="sxs-lookup"><span data-stu-id="d5c6e-474">Boolean</span></span> | <span data-ttu-id="d5c6e-p129">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-p129">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="d5c6e-477">Строка</span><span class="sxs-lookup"><span data-stu-id="d5c6e-477">String</span></span> | <span data-ttu-id="d5c6e-478">Используется, только если свойству `type` присвоено значение `item`.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-478">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="d5c6e-479">Идентификатор элемента EWS существующего сообщения электронной почты, которое необходимо присоединить к новому сообщению.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-479">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="d5c6e-480">Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-480">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="d5c6e-481">Требования</span><span class="sxs-lookup"><span data-stu-id="d5c6e-481">Requirements</span></span>

|<span data-ttu-id="d5c6e-482">Требование</span><span class="sxs-lookup"><span data-stu-id="d5c6e-482">Requirement</span></span>| <span data-ttu-id="d5c6e-483">Значение</span><span class="sxs-lookup"><span data-stu-id="d5c6e-483">Value</span></span>|
|---|---|
|[<span data-ttu-id="d5c6e-484">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="d5c6e-484">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d5c6e-485">1.6</span><span class="sxs-lookup"><span data-stu-id="d5c6e-485">1.6</span></span> |
|[<span data-ttu-id="d5c6e-486">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d5c6e-486">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d5c6e-487">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d5c6e-487">ReadItem</span></span>|
|[<span data-ttu-id="d5c6e-488">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d5c6e-488">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d5c6e-489">Чтение</span><span class="sxs-lookup"><span data-stu-id="d5c6e-489">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d5c6e-490">Пример</span><span class="sxs-lookup"><span data-stu-id="d5c6e-490">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="d5c6e-491">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="d5c6e-491">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="d5c6e-492">Возвращает строку, содержащую маркер, который используется для вызова интерфейсов REST API или веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-492">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="d5c6e-p131">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-p131">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="d5c6e-495">Рекомендуем сделать так, чтобы по мере возможности надстройки использовали интерфейсы REST API, а не веб-службы Exchange.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-495">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="d5c6e-496">Для вызова `getCallbackTokenAsync` метода в режиме чтения требуется минимальный уровень разрешений **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-496">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="d5c6e-497">Для `getCallbackTokenAsync` вызова в режиме создания необходимо сохранить элемент.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-497">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="d5c6e-498">[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) Метод требует наличия минимального уровня разрешений **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-498">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

<span data-ttu-id="d5c6e-499">**Маркеры REST**</span><span class="sxs-lookup"><span data-stu-id="d5c6e-499">**REST Tokens**</span></span>

<span data-ttu-id="d5c6e-p133">Если запрашивается маркер REST (`options.isRest = true`), полученный маркер не подойдет для проверки подлинности при вызовах веб-служб Exchange. Область действия маркера будет ограничена доступом только для чтения к текущему элементу и его вложениям, если в манифесте надстройки не указано разрешение [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission). Если указано разрешение `ReadWriteMailbox`, полученный маркер предоставит доступ на чтение и запись к почте, календарю и контактам, включая возможность отправки почты.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-p133">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="d5c6e-503">С помощью свойства `restUrl` надстройка должна определить правильный URL-адрес для вызовов REST API.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-503">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="d5c6e-504">**Маркеры EWS**</span><span class="sxs-lookup"><span data-stu-id="d5c6e-504">**EWS Tokens**</span></span>

<span data-ttu-id="d5c6e-p134">Если запрашивается маркер EWS (`options.isRest = false`), полученный маркер не подойдет для проверки подлинности при вызовах REST API. Область действия маркера будет ограничена доступом к текущему элементу.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-p134">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="d5c6e-507">С помощью свойства `ewsUrl` надстройка должна определить правильный URL-адрес для вызовов EWS.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-507">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

<span data-ttu-id="d5c6e-508">Можно передать как маркер, так и идентификатор вложения или идентификатор элемента в систему стороннего производителя.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-508">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="d5c6e-509">Третья система использует маркер в качестве маркера авторизации носителя, чтобы вызвать операцию [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) [или GetItem](/exchange/client-developer/web-service-reference/getitem-operation) веб-служб Exchange (EWS) для получения вложения или элемента.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-509">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to retrieve an attachment or item.</span></span> <span data-ttu-id="d5c6e-510">Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="d5c6e-510">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="parameters"></a><span data-ttu-id="d5c6e-511">Параметры</span><span class="sxs-lookup"><span data-stu-id="d5c6e-511">Parameters</span></span>

|<span data-ttu-id="d5c6e-512">Имя</span><span class="sxs-lookup"><span data-stu-id="d5c6e-512">Name</span></span>| <span data-ttu-id="d5c6e-513">Тип</span><span class="sxs-lookup"><span data-stu-id="d5c6e-513">Type</span></span>| <span data-ttu-id="d5c6e-514">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="d5c6e-514">Attributes</span></span>| <span data-ttu-id="d5c6e-515">Описание</span><span class="sxs-lookup"><span data-stu-id="d5c6e-515">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="d5c6e-516">Object</span><span class="sxs-lookup"><span data-stu-id="d5c6e-516">Object</span></span> | <span data-ttu-id="d5c6e-517">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d5c6e-517">&lt;optional&gt;</span></span> | <span data-ttu-id="d5c6e-518">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-518">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="d5c6e-519">Boolean</span><span class="sxs-lookup"><span data-stu-id="d5c6e-519">Boolean</span></span> |  <span data-ttu-id="d5c6e-520">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d5c6e-520">&lt;optional&gt;</span></span> | <span data-ttu-id="d5c6e-p136">Определяет, будет ли предоставленный маркер использоваться для интерфейсов REST API Outlook или веб-служб Exchange. Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-p136">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="d5c6e-523">Object</span><span class="sxs-lookup"><span data-stu-id="d5c6e-523">Object</span></span> |  <span data-ttu-id="d5c6e-524">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d5c6e-524">&lt;optional&gt;</span></span> | <span data-ttu-id="d5c6e-525">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-525">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="d5c6e-526">функция</span><span class="sxs-lookup"><span data-stu-id="d5c6e-526">function</span></span>||<span data-ttu-id="d5c6e-527">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d5c6e-527">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d5c6e-528">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-528">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="d5c6e-529">При наличии ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут предоставлять дополнительные сведения.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-529">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d5c6e-530">Ошибки</span><span class="sxs-lookup"><span data-stu-id="d5c6e-530">Errors</span></span>

|<span data-ttu-id="d5c6e-531">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="d5c6e-531">Error code</span></span>|<span data-ttu-id="d5c6e-532">Описание</span><span class="sxs-lookup"><span data-stu-id="d5c6e-532">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="d5c6e-533">Не удалось выполнить запрос.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-533">The request has failed.</span></span> <span data-ttu-id="d5c6e-534">Просмотрите объект диагностики для кода ошибки HTTP.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-534">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="d5c6e-535">Сервер Exchange Server вернул ошибку.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-535">The Exchange server returned an error.</span></span> <span data-ttu-id="d5c6e-536">Для получения дополнительных сведений просмотрите объект диагностики.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-536">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="d5c6e-537">Пользователь отключен от сети.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-537">The user is no longer connected to the network.</span></span> <span data-ttu-id="d5c6e-538">Проверьте сетевое подключение и повторите попытку.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-538">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d5c6e-539">Требования</span><span class="sxs-lookup"><span data-stu-id="d5c6e-539">Requirements</span></span>

|<span data-ttu-id="d5c6e-540">Требование</span><span class="sxs-lookup"><span data-stu-id="d5c6e-540">Requirement</span></span>| <span data-ttu-id="d5c6e-541">Значение</span><span class="sxs-lookup"><span data-stu-id="d5c6e-541">Value</span></span>|
|---|---|
|[<span data-ttu-id="d5c6e-542">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="d5c6e-542">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d5c6e-543">1.5</span><span class="sxs-lookup"><span data-stu-id="d5c6e-543">1.5</span></span> |
|[<span data-ttu-id="d5c6e-544">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d5c6e-544">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d5c6e-545">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d5c6e-545">ReadItem</span></span>|
|[<span data-ttu-id="d5c6e-546">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d5c6e-546">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d5c6e-547">Создание и чтение</span><span class="sxs-lookup"><span data-stu-id="d5c6e-547">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="d5c6e-548">Пример</span><span class="sxs-lookup"><span data-stu-id="d5c6e-548">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="d5c6e-549">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="d5c6e-549">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="d5c6e-550">Получает строку, содержащую маркер, используемый для получения вложения или элемента с Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-550">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="d5c6e-p140">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-p140">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="d5c6e-553">Можно передать как маркер, так и идентификатор вложения или идентификатор элемента в систему стороннего производителя.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-553">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="d5c6e-554">Третья система использует маркер в качестве маркера авторизации носителя, чтобы вызвать операцию [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) [или GetItem](/exchange/client-developer/web-service-reference/getitem-operation) веб-служб Exchange (EWS) для возврата вложения или элемента.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-554">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="d5c6e-555">Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="d5c6e-555">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="d5c6e-556">Для вызова `getCallbackTokenAsync` метода в режиме чтения требуется минимальный уровень разрешений **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-556">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="d5c6e-557">Для `getCallbackTokenAsync` вызова в режиме создания необходимо сохранить элемент.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-557">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="d5c6e-558">[`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) Метод требует наличия минимального уровня разрешений **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-558">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d5c6e-559">Параметры</span><span class="sxs-lookup"><span data-stu-id="d5c6e-559">Parameters</span></span>

|<span data-ttu-id="d5c6e-560">Имя</span><span class="sxs-lookup"><span data-stu-id="d5c6e-560">Name</span></span>| <span data-ttu-id="d5c6e-561">Тип</span><span class="sxs-lookup"><span data-stu-id="d5c6e-561">Type</span></span>| <span data-ttu-id="d5c6e-562">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="d5c6e-562">Attributes</span></span>| <span data-ttu-id="d5c6e-563">Описание</span><span class="sxs-lookup"><span data-stu-id="d5c6e-563">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="d5c6e-564">функция</span><span class="sxs-lookup"><span data-stu-id="d5c6e-564">function</span></span>||<span data-ttu-id="d5c6e-565">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d5c6e-565">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d5c6e-566">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-566">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="d5c6e-567">При наличии ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут предоставлять дополнительные сведения.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-567">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="d5c6e-568">Объект</span><span class="sxs-lookup"><span data-stu-id="d5c6e-568">Object</span></span>| <span data-ttu-id="d5c6e-569">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d5c6e-569">&lt;optional&gt;</span></span>|<span data-ttu-id="d5c6e-570">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-570">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d5c6e-571">Ошибки</span><span class="sxs-lookup"><span data-stu-id="d5c6e-571">Errors</span></span>

|<span data-ttu-id="d5c6e-572">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="d5c6e-572">Error code</span></span>|<span data-ttu-id="d5c6e-573">Описание</span><span class="sxs-lookup"><span data-stu-id="d5c6e-573">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="d5c6e-574">Не удалось выполнить запрос.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-574">The request has failed.</span></span> <span data-ttu-id="d5c6e-575">Просмотрите объект диагностики для кода ошибки HTTP.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-575">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="d5c6e-576">Сервер Exchange Server вернул ошибку.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-576">The Exchange server returned an error.</span></span> <span data-ttu-id="d5c6e-577">Для получения дополнительных сведений просмотрите объект диагностики.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-577">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="d5c6e-578">Пользователь отключен от сети.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-578">The user is no longer connected to the network.</span></span> <span data-ttu-id="d5c6e-579">Проверьте сетевое подключение и повторите попытку.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-579">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d5c6e-580">Требования</span><span class="sxs-lookup"><span data-stu-id="d5c6e-580">Requirements</span></span>

|<span data-ttu-id="d5c6e-581">Требование</span><span class="sxs-lookup"><span data-stu-id="d5c6e-581">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="d5c6e-582">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d5c6e-582">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d5c6e-583">1.0</span><span class="sxs-lookup"><span data-stu-id="d5c6e-583">1.0</span></span> | <span data-ttu-id="d5c6e-584">1.3</span><span class="sxs-lookup"><span data-stu-id="d5c6e-584">1.3</span></span> |
|[<span data-ttu-id="d5c6e-585">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d5c6e-585">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d5c6e-586">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d5c6e-586">ReadItem</span></span> | <span data-ttu-id="d5c6e-587">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d5c6e-587">ReadItem</span></span> |
|[<span data-ttu-id="d5c6e-588">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d5c6e-588">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d5c6e-589">Чтение</span><span class="sxs-lookup"><span data-stu-id="d5c6e-589">Read</span></span> | <span data-ttu-id="d5c6e-590">Создание</span><span class="sxs-lookup"><span data-stu-id="d5c6e-590">Compose</span></span> |

##### <a name="example"></a><span data-ttu-id="d5c6e-591">Пример</span><span class="sxs-lookup"><span data-stu-id="d5c6e-591">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="d5c6e-592">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="d5c6e-592">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="d5c6e-593">Получает маркер, идентифицирующий пользователя и надстройку Office.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-593">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="d5c6e-594">Метод `getUserIdentityTokenAsync` возвращает маркер, который можно использовать для идентификации, а также [проверки подлинности надстройки и пользователя в сторонней системе](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="d5c6e-594">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="d5c6e-595">Параметры</span><span class="sxs-lookup"><span data-stu-id="d5c6e-595">Parameters</span></span>

|<span data-ttu-id="d5c6e-596">Имя</span><span class="sxs-lookup"><span data-stu-id="d5c6e-596">Name</span></span>| <span data-ttu-id="d5c6e-597">Тип</span><span class="sxs-lookup"><span data-stu-id="d5c6e-597">Type</span></span>| <span data-ttu-id="d5c6e-598">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="d5c6e-598">Attributes</span></span>| <span data-ttu-id="d5c6e-599">Описание</span><span class="sxs-lookup"><span data-stu-id="d5c6e-599">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="d5c6e-600">функция</span><span class="sxs-lookup"><span data-stu-id="d5c6e-600">function</span></span>||<span data-ttu-id="d5c6e-601">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d5c6e-601">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d5c6e-602">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-602">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="d5c6e-603">При наличии ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут предоставлять дополнительные сведения.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-603">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="d5c6e-604">Объект</span><span class="sxs-lookup"><span data-stu-id="d5c6e-604">Object</span></span>| <span data-ttu-id="d5c6e-605">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d5c6e-605">&lt;optional&gt;</span></span>|<span data-ttu-id="d5c6e-606">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-606">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="d5c6e-607">Ошибки</span><span class="sxs-lookup"><span data-stu-id="d5c6e-607">Errors</span></span>

|<span data-ttu-id="d5c6e-608">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="d5c6e-608">Error code</span></span>|<span data-ttu-id="d5c6e-609">Описание</span><span class="sxs-lookup"><span data-stu-id="d5c6e-609">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="d5c6e-610">Не удалось выполнить запрос.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-610">The request has failed.</span></span> <span data-ttu-id="d5c6e-611">Просмотрите объект диагностики для кода ошибки HTTP.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-611">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="d5c6e-612">Сервер Exchange Server вернул ошибку.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-612">The Exchange server returned an error.</span></span> <span data-ttu-id="d5c6e-613">Для получения дополнительных сведений просмотрите объект диагностики.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-613">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="d5c6e-614">Пользователь отключен от сети.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-614">The user is no longer connected to the network.</span></span> <span data-ttu-id="d5c6e-615">Проверьте сетевое подключение и повторите попытку.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-615">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d5c6e-616">Требования</span><span class="sxs-lookup"><span data-stu-id="d5c6e-616">Requirements</span></span>

|<span data-ttu-id="d5c6e-617">Требование</span><span class="sxs-lookup"><span data-stu-id="d5c6e-617">Requirement</span></span>| <span data-ttu-id="d5c6e-618">Значение</span><span class="sxs-lookup"><span data-stu-id="d5c6e-618">Value</span></span>|
|---|---|
|[<span data-ttu-id="d5c6e-619">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d5c6e-619">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d5c6e-620">1.0</span><span class="sxs-lookup"><span data-stu-id="d5c6e-620">1.0</span></span>|
|[<span data-ttu-id="d5c6e-621">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d5c6e-621">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d5c6e-622">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d5c6e-622">ReadItem</span></span>|
|[<span data-ttu-id="d5c6e-623">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d5c6e-623">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d5c6e-624">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d5c6e-624">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d5c6e-625">Пример</span><span class="sxs-lookup"><span data-stu-id="d5c6e-625">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="d5c6e-626">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="d5c6e-626">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="d5c6e-627">Выполняет асинхронный запрос для веб-служб Exchange (EWS) на сервере Exchange Server, на котором размещен почтовый ящик пользователя.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-627">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="d5c6e-628">Этот метод не поддерживается в следующих сценариях:</span><span class="sxs-lookup"><span data-stu-id="d5c6e-628">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="d5c6e-629">В Outlook для iOS и Android</span><span class="sxs-lookup"><span data-stu-id="d5c6e-629">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="d5c6e-630">Если надстройка загружается в почтовый ящик Gmail.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-630">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="d5c6e-631">В таких случаях надстройка должна [использовать REST API](/outlook/add-ins/use-rest-api) для доступа к почтовому ящику пользователя.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-631">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="d5c6e-632">Метод `makeEwsRequestAsync` отправляет запрос EWS от имени надстройки в Exchange.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-632">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="d5c6e-633">Список поддерживаемых операций EWS см. в статье [Вызов веб-служб из надстройки Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support).</span><span class="sxs-lookup"><span data-stu-id="d5c6e-633">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="d5c6e-634">С помощью метода `makeEwsRequestAsync` невозможно запрашивать элементы, связанные с папкой.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-634">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="d5c6e-635">В запросе XML должна быть указана кодировка UTF-8.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-635">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="d5c6e-p150">У вашей надстройки должно быть разрешение **ReadWriteMailbox** для использования метода `makeEwsRequestAsync`. Сведения об использовании разрешения **ReadWriteMailbox** и операций EWS, которые можно вызывать с помощью метода `makeEwsRequestAsync`, см. в статье [Указание разрешений для доступа почтовой надстройки к почтовому ящику пользователя](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="d5c6e-p150">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="d5c6e-638">Администратор сервера должен установить значение true для параметра `OAuthAuthentication` в каталоге сервера клиентского доступа EWS, чтобы метод `makeEwsRequestAsync` мог выполнять запросы EWS.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-638">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="d5c6e-639">Различия версий</span><span class="sxs-lookup"><span data-stu-id="d5c6e-639">Version differences</span></span>

<span data-ttu-id="d5c6e-640">Если вы используете метод `makeEwsRequestAsync` в почтовых приложениях, которые выполняются в Outlook версии более ранней, чем 15.0.4535.1004, указывайте кодировку `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-640">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="d5c6e-641">Значение кодировки не нужно указывать, если почтовое приложение выполняется в Outlook в Интернете.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-641">You do not need to set the encoding value when your mail app is running in Outlook on the web.</span></span> <span data-ttu-id="d5c6e-642">Вы можете определить, работает ли почтовое приложение в Outlook в Интернете или на настольном клиенте с помощью свойства Mailbox. Diagnostics. hostName.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-642">You can determine whether your mail app is running in Outlook on the web or a desktop client by using the mailbox.diagnostics.hostName property.</span></span> <span data-ttu-id="d5c6e-643">Используемую версию Outlook можно определить с помощью свойства mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-643">You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d5c6e-644">Параметры</span><span class="sxs-lookup"><span data-stu-id="d5c6e-644">Parameters</span></span>

|<span data-ttu-id="d5c6e-645">Имя</span><span class="sxs-lookup"><span data-stu-id="d5c6e-645">Name</span></span>| <span data-ttu-id="d5c6e-646">Тип</span><span class="sxs-lookup"><span data-stu-id="d5c6e-646">Type</span></span>| <span data-ttu-id="d5c6e-647">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="d5c6e-647">Attributes</span></span>| <span data-ttu-id="d5c6e-648">Описание</span><span class="sxs-lookup"><span data-stu-id="d5c6e-648">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="d5c6e-649">String</span><span class="sxs-lookup"><span data-stu-id="d5c6e-649">String</span></span>||<span data-ttu-id="d5c6e-650">Запрос EWS.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-650">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="d5c6e-651">function</span><span class="sxs-lookup"><span data-stu-id="d5c6e-651">function</span></span>||<span data-ttu-id="d5c6e-652">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d5c6e-652">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d5c6e-653">Результат XML вызова EWS указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-653">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="d5c6e-654">Если размер результата превышает 1 МБ, возвращается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-654">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="d5c6e-655">Object</span><span class="sxs-lookup"><span data-stu-id="d5c6e-655">Object</span></span>| <span data-ttu-id="d5c6e-656">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d5c6e-656">&lt;optional&gt;</span></span>|<span data-ttu-id="d5c6e-657">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-657">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d5c6e-658">Требования</span><span class="sxs-lookup"><span data-stu-id="d5c6e-658">Requirements</span></span>

|<span data-ttu-id="d5c6e-659">Требование</span><span class="sxs-lookup"><span data-stu-id="d5c6e-659">Requirement</span></span>| <span data-ttu-id="d5c6e-660">Значение</span><span class="sxs-lookup"><span data-stu-id="d5c6e-660">Value</span></span>|
|---|---|
|[<span data-ttu-id="d5c6e-661">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d5c6e-661">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d5c6e-662">1.0</span><span class="sxs-lookup"><span data-stu-id="d5c6e-662">1.0</span></span>|
|[<span data-ttu-id="d5c6e-663">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d5c6e-663">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d5c6e-664">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="d5c6e-664">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="d5c6e-665">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d5c6e-665">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d5c6e-666">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d5c6e-666">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d5c6e-667">Пример</span><span class="sxs-lookup"><span data-stu-id="d5c6e-667">Example</span></span>

<span data-ttu-id="d5c6e-668">В приведенном ниже примере вызывается `makeEwsRequestAsync` для получения темы элемента с помощью операции `GetItem`.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-668">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="d5c6e-669">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d5c6e-669">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="d5c6e-670">Удаляет обработчиков для поддерживаемого типа события.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-670">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="d5c6e-671">В настоящее время поддерживаются типы событий `Office.EventType.ItemChanged` и `Office.EventType.OfficeThemeChanged`.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-671">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d5c6e-672">Параметры</span><span class="sxs-lookup"><span data-stu-id="d5c6e-672">Parameters</span></span>

| <span data-ttu-id="d5c6e-673">Имя</span><span class="sxs-lookup"><span data-stu-id="d5c6e-673">Name</span></span> | <span data-ttu-id="d5c6e-674">Тип</span><span class="sxs-lookup"><span data-stu-id="d5c6e-674">Type</span></span> | <span data-ttu-id="d5c6e-675">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="d5c6e-675">Attributes</span></span> | <span data-ttu-id="d5c6e-676">Описание</span><span class="sxs-lookup"><span data-stu-id="d5c6e-676">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="d5c6e-677">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="d5c6e-677">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="d5c6e-678">Событие, которое должно отменить обработчик.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-678">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="d5c6e-679">Object</span><span class="sxs-lookup"><span data-stu-id="d5c6e-679">Object</span></span> | <span data-ttu-id="d5c6e-680">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d5c6e-680">&lt;optional&gt;</span></span> | <span data-ttu-id="d5c6e-681">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-681">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="d5c6e-682">Object</span><span class="sxs-lookup"><span data-stu-id="d5c6e-682">Object</span></span> | <span data-ttu-id="d5c6e-683">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d5c6e-683">&lt;optional&gt;</span></span> | <span data-ttu-id="d5c6e-684">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="d5c6e-684">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="d5c6e-685">функция</span><span class="sxs-lookup"><span data-stu-id="d5c6e-685">function</span></span>| <span data-ttu-id="d5c6e-686">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d5c6e-686">&lt;optional&gt;</span></span>|<span data-ttu-id="d5c6e-687">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d5c6e-687">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d5c6e-688">Требования</span><span class="sxs-lookup"><span data-stu-id="d5c6e-688">Requirements</span></span>

|<span data-ttu-id="d5c6e-689">Требование</span><span class="sxs-lookup"><span data-stu-id="d5c6e-689">Requirement</span></span>| <span data-ttu-id="d5c6e-690">Значение</span><span class="sxs-lookup"><span data-stu-id="d5c6e-690">Value</span></span>|
|---|---|
|[<span data-ttu-id="d5c6e-691">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="d5c6e-691">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d5c6e-692">1.5</span><span class="sxs-lookup"><span data-stu-id="d5c6e-692">1.5</span></span> |
|[<span data-ttu-id="d5c6e-693">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d5c6e-693">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d5c6e-694">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d5c6e-694">ReadItem</span></span> |
|[<span data-ttu-id="d5c6e-695">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d5c6e-695">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d5c6e-696">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d5c6e-696">Compose or Read</span></span>|
