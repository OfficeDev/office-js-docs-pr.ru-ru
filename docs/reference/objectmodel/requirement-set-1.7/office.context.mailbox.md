---
title: Office. Context. Mailbox — набор обязательных элементов 1,7
description: ''
ms.date: 11/27/2019
localization_priority: Normal
ms.openlocfilehash: c310ad38bb9821955fb0571d3693ce39715376f4
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629674"
---
# <a name="mailbox"></a><span data-ttu-id="19f86-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="19f86-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="19f86-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="19f86-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="19f86-104">Предоставляет для Microsoft Outlook доступ к объектной модели надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="19f86-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="19f86-105">Требования</span><span class="sxs-lookup"><span data-stu-id="19f86-105">Requirements</span></span>

|<span data-ttu-id="19f86-106">Требование</span><span class="sxs-lookup"><span data-stu-id="19f86-106">Requirement</span></span>| <span data-ttu-id="19f86-107">Значение</span><span class="sxs-lookup"><span data-stu-id="19f86-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="19f86-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="19f86-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="19f86-109">1.0</span><span class="sxs-lookup"><span data-stu-id="19f86-109">1.0</span></span>|
|[<span data-ttu-id="19f86-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="19f86-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="19f86-111">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="19f86-111">Restricted</span></span>|
|[<span data-ttu-id="19f86-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="19f86-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="19f86-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="19f86-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="19f86-114">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="19f86-114">Members and methods</span></span>

| <span data-ttu-id="19f86-115">Элемент</span><span class="sxs-lookup"><span data-stu-id="19f86-115">Member</span></span> | <span data-ttu-id="19f86-116">Тип</span><span class="sxs-lookup"><span data-stu-id="19f86-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="19f86-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="19f86-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="19f86-118">Элемент</span><span class="sxs-lookup"><span data-stu-id="19f86-118">Member</span></span> |
| [<span data-ttu-id="19f86-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="19f86-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="19f86-120">Элемент</span><span class="sxs-lookup"><span data-stu-id="19f86-120">Member</span></span> |
| [<span data-ttu-id="19f86-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="19f86-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="19f86-122">Метод</span><span class="sxs-lookup"><span data-stu-id="19f86-122">Method</span></span> |
| [<span data-ttu-id="19f86-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="19f86-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="19f86-124">Метод</span><span class="sxs-lookup"><span data-stu-id="19f86-124">Method</span></span> |
| [<span data-ttu-id="19f86-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="19f86-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="19f86-126">Метод</span><span class="sxs-lookup"><span data-stu-id="19f86-126">Method</span></span> |
| [<span data-ttu-id="19f86-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="19f86-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="19f86-128">Метод</span><span class="sxs-lookup"><span data-stu-id="19f86-128">Method</span></span> |
| [<span data-ttu-id="19f86-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="19f86-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="19f86-130">Метод</span><span class="sxs-lookup"><span data-stu-id="19f86-130">Method</span></span> |
| [<span data-ttu-id="19f86-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="19f86-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="19f86-132">Метод</span><span class="sxs-lookup"><span data-stu-id="19f86-132">Method</span></span> |
| [<span data-ttu-id="19f86-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="19f86-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="19f86-134">Метод</span><span class="sxs-lookup"><span data-stu-id="19f86-134">Method</span></span> |
| [<span data-ttu-id="19f86-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="19f86-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="19f86-136">Метод</span><span class="sxs-lookup"><span data-stu-id="19f86-136">Method</span></span> |
| [<span data-ttu-id="19f86-137">дисплайневмессажеформ</span><span class="sxs-lookup"><span data-stu-id="19f86-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="19f86-138">Метод</span><span class="sxs-lookup"><span data-stu-id="19f86-138">Method</span></span> |
| [<span data-ttu-id="19f86-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="19f86-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="19f86-140">Метод</span><span class="sxs-lookup"><span data-stu-id="19f86-140">Method</span></span> |
| [<span data-ttu-id="19f86-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="19f86-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="19f86-142">Метод</span><span class="sxs-lookup"><span data-stu-id="19f86-142">Method</span></span> |
| [<span data-ttu-id="19f86-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="19f86-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="19f86-144">Метод</span><span class="sxs-lookup"><span data-stu-id="19f86-144">Method</span></span> |
| [<span data-ttu-id="19f86-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="19f86-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="19f86-146">Метод</span><span class="sxs-lookup"><span data-stu-id="19f86-146">Method</span></span> |
| [<span data-ttu-id="19f86-147">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="19f86-147">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="19f86-148">Метод</span><span class="sxs-lookup"><span data-stu-id="19f86-148">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="19f86-149">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="19f86-149">Namespaces</span></span>

<span data-ttu-id="19f86-150">[diagnostics](Office.context.mailbox.diagnostics.md). Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="19f86-150">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="19f86-151">[item](Office.context.mailbox.item.md). Предоставляет методы и свойства для доступа к сообщению или встрече в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="19f86-151">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="19f86-152">[userProfile](Office.context.mailbox.userProfile.md). Предоставляет сведения о пользователе в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="19f86-152">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="19f86-153">Members</span><span class="sxs-lookup"><span data-stu-id="19f86-153">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="19f86-154">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="19f86-154">ewsUrl: String</span></span>

<span data-ttu-id="19f86-p101">Получает URL-адрес конечной точки веб-служб Exchange (EWS) для этой учетной записи электронной почты. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="19f86-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="19f86-157">Этот элемент не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="19f86-157">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="19f86-p102">Удаленная служба может использовать значение `ewsUrl`, чтобы выполнять вызовы EWS для почтового ящика пользователя. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="19f86-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="19f86-160">Чтобы вызвать элемент `ewsUrl` в режиме чтения, в манифесте приложения должно быть указано разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="19f86-160">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="19f86-p103">Перед использованием элемента `ewsUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="19f86-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="19f86-163">Тип</span><span class="sxs-lookup"><span data-stu-id="19f86-163">Type</span></span>

*   <span data-ttu-id="19f86-164">String</span><span class="sxs-lookup"><span data-stu-id="19f86-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="19f86-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="19f86-165">Requirements</span></span>

|<span data-ttu-id="19f86-166">Требование</span><span class="sxs-lookup"><span data-stu-id="19f86-166">Requirement</span></span>| <span data-ttu-id="19f86-167">Значение</span><span class="sxs-lookup"><span data-stu-id="19f86-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="19f86-168">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="19f86-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="19f86-169">1.0</span><span class="sxs-lookup"><span data-stu-id="19f86-169">1.0</span></span>|
|[<span data-ttu-id="19f86-170">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="19f86-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="19f86-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="19f86-171">ReadItem</span></span>|
|[<span data-ttu-id="19f86-172">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="19f86-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="19f86-173">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="19f86-173">Compose or Read</span></span>|

<br>

---
---

#### <a name="resturl-string"></a><span data-ttu-id="19f86-174">restUrl: String</span><span class="sxs-lookup"><span data-stu-id="19f86-174">restUrl: String</span></span>

<span data-ttu-id="19f86-175">Возвращает URL-адрес конечной точки REST для этой учетной записи электронной почты.</span><span class="sxs-lookup"><span data-stu-id="19f86-175">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="19f86-176">С помощью значения `restUrl` можно выполнять вызовы [REST API](/outlook/rest/) для почтового ящика пользователя.</span><span class="sxs-lookup"><span data-stu-id="19f86-176">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="19f86-177">Тип</span><span class="sxs-lookup"><span data-stu-id="19f86-177">Type</span></span>

*   <span data-ttu-id="19f86-178">String</span><span class="sxs-lookup"><span data-stu-id="19f86-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="19f86-179">Requirements</span><span class="sxs-lookup"><span data-stu-id="19f86-179">Requirements</span></span>

|<span data-ttu-id="19f86-180">Требование</span><span class="sxs-lookup"><span data-stu-id="19f86-180">Requirement</span></span>| <span data-ttu-id="19f86-181">Значение</span><span class="sxs-lookup"><span data-stu-id="19f86-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="19f86-182">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="19f86-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="19f86-183">1.5</span><span class="sxs-lookup"><span data-stu-id="19f86-183">1.5</span></span> |
|[<span data-ttu-id="19f86-184">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="19f86-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="19f86-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="19f86-185">ReadItem</span></span>|
|[<span data-ttu-id="19f86-186">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="19f86-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="19f86-187">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="19f86-187">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="19f86-188">Методы</span><span class="sxs-lookup"><span data-stu-id="19f86-188">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="19f86-189">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="19f86-189">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="19f86-190">Добавляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="19f86-190">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="19f86-191">В настоящее время единственный поддерживаемый тип события — `Office.EventType.ItemChanged`.</span><span class="sxs-lookup"><span data-stu-id="19f86-191">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="19f86-192">Параметры</span><span class="sxs-lookup"><span data-stu-id="19f86-192">Parameters</span></span>

| <span data-ttu-id="19f86-193">Имя</span><span class="sxs-lookup"><span data-stu-id="19f86-193">Name</span></span> | <span data-ttu-id="19f86-194">Тип</span><span class="sxs-lookup"><span data-stu-id="19f86-194">Type</span></span> | <span data-ttu-id="19f86-195">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="19f86-195">Attributes</span></span> | <span data-ttu-id="19f86-196">Описание</span><span class="sxs-lookup"><span data-stu-id="19f86-196">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="19f86-197">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="19f86-197">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="19f86-198">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="19f86-198">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="19f86-199">Function</span><span class="sxs-lookup"><span data-stu-id="19f86-199">Function</span></span> || <span data-ttu-id="19f86-p104">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="19f86-p104">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="19f86-203">Объект</span><span class="sxs-lookup"><span data-stu-id="19f86-203">Object</span></span> | <span data-ttu-id="19f86-204">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="19f86-204">&lt;optional&gt;</span></span> | <span data-ttu-id="19f86-205">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="19f86-205">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="19f86-206">Object</span><span class="sxs-lookup"><span data-stu-id="19f86-206">Object</span></span> | <span data-ttu-id="19f86-207">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="19f86-207">&lt;optional&gt;</span></span> | <span data-ttu-id="19f86-208">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="19f86-208">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="19f86-209">функция</span><span class="sxs-lookup"><span data-stu-id="19f86-209">function</span></span>| <span data-ttu-id="19f86-210">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="19f86-210">&lt;optional&gt;</span></span>|<span data-ttu-id="19f86-211">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="19f86-211">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="19f86-212">Requirements</span><span class="sxs-lookup"><span data-stu-id="19f86-212">Requirements</span></span>

|<span data-ttu-id="19f86-213">Требование</span><span class="sxs-lookup"><span data-stu-id="19f86-213">Requirement</span></span>| <span data-ttu-id="19f86-214">Значение</span><span class="sxs-lookup"><span data-stu-id="19f86-214">Value</span></span>|
|---|---|
|[<span data-ttu-id="19f86-215">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="19f86-215">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="19f86-216">1.5</span><span class="sxs-lookup"><span data-stu-id="19f86-216">1.5</span></span> |
|[<span data-ttu-id="19f86-217">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="19f86-217">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="19f86-218">ReadItem</span><span class="sxs-lookup"><span data-stu-id="19f86-218">ReadItem</span></span> |
|[<span data-ttu-id="19f86-219">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="19f86-219">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="19f86-220">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="19f86-220">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="19f86-221">Пример</span><span class="sxs-lookup"><span data-stu-id="19f86-221">Example</span></span>

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
};
```

<br>

---
---

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="19f86-222">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="19f86-222">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="19f86-223">Преобразовывает идентификатор элемента из формата REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="19f86-223">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="19f86-224">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="19f86-224">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="19f86-p105">Формат идентификаторов, извлекаемых через API REST (например, [API Почты Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)), отличается от формата веб-служб Exchange (EWS). Метод `convertToEwsId` преобразовывает идентификатор в формате REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="19f86-p105">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="19f86-227">Параметры</span><span class="sxs-lookup"><span data-stu-id="19f86-227">Parameters</span></span>

|<span data-ttu-id="19f86-228">Имя</span><span class="sxs-lookup"><span data-stu-id="19f86-228">Name</span></span>| <span data-ttu-id="19f86-229">Тип</span><span class="sxs-lookup"><span data-stu-id="19f86-229">Type</span></span>| <span data-ttu-id="19f86-230">Описание</span><span class="sxs-lookup"><span data-stu-id="19f86-230">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="19f86-231">String</span><span class="sxs-lookup"><span data-stu-id="19f86-231">String</span></span>|<span data-ttu-id="19f86-232">Идентификатор элемента в формате REST API для Outlook</span><span class="sxs-lookup"><span data-stu-id="19f86-232">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="19f86-233">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="19f86-233">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.7)|<span data-ttu-id="19f86-234">Значение, определяющее версию REST API для Outlook, которая используется для извлечения идентификатора элемента.</span><span class="sxs-lookup"><span data-stu-id="19f86-234">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="19f86-235">Требования</span><span class="sxs-lookup"><span data-stu-id="19f86-235">Requirements</span></span>

|<span data-ttu-id="19f86-236">Требование</span><span class="sxs-lookup"><span data-stu-id="19f86-236">Requirement</span></span>| <span data-ttu-id="19f86-237">Значение</span><span class="sxs-lookup"><span data-stu-id="19f86-237">Value</span></span>|
|---|---|
|[<span data-ttu-id="19f86-238">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="19f86-238">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="19f86-239">1.3</span><span class="sxs-lookup"><span data-stu-id="19f86-239">1.3</span></span>|
|[<span data-ttu-id="19f86-240">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="19f86-240">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="19f86-241">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="19f86-241">Restricted</span></span>|
|[<span data-ttu-id="19f86-242">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="19f86-242">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="19f86-243">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="19f86-243">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="19f86-244">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="19f86-244">Returns:</span></span>

<span data-ttu-id="19f86-245">Тип: String</span><span class="sxs-lookup"><span data-stu-id="19f86-245">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="19f86-246">Пример</span><span class="sxs-lookup"><span data-stu-id="19f86-246">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-17"></a><span data-ttu-id="19f86-247">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.7)}</span><span class="sxs-lookup"><span data-stu-id="19f86-247">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.7)}</span></span>

<span data-ttu-id="19f86-248">Получает словарь, содержащий сведения о локальном времени клиента.</span><span class="sxs-lookup"><span data-stu-id="19f86-248">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="19f86-p106">Почтовое приложение для классической версии Outlook или версии в Интернете может использовать разные часовые пояса для дат и времени. Классическое приложение Outlook использует часовой пояс клиентского компьютера. Outlook в Интернете использует часовой пояс, заданный в Центре администрирования Exchange (EAC). Значения даты и времени должны обрабатываться так, чтобы значения в пользовательском интерфейсе всегда согласовывались с часовым поясом, ожидаемым пользователем.</span><span class="sxs-lookup"><span data-stu-id="19f86-p106">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="19f86-p107">Если почтовое приложение работает в классическом клиенте Outlook, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса клиентского компьютера. Если почтовое приложение работает в Outlook в Интернете, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса, заданного в Центре администрирования Exchange.</span><span class="sxs-lookup"><span data-stu-id="19f86-p107">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="19f86-254">Параметры</span><span class="sxs-lookup"><span data-stu-id="19f86-254">Parameters</span></span>

|<span data-ttu-id="19f86-255">Имя</span><span class="sxs-lookup"><span data-stu-id="19f86-255">Name</span></span>| <span data-ttu-id="19f86-256">Тип</span><span class="sxs-lookup"><span data-stu-id="19f86-256">Type</span></span>| <span data-ttu-id="19f86-257">Описание</span><span class="sxs-lookup"><span data-stu-id="19f86-257">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="19f86-258">Date</span><span class="sxs-lookup"><span data-stu-id="19f86-258">Date</span></span>|<span data-ttu-id="19f86-259">Объект Date</span><span class="sxs-lookup"><span data-stu-id="19f86-259">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="19f86-260">Requirements</span><span class="sxs-lookup"><span data-stu-id="19f86-260">Requirements</span></span>

|<span data-ttu-id="19f86-261">Требование</span><span class="sxs-lookup"><span data-stu-id="19f86-261">Requirement</span></span>| <span data-ttu-id="19f86-262">Значение</span><span class="sxs-lookup"><span data-stu-id="19f86-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="19f86-263">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="19f86-263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="19f86-264">1.0</span><span class="sxs-lookup"><span data-stu-id="19f86-264">1.0</span></span>|
|[<span data-ttu-id="19f86-265">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="19f86-265">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="19f86-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="19f86-266">ReadItem</span></span>|
|[<span data-ttu-id="19f86-267">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="19f86-267">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="19f86-268">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="19f86-268">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="19f86-269">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="19f86-269">Returns:</span></span>

<span data-ttu-id="19f86-270">Тип: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.7)</span><span class="sxs-lookup"><span data-stu-id="19f86-270">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.7)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="19f86-271">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="19f86-271">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="19f86-272">Преобразовывает идентификатор элемента в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="19f86-272">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="19f86-273">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="19f86-273">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="19f86-p108">Формат идентификаторов, извлекаемых через EWS или свойство `itemId`, отличается от формата API REST (таких как [API Почты Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)). Метод `convertToRestId` преобразовывает идентификатор в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="19f86-p108">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="19f86-276">Параметры</span><span class="sxs-lookup"><span data-stu-id="19f86-276">Parameters</span></span>

|<span data-ttu-id="19f86-277">Имя</span><span class="sxs-lookup"><span data-stu-id="19f86-277">Name</span></span>| <span data-ttu-id="19f86-278">Тип</span><span class="sxs-lookup"><span data-stu-id="19f86-278">Type</span></span>| <span data-ttu-id="19f86-279">Описание</span><span class="sxs-lookup"><span data-stu-id="19f86-279">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="19f86-280">String</span><span class="sxs-lookup"><span data-stu-id="19f86-280">String</span></span>|<span data-ttu-id="19f86-281">Идентификатор элемента в формате EWS</span><span class="sxs-lookup"><span data-stu-id="19f86-281">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="19f86-282">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="19f86-282">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.7)|<span data-ttu-id="19f86-283">Значение, определяющее версию REST API для Outlook, с которой будет использоваться преобразованный идентификатор.</span><span class="sxs-lookup"><span data-stu-id="19f86-283">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="19f86-284">Requirements</span><span class="sxs-lookup"><span data-stu-id="19f86-284">Requirements</span></span>

|<span data-ttu-id="19f86-285">Требование</span><span class="sxs-lookup"><span data-stu-id="19f86-285">Requirement</span></span>| <span data-ttu-id="19f86-286">Значение</span><span class="sxs-lookup"><span data-stu-id="19f86-286">Value</span></span>|
|---|---|
|[<span data-ttu-id="19f86-287">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="19f86-287">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="19f86-288">1.3</span><span class="sxs-lookup"><span data-stu-id="19f86-288">1.3</span></span>|
|[<span data-ttu-id="19f86-289">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="19f86-289">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="19f86-290">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="19f86-290">Restricted</span></span>|
|[<span data-ttu-id="19f86-291">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="19f86-291">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="19f86-292">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="19f86-292">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="19f86-293">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="19f86-293">Returns:</span></span>

<span data-ttu-id="19f86-294">Тип: String</span><span class="sxs-lookup"><span data-stu-id="19f86-294">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="19f86-295">Пример</span><span class="sxs-lookup"><span data-stu-id="19f86-295">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="19f86-296">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="19f86-296">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="19f86-297">Получает объект Date из словаря, содержащего сведения о времени.</span><span class="sxs-lookup"><span data-stu-id="19f86-297">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="19f86-298">Метод `convertToUtcClientTime` преобразует словарь, содержащий локальную дату и время, в объект Date с правильными значениями локальной даты и времени.</span><span class="sxs-lookup"><span data-stu-id="19f86-298">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="19f86-299">Параметры</span><span class="sxs-lookup"><span data-stu-id="19f86-299">Parameters</span></span>

|<span data-ttu-id="19f86-300">Имя</span><span class="sxs-lookup"><span data-stu-id="19f86-300">Name</span></span>| <span data-ttu-id="19f86-301">Тип</span><span class="sxs-lookup"><span data-stu-id="19f86-301">Type</span></span>| <span data-ttu-id="19f86-302">Описание</span><span class="sxs-lookup"><span data-stu-id="19f86-302">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="19f86-303">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="19f86-303">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.7)|<span data-ttu-id="19f86-304">Значение локального времени для преобразования.</span><span class="sxs-lookup"><span data-stu-id="19f86-304">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="19f86-305">Requirements</span><span class="sxs-lookup"><span data-stu-id="19f86-305">Requirements</span></span>

|<span data-ttu-id="19f86-306">Требование</span><span class="sxs-lookup"><span data-stu-id="19f86-306">Requirement</span></span>| <span data-ttu-id="19f86-307">Значение</span><span class="sxs-lookup"><span data-stu-id="19f86-307">Value</span></span>|
|---|---|
|[<span data-ttu-id="19f86-308">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="19f86-308">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="19f86-309">1.0</span><span class="sxs-lookup"><span data-stu-id="19f86-309">1.0</span></span>|
|[<span data-ttu-id="19f86-310">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="19f86-310">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="19f86-311">ReadItem</span><span class="sxs-lookup"><span data-stu-id="19f86-311">ReadItem</span></span>|
|[<span data-ttu-id="19f86-312">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="19f86-312">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="19f86-313">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="19f86-313">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="19f86-314">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="19f86-314">Returns:</span></span>

<span data-ttu-id="19f86-315">Объект Date со временем в формате UTC.</span><span class="sxs-lookup"><span data-stu-id="19f86-315">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="19f86-316">Тип: Date</span><span class="sxs-lookup"><span data-stu-id="19f86-316">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="19f86-317">Пример</span><span class="sxs-lookup"><span data-stu-id="19f86-317">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="19f86-318">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="19f86-318">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="19f86-319">Отображает имеющуюся встречу из календаря.</span><span class="sxs-lookup"><span data-stu-id="19f86-319">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="19f86-320">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="19f86-320">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="19f86-321">Метод `displayAppointmentForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее сведения календаря о существующей встрече.</span><span class="sxs-lookup"><span data-stu-id="19f86-321">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="19f86-p109">В Outlook для Mac с помощью этого метода можно отобразить одну встречу, которая не является частью повторяющегося ряда, или основную встречу такого ряда, но не экземпляр из него, так как в Outlook для Mac невозможно получить доступ к свойствам экземпляра повторяющегося ряда (в том числе к идентификатору элемента).</span><span class="sxs-lookup"><span data-stu-id="19f86-p109">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="19f86-324">В Outlook в Интернете этот метод открывает указанную форму, только если текст формы содержит символы размером не более 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="19f86-324">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="19f86-325">Если указанный идентификатор элемента не определяет существующую встречу, на клиентском компьютере или устройстве открывается пустая страница, и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="19f86-325">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="19f86-326">Параметры</span><span class="sxs-lookup"><span data-stu-id="19f86-326">Parameters</span></span>

|<span data-ttu-id="19f86-327">Имя</span><span class="sxs-lookup"><span data-stu-id="19f86-327">Name</span></span>| <span data-ttu-id="19f86-328">Тип</span><span class="sxs-lookup"><span data-stu-id="19f86-328">Type</span></span>| <span data-ttu-id="19f86-329">Описание</span><span class="sxs-lookup"><span data-stu-id="19f86-329">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="19f86-330">String</span><span class="sxs-lookup"><span data-stu-id="19f86-330">String</span></span>|<span data-ttu-id="19f86-331">Идентификатор веб-служб Exchange для существующей встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="19f86-331">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="19f86-332">Requirements</span><span class="sxs-lookup"><span data-stu-id="19f86-332">Requirements</span></span>

|<span data-ttu-id="19f86-333">Требование</span><span class="sxs-lookup"><span data-stu-id="19f86-333">Requirement</span></span>| <span data-ttu-id="19f86-334">Значение</span><span class="sxs-lookup"><span data-stu-id="19f86-334">Value</span></span>|
|---|---|
|[<span data-ttu-id="19f86-335">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="19f86-335">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="19f86-336">1.0</span><span class="sxs-lookup"><span data-stu-id="19f86-336">1.0</span></span>|
|[<span data-ttu-id="19f86-337">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="19f86-337">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="19f86-338">ReadItem</span><span class="sxs-lookup"><span data-stu-id="19f86-338">ReadItem</span></span>|
|[<span data-ttu-id="19f86-339">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="19f86-339">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="19f86-340">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="19f86-340">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="19f86-341">Пример</span><span class="sxs-lookup"><span data-stu-id="19f86-341">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="19f86-342">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="19f86-342">displayMessageForm(itemId)</span></span>

<span data-ttu-id="19f86-343">Отображает имеющееся сообщение.</span><span class="sxs-lookup"><span data-stu-id="19f86-343">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="19f86-344">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="19f86-344">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="19f86-345">Метод `displayMessageForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее существующее сообщение.</span><span class="sxs-lookup"><span data-stu-id="19f86-345">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="19f86-346">В Outlook в Интернете этот метод открывает указанную форму, только если текст формы содержит символы размером не более 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="19f86-346">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="19f86-347">Если указанный идентификатор элемента не определяет существующее сообщение, окно на клиентском компьютере не открывается и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="19f86-347">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="19f86-p110">Не используйте `displayMessageForm` с параметром `itemId`, который представляет собой встречу. Используйте метод `displayAppointmentForm`, чтобы отобразить сведения о существующей встрече, а метод `displayNewAppointmentForm` — для отображения формы создания встречи.</span><span class="sxs-lookup"><span data-stu-id="19f86-p110">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="19f86-350">Параметры</span><span class="sxs-lookup"><span data-stu-id="19f86-350">Parameters</span></span>

|<span data-ttu-id="19f86-351">Имя</span><span class="sxs-lookup"><span data-stu-id="19f86-351">Name</span></span>| <span data-ttu-id="19f86-352">Тип</span><span class="sxs-lookup"><span data-stu-id="19f86-352">Type</span></span>| <span data-ttu-id="19f86-353">Описание</span><span class="sxs-lookup"><span data-stu-id="19f86-353">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="19f86-354">String</span><span class="sxs-lookup"><span data-stu-id="19f86-354">String</span></span>|<span data-ttu-id="19f86-355">Идентификатор веб-служб Exchange для существующего сообщения.</span><span class="sxs-lookup"><span data-stu-id="19f86-355">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="19f86-356">Требования</span><span class="sxs-lookup"><span data-stu-id="19f86-356">Requirements</span></span>

|<span data-ttu-id="19f86-357">Требование</span><span class="sxs-lookup"><span data-stu-id="19f86-357">Requirement</span></span>| <span data-ttu-id="19f86-358">Значение</span><span class="sxs-lookup"><span data-stu-id="19f86-358">Value</span></span>|
|---|---|
|[<span data-ttu-id="19f86-359">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="19f86-359">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="19f86-360">1.0</span><span class="sxs-lookup"><span data-stu-id="19f86-360">1.0</span></span>|
|[<span data-ttu-id="19f86-361">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="19f86-361">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="19f86-362">ReadItem</span><span class="sxs-lookup"><span data-stu-id="19f86-362">ReadItem</span></span>|
|[<span data-ttu-id="19f86-363">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="19f86-363">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="19f86-364">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="19f86-364">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="19f86-365">Пример</span><span class="sxs-lookup"><span data-stu-id="19f86-365">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="19f86-366">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="19f86-366">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="19f86-367">Отображает форму для создания новой встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="19f86-367">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="19f86-368">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="19f86-368">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="19f86-p111">Метод `displayNewAppointmentForm` открывает форму, в которой пользователь может создать встречу или собрание. Если параметры заданы, поля формы встречи автоматически заполняются их содержимым.</span><span class="sxs-lookup"><span data-stu-id="19f86-p111">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="19f86-p112">В Outlook в Интернете и на мобильных устройствах этот метод всегда отображает форму с полем участников. Если вы не укажете участников в качестве входных аргументов, метод отображает форму с кнопкой **Сохранить**. Если вы укажете участников, форма будет включать участников и кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="19f86-p112">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="19f86-p113">Если вы укажете участников или ресурсы с помощью параметра `requiredAttendees`, `optionalAttendees` или `resources` в клиенте Outlook с расширенными возможностями и Outlook RT, этот метод отобразит форму собрания с кнопкой **Отправить**. Если не указать получателей, этот метод отобразит форму встречи с кнопкой **Сохранить и закрыть**.</span><span class="sxs-lookup"><span data-stu-id="19f86-p113">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="19f86-376">Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.</span><span class="sxs-lookup"><span data-stu-id="19f86-376">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="19f86-377">Параметры</span><span class="sxs-lookup"><span data-stu-id="19f86-377">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="19f86-378">Все параметры являются необязательными.</span><span class="sxs-lookup"><span data-stu-id="19f86-378">All parameters are optional.</span></span>

|<span data-ttu-id="19f86-379">Имя</span><span class="sxs-lookup"><span data-stu-id="19f86-379">Name</span></span>| <span data-ttu-id="19f86-380">Тип</span><span class="sxs-lookup"><span data-stu-id="19f86-380">Type</span></span>| <span data-ttu-id="19f86-381">Описание</span><span class="sxs-lookup"><span data-stu-id="19f86-381">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="19f86-382">Object</span><span class="sxs-lookup"><span data-stu-id="19f86-382">Object</span></span> | <span data-ttu-id="19f86-383">Словарь параметров, описывающий новую встречу.</span><span class="sxs-lookup"><span data-stu-id="19f86-383">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="19f86-384">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span><span class="sxs-lookup"><span data-stu-id="19f86-384">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span></span> | <span data-ttu-id="19f86-p114">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из обязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="19f86-p114">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="19f86-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span><span class="sxs-lookup"><span data-stu-id="19f86-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span></span> | <span data-ttu-id="19f86-p115">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из необязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="19f86-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="19f86-390">Date</span><span class="sxs-lookup"><span data-stu-id="19f86-390">Date</span></span> | <span data-ttu-id="19f86-391">Объект `Date`, указывающий дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="19f86-391">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="19f86-392">Date</span><span class="sxs-lookup"><span data-stu-id="19f86-392">Date</span></span> | <span data-ttu-id="19f86-393">Объект `Date`, указывающий дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="19f86-393">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="19f86-394">String</span><span class="sxs-lookup"><span data-stu-id="19f86-394">String</span></span> | <span data-ttu-id="19f86-p116">Строка со сведениями о месте встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="19f86-p116">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="19f86-397">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="19f86-397">Array.&lt;String&gt;</span></span> | <span data-ttu-id="19f86-p117">Массив строк, содержащий необходимые для встречи ресурсы. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="19f86-p117">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="19f86-400">String</span><span class="sxs-lookup"><span data-stu-id="19f86-400">String</span></span> | <span data-ttu-id="19f86-p118">Строка с темой встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="19f86-p118">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="19f86-403">String</span><span class="sxs-lookup"><span data-stu-id="19f86-403">String</span></span> | <span data-ttu-id="19f86-p119">Текст сообщения о встрече. Максимальный размер содержимого сообщения — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="19f86-p119">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="19f86-406">Требования</span><span class="sxs-lookup"><span data-stu-id="19f86-406">Requirements</span></span>

|<span data-ttu-id="19f86-407">Требование</span><span class="sxs-lookup"><span data-stu-id="19f86-407">Requirement</span></span>| <span data-ttu-id="19f86-408">Значение</span><span class="sxs-lookup"><span data-stu-id="19f86-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="19f86-409">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="19f86-409">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="19f86-410">1.0</span><span class="sxs-lookup"><span data-stu-id="19f86-410">1.0</span></span>|
|[<span data-ttu-id="19f86-411">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="19f86-411">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="19f86-412">ReadItem</span><span class="sxs-lookup"><span data-stu-id="19f86-412">ReadItem</span></span>|
|[<span data-ttu-id="19f86-413">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="19f86-413">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="19f86-414">Чтение</span><span class="sxs-lookup"><span data-stu-id="19f86-414">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="19f86-415">Пример</span><span class="sxs-lookup"><span data-stu-id="19f86-415">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="19f86-416">Дисплайневмессажеформ (Parameters)</span><span class="sxs-lookup"><span data-stu-id="19f86-416">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="19f86-417">Отображает форму для создания нового сообщения.</span><span class="sxs-lookup"><span data-stu-id="19f86-417">Displays a form for creating a new message.</span></span>

<span data-ttu-id="19f86-418">`displayNewMessageForm` Метод открывает форму, которая позволяет пользователю создать новое сообщение.</span><span class="sxs-lookup"><span data-stu-id="19f86-418">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="19f86-419">Если указаны параметры, поля формы сообщения автоматически заполняются содержимым параметров.</span><span class="sxs-lookup"><span data-stu-id="19f86-419">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="19f86-420">Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.</span><span class="sxs-lookup"><span data-stu-id="19f86-420">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="19f86-421">Параметры</span><span class="sxs-lookup"><span data-stu-id="19f86-421">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="19f86-422">Все параметры являются необязательными.</span><span class="sxs-lookup"><span data-stu-id="19f86-422">All parameters are optional.</span></span>

|<span data-ttu-id="19f86-423">Имя</span><span class="sxs-lookup"><span data-stu-id="19f86-423">Name</span></span>| <span data-ttu-id="19f86-424">Тип</span><span class="sxs-lookup"><span data-stu-id="19f86-424">Type</span></span>| <span data-ttu-id="19f86-425">Описание</span><span class="sxs-lookup"><span data-stu-id="19f86-425">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="19f86-426">Object</span><span class="sxs-lookup"><span data-stu-id="19f86-426">Object</span></span> | <span data-ttu-id="19f86-427">Словарь параметров, описывающих новое сообщение.</span><span class="sxs-lookup"><span data-stu-id="19f86-427">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="19f86-428">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span><span class="sxs-lookup"><span data-stu-id="19f86-428">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span></span> | <span data-ttu-id="19f86-429">Массив строк, содержащий адреса электронной почты или массив, содержащий `EmailAddressDetails` объект для каждого из получателей в строке "Кому".</span><span class="sxs-lookup"><span data-stu-id="19f86-429">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="19f86-430">Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="19f86-430">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="19f86-431">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span><span class="sxs-lookup"><span data-stu-id="19f86-431">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span></span> | <span data-ttu-id="19f86-432">Массив строк, содержащий адреса электронной почты или массив, содержащий `EmailAddressDetails` объект для каждого получателя в строке "копия".</span><span class="sxs-lookup"><span data-stu-id="19f86-432">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="19f86-433">Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="19f86-433">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="19f86-434">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span><span class="sxs-lookup"><span data-stu-id="19f86-434">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.7)&gt;</span></span> | <span data-ttu-id="19f86-435">Массив строк, содержащий адреса электронной почты или массив, содержащий `EmailAddressDetails` объект для каждого из получателей, указанных в строке "СК".</span><span class="sxs-lookup"><span data-stu-id="19f86-435">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="19f86-436">Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="19f86-436">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="19f86-437">String</span><span class="sxs-lookup"><span data-stu-id="19f86-437">String</span></span> | <span data-ttu-id="19f86-438">Строка, содержащая тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="19f86-438">A string containing the subject of the message.</span></span> <span data-ttu-id="19f86-439">Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="19f86-439">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="19f86-440">String</span><span class="sxs-lookup"><span data-stu-id="19f86-440">String</span></span> | <span data-ttu-id="19f86-441">Текст сообщения в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="19f86-441">The HTML body of the message.</span></span> <span data-ttu-id="19f86-442">Максимальный размер содержимого сообщения — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="19f86-442">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="19f86-443">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="19f86-443">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="19f86-444">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="19f86-444">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="19f86-445">String</span><span class="sxs-lookup"><span data-stu-id="19f86-445">String</span></span> | <span data-ttu-id="19f86-p126">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="19f86-p126">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="19f86-448">Строка</span><span class="sxs-lookup"><span data-stu-id="19f86-448">String</span></span> | <span data-ttu-id="19f86-449">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="19f86-449">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="19f86-450">String</span><span class="sxs-lookup"><span data-stu-id="19f86-450">String</span></span> | <span data-ttu-id="19f86-p127">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="19f86-p127">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="19f86-453">Логический</span><span class="sxs-lookup"><span data-stu-id="19f86-453">Boolean</span></span> | <span data-ttu-id="19f86-p128">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="19f86-p128">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="19f86-456">Строка</span><span class="sxs-lookup"><span data-stu-id="19f86-456">String</span></span> | <span data-ttu-id="19f86-457">Используется, только если свойству `type` присвоено значение `item`.</span><span class="sxs-lookup"><span data-stu-id="19f86-457">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="19f86-458">Идентификатор элемента EWS существующего сообщения электронной почты, которое необходимо присоединить к новому сообщению.</span><span class="sxs-lookup"><span data-stu-id="19f86-458">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="19f86-459">Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="19f86-459">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="19f86-460">Requirements</span><span class="sxs-lookup"><span data-stu-id="19f86-460">Requirements</span></span>

|<span data-ttu-id="19f86-461">Требование</span><span class="sxs-lookup"><span data-stu-id="19f86-461">Requirement</span></span>| <span data-ttu-id="19f86-462">Значение</span><span class="sxs-lookup"><span data-stu-id="19f86-462">Value</span></span>|
|---|---|
|[<span data-ttu-id="19f86-463">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="19f86-463">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="19f86-464">1.6</span><span class="sxs-lookup"><span data-stu-id="19f86-464">1.6</span></span> |
|[<span data-ttu-id="19f86-465">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="19f86-465">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="19f86-466">ReadItem</span><span class="sxs-lookup"><span data-stu-id="19f86-466">ReadItem</span></span>|
|[<span data-ttu-id="19f86-467">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="19f86-467">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="19f86-468">Чтение</span><span class="sxs-lookup"><span data-stu-id="19f86-468">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="19f86-469">Пример</span><span class="sxs-lookup"><span data-stu-id="19f86-469">Example</span></span>

```js
Office.context.mailbox.displayNewMessageForm(
  {
    // Copy the To line from current item.
    toRecipients: Office.context.mailbox.item.to
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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="19f86-470">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="19f86-470">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="19f86-471">Возвращает строку, содержащую маркер, который используется для вызова интерфейсов REST API или веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="19f86-471">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="19f86-p130">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="19f86-p130">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="19f86-474">Рекомендуем сделать так, чтобы по мере возможности надстройки использовали интерфейсы REST API, а не веб-службы Exchange.</span><span class="sxs-lookup"><span data-stu-id="19f86-474">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="19f86-475">Для вызова метода `getCallbackTokenAsync` в режиме чтения требуется минимальный уровень разрешения **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="19f86-475">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="19f86-476">Для вызова `getCallbackTokenAsync` в режиме создания сообщения требуется сохранить элемент.</span><span class="sxs-lookup"><span data-stu-id="19f86-476">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="19f86-477">Для метода [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) требуется минимальный уровень разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="19f86-477">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

<span data-ttu-id="19f86-478">**Маркеры REST**</span><span class="sxs-lookup"><span data-stu-id="19f86-478">**REST Tokens**</span></span>

<span data-ttu-id="19f86-p132">Если запрашивается маркер REST (`options.isRest = true`), полученный маркер не подойдет для проверки подлинности при вызовах веб-служб Exchange. Область действия маркера будет ограничена доступом только для чтения к текущему элементу и его вложениям, если в манифесте надстройки не указано разрешение [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission). Если указано разрешение `ReadWriteMailbox`, полученный маркер предоставит доступ на чтение и запись к почте, календарю и контактам, включая возможность отправки почты.</span><span class="sxs-lookup"><span data-stu-id="19f86-p132">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="19f86-482">С помощью свойства `restUrl` надстройка должна определить правильный URL-адрес для вызовов REST API.</span><span class="sxs-lookup"><span data-stu-id="19f86-482">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="19f86-483">**Маркеры EWS**</span><span class="sxs-lookup"><span data-stu-id="19f86-483">**EWS Tokens**</span></span>

<span data-ttu-id="19f86-p133">Если запрашивается маркер EWS (`options.isRest = false`), полученный маркер не подойдет для проверки подлинности при вызовах REST API. Область действия маркера будет ограничена доступом к текущему элементу.</span><span class="sxs-lookup"><span data-stu-id="19f86-p133">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="19f86-486">С помощью свойства `ewsUrl` надстройка должна определить правильный URL-адрес для вызовов EWS.</span><span class="sxs-lookup"><span data-stu-id="19f86-486">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

<span data-ttu-id="19f86-487">Вы можете передать сторонней системе маркер и идентификатор вложения или элемента.</span><span class="sxs-lookup"><span data-stu-id="19f86-487">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="19f86-488">Сторонняя система использует этот маркер как маркер авторизации, чтобы вызвать операцию [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) или [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) веб-служб Exchange для возврата вложения или элемента.</span><span class="sxs-lookup"><span data-stu-id="19f86-488">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="19f86-489">Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="19f86-489">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="parameters"></a><span data-ttu-id="19f86-490">Параметры</span><span class="sxs-lookup"><span data-stu-id="19f86-490">Parameters</span></span>

|<span data-ttu-id="19f86-491">Имя</span><span class="sxs-lookup"><span data-stu-id="19f86-491">Name</span></span>| <span data-ttu-id="19f86-492">Тип</span><span class="sxs-lookup"><span data-stu-id="19f86-492">Type</span></span>| <span data-ttu-id="19f86-493">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="19f86-493">Attributes</span></span>| <span data-ttu-id="19f86-494">Описание</span><span class="sxs-lookup"><span data-stu-id="19f86-494">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="19f86-495">Объект</span><span class="sxs-lookup"><span data-stu-id="19f86-495">Object</span></span> | <span data-ttu-id="19f86-496">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="19f86-496">&lt;optional&gt;</span></span> | <span data-ttu-id="19f86-497">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="19f86-497">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="19f86-498">Boolean</span><span class="sxs-lookup"><span data-stu-id="19f86-498">Boolean</span></span> |  <span data-ttu-id="19f86-499">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="19f86-499">&lt;optional&gt;</span></span> | <span data-ttu-id="19f86-p135">Определяет, будет ли предоставленный маркер использоваться для интерфейсов REST API Outlook или веб-служб Exchange. Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="19f86-p135">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="19f86-502">Объект</span><span class="sxs-lookup"><span data-stu-id="19f86-502">Object</span></span> |  <span data-ttu-id="19f86-503">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="19f86-503">&lt;optional&gt;</span></span> | <span data-ttu-id="19f86-504">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="19f86-504">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="19f86-505">функция</span><span class="sxs-lookup"><span data-stu-id="19f86-505">function</span></span>||<span data-ttu-id="19f86-506">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="19f86-506">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="19f86-507">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="19f86-507">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="19f86-508">При наличии ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут предоставлять дополнительные сведения.</span><span class="sxs-lookup"><span data-stu-id="19f86-508">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="19f86-509">Ошибки</span><span class="sxs-lookup"><span data-stu-id="19f86-509">Errors</span></span>

|<span data-ttu-id="19f86-510">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="19f86-510">Error code</span></span>|<span data-ttu-id="19f86-511">Описание</span><span class="sxs-lookup"><span data-stu-id="19f86-511">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="19f86-512">Не удалось выполнить запрос.</span><span class="sxs-lookup"><span data-stu-id="19f86-512">The request has failed.</span></span> <span data-ttu-id="19f86-513">Просмотрите объект диагностики для кода ошибки HTTP.</span><span class="sxs-lookup"><span data-stu-id="19f86-513">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="19f86-514">Сервер Exchange Server вернул ошибку.</span><span class="sxs-lookup"><span data-stu-id="19f86-514">The Exchange server returned an error.</span></span> <span data-ttu-id="19f86-515">Для получения дополнительных сведений просмотрите объект диагностики.</span><span class="sxs-lookup"><span data-stu-id="19f86-515">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="19f86-516">Пользователь отключен от сети.</span><span class="sxs-lookup"><span data-stu-id="19f86-516">The user is no longer connected to the network.</span></span> <span data-ttu-id="19f86-517">Проверьте сетевое подключение и повторите попытку.</span><span class="sxs-lookup"><span data-stu-id="19f86-517">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="19f86-518">Требования</span><span class="sxs-lookup"><span data-stu-id="19f86-518">Requirements</span></span>

|<span data-ttu-id="19f86-519">Требование</span><span class="sxs-lookup"><span data-stu-id="19f86-519">Requirement</span></span>| <span data-ttu-id="19f86-520">Значение</span><span class="sxs-lookup"><span data-stu-id="19f86-520">Value</span></span>|
|---|---|
|[<span data-ttu-id="19f86-521">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="19f86-521">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="19f86-522">1.5</span><span class="sxs-lookup"><span data-stu-id="19f86-522">1.5</span></span> |
|[<span data-ttu-id="19f86-523">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="19f86-523">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="19f86-524">ReadItem</span><span class="sxs-lookup"><span data-stu-id="19f86-524">ReadItem</span></span>|
|[<span data-ttu-id="19f86-525">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="19f86-525">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="19f86-526">Создание и чтение</span><span class="sxs-lookup"><span data-stu-id="19f86-526">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="19f86-527">Пример</span><span class="sxs-lookup"><span data-stu-id="19f86-527">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="19f86-528">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="19f86-528">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="19f86-529">Получает строку, содержащую маркер, используемый для получения вложения или элемента с Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="19f86-529">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="19f86-p139">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="19f86-p139">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="19f86-532">Вы можете передать сторонней системе маркер и идентификатор вложения или элемента.</span><span class="sxs-lookup"><span data-stu-id="19f86-532">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="19f86-533">Сторонняя система использует этот маркер как маркер авторизации, чтобы вызвать операцию [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) или [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) веб-служб Exchange для возврата вложения или элемента.</span><span class="sxs-lookup"><span data-stu-id="19f86-533">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="19f86-534">Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="19f86-534">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="19f86-535">Для вызова метода `getCallbackTokenAsync` в режиме чтения требуется минимальный уровень разрешения **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="19f86-535">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="19f86-536">Для вызова `getCallbackTokenAsync` в режиме создания сообщения требуется сохранить элемент.</span><span class="sxs-lookup"><span data-stu-id="19f86-536">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="19f86-537">Для метода [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) требуется минимальный уровень разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="19f86-537">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="19f86-538">Параметры</span><span class="sxs-lookup"><span data-stu-id="19f86-538">Parameters</span></span>

|<span data-ttu-id="19f86-539">Имя</span><span class="sxs-lookup"><span data-stu-id="19f86-539">Name</span></span>| <span data-ttu-id="19f86-540">Тип</span><span class="sxs-lookup"><span data-stu-id="19f86-540">Type</span></span>| <span data-ttu-id="19f86-541">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="19f86-541">Attributes</span></span>| <span data-ttu-id="19f86-542">Описание</span><span class="sxs-lookup"><span data-stu-id="19f86-542">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="19f86-543">функция</span><span class="sxs-lookup"><span data-stu-id="19f86-543">function</span></span>||<span data-ttu-id="19f86-544">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="19f86-544">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="19f86-545">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="19f86-545">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="19f86-546">При наличии ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут предоставлять дополнительные сведения.</span><span class="sxs-lookup"><span data-stu-id="19f86-546">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="19f86-547">Объект</span><span class="sxs-lookup"><span data-stu-id="19f86-547">Object</span></span>| <span data-ttu-id="19f86-548">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="19f86-548">&lt;optional&gt;</span></span>|<span data-ttu-id="19f86-549">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="19f86-549">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="19f86-550">Ошибки</span><span class="sxs-lookup"><span data-stu-id="19f86-550">Errors</span></span>

|<span data-ttu-id="19f86-551">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="19f86-551">Error code</span></span>|<span data-ttu-id="19f86-552">Описание</span><span class="sxs-lookup"><span data-stu-id="19f86-552">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="19f86-553">Не удалось выполнить запрос.</span><span class="sxs-lookup"><span data-stu-id="19f86-553">The request has failed.</span></span> <span data-ttu-id="19f86-554">Просмотрите объект диагностики для кода ошибки HTTP.</span><span class="sxs-lookup"><span data-stu-id="19f86-554">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="19f86-555">Сервер Exchange Server вернул ошибку.</span><span class="sxs-lookup"><span data-stu-id="19f86-555">The Exchange server returned an error.</span></span> <span data-ttu-id="19f86-556">Для получения дополнительных сведений просмотрите объект диагностики.</span><span class="sxs-lookup"><span data-stu-id="19f86-556">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="19f86-557">Пользователь отключен от сети.</span><span class="sxs-lookup"><span data-stu-id="19f86-557">The user is no longer connected to the network.</span></span> <span data-ttu-id="19f86-558">Проверьте сетевое подключение и повторите попытку.</span><span class="sxs-lookup"><span data-stu-id="19f86-558">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="19f86-559">Требования</span><span class="sxs-lookup"><span data-stu-id="19f86-559">Requirements</span></span>

|<span data-ttu-id="19f86-560">Требование</span><span class="sxs-lookup"><span data-stu-id="19f86-560">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="19f86-561">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="19f86-561">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="19f86-562">1.0</span><span class="sxs-lookup"><span data-stu-id="19f86-562">1.0</span></span> | <span data-ttu-id="19f86-563">1.3</span><span class="sxs-lookup"><span data-stu-id="19f86-563">1.3</span></span> |
|[<span data-ttu-id="19f86-564">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="19f86-564">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="19f86-565">ReadItem</span><span class="sxs-lookup"><span data-stu-id="19f86-565">ReadItem</span></span> | <span data-ttu-id="19f86-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="19f86-566">ReadItem</span></span> |
|[<span data-ttu-id="19f86-567">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="19f86-567">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="19f86-568">Чтение</span><span class="sxs-lookup"><span data-stu-id="19f86-568">Read</span></span> | <span data-ttu-id="19f86-569">Создание</span><span class="sxs-lookup"><span data-stu-id="19f86-569">Compose</span></span> |

##### <a name="example"></a><span data-ttu-id="19f86-570">Пример</span><span class="sxs-lookup"><span data-stu-id="19f86-570">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="19f86-571">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="19f86-571">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="19f86-572">Получает маркер, идентифицирующий пользователя и надстройку Office.</span><span class="sxs-lookup"><span data-stu-id="19f86-572">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="19f86-573">Метод `getUserIdentityTokenAsync` возвращает маркер, который можно использовать для идентификации, а также [проверки подлинности надстройки и пользователя в сторонней системе](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="19f86-573">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="19f86-574">Параметры</span><span class="sxs-lookup"><span data-stu-id="19f86-574">Parameters</span></span>

|<span data-ttu-id="19f86-575">Имя</span><span class="sxs-lookup"><span data-stu-id="19f86-575">Name</span></span>| <span data-ttu-id="19f86-576">Тип</span><span class="sxs-lookup"><span data-stu-id="19f86-576">Type</span></span>| <span data-ttu-id="19f86-577">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="19f86-577">Attributes</span></span>| <span data-ttu-id="19f86-578">Описание</span><span class="sxs-lookup"><span data-stu-id="19f86-578">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="19f86-579">функция</span><span class="sxs-lookup"><span data-stu-id="19f86-579">function</span></span>||<span data-ttu-id="19f86-580">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="19f86-580">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="19f86-581">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="19f86-581">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="19f86-582">При наличии ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут предоставлять дополнительные сведения.</span><span class="sxs-lookup"><span data-stu-id="19f86-582">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="19f86-583">Объект</span><span class="sxs-lookup"><span data-stu-id="19f86-583">Object</span></span>| <span data-ttu-id="19f86-584">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="19f86-584">&lt;optional&gt;</span></span>|<span data-ttu-id="19f86-585">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="19f86-585">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="19f86-586">Ошибки</span><span class="sxs-lookup"><span data-stu-id="19f86-586">Errors</span></span>

|<span data-ttu-id="19f86-587">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="19f86-587">Error code</span></span>|<span data-ttu-id="19f86-588">Описание</span><span class="sxs-lookup"><span data-stu-id="19f86-588">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="19f86-589">Не удалось выполнить запрос.</span><span class="sxs-lookup"><span data-stu-id="19f86-589">The request has failed.</span></span> <span data-ttu-id="19f86-590">Просмотрите объект диагностики для кода ошибки HTTP.</span><span class="sxs-lookup"><span data-stu-id="19f86-590">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="19f86-591">Сервер Exchange Server вернул ошибку.</span><span class="sxs-lookup"><span data-stu-id="19f86-591">The Exchange server returned an error.</span></span> <span data-ttu-id="19f86-592">Для получения дополнительных сведений просмотрите объект диагностики.</span><span class="sxs-lookup"><span data-stu-id="19f86-592">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="19f86-593">Пользователь отключен от сети.</span><span class="sxs-lookup"><span data-stu-id="19f86-593">The user is no longer connected to the network.</span></span> <span data-ttu-id="19f86-594">Проверьте сетевое подключение и повторите попытку.</span><span class="sxs-lookup"><span data-stu-id="19f86-594">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="19f86-595">Требования</span><span class="sxs-lookup"><span data-stu-id="19f86-595">Requirements</span></span>

|<span data-ttu-id="19f86-596">Требование</span><span class="sxs-lookup"><span data-stu-id="19f86-596">Requirement</span></span>| <span data-ttu-id="19f86-597">Значение</span><span class="sxs-lookup"><span data-stu-id="19f86-597">Value</span></span>|
|---|---|
|[<span data-ttu-id="19f86-598">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="19f86-598">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="19f86-599">1.0</span><span class="sxs-lookup"><span data-stu-id="19f86-599">1.0</span></span>|
|[<span data-ttu-id="19f86-600">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="19f86-600">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="19f86-601">ReadItem</span><span class="sxs-lookup"><span data-stu-id="19f86-601">ReadItem</span></span>|
|[<span data-ttu-id="19f86-602">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="19f86-602">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="19f86-603">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="19f86-603">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="19f86-604">Пример</span><span class="sxs-lookup"><span data-stu-id="19f86-604">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="19f86-605">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="19f86-605">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="19f86-606">Выполняет асинхронный запрос для веб-служб Exchange (EWS) на сервере Exchange Server, на котором размещен почтовый ящик пользователя.</span><span class="sxs-lookup"><span data-stu-id="19f86-606">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="19f86-607">Этот метод не поддерживается в следующих сценариях:</span><span class="sxs-lookup"><span data-stu-id="19f86-607">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="19f86-608">В Outlook для iOS и Android</span><span class="sxs-lookup"><span data-stu-id="19f86-608">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="19f86-609">Если надстройка загружается в почтовый ящик Gmail.</span><span class="sxs-lookup"><span data-stu-id="19f86-609">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="19f86-610">В таких случаях надстройка должна [использовать REST API](/outlook/add-ins/use-rest-api) для доступа к почтовому ящику пользователя.</span><span class="sxs-lookup"><span data-stu-id="19f86-610">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="19f86-611">Метод `makeEwsRequestAsync` отправляет запрос EWS от имени надстройки в Exchange.</span><span class="sxs-lookup"><span data-stu-id="19f86-611">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="19f86-612">Список поддерживаемых операций EWS см. в статье [Вызов веб-служб из надстройки Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support).</span><span class="sxs-lookup"><span data-stu-id="19f86-612">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="19f86-613">С помощью метода `makeEwsRequestAsync` невозможно запрашивать элементы, связанные с папкой.</span><span class="sxs-lookup"><span data-stu-id="19f86-613">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="19f86-614">В запросе XML должна быть указана кодировка UTF-8.</span><span class="sxs-lookup"><span data-stu-id="19f86-614">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="19f86-p149">У вашей надстройки должно быть разрешение **ReadWriteMailbox** для использования метода `makeEwsRequestAsync`. Сведения об использовании разрешения **ReadWriteMailbox** и операций EWS, которые можно вызывать с помощью метода `makeEwsRequestAsync`, см. в статье [Указание разрешений для доступа почтовой надстройки к почтовому ящику пользователя](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="19f86-p149">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="19f86-617">Администратор сервера должен установить значение true для параметра `OAuthAuthentication` в каталоге сервера клиентского доступа EWS, чтобы метод `makeEwsRequestAsync` мог выполнять запросы EWS.</span><span class="sxs-lookup"><span data-stu-id="19f86-617">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="19f86-618">Различия версий</span><span class="sxs-lookup"><span data-stu-id="19f86-618">Version differences</span></span>

<span data-ttu-id="19f86-619">Если вы используете метод `makeEwsRequestAsync` в почтовых приложениях, которые выполняются в Outlook версии более ранней, чем 15.0.4535.1004, указывайте кодировку `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="19f86-619">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="19f86-p150">Значение кодировки не нужно указывать, если почтовое приложение выполняется в Outlook в Интернете. Чтобы определить, выполняется ли приложение в Outlook или Outlook в Интернете, используйте свойство mailbox.diagnostics.hostName. Используемую версию Outlook можно определить с помощью свойства mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="19f86-p150">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="19f86-623">Параметры</span><span class="sxs-lookup"><span data-stu-id="19f86-623">Parameters</span></span>

|<span data-ttu-id="19f86-624">Имя</span><span class="sxs-lookup"><span data-stu-id="19f86-624">Name</span></span>| <span data-ttu-id="19f86-625">Тип</span><span class="sxs-lookup"><span data-stu-id="19f86-625">Type</span></span>| <span data-ttu-id="19f86-626">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="19f86-626">Attributes</span></span>| <span data-ttu-id="19f86-627">Описание</span><span class="sxs-lookup"><span data-stu-id="19f86-627">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="19f86-628">String</span><span class="sxs-lookup"><span data-stu-id="19f86-628">String</span></span>||<span data-ttu-id="19f86-629">Запрос EWS.</span><span class="sxs-lookup"><span data-stu-id="19f86-629">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="19f86-630">function</span><span class="sxs-lookup"><span data-stu-id="19f86-630">function</span></span>||<span data-ttu-id="19f86-631">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="19f86-631">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="19f86-632">Результат XML вызова EWS указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="19f86-632">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="19f86-633">Если размер результата превышает 1 МБ, возвращается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="19f86-633">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="19f86-634">Объект</span><span class="sxs-lookup"><span data-stu-id="19f86-634">Object</span></span>| <span data-ttu-id="19f86-635">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="19f86-635">&lt;optional&gt;</span></span>|<span data-ttu-id="19f86-636">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="19f86-636">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="19f86-637">Требования</span><span class="sxs-lookup"><span data-stu-id="19f86-637">Requirements</span></span>

|<span data-ttu-id="19f86-638">Требование</span><span class="sxs-lookup"><span data-stu-id="19f86-638">Requirement</span></span>| <span data-ttu-id="19f86-639">Значение</span><span class="sxs-lookup"><span data-stu-id="19f86-639">Value</span></span>|
|---|---|
|[<span data-ttu-id="19f86-640">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="19f86-640">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="19f86-641">1.0</span><span class="sxs-lookup"><span data-stu-id="19f86-641">1.0</span></span>|
|[<span data-ttu-id="19f86-642">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="19f86-642">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="19f86-643">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="19f86-643">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="19f86-644">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="19f86-644">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="19f86-645">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="19f86-645">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="19f86-646">Пример</span><span class="sxs-lookup"><span data-stu-id="19f86-646">Example</span></span>

<span data-ttu-id="19f86-647">В приведенном ниже примере вызывается `makeEwsRequestAsync` для получения темы элемента с помощью операции `GetItem`.</span><span class="sxs-lookup"><span data-stu-id="19f86-647">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="19f86-648">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="19f86-648">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="19f86-649">Удаляет обработчиков для поддерживаемого типа события.</span><span class="sxs-lookup"><span data-stu-id="19f86-649">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="19f86-650">В настоящее время единственный поддерживаемый тип события — `Office.EventType.ItemChanged`.</span><span class="sxs-lookup"><span data-stu-id="19f86-650">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="19f86-651">Параметры</span><span class="sxs-lookup"><span data-stu-id="19f86-651">Parameters</span></span>

| <span data-ttu-id="19f86-652">Имя</span><span class="sxs-lookup"><span data-stu-id="19f86-652">Name</span></span> | <span data-ttu-id="19f86-653">Тип</span><span class="sxs-lookup"><span data-stu-id="19f86-653">Type</span></span> | <span data-ttu-id="19f86-654">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="19f86-654">Attributes</span></span> | <span data-ttu-id="19f86-655">Описание</span><span class="sxs-lookup"><span data-stu-id="19f86-655">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="19f86-656">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="19f86-656">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="19f86-657">Событие, которое должно отменить обработчик.</span><span class="sxs-lookup"><span data-stu-id="19f86-657">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="19f86-658">Объект</span><span class="sxs-lookup"><span data-stu-id="19f86-658">Object</span></span> | <span data-ttu-id="19f86-659">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="19f86-659">&lt;optional&gt;</span></span> | <span data-ttu-id="19f86-660">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="19f86-660">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="19f86-661">Object</span><span class="sxs-lookup"><span data-stu-id="19f86-661">Object</span></span> | <span data-ttu-id="19f86-662">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="19f86-662">&lt;optional&gt;</span></span> | <span data-ttu-id="19f86-663">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="19f86-663">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="19f86-664">функция</span><span class="sxs-lookup"><span data-stu-id="19f86-664">function</span></span>| <span data-ttu-id="19f86-665">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="19f86-665">&lt;optional&gt;</span></span>|<span data-ttu-id="19f86-666">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="19f86-666">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="19f86-667">Requirements</span><span class="sxs-lookup"><span data-stu-id="19f86-667">Requirements</span></span>

|<span data-ttu-id="19f86-668">Требование</span><span class="sxs-lookup"><span data-stu-id="19f86-668">Requirement</span></span>| <span data-ttu-id="19f86-669">Значение</span><span class="sxs-lookup"><span data-stu-id="19f86-669">Value</span></span>|
|---|---|
|[<span data-ttu-id="19f86-670">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="19f86-670">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="19f86-671">1.5</span><span class="sxs-lookup"><span data-stu-id="19f86-671">1.5</span></span> |
|[<span data-ttu-id="19f86-672">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="19f86-672">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="19f86-673">ReadItem</span><span class="sxs-lookup"><span data-stu-id="19f86-673">ReadItem</span></span> |
|[<span data-ttu-id="19f86-674">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="19f86-674">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="19f86-675">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="19f86-675">Compose or Read</span></span>|
