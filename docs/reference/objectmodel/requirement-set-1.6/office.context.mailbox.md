---
title: Office. Context. Mailbox — набор обязательных элементов 1,6
description: ''
ms.date: 11/27/2019
localization_priority: Normal
ms.openlocfilehash: 09c3930daf6f26edbc38b01f515ee5b1830ce802
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629695"
---
# <a name="mailbox"></a><span data-ttu-id="aa4d8-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="aa4d8-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="aa4d8-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="aa4d8-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="aa4d8-104">Предоставляет для Microsoft Outlook доступ к объектной модели надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa4d8-105">Требования</span><span class="sxs-lookup"><span data-stu-id="aa4d8-105">Requirements</span></span>

|<span data-ttu-id="aa4d8-106">Требование</span><span class="sxs-lookup"><span data-stu-id="aa4d8-106">Requirement</span></span>| <span data-ttu-id="aa4d8-107">Значение</span><span class="sxs-lookup"><span data-stu-id="aa4d8-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa4d8-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa4d8-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa4d8-109">1.0</span><span class="sxs-lookup"><span data-stu-id="aa4d8-109">1.0</span></span>|
|[<span data-ttu-id="aa4d8-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa4d8-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa4d8-111">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="aa4d8-111">Restricted</span></span>|
|[<span data-ttu-id="aa4d8-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa4d8-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa4d8-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="aa4d8-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="aa4d8-114">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="aa4d8-114">Members and methods</span></span>

| <span data-ttu-id="aa4d8-115">Элемент</span><span class="sxs-lookup"><span data-stu-id="aa4d8-115">Member</span></span> | <span data-ttu-id="aa4d8-116">Тип</span><span class="sxs-lookup"><span data-stu-id="aa4d8-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="aa4d8-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="aa4d8-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="aa4d8-118">Элемент</span><span class="sxs-lookup"><span data-stu-id="aa4d8-118">Member</span></span> |
| [<span data-ttu-id="aa4d8-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="aa4d8-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="aa4d8-120">Элемент</span><span class="sxs-lookup"><span data-stu-id="aa4d8-120">Member</span></span> |
| [<span data-ttu-id="aa4d8-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="aa4d8-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="aa4d8-122">Метод</span><span class="sxs-lookup"><span data-stu-id="aa4d8-122">Method</span></span> |
| [<span data-ttu-id="aa4d8-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="aa4d8-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="aa4d8-124">Метод</span><span class="sxs-lookup"><span data-stu-id="aa4d8-124">Method</span></span> |
| [<span data-ttu-id="aa4d8-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="aa4d8-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="aa4d8-126">Метод</span><span class="sxs-lookup"><span data-stu-id="aa4d8-126">Method</span></span> |
| [<span data-ttu-id="aa4d8-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="aa4d8-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="aa4d8-128">Метод</span><span class="sxs-lookup"><span data-stu-id="aa4d8-128">Method</span></span> |
| [<span data-ttu-id="aa4d8-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="aa4d8-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="aa4d8-130">Метод</span><span class="sxs-lookup"><span data-stu-id="aa4d8-130">Method</span></span> |
| [<span data-ttu-id="aa4d8-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="aa4d8-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="aa4d8-132">Метод</span><span class="sxs-lookup"><span data-stu-id="aa4d8-132">Method</span></span> |
| [<span data-ttu-id="aa4d8-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="aa4d8-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="aa4d8-134">Метод</span><span class="sxs-lookup"><span data-stu-id="aa4d8-134">Method</span></span> |
| [<span data-ttu-id="aa4d8-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="aa4d8-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="aa4d8-136">Метод</span><span class="sxs-lookup"><span data-stu-id="aa4d8-136">Method</span></span> |
| [<span data-ttu-id="aa4d8-137">дисплайневмессажеформ</span><span class="sxs-lookup"><span data-stu-id="aa4d8-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="aa4d8-138">Метод</span><span class="sxs-lookup"><span data-stu-id="aa4d8-138">Method</span></span> |
| [<span data-ttu-id="aa4d8-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="aa4d8-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="aa4d8-140">Метод</span><span class="sxs-lookup"><span data-stu-id="aa4d8-140">Method</span></span> |
| [<span data-ttu-id="aa4d8-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="aa4d8-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="aa4d8-142">Метод</span><span class="sxs-lookup"><span data-stu-id="aa4d8-142">Method</span></span> |
| [<span data-ttu-id="aa4d8-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="aa4d8-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="aa4d8-144">Метод</span><span class="sxs-lookup"><span data-stu-id="aa4d8-144">Method</span></span> |
| [<span data-ttu-id="aa4d8-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="aa4d8-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="aa4d8-146">Метод</span><span class="sxs-lookup"><span data-stu-id="aa4d8-146">Method</span></span> |
| [<span data-ttu-id="aa4d8-147">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="aa4d8-147">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="aa4d8-148">Метод</span><span class="sxs-lookup"><span data-stu-id="aa4d8-148">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="aa4d8-149">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="aa4d8-149">Namespaces</span></span>

<span data-ttu-id="aa4d8-150">[diagnostics](Office.context.mailbox.diagnostics.md). Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-150">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="aa4d8-151">[item](Office.context.mailbox.item.md). Предоставляет методы и свойства для доступа к сообщению или встрече в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-151">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="aa4d8-152">[userProfile](Office.context.mailbox.userProfile.md). Предоставляет сведения о пользователе в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-152">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="aa4d8-153">Members</span><span class="sxs-lookup"><span data-stu-id="aa4d8-153">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="aa4d8-154">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="aa4d8-154">ewsUrl: String</span></span>

<span data-ttu-id="aa4d8-p101">Получает URL-адрес конечной точки веб-служб Exchange (EWS) для этой учетной записи электронной почты. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="aa4d8-157">Этот элемент не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-157">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="aa4d8-p102">Удаленная служба может использовать значение `ewsUrl`, чтобы выполнять вызовы EWS для почтового ящика пользователя. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="aa4d8-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="aa4d8-160">Чтобы вызвать элемент `ewsUrl` в режиме чтения, в манифесте приложения должно быть указано разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-160">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="aa4d8-p103">Перед использованием элемента `ewsUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="aa4d8-163">Тип</span><span class="sxs-lookup"><span data-stu-id="aa4d8-163">Type</span></span>

*   <span data-ttu-id="aa4d8-164">String</span><span class="sxs-lookup"><span data-stu-id="aa4d8-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa4d8-165">Requirements</span><span class="sxs-lookup"><span data-stu-id="aa4d8-165">Requirements</span></span>

|<span data-ttu-id="aa4d8-166">Требование</span><span class="sxs-lookup"><span data-stu-id="aa4d8-166">Requirement</span></span>| <span data-ttu-id="aa4d8-167">Значение</span><span class="sxs-lookup"><span data-stu-id="aa4d8-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa4d8-168">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa4d8-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa4d8-169">1.0</span><span class="sxs-lookup"><span data-stu-id="aa4d8-169">1.0</span></span>|
|[<span data-ttu-id="aa4d8-170">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa4d8-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa4d8-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa4d8-171">ReadItem</span></span>|
|[<span data-ttu-id="aa4d8-172">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa4d8-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa4d8-173">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="aa4d8-173">Compose or Read</span></span>|

<br>

---
---

#### <a name="resturl-string"></a><span data-ttu-id="aa4d8-174">restUrl: String</span><span class="sxs-lookup"><span data-stu-id="aa4d8-174">restUrl: String</span></span>

<span data-ttu-id="aa4d8-175">Возвращает URL-адрес конечной точки REST для этой учетной записи электронной почты.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-175">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="aa4d8-176">С помощью значения `restUrl` можно выполнять вызовы [REST API](/outlook/rest/) для почтового ящика пользователя.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-176">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="aa4d8-177">Тип</span><span class="sxs-lookup"><span data-stu-id="aa4d8-177">Type</span></span>

*   <span data-ttu-id="aa4d8-178">String</span><span class="sxs-lookup"><span data-stu-id="aa4d8-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="aa4d8-179">Requirements</span><span class="sxs-lookup"><span data-stu-id="aa4d8-179">Requirements</span></span>

|<span data-ttu-id="aa4d8-180">Требование</span><span class="sxs-lookup"><span data-stu-id="aa4d8-180">Requirement</span></span>| <span data-ttu-id="aa4d8-181">Значение</span><span class="sxs-lookup"><span data-stu-id="aa4d8-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa4d8-182">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="aa4d8-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa4d8-183">1.5</span><span class="sxs-lookup"><span data-stu-id="aa4d8-183">1.5</span></span> |
|[<span data-ttu-id="aa4d8-184">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa4d8-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa4d8-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa4d8-185">ReadItem</span></span>|
|[<span data-ttu-id="aa4d8-186">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa4d8-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa4d8-187">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="aa4d8-187">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="aa4d8-188">Методы</span><span class="sxs-lookup"><span data-stu-id="aa4d8-188">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="aa4d8-189">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="aa4d8-189">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="aa4d8-190">Добавляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-190">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="aa4d8-191">В настоящее время поддерживаются только события типа `Office.EventType.ItemChanged`, которые вызываются, когда пользователь выбирает новый элемент.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-191">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item.</span></span> <span data-ttu-id="aa4d8-192">Это событие используется надстройками, реализующими закрепляемую область задач, и позволяет надстройке обновлять пользовательский интерфейс области задач в соответствии с выбранным в данный момент элементом.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-192">This event is used by add-ins that implement a pinnable task pane, and allows the add-in to refresh the task pane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa4d8-193">Параметры</span><span class="sxs-lookup"><span data-stu-id="aa4d8-193">Parameters</span></span>

| <span data-ttu-id="aa4d8-194">Имя</span><span class="sxs-lookup"><span data-stu-id="aa4d8-194">Name</span></span> | <span data-ttu-id="aa4d8-195">Тип</span><span class="sxs-lookup"><span data-stu-id="aa4d8-195">Type</span></span> | <span data-ttu-id="aa4d8-196">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="aa4d8-196">Attributes</span></span> | <span data-ttu-id="aa4d8-197">Описание</span><span class="sxs-lookup"><span data-stu-id="aa4d8-197">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="aa4d8-198">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="aa4d8-198">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="aa4d8-199">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-199">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="aa4d8-200">Function</span><span class="sxs-lookup"><span data-stu-id="aa4d8-200">Function</span></span> || <span data-ttu-id="aa4d8-p105">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-p105">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="aa4d8-204">Объект</span><span class="sxs-lookup"><span data-stu-id="aa4d8-204">Object</span></span> | <span data-ttu-id="aa4d8-205">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="aa4d8-205">&lt;optional&gt;</span></span> | <span data-ttu-id="aa4d8-206">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-206">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="aa4d8-207">Object</span><span class="sxs-lookup"><span data-stu-id="aa4d8-207">Object</span></span> | <span data-ttu-id="aa4d8-208">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="aa4d8-208">&lt;optional&gt;</span></span> | <span data-ttu-id="aa4d8-209">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-209">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="aa4d8-210">функция</span><span class="sxs-lookup"><span data-stu-id="aa4d8-210">function</span></span>| <span data-ttu-id="aa4d8-211">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="aa4d8-211">&lt;optional&gt;</span></span>|<span data-ttu-id="aa4d8-212">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="aa4d8-212">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aa4d8-213">Requirements</span><span class="sxs-lookup"><span data-stu-id="aa4d8-213">Requirements</span></span>

|<span data-ttu-id="aa4d8-214">Требование</span><span class="sxs-lookup"><span data-stu-id="aa4d8-214">Requirement</span></span>| <span data-ttu-id="aa4d8-215">Значение</span><span class="sxs-lookup"><span data-stu-id="aa4d8-215">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa4d8-216">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="aa4d8-216">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa4d8-217">1.5</span><span class="sxs-lookup"><span data-stu-id="aa4d8-217">1.5</span></span> |
|[<span data-ttu-id="aa4d8-218">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa4d8-218">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa4d8-219">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa4d8-219">ReadItem</span></span> |
|[<span data-ttu-id="aa4d8-220">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa4d8-220">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa4d8-221">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="aa4d8-221">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa4d8-222">Пример</span><span class="sxs-lookup"><span data-stu-id="aa4d8-222">Example</span></span>

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

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="aa4d8-223">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="aa4d8-223">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="aa4d8-224">Преобразовывает идентификатор элемента из формата REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-224">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="aa4d8-225">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-225">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="aa4d8-p106">Формат идентификаторов, извлекаемых через API REST (например, [API Почты Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)), отличается от формата веб-служб Exchange (EWS). Метод `convertToEwsId` преобразовывает идентификатор в формате REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-p106">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa4d8-228">Параметры</span><span class="sxs-lookup"><span data-stu-id="aa4d8-228">Parameters</span></span>

|<span data-ttu-id="aa4d8-229">Имя</span><span class="sxs-lookup"><span data-stu-id="aa4d8-229">Name</span></span>| <span data-ttu-id="aa4d8-230">Тип</span><span class="sxs-lookup"><span data-stu-id="aa4d8-230">Type</span></span>| <span data-ttu-id="aa4d8-231">Описание</span><span class="sxs-lookup"><span data-stu-id="aa4d8-231">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="aa4d8-232">String</span><span class="sxs-lookup"><span data-stu-id="aa4d8-232">String</span></span>|<span data-ttu-id="aa4d8-233">Идентификатор элемента в формате REST API для Outlook</span><span class="sxs-lookup"><span data-stu-id="aa4d8-233">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="aa4d8-234">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="aa4d8-234">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.6)|<span data-ttu-id="aa4d8-235">Значение, определяющее версию REST API для Outlook, которая используется для извлечения идентификатора элемента.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-235">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aa4d8-236">Требования</span><span class="sxs-lookup"><span data-stu-id="aa4d8-236">Requirements</span></span>

|<span data-ttu-id="aa4d8-237">Требование</span><span class="sxs-lookup"><span data-stu-id="aa4d8-237">Requirement</span></span>| <span data-ttu-id="aa4d8-238">Значение</span><span class="sxs-lookup"><span data-stu-id="aa4d8-238">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa4d8-239">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="aa4d8-239">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa4d8-240">1.3</span><span class="sxs-lookup"><span data-stu-id="aa4d8-240">1.3</span></span>|
|[<span data-ttu-id="aa4d8-241">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa4d8-241">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa4d8-242">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="aa4d8-242">Restricted</span></span>|
|[<span data-ttu-id="aa4d8-243">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa4d8-243">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa4d8-244">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="aa4d8-244">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="aa4d8-245">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="aa4d8-245">Returns:</span></span>

<span data-ttu-id="aa4d8-246">Тип: String</span><span class="sxs-lookup"><span data-stu-id="aa4d8-246">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="aa4d8-247">Пример</span><span class="sxs-lookup"><span data-stu-id="aa4d8-247">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-16"></a><span data-ttu-id="aa4d8-248">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)}</span><span class="sxs-lookup"><span data-stu-id="aa4d8-248">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)}</span></span>

<span data-ttu-id="aa4d8-249">Получает словарь, содержащий сведения о локальном времени клиента.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-249">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="aa4d8-p107">Почтовое приложение для классической версии Outlook или версии в Интернете может использовать разные часовые пояса для дат и времени. Классическое приложение Outlook использует часовой пояс клиентского компьютера. Outlook в Интернете использует часовой пояс, заданный в Центре администрирования Exchange (EAC). Значения даты и времени должны обрабатываться так, чтобы значения в пользовательском интерфейсе всегда согласовывались с часовым поясом, ожидаемым пользователем.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-p107">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="aa4d8-p108">Если почтовое приложение работает в классическом клиенте Outlook, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса клиентского компьютера. Если почтовое приложение работает в Outlook в Интернете, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса, заданного в Центре администрирования Exchange.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-p108">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa4d8-255">Параметры</span><span class="sxs-lookup"><span data-stu-id="aa4d8-255">Parameters</span></span>

|<span data-ttu-id="aa4d8-256">Имя</span><span class="sxs-lookup"><span data-stu-id="aa4d8-256">Name</span></span>| <span data-ttu-id="aa4d8-257">Тип</span><span class="sxs-lookup"><span data-stu-id="aa4d8-257">Type</span></span>| <span data-ttu-id="aa4d8-258">Описание</span><span class="sxs-lookup"><span data-stu-id="aa4d8-258">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="aa4d8-259">Date</span><span class="sxs-lookup"><span data-stu-id="aa4d8-259">Date</span></span>|<span data-ttu-id="aa4d8-260">Объект Date</span><span class="sxs-lookup"><span data-stu-id="aa4d8-260">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aa4d8-261">Requirements</span><span class="sxs-lookup"><span data-stu-id="aa4d8-261">Requirements</span></span>

|<span data-ttu-id="aa4d8-262">Требование</span><span class="sxs-lookup"><span data-stu-id="aa4d8-262">Requirement</span></span>| <span data-ttu-id="aa4d8-263">Значение</span><span class="sxs-lookup"><span data-stu-id="aa4d8-263">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa4d8-264">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa4d8-264">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa4d8-265">1.0</span><span class="sxs-lookup"><span data-stu-id="aa4d8-265">1.0</span></span>|
|[<span data-ttu-id="aa4d8-266">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa4d8-266">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa4d8-267">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa4d8-267">ReadItem</span></span>|
|[<span data-ttu-id="aa4d8-268">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa4d8-268">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa4d8-269">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="aa4d8-269">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="aa4d8-270">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="aa4d8-270">Returns:</span></span>

<span data-ttu-id="aa4d8-271">Тип: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="aa4d8-271">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="aa4d8-272">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="aa4d8-272">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="aa4d8-273">Преобразовывает идентификатор элемента в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-273">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="aa4d8-274">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-274">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="aa4d8-p109">Формат идентификаторов, извлекаемых через EWS или свойство `itemId`, отличается от формата API REST (таких как [API Почты Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)). Метод `convertToRestId` преобразовывает идентификатор в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-p109">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa4d8-277">Параметры</span><span class="sxs-lookup"><span data-stu-id="aa4d8-277">Parameters</span></span>

|<span data-ttu-id="aa4d8-278">Имя</span><span class="sxs-lookup"><span data-stu-id="aa4d8-278">Name</span></span>| <span data-ttu-id="aa4d8-279">Тип</span><span class="sxs-lookup"><span data-stu-id="aa4d8-279">Type</span></span>| <span data-ttu-id="aa4d8-280">Описание</span><span class="sxs-lookup"><span data-stu-id="aa4d8-280">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="aa4d8-281">String</span><span class="sxs-lookup"><span data-stu-id="aa4d8-281">String</span></span>|<span data-ttu-id="aa4d8-282">Идентификатор элемента в формате EWS</span><span class="sxs-lookup"><span data-stu-id="aa4d8-282">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="aa4d8-283">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="aa4d8-283">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.6)|<span data-ttu-id="aa4d8-284">Значение, определяющее версию REST API для Outlook, с которой будет использоваться преобразованный идентификатор.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-284">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aa4d8-285">Requirements</span><span class="sxs-lookup"><span data-stu-id="aa4d8-285">Requirements</span></span>

|<span data-ttu-id="aa4d8-286">Требование</span><span class="sxs-lookup"><span data-stu-id="aa4d8-286">Requirement</span></span>| <span data-ttu-id="aa4d8-287">Значение</span><span class="sxs-lookup"><span data-stu-id="aa4d8-287">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa4d8-288">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="aa4d8-288">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa4d8-289">1.3</span><span class="sxs-lookup"><span data-stu-id="aa4d8-289">1.3</span></span>|
|[<span data-ttu-id="aa4d8-290">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa4d8-290">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa4d8-291">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="aa4d8-291">Restricted</span></span>|
|[<span data-ttu-id="aa4d8-292">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa4d8-292">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa4d8-293">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="aa4d8-293">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="aa4d8-294">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="aa4d8-294">Returns:</span></span>

<span data-ttu-id="aa4d8-295">Тип: String</span><span class="sxs-lookup"><span data-stu-id="aa4d8-295">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="aa4d8-296">Пример</span><span class="sxs-lookup"><span data-stu-id="aa4d8-296">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="aa4d8-297">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="aa4d8-297">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="aa4d8-298">Получает объект Date из словаря, содержащего сведения о времени.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-298">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="aa4d8-299">Метод `convertToUtcClientTime` преобразует словарь, содержащий локальную дату и время, в объект Date с правильными значениями локальной даты и времени.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-299">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa4d8-300">Параметры</span><span class="sxs-lookup"><span data-stu-id="aa4d8-300">Parameters</span></span>

|<span data-ttu-id="aa4d8-301">Имя</span><span class="sxs-lookup"><span data-stu-id="aa4d8-301">Name</span></span>| <span data-ttu-id="aa4d8-302">Тип</span><span class="sxs-lookup"><span data-stu-id="aa4d8-302">Type</span></span>| <span data-ttu-id="aa4d8-303">Описание</span><span class="sxs-lookup"><span data-stu-id="aa4d8-303">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="aa4d8-304">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="aa4d8-304">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.6)|<span data-ttu-id="aa4d8-305">Значение локального времени для преобразования.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-305">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aa4d8-306">Requirements</span><span class="sxs-lookup"><span data-stu-id="aa4d8-306">Requirements</span></span>

|<span data-ttu-id="aa4d8-307">Требование</span><span class="sxs-lookup"><span data-stu-id="aa4d8-307">Requirement</span></span>| <span data-ttu-id="aa4d8-308">Значение</span><span class="sxs-lookup"><span data-stu-id="aa4d8-308">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa4d8-309">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa4d8-309">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa4d8-310">1.0</span><span class="sxs-lookup"><span data-stu-id="aa4d8-310">1.0</span></span>|
|[<span data-ttu-id="aa4d8-311">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa4d8-311">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa4d8-312">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa4d8-312">ReadItem</span></span>|
|[<span data-ttu-id="aa4d8-313">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa4d8-313">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa4d8-314">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="aa4d8-314">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="aa4d8-315">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="aa4d8-315">Returns:</span></span>

<span data-ttu-id="aa4d8-316">Объект Date со временем в формате UTC.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-316">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="aa4d8-317">Тип: Date</span><span class="sxs-lookup"><span data-stu-id="aa4d8-317">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="aa4d8-318">Пример</span><span class="sxs-lookup"><span data-stu-id="aa4d8-318">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="aa4d8-319">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="aa4d8-319">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="aa4d8-320">Отображает имеющуюся встречу из календаря.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-320">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="aa4d8-321">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-321">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="aa4d8-322">Метод `displayAppointmentForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее сведения календаря о существующей встрече.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-322">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="aa4d8-p110">В Outlook для Mac с помощью этого метода можно отобразить одну встречу, которая не является частью повторяющегося ряда, или основную встречу такого ряда, но не экземпляр из него, так как в Outlook для Mac невозможно получить доступ к свойствам экземпляра повторяющегося ряда (в том числе к идентификатору элемента).</span><span class="sxs-lookup"><span data-stu-id="aa4d8-p110">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="aa4d8-325">В Outlook в Интернете этот метод открывает указанную форму, только если текст формы содержит символы размером не более 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-325">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="aa4d8-326">Если указанный идентификатор элемента не определяет существующую встречу, на клиентском компьютере или устройстве открывается пустая страница, и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-326">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa4d8-327">Параметры</span><span class="sxs-lookup"><span data-stu-id="aa4d8-327">Parameters</span></span>

|<span data-ttu-id="aa4d8-328">Имя</span><span class="sxs-lookup"><span data-stu-id="aa4d8-328">Name</span></span>| <span data-ttu-id="aa4d8-329">Тип</span><span class="sxs-lookup"><span data-stu-id="aa4d8-329">Type</span></span>| <span data-ttu-id="aa4d8-330">Описание</span><span class="sxs-lookup"><span data-stu-id="aa4d8-330">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="aa4d8-331">String</span><span class="sxs-lookup"><span data-stu-id="aa4d8-331">String</span></span>|<span data-ttu-id="aa4d8-332">Идентификатор веб-служб Exchange для существующей встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-332">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aa4d8-333">Requirements</span><span class="sxs-lookup"><span data-stu-id="aa4d8-333">Requirements</span></span>

|<span data-ttu-id="aa4d8-334">Требование</span><span class="sxs-lookup"><span data-stu-id="aa4d8-334">Requirement</span></span>| <span data-ttu-id="aa4d8-335">Значение</span><span class="sxs-lookup"><span data-stu-id="aa4d8-335">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa4d8-336">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa4d8-336">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa4d8-337">1.0</span><span class="sxs-lookup"><span data-stu-id="aa4d8-337">1.0</span></span>|
|[<span data-ttu-id="aa4d8-338">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa4d8-338">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa4d8-339">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa4d8-339">ReadItem</span></span>|
|[<span data-ttu-id="aa4d8-340">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa4d8-340">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa4d8-341">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="aa4d8-341">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa4d8-342">Пример</span><span class="sxs-lookup"><span data-stu-id="aa4d8-342">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="aa4d8-343">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="aa4d8-343">displayMessageForm(itemId)</span></span>

<span data-ttu-id="aa4d8-344">Отображает имеющееся сообщение.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-344">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="aa4d8-345">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-345">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="aa4d8-346">Метод `displayMessageForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее существующее сообщение.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-346">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="aa4d8-347">В Outlook в Интернете этот метод открывает указанную форму, только если текст формы содержит символы размером не более 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-347">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="aa4d8-348">Если указанный идентификатор элемента не определяет существующее сообщение, окно на клиентском компьютере не открывается и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-348">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="aa4d8-p111">Не используйте `displayMessageForm` с параметром `itemId`, который представляет собой встречу. Используйте метод `displayAppointmentForm`, чтобы отобразить сведения о существующей встрече, а метод `displayNewAppointmentForm` — для отображения формы создания встречи.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-p111">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa4d8-351">Параметры</span><span class="sxs-lookup"><span data-stu-id="aa4d8-351">Parameters</span></span>

|<span data-ttu-id="aa4d8-352">Имя</span><span class="sxs-lookup"><span data-stu-id="aa4d8-352">Name</span></span>| <span data-ttu-id="aa4d8-353">Тип</span><span class="sxs-lookup"><span data-stu-id="aa4d8-353">Type</span></span>| <span data-ttu-id="aa4d8-354">Описание</span><span class="sxs-lookup"><span data-stu-id="aa4d8-354">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="aa4d8-355">String</span><span class="sxs-lookup"><span data-stu-id="aa4d8-355">String</span></span>|<span data-ttu-id="aa4d8-356">Идентификатор веб-служб Exchange для существующего сообщения.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-356">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aa4d8-357">Требования</span><span class="sxs-lookup"><span data-stu-id="aa4d8-357">Requirements</span></span>

|<span data-ttu-id="aa4d8-358">Требование</span><span class="sxs-lookup"><span data-stu-id="aa4d8-358">Requirement</span></span>| <span data-ttu-id="aa4d8-359">Значение</span><span class="sxs-lookup"><span data-stu-id="aa4d8-359">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa4d8-360">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa4d8-360">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa4d8-361">1.0</span><span class="sxs-lookup"><span data-stu-id="aa4d8-361">1.0</span></span>|
|[<span data-ttu-id="aa4d8-362">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa4d8-362">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa4d8-363">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa4d8-363">ReadItem</span></span>|
|[<span data-ttu-id="aa4d8-364">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa4d8-364">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa4d8-365">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="aa4d8-365">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa4d8-366">Пример</span><span class="sxs-lookup"><span data-stu-id="aa4d8-366">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="aa4d8-367">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="aa4d8-367">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="aa4d8-368">Отображает форму для создания новой встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-368">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="aa4d8-369">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-369">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="aa4d8-p112">Метод `displayNewAppointmentForm` открывает форму, в которой пользователь может создать встречу или собрание. Если параметры заданы, поля формы встречи автоматически заполняются их содержимым.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-p112">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="aa4d8-p113">В Outlook в Интернете и на мобильных устройствах этот метод всегда отображает форму с полем участников. Если вы не укажете участников в качестве входных аргументов, метод отображает форму с кнопкой **Сохранить**. Если вы укажете участников, форма будет включать участников и кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-p113">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="aa4d8-p114">Если вы укажете участников или ресурсы с помощью параметра `requiredAttendees`, `optionalAttendees` или `resources` в клиенте Outlook с расширенными возможностями и Outlook RT, этот метод отобразит форму собрания с кнопкой **Отправить**. Если не указать получателей, этот метод отобразит форму встречи с кнопкой **Сохранить и закрыть**.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-p114">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="aa4d8-377">Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-377">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa4d8-378">Параметры</span><span class="sxs-lookup"><span data-stu-id="aa4d8-378">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="aa4d8-379">Все параметры являются необязательными.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-379">All parameters are optional.</span></span>

|<span data-ttu-id="aa4d8-380">Имя</span><span class="sxs-lookup"><span data-stu-id="aa4d8-380">Name</span></span>| <span data-ttu-id="aa4d8-381">Тип</span><span class="sxs-lookup"><span data-stu-id="aa4d8-381">Type</span></span>| <span data-ttu-id="aa4d8-382">Описание</span><span class="sxs-lookup"><span data-stu-id="aa4d8-382">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="aa4d8-383">Object</span><span class="sxs-lookup"><span data-stu-id="aa4d8-383">Object</span></span> | <span data-ttu-id="aa4d8-384">Словарь параметров, описывающий новую встречу.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-384">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="aa4d8-385">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="aa4d8-385">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="aa4d8-p115">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из обязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="aa4d8-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="aa4d8-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="aa4d8-p116">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из необязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="aa4d8-391">Date</span><span class="sxs-lookup"><span data-stu-id="aa4d8-391">Date</span></span> | <span data-ttu-id="aa4d8-392">Объект `Date`, указывающий дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-392">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="aa4d8-393">Date</span><span class="sxs-lookup"><span data-stu-id="aa4d8-393">Date</span></span> | <span data-ttu-id="aa4d8-394">Объект `Date`, указывающий дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-394">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="aa4d8-395">String</span><span class="sxs-lookup"><span data-stu-id="aa4d8-395">String</span></span> | <span data-ttu-id="aa4d8-p117">Строка со сведениями о месте встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-p117">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="aa4d8-398">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="aa4d8-398">Array.&lt;String&gt;</span></span> | <span data-ttu-id="aa4d8-p118">Массив строк, содержащий необходимые для встречи ресурсы. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-p118">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="aa4d8-401">String</span><span class="sxs-lookup"><span data-stu-id="aa4d8-401">String</span></span> | <span data-ttu-id="aa4d8-p119">Строка с темой встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-p119">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="aa4d8-404">String</span><span class="sxs-lookup"><span data-stu-id="aa4d8-404">String</span></span> | <span data-ttu-id="aa4d8-p120">Текст сообщения о встрече. Максимальный размер содержимого сообщения — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-p120">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="aa4d8-407">Требования</span><span class="sxs-lookup"><span data-stu-id="aa4d8-407">Requirements</span></span>

|<span data-ttu-id="aa4d8-408">Требование</span><span class="sxs-lookup"><span data-stu-id="aa4d8-408">Requirement</span></span>| <span data-ttu-id="aa4d8-409">Значение</span><span class="sxs-lookup"><span data-stu-id="aa4d8-409">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa4d8-410">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa4d8-410">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa4d8-411">1.0</span><span class="sxs-lookup"><span data-stu-id="aa4d8-411">1.0</span></span>|
|[<span data-ttu-id="aa4d8-412">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa4d8-412">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa4d8-413">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa4d8-413">ReadItem</span></span>|
|[<span data-ttu-id="aa4d8-414">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa4d8-414">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa4d8-415">Чтение</span><span class="sxs-lookup"><span data-stu-id="aa4d8-415">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa4d8-416">Пример</span><span class="sxs-lookup"><span data-stu-id="aa4d8-416">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="aa4d8-417">Дисплайневмессажеформ (Parameters)</span><span class="sxs-lookup"><span data-stu-id="aa4d8-417">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="aa4d8-418">Отображает форму для создания нового сообщения.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-418">Displays a form for creating a new message.</span></span>

<span data-ttu-id="aa4d8-419">`displayNewMessageForm` Метод открывает форму, которая позволяет пользователю создать новое сообщение.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-419">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="aa4d8-420">Если указаны параметры, поля формы сообщения автоматически заполняются содержимым параметров.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-420">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="aa4d8-421">Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-421">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa4d8-422">Параметры</span><span class="sxs-lookup"><span data-stu-id="aa4d8-422">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="aa4d8-423">Все параметры являются необязательными.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-423">All parameters are optional.</span></span>

|<span data-ttu-id="aa4d8-424">Имя</span><span class="sxs-lookup"><span data-stu-id="aa4d8-424">Name</span></span>| <span data-ttu-id="aa4d8-425">Тип</span><span class="sxs-lookup"><span data-stu-id="aa4d8-425">Type</span></span>| <span data-ttu-id="aa4d8-426">Описание</span><span class="sxs-lookup"><span data-stu-id="aa4d8-426">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="aa4d8-427">Object</span><span class="sxs-lookup"><span data-stu-id="aa4d8-427">Object</span></span> | <span data-ttu-id="aa4d8-428">Словарь параметров, описывающих новое сообщение.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-428">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="aa4d8-429">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="aa4d8-429">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="aa4d8-430">Массив строк, содержащий адреса электронной почты или массив, содержащий `EmailAddressDetails` объект для каждого из получателей в строке "Кому".</span><span class="sxs-lookup"><span data-stu-id="aa4d8-430">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="aa4d8-431">Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-431">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="aa4d8-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="aa4d8-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="aa4d8-433">Массив строк, содержащий адреса электронной почты или массив, содержащий `EmailAddressDetails` объект для каждого получателя в строке "копия".</span><span class="sxs-lookup"><span data-stu-id="aa4d8-433">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="aa4d8-434">Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-434">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="aa4d8-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span><span class="sxs-lookup"><span data-stu-id="aa4d8-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.6)&gt;</span></span> | <span data-ttu-id="aa4d8-436">Массив строк, содержащий адреса электронной почты или массив, содержащий `EmailAddressDetails` объект для каждого из получателей, указанных в строке "СК".</span><span class="sxs-lookup"><span data-stu-id="aa4d8-436">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="aa4d8-437">Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-437">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="aa4d8-438">String</span><span class="sxs-lookup"><span data-stu-id="aa4d8-438">String</span></span> | <span data-ttu-id="aa4d8-439">Строка, содержащая тему сообщения.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-439">A string containing the subject of the message.</span></span> <span data-ttu-id="aa4d8-440">Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-440">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="aa4d8-441">String</span><span class="sxs-lookup"><span data-stu-id="aa4d8-441">String</span></span> | <span data-ttu-id="aa4d8-442">Текст сообщения в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-442">The HTML body of the message.</span></span> <span data-ttu-id="aa4d8-443">Максимальный размер содержимого сообщения — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-443">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="aa4d8-444">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="aa4d8-444">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="aa4d8-445">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-445">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="aa4d8-446">String</span><span class="sxs-lookup"><span data-stu-id="aa4d8-446">String</span></span> | <span data-ttu-id="aa4d8-p127">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-p127">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="aa4d8-449">Строка</span><span class="sxs-lookup"><span data-stu-id="aa4d8-449">String</span></span> | <span data-ttu-id="aa4d8-450">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-450">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="aa4d8-451">String</span><span class="sxs-lookup"><span data-stu-id="aa4d8-451">String</span></span> | <span data-ttu-id="aa4d8-p128">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-p128">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="aa4d8-454">Логический</span><span class="sxs-lookup"><span data-stu-id="aa4d8-454">Boolean</span></span> | <span data-ttu-id="aa4d8-p129">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-p129">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="aa4d8-457">Строка</span><span class="sxs-lookup"><span data-stu-id="aa4d8-457">String</span></span> | <span data-ttu-id="aa4d8-458">Используется, только если свойству `type` присвоено значение `item`.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-458">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="aa4d8-459">Идентификатор элемента EWS существующего сообщения электронной почты, которое необходимо присоединить к новому сообщению.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-459">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="aa4d8-460">Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-460">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="aa4d8-461">Requirements</span><span class="sxs-lookup"><span data-stu-id="aa4d8-461">Requirements</span></span>

|<span data-ttu-id="aa4d8-462">Требование</span><span class="sxs-lookup"><span data-stu-id="aa4d8-462">Requirement</span></span>| <span data-ttu-id="aa4d8-463">Значение</span><span class="sxs-lookup"><span data-stu-id="aa4d8-463">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa4d8-464">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="aa4d8-464">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa4d8-465">1.6</span><span class="sxs-lookup"><span data-stu-id="aa4d8-465">1.6</span></span> |
|[<span data-ttu-id="aa4d8-466">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa4d8-466">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa4d8-467">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa4d8-467">ReadItem</span></span>|
|[<span data-ttu-id="aa4d8-468">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa4d8-468">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa4d8-469">Чтение</span><span class="sxs-lookup"><span data-stu-id="aa4d8-469">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa4d8-470">Пример</span><span class="sxs-lookup"><span data-stu-id="aa4d8-470">Example</span></span>

```js
Office.context.mailbox.displayNewMessageForm(
  {
    toRecipients: Office.context.mailbox.item.to, // Copy the To line from current item
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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="aa4d8-471">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="aa4d8-471">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="aa4d8-472">Возвращает строку, содержащую маркер, который используется для вызова интерфейсов REST API или веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-472">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="aa4d8-p131">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-p131">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="aa4d8-475">Рекомендуем сделать так, чтобы по мере возможности надстройки использовали интерфейсы REST API, а не веб-службы Exchange.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-475">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="aa4d8-476">Для вызова метода `getCallbackTokenAsync` в режиме чтения требуется минимальный уровень разрешения **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-476">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="aa4d8-477">Для вызова `getCallbackTokenAsync` в режиме создания сообщения требуется сохранить элемент.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-477">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="aa4d8-478">Для метода [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) требуется минимальный уровень разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-478">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

<span data-ttu-id="aa4d8-479">**Маркеры REST**</span><span class="sxs-lookup"><span data-stu-id="aa4d8-479">**REST Tokens**</span></span>

<span data-ttu-id="aa4d8-p133">Если запрашивается маркер REST (`options.isRest = true`), полученный маркер не подойдет для проверки подлинности при вызовах веб-служб Exchange. Область действия маркера будет ограничена доступом только для чтения к текущему элементу и его вложениям, если в манифесте надстройки не указано разрешение [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission). Если указано разрешение `ReadWriteMailbox`, полученный маркер предоставит доступ на чтение и запись к почте, календарю и контактам, включая возможность отправки почты.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-p133">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="aa4d8-483">С помощью свойства `restUrl` надстройка должна определить правильный URL-адрес для вызовов REST API.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-483">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="aa4d8-484">**Маркеры EWS**</span><span class="sxs-lookup"><span data-stu-id="aa4d8-484">**EWS Tokens**</span></span>

<span data-ttu-id="aa4d8-p134">Если запрашивается маркер EWS (`options.isRest = false`), полученный маркер не подойдет для проверки подлинности при вызовах REST API. Область действия маркера будет ограничена доступом к текущему элементу.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-p134">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="aa4d8-487">С помощью свойства `ewsUrl` надстройка должна определить правильный URL-адрес для вызовов EWS.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-487">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

<span data-ttu-id="aa4d8-488">Вы можете передать сторонней системе маркер и идентификатор вложения или элемента.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-488">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="aa4d8-489">Сторонняя система использует этот маркер как маркер авторизации, чтобы вызвать операцию [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) или [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) веб-служб Exchange для возврата вложения или элемента.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-489">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="aa4d8-490">Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="aa4d8-490">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa4d8-491">Параметры</span><span class="sxs-lookup"><span data-stu-id="aa4d8-491">Parameters</span></span>

|<span data-ttu-id="aa4d8-492">Имя</span><span class="sxs-lookup"><span data-stu-id="aa4d8-492">Name</span></span>| <span data-ttu-id="aa4d8-493">Тип</span><span class="sxs-lookup"><span data-stu-id="aa4d8-493">Type</span></span>| <span data-ttu-id="aa4d8-494">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="aa4d8-494">Attributes</span></span>| <span data-ttu-id="aa4d8-495">Описание</span><span class="sxs-lookup"><span data-stu-id="aa4d8-495">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="aa4d8-496">Object</span><span class="sxs-lookup"><span data-stu-id="aa4d8-496">Object</span></span> | <span data-ttu-id="aa4d8-497">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="aa4d8-497">&lt;optional&gt;</span></span> | <span data-ttu-id="aa4d8-498">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-498">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="aa4d8-499">Boolean</span><span class="sxs-lookup"><span data-stu-id="aa4d8-499">Boolean</span></span> |  <span data-ttu-id="aa4d8-500">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="aa4d8-500">&lt;optional&gt;</span></span> | <span data-ttu-id="aa4d8-p136">Определяет, будет ли предоставленный маркер использоваться для интерфейсов REST API Outlook или веб-служб Exchange. Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-p136">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="aa4d8-503">Объект</span><span class="sxs-lookup"><span data-stu-id="aa4d8-503">Object</span></span> |  <span data-ttu-id="aa4d8-504">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="aa4d8-504">&lt;optional&gt;</span></span> | <span data-ttu-id="aa4d8-505">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-505">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="aa4d8-506">функция</span><span class="sxs-lookup"><span data-stu-id="aa4d8-506">function</span></span>||<span data-ttu-id="aa4d8-507">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="aa4d8-507">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="aa4d8-508">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-508">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="aa4d8-509">При наличии ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут предоставлять дополнительные сведения.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-509">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="aa4d8-510">Ошибки</span><span class="sxs-lookup"><span data-stu-id="aa4d8-510">Errors</span></span>

|<span data-ttu-id="aa4d8-511">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="aa4d8-511">Error code</span></span>|<span data-ttu-id="aa4d8-512">Описание</span><span class="sxs-lookup"><span data-stu-id="aa4d8-512">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="aa4d8-513">Не удалось выполнить запрос.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-513">The request has failed.</span></span> <span data-ttu-id="aa4d8-514">Просмотрите объект диагностики для кода ошибки HTTP.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-514">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="aa4d8-515">Сервер Exchange Server вернул ошибку.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-515">The Exchange server returned an error.</span></span> <span data-ttu-id="aa4d8-516">Для получения дополнительных сведений просмотрите объект диагностики.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-516">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="aa4d8-517">Пользователь отключен от сети.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-517">The user is no longer connected to the network.</span></span> <span data-ttu-id="aa4d8-518">Проверьте сетевое подключение и повторите попытку.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-518">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aa4d8-519">Требования</span><span class="sxs-lookup"><span data-stu-id="aa4d8-519">Requirements</span></span>

|<span data-ttu-id="aa4d8-520">Требование</span><span class="sxs-lookup"><span data-stu-id="aa4d8-520">Requirement</span></span>| <span data-ttu-id="aa4d8-521">Значение</span><span class="sxs-lookup"><span data-stu-id="aa4d8-521">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa4d8-522">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="aa4d8-522">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa4d8-523">1.5</span><span class="sxs-lookup"><span data-stu-id="aa4d8-523">1.5</span></span> |
|[<span data-ttu-id="aa4d8-524">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa4d8-524">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa4d8-525">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa4d8-525">ReadItem</span></span>|
|[<span data-ttu-id="aa4d8-526">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa4d8-526">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa4d8-527">Создание и чтение</span><span class="sxs-lookup"><span data-stu-id="aa4d8-527">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa4d8-528">Пример</span><span class="sxs-lookup"><span data-stu-id="aa4d8-528">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="aa4d8-529">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="aa4d8-529">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="aa4d8-530">Получает строку, содержащую маркер, используемый для получения вложения или элемента с Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-530">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="aa4d8-p140">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-p140">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="aa4d8-533">Вы можете передать сторонней системе маркер и идентификатор вложения или элемента.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-533">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="aa4d8-534">Сторонняя система использует этот маркер как маркер авторизации, чтобы вызвать операцию [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) или [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) веб-служб Exchange для возврата вложения или элемента.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-534">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="aa4d8-535">Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="aa4d8-535">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="aa4d8-536">Для вызова метода `getCallbackTokenAsync` в режиме чтения требуется минимальный уровень разрешения **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-536">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="aa4d8-537">Для вызова `getCallbackTokenAsync` в режиме создания сообщения требуется сохранить элемент.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-537">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="aa4d8-538">Для метода [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) требуется минимальный уровень разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-538">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa4d8-539">Параметры</span><span class="sxs-lookup"><span data-stu-id="aa4d8-539">Parameters</span></span>

|<span data-ttu-id="aa4d8-540">Имя</span><span class="sxs-lookup"><span data-stu-id="aa4d8-540">Name</span></span>| <span data-ttu-id="aa4d8-541">Тип</span><span class="sxs-lookup"><span data-stu-id="aa4d8-541">Type</span></span>| <span data-ttu-id="aa4d8-542">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="aa4d8-542">Attributes</span></span>| <span data-ttu-id="aa4d8-543">Описание</span><span class="sxs-lookup"><span data-stu-id="aa4d8-543">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="aa4d8-544">функция</span><span class="sxs-lookup"><span data-stu-id="aa4d8-544">function</span></span>||<span data-ttu-id="aa4d8-545">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="aa4d8-545">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="aa4d8-546">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-546">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="aa4d8-547">При наличии ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут предоставлять дополнительные сведения.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-547">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="aa4d8-548">Объект</span><span class="sxs-lookup"><span data-stu-id="aa4d8-548">Object</span></span>| <span data-ttu-id="aa4d8-549">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="aa4d8-549">&lt;optional&gt;</span></span>|<span data-ttu-id="aa4d8-550">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-550">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="aa4d8-551">Ошибки</span><span class="sxs-lookup"><span data-stu-id="aa4d8-551">Errors</span></span>

|<span data-ttu-id="aa4d8-552">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="aa4d8-552">Error code</span></span>|<span data-ttu-id="aa4d8-553">Описание</span><span class="sxs-lookup"><span data-stu-id="aa4d8-553">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="aa4d8-554">Не удалось выполнить запрос.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-554">The request has failed.</span></span> <span data-ttu-id="aa4d8-555">Просмотрите объект диагностики для кода ошибки HTTP.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-555">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="aa4d8-556">Сервер Exchange Server вернул ошибку.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-556">The Exchange server returned an error.</span></span> <span data-ttu-id="aa4d8-557">Для получения дополнительных сведений просмотрите объект диагностики.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-557">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="aa4d8-558">Пользователь отключен от сети.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-558">The user is no longer connected to the network.</span></span> <span data-ttu-id="aa4d8-559">Проверьте сетевое подключение и повторите попытку.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-559">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aa4d8-560">Требования</span><span class="sxs-lookup"><span data-stu-id="aa4d8-560">Requirements</span></span>

|<span data-ttu-id="aa4d8-561">Требование</span><span class="sxs-lookup"><span data-stu-id="aa4d8-561">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="aa4d8-562">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa4d8-562">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa4d8-563">1.0</span><span class="sxs-lookup"><span data-stu-id="aa4d8-563">1.0</span></span> | <span data-ttu-id="aa4d8-564">1.3</span><span class="sxs-lookup"><span data-stu-id="aa4d8-564">1.3</span></span> |
|[<span data-ttu-id="aa4d8-565">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa4d8-565">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa4d8-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa4d8-566">ReadItem</span></span> | <span data-ttu-id="aa4d8-567">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa4d8-567">ReadItem</span></span> |
|[<span data-ttu-id="aa4d8-568">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa4d8-568">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa4d8-569">Чтение</span><span class="sxs-lookup"><span data-stu-id="aa4d8-569">Read</span></span> | <span data-ttu-id="aa4d8-570">Создание</span><span class="sxs-lookup"><span data-stu-id="aa4d8-570">Compose</span></span> |

##### <a name="example"></a><span data-ttu-id="aa4d8-571">Пример</span><span class="sxs-lookup"><span data-stu-id="aa4d8-571">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="aa4d8-572">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="aa4d8-572">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="aa4d8-573">Получает маркер, идентифицирующий пользователя и надстройку Office.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-573">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="aa4d8-574">Метод `getUserIdentityTokenAsync` возвращает маркер, который можно использовать для идентификации, а также [проверки подлинности надстройки и пользователя в сторонней системе](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="aa4d8-574">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa4d8-575">Параметры</span><span class="sxs-lookup"><span data-stu-id="aa4d8-575">Parameters</span></span>

|<span data-ttu-id="aa4d8-576">Имя</span><span class="sxs-lookup"><span data-stu-id="aa4d8-576">Name</span></span>| <span data-ttu-id="aa4d8-577">Тип</span><span class="sxs-lookup"><span data-stu-id="aa4d8-577">Type</span></span>| <span data-ttu-id="aa4d8-578">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="aa4d8-578">Attributes</span></span>| <span data-ttu-id="aa4d8-579">Описание</span><span class="sxs-lookup"><span data-stu-id="aa4d8-579">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="aa4d8-580">функция</span><span class="sxs-lookup"><span data-stu-id="aa4d8-580">function</span></span>||<span data-ttu-id="aa4d8-581">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="aa4d8-581">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="aa4d8-582">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-582">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="aa4d8-583">При наличии ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут предоставлять дополнительные сведения.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-583">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="aa4d8-584">Объект</span><span class="sxs-lookup"><span data-stu-id="aa4d8-584">Object</span></span>| <span data-ttu-id="aa4d8-585">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="aa4d8-585">&lt;optional&gt;</span></span>|<span data-ttu-id="aa4d8-586">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-586">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="aa4d8-587">Ошибки</span><span class="sxs-lookup"><span data-stu-id="aa4d8-587">Errors</span></span>

|<span data-ttu-id="aa4d8-588">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="aa4d8-588">Error code</span></span>|<span data-ttu-id="aa4d8-589">Описание</span><span class="sxs-lookup"><span data-stu-id="aa4d8-589">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="aa4d8-590">Не удалось выполнить запрос.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-590">The request has failed.</span></span> <span data-ttu-id="aa4d8-591">Просмотрите объект диагностики для кода ошибки HTTP.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-591">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="aa4d8-592">Сервер Exchange Server вернул ошибку.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-592">The Exchange server returned an error.</span></span> <span data-ttu-id="aa4d8-593">Для получения дополнительных сведений просмотрите объект диагностики.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-593">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="aa4d8-594">Пользователь отключен от сети.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-594">The user is no longer connected to the network.</span></span> <span data-ttu-id="aa4d8-595">Проверьте сетевое подключение и повторите попытку.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-595">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aa4d8-596">Требования</span><span class="sxs-lookup"><span data-stu-id="aa4d8-596">Requirements</span></span>

|<span data-ttu-id="aa4d8-597">Требование</span><span class="sxs-lookup"><span data-stu-id="aa4d8-597">Requirement</span></span>| <span data-ttu-id="aa4d8-598">Значение</span><span class="sxs-lookup"><span data-stu-id="aa4d8-598">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa4d8-599">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="aa4d8-599">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa4d8-600">1.0</span><span class="sxs-lookup"><span data-stu-id="aa4d8-600">1.0</span></span>|
|[<span data-ttu-id="aa4d8-601">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa4d8-601">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa4d8-602">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa4d8-602">ReadItem</span></span>|
|[<span data-ttu-id="aa4d8-603">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa4d8-603">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa4d8-604">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="aa4d8-604">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa4d8-605">Пример</span><span class="sxs-lookup"><span data-stu-id="aa4d8-605">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="aa4d8-606">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="aa4d8-606">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="aa4d8-607">Выполняет асинхронный запрос для веб-служб Exchange (EWS) на сервере Exchange Server, на котором размещен почтовый ящик пользователя.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-607">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="aa4d8-608">Этот метод не поддерживается в следующих сценариях:</span><span class="sxs-lookup"><span data-stu-id="aa4d8-608">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="aa4d8-609">В Outlook для iOS и Android</span><span class="sxs-lookup"><span data-stu-id="aa4d8-609">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="aa4d8-610">Если надстройка загружается в почтовый ящик Gmail.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-610">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="aa4d8-611">В таких случаях надстройка должна [использовать REST API](/outlook/add-ins/use-rest-api) для доступа к почтовому ящику пользователя.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-611">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="aa4d8-612">Метод `makeEwsRequestAsync` отправляет запрос EWS от имени надстройки в Exchange.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-612">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="aa4d8-613">Список поддерживаемых операций EWS см. в статье [Вызов веб-служб из надстройки Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support).</span><span class="sxs-lookup"><span data-stu-id="aa4d8-613">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="aa4d8-614">С помощью метода `makeEwsRequestAsync` невозможно запрашивать элементы, связанные с папкой.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-614">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="aa4d8-615">В запросе XML должна быть указана кодировка UTF-8.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-615">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="aa4d8-p150">У вашей надстройки должно быть разрешение **ReadWriteMailbox** для использования метода `makeEwsRequestAsync`. Сведения об использовании разрешения **ReadWriteMailbox** и операций EWS, которые можно вызывать с помощью метода `makeEwsRequestAsync`, см. в статье [Указание разрешений для доступа почтовой надстройки к почтовому ящику пользователя](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="aa4d8-p150">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="aa4d8-618">Администратор сервера должен установить значение true для параметра `OAuthAuthentication` в каталоге сервера клиентского доступа EWS, чтобы метод `makeEwsRequestAsync` мог выполнять запросы EWS.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-618">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="aa4d8-619">Различия версий</span><span class="sxs-lookup"><span data-stu-id="aa4d8-619">Version differences</span></span>

<span data-ttu-id="aa4d8-620">Если вы используете метод `makeEwsRequestAsync` в почтовых приложениях, которые выполняются в Outlook версии более ранней, чем 15.0.4535.1004, указывайте кодировку `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-620">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="aa4d8-p151">Значение кодировки не нужно указывать, если почтовое приложение выполняется в Outlook в Интернете. Чтобы определить, выполняется ли приложение в Outlook или Outlook в Интернете, используйте свойство mailbox.diagnostics.hostName. Используемую версию Outlook можно определить с помощью свойства mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-p151">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa4d8-624">Параметры</span><span class="sxs-lookup"><span data-stu-id="aa4d8-624">Parameters</span></span>

|<span data-ttu-id="aa4d8-625">Имя</span><span class="sxs-lookup"><span data-stu-id="aa4d8-625">Name</span></span>| <span data-ttu-id="aa4d8-626">Тип</span><span class="sxs-lookup"><span data-stu-id="aa4d8-626">Type</span></span>| <span data-ttu-id="aa4d8-627">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="aa4d8-627">Attributes</span></span>| <span data-ttu-id="aa4d8-628">Описание</span><span class="sxs-lookup"><span data-stu-id="aa4d8-628">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="aa4d8-629">String</span><span class="sxs-lookup"><span data-stu-id="aa4d8-629">String</span></span>||<span data-ttu-id="aa4d8-630">Запрос EWS.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-630">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="aa4d8-631">function</span><span class="sxs-lookup"><span data-stu-id="aa4d8-631">function</span></span>||<span data-ttu-id="aa4d8-632">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="aa4d8-632">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="aa4d8-633">Результат XML вызова EWS указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-633">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="aa4d8-634">Если размер результата превышает 1 МБ, возвращается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-634">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="aa4d8-635">Объект</span><span class="sxs-lookup"><span data-stu-id="aa4d8-635">Object</span></span>| <span data-ttu-id="aa4d8-636">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="aa4d8-636">&lt;optional&gt;</span></span>|<span data-ttu-id="aa4d8-637">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-637">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aa4d8-638">Требования</span><span class="sxs-lookup"><span data-stu-id="aa4d8-638">Requirements</span></span>

|<span data-ttu-id="aa4d8-639">Требование</span><span class="sxs-lookup"><span data-stu-id="aa4d8-639">Requirement</span></span>| <span data-ttu-id="aa4d8-640">Значение</span><span class="sxs-lookup"><span data-stu-id="aa4d8-640">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa4d8-641">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="aa4d8-641">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa4d8-642">1.0</span><span class="sxs-lookup"><span data-stu-id="aa4d8-642">1.0</span></span>|
|[<span data-ttu-id="aa4d8-643">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa4d8-643">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa4d8-644">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="aa4d8-644">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="aa4d8-645">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa4d8-645">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa4d8-646">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="aa4d8-646">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="aa4d8-647">Пример</span><span class="sxs-lookup"><span data-stu-id="aa4d8-647">Example</span></span>

<span data-ttu-id="aa4d8-648">В приведенном ниже примере вызывается `makeEwsRequestAsync` для получения темы элемента с помощью операции `GetItem`.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-648">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="aa4d8-649">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="aa4d8-649">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="aa4d8-650">Удаляет обработчиков для поддерживаемого типа события.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-650">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="aa4d8-651">В настоящее время единственный поддерживаемый тип события — `Office.EventType.ItemChanged`.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-651">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="aa4d8-652">Параметры</span><span class="sxs-lookup"><span data-stu-id="aa4d8-652">Parameters</span></span>

| <span data-ttu-id="aa4d8-653">Имя</span><span class="sxs-lookup"><span data-stu-id="aa4d8-653">Name</span></span> | <span data-ttu-id="aa4d8-654">Тип</span><span class="sxs-lookup"><span data-stu-id="aa4d8-654">Type</span></span> | <span data-ttu-id="aa4d8-655">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="aa4d8-655">Attributes</span></span> | <span data-ttu-id="aa4d8-656">Описание</span><span class="sxs-lookup"><span data-stu-id="aa4d8-656">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="aa4d8-657">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="aa4d8-657">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="aa4d8-658">Событие, которое должно отменить обработчик.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-658">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="aa4d8-659">Объект</span><span class="sxs-lookup"><span data-stu-id="aa4d8-659">Object</span></span> | <span data-ttu-id="aa4d8-660">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="aa4d8-660">&lt;optional&gt;</span></span> | <span data-ttu-id="aa4d8-661">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-661">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="aa4d8-662">Object</span><span class="sxs-lookup"><span data-stu-id="aa4d8-662">Object</span></span> | <span data-ttu-id="aa4d8-663">&lt;Необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="aa4d8-663">&lt;optional&gt;</span></span> | <span data-ttu-id="aa4d8-664">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="aa4d8-664">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="aa4d8-665">функция</span><span class="sxs-lookup"><span data-stu-id="aa4d8-665">function</span></span>| <span data-ttu-id="aa4d8-666">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="aa4d8-666">&lt;optional&gt;</span></span>|<span data-ttu-id="aa4d8-667">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="aa4d8-667">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="aa4d8-668">Requirements</span><span class="sxs-lookup"><span data-stu-id="aa4d8-668">Requirements</span></span>

|<span data-ttu-id="aa4d8-669">Требование</span><span class="sxs-lookup"><span data-stu-id="aa4d8-669">Requirement</span></span>| <span data-ttu-id="aa4d8-670">Значение</span><span class="sxs-lookup"><span data-stu-id="aa4d8-670">Value</span></span>|
|---|---|
|[<span data-ttu-id="aa4d8-671">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="aa4d8-671">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="aa4d8-672">1.5</span><span class="sxs-lookup"><span data-stu-id="aa4d8-672">1.5</span></span> |
|[<span data-ttu-id="aa4d8-673">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="aa4d8-673">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="aa4d8-674">ReadItem</span><span class="sxs-lookup"><span data-stu-id="aa4d8-674">ReadItem</span></span> |
|[<span data-ttu-id="aa4d8-675">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="aa4d8-675">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="aa4d8-676">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="aa4d8-676">Compose or Read</span></span>|
