---
title: Office.context.mailbox — набор обязательных элементов 1.5
description: ''
ms.date: 11/27/2019
localization_priority: Priority
ms.openlocfilehash: eefeab2cf6fbe78451afae7e588640fe7f50dba4
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629688"
---
# <a name="mailbox"></a><span data-ttu-id="b771c-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="b771c-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="b771c-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="b771c-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="b771c-104">Предоставляет для Microsoft Outlook доступ к объектной модели надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="b771c-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="b771c-105">Требования</span><span class="sxs-lookup"><span data-stu-id="b771c-105">Requirements</span></span>

|<span data-ttu-id="b771c-106">Требование</span><span class="sxs-lookup"><span data-stu-id="b771c-106">Requirement</span></span>| <span data-ttu-id="b771c-107">Значение</span><span class="sxs-lookup"><span data-stu-id="b771c-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="b771c-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b771c-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b771c-109">1.0</span><span class="sxs-lookup"><span data-stu-id="b771c-109">1.0</span></span>|
|[<span data-ttu-id="b771c-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b771c-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b771c-111">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="b771c-111">Restricted</span></span>|
|[<span data-ttu-id="b771c-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b771c-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b771c-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b771c-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="b771c-114">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="b771c-114">Members and methods</span></span>

| <span data-ttu-id="b771c-115">Элемент</span><span class="sxs-lookup"><span data-stu-id="b771c-115">Member</span></span> | <span data-ttu-id="b771c-116">Тип</span><span class="sxs-lookup"><span data-stu-id="b771c-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="b771c-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="b771c-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="b771c-118">Элемент</span><span class="sxs-lookup"><span data-stu-id="b771c-118">Member</span></span> |
| [<span data-ttu-id="b771c-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="b771c-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="b771c-120">Элемент</span><span class="sxs-lookup"><span data-stu-id="b771c-120">Member</span></span> |
| [<span data-ttu-id="b771c-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="b771c-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="b771c-122">Метод</span><span class="sxs-lookup"><span data-stu-id="b771c-122">Method</span></span> |
| [<span data-ttu-id="b771c-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="b771c-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="b771c-124">Метод</span><span class="sxs-lookup"><span data-stu-id="b771c-124">Method</span></span> |
| [<span data-ttu-id="b771c-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="b771c-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="b771c-126">Метод</span><span class="sxs-lookup"><span data-stu-id="b771c-126">Method</span></span> |
| [<span data-ttu-id="b771c-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="b771c-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="b771c-128">Метод</span><span class="sxs-lookup"><span data-stu-id="b771c-128">Method</span></span> |
| [<span data-ttu-id="b771c-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="b771c-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="b771c-130">Метод</span><span class="sxs-lookup"><span data-stu-id="b771c-130">Method</span></span> |
| [<span data-ttu-id="b771c-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="b771c-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="b771c-132">Метод</span><span class="sxs-lookup"><span data-stu-id="b771c-132">Method</span></span> |
| [<span data-ttu-id="b771c-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="b771c-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="b771c-134">Метод</span><span class="sxs-lookup"><span data-stu-id="b771c-134">Method</span></span> |
| [<span data-ttu-id="b771c-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="b771c-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="b771c-136">Метод</span><span class="sxs-lookup"><span data-stu-id="b771c-136">Method</span></span> |
| [<span data-ttu-id="b771c-137">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="b771c-137">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="b771c-138">Метод</span><span class="sxs-lookup"><span data-stu-id="b771c-138">Method</span></span> |
| [<span data-ttu-id="b771c-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="b771c-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="b771c-140">Метод</span><span class="sxs-lookup"><span data-stu-id="b771c-140">Method</span></span> |
| [<span data-ttu-id="b771c-141">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="b771c-141">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="b771c-142">Метод</span><span class="sxs-lookup"><span data-stu-id="b771c-142">Method</span></span> |
| [<span data-ttu-id="b771c-143">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="b771c-143">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="b771c-144">Метод</span><span class="sxs-lookup"><span data-stu-id="b771c-144">Method</span></span> |
| [<span data-ttu-id="b771c-145">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="b771c-145">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="b771c-146">Метод</span><span class="sxs-lookup"><span data-stu-id="b771c-146">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="b771c-147">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="b771c-147">Namespaces</span></span>

<span data-ttu-id="b771c-148">[diagnostics](Office.context.mailbox.diagnostics.md). Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="b771c-148">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="b771c-149">[item](Office.context.mailbox.item.md). Предоставляет методы и свойства для доступа к сообщению или встрече в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="b771c-149">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="b771c-150">[userProfile](Office.context.mailbox.userProfile.md). Предоставляет сведения о пользователе в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="b771c-150">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="b771c-151">Members</span><span class="sxs-lookup"><span data-stu-id="b771c-151">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="b771c-152">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="b771c-152">ewsUrl: String</span></span>

<span data-ttu-id="b771c-p101">Получает URL-адрес конечной точки веб-служб Exchange (EWS) для этой учетной записи электронной почты. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="b771c-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="b771c-155">Этот элемент не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="b771c-155">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="b771c-p102">Удаленная служба может использовать значение `ewsUrl`, чтобы выполнять вызовы EWS для почтового ящика пользователя. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="b771c-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="b771c-158">Чтобы вызвать элемент `ewsUrl` в режиме чтения, в манифесте приложения должно быть указано разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="b771c-158">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="b771c-p103">Перед использованием элемента `ewsUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="b771c-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="b771c-161">Тип</span><span class="sxs-lookup"><span data-stu-id="b771c-161">Type</span></span>

*   <span data-ttu-id="b771c-162">String</span><span class="sxs-lookup"><span data-stu-id="b771c-162">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b771c-163">Требования</span><span class="sxs-lookup"><span data-stu-id="b771c-163">Requirements</span></span>

|<span data-ttu-id="b771c-164">Требование</span><span class="sxs-lookup"><span data-stu-id="b771c-164">Requirement</span></span>| <span data-ttu-id="b771c-165">Значение</span><span class="sxs-lookup"><span data-stu-id="b771c-165">Value</span></span>|
|---|---|
|[<span data-ttu-id="b771c-166">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b771c-166">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b771c-167">1.0</span><span class="sxs-lookup"><span data-stu-id="b771c-167">1.0</span></span>|
|[<span data-ttu-id="b771c-168">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b771c-168">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b771c-169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b771c-169">ReadItem</span></span>|
|[<span data-ttu-id="b771c-170">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b771c-170">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b771c-171">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b771c-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="resturl-string"></a><span data-ttu-id="b771c-172">restUrl: String</span><span class="sxs-lookup"><span data-stu-id="b771c-172">restUrl: String</span></span>

<span data-ttu-id="b771c-173">Возвращает URL-адрес конечной точки REST для этой учетной записи электронной почты.</span><span class="sxs-lookup"><span data-stu-id="b771c-173">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="b771c-174">С помощью значения `restUrl` можно выполнять вызовы [REST API](/outlook/rest/) для почтового ящика пользователя.</span><span class="sxs-lookup"><span data-stu-id="b771c-174">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="b771c-175">Клиенты Outlook, подключенные к локальным установленным версиям Exchange 2016 или более поздним с пользовательским URL-адресом REST, возвращают недопустимое значение `restUrl`.</span><span class="sxs-lookup"><span data-stu-id="b771c-175">Outlook clients connected to on-premises installations of Exchange 2016 or later with a custom REST URL configured will return an invalid value for `restUrl`.</span></span>

##### <a name="type"></a><span data-ttu-id="b771c-176">Тип</span><span class="sxs-lookup"><span data-stu-id="b771c-176">Type</span></span>

*   <span data-ttu-id="b771c-177">String</span><span class="sxs-lookup"><span data-stu-id="b771c-177">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b771c-178">Требования</span><span class="sxs-lookup"><span data-stu-id="b771c-178">Requirements</span></span>

|<span data-ttu-id="b771c-179">Требование</span><span class="sxs-lookup"><span data-stu-id="b771c-179">Requirement</span></span>| <span data-ttu-id="b771c-180">Значение</span><span class="sxs-lookup"><span data-stu-id="b771c-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="b771c-181">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b771c-181">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b771c-182">1.5</span><span class="sxs-lookup"><span data-stu-id="b771c-182">1.5</span></span> |
|[<span data-ttu-id="b771c-183">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b771c-183">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b771c-184">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b771c-184">ReadItem</span></span>|
|[<span data-ttu-id="b771c-185">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b771c-185">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b771c-186">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b771c-186">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="b771c-187">Методы</span><span class="sxs-lookup"><span data-stu-id="b771c-187">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="b771c-188">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="b771c-188">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="b771c-189">Добавляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="b771c-189">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="b771c-190">В настоящее время поддерживаются только события типа `Office.EventType.ItemChanged`, которые вызываются, когда пользователь выбирает новый элемент.</span><span class="sxs-lookup"><span data-stu-id="b771c-190">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item.</span></span> <span data-ttu-id="b771c-191">Это событие используется надстройками, реализующими закрепляемую область задач, и позволяет надстройке обновлять пользовательский интерфейс области задач в соответствии с выбранным в данный момент элементом.</span><span class="sxs-lookup"><span data-stu-id="b771c-191">This event is used by add-ins that implement a pinnable task pane, and allows the add-in to refresh the task pane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b771c-192">Параметры</span><span class="sxs-lookup"><span data-stu-id="b771c-192">Parameters</span></span>

| <span data-ttu-id="b771c-193">Имя</span><span class="sxs-lookup"><span data-stu-id="b771c-193">Name</span></span> | <span data-ttu-id="b771c-194">Тип</span><span class="sxs-lookup"><span data-stu-id="b771c-194">Type</span></span> | <span data-ttu-id="b771c-195">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="b771c-195">Attributes</span></span> | <span data-ttu-id="b771c-196">Описание</span><span class="sxs-lookup"><span data-stu-id="b771c-196">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="b771c-197">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="b771c-197">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="b771c-198">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="b771c-198">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="b771c-199">Function</span><span class="sxs-lookup"><span data-stu-id="b771c-199">Function</span></span> || <span data-ttu-id="b771c-p105">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="b771c-p105">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="b771c-203">Объект</span><span class="sxs-lookup"><span data-stu-id="b771c-203">Object</span></span> | <span data-ttu-id="b771c-204">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b771c-204">&lt;optional&gt;</span></span> | <span data-ttu-id="b771c-205">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="b771c-205">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="b771c-206">Object</span><span class="sxs-lookup"><span data-stu-id="b771c-206">Object</span></span> | <span data-ttu-id="b771c-207">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b771c-207">&lt;optional&gt;</span></span> | <span data-ttu-id="b771c-208">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="b771c-208">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="b771c-209">функция</span><span class="sxs-lookup"><span data-stu-id="b771c-209">function</span></span>| <span data-ttu-id="b771c-210">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b771c-210">&lt;optional&gt;</span></span>|<span data-ttu-id="b771c-211">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b771c-211">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b771c-212">Требования</span><span class="sxs-lookup"><span data-stu-id="b771c-212">Requirements</span></span>

|<span data-ttu-id="b771c-213">Требование</span><span class="sxs-lookup"><span data-stu-id="b771c-213">Requirement</span></span>| <span data-ttu-id="b771c-214">Значение</span><span class="sxs-lookup"><span data-stu-id="b771c-214">Value</span></span>|
|---|---|
|[<span data-ttu-id="b771c-215">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b771c-215">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b771c-216">1.5</span><span class="sxs-lookup"><span data-stu-id="b771c-216">1.5</span></span> |
|[<span data-ttu-id="b771c-217">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b771c-217">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b771c-218">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b771c-218">ReadItem</span></span> |
|[<span data-ttu-id="b771c-219">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b771c-219">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b771c-220">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b771c-220">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b771c-221">Пример</span><span class="sxs-lookup"><span data-stu-id="b771c-221">Example</span></span>

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

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="b771c-222">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="b771c-222">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="b771c-223">Преобразовывает идентификатор элемента из формата REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="b771c-223">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="b771c-224">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="b771c-224">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="b771c-p106">Формат идентификаторов, извлекаемых через API REST (например, [API Почты Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)), отличается от формата веб-служб Exchange (EWS). Метод `convertToEwsId` преобразовывает идентификатор в формате REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="b771c-p106">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b771c-227">Параметры</span><span class="sxs-lookup"><span data-stu-id="b771c-227">Parameters</span></span>

|<span data-ttu-id="b771c-228">Имя</span><span class="sxs-lookup"><span data-stu-id="b771c-228">Name</span></span>| <span data-ttu-id="b771c-229">Тип</span><span class="sxs-lookup"><span data-stu-id="b771c-229">Type</span></span>| <span data-ttu-id="b771c-230">Описание</span><span class="sxs-lookup"><span data-stu-id="b771c-230">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="b771c-231">String</span><span class="sxs-lookup"><span data-stu-id="b771c-231">String</span></span>|<span data-ttu-id="b771c-232">Идентификатор элемента в формате REST API для Outlook</span><span class="sxs-lookup"><span data-stu-id="b771c-232">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="b771c-233">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="b771c-233">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.5)|<span data-ttu-id="b771c-234">Значение, определяющее версию REST API для Outlook, которая используется для извлечения идентификатора элемента.</span><span class="sxs-lookup"><span data-stu-id="b771c-234">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b771c-235">Требования</span><span class="sxs-lookup"><span data-stu-id="b771c-235">Requirements</span></span>

|<span data-ttu-id="b771c-236">Требование</span><span class="sxs-lookup"><span data-stu-id="b771c-236">Requirement</span></span>| <span data-ttu-id="b771c-237">Значение</span><span class="sxs-lookup"><span data-stu-id="b771c-237">Value</span></span>|
|---|---|
|[<span data-ttu-id="b771c-238">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b771c-238">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b771c-239">1.3</span><span class="sxs-lookup"><span data-stu-id="b771c-239">1.3</span></span>|
|[<span data-ttu-id="b771c-240">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b771c-240">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b771c-241">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="b771c-241">Restricted</span></span>|
|[<span data-ttu-id="b771c-242">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b771c-242">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b771c-243">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b771c-243">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b771c-244">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="b771c-244">Returns:</span></span>

<span data-ttu-id="b771c-245">Тип: String</span><span class="sxs-lookup"><span data-stu-id="b771c-245">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="b771c-246">Пример</span><span class="sxs-lookup"><span data-stu-id="b771c-246">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-15"></a><span data-ttu-id="b771c-247">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)}</span><span class="sxs-lookup"><span data-stu-id="b771c-247">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)}</span></span>

<span data-ttu-id="b771c-248">Получает словарь, содержащий сведения о локальном времени клиента.</span><span class="sxs-lookup"><span data-stu-id="b771c-248">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="b771c-p107">Почтовое приложение для классической версии Outlook или версии в Интернете может использовать разные часовые пояса для дат и времени. Классическое приложение Outlook использует часовой пояс клиентского компьютера. Outlook в Интернете использует часовой пояс, заданный в Центре администрирования Exchange (EAC). Значения даты и времени должны обрабатываться так, чтобы значения в пользовательском интерфейсе всегда согласовывались с часовым поясом, ожидаемым пользователем.</span><span class="sxs-lookup"><span data-stu-id="b771c-p107">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="b771c-p108">Если почтовое приложение работает в классическом клиенте Outlook, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса клиентского компьютера. Если почтовое приложение работает в Outlook в Интернете, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса, заданного в Центре администрирования Exchange.</span><span class="sxs-lookup"><span data-stu-id="b771c-p108">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b771c-254">Параметры</span><span class="sxs-lookup"><span data-stu-id="b771c-254">Parameters</span></span>

|<span data-ttu-id="b771c-255">Имя</span><span class="sxs-lookup"><span data-stu-id="b771c-255">Name</span></span>| <span data-ttu-id="b771c-256">Тип</span><span class="sxs-lookup"><span data-stu-id="b771c-256">Type</span></span>| <span data-ttu-id="b771c-257">Описание</span><span class="sxs-lookup"><span data-stu-id="b771c-257">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="b771c-258">Date</span><span class="sxs-lookup"><span data-stu-id="b771c-258">Date</span></span>|<span data-ttu-id="b771c-259">Объект Date</span><span class="sxs-lookup"><span data-stu-id="b771c-259">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b771c-260">Требования</span><span class="sxs-lookup"><span data-stu-id="b771c-260">Requirements</span></span>

|<span data-ttu-id="b771c-261">Требование</span><span class="sxs-lookup"><span data-stu-id="b771c-261">Requirement</span></span>| <span data-ttu-id="b771c-262">Значение</span><span class="sxs-lookup"><span data-stu-id="b771c-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="b771c-263">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b771c-263">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b771c-264">1.0</span><span class="sxs-lookup"><span data-stu-id="b771c-264">1.0</span></span>|
|[<span data-ttu-id="b771c-265">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b771c-265">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b771c-266">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b771c-266">ReadItem</span></span>|
|[<span data-ttu-id="b771c-267">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b771c-267">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b771c-268">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b771c-268">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b771c-269">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="b771c-269">Returns:</span></span>

<span data-ttu-id="b771c-270">Тип: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="b771c-270">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="b771c-271">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="b771c-271">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="b771c-272">Преобразовывает идентификатор элемента в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="b771c-272">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="b771c-273">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="b771c-273">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="b771c-p109">Формат идентификаторов, извлекаемых через EWS или свойство `itemId`, отличается от формата API REST (таких как [API Почты Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)). Метод `convertToRestId` преобразовывает идентификатор в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="b771c-p109">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b771c-276">Параметры</span><span class="sxs-lookup"><span data-stu-id="b771c-276">Parameters</span></span>

|<span data-ttu-id="b771c-277">Имя</span><span class="sxs-lookup"><span data-stu-id="b771c-277">Name</span></span>| <span data-ttu-id="b771c-278">Тип</span><span class="sxs-lookup"><span data-stu-id="b771c-278">Type</span></span>| <span data-ttu-id="b771c-279">Описание</span><span class="sxs-lookup"><span data-stu-id="b771c-279">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="b771c-280">String</span><span class="sxs-lookup"><span data-stu-id="b771c-280">String</span></span>|<span data-ttu-id="b771c-281">Идентификатор элемента в формате EWS</span><span class="sxs-lookup"><span data-stu-id="b771c-281">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="b771c-282">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="b771c-282">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.5)|<span data-ttu-id="b771c-283">Значение, определяющее версию REST API для Outlook, с которой будет использоваться преобразованный идентификатор.</span><span class="sxs-lookup"><span data-stu-id="b771c-283">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b771c-284">Требования</span><span class="sxs-lookup"><span data-stu-id="b771c-284">Requirements</span></span>

|<span data-ttu-id="b771c-285">Требование</span><span class="sxs-lookup"><span data-stu-id="b771c-285">Requirement</span></span>| <span data-ttu-id="b771c-286">Значение</span><span class="sxs-lookup"><span data-stu-id="b771c-286">Value</span></span>|
|---|---|
|[<span data-ttu-id="b771c-287">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b771c-287">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b771c-288">1.3</span><span class="sxs-lookup"><span data-stu-id="b771c-288">1.3</span></span>|
|[<span data-ttu-id="b771c-289">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b771c-289">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b771c-290">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="b771c-290">Restricted</span></span>|
|[<span data-ttu-id="b771c-291">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b771c-291">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b771c-292">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b771c-292">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b771c-293">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="b771c-293">Returns:</span></span>

<span data-ttu-id="b771c-294">Тип: String</span><span class="sxs-lookup"><span data-stu-id="b771c-294">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="b771c-295">Пример</span><span class="sxs-lookup"><span data-stu-id="b771c-295">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="b771c-296">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="b771c-296">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="b771c-297">Получает объект Date из словаря, содержащего сведения о времени.</span><span class="sxs-lookup"><span data-stu-id="b771c-297">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="b771c-298">Метод `convertToUtcClientTime` преобразует словарь, содержащий локальную дату и время, в объект Date с правильными значениями локальной даты и времени.</span><span class="sxs-lookup"><span data-stu-id="b771c-298">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b771c-299">Параметры</span><span class="sxs-lookup"><span data-stu-id="b771c-299">Parameters</span></span>

|<span data-ttu-id="b771c-300">Имя</span><span class="sxs-lookup"><span data-stu-id="b771c-300">Name</span></span>| <span data-ttu-id="b771c-301">Тип</span><span class="sxs-lookup"><span data-stu-id="b771c-301">Type</span></span>| <span data-ttu-id="b771c-302">Описание</span><span class="sxs-lookup"><span data-stu-id="b771c-302">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="b771c-303">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="b771c-303">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)|<span data-ttu-id="b771c-304">Значение локального времени для преобразования.</span><span class="sxs-lookup"><span data-stu-id="b771c-304">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b771c-305">Требования</span><span class="sxs-lookup"><span data-stu-id="b771c-305">Requirements</span></span>

|<span data-ttu-id="b771c-306">Требование</span><span class="sxs-lookup"><span data-stu-id="b771c-306">Requirement</span></span>| <span data-ttu-id="b771c-307">Значение</span><span class="sxs-lookup"><span data-stu-id="b771c-307">Value</span></span>|
|---|---|
|[<span data-ttu-id="b771c-308">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b771c-308">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b771c-309">1.0</span><span class="sxs-lookup"><span data-stu-id="b771c-309">1.0</span></span>|
|[<span data-ttu-id="b771c-310">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b771c-310">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b771c-311">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b771c-311">ReadItem</span></span>|
|[<span data-ttu-id="b771c-312">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b771c-312">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b771c-313">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b771c-313">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="b771c-314">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="b771c-314">Returns:</span></span>

<span data-ttu-id="b771c-315">Объект Date со временем в формате UTC.</span><span class="sxs-lookup"><span data-stu-id="b771c-315">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="b771c-316">Тип: Date</span><span class="sxs-lookup"><span data-stu-id="b771c-316">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="b771c-317">Пример</span><span class="sxs-lookup"><span data-stu-id="b771c-317">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="b771c-318">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="b771c-318">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="b771c-319">Отображает имеющуюся встречу из календаря.</span><span class="sxs-lookup"><span data-stu-id="b771c-319">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="b771c-320">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="b771c-320">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="b771c-321">Метод `displayAppointmentForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее сведения календаря о существующей встрече.</span><span class="sxs-lookup"><span data-stu-id="b771c-321">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="b771c-p110">В Outlook для Mac с помощью этого метода можно отобразить одну встречу, которая не является частью повторяющегося ряда, или основную встречу такого ряда, но не экземпляр из него, так как в Outlook для Mac невозможно получить доступ к свойствам экземпляра повторяющегося ряда (в том числе к идентификатору элемента).</span><span class="sxs-lookup"><span data-stu-id="b771c-p110">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="b771c-324">В Outlook в Интернете этот метод открывает указанную форму, только если текст формы содержит символы размером не более 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="b771c-324">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="b771c-325">Если указанный идентификатор элемента не определяет существующую встречу, на клиентском компьютере или устройстве открывается пустая страница, и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="b771c-325">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b771c-326">Параметры</span><span class="sxs-lookup"><span data-stu-id="b771c-326">Parameters</span></span>

|<span data-ttu-id="b771c-327">Имя</span><span class="sxs-lookup"><span data-stu-id="b771c-327">Name</span></span>| <span data-ttu-id="b771c-328">Тип</span><span class="sxs-lookup"><span data-stu-id="b771c-328">Type</span></span>| <span data-ttu-id="b771c-329">Описание</span><span class="sxs-lookup"><span data-stu-id="b771c-329">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="b771c-330">String</span><span class="sxs-lookup"><span data-stu-id="b771c-330">String</span></span>|<span data-ttu-id="b771c-331">Идентификатор веб-служб Exchange для существующей встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="b771c-331">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b771c-332">Требования</span><span class="sxs-lookup"><span data-stu-id="b771c-332">Requirements</span></span>

|<span data-ttu-id="b771c-333">Требование</span><span class="sxs-lookup"><span data-stu-id="b771c-333">Requirement</span></span>| <span data-ttu-id="b771c-334">Значение</span><span class="sxs-lookup"><span data-stu-id="b771c-334">Value</span></span>|
|---|---|
|[<span data-ttu-id="b771c-335">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b771c-335">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b771c-336">1.0</span><span class="sxs-lookup"><span data-stu-id="b771c-336">1.0</span></span>|
|[<span data-ttu-id="b771c-337">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b771c-337">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b771c-338">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b771c-338">ReadItem</span></span>|
|[<span data-ttu-id="b771c-339">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b771c-339">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b771c-340">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b771c-340">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b771c-341">Пример</span><span class="sxs-lookup"><span data-stu-id="b771c-341">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="b771c-342">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="b771c-342">displayMessageForm(itemId)</span></span>

<span data-ttu-id="b771c-343">Отображает имеющееся сообщение.</span><span class="sxs-lookup"><span data-stu-id="b771c-343">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="b771c-344">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="b771c-344">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="b771c-345">Метод `displayMessageForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее существующее сообщение.</span><span class="sxs-lookup"><span data-stu-id="b771c-345">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="b771c-346">В Outlook в Интернете этот метод открывает указанную форму, только если текст формы содержит символы размером не более 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="b771c-346">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="b771c-347">Если указанный идентификатор элемента не определяет существующее сообщение, окно на клиентском компьютере не открывается и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="b771c-347">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="b771c-p111">Не используйте `displayMessageForm` с параметром `itemId`, который представляет собой встречу. Используйте метод `displayAppointmentForm`, чтобы отобразить сведения о существующей встрече, а метод `displayNewAppointmentForm` — для отображения формы создания встречи.</span><span class="sxs-lookup"><span data-stu-id="b771c-p111">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b771c-350">Параметры</span><span class="sxs-lookup"><span data-stu-id="b771c-350">Parameters</span></span>

|<span data-ttu-id="b771c-351">Имя</span><span class="sxs-lookup"><span data-stu-id="b771c-351">Name</span></span>| <span data-ttu-id="b771c-352">Тип</span><span class="sxs-lookup"><span data-stu-id="b771c-352">Type</span></span>| <span data-ttu-id="b771c-353">Описание</span><span class="sxs-lookup"><span data-stu-id="b771c-353">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="b771c-354">String</span><span class="sxs-lookup"><span data-stu-id="b771c-354">String</span></span>|<span data-ttu-id="b771c-355">Идентификатор веб-служб Exchange для существующего сообщения.</span><span class="sxs-lookup"><span data-stu-id="b771c-355">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b771c-356">Требования</span><span class="sxs-lookup"><span data-stu-id="b771c-356">Requirements</span></span>

|<span data-ttu-id="b771c-357">Требование</span><span class="sxs-lookup"><span data-stu-id="b771c-357">Requirement</span></span>| <span data-ttu-id="b771c-358">Значение</span><span class="sxs-lookup"><span data-stu-id="b771c-358">Value</span></span>|
|---|---|
|[<span data-ttu-id="b771c-359">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b771c-359">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b771c-360">1.0</span><span class="sxs-lookup"><span data-stu-id="b771c-360">1.0</span></span>|
|[<span data-ttu-id="b771c-361">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b771c-361">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b771c-362">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b771c-362">ReadItem</span></span>|
|[<span data-ttu-id="b771c-363">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b771c-363">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b771c-364">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b771c-364">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b771c-365">Пример</span><span class="sxs-lookup"><span data-stu-id="b771c-365">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="b771c-366">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="b771c-366">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="b771c-367">Отображает форму для создания новой встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="b771c-367">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="b771c-368">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="b771c-368">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="b771c-p112">Метод `displayNewAppointmentForm` открывает форму, в которой пользователь может создать встречу или собрание. Если параметры заданы, поля формы встречи автоматически заполняются их содержимым.</span><span class="sxs-lookup"><span data-stu-id="b771c-p112">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="b771c-p113">В Outlook в Интернете и на мобильных устройствах этот метод всегда отображает форму с полем участников. Если вы не укажете участников в качестве входных аргументов, метод отображает форму с кнопкой **Сохранить**. Если вы укажете участников, форма будет включать участников и кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="b771c-p113">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="b771c-p114">Если вы укажете участников или ресурсы с помощью параметра `requiredAttendees`, `optionalAttendees` или `resources` в клиенте Outlook с расширенными возможностями и Outlook RT, этот метод отобразит форму собрания с кнопкой **Отправить**. Если не указать получателей, этот метод отобразит форму встречи с кнопкой **Сохранить и закрыть**.</span><span class="sxs-lookup"><span data-stu-id="b771c-p114">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="b771c-376">Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.</span><span class="sxs-lookup"><span data-stu-id="b771c-376">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b771c-377">Параметры</span><span class="sxs-lookup"><span data-stu-id="b771c-377">Parameters</span></span>

|<span data-ttu-id="b771c-378">Имя</span><span class="sxs-lookup"><span data-stu-id="b771c-378">Name</span></span>| <span data-ttu-id="b771c-379">Тип</span><span class="sxs-lookup"><span data-stu-id="b771c-379">Type</span></span>| <span data-ttu-id="b771c-380">Описание</span><span class="sxs-lookup"><span data-stu-id="b771c-380">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="b771c-381">Object</span><span class="sxs-lookup"><span data-stu-id="b771c-381">Object</span></span> | <span data-ttu-id="b771c-382">Словарь параметров, описывающий новую встречу.</span><span class="sxs-lookup"><span data-stu-id="b771c-382">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="b771c-383">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span><span class="sxs-lookup"><span data-stu-id="b771c-383">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span></span> | <span data-ttu-id="b771c-p115">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из обязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="b771c-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="b771c-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span><span class="sxs-lookup"><span data-stu-id="b771c-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span></span> | <span data-ttu-id="b771c-p116">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из необязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="b771c-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="b771c-389">Date</span><span class="sxs-lookup"><span data-stu-id="b771c-389">Date</span></span> | <span data-ttu-id="b771c-390">Объект `Date`, указывающий дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="b771c-390">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="b771c-391">Date</span><span class="sxs-lookup"><span data-stu-id="b771c-391">Date</span></span> | <span data-ttu-id="b771c-392">Объект `Date`, указывающий дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="b771c-392">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="b771c-393">String</span><span class="sxs-lookup"><span data-stu-id="b771c-393">String</span></span> | <span data-ttu-id="b771c-p117">Строка со сведениями о месте встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="b771c-p117">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="b771c-396">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="b771c-396">Array.&lt;String&gt;</span></span> | <span data-ttu-id="b771c-p118">Массив строк, содержащий необходимые для встречи ресурсы. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="b771c-p118">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="b771c-399">String</span><span class="sxs-lookup"><span data-stu-id="b771c-399">String</span></span> | <span data-ttu-id="b771c-p119">Строка с темой встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="b771c-p119">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="b771c-402">String</span><span class="sxs-lookup"><span data-stu-id="b771c-402">String</span></span> | <span data-ttu-id="b771c-p120">Текст сообщения о встрече. Максимальный размер содержимого сообщения — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="b771c-p120">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="b771c-405">Требования</span><span class="sxs-lookup"><span data-stu-id="b771c-405">Requirements</span></span>

|<span data-ttu-id="b771c-406">Требование</span><span class="sxs-lookup"><span data-stu-id="b771c-406">Requirement</span></span>| <span data-ttu-id="b771c-407">Значение</span><span class="sxs-lookup"><span data-stu-id="b771c-407">Value</span></span>|
|---|---|
|[<span data-ttu-id="b771c-408">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b771c-408">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b771c-409">1.0</span><span class="sxs-lookup"><span data-stu-id="b771c-409">1.0</span></span>|
|[<span data-ttu-id="b771c-410">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b771c-410">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b771c-411">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b771c-411">ReadItem</span></span>|
|[<span data-ttu-id="b771c-412">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b771c-412">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b771c-413">Чтение</span><span class="sxs-lookup"><span data-stu-id="b771c-413">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b771c-414">Пример</span><span class="sxs-lookup"><span data-stu-id="b771c-414">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="b771c-415">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="b771c-415">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="b771c-416">Возвращает строку, содержащую маркер, который используется для вызова интерфейсов REST API или веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="b771c-416">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="b771c-p121">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="b771c-p121">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="b771c-419">Рекомендуем сделать так, чтобы по мере возможности надстройки использовали интерфейсы REST API, а не веб-службы Exchange.</span><span class="sxs-lookup"><span data-stu-id="b771c-419">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="b771c-420">Для вызова метода `getCallbackTokenAsync` в режиме чтения требуется минимальный уровень разрешения **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="b771c-420">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="b771c-421">Для вызова `getCallbackTokenAsync` в режиме создания сообщения требуется сохранить элемент.</span><span class="sxs-lookup"><span data-stu-id="b771c-421">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="b771c-422">Для метода [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) требуется минимальный уровень разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="b771c-422">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

<span data-ttu-id="b771c-423">**Маркеры REST**</span><span class="sxs-lookup"><span data-stu-id="b771c-423">**REST Tokens**</span></span>

<span data-ttu-id="b771c-p123">Если запрашивается маркер REST (`options.isRest = true`), полученный маркер не подойдет для проверки подлинности при вызовах веб-служб Exchange. Область действия маркера будет ограничена доступом только для чтения к текущему элементу и его вложениям, если в манифесте надстройки не указано разрешение [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission). Если указано разрешение `ReadWriteMailbox`, полученный маркер предоставит доступ на чтение и запись к почте, календарю и контактам, включая возможность отправки почты.</span><span class="sxs-lookup"><span data-stu-id="b771c-p123">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="b771c-427">С помощью свойства `restUrl` надстройка должна определить правильный URL-адрес для вызовов REST API.</span><span class="sxs-lookup"><span data-stu-id="b771c-427">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="b771c-428">**Маркеры EWS**</span><span class="sxs-lookup"><span data-stu-id="b771c-428">**EWS Tokens**</span></span>

<span data-ttu-id="b771c-p124">Если запрашивается маркер EWS (`options.isRest = false`), полученный маркер не подойдет для проверки подлинности при вызовах REST API. Область действия маркера будет ограничена доступом к текущему элементу.</span><span class="sxs-lookup"><span data-stu-id="b771c-p124">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="b771c-431">С помощью свойства `ewsUrl` надстройка должна определить правильный URL-адрес для вызовов EWS.</span><span class="sxs-lookup"><span data-stu-id="b771c-431">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

<span data-ttu-id="b771c-432">Вы можете передать сторонней системе маркер и идентификатор вложения или элемента.</span><span class="sxs-lookup"><span data-stu-id="b771c-432">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="b771c-433">Сторонняя система использует этот маркер как маркер авторизации, чтобы вызвать операцию [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) или [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) веб-служб Exchange для возврата вложения или элемента.</span><span class="sxs-lookup"><span data-stu-id="b771c-433">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="b771c-434">Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="b771c-434">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="parameters"></a><span data-ttu-id="b771c-435">Параметры</span><span class="sxs-lookup"><span data-stu-id="b771c-435">Parameters</span></span>

|<span data-ttu-id="b771c-436">Имя</span><span class="sxs-lookup"><span data-stu-id="b771c-436">Name</span></span>| <span data-ttu-id="b771c-437">Тип</span><span class="sxs-lookup"><span data-stu-id="b771c-437">Type</span></span>| <span data-ttu-id="b771c-438">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="b771c-438">Attributes</span></span>| <span data-ttu-id="b771c-439">Описание</span><span class="sxs-lookup"><span data-stu-id="b771c-439">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="b771c-440">Object</span><span class="sxs-lookup"><span data-stu-id="b771c-440">Object</span></span> | <span data-ttu-id="b771c-441">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b771c-441">&lt;optional&gt;</span></span> | <span data-ttu-id="b771c-442">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="b771c-442">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="b771c-443">Boolean</span><span class="sxs-lookup"><span data-stu-id="b771c-443">Boolean</span></span> |  <span data-ttu-id="b771c-444">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b771c-444">&lt;optional&gt;</span></span> | <span data-ttu-id="b771c-p126">Определяет, будет ли предоставленный маркер использоваться для интерфейсов REST API Outlook или веб-служб Exchange. Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="b771c-p126">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="b771c-447">Объект</span><span class="sxs-lookup"><span data-stu-id="b771c-447">Object</span></span> |  <span data-ttu-id="b771c-448">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b771c-448">&lt;optional&gt;</span></span> | <span data-ttu-id="b771c-449">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="b771c-449">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="b771c-450">функция</span><span class="sxs-lookup"><span data-stu-id="b771c-450">function</span></span>||<span data-ttu-id="b771c-451">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b771c-451">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="b771c-452">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="b771c-452">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="b771c-453">При наличии ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут предоставлять дополнительные сведения.</span><span class="sxs-lookup"><span data-stu-id="b771c-453">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="b771c-454">Ошибки</span><span class="sxs-lookup"><span data-stu-id="b771c-454">Errors</span></span>

|<span data-ttu-id="b771c-455">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="b771c-455">Error code</span></span>|<span data-ttu-id="b771c-456">Описание</span><span class="sxs-lookup"><span data-stu-id="b771c-456">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="b771c-457">Не удалось выполнить запрос.</span><span class="sxs-lookup"><span data-stu-id="b771c-457">The request has failed.</span></span> <span data-ttu-id="b771c-458">Просмотрите объект диагностики для кода ошибки HTTP.</span><span class="sxs-lookup"><span data-stu-id="b771c-458">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="b771c-459">Сервер Exchange Server вернул ошибку.</span><span class="sxs-lookup"><span data-stu-id="b771c-459">The Exchange server returned an error.</span></span> <span data-ttu-id="b771c-460">Для получения дополнительных сведений просмотрите объект диагностики.</span><span class="sxs-lookup"><span data-stu-id="b771c-460">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="b771c-461">Пользователь отключен от сети.</span><span class="sxs-lookup"><span data-stu-id="b771c-461">The user is no longer connected to the network.</span></span> <span data-ttu-id="b771c-462">Проверьте сетевое подключение и повторите попытку.</span><span class="sxs-lookup"><span data-stu-id="b771c-462">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b771c-463">Требования</span><span class="sxs-lookup"><span data-stu-id="b771c-463">Requirements</span></span>

|<span data-ttu-id="b771c-464">Требование</span><span class="sxs-lookup"><span data-stu-id="b771c-464">Requirement</span></span>| <span data-ttu-id="b771c-465">Значение</span><span class="sxs-lookup"><span data-stu-id="b771c-465">Value</span></span>|
|---|---|
|[<span data-ttu-id="b771c-466">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b771c-466">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b771c-467">1.5</span><span class="sxs-lookup"><span data-stu-id="b771c-467">1.5</span></span> |
|[<span data-ttu-id="b771c-468">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b771c-468">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b771c-469">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b771c-469">ReadItem</span></span>|
|[<span data-ttu-id="b771c-470">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b771c-470">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b771c-471">Создание и чтение</span><span class="sxs-lookup"><span data-stu-id="b771c-471">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="b771c-472">Пример</span><span class="sxs-lookup"><span data-stu-id="b771c-472">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="b771c-473">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="b771c-473">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="b771c-474">Получает строку, содержащую маркер, используемый для получения вложения или элемента с Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="b771c-474">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="b771c-p130">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="b771c-p130">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="b771c-477">Вы можете передать сторонней системе маркер и идентификатор вложения или элемента.</span><span class="sxs-lookup"><span data-stu-id="b771c-477">You can pass both the token and either an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="b771c-478">Сторонняя система использует этот маркер как маркер авторизации, чтобы вызвать операцию [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) или [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) веб-служб Exchange для возврата вложения или элемента.</span><span class="sxs-lookup"><span data-stu-id="b771c-478">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) operation or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="b771c-479">Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="b771c-479">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="b771c-480">Для вызова метода `getCallbackTokenAsync` в режиме чтения требуется минимальный уровень разрешения **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="b771c-480">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="b771c-481">Для вызова `getCallbackTokenAsync` в режиме создания сообщения требуется сохранить элемент.</span><span class="sxs-lookup"><span data-stu-id="b771c-481">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="b771c-482">Для метода [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) требуется минимальный уровень разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="b771c-482">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b771c-483">Параметры</span><span class="sxs-lookup"><span data-stu-id="b771c-483">Parameters</span></span>

|<span data-ttu-id="b771c-484">Имя</span><span class="sxs-lookup"><span data-stu-id="b771c-484">Name</span></span>| <span data-ttu-id="b771c-485">Тип</span><span class="sxs-lookup"><span data-stu-id="b771c-485">Type</span></span>| <span data-ttu-id="b771c-486">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="b771c-486">Attributes</span></span>| <span data-ttu-id="b771c-487">Описание</span><span class="sxs-lookup"><span data-stu-id="b771c-487">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="b771c-488">function</span><span class="sxs-lookup"><span data-stu-id="b771c-488">function</span></span>||<span data-ttu-id="b771c-489">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b771c-489">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="b771c-490">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="b771c-490">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="b771c-491">При наличии ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут предоставлять дополнительные сведения.</span><span class="sxs-lookup"><span data-stu-id="b771c-491">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="b771c-492">Объект</span><span class="sxs-lookup"><span data-stu-id="b771c-492">Object</span></span>| <span data-ttu-id="b771c-493">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b771c-493">&lt;optional&gt;</span></span>|<span data-ttu-id="b771c-494">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="b771c-494">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="b771c-495">Ошибки</span><span class="sxs-lookup"><span data-stu-id="b771c-495">Errors</span></span>

|<span data-ttu-id="b771c-496">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="b771c-496">Error code</span></span>|<span data-ttu-id="b771c-497">Описание</span><span class="sxs-lookup"><span data-stu-id="b771c-497">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="b771c-498">Не удалось выполнить запрос.</span><span class="sxs-lookup"><span data-stu-id="b771c-498">The request has failed.</span></span> <span data-ttu-id="b771c-499">Просмотрите объект диагностики для кода ошибки HTTP.</span><span class="sxs-lookup"><span data-stu-id="b771c-499">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="b771c-500">Сервер Exchange Server вернул ошибку.</span><span class="sxs-lookup"><span data-stu-id="b771c-500">The Exchange server returned an error.</span></span> <span data-ttu-id="b771c-501">Для получения дополнительных сведений просмотрите объект диагностики.</span><span class="sxs-lookup"><span data-stu-id="b771c-501">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="b771c-502">Пользователь отключен от сети.</span><span class="sxs-lookup"><span data-stu-id="b771c-502">The user is no longer connected to the network.</span></span> <span data-ttu-id="b771c-503">Проверьте сетевое подключение и повторите попытку.</span><span class="sxs-lookup"><span data-stu-id="b771c-503">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b771c-504">Требования</span><span class="sxs-lookup"><span data-stu-id="b771c-504">Requirements</span></span>

|<span data-ttu-id="b771c-505">Требование</span><span class="sxs-lookup"><span data-stu-id="b771c-505">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="b771c-506">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b771c-506">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b771c-507">1.0</span><span class="sxs-lookup"><span data-stu-id="b771c-507">1.0</span></span> | <span data-ttu-id="b771c-508">1.3</span><span class="sxs-lookup"><span data-stu-id="b771c-508">1.3</span></span> |
|[<span data-ttu-id="b771c-509">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b771c-509">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b771c-510">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b771c-510">ReadItem</span></span> | <span data-ttu-id="b771c-511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b771c-511">ReadItem</span></span> |
|[<span data-ttu-id="b771c-512">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b771c-512">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b771c-513">Чтение</span><span class="sxs-lookup"><span data-stu-id="b771c-513">Read</span></span> | <span data-ttu-id="b771c-514">Создание</span><span class="sxs-lookup"><span data-stu-id="b771c-514">Compose</span></span> |

##### <a name="example"></a><span data-ttu-id="b771c-515">Пример</span><span class="sxs-lookup"><span data-stu-id="b771c-515">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="b771c-516">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="b771c-516">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="b771c-517">Получает маркер, идентифицирующий пользователя и надстройку Office.</span><span class="sxs-lookup"><span data-stu-id="b771c-517">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="b771c-518">Метод `getUserIdentityTokenAsync` возвращает маркер, который можно использовать для идентификации, а также [проверки подлинности надстройки и пользователя в сторонней системе](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="b771c-518">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="b771c-519">Параметры</span><span class="sxs-lookup"><span data-stu-id="b771c-519">Parameters</span></span>

|<span data-ttu-id="b771c-520">Имя</span><span class="sxs-lookup"><span data-stu-id="b771c-520">Name</span></span>| <span data-ttu-id="b771c-521">Тип</span><span class="sxs-lookup"><span data-stu-id="b771c-521">Type</span></span>| <span data-ttu-id="b771c-522">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="b771c-522">Attributes</span></span>| <span data-ttu-id="b771c-523">Описание</span><span class="sxs-lookup"><span data-stu-id="b771c-523">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="b771c-524">function</span><span class="sxs-lookup"><span data-stu-id="b771c-524">function</span></span>||<span data-ttu-id="b771c-525">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b771c-525">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="b771c-526">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="b771c-526">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="b771c-527">При наличии ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут предоставлять дополнительные сведения.</span><span class="sxs-lookup"><span data-stu-id="b771c-527">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="b771c-528">Объект</span><span class="sxs-lookup"><span data-stu-id="b771c-528">Object</span></span>| <span data-ttu-id="b771c-529">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b771c-529">&lt;optional&gt;</span></span>|<span data-ttu-id="b771c-530">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="b771c-530">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="b771c-531">Ошибки</span><span class="sxs-lookup"><span data-stu-id="b771c-531">Errors</span></span>

|<span data-ttu-id="b771c-532">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="b771c-532">Error code</span></span>|<span data-ttu-id="b771c-533">Описание</span><span class="sxs-lookup"><span data-stu-id="b771c-533">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="b771c-534">Не удалось выполнить запрос.</span><span class="sxs-lookup"><span data-stu-id="b771c-534">The request has failed.</span></span> <span data-ttu-id="b771c-535">Просмотрите объект диагностики для кода ошибки HTTP.</span><span class="sxs-lookup"><span data-stu-id="b771c-535">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="b771c-536">Сервер Exchange Server вернул ошибку.</span><span class="sxs-lookup"><span data-stu-id="b771c-536">The Exchange server returned an error.</span></span> <span data-ttu-id="b771c-537">Для получения дополнительных сведений просмотрите объект диагностики.</span><span class="sxs-lookup"><span data-stu-id="b771c-537">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="b771c-538">Пользователь отключен от сети.</span><span class="sxs-lookup"><span data-stu-id="b771c-538">The user is no longer connected to the network.</span></span> <span data-ttu-id="b771c-539">Проверьте сетевое подключение и повторите попытку.</span><span class="sxs-lookup"><span data-stu-id="b771c-539">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b771c-540">Требования</span><span class="sxs-lookup"><span data-stu-id="b771c-540">Requirements</span></span>

|<span data-ttu-id="b771c-541">Требование</span><span class="sxs-lookup"><span data-stu-id="b771c-541">Requirement</span></span>| <span data-ttu-id="b771c-542">Значение</span><span class="sxs-lookup"><span data-stu-id="b771c-542">Value</span></span>|
|---|---|
|[<span data-ttu-id="b771c-543">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b771c-543">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b771c-544">1.0</span><span class="sxs-lookup"><span data-stu-id="b771c-544">1.0</span></span>|
|[<span data-ttu-id="b771c-545">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b771c-545">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b771c-546">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b771c-546">ReadItem</span></span>|
|[<span data-ttu-id="b771c-547">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b771c-547">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b771c-548">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b771c-548">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b771c-549">Пример</span><span class="sxs-lookup"><span data-stu-id="b771c-549">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="b771c-550">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="b771c-550">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="b771c-551">Выполняет асинхронный запрос для веб-служб Exchange (EWS) на сервере Exchange Server, на котором размещен почтовый ящик пользователя.</span><span class="sxs-lookup"><span data-stu-id="b771c-551">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="b771c-552">Этот метод не поддерживается в следующих сценариях:</span><span class="sxs-lookup"><span data-stu-id="b771c-552">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="b771c-553">В Outlook для iOS и Android</span><span class="sxs-lookup"><span data-stu-id="b771c-553">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="b771c-554">Если надстройка загружается в почтовый ящик Gmail.</span><span class="sxs-lookup"><span data-stu-id="b771c-554">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="b771c-555">В таких случаях надстройка должна [использовать REST API](/outlook/add-ins/use-rest-api) для доступа к почтовому ящику пользователя.</span><span class="sxs-lookup"><span data-stu-id="b771c-555">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="b771c-556">Метод `makeEwsRequestAsync` отправляет запрос EWS от имени надстройки в Exchange.</span><span class="sxs-lookup"><span data-stu-id="b771c-556">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="b771c-557">Список поддерживаемых операций EWS см. в статье [Вызов веб-служб из надстройки Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support).</span><span class="sxs-lookup"><span data-stu-id="b771c-557">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="b771c-558">С помощью метода `makeEwsRequestAsync` невозможно запрашивать элементы, связанные с папкой.</span><span class="sxs-lookup"><span data-stu-id="b771c-558">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="b771c-559">В запросе XML должна быть указана кодировка UTF-8.</span><span class="sxs-lookup"><span data-stu-id="b771c-559">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="b771c-p140">У вашей надстройки должно быть разрешение **ReadWriteMailbox** для использования метода `makeEwsRequestAsync`. Сведения об использовании разрешения **ReadWriteMailbox** и операций EWS, которые можно вызывать с помощью метода `makeEwsRequestAsync`, см. в статье [Указание разрешений для доступа почтовой надстройки к почтовому ящику пользователя](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="b771c-p140">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="b771c-562">Администратор сервера должен установить значение true для параметра `OAuthAuthentication` в каталоге сервера клиентского доступа EWS, чтобы метод `makeEwsRequestAsync` мог выполнять запросы EWS.</span><span class="sxs-lookup"><span data-stu-id="b771c-562">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="b771c-563">Различия версий</span><span class="sxs-lookup"><span data-stu-id="b771c-563">Version differences</span></span>

<span data-ttu-id="b771c-564">Если вы используете метод `makeEwsRequestAsync` в почтовых приложениях, которые выполняются в Outlook версии более ранней, чем 15.0.4535.1004, указывайте кодировку `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="b771c-564">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="b771c-p141">Значение кодировки не нужно указывать, если почтовое приложение выполняется в Outlook в Интернете. Чтобы определить, выполняется ли приложение в Outlook или Outlook в Интернете, используйте свойство mailbox.diagnostics.hostName. Используемую версию Outlook можно определить с помощью свойства mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="b771c-p141">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b771c-568">Параметры</span><span class="sxs-lookup"><span data-stu-id="b771c-568">Parameters</span></span>

|<span data-ttu-id="b771c-569">Имя</span><span class="sxs-lookup"><span data-stu-id="b771c-569">Name</span></span>| <span data-ttu-id="b771c-570">Тип</span><span class="sxs-lookup"><span data-stu-id="b771c-570">Type</span></span>| <span data-ttu-id="b771c-571">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="b771c-571">Attributes</span></span>| <span data-ttu-id="b771c-572">Описание</span><span class="sxs-lookup"><span data-stu-id="b771c-572">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="b771c-573">String</span><span class="sxs-lookup"><span data-stu-id="b771c-573">String</span></span>||<span data-ttu-id="b771c-574">Запрос EWS.</span><span class="sxs-lookup"><span data-stu-id="b771c-574">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="b771c-575">function</span><span class="sxs-lookup"><span data-stu-id="b771c-575">function</span></span>||<span data-ttu-id="b771c-576">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b771c-576">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="b771c-577">Результат XML вызова EWS указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="b771c-577">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="b771c-578">Если размер результата превышает 1 МБ, возвращается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="b771c-578">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="b771c-579">Объект</span><span class="sxs-lookup"><span data-stu-id="b771c-579">Object</span></span>| <span data-ttu-id="b771c-580">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b771c-580">&lt;optional&gt;</span></span>|<span data-ttu-id="b771c-581">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="b771c-581">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b771c-582">Требования</span><span class="sxs-lookup"><span data-stu-id="b771c-582">Requirements</span></span>

|<span data-ttu-id="b771c-583">Требование</span><span class="sxs-lookup"><span data-stu-id="b771c-583">Requirement</span></span>| <span data-ttu-id="b771c-584">Значение</span><span class="sxs-lookup"><span data-stu-id="b771c-584">Value</span></span>|
|---|---|
|[<span data-ttu-id="b771c-585">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b771c-585">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b771c-586">1.0</span><span class="sxs-lookup"><span data-stu-id="b771c-586">1.0</span></span>|
|[<span data-ttu-id="b771c-587">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b771c-587">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b771c-588">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="b771c-588">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="b771c-589">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b771c-589">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b771c-590">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b771c-590">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b771c-591">Пример</span><span class="sxs-lookup"><span data-stu-id="b771c-591">Example</span></span>

<span data-ttu-id="b771c-592">В приведенном ниже примере вызывается `makeEwsRequestAsync` для получения темы элемента с помощью операции `GetItem`.</span><span class="sxs-lookup"><span data-stu-id="b771c-592">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="b771c-593">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="b771c-593">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="b771c-594">Удаляет обработчиков для поддерживаемого типа события.</span><span class="sxs-lookup"><span data-stu-id="b771c-594">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="b771c-595">В настоящее время единственный поддерживаемый тип события — `Office.EventType.ItemChanged`.</span><span class="sxs-lookup"><span data-stu-id="b771c-595">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="b771c-596">Параметры</span><span class="sxs-lookup"><span data-stu-id="b771c-596">Parameters</span></span>

| <span data-ttu-id="b771c-597">Имя</span><span class="sxs-lookup"><span data-stu-id="b771c-597">Name</span></span> | <span data-ttu-id="b771c-598">Тип</span><span class="sxs-lookup"><span data-stu-id="b771c-598">Type</span></span> | <span data-ttu-id="b771c-599">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="b771c-599">Attributes</span></span> | <span data-ttu-id="b771c-600">Описание</span><span class="sxs-lookup"><span data-stu-id="b771c-600">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="b771c-601">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="b771c-601">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="b771c-602">Событие, которое должно отменить обработчик.</span><span class="sxs-lookup"><span data-stu-id="b771c-602">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="b771c-603">Объект</span><span class="sxs-lookup"><span data-stu-id="b771c-603">Object</span></span> | <span data-ttu-id="b771c-604">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b771c-604">&lt;optional&gt;</span></span> | <span data-ttu-id="b771c-605">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="b771c-605">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="b771c-606">Object</span><span class="sxs-lookup"><span data-stu-id="b771c-606">Object</span></span> | <span data-ttu-id="b771c-607">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b771c-607">&lt;optional&gt;</span></span> | <span data-ttu-id="b771c-608">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="b771c-608">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="b771c-609">функция</span><span class="sxs-lookup"><span data-stu-id="b771c-609">function</span></span>| <span data-ttu-id="b771c-610">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="b771c-610">&lt;optional&gt;</span></span>|<span data-ttu-id="b771c-611">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="b771c-611">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="b771c-612">Требования</span><span class="sxs-lookup"><span data-stu-id="b771c-612">Requirements</span></span>

|<span data-ttu-id="b771c-613">Требование</span><span class="sxs-lookup"><span data-stu-id="b771c-613">Requirement</span></span>| <span data-ttu-id="b771c-614">Значение</span><span class="sxs-lookup"><span data-stu-id="b771c-614">Value</span></span>|
|---|---|
|[<span data-ttu-id="b771c-615">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b771c-615">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b771c-616">1.5</span><span class="sxs-lookup"><span data-stu-id="b771c-616">1.5</span></span> |
|[<span data-ttu-id="b771c-617">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b771c-617">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b771c-618">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b771c-618">ReadItem</span></span> |
|[<span data-ttu-id="b771c-619">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b771c-619">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="b771c-620">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b771c-620">Compose or Read</span></span>|
