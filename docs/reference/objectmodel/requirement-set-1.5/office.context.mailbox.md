---
title: Office.context.mailbox — набор обязательных элементов 1.5
description: ''
ms.date: 10/21/2019
localization_priority: Priority
ms.openlocfilehash: bb63d8186d41d072aa62b180b16958d61ce9a66c
ms.sourcegitcommit: 499bf49b41205f8034c501d4db5fe4b02dab205e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/22/2019
ms.locfileid: "37627015"
---
# <a name="mailbox"></a><span data-ttu-id="3b941-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="3b941-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="3b941-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="3b941-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="3b941-104">Предоставляет для Microsoft Outlook доступ к объектной модели надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="3b941-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="3b941-105">Требования</span><span class="sxs-lookup"><span data-stu-id="3b941-105">Requirements</span></span>

|<span data-ttu-id="3b941-106">Требование</span><span class="sxs-lookup"><span data-stu-id="3b941-106">Requirement</span></span>| <span data-ttu-id="3b941-107">Значение</span><span class="sxs-lookup"><span data-stu-id="3b941-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="3b941-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="3b941-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3b941-109">1.0</span><span class="sxs-lookup"><span data-stu-id="3b941-109">1.0</span></span>|
|[<span data-ttu-id="3b941-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="3b941-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3b941-111">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="3b941-111">Restricted</span></span>|
|[<span data-ttu-id="3b941-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3b941-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3b941-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3b941-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="3b941-114">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="3b941-114">Members and methods</span></span>

| <span data-ttu-id="3b941-115">Элемент</span><span class="sxs-lookup"><span data-stu-id="3b941-115">Member</span></span> | <span data-ttu-id="3b941-116">Тип</span><span class="sxs-lookup"><span data-stu-id="3b941-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="3b941-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="3b941-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="3b941-118">Элемент</span><span class="sxs-lookup"><span data-stu-id="3b941-118">Member</span></span> |
| [<span data-ttu-id="3b941-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="3b941-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="3b941-120">Элемент</span><span class="sxs-lookup"><span data-stu-id="3b941-120">Member</span></span> |
| [<span data-ttu-id="3b941-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="3b941-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="3b941-122">Метод</span><span class="sxs-lookup"><span data-stu-id="3b941-122">Method</span></span> |
| [<span data-ttu-id="3b941-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="3b941-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="3b941-124">Метод</span><span class="sxs-lookup"><span data-stu-id="3b941-124">Method</span></span> |
| [<span data-ttu-id="3b941-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="3b941-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="3b941-126">Метод</span><span class="sxs-lookup"><span data-stu-id="3b941-126">Method</span></span> |
| [<span data-ttu-id="3b941-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="3b941-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="3b941-128">Метод</span><span class="sxs-lookup"><span data-stu-id="3b941-128">Method</span></span> |
| [<span data-ttu-id="3b941-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="3b941-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="3b941-130">Метод</span><span class="sxs-lookup"><span data-stu-id="3b941-130">Method</span></span> |
| [<span data-ttu-id="3b941-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="3b941-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="3b941-132">Метод</span><span class="sxs-lookup"><span data-stu-id="3b941-132">Method</span></span> |
| [<span data-ttu-id="3b941-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="3b941-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="3b941-134">Метод</span><span class="sxs-lookup"><span data-stu-id="3b941-134">Method</span></span> |
| [<span data-ttu-id="3b941-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="3b941-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="3b941-136">Метод</span><span class="sxs-lookup"><span data-stu-id="3b941-136">Method</span></span> |
| [<span data-ttu-id="3b941-137">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="3b941-137">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="3b941-138">Метод</span><span class="sxs-lookup"><span data-stu-id="3b941-138">Method</span></span> |
| [<span data-ttu-id="3b941-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="3b941-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="3b941-140">Метод</span><span class="sxs-lookup"><span data-stu-id="3b941-140">Method</span></span> |
| [<span data-ttu-id="3b941-141">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="3b941-141">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="3b941-142">Метод</span><span class="sxs-lookup"><span data-stu-id="3b941-142">Method</span></span> |
| [<span data-ttu-id="3b941-143">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="3b941-143">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="3b941-144">Метод</span><span class="sxs-lookup"><span data-stu-id="3b941-144">Method</span></span> |
| [<span data-ttu-id="3b941-145">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="3b941-145">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="3b941-146">Метод</span><span class="sxs-lookup"><span data-stu-id="3b941-146">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="3b941-147">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="3b941-147">Namespaces</span></span>

<span data-ttu-id="3b941-148">[diagnostics](Office.context.mailbox.diagnostics.md). Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="3b941-148">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="3b941-149">[item](Office.context.mailbox.item.md). Предоставляет методы и свойства для доступа к сообщению или встрече в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="3b941-149">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="3b941-150">[userProfile](Office.context.mailbox.userProfile.md). Предоставляет сведения о пользователе в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="3b941-150">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="3b941-151">Members</span><span class="sxs-lookup"><span data-stu-id="3b941-151">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="3b941-152">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="3b941-152">ewsUrl: String</span></span>

<span data-ttu-id="3b941-p101">Получает URL-адрес конечной точки веб-служб Exchange (EWS) для этой учетной записи электронной почты. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="3b941-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="3b941-155">Этот элемент не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="3b941-155">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="3b941-p102">Удаленная служба может использовать значение `ewsUrl`, чтобы выполнять вызовы EWS для почтового ящика пользователя. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="3b941-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="3b941-158">Чтобы вызвать элемент `ewsUrl` в режиме чтения, в манифесте приложения должно быть указано разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="3b941-158">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="3b941-p103">Перед использованием элемента `ewsUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="3b941-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="3b941-161">Тип</span><span class="sxs-lookup"><span data-stu-id="3b941-161">Type</span></span>

*   <span data-ttu-id="3b941-162">String</span><span class="sxs-lookup"><span data-stu-id="3b941-162">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3b941-163">Требования</span><span class="sxs-lookup"><span data-stu-id="3b941-163">Requirements</span></span>

|<span data-ttu-id="3b941-164">Требование</span><span class="sxs-lookup"><span data-stu-id="3b941-164">Requirement</span></span>| <span data-ttu-id="3b941-165">Значение</span><span class="sxs-lookup"><span data-stu-id="3b941-165">Value</span></span>|
|---|---|
|[<span data-ttu-id="3b941-166">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="3b941-166">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3b941-167">1.0</span><span class="sxs-lookup"><span data-stu-id="3b941-167">1.0</span></span>|
|[<span data-ttu-id="3b941-168">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="3b941-168">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3b941-169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3b941-169">ReadItem</span></span>|
|[<span data-ttu-id="3b941-170">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3b941-170">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3b941-171">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3b941-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="resturl-string"></a><span data-ttu-id="3b941-172">restUrl: String</span><span class="sxs-lookup"><span data-stu-id="3b941-172">restUrl: String</span></span>

<span data-ttu-id="3b941-173">Возвращает URL-адрес конечной точки REST для этой учетной записи электронной почты.</span><span class="sxs-lookup"><span data-stu-id="3b941-173">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="3b941-174">С помощью значения `restUrl` можно выполнять вызовы [REST API](/outlook/rest/) для почтового ящика пользователя.</span><span class="sxs-lookup"><span data-stu-id="3b941-174">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="3b941-175">Чтобы вызвать элемент `restUrl` в режиме чтения, в манифесте приложения необходимо указать разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="3b941-175">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="3b941-p104">Перед использованием элемента `restUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="3b941-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

> [!NOTE]
> <span data-ttu-id="3b941-178">Клиенты Outlook, подключенные к локальным установленным версиям Exchange 2016 или более поздним с пользовательским URL-адресом REST, возвращают недопустимое значение `restUrl`.</span><span class="sxs-lookup"><span data-stu-id="3b941-178">Outlook clients connected to on-premises installations of Exchange 2016 or later with a custom REST URL configured will return an invalid value for `restUrl`.</span></span>

##### <a name="type"></a><span data-ttu-id="3b941-179">Тип</span><span class="sxs-lookup"><span data-stu-id="3b941-179">Type</span></span>

*   <span data-ttu-id="3b941-180">String</span><span class="sxs-lookup"><span data-stu-id="3b941-180">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3b941-181">Требования</span><span class="sxs-lookup"><span data-stu-id="3b941-181">Requirements</span></span>

|<span data-ttu-id="3b941-182">Требование</span><span class="sxs-lookup"><span data-stu-id="3b941-182">Requirement</span></span>| <span data-ttu-id="3b941-183">Значение</span><span class="sxs-lookup"><span data-stu-id="3b941-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="3b941-184">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="3b941-184">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3b941-185">1.5</span><span class="sxs-lookup"><span data-stu-id="3b941-185">1.5</span></span> |
|[<span data-ttu-id="3b941-186">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="3b941-186">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3b941-187">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3b941-187">ReadItem</span></span>|
|[<span data-ttu-id="3b941-188">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3b941-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3b941-189">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3b941-189">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="3b941-190">Методы</span><span class="sxs-lookup"><span data-stu-id="3b941-190">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="3b941-191">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="3b941-191">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="3b941-192">Добавляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="3b941-192">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="3b941-193">В настоящее время поддерживаются только события типа `Office.EventType.ItemChanged`, которые вызываются, когда пользователь выбирает новый элемент.</span><span class="sxs-lookup"><span data-stu-id="3b941-193">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item.</span></span> <span data-ttu-id="3b941-194">Это событие используется надстройками, реализующими закрепляемую область задач, и позволяет надстройке обновлять пользовательский интерфейс области задач в соответствии с выбранным в данный момент элементом.</span><span class="sxs-lookup"><span data-stu-id="3b941-194">This event is used by add-ins that implement a pinnable task pane, and allows the add-in to refresh the task pane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3b941-195">Параметры</span><span class="sxs-lookup"><span data-stu-id="3b941-195">Parameters</span></span>

| <span data-ttu-id="3b941-196">Имя</span><span class="sxs-lookup"><span data-stu-id="3b941-196">Name</span></span> | <span data-ttu-id="3b941-197">Тип</span><span class="sxs-lookup"><span data-stu-id="3b941-197">Type</span></span> | <span data-ttu-id="3b941-198">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="3b941-198">Attributes</span></span> | <span data-ttu-id="3b941-199">Описание</span><span class="sxs-lookup"><span data-stu-id="3b941-199">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="3b941-200">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="3b941-200">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="3b941-201">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="3b941-201">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="3b941-202">Function</span><span class="sxs-lookup"><span data-stu-id="3b941-202">Function</span></span> || <span data-ttu-id="3b941-p106">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="3b941-p106">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="3b941-206">Объект</span><span class="sxs-lookup"><span data-stu-id="3b941-206">Object</span></span> | <span data-ttu-id="3b941-207">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="3b941-207">&lt;optional&gt;</span></span> | <span data-ttu-id="3b941-208">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="3b941-208">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="3b941-209">Object</span><span class="sxs-lookup"><span data-stu-id="3b941-209">Object</span></span> | <span data-ttu-id="3b941-210">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="3b941-210">&lt;optional&gt;</span></span> | <span data-ttu-id="3b941-211">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="3b941-211">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="3b941-212">функция</span><span class="sxs-lookup"><span data-stu-id="3b941-212">function</span></span>| <span data-ttu-id="3b941-213">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="3b941-213">&lt;optional&gt;</span></span>|<span data-ttu-id="3b941-214">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="3b941-214">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3b941-215">Требования</span><span class="sxs-lookup"><span data-stu-id="3b941-215">Requirements</span></span>

|<span data-ttu-id="3b941-216">Требование</span><span class="sxs-lookup"><span data-stu-id="3b941-216">Requirement</span></span>| <span data-ttu-id="3b941-217">Значение</span><span class="sxs-lookup"><span data-stu-id="3b941-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="3b941-218">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="3b941-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3b941-219">1.5</span><span class="sxs-lookup"><span data-stu-id="3b941-219">1.5</span></span> |
|[<span data-ttu-id="3b941-220">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="3b941-220">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3b941-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3b941-221">ReadItem</span></span> |
|[<span data-ttu-id="3b941-222">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3b941-222">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3b941-223">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3b941-223">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3b941-224">Пример</span><span class="sxs-lookup"><span data-stu-id="3b941-224">Example</span></span>

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

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="3b941-225">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="3b941-225">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="3b941-226">Преобразовывает идентификатор элемента из формата REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="3b941-226">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="3b941-227">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="3b941-227">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="3b941-p107">Формат идентификаторов, извлекаемых через API REST (например, [API Почты Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)), отличается от формата веб-служб Exchange (EWS). Метод `convertToEwsId` преобразовывает идентификатор в формате REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="3b941-p107">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3b941-230">Параметры</span><span class="sxs-lookup"><span data-stu-id="3b941-230">Parameters</span></span>

|<span data-ttu-id="3b941-231">Имя</span><span class="sxs-lookup"><span data-stu-id="3b941-231">Name</span></span>| <span data-ttu-id="3b941-232">Тип</span><span class="sxs-lookup"><span data-stu-id="3b941-232">Type</span></span>| <span data-ttu-id="3b941-233">Описание</span><span class="sxs-lookup"><span data-stu-id="3b941-233">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="3b941-234">String</span><span class="sxs-lookup"><span data-stu-id="3b941-234">String</span></span>|<span data-ttu-id="3b941-235">Идентификатор элемента в формате REST API для Outlook</span><span class="sxs-lookup"><span data-stu-id="3b941-235">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="3b941-236">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="3b941-236">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.5)|<span data-ttu-id="3b941-237">Значение, определяющее версию REST API для Outlook, которая используется для извлечения идентификатора элемента.</span><span class="sxs-lookup"><span data-stu-id="3b941-237">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3b941-238">Требования</span><span class="sxs-lookup"><span data-stu-id="3b941-238">Requirements</span></span>

|<span data-ttu-id="3b941-239">Требование</span><span class="sxs-lookup"><span data-stu-id="3b941-239">Requirement</span></span>| <span data-ttu-id="3b941-240">Значение</span><span class="sxs-lookup"><span data-stu-id="3b941-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="3b941-241">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="3b941-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3b941-242">1.3</span><span class="sxs-lookup"><span data-stu-id="3b941-242">1.3</span></span>|
|[<span data-ttu-id="3b941-243">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="3b941-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3b941-244">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="3b941-244">Restricted</span></span>|
|[<span data-ttu-id="3b941-245">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3b941-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3b941-246">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3b941-246">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3b941-247">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="3b941-247">Returns:</span></span>

<span data-ttu-id="3b941-248">Тип: String</span><span class="sxs-lookup"><span data-stu-id="3b941-248">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="3b941-249">Пример</span><span class="sxs-lookup"><span data-stu-id="3b941-249">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-15"></a><span data-ttu-id="3b941-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)}</span><span class="sxs-lookup"><span data-stu-id="3b941-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)}</span></span>

<span data-ttu-id="3b941-251">Получает словарь, содержащий сведения о локальном времени клиента.</span><span class="sxs-lookup"><span data-stu-id="3b941-251">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="3b941-p108">Почтовое приложение для классической версии Outlook или версии в Интернете может использовать разные часовые пояса для дат и времени. Классическое приложение Outlook использует часовой пояс клиентского компьютера. Outlook в Интернете использует часовой пояс, заданный в Центре администрирования Exchange (EAC). Значения даты и времени должны обрабатываться так, чтобы значения в пользовательском интерфейсе всегда согласовывались с часовым поясом, ожидаемым пользователем.</span><span class="sxs-lookup"><span data-stu-id="3b941-p108">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times. Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="3b941-p109">Если почтовое приложение работает в классическом клиенте Outlook, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса клиентского компьютера. Если почтовое приложение работает в Outlook в Интернете, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса, заданного в Центре администрирования Exchange.</span><span class="sxs-lookup"><span data-stu-id="3b941-p109">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3b941-257">Параметры</span><span class="sxs-lookup"><span data-stu-id="3b941-257">Parameters</span></span>

|<span data-ttu-id="3b941-258">Имя</span><span class="sxs-lookup"><span data-stu-id="3b941-258">Name</span></span>| <span data-ttu-id="3b941-259">Тип</span><span class="sxs-lookup"><span data-stu-id="3b941-259">Type</span></span>| <span data-ttu-id="3b941-260">Описание</span><span class="sxs-lookup"><span data-stu-id="3b941-260">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="3b941-261">Date</span><span class="sxs-lookup"><span data-stu-id="3b941-261">Date</span></span>|<span data-ttu-id="3b941-262">Объект Date</span><span class="sxs-lookup"><span data-stu-id="3b941-262">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3b941-263">Требования</span><span class="sxs-lookup"><span data-stu-id="3b941-263">Requirements</span></span>

|<span data-ttu-id="3b941-264">Требование</span><span class="sxs-lookup"><span data-stu-id="3b941-264">Requirement</span></span>| <span data-ttu-id="3b941-265">Значение</span><span class="sxs-lookup"><span data-stu-id="3b941-265">Value</span></span>|
|---|---|
|[<span data-ttu-id="3b941-266">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="3b941-266">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3b941-267">1.0</span><span class="sxs-lookup"><span data-stu-id="3b941-267">1.0</span></span>|
|[<span data-ttu-id="3b941-268">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="3b941-268">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3b941-269">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3b941-269">ReadItem</span></span>|
|[<span data-ttu-id="3b941-270">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3b941-270">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3b941-271">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3b941-271">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3b941-272">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="3b941-272">Returns:</span></span>

<span data-ttu-id="3b941-273">Тип: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="3b941-273">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="3b941-274">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="3b941-274">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="3b941-275">Преобразовывает идентификатор элемента в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="3b941-275">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="3b941-276">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="3b941-276">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="3b941-p110">Формат идентификаторов, извлекаемых через EWS или свойство `itemId`, отличается от формата API REST (таких как [API Почты Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)). Метод `convertToRestId` преобразовывает идентификатор в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="3b941-p110">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3b941-279">Параметры</span><span class="sxs-lookup"><span data-stu-id="3b941-279">Parameters</span></span>

|<span data-ttu-id="3b941-280">Имя</span><span class="sxs-lookup"><span data-stu-id="3b941-280">Name</span></span>| <span data-ttu-id="3b941-281">Тип</span><span class="sxs-lookup"><span data-stu-id="3b941-281">Type</span></span>| <span data-ttu-id="3b941-282">Описание</span><span class="sxs-lookup"><span data-stu-id="3b941-282">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="3b941-283">String</span><span class="sxs-lookup"><span data-stu-id="3b941-283">String</span></span>|<span data-ttu-id="3b941-284">Идентификатор элемента в формате EWS</span><span class="sxs-lookup"><span data-stu-id="3b941-284">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="3b941-285">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="3b941-285">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.5)|<span data-ttu-id="3b941-286">Значение, определяющее версию REST API для Outlook, с которой будет использоваться преобразованный идентификатор.</span><span class="sxs-lookup"><span data-stu-id="3b941-286">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3b941-287">Требования</span><span class="sxs-lookup"><span data-stu-id="3b941-287">Requirements</span></span>

|<span data-ttu-id="3b941-288">Требование</span><span class="sxs-lookup"><span data-stu-id="3b941-288">Requirement</span></span>| <span data-ttu-id="3b941-289">Значение</span><span class="sxs-lookup"><span data-stu-id="3b941-289">Value</span></span>|
|---|---|
|[<span data-ttu-id="3b941-290">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="3b941-290">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3b941-291">1.3</span><span class="sxs-lookup"><span data-stu-id="3b941-291">1.3</span></span>|
|[<span data-ttu-id="3b941-292">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="3b941-292">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3b941-293">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="3b941-293">Restricted</span></span>|
|[<span data-ttu-id="3b941-294">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3b941-294">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3b941-295">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3b941-295">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3b941-296">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="3b941-296">Returns:</span></span>

<span data-ttu-id="3b941-297">Тип: String</span><span class="sxs-lookup"><span data-stu-id="3b941-297">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="3b941-298">Пример</span><span class="sxs-lookup"><span data-stu-id="3b941-298">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="3b941-299">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="3b941-299">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="3b941-300">Получает объект Date из словаря, содержащего сведения о времени.</span><span class="sxs-lookup"><span data-stu-id="3b941-300">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="3b941-301">Метод `convertToUtcClientTime` преобразует словарь, содержащий локальную дату и время, в объект Date с правильными значениями локальной даты и времени.</span><span class="sxs-lookup"><span data-stu-id="3b941-301">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3b941-302">Параметры</span><span class="sxs-lookup"><span data-stu-id="3b941-302">Parameters</span></span>

|<span data-ttu-id="3b941-303">Имя</span><span class="sxs-lookup"><span data-stu-id="3b941-303">Name</span></span>| <span data-ttu-id="3b941-304">Тип</span><span class="sxs-lookup"><span data-stu-id="3b941-304">Type</span></span>| <span data-ttu-id="3b941-305">Описание</span><span class="sxs-lookup"><span data-stu-id="3b941-305">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="3b941-306">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="3b941-306">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)|<span data-ttu-id="3b941-307">Значение локального времени для преобразования.</span><span class="sxs-lookup"><span data-stu-id="3b941-307">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3b941-308">Требования</span><span class="sxs-lookup"><span data-stu-id="3b941-308">Requirements</span></span>

|<span data-ttu-id="3b941-309">Требование</span><span class="sxs-lookup"><span data-stu-id="3b941-309">Requirement</span></span>| <span data-ttu-id="3b941-310">Значение</span><span class="sxs-lookup"><span data-stu-id="3b941-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="3b941-311">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="3b941-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3b941-312">1.0</span><span class="sxs-lookup"><span data-stu-id="3b941-312">1.0</span></span>|
|[<span data-ttu-id="3b941-313">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="3b941-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3b941-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3b941-314">ReadItem</span></span>|
|[<span data-ttu-id="3b941-315">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3b941-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3b941-316">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3b941-316">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="3b941-317">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="3b941-317">Returns:</span></span>

<span data-ttu-id="3b941-318">Объект Date со временем в формате UTC.</span><span class="sxs-lookup"><span data-stu-id="3b941-318">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="3b941-319">Тип: Date</span><span class="sxs-lookup"><span data-stu-id="3b941-319">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="3b941-320">Пример</span><span class="sxs-lookup"><span data-stu-id="3b941-320">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="3b941-321">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="3b941-321">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="3b941-322">Отображает имеющуюся встречу из календаря.</span><span class="sxs-lookup"><span data-stu-id="3b941-322">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="3b941-323">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="3b941-323">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="3b941-324">Метод `displayAppointmentForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее сведения календаря о существующей встрече.</span><span class="sxs-lookup"><span data-stu-id="3b941-324">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="3b941-p111">В Outlook для Mac с помощью этого метода можно отобразить одну встречу, которая не является частью повторяющегося ряда, или основную встречу такого ряда, но не экземпляр из него, так как в Outlook для Mac невозможно получить доступ к свойствам экземпляра повторяющегося ряда (в том числе к идентификатору элемента).</span><span class="sxs-lookup"><span data-stu-id="3b941-p111">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="3b941-327">В Outlook в Интернете этот метод открывает указанную форму, только если текст формы содержит символы размером не более 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="3b941-327">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="3b941-328">Если указанный идентификатор элемента не определяет существующую встречу, на клиентском компьютере или устройстве открывается пустая страница, и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="3b941-328">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3b941-329">Параметры</span><span class="sxs-lookup"><span data-stu-id="3b941-329">Parameters</span></span>

|<span data-ttu-id="3b941-330">Имя</span><span class="sxs-lookup"><span data-stu-id="3b941-330">Name</span></span>| <span data-ttu-id="3b941-331">Тип</span><span class="sxs-lookup"><span data-stu-id="3b941-331">Type</span></span>| <span data-ttu-id="3b941-332">Описание</span><span class="sxs-lookup"><span data-stu-id="3b941-332">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="3b941-333">String</span><span class="sxs-lookup"><span data-stu-id="3b941-333">String</span></span>|<span data-ttu-id="3b941-334">Идентификатор веб-служб Exchange для существующей встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="3b941-334">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3b941-335">Требования</span><span class="sxs-lookup"><span data-stu-id="3b941-335">Requirements</span></span>

|<span data-ttu-id="3b941-336">Требование</span><span class="sxs-lookup"><span data-stu-id="3b941-336">Requirement</span></span>| <span data-ttu-id="3b941-337">Значение</span><span class="sxs-lookup"><span data-stu-id="3b941-337">Value</span></span>|
|---|---|
|[<span data-ttu-id="3b941-338">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="3b941-338">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3b941-339">1.0</span><span class="sxs-lookup"><span data-stu-id="3b941-339">1.0</span></span>|
|[<span data-ttu-id="3b941-340">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="3b941-340">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3b941-341">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3b941-341">ReadItem</span></span>|
|[<span data-ttu-id="3b941-342">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3b941-342">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3b941-343">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3b941-343">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3b941-344">Пример</span><span class="sxs-lookup"><span data-stu-id="3b941-344">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="3b941-345">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="3b941-345">displayMessageForm(itemId)</span></span>

<span data-ttu-id="3b941-346">Отображает имеющееся сообщение.</span><span class="sxs-lookup"><span data-stu-id="3b941-346">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="3b941-347">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="3b941-347">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="3b941-348">Метод `displayMessageForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее существующее сообщение.</span><span class="sxs-lookup"><span data-stu-id="3b941-348">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="3b941-349">В Outlook в Интернете этот метод открывает указанную форму, только если текст формы содержит символы размером не более 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="3b941-349">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="3b941-350">Если указанный идентификатор элемента не определяет существующее сообщение, окно на клиентском компьютере не открывается и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="3b941-350">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="3b941-p112">Не используйте `displayMessageForm` с параметром `itemId`, который представляет собой встречу. Используйте метод `displayAppointmentForm`, чтобы отобразить сведения о существующей встрече, а метод `displayNewAppointmentForm` — для отображения формы создания встречи.</span><span class="sxs-lookup"><span data-stu-id="3b941-p112">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3b941-353">Параметры</span><span class="sxs-lookup"><span data-stu-id="3b941-353">Parameters</span></span>

|<span data-ttu-id="3b941-354">Имя</span><span class="sxs-lookup"><span data-stu-id="3b941-354">Name</span></span>| <span data-ttu-id="3b941-355">Тип</span><span class="sxs-lookup"><span data-stu-id="3b941-355">Type</span></span>| <span data-ttu-id="3b941-356">Описание</span><span class="sxs-lookup"><span data-stu-id="3b941-356">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="3b941-357">String</span><span class="sxs-lookup"><span data-stu-id="3b941-357">String</span></span>|<span data-ttu-id="3b941-358">Идентификатор веб-служб Exchange для существующего сообщения.</span><span class="sxs-lookup"><span data-stu-id="3b941-358">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3b941-359">Требования</span><span class="sxs-lookup"><span data-stu-id="3b941-359">Requirements</span></span>

|<span data-ttu-id="3b941-360">Требование</span><span class="sxs-lookup"><span data-stu-id="3b941-360">Requirement</span></span>| <span data-ttu-id="3b941-361">Значение</span><span class="sxs-lookup"><span data-stu-id="3b941-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="3b941-362">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="3b941-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3b941-363">1.0</span><span class="sxs-lookup"><span data-stu-id="3b941-363">1.0</span></span>|
|[<span data-ttu-id="3b941-364">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="3b941-364">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3b941-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3b941-365">ReadItem</span></span>|
|[<span data-ttu-id="3b941-366">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3b941-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3b941-367">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3b941-367">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3b941-368">Пример</span><span class="sxs-lookup"><span data-stu-id="3b941-368">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="3b941-369">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="3b941-369">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="3b941-370">Отображает форму для создания новой встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="3b941-370">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="3b941-371">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="3b941-371">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="3b941-p113">Метод `displayNewAppointmentForm` открывает форму, в которой пользователь может создать встречу или собрание. Если параметры заданы, поля формы встречи автоматически заполняются их содержимым.</span><span class="sxs-lookup"><span data-stu-id="3b941-p113">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="3b941-p114">В Outlook в Интернете и на мобильных устройствах этот метод всегда отображает форму с полем участников. Если вы не укажете участников в качестве входных аргументов, метод отображает форму с кнопкой **Сохранить**. Если вы укажете участников, форма будет включать участников и кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="3b941-p114">In Outlook on the web and mobile devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="3b941-p115">Если вы укажете участников или ресурсы с помощью параметра `requiredAttendees`, `optionalAttendees` или `resources` в клиенте Outlook с расширенными возможностями и Outlook RT, этот метод отобразит форму собрания с кнопкой **Отправить**. Если не указать получателей, этот метод отобразит форму встречи с кнопкой **Сохранить и закрыть**.</span><span class="sxs-lookup"><span data-stu-id="3b941-p115">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="3b941-379">Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.</span><span class="sxs-lookup"><span data-stu-id="3b941-379">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3b941-380">Параметры</span><span class="sxs-lookup"><span data-stu-id="3b941-380">Parameters</span></span>

|<span data-ttu-id="3b941-381">Имя</span><span class="sxs-lookup"><span data-stu-id="3b941-381">Name</span></span>| <span data-ttu-id="3b941-382">Тип</span><span class="sxs-lookup"><span data-stu-id="3b941-382">Type</span></span>| <span data-ttu-id="3b941-383">Описание</span><span class="sxs-lookup"><span data-stu-id="3b941-383">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="3b941-384">Object</span><span class="sxs-lookup"><span data-stu-id="3b941-384">Object</span></span> | <span data-ttu-id="3b941-385">Словарь параметров, описывающий новую встречу.</span><span class="sxs-lookup"><span data-stu-id="3b941-385">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="3b941-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span><span class="sxs-lookup"><span data-stu-id="3b941-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span></span> | <span data-ttu-id="3b941-p116">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из обязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="3b941-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="3b941-389">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span><span class="sxs-lookup"><span data-stu-id="3b941-389">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span></span> | <span data-ttu-id="3b941-p117">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из необязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="3b941-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="3b941-392">Date</span><span class="sxs-lookup"><span data-stu-id="3b941-392">Date</span></span> | <span data-ttu-id="3b941-393">Объект `Date`, указывающий дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="3b941-393">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="3b941-394">Date</span><span class="sxs-lookup"><span data-stu-id="3b941-394">Date</span></span> | <span data-ttu-id="3b941-395">Объект `Date`, указывающий дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="3b941-395">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="3b941-396">String</span><span class="sxs-lookup"><span data-stu-id="3b941-396">String</span></span> | <span data-ttu-id="3b941-p118">Строка со сведениями о месте встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="3b941-p118">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="3b941-399">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="3b941-399">Array.&lt;String&gt;</span></span> | <span data-ttu-id="3b941-p119">Массив строк, содержащий необходимые для встречи ресурсы. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="3b941-p119">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="3b941-402">String</span><span class="sxs-lookup"><span data-stu-id="3b941-402">String</span></span> | <span data-ttu-id="3b941-p120">Строка с темой встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="3b941-p120">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="3b941-405">String</span><span class="sxs-lookup"><span data-stu-id="3b941-405">String</span></span> | <span data-ttu-id="3b941-p121">Текст сообщения о встрече. Максимальный размер содержимого сообщения — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="3b941-p121">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="3b941-408">Требования</span><span class="sxs-lookup"><span data-stu-id="3b941-408">Requirements</span></span>

|<span data-ttu-id="3b941-409">Требование</span><span class="sxs-lookup"><span data-stu-id="3b941-409">Requirement</span></span>| <span data-ttu-id="3b941-410">Значение</span><span class="sxs-lookup"><span data-stu-id="3b941-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="3b941-411">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="3b941-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3b941-412">1.0</span><span class="sxs-lookup"><span data-stu-id="3b941-412">1.0</span></span>|
|[<span data-ttu-id="3b941-413">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="3b941-413">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3b941-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3b941-414">ReadItem</span></span>|
|[<span data-ttu-id="3b941-415">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3b941-415">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3b941-416">Чтение</span><span class="sxs-lookup"><span data-stu-id="3b941-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3b941-417">Пример</span><span class="sxs-lookup"><span data-stu-id="3b941-417">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="3b941-418">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="3b941-418">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="3b941-419">Возвращает строку, содержащую маркер, который используется для вызова интерфейсов REST API или веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="3b941-419">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="3b941-p122">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="3b941-p122">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="3b941-422">Рекомендуем сделать так, чтобы по мере возможности надстройки использовали интерфейсы REST API, а не веб-службы Exchange.</span><span class="sxs-lookup"><span data-stu-id="3b941-422">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span>

<span data-ttu-id="3b941-423">Для вызова метода `getCallbackTokenAsync` в режиме чтения требуется минимальный уровень разрешения **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="3b941-423">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="3b941-424">Для вызова `getCallbackTokenAsync` в режиме создания сообщения требуется сохранить элемент.</span><span class="sxs-lookup"><span data-stu-id="3b941-424">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="3b941-425">Для метода [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) требуется минимальный уровень разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="3b941-425">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

<span data-ttu-id="3b941-426">**Маркеры REST**</span><span class="sxs-lookup"><span data-stu-id="3b941-426">**REST Tokens**</span></span>

<span data-ttu-id="3b941-p124">Если запрашивается маркер REST (`options.isRest = true`), полученный маркер не подойдет для проверки подлинности при вызовах веб-служб Exchange. Область действия маркера будет ограничена доступом только для чтения к текущему элементу и его вложениям, если в манифесте надстройки не указано разрешение [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission). Если указано разрешение `ReadWriteMailbox`, полученный маркер предоставит доступ на чтение и запись к почте, календарю и контактам, включая возможность отправки почты.</span><span class="sxs-lookup"><span data-stu-id="3b941-p124">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="3b941-430">С помощью свойства `restUrl` надстройка должна определить правильный URL-адрес для вызовов REST API.</span><span class="sxs-lookup"><span data-stu-id="3b941-430">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="3b941-431">**Маркеры EWS**</span><span class="sxs-lookup"><span data-stu-id="3b941-431">**EWS Tokens**</span></span>

<span data-ttu-id="3b941-p125">Если запрашивается маркер EWS (`options.isRest = false`), полученный маркер не подойдет для проверки подлинности при вызовах REST API. Область действия маркера будет ограничена доступом к текущему элементу.</span><span class="sxs-lookup"><span data-stu-id="3b941-p125">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="3b941-434">С помощью свойства `ewsUrl` надстройка должна определить правильный URL-адрес для вызовов EWS.</span><span class="sxs-lookup"><span data-stu-id="3b941-434">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

<span data-ttu-id="3b941-435">Вы можете передать сторонней системе маркер и идентификатор вложения или элемента.</span><span class="sxs-lookup"><span data-stu-id="3b941-435">You can pass the token and an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="3b941-436">Сторонняя система использует этот маркер как маркер авторизации, чтобы вызвать операцию [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) или [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) веб-служб Exchange для возврата вложения или элемента.</span><span class="sxs-lookup"><span data-stu-id="3b941-436">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="3b941-437">Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="3b941-437">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="parameters"></a><span data-ttu-id="3b941-438">Параметры</span><span class="sxs-lookup"><span data-stu-id="3b941-438">Parameters</span></span>

|<span data-ttu-id="3b941-439">Имя</span><span class="sxs-lookup"><span data-stu-id="3b941-439">Name</span></span>| <span data-ttu-id="3b941-440">Тип</span><span class="sxs-lookup"><span data-stu-id="3b941-440">Type</span></span>| <span data-ttu-id="3b941-441">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="3b941-441">Attributes</span></span>| <span data-ttu-id="3b941-442">Описание</span><span class="sxs-lookup"><span data-stu-id="3b941-442">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="3b941-443">Object</span><span class="sxs-lookup"><span data-stu-id="3b941-443">Object</span></span> | <span data-ttu-id="3b941-444">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="3b941-444">&lt;optional&gt;</span></span> | <span data-ttu-id="3b941-445">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="3b941-445">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="3b941-446">Boolean</span><span class="sxs-lookup"><span data-stu-id="3b941-446">Boolean</span></span> |  <span data-ttu-id="3b941-447">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="3b941-447">&lt;optional&gt;</span></span> | <span data-ttu-id="3b941-p127">Определяет, будет ли предоставленный маркер использоваться для интерфейсов REST API Outlook или веб-служб Exchange. Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="3b941-p127">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="3b941-450">Объект</span><span class="sxs-lookup"><span data-stu-id="3b941-450">Object</span></span> |  <span data-ttu-id="3b941-451">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="3b941-451">&lt;optional&gt;</span></span> | <span data-ttu-id="3b941-452">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="3b941-452">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="3b941-453">функция</span><span class="sxs-lookup"><span data-stu-id="3b941-453">function</span></span>||<span data-ttu-id="3b941-454">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="3b941-454">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="3b941-455">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="3b941-455">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="3b941-456">При наличии ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут предоставлять дополнительные сведения.</span><span class="sxs-lookup"><span data-stu-id="3b941-456">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="3b941-457">Ошибки</span><span class="sxs-lookup"><span data-stu-id="3b941-457">Errors</span></span>

|<span data-ttu-id="3b941-458">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="3b941-458">Error code</span></span>|<span data-ttu-id="3b941-459">Описание</span><span class="sxs-lookup"><span data-stu-id="3b941-459">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="3b941-460">Не удалось выполнить запрос.</span><span class="sxs-lookup"><span data-stu-id="3b941-460">The request has failed.</span></span> <span data-ttu-id="3b941-461">Просмотрите объект диагностики для кода ошибки HTTP.</span><span class="sxs-lookup"><span data-stu-id="3b941-461">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="3b941-462">Сервер Exchange Server вернул ошибку.</span><span class="sxs-lookup"><span data-stu-id="3b941-462">The Exchange server returned an error.</span></span> <span data-ttu-id="3b941-463">Для получения дополнительных сведений просмотрите объект диагностики.</span><span class="sxs-lookup"><span data-stu-id="3b941-463">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="3b941-464">Пользователь отключен от сети.</span><span class="sxs-lookup"><span data-stu-id="3b941-464">The user is no longer connected to the network.</span></span> <span data-ttu-id="3b941-465">Проверьте сетевое подключение и повторите попытку.</span><span class="sxs-lookup"><span data-stu-id="3b941-465">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3b941-466">Требования</span><span class="sxs-lookup"><span data-stu-id="3b941-466">Requirements</span></span>

|<span data-ttu-id="3b941-467">Требование</span><span class="sxs-lookup"><span data-stu-id="3b941-467">Requirement</span></span>| <span data-ttu-id="3b941-468">Значение</span><span class="sxs-lookup"><span data-stu-id="3b941-468">Value</span></span>|
|---|---|
|[<span data-ttu-id="3b941-469">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="3b941-469">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3b941-470">1.5</span><span class="sxs-lookup"><span data-stu-id="3b941-470">1.5</span></span> |
|[<span data-ttu-id="3b941-471">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="3b941-471">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3b941-472">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3b941-472">ReadItem</span></span>|
|[<span data-ttu-id="3b941-473">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3b941-473">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3b941-474">Создание и чтение</span><span class="sxs-lookup"><span data-stu-id="3b941-474">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="3b941-475">Пример</span><span class="sxs-lookup"><span data-stu-id="3b941-475">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="3b941-476">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="3b941-476">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="3b941-477">Получает строку, содержащую маркер, используемый для получения вложения или элемента с Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="3b941-477">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="3b941-p131">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="3b941-p131">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="3b941-480">Вы можете передать сторонней системе маркер и идентификатор вложения или элемента.</span><span class="sxs-lookup"><span data-stu-id="3b941-480">You can pass the token and an attachment identifier or item identifier to a third-party system.</span></span> <span data-ttu-id="3b941-481">Сторонняя система использует этот маркер как маркер авторизации, чтобы вызвать операцию [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) или [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) веб-служб Exchange для возврата вложения или элемента.</span><span class="sxs-lookup"><span data-stu-id="3b941-481">The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item.</span></span> <span data-ttu-id="3b941-482">Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="3b941-482">For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="3b941-483">Для вызова метода `getCallbackTokenAsync` в режиме чтения требуется минимальный уровень разрешения **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="3b941-483">Calling the `getCallbackTokenAsync` method in read mode requires a minimum permission level of **ReadItem**.</span></span>

<span data-ttu-id="3b941-484">Для вызова `getCallbackTokenAsync` в режиме создания сообщения требуется сохранить элемент.</span><span class="sxs-lookup"><span data-stu-id="3b941-484">Calling `getCallbackTokenAsync` in compose mode requires you to have saved the item.</span></span> <span data-ttu-id="3b941-485">Для метода [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) требуется минимальный уровень разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="3b941-485">The [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method requires a minimum permission level of **ReadWriteItem**.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3b941-486">Параметры</span><span class="sxs-lookup"><span data-stu-id="3b941-486">Parameters</span></span>

|<span data-ttu-id="3b941-487">Имя</span><span class="sxs-lookup"><span data-stu-id="3b941-487">Name</span></span>| <span data-ttu-id="3b941-488">Тип</span><span class="sxs-lookup"><span data-stu-id="3b941-488">Type</span></span>| <span data-ttu-id="3b941-489">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="3b941-489">Attributes</span></span>| <span data-ttu-id="3b941-490">Описание</span><span class="sxs-lookup"><span data-stu-id="3b941-490">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="3b941-491">function</span><span class="sxs-lookup"><span data-stu-id="3b941-491">function</span></span>||<span data-ttu-id="3b941-492">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="3b941-492">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="3b941-493">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="3b941-493">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="3b941-494">При наличии ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут предоставлять дополнительные сведения.</span><span class="sxs-lookup"><span data-stu-id="3b941-494">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="3b941-495">Объект</span><span class="sxs-lookup"><span data-stu-id="3b941-495">Object</span></span>| <span data-ttu-id="3b941-496">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="3b941-496">&lt;optional&gt;</span></span>|<span data-ttu-id="3b941-497">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="3b941-497">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="3b941-498">Ошибки</span><span class="sxs-lookup"><span data-stu-id="3b941-498">Errors</span></span>

|<span data-ttu-id="3b941-499">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="3b941-499">Error code</span></span>|<span data-ttu-id="3b941-500">Описание</span><span class="sxs-lookup"><span data-stu-id="3b941-500">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="3b941-501">Не удалось выполнить запрос.</span><span class="sxs-lookup"><span data-stu-id="3b941-501">The request has failed.</span></span> <span data-ttu-id="3b941-502">Просмотрите объект диагностики для кода ошибки HTTP.</span><span class="sxs-lookup"><span data-stu-id="3b941-502">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="3b941-503">Сервер Exchange Server вернул ошибку.</span><span class="sxs-lookup"><span data-stu-id="3b941-503">The Exchange server returned an error.</span></span> <span data-ttu-id="3b941-504">Для получения дополнительных сведений просмотрите объект диагностики.</span><span class="sxs-lookup"><span data-stu-id="3b941-504">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="3b941-505">Пользователь отключен от сети.</span><span class="sxs-lookup"><span data-stu-id="3b941-505">The user is no longer connected to the network.</span></span> <span data-ttu-id="3b941-506">Проверьте сетевое подключение и повторите попытку.</span><span class="sxs-lookup"><span data-stu-id="3b941-506">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3b941-507">Требования</span><span class="sxs-lookup"><span data-stu-id="3b941-507">Requirements</span></span>

|<span data-ttu-id="3b941-508">Требование</span><span class="sxs-lookup"><span data-stu-id="3b941-508">Requirement</span></span>|||
|---|---|---|
|[<span data-ttu-id="3b941-509">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="3b941-509">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3b941-510">1.0</span><span class="sxs-lookup"><span data-stu-id="3b941-510">1.0</span></span> | <span data-ttu-id="3b941-511">1.3</span><span class="sxs-lookup"><span data-stu-id="3b941-511">1.3</span></span> |
|[<span data-ttu-id="3b941-512">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="3b941-512">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3b941-513">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3b941-513">ReadItem</span></span> | <span data-ttu-id="3b941-514">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3b941-514">ReadItem</span></span> |
|[<span data-ttu-id="3b941-515">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3b941-515">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3b941-516">Чтение</span><span class="sxs-lookup"><span data-stu-id="3b941-516">Read</span></span> | <span data-ttu-id="3b941-517">Создание</span><span class="sxs-lookup"><span data-stu-id="3b941-517">Compose</span></span> |

##### <a name="example"></a><span data-ttu-id="3b941-518">Пример</span><span class="sxs-lookup"><span data-stu-id="3b941-518">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="3b941-519">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="3b941-519">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="3b941-520">Получает маркер, идентифицирующий пользователя и надстройку Office.</span><span class="sxs-lookup"><span data-stu-id="3b941-520">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="3b941-521">Метод `getUserIdentityTokenAsync` возвращает маркер, который можно использовать для идентификации, а также [проверки подлинности надстройки и пользователя в сторонней системе](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="3b941-521">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="3b941-522">Параметры</span><span class="sxs-lookup"><span data-stu-id="3b941-522">Parameters</span></span>

|<span data-ttu-id="3b941-523">Имя</span><span class="sxs-lookup"><span data-stu-id="3b941-523">Name</span></span>| <span data-ttu-id="3b941-524">Тип</span><span class="sxs-lookup"><span data-stu-id="3b941-524">Type</span></span>| <span data-ttu-id="3b941-525">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="3b941-525">Attributes</span></span>| <span data-ttu-id="3b941-526">Описание</span><span class="sxs-lookup"><span data-stu-id="3b941-526">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="3b941-527">function</span><span class="sxs-lookup"><span data-stu-id="3b941-527">function</span></span>||<span data-ttu-id="3b941-528">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="3b941-528">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="3b941-529">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="3b941-529">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="3b941-530">При наличии ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут предоставлять дополнительные сведения.</span><span class="sxs-lookup"><span data-stu-id="3b941-530">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="3b941-531">Объект</span><span class="sxs-lookup"><span data-stu-id="3b941-531">Object</span></span>| <span data-ttu-id="3b941-532">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="3b941-532">&lt;optional&gt;</span></span>|<span data-ttu-id="3b941-533">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="3b941-533">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="3b941-534">Ошибки</span><span class="sxs-lookup"><span data-stu-id="3b941-534">Errors</span></span>

|<span data-ttu-id="3b941-535">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="3b941-535">Error code</span></span>|<span data-ttu-id="3b941-536">Описание</span><span class="sxs-lookup"><span data-stu-id="3b941-536">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="3b941-537">Не удалось выполнить запрос.</span><span class="sxs-lookup"><span data-stu-id="3b941-537">The request has failed.</span></span> <span data-ttu-id="3b941-538">Просмотрите объект диагностики для кода ошибки HTTP.</span><span class="sxs-lookup"><span data-stu-id="3b941-538">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="3b941-539">Сервер Exchange Server вернул ошибку.</span><span class="sxs-lookup"><span data-stu-id="3b941-539">The Exchange server returned an error.</span></span> <span data-ttu-id="3b941-540">Для получения дополнительных сведений просмотрите объект диагностики.</span><span class="sxs-lookup"><span data-stu-id="3b941-540">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="3b941-541">Пользователь отключен от сети.</span><span class="sxs-lookup"><span data-stu-id="3b941-541">The user is no longer connected to the network.</span></span> <span data-ttu-id="3b941-542">Проверьте сетевое подключение и повторите попытку.</span><span class="sxs-lookup"><span data-stu-id="3b941-542">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3b941-543">Требования</span><span class="sxs-lookup"><span data-stu-id="3b941-543">Requirements</span></span>

|<span data-ttu-id="3b941-544">Требование</span><span class="sxs-lookup"><span data-stu-id="3b941-544">Requirement</span></span>| <span data-ttu-id="3b941-545">Значение</span><span class="sxs-lookup"><span data-stu-id="3b941-545">Value</span></span>|
|---|---|
|[<span data-ttu-id="3b941-546">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="3b941-546">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3b941-547">1.0</span><span class="sxs-lookup"><span data-stu-id="3b941-547">1.0</span></span>|
|[<span data-ttu-id="3b941-548">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="3b941-548">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3b941-549">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3b941-549">ReadItem</span></span>|
|[<span data-ttu-id="3b941-550">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3b941-550">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3b941-551">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3b941-551">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3b941-552">Пример</span><span class="sxs-lookup"><span data-stu-id="3b941-552">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="3b941-553">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="3b941-553">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="3b941-554">Выполняет асинхронный запрос для веб-служб Exchange (EWS) на сервере Exchange Server, на котором размещен почтовый ящик пользователя.</span><span class="sxs-lookup"><span data-stu-id="3b941-554">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="3b941-555">Этот метод не поддерживается в следующих сценариях:</span><span class="sxs-lookup"><span data-stu-id="3b941-555">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="3b941-556">В Outlook для iOS и Android</span><span class="sxs-lookup"><span data-stu-id="3b941-556">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="3b941-557">Если надстройка загружается в почтовый ящик Gmail.</span><span class="sxs-lookup"><span data-stu-id="3b941-557">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="3b941-558">В таких случаях надстройка должна [использовать REST API](/outlook/add-ins/use-rest-api) для доступа к почтовому ящику пользователя.</span><span class="sxs-lookup"><span data-stu-id="3b941-558">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="3b941-559">Метод `makeEwsRequestAsync` отправляет запрос EWS от имени надстройки в Exchange.</span><span class="sxs-lookup"><span data-stu-id="3b941-559">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="3b941-560">Список поддерживаемых операций EWS см. в статье [Вызов веб-служб из надстройки Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support).</span><span class="sxs-lookup"><span data-stu-id="3b941-560">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="3b941-561">С помощью метода `makeEwsRequestAsync` невозможно запрашивать элементы, связанные с папкой.</span><span class="sxs-lookup"><span data-stu-id="3b941-561">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="3b941-562">В запросе XML должна быть указана кодировка UTF-8.</span><span class="sxs-lookup"><span data-stu-id="3b941-562">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="3b941-p141">У вашей надстройки должно быть разрешение **ReadWriteMailbox** для использования метода `makeEwsRequestAsync`. Сведения об использовании разрешения **ReadWriteMailbox** и операций EWS, которые можно вызывать с помощью метода `makeEwsRequestAsync`, см. в статье [Указание разрешений для доступа почтовой надстройки к почтовому ящику пользователя](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="3b941-p141">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="3b941-565">Администратор сервера должен установить значение true для параметра `OAuthAuthentication` в каталоге сервера клиентского доступа EWS, чтобы метод `makeEwsRequestAsync` мог выполнять запросы EWS.</span><span class="sxs-lookup"><span data-stu-id="3b941-565">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="3b941-566">Различия версий</span><span class="sxs-lookup"><span data-stu-id="3b941-566">Version differences</span></span>

<span data-ttu-id="3b941-567">Если вы используете метод `makeEwsRequestAsync` в почтовых приложениях, которые выполняются в Outlook версии более ранней, чем 15.0.4535.1004, указывайте кодировку `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="3b941-567">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="3b941-p142">Значение кодировки не нужно указывать, если почтовое приложение выполняется в Outlook в Интернете. Чтобы определить, выполняется ли приложение в Outlook или Outlook в Интернете, используйте свойство mailbox.diagnostics.hostName. Используемую версию Outlook можно определить с помощью свойства mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="3b941-p142">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3b941-571">Параметры</span><span class="sxs-lookup"><span data-stu-id="3b941-571">Parameters</span></span>

|<span data-ttu-id="3b941-572">Имя</span><span class="sxs-lookup"><span data-stu-id="3b941-572">Name</span></span>| <span data-ttu-id="3b941-573">Тип</span><span class="sxs-lookup"><span data-stu-id="3b941-573">Type</span></span>| <span data-ttu-id="3b941-574">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="3b941-574">Attributes</span></span>| <span data-ttu-id="3b941-575">Описание</span><span class="sxs-lookup"><span data-stu-id="3b941-575">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="3b941-576">String</span><span class="sxs-lookup"><span data-stu-id="3b941-576">String</span></span>||<span data-ttu-id="3b941-577">Запрос EWS.</span><span class="sxs-lookup"><span data-stu-id="3b941-577">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="3b941-578">function</span><span class="sxs-lookup"><span data-stu-id="3b941-578">function</span></span>||<span data-ttu-id="3b941-579">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="3b941-579">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="3b941-580">Результат XML вызова EWS указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="3b941-580">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="3b941-581">Если размер результата превышает 1 МБ, возвращается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="3b941-581">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="3b941-582">Объект</span><span class="sxs-lookup"><span data-stu-id="3b941-582">Object</span></span>| <span data-ttu-id="3b941-583">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="3b941-583">&lt;optional&gt;</span></span>|<span data-ttu-id="3b941-584">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="3b941-584">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3b941-585">Требования</span><span class="sxs-lookup"><span data-stu-id="3b941-585">Requirements</span></span>

|<span data-ttu-id="3b941-586">Требование</span><span class="sxs-lookup"><span data-stu-id="3b941-586">Requirement</span></span>| <span data-ttu-id="3b941-587">Значение</span><span class="sxs-lookup"><span data-stu-id="3b941-587">Value</span></span>|
|---|---|
|[<span data-ttu-id="3b941-588">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="3b941-588">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3b941-589">1.0</span><span class="sxs-lookup"><span data-stu-id="3b941-589">1.0</span></span>|
|[<span data-ttu-id="3b941-590">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="3b941-590">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3b941-591">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="3b941-591">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="3b941-592">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3b941-592">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3b941-593">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3b941-593">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="3b941-594">Пример</span><span class="sxs-lookup"><span data-stu-id="3b941-594">Example</span></span>

<span data-ttu-id="3b941-595">В приведенном ниже примере вызывается `makeEwsRequestAsync` для получения темы элемента с помощью операции `GetItem`.</span><span class="sxs-lookup"><span data-stu-id="3b941-595">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="3b941-596">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="3b941-596">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="3b941-597">Удаляет обработчиков для поддерживаемого типа события.</span><span class="sxs-lookup"><span data-stu-id="3b941-597">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="3b941-598">В настоящее время единственный поддерживаемый тип события — `Office.EventType.ItemChanged`.</span><span class="sxs-lookup"><span data-stu-id="3b941-598">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="3b941-599">Параметры</span><span class="sxs-lookup"><span data-stu-id="3b941-599">Parameters</span></span>

| <span data-ttu-id="3b941-600">Имя</span><span class="sxs-lookup"><span data-stu-id="3b941-600">Name</span></span> | <span data-ttu-id="3b941-601">Тип</span><span class="sxs-lookup"><span data-stu-id="3b941-601">Type</span></span> | <span data-ttu-id="3b941-602">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="3b941-602">Attributes</span></span> | <span data-ttu-id="3b941-603">Описание</span><span class="sxs-lookup"><span data-stu-id="3b941-603">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="3b941-604">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="3b941-604">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="3b941-605">Событие, которое должно отменить обработчик.</span><span class="sxs-lookup"><span data-stu-id="3b941-605">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="3b941-606">Объект</span><span class="sxs-lookup"><span data-stu-id="3b941-606">Object</span></span> | <span data-ttu-id="3b941-607">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="3b941-607">&lt;optional&gt;</span></span> | <span data-ttu-id="3b941-608">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="3b941-608">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="3b941-609">Object</span><span class="sxs-lookup"><span data-stu-id="3b941-609">Object</span></span> | <span data-ttu-id="3b941-610">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="3b941-610">&lt;optional&gt;</span></span> | <span data-ttu-id="3b941-611">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="3b941-611">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="3b941-612">функция</span><span class="sxs-lookup"><span data-stu-id="3b941-612">function</span></span>| <span data-ttu-id="3b941-613">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="3b941-613">&lt;optional&gt;</span></span>|<span data-ttu-id="3b941-614">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="3b941-614">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="3b941-615">Требования</span><span class="sxs-lookup"><span data-stu-id="3b941-615">Requirements</span></span>

|<span data-ttu-id="3b941-616">Требование</span><span class="sxs-lookup"><span data-stu-id="3b941-616">Requirement</span></span>| <span data-ttu-id="3b941-617">Значение</span><span class="sxs-lookup"><span data-stu-id="3b941-617">Value</span></span>|
|---|---|
|[<span data-ttu-id="3b941-618">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="3b941-618">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3b941-619">1.5</span><span class="sxs-lookup"><span data-stu-id="3b941-619">1.5</span></span> |
|[<span data-ttu-id="3b941-620">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="3b941-620">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3b941-621">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3b941-621">ReadItem</span></span> |
|[<span data-ttu-id="3b941-622">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3b941-622">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="3b941-623">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3b941-623">Compose or Read</span></span>|
