---
title: Office.context.mailbox — набор обязательных элементов 1.5
description: ''
ms.date: 04/24/2019
localization_priority: Priority
ms.openlocfilehash: a0c1a45fd3eaa9cf324a6854120d642eb7520132
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127284"
---
# <a name="mailbox"></a><span data-ttu-id="d06cf-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="d06cf-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="d06cf-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="d06cf-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="d06cf-104">Предоставляет для Microsoft Outlook доступ к объектной модели надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="d06cf-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="d06cf-105">Требования</span><span class="sxs-lookup"><span data-stu-id="d06cf-105">Requirements</span></span>

|<span data-ttu-id="d06cf-106">Требование</span><span class="sxs-lookup"><span data-stu-id="d06cf-106">Requirement</span></span>| <span data-ttu-id="d06cf-107">Значение</span><span class="sxs-lookup"><span data-stu-id="d06cf-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="d06cf-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d06cf-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d06cf-109">1.0</span><span class="sxs-lookup"><span data-stu-id="d06cf-109">1.0</span></span>|
|[<span data-ttu-id="d06cf-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d06cf-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d06cf-111">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="d06cf-111">Restricted</span></span>|
|[<span data-ttu-id="d06cf-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d06cf-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d06cf-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d06cf-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="d06cf-114">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="d06cf-114">Members and methods</span></span>

| <span data-ttu-id="d06cf-115">Элемент</span><span class="sxs-lookup"><span data-stu-id="d06cf-115">Member</span></span> | <span data-ttu-id="d06cf-116">Тип</span><span class="sxs-lookup"><span data-stu-id="d06cf-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="d06cf-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="d06cf-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="d06cf-118">Элемент</span><span class="sxs-lookup"><span data-stu-id="d06cf-118">Member</span></span> |
| [<span data-ttu-id="d06cf-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="d06cf-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="d06cf-120">Элемент</span><span class="sxs-lookup"><span data-stu-id="d06cf-120">Member</span></span> |
| [<span data-ttu-id="d06cf-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="d06cf-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="d06cf-122">Метод</span><span class="sxs-lookup"><span data-stu-id="d06cf-122">Method</span></span> |
| [<span data-ttu-id="d06cf-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="d06cf-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="d06cf-124">Метод</span><span class="sxs-lookup"><span data-stu-id="d06cf-124">Method</span></span> |
| [<span data-ttu-id="d06cf-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="d06cf-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="d06cf-126">Метод</span><span class="sxs-lookup"><span data-stu-id="d06cf-126">Method</span></span> |
| [<span data-ttu-id="d06cf-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="d06cf-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="d06cf-128">Метод</span><span class="sxs-lookup"><span data-stu-id="d06cf-128">Method</span></span> |
| [<span data-ttu-id="d06cf-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="d06cf-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="d06cf-130">Метод</span><span class="sxs-lookup"><span data-stu-id="d06cf-130">Method</span></span> |
| [<span data-ttu-id="d06cf-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="d06cf-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="d06cf-132">Метод</span><span class="sxs-lookup"><span data-stu-id="d06cf-132">Method</span></span> |
| [<span data-ttu-id="d06cf-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="d06cf-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="d06cf-134">Метод</span><span class="sxs-lookup"><span data-stu-id="d06cf-134">Method</span></span> |
| [<span data-ttu-id="d06cf-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="d06cf-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="d06cf-136">Метод</span><span class="sxs-lookup"><span data-stu-id="d06cf-136">Method</span></span> |
| [<span data-ttu-id="d06cf-137">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="d06cf-137">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="d06cf-138">Метод</span><span class="sxs-lookup"><span data-stu-id="d06cf-138">Method</span></span> |
| [<span data-ttu-id="d06cf-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="d06cf-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="d06cf-140">Метод</span><span class="sxs-lookup"><span data-stu-id="d06cf-140">Method</span></span> |
| [<span data-ttu-id="d06cf-141">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="d06cf-141">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="d06cf-142">Метод</span><span class="sxs-lookup"><span data-stu-id="d06cf-142">Method</span></span> |
| [<span data-ttu-id="d06cf-143">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="d06cf-143">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="d06cf-144">Метод</span><span class="sxs-lookup"><span data-stu-id="d06cf-144">Method</span></span> |
| [<span data-ttu-id="d06cf-145">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="d06cf-145">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="d06cf-146">Метод</span><span class="sxs-lookup"><span data-stu-id="d06cf-146">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="d06cf-147">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="d06cf-147">Namespaces</span></span>

<span data-ttu-id="d06cf-148">[diagnostics](Office.context.mailbox.diagnostics.md). Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="d06cf-148">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="d06cf-149">[item](Office.context.mailbox.item.md). Предоставляет методы и свойства для доступа к сообщению или встрече в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="d06cf-149">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="d06cf-150">[userProfile](Office.context.mailbox.userProfile.md). Предоставляет сведения о пользователе в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="d06cf-150">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="d06cf-151">Members</span><span class="sxs-lookup"><span data-stu-id="d06cf-151">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="d06cf-152">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="d06cf-152">ewsUrl :String</span></span>

<span data-ttu-id="d06cf-p101">Получает URL-адрес конечной точки веб-служб Exchange (EWS) для этой учетной записи электронной почты. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="d06cf-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="d06cf-155">Этот элемент не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="d06cf-155">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d06cf-p102">Удаленная служба может использовать значение `ewsUrl`, чтобы выполнять вызовы EWS для почтового ящика пользователя. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="d06cf-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="d06cf-158">Чтобы вызвать элемент `ewsUrl` в режиме чтения, в манифесте приложения должно быть указано разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="d06cf-158">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="d06cf-p103">Перед использованием элемента `ewsUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="d06cf-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="d06cf-161">Тип</span><span class="sxs-lookup"><span data-stu-id="d06cf-161">Type</span></span>

*   <span data-ttu-id="d06cf-162">String</span><span class="sxs-lookup"><span data-stu-id="d06cf-162">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d06cf-163">Требования</span><span class="sxs-lookup"><span data-stu-id="d06cf-163">Requirements</span></span>

|<span data-ttu-id="d06cf-164">Требование</span><span class="sxs-lookup"><span data-stu-id="d06cf-164">Requirement</span></span>| <span data-ttu-id="d06cf-165">Значение</span><span class="sxs-lookup"><span data-stu-id="d06cf-165">Value</span></span>|
|---|---|
|[<span data-ttu-id="d06cf-166">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d06cf-166">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d06cf-167">1.0</span><span class="sxs-lookup"><span data-stu-id="d06cf-167">1.0</span></span>|
|[<span data-ttu-id="d06cf-168">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d06cf-168">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d06cf-169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d06cf-169">ReadItem</span></span>|
|[<span data-ttu-id="d06cf-170">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d06cf-170">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d06cf-171">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d06cf-171">Compose or Read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="d06cf-172">restUrl: String</span><span class="sxs-lookup"><span data-stu-id="d06cf-172">restUrl :String</span></span>

<span data-ttu-id="d06cf-173">Возвращает URL-адрес конечной точки REST для этой учетной записи электронной почты.</span><span class="sxs-lookup"><span data-stu-id="d06cf-173">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="d06cf-174">С помощью значения `restUrl` можно выполнять вызовы [REST API](/outlook/rest/) для почтового ящика пользователя.</span><span class="sxs-lookup"><span data-stu-id="d06cf-174">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="d06cf-175">Чтобы вызвать элемент `restUrl` в режиме чтения, в манифесте приложения необходимо указать разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="d06cf-175">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="d06cf-p104">Перед использованием элемента `restUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="d06cf-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

> [!NOTE]
> <span data-ttu-id="d06cf-178">Клиенты Outlook, подключенные к локальным установленным версиям Exchange 2016 или более поздним с пользовательским URL-адресом REST, возвращают недопустимое значение `restUrl`.</span><span class="sxs-lookup"><span data-stu-id="d06cf-178">Outlook clients connected to on-premises installations of Exchange 2016 or later with a custom REST URL configured will return an invalid value for `restUrl`.</span></span>

##### <a name="type"></a><span data-ttu-id="d06cf-179">Тип</span><span class="sxs-lookup"><span data-stu-id="d06cf-179">Type</span></span>

*   <span data-ttu-id="d06cf-180">String</span><span class="sxs-lookup"><span data-stu-id="d06cf-180">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d06cf-181">Требования</span><span class="sxs-lookup"><span data-stu-id="d06cf-181">Requirements</span></span>

|<span data-ttu-id="d06cf-182">Требование</span><span class="sxs-lookup"><span data-stu-id="d06cf-182">Requirement</span></span>| <span data-ttu-id="d06cf-183">Значение</span><span class="sxs-lookup"><span data-stu-id="d06cf-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="d06cf-184">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="d06cf-184">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d06cf-185">1.5</span><span class="sxs-lookup"><span data-stu-id="d06cf-185">1.5</span></span> |
|[<span data-ttu-id="d06cf-186">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d06cf-186">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d06cf-187">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d06cf-187">ReadItem</span></span>|
|[<span data-ttu-id="d06cf-188">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d06cf-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d06cf-189">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d06cf-189">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="d06cf-190">Методы</span><span class="sxs-lookup"><span data-stu-id="d06cf-190">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="d06cf-191">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d06cf-191">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="d06cf-192">Добавляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="d06cf-192">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="d06cf-193">В настоящее время поддерживаются только события типа `Office.EventType.ItemChanged`, которые вызываются, когда пользователь выбирает новый элемент.</span><span class="sxs-lookup"><span data-stu-id="d06cf-193">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item.</span></span> <span data-ttu-id="d06cf-194">Это событие используется надстройками, реализующими закрепляемую область задач, и позволяет надстройке обновлять пользовательский интерфейс области задач в соответствии с выбранным в данный момент элементом.</span><span class="sxs-lookup"><span data-stu-id="d06cf-194">This event is used by add-ins that implement a pinnable task pane, and allows the add-in to refresh the task pane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d06cf-195">Параметры</span><span class="sxs-lookup"><span data-stu-id="d06cf-195">Parameters</span></span>

| <span data-ttu-id="d06cf-196">Имя</span><span class="sxs-lookup"><span data-stu-id="d06cf-196">Name</span></span> | <span data-ttu-id="d06cf-197">Тип</span><span class="sxs-lookup"><span data-stu-id="d06cf-197">Type</span></span> | <span data-ttu-id="d06cf-198">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="d06cf-198">Attributes</span></span> | <span data-ttu-id="d06cf-199">Описание</span><span class="sxs-lookup"><span data-stu-id="d06cf-199">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="d06cf-200">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="d06cf-200">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="d06cf-201">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="d06cf-201">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="d06cf-202">Function</span><span class="sxs-lookup"><span data-stu-id="d06cf-202">Function</span></span> || <span data-ttu-id="d06cf-p106">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="d06cf-p106">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="d06cf-206">Объект</span><span class="sxs-lookup"><span data-stu-id="d06cf-206">Object</span></span> | <span data-ttu-id="d06cf-207">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d06cf-207">&lt;optional&gt;</span></span> | <span data-ttu-id="d06cf-208">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="d06cf-208">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="d06cf-209">Object</span><span class="sxs-lookup"><span data-stu-id="d06cf-209">Object</span></span> | <span data-ttu-id="d06cf-210">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d06cf-210">&lt;optional&gt;</span></span> | <span data-ttu-id="d06cf-211">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="d06cf-211">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="d06cf-212">функция</span><span class="sxs-lookup"><span data-stu-id="d06cf-212">function</span></span>| <span data-ttu-id="d06cf-213">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d06cf-213">&lt;optional&gt;</span></span>|<span data-ttu-id="d06cf-214">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d06cf-214">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d06cf-215">Требования</span><span class="sxs-lookup"><span data-stu-id="d06cf-215">Requirements</span></span>

|<span data-ttu-id="d06cf-216">Требование</span><span class="sxs-lookup"><span data-stu-id="d06cf-216">Requirement</span></span>| <span data-ttu-id="d06cf-217">Значение</span><span class="sxs-lookup"><span data-stu-id="d06cf-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="d06cf-218">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="d06cf-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d06cf-219">1.5</span><span class="sxs-lookup"><span data-stu-id="d06cf-219">1.5</span></span> |
|[<span data-ttu-id="d06cf-220">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d06cf-220">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d06cf-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d06cf-221">ReadItem</span></span> |
|[<span data-ttu-id="d06cf-222">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d06cf-222">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d06cf-223">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d06cf-223">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d06cf-224">Пример</span><span class="sxs-lookup"><span data-stu-id="d06cf-224">Example</span></span>

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
};
```

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="d06cf-225">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="d06cf-225">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="d06cf-226">Преобразовывает идентификатор элемента из формата REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="d06cf-226">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="d06cf-227">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="d06cf-227">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d06cf-p107">Формат идентификаторов, извлекаемых через API REST (например, [API Почты Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)), отличается от формата веб-служб Exchange (EWS). Метод `convertToEwsId` преобразовывает идентификатор в формате REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="d06cf-p107">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d06cf-230">Параметры</span><span class="sxs-lookup"><span data-stu-id="d06cf-230">Parameters</span></span>

|<span data-ttu-id="d06cf-231">Имя</span><span class="sxs-lookup"><span data-stu-id="d06cf-231">Name</span></span>| <span data-ttu-id="d06cf-232">Тип</span><span class="sxs-lookup"><span data-stu-id="d06cf-232">Type</span></span>| <span data-ttu-id="d06cf-233">Описание</span><span class="sxs-lookup"><span data-stu-id="d06cf-233">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="d06cf-234">String</span><span class="sxs-lookup"><span data-stu-id="d06cf-234">String</span></span>|<span data-ttu-id="d06cf-235">Идентификатор элемента в формате REST API для Outlook</span><span class="sxs-lookup"><span data-stu-id="d06cf-235">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="d06cf-236">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="d06cf-236">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.restversion)|<span data-ttu-id="d06cf-237">Значение, определяющее версию REST API для Outlook, которая используется для извлечения идентификатора элемента.</span><span class="sxs-lookup"><span data-stu-id="d06cf-237">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d06cf-238">Требования</span><span class="sxs-lookup"><span data-stu-id="d06cf-238">Requirements</span></span>

|<span data-ttu-id="d06cf-239">Требование</span><span class="sxs-lookup"><span data-stu-id="d06cf-239">Requirement</span></span>| <span data-ttu-id="d06cf-240">Значение</span><span class="sxs-lookup"><span data-stu-id="d06cf-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="d06cf-241">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="d06cf-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d06cf-242">1.3</span><span class="sxs-lookup"><span data-stu-id="d06cf-242">1.3</span></span>|
|[<span data-ttu-id="d06cf-243">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d06cf-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d06cf-244">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="d06cf-244">Restricted</span></span>|
|[<span data-ttu-id="d06cf-245">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d06cf-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d06cf-246">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d06cf-246">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d06cf-247">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="d06cf-247">Returns:</span></span>

<span data-ttu-id="d06cf-248">Тип: String</span><span class="sxs-lookup"><span data-stu-id="d06cf-248">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="d06cf-249">Пример</span><span class="sxs-lookup"><span data-stu-id="d06cf-249">Example</span></span>

```javascript
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook15officelocalclienttime"></a><span data-ttu-id="d06cf-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="d06cf-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)}</span></span>

<span data-ttu-id="d06cf-251">Получает словарь, содержащий сведения о локальном времени клиента.</span><span class="sxs-lookup"><span data-stu-id="d06cf-251">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="d06cf-p108">Почтовое приложение для классической версии Outlook или версии в Интернете может использовать разные часовые пояса для дат и времени. Классическое приложение Outlook использует часовой пояс клиентского компьютера. Outlook в Интернете использует часовой пояс, заданный в Центре администрирования Exchange (EAC). Значения даты и времени должны обрабатываться так, чтобы значения в пользовательском интерфейсе всегда согласовывались с часовым поясом, ожидаемым пользователем.</span><span class="sxs-lookup"><span data-stu-id="d06cf-p108">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="d06cf-p109">Если почтовое приложение работает в классическом клиенте Outlook, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса клиентского компьютера. Если почтовое приложение работает в Outlook в Интернете, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса, заданного в Центре администрирования Exchange.</span><span class="sxs-lookup"><span data-stu-id="d06cf-p109">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d06cf-257">Параметры</span><span class="sxs-lookup"><span data-stu-id="d06cf-257">Parameters</span></span>

|<span data-ttu-id="d06cf-258">Имя</span><span class="sxs-lookup"><span data-stu-id="d06cf-258">Name</span></span>| <span data-ttu-id="d06cf-259">Тип</span><span class="sxs-lookup"><span data-stu-id="d06cf-259">Type</span></span>| <span data-ttu-id="d06cf-260">Описание</span><span class="sxs-lookup"><span data-stu-id="d06cf-260">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="d06cf-261">Date</span><span class="sxs-lookup"><span data-stu-id="d06cf-261">Date</span></span>|<span data-ttu-id="d06cf-262">Объект Date</span><span class="sxs-lookup"><span data-stu-id="d06cf-262">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d06cf-263">Требования</span><span class="sxs-lookup"><span data-stu-id="d06cf-263">Requirements</span></span>

|<span data-ttu-id="d06cf-264">Требование</span><span class="sxs-lookup"><span data-stu-id="d06cf-264">Requirement</span></span>| <span data-ttu-id="d06cf-265">Значение</span><span class="sxs-lookup"><span data-stu-id="d06cf-265">Value</span></span>|
|---|---|
|[<span data-ttu-id="d06cf-266">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d06cf-266">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d06cf-267">1.0</span><span class="sxs-lookup"><span data-stu-id="d06cf-267">1.0</span></span>|
|[<span data-ttu-id="d06cf-268">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d06cf-268">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d06cf-269">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d06cf-269">ReadItem</span></span>|
|[<span data-ttu-id="d06cf-270">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d06cf-270">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d06cf-271">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d06cf-271">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d06cf-272">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="d06cf-272">Returns:</span></span>

<span data-ttu-id="d06cf-273">Тип: [LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="d06cf-273">Type: [LocalClientTime](/javascript/api/outlook_1_5/office.LocalClientTime)</span></span>

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="d06cf-274">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="d06cf-274">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="d06cf-275">Преобразовывает идентификатор элемента в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="d06cf-275">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="d06cf-276">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="d06cf-276">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d06cf-p110">Формат идентификаторов, извлекаемых через EWS или свойство `itemId`, отличается от формата API REST (таких как [API Почты Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)). Метод `convertToRestId` преобразовывает идентификатор в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="d06cf-p110">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d06cf-279">Параметры</span><span class="sxs-lookup"><span data-stu-id="d06cf-279">Parameters</span></span>

|<span data-ttu-id="d06cf-280">Имя</span><span class="sxs-lookup"><span data-stu-id="d06cf-280">Name</span></span>| <span data-ttu-id="d06cf-281">Тип</span><span class="sxs-lookup"><span data-stu-id="d06cf-281">Type</span></span>| <span data-ttu-id="d06cf-282">Описание</span><span class="sxs-lookup"><span data-stu-id="d06cf-282">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="d06cf-283">String</span><span class="sxs-lookup"><span data-stu-id="d06cf-283">String</span></span>|<span data-ttu-id="d06cf-284">Идентификатор элемента в формате EWS</span><span class="sxs-lookup"><span data-stu-id="d06cf-284">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="d06cf-285">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="d06cf-285">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_5/office.mailboxenums.restversion)|<span data-ttu-id="d06cf-286">Значение, определяющее версию REST API для Outlook, с которой будет использоваться преобразованный идентификатор.</span><span class="sxs-lookup"><span data-stu-id="d06cf-286">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d06cf-287">Требования</span><span class="sxs-lookup"><span data-stu-id="d06cf-287">Requirements</span></span>

|<span data-ttu-id="d06cf-288">Требование</span><span class="sxs-lookup"><span data-stu-id="d06cf-288">Requirement</span></span>| <span data-ttu-id="d06cf-289">Значение</span><span class="sxs-lookup"><span data-stu-id="d06cf-289">Value</span></span>|
|---|---|
|[<span data-ttu-id="d06cf-290">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="d06cf-290">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d06cf-291">1.3</span><span class="sxs-lookup"><span data-stu-id="d06cf-291">1.3</span></span>|
|[<span data-ttu-id="d06cf-292">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d06cf-292">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d06cf-293">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="d06cf-293">Restricted</span></span>|
|[<span data-ttu-id="d06cf-294">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d06cf-294">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d06cf-295">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d06cf-295">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d06cf-296">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="d06cf-296">Returns:</span></span>

<span data-ttu-id="d06cf-297">Тип: String</span><span class="sxs-lookup"><span data-stu-id="d06cf-297">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="d06cf-298">Пример</span><span class="sxs-lookup"><span data-stu-id="d06cf-298">Example</span></span>

```javascript
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="d06cf-299">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="d06cf-299">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="d06cf-300">Получает объект Date из словаря, содержащего сведения о времени.</span><span class="sxs-lookup"><span data-stu-id="d06cf-300">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="d06cf-301">Метод `convertToUtcClientTime` преобразует словарь, содержащий локальную дату и время, в объект Date с правильными значениями локальной даты и времени.</span><span class="sxs-lookup"><span data-stu-id="d06cf-301">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d06cf-302">Параметры</span><span class="sxs-lookup"><span data-stu-id="d06cf-302">Parameters</span></span>

|<span data-ttu-id="d06cf-303">Имя</span><span class="sxs-lookup"><span data-stu-id="d06cf-303">Name</span></span>| <span data-ttu-id="d06cf-304">Тип</span><span class="sxs-lookup"><span data-stu-id="d06cf-304">Type</span></span>| <span data-ttu-id="d06cf-305">Описание</span><span class="sxs-lookup"><span data-stu-id="d06cf-305">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="d06cf-306">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="d06cf-306">LocalClientTime</span></span>](/javascript/api/outlook_1_5/office.LocalClientTime)|<span data-ttu-id="d06cf-307">Значение локального времени для преобразования.</span><span class="sxs-lookup"><span data-stu-id="d06cf-307">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d06cf-308">Требования</span><span class="sxs-lookup"><span data-stu-id="d06cf-308">Requirements</span></span>

|<span data-ttu-id="d06cf-309">Требование</span><span class="sxs-lookup"><span data-stu-id="d06cf-309">Requirement</span></span>| <span data-ttu-id="d06cf-310">Значение</span><span class="sxs-lookup"><span data-stu-id="d06cf-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="d06cf-311">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d06cf-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d06cf-312">1.0</span><span class="sxs-lookup"><span data-stu-id="d06cf-312">1.0</span></span>|
|[<span data-ttu-id="d06cf-313">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d06cf-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d06cf-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d06cf-314">ReadItem</span></span>|
|[<span data-ttu-id="d06cf-315">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d06cf-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d06cf-316">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d06cf-316">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="d06cf-317">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="d06cf-317">Returns:</span></span>

<span data-ttu-id="d06cf-318">Объект Date со временем в формате UTC.</span><span class="sxs-lookup"><span data-stu-id="d06cf-318">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="d06cf-319">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="d06cf-319">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="d06cf-320">Date</span><span class="sxs-lookup"><span data-stu-id="d06cf-320">Date</span></span></dd>

</dl>

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="d06cf-321">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="d06cf-321">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="d06cf-322">Отображает имеющуюся встречу из календаря.</span><span class="sxs-lookup"><span data-stu-id="d06cf-322">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="d06cf-323">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="d06cf-323">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d06cf-324">Метод `displayAppointmentForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее сведения календаря о существующей встрече.</span><span class="sxs-lookup"><span data-stu-id="d06cf-324">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="d06cf-p111">В Outlook для Mac с помощью этого метода можно отобразить одну встречу, которая не является частью повторяющегося ряда, или основную встречу такого ряда, но не экземпляр из него, так как в Outlook для Mac невозможно получить доступ к свойствам экземпляра повторяющегося ряда (в том числе к идентификатору элемента).</span><span class="sxs-lookup"><span data-stu-id="d06cf-p111">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="d06cf-327">В Outlook в Интернете этот метод открывает указанную форму, только если текст формы содержит символы размером не более 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="d06cf-327">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="d06cf-328">Если указанный идентификатор элемента не определяет существующую встречу, на клиентском компьютере или устройстве открывается пустая страница, и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="d06cf-328">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d06cf-329">Параметры</span><span class="sxs-lookup"><span data-stu-id="d06cf-329">Parameters</span></span>

|<span data-ttu-id="d06cf-330">Имя</span><span class="sxs-lookup"><span data-stu-id="d06cf-330">Name</span></span>| <span data-ttu-id="d06cf-331">Тип</span><span class="sxs-lookup"><span data-stu-id="d06cf-331">Type</span></span>| <span data-ttu-id="d06cf-332">Описание</span><span class="sxs-lookup"><span data-stu-id="d06cf-332">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="d06cf-333">String</span><span class="sxs-lookup"><span data-stu-id="d06cf-333">String</span></span>|<span data-ttu-id="d06cf-334">Идентификатор веб-служб Exchange для существующей встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="d06cf-334">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d06cf-335">Требования</span><span class="sxs-lookup"><span data-stu-id="d06cf-335">Requirements</span></span>

|<span data-ttu-id="d06cf-336">Требование</span><span class="sxs-lookup"><span data-stu-id="d06cf-336">Requirement</span></span>| <span data-ttu-id="d06cf-337">Значение</span><span class="sxs-lookup"><span data-stu-id="d06cf-337">Value</span></span>|
|---|---|
|[<span data-ttu-id="d06cf-338">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d06cf-338">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d06cf-339">1.0</span><span class="sxs-lookup"><span data-stu-id="d06cf-339">1.0</span></span>|
|[<span data-ttu-id="d06cf-340">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d06cf-340">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d06cf-341">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d06cf-341">ReadItem</span></span>|
|[<span data-ttu-id="d06cf-342">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d06cf-342">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d06cf-343">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d06cf-343">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d06cf-344">Пример</span><span class="sxs-lookup"><span data-stu-id="d06cf-344">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

#### <a name="displaymessageformitemid"></a><span data-ttu-id="d06cf-345">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="d06cf-345">displayMessageForm(itemId)</span></span>

<span data-ttu-id="d06cf-346">Отображает имеющееся сообщение.</span><span class="sxs-lookup"><span data-stu-id="d06cf-346">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="d06cf-347">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="d06cf-347">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d06cf-348">Метод `displayMessageForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее существующее сообщение.</span><span class="sxs-lookup"><span data-stu-id="d06cf-348">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="d06cf-349">В Outlook в Интернете этот метод открывает указанную форму, только если текст формы содержит символы размером не более 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="d06cf-349">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="d06cf-350">Если указанный идентификатор элемента не определяет существующее сообщение, окно на клиентском компьютере не открывается и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="d06cf-350">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="d06cf-p112">Не используйте `displayMessageForm` с параметром `itemId`, который представляет собой встречу. Используйте метод `displayAppointmentForm`, чтобы отобразить сведения о существующей встрече, а метод `displayNewAppointmentForm` — для отображения формы создания встречи.</span><span class="sxs-lookup"><span data-stu-id="d06cf-p112">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d06cf-353">Параметры</span><span class="sxs-lookup"><span data-stu-id="d06cf-353">Parameters</span></span>

|<span data-ttu-id="d06cf-354">Имя</span><span class="sxs-lookup"><span data-stu-id="d06cf-354">Name</span></span>| <span data-ttu-id="d06cf-355">Тип</span><span class="sxs-lookup"><span data-stu-id="d06cf-355">Type</span></span>| <span data-ttu-id="d06cf-356">Описание</span><span class="sxs-lookup"><span data-stu-id="d06cf-356">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="d06cf-357">String</span><span class="sxs-lookup"><span data-stu-id="d06cf-357">String</span></span>|<span data-ttu-id="d06cf-358">Идентификатор веб-служб Exchange для существующего сообщения.</span><span class="sxs-lookup"><span data-stu-id="d06cf-358">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d06cf-359">Требования</span><span class="sxs-lookup"><span data-stu-id="d06cf-359">Requirements</span></span>

|<span data-ttu-id="d06cf-360">Требование</span><span class="sxs-lookup"><span data-stu-id="d06cf-360">Requirement</span></span>| <span data-ttu-id="d06cf-361">Значение</span><span class="sxs-lookup"><span data-stu-id="d06cf-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="d06cf-362">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d06cf-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d06cf-363">1.0</span><span class="sxs-lookup"><span data-stu-id="d06cf-363">1.0</span></span>|
|[<span data-ttu-id="d06cf-364">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d06cf-364">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d06cf-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d06cf-365">ReadItem</span></span>|
|[<span data-ttu-id="d06cf-366">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d06cf-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d06cf-367">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d06cf-367">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d06cf-368">Пример</span><span class="sxs-lookup"><span data-stu-id="d06cf-368">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="d06cf-369">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="d06cf-369">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="d06cf-370">Отображает форму для создания новой встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="d06cf-370">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="d06cf-371">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="d06cf-371">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="d06cf-p113">Метод `displayNewAppointmentForm` открывает форму, в которой пользователь может создать встречу или собрание. Если параметры заданы, поля формы встречи автоматически заполняются их содержимым.</span><span class="sxs-lookup"><span data-stu-id="d06cf-p113">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="d06cf-p114">В Outlook в Интернете и на мобильных устройствах этот метод всегда отображает форму с полем участников. Если вы не укажете участников в качестве входных аргументов, метод отображает форму с кнопкой **Сохранить**. Если вы укажете участников, форма будет включать участников и кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="d06cf-p114">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="d06cf-p115">Если вы укажете участников или ресурсы с помощью параметра `requiredAttendees`, `optionalAttendees` или `resources` в клиенте Outlook с расширенными возможностями и Outlook RT, этот метод отобразит форму собрания с кнопкой **Отправить**. Если не указать получателей, этот метод отобразит форму встречи с кнопкой **Сохранить и закрыть**.</span><span class="sxs-lookup"><span data-stu-id="d06cf-p115">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="d06cf-379">Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.</span><span class="sxs-lookup"><span data-stu-id="d06cf-379">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d06cf-380">Параметры</span><span class="sxs-lookup"><span data-stu-id="d06cf-380">Parameters</span></span>

|<span data-ttu-id="d06cf-381">Имя</span><span class="sxs-lookup"><span data-stu-id="d06cf-381">Name</span></span>| <span data-ttu-id="d06cf-382">Тип</span><span class="sxs-lookup"><span data-stu-id="d06cf-382">Type</span></span>| <span data-ttu-id="d06cf-383">Описание</span><span class="sxs-lookup"><span data-stu-id="d06cf-383">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="d06cf-384">Object</span><span class="sxs-lookup"><span data-stu-id="d06cf-384">Object</span></span> | <span data-ttu-id="d06cf-385">Словарь параметров, описывающий новую встречу.</span><span class="sxs-lookup"><span data-stu-id="d06cf-385">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="d06cf-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="d06cf-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="d06cf-p116">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из обязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="d06cf-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="d06cf-389">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="d06cf-389">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_5/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="d06cf-p117">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из необязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="d06cf-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="d06cf-392">Date</span><span class="sxs-lookup"><span data-stu-id="d06cf-392">Date</span></span> | <span data-ttu-id="d06cf-393">Объект `Date`, указывающий дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="d06cf-393">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="d06cf-394">Date</span><span class="sxs-lookup"><span data-stu-id="d06cf-394">Date</span></span> | <span data-ttu-id="d06cf-395">Объект `Date`, указывающий дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="d06cf-395">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="d06cf-396">String</span><span class="sxs-lookup"><span data-stu-id="d06cf-396">String</span></span> | <span data-ttu-id="d06cf-p118">Строка со сведениями о месте встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="d06cf-p118">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="d06cf-399">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="d06cf-399">Array.&lt;String&gt;</span></span> | <span data-ttu-id="d06cf-p119">Массив строк, содержащий необходимые для встречи ресурсы. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="d06cf-p119">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="d06cf-402">String</span><span class="sxs-lookup"><span data-stu-id="d06cf-402">String</span></span> | <span data-ttu-id="d06cf-p120">Строка с темой встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="d06cf-p120">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="d06cf-405">String</span><span class="sxs-lookup"><span data-stu-id="d06cf-405">String</span></span> | <span data-ttu-id="d06cf-p121">Текст сообщения о встрече. Максимальный размер содержимого сообщения — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="d06cf-p121">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d06cf-408">Требования</span><span class="sxs-lookup"><span data-stu-id="d06cf-408">Requirements</span></span>

|<span data-ttu-id="d06cf-409">Требование</span><span class="sxs-lookup"><span data-stu-id="d06cf-409">Requirement</span></span>| <span data-ttu-id="d06cf-410">Значение</span><span class="sxs-lookup"><span data-stu-id="d06cf-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="d06cf-411">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d06cf-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d06cf-412">1.0</span><span class="sxs-lookup"><span data-stu-id="d06cf-412">1.0</span></span>|
|[<span data-ttu-id="d06cf-413">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d06cf-413">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d06cf-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d06cf-414">ReadItem</span></span>|
|[<span data-ttu-id="d06cf-415">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d06cf-415">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d06cf-416">Чтение</span><span class="sxs-lookup"><span data-stu-id="d06cf-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d06cf-417">Пример</span><span class="sxs-lookup"><span data-stu-id="d06cf-417">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="d06cf-418">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="d06cf-418">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="d06cf-419">Возвращает строку, содержащую маркер, который используется для вызова интерфейсов REST API или веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="d06cf-419">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="d06cf-p122">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="d06cf-p122">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="d06cf-422">Рекомендуем сделать так, чтобы по мере возможности надстройки использовали интерфейсы REST API, а не веб-службы Exchange.</span><span class="sxs-lookup"><span data-stu-id="d06cf-422">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="d06cf-423">**Маркеры REST**</span><span class="sxs-lookup"><span data-stu-id="d06cf-423">**REST Tokens**</span></span>

<span data-ttu-id="d06cf-p123">Если запрашивается маркер REST (`options.isRest = true`), полученный маркер не подойдет для проверки подлинности при вызовах веб-служб Exchange. Область действия маркера будет ограничена доступом только для чтения к текущему элементу и его вложениям, если в манифесте надстройки не указано разрешение [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission). Если указано разрешение `ReadWriteMailbox`, полученный маркер предоставит доступ на чтение и запись к почте, календарю и контактам, включая возможность отправки почты.</span><span class="sxs-lookup"><span data-stu-id="d06cf-p123">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="d06cf-427">С помощью свойства `restUrl` надстройка должна определить правильный URL-адрес для вызовов REST API.</span><span class="sxs-lookup"><span data-stu-id="d06cf-427">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="d06cf-428">**Маркеры EWS**</span><span class="sxs-lookup"><span data-stu-id="d06cf-428">**EWS Tokens**</span></span>

<span data-ttu-id="d06cf-p124">Если запрашивается маркер EWS (`options.isRest = false`), полученный маркер не подойдет для проверки подлинности при вызовах REST API. Область действия маркера будет ограничена доступом к текущему элементу.</span><span class="sxs-lookup"><span data-stu-id="d06cf-p124">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="d06cf-431">С помощью свойства `ewsUrl` надстройка должна определить правильный URL-адрес для вызовов EWS.</span><span class="sxs-lookup"><span data-stu-id="d06cf-431">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d06cf-432">Параметры</span><span class="sxs-lookup"><span data-stu-id="d06cf-432">Parameters</span></span>

|<span data-ttu-id="d06cf-433">Имя</span><span class="sxs-lookup"><span data-stu-id="d06cf-433">Name</span></span>| <span data-ttu-id="d06cf-434">Тип</span><span class="sxs-lookup"><span data-stu-id="d06cf-434">Type</span></span>| <span data-ttu-id="d06cf-435">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="d06cf-435">Attributes</span></span>| <span data-ttu-id="d06cf-436">Описание</span><span class="sxs-lookup"><span data-stu-id="d06cf-436">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="d06cf-437">Object</span><span class="sxs-lookup"><span data-stu-id="d06cf-437">Object</span></span> | <span data-ttu-id="d06cf-438">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d06cf-438">&lt;optional&gt;</span></span> | <span data-ttu-id="d06cf-439">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="d06cf-439">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="d06cf-440">Boolean</span><span class="sxs-lookup"><span data-stu-id="d06cf-440">Boolean</span></span> |  <span data-ttu-id="d06cf-441">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d06cf-441">&lt;optional&gt;</span></span> | <span data-ttu-id="d06cf-p125">Определяет, будет ли предоставленный маркер использоваться для интерфейсов REST API Outlook или веб-служб Exchange. Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="d06cf-p125">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="d06cf-444">Объект</span><span class="sxs-lookup"><span data-stu-id="d06cf-444">Object</span></span> |  <span data-ttu-id="d06cf-445">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d06cf-445">&lt;optional&gt;</span></span> | <span data-ttu-id="d06cf-446">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="d06cf-446">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="d06cf-447">function</span><span class="sxs-lookup"><span data-stu-id="d06cf-447">function</span></span>||<span data-ttu-id="d06cf-p126">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult). Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="d06cf-p126">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d06cf-450">Требования</span><span class="sxs-lookup"><span data-stu-id="d06cf-450">Requirements</span></span>

|<span data-ttu-id="d06cf-451">Требование</span><span class="sxs-lookup"><span data-stu-id="d06cf-451">Requirement</span></span>| <span data-ttu-id="d06cf-452">Значение</span><span class="sxs-lookup"><span data-stu-id="d06cf-452">Value</span></span>|
|---|---|
|[<span data-ttu-id="d06cf-453">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="d06cf-453">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d06cf-454">1.5</span><span class="sxs-lookup"><span data-stu-id="d06cf-454">1.5</span></span> |
|[<span data-ttu-id="d06cf-455">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d06cf-455">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d06cf-456">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d06cf-456">ReadItem</span></span>|
|[<span data-ttu-id="d06cf-457">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d06cf-457">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d06cf-458">Создание и чтение</span><span class="sxs-lookup"><span data-stu-id="d06cf-458">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="d06cf-459">Пример</span><span class="sxs-lookup"><span data-stu-id="d06cf-459">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="d06cf-460">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="d06cf-460">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="d06cf-461">Получает строку, содержащую маркер, используемый для получения вложения или элемента с Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="d06cf-461">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="d06cf-p127">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="d06cf-p127">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="d06cf-p128">Вы можете передать сторонней системе маркер и идентификатор вложения или элемента. Сторонняя система использует этот маркер как маркер авторизации, чтобы вызвать операцию [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) или [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) веб-служб Exchange для возврата вложения или элемента. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="d06cf-p128">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="d06cf-467">Для вызова метода `getCallbackTokenAsync` в режиме чтения манифесте приложения должно быть указано разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="d06cf-467">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="d06cf-p129">Чтобы получить идентификатор элемента для передачи в метод `getCallbackTokenAsync`, в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="d06cf-p129">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d06cf-470">Параметры</span><span class="sxs-lookup"><span data-stu-id="d06cf-470">Parameters</span></span>

|<span data-ttu-id="d06cf-471">Имя</span><span class="sxs-lookup"><span data-stu-id="d06cf-471">Name</span></span>| <span data-ttu-id="d06cf-472">Тип</span><span class="sxs-lookup"><span data-stu-id="d06cf-472">Type</span></span>| <span data-ttu-id="d06cf-473">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="d06cf-473">Attributes</span></span>| <span data-ttu-id="d06cf-474">Описание</span><span class="sxs-lookup"><span data-stu-id="d06cf-474">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="d06cf-475">function</span><span class="sxs-lookup"><span data-stu-id="d06cf-475">function</span></span>||<span data-ttu-id="d06cf-p130">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult). Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="d06cf-p130">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="d06cf-478">Объект</span><span class="sxs-lookup"><span data-stu-id="d06cf-478">Object</span></span>| <span data-ttu-id="d06cf-479">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d06cf-479">&lt;optional&gt;</span></span>|<span data-ttu-id="d06cf-480">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="d06cf-480">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d06cf-481">Требования</span><span class="sxs-lookup"><span data-stu-id="d06cf-481">Requirements</span></span>

|<span data-ttu-id="d06cf-482">Требование</span><span class="sxs-lookup"><span data-stu-id="d06cf-482">Requirement</span></span>| <span data-ttu-id="d06cf-483">Значение</span><span class="sxs-lookup"><span data-stu-id="d06cf-483">Value</span></span>|
|---|---|
|[<span data-ttu-id="d06cf-484">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="d06cf-484">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d06cf-485">1.3</span><span class="sxs-lookup"><span data-stu-id="d06cf-485">1.3</span></span>|
|[<span data-ttu-id="d06cf-486">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d06cf-486">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d06cf-487">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d06cf-487">ReadItem</span></span>|
|[<span data-ttu-id="d06cf-488">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d06cf-488">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d06cf-489">Создание и чтение</span><span class="sxs-lookup"><span data-stu-id="d06cf-489">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="d06cf-490">Пример</span><span class="sxs-lookup"><span data-stu-id="d06cf-490">Example</span></span>

```javascript
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="d06cf-491">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="d06cf-491">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="d06cf-492">Получает маркер, идентифицирующий пользователя и надстройку Office.</span><span class="sxs-lookup"><span data-stu-id="d06cf-492">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="d06cf-493">Метод `getUserIdentityTokenAsync` возвращает маркер, который можно использовать для идентификации, а также [проверки подлинности надстройки и пользователя в сторонней системе](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="d06cf-493">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="d06cf-494">Параметры</span><span class="sxs-lookup"><span data-stu-id="d06cf-494">Parameters</span></span>

|<span data-ttu-id="d06cf-495">Имя</span><span class="sxs-lookup"><span data-stu-id="d06cf-495">Name</span></span>| <span data-ttu-id="d06cf-496">Тип</span><span class="sxs-lookup"><span data-stu-id="d06cf-496">Type</span></span>| <span data-ttu-id="d06cf-497">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="d06cf-497">Attributes</span></span>| <span data-ttu-id="d06cf-498">Описание</span><span class="sxs-lookup"><span data-stu-id="d06cf-498">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="d06cf-499">function</span><span class="sxs-lookup"><span data-stu-id="d06cf-499">function</span></span>||<span data-ttu-id="d06cf-500">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d06cf-500">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d06cf-501">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="d06cf-501">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="d06cf-502">Object</span><span class="sxs-lookup"><span data-stu-id="d06cf-502">Object</span></span>| <span data-ttu-id="d06cf-503">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d06cf-503">&lt;optional&gt;</span></span>|<span data-ttu-id="d06cf-504">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="d06cf-504">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d06cf-505">Требования</span><span class="sxs-lookup"><span data-stu-id="d06cf-505">Requirements</span></span>

|<span data-ttu-id="d06cf-506">Требование</span><span class="sxs-lookup"><span data-stu-id="d06cf-506">Requirement</span></span>| <span data-ttu-id="d06cf-507">Значение</span><span class="sxs-lookup"><span data-stu-id="d06cf-507">Value</span></span>|
|---|---|
|[<span data-ttu-id="d06cf-508">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d06cf-508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d06cf-509">1.0</span><span class="sxs-lookup"><span data-stu-id="d06cf-509">1.0</span></span>|
|[<span data-ttu-id="d06cf-510">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d06cf-510">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d06cf-511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d06cf-511">ReadItem</span></span>|
|[<span data-ttu-id="d06cf-512">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d06cf-512">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d06cf-513">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d06cf-513">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d06cf-514">Пример</span><span class="sxs-lookup"><span data-stu-id="d06cf-514">Example</span></span>

```javascript
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="d06cf-515">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="d06cf-515">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="d06cf-516">Выполняет асинхронный запрос для веб-служб Exchange (EWS) на сервере Exchange Server, на котором размещен почтовый ящик пользователя.</span><span class="sxs-lookup"><span data-stu-id="d06cf-516">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="d06cf-517">Этот метод не поддерживается в следующих сценариях:</span><span class="sxs-lookup"><span data-stu-id="d06cf-517">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="d06cf-518">В Outlook для iOS и Android</span><span class="sxs-lookup"><span data-stu-id="d06cf-518">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="d06cf-519">Если надстройка загружается в почтовый ящик Gmail.</span><span class="sxs-lookup"><span data-stu-id="d06cf-519">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="d06cf-520">В таких случаях надстройка должна [использовать REST API](/outlook/add-ins/use-rest-api) для доступа к почтовому ящику пользователя.</span><span class="sxs-lookup"><span data-stu-id="d06cf-520">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="d06cf-521">Метод `makeEwsRequestAsync` отправляет запрос EWS от имени надстройки в Exchange.</span><span class="sxs-lookup"><span data-stu-id="d06cf-521">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="d06cf-522">Список поддерживаемых операций EWS см. в статье [Вызов веб-служб из надстройки Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support).</span><span class="sxs-lookup"><span data-stu-id="d06cf-522">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="d06cf-523">С помощью метода `makeEwsRequestAsync` невозможно запрашивать элементы, связанные с папкой.</span><span class="sxs-lookup"><span data-stu-id="d06cf-523">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="d06cf-524">В запросе XML должна быть указана кодировка UTF-8.</span><span class="sxs-lookup"><span data-stu-id="d06cf-524">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="d06cf-p132">У вашей надстройки должно быть разрешение **ReadWriteMailbox** для использования метода `makeEwsRequestAsync`. Сведения об использовании разрешения **ReadWriteMailbox** и операций EWS, которые можно вызывать с помощью метода `makeEwsRequestAsync`, см. в статье [Указание разрешений для доступа почтовой надстройки к почтовому ящику пользователя](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="d06cf-p132">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="d06cf-527">Администратор сервера должен установить значение true для параметра `OAuthAuthentication` в каталоге сервера клиентского доступа EWS, чтобы метод `makeEwsRequestAsync` мог выполнять запросы EWS.</span><span class="sxs-lookup"><span data-stu-id="d06cf-527">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="d06cf-528">Различия версий</span><span class="sxs-lookup"><span data-stu-id="d06cf-528">Version differences</span></span>

<span data-ttu-id="d06cf-529">Если вы используете метод `makeEwsRequestAsync` в почтовых приложениях, которые выполняются в Outlook версии более ранней, чем 15.0.4535.1004, указывайте кодировку `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="d06cf-529">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="d06cf-p133">Значение кодировки не нужно указывать, если почтовое приложение выполняется в Outlook в Интернете. Чтобы определить, выполняется ли приложение в Outlook или Outlook в Интернете, используйте свойство mailbox.diagnostics.hostName. Используемую версию Outlook можно определить с помощью свойства mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="d06cf-p133">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d06cf-533">Параметры</span><span class="sxs-lookup"><span data-stu-id="d06cf-533">Parameters</span></span>

|<span data-ttu-id="d06cf-534">Имя</span><span class="sxs-lookup"><span data-stu-id="d06cf-534">Name</span></span>| <span data-ttu-id="d06cf-535">Тип</span><span class="sxs-lookup"><span data-stu-id="d06cf-535">Type</span></span>| <span data-ttu-id="d06cf-536">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="d06cf-536">Attributes</span></span>| <span data-ttu-id="d06cf-537">Описание</span><span class="sxs-lookup"><span data-stu-id="d06cf-537">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="d06cf-538">String</span><span class="sxs-lookup"><span data-stu-id="d06cf-538">String</span></span>||<span data-ttu-id="d06cf-539">Запрос EWS.</span><span class="sxs-lookup"><span data-stu-id="d06cf-539">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="d06cf-540">function</span><span class="sxs-lookup"><span data-stu-id="d06cf-540">function</span></span>||<span data-ttu-id="d06cf-541">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d06cf-541">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="d06cf-542">Результат XML вызова EWS указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="d06cf-542">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="d06cf-543">Если размер результата превышает 1 МБ, возвращается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="d06cf-543">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="d06cf-544">Объект</span><span class="sxs-lookup"><span data-stu-id="d06cf-544">Object</span></span>| <span data-ttu-id="d06cf-545">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d06cf-545">&lt;optional&gt;</span></span>|<span data-ttu-id="d06cf-546">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="d06cf-546">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d06cf-547">Требования</span><span class="sxs-lookup"><span data-stu-id="d06cf-547">Requirements</span></span>

|<span data-ttu-id="d06cf-548">Требование</span><span class="sxs-lookup"><span data-stu-id="d06cf-548">Requirement</span></span>| <span data-ttu-id="d06cf-549">Значение</span><span class="sxs-lookup"><span data-stu-id="d06cf-549">Value</span></span>|
|---|---|
|[<span data-ttu-id="d06cf-550">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d06cf-550">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d06cf-551">1.0</span><span class="sxs-lookup"><span data-stu-id="d06cf-551">1.0</span></span>|
|[<span data-ttu-id="d06cf-552">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d06cf-552">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d06cf-553">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="d06cf-553">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="d06cf-554">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d06cf-554">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d06cf-555">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d06cf-555">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d06cf-556">Пример</span><span class="sxs-lookup"><span data-stu-id="d06cf-556">Example</span></span>

<span data-ttu-id="d06cf-557">В приведенном ниже примере вызывается `makeEwsRequestAsync` для получения темы элемента с помощью операции `GetItem`.</span><span class="sxs-lookup"><span data-stu-id="d06cf-557">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="d06cf-558">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="d06cf-558">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="d06cf-559">Удаляет обработчиков для поддерживаемого типа события.</span><span class="sxs-lookup"><span data-stu-id="d06cf-559">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="d06cf-560">В настоящее время единственный поддерживаемый тип события — `Office.EventType.ItemChanged`.</span><span class="sxs-lookup"><span data-stu-id="d06cf-560">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="d06cf-561">Параметры</span><span class="sxs-lookup"><span data-stu-id="d06cf-561">Parameters</span></span>

| <span data-ttu-id="d06cf-562">Имя</span><span class="sxs-lookup"><span data-stu-id="d06cf-562">Name</span></span> | <span data-ttu-id="d06cf-563">Тип</span><span class="sxs-lookup"><span data-stu-id="d06cf-563">Type</span></span> | <span data-ttu-id="d06cf-564">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="d06cf-564">Attributes</span></span> | <span data-ttu-id="d06cf-565">Описание</span><span class="sxs-lookup"><span data-stu-id="d06cf-565">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="d06cf-566">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="d06cf-566">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="d06cf-567">Событие, которое должно отменить обработчик.</span><span class="sxs-lookup"><span data-stu-id="d06cf-567">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="d06cf-568">Объект</span><span class="sxs-lookup"><span data-stu-id="d06cf-568">Object</span></span> | <span data-ttu-id="d06cf-569">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d06cf-569">&lt;optional&gt;</span></span> | <span data-ttu-id="d06cf-570">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="d06cf-570">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="d06cf-571">Object</span><span class="sxs-lookup"><span data-stu-id="d06cf-571">Object</span></span> | <span data-ttu-id="d06cf-572">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d06cf-572">&lt;optional&gt;</span></span> | <span data-ttu-id="d06cf-573">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="d06cf-573">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="d06cf-574">функция</span><span class="sxs-lookup"><span data-stu-id="d06cf-574">function</span></span>| <span data-ttu-id="d06cf-575">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="d06cf-575">&lt;optional&gt;</span></span>|<span data-ttu-id="d06cf-576">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="d06cf-576">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d06cf-577">Требования</span><span class="sxs-lookup"><span data-stu-id="d06cf-577">Requirements</span></span>

|<span data-ttu-id="d06cf-578">Требование</span><span class="sxs-lookup"><span data-stu-id="d06cf-578">Requirement</span></span>| <span data-ttu-id="d06cf-579">Значение</span><span class="sxs-lookup"><span data-stu-id="d06cf-579">Value</span></span>|
|---|---|
|[<span data-ttu-id="d06cf-580">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="d06cf-580">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d06cf-581">1.5</span><span class="sxs-lookup"><span data-stu-id="d06cf-581">1.5</span></span> |
|[<span data-ttu-id="d06cf-582">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d06cf-582">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d06cf-583">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d06cf-583">ReadItem</span></span> |
|[<span data-ttu-id="d06cf-584">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d06cf-584">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d06cf-585">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d06cf-585">Compose or Read</span></span>|
