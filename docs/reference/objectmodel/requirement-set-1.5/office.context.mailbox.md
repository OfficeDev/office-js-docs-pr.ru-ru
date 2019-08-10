---
title: Office.context.mailbox — набор обязательных элементов 1.5
description: ''
ms.date: 08/06/2019
localization_priority: Priority
ms.openlocfilehash: 68912520250789b2259d59fb14387f97b24c5c7d
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268644"
---
# <a name="mailbox"></a><span data-ttu-id="14fdd-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="14fdd-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="14fdd-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="14fdd-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="14fdd-104">Предоставляет для Microsoft Outlook доступ к объектной модели надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="14fdd-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="14fdd-105">Требования</span><span class="sxs-lookup"><span data-stu-id="14fdd-105">Requirements</span></span>

|<span data-ttu-id="14fdd-106">Требование</span><span class="sxs-lookup"><span data-stu-id="14fdd-106">Requirement</span></span>| <span data-ttu-id="14fdd-107">Значение</span><span class="sxs-lookup"><span data-stu-id="14fdd-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="14fdd-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="14fdd-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14fdd-109">1.0</span><span class="sxs-lookup"><span data-stu-id="14fdd-109">1.0</span></span>|
|[<span data-ttu-id="14fdd-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14fdd-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14fdd-111">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="14fdd-111">Restricted</span></span>|
|[<span data-ttu-id="14fdd-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14fdd-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14fdd-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="14fdd-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="14fdd-114">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="14fdd-114">Members and methods</span></span>

| <span data-ttu-id="14fdd-115">Элемент</span><span class="sxs-lookup"><span data-stu-id="14fdd-115">Member</span></span> | <span data-ttu-id="14fdd-116">Тип</span><span class="sxs-lookup"><span data-stu-id="14fdd-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="14fdd-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="14fdd-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="14fdd-118">Элемент</span><span class="sxs-lookup"><span data-stu-id="14fdd-118">Member</span></span> |
| [<span data-ttu-id="14fdd-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="14fdd-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="14fdd-120">Элемент</span><span class="sxs-lookup"><span data-stu-id="14fdd-120">Member</span></span> |
| [<span data-ttu-id="14fdd-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="14fdd-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="14fdd-122">Метод</span><span class="sxs-lookup"><span data-stu-id="14fdd-122">Method</span></span> |
| [<span data-ttu-id="14fdd-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="14fdd-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="14fdd-124">Метод</span><span class="sxs-lookup"><span data-stu-id="14fdd-124">Method</span></span> |
| [<span data-ttu-id="14fdd-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="14fdd-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="14fdd-126">Метод</span><span class="sxs-lookup"><span data-stu-id="14fdd-126">Method</span></span> |
| [<span data-ttu-id="14fdd-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="14fdd-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="14fdd-128">Метод</span><span class="sxs-lookup"><span data-stu-id="14fdd-128">Method</span></span> |
| [<span data-ttu-id="14fdd-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="14fdd-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="14fdd-130">Метод</span><span class="sxs-lookup"><span data-stu-id="14fdd-130">Method</span></span> |
| [<span data-ttu-id="14fdd-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="14fdd-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="14fdd-132">Метод</span><span class="sxs-lookup"><span data-stu-id="14fdd-132">Method</span></span> |
| [<span data-ttu-id="14fdd-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="14fdd-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="14fdd-134">Метод</span><span class="sxs-lookup"><span data-stu-id="14fdd-134">Method</span></span> |
| [<span data-ttu-id="14fdd-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="14fdd-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="14fdd-136">Метод</span><span class="sxs-lookup"><span data-stu-id="14fdd-136">Method</span></span> |
| [<span data-ttu-id="14fdd-137">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="14fdd-137">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="14fdd-138">Метод</span><span class="sxs-lookup"><span data-stu-id="14fdd-138">Method</span></span> |
| [<span data-ttu-id="14fdd-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="14fdd-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="14fdd-140">Метод</span><span class="sxs-lookup"><span data-stu-id="14fdd-140">Method</span></span> |
| [<span data-ttu-id="14fdd-141">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="14fdd-141">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="14fdd-142">Метод</span><span class="sxs-lookup"><span data-stu-id="14fdd-142">Method</span></span> |
| [<span data-ttu-id="14fdd-143">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="14fdd-143">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="14fdd-144">Метод</span><span class="sxs-lookup"><span data-stu-id="14fdd-144">Method</span></span> |
| [<span data-ttu-id="14fdd-145">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="14fdd-145">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="14fdd-146">Метод</span><span class="sxs-lookup"><span data-stu-id="14fdd-146">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="14fdd-147">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="14fdd-147">Namespaces</span></span>

<span data-ttu-id="14fdd-148">[diagnostics](Office.context.mailbox.diagnostics.md). Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="14fdd-148">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="14fdd-149">[item](Office.context.mailbox.item.md). Предоставляет методы и свойства для доступа к сообщению или встрече в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="14fdd-149">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="14fdd-150">[userProfile](Office.context.mailbox.userProfile.md). Предоставляет сведения о пользователе в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="14fdd-150">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="14fdd-151">Members</span><span class="sxs-lookup"><span data-stu-id="14fdd-151">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="14fdd-152">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="14fdd-152">ewsUrl :String</span></span>

<span data-ttu-id="14fdd-p101">Получает URL-адрес конечной точки веб-служб Exchange (EWS) для этой учетной записи электронной почты. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="14fdd-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="14fdd-155">Этот элемент не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="14fdd-155">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="14fdd-p102">Удаленная служба может использовать значение `ewsUrl`, чтобы выполнять вызовы EWS для почтового ящика пользователя. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="14fdd-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="14fdd-158">Чтобы вызвать элемент `ewsUrl` в режиме чтения, в манифесте приложения должно быть указано разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="14fdd-158">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="14fdd-p103">Перед использованием элемента `ewsUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="14fdd-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="14fdd-161">Тип</span><span class="sxs-lookup"><span data-stu-id="14fdd-161">Type</span></span>

*   <span data-ttu-id="14fdd-162">String</span><span class="sxs-lookup"><span data-stu-id="14fdd-162">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="14fdd-163">Требования</span><span class="sxs-lookup"><span data-stu-id="14fdd-163">Requirements</span></span>

|<span data-ttu-id="14fdd-164">Требование</span><span class="sxs-lookup"><span data-stu-id="14fdd-164">Requirement</span></span>| <span data-ttu-id="14fdd-165">Значение</span><span class="sxs-lookup"><span data-stu-id="14fdd-165">Value</span></span>|
|---|---|
|[<span data-ttu-id="14fdd-166">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="14fdd-166">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14fdd-167">1.0</span><span class="sxs-lookup"><span data-stu-id="14fdd-167">1.0</span></span>|
|[<span data-ttu-id="14fdd-168">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14fdd-168">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14fdd-169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14fdd-169">ReadItem</span></span>|
|[<span data-ttu-id="14fdd-170">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14fdd-170">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14fdd-171">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="14fdd-171">Compose or Read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="14fdd-172">restUrl: String</span><span class="sxs-lookup"><span data-stu-id="14fdd-172">restUrl :String</span></span>

<span data-ttu-id="14fdd-173">Возвращает URL-адрес конечной точки REST для этой учетной записи электронной почты.</span><span class="sxs-lookup"><span data-stu-id="14fdd-173">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="14fdd-174">С помощью значения `restUrl` можно выполнять вызовы [REST API](/outlook/rest/) для почтового ящика пользователя.</span><span class="sxs-lookup"><span data-stu-id="14fdd-174">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="14fdd-175">Чтобы вызвать элемент `restUrl` в режиме чтения, в манифесте приложения необходимо указать разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="14fdd-175">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="14fdd-p104">Перед использованием элемента `restUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="14fdd-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

> [!NOTE]
> <span data-ttu-id="14fdd-178">Клиенты Outlook, подключенные к локальным установленным версиям Exchange 2016 или более поздним с пользовательским URL-адресом REST, возвращают недопустимое значение `restUrl`.</span><span class="sxs-lookup"><span data-stu-id="14fdd-178">Outlook clients connected to on-premises installations of Exchange 2016 or later with a custom REST URL configured will return an invalid value for `restUrl`.</span></span>

##### <a name="type"></a><span data-ttu-id="14fdd-179">Тип</span><span class="sxs-lookup"><span data-stu-id="14fdd-179">Type</span></span>

*   <span data-ttu-id="14fdd-180">String</span><span class="sxs-lookup"><span data-stu-id="14fdd-180">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="14fdd-181">Требования</span><span class="sxs-lookup"><span data-stu-id="14fdd-181">Requirements</span></span>

|<span data-ttu-id="14fdd-182">Требование</span><span class="sxs-lookup"><span data-stu-id="14fdd-182">Requirement</span></span>| <span data-ttu-id="14fdd-183">Значение</span><span class="sxs-lookup"><span data-stu-id="14fdd-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="14fdd-184">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="14fdd-184">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14fdd-185">1.5</span><span class="sxs-lookup"><span data-stu-id="14fdd-185">1.5</span></span> |
|[<span data-ttu-id="14fdd-186">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14fdd-186">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14fdd-187">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14fdd-187">ReadItem</span></span>|
|[<span data-ttu-id="14fdd-188">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14fdd-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14fdd-189">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="14fdd-189">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="14fdd-190">Методы</span><span class="sxs-lookup"><span data-stu-id="14fdd-190">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="14fdd-191">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="14fdd-191">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="14fdd-192">Добавляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="14fdd-192">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="14fdd-193">В настоящее время поддерживаются только события типа `Office.EventType.ItemChanged`, которые вызываются, когда пользователь выбирает новый элемент.</span><span class="sxs-lookup"><span data-stu-id="14fdd-193">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item.</span></span> <span data-ttu-id="14fdd-194">Это событие используется надстройками, реализующими закрепляемую область задач, и позволяет надстройке обновлять пользовательский интерфейс области задач в соответствии с выбранным в данный момент элементом.</span><span class="sxs-lookup"><span data-stu-id="14fdd-194">This event is used by add-ins that implement a pinnable task pane, and allows the add-in to refresh the task pane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="14fdd-195">Параметры</span><span class="sxs-lookup"><span data-stu-id="14fdd-195">Parameters</span></span>

| <span data-ttu-id="14fdd-196">Имя</span><span class="sxs-lookup"><span data-stu-id="14fdd-196">Name</span></span> | <span data-ttu-id="14fdd-197">Тип</span><span class="sxs-lookup"><span data-stu-id="14fdd-197">Type</span></span> | <span data-ttu-id="14fdd-198">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="14fdd-198">Attributes</span></span> | <span data-ttu-id="14fdd-199">Описание</span><span class="sxs-lookup"><span data-stu-id="14fdd-199">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="14fdd-200">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="14fdd-200">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="14fdd-201">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="14fdd-201">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="14fdd-202">Function</span><span class="sxs-lookup"><span data-stu-id="14fdd-202">Function</span></span> || <span data-ttu-id="14fdd-p106">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="14fdd-p106">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="14fdd-206">Объект</span><span class="sxs-lookup"><span data-stu-id="14fdd-206">Object</span></span> | <span data-ttu-id="14fdd-207">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="14fdd-207">&lt;optional&gt;</span></span> | <span data-ttu-id="14fdd-208">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="14fdd-208">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="14fdd-209">Object</span><span class="sxs-lookup"><span data-stu-id="14fdd-209">Object</span></span> | <span data-ttu-id="14fdd-210">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="14fdd-210">&lt;optional&gt;</span></span> | <span data-ttu-id="14fdd-211">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="14fdd-211">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="14fdd-212">функция</span><span class="sxs-lookup"><span data-stu-id="14fdd-212">function</span></span>| <span data-ttu-id="14fdd-213">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="14fdd-213">&lt;optional&gt;</span></span>|<span data-ttu-id="14fdd-214">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="14fdd-214">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="14fdd-215">Требования</span><span class="sxs-lookup"><span data-stu-id="14fdd-215">Requirements</span></span>

|<span data-ttu-id="14fdd-216">Требование</span><span class="sxs-lookup"><span data-stu-id="14fdd-216">Requirement</span></span>| <span data-ttu-id="14fdd-217">Значение</span><span class="sxs-lookup"><span data-stu-id="14fdd-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="14fdd-218">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="14fdd-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14fdd-219">1.5</span><span class="sxs-lookup"><span data-stu-id="14fdd-219">1.5</span></span> |
|[<span data-ttu-id="14fdd-220">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14fdd-220">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14fdd-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14fdd-221">ReadItem</span></span> |
|[<span data-ttu-id="14fdd-222">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14fdd-222">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14fdd-223">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="14fdd-223">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="14fdd-224">Пример</span><span class="sxs-lookup"><span data-stu-id="14fdd-224">Example</span></span>

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

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="14fdd-225">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="14fdd-225">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="14fdd-226">Преобразовывает идентификатор элемента из формата REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="14fdd-226">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="14fdd-227">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="14fdd-227">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="14fdd-p107">Формат идентификаторов, извлекаемых через API REST (например, [API Почты Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)), отличается от формата веб-служб Exchange (EWS). Метод `convertToEwsId` преобразовывает идентификатор в формате REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="14fdd-p107">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="14fdd-230">Параметры</span><span class="sxs-lookup"><span data-stu-id="14fdd-230">Parameters</span></span>

|<span data-ttu-id="14fdd-231">Имя</span><span class="sxs-lookup"><span data-stu-id="14fdd-231">Name</span></span>| <span data-ttu-id="14fdd-232">Тип</span><span class="sxs-lookup"><span data-stu-id="14fdd-232">Type</span></span>| <span data-ttu-id="14fdd-233">Описание</span><span class="sxs-lookup"><span data-stu-id="14fdd-233">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="14fdd-234">String</span><span class="sxs-lookup"><span data-stu-id="14fdd-234">String</span></span>|<span data-ttu-id="14fdd-235">Идентификатор элемента в формате REST API для Outlook</span><span class="sxs-lookup"><span data-stu-id="14fdd-235">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="14fdd-236">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="14fdd-236">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.5)|<span data-ttu-id="14fdd-237">Значение, определяющее версию REST API для Outlook, которая используется для извлечения идентификатора элемента.</span><span class="sxs-lookup"><span data-stu-id="14fdd-237">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="14fdd-238">Требования</span><span class="sxs-lookup"><span data-stu-id="14fdd-238">Requirements</span></span>

|<span data-ttu-id="14fdd-239">Требование</span><span class="sxs-lookup"><span data-stu-id="14fdd-239">Requirement</span></span>| <span data-ttu-id="14fdd-240">Значение</span><span class="sxs-lookup"><span data-stu-id="14fdd-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="14fdd-241">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="14fdd-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14fdd-242">1.3</span><span class="sxs-lookup"><span data-stu-id="14fdd-242">1.3</span></span>|
|[<span data-ttu-id="14fdd-243">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14fdd-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14fdd-244">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="14fdd-244">Restricted</span></span>|
|[<span data-ttu-id="14fdd-245">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14fdd-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14fdd-246">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="14fdd-246">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="14fdd-247">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="14fdd-247">Returns:</span></span>

<span data-ttu-id="14fdd-248">Тип: String</span><span class="sxs-lookup"><span data-stu-id="14fdd-248">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="14fdd-249">Пример</span><span class="sxs-lookup"><span data-stu-id="14fdd-249">Example</span></span>

```javascript
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-15"></a><span data-ttu-id="14fdd-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)}</span><span class="sxs-lookup"><span data-stu-id="14fdd-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)}</span></span>

<span data-ttu-id="14fdd-251">Получает словарь, содержащий сведения о локальном времени клиента.</span><span class="sxs-lookup"><span data-stu-id="14fdd-251">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="14fdd-p108">Почтовое приложение для классической версии Outlook или версии в Интернете может использовать разные часовые пояса для дат и времени. Классическое приложение Outlook использует часовой пояс клиентского компьютера. Outlook в Интернете использует часовой пояс, заданный в Центре администрирования Exchange (EAC). Значения даты и времени должны обрабатываться так, чтобы значения в пользовательском интерфейсе всегда согласовывались с часовым поясом, ожидаемым пользователем.</span><span class="sxs-lookup"><span data-stu-id="14fdd-p108">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="14fdd-p109">Если почтовое приложение работает в классическом клиенте Outlook, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса клиентского компьютера. Если почтовое приложение работает в Outlook в Интернете, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса, заданного в Центре администрирования Exchange.</span><span class="sxs-lookup"><span data-stu-id="14fdd-p109">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="14fdd-257">Параметры</span><span class="sxs-lookup"><span data-stu-id="14fdd-257">Parameters</span></span>

|<span data-ttu-id="14fdd-258">Имя</span><span class="sxs-lookup"><span data-stu-id="14fdd-258">Name</span></span>| <span data-ttu-id="14fdd-259">Тип</span><span class="sxs-lookup"><span data-stu-id="14fdd-259">Type</span></span>| <span data-ttu-id="14fdd-260">Описание</span><span class="sxs-lookup"><span data-stu-id="14fdd-260">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="14fdd-261">Date</span><span class="sxs-lookup"><span data-stu-id="14fdd-261">Date</span></span>|<span data-ttu-id="14fdd-262">Объект Date</span><span class="sxs-lookup"><span data-stu-id="14fdd-262">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="14fdd-263">Требования</span><span class="sxs-lookup"><span data-stu-id="14fdd-263">Requirements</span></span>

|<span data-ttu-id="14fdd-264">Требование</span><span class="sxs-lookup"><span data-stu-id="14fdd-264">Requirement</span></span>| <span data-ttu-id="14fdd-265">Значение</span><span class="sxs-lookup"><span data-stu-id="14fdd-265">Value</span></span>|
|---|---|
|[<span data-ttu-id="14fdd-266">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="14fdd-266">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14fdd-267">1.0</span><span class="sxs-lookup"><span data-stu-id="14fdd-267">1.0</span></span>|
|[<span data-ttu-id="14fdd-268">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14fdd-268">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14fdd-269">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14fdd-269">ReadItem</span></span>|
|[<span data-ttu-id="14fdd-270">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14fdd-270">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14fdd-271">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="14fdd-271">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="14fdd-272">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="14fdd-272">Returns:</span></span>

<span data-ttu-id="14fdd-273">Тип: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="14fdd-273">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)</span></span>

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="14fdd-274">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="14fdd-274">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="14fdd-275">Преобразовывает идентификатор элемента в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="14fdd-275">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="14fdd-276">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="14fdd-276">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="14fdd-p110">Формат идентификаторов, извлекаемых через EWS или свойство `itemId`, отличается от формата API REST (таких как [API Почты Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)). Метод `convertToRestId` преобразовывает идентификатор в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="14fdd-p110">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="14fdd-279">Параметры</span><span class="sxs-lookup"><span data-stu-id="14fdd-279">Parameters</span></span>

|<span data-ttu-id="14fdd-280">Имя</span><span class="sxs-lookup"><span data-stu-id="14fdd-280">Name</span></span>| <span data-ttu-id="14fdd-281">Тип</span><span class="sxs-lookup"><span data-stu-id="14fdd-281">Type</span></span>| <span data-ttu-id="14fdd-282">Описание</span><span class="sxs-lookup"><span data-stu-id="14fdd-282">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="14fdd-283">String</span><span class="sxs-lookup"><span data-stu-id="14fdd-283">String</span></span>|<span data-ttu-id="14fdd-284">Идентификатор элемента в формате EWS</span><span class="sxs-lookup"><span data-stu-id="14fdd-284">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="14fdd-285">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="14fdd-285">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.5)|<span data-ttu-id="14fdd-286">Значение, определяющее версию REST API для Outlook, с которой будет использоваться преобразованный идентификатор.</span><span class="sxs-lookup"><span data-stu-id="14fdd-286">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="14fdd-287">Требования</span><span class="sxs-lookup"><span data-stu-id="14fdd-287">Requirements</span></span>

|<span data-ttu-id="14fdd-288">Требование</span><span class="sxs-lookup"><span data-stu-id="14fdd-288">Requirement</span></span>| <span data-ttu-id="14fdd-289">Значение</span><span class="sxs-lookup"><span data-stu-id="14fdd-289">Value</span></span>|
|---|---|
|[<span data-ttu-id="14fdd-290">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="14fdd-290">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14fdd-291">1.3</span><span class="sxs-lookup"><span data-stu-id="14fdd-291">1.3</span></span>|
|[<span data-ttu-id="14fdd-292">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14fdd-292">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14fdd-293">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="14fdd-293">Restricted</span></span>|
|[<span data-ttu-id="14fdd-294">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14fdd-294">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14fdd-295">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="14fdd-295">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="14fdd-296">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="14fdd-296">Returns:</span></span>

<span data-ttu-id="14fdd-297">Тип: String</span><span class="sxs-lookup"><span data-stu-id="14fdd-297">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="14fdd-298">Пример</span><span class="sxs-lookup"><span data-stu-id="14fdd-298">Example</span></span>

```javascript
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="14fdd-299">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="14fdd-299">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="14fdd-300">Получает объект Date из словаря, содержащего сведения о времени.</span><span class="sxs-lookup"><span data-stu-id="14fdd-300">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="14fdd-301">Метод `convertToUtcClientTime` преобразует словарь, содержащий локальную дату и время, в объект Date с правильными значениями локальной даты и времени.</span><span class="sxs-lookup"><span data-stu-id="14fdd-301">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="14fdd-302">Параметры</span><span class="sxs-lookup"><span data-stu-id="14fdd-302">Parameters</span></span>

|<span data-ttu-id="14fdd-303">Имя</span><span class="sxs-lookup"><span data-stu-id="14fdd-303">Name</span></span>| <span data-ttu-id="14fdd-304">Тип</span><span class="sxs-lookup"><span data-stu-id="14fdd-304">Type</span></span>| <span data-ttu-id="14fdd-305">Описание</span><span class="sxs-lookup"><span data-stu-id="14fdd-305">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="14fdd-306">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="14fdd-306">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)|<span data-ttu-id="14fdd-307">Значение локального времени для преобразования.</span><span class="sxs-lookup"><span data-stu-id="14fdd-307">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="14fdd-308">Требования</span><span class="sxs-lookup"><span data-stu-id="14fdd-308">Requirements</span></span>

|<span data-ttu-id="14fdd-309">Требование</span><span class="sxs-lookup"><span data-stu-id="14fdd-309">Requirement</span></span>| <span data-ttu-id="14fdd-310">Значение</span><span class="sxs-lookup"><span data-stu-id="14fdd-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="14fdd-311">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="14fdd-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14fdd-312">1.0</span><span class="sxs-lookup"><span data-stu-id="14fdd-312">1.0</span></span>|
|[<span data-ttu-id="14fdd-313">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14fdd-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14fdd-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14fdd-314">ReadItem</span></span>|
|[<span data-ttu-id="14fdd-315">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14fdd-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14fdd-316">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="14fdd-316">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="14fdd-317">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="14fdd-317">Returns:</span></span>

<span data-ttu-id="14fdd-318">Объект Date со временем в формате UTC.</span><span class="sxs-lookup"><span data-stu-id="14fdd-318">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="14fdd-319">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="14fdd-319">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="14fdd-320">Date</span><span class="sxs-lookup"><span data-stu-id="14fdd-320">Date</span></span></dd>

</dl>

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="14fdd-321">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="14fdd-321">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="14fdd-322">Отображает имеющуюся встречу из календаря.</span><span class="sxs-lookup"><span data-stu-id="14fdd-322">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="14fdd-323">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="14fdd-323">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="14fdd-324">Метод `displayAppointmentForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее сведения календаря о существующей встрече.</span><span class="sxs-lookup"><span data-stu-id="14fdd-324">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="14fdd-p111">В Outlook для Mac с помощью этого метода можно отобразить одну встречу, которая не является частью повторяющегося ряда, или основную встречу такого ряда, но не экземпляр из него, так как в Outlook для Mac невозможно получить доступ к свойствам экземпляра повторяющегося ряда (в том числе к идентификатору элемента).</span><span class="sxs-lookup"><span data-stu-id="14fdd-p111">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="14fdd-327">В Outlook в Интернете этот метод открывает указанную форму, только если текст формы содержит символы размером не более 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="14fdd-327">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="14fdd-328">Если указанный идентификатор элемента не определяет существующую встречу, на клиентском компьютере или устройстве открывается пустая страница, и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="14fdd-328">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="14fdd-329">Параметры</span><span class="sxs-lookup"><span data-stu-id="14fdd-329">Parameters</span></span>

|<span data-ttu-id="14fdd-330">Имя</span><span class="sxs-lookup"><span data-stu-id="14fdd-330">Name</span></span>| <span data-ttu-id="14fdd-331">Тип</span><span class="sxs-lookup"><span data-stu-id="14fdd-331">Type</span></span>| <span data-ttu-id="14fdd-332">Описание</span><span class="sxs-lookup"><span data-stu-id="14fdd-332">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="14fdd-333">String</span><span class="sxs-lookup"><span data-stu-id="14fdd-333">String</span></span>|<span data-ttu-id="14fdd-334">Идентификатор веб-служб Exchange для существующей встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="14fdd-334">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="14fdd-335">Требования</span><span class="sxs-lookup"><span data-stu-id="14fdd-335">Requirements</span></span>

|<span data-ttu-id="14fdd-336">Требование</span><span class="sxs-lookup"><span data-stu-id="14fdd-336">Requirement</span></span>| <span data-ttu-id="14fdd-337">Значение</span><span class="sxs-lookup"><span data-stu-id="14fdd-337">Value</span></span>|
|---|---|
|[<span data-ttu-id="14fdd-338">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="14fdd-338">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14fdd-339">1.0</span><span class="sxs-lookup"><span data-stu-id="14fdd-339">1.0</span></span>|
|[<span data-ttu-id="14fdd-340">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14fdd-340">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14fdd-341">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14fdd-341">ReadItem</span></span>|
|[<span data-ttu-id="14fdd-342">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14fdd-342">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14fdd-343">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="14fdd-343">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="14fdd-344">Пример</span><span class="sxs-lookup"><span data-stu-id="14fdd-344">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

#### <a name="displaymessageformitemid"></a><span data-ttu-id="14fdd-345">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="14fdd-345">displayMessageForm(itemId)</span></span>

<span data-ttu-id="14fdd-346">Отображает имеющееся сообщение.</span><span class="sxs-lookup"><span data-stu-id="14fdd-346">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="14fdd-347">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="14fdd-347">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="14fdd-348">Метод `displayMessageForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее существующее сообщение.</span><span class="sxs-lookup"><span data-stu-id="14fdd-348">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="14fdd-349">В Outlook в Интернете этот метод открывает указанную форму, только если текст формы содержит символы размером не более 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="14fdd-349">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="14fdd-350">Если указанный идентификатор элемента не определяет существующее сообщение, окно на клиентском компьютере не открывается и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="14fdd-350">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="14fdd-p112">Не используйте `displayMessageForm` с параметром `itemId`, который представляет собой встречу. Используйте метод `displayAppointmentForm`, чтобы отобразить сведения о существующей встрече, а метод `displayNewAppointmentForm` — для отображения формы создания встречи.</span><span class="sxs-lookup"><span data-stu-id="14fdd-p112">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="14fdd-353">Параметры</span><span class="sxs-lookup"><span data-stu-id="14fdd-353">Parameters</span></span>

|<span data-ttu-id="14fdd-354">Имя</span><span class="sxs-lookup"><span data-stu-id="14fdd-354">Name</span></span>| <span data-ttu-id="14fdd-355">Тип</span><span class="sxs-lookup"><span data-stu-id="14fdd-355">Type</span></span>| <span data-ttu-id="14fdd-356">Описание</span><span class="sxs-lookup"><span data-stu-id="14fdd-356">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="14fdd-357">String</span><span class="sxs-lookup"><span data-stu-id="14fdd-357">String</span></span>|<span data-ttu-id="14fdd-358">Идентификатор веб-служб Exchange для существующего сообщения.</span><span class="sxs-lookup"><span data-stu-id="14fdd-358">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="14fdd-359">Требования</span><span class="sxs-lookup"><span data-stu-id="14fdd-359">Requirements</span></span>

|<span data-ttu-id="14fdd-360">Требование</span><span class="sxs-lookup"><span data-stu-id="14fdd-360">Requirement</span></span>| <span data-ttu-id="14fdd-361">Значение</span><span class="sxs-lookup"><span data-stu-id="14fdd-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="14fdd-362">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="14fdd-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14fdd-363">1.0</span><span class="sxs-lookup"><span data-stu-id="14fdd-363">1.0</span></span>|
|[<span data-ttu-id="14fdd-364">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14fdd-364">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14fdd-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14fdd-365">ReadItem</span></span>|
|[<span data-ttu-id="14fdd-366">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14fdd-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14fdd-367">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="14fdd-367">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="14fdd-368">Пример</span><span class="sxs-lookup"><span data-stu-id="14fdd-368">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="14fdd-369">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="14fdd-369">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="14fdd-370">Отображает форму для создания новой встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="14fdd-370">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="14fdd-371">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="14fdd-371">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="14fdd-p113">Метод `displayNewAppointmentForm` открывает форму, в которой пользователь может создать встречу или собрание. Если параметры заданы, поля формы встречи автоматически заполняются их содержимым.</span><span class="sxs-lookup"><span data-stu-id="14fdd-p113">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="14fdd-p114">В Outlook в Интернете и на мобильных устройствах этот метод всегда отображает форму с полем участников. Если вы не укажете участников в качестве входных аргументов, метод отображает форму с кнопкой **Сохранить**. Если вы укажете участников, форма будет включать участников и кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="14fdd-p114">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="14fdd-p115">Если вы укажете участников или ресурсы с помощью параметра `requiredAttendees`, `optionalAttendees` или `resources` в клиенте Outlook с расширенными возможностями и Outlook RT, этот метод отобразит форму собрания с кнопкой **Отправить**. Если не указать получателей, этот метод отобразит форму встречи с кнопкой **Сохранить и закрыть**.</span><span class="sxs-lookup"><span data-stu-id="14fdd-p115">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="14fdd-379">Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.</span><span class="sxs-lookup"><span data-stu-id="14fdd-379">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="14fdd-380">Параметры</span><span class="sxs-lookup"><span data-stu-id="14fdd-380">Parameters</span></span>

|<span data-ttu-id="14fdd-381">Имя</span><span class="sxs-lookup"><span data-stu-id="14fdd-381">Name</span></span>| <span data-ttu-id="14fdd-382">Тип</span><span class="sxs-lookup"><span data-stu-id="14fdd-382">Type</span></span>| <span data-ttu-id="14fdd-383">Описание</span><span class="sxs-lookup"><span data-stu-id="14fdd-383">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="14fdd-384">Object</span><span class="sxs-lookup"><span data-stu-id="14fdd-384">Object</span></span> | <span data-ttu-id="14fdd-385">Словарь параметров, описывающий новую встречу.</span><span class="sxs-lookup"><span data-stu-id="14fdd-385">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="14fdd-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span><span class="sxs-lookup"><span data-stu-id="14fdd-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span></span> | <span data-ttu-id="14fdd-p116">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из обязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="14fdd-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="14fdd-389">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span><span class="sxs-lookup"><span data-stu-id="14fdd-389">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span></span> | <span data-ttu-id="14fdd-p117">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из необязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="14fdd-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="14fdd-392">Date</span><span class="sxs-lookup"><span data-stu-id="14fdd-392">Date</span></span> | <span data-ttu-id="14fdd-393">Объект `Date`, указывающий дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="14fdd-393">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="14fdd-394">Date</span><span class="sxs-lookup"><span data-stu-id="14fdd-394">Date</span></span> | <span data-ttu-id="14fdd-395">Объект `Date`, указывающий дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="14fdd-395">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="14fdd-396">String</span><span class="sxs-lookup"><span data-stu-id="14fdd-396">String</span></span> | <span data-ttu-id="14fdd-p118">Строка со сведениями о месте встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="14fdd-p118">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="14fdd-399">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="14fdd-399">Array.&lt;String&gt;</span></span> | <span data-ttu-id="14fdd-p119">Массив строк, содержащий необходимые для встречи ресурсы. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="14fdd-p119">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="14fdd-402">String</span><span class="sxs-lookup"><span data-stu-id="14fdd-402">String</span></span> | <span data-ttu-id="14fdd-p120">Строка с темой встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="14fdd-p120">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="14fdd-405">String</span><span class="sxs-lookup"><span data-stu-id="14fdd-405">String</span></span> | <span data-ttu-id="14fdd-p121">Текст сообщения о встрече. Максимальный размер содержимого сообщения — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="14fdd-p121">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="14fdd-408">Требования</span><span class="sxs-lookup"><span data-stu-id="14fdd-408">Requirements</span></span>

|<span data-ttu-id="14fdd-409">Требование</span><span class="sxs-lookup"><span data-stu-id="14fdd-409">Requirement</span></span>| <span data-ttu-id="14fdd-410">Значение</span><span class="sxs-lookup"><span data-stu-id="14fdd-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="14fdd-411">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="14fdd-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14fdd-412">1.0</span><span class="sxs-lookup"><span data-stu-id="14fdd-412">1.0</span></span>|
|[<span data-ttu-id="14fdd-413">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14fdd-413">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14fdd-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14fdd-414">ReadItem</span></span>|
|[<span data-ttu-id="14fdd-415">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14fdd-415">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14fdd-416">Чтение</span><span class="sxs-lookup"><span data-stu-id="14fdd-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="14fdd-417">Пример</span><span class="sxs-lookup"><span data-stu-id="14fdd-417">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="14fdd-418">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="14fdd-418">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="14fdd-419">Возвращает строку, содержащую маркер, который используется для вызова интерфейсов REST API или веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="14fdd-419">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="14fdd-p122">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="14fdd-p122">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="14fdd-422">Рекомендуем сделать так, чтобы по мере возможности надстройки использовали интерфейсы REST API, а не веб-службы Exchange.</span><span class="sxs-lookup"><span data-stu-id="14fdd-422">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="14fdd-423">**Маркеры REST**</span><span class="sxs-lookup"><span data-stu-id="14fdd-423">**REST Tokens**</span></span>

<span data-ttu-id="14fdd-p123">Если запрашивается маркер REST (`options.isRest = true`), полученный маркер не подойдет для проверки подлинности при вызовах веб-служб Exchange. Область действия маркера будет ограничена доступом только для чтения к текущему элементу и его вложениям, если в манифесте надстройки не указано разрешение [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission). Если указано разрешение `ReadWriteMailbox`, полученный маркер предоставит доступ на чтение и запись к почте, календарю и контактам, включая возможность отправки почты.</span><span class="sxs-lookup"><span data-stu-id="14fdd-p123">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="14fdd-427">С помощью свойства `restUrl` надстройка должна определить правильный URL-адрес для вызовов REST API.</span><span class="sxs-lookup"><span data-stu-id="14fdd-427">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="14fdd-428">**Маркеры EWS**</span><span class="sxs-lookup"><span data-stu-id="14fdd-428">**EWS Tokens**</span></span>

<span data-ttu-id="14fdd-p124">Если запрашивается маркер EWS (`options.isRest = false`), полученный маркер не подойдет для проверки подлинности при вызовах REST API. Область действия маркера будет ограничена доступом к текущему элементу.</span><span class="sxs-lookup"><span data-stu-id="14fdd-p124">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="14fdd-431">С помощью свойства `ewsUrl` надстройка должна определить правильный URL-адрес для вызовов EWS.</span><span class="sxs-lookup"><span data-stu-id="14fdd-431">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="14fdd-432">Параметры</span><span class="sxs-lookup"><span data-stu-id="14fdd-432">Parameters</span></span>

|<span data-ttu-id="14fdd-433">Имя</span><span class="sxs-lookup"><span data-stu-id="14fdd-433">Name</span></span>| <span data-ttu-id="14fdd-434">Тип</span><span class="sxs-lookup"><span data-stu-id="14fdd-434">Type</span></span>| <span data-ttu-id="14fdd-435">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="14fdd-435">Attributes</span></span>| <span data-ttu-id="14fdd-436">Описание</span><span class="sxs-lookup"><span data-stu-id="14fdd-436">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="14fdd-437">Object</span><span class="sxs-lookup"><span data-stu-id="14fdd-437">Object</span></span> | <span data-ttu-id="14fdd-438">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="14fdd-438">&lt;optional&gt;</span></span> | <span data-ttu-id="14fdd-439">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="14fdd-439">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="14fdd-440">Boolean</span><span class="sxs-lookup"><span data-stu-id="14fdd-440">Boolean</span></span> |  <span data-ttu-id="14fdd-441">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="14fdd-441">&lt;optional&gt;</span></span> | <span data-ttu-id="14fdd-p125">Определяет, будет ли предоставленный маркер использоваться для интерфейсов REST API Outlook или веб-служб Exchange. Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="14fdd-p125">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="14fdd-444">Объект</span><span class="sxs-lookup"><span data-stu-id="14fdd-444">Object</span></span> |  <span data-ttu-id="14fdd-445">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="14fdd-445">&lt;optional&gt;</span></span> | <span data-ttu-id="14fdd-446">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="14fdd-446">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="14fdd-447">функция</span><span class="sxs-lookup"><span data-stu-id="14fdd-447">function</span></span>||<span data-ttu-id="14fdd-448">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="14fdd-448">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="14fdd-449">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="14fdd-449">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="14fdd-450">При наличии ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут предоставлять дополнительные сведения.</span><span class="sxs-lookup"><span data-stu-id="14fdd-450">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="14fdd-451">Ошибки</span><span class="sxs-lookup"><span data-stu-id="14fdd-451">Errors</span></span>

|<span data-ttu-id="14fdd-452">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="14fdd-452">Error code</span></span>|<span data-ttu-id="14fdd-453">Описание</span><span class="sxs-lookup"><span data-stu-id="14fdd-453">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="14fdd-454">Не удалось выполнить запрос.</span><span class="sxs-lookup"><span data-stu-id="14fdd-454">The request has failed.</span></span> <span data-ttu-id="14fdd-455">Просмотрите объект диагностики для кода ошибки HTTP.</span><span class="sxs-lookup"><span data-stu-id="14fdd-455">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="14fdd-456">Сервер Exchange Server вернул ошибку.</span><span class="sxs-lookup"><span data-stu-id="14fdd-456">The RMS server returned an error.</span></span> <span data-ttu-id="14fdd-457">Для получения дополнительных сведений просмотрите объект диагностики.</span><span class="sxs-lookup"><span data-stu-id="14fdd-457">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="14fdd-458">Пользователь отключен от сети.</span><span class="sxs-lookup"><span data-stu-id="14fdd-458">The user is no longer connected to the network.</span></span> <span data-ttu-id="14fdd-459">Проверьте сетевое подключение и повторите попытку.</span><span class="sxs-lookup"><span data-stu-id="14fdd-459">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="14fdd-460">Требования</span><span class="sxs-lookup"><span data-stu-id="14fdd-460">Requirements</span></span>

|<span data-ttu-id="14fdd-461">Требование</span><span class="sxs-lookup"><span data-stu-id="14fdd-461">Requirement</span></span>| <span data-ttu-id="14fdd-462">Значение</span><span class="sxs-lookup"><span data-stu-id="14fdd-462">Value</span></span>|
|---|---|
|[<span data-ttu-id="14fdd-463">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="14fdd-463">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14fdd-464">1.5</span><span class="sxs-lookup"><span data-stu-id="14fdd-464">1.5</span></span> |
|[<span data-ttu-id="14fdd-465">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14fdd-465">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14fdd-466">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14fdd-466">ReadItem</span></span>|
|[<span data-ttu-id="14fdd-467">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14fdd-467">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14fdd-468">Создание и чтение</span><span class="sxs-lookup"><span data-stu-id="14fdd-468">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="14fdd-469">Пример</span><span class="sxs-lookup"><span data-stu-id="14fdd-469">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="14fdd-470">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="14fdd-470">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="14fdd-471">Получает строку, содержащую маркер, используемый для получения вложения или элемента с Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="14fdd-471">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="14fdd-p129">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="14fdd-p129">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="14fdd-p130">Вы можете передать сторонней системе маркер и идентификатор вложения или элемента. Сторонняя система использует этот маркер как маркер авторизации, чтобы вызвать операцию [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) или [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) веб-служб Exchange для возврата вложения или элемента. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="14fdd-p130">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="14fdd-477">Для вызова метода `getCallbackTokenAsync` в режиме чтения манифесте приложения должно быть указано разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="14fdd-477">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="14fdd-p131">Чтобы получить идентификатор элемента для передачи в метод `getCallbackTokenAsync`, в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="14fdd-p131">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="14fdd-480">Параметры</span><span class="sxs-lookup"><span data-stu-id="14fdd-480">Parameters</span></span>

|<span data-ttu-id="14fdd-481">Имя</span><span class="sxs-lookup"><span data-stu-id="14fdd-481">Name</span></span>| <span data-ttu-id="14fdd-482">Тип</span><span class="sxs-lookup"><span data-stu-id="14fdd-482">Type</span></span>| <span data-ttu-id="14fdd-483">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="14fdd-483">Attributes</span></span>| <span data-ttu-id="14fdd-484">Описание</span><span class="sxs-lookup"><span data-stu-id="14fdd-484">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="14fdd-485">function</span><span class="sxs-lookup"><span data-stu-id="14fdd-485">function</span></span>||<span data-ttu-id="14fdd-486">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="14fdd-486">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="14fdd-487">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="14fdd-487">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="14fdd-488">При наличии ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут предоставлять дополнительные сведения.</span><span class="sxs-lookup"><span data-stu-id="14fdd-488">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="14fdd-489">Объект</span><span class="sxs-lookup"><span data-stu-id="14fdd-489">Object</span></span>| <span data-ttu-id="14fdd-490">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="14fdd-490">&lt;optional&gt;</span></span>|<span data-ttu-id="14fdd-491">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="14fdd-491">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="14fdd-492">Ошибки</span><span class="sxs-lookup"><span data-stu-id="14fdd-492">Errors</span></span>

|<span data-ttu-id="14fdd-493">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="14fdd-493">Error code</span></span>|<span data-ttu-id="14fdd-494">Описание</span><span class="sxs-lookup"><span data-stu-id="14fdd-494">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="14fdd-495">Не удалось выполнить запрос.</span><span class="sxs-lookup"><span data-stu-id="14fdd-495">The request has failed.</span></span> <span data-ttu-id="14fdd-496">Просмотрите объект диагностики для кода ошибки HTTP.</span><span class="sxs-lookup"><span data-stu-id="14fdd-496">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="14fdd-497">Сервер Exchange Server вернул ошибку.</span><span class="sxs-lookup"><span data-stu-id="14fdd-497">The RMS server returned an error.</span></span> <span data-ttu-id="14fdd-498">Для получения дополнительных сведений просмотрите объект диагностики.</span><span class="sxs-lookup"><span data-stu-id="14fdd-498">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="14fdd-499">Пользователь отключен от сети.</span><span class="sxs-lookup"><span data-stu-id="14fdd-499">The user is no longer connected to the network.</span></span> <span data-ttu-id="14fdd-500">Проверьте сетевое подключение и повторите попытку.</span><span class="sxs-lookup"><span data-stu-id="14fdd-500">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="14fdd-501">Требования</span><span class="sxs-lookup"><span data-stu-id="14fdd-501">Requirements</span></span>

|<span data-ttu-id="14fdd-502">Требование</span><span class="sxs-lookup"><span data-stu-id="14fdd-502">Requirement</span></span>| <span data-ttu-id="14fdd-503">Значение</span><span class="sxs-lookup"><span data-stu-id="14fdd-503">Value</span></span>|
|---|---|
|[<span data-ttu-id="14fdd-504">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="14fdd-504">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14fdd-505">1.0</span><span class="sxs-lookup"><span data-stu-id="14fdd-505">1.0</span></span>|
|[<span data-ttu-id="14fdd-506">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14fdd-506">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14fdd-507">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14fdd-507">ReadItem</span></span>|
|[<span data-ttu-id="14fdd-508">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14fdd-508">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14fdd-509">Создание и чтение</span><span class="sxs-lookup"><span data-stu-id="14fdd-509">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="14fdd-510">Пример</span><span class="sxs-lookup"><span data-stu-id="14fdd-510">Example</span></span>

```javascript
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="14fdd-511">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="14fdd-511">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="14fdd-512">Получает маркер, идентифицирующий пользователя и надстройку Office.</span><span class="sxs-lookup"><span data-stu-id="14fdd-512">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="14fdd-513">Метод `getUserIdentityTokenAsync` возвращает маркер, который можно использовать для идентификации, а также [проверки подлинности надстройки и пользователя в сторонней системе](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="14fdd-513">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="14fdd-514">Параметры</span><span class="sxs-lookup"><span data-stu-id="14fdd-514">Parameters</span></span>

|<span data-ttu-id="14fdd-515">Имя</span><span class="sxs-lookup"><span data-stu-id="14fdd-515">Name</span></span>| <span data-ttu-id="14fdd-516">Тип</span><span class="sxs-lookup"><span data-stu-id="14fdd-516">Type</span></span>| <span data-ttu-id="14fdd-517">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="14fdd-517">Attributes</span></span>| <span data-ttu-id="14fdd-518">Описание</span><span class="sxs-lookup"><span data-stu-id="14fdd-518">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="14fdd-519">function</span><span class="sxs-lookup"><span data-stu-id="14fdd-519">function</span></span>||<span data-ttu-id="14fdd-520">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="14fdd-520">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="14fdd-521">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="14fdd-521">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="14fdd-522">При наличии ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут предоставлять дополнительные сведения.</span><span class="sxs-lookup"><span data-stu-id="14fdd-522">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="14fdd-523">Объект</span><span class="sxs-lookup"><span data-stu-id="14fdd-523">Object</span></span>| <span data-ttu-id="14fdd-524">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="14fdd-524">&lt;optional&gt;</span></span>|<span data-ttu-id="14fdd-525">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="14fdd-525">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="14fdd-526">Ошибки</span><span class="sxs-lookup"><span data-stu-id="14fdd-526">Errors</span></span>

|<span data-ttu-id="14fdd-527">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="14fdd-527">Error code</span></span>|<span data-ttu-id="14fdd-528">Описание</span><span class="sxs-lookup"><span data-stu-id="14fdd-528">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="14fdd-529">Не удалось выполнить запрос.</span><span class="sxs-lookup"><span data-stu-id="14fdd-529">The request has failed.</span></span> <span data-ttu-id="14fdd-530">Просмотрите объект диагностики для кода ошибки HTTP.</span><span class="sxs-lookup"><span data-stu-id="14fdd-530">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="14fdd-531">Сервер Exchange Server вернул ошибку.</span><span class="sxs-lookup"><span data-stu-id="14fdd-531">The RMS server returned an error.</span></span> <span data-ttu-id="14fdd-532">Для получения дополнительных сведений просмотрите объект диагностики.</span><span class="sxs-lookup"><span data-stu-id="14fdd-532">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="14fdd-533">Пользователь отключен от сети.</span><span class="sxs-lookup"><span data-stu-id="14fdd-533">The user is no longer connected to the network.</span></span> <span data-ttu-id="14fdd-534">Проверьте сетевое подключение и повторите попытку.</span><span class="sxs-lookup"><span data-stu-id="14fdd-534">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="14fdd-535">Требования</span><span class="sxs-lookup"><span data-stu-id="14fdd-535">Requirements</span></span>

|<span data-ttu-id="14fdd-536">Требование</span><span class="sxs-lookup"><span data-stu-id="14fdd-536">Requirement</span></span>| <span data-ttu-id="14fdd-537">Значение</span><span class="sxs-lookup"><span data-stu-id="14fdd-537">Value</span></span>|
|---|---|
|[<span data-ttu-id="14fdd-538">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="14fdd-538">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14fdd-539">1.0</span><span class="sxs-lookup"><span data-stu-id="14fdd-539">1.0</span></span>|
|[<span data-ttu-id="14fdd-540">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14fdd-540">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14fdd-541">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14fdd-541">ReadItem</span></span>|
|[<span data-ttu-id="14fdd-542">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14fdd-542">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14fdd-543">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="14fdd-543">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="14fdd-544">Пример</span><span class="sxs-lookup"><span data-stu-id="14fdd-544">Example</span></span>

```javascript
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="14fdd-545">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="14fdd-545">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="14fdd-546">Выполняет асинхронный запрос для веб-служб Exchange (EWS) на сервере Exchange Server, на котором размещен почтовый ящик пользователя.</span><span class="sxs-lookup"><span data-stu-id="14fdd-546">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="14fdd-547">Этот метод не поддерживается в следующих сценариях:</span><span class="sxs-lookup"><span data-stu-id="14fdd-547">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="14fdd-548">В Outlook для iOS и Android</span><span class="sxs-lookup"><span data-stu-id="14fdd-548">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="14fdd-549">Если надстройка загружается в почтовый ящик Gmail.</span><span class="sxs-lookup"><span data-stu-id="14fdd-549">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="14fdd-550">В таких случаях надстройка должна [использовать REST API](/outlook/add-ins/use-rest-api) для доступа к почтовому ящику пользователя.</span><span class="sxs-lookup"><span data-stu-id="14fdd-550">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="14fdd-551">Метод `makeEwsRequestAsync` отправляет запрос EWS от имени надстройки в Exchange.</span><span class="sxs-lookup"><span data-stu-id="14fdd-551">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="14fdd-552">Список поддерживаемых операций EWS см. в статье [Вызов веб-служб из надстройки Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support).</span><span class="sxs-lookup"><span data-stu-id="14fdd-552">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="14fdd-553">С помощью метода `makeEwsRequestAsync` невозможно запрашивать элементы, связанные с папкой.</span><span class="sxs-lookup"><span data-stu-id="14fdd-553">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="14fdd-554">В запросе XML должна быть указана кодировка UTF-8.</span><span class="sxs-lookup"><span data-stu-id="14fdd-554">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="14fdd-p139">У вашей надстройки должно быть разрешение **ReadWriteMailbox** для использования метода `makeEwsRequestAsync`. Сведения об использовании разрешения **ReadWriteMailbox** и операций EWS, которые можно вызывать с помощью метода `makeEwsRequestAsync`, см. в статье [Указание разрешений для доступа почтовой надстройки к почтовому ящику пользователя](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="14fdd-p139">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="14fdd-557">Администратор сервера должен установить значение true для параметра `OAuthAuthentication` в каталоге сервера клиентского доступа EWS, чтобы метод `makeEwsRequestAsync` мог выполнять запросы EWS.</span><span class="sxs-lookup"><span data-stu-id="14fdd-557">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="14fdd-558">Различия версий</span><span class="sxs-lookup"><span data-stu-id="14fdd-558">Version differences</span></span>

<span data-ttu-id="14fdd-559">Если вы используете метод `makeEwsRequestAsync` в почтовых приложениях, которые выполняются в Outlook версии более ранней, чем 15.0.4535.1004, указывайте кодировку `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="14fdd-559">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="14fdd-p140">Значение кодировки не нужно указывать, если почтовое приложение выполняется в Outlook в Интернете. Чтобы определить, выполняется ли приложение в Outlook или Outlook в Интернете, используйте свойство mailbox.diagnostics.hostName. Используемую версию Outlook можно определить с помощью свойства mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="14fdd-p140">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="14fdd-563">Параметры</span><span class="sxs-lookup"><span data-stu-id="14fdd-563">Parameters</span></span>

|<span data-ttu-id="14fdd-564">Имя</span><span class="sxs-lookup"><span data-stu-id="14fdd-564">Name</span></span>| <span data-ttu-id="14fdd-565">Тип</span><span class="sxs-lookup"><span data-stu-id="14fdd-565">Type</span></span>| <span data-ttu-id="14fdd-566">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="14fdd-566">Attributes</span></span>| <span data-ttu-id="14fdd-567">Описание</span><span class="sxs-lookup"><span data-stu-id="14fdd-567">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="14fdd-568">String</span><span class="sxs-lookup"><span data-stu-id="14fdd-568">String</span></span>||<span data-ttu-id="14fdd-569">Запрос EWS.</span><span class="sxs-lookup"><span data-stu-id="14fdd-569">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="14fdd-570">function</span><span class="sxs-lookup"><span data-stu-id="14fdd-570">function</span></span>||<span data-ttu-id="14fdd-571">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="14fdd-571">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="14fdd-572">Результат XML вызова EWS указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="14fdd-572">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="14fdd-573">Если размер результата превышает 1 МБ, возвращается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="14fdd-573">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="14fdd-574">Объект</span><span class="sxs-lookup"><span data-stu-id="14fdd-574">Object</span></span>| <span data-ttu-id="14fdd-575">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="14fdd-575">&lt;optional&gt;</span></span>|<span data-ttu-id="14fdd-576">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="14fdd-576">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="14fdd-577">Требования</span><span class="sxs-lookup"><span data-stu-id="14fdd-577">Requirements</span></span>

|<span data-ttu-id="14fdd-578">Требование</span><span class="sxs-lookup"><span data-stu-id="14fdd-578">Requirement</span></span>| <span data-ttu-id="14fdd-579">Значение</span><span class="sxs-lookup"><span data-stu-id="14fdd-579">Value</span></span>|
|---|---|
|[<span data-ttu-id="14fdd-580">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="14fdd-580">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14fdd-581">1.0</span><span class="sxs-lookup"><span data-stu-id="14fdd-581">1.0</span></span>|
|[<span data-ttu-id="14fdd-582">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14fdd-582">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14fdd-583">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="14fdd-583">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="14fdd-584">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14fdd-584">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14fdd-585">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="14fdd-585">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="14fdd-586">Пример</span><span class="sxs-lookup"><span data-stu-id="14fdd-586">Example</span></span>

<span data-ttu-id="14fdd-587">В приведенном ниже примере вызывается `makeEwsRequestAsync` для получения темы элемента с помощью операции `GetItem`.</span><span class="sxs-lookup"><span data-stu-id="14fdd-587">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="14fdd-588">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="14fdd-588">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="14fdd-589">Удаляет обработчиков для поддерживаемого типа события.</span><span class="sxs-lookup"><span data-stu-id="14fdd-589">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="14fdd-590">В настоящее время единственный поддерживаемый тип события — `Office.EventType.ItemChanged`.</span><span class="sxs-lookup"><span data-stu-id="14fdd-590">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="14fdd-591">Параметры</span><span class="sxs-lookup"><span data-stu-id="14fdd-591">Parameters</span></span>

| <span data-ttu-id="14fdd-592">Имя</span><span class="sxs-lookup"><span data-stu-id="14fdd-592">Name</span></span> | <span data-ttu-id="14fdd-593">Тип</span><span class="sxs-lookup"><span data-stu-id="14fdd-593">Type</span></span> | <span data-ttu-id="14fdd-594">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="14fdd-594">Attributes</span></span> | <span data-ttu-id="14fdd-595">Описание</span><span class="sxs-lookup"><span data-stu-id="14fdd-595">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="14fdd-596">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="14fdd-596">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="14fdd-597">Событие, которое должно отменить обработчик.</span><span class="sxs-lookup"><span data-stu-id="14fdd-597">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="14fdd-598">Объект</span><span class="sxs-lookup"><span data-stu-id="14fdd-598">Object</span></span> | <span data-ttu-id="14fdd-599">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="14fdd-599">&lt;optional&gt;</span></span> | <span data-ttu-id="14fdd-600">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="14fdd-600">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="14fdd-601">Object</span><span class="sxs-lookup"><span data-stu-id="14fdd-601">Object</span></span> | <span data-ttu-id="14fdd-602">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="14fdd-602">&lt;optional&gt;</span></span> | <span data-ttu-id="14fdd-603">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="14fdd-603">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="14fdd-604">функция</span><span class="sxs-lookup"><span data-stu-id="14fdd-604">function</span></span>| <span data-ttu-id="14fdd-605">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="14fdd-605">&lt;optional&gt;</span></span>|<span data-ttu-id="14fdd-606">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="14fdd-606">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="14fdd-607">Требования</span><span class="sxs-lookup"><span data-stu-id="14fdd-607">Requirements</span></span>

|<span data-ttu-id="14fdd-608">Требование</span><span class="sxs-lookup"><span data-stu-id="14fdd-608">Requirement</span></span>| <span data-ttu-id="14fdd-609">Значение</span><span class="sxs-lookup"><span data-stu-id="14fdd-609">Value</span></span>|
|---|---|
|[<span data-ttu-id="14fdd-610">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="14fdd-610">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="14fdd-611">1.5</span><span class="sxs-lookup"><span data-stu-id="14fdd-611">1.5</span></span> |
|[<span data-ttu-id="14fdd-612">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="14fdd-612">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="14fdd-613">ReadItem</span><span class="sxs-lookup"><span data-stu-id="14fdd-613">ReadItem</span></span> |
|[<span data-ttu-id="14fdd-614">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="14fdd-614">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="14fdd-615">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="14fdd-615">Compose or Read</span></span>|
