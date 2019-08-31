---
title: Office.context.mailbox — набор обязательных элементов 1.5
description: ''
ms.date: 08/30/2019
localization_priority: Priority
ms.openlocfilehash: 62834db09742f2f11eb73d571f22c7a249f36763
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696101"
---
# <a name="mailbox"></a><span data-ttu-id="9c901-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="9c901-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="9c901-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="9c901-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="9c901-104">Предоставляет для Microsoft Outlook доступ к объектной модели надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="9c901-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="9c901-105">Требования</span><span class="sxs-lookup"><span data-stu-id="9c901-105">Requirements</span></span>

|<span data-ttu-id="9c901-106">Требование</span><span class="sxs-lookup"><span data-stu-id="9c901-106">Requirement</span></span>| <span data-ttu-id="9c901-107">Значение</span><span class="sxs-lookup"><span data-stu-id="9c901-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c901-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9c901-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c901-109">1.0</span><span class="sxs-lookup"><span data-stu-id="9c901-109">1.0</span></span>|
|[<span data-ttu-id="9c901-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9c901-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c901-111">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="9c901-111">Restricted</span></span>|
|[<span data-ttu-id="9c901-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9c901-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9c901-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9c901-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="9c901-114">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="9c901-114">Members and methods</span></span>

| <span data-ttu-id="9c901-115">Элемент</span><span class="sxs-lookup"><span data-stu-id="9c901-115">Member</span></span> | <span data-ttu-id="9c901-116">Тип</span><span class="sxs-lookup"><span data-stu-id="9c901-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="9c901-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="9c901-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="9c901-118">Элемент</span><span class="sxs-lookup"><span data-stu-id="9c901-118">Member</span></span> |
| [<span data-ttu-id="9c901-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="9c901-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="9c901-120">Элемент</span><span class="sxs-lookup"><span data-stu-id="9c901-120">Member</span></span> |
| [<span data-ttu-id="9c901-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="9c901-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="9c901-122">Метод</span><span class="sxs-lookup"><span data-stu-id="9c901-122">Method</span></span> |
| [<span data-ttu-id="9c901-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="9c901-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="9c901-124">Метод</span><span class="sxs-lookup"><span data-stu-id="9c901-124">Method</span></span> |
| [<span data-ttu-id="9c901-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="9c901-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="9c901-126">Метод</span><span class="sxs-lookup"><span data-stu-id="9c901-126">Method</span></span> |
| [<span data-ttu-id="9c901-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="9c901-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="9c901-128">Метод</span><span class="sxs-lookup"><span data-stu-id="9c901-128">Method</span></span> |
| [<span data-ttu-id="9c901-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="9c901-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="9c901-130">Метод</span><span class="sxs-lookup"><span data-stu-id="9c901-130">Method</span></span> |
| [<span data-ttu-id="9c901-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="9c901-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="9c901-132">Метод</span><span class="sxs-lookup"><span data-stu-id="9c901-132">Method</span></span> |
| [<span data-ttu-id="9c901-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="9c901-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="9c901-134">Метод</span><span class="sxs-lookup"><span data-stu-id="9c901-134">Method</span></span> |
| [<span data-ttu-id="9c901-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="9c901-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="9c901-136">Метод</span><span class="sxs-lookup"><span data-stu-id="9c901-136">Method</span></span> |
| [<span data-ttu-id="9c901-137">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="9c901-137">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="9c901-138">Метод</span><span class="sxs-lookup"><span data-stu-id="9c901-138">Method</span></span> |
| [<span data-ttu-id="9c901-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="9c901-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="9c901-140">Метод</span><span class="sxs-lookup"><span data-stu-id="9c901-140">Method</span></span> |
| [<span data-ttu-id="9c901-141">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="9c901-141">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="9c901-142">Метод</span><span class="sxs-lookup"><span data-stu-id="9c901-142">Method</span></span> |
| [<span data-ttu-id="9c901-143">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="9c901-143">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="9c901-144">Метод</span><span class="sxs-lookup"><span data-stu-id="9c901-144">Method</span></span> |
| [<span data-ttu-id="9c901-145">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="9c901-145">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="9c901-146">Метод</span><span class="sxs-lookup"><span data-stu-id="9c901-146">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="9c901-147">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="9c901-147">Namespaces</span></span>

<span data-ttu-id="9c901-148">[diagnostics](Office.context.mailbox.diagnostics.md). Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="9c901-148">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="9c901-149">[item](Office.context.mailbox.item.md). Предоставляет методы и свойства для доступа к сообщению или встрече в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="9c901-149">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="9c901-150">[userProfile](Office.context.mailbox.userProfile.md). Предоставляет сведения о пользователе в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="9c901-150">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="9c901-151">Members</span><span class="sxs-lookup"><span data-stu-id="9c901-151">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="9c901-152">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="9c901-152">ewsUrl :String</span></span>

<span data-ttu-id="9c901-p101">Получает URL-адрес конечной точки веб-служб Exchange (EWS) для этой учетной записи электронной почты. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="9c901-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="9c901-155">Этот элемент не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="9c901-155">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9c901-p102">Удаленная служба может использовать значение `ewsUrl`, чтобы выполнять вызовы EWS для почтового ящика пользователя. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="9c901-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="9c901-158">Чтобы вызвать элемент `ewsUrl` в режиме чтения, в манифесте приложения должно быть указано разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="9c901-158">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="9c901-p103">Перед использованием элемента `ewsUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="9c901-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="9c901-161">Тип</span><span class="sxs-lookup"><span data-stu-id="9c901-161">Type</span></span>

*   <span data-ttu-id="9c901-162">String</span><span class="sxs-lookup"><span data-stu-id="9c901-162">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9c901-163">Требования</span><span class="sxs-lookup"><span data-stu-id="9c901-163">Requirements</span></span>

|<span data-ttu-id="9c901-164">Требование</span><span class="sxs-lookup"><span data-stu-id="9c901-164">Requirement</span></span>| <span data-ttu-id="9c901-165">Значение</span><span class="sxs-lookup"><span data-stu-id="9c901-165">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c901-166">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9c901-166">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c901-167">1.0</span><span class="sxs-lookup"><span data-stu-id="9c901-167">1.0</span></span>|
|[<span data-ttu-id="9c901-168">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9c901-168">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c901-169">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c901-169">ReadItem</span></span>|
|[<span data-ttu-id="9c901-170">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9c901-170">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9c901-171">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9c901-171">Compose or Read</span></span>|

<br>

---
---

#### <a name="resturl-string"></a><span data-ttu-id="9c901-172">restUrl: String</span><span class="sxs-lookup"><span data-stu-id="9c901-172">restUrl :String</span></span>

<span data-ttu-id="9c901-173">Возвращает URL-адрес конечной точки REST для этой учетной записи электронной почты.</span><span class="sxs-lookup"><span data-stu-id="9c901-173">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="9c901-174">С помощью значения `restUrl` можно выполнять вызовы [REST API](/outlook/rest/) для почтового ящика пользователя.</span><span class="sxs-lookup"><span data-stu-id="9c901-174">The `restUrl` value can be used to make [REST API](/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="9c901-175">Чтобы вызвать элемент `restUrl` в режиме чтения, в манифесте приложения необходимо указать разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="9c901-175">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="9c901-p104">Перед использованием элемента `restUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="9c901-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

> [!NOTE]
> <span data-ttu-id="9c901-178">Клиенты Outlook, подключенные к локальным установленным версиям Exchange 2016 или более поздним с пользовательским URL-адресом REST, возвращают недопустимое значение `restUrl`.</span><span class="sxs-lookup"><span data-stu-id="9c901-178">Outlook clients connected to on-premises installations of Exchange 2016 or later with a custom REST URL configured will return an invalid value for `restUrl`.</span></span>

##### <a name="type"></a><span data-ttu-id="9c901-179">Тип</span><span class="sxs-lookup"><span data-stu-id="9c901-179">Type</span></span>

*   <span data-ttu-id="9c901-180">String</span><span class="sxs-lookup"><span data-stu-id="9c901-180">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9c901-181">Требования</span><span class="sxs-lookup"><span data-stu-id="9c901-181">Requirements</span></span>

|<span data-ttu-id="9c901-182">Требование</span><span class="sxs-lookup"><span data-stu-id="9c901-182">Requirement</span></span>| <span data-ttu-id="9c901-183">Значение</span><span class="sxs-lookup"><span data-stu-id="9c901-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c901-184">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9c901-184">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c901-185">1.5</span><span class="sxs-lookup"><span data-stu-id="9c901-185">1.5</span></span> |
|[<span data-ttu-id="9c901-186">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9c901-186">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c901-187">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c901-187">ReadItem</span></span>|
|[<span data-ttu-id="9c901-188">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9c901-188">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9c901-189">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9c901-189">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="9c901-190">Методы</span><span class="sxs-lookup"><span data-stu-id="9c901-190">Methods</span></span>

#### <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="9c901-191">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="9c901-191">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="9c901-192">Добавляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="9c901-192">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="9c901-193">В настоящее время поддерживаются только события типа `Office.EventType.ItemChanged`, которые вызываются, когда пользователь выбирает новый элемент.</span><span class="sxs-lookup"><span data-stu-id="9c901-193">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item.</span></span> <span data-ttu-id="9c901-194">Это событие используется надстройками, реализующими закрепляемую область задач, и позволяет надстройке обновлять пользовательский интерфейс области задач в соответствии с выбранным в данный момент элементом.</span><span class="sxs-lookup"><span data-stu-id="9c901-194">This event is used by add-ins that implement a pinnable task pane, and allows the add-in to refresh the task pane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9c901-195">Параметры</span><span class="sxs-lookup"><span data-stu-id="9c901-195">Parameters</span></span>

| <span data-ttu-id="9c901-196">Имя</span><span class="sxs-lookup"><span data-stu-id="9c901-196">Name</span></span> | <span data-ttu-id="9c901-197">Тип</span><span class="sxs-lookup"><span data-stu-id="9c901-197">Type</span></span> | <span data-ttu-id="9c901-198">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="9c901-198">Attributes</span></span> | <span data-ttu-id="9c901-199">Описание</span><span class="sxs-lookup"><span data-stu-id="9c901-199">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="9c901-200">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="9c901-200">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="9c901-201">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="9c901-201">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="9c901-202">Function</span><span class="sxs-lookup"><span data-stu-id="9c901-202">Function</span></span> || <span data-ttu-id="9c901-p106">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="9c901-p106">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="9c901-206">Объект</span><span class="sxs-lookup"><span data-stu-id="9c901-206">Object</span></span> | <span data-ttu-id="9c901-207">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9c901-207">&lt;optional&gt;</span></span> | <span data-ttu-id="9c901-208">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="9c901-208">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="9c901-209">Object</span><span class="sxs-lookup"><span data-stu-id="9c901-209">Object</span></span> | <span data-ttu-id="9c901-210">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9c901-210">&lt;optional&gt;</span></span> | <span data-ttu-id="9c901-211">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="9c901-211">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="9c901-212">функция</span><span class="sxs-lookup"><span data-stu-id="9c901-212">function</span></span>| <span data-ttu-id="9c901-213">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9c901-213">&lt;optional&gt;</span></span>|<span data-ttu-id="9c901-214">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9c901-214">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9c901-215">Требования</span><span class="sxs-lookup"><span data-stu-id="9c901-215">Requirements</span></span>

|<span data-ttu-id="9c901-216">Требование</span><span class="sxs-lookup"><span data-stu-id="9c901-216">Requirement</span></span>| <span data-ttu-id="9c901-217">Значение</span><span class="sxs-lookup"><span data-stu-id="9c901-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c901-218">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9c901-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c901-219">1.5</span><span class="sxs-lookup"><span data-stu-id="9c901-219">1.5</span></span> |
|[<span data-ttu-id="9c901-220">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9c901-220">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c901-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c901-221">ReadItem</span></span> |
|[<span data-ttu-id="9c901-222">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9c901-222">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9c901-223">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9c901-223">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9c901-224">Пример</span><span class="sxs-lookup"><span data-stu-id="9c901-224">Example</span></span>

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

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="9c901-225">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="9c901-225">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="9c901-226">Преобразовывает идентификатор элемента из формата REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="9c901-226">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="9c901-227">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="9c901-227">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9c901-p107">Формат идентификаторов, извлекаемых через API REST (например, [API Почты Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)), отличается от формата веб-служб Exchange (EWS). Метод `convertToEwsId` преобразовывает идентификатор в формате REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="9c901-p107">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9c901-230">Параметры</span><span class="sxs-lookup"><span data-stu-id="9c901-230">Parameters</span></span>

|<span data-ttu-id="9c901-231">Имя</span><span class="sxs-lookup"><span data-stu-id="9c901-231">Name</span></span>| <span data-ttu-id="9c901-232">Тип</span><span class="sxs-lookup"><span data-stu-id="9c901-232">Type</span></span>| <span data-ttu-id="9c901-233">Описание</span><span class="sxs-lookup"><span data-stu-id="9c901-233">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="9c901-234">String</span><span class="sxs-lookup"><span data-stu-id="9c901-234">String</span></span>|<span data-ttu-id="9c901-235">Идентификатор элемента в формате REST API для Outlook</span><span class="sxs-lookup"><span data-stu-id="9c901-235">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="9c901-236">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="9c901-236">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.5)|<span data-ttu-id="9c901-237">Значение, определяющее версию REST API для Outlook, которая используется для извлечения идентификатора элемента.</span><span class="sxs-lookup"><span data-stu-id="9c901-237">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9c901-238">Требования</span><span class="sxs-lookup"><span data-stu-id="9c901-238">Requirements</span></span>

|<span data-ttu-id="9c901-239">Требование</span><span class="sxs-lookup"><span data-stu-id="9c901-239">Requirement</span></span>| <span data-ttu-id="9c901-240">Значение</span><span class="sxs-lookup"><span data-stu-id="9c901-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c901-241">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9c901-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c901-242">1.3</span><span class="sxs-lookup"><span data-stu-id="9c901-242">1.3</span></span>|
|[<span data-ttu-id="9c901-243">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9c901-243">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c901-244">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="9c901-244">Restricted</span></span>|
|[<span data-ttu-id="9c901-245">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9c901-245">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9c901-246">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9c901-246">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9c901-247">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="9c901-247">Returns:</span></span>

<span data-ttu-id="9c901-248">Тип: String</span><span class="sxs-lookup"><span data-stu-id="9c901-248">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="9c901-249">Пример</span><span class="sxs-lookup"><span data-stu-id="9c901-249">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-15"></a><span data-ttu-id="9c901-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)}</span><span class="sxs-lookup"><span data-stu-id="9c901-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)}</span></span>

<span data-ttu-id="9c901-251">Получает словарь, содержащий сведения о локальном времени клиента.</span><span class="sxs-lookup"><span data-stu-id="9c901-251">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="9c901-p108">Почтовое приложение для классической версии Outlook или версии в Интернете может использовать разные часовые пояса для дат и времени. Классическое приложение Outlook использует часовой пояс клиентского компьютера. Outlook в Интернете использует часовой пояс, заданный в Центре администрирования Exchange (EAC). Значения даты и времени должны обрабатываться так, чтобы значения в пользовательском интерфейсе всегда согласовывались с часовым поясом, ожидаемым пользователем.</span><span class="sxs-lookup"><span data-stu-id="9c901-p108">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="9c901-p109">Если почтовое приложение работает в классическом клиенте Outlook, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса клиентского компьютера. Если почтовое приложение работает в Outlook в Интернете, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса, заданного в Центре администрирования Exchange.</span><span class="sxs-lookup"><span data-stu-id="9c901-p109">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9c901-257">Параметры</span><span class="sxs-lookup"><span data-stu-id="9c901-257">Parameters</span></span>

|<span data-ttu-id="9c901-258">Имя</span><span class="sxs-lookup"><span data-stu-id="9c901-258">Name</span></span>| <span data-ttu-id="9c901-259">Тип</span><span class="sxs-lookup"><span data-stu-id="9c901-259">Type</span></span>| <span data-ttu-id="9c901-260">Описание</span><span class="sxs-lookup"><span data-stu-id="9c901-260">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="9c901-261">Date</span><span class="sxs-lookup"><span data-stu-id="9c901-261">Date</span></span>|<span data-ttu-id="9c901-262">Объект Date</span><span class="sxs-lookup"><span data-stu-id="9c901-262">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9c901-263">Требования</span><span class="sxs-lookup"><span data-stu-id="9c901-263">Requirements</span></span>

|<span data-ttu-id="9c901-264">Требование</span><span class="sxs-lookup"><span data-stu-id="9c901-264">Requirement</span></span>| <span data-ttu-id="9c901-265">Значение</span><span class="sxs-lookup"><span data-stu-id="9c901-265">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c901-266">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9c901-266">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c901-267">1.0</span><span class="sxs-lookup"><span data-stu-id="9c901-267">1.0</span></span>|
|[<span data-ttu-id="9c901-268">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9c901-268">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c901-269">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c901-269">ReadItem</span></span>|
|[<span data-ttu-id="9c901-270">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9c901-270">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9c901-271">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9c901-271">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9c901-272">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="9c901-272">Returns:</span></span>

<span data-ttu-id="9c901-273">Тип: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)</span><span class="sxs-lookup"><span data-stu-id="9c901-273">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="9c901-274">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="9c901-274">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="9c901-275">Преобразовывает идентификатор элемента в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="9c901-275">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="9c901-276">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="9c901-276">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9c901-p110">Формат идентификаторов, извлекаемых через EWS или свойство `itemId`, отличается от формата API REST (таких как [API Почты Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)). Метод `convertToRestId` преобразовывает идентификатор в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="9c901-p110">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9c901-279">Параметры</span><span class="sxs-lookup"><span data-stu-id="9c901-279">Parameters</span></span>

|<span data-ttu-id="9c901-280">Имя</span><span class="sxs-lookup"><span data-stu-id="9c901-280">Name</span></span>| <span data-ttu-id="9c901-281">Тип</span><span class="sxs-lookup"><span data-stu-id="9c901-281">Type</span></span>| <span data-ttu-id="9c901-282">Описание</span><span class="sxs-lookup"><span data-stu-id="9c901-282">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="9c901-283">String</span><span class="sxs-lookup"><span data-stu-id="9c901-283">String</span></span>|<span data-ttu-id="9c901-284">Идентификатор элемента в формате EWS</span><span class="sxs-lookup"><span data-stu-id="9c901-284">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="9c901-285">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="9c901-285">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.5)|<span data-ttu-id="9c901-286">Значение, определяющее версию REST API для Outlook, с которой будет использоваться преобразованный идентификатор.</span><span class="sxs-lookup"><span data-stu-id="9c901-286">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9c901-287">Требования</span><span class="sxs-lookup"><span data-stu-id="9c901-287">Requirements</span></span>

|<span data-ttu-id="9c901-288">Требование</span><span class="sxs-lookup"><span data-stu-id="9c901-288">Requirement</span></span>| <span data-ttu-id="9c901-289">Значение</span><span class="sxs-lookup"><span data-stu-id="9c901-289">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c901-290">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9c901-290">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c901-291">1.3</span><span class="sxs-lookup"><span data-stu-id="9c901-291">1.3</span></span>|
|[<span data-ttu-id="9c901-292">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9c901-292">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c901-293">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="9c901-293">Restricted</span></span>|
|[<span data-ttu-id="9c901-294">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9c901-294">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9c901-295">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9c901-295">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9c901-296">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="9c901-296">Returns:</span></span>

<span data-ttu-id="9c901-297">Тип: String</span><span class="sxs-lookup"><span data-stu-id="9c901-297">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="9c901-298">Пример</span><span class="sxs-lookup"><span data-stu-id="9c901-298">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="9c901-299">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="9c901-299">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="9c901-300">Получает объект Date из словаря, содержащего сведения о времени.</span><span class="sxs-lookup"><span data-stu-id="9c901-300">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="9c901-301">Метод `convertToUtcClientTime` преобразует словарь, содержащий локальную дату и время, в объект Date с правильными значениями локальной даты и времени.</span><span class="sxs-lookup"><span data-stu-id="9c901-301">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9c901-302">Параметры</span><span class="sxs-lookup"><span data-stu-id="9c901-302">Parameters</span></span>

|<span data-ttu-id="9c901-303">Имя</span><span class="sxs-lookup"><span data-stu-id="9c901-303">Name</span></span>| <span data-ttu-id="9c901-304">Тип</span><span class="sxs-lookup"><span data-stu-id="9c901-304">Type</span></span>| <span data-ttu-id="9c901-305">Описание</span><span class="sxs-lookup"><span data-stu-id="9c901-305">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="9c901-306">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="9c901-306">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.5)|<span data-ttu-id="9c901-307">Значение локального времени для преобразования.</span><span class="sxs-lookup"><span data-stu-id="9c901-307">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9c901-308">Требования</span><span class="sxs-lookup"><span data-stu-id="9c901-308">Requirements</span></span>

|<span data-ttu-id="9c901-309">Требование</span><span class="sxs-lookup"><span data-stu-id="9c901-309">Requirement</span></span>| <span data-ttu-id="9c901-310">Значение</span><span class="sxs-lookup"><span data-stu-id="9c901-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c901-311">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9c901-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c901-312">1.0</span><span class="sxs-lookup"><span data-stu-id="9c901-312">1.0</span></span>|
|[<span data-ttu-id="9c901-313">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9c901-313">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c901-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c901-314">ReadItem</span></span>|
|[<span data-ttu-id="9c901-315">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9c901-315">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9c901-316">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9c901-316">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="9c901-317">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="9c901-317">Returns:</span></span>

<span data-ttu-id="9c901-318">Объект Date со временем в формате UTC.</span><span class="sxs-lookup"><span data-stu-id="9c901-318">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="9c901-319">Тип: Date</span><span class="sxs-lookup"><span data-stu-id="9c901-319">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="9c901-320">Пример</span><span class="sxs-lookup"><span data-stu-id="9c901-320">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="9c901-321">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="9c901-321">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="9c901-322">Отображает имеющуюся встречу из календаря.</span><span class="sxs-lookup"><span data-stu-id="9c901-322">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="9c901-323">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="9c901-323">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9c901-324">Метод `displayAppointmentForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее сведения календаря о существующей встрече.</span><span class="sxs-lookup"><span data-stu-id="9c901-324">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="9c901-p111">В Outlook для Mac с помощью этого метода можно отобразить одну встречу, которая не является частью повторяющегося ряда, или основную встречу такого ряда, но не экземпляр из него, так как в Outlook для Mac невозможно получить доступ к свойствам экземпляра повторяющегося ряда (в том числе к идентификатору элемента).</span><span class="sxs-lookup"><span data-stu-id="9c901-p111">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="9c901-327">В Outlook в Интернете этот метод открывает указанную форму, только если текст формы содержит символы размером не более 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="9c901-327">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="9c901-328">Если указанный идентификатор элемента не определяет существующую встречу, на клиентском компьютере или устройстве открывается пустая страница, и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="9c901-328">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9c901-329">Параметры</span><span class="sxs-lookup"><span data-stu-id="9c901-329">Parameters</span></span>

|<span data-ttu-id="9c901-330">Имя</span><span class="sxs-lookup"><span data-stu-id="9c901-330">Name</span></span>| <span data-ttu-id="9c901-331">Тип</span><span class="sxs-lookup"><span data-stu-id="9c901-331">Type</span></span>| <span data-ttu-id="9c901-332">Описание</span><span class="sxs-lookup"><span data-stu-id="9c901-332">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="9c901-333">String</span><span class="sxs-lookup"><span data-stu-id="9c901-333">String</span></span>|<span data-ttu-id="9c901-334">Идентификатор веб-служб Exchange для существующей встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="9c901-334">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9c901-335">Требования</span><span class="sxs-lookup"><span data-stu-id="9c901-335">Requirements</span></span>

|<span data-ttu-id="9c901-336">Требование</span><span class="sxs-lookup"><span data-stu-id="9c901-336">Requirement</span></span>| <span data-ttu-id="9c901-337">Значение</span><span class="sxs-lookup"><span data-stu-id="9c901-337">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c901-338">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9c901-338">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c901-339">1.0</span><span class="sxs-lookup"><span data-stu-id="9c901-339">1.0</span></span>|
|[<span data-ttu-id="9c901-340">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9c901-340">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c901-341">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c901-341">ReadItem</span></span>|
|[<span data-ttu-id="9c901-342">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9c901-342">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9c901-343">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9c901-343">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9c901-344">Пример</span><span class="sxs-lookup"><span data-stu-id="9c901-344">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="9c901-345">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="9c901-345">displayMessageForm(itemId)</span></span>

<span data-ttu-id="9c901-346">Отображает имеющееся сообщение.</span><span class="sxs-lookup"><span data-stu-id="9c901-346">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="9c901-347">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="9c901-347">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9c901-348">Метод `displayMessageForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее существующее сообщение.</span><span class="sxs-lookup"><span data-stu-id="9c901-348">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="9c901-349">В Outlook в Интернете этот метод открывает указанную форму, только если текст формы содержит символы размером не более 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="9c901-349">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="9c901-350">Если указанный идентификатор элемента не определяет существующее сообщение, окно на клиентском компьютере не открывается и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="9c901-350">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="9c901-p112">Не используйте `displayMessageForm` с параметром `itemId`, который представляет собой встречу. Используйте метод `displayAppointmentForm`, чтобы отобразить сведения о существующей встрече, а метод `displayNewAppointmentForm` — для отображения формы создания встречи.</span><span class="sxs-lookup"><span data-stu-id="9c901-p112">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9c901-353">Параметры</span><span class="sxs-lookup"><span data-stu-id="9c901-353">Parameters</span></span>

|<span data-ttu-id="9c901-354">Имя</span><span class="sxs-lookup"><span data-stu-id="9c901-354">Name</span></span>| <span data-ttu-id="9c901-355">Тип</span><span class="sxs-lookup"><span data-stu-id="9c901-355">Type</span></span>| <span data-ttu-id="9c901-356">Описание</span><span class="sxs-lookup"><span data-stu-id="9c901-356">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="9c901-357">String</span><span class="sxs-lookup"><span data-stu-id="9c901-357">String</span></span>|<span data-ttu-id="9c901-358">Идентификатор веб-служб Exchange для существующего сообщения.</span><span class="sxs-lookup"><span data-stu-id="9c901-358">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9c901-359">Требования</span><span class="sxs-lookup"><span data-stu-id="9c901-359">Requirements</span></span>

|<span data-ttu-id="9c901-360">Требование</span><span class="sxs-lookup"><span data-stu-id="9c901-360">Requirement</span></span>| <span data-ttu-id="9c901-361">Значение</span><span class="sxs-lookup"><span data-stu-id="9c901-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c901-362">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9c901-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c901-363">1.0</span><span class="sxs-lookup"><span data-stu-id="9c901-363">1.0</span></span>|
|[<span data-ttu-id="9c901-364">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9c901-364">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c901-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c901-365">ReadItem</span></span>|
|[<span data-ttu-id="9c901-366">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9c901-366">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9c901-367">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9c901-367">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9c901-368">Пример</span><span class="sxs-lookup"><span data-stu-id="9c901-368">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="9c901-369">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="9c901-369">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="9c901-370">Отображает форму для создания новой встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="9c901-370">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="9c901-371">Этот метод не поддерживается в Outlook для iOS и Android.</span><span class="sxs-lookup"><span data-stu-id="9c901-371">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="9c901-p113">Метод `displayNewAppointmentForm` открывает форму, в которой пользователь может создать встречу или собрание. Если параметры заданы, поля формы встречи автоматически заполняются их содержимым.</span><span class="sxs-lookup"><span data-stu-id="9c901-p113">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="9c901-p114">В Outlook в Интернете и на мобильных устройствах этот метод всегда отображает форму с полем участников. Если вы не укажете участников в качестве входных аргументов, метод отображает форму с кнопкой **Сохранить**. Если вы укажете участников, форма будет включать участников и кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="9c901-p114">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="9c901-p115">Если вы укажете участников или ресурсы с помощью параметра `requiredAttendees`, `optionalAttendees` или `resources` в клиенте Outlook с расширенными возможностями и Outlook RT, этот метод отобразит форму собрания с кнопкой **Отправить**. Если не указать получателей, этот метод отобразит форму встречи с кнопкой **Сохранить и закрыть**.</span><span class="sxs-lookup"><span data-stu-id="9c901-p115">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="9c901-379">Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.</span><span class="sxs-lookup"><span data-stu-id="9c901-379">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9c901-380">Параметры</span><span class="sxs-lookup"><span data-stu-id="9c901-380">Parameters</span></span>

|<span data-ttu-id="9c901-381">Имя</span><span class="sxs-lookup"><span data-stu-id="9c901-381">Name</span></span>| <span data-ttu-id="9c901-382">Тип</span><span class="sxs-lookup"><span data-stu-id="9c901-382">Type</span></span>| <span data-ttu-id="9c901-383">Описание</span><span class="sxs-lookup"><span data-stu-id="9c901-383">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="9c901-384">Object</span><span class="sxs-lookup"><span data-stu-id="9c901-384">Object</span></span> | <span data-ttu-id="9c901-385">Словарь параметров, описывающий новую встречу.</span><span class="sxs-lookup"><span data-stu-id="9c901-385">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="9c901-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span><span class="sxs-lookup"><span data-stu-id="9c901-386">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span></span> | <span data-ttu-id="9c901-p116">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из обязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="9c901-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="9c901-389">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span><span class="sxs-lookup"><span data-stu-id="9c901-389">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.5)&gt;</span></span> | <span data-ttu-id="9c901-p117">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из необязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="9c901-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="9c901-392">Date</span><span class="sxs-lookup"><span data-stu-id="9c901-392">Date</span></span> | <span data-ttu-id="9c901-393">Объект `Date`, указывающий дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="9c901-393">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="9c901-394">Date</span><span class="sxs-lookup"><span data-stu-id="9c901-394">Date</span></span> | <span data-ttu-id="9c901-395">Объект `Date`, указывающий дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="9c901-395">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="9c901-396">String</span><span class="sxs-lookup"><span data-stu-id="9c901-396">String</span></span> | <span data-ttu-id="9c901-p118">Строка со сведениями о месте встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="9c901-p118">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="9c901-399">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="9c901-399">Array.&lt;String&gt;</span></span> | <span data-ttu-id="9c901-p119">Массив строк, содержащий необходимые для встречи ресурсы. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="9c901-p119">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="9c901-402">String</span><span class="sxs-lookup"><span data-stu-id="9c901-402">String</span></span> | <span data-ttu-id="9c901-p120">Строка с темой встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="9c901-p120">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="9c901-405">String</span><span class="sxs-lookup"><span data-stu-id="9c901-405">String</span></span> | <span data-ttu-id="9c901-p121">Текст сообщения о встрече. Максимальный размер содержимого сообщения — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="9c901-p121">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9c901-408">Требования</span><span class="sxs-lookup"><span data-stu-id="9c901-408">Requirements</span></span>

|<span data-ttu-id="9c901-409">Требование</span><span class="sxs-lookup"><span data-stu-id="9c901-409">Requirement</span></span>| <span data-ttu-id="9c901-410">Значение</span><span class="sxs-lookup"><span data-stu-id="9c901-410">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c901-411">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9c901-411">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c901-412">1.0</span><span class="sxs-lookup"><span data-stu-id="9c901-412">1.0</span></span>|
|[<span data-ttu-id="9c901-413">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9c901-413">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c901-414">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c901-414">ReadItem</span></span>|
|[<span data-ttu-id="9c901-415">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9c901-415">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9c901-416">Чтение</span><span class="sxs-lookup"><span data-stu-id="9c901-416">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9c901-417">Пример</span><span class="sxs-lookup"><span data-stu-id="9c901-417">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="9c901-418">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="9c901-418">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="9c901-419">Возвращает строку, содержащую маркер, который используется для вызова интерфейсов REST API или веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="9c901-419">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="9c901-p122">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="9c901-p122">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="9c901-422">Рекомендуем сделать так, чтобы по мере возможности надстройки использовали интерфейсы REST API, а не веб-службы Exchange.</span><span class="sxs-lookup"><span data-stu-id="9c901-422">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="9c901-423">**Маркеры REST**</span><span class="sxs-lookup"><span data-stu-id="9c901-423">**REST Tokens**</span></span>

<span data-ttu-id="9c901-p123">Если запрашивается маркер REST (`options.isRest = true`), полученный маркер не подойдет для проверки подлинности при вызовах веб-служб Exchange. Область действия маркера будет ограничена доступом только для чтения к текущему элементу и его вложениям, если в манифесте надстройки не указано разрешение [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission). Если указано разрешение `ReadWriteMailbox`, полученный маркер предоставит доступ на чтение и запись к почте, календарю и контактам, включая возможность отправки почты.</span><span class="sxs-lookup"><span data-stu-id="9c901-p123">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="9c901-427">С помощью свойства `restUrl` надстройка должна определить правильный URL-адрес для вызовов REST API.</span><span class="sxs-lookup"><span data-stu-id="9c901-427">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="9c901-428">**Маркеры EWS**</span><span class="sxs-lookup"><span data-stu-id="9c901-428">**EWS Tokens**</span></span>

<span data-ttu-id="9c901-p124">Если запрашивается маркер EWS (`options.isRest = false`), полученный маркер не подойдет для проверки подлинности при вызовах REST API. Область действия маркера будет ограничена доступом к текущему элементу.</span><span class="sxs-lookup"><span data-stu-id="9c901-p124">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="9c901-431">С помощью свойства `ewsUrl` надстройка должна определить правильный URL-адрес для вызовов EWS.</span><span class="sxs-lookup"><span data-stu-id="9c901-431">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9c901-432">Параметры</span><span class="sxs-lookup"><span data-stu-id="9c901-432">Parameters</span></span>

|<span data-ttu-id="9c901-433">Имя</span><span class="sxs-lookup"><span data-stu-id="9c901-433">Name</span></span>| <span data-ttu-id="9c901-434">Тип</span><span class="sxs-lookup"><span data-stu-id="9c901-434">Type</span></span>| <span data-ttu-id="9c901-435">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="9c901-435">Attributes</span></span>| <span data-ttu-id="9c901-436">Описание</span><span class="sxs-lookup"><span data-stu-id="9c901-436">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="9c901-437">Object</span><span class="sxs-lookup"><span data-stu-id="9c901-437">Object</span></span> | <span data-ttu-id="9c901-438">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9c901-438">&lt;optional&gt;</span></span> | <span data-ttu-id="9c901-439">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="9c901-439">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="9c901-440">Boolean</span><span class="sxs-lookup"><span data-stu-id="9c901-440">Boolean</span></span> |  <span data-ttu-id="9c901-441">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9c901-441">&lt;optional&gt;</span></span> | <span data-ttu-id="9c901-p125">Определяет, будет ли предоставленный маркер использоваться для интерфейсов REST API Outlook или веб-служб Exchange. Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="9c901-p125">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="9c901-444">Объект</span><span class="sxs-lookup"><span data-stu-id="9c901-444">Object</span></span> |  <span data-ttu-id="9c901-445">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9c901-445">&lt;optional&gt;</span></span> | <span data-ttu-id="9c901-446">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="9c901-446">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="9c901-447">функция</span><span class="sxs-lookup"><span data-stu-id="9c901-447">function</span></span>||<span data-ttu-id="9c901-448">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9c901-448">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="9c901-449">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="9c901-449">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="9c901-450">При наличии ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут предоставлять дополнительные сведения.</span><span class="sxs-lookup"><span data-stu-id="9c901-450">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|

##### <a name="errors"></a><span data-ttu-id="9c901-451">Ошибки</span><span class="sxs-lookup"><span data-stu-id="9c901-451">Errors</span></span>

|<span data-ttu-id="9c901-452">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="9c901-452">Error code</span></span>|<span data-ttu-id="9c901-453">Описание</span><span class="sxs-lookup"><span data-stu-id="9c901-453">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="9c901-454">Не удалось выполнить запрос.</span><span class="sxs-lookup"><span data-stu-id="9c901-454">The request has failed.</span></span> <span data-ttu-id="9c901-455">Просмотрите объект диагностики для кода ошибки HTTP.</span><span class="sxs-lookup"><span data-stu-id="9c901-455">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="9c901-456">Сервер Exchange Server вернул ошибку.</span><span class="sxs-lookup"><span data-stu-id="9c901-456">The RMS server returned an error.</span></span> <span data-ttu-id="9c901-457">Для получения дополнительных сведений просмотрите объект диагностики.</span><span class="sxs-lookup"><span data-stu-id="9c901-457">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="9c901-458">Пользователь отключен от сети.</span><span class="sxs-lookup"><span data-stu-id="9c901-458">The user is no longer connected to the network.</span></span> <span data-ttu-id="9c901-459">Проверьте сетевое подключение и повторите попытку.</span><span class="sxs-lookup"><span data-stu-id="9c901-459">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9c901-460">Требования</span><span class="sxs-lookup"><span data-stu-id="9c901-460">Requirements</span></span>

|<span data-ttu-id="9c901-461">Требование</span><span class="sxs-lookup"><span data-stu-id="9c901-461">Requirement</span></span>| <span data-ttu-id="9c901-462">Значение</span><span class="sxs-lookup"><span data-stu-id="9c901-462">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c901-463">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9c901-463">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c901-464">1.5</span><span class="sxs-lookup"><span data-stu-id="9c901-464">1.5</span></span> |
|[<span data-ttu-id="9c901-465">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9c901-465">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c901-466">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c901-466">ReadItem</span></span>|
|[<span data-ttu-id="9c901-467">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9c901-467">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9c901-468">Создание и чтение</span><span class="sxs-lookup"><span data-stu-id="9c901-468">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="9c901-469">Пример</span><span class="sxs-lookup"><span data-stu-id="9c901-469">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="9c901-470">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="9c901-470">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="9c901-471">Получает строку, содержащую маркер, используемый для получения вложения или элемента с Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="9c901-471">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="9c901-p129">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="9c901-p129">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="9c901-p130">Вы можете передать сторонней системе маркер и идентификатор вложения или элемента. Сторонняя система использует этот маркер как маркер авторизации, чтобы вызвать операцию [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) или [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) веб-служб Exchange для возврата вложения или элемента. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="9c901-p130">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="9c901-477">Для вызова метода `getCallbackTokenAsync` в режиме чтения манифесте приложения должно быть указано разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="9c901-477">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="9c901-p131">Чтобы получить идентификатор элемента для передачи в метод `getCallbackTokenAsync`, в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="9c901-p131">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9c901-480">Параметры</span><span class="sxs-lookup"><span data-stu-id="9c901-480">Parameters</span></span>

|<span data-ttu-id="9c901-481">Имя</span><span class="sxs-lookup"><span data-stu-id="9c901-481">Name</span></span>| <span data-ttu-id="9c901-482">Тип</span><span class="sxs-lookup"><span data-stu-id="9c901-482">Type</span></span>| <span data-ttu-id="9c901-483">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="9c901-483">Attributes</span></span>| <span data-ttu-id="9c901-484">Описание</span><span class="sxs-lookup"><span data-stu-id="9c901-484">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="9c901-485">function</span><span class="sxs-lookup"><span data-stu-id="9c901-485">function</span></span>||<span data-ttu-id="9c901-486">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9c901-486">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="9c901-487">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="9c901-487">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="9c901-488">При наличии ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут предоставлять дополнительные сведения.</span><span class="sxs-lookup"><span data-stu-id="9c901-488">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="9c901-489">Объект</span><span class="sxs-lookup"><span data-stu-id="9c901-489">Object</span></span>| <span data-ttu-id="9c901-490">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9c901-490">&lt;optional&gt;</span></span>|<span data-ttu-id="9c901-491">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="9c901-491">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="9c901-492">Ошибки</span><span class="sxs-lookup"><span data-stu-id="9c901-492">Errors</span></span>

|<span data-ttu-id="9c901-493">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="9c901-493">Error code</span></span>|<span data-ttu-id="9c901-494">Описание</span><span class="sxs-lookup"><span data-stu-id="9c901-494">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="9c901-495">Не удалось выполнить запрос.</span><span class="sxs-lookup"><span data-stu-id="9c901-495">The request has failed.</span></span> <span data-ttu-id="9c901-496">Просмотрите объект диагностики для кода ошибки HTTP.</span><span class="sxs-lookup"><span data-stu-id="9c901-496">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="9c901-497">Сервер Exchange Server вернул ошибку.</span><span class="sxs-lookup"><span data-stu-id="9c901-497">The RMS server returned an error.</span></span> <span data-ttu-id="9c901-498">Для получения дополнительных сведений просмотрите объект диагностики.</span><span class="sxs-lookup"><span data-stu-id="9c901-498">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="9c901-499">Пользователь отключен от сети.</span><span class="sxs-lookup"><span data-stu-id="9c901-499">The user is no longer connected to the network.</span></span> <span data-ttu-id="9c901-500">Проверьте сетевое подключение и повторите попытку.</span><span class="sxs-lookup"><span data-stu-id="9c901-500">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9c901-501">Требования</span><span class="sxs-lookup"><span data-stu-id="9c901-501">Requirements</span></span>

|<span data-ttu-id="9c901-502">Требование</span><span class="sxs-lookup"><span data-stu-id="9c901-502">Requirement</span></span>| <span data-ttu-id="9c901-503">Значение</span><span class="sxs-lookup"><span data-stu-id="9c901-503">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c901-504">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9c901-504">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c901-505">1.0</span><span class="sxs-lookup"><span data-stu-id="9c901-505">1.0</span></span>|
|[<span data-ttu-id="9c901-506">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9c901-506">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c901-507">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c901-507">ReadItem</span></span>|
|[<span data-ttu-id="9c901-508">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9c901-508">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9c901-509">Создание и чтение</span><span class="sxs-lookup"><span data-stu-id="9c901-509">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="9c901-510">Пример</span><span class="sxs-lookup"><span data-stu-id="9c901-510">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="9c901-511">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="9c901-511">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="9c901-512">Получает маркер, идентифицирующий пользователя и надстройку Office.</span><span class="sxs-lookup"><span data-stu-id="9c901-512">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="9c901-513">Метод `getUserIdentityTokenAsync` возвращает маркер, который можно использовать для идентификации, а также [проверки подлинности надстройки и пользователя в сторонней системе](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="9c901-513">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="9c901-514">Параметры</span><span class="sxs-lookup"><span data-stu-id="9c901-514">Parameters</span></span>

|<span data-ttu-id="9c901-515">Имя</span><span class="sxs-lookup"><span data-stu-id="9c901-515">Name</span></span>| <span data-ttu-id="9c901-516">Тип</span><span class="sxs-lookup"><span data-stu-id="9c901-516">Type</span></span>| <span data-ttu-id="9c901-517">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="9c901-517">Attributes</span></span>| <span data-ttu-id="9c901-518">Описание</span><span class="sxs-lookup"><span data-stu-id="9c901-518">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="9c901-519">function</span><span class="sxs-lookup"><span data-stu-id="9c901-519">function</span></span>||<span data-ttu-id="9c901-520">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9c901-520">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="9c901-521">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="9c901-521">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="9c901-522">При наличии ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут предоставлять дополнительные сведения.</span><span class="sxs-lookup"><span data-stu-id="9c901-522">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="9c901-523">Объект</span><span class="sxs-lookup"><span data-stu-id="9c901-523">Object</span></span>| <span data-ttu-id="9c901-524">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9c901-524">&lt;optional&gt;</span></span>|<span data-ttu-id="9c901-525">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="9c901-525">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="9c901-526">Ошибки</span><span class="sxs-lookup"><span data-stu-id="9c901-526">Errors</span></span>

|<span data-ttu-id="9c901-527">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="9c901-527">Error code</span></span>|<span data-ttu-id="9c901-528">Описание</span><span class="sxs-lookup"><span data-stu-id="9c901-528">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="9c901-529">Не удалось выполнить запрос.</span><span class="sxs-lookup"><span data-stu-id="9c901-529">The request has failed.</span></span> <span data-ttu-id="9c901-530">Просмотрите объект диагностики для кода ошибки HTTP.</span><span class="sxs-lookup"><span data-stu-id="9c901-530">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="9c901-531">Сервер Exchange Server вернул ошибку.</span><span class="sxs-lookup"><span data-stu-id="9c901-531">The RMS server returned an error.</span></span> <span data-ttu-id="9c901-532">Для получения дополнительных сведений просмотрите объект диагностики.</span><span class="sxs-lookup"><span data-stu-id="9c901-532">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="9c901-533">Пользователь отключен от сети.</span><span class="sxs-lookup"><span data-stu-id="9c901-533">The user is no longer connected to the network.</span></span> <span data-ttu-id="9c901-534">Проверьте сетевое подключение и повторите попытку.</span><span class="sxs-lookup"><span data-stu-id="9c901-534">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9c901-535">Требования</span><span class="sxs-lookup"><span data-stu-id="9c901-535">Requirements</span></span>

|<span data-ttu-id="9c901-536">Требование</span><span class="sxs-lookup"><span data-stu-id="9c901-536">Requirement</span></span>| <span data-ttu-id="9c901-537">Значение</span><span class="sxs-lookup"><span data-stu-id="9c901-537">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c901-538">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9c901-538">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c901-539">1.0</span><span class="sxs-lookup"><span data-stu-id="9c901-539">1.0</span></span>|
|[<span data-ttu-id="9c901-540">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9c901-540">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c901-541">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c901-541">ReadItem</span></span>|
|[<span data-ttu-id="9c901-542">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9c901-542">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9c901-543">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9c901-543">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9c901-544">Пример</span><span class="sxs-lookup"><span data-stu-id="9c901-544">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="9c901-545">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="9c901-545">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="9c901-546">Выполняет асинхронный запрос для веб-служб Exchange (EWS) на сервере Exchange Server, на котором размещен почтовый ящик пользователя.</span><span class="sxs-lookup"><span data-stu-id="9c901-546">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="9c901-547">Этот метод не поддерживается в следующих сценариях:</span><span class="sxs-lookup"><span data-stu-id="9c901-547">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="9c901-548">В Outlook для iOS и Android</span><span class="sxs-lookup"><span data-stu-id="9c901-548">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="9c901-549">Если надстройка загружается в почтовый ящик Gmail.</span><span class="sxs-lookup"><span data-stu-id="9c901-549">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="9c901-550">В таких случаях надстройка должна [использовать REST API](/outlook/add-ins/use-rest-api) для доступа к почтовому ящику пользователя.</span><span class="sxs-lookup"><span data-stu-id="9c901-550">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="9c901-551">Метод `makeEwsRequestAsync` отправляет запрос EWS от имени надстройки в Exchange.</span><span class="sxs-lookup"><span data-stu-id="9c901-551">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="9c901-552">Список поддерживаемых операций EWS см. в статье [Вызов веб-служб из надстройки Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support).</span><span class="sxs-lookup"><span data-stu-id="9c901-552">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="9c901-553">С помощью метода `makeEwsRequestAsync` невозможно запрашивать элементы, связанные с папкой.</span><span class="sxs-lookup"><span data-stu-id="9c901-553">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="9c901-554">В запросе XML должна быть указана кодировка UTF-8.</span><span class="sxs-lookup"><span data-stu-id="9c901-554">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="9c901-p139">У вашей надстройки должно быть разрешение **ReadWriteMailbox** для использования метода `makeEwsRequestAsync`. Сведения об использовании разрешения **ReadWriteMailbox** и операций EWS, которые можно вызывать с помощью метода `makeEwsRequestAsync`, см. в статье [Указание разрешений для доступа почтовой надстройки к почтовому ящику пользователя](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="9c901-p139">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="9c901-557">Администратор сервера должен установить значение true для параметра `OAuthAuthentication` в каталоге сервера клиентского доступа EWS, чтобы метод `makeEwsRequestAsync` мог выполнять запросы EWS.</span><span class="sxs-lookup"><span data-stu-id="9c901-557">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="9c901-558">Различия версий</span><span class="sxs-lookup"><span data-stu-id="9c901-558">Version differences</span></span>

<span data-ttu-id="9c901-559">Если вы используете метод `makeEwsRequestAsync` в почтовых приложениях, которые выполняются в Outlook версии более ранней, чем 15.0.4535.1004, указывайте кодировку `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="9c901-559">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="9c901-p140">Значение кодировки не нужно указывать, если почтовое приложение выполняется в Outlook в Интернете. Чтобы определить, выполняется ли приложение в Outlook или Outlook в Интернете, используйте свойство mailbox.diagnostics.hostName. Используемую версию Outlook можно определить с помощью свойства mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="9c901-p140">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9c901-563">Параметры</span><span class="sxs-lookup"><span data-stu-id="9c901-563">Parameters</span></span>

|<span data-ttu-id="9c901-564">Имя</span><span class="sxs-lookup"><span data-stu-id="9c901-564">Name</span></span>| <span data-ttu-id="9c901-565">Тип</span><span class="sxs-lookup"><span data-stu-id="9c901-565">Type</span></span>| <span data-ttu-id="9c901-566">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="9c901-566">Attributes</span></span>| <span data-ttu-id="9c901-567">Описание</span><span class="sxs-lookup"><span data-stu-id="9c901-567">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="9c901-568">String</span><span class="sxs-lookup"><span data-stu-id="9c901-568">String</span></span>||<span data-ttu-id="9c901-569">Запрос EWS.</span><span class="sxs-lookup"><span data-stu-id="9c901-569">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="9c901-570">function</span><span class="sxs-lookup"><span data-stu-id="9c901-570">function</span></span>||<span data-ttu-id="9c901-571">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9c901-571">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="9c901-572">Результат XML вызова EWS указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="9c901-572">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="9c901-573">Если размер результата превышает 1 МБ, возвращается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="9c901-573">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="9c901-574">Объект</span><span class="sxs-lookup"><span data-stu-id="9c901-574">Object</span></span>| <span data-ttu-id="9c901-575">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9c901-575">&lt;optional&gt;</span></span>|<span data-ttu-id="9c901-576">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="9c901-576">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9c901-577">Требования</span><span class="sxs-lookup"><span data-stu-id="9c901-577">Requirements</span></span>

|<span data-ttu-id="9c901-578">Требование</span><span class="sxs-lookup"><span data-stu-id="9c901-578">Requirement</span></span>| <span data-ttu-id="9c901-579">Значение</span><span class="sxs-lookup"><span data-stu-id="9c901-579">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c901-580">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9c901-580">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c901-581">1.0</span><span class="sxs-lookup"><span data-stu-id="9c901-581">1.0</span></span>|
|[<span data-ttu-id="9c901-582">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9c901-582">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c901-583">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="9c901-583">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="9c901-584">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9c901-584">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9c901-585">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9c901-585">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9c901-586">Пример</span><span class="sxs-lookup"><span data-stu-id="9c901-586">Example</span></span>

<span data-ttu-id="9c901-587">В приведенном ниже примере вызывается `makeEwsRequestAsync` для получения темы элемента с помощью операции `GetItem`.</span><span class="sxs-lookup"><span data-stu-id="9c901-587">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

#### <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="9c901-588">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="9c901-588">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="9c901-589">Удаляет обработчиков для поддерживаемого типа события.</span><span class="sxs-lookup"><span data-stu-id="9c901-589">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="9c901-590">В настоящее время единственный поддерживаемый тип события — `Office.EventType.ItemChanged`.</span><span class="sxs-lookup"><span data-stu-id="9c901-590">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="9c901-591">Параметры</span><span class="sxs-lookup"><span data-stu-id="9c901-591">Parameters</span></span>

| <span data-ttu-id="9c901-592">Имя</span><span class="sxs-lookup"><span data-stu-id="9c901-592">Name</span></span> | <span data-ttu-id="9c901-593">Тип</span><span class="sxs-lookup"><span data-stu-id="9c901-593">Type</span></span> | <span data-ttu-id="9c901-594">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="9c901-594">Attributes</span></span> | <span data-ttu-id="9c901-595">Описание</span><span class="sxs-lookup"><span data-stu-id="9c901-595">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="9c901-596">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="9c901-596">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="9c901-597">Событие, которое должно отменить обработчик.</span><span class="sxs-lookup"><span data-stu-id="9c901-597">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="9c901-598">Объект</span><span class="sxs-lookup"><span data-stu-id="9c901-598">Object</span></span> | <span data-ttu-id="9c901-599">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9c901-599">&lt;optional&gt;</span></span> | <span data-ttu-id="9c901-600">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="9c901-600">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="9c901-601">Object</span><span class="sxs-lookup"><span data-stu-id="9c901-601">Object</span></span> | <span data-ttu-id="9c901-602">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9c901-602">&lt;optional&gt;</span></span> | <span data-ttu-id="9c901-603">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="9c901-603">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="9c901-604">функция</span><span class="sxs-lookup"><span data-stu-id="9c901-604">function</span></span>| <span data-ttu-id="9c901-605">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="9c901-605">&lt;optional&gt;</span></span>|<span data-ttu-id="9c901-606">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="9c901-606">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9c901-607">Требования</span><span class="sxs-lookup"><span data-stu-id="9c901-607">Requirements</span></span>

|<span data-ttu-id="9c901-608">Требование</span><span class="sxs-lookup"><span data-stu-id="9c901-608">Requirement</span></span>| <span data-ttu-id="9c901-609">Значение</span><span class="sxs-lookup"><span data-stu-id="9c901-609">Value</span></span>|
|---|---|
|[<span data-ttu-id="9c901-610">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9c901-610">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9c901-611">1.5</span><span class="sxs-lookup"><span data-stu-id="9c901-611">1.5</span></span> |
|[<span data-ttu-id="9c901-612">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9c901-612">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="9c901-613">ReadItem</span><span class="sxs-lookup"><span data-stu-id="9c901-613">ReadItem</span></span> |
|[<span data-ttu-id="9c901-614">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9c901-614">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="9c901-615">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9c901-615">Compose or Read</span></span>|
