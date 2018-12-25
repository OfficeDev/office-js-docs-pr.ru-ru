---
title: Office.context.mailbox — набор обязательных элементов 1.7
description: ''
ms.date: 10/31/2018
ms.openlocfilehash: 3f4cee427e49a691b42d47d6821c54cd428207d9
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433343"
---
# <a name="mailbox"></a><span data-ttu-id="27126-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="27126-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="27126-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="27126-103">Office.context.mailbox</span></span>

<span data-ttu-id="27126-104">Предоставляет для Microsoft Outlook и Microsoft Outlook в Интернете доступ к объектной модели надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="27126-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="27126-105">Требования</span><span class="sxs-lookup"><span data-stu-id="27126-105">Requirements</span></span>

|<span data-ttu-id="27126-106">Требование</span><span class="sxs-lookup"><span data-stu-id="27126-106">Requirement</span></span>| <span data-ttu-id="27126-107">Значение</span><span class="sxs-lookup"><span data-stu-id="27126-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="27126-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="27126-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="27126-109">1.0</span><span class="sxs-lookup"><span data-stu-id="27126-109">1.0</span></span>|
|[<span data-ttu-id="27126-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="27126-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="27126-111">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="27126-111">Restricted</span></span>|
|[<span data-ttu-id="27126-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="27126-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="27126-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="27126-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="27126-114">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="27126-114">Members and methods</span></span>

| <span data-ttu-id="27126-115">Элемент</span><span class="sxs-lookup"><span data-stu-id="27126-115">Member</span></span> | <span data-ttu-id="27126-116">Тип</span><span class="sxs-lookup"><span data-stu-id="27126-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="27126-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="27126-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="27126-118">Элемент</span><span class="sxs-lookup"><span data-stu-id="27126-118">Member</span></span> |
| [<span data-ttu-id="27126-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="27126-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="27126-120">Элемент</span><span class="sxs-lookup"><span data-stu-id="27126-120">Member</span></span> |
| [<span data-ttu-id="27126-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="27126-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="27126-122">Метод</span><span class="sxs-lookup"><span data-stu-id="27126-122">Method</span></span> |
| [<span data-ttu-id="27126-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="27126-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="27126-124">Метод</span><span class="sxs-lookup"><span data-stu-id="27126-124">Method</span></span> |
| [<span data-ttu-id="27126-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="27126-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime) | <span data-ttu-id="27126-126">Метод</span><span class="sxs-lookup"><span data-stu-id="27126-126">Method</span></span> |
| [<span data-ttu-id="27126-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="27126-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="27126-128">Метод</span><span class="sxs-lookup"><span data-stu-id="27126-128">Method</span></span> |
| [<span data-ttu-id="27126-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="27126-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="27126-130">Метод</span><span class="sxs-lookup"><span data-stu-id="27126-130">Method</span></span> |
| [<span data-ttu-id="27126-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="27126-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="27126-132">Метод</span><span class="sxs-lookup"><span data-stu-id="27126-132">Method</span></span> |
| [<span data-ttu-id="27126-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="27126-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="27126-134">Метод</span><span class="sxs-lookup"><span data-stu-id="27126-134">Method</span></span> |
| [<span data-ttu-id="27126-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="27126-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="27126-136">Метод</span><span class="sxs-lookup"><span data-stu-id="27126-136">Method</span></span> |
| [<span data-ttu-id="27126-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="27126-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="27126-138">Метод</span><span class="sxs-lookup"><span data-stu-id="27126-138">Method</span></span> |
| [<span data-ttu-id="27126-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="27126-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="27126-140">Метод</span><span class="sxs-lookup"><span data-stu-id="27126-140">Method</span></span> |
| [<span data-ttu-id="27126-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="27126-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="27126-142">Метод</span><span class="sxs-lookup"><span data-stu-id="27126-142">Method</span></span> |
| [<span data-ttu-id="27126-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="27126-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="27126-144">Метод</span><span class="sxs-lookup"><span data-stu-id="27126-144">Method</span></span> |
| [<span data-ttu-id="27126-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="27126-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="27126-146">Метод</span><span class="sxs-lookup"><span data-stu-id="27126-146">Method</span></span> |
| [<span data-ttu-id="27126-147">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="27126-147">removeHandlerAsync</span></span>](#removehandlerasynceventtype-handler-options-callback) | <span data-ttu-id="27126-148">Метод</span><span class="sxs-lookup"><span data-stu-id="27126-148">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="27126-149">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="27126-149">Namespaces</span></span>

<span data-ttu-id="27126-150">[diagnostics](Office.context.mailbox.diagnostics.md). Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="27126-150">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="27126-151">[item](Office.context.mailbox.item.md). Предоставляет методы и свойства для доступа к сообщению или встрече в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="27126-151">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="27126-152">[userProfile](Office.context.mailbox.userProfile.md). Предоставляет сведения о пользователе в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="27126-152">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="27126-153">Элементы</span><span class="sxs-lookup"><span data-stu-id="27126-153">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="27126-154">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="27126-154">ewsUrl :String</span></span>

<span data-ttu-id="27126-p101">Получает URL-адрес конечной точки веб-служб Exchange (EWS) для этой учетной записи электронной почты. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="27126-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="27126-157">Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="27126-157">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="27126-p102">Удаленная служба может использовать значение `ewsUrl`, чтобы выполнять вызовы EWS для почтового ящика пользователя. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="27126-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="27126-160">Чтобы вызвать элемент `ewsUrl` в режиме чтения, в манифесте приложения должно быть указано разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="27126-160">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="27126-p103">Перед использованием элемента `ewsUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="27126-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="27126-163">Тип:</span><span class="sxs-lookup"><span data-stu-id="27126-163">Type:</span></span>

*   <span data-ttu-id="27126-164">String</span><span class="sxs-lookup"><span data-stu-id="27126-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="27126-165">Требования</span><span class="sxs-lookup"><span data-stu-id="27126-165">Requirements</span></span>

|<span data-ttu-id="27126-166">Требование</span><span class="sxs-lookup"><span data-stu-id="27126-166">Requirement</span></span>| <span data-ttu-id="27126-167">Значение</span><span class="sxs-lookup"><span data-stu-id="27126-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="27126-168">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="27126-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="27126-169">1.0</span><span class="sxs-lookup"><span data-stu-id="27126-169">1.0</span></span>|
|[<span data-ttu-id="27126-170">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="27126-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="27126-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27126-171">ReadItem</span></span>|
|[<span data-ttu-id="27126-172">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="27126-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="27126-173">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="27126-173">Compose or read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="27126-174">restUrl :String</span><span class="sxs-lookup"><span data-stu-id="27126-174">restUrl :String</span></span>

<span data-ttu-id="27126-175">Возвращает URL-адрес конечной точки REST для этой учетной записи электронной почты.</span><span class="sxs-lookup"><span data-stu-id="27126-175">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="27126-176">С помощью значения `restUrl` можно выполнять вызовы [REST API](https://docs.microsoft.com/outlook/rest/) для почтового ящика пользователя.</span><span class="sxs-lookup"><span data-stu-id="27126-176">The `restUrl` value can be used to make [REST API](https://docs.microsoft.com/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="27126-177">Чтобы вызвать элемент `restUrl` в режиме чтения, в манифесте приложения необходимо указать разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="27126-177">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="27126-p104">Перед использованием элемента `restUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="27126-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="27126-180">Тип:</span><span class="sxs-lookup"><span data-stu-id="27126-180">Type:</span></span>

*   <span data-ttu-id="27126-181">String</span><span class="sxs-lookup"><span data-stu-id="27126-181">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="27126-182">Требования</span><span class="sxs-lookup"><span data-stu-id="27126-182">Requirements</span></span>

|<span data-ttu-id="27126-183">Требование</span><span class="sxs-lookup"><span data-stu-id="27126-183">Requirement</span></span>| <span data-ttu-id="27126-184">Значение</span><span class="sxs-lookup"><span data-stu-id="27126-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="27126-185">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="27126-185">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="27126-186">1.5</span><span class="sxs-lookup"><span data-stu-id="27126-186">1.5</span></span> |
|[<span data-ttu-id="27126-187">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="27126-187">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="27126-188">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27126-188">ReadItem</span></span>|
|[<span data-ttu-id="27126-189">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="27126-189">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="27126-190">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="27126-190">Compose or read</span></span>|

### <a name="methods"></a><span data-ttu-id="27126-191">Методы</span><span class="sxs-lookup"><span data-stu-id="27126-191">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="27126-192">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="27126-192">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="27126-193">Добавляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="27126-193">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="27126-194">В настоящее время единственный поддерживаемый тип события — `Office.EventType.ItemChanged`.</span><span class="sxs-lookup"><span data-stu-id="27126-194">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="27126-195">Параметры:</span><span class="sxs-lookup"><span data-stu-id="27126-195">Parameters:</span></span>

| <span data-ttu-id="27126-196">Имя</span><span class="sxs-lookup"><span data-stu-id="27126-196">Name</span></span> | <span data-ttu-id="27126-197">Тип</span><span class="sxs-lookup"><span data-stu-id="27126-197">Type</span></span> | <span data-ttu-id="27126-198">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="27126-198">Attributes</span></span> | <span data-ttu-id="27126-199">Описание</span><span class="sxs-lookup"><span data-stu-id="27126-199">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="27126-200">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="27126-200">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="27126-201">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="27126-201">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="27126-202">Function</span><span class="sxs-lookup"><span data-stu-id="27126-202">Function</span></span> || <span data-ttu-id="27126-p105">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="27126-p105">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="27126-206">Объект</span><span class="sxs-lookup"><span data-stu-id="27126-206">Object</span></span> | <span data-ttu-id="27126-207">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="27126-207">&lt;optional&gt;</span></span> | <span data-ttu-id="27126-208">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="27126-208">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="27126-209">Object</span><span class="sxs-lookup"><span data-stu-id="27126-209">Object</span></span> | <span data-ttu-id="27126-210">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="27126-210">&lt;optional&gt;</span></span> | <span data-ttu-id="27126-211">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="27126-211">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="27126-212">функция</span><span class="sxs-lookup"><span data-stu-id="27126-212">function</span></span>| <span data-ttu-id="27126-213">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="27126-213">&lt;optional&gt;</span></span>|<span data-ttu-id="27126-214">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="27126-214">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="27126-215">Требования</span><span class="sxs-lookup"><span data-stu-id="27126-215">Requirements</span></span>

|<span data-ttu-id="27126-216">Требование</span><span class="sxs-lookup"><span data-stu-id="27126-216">Requirement</span></span>| <span data-ttu-id="27126-217">Значение</span><span class="sxs-lookup"><span data-stu-id="27126-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="27126-218">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="27126-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="27126-219">1.5</span><span class="sxs-lookup"><span data-stu-id="27126-219">1.5</span></span> |
|[<span data-ttu-id="27126-220">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="27126-220">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="27126-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27126-221">ReadItem</span></span> |
|[<span data-ttu-id="27126-222">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="27126-222">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="27126-223">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="27126-223">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="27126-224">Пример</span><span class="sxs-lookup"><span data-stu-id="27126-224">Example</span></span>

```js
Office.initialize = function (reason) {
  $(document).ready(function () {
    Office.context.mailbox.addHandlerAsync(Office.EventType.ItemChanged, loadNewItem, function (result) {
      if (result.status === Office.AsyncResultStatus.Failed) {
        // Handle error
      }
    });
  });
};

function loadNewItem(eventArgs) {
  // Load the properties of the newly selected item
  loadProps(Office.context.mailbox.item);
};
```

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="27126-225">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="27126-225">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="27126-226">Преобразовывает идентификатор элемента из формата REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="27126-226">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="27126-227">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="27126-227">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="27126-p106">Формат идентификаторов, извлекаемых через API REST (например, [API Почты Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)), отличается от формата веб-служб Exchange (EWS). Метод `convertToEwsId` преобразовывает идентификатор в формате REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="27126-p106">Item IDs retrieved via a REST API (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="27126-230">Параметры</span><span class="sxs-lookup"><span data-stu-id="27126-230">Parameters:</span></span>

|<span data-ttu-id="27126-231">Имя</span><span class="sxs-lookup"><span data-stu-id="27126-231">Name</span></span>| <span data-ttu-id="27126-232">Тип</span><span class="sxs-lookup"><span data-stu-id="27126-232">Type</span></span>| <span data-ttu-id="27126-233">Описание</span><span class="sxs-lookup"><span data-stu-id="27126-233">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="27126-234">String</span><span class="sxs-lookup"><span data-stu-id="27126-234">String</span></span>|<span data-ttu-id="27126-235">Идентификатор элемента в формате REST API для Outlook</span><span class="sxs-lookup"><span data-stu-id="27126-235">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="27126-236">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="27126-236">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.restversion)|<span data-ttu-id="27126-237">Значение, определяющее версию REST API для Outlook, которая используется для извлечения идентификатора элемента.</span><span class="sxs-lookup"><span data-stu-id="27126-237">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="27126-238">Требования</span><span class="sxs-lookup"><span data-stu-id="27126-238">Requirements</span></span>

|<span data-ttu-id="27126-239">Требование</span><span class="sxs-lookup"><span data-stu-id="27126-239">Requirement</span></span>| <span data-ttu-id="27126-240">Значение</span><span class="sxs-lookup"><span data-stu-id="27126-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="27126-241">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="27126-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="27126-242">1.3</span><span class="sxs-lookup"><span data-stu-id="27126-242">1.3</span></span>|
|[<span data-ttu-id="27126-243">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="27126-243">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="27126-244">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="27126-244">Restricted</span></span>|
|[<span data-ttu-id="27126-245">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="27126-245">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="27126-246">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="27126-246">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="27126-247">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="27126-247">Returns:</span></span>

<span data-ttu-id="27126-248">Тип: String</span><span class="sxs-lookup"><span data-stu-id="27126-248">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="27126-249">Пример</span><span class="sxs-lookup"><span data-stu-id="27126-249">Example</span></span>

```js
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook17officelocalclienttime"></a><span data-ttu-id="27126-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="27126-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)}</span></span>

<span data-ttu-id="27126-251">Получает словарь, содержащий сведения о локальном времени клиента.</span><span class="sxs-lookup"><span data-stu-id="27126-251">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="27126-p107">В случае дат и времени в почтовом приложении для Outlook или Outlook Web App могут использоваться разные часовые пояса. Outlook использует часовой пояс клиентского компьютера. Outlook Web App использует часовой пояс, заданный в Центре администрирования Exchange (EAC). Значения даты и времени должны обрабатываться так, чтобы значения в пользовательском интерфейсе всегда согласовывались с часовым поясом, ожидаемым пользователем.</span><span class="sxs-lookup"><span data-stu-id="27126-p107">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="27126-p108">Если почтовое приложение работает в Outlook, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса клиентского компьютера. Если почтовое приложение работает в Outlook Web App, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса, заданного в Центре администрирования Exchange.</span><span class="sxs-lookup"><span data-stu-id="27126-p108">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="27126-257">Параметры</span><span class="sxs-lookup"><span data-stu-id="27126-257">Parameters:</span></span>

|<span data-ttu-id="27126-258">Имя</span><span class="sxs-lookup"><span data-stu-id="27126-258">Name</span></span>| <span data-ttu-id="27126-259">Тип</span><span class="sxs-lookup"><span data-stu-id="27126-259">Type</span></span>| <span data-ttu-id="27126-260">Описание</span><span class="sxs-lookup"><span data-stu-id="27126-260">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="27126-261">Date</span><span class="sxs-lookup"><span data-stu-id="27126-261">Date</span></span>|<span data-ttu-id="27126-262">Объект Date</span><span class="sxs-lookup"><span data-stu-id="27126-262">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="27126-263">Требования</span><span class="sxs-lookup"><span data-stu-id="27126-263">Requirements</span></span>

|<span data-ttu-id="27126-264">Требование</span><span class="sxs-lookup"><span data-stu-id="27126-264">Requirement</span></span>| <span data-ttu-id="27126-265">Значение</span><span class="sxs-lookup"><span data-stu-id="27126-265">Value</span></span>|
|---|---|
|[<span data-ttu-id="27126-266">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="27126-266">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="27126-267">1.0</span><span class="sxs-lookup"><span data-stu-id="27126-267">1.0</span></span>|
|[<span data-ttu-id="27126-268">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="27126-268">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="27126-269">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27126-269">ReadItem</span></span>|
|[<span data-ttu-id="27126-270">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="27126-270">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="27126-271">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="27126-271">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="27126-272">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="27126-272">Returns:</span></span>

<span data-ttu-id="27126-273">Тип: [LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="27126-273">Type: [LocalClientTime](/javascript/api/outlook_1_7/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="27126-274">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="27126-274">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="27126-275">Преобразовывает идентификатор элемента в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="27126-275">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="27126-276">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="27126-276">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="27126-p109">Формат идентификаторов, извлекаемых через EWS или свойство `itemId`, отличается от формата API REST (таких как [API Почты Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)). Метод `convertToRestId` преобразовывает идентификатор в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="27126-p109">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="27126-279">Параметры</span><span class="sxs-lookup"><span data-stu-id="27126-279">Parameters:</span></span>

|<span data-ttu-id="27126-280">Имя</span><span class="sxs-lookup"><span data-stu-id="27126-280">Name</span></span>| <span data-ttu-id="27126-281">Тип</span><span class="sxs-lookup"><span data-stu-id="27126-281">Type</span></span>| <span data-ttu-id="27126-282">Описание</span><span class="sxs-lookup"><span data-stu-id="27126-282">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="27126-283">String</span><span class="sxs-lookup"><span data-stu-id="27126-283">String</span></span>|<span data-ttu-id="27126-284">Идентификатор элемента в формате EWS</span><span class="sxs-lookup"><span data-stu-id="27126-284">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="27126-285">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="27126-285">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_7/office.mailboxenums.restversion)|<span data-ttu-id="27126-286">Значение, определяющее версию REST API для Outlook, с которой будет использоваться преобразованный идентификатор.</span><span class="sxs-lookup"><span data-stu-id="27126-286">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="27126-287">Требования</span><span class="sxs-lookup"><span data-stu-id="27126-287">Requirements</span></span>

|<span data-ttu-id="27126-288">Требование</span><span class="sxs-lookup"><span data-stu-id="27126-288">Requirement</span></span>| <span data-ttu-id="27126-289">Значение</span><span class="sxs-lookup"><span data-stu-id="27126-289">Value</span></span>|
|---|---|
|[<span data-ttu-id="27126-290">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="27126-290">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="27126-291">1.3</span><span class="sxs-lookup"><span data-stu-id="27126-291">1.3</span></span>|
|[<span data-ttu-id="27126-292">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="27126-292">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="27126-293">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="27126-293">Restricted</span></span>|
|[<span data-ttu-id="27126-294">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="27126-294">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="27126-295">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="27126-295">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="27126-296">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="27126-296">Returns:</span></span>

<span data-ttu-id="27126-297">Тип: String</span><span class="sxs-lookup"><span data-stu-id="27126-297">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="27126-298">Пример</span><span class="sxs-lookup"><span data-stu-id="27126-298">Example</span></span>

```js
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="27126-299">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="27126-299">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="27126-300">Получает объект Date из словаря, содержащего сведения о времени.</span><span class="sxs-lookup"><span data-stu-id="27126-300">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="27126-301">Метод `convertToUtcClientTime` преобразует словарь, содержащий локальную дату и время, в объект Date с правильными значениями локальной даты и времени.</span><span class="sxs-lookup"><span data-stu-id="27126-301">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="27126-302">Параметры</span><span class="sxs-lookup"><span data-stu-id="27126-302">Parameters:</span></span>

|<span data-ttu-id="27126-303">Имя</span><span class="sxs-lookup"><span data-stu-id="27126-303">Name</span></span>| <span data-ttu-id="27126-304">Тип</span><span class="sxs-lookup"><span data-stu-id="27126-304">Type</span></span>| <span data-ttu-id="27126-305">Описание</span><span class="sxs-lookup"><span data-stu-id="27126-305">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="27126-306">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="27126-306">LocalClientTime</span></span>](/javascript/api/outlook_1_7/office.LocalClientTime)|<span data-ttu-id="27126-307">Значение локального времени для преобразования.</span><span class="sxs-lookup"><span data-stu-id="27126-307">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="27126-308">Требования</span><span class="sxs-lookup"><span data-stu-id="27126-308">Requirements</span></span>

|<span data-ttu-id="27126-309">Требование</span><span class="sxs-lookup"><span data-stu-id="27126-309">Requirement</span></span>| <span data-ttu-id="27126-310">Значение</span><span class="sxs-lookup"><span data-stu-id="27126-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="27126-311">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="27126-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="27126-312">1.0</span><span class="sxs-lookup"><span data-stu-id="27126-312">1.0</span></span>|
|[<span data-ttu-id="27126-313">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="27126-313">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="27126-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27126-314">ReadItem</span></span>|
|[<span data-ttu-id="27126-315">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="27126-315">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="27126-316">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="27126-316">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="27126-317">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="27126-317">Returns:</span></span>

<span data-ttu-id="27126-318">Объект Date со временем в формате UTC.</span><span class="sxs-lookup"><span data-stu-id="27126-318">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="27126-319">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="27126-319">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="27126-320">Date</span><span class="sxs-lookup"><span data-stu-id="27126-320">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="27126-321">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="27126-321">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="27126-322">Отображает имеющуюся встречу из календаря.</span><span class="sxs-lookup"><span data-stu-id="27126-322">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="27126-323">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="27126-323">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="27126-324">Метод `displayAppointmentForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее сведения календаря о существующей встрече.</span><span class="sxs-lookup"><span data-stu-id="27126-324">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="27126-p110">В Outlook для Mac с помощью этого метода можно отобразить одну встречу, которая не является частью повторяющегося ряда, или основную встречу такого ряда, но не экземпляр из него, так как в Outlook для Mac невозможно получить доступ к свойствам экземпляра повторяющегося ряда (в том числе к идентификатору элемента).</span><span class="sxs-lookup"><span data-stu-id="27126-p110">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="27126-327">В Outlook Web App этот метод открывает указанную форму, только если текст формы содержит символы размером не более 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="27126-327">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="27126-328">Если указанный идентификатор элемента не определяет существующую встречу, на клиентском компьютере или устройстве открывается пустая страница, и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="27126-328">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="27126-329">Параметры</span><span class="sxs-lookup"><span data-stu-id="27126-329">Parameters:</span></span>

|<span data-ttu-id="27126-330">Имя</span><span class="sxs-lookup"><span data-stu-id="27126-330">Name</span></span>| <span data-ttu-id="27126-331">Тип</span><span class="sxs-lookup"><span data-stu-id="27126-331">Type</span></span>| <span data-ttu-id="27126-332">Описание</span><span class="sxs-lookup"><span data-stu-id="27126-332">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="27126-333">String</span><span class="sxs-lookup"><span data-stu-id="27126-333">String</span></span>|<span data-ttu-id="27126-334">Идентификатор веб-служб Exchange для существующей встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="27126-334">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="27126-335">Требования</span><span class="sxs-lookup"><span data-stu-id="27126-335">Requirements</span></span>

|<span data-ttu-id="27126-336">Требование</span><span class="sxs-lookup"><span data-stu-id="27126-336">Requirement</span></span>| <span data-ttu-id="27126-337">Значение</span><span class="sxs-lookup"><span data-stu-id="27126-337">Value</span></span>|
|---|---|
|[<span data-ttu-id="27126-338">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="27126-338">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="27126-339">1.0</span><span class="sxs-lookup"><span data-stu-id="27126-339">1.0</span></span>|
|[<span data-ttu-id="27126-340">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="27126-340">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="27126-341">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27126-341">ReadItem</span></span>|
|[<span data-ttu-id="27126-342">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="27126-342">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="27126-343">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="27126-343">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="27126-344">Пример</span><span class="sxs-lookup"><span data-stu-id="27126-344">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="27126-345">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="27126-345">displayMessageForm(itemId)</span></span>

<span data-ttu-id="27126-346">Отображает имеющееся сообщение.</span><span class="sxs-lookup"><span data-stu-id="27126-346">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="27126-347">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="27126-347">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="27126-348">Метод `displayMessageForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее существующее сообщение.</span><span class="sxs-lookup"><span data-stu-id="27126-348">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="27126-349">В Outlook Web App этот метод открывает указанную форму, только если текст формы содержит символы размером не более 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="27126-349">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="27126-350">Если указанный идентификатор элемента не определяет существующее сообщение, окно на клиентском компьютере не открывается и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="27126-350">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="27126-p111">Не используйте `displayMessageForm` с параметром `itemId`, который представляет собой встречу. Используйте метод `displayAppointmentForm`, чтобы отобразить сведения о существующей встрече, а метод `displayNewAppointmentForm` — для отображения формы создания встречи.</span><span class="sxs-lookup"><span data-stu-id="27126-p111">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="27126-353">Параметры</span><span class="sxs-lookup"><span data-stu-id="27126-353">Parameters:</span></span>

|<span data-ttu-id="27126-354">Имя</span><span class="sxs-lookup"><span data-stu-id="27126-354">Name</span></span>| <span data-ttu-id="27126-355">Тип</span><span class="sxs-lookup"><span data-stu-id="27126-355">Type</span></span>| <span data-ttu-id="27126-356">Описание</span><span class="sxs-lookup"><span data-stu-id="27126-356">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="27126-357">String</span><span class="sxs-lookup"><span data-stu-id="27126-357">String</span></span>|<span data-ttu-id="27126-358">Идентификатор веб-служб Exchange для существующего сообщения.</span><span class="sxs-lookup"><span data-stu-id="27126-358">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="27126-359">Требования</span><span class="sxs-lookup"><span data-stu-id="27126-359">Requirements</span></span>

|<span data-ttu-id="27126-360">Требование</span><span class="sxs-lookup"><span data-stu-id="27126-360">Requirement</span></span>| <span data-ttu-id="27126-361">Значение</span><span class="sxs-lookup"><span data-stu-id="27126-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="27126-362">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="27126-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="27126-363">1.0</span><span class="sxs-lookup"><span data-stu-id="27126-363">1.0</span></span>|
|[<span data-ttu-id="27126-364">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="27126-364">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="27126-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27126-365">ReadItem</span></span>|
|[<span data-ttu-id="27126-366">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="27126-366">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="27126-367">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="27126-367">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="27126-368">Пример</span><span class="sxs-lookup"><span data-stu-id="27126-368">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="27126-369">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="27126-369">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="27126-370">Отображает форму для создания новой встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="27126-370">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="27126-371">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="27126-371">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="27126-p112">Метод `displayNewAppointmentForm` открывает форму, в которой пользователь может создать встречу или собрание. Если параметры заданы, поля формы встречи автоматически заполняются их содержимым.</span><span class="sxs-lookup"><span data-stu-id="27126-p112">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="27126-p113">В Outlook Web App и Outlook Web App для устройств этот метод всегда отображает форму с полем участников. Если вы не укажете участников в качестве входных аргументов, метод отображает форму с кнопкой **Сохранить**. Если вы укажете участников, форма будет включать участников и кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="27126-p113">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="27126-p114">Если вы укажете участников или ресурсы с помощью параметра `requiredAttendees`, `optionalAttendees` или `resources` в клиенте Outlook с расширенными возможностями и Outlook RT, этот метод отобразит форму собрания с кнопкой **Отправить**. Если не указать получателей, этот метод отобразит форму встречи с кнопкой **Сохранить и закрыть**.</span><span class="sxs-lookup"><span data-stu-id="27126-p114">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="27126-379">Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.</span><span class="sxs-lookup"><span data-stu-id="27126-379">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="27126-380">Параметры:</span><span class="sxs-lookup"><span data-stu-id="27126-380">Parameters:</span></span>

> [!NOTE]
> <span data-ttu-id="27126-381">Все параметры являются необязательными.</span><span class="sxs-lookup"><span data-stu-id="27126-381">All parameters are optional.</span></span>

|<span data-ttu-id="27126-382">Имя</span><span class="sxs-lookup"><span data-stu-id="27126-382">Name</span></span>| <span data-ttu-id="27126-383">Тип</span><span class="sxs-lookup"><span data-stu-id="27126-383">Type</span></span>| <span data-ttu-id="27126-384">Описание</span><span class="sxs-lookup"><span data-stu-id="27126-384">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="27126-385">Object</span><span class="sxs-lookup"><span data-stu-id="27126-385">Object</span></span> | <span data-ttu-id="27126-386">Словарь параметров, описывающий новую встречу.</span><span class="sxs-lookup"><span data-stu-id="27126-386">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="27126-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="27126-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="27126-p115">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из обязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="27126-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="27126-390">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="27126-390">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="27126-p116">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из необязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="27126-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="27126-393">Date</span><span class="sxs-lookup"><span data-stu-id="27126-393">Date</span></span> | <span data-ttu-id="27126-394">Объект `Date`, указывающий дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="27126-394">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="27126-395">Date</span><span class="sxs-lookup"><span data-stu-id="27126-395">Date</span></span> | <span data-ttu-id="27126-396">Объект `Date`, указывающий дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="27126-396">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="27126-397">String</span><span class="sxs-lookup"><span data-stu-id="27126-397">String</span></span> | <span data-ttu-id="27126-p117">Строка со сведениями о месте встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="27126-p117">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="27126-400">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="27126-400">Array.&lt;String&gt;</span></span> | <span data-ttu-id="27126-p118">Массив строк, содержащий необходимые для встречи ресурсы. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="27126-p118">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="27126-403">String</span><span class="sxs-lookup"><span data-stu-id="27126-403">String</span></span> | <span data-ttu-id="27126-p119">Строка с темой встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="27126-p119">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="27126-406">String</span><span class="sxs-lookup"><span data-stu-id="27126-406">String</span></span> | <span data-ttu-id="27126-p120">Текст сообщения о встрече. Максимальный размер содержимого сообщения — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="27126-p120">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="27126-409">Требования</span><span class="sxs-lookup"><span data-stu-id="27126-409">Requirements</span></span>

|<span data-ttu-id="27126-410">Требование</span><span class="sxs-lookup"><span data-stu-id="27126-410">Requirement</span></span>| <span data-ttu-id="27126-411">Значение</span><span class="sxs-lookup"><span data-stu-id="27126-411">Value</span></span>|
|---|---|
|[<span data-ttu-id="27126-412">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="27126-412">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="27126-413">1.0</span><span class="sxs-lookup"><span data-stu-id="27126-413">1.0</span></span>|
|[<span data-ttu-id="27126-414">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="27126-414">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="27126-415">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27126-415">ReadItem</span></span>|
|[<span data-ttu-id="27126-416">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="27126-416">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="27126-417">Чтение</span><span class="sxs-lookup"><span data-stu-id="27126-417">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="27126-418">Пример</span><span class="sxs-lookup"><span data-stu-id="27126-418">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="27126-419">displayNewMessageForm(параметры)</span><span class="sxs-lookup"><span data-stu-id="27126-419">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="27126-420">Отображает форму для создания сообщения.</span><span class="sxs-lookup"><span data-stu-id="27126-420">Displays a form for creating a new message.</span></span>

<span data-ttu-id="27126-421">Метод `displayNewMessageForm` открывает форму, в которой пользователь может создать сообщение.</span><span class="sxs-lookup"><span data-stu-id="27126-421">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="27126-422">Если параметры заданы, поля формы сообщения автоматически заполняются их содержимым.</span><span class="sxs-lookup"><span data-stu-id="27126-422">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="27126-423">Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.</span><span class="sxs-lookup"><span data-stu-id="27126-423">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="27126-424">Параметры:</span><span class="sxs-lookup"><span data-stu-id="27126-424">Parameters:</span></span>

> [!NOTE]
> <span data-ttu-id="27126-425">Все параметры являются необязательными.</span><span class="sxs-lookup"><span data-stu-id="27126-425">All parameters are optional.</span></span>

|<span data-ttu-id="27126-426">Имя</span><span class="sxs-lookup"><span data-stu-id="27126-426">Name</span></span>| <span data-ttu-id="27126-427">Тип</span><span class="sxs-lookup"><span data-stu-id="27126-427">Type</span></span>| <span data-ttu-id="27126-428">Описание</span><span class="sxs-lookup"><span data-stu-id="27126-428">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="27126-429">Объект</span><span class="sxs-lookup"><span data-stu-id="27126-429">Object</span></span> | <span data-ttu-id="27126-430">Словарь параметров, описывающий новое сообщение.</span><span class="sxs-lookup"><span data-stu-id="27126-430">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="27126-431">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="27126-431">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="27126-432">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из получателей, указанных в строке "Кому".</span><span class="sxs-lookup"><span data-stu-id="27126-432">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="27126-433">Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="27126-433">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="27126-434">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="27126-434">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="27126-435">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из получателей, указанных в строке "Копия".</span><span class="sxs-lookup"><span data-stu-id="27126-435">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="27126-436">Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="27126-436">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="27126-437">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="27126-437">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_7/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="27126-438">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из получателей, указанных в строке "Скрытая копия".</span><span class="sxs-lookup"><span data-stu-id="27126-438">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="27126-439">Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="27126-439">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="27126-440">String</span><span class="sxs-lookup"><span data-stu-id="27126-440">String</span></span> | <span data-ttu-id="27126-441">Строка с темой сообщения.</span><span class="sxs-lookup"><span data-stu-id="27126-441">A string containing the subject of the message.</span></span> <span data-ttu-id="27126-442">Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="27126-442">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="27126-443">String</span><span class="sxs-lookup"><span data-stu-id="27126-443">String</span></span> | <span data-ttu-id="27126-444">Текст сообщения в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="27126-444">The HTML body of the message.</span></span> <span data-ttu-id="27126-445">Максимальный размер содержимого сообщения — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="27126-445">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="27126-446">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="27126-446">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="27126-447">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="27126-447">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="27126-448">String</span><span class="sxs-lookup"><span data-stu-id="27126-448">String</span></span> | <span data-ttu-id="27126-p127">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="27126-p127">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="27126-451">String</span><span class="sxs-lookup"><span data-stu-id="27126-451">String</span></span> | <span data-ttu-id="27126-452">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="27126-452">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="27126-453">String</span><span class="sxs-lookup"><span data-stu-id="27126-453">String</span></span> | <span data-ttu-id="27126-p128">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="27126-p128">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="27126-456">Логический</span><span class="sxs-lookup"><span data-stu-id="27126-456">Boolean</span></span> | <span data-ttu-id="27126-p129">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="27126-p129">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="27126-459">String</span><span class="sxs-lookup"><span data-stu-id="27126-459">String</span></span> | <span data-ttu-id="27126-460">Используется, только если для `type` задано значение `item`.</span><span class="sxs-lookup"><span data-stu-id="27126-460">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="27126-461">Идентификатор элемента EWS существующего письма, которое нужно вложить в новое сообщение.</span><span class="sxs-lookup"><span data-stu-id="27126-461">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="27126-462">Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="27126-462">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="27126-463">Требования</span><span class="sxs-lookup"><span data-stu-id="27126-463">Requirements</span></span>

|<span data-ttu-id="27126-464">Требование</span><span class="sxs-lookup"><span data-stu-id="27126-464">Requirement</span></span>| <span data-ttu-id="27126-465">Значение</span><span class="sxs-lookup"><span data-stu-id="27126-465">Value</span></span>|
|---|---|
|[<span data-ttu-id="27126-466">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="27126-466">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="27126-467">1.6</span><span class="sxs-lookup"><span data-stu-id="27126-467">1.6</span></span> |
|[<span data-ttu-id="27126-468">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="27126-468">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="27126-469">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27126-469">ReadItem</span></span>|
|[<span data-ttu-id="27126-470">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="27126-470">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="27126-471">Чтение</span><span class="sxs-lookup"><span data-stu-id="27126-471">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="27126-472">Пример</span><span class="sxs-lookup"><span data-stu-id="27126-472">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="27126-473">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="27126-473">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="27126-474">Возвращает строку, содержащую маркер, который используется для вызова интерфейсов REST API или веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="27126-474">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="27126-p131">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="27126-p131">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="27126-477">Рекомендуем сделать так, чтобы по мере возможности надстройки использовали интерфейсы REST API, а не веб-службы Exchange.</span><span class="sxs-lookup"><span data-stu-id="27126-477">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="27126-478">**Маркеры REST**</span><span class="sxs-lookup"><span data-stu-id="27126-478">**REST Tokens**</span></span>

<span data-ttu-id="27126-p132">Если запрашивается маркер REST (`options.isRest = true`), полученный маркер не подойдет для проверки подлинности при вызовах веб-служб Exchange. Область действия маркера будет ограничена доступом только для чтения к текущему элементу и его вложениям, если в манифесте надстройки не указано разрешение [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission). Если указано разрешение `ReadWriteMailbox`, полученный маркер предоставит доступ на чтение и запись к почте, календарю и контактам, включая возможность отправки почты.</span><span class="sxs-lookup"><span data-stu-id="27126-p132">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="27126-482">С помощью свойства `restUrl` надстройка должна определить правильный URL-адрес для вызовов REST API.</span><span class="sxs-lookup"><span data-stu-id="27126-482">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="27126-483">**Маркеры EWS**</span><span class="sxs-lookup"><span data-stu-id="27126-483">**EWS Tokens**</span></span>

<span data-ttu-id="27126-p133">Если запрашивается маркер EWS (`options.isRest = false`), полученный маркер не подойдет для проверки подлинности при вызовах REST API. Область действия маркера будет ограничена доступом к текущему элементу.</span><span class="sxs-lookup"><span data-stu-id="27126-p133">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="27126-486">С помощью свойства `ewsUrl` надстройка должна определить правильный URL-адрес для вызовов EWS.</span><span class="sxs-lookup"><span data-stu-id="27126-486">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="27126-487">Параметры:</span><span class="sxs-lookup"><span data-stu-id="27126-487">Parameters:</span></span>

|<span data-ttu-id="27126-488">Имя</span><span class="sxs-lookup"><span data-stu-id="27126-488">Name</span></span>| <span data-ttu-id="27126-489">Тип</span><span class="sxs-lookup"><span data-stu-id="27126-489">Type</span></span>| <span data-ttu-id="27126-490">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="27126-490">Attributes</span></span>| <span data-ttu-id="27126-491">Описание</span><span class="sxs-lookup"><span data-stu-id="27126-491">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="27126-492">Object</span><span class="sxs-lookup"><span data-stu-id="27126-492">Object</span></span> | <span data-ttu-id="27126-493">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="27126-493">&lt;optional&gt;</span></span> | <span data-ttu-id="27126-494">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="27126-494">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="27126-495">Boolean</span><span class="sxs-lookup"><span data-stu-id="27126-495">Boolean</span></span> |  <span data-ttu-id="27126-496">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="27126-496">&lt;optional&gt;</span></span> | <span data-ttu-id="27126-p134">Определяет, будет ли предоставленный маркер использоваться для интерфейсов REST API Outlook или веб-служб Exchange. Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="27126-p134">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="27126-499">Объект</span><span class="sxs-lookup"><span data-stu-id="27126-499">Object</span></span> |  <span data-ttu-id="27126-500">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="27126-500">&lt;optional&gt;</span></span> | <span data-ttu-id="27126-501">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="27126-501">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="27126-502">function</span><span class="sxs-lookup"><span data-stu-id="27126-502">function</span></span>||<span data-ttu-id="27126-p135">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult). Токен указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="27126-p135">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="27126-505">Требования</span><span class="sxs-lookup"><span data-stu-id="27126-505">Requirements</span></span>

|<span data-ttu-id="27126-506">Требование</span><span class="sxs-lookup"><span data-stu-id="27126-506">Requirement</span></span>| <span data-ttu-id="27126-507">Значение</span><span class="sxs-lookup"><span data-stu-id="27126-507">Value</span></span>|
|---|---|
|[<span data-ttu-id="27126-508">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="27126-508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="27126-509">1.5</span><span class="sxs-lookup"><span data-stu-id="27126-509">1.5</span></span> |
|[<span data-ttu-id="27126-510">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="27126-510">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="27126-511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27126-511">ReadItem</span></span>|
|[<span data-ttu-id="27126-512">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="27126-512">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="27126-513">Создание и чтение</span><span class="sxs-lookup"><span data-stu-id="27126-513">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="27126-514">Пример</span><span class="sxs-lookup"><span data-stu-id="27126-514">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="27126-515">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="27126-515">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="27126-516">Получает строку, содержащую маркер, используемый для получения вложения или элемента с Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="27126-516">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="27126-p136">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный токен с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="27126-p136">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="27126-p137">Вы можете передать сторонней системе токен и идентификатор вложения или элемента. Сторонняя система использует этот токен как токен авторизации, чтобы вызвать операцию [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) или [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) веб-служб Exchange для возврата вложения или элемента. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="27126-p137">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="27126-522">Для вызова метода `getCallbackTokenAsync` в режиме чтения манифесте приложения должно быть указано разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="27126-522">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="27126-p138">Чтобы получить идентификатор элемента для передачи в метод `getCallbackTokenAsync`, в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="27126-p138">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="27126-525">Параметры</span><span class="sxs-lookup"><span data-stu-id="27126-525">Parameters:</span></span>

|<span data-ttu-id="27126-526">Имя</span><span class="sxs-lookup"><span data-stu-id="27126-526">Name</span></span>| <span data-ttu-id="27126-527">Тип</span><span class="sxs-lookup"><span data-stu-id="27126-527">Type</span></span>| <span data-ttu-id="27126-528">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="27126-528">Attributes</span></span>| <span data-ttu-id="27126-529">Описание</span><span class="sxs-lookup"><span data-stu-id="27126-529">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="27126-530">function</span><span class="sxs-lookup"><span data-stu-id="27126-530">function</span></span>||<span data-ttu-id="27126-p139">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult). Токен указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="27126-p139">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="27126-533">Объект</span><span class="sxs-lookup"><span data-stu-id="27126-533">Object</span></span>| <span data-ttu-id="27126-534">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="27126-534">&lt;optional&gt;</span></span>|<span data-ttu-id="27126-535">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="27126-535">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="27126-536">Требования</span><span class="sxs-lookup"><span data-stu-id="27126-536">Requirements</span></span>

|<span data-ttu-id="27126-537">Требование</span><span class="sxs-lookup"><span data-stu-id="27126-537">Requirement</span></span>| <span data-ttu-id="27126-538">Значение</span><span class="sxs-lookup"><span data-stu-id="27126-538">Value</span></span>|
|---|---|
|[<span data-ttu-id="27126-539">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="27126-539">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="27126-540">1.3</span><span class="sxs-lookup"><span data-stu-id="27126-540">1.3</span></span>|
|[<span data-ttu-id="27126-541">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="27126-541">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="27126-542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27126-542">ReadItem</span></span>|
|[<span data-ttu-id="27126-543">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="27126-543">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="27126-544">Создание и чтение</span><span class="sxs-lookup"><span data-stu-id="27126-544">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="27126-545">Пример</span><span class="sxs-lookup"><span data-stu-id="27126-545">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="27126-546">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="27126-546">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="27126-547">Получает маркер, идентифицирующий пользователя и надстройку Office.</span><span class="sxs-lookup"><span data-stu-id="27126-547">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="27126-548">Метод `getUserIdentityTokenAsync` возвращает токен, который можно использовать для идентификации, а также [проверки подлинности надстройки и пользователя в сторонней системе](https://docs.microsoft.com/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="27126-548">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="27126-549">Параметры</span><span class="sxs-lookup"><span data-stu-id="27126-549">Parameters:</span></span>

|<span data-ttu-id="27126-550">Имя</span><span class="sxs-lookup"><span data-stu-id="27126-550">Name</span></span>| <span data-ttu-id="27126-551">Тип</span><span class="sxs-lookup"><span data-stu-id="27126-551">Type</span></span>| <span data-ttu-id="27126-552">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="27126-552">Attributes</span></span>| <span data-ttu-id="27126-553">Описание</span><span class="sxs-lookup"><span data-stu-id="27126-553">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="27126-554">function</span><span class="sxs-lookup"><span data-stu-id="27126-554">function</span></span>||<span data-ttu-id="27126-555">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="27126-555">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="27126-556">Токен указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="27126-556">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="27126-557">Object</span><span class="sxs-lookup"><span data-stu-id="27126-557">Object</span></span>| <span data-ttu-id="27126-558">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="27126-558">&lt;optional&gt;</span></span>|<span data-ttu-id="27126-559">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="27126-559">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="27126-560">Требования</span><span class="sxs-lookup"><span data-stu-id="27126-560">Requirements</span></span>

|<span data-ttu-id="27126-561">Требование</span><span class="sxs-lookup"><span data-stu-id="27126-561">Requirement</span></span>| <span data-ttu-id="27126-562">Значение</span><span class="sxs-lookup"><span data-stu-id="27126-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="27126-563">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="27126-563">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="27126-564">1.0</span><span class="sxs-lookup"><span data-stu-id="27126-564">1.0</span></span>|
|[<span data-ttu-id="27126-565">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="27126-565">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="27126-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27126-566">ReadItem</span></span>|
|[<span data-ttu-id="27126-567">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="27126-567">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="27126-568">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="27126-568">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="27126-569">Пример</span><span class="sxs-lookup"><span data-stu-id="27126-569">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="27126-570">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="27126-570">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="27126-571">Выполняет асинхронный запрос для веб-служб Exchange (EWS) на сервере Exchange Server, на котором размещен почтовый ящик пользователя.</span><span class="sxs-lookup"><span data-stu-id="27126-571">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="27126-572">Этот метод не поддерживается в следующих сценариях:</span><span class="sxs-lookup"><span data-stu-id="27126-572">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="27126-573">В Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="27126-573">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="27126-574">Если надстройка загружается в почтовый ящик Gmail.</span><span class="sxs-lookup"><span data-stu-id="27126-574">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="27126-575">В таких случаях надстройка должна [использовать REST API](https://docs.microsoft.com/outlook/add-ins/use-rest-api) для доступа к почтовому ящику пользователя.</span><span class="sxs-lookup"><span data-stu-id="27126-575">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="27126-576">Метод `makeEwsRequestAsync` отправляет запрос EWS от имени надстройки в Exchange.</span><span class="sxs-lookup"><span data-stu-id="27126-576">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="27126-577">Список поддерживаемых операций EWS см. в статье [Вызов веб-служб из надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support).</span><span class="sxs-lookup"><span data-stu-id="27126-577">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="27126-578">С помощью метода `makeEwsRequestAsync` невозможно запрашивать элементы, связанные с папкой.</span><span class="sxs-lookup"><span data-stu-id="27126-578">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="27126-579">В запросе XML должна быть указана кодировка UTF-8.</span><span class="sxs-lookup"><span data-stu-id="27126-579">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="27126-p141">У вашей надстройки должно быть разрешение **ReadWriteMailbox** для использования метода `makeEwsRequestAsync`. Сведения об использовании разрешения **ReadWriteMailbox** и операций EWS, которые можно вызывать с помощью метода `makeEwsRequestAsync`, см. в статье [Указание разрешений для доступа почтовой надстройки к почтовому ящику пользователя](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="27126-p141">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="27126-582">Администратор сервера должен установить значение true для параметра `OAuthAuthentication` в каталоге сервера клиентского доступа EWS, чтобы метод `makeEwsRequestAsync` мог выполнять запросы EWS.</span><span class="sxs-lookup"><span data-stu-id="27126-582">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="27126-583">Различия версий</span><span class="sxs-lookup"><span data-stu-id="27126-583">Version differences</span></span>

<span data-ttu-id="27126-584">Если вы используете метод `makeEwsRequestAsync` в почтовых приложениях, которые выполняются в Outlook версии более ранней, чем 15.0.4535.1004, указывайте кодировку `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="27126-584">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="27126-p142">Значение кодировки не нужно указывать, если почтовое приложение выполняется в Outlook в Интернете. Чтобы определить, выполняется ли приложение в Outlook или Outlook в Интернете, используйте свойство mailbox.diagnostics.hostName. Используемую версию Outlook можно определить с помощью свойства mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="27126-p142">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="27126-588">Параметры</span><span class="sxs-lookup"><span data-stu-id="27126-588">Parameters:</span></span>

|<span data-ttu-id="27126-589">Имя</span><span class="sxs-lookup"><span data-stu-id="27126-589">Name</span></span>| <span data-ttu-id="27126-590">Тип</span><span class="sxs-lookup"><span data-stu-id="27126-590">Type</span></span>| <span data-ttu-id="27126-591">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="27126-591">Attributes</span></span>| <span data-ttu-id="27126-592">Описание</span><span class="sxs-lookup"><span data-stu-id="27126-592">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="27126-593">String</span><span class="sxs-lookup"><span data-stu-id="27126-593">String</span></span>||<span data-ttu-id="27126-594">Запрос EWS.</span><span class="sxs-lookup"><span data-stu-id="27126-594">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="27126-595">function</span><span class="sxs-lookup"><span data-stu-id="27126-595">function</span></span>||<span data-ttu-id="27126-596">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="27126-596">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="27126-597">Результат XML вызова EWS указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="27126-597">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="27126-598">Если размер результата превышает 1 МБ, возвращается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="27126-598">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="27126-599">Объект</span><span class="sxs-lookup"><span data-stu-id="27126-599">Object</span></span>| <span data-ttu-id="27126-600">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="27126-600">&lt;optional&gt;</span></span>|<span data-ttu-id="27126-601">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="27126-601">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="27126-602">Требования</span><span class="sxs-lookup"><span data-stu-id="27126-602">Requirements</span></span>

|<span data-ttu-id="27126-603">Требование</span><span class="sxs-lookup"><span data-stu-id="27126-603">Requirement</span></span>| <span data-ttu-id="27126-604">Значение</span><span class="sxs-lookup"><span data-stu-id="27126-604">Value</span></span>|
|---|---|
|[<span data-ttu-id="27126-605">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="27126-605">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="27126-606">1.0</span><span class="sxs-lookup"><span data-stu-id="27126-606">1.0</span></span>|
|[<span data-ttu-id="27126-607">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="27126-607">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="27126-608">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="27126-608">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="27126-609">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="27126-609">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="27126-610">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="27126-610">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="27126-611">Пример</span><span class="sxs-lookup"><span data-stu-id="27126-611">Example</span></span>

<span data-ttu-id="27126-612">В следующем примере вызывается `makeEwsRequestAsync` для получения темы элемента с помощью операции `GetItem`.</span><span class="sxs-lookup"><span data-stu-id="27126-612">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

####  <a name="removehandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="27126-613">removeHandlerAsync(eventType, handler, [параметры], [функция обратного вызова])</span><span class="sxs-lookup"><span data-stu-id="27126-613">removeHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="27126-614">Удаляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="27126-614">Removes an event handler for a supported event.</span></span>

<span data-ttu-id="27126-615">В настоящее время единственный поддерживаемый тип события — `Office.EventType.ItemChanged`.</span><span class="sxs-lookup"><span data-stu-id="27126-615">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="27126-616">Параметры:</span><span class="sxs-lookup"><span data-stu-id="27126-616">Parameters:</span></span>

| <span data-ttu-id="27126-617">Имя</span><span class="sxs-lookup"><span data-stu-id="27126-617">Name</span></span> | <span data-ttu-id="27126-618">Тип</span><span class="sxs-lookup"><span data-stu-id="27126-618">Type</span></span> | <span data-ttu-id="27126-619">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="27126-619">Attributes</span></span> | <span data-ttu-id="27126-620">Описание</span><span class="sxs-lookup"><span data-stu-id="27126-620">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="27126-621">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="27126-621">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="27126-622">Событие, которое должно отменить обработчик.</span><span class="sxs-lookup"><span data-stu-id="27126-622">The event that should revoke the handler.</span></span> |
| `handler` | <span data-ttu-id="27126-623">Функция</span><span class="sxs-lookup"><span data-stu-id="27126-623">Function</span></span> || <span data-ttu-id="27126-p144">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="27126-p144">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="27126-627">Объект</span><span class="sxs-lookup"><span data-stu-id="27126-627">Object</span></span> | <span data-ttu-id="27126-628">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="27126-628">&lt;optional&gt;</span></span> | <span data-ttu-id="27126-629">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="27126-629">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="27126-630">Object</span><span class="sxs-lookup"><span data-stu-id="27126-630">Object</span></span> | <span data-ttu-id="27126-631">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="27126-631">&lt;optional&gt;</span></span> | <span data-ttu-id="27126-632">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="27126-632">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="27126-633">функция</span><span class="sxs-lookup"><span data-stu-id="27126-633">function</span></span>| <span data-ttu-id="27126-634">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="27126-634">&lt;optional&gt;</span></span>|<span data-ttu-id="27126-635">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="27126-635">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="27126-636">Требования</span><span class="sxs-lookup"><span data-stu-id="27126-636">Requirements</span></span>

|<span data-ttu-id="27126-637">Требование</span><span class="sxs-lookup"><span data-stu-id="27126-637">Requirement</span></span>| <span data-ttu-id="27126-638">Значение</span><span class="sxs-lookup"><span data-stu-id="27126-638">Value</span></span>|
|---|---|
|[<span data-ttu-id="27126-639">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="27126-639">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="27126-640">1.5</span><span class="sxs-lookup"><span data-stu-id="27126-640">1.5</span></span> |
|[<span data-ttu-id="27126-641">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="27126-641">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="27126-642">ReadItem</span><span class="sxs-lookup"><span data-stu-id="27126-642">ReadItem</span></span> |
|[<span data-ttu-id="27126-643">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="27126-643">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="27126-644">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="27126-644">Compose or read</span></span>|