---
title: Office.context.mailbox — набор обязательных элементов 1.6
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: 91ba2945b3c6390f4623e1d716c516790be26ad3
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/16/2019
ms.locfileid: "30068100"
---
# <a name="mailbox"></a><span data-ttu-id="38be1-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="38be1-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="38be1-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="38be1-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="38be1-104">Предоставляет для Microsoft Outlook и Microsoft Outlook в Интернете доступ к объектной модели надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="38be1-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="38be1-105">Требования</span><span class="sxs-lookup"><span data-stu-id="38be1-105">Requirements</span></span>

|<span data-ttu-id="38be1-106">Требование</span><span class="sxs-lookup"><span data-stu-id="38be1-106">Requirement</span></span>| <span data-ttu-id="38be1-107">Значение</span><span class="sxs-lookup"><span data-stu-id="38be1-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="38be1-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="38be1-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="38be1-109">1.0</span><span class="sxs-lookup"><span data-stu-id="38be1-109">1.0</span></span>|
|[<span data-ttu-id="38be1-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="38be1-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="38be1-111">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="38be1-111">Restricted</span></span>|
|[<span data-ttu-id="38be1-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="38be1-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="38be1-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="38be1-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="38be1-114">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="38be1-114">Members and methods</span></span>

| <span data-ttu-id="38be1-115">Элемент</span><span class="sxs-lookup"><span data-stu-id="38be1-115">Member</span></span> | <span data-ttu-id="38be1-116">Тип</span><span class="sxs-lookup"><span data-stu-id="38be1-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="38be1-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="38be1-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="38be1-118">Элемент</span><span class="sxs-lookup"><span data-stu-id="38be1-118">Member</span></span> |
| [<span data-ttu-id="38be1-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="38be1-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="38be1-120">Элемент</span><span class="sxs-lookup"><span data-stu-id="38be1-120">Member</span></span> |
| [<span data-ttu-id="38be1-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="38be1-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="38be1-122">Метод</span><span class="sxs-lookup"><span data-stu-id="38be1-122">Method</span></span> |
| [<span data-ttu-id="38be1-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="38be1-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="38be1-124">Метод</span><span class="sxs-lookup"><span data-stu-id="38be1-124">Method</span></span> |
| [<span data-ttu-id="38be1-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="38be1-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime) | <span data-ttu-id="38be1-126">Метод</span><span class="sxs-lookup"><span data-stu-id="38be1-126">Method</span></span> |
| [<span data-ttu-id="38be1-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="38be1-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="38be1-128">Метод</span><span class="sxs-lookup"><span data-stu-id="38be1-128">Method</span></span> |
| [<span data-ttu-id="38be1-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="38be1-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="38be1-130">Метод</span><span class="sxs-lookup"><span data-stu-id="38be1-130">Method</span></span> |
| [<span data-ttu-id="38be1-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="38be1-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="38be1-132">Метод</span><span class="sxs-lookup"><span data-stu-id="38be1-132">Method</span></span> |
| [<span data-ttu-id="38be1-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="38be1-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="38be1-134">Метод</span><span class="sxs-lookup"><span data-stu-id="38be1-134">Method</span></span> |
| [<span data-ttu-id="38be1-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="38be1-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="38be1-136">Метод</span><span class="sxs-lookup"><span data-stu-id="38be1-136">Method</span></span> |
| [<span data-ttu-id="38be1-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="38be1-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="38be1-138">Метод</span><span class="sxs-lookup"><span data-stu-id="38be1-138">Method</span></span> |
| [<span data-ttu-id="38be1-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="38be1-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="38be1-140">Метод</span><span class="sxs-lookup"><span data-stu-id="38be1-140">Method</span></span> |
| [<span data-ttu-id="38be1-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="38be1-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="38be1-142">Метод</span><span class="sxs-lookup"><span data-stu-id="38be1-142">Method</span></span> |
| [<span data-ttu-id="38be1-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="38be1-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="38be1-144">Метод</span><span class="sxs-lookup"><span data-stu-id="38be1-144">Method</span></span> |
| [<span data-ttu-id="38be1-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="38be1-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="38be1-146">Метод</span><span class="sxs-lookup"><span data-stu-id="38be1-146">Method</span></span> |
| [<span data-ttu-id="38be1-147">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="38be1-147">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="38be1-148">Метод</span><span class="sxs-lookup"><span data-stu-id="38be1-148">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="38be1-149">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="38be1-149">Namespaces</span></span>

<span data-ttu-id="38be1-150">[diagnostics](Office.context.mailbox.diagnostics.md). Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="38be1-150">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="38be1-151">[item](Office.context.mailbox.item.md). Предоставляет методы и свойства для доступа к сообщению или встрече в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="38be1-151">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="38be1-152">[userProfile](Office.context.mailbox.userProfile.md). Предоставляет сведения о пользователе в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="38be1-152">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="38be1-153">Элементы</span><span class="sxs-lookup"><span data-stu-id="38be1-153">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="38be1-154">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="38be1-154">ewsUrl :String</span></span>

<span data-ttu-id="38be1-p101">Получает URL-адрес конечной точки веб-служб Exchange (EWS) для этой учетной записи электронной почты. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="38be1-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="38be1-157">Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="38be1-157">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="38be1-p102">Удаленная служба может использовать значение `ewsUrl`, чтобы выполнять вызовы EWS для почтового ящика пользователя. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="38be1-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="38be1-160">Чтобы вызвать элемент `ewsUrl` в режиме чтения, в манифесте приложения должно быть указано разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="38be1-160">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="38be1-p103">Перед использованием элемента `ewsUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="38be1-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="38be1-163">Тип</span><span class="sxs-lookup"><span data-stu-id="38be1-163">Type</span></span>

*   <span data-ttu-id="38be1-164">String</span><span class="sxs-lookup"><span data-stu-id="38be1-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="38be1-165">Требования</span><span class="sxs-lookup"><span data-stu-id="38be1-165">Requirements</span></span>

|<span data-ttu-id="38be1-166">Требование</span><span class="sxs-lookup"><span data-stu-id="38be1-166">Requirement</span></span>| <span data-ttu-id="38be1-167">Значение</span><span class="sxs-lookup"><span data-stu-id="38be1-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="38be1-168">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="38be1-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="38be1-169">1.0</span><span class="sxs-lookup"><span data-stu-id="38be1-169">1.0</span></span>|
|[<span data-ttu-id="38be1-170">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="38be1-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="38be1-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="38be1-171">ReadItem</span></span>|
|[<span data-ttu-id="38be1-172">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="38be1-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="38be1-173">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="38be1-173">Compose or Read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="38be1-174">restUrl :String</span><span class="sxs-lookup"><span data-stu-id="38be1-174">restUrl :String</span></span>

<span data-ttu-id="38be1-175">Возвращает URL-адрес конечной точки REST для этой учетной записи электронной почты.</span><span class="sxs-lookup"><span data-stu-id="38be1-175">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="38be1-176">С помощью значения `restUrl` можно выполнять вызовы [REST API](https://docs.microsoft.com/outlook/rest/) для почтового ящика пользователя.</span><span class="sxs-lookup"><span data-stu-id="38be1-176">The `restUrl` value can be used to make [REST API](https://docs.microsoft.com/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="38be1-177">Чтобы вызвать элемент `restUrl` в режиме чтения, в манифесте приложения необходимо указать разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="38be1-177">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="38be1-p104">Перед использованием элемента `restUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="38be1-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="38be1-180">Тип</span><span class="sxs-lookup"><span data-stu-id="38be1-180">Type</span></span>

*   <span data-ttu-id="38be1-181">String</span><span class="sxs-lookup"><span data-stu-id="38be1-181">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="38be1-182">Требования</span><span class="sxs-lookup"><span data-stu-id="38be1-182">Requirements</span></span>

|<span data-ttu-id="38be1-183">Требование</span><span class="sxs-lookup"><span data-stu-id="38be1-183">Requirement</span></span>| <span data-ttu-id="38be1-184">Значение</span><span class="sxs-lookup"><span data-stu-id="38be1-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="38be1-185">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="38be1-185">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="38be1-186">1.5</span><span class="sxs-lookup"><span data-stu-id="38be1-186">1.5</span></span> |
|[<span data-ttu-id="38be1-187">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="38be1-187">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="38be1-188">ReadItem</span><span class="sxs-lookup"><span data-stu-id="38be1-188">ReadItem</span></span>|
|[<span data-ttu-id="38be1-189">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="38be1-189">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="38be1-190">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="38be1-190">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="38be1-191">Методы</span><span class="sxs-lookup"><span data-stu-id="38be1-191">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="38be1-192">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="38be1-192">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="38be1-193">Добавляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="38be1-193">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="38be1-194">В настоящее время поддерживаются только события типа `Office.EventType.ItemChanged`, которые вызываются, когда пользователь выбирает новый элемент.</span><span class="sxs-lookup"><span data-stu-id="38be1-194">Currently the only supported event type is `Office.EventType.ItemChanged`, which is invoked when the user selects a new item.</span></span> <span data-ttu-id="38be1-195">Это событие используется надстройками, реализующими закрепляемую область задач, и позволяет надстройке обновлять пользовательский интерфейс области задач в соответствии с выбранным в данный момент элементом.</span><span class="sxs-lookup"><span data-stu-id="38be1-195">This event is used by add-ins that implement a pinnable task pane, and allows the add-in to refresh the task pane UI based on the currently selected item.</span></span>

##### <a name="parameters"></a><span data-ttu-id="38be1-196">Параметры</span><span class="sxs-lookup"><span data-stu-id="38be1-196">Parameters</span></span>

| <span data-ttu-id="38be1-197">Имя</span><span class="sxs-lookup"><span data-stu-id="38be1-197">Name</span></span> | <span data-ttu-id="38be1-198">Тип</span><span class="sxs-lookup"><span data-stu-id="38be1-198">Type</span></span> | <span data-ttu-id="38be1-199">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="38be1-199">Attributes</span></span> | <span data-ttu-id="38be1-200">Описание</span><span class="sxs-lookup"><span data-stu-id="38be1-200">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="38be1-201">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="38be1-201">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="38be1-202">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="38be1-202">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="38be1-203">Function</span><span class="sxs-lookup"><span data-stu-id="38be1-203">Function</span></span> || <span data-ttu-id="38be1-p106">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="38be1-p106">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="38be1-207">Объект</span><span class="sxs-lookup"><span data-stu-id="38be1-207">Object</span></span> | <span data-ttu-id="38be1-208">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="38be1-208">&lt;optional&gt;</span></span> | <span data-ttu-id="38be1-209">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="38be1-209">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="38be1-210">Object</span><span class="sxs-lookup"><span data-stu-id="38be1-210">Object</span></span> | <span data-ttu-id="38be1-211">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="38be1-211">&lt;optional&gt;</span></span> | <span data-ttu-id="38be1-212">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="38be1-212">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="38be1-213">функция</span><span class="sxs-lookup"><span data-stu-id="38be1-213">function</span></span>| <span data-ttu-id="38be1-214">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="38be1-214">&lt;optional&gt;</span></span>|<span data-ttu-id="38be1-215">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="38be1-215">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="38be1-216">Требования</span><span class="sxs-lookup"><span data-stu-id="38be1-216">Requirements</span></span>

|<span data-ttu-id="38be1-217">Требование</span><span class="sxs-lookup"><span data-stu-id="38be1-217">Requirement</span></span>| <span data-ttu-id="38be1-218">Значение</span><span class="sxs-lookup"><span data-stu-id="38be1-218">Value</span></span>|
|---|---|
|[<span data-ttu-id="38be1-219">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="38be1-219">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="38be1-220">1.5</span><span class="sxs-lookup"><span data-stu-id="38be1-220">1.5</span></span> |
|[<span data-ttu-id="38be1-221">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="38be1-221">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="38be1-222">ReadItem</span><span class="sxs-lookup"><span data-stu-id="38be1-222">ReadItem</span></span> |
|[<span data-ttu-id="38be1-223">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="38be1-223">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="38be1-224">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="38be1-224">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="38be1-225">Пример</span><span class="sxs-lookup"><span data-stu-id="38be1-225">Example</span></span>

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

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="38be1-226">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="38be1-226">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="38be1-227">Преобразовывает идентификатор элемента из формата REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="38be1-227">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="38be1-228">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="38be1-228">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="38be1-p107">Формат идентификаторов, извлекаемых через API REST (например, [API Почты Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)), отличается от формата веб-служб Exchange (EWS). Метод `convertToEwsId` преобразовывает идентификатор в формате REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="38be1-p107">Item IDs retrieved via a REST API (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="38be1-231">Параметры</span><span class="sxs-lookup"><span data-stu-id="38be1-231">Parameters</span></span>

|<span data-ttu-id="38be1-232">Имя</span><span class="sxs-lookup"><span data-stu-id="38be1-232">Name</span></span>| <span data-ttu-id="38be1-233">Тип</span><span class="sxs-lookup"><span data-stu-id="38be1-233">Type</span></span>| <span data-ttu-id="38be1-234">Описание</span><span class="sxs-lookup"><span data-stu-id="38be1-234">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="38be1-235">String</span><span class="sxs-lookup"><span data-stu-id="38be1-235">String</span></span>|<span data-ttu-id="38be1-236">Идентификатор элемента в формате REST API для Outlook</span><span class="sxs-lookup"><span data-stu-id="38be1-236">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="38be1-237">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="38be1-237">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.restversion)|<span data-ttu-id="38be1-238">Значение, определяющее версию REST API для Outlook, которая используется для извлечения идентификатора элемента.</span><span class="sxs-lookup"><span data-stu-id="38be1-238">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="38be1-239">Требования</span><span class="sxs-lookup"><span data-stu-id="38be1-239">Requirements</span></span>

|<span data-ttu-id="38be1-240">Требование</span><span class="sxs-lookup"><span data-stu-id="38be1-240">Requirement</span></span>| <span data-ttu-id="38be1-241">Значение</span><span class="sxs-lookup"><span data-stu-id="38be1-241">Value</span></span>|
|---|---|
|[<span data-ttu-id="38be1-242">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="38be1-242">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="38be1-243">1.3</span><span class="sxs-lookup"><span data-stu-id="38be1-243">1.3</span></span>|
|[<span data-ttu-id="38be1-244">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="38be1-244">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="38be1-245">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="38be1-245">Restricted</span></span>|
|[<span data-ttu-id="38be1-246">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="38be1-246">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="38be1-247">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="38be1-247">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="38be1-248">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="38be1-248">Returns:</span></span>

<span data-ttu-id="38be1-249">Тип: String</span><span class="sxs-lookup"><span data-stu-id="38be1-249">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="38be1-250">Пример</span><span class="sxs-lookup"><span data-stu-id="38be1-250">Example</span></span>

```javascript
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook16officelocalclienttime"></a><span data-ttu-id="38be1-251">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="38be1-251">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)}</span></span>

<span data-ttu-id="38be1-252">Получает словарь, содержащий сведения о локальном времени клиента.</span><span class="sxs-lookup"><span data-stu-id="38be1-252">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="38be1-p108">В случае дат и времени в почтовом приложении для Outlook или Outlook Web App могут использоваться разные часовые пояса. Outlook использует часовой пояс клиентского компьютера. Outlook Web App использует часовой пояс, заданный в Центре администрирования Exchange (EAC). Значения даты и времени должны обрабатываться так, чтобы значения в пользовательском интерфейсе всегда согласовывались с часовым поясом, ожидаемым пользователем.</span><span class="sxs-lookup"><span data-stu-id="38be1-p108">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="38be1-p109">Если почтовое приложение работает в Outlook, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса клиентского компьютера. Если почтовое приложение работает в Outlook Web App, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса, заданного в Центре администрирования Exchange.</span><span class="sxs-lookup"><span data-stu-id="38be1-p109">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="38be1-258">Параметры</span><span class="sxs-lookup"><span data-stu-id="38be1-258">Parameters</span></span>

|<span data-ttu-id="38be1-259">Имя</span><span class="sxs-lookup"><span data-stu-id="38be1-259">Name</span></span>| <span data-ttu-id="38be1-260">Тип</span><span class="sxs-lookup"><span data-stu-id="38be1-260">Type</span></span>| <span data-ttu-id="38be1-261">Описание</span><span class="sxs-lookup"><span data-stu-id="38be1-261">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="38be1-262">Date</span><span class="sxs-lookup"><span data-stu-id="38be1-262">Date</span></span>|<span data-ttu-id="38be1-263">Объект Date</span><span class="sxs-lookup"><span data-stu-id="38be1-263">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="38be1-264">Требования</span><span class="sxs-lookup"><span data-stu-id="38be1-264">Requirements</span></span>

|<span data-ttu-id="38be1-265">Требование</span><span class="sxs-lookup"><span data-stu-id="38be1-265">Requirement</span></span>| <span data-ttu-id="38be1-266">Значение</span><span class="sxs-lookup"><span data-stu-id="38be1-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="38be1-267">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="38be1-267">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="38be1-268">1.0</span><span class="sxs-lookup"><span data-stu-id="38be1-268">1.0</span></span>|
|[<span data-ttu-id="38be1-269">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="38be1-269">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="38be1-270">ReadItem</span><span class="sxs-lookup"><span data-stu-id="38be1-270">ReadItem</span></span>|
|[<span data-ttu-id="38be1-271">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="38be1-271">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="38be1-272">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="38be1-272">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="38be1-273">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="38be1-273">Returns:</span></span>

<span data-ttu-id="38be1-274">Тип: [LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="38be1-274">Type: [LocalClientTime](/javascript/api/outlook_1_6/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="38be1-275">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="38be1-275">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="38be1-276">Преобразовывает идентификатор элемента в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="38be1-276">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="38be1-277">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="38be1-277">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="38be1-p110">Формат идентификаторов, извлекаемых через EWS или свойство `itemId`, отличается от формата API REST (таких как [API Почты Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)). Метод `convertToRestId` преобразовывает идентификатор в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="38be1-p110">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="38be1-280">Параметры</span><span class="sxs-lookup"><span data-stu-id="38be1-280">Parameters</span></span>

|<span data-ttu-id="38be1-281">Имя</span><span class="sxs-lookup"><span data-stu-id="38be1-281">Name</span></span>| <span data-ttu-id="38be1-282">Тип</span><span class="sxs-lookup"><span data-stu-id="38be1-282">Type</span></span>| <span data-ttu-id="38be1-283">Описание</span><span class="sxs-lookup"><span data-stu-id="38be1-283">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="38be1-284">String</span><span class="sxs-lookup"><span data-stu-id="38be1-284">String</span></span>|<span data-ttu-id="38be1-285">Идентификатор элемента в формате EWS</span><span class="sxs-lookup"><span data-stu-id="38be1-285">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="38be1-286">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="38be1-286">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook_1_6/office.mailboxenums.restversion)|<span data-ttu-id="38be1-287">Значение, определяющее версию REST API для Outlook, с которой будет использоваться преобразованный идентификатор.</span><span class="sxs-lookup"><span data-stu-id="38be1-287">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="38be1-288">Требования</span><span class="sxs-lookup"><span data-stu-id="38be1-288">Requirements</span></span>

|<span data-ttu-id="38be1-289">Требование</span><span class="sxs-lookup"><span data-stu-id="38be1-289">Requirement</span></span>| <span data-ttu-id="38be1-290">Значение</span><span class="sxs-lookup"><span data-stu-id="38be1-290">Value</span></span>|
|---|---|
|[<span data-ttu-id="38be1-291">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="38be1-291">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="38be1-292">1.3</span><span class="sxs-lookup"><span data-stu-id="38be1-292">1.3</span></span>|
|[<span data-ttu-id="38be1-293">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="38be1-293">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="38be1-294">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="38be1-294">Restricted</span></span>|
|[<span data-ttu-id="38be1-295">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="38be1-295">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="38be1-296">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="38be1-296">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="38be1-297">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="38be1-297">Returns:</span></span>

<span data-ttu-id="38be1-298">Тип: String</span><span class="sxs-lookup"><span data-stu-id="38be1-298">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="38be1-299">Пример</span><span class="sxs-lookup"><span data-stu-id="38be1-299">Example</span></span>

```javascript
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="38be1-300">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="38be1-300">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="38be1-301">Получает объект Date из словаря, содержащего сведения о времени.</span><span class="sxs-lookup"><span data-stu-id="38be1-301">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="38be1-302">Метод `convertToUtcClientTime` преобразует словарь, содержащий локальную дату и время, в объект Date с правильными значениями локальной даты и времени.</span><span class="sxs-lookup"><span data-stu-id="38be1-302">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="38be1-303">Параметры</span><span class="sxs-lookup"><span data-stu-id="38be1-303">Parameters</span></span>

|<span data-ttu-id="38be1-304">Имя</span><span class="sxs-lookup"><span data-stu-id="38be1-304">Name</span></span>| <span data-ttu-id="38be1-305">Тип</span><span class="sxs-lookup"><span data-stu-id="38be1-305">Type</span></span>| <span data-ttu-id="38be1-306">Описание</span><span class="sxs-lookup"><span data-stu-id="38be1-306">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="38be1-307">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="38be1-307">LocalClientTime</span></span>](/javascript/api/outlook_1_6/office.LocalClientTime)|<span data-ttu-id="38be1-308">Значение локального времени для преобразования.</span><span class="sxs-lookup"><span data-stu-id="38be1-308">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="38be1-309">Требования</span><span class="sxs-lookup"><span data-stu-id="38be1-309">Requirements</span></span>

|<span data-ttu-id="38be1-310">Требование</span><span class="sxs-lookup"><span data-stu-id="38be1-310">Requirement</span></span>| <span data-ttu-id="38be1-311">Значение</span><span class="sxs-lookup"><span data-stu-id="38be1-311">Value</span></span>|
|---|---|
|[<span data-ttu-id="38be1-312">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="38be1-312">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="38be1-313">1.0</span><span class="sxs-lookup"><span data-stu-id="38be1-313">1.0</span></span>|
|[<span data-ttu-id="38be1-314">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="38be1-314">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="38be1-315">ReadItem</span><span class="sxs-lookup"><span data-stu-id="38be1-315">ReadItem</span></span>|
|[<span data-ttu-id="38be1-316">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="38be1-316">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="38be1-317">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="38be1-317">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="38be1-318">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="38be1-318">Returns:</span></span>

<span data-ttu-id="38be1-319">Объект Date со временем в формате UTC.</span><span class="sxs-lookup"><span data-stu-id="38be1-319">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="38be1-320">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="38be1-320">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="38be1-321">Date</span><span class="sxs-lookup"><span data-stu-id="38be1-321">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="38be1-322">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="38be1-322">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="38be1-323">Отображает имеющуюся встречу из календаря.</span><span class="sxs-lookup"><span data-stu-id="38be1-323">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="38be1-324">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="38be1-324">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="38be1-325">Метод `displayAppointmentForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее сведения календаря о существующей встрече.</span><span class="sxs-lookup"><span data-stu-id="38be1-325">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="38be1-p111">В Outlook для Mac с помощью этого метода можно отобразить одну встречу, которая не является частью повторяющегося ряда, или основную встречу такого ряда, но не экземпляр из него, так как в Outlook для Mac невозможно получить доступ к свойствам экземпляра повторяющегося ряда (в том числе к идентификатору элемента).</span><span class="sxs-lookup"><span data-stu-id="38be1-p111">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="38be1-328">В Outlook Web App этот метод открывает указанную форму, только если текст формы содержит символы размером не более 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="38be1-328">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="38be1-329">Если указанный идентификатор элемента не определяет существующую встречу, на клиентском компьютере или устройстве открывается пустая страница, и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="38be1-329">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="38be1-330">Параметры</span><span class="sxs-lookup"><span data-stu-id="38be1-330">Parameters</span></span>

|<span data-ttu-id="38be1-331">Имя</span><span class="sxs-lookup"><span data-stu-id="38be1-331">Name</span></span>| <span data-ttu-id="38be1-332">Тип</span><span class="sxs-lookup"><span data-stu-id="38be1-332">Type</span></span>| <span data-ttu-id="38be1-333">Описание</span><span class="sxs-lookup"><span data-stu-id="38be1-333">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="38be1-334">String</span><span class="sxs-lookup"><span data-stu-id="38be1-334">String</span></span>|<span data-ttu-id="38be1-335">Идентификатор веб-служб Exchange для существующей встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="38be1-335">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="38be1-336">Требования</span><span class="sxs-lookup"><span data-stu-id="38be1-336">Requirements</span></span>

|<span data-ttu-id="38be1-337">Требование</span><span class="sxs-lookup"><span data-stu-id="38be1-337">Requirement</span></span>| <span data-ttu-id="38be1-338">Значение</span><span class="sxs-lookup"><span data-stu-id="38be1-338">Value</span></span>|
|---|---|
|[<span data-ttu-id="38be1-339">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="38be1-339">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="38be1-340">1.0</span><span class="sxs-lookup"><span data-stu-id="38be1-340">1.0</span></span>|
|[<span data-ttu-id="38be1-341">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="38be1-341">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="38be1-342">ReadItem</span><span class="sxs-lookup"><span data-stu-id="38be1-342">ReadItem</span></span>|
|[<span data-ttu-id="38be1-343">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="38be1-343">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="38be1-344">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="38be1-344">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="38be1-345">Пример</span><span class="sxs-lookup"><span data-stu-id="38be1-345">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="38be1-346">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="38be1-346">displayMessageForm(itemId)</span></span>

<span data-ttu-id="38be1-347">Отображает имеющееся сообщение.</span><span class="sxs-lookup"><span data-stu-id="38be1-347">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="38be1-348">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="38be1-348">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="38be1-349">Метод `displayMessageForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее существующее сообщение.</span><span class="sxs-lookup"><span data-stu-id="38be1-349">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="38be1-350">В Outlook Web App этот метод открывает указанную форму, только если текст формы содержит символы размером не более 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="38be1-350">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="38be1-351">Если указанный идентификатор элемента не определяет существующее сообщение, окно на клиентском компьютере не открывается и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="38be1-351">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="38be1-p112">Не используйте `displayMessageForm` с параметром `itemId`, который представляет собой встречу. Используйте метод `displayAppointmentForm`, чтобы отобразить сведения о существующей встрече, а метод `displayNewAppointmentForm` — для отображения формы создания встречи.</span><span class="sxs-lookup"><span data-stu-id="38be1-p112">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="38be1-354">Параметры</span><span class="sxs-lookup"><span data-stu-id="38be1-354">Parameters</span></span>

|<span data-ttu-id="38be1-355">Имя</span><span class="sxs-lookup"><span data-stu-id="38be1-355">Name</span></span>| <span data-ttu-id="38be1-356">Тип</span><span class="sxs-lookup"><span data-stu-id="38be1-356">Type</span></span>| <span data-ttu-id="38be1-357">Описание</span><span class="sxs-lookup"><span data-stu-id="38be1-357">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="38be1-358">String</span><span class="sxs-lookup"><span data-stu-id="38be1-358">String</span></span>|<span data-ttu-id="38be1-359">Идентификатор веб-служб Exchange для существующего сообщения.</span><span class="sxs-lookup"><span data-stu-id="38be1-359">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="38be1-360">Требования</span><span class="sxs-lookup"><span data-stu-id="38be1-360">Requirements</span></span>

|<span data-ttu-id="38be1-361">Требование</span><span class="sxs-lookup"><span data-stu-id="38be1-361">Requirement</span></span>| <span data-ttu-id="38be1-362">Значение</span><span class="sxs-lookup"><span data-stu-id="38be1-362">Value</span></span>|
|---|---|
|[<span data-ttu-id="38be1-363">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="38be1-363">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="38be1-364">1.0</span><span class="sxs-lookup"><span data-stu-id="38be1-364">1.0</span></span>|
|[<span data-ttu-id="38be1-365">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="38be1-365">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="38be1-366">ReadItem</span><span class="sxs-lookup"><span data-stu-id="38be1-366">ReadItem</span></span>|
|[<span data-ttu-id="38be1-367">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="38be1-367">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="38be1-368">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="38be1-368">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="38be1-369">Пример</span><span class="sxs-lookup"><span data-stu-id="38be1-369">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="38be1-370">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="38be1-370">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="38be1-371">Отображает форму для создания новой встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="38be1-371">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="38be1-372">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="38be1-372">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="38be1-p113">Метод `displayNewAppointmentForm` открывает форму, в которой пользователь может создать встречу или собрание. Если параметры заданы, поля формы встречи автоматически заполняются их содержимым.</span><span class="sxs-lookup"><span data-stu-id="38be1-p113">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="38be1-p114">В Outlook Web App и Outlook Web App для устройств этот метод всегда отображает форму с полем участников. Если вы не укажете участников в качестве входных аргументов, метод отображает форму с кнопкой **Сохранить**. Если вы укажете участников, форма будет включать участников и кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="38be1-p114">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="38be1-p115">Если вы укажете участников или ресурсы с помощью параметра `requiredAttendees`, `optionalAttendees` или `resources` в клиенте Outlook с расширенными возможностями и Outlook RT, этот метод отобразит форму собрания с кнопкой **Отправить**. Если не указать получателей, этот метод отобразит форму встречи с кнопкой **Сохранить и закрыть**.</span><span class="sxs-lookup"><span data-stu-id="38be1-p115">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="38be1-380">Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.</span><span class="sxs-lookup"><span data-stu-id="38be1-380">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="38be1-381">Parameters</span><span class="sxs-lookup"><span data-stu-id="38be1-381">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="38be1-382">Все параметры являются необязательными.</span><span class="sxs-lookup"><span data-stu-id="38be1-382">All parameters are optional.</span></span>

|<span data-ttu-id="38be1-383">Имя</span><span class="sxs-lookup"><span data-stu-id="38be1-383">Name</span></span>| <span data-ttu-id="38be1-384">Тип</span><span class="sxs-lookup"><span data-stu-id="38be1-384">Type</span></span>| <span data-ttu-id="38be1-385">Описание</span><span class="sxs-lookup"><span data-stu-id="38be1-385">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="38be1-386">Object</span><span class="sxs-lookup"><span data-stu-id="38be1-386">Object</span></span> | <span data-ttu-id="38be1-387">Словарь параметров, описывающий новую встречу.</span><span class="sxs-lookup"><span data-stu-id="38be1-387">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="38be1-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="38be1-388">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="38be1-p116">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из обязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="38be1-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="38be1-391">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="38be1-391">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="38be1-p117">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из необязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="38be1-p117">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="38be1-394">Date</span><span class="sxs-lookup"><span data-stu-id="38be1-394">Date</span></span> | <span data-ttu-id="38be1-395">Объект `Date`, указывающий дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="38be1-395">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="38be1-396">Date</span><span class="sxs-lookup"><span data-stu-id="38be1-396">Date</span></span> | <span data-ttu-id="38be1-397">Объект `Date`, указывающий дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="38be1-397">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="38be1-398">String</span><span class="sxs-lookup"><span data-stu-id="38be1-398">String</span></span> | <span data-ttu-id="38be1-p118">Строка со сведениями о месте встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="38be1-p118">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="38be1-401">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="38be1-401">Array.&lt;String&gt;</span></span> | <span data-ttu-id="38be1-p119">Массив строк, содержащий необходимые для встречи ресурсы. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="38be1-p119">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="38be1-404">String</span><span class="sxs-lookup"><span data-stu-id="38be1-404">String</span></span> | <span data-ttu-id="38be1-p120">Строка с темой встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="38be1-p120">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="38be1-407">String</span><span class="sxs-lookup"><span data-stu-id="38be1-407">String</span></span> | <span data-ttu-id="38be1-p121">Текст сообщения о встрече. Максимальный размер содержимого сообщения — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="38be1-p121">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="38be1-410">Требования</span><span class="sxs-lookup"><span data-stu-id="38be1-410">Requirements</span></span>

|<span data-ttu-id="38be1-411">Требование</span><span class="sxs-lookup"><span data-stu-id="38be1-411">Requirement</span></span>| <span data-ttu-id="38be1-412">Значение</span><span class="sxs-lookup"><span data-stu-id="38be1-412">Value</span></span>|
|---|---|
|[<span data-ttu-id="38be1-413">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="38be1-413">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="38be1-414">1.0</span><span class="sxs-lookup"><span data-stu-id="38be1-414">1.0</span></span>|
|[<span data-ttu-id="38be1-415">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="38be1-415">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="38be1-416">ReadItem</span><span class="sxs-lookup"><span data-stu-id="38be1-416">ReadItem</span></span>|
|[<span data-ttu-id="38be1-417">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="38be1-417">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="38be1-418">Чтение</span><span class="sxs-lookup"><span data-stu-id="38be1-418">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="38be1-419">Пример</span><span class="sxs-lookup"><span data-stu-id="38be1-419">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="38be1-420">displayNewMessageForm(параметры)</span><span class="sxs-lookup"><span data-stu-id="38be1-420">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="38be1-421">Отображает форму для создания сообщения.</span><span class="sxs-lookup"><span data-stu-id="38be1-421">Displays a form for creating a new message.</span></span>

<span data-ttu-id="38be1-422">Метод `displayNewMessageForm` открывает форму, в которой пользователь может создать сообщение.</span><span class="sxs-lookup"><span data-stu-id="38be1-422">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="38be1-423">Если параметры заданы, поля формы сообщения автоматически заполняются их содержимым.</span><span class="sxs-lookup"><span data-stu-id="38be1-423">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="38be1-424">Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.</span><span class="sxs-lookup"><span data-stu-id="38be1-424">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="38be1-425">Parameters</span><span class="sxs-lookup"><span data-stu-id="38be1-425">Parameters</span></span>

> [!NOTE]
> <span data-ttu-id="38be1-426">Все параметры являются необязательными.</span><span class="sxs-lookup"><span data-stu-id="38be1-426">All parameters are optional.</span></span>

|<span data-ttu-id="38be1-427">Имя</span><span class="sxs-lookup"><span data-stu-id="38be1-427">Name</span></span>| <span data-ttu-id="38be1-428">Тип</span><span class="sxs-lookup"><span data-stu-id="38be1-428">Type</span></span>| <span data-ttu-id="38be1-429">Описание</span><span class="sxs-lookup"><span data-stu-id="38be1-429">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="38be1-430">Объект</span><span class="sxs-lookup"><span data-stu-id="38be1-430">Object</span></span> | <span data-ttu-id="38be1-431">Словарь параметров, описывающий новое сообщение.</span><span class="sxs-lookup"><span data-stu-id="38be1-431">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="38be1-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="38be1-432">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="38be1-433">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из получателей, указанных в строке "Кому".</span><span class="sxs-lookup"><span data-stu-id="38be1-433">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="38be1-434">Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="38be1-434">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="38be1-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="38be1-435">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="38be1-436">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из получателей, указанных в строке "Копия".</span><span class="sxs-lookup"><span data-stu-id="38be1-436">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="38be1-437">Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="38be1-437">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="38be1-438">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="38be1-438">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_6/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="38be1-439">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из получателей, указанных в строке "Скрытая копия".</span><span class="sxs-lookup"><span data-stu-id="38be1-439">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="38be1-440">Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="38be1-440">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="38be1-441">String</span><span class="sxs-lookup"><span data-stu-id="38be1-441">String</span></span> | <span data-ttu-id="38be1-442">Строка с темой сообщения.</span><span class="sxs-lookup"><span data-stu-id="38be1-442">A string containing the subject of the message.</span></span> <span data-ttu-id="38be1-443">Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="38be1-443">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="38be1-444">String</span><span class="sxs-lookup"><span data-stu-id="38be1-444">String</span></span> | <span data-ttu-id="38be1-445">Текст сообщения в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="38be1-445">The HTML body of the message.</span></span> <span data-ttu-id="38be1-446">Максимальный размер содержимого сообщения — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="38be1-446">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="38be1-447">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="38be1-447">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="38be1-448">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="38be1-448">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="38be1-449">String</span><span class="sxs-lookup"><span data-stu-id="38be1-449">String</span></span> | <span data-ttu-id="38be1-p128">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="38be1-p128">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="38be1-452">Строка</span><span class="sxs-lookup"><span data-stu-id="38be1-452">String</span></span> | <span data-ttu-id="38be1-453">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="38be1-453">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="38be1-454">Строка</span><span class="sxs-lookup"><span data-stu-id="38be1-454">String</span></span> | <span data-ttu-id="38be1-p129">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="38be1-p129">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="38be1-457">Boolean</span><span class="sxs-lookup"><span data-stu-id="38be1-457">Boolean</span></span> | <span data-ttu-id="38be1-p130">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="38be1-p130">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="38be1-460">String</span><span class="sxs-lookup"><span data-stu-id="38be1-460">String</span></span> | <span data-ttu-id="38be1-461">Используется, только если для `type` задано значение `item`.</span><span class="sxs-lookup"><span data-stu-id="38be1-461">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="38be1-462">Идентификатор элемента EWS существующего письма, которое нужно вложить в новое сообщение.</span><span class="sxs-lookup"><span data-stu-id="38be1-462">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="38be1-463">Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="38be1-463">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="38be1-464">Требования</span><span class="sxs-lookup"><span data-stu-id="38be1-464">Requirements</span></span>

|<span data-ttu-id="38be1-465">Требование</span><span class="sxs-lookup"><span data-stu-id="38be1-465">Requirement</span></span>| <span data-ttu-id="38be1-466">Значение</span><span class="sxs-lookup"><span data-stu-id="38be1-466">Value</span></span>|
|---|---|
|[<span data-ttu-id="38be1-467">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="38be1-467">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="38be1-468">1.6</span><span class="sxs-lookup"><span data-stu-id="38be1-468">1.6</span></span> |
|[<span data-ttu-id="38be1-469">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="38be1-469">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="38be1-470">ReadItem</span><span class="sxs-lookup"><span data-stu-id="38be1-470">ReadItem</span></span>|
|[<span data-ttu-id="38be1-471">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="38be1-471">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="38be1-472">Чтение</span><span class="sxs-lookup"><span data-stu-id="38be1-472">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="38be1-473">Пример</span><span class="sxs-lookup"><span data-stu-id="38be1-473">Example</span></span>

```javascript
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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="38be1-474">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="38be1-474">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="38be1-475">Возвращает строку, содержащую маркер, который используется для вызова интерфейсов REST API или веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="38be1-475">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="38be1-p132">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="38be1-p132">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="38be1-478">Рекомендуем сделать так, чтобы по мере возможности надстройки использовали интерфейсы REST API, а не веб-службы Exchange.</span><span class="sxs-lookup"><span data-stu-id="38be1-478">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="38be1-479">**Маркеры REST**</span><span class="sxs-lookup"><span data-stu-id="38be1-479">**REST Tokens**</span></span>

<span data-ttu-id="38be1-p133">Если запрашивается маркер REST (`options.isRest = true`), полученный маркер не подойдет для проверки подлинности при вызовах веб-служб Exchange. Область действия маркера будет ограничена доступом только для чтения к текущему элементу и его вложениям, если в манифесте надстройки не указано разрешение [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission). Если указано разрешение `ReadWriteMailbox`, полученный маркер предоставит доступ на чтение и запись к почте, календарю и контактам, включая возможность отправки почты.</span><span class="sxs-lookup"><span data-stu-id="38be1-p133">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="38be1-483">С помощью свойства `restUrl` надстройка должна определить правильный URL-адрес для вызовов REST API.</span><span class="sxs-lookup"><span data-stu-id="38be1-483">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="38be1-484">**Маркеры EWS**</span><span class="sxs-lookup"><span data-stu-id="38be1-484">**EWS Tokens**</span></span>

<span data-ttu-id="38be1-p134">Если запрашивается маркер EWS (`options.isRest = false`), полученный маркер не подойдет для проверки подлинности при вызовах REST API. Область действия маркера будет ограничена доступом к текущему элементу.</span><span class="sxs-lookup"><span data-stu-id="38be1-p134">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="38be1-487">С помощью свойства `ewsUrl` надстройка должна определить правильный URL-адрес для вызовов EWS.</span><span class="sxs-lookup"><span data-stu-id="38be1-487">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="38be1-488">Параметры</span><span class="sxs-lookup"><span data-stu-id="38be1-488">Parameters</span></span>

|<span data-ttu-id="38be1-489">Имя</span><span class="sxs-lookup"><span data-stu-id="38be1-489">Name</span></span>| <span data-ttu-id="38be1-490">Тип</span><span class="sxs-lookup"><span data-stu-id="38be1-490">Type</span></span>| <span data-ttu-id="38be1-491">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="38be1-491">Attributes</span></span>| <span data-ttu-id="38be1-492">Описание</span><span class="sxs-lookup"><span data-stu-id="38be1-492">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="38be1-493">Object</span><span class="sxs-lookup"><span data-stu-id="38be1-493">Object</span></span> | <span data-ttu-id="38be1-494">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="38be1-494">&lt;optional&gt;</span></span> | <span data-ttu-id="38be1-495">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="38be1-495">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="38be1-496">Boolean</span><span class="sxs-lookup"><span data-stu-id="38be1-496">Boolean</span></span> |  <span data-ttu-id="38be1-497">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="38be1-497">&lt;optional&gt;</span></span> | <span data-ttu-id="38be1-p135">Определяет, будет ли предоставленный маркер использоваться для интерфейсов REST API Outlook или веб-служб Exchange. Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="38be1-p135">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="38be1-500">Object</span><span class="sxs-lookup"><span data-stu-id="38be1-500">Object</span></span> |  <span data-ttu-id="38be1-501">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="38be1-501">&lt;optional&gt;</span></span> | <span data-ttu-id="38be1-502">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="38be1-502">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="38be1-503">function</span><span class="sxs-lookup"><span data-stu-id="38be1-503">function</span></span>||<span data-ttu-id="38be1-p136">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult). Токен указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="38be1-p136">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="38be1-506">Требования</span><span class="sxs-lookup"><span data-stu-id="38be1-506">Requirements</span></span>

|<span data-ttu-id="38be1-507">Требование</span><span class="sxs-lookup"><span data-stu-id="38be1-507">Requirement</span></span>| <span data-ttu-id="38be1-508">Значение</span><span class="sxs-lookup"><span data-stu-id="38be1-508">Value</span></span>|
|---|---|
|[<span data-ttu-id="38be1-509">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="38be1-509">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="38be1-510">1.5</span><span class="sxs-lookup"><span data-stu-id="38be1-510">1.5</span></span> |
|[<span data-ttu-id="38be1-511">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="38be1-511">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="38be1-512">ReadItem</span><span class="sxs-lookup"><span data-stu-id="38be1-512">ReadItem</span></span>|
|[<span data-ttu-id="38be1-513">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="38be1-513">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="38be1-514">Создание и чтение</span><span class="sxs-lookup"><span data-stu-id="38be1-514">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="38be1-515">Пример</span><span class="sxs-lookup"><span data-stu-id="38be1-515">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="38be1-516">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="38be1-516">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="38be1-517">Получает строку, содержащую маркер, используемый для получения вложения или элемента с Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="38be1-517">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="38be1-p137">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный токен с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="38be1-p137">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="38be1-p138">Вы можете передать сторонней системе токен и идентификатор вложения или элемента. Сторонняя система использует этот токен как токен авторизации, чтобы вызвать операцию [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) или [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) веб-служб Exchange для возврата вложения или элемента. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="38be1-p138">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="38be1-523">Для вызова метода `getCallbackTokenAsync` в режиме чтения манифесте приложения должно быть указано разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="38be1-523">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="38be1-p139">Чтобы получить идентификатор элемента для передачи в метод `getCallbackTokenAsync`, в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="38be1-p139">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="38be1-526">Параметры</span><span class="sxs-lookup"><span data-stu-id="38be1-526">Parameters</span></span>

|<span data-ttu-id="38be1-527">Имя</span><span class="sxs-lookup"><span data-stu-id="38be1-527">Name</span></span>| <span data-ttu-id="38be1-528">Тип</span><span class="sxs-lookup"><span data-stu-id="38be1-528">Type</span></span>| <span data-ttu-id="38be1-529">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="38be1-529">Attributes</span></span>| <span data-ttu-id="38be1-530">Описание</span><span class="sxs-lookup"><span data-stu-id="38be1-530">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="38be1-531">функция</span><span class="sxs-lookup"><span data-stu-id="38be1-531">function</span></span>||<span data-ttu-id="38be1-p140">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult). Токен указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="38be1-p140">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="38be1-534">Object</span><span class="sxs-lookup"><span data-stu-id="38be1-534">Object</span></span>| <span data-ttu-id="38be1-535">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="38be1-535">&lt;optional&gt;</span></span>|<span data-ttu-id="38be1-536">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="38be1-536">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="38be1-537">Требования</span><span class="sxs-lookup"><span data-stu-id="38be1-537">Requirements</span></span>

|<span data-ttu-id="38be1-538">Требование</span><span class="sxs-lookup"><span data-stu-id="38be1-538">Requirement</span></span>| <span data-ttu-id="38be1-539">Значение</span><span class="sxs-lookup"><span data-stu-id="38be1-539">Value</span></span>|
|---|---|
|[<span data-ttu-id="38be1-540">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="38be1-540">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="38be1-541">1.3</span><span class="sxs-lookup"><span data-stu-id="38be1-541">1.3</span></span>|
|[<span data-ttu-id="38be1-542">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="38be1-542">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="38be1-543">ReadItem</span><span class="sxs-lookup"><span data-stu-id="38be1-543">ReadItem</span></span>|
|[<span data-ttu-id="38be1-544">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="38be1-544">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="38be1-545">Создание и чтение</span><span class="sxs-lookup"><span data-stu-id="38be1-545">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="38be1-546">Пример</span><span class="sxs-lookup"><span data-stu-id="38be1-546">Example</span></span>

```javascript
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="38be1-547">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="38be1-547">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="38be1-548">Получает маркер, идентифицирующий пользователя и надстройку Office.</span><span class="sxs-lookup"><span data-stu-id="38be1-548">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="38be1-549">Метод `getUserIdentityTokenAsync` возвращает токен, который можно использовать для идентификации, а также [проверки подлинности надстройки и пользователя в сторонней системе](https://docs.microsoft.com/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="38be1-549">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="38be1-550">Параметры</span><span class="sxs-lookup"><span data-stu-id="38be1-550">Parameters</span></span>

|<span data-ttu-id="38be1-551">Имя</span><span class="sxs-lookup"><span data-stu-id="38be1-551">Name</span></span>| <span data-ttu-id="38be1-552">Тип</span><span class="sxs-lookup"><span data-stu-id="38be1-552">Type</span></span>| <span data-ttu-id="38be1-553">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="38be1-553">Attributes</span></span>| <span data-ttu-id="38be1-554">Описание</span><span class="sxs-lookup"><span data-stu-id="38be1-554">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="38be1-555">функция</span><span class="sxs-lookup"><span data-stu-id="38be1-555">function</span></span>||<span data-ttu-id="38be1-556">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="38be1-556">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="38be1-557">Токен указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="38be1-557">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="38be1-558">Object</span><span class="sxs-lookup"><span data-stu-id="38be1-558">Object</span></span>| <span data-ttu-id="38be1-559">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="38be1-559">&lt;optional&gt;</span></span>|<span data-ttu-id="38be1-560">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="38be1-560">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="38be1-561">Требования</span><span class="sxs-lookup"><span data-stu-id="38be1-561">Requirements</span></span>

|<span data-ttu-id="38be1-562">Требование</span><span class="sxs-lookup"><span data-stu-id="38be1-562">Requirement</span></span>| <span data-ttu-id="38be1-563">Значение</span><span class="sxs-lookup"><span data-stu-id="38be1-563">Value</span></span>|
|---|---|
|[<span data-ttu-id="38be1-564">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="38be1-564">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="38be1-565">1.0</span><span class="sxs-lookup"><span data-stu-id="38be1-565">1.0</span></span>|
|[<span data-ttu-id="38be1-566">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="38be1-566">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="38be1-567">ReadItem</span><span class="sxs-lookup"><span data-stu-id="38be1-567">ReadItem</span></span>|
|[<span data-ttu-id="38be1-568">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="38be1-568">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="38be1-569">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="38be1-569">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="38be1-570">Пример</span><span class="sxs-lookup"><span data-stu-id="38be1-570">Example</span></span>

```javascript
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="38be1-571">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="38be1-571">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="38be1-572">Выполняет асинхронный запрос для веб-служб Exchange (EWS) на сервере Exchange Server, на котором размещен почтовый ящик пользователя.</span><span class="sxs-lookup"><span data-stu-id="38be1-572">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="38be1-573">Этот метод не поддерживается в следующих сценариях:</span><span class="sxs-lookup"><span data-stu-id="38be1-573">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="38be1-574">В Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="38be1-574">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="38be1-575">Если надстройка загружается в почтовый ящик Gmail.</span><span class="sxs-lookup"><span data-stu-id="38be1-575">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="38be1-576">В таких случаях надстройка должна [использовать REST API](https://docs.microsoft.com/outlook/add-ins/use-rest-api) для доступа к почтовому ящику пользователя.</span><span class="sxs-lookup"><span data-stu-id="38be1-576">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="38be1-577">Метод `makeEwsRequestAsync` отправляет запрос EWS от имени надстройки в Exchange.</span><span class="sxs-lookup"><span data-stu-id="38be1-577">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="38be1-578">Список поддерживаемых операций EWS см. в статье [Вызов веб-служб из надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support).</span><span class="sxs-lookup"><span data-stu-id="38be1-578">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="38be1-579">С помощью метода `makeEwsRequestAsync` невозможно запрашивать элементы, связанные с папкой.</span><span class="sxs-lookup"><span data-stu-id="38be1-579">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="38be1-580">В запросе XML должна быть указана кодировка UTF-8.</span><span class="sxs-lookup"><span data-stu-id="38be1-580">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="38be1-p142">У вашей надстройки должно быть разрешение **ReadWriteMailbox** для использования метода `makeEwsRequestAsync`. Сведения об использовании разрешения **ReadWriteMailbox** и операций EWS, которые можно вызывать с помощью метода `makeEwsRequestAsync`, см. в статье [Указание разрешений для доступа почтовой надстройки к почтовому ящику пользователя](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="38be1-p142">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="38be1-583">Администратор сервера должен установить значение true для параметра `OAuthAuthentication` в каталоге сервера клиентского доступа EWS, чтобы метод `makeEwsRequestAsync` мог выполнять запросы EWS.</span><span class="sxs-lookup"><span data-stu-id="38be1-583">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="38be1-584">Различия версий</span><span class="sxs-lookup"><span data-stu-id="38be1-584">Version differences</span></span>

<span data-ttu-id="38be1-585">Если вы используете метод `makeEwsRequestAsync` в почтовых приложениях, которые выполняются в Outlook версии более ранней, чем 15.0.4535.1004, указывайте кодировку `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="38be1-585">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="38be1-p143">Значение кодировки не нужно указывать, если почтовое приложение выполняется в Outlook в Интернете. Чтобы определить, выполняется ли приложение в Outlook или Outlook в Интернете, используйте свойство mailbox.diagnostics.hostName. Используемую версию Outlook можно определить с помощью свойства mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="38be1-p143">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="38be1-589">Параметры</span><span class="sxs-lookup"><span data-stu-id="38be1-589">Parameters</span></span>

|<span data-ttu-id="38be1-590">Имя</span><span class="sxs-lookup"><span data-stu-id="38be1-590">Name</span></span>| <span data-ttu-id="38be1-591">Тип</span><span class="sxs-lookup"><span data-stu-id="38be1-591">Type</span></span>| <span data-ttu-id="38be1-592">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="38be1-592">Attributes</span></span>| <span data-ttu-id="38be1-593">Описание</span><span class="sxs-lookup"><span data-stu-id="38be1-593">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="38be1-594">String</span><span class="sxs-lookup"><span data-stu-id="38be1-594">String</span></span>||<span data-ttu-id="38be1-595">Запрос EWS.</span><span class="sxs-lookup"><span data-stu-id="38be1-595">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="38be1-596">функция</span><span class="sxs-lookup"><span data-stu-id="38be1-596">function</span></span>||<span data-ttu-id="38be1-597">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="38be1-597">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="38be1-598">Результат XML вызова EWS указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="38be1-598">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="38be1-599">Если размер результата превышает 1 МБ, возвращается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="38be1-599">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="38be1-600">Объект</span><span class="sxs-lookup"><span data-stu-id="38be1-600">Object</span></span>| <span data-ttu-id="38be1-601">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="38be1-601">&lt;optional&gt;</span></span>|<span data-ttu-id="38be1-602">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="38be1-602">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="38be1-603">Требования</span><span class="sxs-lookup"><span data-stu-id="38be1-603">Requirements</span></span>

|<span data-ttu-id="38be1-604">Требование</span><span class="sxs-lookup"><span data-stu-id="38be1-604">Requirement</span></span>| <span data-ttu-id="38be1-605">Значение</span><span class="sxs-lookup"><span data-stu-id="38be1-605">Value</span></span>|
|---|---|
|[<span data-ttu-id="38be1-606">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="38be1-606">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="38be1-607">1.0</span><span class="sxs-lookup"><span data-stu-id="38be1-607">1.0</span></span>|
|[<span data-ttu-id="38be1-608">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="38be1-608">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="38be1-609">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="38be1-609">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="38be1-610">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="38be1-610">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="38be1-611">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="38be1-611">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="38be1-612">Пример</span><span class="sxs-lookup"><span data-stu-id="38be1-612">Example</span></span>

<span data-ttu-id="38be1-613">В приведенном ниже примере вызывается `makeEwsRequestAsync` для получения темы элемента с помощью операции `GetItem`.</span><span class="sxs-lookup"><span data-stu-id="38be1-613">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="38be1-614">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="38be1-614">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="38be1-615">Удаляет обработчиков для поддерживаемого типа события.</span><span class="sxs-lookup"><span data-stu-id="38be1-615">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="38be1-616">В настоящее время единственный поддерживаемый тип события — `Office.EventType.ItemChanged`.</span><span class="sxs-lookup"><span data-stu-id="38be1-616">Currently, the only supported event type is `Office.EventType.ItemChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="38be1-617">Параметры</span><span class="sxs-lookup"><span data-stu-id="38be1-617">Parameters</span></span>

| <span data-ttu-id="38be1-618">Имя</span><span class="sxs-lookup"><span data-stu-id="38be1-618">Name</span></span> | <span data-ttu-id="38be1-619">Тип</span><span class="sxs-lookup"><span data-stu-id="38be1-619">Type</span></span> | <span data-ttu-id="38be1-620">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="38be1-620">Attributes</span></span> | <span data-ttu-id="38be1-621">Описание</span><span class="sxs-lookup"><span data-stu-id="38be1-621">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="38be1-622">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="38be1-622">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="38be1-623">Событие, которое должно отменить обработчик.</span><span class="sxs-lookup"><span data-stu-id="38be1-623">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="38be1-624">Объект</span><span class="sxs-lookup"><span data-stu-id="38be1-624">Object</span></span> | <span data-ttu-id="38be1-625">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="38be1-625">&lt;optional&gt;</span></span> | <span data-ttu-id="38be1-626">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="38be1-626">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="38be1-627">Object</span><span class="sxs-lookup"><span data-stu-id="38be1-627">Object</span></span> | <span data-ttu-id="38be1-628">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="38be1-628">&lt;optional&gt;</span></span> | <span data-ttu-id="38be1-629">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="38be1-629">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="38be1-630">функция</span><span class="sxs-lookup"><span data-stu-id="38be1-630">function</span></span>| <span data-ttu-id="38be1-631">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="38be1-631">&lt;optional&gt;</span></span>|<span data-ttu-id="38be1-632">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="38be1-632">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="38be1-633">Требования</span><span class="sxs-lookup"><span data-stu-id="38be1-633">Requirements</span></span>

|<span data-ttu-id="38be1-634">Требование</span><span class="sxs-lookup"><span data-stu-id="38be1-634">Requirement</span></span>| <span data-ttu-id="38be1-635">Значение</span><span class="sxs-lookup"><span data-stu-id="38be1-635">Value</span></span>|
|---|---|
|[<span data-ttu-id="38be1-636">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="38be1-636">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="38be1-637">1.5</span><span class="sxs-lookup"><span data-stu-id="38be1-637">1.5</span></span> |
|[<span data-ttu-id="38be1-638">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="38be1-638">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="38be1-639">ReadItem</span><span class="sxs-lookup"><span data-stu-id="38be1-639">ReadItem</span></span> |
|[<span data-ttu-id="38be1-640">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="38be1-640">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="38be1-641">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="38be1-641">Compose or Read</span></span>|
