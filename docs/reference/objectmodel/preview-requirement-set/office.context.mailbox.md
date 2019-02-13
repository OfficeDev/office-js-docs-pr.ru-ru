---
title: Office.context.mailbox — предварительная версия набора обязательных элементов
description: ''
ms.date: 01/16/2019
localization_priority: Normal
ms.openlocfilehash: a1b6c66f34cebe936614ff3c37a888a12b9e21b6
ms.sourcegitcommit: a59f4e322238efa187f388a75b7709462c71e668
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/13/2019
ms.locfileid: "29388222"
---
# <a name="mailbox"></a><span data-ttu-id="62aab-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="62aab-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="62aab-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="62aab-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="62aab-104">Предоставляет для Microsoft Outlook и Microsoft Outlook в Интернете доступ к объектной модели надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="62aab-104">Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="62aab-105">Требования</span><span class="sxs-lookup"><span data-stu-id="62aab-105">Requirements</span></span>

|<span data-ttu-id="62aab-106">Требование</span><span class="sxs-lookup"><span data-stu-id="62aab-106">Requirement</span></span>| <span data-ttu-id="62aab-107">Значение</span><span class="sxs-lookup"><span data-stu-id="62aab-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="62aab-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="62aab-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="62aab-109">1.0</span><span class="sxs-lookup"><span data-stu-id="62aab-109">1.0</span></span>|
|[<span data-ttu-id="62aab-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="62aab-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="62aab-111">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="62aab-111">Restricted</span></span>|
|[<span data-ttu-id="62aab-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="62aab-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="62aab-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="62aab-113">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="62aab-114">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="62aab-114">Members and methods</span></span>

| <span data-ttu-id="62aab-115">Элемент</span><span class="sxs-lookup"><span data-stu-id="62aab-115">Member</span></span> | <span data-ttu-id="62aab-116">Тип</span><span class="sxs-lookup"><span data-stu-id="62aab-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="62aab-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="62aab-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="62aab-118">Элемент</span><span class="sxs-lookup"><span data-stu-id="62aab-118">Member</span></span> |
| [<span data-ttu-id="62aab-119">restUrl</span><span class="sxs-lookup"><span data-stu-id="62aab-119">restUrl</span></span>](#resturl-string) | <span data-ttu-id="62aab-120">Элемент</span><span class="sxs-lookup"><span data-stu-id="62aab-120">Member</span></span> |
| [<span data-ttu-id="62aab-121">addHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="62aab-121">addHandlerAsync</span></span>](#addhandlerasynceventtype-handler-options-callback) | <span data-ttu-id="62aab-122">Метод</span><span class="sxs-lookup"><span data-stu-id="62aab-122">Method</span></span> |
| [<span data-ttu-id="62aab-123">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="62aab-123">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="62aab-124">Метод</span><span class="sxs-lookup"><span data-stu-id="62aab-124">Method</span></span> |
| [<span data-ttu-id="62aab-125">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="62aab-125">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime) | <span data-ttu-id="62aab-126">Метод</span><span class="sxs-lookup"><span data-stu-id="62aab-126">Method</span></span> |
| [<span data-ttu-id="62aab-127">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="62aab-127">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="62aab-128">Метод</span><span class="sxs-lookup"><span data-stu-id="62aab-128">Method</span></span> |
| [<span data-ttu-id="62aab-129">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="62aab-129">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="62aab-130">Метод</span><span class="sxs-lookup"><span data-stu-id="62aab-130">Method</span></span> |
| [<span data-ttu-id="62aab-131">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="62aab-131">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="62aab-132">Метод</span><span class="sxs-lookup"><span data-stu-id="62aab-132">Method</span></span> |
| [<span data-ttu-id="62aab-133">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="62aab-133">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="62aab-134">Метод</span><span class="sxs-lookup"><span data-stu-id="62aab-134">Method</span></span> |
| [<span data-ttu-id="62aab-135">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="62aab-135">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="62aab-136">Метод</span><span class="sxs-lookup"><span data-stu-id="62aab-136">Method</span></span> |
| [<span data-ttu-id="62aab-137">displayNewMessageForm</span><span class="sxs-lookup"><span data-stu-id="62aab-137">displayNewMessageForm</span></span>](#displaynewmessageformparameters) | <span data-ttu-id="62aab-138">Метод</span><span class="sxs-lookup"><span data-stu-id="62aab-138">Method</span></span> |
| [<span data-ttu-id="62aab-139">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="62aab-139">getCallbackTokenAsync</span></span>](#getcallbacktokenasyncoptions-callback) | <span data-ttu-id="62aab-140">Метод</span><span class="sxs-lookup"><span data-stu-id="62aab-140">Method</span></span> |
| [<span data-ttu-id="62aab-141">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="62aab-141">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="62aab-142">Метод</span><span class="sxs-lookup"><span data-stu-id="62aab-142">Method</span></span> |
| [<span data-ttu-id="62aab-143">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="62aab-143">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="62aab-144">Метод</span><span class="sxs-lookup"><span data-stu-id="62aab-144">Method</span></span> |
| [<span data-ttu-id="62aab-145">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="62aab-145">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="62aab-146">Метод</span><span class="sxs-lookup"><span data-stu-id="62aab-146">Method</span></span> |
| [<span data-ttu-id="62aab-147">removeHandlerAsync</span><span class="sxs-lookup"><span data-stu-id="62aab-147">removeHandlerAsync</span></span>](#removehandlerasynceventtype-options-callback) | <span data-ttu-id="62aab-148">Метод</span><span class="sxs-lookup"><span data-stu-id="62aab-148">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="62aab-149">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="62aab-149">Namespaces</span></span>

<span data-ttu-id="62aab-150">[diagnostics](Office.context.mailbox.diagnostics.md). Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="62aab-150">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="62aab-151">[item](Office.context.mailbox.item.md). Предоставляет методы и свойства для доступа к сообщению или встрече в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="62aab-151">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="62aab-152">[userProfile](Office.context.mailbox.userProfile.md). Предоставляет сведения о пользователе в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="62aab-152">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="62aab-153">Элементы</span><span class="sxs-lookup"><span data-stu-id="62aab-153">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="62aab-154">ewsUrl :String</span><span class="sxs-lookup"><span data-stu-id="62aab-154">ewsUrl :String</span></span>

<span data-ttu-id="62aab-p101">Получает URL-адрес конечной точки веб-служб Exchange (EWS) для этой учетной записи электронной почты. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="62aab-p101">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="62aab-157">Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="62aab-157">This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="62aab-p102">Удаленная служба может использовать значение `ewsUrl`, чтобы выполнять вызовы EWS для почтового ящика пользователя. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="62aab-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="62aab-160">Чтобы вызвать элемент `ewsUrl` в режиме чтения, в манифесте приложения должно быть указано разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="62aab-160">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="62aab-p103">Перед использованием элемента `ewsUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="62aab-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="62aab-163">Тип:</span><span class="sxs-lookup"><span data-stu-id="62aab-163">Type:</span></span>

*   <span data-ttu-id="62aab-164">String</span><span class="sxs-lookup"><span data-stu-id="62aab-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="62aab-165">Требования</span><span class="sxs-lookup"><span data-stu-id="62aab-165">Requirements</span></span>

|<span data-ttu-id="62aab-166">Требование</span><span class="sxs-lookup"><span data-stu-id="62aab-166">Requirement</span></span>| <span data-ttu-id="62aab-167">Значение</span><span class="sxs-lookup"><span data-stu-id="62aab-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="62aab-168">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="62aab-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="62aab-169">1.0</span><span class="sxs-lookup"><span data-stu-id="62aab-169">1.0</span></span>|
|[<span data-ttu-id="62aab-170">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="62aab-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="62aab-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="62aab-171">ReadItem</span></span>|
|[<span data-ttu-id="62aab-172">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="62aab-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="62aab-173">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="62aab-173">Compose or read</span></span>|

#### <a name="resturl-string"></a><span data-ttu-id="62aab-174">restUrl :String</span><span class="sxs-lookup"><span data-stu-id="62aab-174">restUrl :String</span></span>

<span data-ttu-id="62aab-175">Возвращает URL-адрес конечной точки REST для этой учетной записи электронной почты.</span><span class="sxs-lookup"><span data-stu-id="62aab-175">Gets the URL of the REST endpoint for this email account.</span></span>

<span data-ttu-id="62aab-176">С помощью значения `restUrl` можно выполнять вызовы [REST API](https://docs.microsoft.com/outlook/rest/) для почтового ящика пользователя.</span><span class="sxs-lookup"><span data-stu-id="62aab-176">The `restUrl` value can be used to make [REST API](https://docs.microsoft.com/outlook/rest/) calls to the user's mailbox.</span></span>

<span data-ttu-id="62aab-177">Чтобы вызвать элемент `restUrl` в режиме чтения, в манифесте приложения необходимо указать разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="62aab-177">Your app must have the **ReadItem** permission specified in its manifest to call the `restUrl` member in read mode.</span></span>

<span data-ttu-id="62aab-p104">Перед использованием элемента `restUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="62aab-p104">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `restUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="62aab-180">Тип:</span><span class="sxs-lookup"><span data-stu-id="62aab-180">Type:</span></span>

*   <span data-ttu-id="62aab-181">String</span><span class="sxs-lookup"><span data-stu-id="62aab-181">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="62aab-182">Требования</span><span class="sxs-lookup"><span data-stu-id="62aab-182">Requirements</span></span>

|<span data-ttu-id="62aab-183">Требование</span><span class="sxs-lookup"><span data-stu-id="62aab-183">Requirement</span></span>| <span data-ttu-id="62aab-184">Значение</span><span class="sxs-lookup"><span data-stu-id="62aab-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="62aab-185">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="62aab-185">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="62aab-186">1.5</span><span class="sxs-lookup"><span data-stu-id="62aab-186">1.5</span></span> |
|[<span data-ttu-id="62aab-187">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="62aab-187">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="62aab-188">ReadItem</span><span class="sxs-lookup"><span data-stu-id="62aab-188">ReadItem</span></span>|
|[<span data-ttu-id="62aab-189">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="62aab-189">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="62aab-190">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="62aab-190">Compose or read</span></span>|

### <a name="methods"></a><span data-ttu-id="62aab-191">Методы</span><span class="sxs-lookup"><span data-stu-id="62aab-191">Methods</span></span>

####  <a name="addhandlerasynceventtype-handler-options-callback"></a><span data-ttu-id="62aab-192">addHandlerAsync(eventType, handler, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="62aab-192">addHandlerAsync(eventType, handler, [options], [callback])</span></span>

<span data-ttu-id="62aab-193">Добавляет обработчик для поддерживаемого события.</span><span class="sxs-lookup"><span data-stu-id="62aab-193">Adds an event handler for a supported event.</span></span>

<span data-ttu-id="62aab-194">В настоящее время поддерживаются следующие типы событий: `Office.EventType.ItemChanged` и `Office.EventType.OfficeThemeChanged`.</span><span class="sxs-lookup"><span data-stu-id="62aab-194">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="62aab-195">Параметры</span><span class="sxs-lookup"><span data-stu-id="62aab-195">Parameters:</span></span>

| <span data-ttu-id="62aab-196">Имя</span><span class="sxs-lookup"><span data-stu-id="62aab-196">Name</span></span> | <span data-ttu-id="62aab-197">Тип</span><span class="sxs-lookup"><span data-stu-id="62aab-197">Type</span></span> | <span data-ttu-id="62aab-198">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="62aab-198">Attributes</span></span> | <span data-ttu-id="62aab-199">Описание</span><span class="sxs-lookup"><span data-stu-id="62aab-199">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="62aab-200">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="62aab-200">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="62aab-201">Событие, которое должно вызвать обработчик.</span><span class="sxs-lookup"><span data-stu-id="62aab-201">The event that should invoke the handler.</span></span> |
| `handler` | <span data-ttu-id="62aab-202">Function</span><span class="sxs-lookup"><span data-stu-id="62aab-202">Function</span></span> || <span data-ttu-id="62aab-p105">Функция для обработки события. Функция должна принимать один параметр, представляющий собой объектный литерал. Значение свойства `type` параметра совпадет со значением параметра `eventType`, переданного методу `addHandlerAsync`.</span><span class="sxs-lookup"><span data-stu-id="62aab-p105">The function to handle the event. The function must accept a single parameter, which is an object literal. The `type` property on the parameter will match the `eventType` parameter passed to `addHandlerAsync`.</span></span> |
| `options` | <span data-ttu-id="62aab-206">Объект</span><span class="sxs-lookup"><span data-stu-id="62aab-206">Object</span></span> | <span data-ttu-id="62aab-207">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="62aab-207">&lt;optional&gt;</span></span> | <span data-ttu-id="62aab-208">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="62aab-208">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="62aab-209">Object</span><span class="sxs-lookup"><span data-stu-id="62aab-209">Object</span></span> | <span data-ttu-id="62aab-210">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="62aab-210">&lt;optional&gt;</span></span> | <span data-ttu-id="62aab-211">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="62aab-211">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="62aab-212">функция</span><span class="sxs-lookup"><span data-stu-id="62aab-212">function</span></span>| <span data-ttu-id="62aab-213">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="62aab-213">&lt;optional&gt;</span></span>|<span data-ttu-id="62aab-214">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="62aab-214">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="62aab-215">Требования</span><span class="sxs-lookup"><span data-stu-id="62aab-215">Requirements</span></span>

|<span data-ttu-id="62aab-216">Требование</span><span class="sxs-lookup"><span data-stu-id="62aab-216">Requirement</span></span>| <span data-ttu-id="62aab-217">Значение</span><span class="sxs-lookup"><span data-stu-id="62aab-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="62aab-218">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="62aab-218">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="62aab-219">1.5</span><span class="sxs-lookup"><span data-stu-id="62aab-219">1.5</span></span> |
|[<span data-ttu-id="62aab-220">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="62aab-220">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="62aab-221">ReadItem</span><span class="sxs-lookup"><span data-stu-id="62aab-221">ReadItem</span></span> |
|[<span data-ttu-id="62aab-222">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="62aab-222">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="62aab-223">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="62aab-223">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="62aab-224">Пример</span><span class="sxs-lookup"><span data-stu-id="62aab-224">Example</span></span>

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

####  <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="62aab-225">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="62aab-225">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="62aab-226">Преобразовывает идентификатор элемента из формата REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="62aab-226">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="62aab-227">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="62aab-227">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="62aab-p106">Формат идентификаторов, извлекаемых через API REST (например, [API Почты Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)), отличается от формата веб-служб Exchange (EWS). Метод `convertToEwsId` преобразовывает идентификатор в формате REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="62aab-p106">Item IDs retrieved via a REST API (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="62aab-230">Параметры:</span><span class="sxs-lookup"><span data-stu-id="62aab-230">Parameters:</span></span>

|<span data-ttu-id="62aab-231">Имя</span><span class="sxs-lookup"><span data-stu-id="62aab-231">Name</span></span>| <span data-ttu-id="62aab-232">Тип</span><span class="sxs-lookup"><span data-stu-id="62aab-232">Type</span></span>| <span data-ttu-id="62aab-233">Описание</span><span class="sxs-lookup"><span data-stu-id="62aab-233">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="62aab-234">String</span><span class="sxs-lookup"><span data-stu-id="62aab-234">String</span></span>|<span data-ttu-id="62aab-235">Идентификатор элемента в формате REST API для Outlook</span><span class="sxs-lookup"><span data-stu-id="62aab-235">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="62aab-236">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="62aab-236">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="62aab-237">Значение, определяющее версию REST API для Outlook, которая используется для извлечения идентификатора элемента.</span><span class="sxs-lookup"><span data-stu-id="62aab-237">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="62aab-238">Требования</span><span class="sxs-lookup"><span data-stu-id="62aab-238">Requirements</span></span>

|<span data-ttu-id="62aab-239">Требование</span><span class="sxs-lookup"><span data-stu-id="62aab-239">Requirement</span></span>| <span data-ttu-id="62aab-240">Значение</span><span class="sxs-lookup"><span data-stu-id="62aab-240">Value</span></span>|
|---|---|
|[<span data-ttu-id="62aab-241">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="62aab-241">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="62aab-242">1.3</span><span class="sxs-lookup"><span data-stu-id="62aab-242">1.3</span></span>|
|[<span data-ttu-id="62aab-243">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="62aab-243">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="62aab-244">Restricted</span><span class="sxs-lookup"><span data-stu-id="62aab-244">Restricted</span></span>|
|[<span data-ttu-id="62aab-245">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="62aab-245">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="62aab-246">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="62aab-246">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="62aab-247">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="62aab-247">Returns:</span></span>

<span data-ttu-id="62aab-248">Тип: String</span><span class="sxs-lookup"><span data-stu-id="62aab-248">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="62aab-249">Пример</span><span class="sxs-lookup"><span data-stu-id="62aab-249">Example</span></span>

```js
// Get an item's ID from a REST API
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the
// Outlook Mail API
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttime"></a><span data-ttu-id="62aab-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="62aab-250">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime)}</span></span>

<span data-ttu-id="62aab-251">Получает словарь, содержащий сведения о локальном времени клиента.</span><span class="sxs-lookup"><span data-stu-id="62aab-251">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="62aab-p107">В случае дат и времени в почтовом приложении для Outlook или Outlook Web App могут использоваться разные часовые пояса. Outlook использует часовой пояс клиентского компьютера. Outlook Web App использует часовой пояс, заданный в Центре администрирования Exchange (EAC). Значения даты и времени должны обрабатываться так, чтобы значения в пользовательском интерфейсе всегда согласовывались с часовым поясом, ожидаемым пользователем.</span><span class="sxs-lookup"><span data-stu-id="62aab-p107">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="62aab-p108">Если почтовое приложение работает в Outlook, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса клиентского компьютера. Если почтовое приложение работает в Outlook Web App, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса, заданного в Центре администрирования Exchange.</span><span class="sxs-lookup"><span data-stu-id="62aab-p108">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="62aab-257">Параметры:</span><span class="sxs-lookup"><span data-stu-id="62aab-257">Parameters:</span></span>

|<span data-ttu-id="62aab-258">Имя</span><span class="sxs-lookup"><span data-stu-id="62aab-258">Name</span></span>| <span data-ttu-id="62aab-259">Тип</span><span class="sxs-lookup"><span data-stu-id="62aab-259">Type</span></span>| <span data-ttu-id="62aab-260">Описание</span><span class="sxs-lookup"><span data-stu-id="62aab-260">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="62aab-261">Date</span><span class="sxs-lookup"><span data-stu-id="62aab-261">Date</span></span>|<span data-ttu-id="62aab-262">Объект Date</span><span class="sxs-lookup"><span data-stu-id="62aab-262">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="62aab-263">Требования</span><span class="sxs-lookup"><span data-stu-id="62aab-263">Requirements</span></span>

|<span data-ttu-id="62aab-264">Требование</span><span class="sxs-lookup"><span data-stu-id="62aab-264">Requirement</span></span>| <span data-ttu-id="62aab-265">Значение</span><span class="sxs-lookup"><span data-stu-id="62aab-265">Value</span></span>|
|---|---|
|[<span data-ttu-id="62aab-266">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="62aab-266">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="62aab-267">1.0</span><span class="sxs-lookup"><span data-stu-id="62aab-267">1.0</span></span>|
|[<span data-ttu-id="62aab-268">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="62aab-268">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="62aab-269">ReadItem</span><span class="sxs-lookup"><span data-stu-id="62aab-269">ReadItem</span></span>|
|[<span data-ttu-id="62aab-270">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="62aab-270">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="62aab-271">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="62aab-271">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="62aab-272">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="62aab-272">Returns:</span></span>

<span data-ttu-id="62aab-273">Тип: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="62aab-273">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime)</span></span>

####  <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="62aab-274">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="62aab-274">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="62aab-275">Преобразовывает идентификатор элемента в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="62aab-275">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="62aab-276">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="62aab-276">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="62aab-p109">Формат идентификаторов, извлекаемых через EWS или свойство `itemId`, отличается от формата API REST (таких как [API Почты Outlook](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)). Метод `convertToRestId` преобразовывает идентификатор в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="62aab-p109">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](https://docs.microsoft.com/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="62aab-279">Параметры:</span><span class="sxs-lookup"><span data-stu-id="62aab-279">Parameters:</span></span>

|<span data-ttu-id="62aab-280">Имя</span><span class="sxs-lookup"><span data-stu-id="62aab-280">Name</span></span>| <span data-ttu-id="62aab-281">Тип</span><span class="sxs-lookup"><span data-stu-id="62aab-281">Type</span></span>| <span data-ttu-id="62aab-282">Описание</span><span class="sxs-lookup"><span data-stu-id="62aab-282">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="62aab-283">String</span><span class="sxs-lookup"><span data-stu-id="62aab-283">String</span></span>|<span data-ttu-id="62aab-284">Идентификатор элемента в формате EWS</span><span class="sxs-lookup"><span data-stu-id="62aab-284">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="62aab-285">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="62aab-285">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion)|<span data-ttu-id="62aab-286">Значение, определяющее версию REST API для Outlook, с которой будет использоваться преобразованный идентификатор.</span><span class="sxs-lookup"><span data-stu-id="62aab-286">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="62aab-287">Требования</span><span class="sxs-lookup"><span data-stu-id="62aab-287">Requirements</span></span>

|<span data-ttu-id="62aab-288">Требование</span><span class="sxs-lookup"><span data-stu-id="62aab-288">Requirement</span></span>| <span data-ttu-id="62aab-289">Значение</span><span class="sxs-lookup"><span data-stu-id="62aab-289">Value</span></span>|
|---|---|
|[<span data-ttu-id="62aab-290">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="62aab-290">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="62aab-291">1.3</span><span class="sxs-lookup"><span data-stu-id="62aab-291">1.3</span></span>|
|[<span data-ttu-id="62aab-292">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="62aab-292">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="62aab-293">Restricted</span><span class="sxs-lookup"><span data-stu-id="62aab-293">Restricted</span></span>|
|[<span data-ttu-id="62aab-294">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="62aab-294">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="62aab-295">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="62aab-295">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="62aab-296">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="62aab-296">Returns:</span></span>

<span data-ttu-id="62aab-297">Тип: String</span><span class="sxs-lookup"><span data-stu-id="62aab-297">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="62aab-298">Пример</span><span class="sxs-lookup"><span data-stu-id="62aab-298">Example</span></span>

```js
// Get the currently selected item's ID
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the
// Outlook Mail API
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="62aab-299">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="62aab-299">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="62aab-300">Получает объект Date из словаря, содержащего сведения о времени.</span><span class="sxs-lookup"><span data-stu-id="62aab-300">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="62aab-301">Метод `convertToUtcClientTime` преобразует словарь, содержащий локальную дату и время, в объект Date с правильными значениями локальной даты и времени.</span><span class="sxs-lookup"><span data-stu-id="62aab-301">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="62aab-302">Параметры:</span><span class="sxs-lookup"><span data-stu-id="62aab-302">Parameters:</span></span>

|<span data-ttu-id="62aab-303">Имя</span><span class="sxs-lookup"><span data-stu-id="62aab-303">Name</span></span>| <span data-ttu-id="62aab-304">Тип</span><span class="sxs-lookup"><span data-stu-id="62aab-304">Type</span></span>| <span data-ttu-id="62aab-305">Описание</span><span class="sxs-lookup"><span data-stu-id="62aab-305">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="62aab-306">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="62aab-306">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime)|<span data-ttu-id="62aab-307">Значение локального времени для преобразования.</span><span class="sxs-lookup"><span data-stu-id="62aab-307">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="62aab-308">Требования</span><span class="sxs-lookup"><span data-stu-id="62aab-308">Requirements</span></span>

|<span data-ttu-id="62aab-309">Требование</span><span class="sxs-lookup"><span data-stu-id="62aab-309">Requirement</span></span>| <span data-ttu-id="62aab-310">Значение</span><span class="sxs-lookup"><span data-stu-id="62aab-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="62aab-311">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="62aab-311">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="62aab-312">1.0</span><span class="sxs-lookup"><span data-stu-id="62aab-312">1.0</span></span>|
|[<span data-ttu-id="62aab-313">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="62aab-313">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="62aab-314">ReadItem</span><span class="sxs-lookup"><span data-stu-id="62aab-314">ReadItem</span></span>|
|[<span data-ttu-id="62aab-315">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="62aab-315">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="62aab-316">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="62aab-316">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="62aab-317">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="62aab-317">Returns:</span></span>

<span data-ttu-id="62aab-318">Объект Date со временем в формате UTC.</span><span class="sxs-lookup"><span data-stu-id="62aab-318">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="62aab-319">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="62aab-319">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="62aab-320">Date</span><span class="sxs-lookup"><span data-stu-id="62aab-320">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="62aab-321">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="62aab-321">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="62aab-322">Отображает имеющуюся встречу из календаря.</span><span class="sxs-lookup"><span data-stu-id="62aab-322">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="62aab-323">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="62aab-323">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="62aab-324">Метод `displayAppointmentForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее сведения календаря о существующей встрече.</span><span class="sxs-lookup"><span data-stu-id="62aab-324">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="62aab-p110">В Outlook для Mac с помощью этого метода можно отобразить одну встречу, которая не является частью повторяющегося ряда, или основную встречу такого ряда, но не экземпляр из него, так как в Outlook для Mac невозможно получить доступ к свойствам экземпляра повторяющегося ряда (в том числе к идентификатору элемента).</span><span class="sxs-lookup"><span data-stu-id="62aab-p110">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="62aab-327">В Outlook Web App этот метод открывает указанную форму, только если текст формы содержит символы размером не более 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="62aab-327">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="62aab-328">Если указанный идентификатор элемента не определяет существующую встречу, на клиентском компьютере или устройстве открывается пустая страница, и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="62aab-328">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="62aab-329">Параметры:</span><span class="sxs-lookup"><span data-stu-id="62aab-329">Parameters:</span></span>

|<span data-ttu-id="62aab-330">Имя</span><span class="sxs-lookup"><span data-stu-id="62aab-330">Name</span></span>| <span data-ttu-id="62aab-331">Тип</span><span class="sxs-lookup"><span data-stu-id="62aab-331">Type</span></span>| <span data-ttu-id="62aab-332">Описание</span><span class="sxs-lookup"><span data-stu-id="62aab-332">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="62aab-333">String</span><span class="sxs-lookup"><span data-stu-id="62aab-333">String</span></span>|<span data-ttu-id="62aab-334">Идентификатор веб-служб Exchange для существующей встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="62aab-334">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="62aab-335">Требования</span><span class="sxs-lookup"><span data-stu-id="62aab-335">Requirements</span></span>

|<span data-ttu-id="62aab-336">Требование</span><span class="sxs-lookup"><span data-stu-id="62aab-336">Requirement</span></span>| <span data-ttu-id="62aab-337">Значение</span><span class="sxs-lookup"><span data-stu-id="62aab-337">Value</span></span>|
|---|---|
|[<span data-ttu-id="62aab-338">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="62aab-338">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="62aab-339">1.0</span><span class="sxs-lookup"><span data-stu-id="62aab-339">1.0</span></span>|
|[<span data-ttu-id="62aab-340">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="62aab-340">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="62aab-341">ReadItem</span><span class="sxs-lookup"><span data-stu-id="62aab-341">ReadItem</span></span>|
|[<span data-ttu-id="62aab-342">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="62aab-342">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="62aab-343">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="62aab-343">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="62aab-344">Пример</span><span class="sxs-lookup"><span data-stu-id="62aab-344">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="62aab-345">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="62aab-345">displayMessageForm(itemId)</span></span>

<span data-ttu-id="62aab-346">Отображает имеющееся сообщение.</span><span class="sxs-lookup"><span data-stu-id="62aab-346">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="62aab-347">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="62aab-347">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="62aab-348">Метод `displayMessageForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее существующее сообщение.</span><span class="sxs-lookup"><span data-stu-id="62aab-348">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="62aab-349">В Outlook Web App этот метод открывает указанную форму, только если текст формы содержит символы размером не более 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="62aab-349">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="62aab-350">Если указанный идентификатор элемента не определяет существующее сообщение, окно на клиентском компьютере не открывается и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="62aab-350">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="62aab-p111">Не используйте `displayMessageForm` с параметром `itemId`, который представляет собой встречу. Используйте метод `displayAppointmentForm`, чтобы отобразить сведения о существующей встрече, а метод `displayNewAppointmentForm` — для отображения формы создания встречи.</span><span class="sxs-lookup"><span data-stu-id="62aab-p111">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="62aab-353">Параметры</span><span class="sxs-lookup"><span data-stu-id="62aab-353">Parameters:</span></span>

|<span data-ttu-id="62aab-354">Имя</span><span class="sxs-lookup"><span data-stu-id="62aab-354">Name</span></span>| <span data-ttu-id="62aab-355">Тип</span><span class="sxs-lookup"><span data-stu-id="62aab-355">Type</span></span>| <span data-ttu-id="62aab-356">Описание</span><span class="sxs-lookup"><span data-stu-id="62aab-356">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="62aab-357">Строка</span><span class="sxs-lookup"><span data-stu-id="62aab-357">String</span></span>|<span data-ttu-id="62aab-358">Идентификатор веб-служб Exchange для существующего сообщения.</span><span class="sxs-lookup"><span data-stu-id="62aab-358">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="62aab-359">Требования</span><span class="sxs-lookup"><span data-stu-id="62aab-359">Requirements</span></span>

|<span data-ttu-id="62aab-360">Требование</span><span class="sxs-lookup"><span data-stu-id="62aab-360">Requirement</span></span>| <span data-ttu-id="62aab-361">Значение</span><span class="sxs-lookup"><span data-stu-id="62aab-361">Value</span></span>|
|---|---|
|[<span data-ttu-id="62aab-362">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="62aab-362">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="62aab-363">1.0</span><span class="sxs-lookup"><span data-stu-id="62aab-363">1.0</span></span>|
|[<span data-ttu-id="62aab-364">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="62aab-364">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="62aab-365">ReadItem</span><span class="sxs-lookup"><span data-stu-id="62aab-365">ReadItem</span></span>|
|[<span data-ttu-id="62aab-366">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="62aab-366">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="62aab-367">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="62aab-367">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="62aab-368">Пример</span><span class="sxs-lookup"><span data-stu-id="62aab-368">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="62aab-369">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="62aab-369">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="62aab-370">Отображает форму для создания новой встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="62aab-370">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="62aab-371">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="62aab-371">This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="62aab-p112">Метод `displayNewAppointmentForm` открывает форму, в которой пользователь может создать встречу или собрание. Если параметры заданы, поля формы встречи автоматически заполняются их содержимым.</span><span class="sxs-lookup"><span data-stu-id="62aab-p112">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="62aab-p113">В Outlook Web App и Outlook Web App для устройств этот метод всегда отображает форму с полем участников. Если вы не укажете участников в качестве входных аргументов, метод отображает форму с кнопкой **Сохранить**. Если вы укажете участников, форма будет включать участников и кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="62aab-p113">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="62aab-p114">Если вы укажете участников или ресурсы с помощью параметра `requiredAttendees`, `optionalAttendees` или `resources` в клиенте Outlook с расширенными возможностями и Outlook RT, этот метод отобразит форму собрания с кнопкой **Отправить**. Если не указать получателей, этот метод отобразит форму встречи с кнопкой **Сохранить и закрыть**.</span><span class="sxs-lookup"><span data-stu-id="62aab-p114">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="62aab-379">Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.</span><span class="sxs-lookup"><span data-stu-id="62aab-379">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="62aab-380">Параметры</span><span class="sxs-lookup"><span data-stu-id="62aab-380">Parameters:</span></span>

> [!NOTE]
> <span data-ttu-id="62aab-381">Все параметры являются необязательными.</span><span class="sxs-lookup"><span data-stu-id="62aab-381">All parameters are optional.</span></span>

|<span data-ttu-id="62aab-382">Имя</span><span class="sxs-lookup"><span data-stu-id="62aab-382">Name</span></span>| <span data-ttu-id="62aab-383">Тип</span><span class="sxs-lookup"><span data-stu-id="62aab-383">Type</span></span>| <span data-ttu-id="62aab-384">Описание</span><span class="sxs-lookup"><span data-stu-id="62aab-384">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="62aab-385">Object</span><span class="sxs-lookup"><span data-stu-id="62aab-385">Object</span></span> | <span data-ttu-id="62aab-386">Словарь параметров, описывающий новую встречу.</span><span class="sxs-lookup"><span data-stu-id="62aab-386">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="62aab-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="62aab-387">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="62aab-p115">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из обязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="62aab-p115">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="62aab-390">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="62aab-390">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="62aab-p116">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из необязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="62aab-p116">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="62aab-393">Date</span><span class="sxs-lookup"><span data-stu-id="62aab-393">Date</span></span> | <span data-ttu-id="62aab-394">Объект `Date`, указывающий дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="62aab-394">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="62aab-395">Date</span><span class="sxs-lookup"><span data-stu-id="62aab-395">Date</span></span> | <span data-ttu-id="62aab-396">Объект `Date`, указывающий дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="62aab-396">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="62aab-397">String</span><span class="sxs-lookup"><span data-stu-id="62aab-397">String</span></span> | <span data-ttu-id="62aab-p117">Строка со сведениями о месте встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="62aab-p117">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="62aab-400">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="62aab-400">Array.&lt;String&gt;</span></span> | <span data-ttu-id="62aab-p118">Массив строк, содержащий необходимые для встречи ресурсы. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="62aab-p118">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="62aab-403">Строка</span><span class="sxs-lookup"><span data-stu-id="62aab-403">String</span></span> | <span data-ttu-id="62aab-p119">Строка с темой встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="62aab-p119">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="62aab-406">Строка</span><span class="sxs-lookup"><span data-stu-id="62aab-406">String</span></span> | <span data-ttu-id="62aab-p120">Текст сообщения о встрече. Максимальный размер содержимого сообщения — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="62aab-p120">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="62aab-409">Требования</span><span class="sxs-lookup"><span data-stu-id="62aab-409">Requirements</span></span>

|<span data-ttu-id="62aab-410">Требование</span><span class="sxs-lookup"><span data-stu-id="62aab-410">Requirement</span></span>| <span data-ttu-id="62aab-411">Значение</span><span class="sxs-lookup"><span data-stu-id="62aab-411">Value</span></span>|
|---|---|
|[<span data-ttu-id="62aab-412">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="62aab-412">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="62aab-413">1.0</span><span class="sxs-lookup"><span data-stu-id="62aab-413">1.0</span></span>|
|[<span data-ttu-id="62aab-414">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="62aab-414">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="62aab-415">ReadItem</span><span class="sxs-lookup"><span data-stu-id="62aab-415">ReadItem</span></span>|
|[<span data-ttu-id="62aab-416">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="62aab-416">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="62aab-417">Чтение</span><span class="sxs-lookup"><span data-stu-id="62aab-417">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="62aab-418">Пример</span><span class="sxs-lookup"><span data-stu-id="62aab-418">Example</span></span>

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

#### <a name="displaynewmessageformparameters"></a><span data-ttu-id="62aab-419">displayNewMessageForm(параметры)</span><span class="sxs-lookup"><span data-stu-id="62aab-419">displayNewMessageForm(parameters)</span></span>

<span data-ttu-id="62aab-420">Отображает форму для создания сообщения.</span><span class="sxs-lookup"><span data-stu-id="62aab-420">Displays a form for creating a new message.</span></span>

<span data-ttu-id="62aab-421">Метод `displayNewMessageForm` открывает форму, в которой пользователь может создать сообщение.</span><span class="sxs-lookup"><span data-stu-id="62aab-421">The `displayNewMessageForm` method opens a form that enables the user to create a new message.</span></span> <span data-ttu-id="62aab-422">Если параметры заданы, поля формы сообщения автоматически заполняются их содержимым.</span><span class="sxs-lookup"><span data-stu-id="62aab-422">If parameters are specified, the message form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="62aab-423">Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.</span><span class="sxs-lookup"><span data-stu-id="62aab-423">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="62aab-424">Параметры</span><span class="sxs-lookup"><span data-stu-id="62aab-424">Parameters:</span></span>

> [!NOTE]
> <span data-ttu-id="62aab-425">Все параметры являются необязательными.</span><span class="sxs-lookup"><span data-stu-id="62aab-425">All parameters are optional.</span></span>

|<span data-ttu-id="62aab-426">Имя</span><span class="sxs-lookup"><span data-stu-id="62aab-426">Name</span></span>| <span data-ttu-id="62aab-427">Тип</span><span class="sxs-lookup"><span data-stu-id="62aab-427">Type</span></span>| <span data-ttu-id="62aab-428">Описание</span><span class="sxs-lookup"><span data-stu-id="62aab-428">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="62aab-429">Object</span><span class="sxs-lookup"><span data-stu-id="62aab-429">Object</span></span> | <span data-ttu-id="62aab-430">Словарь параметров, описывающий новое сообщение.</span><span class="sxs-lookup"><span data-stu-id="62aab-430">A dictionary of parameters describing the new message.</span></span> |
| `parameters.toRecipients` | <span data-ttu-id="62aab-431">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="62aab-431">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="62aab-432">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из получателей, указанных в строке "Кому".</span><span class="sxs-lookup"><span data-stu-id="62aab-432">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the To line.</span></span> <span data-ttu-id="62aab-433">Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="62aab-433">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.ccRecipients` | <span data-ttu-id="62aab-434">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="62aab-434">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="62aab-435">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из получателей, указанных в строке "Копия".</span><span class="sxs-lookup"><span data-stu-id="62aab-435">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Cc line.</span></span> <span data-ttu-id="62aab-436">Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="62aab-436">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.bccRecipients` | <span data-ttu-id="62aab-437">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="62aab-437">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="62aab-438">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из получателей, указанных в строке "Скрытая копия".</span><span class="sxs-lookup"><span data-stu-id="62aab-438">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the recipients on the Bcc line.</span></span> <span data-ttu-id="62aab-439">Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="62aab-439">The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="62aab-440">Строка</span><span class="sxs-lookup"><span data-stu-id="62aab-440">String</span></span> | <span data-ttu-id="62aab-441">Строка с темой сообщения.</span><span class="sxs-lookup"><span data-stu-id="62aab-441">A string containing the subject of the message.</span></span> <span data-ttu-id="62aab-442">Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="62aab-442">The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.htmlBody` | <span data-ttu-id="62aab-443">String</span><span class="sxs-lookup"><span data-stu-id="62aab-443">String</span></span> | <span data-ttu-id="62aab-444">Текст сообщения в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="62aab-444">The HTML body of the message.</span></span> <span data-ttu-id="62aab-445">Максимальный размер содержимого сообщения — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="62aab-445">The body content is limited to a maximum size of 32 KB.</span></span> |
| `parameters.attachments` | <span data-ttu-id="62aab-446">Array.&lt;Object&gt;</span><span class="sxs-lookup"><span data-stu-id="62aab-446">Array.&lt;Object&gt;</span></span> | <span data-ttu-id="62aab-447">Массив объектов JSON, представляющих собой вложенные файлы или элементы.</span><span class="sxs-lookup"><span data-stu-id="62aab-447">An array of JSON objects that are either file or item attachments.</span></span> |
| `parameters.attachments.type` | <span data-ttu-id="62aab-448">Строка</span><span class="sxs-lookup"><span data-stu-id="62aab-448">String</span></span> | <span data-ttu-id="62aab-p127">Указывает тип вложения. Допустимые значения: `file` для вложенного файла и `item` для вложенного элемента.</span><span class="sxs-lookup"><span data-stu-id="62aab-p127">Indicates the type of attachment. Must be `file` for a file attachment or `item` for an item attachment.</span></span> |
| `parameters.attachments.name` | <span data-ttu-id="62aab-451">Строка</span><span class="sxs-lookup"><span data-stu-id="62aab-451">String</span></span> | <span data-ttu-id="62aab-452">Строка, содержащая имя вложения, длиной до 255 символов.</span><span class="sxs-lookup"><span data-stu-id="62aab-452">A string that contains the name of the attachment, up to 255 characters in length.</span></span>|
| `parameters.attachments.url` | <span data-ttu-id="62aab-453">Строка</span><span class="sxs-lookup"><span data-stu-id="62aab-453">String</span></span> | <span data-ttu-id="62aab-p128">Используется, только если свойству `type` задано значение `file`. URI расположения файла.</span><span class="sxs-lookup"><span data-stu-id="62aab-p128">Only used if `type` is set to `file`. The URI of the location for the file.</span></span> |
| `parameters.attachments.isInline` | <span data-ttu-id="62aab-456">Boolean</span><span class="sxs-lookup"><span data-stu-id="62aab-456">Boolean</span></span> | <span data-ttu-id="62aab-p129">Используется, только если свойству `type` задано значение `file`. Значение `true` указывает на то, что вложение будет встроено в текст сообщения и не должно отображаться в списке вложений.</span><span class="sxs-lookup"><span data-stu-id="62aab-p129">Only used if `type` is set to `file`. If `true`, indicates that the attachment will be shown inline in the message body, and should not be displayed in the attachment list.</span></span> |
| `parameters.attachments.itemId` | <span data-ttu-id="62aab-459">String</span><span class="sxs-lookup"><span data-stu-id="62aab-459">String</span></span> | <span data-ttu-id="62aab-460">Используется, только если для `type` задано значение `item`.</span><span class="sxs-lookup"><span data-stu-id="62aab-460">Only used if `type` is set to `item`.</span></span> <span data-ttu-id="62aab-461">Идентификатор элемента EWS существующего письма, которое нужно вложить в новое сообщение.</span><span class="sxs-lookup"><span data-stu-id="62aab-461">The EWS item id of the existing e-mail you want to attach to the new message.</span></span> <span data-ttu-id="62aab-462">Это строка длиной до 100 символов.</span><span class="sxs-lookup"><span data-stu-id="62aab-462">This is a string up to 100 characters.</span></span> |


##### <a name="requirements"></a><span data-ttu-id="62aab-463">Требования</span><span class="sxs-lookup"><span data-stu-id="62aab-463">Requirements</span></span>

|<span data-ttu-id="62aab-464">Требование</span><span class="sxs-lookup"><span data-stu-id="62aab-464">Requirement</span></span>| <span data-ttu-id="62aab-465">Значение</span><span class="sxs-lookup"><span data-stu-id="62aab-465">Value</span></span>|
|---|---|
|[<span data-ttu-id="62aab-466">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="62aab-466">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="62aab-467">1.6</span><span class="sxs-lookup"><span data-stu-id="62aab-467">1.6</span></span> |
|[<span data-ttu-id="62aab-468">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="62aab-468">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="62aab-469">ReadItem</span><span class="sxs-lookup"><span data-stu-id="62aab-469">ReadItem</span></span>|
|[<span data-ttu-id="62aab-470">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="62aab-470">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="62aab-471">Чтение</span><span class="sxs-lookup"><span data-stu-id="62aab-471">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="62aab-472">Пример</span><span class="sxs-lookup"><span data-stu-id="62aab-472">Example</span></span>

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

#### <a name="getcallbacktokenasyncoptions-callback"></a><span data-ttu-id="62aab-473">getCallbackTokenAsync([options], callback)</span><span class="sxs-lookup"><span data-stu-id="62aab-473">getCallbackTokenAsync([options], callback)</span></span>

<span data-ttu-id="62aab-474">Возвращает строку, содержащую маркер, который используется для вызова интерфейсов REST API или веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="62aab-474">Gets a string that contains a token used to call REST APIs or Exchange Web Services.</span></span>

<span data-ttu-id="62aab-p131">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="62aab-p131">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

> [!NOTE]
> <span data-ttu-id="62aab-477">Рекомендуем сделать так, чтобы по мере возможности надстройки использовали интерфейсы REST API, а не веб-службы Exchange.</span><span class="sxs-lookup"><span data-stu-id="62aab-477">It is recommended that add-ins use the REST APIs instead of Exchange Web Services whenever possible.</span></span> 

<span data-ttu-id="62aab-478">**Маркеры REST**</span><span class="sxs-lookup"><span data-stu-id="62aab-478">**REST Tokens**</span></span>

<span data-ttu-id="62aab-p132">Если запрашивается маркер REST (`options.isRest = true`), полученный маркер не подойдет для проверки подлинности при вызовах веб-служб Exchange. Область действия маркера будет ограничена доступом только для чтения к текущему элементу и его вложениям, если в манифесте надстройки не указано разрешение [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission). Если указано разрешение `ReadWriteMailbox`, полученный маркер предоставит доступ на чтение и запись к почте, календарю и контактам, включая возможность отправки почты.</span><span class="sxs-lookup"><span data-stu-id="62aab-p132">When a REST token is requested (`options.isRest = true`), the resulting token will not work to authenticate Exchange Web Services calls. The token will be limited in scope to read-only access to the current item and its attachments, unless the add-in has specified the [`ReadWriteMailbox`](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions#readwritemailbox-permission) permission in its manifest. If the `ReadWriteMailbox` permission is specified, the resulting token will grant read/write access to mail, calendar, and contacts, including the ability to send mail.</span></span>

<span data-ttu-id="62aab-482">С помощью свойства `restUrl` надстройка должна определить правильный URL-адрес для вызовов REST API.</span><span class="sxs-lookup"><span data-stu-id="62aab-482">The add-in should use the `restUrl` property to determine the correct URL to use when making REST API calls.</span></span>

<span data-ttu-id="62aab-483">**Маркеры EWS**</span><span class="sxs-lookup"><span data-stu-id="62aab-483">**EWS Tokens**</span></span>

<span data-ttu-id="62aab-p133">Если запрашивается маркер EWS (`options.isRest = false`), полученный маркер не подойдет для проверки подлинности при вызовах REST API. Область действия маркера будет ограничена доступом к текущему элементу.</span><span class="sxs-lookup"><span data-stu-id="62aab-p133">When an EWS token is requested (`options.isRest = false`), the resulting token will not work to authenticate REST API calls. The token will be limited in scope to accessing the current item.</span></span>

<span data-ttu-id="62aab-486">С помощью свойства `ewsUrl` надстройка должна определить правильный URL-адрес для вызовов EWS.</span><span class="sxs-lookup"><span data-stu-id="62aab-486">The add-in should use the `ewsUrl` property to determine the correct URL to use when making EWS calls.</span></span>

##### <a name="parameters"></a><span data-ttu-id="62aab-487">Параметры:</span><span class="sxs-lookup"><span data-stu-id="62aab-487">Parameters:</span></span>

|<span data-ttu-id="62aab-488">Имя</span><span class="sxs-lookup"><span data-stu-id="62aab-488">Name</span></span>| <span data-ttu-id="62aab-489">Тип</span><span class="sxs-lookup"><span data-stu-id="62aab-489">Type</span></span>| <span data-ttu-id="62aab-490">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="62aab-490">Attributes</span></span>| <span data-ttu-id="62aab-491">Описание</span><span class="sxs-lookup"><span data-stu-id="62aab-491">Description</span></span>|
|---|---|---|---|
| `options` | <span data-ttu-id="62aab-492">Object</span><span class="sxs-lookup"><span data-stu-id="62aab-492">Object</span></span> | <span data-ttu-id="62aab-493">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="62aab-493">&lt;optional&gt;</span></span> | <span data-ttu-id="62aab-494">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="62aab-494">An object literal that contains one or more of the following properties.</span></span> |
| `options.isRest` | <span data-ttu-id="62aab-495">Boolean</span><span class="sxs-lookup"><span data-stu-id="62aab-495">Boolean</span></span> |  <span data-ttu-id="62aab-496">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="62aab-496">&lt;optional&gt;</span></span> | <span data-ttu-id="62aab-p134">Определяет, будет ли предоставленный маркер использоваться для интерфейсов REST API Outlook или веб-служб Exchange. Значение по умолчанию: `false`.</span><span class="sxs-lookup"><span data-stu-id="62aab-p134">Determines if the token provided will be used for the Outlook REST APIs or Exchange Web Services. Default value is `false`.</span></span> |
| `options.asyncContext` | <span data-ttu-id="62aab-499">Object</span><span class="sxs-lookup"><span data-stu-id="62aab-499">Object</span></span> |  <span data-ttu-id="62aab-500">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="62aab-500">&lt;optional&gt;</span></span> | <span data-ttu-id="62aab-501">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="62aab-501">Any state data that is passed to the asynchronous method.</span></span> |
|`callback`| <span data-ttu-id="62aab-502">function</span><span class="sxs-lookup"><span data-stu-id="62aab-502">function</span></span>||<span data-ttu-id="62aab-p135">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult). Токен указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="62aab-p135">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="62aab-505">Требования</span><span class="sxs-lookup"><span data-stu-id="62aab-505">Requirements</span></span>

|<span data-ttu-id="62aab-506">Требование</span><span class="sxs-lookup"><span data-stu-id="62aab-506">Requirement</span></span>| <span data-ttu-id="62aab-507">Значение</span><span class="sxs-lookup"><span data-stu-id="62aab-507">Value</span></span>|
|---|---|
|[<span data-ttu-id="62aab-508">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="62aab-508">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="62aab-509">1.5</span><span class="sxs-lookup"><span data-stu-id="62aab-509">1.5</span></span> |
|[<span data-ttu-id="62aab-510">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="62aab-510">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="62aab-511">ReadItem</span><span class="sxs-lookup"><span data-stu-id="62aab-511">ReadItem</span></span>|
|[<span data-ttu-id="62aab-512">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="62aab-512">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="62aab-513">Создание и чтение</span><span class="sxs-lookup"><span data-stu-id="62aab-513">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="62aab-514">Пример</span><span class="sxs-lookup"><span data-stu-id="62aab-514">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="62aab-515">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="62aab-515">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="62aab-516">Получает строку, содержащую маркер, используемый для получения вложения или элемента с Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="62aab-516">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="62aab-p136">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный токен с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="62aab-p136">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="62aab-p137">Вы можете передать сторонней системе токен и идентификатор вложения или элемента. Сторонняя система использует этот токен как токен авторизации, чтобы вызвать операцию [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) или [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) веб-служб Exchange для возврата вложения или элемента. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="62aab-p137">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="62aab-522">Для вызова метода `getCallbackTokenAsync` в режиме чтения манифесте приложения должно быть указано разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="62aab-522">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="62aab-p138">Чтобы получить идентификатор элемента для передачи в метод `getCallbackTokenAsync`, в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="62aab-p138">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="62aab-525">Параметры:</span><span class="sxs-lookup"><span data-stu-id="62aab-525">Parameters:</span></span>

|<span data-ttu-id="62aab-526">Имя</span><span class="sxs-lookup"><span data-stu-id="62aab-526">Name</span></span>| <span data-ttu-id="62aab-527">Тип</span><span class="sxs-lookup"><span data-stu-id="62aab-527">Type</span></span>| <span data-ttu-id="62aab-528">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="62aab-528">Attributes</span></span>| <span data-ttu-id="62aab-529">Описание</span><span class="sxs-lookup"><span data-stu-id="62aab-529">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="62aab-530">функция</span><span class="sxs-lookup"><span data-stu-id="62aab-530">function</span></span>||<span data-ttu-id="62aab-p139">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult). Токен указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="62aab-p139">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object. The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="62aab-533">Object</span><span class="sxs-lookup"><span data-stu-id="62aab-533">Object</span></span>| <span data-ttu-id="62aab-534">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="62aab-534">&lt;optional&gt;</span></span>|<span data-ttu-id="62aab-535">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="62aab-535">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="62aab-536">Требования</span><span class="sxs-lookup"><span data-stu-id="62aab-536">Requirements</span></span>

|<span data-ttu-id="62aab-537">Требование</span><span class="sxs-lookup"><span data-stu-id="62aab-537">Requirement</span></span>| <span data-ttu-id="62aab-538">Значение</span><span class="sxs-lookup"><span data-stu-id="62aab-538">Value</span></span>|
|---|---|
|[<span data-ttu-id="62aab-539">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="62aab-539">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="62aab-540">1.3</span><span class="sxs-lookup"><span data-stu-id="62aab-540">1.3</span></span>|
|[<span data-ttu-id="62aab-541">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="62aab-541">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="62aab-542">ReadItem</span><span class="sxs-lookup"><span data-stu-id="62aab-542">ReadItem</span></span>|
|[<span data-ttu-id="62aab-543">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="62aab-543">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="62aab-544">Создание и чтение</span><span class="sxs-lookup"><span data-stu-id="62aab-544">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="62aab-545">Пример</span><span class="sxs-lookup"><span data-stu-id="62aab-545">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="62aab-546">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="62aab-546">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="62aab-547">Получает маркер, идентифицирующий пользователя и надстройку Office.</span><span class="sxs-lookup"><span data-stu-id="62aab-547">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="62aab-548">Метод `getUserIdentityTokenAsync` возвращает токен, который можно использовать для идентификации, а также [проверки подлинности надстройки и пользователя в сторонней системе](https://docs.microsoft.com/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="62aab-548">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="62aab-549">Параметры</span><span class="sxs-lookup"><span data-stu-id="62aab-549">Parameters:</span></span>

|<span data-ttu-id="62aab-550">Имя</span><span class="sxs-lookup"><span data-stu-id="62aab-550">Name</span></span>| <span data-ttu-id="62aab-551">Тип</span><span class="sxs-lookup"><span data-stu-id="62aab-551">Type</span></span>| <span data-ttu-id="62aab-552">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="62aab-552">Attributes</span></span>| <span data-ttu-id="62aab-553">Описание</span><span class="sxs-lookup"><span data-stu-id="62aab-553">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="62aab-554">функция</span><span class="sxs-lookup"><span data-stu-id="62aab-554">function</span></span>||<span data-ttu-id="62aab-555">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="62aab-555">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="62aab-556">Токен указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="62aab-556">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="62aab-557">Object</span><span class="sxs-lookup"><span data-stu-id="62aab-557">Object</span></span>| <span data-ttu-id="62aab-558">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="62aab-558">&lt;optional&gt;</span></span>|<span data-ttu-id="62aab-559">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="62aab-559">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="62aab-560">Требования</span><span class="sxs-lookup"><span data-stu-id="62aab-560">Requirements</span></span>

|<span data-ttu-id="62aab-561">Требование</span><span class="sxs-lookup"><span data-stu-id="62aab-561">Requirement</span></span>| <span data-ttu-id="62aab-562">Значение</span><span class="sxs-lookup"><span data-stu-id="62aab-562">Value</span></span>|
|---|---|
|[<span data-ttu-id="62aab-563">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="62aab-563">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="62aab-564">1.0</span><span class="sxs-lookup"><span data-stu-id="62aab-564">1.0</span></span>|
|[<span data-ttu-id="62aab-565">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="62aab-565">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="62aab-566">ReadItem</span><span class="sxs-lookup"><span data-stu-id="62aab-566">ReadItem</span></span>|
|[<span data-ttu-id="62aab-567">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="62aab-567">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="62aab-568">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="62aab-568">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="62aab-569">Пример</span><span class="sxs-lookup"><span data-stu-id="62aab-569">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="62aab-570">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="62aab-570">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="62aab-571">Выполняет асинхронный запрос для веб-служб Exchange (EWS) на сервере Exchange Server, на котором размещен почтовый ящик пользователя.</span><span class="sxs-lookup"><span data-stu-id="62aab-571">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="62aab-572">Этот метод не поддерживается в следующих сценариях:</span><span class="sxs-lookup"><span data-stu-id="62aab-572">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="62aab-573">В Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="62aab-573">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="62aab-574">Если надстройка загружается в почтовый ящик Gmail.</span><span class="sxs-lookup"><span data-stu-id="62aab-574">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="62aab-575">В таких случаях надстройка должна [использовать REST API](https://docs.microsoft.com/outlook/add-ins/use-rest-api) для доступа к почтовому ящику пользователя.</span><span class="sxs-lookup"><span data-stu-id="62aab-575">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="62aab-576">Метод `makeEwsRequestAsync` отправляет запрос EWS от имени надстройки в Exchange.</span><span class="sxs-lookup"><span data-stu-id="62aab-576">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="62aab-577">Список поддерживаемых операций EWS см. в статье [Вызов веб-служб из надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support).</span><span class="sxs-lookup"><span data-stu-id="62aab-577">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="62aab-578">С помощью метода `makeEwsRequestAsync` невозможно запрашивать элементы, связанные с папкой.</span><span class="sxs-lookup"><span data-stu-id="62aab-578">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="62aab-579">В запросе XML должна быть указана кодировка UTF-8.</span><span class="sxs-lookup"><span data-stu-id="62aab-579">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="62aab-p141">У вашей надстройки должно быть разрешение **ReadWriteMailbox** для использования метода `makeEwsRequestAsync`. Сведения об использовании разрешения **ReadWriteMailbox** и операций EWS, которые можно вызывать с помощью метода `makeEwsRequestAsync`, см. в статье [Указание разрешений для доступа почтовой надстройки к почтовому ящику пользователя](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="62aab-p141">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="62aab-582">Администратор сервера должен установить значение true для параметра `OAuthAuthentication` в каталоге сервера клиентского доступа EWS, чтобы метод `makeEwsRequestAsync` мог выполнять запросы EWS.</span><span class="sxs-lookup"><span data-stu-id="62aab-582">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="62aab-583">Различия версий</span><span class="sxs-lookup"><span data-stu-id="62aab-583">Version differences</span></span>

<span data-ttu-id="62aab-584">Если вы используете метод `makeEwsRequestAsync` в почтовых приложениях, которые выполняются в Outlook версии более ранней, чем 15.0.4535.1004, указывайте кодировку `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="62aab-584">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="62aab-p142">Значение кодировки не нужно указывать, если почтовое приложение выполняется в Outlook в Интернете. Чтобы определить, выполняется ли приложение в Outlook или Outlook в Интернете, используйте свойство mailbox.diagnostics.hostName. Используемую версию Outlook можно определить с помощью свойства mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="62aab-p142">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="62aab-588">Параметры:</span><span class="sxs-lookup"><span data-stu-id="62aab-588">Parameters:</span></span>

|<span data-ttu-id="62aab-589">Имя</span><span class="sxs-lookup"><span data-stu-id="62aab-589">Name</span></span>| <span data-ttu-id="62aab-590">Тип</span><span class="sxs-lookup"><span data-stu-id="62aab-590">Type</span></span>| <span data-ttu-id="62aab-591">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="62aab-591">Attributes</span></span>| <span data-ttu-id="62aab-592">Описание</span><span class="sxs-lookup"><span data-stu-id="62aab-592">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="62aab-593">String</span><span class="sxs-lookup"><span data-stu-id="62aab-593">String</span></span>||<span data-ttu-id="62aab-594">Запрос EWS.</span><span class="sxs-lookup"><span data-stu-id="62aab-594">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="62aab-595">функция</span><span class="sxs-lookup"><span data-stu-id="62aab-595">function</span></span>||<span data-ttu-id="62aab-596">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="62aab-596">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="62aab-597">Результат XML вызова EWS указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="62aab-597">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="62aab-598">Если размер результата превышает 1 МБ, возвращается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="62aab-598">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="62aab-599">Объект</span><span class="sxs-lookup"><span data-stu-id="62aab-599">Object</span></span>| <span data-ttu-id="62aab-600">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="62aab-600">&lt;optional&gt;</span></span>|<span data-ttu-id="62aab-601">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="62aab-601">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="62aab-602">Требования</span><span class="sxs-lookup"><span data-stu-id="62aab-602">Requirements</span></span>

|<span data-ttu-id="62aab-603">Требование</span><span class="sxs-lookup"><span data-stu-id="62aab-603">Requirement</span></span>| <span data-ttu-id="62aab-604">Значение</span><span class="sxs-lookup"><span data-stu-id="62aab-604">Value</span></span>|
|---|---|
|[<span data-ttu-id="62aab-605">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="62aab-605">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="62aab-606">1.0</span><span class="sxs-lookup"><span data-stu-id="62aab-606">1.0</span></span>|
|[<span data-ttu-id="62aab-607">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="62aab-607">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="62aab-608">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="62aab-608">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="62aab-609">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="62aab-609">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="62aab-610">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="62aab-610">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="62aab-611">Пример</span><span class="sxs-lookup"><span data-stu-id="62aab-611">Example</span></span>

<span data-ttu-id="62aab-612">В приведенном ниже примере вызывается `makeEwsRequestAsync` для получения темы элемента с помощью операции `GetItem`.</span><span class="sxs-lookup"><span data-stu-id="62aab-612">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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

####  <a name="removehandlerasynceventtype-options-callback"></a><span data-ttu-id="62aab-613">removeHandlerAsync(eventType, [options], [callback])</span><span class="sxs-lookup"><span data-stu-id="62aab-613">removeHandlerAsync(eventType, [options], [callback])</span></span>

<span data-ttu-id="62aab-614">Удаляет обработчиков для поддерживаемого типа события.</span><span class="sxs-lookup"><span data-stu-id="62aab-614">Removes the event handlers for a supported event type.</span></span>

<span data-ttu-id="62aab-615">В настоящее время поддерживаются следующие типы событий: `Office.EventType.ItemChanged` и `Office.EventType.OfficeThemeChanged`.</span><span class="sxs-lookup"><span data-stu-id="62aab-615">Currently, the supported event types are `Office.EventType.ItemChanged` and `Office.EventType.OfficeThemeChanged`.</span></span>

##### <a name="parameters"></a><span data-ttu-id="62aab-616">Параметры</span><span class="sxs-lookup"><span data-stu-id="62aab-616">Parameters:</span></span>

| <span data-ttu-id="62aab-617">Имя</span><span class="sxs-lookup"><span data-stu-id="62aab-617">Name</span></span> | <span data-ttu-id="62aab-618">Тип</span><span class="sxs-lookup"><span data-stu-id="62aab-618">Type</span></span> | <span data-ttu-id="62aab-619">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="62aab-619">Attributes</span></span> | <span data-ttu-id="62aab-620">Описание</span><span class="sxs-lookup"><span data-stu-id="62aab-620">Description</span></span> |
|---|---|---|---|
| `eventType` | [<span data-ttu-id="62aab-621">Office.EventType</span><span class="sxs-lookup"><span data-stu-id="62aab-621">Office.EventType</span></span>](office.md#eventtype-string) || <span data-ttu-id="62aab-622">Событие, которое должно отменить обработчик.</span><span class="sxs-lookup"><span data-stu-id="62aab-622">The event that should revoke the handler.</span></span> |
| `options` | <span data-ttu-id="62aab-623">Объект</span><span class="sxs-lookup"><span data-stu-id="62aab-623">Object</span></span> | <span data-ttu-id="62aab-624">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="62aab-624">&lt;optional&gt;</span></span> | <span data-ttu-id="62aab-625">Объектный литерал, содержащий одно или несколько из указанных ниже свойств.</span><span class="sxs-lookup"><span data-stu-id="62aab-625">An object literal that contains one or more of the following properties.</span></span> |
| `options.asyncContext` | <span data-ttu-id="62aab-626">Object</span><span class="sxs-lookup"><span data-stu-id="62aab-626">Object</span></span> | <span data-ttu-id="62aab-627">&lt;необязательный&gt;</span><span class="sxs-lookup"><span data-stu-id="62aab-627">&lt;optional&gt;</span></span> | <span data-ttu-id="62aab-628">Разработчики могут указать любой объект, к которому необходимо получить доступ, в методе обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="62aab-628">Developers can provide any object they wish to access in the callback method.</span></span> |
| `callback` | <span data-ttu-id="62aab-629">функция</span><span class="sxs-lookup"><span data-stu-id="62aab-629">function</span></span>| <span data-ttu-id="62aab-630">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="62aab-630">&lt;optional&gt;</span></span>|<span data-ttu-id="62aab-631">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="62aab-631">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="62aab-632">Требования</span><span class="sxs-lookup"><span data-stu-id="62aab-632">Requirements</span></span>

|<span data-ttu-id="62aab-633">Требование</span><span class="sxs-lookup"><span data-stu-id="62aab-633">Requirement</span></span>| <span data-ttu-id="62aab-634">Значение</span><span class="sxs-lookup"><span data-stu-id="62aab-634">Value</span></span>|
|---|---|
|[<span data-ttu-id="62aab-635">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="62aab-635">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="62aab-636">1.5</span><span class="sxs-lookup"><span data-stu-id="62aab-636">1.5</span></span> |
|[<span data-ttu-id="62aab-637">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="62aab-637">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="62aab-638">ReadItem</span><span class="sxs-lookup"><span data-stu-id="62aab-638">ReadItem</span></span> |
|[<span data-ttu-id="62aab-639">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="62aab-639">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="62aab-640">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="62aab-640">Compose or read</span></span>|
