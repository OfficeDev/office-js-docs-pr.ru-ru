---
title: Office. Context. Mailbox — набор обязательных элементов 1,4
description: ''
ms.date: 08/30/2019
localization_priority: Normal
ms.openlocfilehash: 66ae7cb05ac56224fd7461c5c29587e21a24020a
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696213"
---
# <a name="mailbox"></a><span data-ttu-id="eaabf-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="eaabf-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="eaabf-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="eaabf-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="eaabf-104">Предоставляет доступ к объектной модели надстройки Outlook для Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="eaabf-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="eaabf-105">Требования</span><span class="sxs-lookup"><span data-stu-id="eaabf-105">Requirements</span></span>

|<span data-ttu-id="eaabf-106">Требование</span><span class="sxs-lookup"><span data-stu-id="eaabf-106">Requirement</span></span>| <span data-ttu-id="eaabf-107">Значение</span><span class="sxs-lookup"><span data-stu-id="eaabf-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="eaabf-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="eaabf-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eaabf-109">1.0</span><span class="sxs-lookup"><span data-stu-id="eaabf-109">1.0</span></span>|
|[<span data-ttu-id="eaabf-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="eaabf-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="eaabf-111">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="eaabf-111">Restricted</span></span>|
|[<span data-ttu-id="eaabf-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="eaabf-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="eaabf-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="eaabf-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="eaabf-114">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="eaabf-114">Members and methods</span></span>

| <span data-ttu-id="eaabf-115">Элемент</span><span class="sxs-lookup"><span data-stu-id="eaabf-115">Member</span></span> | <span data-ttu-id="eaabf-116">Тип</span><span class="sxs-lookup"><span data-stu-id="eaabf-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="eaabf-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="eaabf-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="eaabf-118">Элемент</span><span class="sxs-lookup"><span data-stu-id="eaabf-118">Member</span></span> |
| [<span data-ttu-id="eaabf-119">convertToEwsId</span><span class="sxs-lookup"><span data-stu-id="eaabf-119">convertToEwsId</span></span>](#converttoewsiditemid-restversion--string) | <span data-ttu-id="eaabf-120">Метод</span><span class="sxs-lookup"><span data-stu-id="eaabf-120">Method</span></span> |
| [<span data-ttu-id="eaabf-121">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="eaabf-121">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="eaabf-122">Метод</span><span class="sxs-lookup"><span data-stu-id="eaabf-122">Method</span></span> |
| [<span data-ttu-id="eaabf-123">convertToRestId</span><span class="sxs-lookup"><span data-stu-id="eaabf-123">convertToRestId</span></span>](#converttorestiditemid-restversion--string) | <span data-ttu-id="eaabf-124">Метод</span><span class="sxs-lookup"><span data-stu-id="eaabf-124">Method</span></span> |
| [<span data-ttu-id="eaabf-125">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="eaabf-125">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="eaabf-126">Метод</span><span class="sxs-lookup"><span data-stu-id="eaabf-126">Method</span></span> |
| [<span data-ttu-id="eaabf-127">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="eaabf-127">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="eaabf-128">Метод</span><span class="sxs-lookup"><span data-stu-id="eaabf-128">Method</span></span> |
| [<span data-ttu-id="eaabf-129">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="eaabf-129">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="eaabf-130">Метод</span><span class="sxs-lookup"><span data-stu-id="eaabf-130">Method</span></span> |
| [<span data-ttu-id="eaabf-131">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="eaabf-131">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="eaabf-132">Метод</span><span class="sxs-lookup"><span data-stu-id="eaabf-132">Method</span></span> |
| [<span data-ttu-id="eaabf-133">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="eaabf-133">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="eaabf-134">Метод</span><span class="sxs-lookup"><span data-stu-id="eaabf-134">Method</span></span> |
| [<span data-ttu-id="eaabf-135">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="eaabf-135">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="eaabf-136">Метод</span><span class="sxs-lookup"><span data-stu-id="eaabf-136">Method</span></span> |
| [<span data-ttu-id="eaabf-137">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="eaabf-137">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="eaabf-138">Метод</span><span class="sxs-lookup"><span data-stu-id="eaabf-138">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="eaabf-139">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="eaabf-139">Namespaces</span></span>

<span data-ttu-id="eaabf-140">[diagnostics](Office.context.mailbox.diagnostics.md). Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="eaabf-140">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="eaabf-141">[item](Office.context.mailbox.item.md). Предоставляет методы и свойства для доступа к сообщению или встрече в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="eaabf-141">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="eaabf-142">[userProfile](Office.context.mailbox.userProfile.md). Предоставляет сведения о пользователе в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="eaabf-142">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="eaabf-143">Элементы</span><span class="sxs-lookup"><span data-stu-id="eaabf-143">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="eaabf-144">ewsUrl: строка</span><span class="sxs-lookup"><span data-stu-id="eaabf-144">ewsUrl: String</span></span>

<span data-ttu-id="eaabf-145">Получает URL-адрес конечной точки веб-служб Exchange (EWS) для этой учетной записи электронной почты.</span><span class="sxs-lookup"><span data-stu-id="eaabf-145">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account.</span></span> <span data-ttu-id="eaabf-146">Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="eaabf-146">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="eaabf-147">Этот элемент не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="eaabf-147">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="eaabf-p102">Удаленная служба может использовать значение `ewsUrl`, чтобы выполнять вызовы EWS для почтового ящика пользователя. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="eaabf-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="eaabf-150">Чтобы вызвать элемент `ewsUrl` в режиме чтения, в манифесте приложения должно быть указано разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="eaabf-150">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="eaabf-p103">Перед использованием элемента `ewsUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="eaabf-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="eaabf-153">Тип</span><span class="sxs-lookup"><span data-stu-id="eaabf-153">Type</span></span>

*   <span data-ttu-id="eaabf-154">String</span><span class="sxs-lookup"><span data-stu-id="eaabf-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="eaabf-155">Требования</span><span class="sxs-lookup"><span data-stu-id="eaabf-155">Requirements</span></span>

|<span data-ttu-id="eaabf-156">Требование</span><span class="sxs-lookup"><span data-stu-id="eaabf-156">Requirement</span></span>| <span data-ttu-id="eaabf-157">Значение</span><span class="sxs-lookup"><span data-stu-id="eaabf-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="eaabf-158">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="eaabf-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eaabf-159">1.0</span><span class="sxs-lookup"><span data-stu-id="eaabf-159">1.0</span></span>|
|[<span data-ttu-id="eaabf-160">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="eaabf-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="eaabf-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="eaabf-161">ReadItem</span></span>|
|[<span data-ttu-id="eaabf-162">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="eaabf-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="eaabf-163">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="eaabf-163">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="eaabf-164">Методы</span><span class="sxs-lookup"><span data-stu-id="eaabf-164">Methods</span></span>

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="eaabf-165">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="eaabf-165">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="eaabf-166">Преобразовывает идентификатор элемента из формата REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="eaabf-166">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="eaabf-167">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="eaabf-167">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="eaabf-p104">Формат идентификаторов, извлекаемых через API REST (например, [API Почты Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)), отличается от формата веб-служб Exchange (EWS). Метод `convertToEwsId` преобразовывает идентификатор в формате REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="eaabf-p104">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="eaabf-170">Параметры</span><span class="sxs-lookup"><span data-stu-id="eaabf-170">Parameters</span></span>

|<span data-ttu-id="eaabf-171">Имя</span><span class="sxs-lookup"><span data-stu-id="eaabf-171">Name</span></span>| <span data-ttu-id="eaabf-172">Тип</span><span class="sxs-lookup"><span data-stu-id="eaabf-172">Type</span></span>| <span data-ttu-id="eaabf-173">Описание</span><span class="sxs-lookup"><span data-stu-id="eaabf-173">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="eaabf-174">String</span><span class="sxs-lookup"><span data-stu-id="eaabf-174">String</span></span>|<span data-ttu-id="eaabf-175">Идентификатор элемента в формате REST API для Outlook</span><span class="sxs-lookup"><span data-stu-id="eaabf-175">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="eaabf-176">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="eaabf-176">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.4)|<span data-ttu-id="eaabf-177">Значение, определяющее версию REST API для Outlook, которая используется для извлечения идентификатора элемента.</span><span class="sxs-lookup"><span data-stu-id="eaabf-177">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="eaabf-178">Требования</span><span class="sxs-lookup"><span data-stu-id="eaabf-178">Requirements</span></span>

|<span data-ttu-id="eaabf-179">Требование</span><span class="sxs-lookup"><span data-stu-id="eaabf-179">Requirement</span></span>| <span data-ttu-id="eaabf-180">Значение</span><span class="sxs-lookup"><span data-stu-id="eaabf-180">Value</span></span>|
|---|---|
|[<span data-ttu-id="eaabf-181">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="eaabf-181">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eaabf-182">1.3</span><span class="sxs-lookup"><span data-stu-id="eaabf-182">1.3</span></span>|
|[<span data-ttu-id="eaabf-183">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="eaabf-183">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="eaabf-184">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="eaabf-184">Restricted</span></span>|
|[<span data-ttu-id="eaabf-185">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="eaabf-185">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="eaabf-186">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="eaabf-186">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="eaabf-187">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="eaabf-187">Returns:</span></span>

<span data-ttu-id="eaabf-188">Тип: String</span><span class="sxs-lookup"><span data-stu-id="eaabf-188">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="eaabf-189">Пример</span><span class="sxs-lookup"><span data-stu-id="eaabf-189">Example</span></span>

```js
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-14"></a><span data-ttu-id="eaabf-190">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.4)}</span><span class="sxs-lookup"><span data-stu-id="eaabf-190">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.4)}</span></span>

<span data-ttu-id="eaabf-191">Получает словарь, содержащий сведения о локальном времени клиента.</span><span class="sxs-lookup"><span data-stu-id="eaabf-191">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="eaabf-192">Почтовое приложение для Outlook на настольном компьютере или в Интернете может использовать разные часовые пояса для дат и времени.</span><span class="sxs-lookup"><span data-stu-id="eaabf-192">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times.</span></span> <span data-ttu-id="eaabf-193">Outlook на рабочем столе использует часовой пояс клиентского компьютера; В Outlook в Интернете используется часовой пояс, установленный в центре администрирования Exchange.</span><span class="sxs-lookup"><span data-stu-id="eaabf-193">Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC).</span></span> <span data-ttu-id="eaabf-194">Значения даты и времени должны обрабатываться таким образом, чтобы значения, отображаемые в интерфейсе пользователя, всегда согласовывались с часовым поясом, ожидаемым пользователем.</span><span class="sxs-lookup"><span data-stu-id="eaabf-194">You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="eaabf-195">Если почтовое приложение запущено в Outlook на настольном клиенте `convertToLocalClientTime` , метод возвратит объект Dictionary со значениями, заданными для часового пояса клиентского компьютера.</span><span class="sxs-lookup"><span data-stu-id="eaabf-195">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone.</span></span> <span data-ttu-id="eaabf-196">Если почтовое приложение запущено в Outlook в Интернете, `convertToLocalClientTime` метод возвратит объект Dictionary со значениями, заданными в часовом поясе, заданном в центре администрирования Exchange.</span><span class="sxs-lookup"><span data-stu-id="eaabf-196">If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="eaabf-197">Параметры</span><span class="sxs-lookup"><span data-stu-id="eaabf-197">Parameters</span></span>

|<span data-ttu-id="eaabf-198">Имя</span><span class="sxs-lookup"><span data-stu-id="eaabf-198">Name</span></span>| <span data-ttu-id="eaabf-199">Тип</span><span class="sxs-lookup"><span data-stu-id="eaabf-199">Type</span></span>| <span data-ttu-id="eaabf-200">Описание</span><span class="sxs-lookup"><span data-stu-id="eaabf-200">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="eaabf-201">Date</span><span class="sxs-lookup"><span data-stu-id="eaabf-201">Date</span></span>|<span data-ttu-id="eaabf-202">Объект Date</span><span class="sxs-lookup"><span data-stu-id="eaabf-202">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="eaabf-203">Требования</span><span class="sxs-lookup"><span data-stu-id="eaabf-203">Requirements</span></span>

|<span data-ttu-id="eaabf-204">Требование</span><span class="sxs-lookup"><span data-stu-id="eaabf-204">Requirement</span></span>| <span data-ttu-id="eaabf-205">Значение</span><span class="sxs-lookup"><span data-stu-id="eaabf-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="eaabf-206">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="eaabf-206">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eaabf-207">1.0</span><span class="sxs-lookup"><span data-stu-id="eaabf-207">1.0</span></span>|
|[<span data-ttu-id="eaabf-208">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="eaabf-208">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="eaabf-209">ReadItem</span><span class="sxs-lookup"><span data-stu-id="eaabf-209">ReadItem</span></span>|
|[<span data-ttu-id="eaabf-210">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="eaabf-210">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="eaabf-211">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="eaabf-211">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="eaabf-212">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="eaabf-212">Returns:</span></span>

<span data-ttu-id="eaabf-213">Тип: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.4)</span><span class="sxs-lookup"><span data-stu-id="eaabf-213">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.4)</span></span>

<br>

---
---

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="eaabf-214">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="eaabf-214">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="eaabf-215">Преобразовывает идентификатор элемента в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="eaabf-215">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="eaabf-216">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="eaabf-216">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="eaabf-p107">Формат идентификаторов, извлекаемых через EWS или свойство `itemId`, отличается от формата API REST (таких как [API Почты Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)). Метод `convertToRestId` преобразовывает идентификатор в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="eaabf-p107">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="eaabf-219">Параметры</span><span class="sxs-lookup"><span data-stu-id="eaabf-219">Parameters</span></span>

|<span data-ttu-id="eaabf-220">Имя</span><span class="sxs-lookup"><span data-stu-id="eaabf-220">Name</span></span>| <span data-ttu-id="eaabf-221">Тип</span><span class="sxs-lookup"><span data-stu-id="eaabf-221">Type</span></span>| <span data-ttu-id="eaabf-222">Описание</span><span class="sxs-lookup"><span data-stu-id="eaabf-222">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="eaabf-223">String</span><span class="sxs-lookup"><span data-stu-id="eaabf-223">String</span></span>|<span data-ttu-id="eaabf-224">Идентификатор элемента в формате EWS</span><span class="sxs-lookup"><span data-stu-id="eaabf-224">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="eaabf-225">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="eaabf-225">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.4)|<span data-ttu-id="eaabf-226">Значение, определяющее версию REST API для Outlook, с которой будет использоваться преобразованный идентификатор.</span><span class="sxs-lookup"><span data-stu-id="eaabf-226">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="eaabf-227">Требования</span><span class="sxs-lookup"><span data-stu-id="eaabf-227">Requirements</span></span>

|<span data-ttu-id="eaabf-228">Требование</span><span class="sxs-lookup"><span data-stu-id="eaabf-228">Requirement</span></span>| <span data-ttu-id="eaabf-229">Значение</span><span class="sxs-lookup"><span data-stu-id="eaabf-229">Value</span></span>|
|---|---|
|[<span data-ttu-id="eaabf-230">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="eaabf-230">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eaabf-231">1.3</span><span class="sxs-lookup"><span data-stu-id="eaabf-231">1.3</span></span>|
|[<span data-ttu-id="eaabf-232">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="eaabf-232">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="eaabf-233">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="eaabf-233">Restricted</span></span>|
|[<span data-ttu-id="eaabf-234">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="eaabf-234">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="eaabf-235">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="eaabf-235">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="eaabf-236">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="eaabf-236">Returns:</span></span>

<span data-ttu-id="eaabf-237">Тип: String</span><span class="sxs-lookup"><span data-stu-id="eaabf-237">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="eaabf-238">Пример</span><span class="sxs-lookup"><span data-stu-id="eaabf-238">Example</span></span>

```js
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="eaabf-239">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="eaabf-239">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="eaabf-240">Получает объект Date из словаря, содержащего сведения о времени.</span><span class="sxs-lookup"><span data-stu-id="eaabf-240">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="eaabf-241">Метод `convertToUtcClientTime` преобразует словарь, содержащий локальную дату и время, в объект Date с правильными значениями локальной даты и времени.</span><span class="sxs-lookup"><span data-stu-id="eaabf-241">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="eaabf-242">Параметры</span><span class="sxs-lookup"><span data-stu-id="eaabf-242">Parameters</span></span>

|<span data-ttu-id="eaabf-243">Имя</span><span class="sxs-lookup"><span data-stu-id="eaabf-243">Name</span></span>| <span data-ttu-id="eaabf-244">Тип</span><span class="sxs-lookup"><span data-stu-id="eaabf-244">Type</span></span>| <span data-ttu-id="eaabf-245">Описание</span><span class="sxs-lookup"><span data-stu-id="eaabf-245">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="eaabf-246">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="eaabf-246">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.4)|<span data-ttu-id="eaabf-247">Значение локального времени для преобразования.</span><span class="sxs-lookup"><span data-stu-id="eaabf-247">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="eaabf-248">Требования</span><span class="sxs-lookup"><span data-stu-id="eaabf-248">Requirements</span></span>

|<span data-ttu-id="eaabf-249">Требование</span><span class="sxs-lookup"><span data-stu-id="eaabf-249">Requirement</span></span>| <span data-ttu-id="eaabf-250">Значение</span><span class="sxs-lookup"><span data-stu-id="eaabf-250">Value</span></span>|
|---|---|
|[<span data-ttu-id="eaabf-251">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="eaabf-251">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eaabf-252">1.0</span><span class="sxs-lookup"><span data-stu-id="eaabf-252">1.0</span></span>|
|[<span data-ttu-id="eaabf-253">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="eaabf-253">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="eaabf-254">ReadItem</span><span class="sxs-lookup"><span data-stu-id="eaabf-254">ReadItem</span></span>|
|[<span data-ttu-id="eaabf-255">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="eaabf-255">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="eaabf-256">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="eaabf-256">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="eaabf-257">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="eaabf-257">Returns:</span></span>

<span data-ttu-id="eaabf-258">Объект Date со временем в формате UTC.</span><span class="sxs-lookup"><span data-stu-id="eaabf-258">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="eaabf-259">Тип: Date</span><span class="sxs-lookup"><span data-stu-id="eaabf-259">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="eaabf-260">Пример</span><span class="sxs-lookup"><span data-stu-id="eaabf-260">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="eaabf-261">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="eaabf-261">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="eaabf-262">Отображает имеющуюся встречу из календаря.</span><span class="sxs-lookup"><span data-stu-id="eaabf-262">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="eaabf-263">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="eaabf-263">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="eaabf-264">Метод `displayAppointmentForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее сведения календаря о существующей встрече.</span><span class="sxs-lookup"><span data-stu-id="eaabf-264">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="eaabf-265">В Outlook на Mac Этот метод можно использовать для отображения одной встречи, которая не является частью повторяющегося ряда, или главной встречи повторяющейся серии, но невозможно отобразить экземпляр ряда.</span><span class="sxs-lookup"><span data-stu-id="eaabf-265">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series.</span></span> <span data-ttu-id="eaabf-266">Это связано с тем, что в Outlook на Mac-адресе невозможно получить доступ к свойствам (включая идентификатор элемента) повторяющихся рядов.</span><span class="sxs-lookup"><span data-stu-id="eaabf-266">This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="eaabf-267">В Outlook в Интернете этот метод открывает указанную форму, только если текст формы меньше или равен 32 КБ числу символов.</span><span class="sxs-lookup"><span data-stu-id="eaabf-267">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="eaabf-268">Если указанный идентификатор элемента не определяет существующую встречу, на клиентском компьютере или устройстве открывается пустая страница, и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="eaabf-268">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="eaabf-269">Параметры</span><span class="sxs-lookup"><span data-stu-id="eaabf-269">Parameters</span></span>

|<span data-ttu-id="eaabf-270">Имя</span><span class="sxs-lookup"><span data-stu-id="eaabf-270">Name</span></span>| <span data-ttu-id="eaabf-271">Тип</span><span class="sxs-lookup"><span data-stu-id="eaabf-271">Type</span></span>| <span data-ttu-id="eaabf-272">Описание</span><span class="sxs-lookup"><span data-stu-id="eaabf-272">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="eaabf-273">String</span><span class="sxs-lookup"><span data-stu-id="eaabf-273">String</span></span>|<span data-ttu-id="eaabf-274">Идентификатор веб-служб Exchange для существующей встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="eaabf-274">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="eaabf-275">Требования</span><span class="sxs-lookup"><span data-stu-id="eaabf-275">Requirements</span></span>

|<span data-ttu-id="eaabf-276">Требование</span><span class="sxs-lookup"><span data-stu-id="eaabf-276">Requirement</span></span>| <span data-ttu-id="eaabf-277">Значение</span><span class="sxs-lookup"><span data-stu-id="eaabf-277">Value</span></span>|
|---|---|
|[<span data-ttu-id="eaabf-278">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="eaabf-278">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eaabf-279">1.0</span><span class="sxs-lookup"><span data-stu-id="eaabf-279">1.0</span></span>|
|[<span data-ttu-id="eaabf-280">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="eaabf-280">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="eaabf-281">ReadItem</span><span class="sxs-lookup"><span data-stu-id="eaabf-281">ReadItem</span></span>|
|[<span data-ttu-id="eaabf-282">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="eaabf-282">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="eaabf-283">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="eaabf-283">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="eaabf-284">Пример</span><span class="sxs-lookup"><span data-stu-id="eaabf-284">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="eaabf-285">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="eaabf-285">displayMessageForm(itemId)</span></span>

<span data-ttu-id="eaabf-286">Отображает имеющееся сообщение.</span><span class="sxs-lookup"><span data-stu-id="eaabf-286">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="eaabf-287">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="eaabf-287">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="eaabf-288">Метод `displayMessageForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее существующее сообщение.</span><span class="sxs-lookup"><span data-stu-id="eaabf-288">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="eaabf-289">В Outlook в Интернете этот метод открывает указанную форму только в том случае, если размер текста формы меньше или равен 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="eaabf-289">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="eaabf-290">Если указанный идентификатор элемента не определяет существующее сообщение, окно на клиентском компьютере не открывается и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="eaabf-290">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="eaabf-p109">Не используйте `displayMessageForm` с параметром `itemId`, который представляет собой встречу. Используйте метод `displayAppointmentForm`, чтобы отобразить сведения о существующей встрече, а метод `displayNewAppointmentForm` — для отображения формы создания встречи.</span><span class="sxs-lookup"><span data-stu-id="eaabf-p109">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="eaabf-293">Параметры</span><span class="sxs-lookup"><span data-stu-id="eaabf-293">Parameters</span></span>

|<span data-ttu-id="eaabf-294">Имя</span><span class="sxs-lookup"><span data-stu-id="eaabf-294">Name</span></span>| <span data-ttu-id="eaabf-295">Тип</span><span class="sxs-lookup"><span data-stu-id="eaabf-295">Type</span></span>| <span data-ttu-id="eaabf-296">Описание</span><span class="sxs-lookup"><span data-stu-id="eaabf-296">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="eaabf-297">String</span><span class="sxs-lookup"><span data-stu-id="eaabf-297">String</span></span>|<span data-ttu-id="eaabf-298">Идентификатор веб-служб Exchange для существующего сообщения.</span><span class="sxs-lookup"><span data-stu-id="eaabf-298">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="eaabf-299">Требования</span><span class="sxs-lookup"><span data-stu-id="eaabf-299">Requirements</span></span>

|<span data-ttu-id="eaabf-300">Требование</span><span class="sxs-lookup"><span data-stu-id="eaabf-300">Requirement</span></span>| <span data-ttu-id="eaabf-301">Значение</span><span class="sxs-lookup"><span data-stu-id="eaabf-301">Value</span></span>|
|---|---|
|[<span data-ttu-id="eaabf-302">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="eaabf-302">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eaabf-303">1.0</span><span class="sxs-lookup"><span data-stu-id="eaabf-303">1.0</span></span>|
|[<span data-ttu-id="eaabf-304">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="eaabf-304">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="eaabf-305">ReadItem</span><span class="sxs-lookup"><span data-stu-id="eaabf-305">ReadItem</span></span>|
|[<span data-ttu-id="eaabf-306">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="eaabf-306">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="eaabf-307">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="eaabf-307">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="eaabf-308">Пример</span><span class="sxs-lookup"><span data-stu-id="eaabf-308">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="eaabf-309">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="eaabf-309">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="eaabf-310">Отображает форму для создания новой встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="eaabf-310">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="eaabf-311">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="eaabf-311">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="eaabf-p110">Метод `displayNewAppointmentForm` открывает форму, в которой пользователь может создать встречу или собрание. Если параметры заданы, поля формы встречи автоматически заполняются их содержимым.</span><span class="sxs-lookup"><span data-stu-id="eaabf-p110">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="eaabf-314">В Outlook в Интернете и на мобильных устройствах этот метод всегда отображает форму с полем участников.</span><span class="sxs-lookup"><span data-stu-id="eaabf-314">In Outlook on the web and mobile devices, this method always displays a form with an attendees field.</span></span> <span data-ttu-id="eaabf-315">Если вы не укажете участников в качестве входных аргументов, метод отображает форму с кнопкой **Сохранить**.</span><span class="sxs-lookup"><span data-stu-id="eaabf-315">If you do not specify any attendees as input arguments, the method displays a form with a **Save** button.</span></span> <span data-ttu-id="eaabf-316">Если вы укажете участников, форма будет включать участников и кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="eaabf-316">If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="eaabf-p112">Если вы укажете участников или ресурсы с помощью параметра `requiredAttendees`, `optionalAttendees` или `resources` в клиенте Outlook с расширенными возможностями и Outlook RT, этот метод отобразит форму собрания с кнопкой **Отправить**. Если не указать получателей, этот метод отобразит форму встречи с кнопкой **Сохранить и закрыть**.</span><span class="sxs-lookup"><span data-stu-id="eaabf-p112">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="eaabf-319">Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.</span><span class="sxs-lookup"><span data-stu-id="eaabf-319">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="eaabf-320">Параметры</span><span class="sxs-lookup"><span data-stu-id="eaabf-320">Parameters</span></span>

|<span data-ttu-id="eaabf-321">Имя</span><span class="sxs-lookup"><span data-stu-id="eaabf-321">Name</span></span>| <span data-ttu-id="eaabf-322">Тип</span><span class="sxs-lookup"><span data-stu-id="eaabf-322">Type</span></span>| <span data-ttu-id="eaabf-323">Описание</span><span class="sxs-lookup"><span data-stu-id="eaabf-323">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="eaabf-324">Object</span><span class="sxs-lookup"><span data-stu-id="eaabf-324">Object</span></span> | <span data-ttu-id="eaabf-325">Словарь параметров, описывающий новую встречу.</span><span class="sxs-lookup"><span data-stu-id="eaabf-325">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="eaabf-326">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)&gt;</span><span class="sxs-lookup"><span data-stu-id="eaabf-326">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)&gt;</span></span> | <span data-ttu-id="eaabf-p113">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из обязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="eaabf-p113">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="eaabf-329">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)&gt;</span><span class="sxs-lookup"><span data-stu-id="eaabf-329">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.4)&gt;</span></span> | <span data-ttu-id="eaabf-p114">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из необязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="eaabf-p114">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="eaabf-332">Date</span><span class="sxs-lookup"><span data-stu-id="eaabf-332">Date</span></span> | <span data-ttu-id="eaabf-333">Объект `Date`, указывающий дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="eaabf-333">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="eaabf-334">Date</span><span class="sxs-lookup"><span data-stu-id="eaabf-334">Date</span></span> | <span data-ttu-id="eaabf-335">Объект `Date`, указывающий дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="eaabf-335">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="eaabf-336">Строка</span><span class="sxs-lookup"><span data-stu-id="eaabf-336">String</span></span> | <span data-ttu-id="eaabf-p115">Строка со сведениями о месте встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="eaabf-p115">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="eaabf-339">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="eaabf-339">Array.&lt;String&gt;</span></span> | <span data-ttu-id="eaabf-p116">Массив строк, содержащий необходимые для встречи ресурсы. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="eaabf-p116">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="eaabf-342">String</span><span class="sxs-lookup"><span data-stu-id="eaabf-342">String</span></span> | <span data-ttu-id="eaabf-p117">Строка с темой встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="eaabf-p117">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="eaabf-345">String</span><span class="sxs-lookup"><span data-stu-id="eaabf-345">String</span></span> | <span data-ttu-id="eaabf-p118">Текст сообщения о встрече. Максимальный размер содержимого сообщения — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="eaabf-p118">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="eaabf-348">Требования</span><span class="sxs-lookup"><span data-stu-id="eaabf-348">Requirements</span></span>

|<span data-ttu-id="eaabf-349">Требование</span><span class="sxs-lookup"><span data-stu-id="eaabf-349">Requirement</span></span>| <span data-ttu-id="eaabf-350">Значение</span><span class="sxs-lookup"><span data-stu-id="eaabf-350">Value</span></span>|
|---|---|
|[<span data-ttu-id="eaabf-351">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="eaabf-351">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eaabf-352">1.0</span><span class="sxs-lookup"><span data-stu-id="eaabf-352">1.0</span></span>|
|[<span data-ttu-id="eaabf-353">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="eaabf-353">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="eaabf-354">ReadItem</span><span class="sxs-lookup"><span data-stu-id="eaabf-354">ReadItem</span></span>|
|[<span data-ttu-id="eaabf-355">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="eaabf-355">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="eaabf-356">Чтение</span><span class="sxs-lookup"><span data-stu-id="eaabf-356">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="eaabf-357">Пример</span><span class="sxs-lookup"><span data-stu-id="eaabf-357">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="eaabf-358">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="eaabf-358">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="eaabf-359">Получает строку, содержащую маркер, используемый для получения вложения или элемента с Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="eaabf-359">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="eaabf-p119">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="eaabf-p119">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="eaabf-p120">Вы можете передать сторонней системе маркер и идентификатор вложения или элемента. Сторонняя система использует этот маркер как маркер авторизации, чтобы вызвать операцию [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) или [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) веб-служб Exchange для возврата вложения или элемента. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="eaabf-p120">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="eaabf-365">Для вызова метода `getCallbackTokenAsync` в режиме чтения манифесте приложения должно быть указано разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="eaabf-365">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="eaabf-p121">Чтобы получить идентификатор элемента для передачи в метод `getCallbackTokenAsync`, в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="eaabf-p121">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="eaabf-368">Параметры</span><span class="sxs-lookup"><span data-stu-id="eaabf-368">Parameters</span></span>

|<span data-ttu-id="eaabf-369">Имя</span><span class="sxs-lookup"><span data-stu-id="eaabf-369">Name</span></span>| <span data-ttu-id="eaabf-370">Тип</span><span class="sxs-lookup"><span data-stu-id="eaabf-370">Type</span></span>| <span data-ttu-id="eaabf-371">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="eaabf-371">Attributes</span></span>| <span data-ttu-id="eaabf-372">Описание</span><span class="sxs-lookup"><span data-stu-id="eaabf-372">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="eaabf-373">функция</span><span class="sxs-lookup"><span data-stu-id="eaabf-373">function</span></span>||<span data-ttu-id="eaabf-374">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="eaabf-374">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="eaabf-375">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="eaabf-375">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="eaabf-376">При возникновении ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут содержать дополнительные сведения.</span><span class="sxs-lookup"><span data-stu-id="eaabf-376">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="eaabf-377">Объект</span><span class="sxs-lookup"><span data-stu-id="eaabf-377">Object</span></span>| <span data-ttu-id="eaabf-378">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="eaabf-378">&lt;optional&gt;</span></span>|<span data-ttu-id="eaabf-379">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="eaabf-379">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="eaabf-380">Ошибки</span><span class="sxs-lookup"><span data-stu-id="eaabf-380">Errors</span></span>

|<span data-ttu-id="eaabf-381">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="eaabf-381">Error code</span></span>|<span data-ttu-id="eaabf-382">Описание</span><span class="sxs-lookup"><span data-stu-id="eaabf-382">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="eaabf-383">Запрос не выполнен.</span><span class="sxs-lookup"><span data-stu-id="eaabf-383">The request has failed.</span></span> <span data-ttu-id="eaabf-384">Просмотрите объект Diagnostics для кода ошибки HTTP.</span><span class="sxs-lookup"><span data-stu-id="eaabf-384">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="eaabf-385">Сервер Exchange возвратил ошибку.</span><span class="sxs-lookup"><span data-stu-id="eaabf-385">The Exchange server returned an error.</span></span> <span data-ttu-id="eaabf-386">Дополнительные сведения можно найти в объекте диагностики.</span><span class="sxs-lookup"><span data-stu-id="eaabf-386">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="eaabf-387">Пользователь больше не подключен к сети.</span><span class="sxs-lookup"><span data-stu-id="eaabf-387">The user is no longer connected to the network.</span></span> <span data-ttu-id="eaabf-388">Проверьте сетевое подключение и повторите попытку.</span><span class="sxs-lookup"><span data-stu-id="eaabf-388">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="eaabf-389">Требования</span><span class="sxs-lookup"><span data-stu-id="eaabf-389">Requirements</span></span>

|<span data-ttu-id="eaabf-390">Требование</span><span class="sxs-lookup"><span data-stu-id="eaabf-390">Requirement</span></span>| <span data-ttu-id="eaabf-391">Значение</span><span class="sxs-lookup"><span data-stu-id="eaabf-391">Value</span></span>|
|---|---|
|[<span data-ttu-id="eaabf-392">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="eaabf-392">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eaabf-393">1.0</span><span class="sxs-lookup"><span data-stu-id="eaabf-393">1.0</span></span>|
|[<span data-ttu-id="eaabf-394">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="eaabf-394">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="eaabf-395">ReadItem</span><span class="sxs-lookup"><span data-stu-id="eaabf-395">ReadItem</span></span>|
|[<span data-ttu-id="eaabf-396">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="eaabf-396">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="eaabf-397">Создание и чтение</span><span class="sxs-lookup"><span data-stu-id="eaabf-397">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="eaabf-398">Пример</span><span class="sxs-lookup"><span data-stu-id="eaabf-398">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="eaabf-399">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="eaabf-399">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="eaabf-400">Получает маркер, идентифицирующий пользователя и надстройку Office.</span><span class="sxs-lookup"><span data-stu-id="eaabf-400">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="eaabf-401">Метод `getUserIdentityTokenAsync` возвращает маркер, который можно использовать для идентификации, а также [проверки подлинности надстройки и пользователя в сторонней системе](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="eaabf-401">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="eaabf-402">Параметры</span><span class="sxs-lookup"><span data-stu-id="eaabf-402">Parameters</span></span>

|<span data-ttu-id="eaabf-403">Имя</span><span class="sxs-lookup"><span data-stu-id="eaabf-403">Name</span></span>| <span data-ttu-id="eaabf-404">Тип</span><span class="sxs-lookup"><span data-stu-id="eaabf-404">Type</span></span>| <span data-ttu-id="eaabf-405">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="eaabf-405">Attributes</span></span>| <span data-ttu-id="eaabf-406">Описание</span><span class="sxs-lookup"><span data-stu-id="eaabf-406">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="eaabf-407">функция</span><span class="sxs-lookup"><span data-stu-id="eaabf-407">function</span></span>||<span data-ttu-id="eaabf-408">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="eaabf-408">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="eaabf-409">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="eaabf-409">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="eaabf-410">При возникновении ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут содержать дополнительные сведения.</span><span class="sxs-lookup"><span data-stu-id="eaabf-410">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="eaabf-411">Объект</span><span class="sxs-lookup"><span data-stu-id="eaabf-411">Object</span></span>| <span data-ttu-id="eaabf-412">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="eaabf-412">&lt;optional&gt;</span></span>|<span data-ttu-id="eaabf-413">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="eaabf-413">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="eaabf-414">Ошибки</span><span class="sxs-lookup"><span data-stu-id="eaabf-414">Errors</span></span>

|<span data-ttu-id="eaabf-415">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="eaabf-415">Error code</span></span>|<span data-ttu-id="eaabf-416">Описание</span><span class="sxs-lookup"><span data-stu-id="eaabf-416">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="eaabf-417">Запрос не выполнен.</span><span class="sxs-lookup"><span data-stu-id="eaabf-417">The request has failed.</span></span> <span data-ttu-id="eaabf-418">Просмотрите объект Diagnostics для кода ошибки HTTP.</span><span class="sxs-lookup"><span data-stu-id="eaabf-418">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="eaabf-419">Сервер Exchange возвратил ошибку.</span><span class="sxs-lookup"><span data-stu-id="eaabf-419">The Exchange server returned an error.</span></span> <span data-ttu-id="eaabf-420">Дополнительные сведения можно найти в объекте диагностики.</span><span class="sxs-lookup"><span data-stu-id="eaabf-420">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="eaabf-421">Пользователь больше не подключен к сети.</span><span class="sxs-lookup"><span data-stu-id="eaabf-421">The user is no longer connected to the network.</span></span> <span data-ttu-id="eaabf-422">Проверьте сетевое подключение и повторите попытку.</span><span class="sxs-lookup"><span data-stu-id="eaabf-422">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="eaabf-423">Требования</span><span class="sxs-lookup"><span data-stu-id="eaabf-423">Requirements</span></span>

|<span data-ttu-id="eaabf-424">Требование</span><span class="sxs-lookup"><span data-stu-id="eaabf-424">Requirement</span></span>| <span data-ttu-id="eaabf-425">Значение</span><span class="sxs-lookup"><span data-stu-id="eaabf-425">Value</span></span>|
|---|---|
|[<span data-ttu-id="eaabf-426">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="eaabf-426">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eaabf-427">1.0</span><span class="sxs-lookup"><span data-stu-id="eaabf-427">1.0</span></span>|
|[<span data-ttu-id="eaabf-428">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="eaabf-428">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="eaabf-429">ReadItem</span><span class="sxs-lookup"><span data-stu-id="eaabf-429">ReadItem</span></span>|
|[<span data-ttu-id="eaabf-430">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="eaabf-430">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="eaabf-431">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="eaabf-431">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="eaabf-432">Пример</span><span class="sxs-lookup"><span data-stu-id="eaabf-432">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="eaabf-433">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="eaabf-433">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="eaabf-434">Выполняет асинхронный запрос для веб-служб Exchange (EWS) на сервере Exchange Server, на котором размещен почтовый ящик пользователя.</span><span class="sxs-lookup"><span data-stu-id="eaabf-434">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="eaabf-435">Этот метод не поддерживается в следующих сценариях:</span><span class="sxs-lookup"><span data-stu-id="eaabf-435">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="eaabf-436">В Outlook на iOS или Android</span><span class="sxs-lookup"><span data-stu-id="eaabf-436">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="eaabf-437">Если надстройка загружается в почтовый ящик Gmail.</span><span class="sxs-lookup"><span data-stu-id="eaabf-437">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="eaabf-438">В таких случаях надстройка должна [использовать REST API](/outlook/add-ins/use-rest-api) для доступа к почтовому ящику пользователя.</span><span class="sxs-lookup"><span data-stu-id="eaabf-438">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="eaabf-439">Метод `makeEwsRequestAsync` отправляет запрос EWS от имени надстройки в Exchange.</span><span class="sxs-lookup"><span data-stu-id="eaabf-439">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="eaabf-440">Список поддерживаемых операций EWS см. в статье [Вызов веб-служб из надстройки Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support).</span><span class="sxs-lookup"><span data-stu-id="eaabf-440">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="eaabf-441">С помощью метода `makeEwsRequestAsync` невозможно запрашивать элементы, связанные с папкой.</span><span class="sxs-lookup"><span data-stu-id="eaabf-441">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="eaabf-442">В запросе XML должна быть указана кодировка UTF-8.</span><span class="sxs-lookup"><span data-stu-id="eaabf-442">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="eaabf-p129">У вашей надстройки должно быть разрешение **ReadWriteMailbox** для использования метода `makeEwsRequestAsync`. Сведения об использовании разрешения **ReadWriteMailbox** и операций EWS, которые можно вызывать с помощью метода `makeEwsRequestAsync`, см. в статье [Указание разрешений для доступа почтовой надстройки к почтовому ящику пользователя](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="eaabf-p129">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="eaabf-445">Администратор сервера должен установить значение true для параметра `OAuthAuthentication` в каталоге сервера клиентского доступа EWS, чтобы метод `makeEwsRequestAsync` мог выполнять запросы EWS.</span><span class="sxs-lookup"><span data-stu-id="eaabf-445">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="eaabf-446">Различия версий</span><span class="sxs-lookup"><span data-stu-id="eaabf-446">Version differences</span></span>

<span data-ttu-id="eaabf-447">Если вы используете метод `makeEwsRequestAsync` в почтовых приложениях, которые выполняются в Outlook версии более ранней, чем 15.0.4535.1004, указывайте кодировку `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="eaabf-447">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="eaabf-p130">Значение кодировки не нужно указывать, если почтовое приложение выполняется в Outlook в Интернете. Чтобы определить, выполняется ли приложение в Outlook или Outlook в Интернете, используйте свойство mailbox.diagnostics.hostName. Используемую версию Outlook можно определить с помощью свойства mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="eaabf-p130">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="eaabf-451">Параметры</span><span class="sxs-lookup"><span data-stu-id="eaabf-451">Parameters</span></span>

|<span data-ttu-id="eaabf-452">Имя</span><span class="sxs-lookup"><span data-stu-id="eaabf-452">Name</span></span>| <span data-ttu-id="eaabf-453">Тип</span><span class="sxs-lookup"><span data-stu-id="eaabf-453">Type</span></span>| <span data-ttu-id="eaabf-454">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="eaabf-454">Attributes</span></span>| <span data-ttu-id="eaabf-455">Описание</span><span class="sxs-lookup"><span data-stu-id="eaabf-455">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="eaabf-456">String</span><span class="sxs-lookup"><span data-stu-id="eaabf-456">String</span></span>||<span data-ttu-id="eaabf-457">Запрос EWS.</span><span class="sxs-lookup"><span data-stu-id="eaabf-457">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="eaabf-458">function</span><span class="sxs-lookup"><span data-stu-id="eaabf-458">function</span></span>||<span data-ttu-id="eaabf-459">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="eaabf-459">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="eaabf-460">Результат XML вызова EWS указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="eaabf-460">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="eaabf-461">Если размер результата превышает 1 МБ, возвращается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="eaabf-461">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="eaabf-462">Объект</span><span class="sxs-lookup"><span data-stu-id="eaabf-462">Object</span></span>| <span data-ttu-id="eaabf-463">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="eaabf-463">&lt;optional&gt;</span></span>|<span data-ttu-id="eaabf-464">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="eaabf-464">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="eaabf-465">Требования</span><span class="sxs-lookup"><span data-stu-id="eaabf-465">Requirements</span></span>

|<span data-ttu-id="eaabf-466">Требование</span><span class="sxs-lookup"><span data-stu-id="eaabf-466">Requirement</span></span>| <span data-ttu-id="eaabf-467">Значение</span><span class="sxs-lookup"><span data-stu-id="eaabf-467">Value</span></span>|
|---|---|
|[<span data-ttu-id="eaabf-468">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="eaabf-468">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="eaabf-469">1.0</span><span class="sxs-lookup"><span data-stu-id="eaabf-469">1.0</span></span>|
|[<span data-ttu-id="eaabf-470">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="eaabf-470">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="eaabf-471">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="eaabf-471">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="eaabf-472">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="eaabf-472">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="eaabf-473">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="eaabf-473">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="eaabf-474">Пример</span><span class="sxs-lookup"><span data-stu-id="eaabf-474">Example</span></span>

<span data-ttu-id="eaabf-475">В приведенном ниже примере вызывается `makeEwsRequestAsync` для получения темы элемента с помощью операции `GetItem`.</span><span class="sxs-lookup"><span data-stu-id="eaabf-475">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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
