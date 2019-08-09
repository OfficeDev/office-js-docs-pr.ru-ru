---
title: Office. Context. Mailbox — набор обязательных элементов 1,3
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: acc304e302e3bb4d912ecbafee51cc35c88c091c
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268672"
---
# <a name="mailbox"></a><span data-ttu-id="bd8f2-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="bd8f2-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="bd8f2-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="bd8f2-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="bd8f2-104">Предоставляет доступ к объектной модели надстройки Outlook для Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="bd8f2-105">Требования</span><span class="sxs-lookup"><span data-stu-id="bd8f2-105">Requirements</span></span>

|<span data-ttu-id="bd8f2-106">Требование</span><span class="sxs-lookup"><span data-stu-id="bd8f2-106">Requirement</span></span>| <span data-ttu-id="bd8f2-107">Значение</span><span class="sxs-lookup"><span data-stu-id="bd8f2-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd8f2-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="bd8f2-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bd8f2-109">1.0</span><span class="sxs-lookup"><span data-stu-id="bd8f2-109">1.0</span></span>|
|[<span data-ttu-id="bd8f2-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="bd8f2-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bd8f2-111">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="bd8f2-111">Restricted</span></span>|
|[<span data-ttu-id="bd8f2-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="bd8f2-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bd8f2-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="bd8f2-113">Compose or Read</span></span>|

<span data-ttu-id="bd8f2-114">| [ewsUrl](#ewsurl-string) | Участник | | [конверттоевсид](#converttoewsiditemid-restversion--string) | Метод | | [convertToLocalClientTime](#converttolocalclienttimetimevalue--localclienttime) | Метод | | [convertToRestId](#converttorestiditemid-restversion--string) | Метод | | [convertToUtcClientTime](#converttoutcclienttimeinput--date) | Метод | | [displayAppointmentForm](#displayappointmentformitemid) | Метод | | [displayMessageForm](#displaymessageformitemid) | Метод | | [displayNewAppointmentForm](#displaynewappointmentformparameters) | Метод | | [getCallbackTokenAsync](#getcallbacktokenasynccallback-usercontext) | Метод | | [getUserIdentityTokenAsync](#getuseridentitytokenasynccallback-usercontext) | Метод | | [makeEwsRequestAsync](#makeewsrequestasyncdata-callback-usercontext) | Метод |</span><span class="sxs-lookup"><span data-stu-id="bd8f2-114">| [ewsUrl](#ewsurl-string) | Member | | [convertToEwsId](#converttoewsiditemid-restversion--string) | Method | | [convertToLocalClientTime](#converttolocalclienttimetimevalue--localclienttime) | Method | | [convertToRestId](#converttorestiditemid-restversion--string) | Method | | [convertToUtcClientTime](#converttoutcclienttimeinput--date) | Method | | [displayAppointmentForm](#displayappointmentformitemid) | Method | | [displayMessageForm](#displaymessageformitemid) | Method | | [displayNewAppointmentForm](#displaynewappointmentformparameters) | Method | | [getCallbackTokenAsync](#getcallbacktokenasynccallback-usercontext) | Method | | [getUserIdentityTokenAsync](#getuseridentitytokenasynccallback-usercontext) | Method | | [makeEwsRequestAsync](#makeewsrequestasyncdata-callback-usercontext) | Method |</span></span>

### <a name="namespaces"></a><span data-ttu-id="bd8f2-115">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="bd8f2-115">Namespaces</span></span>

<span data-ttu-id="bd8f2-116">[diagnostics](Office.context.mailbox.diagnostics.md). Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-116">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="bd8f2-117">[item](Office.context.mailbox.item.md). Предоставляет методы и свойства для доступа к сообщению или встрече в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-117">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="bd8f2-118">[userProfile](Office.context.mailbox.userProfile.md). Предоставляет сведения о пользователе в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-118">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="bd8f2-119">Элементы</span><span class="sxs-lookup"><span data-stu-id="bd8f2-119">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="bd8f2-120">ewsUrl: строка</span><span class="sxs-lookup"><span data-stu-id="bd8f2-120">ewsUrl: String</span></span>

<span data-ttu-id="bd8f2-121">Получает URL-адрес конечной точки веб-служб Exchange (EWS) для этой учетной записи электронной почты.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-121">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account.</span></span> <span data-ttu-id="bd8f2-122">Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-122">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="bd8f2-123">Этот элемент не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-123">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="bd8f2-p102">Удаленная служба может использовать значение `ewsUrl`, чтобы выполнять вызовы EWS для почтового ящика пользователя. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="bd8f2-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="bd8f2-126">Чтобы вызвать элемент `ewsUrl` в режиме чтения, в манифесте приложения должно быть указано разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-126">Your app must have the **ReadItem** permission specified in its manifest to call the `ewsUrl` member in read mode.</span></span>

<span data-ttu-id="bd8f2-p103">Перед использованием элемента `ewsUrl` в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-p103">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method before you can use the `ewsUrl` member. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="type"></a><span data-ttu-id="bd8f2-129">Тип</span><span class="sxs-lookup"><span data-stu-id="bd8f2-129">Type</span></span>

*   <span data-ttu-id="bd8f2-130">String</span><span class="sxs-lookup"><span data-stu-id="bd8f2-130">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bd8f2-131">Требования</span><span class="sxs-lookup"><span data-stu-id="bd8f2-131">Requirements</span></span>

|<span data-ttu-id="bd8f2-132">Требование</span><span class="sxs-lookup"><span data-stu-id="bd8f2-132">Requirement</span></span>| <span data-ttu-id="bd8f2-133">Значение</span><span class="sxs-lookup"><span data-stu-id="bd8f2-133">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd8f2-134">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="bd8f2-134">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bd8f2-135">1.0</span><span class="sxs-lookup"><span data-stu-id="bd8f2-135">1.0</span></span>|
|[<span data-ttu-id="bd8f2-136">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="bd8f2-136">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bd8f2-137">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd8f2-137">ReadItem</span></span>|
|[<span data-ttu-id="bd8f2-138">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="bd8f2-138">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bd8f2-139">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="bd8f2-139">Compose or Read</span></span>|

### <a name="methods"></a><span data-ttu-id="bd8f2-140">Методы</span><span class="sxs-lookup"><span data-stu-id="bd8f2-140">Methods</span></span>

#### <a name="converttoewsiditemid-restversion--string"></a><span data-ttu-id="bd8f2-141">convertToEwsId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="bd8f2-141">convertToEwsId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="bd8f2-142">Преобразовывает идентификатор элемента из формата REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-142">Converts an item ID formatted for REST into EWS format.</span></span>

> [!NOTE]
> <span data-ttu-id="bd8f2-143">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-143">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="bd8f2-p104">Формат идентификаторов, извлекаемых через API REST (например, [API Почты Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)), отличается от формата веб-служб Exchange (EWS). Метод `convertToEwsId` преобразовывает идентификатор в формате REST в формат EWS.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-p104">Item IDs retrieved via a REST API (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)) use a different format than the format used by Exchange Web Services (EWS). The `convertToEwsId` method converts a REST-formatted ID into the proper format for EWS.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bd8f2-146">Параметры</span><span class="sxs-lookup"><span data-stu-id="bd8f2-146">Parameters</span></span>

|<span data-ttu-id="bd8f2-147">Имя</span><span class="sxs-lookup"><span data-stu-id="bd8f2-147">Name</span></span>| <span data-ttu-id="bd8f2-148">Тип</span><span class="sxs-lookup"><span data-stu-id="bd8f2-148">Type</span></span>| <span data-ttu-id="bd8f2-149">Описание</span><span class="sxs-lookup"><span data-stu-id="bd8f2-149">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="bd8f2-150">String</span><span class="sxs-lookup"><span data-stu-id="bd8f2-150">String</span></span>|<span data-ttu-id="bd8f2-151">Идентификатор элемента в формате REST API для Outlook</span><span class="sxs-lookup"><span data-stu-id="bd8f2-151">An item ID formatted for the Outlook REST APIs</span></span>|
|`restVersion`| [<span data-ttu-id="bd8f2-152">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="bd8f2-152">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.3)|<span data-ttu-id="bd8f2-153">Значение, определяющее версию REST API для Outlook, которая используется для извлечения идентификатора элемента.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-153">A value indicating the version of the Outlook REST API used to retrieve the item ID.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bd8f2-154">Требования</span><span class="sxs-lookup"><span data-stu-id="bd8f2-154">Requirements</span></span>

|<span data-ttu-id="bd8f2-155">Требование</span><span class="sxs-lookup"><span data-stu-id="bd8f2-155">Requirement</span></span>| <span data-ttu-id="bd8f2-156">Значение</span><span class="sxs-lookup"><span data-stu-id="bd8f2-156">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd8f2-157">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="bd8f2-157">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bd8f2-158">1.3</span><span class="sxs-lookup"><span data-stu-id="bd8f2-158">1.3</span></span>|
|[<span data-ttu-id="bd8f2-159">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="bd8f2-159">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bd8f2-160">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="bd8f2-160">Restricted</span></span>|
|[<span data-ttu-id="bd8f2-161">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="bd8f2-161">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bd8f2-162">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="bd8f2-162">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="bd8f2-163">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="bd8f2-163">Returns:</span></span>

<span data-ttu-id="bd8f2-164">Тип: String</span><span class="sxs-lookup"><span data-stu-id="bd8f2-164">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="bd8f2-165">Пример</span><span class="sxs-lookup"><span data-stu-id="bd8f2-165">Example</span></span>

```javascript
// Get an item's ID from a REST API.
var restId = 'AAMkAGVlOTZjNTM3LW...';

// Treat restId as coming from the v2.0 version of the Outlook Mail API.
var ewsId = Office.context.mailbox.convertToEwsId(restId, Office.MailboxEnums.RestVersion.v2_0);
```

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-13"></a><span data-ttu-id="bd8f2-166">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)}</span><span class="sxs-lookup"><span data-stu-id="bd8f2-166">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)}</span></span>

<span data-ttu-id="bd8f2-167">Получает словарь, содержащий сведения о локальном времени клиента.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-167">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="bd8f2-168">Почтовое приложение для Outlook на настольном компьютере или в Интернете может использовать разные часовые пояса для дат и времени.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-168">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times.</span></span> <span data-ttu-id="bd8f2-169">Outlook на рабочем столе использует часовой пояс клиентского компьютера; В Outlook в Интернете используется часовой пояс, установленный в центре администрирования Exchange.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-169">Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC).</span></span> <span data-ttu-id="bd8f2-170">Значения даты и времени должны обрабатываться таким образом, чтобы значения, отображаемые в интерфейсе пользователя, всегда согласовывались с часовым поясом, ожидаемым пользователем.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-170">You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="bd8f2-171">Если почтовое приложение запущено в Outlook на настольном клиенте `convertToLocalClientTime` , метод возвратит объект Dictionary со значениями, заданными для часового пояса клиентского компьютера.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-171">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone.</span></span> <span data-ttu-id="bd8f2-172">Если почтовое приложение запущено в Outlook в Интернете, `convertToLocalClientTime` метод возвратит объект Dictionary со значениями, заданными в часовом поясе, заданном в центре администрирования Exchange.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-172">If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bd8f2-173">Параметры</span><span class="sxs-lookup"><span data-stu-id="bd8f2-173">Parameters</span></span>

|<span data-ttu-id="bd8f2-174">Имя</span><span class="sxs-lookup"><span data-stu-id="bd8f2-174">Name</span></span>| <span data-ttu-id="bd8f2-175">Тип</span><span class="sxs-lookup"><span data-stu-id="bd8f2-175">Type</span></span>| <span data-ttu-id="bd8f2-176">Описание</span><span class="sxs-lookup"><span data-stu-id="bd8f2-176">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="bd8f2-177">Date</span><span class="sxs-lookup"><span data-stu-id="bd8f2-177">Date</span></span>|<span data-ttu-id="bd8f2-178">Объект Date</span><span class="sxs-lookup"><span data-stu-id="bd8f2-178">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bd8f2-179">Требования</span><span class="sxs-lookup"><span data-stu-id="bd8f2-179">Requirements</span></span>

|<span data-ttu-id="bd8f2-180">Требование</span><span class="sxs-lookup"><span data-stu-id="bd8f2-180">Requirement</span></span>| <span data-ttu-id="bd8f2-181">Значение</span><span class="sxs-lookup"><span data-stu-id="bd8f2-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd8f2-182">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="bd8f2-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bd8f2-183">1.0</span><span class="sxs-lookup"><span data-stu-id="bd8f2-183">1.0</span></span>|
|[<span data-ttu-id="bd8f2-184">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="bd8f2-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bd8f2-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd8f2-185">ReadItem</span></span>|
|[<span data-ttu-id="bd8f2-186">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="bd8f2-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bd8f2-187">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="bd8f2-187">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="bd8f2-188">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="bd8f2-188">Returns:</span></span>

<span data-ttu-id="bd8f2-189">Тип: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="bd8f2-189">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)</span></span>

#### <a name="converttorestiditemid-restversion--string"></a><span data-ttu-id="bd8f2-190">convertToRestId(itemId, restVersion) → {String}</span><span class="sxs-lookup"><span data-stu-id="bd8f2-190">convertToRestId(itemId, restVersion) → {String}</span></span>

<span data-ttu-id="bd8f2-191">Преобразовывает идентификатор элемента в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-191">Converts an item ID formatted for EWS into REST format.</span></span>

> [!NOTE]
> <span data-ttu-id="bd8f2-192">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-192">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="bd8f2-p107">Формат идентификаторов, извлекаемых через EWS или свойство `itemId`, отличается от формата API REST (таких как [API Почты Outlook](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) или [Microsoft Graph](https://graph.microsoft.io/)). Метод `convertToRestId` преобразовывает идентификатор в формате EWS в формат REST.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-p107">Item IDs retrieved via EWS or via the `itemId` property use a different format than the format used by REST APIs (such as the [Outlook Mail API](/previous-versions/office/office-365-api/api/version-2.0/mail-rest-operations) or the [Microsoft Graph](https://graph.microsoft.io/)). The `convertToRestId` method converts an EWS-formatted ID into the proper format for REST.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bd8f2-195">Параметры</span><span class="sxs-lookup"><span data-stu-id="bd8f2-195">Parameters</span></span>

|<span data-ttu-id="bd8f2-196">Имя</span><span class="sxs-lookup"><span data-stu-id="bd8f2-196">Name</span></span>| <span data-ttu-id="bd8f2-197">Тип</span><span class="sxs-lookup"><span data-stu-id="bd8f2-197">Type</span></span>| <span data-ttu-id="bd8f2-198">Описание</span><span class="sxs-lookup"><span data-stu-id="bd8f2-198">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="bd8f2-199">String</span><span class="sxs-lookup"><span data-stu-id="bd8f2-199">String</span></span>|<span data-ttu-id="bd8f2-200">Идентификатор элемента в формате EWS</span><span class="sxs-lookup"><span data-stu-id="bd8f2-200">An item ID formatted for Exchange Web Services (EWS)</span></span>|
|`restVersion`| [<span data-ttu-id="bd8f2-201">Office.MailboxEnums.RestVersion</span><span class="sxs-lookup"><span data-stu-id="bd8f2-201">Office.MailboxEnums.RestVersion</span></span>](/javascript/api/outlook/office.mailboxenums.restversion?view=outlook-js-1.3)|<span data-ttu-id="bd8f2-202">Значение, определяющее версию REST API для Outlook, с которой будет использоваться преобразованный идентификатор.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-202">A value indicating the version of the Outlook REST API that the converted ID will be used with.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bd8f2-203">Требования</span><span class="sxs-lookup"><span data-stu-id="bd8f2-203">Requirements</span></span>

|<span data-ttu-id="bd8f2-204">Требование</span><span class="sxs-lookup"><span data-stu-id="bd8f2-204">Requirement</span></span>| <span data-ttu-id="bd8f2-205">Значение</span><span class="sxs-lookup"><span data-stu-id="bd8f2-205">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd8f2-206">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="bd8f2-206">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bd8f2-207">1.3</span><span class="sxs-lookup"><span data-stu-id="bd8f2-207">1.3</span></span>|
|[<span data-ttu-id="bd8f2-208">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="bd8f2-208">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bd8f2-209">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="bd8f2-209">Restricted</span></span>|
|[<span data-ttu-id="bd8f2-210">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="bd8f2-210">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bd8f2-211">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="bd8f2-211">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="bd8f2-212">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="bd8f2-212">Returns:</span></span>

<span data-ttu-id="bd8f2-213">Тип: String</span><span class="sxs-lookup"><span data-stu-id="bd8f2-213">Type: String</span></span>

##### <a name="example"></a><span data-ttu-id="bd8f2-214">Пример</span><span class="sxs-lookup"><span data-stu-id="bd8f2-214">Example</span></span>

```javascript
// Get the currently selected item's ID.
var ewsId = Office.context.mailbox.item.itemId;

// Convert to a REST ID for the v2.0 version of the Outlook Mail API.
var restId = Office.context.mailbox.convertToRestId(ewsId, Office.MailboxEnums.RestVersion.v2_0);
```

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="bd8f2-215">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="bd8f2-215">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="bd8f2-216">Получает объект Date из словаря, содержащего сведения о времени.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-216">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="bd8f2-217">Метод `convertToUtcClientTime` преобразует словарь, содержащий локальную дату и время, в объект Date с правильными значениями локальной даты и времени.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-217">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bd8f2-218">Параметры</span><span class="sxs-lookup"><span data-stu-id="bd8f2-218">Parameters</span></span>

|<span data-ttu-id="bd8f2-219">Имя</span><span class="sxs-lookup"><span data-stu-id="bd8f2-219">Name</span></span>| <span data-ttu-id="bd8f2-220">Тип</span><span class="sxs-lookup"><span data-stu-id="bd8f2-220">Type</span></span>| <span data-ttu-id="bd8f2-221">Описание</span><span class="sxs-lookup"><span data-stu-id="bd8f2-221">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="bd8f2-222">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="bd8f2-222">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.3)|<span data-ttu-id="bd8f2-223">Значение локального времени для преобразования.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-223">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bd8f2-224">Требования</span><span class="sxs-lookup"><span data-stu-id="bd8f2-224">Requirements</span></span>

|<span data-ttu-id="bd8f2-225">Требование</span><span class="sxs-lookup"><span data-stu-id="bd8f2-225">Requirement</span></span>| <span data-ttu-id="bd8f2-226">Значение</span><span class="sxs-lookup"><span data-stu-id="bd8f2-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd8f2-227">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="bd8f2-227">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bd8f2-228">1.0</span><span class="sxs-lookup"><span data-stu-id="bd8f2-228">1.0</span></span>|
|[<span data-ttu-id="bd8f2-229">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="bd8f2-229">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bd8f2-230">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd8f2-230">ReadItem</span></span>|
|[<span data-ttu-id="bd8f2-231">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="bd8f2-231">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bd8f2-232">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="bd8f2-232">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="bd8f2-233">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="bd8f2-233">Returns:</span></span>

<span data-ttu-id="bd8f2-234">Объект Date со временем в формате UTC.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-234">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="bd8f2-235">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="bd8f2-235">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="bd8f2-236">Дата</span><span class="sxs-lookup"><span data-stu-id="bd8f2-236">Date</span></span></dd>

</dl>

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="bd8f2-237">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="bd8f2-237">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="bd8f2-238">Отображает имеющуюся встречу из календаря.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-238">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="bd8f2-239">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-239">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="bd8f2-240">Метод `displayAppointmentForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее сведения календаря о существующей встрече.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-240">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="bd8f2-241">В Outlook на Mac Этот метод можно использовать для отображения одной встречи, которая не является частью повторяющегося ряда, или главной встречи повторяющейся серии, но невозможно отобразить экземпляр ряда.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-241">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series.</span></span> <span data-ttu-id="bd8f2-242">Это связано с тем, что в Outlook на Mac-адресе невозможно получить доступ к свойствам (включая идентификатор элемента) повторяющихся рядов.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-242">This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="bd8f2-243">В Outlook в Интернете этот метод открывает указанную форму, только если текст формы меньше или равен 32 КБ числу символов.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-243">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="bd8f2-244">Если указанный идентификатор элемента не определяет существующую встречу, на клиентском компьютере или устройстве открывается пустая страница, и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-244">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bd8f2-245">Параметры</span><span class="sxs-lookup"><span data-stu-id="bd8f2-245">Parameters</span></span>

|<span data-ttu-id="bd8f2-246">Имя</span><span class="sxs-lookup"><span data-stu-id="bd8f2-246">Name</span></span>| <span data-ttu-id="bd8f2-247">Тип</span><span class="sxs-lookup"><span data-stu-id="bd8f2-247">Type</span></span>| <span data-ttu-id="bd8f2-248">Описание</span><span class="sxs-lookup"><span data-stu-id="bd8f2-248">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="bd8f2-249">String</span><span class="sxs-lookup"><span data-stu-id="bd8f2-249">String</span></span>|<span data-ttu-id="bd8f2-250">Идентификатор веб-служб Exchange для существующей встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-250">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bd8f2-251">Требования</span><span class="sxs-lookup"><span data-stu-id="bd8f2-251">Requirements</span></span>

|<span data-ttu-id="bd8f2-252">Требование</span><span class="sxs-lookup"><span data-stu-id="bd8f2-252">Requirement</span></span>| <span data-ttu-id="bd8f2-253">Значение</span><span class="sxs-lookup"><span data-stu-id="bd8f2-253">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd8f2-254">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="bd8f2-254">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bd8f2-255">1.0</span><span class="sxs-lookup"><span data-stu-id="bd8f2-255">1.0</span></span>|
|[<span data-ttu-id="bd8f2-256">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="bd8f2-256">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bd8f2-257">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd8f2-257">ReadItem</span></span>|
|[<span data-ttu-id="bd8f2-258">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="bd8f2-258">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bd8f2-259">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="bd8f2-259">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bd8f2-260">Пример</span><span class="sxs-lookup"><span data-stu-id="bd8f2-260">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

#### <a name="displaymessageformitemid"></a><span data-ttu-id="bd8f2-261">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="bd8f2-261">displayMessageForm(itemId)</span></span>

<span data-ttu-id="bd8f2-262">Отображает имеющееся сообщение.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-262">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="bd8f2-263">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-263">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="bd8f2-264">Метод `displayMessageForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее существующее сообщение.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-264">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="bd8f2-265">В Outlook в Интернете этот метод открывает указанную форму только в том случае, если размер текста формы меньше или равен 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-265">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="bd8f2-266">Если указанный идентификатор элемента не определяет существующее сообщение, окно на клиентском компьютере не открывается и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-266">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="bd8f2-p109">Не используйте `displayMessageForm` с параметром `itemId`, который представляет собой встречу. Используйте метод `displayAppointmentForm`, чтобы отобразить сведения о существующей встрече, а метод `displayNewAppointmentForm` — для отображения формы создания встречи.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-p109">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bd8f2-269">Параметры</span><span class="sxs-lookup"><span data-stu-id="bd8f2-269">Parameters</span></span>

|<span data-ttu-id="bd8f2-270">Имя</span><span class="sxs-lookup"><span data-stu-id="bd8f2-270">Name</span></span>| <span data-ttu-id="bd8f2-271">Тип</span><span class="sxs-lookup"><span data-stu-id="bd8f2-271">Type</span></span>| <span data-ttu-id="bd8f2-272">Описание</span><span class="sxs-lookup"><span data-stu-id="bd8f2-272">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="bd8f2-273">String</span><span class="sxs-lookup"><span data-stu-id="bd8f2-273">String</span></span>|<span data-ttu-id="bd8f2-274">Идентификатор веб-служб Exchange для существующего сообщения.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-274">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bd8f2-275">Требования</span><span class="sxs-lookup"><span data-stu-id="bd8f2-275">Requirements</span></span>

|<span data-ttu-id="bd8f2-276">Требование</span><span class="sxs-lookup"><span data-stu-id="bd8f2-276">Requirement</span></span>| <span data-ttu-id="bd8f2-277">Значение</span><span class="sxs-lookup"><span data-stu-id="bd8f2-277">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd8f2-278">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="bd8f2-278">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bd8f2-279">1.0</span><span class="sxs-lookup"><span data-stu-id="bd8f2-279">1.0</span></span>|
|[<span data-ttu-id="bd8f2-280">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="bd8f2-280">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bd8f2-281">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd8f2-281">ReadItem</span></span>|
|[<span data-ttu-id="bd8f2-282">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="bd8f2-282">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bd8f2-283">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="bd8f2-283">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bd8f2-284">Пример</span><span class="sxs-lookup"><span data-stu-id="bd8f2-284">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="bd8f2-285">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="bd8f2-285">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="bd8f2-286">Отображает форму для создания новой встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-286">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="bd8f2-287">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-287">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="bd8f2-p110">Метод `displayNewAppointmentForm` открывает форму, в которой пользователь может создать встречу или собрание. Если параметры заданы, поля формы встречи автоматически заполняются их содержимым.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-p110">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="bd8f2-290">В Outlook в Интернете и на мобильных устройствах этот метод всегда отображает форму с полем участников.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-290">In Outlook on the web and mobile devices, this method always displays a form with an attendees field.</span></span> <span data-ttu-id="bd8f2-291">Если вы не укажете участников в качестве входных аргументов, метод отображает форму с кнопкой **Сохранить**.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-291">If you do not specify any attendees as input arguments, the method displays a form with a **Save** button.</span></span> <span data-ttu-id="bd8f2-292">Если вы укажете участников, форма будет включать участников и кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-292">If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="bd8f2-p112">Если вы укажете участников или ресурсы с помощью параметра `requiredAttendees`, `optionalAttendees` или `resources` в клиенте Outlook с расширенными возможностями и Outlook RT, этот метод отобразит форму собрания с кнопкой **Отправить**. Если не указать получателей, этот метод отобразит форму встречи с кнопкой **Сохранить и закрыть**.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-p112">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="bd8f2-295">Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-295">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bd8f2-296">Параметры</span><span class="sxs-lookup"><span data-stu-id="bd8f2-296">Parameters</span></span>

|<span data-ttu-id="bd8f2-297">Имя</span><span class="sxs-lookup"><span data-stu-id="bd8f2-297">Name</span></span>| <span data-ttu-id="bd8f2-298">Тип</span><span class="sxs-lookup"><span data-stu-id="bd8f2-298">Type</span></span>| <span data-ttu-id="bd8f2-299">Описание</span><span class="sxs-lookup"><span data-stu-id="bd8f2-299">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="bd8f2-300">Object</span><span class="sxs-lookup"><span data-stu-id="bd8f2-300">Object</span></span> | <span data-ttu-id="bd8f2-301">Словарь параметров, описывающий новую встречу.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-301">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="bd8f2-302">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)&gt;</span><span class="sxs-lookup"><span data-stu-id="bd8f2-302">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)&gt;</span></span> | <span data-ttu-id="bd8f2-p113">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из обязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-p113">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="bd8f2-305">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)&gt;</span><span class="sxs-lookup"><span data-stu-id="bd8f2-305">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.3)&gt;</span></span> | <span data-ttu-id="bd8f2-p114">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из необязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-p114">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="bd8f2-308">Date</span><span class="sxs-lookup"><span data-stu-id="bd8f2-308">Date</span></span> | <span data-ttu-id="bd8f2-309">Объект `Date`, указывающий дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-309">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="bd8f2-310">Date</span><span class="sxs-lookup"><span data-stu-id="bd8f2-310">Date</span></span> | <span data-ttu-id="bd8f2-311">Объект `Date`, указывающий дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-311">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="bd8f2-312">Строка</span><span class="sxs-lookup"><span data-stu-id="bd8f2-312">String</span></span> | <span data-ttu-id="bd8f2-p115">Строка со сведениями о месте встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-p115">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="bd8f2-315">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="bd8f2-315">Array.&lt;String&gt;</span></span> | <span data-ttu-id="bd8f2-p116">Массив строк, содержащий необходимые для встречи ресурсы. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-p116">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="bd8f2-318">String</span><span class="sxs-lookup"><span data-stu-id="bd8f2-318">String</span></span> | <span data-ttu-id="bd8f2-p117">Строка с темой встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-p117">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="bd8f2-321">String</span><span class="sxs-lookup"><span data-stu-id="bd8f2-321">String</span></span> | <span data-ttu-id="bd8f2-p118">Текст сообщения о встрече. Максимальный размер содержимого сообщения — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-p118">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="bd8f2-324">Требования</span><span class="sxs-lookup"><span data-stu-id="bd8f2-324">Requirements</span></span>

|<span data-ttu-id="bd8f2-325">Требование</span><span class="sxs-lookup"><span data-stu-id="bd8f2-325">Requirement</span></span>| <span data-ttu-id="bd8f2-326">Значение</span><span class="sxs-lookup"><span data-stu-id="bd8f2-326">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd8f2-327">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="bd8f2-327">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bd8f2-328">1.0</span><span class="sxs-lookup"><span data-stu-id="bd8f2-328">1.0</span></span>|
|[<span data-ttu-id="bd8f2-329">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="bd8f2-329">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bd8f2-330">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd8f2-330">ReadItem</span></span>|
|[<span data-ttu-id="bd8f2-331">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="bd8f2-331">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bd8f2-332">Чтение</span><span class="sxs-lookup"><span data-stu-id="bd8f2-332">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bd8f2-333">Пример</span><span class="sxs-lookup"><span data-stu-id="bd8f2-333">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="bd8f2-334">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="bd8f2-334">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="bd8f2-335">Получает строку, содержащую маркер, используемый для получения вложения или элемента с Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-335">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="bd8f2-p119">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-p119">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="bd8f2-p120">Вы можете передать сторонней системе маркер и идентификатор вложения или элемента. Сторонняя система использует этот маркер как маркер авторизации, чтобы вызвать операцию [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) или [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) веб-служб Exchange для возврата вложения или элемента. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="bd8f2-p120">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="bd8f2-341">Для вызова метода `getCallbackTokenAsync` в режиме чтения манифесте приложения должно быть указано разрешение **ReadItem**.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-341">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method in read mode.</span></span>

<span data-ttu-id="bd8f2-p121">Чтобы получить идентификатор элемента для передачи в метод `getCallbackTokenAsync`, в режиме создания необходимо вызвать метод [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback). Для вызова метода `saveAsync` приложение должно иметь разрешения **ReadWriteItem**.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-p121">In compose mode you must call the [`saveAsync`](Office.context.mailbox.item.md#saveasyncoptions-callback) method to get an item identifier to pass to the `getCallbackTokenAsync` method. Your app must have **ReadWriteItem** permissions to call the `saveAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bd8f2-344">Параметры</span><span class="sxs-lookup"><span data-stu-id="bd8f2-344">Parameters</span></span>

|<span data-ttu-id="bd8f2-345">Имя</span><span class="sxs-lookup"><span data-stu-id="bd8f2-345">Name</span></span>| <span data-ttu-id="bd8f2-346">Тип</span><span class="sxs-lookup"><span data-stu-id="bd8f2-346">Type</span></span>| <span data-ttu-id="bd8f2-347">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="bd8f2-347">Attributes</span></span>| <span data-ttu-id="bd8f2-348">Описание</span><span class="sxs-lookup"><span data-stu-id="bd8f2-348">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="bd8f2-349">функция</span><span class="sxs-lookup"><span data-stu-id="bd8f2-349">function</span></span>||<span data-ttu-id="bd8f2-350">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="bd8f2-350">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="bd8f2-351">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-351">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="bd8f2-352">При возникновении ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут содержать дополнительные сведения.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-352">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="bd8f2-353">Объект</span><span class="sxs-lookup"><span data-stu-id="bd8f2-353">Object</span></span>| <span data-ttu-id="bd8f2-354">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="bd8f2-354">&lt;optional&gt;</span></span>|<span data-ttu-id="bd8f2-355">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-355">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="bd8f2-356">Ошибки</span><span class="sxs-lookup"><span data-stu-id="bd8f2-356">Errors</span></span>

|<span data-ttu-id="bd8f2-357">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="bd8f2-357">Error code</span></span>|<span data-ttu-id="bd8f2-358">Описание</span><span class="sxs-lookup"><span data-stu-id="bd8f2-358">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="bd8f2-359">Запрос не выполнен.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-359">The request has failed.</span></span> <span data-ttu-id="bd8f2-360">Просмотрите объект Diagnostics для кода ошибки HTTP.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-360">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="bd8f2-361">Сервер Exchange возвратил ошибку.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-361">The Exchange server returned an error.</span></span> <span data-ttu-id="bd8f2-362">Дополнительные сведения можно найти в объекте диагностики.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-362">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="bd8f2-363">Пользователь больше не подключен к сети.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-363">The user is no longer connected to the network.</span></span> <span data-ttu-id="bd8f2-364">Проверьте сетевое подключение и повторите попытку.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-364">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bd8f2-365">Требования</span><span class="sxs-lookup"><span data-stu-id="bd8f2-365">Requirements</span></span>

|<span data-ttu-id="bd8f2-366">Требование</span><span class="sxs-lookup"><span data-stu-id="bd8f2-366">Requirement</span></span>| <span data-ttu-id="bd8f2-367">Значение</span><span class="sxs-lookup"><span data-stu-id="bd8f2-367">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd8f2-368">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="bd8f2-368">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bd8f2-369">1.0</span><span class="sxs-lookup"><span data-stu-id="bd8f2-369">1.0</span></span>|
|[<span data-ttu-id="bd8f2-370">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="bd8f2-370">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bd8f2-371">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd8f2-371">ReadItem</span></span>|
|[<span data-ttu-id="bd8f2-372">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="bd8f2-372">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bd8f2-373">Создание и чтение</span><span class="sxs-lookup"><span data-stu-id="bd8f2-373">Compose and read</span></span>|

##### <a name="example"></a><span data-ttu-id="bd8f2-374">Пример</span><span class="sxs-lookup"><span data-stu-id="bd8f2-374">Example</span></span>

```javascript
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="bd8f2-375">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="bd8f2-375">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="bd8f2-376">Получает маркер, идентифицирующий пользователя и надстройку Office.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-376">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="bd8f2-377">Метод `getUserIdentityTokenAsync` возвращает маркер, который можно использовать для идентификации, а также [проверки подлинности надстройки и пользователя в сторонней системе](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="bd8f2-377">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="bd8f2-378">Параметры</span><span class="sxs-lookup"><span data-stu-id="bd8f2-378">Parameters</span></span>

|<span data-ttu-id="bd8f2-379">Имя</span><span class="sxs-lookup"><span data-stu-id="bd8f2-379">Name</span></span>| <span data-ttu-id="bd8f2-380">Тип</span><span class="sxs-lookup"><span data-stu-id="bd8f2-380">Type</span></span>| <span data-ttu-id="bd8f2-381">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="bd8f2-381">Attributes</span></span>| <span data-ttu-id="bd8f2-382">Описание</span><span class="sxs-lookup"><span data-stu-id="bd8f2-382">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="bd8f2-383">функция</span><span class="sxs-lookup"><span data-stu-id="bd8f2-383">function</span></span>||<span data-ttu-id="bd8f2-384">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="bd8f2-384">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="bd8f2-385">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-385">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="bd8f2-386">При возникновении ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут содержать дополнительные сведения.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-386">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="bd8f2-387">Объект</span><span class="sxs-lookup"><span data-stu-id="bd8f2-387">Object</span></span>| <span data-ttu-id="bd8f2-388">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="bd8f2-388">&lt;optional&gt;</span></span>|<span data-ttu-id="bd8f2-389">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-389">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="bd8f2-390">Ошибки</span><span class="sxs-lookup"><span data-stu-id="bd8f2-390">Errors</span></span>

|<span data-ttu-id="bd8f2-391">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="bd8f2-391">Error code</span></span>|<span data-ttu-id="bd8f2-392">Описание</span><span class="sxs-lookup"><span data-stu-id="bd8f2-392">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="bd8f2-393">Запрос не выполнен.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-393">The request has failed.</span></span> <span data-ttu-id="bd8f2-394">Просмотрите объект Diagnostics для кода ошибки HTTP.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-394">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="bd8f2-395">Сервер Exchange возвратил ошибку.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-395">The Exchange server returned an error.</span></span> <span data-ttu-id="bd8f2-396">Дополнительные сведения можно найти в объекте диагностики.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-396">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="bd8f2-397">Пользователь больше не подключен к сети.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-397">The user is no longer connected to the network.</span></span> <span data-ttu-id="bd8f2-398">Проверьте сетевое подключение и повторите попытку.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-398">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bd8f2-399">Требования</span><span class="sxs-lookup"><span data-stu-id="bd8f2-399">Requirements</span></span>

|<span data-ttu-id="bd8f2-400">Требование</span><span class="sxs-lookup"><span data-stu-id="bd8f2-400">Requirement</span></span>| <span data-ttu-id="bd8f2-401">Значение</span><span class="sxs-lookup"><span data-stu-id="bd8f2-401">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd8f2-402">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="bd8f2-402">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bd8f2-403">1.0</span><span class="sxs-lookup"><span data-stu-id="bd8f2-403">1.0</span></span>|
|[<span data-ttu-id="bd8f2-404">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="bd8f2-404">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bd8f2-405">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bd8f2-405">ReadItem</span></span>|
|[<span data-ttu-id="bd8f2-406">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="bd8f2-406">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bd8f2-407">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="bd8f2-407">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bd8f2-408">Пример</span><span class="sxs-lookup"><span data-stu-id="bd8f2-408">Example</span></span>

```javascript
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="bd8f2-409">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="bd8f2-409">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="bd8f2-410">Выполняет асинхронный запрос для веб-служб Exchange (EWS) на сервере Exchange Server, на котором размещен почтовый ящик пользователя.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-410">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="bd8f2-411">Этот метод не поддерживается в следующих сценариях:</span><span class="sxs-lookup"><span data-stu-id="bd8f2-411">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="bd8f2-412">В Outlook на iOS или Android</span><span class="sxs-lookup"><span data-stu-id="bd8f2-412">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="bd8f2-413">Если надстройка загружается в почтовый ящик Gmail.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-413">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="bd8f2-414">В таких случаях надстройка должна [использовать REST API](/outlook/add-ins/use-rest-api) для доступа к почтовому ящику пользователя.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-414">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="bd8f2-415">Метод `makeEwsRequestAsync` отправляет запрос EWS от имени надстройки в Exchange.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-415">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="bd8f2-416">Список поддерживаемых операций EWS см. в статье [Вызов веб-служб из надстройки Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support).</span><span class="sxs-lookup"><span data-stu-id="bd8f2-416">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="bd8f2-417">С помощью метода `makeEwsRequestAsync` невозможно запрашивать элементы, связанные с папкой.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-417">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="bd8f2-418">В запросе XML должна быть указана кодировка UTF-8.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-418">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="bd8f2-p129">У вашей надстройки должно быть разрешение **ReadWriteMailbox** для использования метода `makeEwsRequestAsync`. Сведения об использовании разрешения **ReadWriteMailbox** и операций EWS, которые можно вызывать с помощью метода `makeEwsRequestAsync`, см. в статье [Указание разрешений для доступа почтовой надстройки к почтовому ящику пользователя](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="bd8f2-p129">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="bd8f2-421">Администратор сервера должен установить значение true для параметра `OAuthAuthentication` в каталоге сервера клиентского доступа EWS, чтобы метод `makeEwsRequestAsync` мог выполнять запросы EWS.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-421">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="bd8f2-422">Различия версий</span><span class="sxs-lookup"><span data-stu-id="bd8f2-422">Version differences</span></span>

<span data-ttu-id="bd8f2-423">Если вы используете метод `makeEwsRequestAsync` в почтовых приложениях, которые выполняются в Outlook версии более ранней, чем 15.0.4535.1004, указывайте кодировку `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-423">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="bd8f2-p130">Значение кодировки не нужно указывать, если почтовое приложение выполняется в Outlook в Интернете. Чтобы определить, выполняется ли приложение в Outlook или Outlook в Интернете, используйте свойство mailbox.diagnostics.hostName. Используемую версию Outlook можно определить с помощью свойства mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-p130">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="bd8f2-427">Параметры</span><span class="sxs-lookup"><span data-stu-id="bd8f2-427">Parameters</span></span>

|<span data-ttu-id="bd8f2-428">Имя</span><span class="sxs-lookup"><span data-stu-id="bd8f2-428">Name</span></span>| <span data-ttu-id="bd8f2-429">Тип</span><span class="sxs-lookup"><span data-stu-id="bd8f2-429">Type</span></span>| <span data-ttu-id="bd8f2-430">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="bd8f2-430">Attributes</span></span>| <span data-ttu-id="bd8f2-431">Описание</span><span class="sxs-lookup"><span data-stu-id="bd8f2-431">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="bd8f2-432">String</span><span class="sxs-lookup"><span data-stu-id="bd8f2-432">String</span></span>||<span data-ttu-id="bd8f2-433">Запрос EWS.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-433">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="bd8f2-434">function</span><span class="sxs-lookup"><span data-stu-id="bd8f2-434">function</span></span>||<span data-ttu-id="bd8f2-435">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="bd8f2-435">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="bd8f2-436">Результат XML вызова EWS указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-436">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="bd8f2-437">Если размер результата превышает 1 МБ, возвращается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-437">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="bd8f2-438">Объект</span><span class="sxs-lookup"><span data-stu-id="bd8f2-438">Object</span></span>| <span data-ttu-id="bd8f2-439">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="bd8f2-439">&lt;optional&gt;</span></span>|<span data-ttu-id="bd8f2-440">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-440">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="bd8f2-441">Требования</span><span class="sxs-lookup"><span data-stu-id="bd8f2-441">Requirements</span></span>

|<span data-ttu-id="bd8f2-442">Требование</span><span class="sxs-lookup"><span data-stu-id="bd8f2-442">Requirement</span></span>| <span data-ttu-id="bd8f2-443">Значение</span><span class="sxs-lookup"><span data-stu-id="bd8f2-443">Value</span></span>|
|---|---|
|[<span data-ttu-id="bd8f2-444">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="bd8f2-444">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bd8f2-445">1.0</span><span class="sxs-lookup"><span data-stu-id="bd8f2-445">1.0</span></span>|
|[<span data-ttu-id="bd8f2-446">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="bd8f2-446">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bd8f2-447">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="bd8f2-447">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="bd8f2-448">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="bd8f2-448">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bd8f2-449">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="bd8f2-449">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bd8f2-450">Пример</span><span class="sxs-lookup"><span data-stu-id="bd8f2-450">Example</span></span>

<span data-ttu-id="bd8f2-451">В приведенном ниже примере вызывается `makeEwsRequestAsync` для получения темы элемента с помощью операции `GetItem`.</span><span class="sxs-lookup"><span data-stu-id="bd8f2-451">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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
