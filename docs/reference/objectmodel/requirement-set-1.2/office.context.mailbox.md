---
title: Office. Context. Mailbox — набор обязательных элементов 1,2
description: ''
ms.date: 08/30/2019
localization_priority: Normal
ms.openlocfilehash: 2002b7784d0d7295762d1f692e7a0115f1f97059
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696346"
---
# <a name="mailbox"></a><span data-ttu-id="8e6fb-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="8e6fb-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="8e6fb-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="8e6fb-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="8e6fb-104">Предоставляет доступ к объектной модели надстройки Outlook для Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e6fb-105">Требования</span><span class="sxs-lookup"><span data-stu-id="8e6fb-105">Requirements</span></span>

|<span data-ttu-id="8e6fb-106">Требование</span><span class="sxs-lookup"><span data-stu-id="8e6fb-106">Requirement</span></span>| <span data-ttu-id="8e6fb-107">Значение</span><span class="sxs-lookup"><span data-stu-id="8e6fb-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e6fb-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8e6fb-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8e6fb-109">1.0</span><span class="sxs-lookup"><span data-stu-id="8e6fb-109">1.0</span></span>|
|[<span data-ttu-id="8e6fb-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8e6fb-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8e6fb-111">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="8e6fb-111">Restricted</span></span>|
|[<span data-ttu-id="8e6fb-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8e6fb-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8e6fb-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8e6fb-113">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="8e6fb-114">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="8e6fb-114">Members and methods</span></span>

| <span data-ttu-id="8e6fb-115">Элемент</span><span class="sxs-lookup"><span data-stu-id="8e6fb-115">Member</span></span> | <span data-ttu-id="8e6fb-116">Тип</span><span class="sxs-lookup"><span data-stu-id="8e6fb-116">Type</span></span> |
|--------|------|
| [<span data-ttu-id="8e6fb-117">ewsUrl</span><span class="sxs-lookup"><span data-stu-id="8e6fb-117">ewsUrl</span></span>](#ewsurl-string) | <span data-ttu-id="8e6fb-118">Элемент</span><span class="sxs-lookup"><span data-stu-id="8e6fb-118">Member</span></span> |
| [<span data-ttu-id="8e6fb-119">convertToLocalClientTime</span><span class="sxs-lookup"><span data-stu-id="8e6fb-119">convertToLocalClientTime</span></span>](#converttolocalclienttimetimevalue--localclienttime) | <span data-ttu-id="8e6fb-120">Метод</span><span class="sxs-lookup"><span data-stu-id="8e6fb-120">Method</span></span> |
| [<span data-ttu-id="8e6fb-121">convertToUtcClientTime</span><span class="sxs-lookup"><span data-stu-id="8e6fb-121">convertToUtcClientTime</span></span>](#converttoutcclienttimeinput--date) | <span data-ttu-id="8e6fb-122">Метод</span><span class="sxs-lookup"><span data-stu-id="8e6fb-122">Method</span></span> |
| [<span data-ttu-id="8e6fb-123">displayAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="8e6fb-123">displayAppointmentForm</span></span>](#displayappointmentformitemid) | <span data-ttu-id="8e6fb-124">Метод</span><span class="sxs-lookup"><span data-stu-id="8e6fb-124">Method</span></span> |
| [<span data-ttu-id="8e6fb-125">displayMessageForm</span><span class="sxs-lookup"><span data-stu-id="8e6fb-125">displayMessageForm</span></span>](#displaymessageformitemid) | <span data-ttu-id="8e6fb-126">Метод</span><span class="sxs-lookup"><span data-stu-id="8e6fb-126">Method</span></span> |
| [<span data-ttu-id="8e6fb-127">displayNewAppointmentForm</span><span class="sxs-lookup"><span data-stu-id="8e6fb-127">displayNewAppointmentForm</span></span>](#displaynewappointmentformparameters) | <span data-ttu-id="8e6fb-128">Метод</span><span class="sxs-lookup"><span data-stu-id="8e6fb-128">Method</span></span> |
| [<span data-ttu-id="8e6fb-129">getCallbackTokenAsync</span><span class="sxs-lookup"><span data-stu-id="8e6fb-129">getCallbackTokenAsync</span></span>](#getcallbacktokenasynccallback-usercontext) | <span data-ttu-id="8e6fb-130">Метод</span><span class="sxs-lookup"><span data-stu-id="8e6fb-130">Method</span></span> |
| [<span data-ttu-id="8e6fb-131">getUserIdentityTokenAsync</span><span class="sxs-lookup"><span data-stu-id="8e6fb-131">getUserIdentityTokenAsync</span></span>](#getuseridentitytokenasynccallback-usercontext) | <span data-ttu-id="8e6fb-132">Метод</span><span class="sxs-lookup"><span data-stu-id="8e6fb-132">Method</span></span> |
| [<span data-ttu-id="8e6fb-133">makeEwsRequestAsync</span><span class="sxs-lookup"><span data-stu-id="8e6fb-133">makeEwsRequestAsync</span></span>](#makeewsrequestasyncdata-callback-usercontext) | <span data-ttu-id="8e6fb-134">Метод</span><span class="sxs-lookup"><span data-stu-id="8e6fb-134">Method</span></span> |

### <a name="namespaces"></a><span data-ttu-id="8e6fb-135">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="8e6fb-135">Namespaces</span></span>

<span data-ttu-id="8e6fb-136">[diagnostics](Office.context.mailbox.diagnostics.md). Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-136">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="8e6fb-137">[item](Office.context.mailbox.item.md). Предоставляет методы и свойства для доступа к сообщению или встрече в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-137">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="8e6fb-138">[userProfile](Office.context.mailbox.userProfile.md). Предоставляет сведения о пользователе в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-138">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="8e6fb-139">Элементы</span><span class="sxs-lookup"><span data-stu-id="8e6fb-139">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="8e6fb-140">ewsUrl: строка</span><span class="sxs-lookup"><span data-stu-id="8e6fb-140">ewsUrl: String</span></span>

<span data-ttu-id="8e6fb-141">Получает URL-адрес конечной точки веб-служб Exchange (EWS) для этой учетной записи электронной почты.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-141">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account.</span></span> <span data-ttu-id="8e6fb-142">Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-142">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="8e6fb-143">Этот элемент не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-143">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="8e6fb-p102">Удаленная служба может использовать значение `ewsUrl`, чтобы выполнять вызовы EWS для почтового ящика пользователя. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="8e6fb-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="type"></a><span data-ttu-id="8e6fb-146">Тип</span><span class="sxs-lookup"><span data-stu-id="8e6fb-146">Type</span></span>

*   <span data-ttu-id="8e6fb-147">String</span><span class="sxs-lookup"><span data-stu-id="8e6fb-147">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e6fb-148">Требования</span><span class="sxs-lookup"><span data-stu-id="8e6fb-148">Requirements</span></span>

|<span data-ttu-id="8e6fb-149">Требование</span><span class="sxs-lookup"><span data-stu-id="8e6fb-149">Requirement</span></span>| <span data-ttu-id="8e6fb-150">Значение</span><span class="sxs-lookup"><span data-stu-id="8e6fb-150">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e6fb-151">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8e6fb-151">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8e6fb-152">1.0</span><span class="sxs-lookup"><span data-stu-id="8e6fb-152">1.0</span></span>|
|[<span data-ttu-id="8e6fb-153">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8e6fb-153">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8e6fb-154">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e6fb-154">ReadItem</span></span>|
|[<span data-ttu-id="8e6fb-155">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8e6fb-155">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8e6fb-156">Чтение</span><span class="sxs-lookup"><span data-stu-id="8e6fb-156">Read</span></span>|

### <a name="methods"></a><span data-ttu-id="8e6fb-157">Методы</span><span class="sxs-lookup"><span data-stu-id="8e6fb-157">Methods</span></span>

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-12"></a><span data-ttu-id="8e6fb-158">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)}</span><span class="sxs-lookup"><span data-stu-id="8e6fb-158">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)}</span></span>

<span data-ttu-id="8e6fb-159">Получает словарь, содержащий сведения о локальном времени клиента.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-159">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="8e6fb-160">Почтовое приложение для Outlook на настольном компьютере или в Интернете может использовать разные часовые пояса для дат и времени.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-160">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times.</span></span> <span data-ttu-id="8e6fb-161">Outlook на рабочем столе использует часовой пояс клиентского компьютера; В Outlook в Интернете используется часовой пояс, установленный в центре администрирования Exchange.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-161">Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC).</span></span> <span data-ttu-id="8e6fb-162">Значения даты и времени должны обрабатываться таким образом, чтобы значения, отображаемые в интерфейсе пользователя, всегда согласовывались с часовым поясом, ожидаемым пользователем.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-162">You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="8e6fb-163">Если почтовое приложение запущено в Outlook на настольном клиенте `convertToLocalClientTime` , метод возвратит объект Dictionary со значениями, заданными для часового пояса клиентского компьютера.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-163">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone.</span></span> <span data-ttu-id="8e6fb-164">Если почтовое приложение запущено в Outlook в Интернете, `convertToLocalClientTime` метод возвратит объект Dictionary со значениями, заданными в часовом поясе, заданном в центре администрирования Exchange.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-164">If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e6fb-165">Параметры</span><span class="sxs-lookup"><span data-stu-id="8e6fb-165">Parameters</span></span>

|<span data-ttu-id="8e6fb-166">Имя</span><span class="sxs-lookup"><span data-stu-id="8e6fb-166">Name</span></span>| <span data-ttu-id="8e6fb-167">Тип</span><span class="sxs-lookup"><span data-stu-id="8e6fb-167">Type</span></span>| <span data-ttu-id="8e6fb-168">Описание</span><span class="sxs-lookup"><span data-stu-id="8e6fb-168">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="8e6fb-169">Date</span><span class="sxs-lookup"><span data-stu-id="8e6fb-169">Date</span></span>|<span data-ttu-id="8e6fb-170">Объект Date</span><span class="sxs-lookup"><span data-stu-id="8e6fb-170">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e6fb-171">Требования</span><span class="sxs-lookup"><span data-stu-id="8e6fb-171">Requirements</span></span>

|<span data-ttu-id="8e6fb-172">Требование</span><span class="sxs-lookup"><span data-stu-id="8e6fb-172">Requirement</span></span>| <span data-ttu-id="8e6fb-173">Значение</span><span class="sxs-lookup"><span data-stu-id="8e6fb-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e6fb-174">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8e6fb-174">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8e6fb-175">1.0</span><span class="sxs-lookup"><span data-stu-id="8e6fb-175">1.0</span></span>|
|[<span data-ttu-id="8e6fb-176">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8e6fb-176">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8e6fb-177">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e6fb-177">ReadItem</span></span>|
|[<span data-ttu-id="8e6fb-178">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8e6fb-178">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8e6fb-179">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8e6fb-179">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8e6fb-180">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="8e6fb-180">Returns:</span></span>

<span data-ttu-id="8e6fb-181">Тип: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="8e6fb-181">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)</span></span>

<br>

---
---

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="8e6fb-182">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="8e6fb-182">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="8e6fb-183">Получает объект Date из словаря, содержащего сведения о времени.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-183">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="8e6fb-184">Метод `convertToUtcClientTime` преобразует словарь, содержащий локальную дату и время, в объект Date с правильными значениями локальной даты и времени.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-184">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e6fb-185">Параметры</span><span class="sxs-lookup"><span data-stu-id="8e6fb-185">Parameters</span></span>

|<span data-ttu-id="8e6fb-186">Имя</span><span class="sxs-lookup"><span data-stu-id="8e6fb-186">Name</span></span>| <span data-ttu-id="8e6fb-187">Тип</span><span class="sxs-lookup"><span data-stu-id="8e6fb-187">Type</span></span>| <span data-ttu-id="8e6fb-188">Описание</span><span class="sxs-lookup"><span data-stu-id="8e6fb-188">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="8e6fb-189">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="8e6fb-189">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)|<span data-ttu-id="8e6fb-190">Значение локального времени для преобразования.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-190">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e6fb-191">Требования</span><span class="sxs-lookup"><span data-stu-id="8e6fb-191">Requirements</span></span>

|<span data-ttu-id="8e6fb-192">Требование</span><span class="sxs-lookup"><span data-stu-id="8e6fb-192">Requirement</span></span>| <span data-ttu-id="8e6fb-193">Значение</span><span class="sxs-lookup"><span data-stu-id="8e6fb-193">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e6fb-194">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8e6fb-194">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8e6fb-195">1.0</span><span class="sxs-lookup"><span data-stu-id="8e6fb-195">1.0</span></span>|
|[<span data-ttu-id="8e6fb-196">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8e6fb-196">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8e6fb-197">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e6fb-197">ReadItem</span></span>|
|[<span data-ttu-id="8e6fb-198">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8e6fb-198">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8e6fb-199">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8e6fb-199">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="8e6fb-200">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="8e6fb-200">Returns:</span></span>

<span data-ttu-id="8e6fb-201">Объект Date со временем в формате UTC.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-201">A Date object with the time expressed in UTC.</span></span>

<span data-ttu-id="8e6fb-202">Тип: Date</span><span class="sxs-lookup"><span data-stu-id="8e6fb-202">Type: Date</span></span>

##### <a name="example"></a><span data-ttu-id="8e6fb-203">Пример</span><span class="sxs-lookup"><span data-stu-id="8e6fb-203">Example</span></span>

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

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="8e6fb-204">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="8e6fb-204">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="8e6fb-205">Отображает имеющуюся встречу из календаря.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-205">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="8e6fb-206">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-206">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="8e6fb-207">Метод `displayAppointmentForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее сведения календаря о существующей встрече.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-207">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="8e6fb-208">В Outlook на Mac Этот метод можно использовать для отображения одной встречи, которая не является частью повторяющегося ряда, или главной встречи повторяющейся серии, но невозможно отобразить экземпляр ряда.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-208">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series.</span></span> <span data-ttu-id="8e6fb-209">Это связано с тем, что в Outlook на Mac-адресе невозможно получить доступ к свойствам (включая идентификатор элемента) повторяющихся рядов.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-209">This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="8e6fb-210">В Outlook в Интернете этот метод открывает указанную форму, только если текст формы меньше или равен 32 КБ числу символов.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-210">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="8e6fb-211">Если указанный идентификатор элемента не определяет существующую встречу, на клиентском компьютере или устройстве открывается пустая страница, и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-211">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e6fb-212">Параметры</span><span class="sxs-lookup"><span data-stu-id="8e6fb-212">Parameters</span></span>

|<span data-ttu-id="8e6fb-213">Имя</span><span class="sxs-lookup"><span data-stu-id="8e6fb-213">Name</span></span>| <span data-ttu-id="8e6fb-214">Тип</span><span class="sxs-lookup"><span data-stu-id="8e6fb-214">Type</span></span>| <span data-ttu-id="8e6fb-215">Описание</span><span class="sxs-lookup"><span data-stu-id="8e6fb-215">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="8e6fb-216">String</span><span class="sxs-lookup"><span data-stu-id="8e6fb-216">String</span></span>|<span data-ttu-id="8e6fb-217">Идентификатор веб-служб Exchange для существующей встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-217">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e6fb-218">Требования</span><span class="sxs-lookup"><span data-stu-id="8e6fb-218">Requirements</span></span>

|<span data-ttu-id="8e6fb-219">Требование</span><span class="sxs-lookup"><span data-stu-id="8e6fb-219">Requirement</span></span>| <span data-ttu-id="8e6fb-220">Значение</span><span class="sxs-lookup"><span data-stu-id="8e6fb-220">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e6fb-221">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8e6fb-221">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8e6fb-222">1.0</span><span class="sxs-lookup"><span data-stu-id="8e6fb-222">1.0</span></span>|
|[<span data-ttu-id="8e6fb-223">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8e6fb-223">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8e6fb-224">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e6fb-224">ReadItem</span></span>|
|[<span data-ttu-id="8e6fb-225">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8e6fb-225">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8e6fb-226">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8e6fb-226">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e6fb-227">Пример</span><span class="sxs-lookup"><span data-stu-id="8e6fb-227">Example</span></span>

```js
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

<br>

---
---

#### <a name="displaymessageformitemid"></a><span data-ttu-id="8e6fb-228">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="8e6fb-228">displayMessageForm(itemId)</span></span>

<span data-ttu-id="8e6fb-229">Отображает имеющееся сообщение.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-229">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="8e6fb-230">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-230">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="8e6fb-231">Метод `displayMessageForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее существующее сообщение.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-231">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="8e6fb-232">В Outlook в Интернете этот метод открывает указанную форму только в том случае, если размер текста формы меньше или равен 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-232">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="8e6fb-233">Если указанный идентификатор элемента не определяет существующее сообщение, окно на клиентском компьютере не открывается и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-233">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="8e6fb-p106">Не используйте `displayMessageForm` с параметром `itemId`, который представляет собой встречу. Используйте метод `displayAppointmentForm`, чтобы отобразить сведения о существующей встрече, а метод `displayNewAppointmentForm` — для отображения формы создания встречи.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-p106">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e6fb-236">Параметры</span><span class="sxs-lookup"><span data-stu-id="8e6fb-236">Parameters</span></span>

|<span data-ttu-id="8e6fb-237">Имя</span><span class="sxs-lookup"><span data-stu-id="8e6fb-237">Name</span></span>| <span data-ttu-id="8e6fb-238">Тип</span><span class="sxs-lookup"><span data-stu-id="8e6fb-238">Type</span></span>| <span data-ttu-id="8e6fb-239">Описание</span><span class="sxs-lookup"><span data-stu-id="8e6fb-239">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="8e6fb-240">String</span><span class="sxs-lookup"><span data-stu-id="8e6fb-240">String</span></span>|<span data-ttu-id="8e6fb-241">Идентификатор веб-служб Exchange для существующего сообщения.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-241">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e6fb-242">Требования</span><span class="sxs-lookup"><span data-stu-id="8e6fb-242">Requirements</span></span>

|<span data-ttu-id="8e6fb-243">Требование</span><span class="sxs-lookup"><span data-stu-id="8e6fb-243">Requirement</span></span>| <span data-ttu-id="8e6fb-244">Значение</span><span class="sxs-lookup"><span data-stu-id="8e6fb-244">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e6fb-245">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8e6fb-245">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8e6fb-246">1.0</span><span class="sxs-lookup"><span data-stu-id="8e6fb-246">1.0</span></span>|
|[<span data-ttu-id="8e6fb-247">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8e6fb-247">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8e6fb-248">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e6fb-248">ReadItem</span></span>|
|[<span data-ttu-id="8e6fb-249">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8e6fb-249">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8e6fb-250">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8e6fb-250">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e6fb-251">Пример</span><span class="sxs-lookup"><span data-stu-id="8e6fb-251">Example</span></span>

```js
Office.context.mailbox.displayMessageForm(messageId);
```

<br>

---
---

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="8e6fb-252">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="8e6fb-252">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="8e6fb-253">Отображает форму для создания новой встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-253">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="8e6fb-254">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-254">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="8e6fb-p107">Метод `displayNewAppointmentForm` открывает форму, в которой пользователь может создать встречу или собрание. Если параметры заданы, поля формы встречи автоматически заполняются их содержимым.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-p107">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="8e6fb-257">В Outlook в Интернете и на мобильных устройствах этот метод всегда отображает форму с полем участников.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-257">In Outlook on the web and mobile devices, this method always displays a form with an attendees field.</span></span> <span data-ttu-id="8e6fb-258">Если вы не укажете участников в качестве входных аргументов, метод отображает форму с кнопкой **Сохранить**.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-258">If you do not specify any attendees as input arguments, the method displays a form with a **Save** button.</span></span> <span data-ttu-id="8e6fb-259">Если вы укажете участников, форма будет включать участников и кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-259">If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="8e6fb-p109">Если вы укажете участников или ресурсы с помощью параметра `requiredAttendees`, `optionalAttendees` или `resources` в клиенте Outlook с расширенными возможностями и Outlook RT, этот метод отобразит форму собрания с кнопкой **Отправить**. Если не указать получателей, этот метод отобразит форму встречи с кнопкой **Сохранить и закрыть**.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-p109">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="8e6fb-262">Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-262">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e6fb-263">Параметры</span><span class="sxs-lookup"><span data-stu-id="8e6fb-263">Parameters</span></span>

|<span data-ttu-id="8e6fb-264">Имя</span><span class="sxs-lookup"><span data-stu-id="8e6fb-264">Name</span></span>| <span data-ttu-id="8e6fb-265">Тип</span><span class="sxs-lookup"><span data-stu-id="8e6fb-265">Type</span></span>| <span data-ttu-id="8e6fb-266">Описание</span><span class="sxs-lookup"><span data-stu-id="8e6fb-266">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="8e6fb-267">Object</span><span class="sxs-lookup"><span data-stu-id="8e6fb-267">Object</span></span> | <span data-ttu-id="8e6fb-268">Словарь параметров, описывающий новую встречу.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-268">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="8e6fb-269">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)&gt;</span><span class="sxs-lookup"><span data-stu-id="8e6fb-269">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)&gt;</span></span> | <span data-ttu-id="8e6fb-p110">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из обязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-p110">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="8e6fb-272">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)&gt;</span><span class="sxs-lookup"><span data-stu-id="8e6fb-272">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)&gt;</span></span> | <span data-ttu-id="8e6fb-p111">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из необязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-p111">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="8e6fb-275">Date</span><span class="sxs-lookup"><span data-stu-id="8e6fb-275">Date</span></span> | <span data-ttu-id="8e6fb-276">Объект `Date`, указывающий дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-276">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="8e6fb-277">Date</span><span class="sxs-lookup"><span data-stu-id="8e6fb-277">Date</span></span> | <span data-ttu-id="8e6fb-278">Объект `Date`, указывающий дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-278">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="8e6fb-279">Строка</span><span class="sxs-lookup"><span data-stu-id="8e6fb-279">String</span></span> | <span data-ttu-id="8e6fb-p112">Строка со сведениями о месте встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-p112">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="8e6fb-282">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="8e6fb-282">Array.&lt;String&gt;</span></span> | <span data-ttu-id="8e6fb-p113">Массив строк, содержащий необходимые для встречи ресурсы. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-p113">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="8e6fb-285">String</span><span class="sxs-lookup"><span data-stu-id="8e6fb-285">String</span></span> | <span data-ttu-id="8e6fb-p114">Строка с темой встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-p114">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="8e6fb-288">String</span><span class="sxs-lookup"><span data-stu-id="8e6fb-288">String</span></span> | <span data-ttu-id="8e6fb-p115">Текст сообщения о встрече. Максимальный размер содержимого сообщения — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-p115">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8e6fb-291">Требования</span><span class="sxs-lookup"><span data-stu-id="8e6fb-291">Requirements</span></span>

|<span data-ttu-id="8e6fb-292">Требование</span><span class="sxs-lookup"><span data-stu-id="8e6fb-292">Requirement</span></span>| <span data-ttu-id="8e6fb-293">Значение</span><span class="sxs-lookup"><span data-stu-id="8e6fb-293">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e6fb-294">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8e6fb-294">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8e6fb-295">1.0</span><span class="sxs-lookup"><span data-stu-id="8e6fb-295">1.0</span></span>|
|[<span data-ttu-id="8e6fb-296">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8e6fb-296">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8e6fb-297">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e6fb-297">ReadItem</span></span>|
|[<span data-ttu-id="8e6fb-298">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8e6fb-298">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8e6fb-299">Чтение</span><span class="sxs-lookup"><span data-stu-id="8e6fb-299">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e6fb-300">Пример</span><span class="sxs-lookup"><span data-stu-id="8e6fb-300">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="8e6fb-301">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="8e6fb-301">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="8e6fb-302">Получает строку, содержащую маркер, используемый для получения вложения или элемента с Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-302">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="8e6fb-p116">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-p116">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="8e6fb-p117">Вы можете передать сторонней системе маркер и идентификатор вложения или элемента. Сторонняя система использует этот маркер как маркер авторизации, чтобы вызвать операцию [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) или [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) веб-служб Exchange для возврата вложения или элемента. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="8e6fb-p117">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="8e6fb-308">Чтобы вызвать метод \*\*\*\*, у вашего приложения должно быть разрешение `getCallbackTokenAsync`, указанное в его манифесте.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-308">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e6fb-309">Параметры</span><span class="sxs-lookup"><span data-stu-id="8e6fb-309">Parameters</span></span>

|<span data-ttu-id="8e6fb-310">Имя</span><span class="sxs-lookup"><span data-stu-id="8e6fb-310">Name</span></span>| <span data-ttu-id="8e6fb-311">Тип</span><span class="sxs-lookup"><span data-stu-id="8e6fb-311">Type</span></span>| <span data-ttu-id="8e6fb-312">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="8e6fb-312">Attributes</span></span>| <span data-ttu-id="8e6fb-313">Описание</span><span class="sxs-lookup"><span data-stu-id="8e6fb-313">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="8e6fb-314">function</span><span class="sxs-lookup"><span data-stu-id="8e6fb-314">function</span></span>||<span data-ttu-id="8e6fb-315">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8e6fb-315">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8e6fb-316">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-316">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="8e6fb-317">При возникновении ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут содержать дополнительные сведения.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-317">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="8e6fb-318">Объект</span><span class="sxs-lookup"><span data-stu-id="8e6fb-318">Object</span></span>| <span data-ttu-id="8e6fb-319">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8e6fb-319">&lt;optional&gt;</span></span>|<span data-ttu-id="8e6fb-320">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-320">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8e6fb-321">Ошибки</span><span class="sxs-lookup"><span data-stu-id="8e6fb-321">Errors</span></span>

|<span data-ttu-id="8e6fb-322">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="8e6fb-322">Error code</span></span>|<span data-ttu-id="8e6fb-323">Описание</span><span class="sxs-lookup"><span data-stu-id="8e6fb-323">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="8e6fb-324">Запрос не выполнен.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-324">The request has failed.</span></span> <span data-ttu-id="8e6fb-325">Просмотрите объект Diagnostics для кода ошибки HTTP.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-325">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="8e6fb-326">Сервер Exchange возвратил ошибку.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-326">The Exchange server returned an error.</span></span> <span data-ttu-id="8e6fb-327">Дополнительные сведения можно найти в объекте диагностики.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-327">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="8e6fb-328">Пользователь больше не подключен к сети.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-328">The user is no longer connected to the network.</span></span> <span data-ttu-id="8e6fb-329">Проверьте сетевое подключение и повторите попытку.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-329">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e6fb-330">Требования</span><span class="sxs-lookup"><span data-stu-id="8e6fb-330">Requirements</span></span>

|<span data-ttu-id="8e6fb-331">Требование</span><span class="sxs-lookup"><span data-stu-id="8e6fb-331">Requirement</span></span>| <span data-ttu-id="8e6fb-332">Значение</span><span class="sxs-lookup"><span data-stu-id="8e6fb-332">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e6fb-333">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8e6fb-333">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8e6fb-334">1.0</span><span class="sxs-lookup"><span data-stu-id="8e6fb-334">1.0</span></span>|
|[<span data-ttu-id="8e6fb-335">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8e6fb-335">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8e6fb-336">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e6fb-336">ReadItem</span></span>|
|[<span data-ttu-id="8e6fb-337">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8e6fb-337">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8e6fb-338">Чтение</span><span class="sxs-lookup"><span data-stu-id="8e6fb-338">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e6fb-339">Пример</span><span class="sxs-lookup"><span data-stu-id="8e6fb-339">Example</span></span>

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

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="8e6fb-340">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="8e6fb-340">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="8e6fb-341">Получает маркер, идентифицирующий пользователя и надстройку Office.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-341">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="8e6fb-342">Метод `getUserIdentityTokenAsync` возвращает маркер, который можно использовать для идентификации, а также [проверки подлинности надстройки и пользователя в сторонней системе](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="8e6fb-342">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e6fb-343">Параметры</span><span class="sxs-lookup"><span data-stu-id="8e6fb-343">Parameters</span></span>

|<span data-ttu-id="8e6fb-344">Имя</span><span class="sxs-lookup"><span data-stu-id="8e6fb-344">Name</span></span>| <span data-ttu-id="8e6fb-345">Тип</span><span class="sxs-lookup"><span data-stu-id="8e6fb-345">Type</span></span>| <span data-ttu-id="8e6fb-346">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="8e6fb-346">Attributes</span></span>| <span data-ttu-id="8e6fb-347">Описание</span><span class="sxs-lookup"><span data-stu-id="8e6fb-347">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="8e6fb-348">function</span><span class="sxs-lookup"><span data-stu-id="8e6fb-348">function</span></span>||<span data-ttu-id="8e6fb-349">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8e6fb-349">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8e6fb-350">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-350">The token is provided as a string in the `asyncResult.value` property.</span></span><br><br><span data-ttu-id="8e6fb-351">При возникновении ошибки свойства `asyncResult.error` и `asyncResult.diagnostics` могут содержать дополнительные сведения.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-351">If there was an error, the `asyncResult.error` and `asyncResult.diagnostics` properties may provide additional information.</span></span>|
|`userContext`| <span data-ttu-id="8e6fb-352">Объект</span><span class="sxs-lookup"><span data-stu-id="8e6fb-352">Object</span></span>| <span data-ttu-id="8e6fb-353">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8e6fb-353">&lt;optional&gt;</span></span>|<span data-ttu-id="8e6fb-354">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-354">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="errors"></a><span data-ttu-id="8e6fb-355">Ошибки</span><span class="sxs-lookup"><span data-stu-id="8e6fb-355">Errors</span></span>

|<span data-ttu-id="8e6fb-356">Код ошибки</span><span class="sxs-lookup"><span data-stu-id="8e6fb-356">Error code</span></span>|<span data-ttu-id="8e6fb-357">Описание</span><span class="sxs-lookup"><span data-stu-id="8e6fb-357">Description</span></span>|
|------------|-------------|
|`HTTPRequestFailure`|<span data-ttu-id="8e6fb-358">Запрос не выполнен.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-358">The request has failed.</span></span> <span data-ttu-id="8e6fb-359">Просмотрите объект Diagnostics для кода ошибки HTTP.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-359">Please look at the diagnostics object for the HTTP error code.</span></span>|
|`InternalServerError`|<span data-ttu-id="8e6fb-360">Сервер Exchange возвратил ошибку.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-360">The Exchange server returned an error.</span></span> <span data-ttu-id="8e6fb-361">Дополнительные сведения можно найти в объекте диагностики.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-361">Please look at the diagnostics object for more information.</span></span>|
|`NetworkError`|<span data-ttu-id="8e6fb-362">Пользователь больше не подключен к сети.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-362">The user is no longer connected to the network.</span></span> <span data-ttu-id="8e6fb-363">Проверьте сетевое подключение и повторите попытку.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-363">Please check your network connection and try again.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e6fb-364">Требования</span><span class="sxs-lookup"><span data-stu-id="8e6fb-364">Requirements</span></span>

|<span data-ttu-id="8e6fb-365">Требование</span><span class="sxs-lookup"><span data-stu-id="8e6fb-365">Requirement</span></span>| <span data-ttu-id="8e6fb-366">Значение</span><span class="sxs-lookup"><span data-stu-id="8e6fb-366">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e6fb-367">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8e6fb-367">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8e6fb-368">1.0</span><span class="sxs-lookup"><span data-stu-id="8e6fb-368">1.0</span></span>|
|[<span data-ttu-id="8e6fb-369">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8e6fb-369">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8e6fb-370">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8e6fb-370">ReadItem</span></span>|
|[<span data-ttu-id="8e6fb-371">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8e6fb-371">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8e6fb-372">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8e6fb-372">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e6fb-373">Пример</span><span class="sxs-lookup"><span data-stu-id="8e6fb-373">Example</span></span>

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

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="8e6fb-374">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="8e6fb-374">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="8e6fb-375">Выполняет асинхронный запрос для веб-служб Exchange (EWS) на сервере Exchange Server, на котором размещен почтовый ящик пользователя.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-375">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="8e6fb-376">Этот метод не поддерживается в следующих сценариях:</span><span class="sxs-lookup"><span data-stu-id="8e6fb-376">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="8e6fb-377">В Outlook на iOS или Android</span><span class="sxs-lookup"><span data-stu-id="8e6fb-377">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="8e6fb-378">Если надстройка загружается в почтовый ящик Gmail.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-378">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="8e6fb-379">В таких случаях надстройка должна [использовать REST API](/outlook/add-ins/use-rest-api) для доступа к почтовому ящику пользователя.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-379">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="8e6fb-380">Метод `makeEwsRequestAsync` отправляет запрос EWS от имени надстройки в Exchange.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-380">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="8e6fb-381">Список поддерживаемых операций EWS см. в статье [Вызов веб-служб из надстройки Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support).</span><span class="sxs-lookup"><span data-stu-id="8e6fb-381">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="8e6fb-382">С помощью метода `makeEwsRequestAsync` невозможно запрашивать элементы, связанные с папкой.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-382">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="8e6fb-383">В запросе XML должна быть указана кодировка UTF-8.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-383">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="8e6fb-p125">У вашей надстройки должно быть разрешение **ReadWriteMailbox** для использования метода `makeEwsRequestAsync`. Сведения об использовании разрешения **ReadWriteMailbox** и операций EWS, которые можно вызывать с помощью метода `makeEwsRequestAsync`, см. в статье [Указание разрешений для доступа почтовой надстройки к почтовому ящику пользователя](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="8e6fb-p125">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="8e6fb-386">Администратор сервера должен установить значение true для параметра `OAuthAuthentication` в каталоге сервера клиентского доступа EWS, чтобы метод `makeEwsRequestAsync` мог выполнять запросы EWS.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-386">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="8e6fb-387">Различия версий</span><span class="sxs-lookup"><span data-stu-id="8e6fb-387">Version differences</span></span>

<span data-ttu-id="8e6fb-388">Если вы используете метод `makeEwsRequestAsync` в почтовых приложениях, которые выполняются в Outlook версии более ранней, чем 15.0.4535.1004, указывайте кодировку `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-388">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="8e6fb-p126">Значение кодировки не нужно указывать, если почтовое приложение выполняется в Outlook в Интернете. Чтобы определить, выполняется ли приложение в Outlook или Outlook в Интернете, используйте свойство mailbox.diagnostics.hostName. Используемую версию Outlook можно определить с помощью свойства mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-p126">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="8e6fb-392">Параметры</span><span class="sxs-lookup"><span data-stu-id="8e6fb-392">Parameters</span></span>

|<span data-ttu-id="8e6fb-393">Имя</span><span class="sxs-lookup"><span data-stu-id="8e6fb-393">Name</span></span>| <span data-ttu-id="8e6fb-394">Тип</span><span class="sxs-lookup"><span data-stu-id="8e6fb-394">Type</span></span>| <span data-ttu-id="8e6fb-395">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="8e6fb-395">Attributes</span></span>| <span data-ttu-id="8e6fb-396">Описание</span><span class="sxs-lookup"><span data-stu-id="8e6fb-396">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="8e6fb-397">String</span><span class="sxs-lookup"><span data-stu-id="8e6fb-397">String</span></span>||<span data-ttu-id="8e6fb-398">Запрос EWS.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-398">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="8e6fb-399">function</span><span class="sxs-lookup"><span data-stu-id="8e6fb-399">function</span></span>||<span data-ttu-id="8e6fb-400">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="8e6fb-400">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="8e6fb-401">Результат XML вызова EWS указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-401">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="8e6fb-402">Если размер результата превышает 1 МБ, возвращается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-402">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="8e6fb-403">Объект</span><span class="sxs-lookup"><span data-stu-id="8e6fb-403">Object</span></span>| <span data-ttu-id="8e6fb-404">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="8e6fb-404">&lt;optional&gt;</span></span>|<span data-ttu-id="8e6fb-405">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-405">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e6fb-406">Требования</span><span class="sxs-lookup"><span data-stu-id="8e6fb-406">Requirements</span></span>

|<span data-ttu-id="8e6fb-407">Требование</span><span class="sxs-lookup"><span data-stu-id="8e6fb-407">Requirement</span></span>| <span data-ttu-id="8e6fb-408">Значение</span><span class="sxs-lookup"><span data-stu-id="8e6fb-408">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e6fb-409">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8e6fb-409">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8e6fb-410">1.0</span><span class="sxs-lookup"><span data-stu-id="8e6fb-410">1.0</span></span>|
|[<span data-ttu-id="8e6fb-411">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8e6fb-411">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8e6fb-412">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="8e6fb-412">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="8e6fb-413">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8e6fb-413">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8e6fb-414">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8e6fb-414">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8e6fb-415">Пример</span><span class="sxs-lookup"><span data-stu-id="8e6fb-415">Example</span></span>

<span data-ttu-id="8e6fb-416">В приведенном ниже примере вызывается `makeEwsRequestAsync` для получения темы элемента с помощью операции `GetItem`.</span><span class="sxs-lookup"><span data-stu-id="8e6fb-416">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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
