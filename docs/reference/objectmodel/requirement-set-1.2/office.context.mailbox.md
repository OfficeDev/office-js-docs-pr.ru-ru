---
title: Office. Context. Mailbox — набор обязательных элементов 1,2
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: de7c48faf966f9b3f5d1bb76f69aa16810a5381f
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064370"
---
# <a name="mailbox"></a><span data-ttu-id="19040-102">mailbox</span><span class="sxs-lookup"><span data-stu-id="19040-102">mailbox</span></span>

### <a name="officeofficemdcontextofficecontextmdmailbox"></a><span data-ttu-id="19040-103">[Office](Office.md)[.context](Office.context.md).mailbox</span><span class="sxs-lookup"><span data-stu-id="19040-103">[Office](Office.md)[.context](Office.context.md).mailbox</span></span>

<span data-ttu-id="19040-104">Предоставляет доступ к объектной модели надстройки Outlook для Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="19040-104">Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

##### <a name="requirements"></a><span data-ttu-id="19040-105">Требования</span><span class="sxs-lookup"><span data-stu-id="19040-105">Requirements</span></span>

|<span data-ttu-id="19040-106">Требование</span><span class="sxs-lookup"><span data-stu-id="19040-106">Requirement</span></span>| <span data-ttu-id="19040-107">Значение</span><span class="sxs-lookup"><span data-stu-id="19040-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="19040-108">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="19040-108">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="19040-109">1.0</span><span class="sxs-lookup"><span data-stu-id="19040-109">1.0</span></span>|
|[<span data-ttu-id="19040-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="19040-110">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="19040-111">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="19040-111">Restricted</span></span>|
|[<span data-ttu-id="19040-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="19040-112">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="19040-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="19040-113">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="19040-114">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="19040-114">Namespaces</span></span>

<span data-ttu-id="19040-115">[diagnostics](Office.context.mailbox.diagnostics.md). Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="19040-115">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="19040-116">[item](Office.context.mailbox.item.md). Предоставляет методы и свойства для доступа к сообщению или встрече в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="19040-116">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="19040-117">[userProfile](Office.context.mailbox.userProfile.md). Предоставляет сведения о пользователе в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="19040-117">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="19040-118">Элементы</span><span class="sxs-lookup"><span data-stu-id="19040-118">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="19040-119">ewsUrl: строка</span><span class="sxs-lookup"><span data-stu-id="19040-119">ewsUrl: String</span></span>

<span data-ttu-id="19040-120">Получает URL-адрес конечной точки веб-служб Exchange (EWS) для этой учетной записи электронной почты.</span><span class="sxs-lookup"><span data-stu-id="19040-120">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account.</span></span> <span data-ttu-id="19040-121">Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="19040-121">Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="19040-122">Этот элемент не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="19040-122">This member is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="19040-p102">Удаленная служба может использовать значение `ewsUrl`, чтобы выполнять вызовы EWS для почтового ящика пользователя. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="19040-p102">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="type"></a><span data-ttu-id="19040-125">Тип</span><span class="sxs-lookup"><span data-stu-id="19040-125">Type</span></span>

*   <span data-ttu-id="19040-126">String</span><span class="sxs-lookup"><span data-stu-id="19040-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="19040-127">Требования</span><span class="sxs-lookup"><span data-stu-id="19040-127">Requirements</span></span>

|<span data-ttu-id="19040-128">Требование</span><span class="sxs-lookup"><span data-stu-id="19040-128">Requirement</span></span>| <span data-ttu-id="19040-129">Значение</span><span class="sxs-lookup"><span data-stu-id="19040-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="19040-130">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="19040-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="19040-131">1.0</span><span class="sxs-lookup"><span data-stu-id="19040-131">1.0</span></span>|
|[<span data-ttu-id="19040-132">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="19040-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="19040-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="19040-133">ReadItem</span></span>|
|[<span data-ttu-id="19040-134">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="19040-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="19040-135">Чтение</span><span class="sxs-lookup"><span data-stu-id="19040-135">Read</span></span>|

### <a name="methods"></a><span data-ttu-id="19040-136">Методы</span><span class="sxs-lookup"><span data-stu-id="19040-136">Methods</span></span>

#### <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlookofficelocalclienttimeviewoutlook-js-12"></a><span data-ttu-id="19040-137">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)}</span><span class="sxs-lookup"><span data-stu-id="19040-137">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)}</span></span>

<span data-ttu-id="19040-138">Получает словарь, содержащий сведения о локальном времени клиента.</span><span class="sxs-lookup"><span data-stu-id="19040-138">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="19040-139">Почтовое приложение для Outlook на настольном компьютере или в Интернете может использовать разные часовые пояса для дат и времени.</span><span class="sxs-lookup"><span data-stu-id="19040-139">A mail app for Outlook on a desktop or on the web can use different time zones for the dates and times.</span></span> <span data-ttu-id="19040-140">Outlook на рабочем столе использует часовой пояс клиентского компьютера; В Outlook в Интернете используется часовой пояс, установленный в центре администрирования Exchange.</span><span class="sxs-lookup"><span data-stu-id="19040-140">Outlook on a desktop uses the client computer time zone; Outlook on the web uses the time zone set on the Exchange Admin Center (EAC).</span></span> <span data-ttu-id="19040-141">Значения даты и времени должны обрабатываться таким образом, чтобы значения, отображаемые в интерфейсе пользователя, всегда согласовывались с часовым поясом, ожидаемым пользователем.</span><span class="sxs-lookup"><span data-stu-id="19040-141">You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="19040-142">Если почтовое приложение запущено в Outlook на настольном клиенте `convertToLocalClientTime` , метод возвратит объект Dictionary со значениями, заданными для часового пояса клиентского компьютера.</span><span class="sxs-lookup"><span data-stu-id="19040-142">If the mail app is running in Outlook on a desktop client, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone.</span></span> <span data-ttu-id="19040-143">Если почтовое приложение запущено в Outlook в Интернете, `convertToLocalClientTime` метод возвратит объект Dictionary со значениями, заданными в часовом поясе, заданном в центре администрирования Exchange.</span><span class="sxs-lookup"><span data-stu-id="19040-143">If the mail app is running in Outlook on the web, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="19040-144">Параметры</span><span class="sxs-lookup"><span data-stu-id="19040-144">Parameters</span></span>

|<span data-ttu-id="19040-145">Имя</span><span class="sxs-lookup"><span data-stu-id="19040-145">Name</span></span>| <span data-ttu-id="19040-146">Тип</span><span class="sxs-lookup"><span data-stu-id="19040-146">Type</span></span>| <span data-ttu-id="19040-147">Описание</span><span class="sxs-lookup"><span data-stu-id="19040-147">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="19040-148">Дата</span><span class="sxs-lookup"><span data-stu-id="19040-148">Date</span></span>|<span data-ttu-id="19040-149">Объект Date</span><span class="sxs-lookup"><span data-stu-id="19040-149">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="19040-150">Требования</span><span class="sxs-lookup"><span data-stu-id="19040-150">Requirements</span></span>

|<span data-ttu-id="19040-151">Требование</span><span class="sxs-lookup"><span data-stu-id="19040-151">Requirement</span></span>| <span data-ttu-id="19040-152">Значение</span><span class="sxs-lookup"><span data-stu-id="19040-152">Value</span></span>|
|---|---|
|[<span data-ttu-id="19040-153">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="19040-153">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="19040-154">1.0</span><span class="sxs-lookup"><span data-stu-id="19040-154">1.0</span></span>|
|[<span data-ttu-id="19040-155">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="19040-155">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="19040-156">ReadItem</span><span class="sxs-lookup"><span data-stu-id="19040-156">ReadItem</span></span>|
|[<span data-ttu-id="19040-157">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="19040-157">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="19040-158">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="19040-158">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="19040-159">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="19040-159">Returns:</span></span>

<span data-ttu-id="19040-160">Тип: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="19040-160">Type: [LocalClientTime](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)</span></span>

#### <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="19040-161">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="19040-161">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="19040-162">Получает объект Date из словаря, содержащего сведения о времени.</span><span class="sxs-lookup"><span data-stu-id="19040-162">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="19040-163">Метод `convertToUtcClientTime` преобразует словарь, содержащий локальную дату и время, в объект Date с правильными значениями локальной даты и времени.</span><span class="sxs-lookup"><span data-stu-id="19040-163">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="19040-164">Параметры</span><span class="sxs-lookup"><span data-stu-id="19040-164">Parameters</span></span>

|<span data-ttu-id="19040-165">Имя</span><span class="sxs-lookup"><span data-stu-id="19040-165">Name</span></span>| <span data-ttu-id="19040-166">Тип</span><span class="sxs-lookup"><span data-stu-id="19040-166">Type</span></span>| <span data-ttu-id="19040-167">Описание</span><span class="sxs-lookup"><span data-stu-id="19040-167">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="19040-168">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="19040-168">LocalClientTime</span></span>](/javascript/api/outlook/office.LocalClientTime?view=outlook-js-1.2)|<span data-ttu-id="19040-169">Значение локального времени для преобразования.</span><span class="sxs-lookup"><span data-stu-id="19040-169">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="19040-170">Требования</span><span class="sxs-lookup"><span data-stu-id="19040-170">Requirements</span></span>

|<span data-ttu-id="19040-171">Требование</span><span class="sxs-lookup"><span data-stu-id="19040-171">Requirement</span></span>| <span data-ttu-id="19040-172">Значение</span><span class="sxs-lookup"><span data-stu-id="19040-172">Value</span></span>|
|---|---|
|[<span data-ttu-id="19040-173">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="19040-173">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="19040-174">1.0</span><span class="sxs-lookup"><span data-stu-id="19040-174">1.0</span></span>|
|[<span data-ttu-id="19040-175">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="19040-175">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="19040-176">ReadItem</span><span class="sxs-lookup"><span data-stu-id="19040-176">ReadItem</span></span>|
|[<span data-ttu-id="19040-177">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="19040-177">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="19040-178">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="19040-178">Compose or Read</span></span>|

##### <a name="returns"></a><span data-ttu-id="19040-179">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="19040-179">Returns:</span></span>

<span data-ttu-id="19040-180">Объект Date со временем в формате UTC.</span><span class="sxs-lookup"><span data-stu-id="19040-180">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="19040-181">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="19040-181">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="19040-182">Date</span><span class="sxs-lookup"><span data-stu-id="19040-182">Date</span></span></dd>

</dl>

#### <a name="displayappointmentformitemid"></a><span data-ttu-id="19040-183">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="19040-183">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="19040-184">Отображает имеющуюся встречу из календаря.</span><span class="sxs-lookup"><span data-stu-id="19040-184">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="19040-185">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="19040-185">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="19040-186">Метод `displayAppointmentForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее сведения календаря о существующей встрече.</span><span class="sxs-lookup"><span data-stu-id="19040-186">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="19040-187">В Outlook на Mac Этот метод можно использовать для отображения одной встречи, которая не является частью повторяющегося ряда, или главной встречи повторяющейся серии, но невозможно отобразить экземпляр ряда.</span><span class="sxs-lookup"><span data-stu-id="19040-187">In Outlook on Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series.</span></span> <span data-ttu-id="19040-188">Это связано с тем, что в Outlook на Mac-адресе невозможно получить доступ к свойствам (включая идентификатор элемента) повторяющихся рядов.</span><span class="sxs-lookup"><span data-stu-id="19040-188">This is because in Outlook on Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="19040-189">В Outlook в Интернете этот метод открывает указанную форму, только если текст формы меньше или равен 32 КБ числу символов.</span><span class="sxs-lookup"><span data-stu-id="19040-189">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="19040-190">Если указанный идентификатор элемента не определяет существующую встречу, на клиентском компьютере или устройстве открывается пустая страница, и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="19040-190">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="19040-191">Параметры</span><span class="sxs-lookup"><span data-stu-id="19040-191">Parameters</span></span>

|<span data-ttu-id="19040-192">Имя</span><span class="sxs-lookup"><span data-stu-id="19040-192">Name</span></span>| <span data-ttu-id="19040-193">Тип</span><span class="sxs-lookup"><span data-stu-id="19040-193">Type</span></span>| <span data-ttu-id="19040-194">Описание</span><span class="sxs-lookup"><span data-stu-id="19040-194">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="19040-195">String</span><span class="sxs-lookup"><span data-stu-id="19040-195">String</span></span>|<span data-ttu-id="19040-196">Идентификатор веб-служб Exchange для существующей встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="19040-196">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="19040-197">Требования</span><span class="sxs-lookup"><span data-stu-id="19040-197">Requirements</span></span>

|<span data-ttu-id="19040-198">Требование</span><span class="sxs-lookup"><span data-stu-id="19040-198">Requirement</span></span>| <span data-ttu-id="19040-199">Значение</span><span class="sxs-lookup"><span data-stu-id="19040-199">Value</span></span>|
|---|---|
|[<span data-ttu-id="19040-200">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="19040-200">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="19040-201">1.0</span><span class="sxs-lookup"><span data-stu-id="19040-201">1.0</span></span>|
|[<span data-ttu-id="19040-202">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="19040-202">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="19040-203">ReadItem</span><span class="sxs-lookup"><span data-stu-id="19040-203">ReadItem</span></span>|
|[<span data-ttu-id="19040-204">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="19040-204">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="19040-205">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="19040-205">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="19040-206">Пример</span><span class="sxs-lookup"><span data-stu-id="19040-206">Example</span></span>

```javascript
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

#### <a name="displaymessageformitemid"></a><span data-ttu-id="19040-207">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="19040-207">displayMessageForm(itemId)</span></span>

<span data-ttu-id="19040-208">Отображает имеющееся сообщение.</span><span class="sxs-lookup"><span data-stu-id="19040-208">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="19040-209">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="19040-209">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="19040-210">Метод `displayMessageForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее существующее сообщение.</span><span class="sxs-lookup"><span data-stu-id="19040-210">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="19040-211">В Outlook в Интернете этот метод открывает указанную форму только в том случае, если размер текста формы меньше или равен 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="19040-211">In Outlook on the web, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="19040-212">Если указанный идентификатор элемента не определяет существующее сообщение, окно на клиентском компьютере не открывается и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="19040-212">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="19040-p106">Не используйте `displayMessageForm` с параметром `itemId`, который представляет собой встречу. Используйте метод `displayAppointmentForm`, чтобы отобразить сведения о существующей встрече, а метод `displayNewAppointmentForm` — для отображения формы создания встречи.</span><span class="sxs-lookup"><span data-stu-id="19040-p106">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="19040-215">Параметры</span><span class="sxs-lookup"><span data-stu-id="19040-215">Parameters</span></span>

|<span data-ttu-id="19040-216">Имя</span><span class="sxs-lookup"><span data-stu-id="19040-216">Name</span></span>| <span data-ttu-id="19040-217">Тип</span><span class="sxs-lookup"><span data-stu-id="19040-217">Type</span></span>| <span data-ttu-id="19040-218">Описание</span><span class="sxs-lookup"><span data-stu-id="19040-218">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="19040-219">String</span><span class="sxs-lookup"><span data-stu-id="19040-219">String</span></span>|<span data-ttu-id="19040-220">Идентификатор веб-служб Exchange для существующего сообщения.</span><span class="sxs-lookup"><span data-stu-id="19040-220">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="19040-221">Требования</span><span class="sxs-lookup"><span data-stu-id="19040-221">Requirements</span></span>

|<span data-ttu-id="19040-222">Требование</span><span class="sxs-lookup"><span data-stu-id="19040-222">Requirement</span></span>| <span data-ttu-id="19040-223">Значение</span><span class="sxs-lookup"><span data-stu-id="19040-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="19040-224">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="19040-224">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="19040-225">1.0</span><span class="sxs-lookup"><span data-stu-id="19040-225">1.0</span></span>|
|[<span data-ttu-id="19040-226">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="19040-226">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="19040-227">ReadItem</span><span class="sxs-lookup"><span data-stu-id="19040-227">ReadItem</span></span>|
|[<span data-ttu-id="19040-228">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="19040-228">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="19040-229">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="19040-229">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="19040-230">Пример</span><span class="sxs-lookup"><span data-stu-id="19040-230">Example</span></span>

```javascript
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="19040-231">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="19040-231">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="19040-232">Отображает форму для создания новой встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="19040-232">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="19040-233">Этот метод не поддерживается в Outlook на iOS или Android.</span><span class="sxs-lookup"><span data-stu-id="19040-233">This method is not supported in Outlook on iOS or Android.</span></span>

<span data-ttu-id="19040-p107">Метод `displayNewAppointmentForm` открывает форму, в которой пользователь может создать встречу или собрание. Если параметры заданы, поля формы встречи автоматически заполняются их содержимым.</span><span class="sxs-lookup"><span data-stu-id="19040-p107">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="19040-236">В Outlook в Интернете и на мобильных устройствах этот метод всегда отображает форму с полем участников.</span><span class="sxs-lookup"><span data-stu-id="19040-236">In Outlook on the web and mobile devices, this method always displays a form with an attendees field.</span></span> <span data-ttu-id="19040-237">Если вы не укажете участников в качестве входных аргументов, метод отображает форму с кнопкой **Сохранить**.</span><span class="sxs-lookup"><span data-stu-id="19040-237">If you do not specify any attendees as input arguments, the method displays a form with a **Save** button.</span></span> <span data-ttu-id="19040-238">Если вы укажете участников, форма будет включать участников и кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="19040-238">If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="19040-p109">Если вы укажете участников или ресурсы с помощью параметра `requiredAttendees`, `optionalAttendees` или `resources` в клиенте Outlook с расширенными возможностями и Outlook RT, этот метод отобразит форму собрания с кнопкой **Отправить**. Если не указать получателей, этот метод отобразит форму встречи с кнопкой **Сохранить и закрыть**.</span><span class="sxs-lookup"><span data-stu-id="19040-p109">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="19040-241">Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.</span><span class="sxs-lookup"><span data-stu-id="19040-241">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="19040-242">Параметры</span><span class="sxs-lookup"><span data-stu-id="19040-242">Parameters</span></span>

|<span data-ttu-id="19040-243">Имя</span><span class="sxs-lookup"><span data-stu-id="19040-243">Name</span></span>| <span data-ttu-id="19040-244">Тип</span><span class="sxs-lookup"><span data-stu-id="19040-244">Type</span></span>| <span data-ttu-id="19040-245">Описание</span><span class="sxs-lookup"><span data-stu-id="19040-245">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="19040-246">Object</span><span class="sxs-lookup"><span data-stu-id="19040-246">Object</span></span> | <span data-ttu-id="19040-247">Словарь параметров, описывающий новую встречу.</span><span class="sxs-lookup"><span data-stu-id="19040-247">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="19040-248">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)&gt;</span><span class="sxs-lookup"><span data-stu-id="19040-248">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)&gt;</span></span> | <span data-ttu-id="19040-p110">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из обязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="19040-p110">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="19040-251">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)&gt;</span><span class="sxs-lookup"><span data-stu-id="19040-251">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook/office.emailaddressdetails?view=outlook-js-1.2)&gt;</span></span> | <span data-ttu-id="19040-p111">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из необязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="19040-p111">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="19040-254">Date</span><span class="sxs-lookup"><span data-stu-id="19040-254">Date</span></span> | <span data-ttu-id="19040-255">Объект `Date`, указывающий дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="19040-255">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="19040-256">Date</span><span class="sxs-lookup"><span data-stu-id="19040-256">Date</span></span> | <span data-ttu-id="19040-257">Объект `Date`, указывающий дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="19040-257">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="19040-258">Строка</span><span class="sxs-lookup"><span data-stu-id="19040-258">String</span></span> | <span data-ttu-id="19040-p112">Строка со сведениями о месте встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="19040-p112">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="19040-261">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="19040-261">Array.&lt;String&gt;</span></span> | <span data-ttu-id="19040-p113">Массив строк, содержащий необходимые для встречи ресурсы. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="19040-p113">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="19040-264">String</span><span class="sxs-lookup"><span data-stu-id="19040-264">String</span></span> | <span data-ttu-id="19040-p114">Строка с темой встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="19040-p114">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="19040-267">String</span><span class="sxs-lookup"><span data-stu-id="19040-267">String</span></span> | <span data-ttu-id="19040-p115">Текст сообщения о встрече. Максимальный размер содержимого сообщения — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="19040-p115">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="19040-270">Требования</span><span class="sxs-lookup"><span data-stu-id="19040-270">Requirements</span></span>

|<span data-ttu-id="19040-271">Требование</span><span class="sxs-lookup"><span data-stu-id="19040-271">Requirement</span></span>| <span data-ttu-id="19040-272">Значение</span><span class="sxs-lookup"><span data-stu-id="19040-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="19040-273">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="19040-273">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="19040-274">1.0</span><span class="sxs-lookup"><span data-stu-id="19040-274">1.0</span></span>|
|[<span data-ttu-id="19040-275">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="19040-275">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="19040-276">ReadItem</span><span class="sxs-lookup"><span data-stu-id="19040-276">ReadItem</span></span>|
|[<span data-ttu-id="19040-277">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="19040-277">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="19040-278">Чтение</span><span class="sxs-lookup"><span data-stu-id="19040-278">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="19040-279">Пример</span><span class="sxs-lookup"><span data-stu-id="19040-279">Example</span></span>

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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="19040-280">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="19040-280">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="19040-281">Получает строку, содержащую маркер, используемый для получения вложения или элемента с Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="19040-281">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="19040-p116">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="19040-p116">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="19040-p117">Вы можете передать сторонней системе маркер и идентификатор вложения или элемента. Сторонняя система использует этот маркер как маркер авторизации, чтобы вызвать операцию [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) или [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) веб-служб Exchange для возврата вложения или элемента. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="19040-p117">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="19040-287">Чтобы вызвать метод \*\*\*\*, у вашего приложения должно быть разрешение `getCallbackTokenAsync`, указанное в его манифесте.</span><span class="sxs-lookup"><span data-stu-id="19040-287">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="19040-288">Параметры</span><span class="sxs-lookup"><span data-stu-id="19040-288">Parameters</span></span>

|<span data-ttu-id="19040-289">Имя</span><span class="sxs-lookup"><span data-stu-id="19040-289">Name</span></span>| <span data-ttu-id="19040-290">Тип</span><span class="sxs-lookup"><span data-stu-id="19040-290">Type</span></span>| <span data-ttu-id="19040-291">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="19040-291">Attributes</span></span>| <span data-ttu-id="19040-292">Описание</span><span class="sxs-lookup"><span data-stu-id="19040-292">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="19040-293">function</span><span class="sxs-lookup"><span data-stu-id="19040-293">function</span></span>||<span data-ttu-id="19040-294">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="19040-294">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="19040-295">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="19040-295">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="19040-296">Объект</span><span class="sxs-lookup"><span data-stu-id="19040-296">Object</span></span>| <span data-ttu-id="19040-297">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="19040-297">&lt;optional&gt;</span></span>|<span data-ttu-id="19040-298">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="19040-298">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="19040-299">Требования</span><span class="sxs-lookup"><span data-stu-id="19040-299">Requirements</span></span>

|<span data-ttu-id="19040-300">Требование</span><span class="sxs-lookup"><span data-stu-id="19040-300">Requirement</span></span>| <span data-ttu-id="19040-301">Значение</span><span class="sxs-lookup"><span data-stu-id="19040-301">Value</span></span>|
|---|---|
|[<span data-ttu-id="19040-302">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="19040-302">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="19040-303">1.0</span><span class="sxs-lookup"><span data-stu-id="19040-303">1.0</span></span>|
|[<span data-ttu-id="19040-304">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="19040-304">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="19040-305">ReadItem</span><span class="sxs-lookup"><span data-stu-id="19040-305">ReadItem</span></span>|
|[<span data-ttu-id="19040-306">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="19040-306">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="19040-307">Чтение</span><span class="sxs-lookup"><span data-stu-id="19040-307">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="19040-308">Пример</span><span class="sxs-lookup"><span data-stu-id="19040-308">Example</span></span>

```javascript
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="19040-309">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="19040-309">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="19040-310">Получает маркер, идентифицирующий пользователя и надстройку Office.</span><span class="sxs-lookup"><span data-stu-id="19040-310">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="19040-311">Метод `getUserIdentityTokenAsync` возвращает маркер, который можно использовать для идентификации, а также [проверки подлинности надстройки и пользователя в сторонней системе](/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="19040-311">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="19040-312">Параметры</span><span class="sxs-lookup"><span data-stu-id="19040-312">Parameters</span></span>

|<span data-ttu-id="19040-313">Имя</span><span class="sxs-lookup"><span data-stu-id="19040-313">Name</span></span>| <span data-ttu-id="19040-314">Тип</span><span class="sxs-lookup"><span data-stu-id="19040-314">Type</span></span>| <span data-ttu-id="19040-315">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="19040-315">Attributes</span></span>| <span data-ttu-id="19040-316">Описание</span><span class="sxs-lookup"><span data-stu-id="19040-316">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="19040-317">function</span><span class="sxs-lookup"><span data-stu-id="19040-317">function</span></span>||<span data-ttu-id="19040-318">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="19040-318">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="19040-319">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="19040-319">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="19040-320">Объект</span><span class="sxs-lookup"><span data-stu-id="19040-320">Object</span></span>| <span data-ttu-id="19040-321">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="19040-321">&lt;optional&gt;</span></span>|<span data-ttu-id="19040-322">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="19040-322">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="19040-323">Требования</span><span class="sxs-lookup"><span data-stu-id="19040-323">Requirements</span></span>

|<span data-ttu-id="19040-324">Требование</span><span class="sxs-lookup"><span data-stu-id="19040-324">Requirement</span></span>| <span data-ttu-id="19040-325">Значение</span><span class="sxs-lookup"><span data-stu-id="19040-325">Value</span></span>|
|---|---|
|[<span data-ttu-id="19040-326">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="19040-326">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="19040-327">1.0</span><span class="sxs-lookup"><span data-stu-id="19040-327">1.0</span></span>|
|[<span data-ttu-id="19040-328">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="19040-328">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="19040-329">ReadItem</span><span class="sxs-lookup"><span data-stu-id="19040-329">ReadItem</span></span>|
|[<span data-ttu-id="19040-330">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="19040-330">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="19040-331">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="19040-331">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="19040-332">Пример</span><span class="sxs-lookup"><span data-stu-id="19040-332">Example</span></span>

```javascript
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

#### <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="19040-333">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="19040-333">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="19040-334">Выполняет асинхронный запрос для веб-служб Exchange (EWS) на сервере Exchange Server, на котором размещен почтовый ящик пользователя.</span><span class="sxs-lookup"><span data-stu-id="19040-334">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="19040-335">Этот метод не поддерживается в следующих сценариях:</span><span class="sxs-lookup"><span data-stu-id="19040-335">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="19040-336">В Outlook на iOS или Android</span><span class="sxs-lookup"><span data-stu-id="19040-336">In Outlook on iOS or Android</span></span>
> - <span data-ttu-id="19040-337">Если надстройка загружается в почтовый ящик Gmail.</span><span class="sxs-lookup"><span data-stu-id="19040-337">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="19040-338">В таких случаях надстройка должна [использовать REST API](/outlook/add-ins/use-rest-api) для доступа к почтовому ящику пользователя.</span><span class="sxs-lookup"><span data-stu-id="19040-338">In these cases, add-ins should [use REST APIs](/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="19040-339">Метод `makeEwsRequestAsync` отправляет запрос EWS от имени надстройки в Exchange.</span><span class="sxs-lookup"><span data-stu-id="19040-339">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="19040-340">Список поддерживаемых операций EWS см. в статье [Вызов веб-служб из надстройки Outlook](/outlook/add-ins/web-services#ews-operations-that-add-ins-support).</span><span class="sxs-lookup"><span data-stu-id="19040-340">See [Call web services from an Outlook add-in](/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="19040-341">С помощью метода `makeEwsRequestAsync` невозможно запрашивать элементы, связанные с папкой.</span><span class="sxs-lookup"><span data-stu-id="19040-341">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="19040-342">В запросе XML должна быть указана кодировка UTF-8.</span><span class="sxs-lookup"><span data-stu-id="19040-342">The XML request must specify UTF-8 encoding.</span></span>

```xml
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="19040-p119">У вашей надстройки должно быть разрешение **ReadWriteMailbox** для использования метода `makeEwsRequestAsync`. Сведения об использовании разрешения **ReadWriteMailbox** и операций EWS, которые можно вызывать с помощью метода `makeEwsRequestAsync`, см. в статье [Указание разрешений для доступа почтовой надстройки к почтовому ящику пользователя](/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="19040-p119">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="19040-345">Администратор сервера должен установить значение true для параметра `OAuthAuthentication` в каталоге сервера клиентского доступа EWS, чтобы метод `makeEwsRequestAsync` мог выполнять запросы EWS.</span><span class="sxs-lookup"><span data-stu-id="19040-345">The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="19040-346">Различия версий</span><span class="sxs-lookup"><span data-stu-id="19040-346">Version differences</span></span>

<span data-ttu-id="19040-347">Если вы используете метод `makeEwsRequestAsync` в почтовых приложениях, которые выполняются в Outlook версии более ранней, чем 15.0.4535.1004, указывайте кодировку `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="19040-347">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```xml
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="19040-p120">Значение кодировки не нужно указывать, если почтовое приложение выполняется в Outlook в Интернете. Чтобы определить, выполняется ли приложение в Outlook или Outlook в Интернете, используйте свойство mailbox.diagnostics.hostName. Используемую версию Outlook можно определить с помощью свойства mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="19040-p120">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="19040-351">Параметры</span><span class="sxs-lookup"><span data-stu-id="19040-351">Parameters</span></span>

|<span data-ttu-id="19040-352">Имя</span><span class="sxs-lookup"><span data-stu-id="19040-352">Name</span></span>| <span data-ttu-id="19040-353">Тип</span><span class="sxs-lookup"><span data-stu-id="19040-353">Type</span></span>| <span data-ttu-id="19040-354">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="19040-354">Attributes</span></span>| <span data-ttu-id="19040-355">Описание</span><span class="sxs-lookup"><span data-stu-id="19040-355">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="19040-356">String</span><span class="sxs-lookup"><span data-stu-id="19040-356">String</span></span>||<span data-ttu-id="19040-357">Запрос EWS.</span><span class="sxs-lookup"><span data-stu-id="19040-357">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="19040-358">function</span><span class="sxs-lookup"><span data-stu-id="19040-358">function</span></span>||<span data-ttu-id="19040-359">После выполнения метода функция, переданная в параметре `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="19040-359">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="19040-360">Результат XML вызова EWS указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="19040-360">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="19040-361">Если размер результата превышает 1 МБ, возвращается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="19040-361">If the result exceeds 1 MB in size, an error message is returned instead.</span></span>|
|`userContext`| <span data-ttu-id="19040-362">Объект</span><span class="sxs-lookup"><span data-stu-id="19040-362">Object</span></span>| <span data-ttu-id="19040-363">&lt;необязательно&gt;</span><span class="sxs-lookup"><span data-stu-id="19040-363">&lt;optional&gt;</span></span>|<span data-ttu-id="19040-364">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="19040-364">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="19040-365">Требования</span><span class="sxs-lookup"><span data-stu-id="19040-365">Requirements</span></span>

|<span data-ttu-id="19040-366">Требование</span><span class="sxs-lookup"><span data-stu-id="19040-366">Requirement</span></span>| <span data-ttu-id="19040-367">Значение</span><span class="sxs-lookup"><span data-stu-id="19040-367">Value</span></span>|
|---|---|
|[<span data-ttu-id="19040-368">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="19040-368">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="19040-369">1.0</span><span class="sxs-lookup"><span data-stu-id="19040-369">1.0</span></span>|
|[<span data-ttu-id="19040-370">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="19040-370">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="19040-371">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="19040-371">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="19040-372">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="19040-372">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="19040-373">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="19040-373">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="19040-374">Пример</span><span class="sxs-lookup"><span data-stu-id="19040-374">Example</span></span>

<span data-ttu-id="19040-375">В приведенном ниже примере вызывается `makeEwsRequestAsync` для получения темы элемента с помощью операции `GetItem`.</span><span class="sxs-lookup"><span data-stu-id="19040-375">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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
