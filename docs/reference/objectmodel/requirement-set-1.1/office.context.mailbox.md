
# <a name="mailbox"></a><span data-ttu-id="4c04c-101">mailbox</span><span class="sxs-lookup"><span data-stu-id="4c04c-101">mailbox</span></span>

### <span data-ttu-id="4c04c-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span><span class="sxs-lookup"><span data-stu-id="4c04c-p101">[Office](Office.md)[.context](Office.context.md). mailbox</span></span>

<span data-ttu-id="4c04c-104">Предоставляет для Microsoft Outlook и Microsoft Outlook в Интернете доступ к объектной модели надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="4c04c-104">Provides access to the Outlook Add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

##### <a name="requirements"></a><span data-ttu-id="4c04c-105">Требования</span><span class="sxs-lookup"><span data-stu-id="4c04c-105">Requirements</span></span>

|<span data-ttu-id="4c04c-106">Требование</span><span class="sxs-lookup"><span data-stu-id="4c04c-106">Requirement</span></span>| <span data-ttu-id="4c04c-107">Значение</span><span class="sxs-lookup"><span data-stu-id="4c04c-107">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c04c-108">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="4c04c-108">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c04c-109">1.0</span><span class="sxs-lookup"><span data-stu-id="4c04c-109">1.0</span></span>|
|[<span data-ttu-id="4c04c-110">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4c04c-110">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c04c-111">Restricted (с ограничениями)</span><span class="sxs-lookup"><span data-stu-id="4c04c-111">Restricted</span></span>|
|[<span data-ttu-id="4c04c-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4c04c-112">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c04c-113">Compose или read</span><span class="sxs-lookup"><span data-stu-id="4c04c-113">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="4c04c-114">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="4c04c-114">Namespaces</span></span>

<span data-ttu-id="4c04c-115">[diagnostics](Office.context.mailbox.diagnostics.md). Предоставляет надстройке Outlook диагностические сведения.</span><span class="sxs-lookup"><span data-stu-id="4c04c-115">[diagnostics](Office.context.mailbox.diagnostics.md): Provides diagnostic information to an Outlook add-in.</span></span>

<span data-ttu-id="4c04c-116">[item](Office.context.mailbox.item.md). Предоставляет методы и свойства для доступа к сообщению или встрече в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="4c04c-116">[item](Office.context.mailbox.item.md): Provides methods and properties for accessing a message or appointment in an Outlook add-in.</span></span>

<span data-ttu-id="4c04c-117">[userProfile](Office.context.mailbox.userProfile.md). Предоставляет сведения о пользователе в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="4c04c-117">[userProfile](Office.context.mailbox.userProfile.md): Provides information about the user in an Outlook add-in.</span></span>

### <a name="members"></a><span data-ttu-id="4c04c-118">Элементы</span><span class="sxs-lookup"><span data-stu-id="4c04c-118">Members</span></span>

#### <a name="ewsurl-string"></a><span data-ttu-id="4c04c-119">ewsUrl: String</span><span class="sxs-lookup"><span data-stu-id="4c04c-119">ewsUrl :String</span></span>

<span data-ttu-id="4c04c-p102">Получает URL-адрес конечной точки веб-служб Exchange (EWS) для конкретной учетной записи электронной почты. Только в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="4c04c-p102">Gets the URL of the Exchange Web Services (EWS) endpoint for this email account. Read mode only.</span></span>

> [!NOTE]
> <span data-ttu-id="4c04c-122">Этот элемент не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="4c04c-122">Note: This member is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4c04c-p103">Удаленная служба может использовать значение `ewsUrl`, чтобы выполнять вызовы EWS для почтового ящика пользователя. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="4c04c-p103">The `ewsUrl` value can be used by a remote service to make EWS calls to the user's mailbox. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

##### <a name="type"></a><span data-ttu-id="4c04c-125">Тип:</span><span class="sxs-lookup"><span data-stu-id="4c04c-125">Type:</span></span>

*   <span data-ttu-id="4c04c-126">String</span><span class="sxs-lookup"><span data-stu-id="4c04c-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4c04c-127">Требования</span><span class="sxs-lookup"><span data-stu-id="4c04c-127">Requirements</span></span>

|<span data-ttu-id="4c04c-128">Требование</span><span class="sxs-lookup"><span data-stu-id="4c04c-128">Requirement</span></span>| <span data-ttu-id="4c04c-129">Значение</span><span class="sxs-lookup"><span data-stu-id="4c04c-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c04c-130">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="4c04c-130">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c04c-131">1.0</span><span class="sxs-lookup"><span data-stu-id="4c04c-131">1.0</span></span>|
|[<span data-ttu-id="4c04c-132">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4c04c-132">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c04c-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4c04c-133">ReadItem</span></span>|
|[<span data-ttu-id="4c04c-134">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4c04c-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c04c-135">Read (чтение)</span><span class="sxs-lookup"><span data-stu-id="4c04c-135">Read</span></span>|

### <a name="methods"></a><span data-ttu-id="4c04c-136">Методы</span><span class="sxs-lookup"><span data-stu-id="4c04c-136">Methods</span></span>

####  <a name="converttolocalclienttimetimevalue--localclienttimejavascriptapioutlook11officelocalclienttime"></a><span data-ttu-id="4c04c-137">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_1/office.LocalClientTime)}</span><span class="sxs-lookup"><span data-stu-id="4c04c-137">convertToLocalClientTime(timeValue) → {[LocalClientTime](/javascript/api/outlook_1_1/office.LocalClientTime)}</span></span>

<span data-ttu-id="4c04c-138">Получает словарь, содержащий сведения о локальном времени клиента.</span><span class="sxs-lookup"><span data-stu-id="4c04c-138">Gets a dictionary containing time information in local client time.</span></span>

<span data-ttu-id="4c04c-p104">Для даты и времени в почтовом приложении для Outlook или Outlook Web App могут использоваться разные часовые пояса. Outlook использует часовой пояс клиентского компьютера. Outlook Web App использует часовой пояс, заданный в Центре администрирования Exchange (EAC). Значения даты и времени должны обрабатываться так, чтобы значения в пользовательском интерфейсе всегда согласовывались с часовым поясом, ожидаемым пользователем.</span><span class="sxs-lookup"><span data-stu-id="4c04c-p104">The dates and times used by a mail app for Outlook or Outlook Web App can use different time zones. Outlook uses the client computer time zone; Outlook Web App uses the time zone set on the Exchange Admin Center (EAC). You should handle date and time values so that the values you display on the user interface are always consistent with the time zone that the user expects.</span></span>

<span data-ttu-id="4c04c-p105">Если почтовое приложение работает в Outlook, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса клиентского компьютера. Если почтовое приложение работает в Outlook Web App, метод `convertToLocalClientTime` вернет объект словаря со значениями часового пояса, заданного в Центре администрирования Exchange.</span><span class="sxs-lookup"><span data-stu-id="4c04c-p105">If the mail app is running in Outlook, the `convertToLocalClientTime` method will return a dictionary object with the values set to the client computer time zone. If the mail app is running in Outlook Web App, the `convertToLocalClientTime` method will return a dictionary object with the values set to the time zone specified in the EAC.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4c04c-144">Параметры:</span><span class="sxs-lookup"><span data-stu-id="4c04c-144">Parameters:</span></span>

|<span data-ttu-id="4c04c-145">Имя</span><span class="sxs-lookup"><span data-stu-id="4c04c-145">Name</span></span>| <span data-ttu-id="4c04c-146">Тип</span><span class="sxs-lookup"><span data-stu-id="4c04c-146">Type</span></span>| <span data-ttu-id="4c04c-147">Описание</span><span class="sxs-lookup"><span data-stu-id="4c04c-147">Description</span></span>|
|---|---|---|
|`timeValue`| <span data-ttu-id="4c04c-148">Date</span><span class="sxs-lookup"><span data-stu-id="4c04c-148">Date</span></span>|<span data-ttu-id="4c04c-149">Объект Date</span><span class="sxs-lookup"><span data-stu-id="4c04c-149">A Date object</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4c04c-150">Требования</span><span class="sxs-lookup"><span data-stu-id="4c04c-150">Requirements</span></span>

|<span data-ttu-id="4c04c-151">Требование</span><span class="sxs-lookup"><span data-stu-id="4c04c-151">Requirement</span></span>| <span data-ttu-id="4c04c-152">Значение</span><span class="sxs-lookup"><span data-stu-id="4c04c-152">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c04c-153">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="4c04c-153">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c04c-154">1.0</span><span class="sxs-lookup"><span data-stu-id="4c04c-154">1.0</span></span>|
|[<span data-ttu-id="4c04c-155">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4c04c-155">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c04c-156">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4c04c-156">ReadItem</span></span>|
|[<span data-ttu-id="4c04c-157">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4c04c-157">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c04c-158">Compose или read</span><span class="sxs-lookup"><span data-stu-id="4c04c-158">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4c04c-159">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="4c04c-159">Returns:</span></span>

<span data-ttu-id="4c04c-160">Тип: [LocalClientTime](/javascript/api/outlook_1_1/office.LocalClientTime)</span><span class="sxs-lookup"><span data-stu-id="4c04c-160">Type: [LocalClientTime](/javascript/api/outlook_1_1/office.LocalClientTime)</span></span>

####  <a name="converttoutcclienttimeinput--date"></a><span data-ttu-id="4c04c-161">convertToUtcClientTime(input) → {Date}</span><span class="sxs-lookup"><span data-stu-id="4c04c-161">convertToUtcClientTime(input) → {Date}</span></span>

<span data-ttu-id="4c04c-162">Получает объект Date из словаря, содержащего сведения о времени.</span><span class="sxs-lookup"><span data-stu-id="4c04c-162">Gets a Date object from a dictionary containing time information.</span></span>

<span data-ttu-id="4c04c-163">Метод `convertToUtcClientTime` преобразует словарь, содержащий локальную дату и время, в объект Date с правильными значениями локальной даты и времени.</span><span class="sxs-lookup"><span data-stu-id="4c04c-163">The `convertToUtcClientTime` method converts a dictionary containing a local date and time to a Date object with the correct values for the local date and time.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4c04c-164">Параметры:</span><span class="sxs-lookup"><span data-stu-id="4c04c-164">Parameters:</span></span>

|<span data-ttu-id="4c04c-165">Имя</span><span class="sxs-lookup"><span data-stu-id="4c04c-165">Name</span></span>| <span data-ttu-id="4c04c-166">Тип</span><span class="sxs-lookup"><span data-stu-id="4c04c-166">Type</span></span>| <span data-ttu-id="4c04c-167">Описание</span><span class="sxs-lookup"><span data-stu-id="4c04c-167">Description</span></span>|
|---|---|---|
|`input`| [<span data-ttu-id="4c04c-168">LocalClientTime</span><span class="sxs-lookup"><span data-stu-id="4c04c-168">LocalClientTime</span></span>](/javascript/api/outlook_1_1/office.LocalClientTime)|<span data-ttu-id="4c04c-169">Значение локального времени для преобразования.</span><span class="sxs-lookup"><span data-stu-id="4c04c-169">The local time value to convert.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4c04c-170">Требования</span><span class="sxs-lookup"><span data-stu-id="4c04c-170">Requirements</span></span>

|<span data-ttu-id="4c04c-171">Требование</span><span class="sxs-lookup"><span data-stu-id="4c04c-171">Requirement</span></span>| <span data-ttu-id="4c04c-172">Значение</span><span class="sxs-lookup"><span data-stu-id="4c04c-172">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c04c-173">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="4c04c-173">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c04c-174">1.0</span><span class="sxs-lookup"><span data-stu-id="4c04c-174">1.0</span></span>|
|[<span data-ttu-id="4c04c-175">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4c04c-175">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c04c-176">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4c04c-176">ReadItem</span></span>|
|[<span data-ttu-id="4c04c-177">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4c04c-177">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c04c-178">Compose или read</span><span class="sxs-lookup"><span data-stu-id="4c04c-178">Compose or read</span></span>|

##### <a name="returns"></a><span data-ttu-id="4c04c-179">Возвращаемое значение:</span><span class="sxs-lookup"><span data-stu-id="4c04c-179">Returns:</span></span>

<span data-ttu-id="4c04c-180">Объект Date со временем в формате UTC.</span><span class="sxs-lookup"><span data-stu-id="4c04c-180">A Date object with the time expressed in UTC.</span></span>

<dl class="param-type"><span data-ttu-id="4c04c-181">

<dt>Тип</dt>

</span><span class="sxs-lookup"><span data-stu-id="4c04c-181">

<dt>Type</dt>

</span></span><dd><span data-ttu-id="4c04c-182">Date</span><span class="sxs-lookup"><span data-stu-id="4c04c-182">Date</span></span></dd>

</dl>

####  <a name="displayappointmentformitemid"></a><span data-ttu-id="4c04c-183">displayAppointmentForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="4c04c-183">displayAppointmentForm(itemId)</span></span>

<span data-ttu-id="4c04c-184">Отображает имеющуюся встречу из календаря.</span><span class="sxs-lookup"><span data-stu-id="4c04c-184">Displays an existing calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="4c04c-185">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="4c04c-185">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4c04c-186">Метод `displayAppointmentForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее сведения календаря о существующей встрече.</span><span class="sxs-lookup"><span data-stu-id="4c04c-186">The `displayAppointmentForm` method opens an existing calendar appointment in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="4c04c-p106">В Outlook для Mac с помощью этого метода можно отобразить одну встречу, которая не является частью повторяющегося ряда, или основную встречу такого ряда, но не экземпляр из него, так как в Outlook для Mac невозможно получить доступ к свойствам экземпляра повторяющегося ряда (в том числе к идентификатору элемента).</span><span class="sxs-lookup"><span data-stu-id="4c04c-p106">In Outlook for Mac, you can use this method to display a single appointment that is not part of a recurring series, or the master appointment of a recurring series, but you cannot display an instance of the series. This is because in Outlook for Mac, you cannot access the properties (including the item ID) of instances of a recurring series.</span></span>

<span data-ttu-id="4c04c-189">В Outlook Web App этот метод открывает указанную форму, только если текст формы содержит символы размером не более 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="4c04c-189">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32KB number of characters.</span></span>

<span data-ttu-id="4c04c-190">Если указанный идентификатор элемента не определяет существующую встречу, на клиентском компьютере или устройстве открывается пустая страница, и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="4c04c-190">If the specified item identifier does not identify an existing appointment, a blank pane opens on the client computer or device, and no error message will be returned.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4c04c-191">Параметры:</span><span class="sxs-lookup"><span data-stu-id="4c04c-191">Parameters:</span></span>

|<span data-ttu-id="4c04c-192">Имя</span><span class="sxs-lookup"><span data-stu-id="4c04c-192">Name</span></span>| <span data-ttu-id="4c04c-193">Тип</span><span class="sxs-lookup"><span data-stu-id="4c04c-193">Type</span></span>| <span data-ttu-id="4c04c-194">Описание</span><span class="sxs-lookup"><span data-stu-id="4c04c-194">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="4c04c-195">String</span><span class="sxs-lookup"><span data-stu-id="4c04c-195">String</span></span>|<span data-ttu-id="4c04c-196">Идентификатор веб-служб Exchange для существующей встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="4c04c-196">The Exchange Web Services (EWS) identifier for an existing calendar appointment.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4c04c-197">Требования</span><span class="sxs-lookup"><span data-stu-id="4c04c-197">Requirements</span></span>

|<span data-ttu-id="4c04c-198">Требование</span><span class="sxs-lookup"><span data-stu-id="4c04c-198">Requirement</span></span>| <span data-ttu-id="4c04c-199">Значение</span><span class="sxs-lookup"><span data-stu-id="4c04c-199">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c04c-200">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="4c04c-200">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c04c-201">1.0</span><span class="sxs-lookup"><span data-stu-id="4c04c-201">1.0</span></span>|
|[<span data-ttu-id="4c04c-202">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4c04c-202">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c04c-203">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4c04c-203">ReadItem</span></span>|
|[<span data-ttu-id="4c04c-204">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4c04c-204">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c04c-205">Compose или read</span><span class="sxs-lookup"><span data-stu-id="4c04c-205">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4c04c-206">Пример</span><span class="sxs-lookup"><span data-stu-id="4c04c-206">Example</span></span>

```
Office.context.mailbox.displayAppointmentForm(appointmentId);
```

####  <a name="displaymessageformitemid"></a><span data-ttu-id="4c04c-207">displayMessageForm(itemId)</span><span class="sxs-lookup"><span data-stu-id="4c04c-207">displayMessageForm(itemId)</span></span>

<span data-ttu-id="4c04c-208">Отображает имеющееся сообщение.</span><span class="sxs-lookup"><span data-stu-id="4c04c-208">Displays an existing message.</span></span>

> [!NOTE]
> <span data-ttu-id="4c04c-209">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="4c04c-209">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4c04c-210">Метод `displayMessageForm` открывает новое окно на компьютере или диалоговое окно на мобильном устройстве, содержащее существующее сообщение.</span><span class="sxs-lookup"><span data-stu-id="4c04c-210">The `displayMessageForm` method opens an existing message in a new window on the desktop or in a dialog box on mobile devices.</span></span>

<span data-ttu-id="4c04c-211">В Outlook Web App этот метод открывает указанную форму, только если текст формы содержит символы размером не более 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="4c04c-211">In Outlook Web App, this method opens the specified form only if the body of the form is less than or equal to 32 KB number of characters.</span></span>

<span data-ttu-id="4c04c-212">Если указанный идентификатор элемента не определяет существующее сообщение, окно на клиентском компьютере не открывается и сообщение об ошибке не возвращается.</span><span class="sxs-lookup"><span data-stu-id="4c04c-212">If the specified item identifier does not identify an existing message, no message will be displayed on the client computer, and no error message will be returned.</span></span>

<span data-ttu-id="4c04c-p107">Не используйте `displayMessageForm` с параметром `itemId`, который представляет собой встречу. Используйте метод `displayAppointmentForm`, чтобы отобразить сведения о существующей встрече, а метод `displayNewAppointmentForm` — для отображения формы создания встречи.</span><span class="sxs-lookup"><span data-stu-id="4c04c-p107">Do not use the `displayMessageForm` with an `itemId` that represents an appointment. Use the `displayAppointmentForm` method to display an existing appointment, and `displayNewAppointmentForm` to display a form to create a new appointment.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4c04c-215">Параметры:</span><span class="sxs-lookup"><span data-stu-id="4c04c-215">Parameters:</span></span>

|<span data-ttu-id="4c04c-216">Имя</span><span class="sxs-lookup"><span data-stu-id="4c04c-216">Name</span></span>| <span data-ttu-id="4c04c-217">Тип</span><span class="sxs-lookup"><span data-stu-id="4c04c-217">Type</span></span>| <span data-ttu-id="4c04c-218">Описание</span><span class="sxs-lookup"><span data-stu-id="4c04c-218">Description</span></span>|
|---|---|---|
|`itemId`| <span data-ttu-id="4c04c-219">String</span><span class="sxs-lookup"><span data-stu-id="4c04c-219">String</span></span>|<span data-ttu-id="4c04c-220">Идентификатор веб-служб Exchange для существующего сообщения.</span><span class="sxs-lookup"><span data-stu-id="4c04c-220">The Exchange Web Services (EWS) identifier for an existing message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4c04c-221">Требования</span><span class="sxs-lookup"><span data-stu-id="4c04c-221">Requirements</span></span>

|<span data-ttu-id="4c04c-222">Требование</span><span class="sxs-lookup"><span data-stu-id="4c04c-222">Requirement</span></span>| <span data-ttu-id="4c04c-223">Значение</span><span class="sxs-lookup"><span data-stu-id="4c04c-223">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c04c-224">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="4c04c-224">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c04c-225">1.0</span><span class="sxs-lookup"><span data-stu-id="4c04c-225">1.0</span></span>|
|[<span data-ttu-id="4c04c-226">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4c04c-226">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c04c-227">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4c04c-227">ReadItem</span></span>|
|[<span data-ttu-id="4c04c-228">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4c04c-228">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c04c-229">Compose или read</span><span class="sxs-lookup"><span data-stu-id="4c04c-229">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4c04c-230">Пример</span><span class="sxs-lookup"><span data-stu-id="4c04c-230">Example</span></span>

```
Office.context.mailbox.displayMessageForm(messageId);
```

#### <a name="displaynewappointmentformparameters"></a><span data-ttu-id="4c04c-231">displayNewAppointmentForm(parameters)</span><span class="sxs-lookup"><span data-stu-id="4c04c-231">displayNewAppointmentForm(parameters)</span></span>

<span data-ttu-id="4c04c-232">Отображает форму для создания новой встречи в календаре.</span><span class="sxs-lookup"><span data-stu-id="4c04c-232">Displays a form for creating a new calendar appointment.</span></span>

> [!NOTE]
> <span data-ttu-id="4c04c-233">Этот метод не поддерживается в Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="4c04c-233">Note: This method is not supported in Outlook for iOS or Outlook for Android.</span></span>

<span data-ttu-id="4c04c-p108">Метод `displayNewAppointmentForm` открывает форму, в которой пользователь может создать встречу или собрание. Если параметры заданы, поля формы встречи автоматически заполняются их содержимым.</span><span class="sxs-lookup"><span data-stu-id="4c04c-p108">The `displayNewAppointmentForm` method opens a form that enables the user to create a new appointment or meeting. If parameters are specified, the appointment form fields are automatically populated with the contents of the parameters.</span></span>

<span data-ttu-id="4c04c-p109">В Outlook Web App и Outlook Web App для устройств этот метод всегда отображает форму с полем участников. Если вы не укажете участников в качестве входных аргументов, метод отображает форму с кнопкой **Сохранить**. Если вы укажете участников, форма будет включать участников и кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="4c04c-p109">In Outlook Web App and OWA for Devices, this method always displays a form with an attendees field. If you do not specify any attendees as input arguments, the method displays a form with a **Save** button. If you have specified attendees, the form would include the attendees and a **Send** button.</span></span>

<span data-ttu-id="4c04c-p110">Если вы укажете участников или ресурсы с помощью параметра `requiredAttendees`, `optionalAttendees` или `resources` в клиенте Outlook с расширенными возможностями и Outlook RT, этот метод отобразит форму собрания с кнопкой **Отправить**. Если не указать получателей, этот метод отобразит форму встречи с кнопкой **Сохранить и закрыть**.</span><span class="sxs-lookup"><span data-stu-id="4c04c-p110">In the Outlook rich client and Outlook RT, if you specify any attendees or resources in the `requiredAttendees`, `optionalAttendees`, or `resources` parameter, this method displays a meeting form with a **Send** button. If you don't specify any recipients, this method displays an appointment form with a **Save & Close** button.</span></span>

<span data-ttu-id="4c04c-241">Если параметры превышают указанные ограничения размера или если указано неизвестное имя параметра, вызывается исключение.</span><span class="sxs-lookup"><span data-stu-id="4c04c-241">If any of the parameters exceed the specified size limits, or if an unknown parameter name is specified, an exception is thrown.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4c04c-242">Параметры:</span><span class="sxs-lookup"><span data-stu-id="4c04c-242">Parameters:</span></span>

|<span data-ttu-id="4c04c-243">Имя</span><span class="sxs-lookup"><span data-stu-id="4c04c-243">Name</span></span>| <span data-ttu-id="4c04c-244">Тип</span><span class="sxs-lookup"><span data-stu-id="4c04c-244">Type</span></span>| <span data-ttu-id="4c04c-245">Описание</span><span class="sxs-lookup"><span data-stu-id="4c04c-245">Description</span></span>|
|---|---|---|
| `parameters` | <span data-ttu-id="4c04c-246">Object</span><span class="sxs-lookup"><span data-stu-id="4c04c-246">Object</span></span> | <span data-ttu-id="4c04c-247">Словарь параметров, описывающий новую встречу.</span><span class="sxs-lookup"><span data-stu-id="4c04c-247">A dictionary of parameters describing the new appointment.</span></span> |
| `parameters.requiredAttendees` | <span data-ttu-id="4c04c-248">Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="4c04c-248">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="4c04c-p111">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из обязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="4c04c-p111">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the required attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.optionalAttendees` | <span data-ttu-id="4c04c-251">Array.&lt;String&gt; | Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)&gt;</span><span class="sxs-lookup"><span data-stu-id="4c04c-251">Array.&lt;String&gt; &#124; Array.&lt;[EmailAddressDetails](/javascript/api/outlook_1_1/office.emailaddressdetails)&gt;</span></span> | <span data-ttu-id="4c04c-p112">Массив строк, содержащий электронные адреса, или массив, содержащий объекты `EmailAddressDetails` для каждого из необязательных участников встречи. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="4c04c-p112">An array of strings containing the email addresses or an array containing an `EmailAddressDetails` object for each of the optional attendees for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.start` | <span data-ttu-id="4c04c-254">Date</span><span class="sxs-lookup"><span data-stu-id="4c04c-254">Date</span></span> | <span data-ttu-id="4c04c-255">Объект `Date`, указывающий дату и время начала встречи.</span><span class="sxs-lookup"><span data-stu-id="4c04c-255">A `Date` object specifying the start date and time of the appointment.</span></span> |
| `parameters.end` | <span data-ttu-id="4c04c-256">Date</span><span class="sxs-lookup"><span data-stu-id="4c04c-256">Date</span></span> | <span data-ttu-id="4c04c-257">Объект `Date`, указывающий дату и время окончания встречи.</span><span class="sxs-lookup"><span data-stu-id="4c04c-257">A `Date` object specifying the end date and time of the appointment.</span></span> |
| `parameters.location` | <span data-ttu-id="4c04c-258">String</span><span class="sxs-lookup"><span data-stu-id="4c04c-258">String</span></span> | <span data-ttu-id="4c04c-p113">Строка со сведениями о месте встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="4c04c-p113">A string containing the location of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.resources` | <span data-ttu-id="4c04c-261">Array.&lt;String&gt;</span><span class="sxs-lookup"><span data-stu-id="4c04c-261">Array.&lt;String&gt;</span></span> | <span data-ttu-id="4c04c-p114">Массив строк, содержащий необходимые для встречи ресурсы. Массив может включать не более 100 записей.</span><span class="sxs-lookup"><span data-stu-id="4c04c-p114">An array of strings containing the resources required for the appointment. The array is limited to a maximum of 100 entries.</span></span> |
| `parameters.subject` | <span data-ttu-id="4c04c-264">String</span><span class="sxs-lookup"><span data-stu-id="4c04c-264">String</span></span> | <span data-ttu-id="4c04c-p115">Строка с темой встречи. Максимальное количество символов в строке — 255.</span><span class="sxs-lookup"><span data-stu-id="4c04c-p115">A string containing the subject of the appointment. The string is limited to a maximum of 255 characters.</span></span> |
| `parameters.body` | <span data-ttu-id="4c04c-267">String</span><span class="sxs-lookup"><span data-stu-id="4c04c-267">String</span></span> | <span data-ttu-id="4c04c-p116">Текст сообщения о встрече. Максимальный размер содержимого сообщения — 32 КБ.</span><span class="sxs-lookup"><span data-stu-id="4c04c-p116">The body of the appointment. The body content is limited to a maximum size of 32 KB.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="4c04c-270">Требования</span><span class="sxs-lookup"><span data-stu-id="4c04c-270">Requirements</span></span>

|<span data-ttu-id="4c04c-271">Требование</span><span class="sxs-lookup"><span data-stu-id="4c04c-271">Requirement</span></span>| <span data-ttu-id="4c04c-272">Значение</span><span class="sxs-lookup"><span data-stu-id="4c04c-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c04c-273">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="4c04c-273">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c04c-274">1.0</span><span class="sxs-lookup"><span data-stu-id="4c04c-274">1.0</span></span>|
|[<span data-ttu-id="4c04c-275">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4c04c-275">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c04c-276">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4c04c-276">ReadItem</span></span>|
|[<span data-ttu-id="4c04c-277">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4c04c-277">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c04c-278">Read (чтение)</span><span class="sxs-lookup"><span data-stu-id="4c04c-278">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4c04c-279">Пример</span><span class="sxs-lookup"><span data-stu-id="4c04c-279">Example</span></span>

```
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

#### <a name="getcallbacktokenasynccallback-usercontext"></a><span data-ttu-id="4c04c-280">getCallbackTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="4c04c-280">getCallbackTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="4c04c-281">Получает строку, содержащую маркер, используемый для получения вложения или элемента с Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="4c04c-281">Gets a string that contains a token used to get an attachment or item from an Exchange Server.</span></span>

<span data-ttu-id="4c04c-p117">Метод `getCallbackTokenAsync` совершает асинхронный вызов, чтобы получить непрозрачный маркер с сервера Exchange Server, на котором размещен почтовый ящик пользователя. Время существования маркера обратного вызова составляет 5 минут.</span><span class="sxs-lookup"><span data-stu-id="4c04c-p117">The `getCallbackTokenAsync` method makes an asynchronous call to get an opaque token from the Exchange Server that hosts the user's mailbox. The lifetime of the callback token is 5 minutes.</span></span>

<span data-ttu-id="4c04c-p118">Вы можете передать сторонней системе маркер и идентификатор вложения или элемента. Сторонняя система использует этот маркер как маркер авторизации, чтобы вызвать операцию [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) или [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) веб-служб Exchange для возврата вложения или элемента. Например, вы можете создать удаленную службу, чтобы [получить вложения из выбранного элемента](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span><span class="sxs-lookup"><span data-stu-id="4c04c-p118">You can pass the token and an attachment identifier or item identifier to a third-party system. The third-party system uses the token as a bearer authorization token to call the Exchange Web Services (EWS) [GetAttachment](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getattachment-operation) or [GetItem](https://docs.microsoft.com/exchange/client-developer/web-service-reference/getitem-operation) operation to return an attachment or item. For example, you can create a remote service to [get attachments from the selected item](https://docs.microsoft.com/outlook/add-ins/get-attachments-of-an-outlook-item).</span></span>

<span data-ttu-id="4c04c-287">Чтобы вызвать метод `getCallbackTokenAsync`, у вашего приложения должно быть разрешение **ReadItem**, указанное в его манифесте.</span><span class="sxs-lookup"><span data-stu-id="4c04c-287">Your app must have the **ReadItem** permission specified in its manifest to call the `getCallbackTokenAsync` method.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4c04c-288">Параметры:</span><span class="sxs-lookup"><span data-stu-id="4c04c-288">Parameters:</span></span>

|<span data-ttu-id="4c04c-289">Имя</span><span class="sxs-lookup"><span data-stu-id="4c04c-289">Name</span></span>| <span data-ttu-id="4c04c-290">Тип</span><span class="sxs-lookup"><span data-stu-id="4c04c-290">Type</span></span>| <span data-ttu-id="4c04c-291">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4c04c-291">Attributes</span></span>| <span data-ttu-id="4c04c-292">Описание</span><span class="sxs-lookup"><span data-stu-id="4c04c-292">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="4c04c-293">function</span><span class="sxs-lookup"><span data-stu-id="4c04c-293">function</span></span>||<span data-ttu-id="4c04c-294">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4c04c-294">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4c04c-295">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4c04c-295">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="4c04c-296">Object</span><span class="sxs-lookup"><span data-stu-id="4c04c-296">Object</span></span>| <span data-ttu-id="4c04c-297">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4c04c-297">&lt;optional&gt;</span></span>|<span data-ttu-id="4c04c-298">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="4c04c-298">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4c04c-299">Требования</span><span class="sxs-lookup"><span data-stu-id="4c04c-299">Requirements</span></span>

|<span data-ttu-id="4c04c-300">Требование</span><span class="sxs-lookup"><span data-stu-id="4c04c-300">Requirement</span></span>| <span data-ttu-id="4c04c-301">Значение</span><span class="sxs-lookup"><span data-stu-id="4c04c-301">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c04c-302">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="4c04c-302">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c04c-303">1.0</span><span class="sxs-lookup"><span data-stu-id="4c04c-303">1.0</span></span>|
|[<span data-ttu-id="4c04c-304">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4c04c-304">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c04c-305">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4c04c-305">ReadItem</span></span>|
|[<span data-ttu-id="4c04c-306">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4c04c-306">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c04c-307">Read (чтение)</span><span class="sxs-lookup"><span data-stu-id="4c04c-307">Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4c04c-308">Пример</span><span class="sxs-lookup"><span data-stu-id="4c04c-308">Example</span></span>

```js
function getCallbackToken() {
  Office.context.mailbox.getCallbackTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="getuseridentitytokenasynccallback-usercontext"></a><span data-ttu-id="4c04c-309">getUserIdentityTokenAsync(callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="4c04c-309">getUserIdentityTokenAsync(callback, [userContext])</span></span>

<span data-ttu-id="4c04c-310">Получает маркер, идентифицирующий пользователя и надстройку Office.</span><span class="sxs-lookup"><span data-stu-id="4c04c-310">Gets a token identifying the user and the Office Add-in.</span></span>

<span data-ttu-id="4c04c-311">Метод `getUserIdentityTokenAsync` возвращает маркер, который можно использовать для идентификации, а также [проверки подлинности надстройки и пользователя в сторонней системе](https://docs.microsoft.com/outlook/add-ins/authentication).</span><span class="sxs-lookup"><span data-stu-id="4c04c-311">The `getUserIdentityTokenAsync` method returns a token that you can use to identify and [authenticate the add-in and user with a third-party system](https://docs.microsoft.com/outlook/add-ins/authentication).</span></span>

##### <a name="parameters"></a><span data-ttu-id="4c04c-312">Параметры:</span><span class="sxs-lookup"><span data-stu-id="4c04c-312">Parameters:</span></span>

|<span data-ttu-id="4c04c-313">Имя</span><span class="sxs-lookup"><span data-stu-id="4c04c-313">Name</span></span>| <span data-ttu-id="4c04c-314">Тип</span><span class="sxs-lookup"><span data-stu-id="4c04c-314">Type</span></span>| <span data-ttu-id="4c04c-315">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4c04c-315">Attributes</span></span>| <span data-ttu-id="4c04c-316">Описание</span><span class="sxs-lookup"><span data-stu-id="4c04c-316">Description</span></span>|
|---|---|---|---|
|`callback`| <span data-ttu-id="4c04c-317">function</span><span class="sxs-lookup"><span data-stu-id="4c04c-317">function</span></span>||<span data-ttu-id="4c04c-318">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4c04c-318">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4c04c-319">Маркер указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4c04c-319">The token is provided as a string in the `asyncResult.value` property.</span></span>|
|`userContext`| <span data-ttu-id="4c04c-320">Object</span><span class="sxs-lookup"><span data-stu-id="4c04c-320">Object</span></span>| <span data-ttu-id="4c04c-321">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4c04c-321">&lt;optional&gt;</span></span>|<span data-ttu-id="4c04c-322">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="4c04c-322">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4c04c-323">Требования</span><span class="sxs-lookup"><span data-stu-id="4c04c-323">Requirements</span></span>

|<span data-ttu-id="4c04c-324">Требование</span><span class="sxs-lookup"><span data-stu-id="4c04c-324">Requirement</span></span>| <span data-ttu-id="4c04c-325">Значение</span><span class="sxs-lookup"><span data-stu-id="4c04c-325">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c04c-326">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="4c04c-326">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c04c-327">1.0</span><span class="sxs-lookup"><span data-stu-id="4c04c-327">1.0</span></span>|
|[<span data-ttu-id="4c04c-328">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4c04c-328">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c04c-329">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4c04c-329">ReadItem</span></span>|
|[<span data-ttu-id="4c04c-330">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4c04c-330">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c04c-331">Compose или read</span><span class="sxs-lookup"><span data-stu-id="4c04c-331">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4c04c-332">Пример</span><span class="sxs-lookup"><span data-stu-id="4c04c-332">Example</span></span>

```js
function getIdentityToken() {
  Office.context.mailbox.getUserIdentityTokenAsync(cb);
}

function cb(asyncResult) {
  var token = asyncResult.value;
}
```

####  <a name="makeewsrequestasyncdata-callback-usercontext"></a><span data-ttu-id="4c04c-333">makeEwsRequestAsync(data, callback, [userContext])</span><span class="sxs-lookup"><span data-stu-id="4c04c-333">makeEwsRequestAsync(data, callback, [userContext])</span></span>

<span data-ttu-id="4c04c-334">Выполняет асинхронный запрос для веб-служб Exchange (EWS) на сервере Exchange Server, на котором размещен почтовый ящик пользователя.</span><span class="sxs-lookup"><span data-stu-id="4c04c-334">Makes an asynchronous request to an Exchange Web Services (EWS) service on the Exchange server that hosts the user’s mailbox.</span></span>

> [!NOTE]
> <span data-ttu-id="4c04c-335">Этот метод не поддерживается в следующих сценариях.</span><span class="sxs-lookup"><span data-stu-id="4c04c-335">This method is not supported in the following scenarios.</span></span>
> - <span data-ttu-id="4c04c-336">В Outlook для iOS или Outlook для Android.</span><span class="sxs-lookup"><span data-stu-id="4c04c-336">In Outlook for iOS or Outlook for Android</span></span>
> - <span data-ttu-id="4c04c-337">Когда надстройка загружается в почтовом ящике Gmail</span><span class="sxs-lookup"><span data-stu-id="4c04c-337">When the add-in is loaded in a Gmail mailbox</span></span>
> 
> <span data-ttu-id="4c04c-338">Вместо этого надстройкам следует использовать [API-интерфейсы REST](https://docs.microsoft.com/outlook/add-ins/use-rest-api) для доступа к почтовому ящику пользователя.</span><span class="sxs-lookup"><span data-stu-id="4c04c-338">In these cases, add-ins should [use REST APIs](https://docs.microsoft.com/outlook/add-ins/use-rest-api) to access the user's mailbox instead.</span></span>

<span data-ttu-id="4c04c-339">Метод `makeEwsRequestAsync` отправляет запрос EWS от имени надстройки в Exchange.</span><span class="sxs-lookup"><span data-stu-id="4c04c-339">The `makeEwsRequestAsync` method sends an EWS request on behalf of the add-in to Exchange.</span></span> <span data-ttu-id="4c04c-340">См. [Вызов веб-служб из надстройки Outlook](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) для ознакомления с информацией о списке поддерживаемых операций веб-служб Exchange.</span><span class="sxs-lookup"><span data-stu-id="4c04c-340">See [Call web services from an Outlook add-in](https://docs.microsoft.com/outlook/add-ins/web-services#ews-operations-that-add-ins-support) for a list of the supported EWS operations.</span></span>

<span data-ttu-id="4c04c-341">С помощью метода `makeEwsRequestAsync` невозможно запрашивать элементы, связанные с папкой.</span><span class="sxs-lookup"><span data-stu-id="4c04c-341">You cannot request Folder Associated Items with the `makeEwsRequestAsync` method.</span></span>

<span data-ttu-id="4c04c-342">XML-запрос должен указывать кодировку UTF-8.</span><span class="sxs-lookup"><span data-stu-id="4c04c-342">The XML request must specify UTF-8 encoding.</span></span>

```
<?xml version="1.0" encoding="utf-8"?>
```

<span data-ttu-id="4c04c-p120">У вашей надстройки должно быть разрешение **ReadWriteMailbox** для использования метода `makeEwsRequestAsync`. Сведения об использовании разрешения **ReadWriteMailbox** и операций EWS, которые можно вызывать с помощью метода `makeEwsRequestAsync`, см. в статье [Указание разрешений для доступа почтовой надстройки к почтовому ящику пользователя](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span><span class="sxs-lookup"><span data-stu-id="4c04c-p120">Your add-in must have the **ReadWriteMailbox** permission to use the `makeEwsRequestAsync` method. For information about using the **ReadWriteMailbox** permission and the EWS operations that you can call with the `makeEwsRequestAsync` method, see [Specify permissions for mail add-in access to the user's mailbox](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions).</span></span>

> [!NOTE]
> <span data-ttu-id="4c04c-345">Администратор сервера должен установить для `OAuthAuthentication` значение true в каталоге сервера клиентского доступа EWS, чтобы включить метод  `makeEwsRequestAsync` для запросов служб EWS.</span><span class="sxs-lookup"><span data-stu-id="4c04c-345">NOTE: The server administrator must set `OAuthAuthentication` to true on the Client Access Server EWS directory to enable the `makeEwsRequestAsync` method to make EWS requests.</span></span>

##### <a name="version-differences"></a><span data-ttu-id="4c04c-346">Различия версий</span><span class="sxs-lookup"><span data-stu-id="4c04c-346">Version differences</span></span>

<span data-ttu-id="4c04c-347">Если вы используете метод `makeEwsRequestAsync` в почтовых приложениях, которые выполняются в Outlook версии, предшествующей 15.0.4535.1004, указывайте кодировку `ISO-8859-1`.</span><span class="sxs-lookup"><span data-stu-id="4c04c-347">When you use the `makeEwsRequestAsync` method in mail apps running in Outlook versions earlier than version 15.0.4535.1004, you should set the encoding value to `ISO-8859-1`.</span></span>

```
<?xml version="1.0" encoding="iso-8859-1"?>
```

<span data-ttu-id="4c04c-p121">Значение кодировки не нужно указывать, если почтовое приложение выполняется в Outlook в Интернете. Чтобы определить, выполняется ли приложение в Outlook или Outlook в Интернете, используйте свойство mailbox.diagnostics.hostName. Используемую версию Outlook можно определить с помощью свойства mailbox.diagnostics.hostVersion.</span><span class="sxs-lookup"><span data-stu-id="4c04c-p121">You do not need to set the encoding value when your mail app is running in Outlook on the web. You can determine whether your mail app is running in Outlook or Outlook on the web by using the mailbox.diagnostics.hostName property. You can determine what version of Outlook is running by using the mailbox.diagnostics.hostVersion property.</span></span>

##### <a name="parameters"></a><span data-ttu-id="4c04c-351">Параметры:</span><span class="sxs-lookup"><span data-stu-id="4c04c-351">Parameters:</span></span>

|<span data-ttu-id="4c04c-352">Имя</span><span class="sxs-lookup"><span data-stu-id="4c04c-352">Name</span></span>| <span data-ttu-id="4c04c-353">Тип</span><span class="sxs-lookup"><span data-stu-id="4c04c-353">Type</span></span>| <span data-ttu-id="4c04c-354">Атрибуты</span><span class="sxs-lookup"><span data-stu-id="4c04c-354">Attributes</span></span>| <span data-ttu-id="4c04c-355">Описание</span><span class="sxs-lookup"><span data-stu-id="4c04c-355">Description</span></span>|
|---|---|---|---|
|`data`| <span data-ttu-id="4c04c-356">String</span><span class="sxs-lookup"><span data-stu-id="4c04c-356">String</span></span>||<span data-ttu-id="4c04c-357">Запрос EWS.</span><span class="sxs-lookup"><span data-stu-id="4c04c-357">The EWS request.</span></span>|
|`callback`| <span data-ttu-id="4c04c-358">function</span><span class="sxs-lookup"><span data-stu-id="4c04c-358">function</span></span>||<span data-ttu-id="4c04c-359">После применения метода функция, переданная в параметр `callback`, вызывается с помощью параметра `asyncResult`, который представляет собой объект [`AsyncResult`](/javascript/api/office/office.asyncresult).</span><span class="sxs-lookup"><span data-stu-id="4c04c-359">When the method completes, the function passed in the `callback` parameter is called with a single parameter, `asyncResult`, which is an [`AsyncResult`](/javascript/api/office/office.asyncresult) object.</span></span><br/><br/><span data-ttu-id="4c04c-360">Результат XML вызова EWS указывается в виде строки в свойстве `asyncResult.value`.</span><span class="sxs-lookup"><span data-stu-id="4c04c-360">The XML result of the EWS call is provided as a string in the `asyncResult.value` property.</span></span> <span data-ttu-id="4c04c-361">Если размер результата превышает 1 МБ, возвращается сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="4c04c-361">If the result exceeds 1 MB in size, an error message is returned instead.| | Object| optional|Any state data that is passed to the asynchronous method.|</span></span>|
|`userContext`| <span data-ttu-id="4c04c-362">Object</span><span class="sxs-lookup"><span data-stu-id="4c04c-362">Object</span></span>| <span data-ttu-id="4c04c-363">&lt;optional&gt;</span><span class="sxs-lookup"><span data-stu-id="4c04c-363">&lt;optional&gt;</span></span>|<span data-ttu-id="4c04c-364">Данные о состоянии, передаваемые в асинхронный метод.</span><span class="sxs-lookup"><span data-stu-id="4c04c-364">Any state data that is passed to the asynchronous method.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="4c04c-365">Требования</span><span class="sxs-lookup"><span data-stu-id="4c04c-365">Requirements</span></span>

|<span data-ttu-id="4c04c-366">Требование</span><span class="sxs-lookup"><span data-stu-id="4c04c-366">Requirement</span></span>| <span data-ttu-id="4c04c-367">Значение</span><span class="sxs-lookup"><span data-stu-id="4c04c-367">Value</span></span>|
|---|---|
|[<span data-ttu-id="4c04c-368">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="4c04c-368">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4c04c-369">1.0</span><span class="sxs-lookup"><span data-stu-id="4c04c-369">1.0</span></span>|
|[<span data-ttu-id="4c04c-370">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4c04c-370">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4c04c-371">ReadWriteMailbox</span><span class="sxs-lookup"><span data-stu-id="4c04c-371">ReadWriteMailbox</span></span>|
|[<span data-ttu-id="4c04c-372">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4c04c-372">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4c04c-373">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="4c04c-373">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4c04c-374">Пример</span><span class="sxs-lookup"><span data-stu-id="4c04c-374">Example</span></span>

<span data-ttu-id="4c04c-375">В следующем примере вызывается `makeEwsRequestAsync` для получения темы элемента с помощью операции `GetItem`.</span><span class="sxs-lookup"><span data-stu-id="4c04c-375">The following example calls `makeEwsRequestAsync` to use the `GetItem` operation to get the subject of an item.</span></span>

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