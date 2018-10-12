 

# <a name="office"></a><span data-ttu-id="6954b-101">Office</span><span class="sxs-lookup"><span data-stu-id="6954b-101">Office</span></span>

<span data-ttu-id="6954b-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="6954b-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="6954b-104">Требования</span><span class="sxs-lookup"><span data-stu-id="6954b-104">Requirements</span></span>

|<span data-ttu-id="6954b-105">Требование</span><span class="sxs-lookup"><span data-stu-id="6954b-105">Requirement</span></span>| <span data-ttu-id="6954b-106">Значение</span><span class="sxs-lookup"><span data-stu-id="6954b-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="6954b-107">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="6954b-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6954b-108">1.0</span><span class="sxs-lookup"><span data-stu-id="6954b-108">1.0</span></span>|
|[<span data-ttu-id="6954b-109">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6954b-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6954b-110">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6954b-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="6954b-111">Члены и методы</span><span class="sxs-lookup"><span data-stu-id="6954b-111">Members and methods</span></span>

| <span data-ttu-id="6954b-112">Член</span><span class="sxs-lookup"><span data-stu-id="6954b-112">Member</span></span> | <span data-ttu-id="6954b-113">Тип</span><span class="sxs-lookup"><span data-stu-id="6954b-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="6954b-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="6954b-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="6954b-115">Член</span><span class="sxs-lookup"><span data-stu-id="6954b-115">Member</span></span> |
| [<span data-ttu-id="6954b-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="6954b-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="6954b-117">Член</span><span class="sxs-lookup"><span data-stu-id="6954b-117">Member</span></span> |
| [<span data-ttu-id="6954b-118">EventType</span><span class="sxs-lookup"><span data-stu-id="6954b-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="6954b-119">Член</span><span class="sxs-lookup"><span data-stu-id="6954b-119">Member</span></span> |
| [<span data-ttu-id="6954b-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="6954b-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="6954b-121">Член</span><span class="sxs-lookup"><span data-stu-id="6954b-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="6954b-122">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="6954b-122">Namespaces</span></span>

<span data-ttu-id="6954b-123">[context](office.context.md) — предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="6954b-123">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="6954b-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype) — включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="6954b-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="6954b-125">Члены</span><span class="sxs-lookup"><span data-stu-id="6954b-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="6954b-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="6954b-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="6954b-127">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="6954b-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="6954b-128">Тип:</span><span class="sxs-lookup"><span data-stu-id="6954b-128">Type:</span></span>

*   <span data-ttu-id="6954b-129">Строка​</span><span class="sxs-lookup"><span data-stu-id="6954b-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6954b-130">Свойства:</span><span class="sxs-lookup"><span data-stu-id="6954b-130">Properties:</span></span>

|<span data-ttu-id="6954b-131">Name</span><span class="sxs-lookup"><span data-stu-id="6954b-131">Name</span></span>| <span data-ttu-id="6954b-132">Тип</span><span class="sxs-lookup"><span data-stu-id="6954b-132">Type</span></span>| <span data-ttu-id="6954b-133">Описание</span><span class="sxs-lookup"><span data-stu-id="6954b-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="6954b-134">Строка​</span><span class="sxs-lookup"><span data-stu-id="6954b-134">String</span></span>|<span data-ttu-id="6954b-135">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="6954b-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="6954b-136">Строка​</span><span class="sxs-lookup"><span data-stu-id="6954b-136">String</span></span>|<span data-ttu-id="6954b-137">Вызов не удался.</span><span class="sxs-lookup"><span data-stu-id="6954b-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6954b-138">Требования</span><span class="sxs-lookup"><span data-stu-id="6954b-138">Requirements</span></span>

|<span data-ttu-id="6954b-139">Требование</span><span class="sxs-lookup"><span data-stu-id="6954b-139">Requirement</span></span>| <span data-ttu-id="6954b-140">Значение</span><span class="sxs-lookup"><span data-stu-id="6954b-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="6954b-141">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="6954b-141">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6954b-142">1.0</span><span class="sxs-lookup"><span data-stu-id="6954b-142">1.0</span></span>|
|[<span data-ttu-id="6954b-143">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6954b-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6954b-144">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6954b-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="6954b-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="6954b-145">CoercionType :String</span></span>

<span data-ttu-id="6954b-146">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="6954b-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="6954b-147">Тип:</span><span class="sxs-lookup"><span data-stu-id="6954b-147">Type:</span></span>

*   <span data-ttu-id="6954b-148">Строка​</span><span class="sxs-lookup"><span data-stu-id="6954b-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6954b-149">Свойства:</span><span class="sxs-lookup"><span data-stu-id="6954b-149">Properties:</span></span>

|<span data-ttu-id="6954b-150">Имя</span><span class="sxs-lookup"><span data-stu-id="6954b-150">Name</span></span>| <span data-ttu-id="6954b-151">Тип</span><span class="sxs-lookup"><span data-stu-id="6954b-151">Type</span></span>| <span data-ttu-id="6954b-152">Описание</span><span class="sxs-lookup"><span data-stu-id="6954b-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="6954b-153">Строка​</span><span class="sxs-lookup"><span data-stu-id="6954b-153">String</span></span>|<span data-ttu-id="6954b-154">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="6954b-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="6954b-155">Строка​</span><span class="sxs-lookup"><span data-stu-id="6954b-155">String</span></span>|<span data-ttu-id="6954b-156">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="6954b-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6954b-157">Требования</span><span class="sxs-lookup"><span data-stu-id="6954b-157">Requirements</span></span>

|<span data-ttu-id="6954b-158">Требование</span><span class="sxs-lookup"><span data-stu-id="6954b-158">Requirement</span></span>| <span data-ttu-id="6954b-159">Значение</span><span class="sxs-lookup"><span data-stu-id="6954b-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="6954b-160">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="6954b-160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6954b-161">1.0</span><span class="sxs-lookup"><span data-stu-id="6954b-161">1.0</span></span>|
|[<span data-ttu-id="6954b-162">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6954b-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6954b-163">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6954b-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="6954b-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="6954b-164">EventType :String</span></span>

<span data-ttu-id="6954b-165">Указывает событие, связанное с обработчиком событий.</span><span class="sxs-lookup"><span data-stu-id="6954b-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="6954b-166">Тип:</span><span class="sxs-lookup"><span data-stu-id="6954b-166">Type:</span></span>

*   <span data-ttu-id="6954b-167">Строка​</span><span class="sxs-lookup"><span data-stu-id="6954b-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6954b-168">Свойства:</span><span class="sxs-lookup"><span data-stu-id="6954b-168">Properties:</span></span>

| <span data-ttu-id="6954b-169">Name</span><span class="sxs-lookup"><span data-stu-id="6954b-169">Name</span></span> | <span data-ttu-id="6954b-170">Тип</span><span class="sxs-lookup"><span data-stu-id="6954b-170">Type</span></span> | <span data-ttu-id="6954b-171">Описание</span><span class="sxs-lookup"><span data-stu-id="6954b-171">Description</span></span> | <span data-ttu-id="6954b-172">Минимальный набор требований</span><span class="sxs-lookup"><span data-stu-id="6954b-172">Minimum mailbox requirement set version</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="6954b-173">Строка​</span><span class="sxs-lookup"><span data-stu-id="6954b-173">String</span></span> | <span data-ttu-id="6954b-174">Дата или время выбранной встречи или серии была изменена.</span><span class="sxs-lookup"><span data-stu-id="6954b-174">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="6954b-175">1.7</span><span class="sxs-lookup"><span data-stu-id="6954b-175">17 </span></span> |
|`ItemChanged`| <span data-ttu-id="6954b-176">Строка​</span><span class="sxs-lookup"><span data-stu-id="6954b-176">String</span></span> | <span data-ttu-id="6954b-177">Выбранный элемент изменился.</span><span class="sxs-lookup"><span data-stu-id="6954b-177">The selected item has changed.</span></span> | <span data-ttu-id="6954b-178">1.5</span><span class="sxs-lookup"><span data-stu-id="6954b-178">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="6954b-179">Строка​</span><span class="sxs-lookup"><span data-stu-id="6954b-179">String</span></span> | <span data-ttu-id="6954b-180">Выбранный элемент изменился.</span><span class="sxs-lookup"><span data-stu-id="6954b-180">The selected item has changed.</span></span> | <span data-ttu-id="6954b-181">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="6954b-181">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="6954b-182">Строка​</span><span class="sxs-lookup"><span data-stu-id="6954b-182">String</span></span> | <span data-ttu-id="6954b-183">Список получателей в выбранном элементе или расположение встречи изменен(-о).</span><span class="sxs-lookup"><span data-stu-id="6954b-183">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="6954b-184">1.7</span><span class="sxs-lookup"><span data-stu-id="6954b-184">17 </span></span> |
|`RecurrenceChanged`| <span data-ttu-id="6954b-185">Строка​</span><span class="sxs-lookup"><span data-stu-id="6954b-185">String</span></span> | <span data-ttu-id="6954b-186">Расписание повторения выбранной серии было изменено.</span><span class="sxs-lookup"><span data-stu-id="6954b-186">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="6954b-187">1.7</span><span class="sxs-lookup"><span data-stu-id="6954b-187">17 </span></span> |

##### <a name="requirements"></a><span data-ttu-id="6954b-188">Требования</span><span class="sxs-lookup"><span data-stu-id="6954b-188">Requirements</span></span>

|<span data-ttu-id="6954b-189">Требование</span><span class="sxs-lookup"><span data-stu-id="6954b-189">Requirement</span></span>| <span data-ttu-id="6954b-190">Значение</span><span class="sxs-lookup"><span data-stu-id="6954b-190">Value</span></span>|
|---|---|
|[<span data-ttu-id="6954b-191">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="6954b-191">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6954b-192">1.5</span><span class="sxs-lookup"><span data-stu-id="6954b-192">1.5</span></span> |
|[<span data-ttu-id="6954b-193">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6954b-193">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6954b-194">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6954b-194">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="6954b-195">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="6954b-195">SourceProperty :String</span></span>

<span data-ttu-id="6954b-196">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="6954b-196">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="6954b-197">Тип:</span><span class="sxs-lookup"><span data-stu-id="6954b-197">Type:</span></span>

*   <span data-ttu-id="6954b-198">Строка​</span><span class="sxs-lookup"><span data-stu-id="6954b-198">String</span></span>

##### <a name="properties"></a><span data-ttu-id="6954b-199">Свойства:</span><span class="sxs-lookup"><span data-stu-id="6954b-199">Properties:</span></span>

|<span data-ttu-id="6954b-200">Name</span><span class="sxs-lookup"><span data-stu-id="6954b-200">Name</span></span>| <span data-ttu-id="6954b-201">Тип</span><span class="sxs-lookup"><span data-stu-id="6954b-201">Type</span></span>| <span data-ttu-id="6954b-202">Описание</span><span class="sxs-lookup"><span data-stu-id="6954b-202">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="6954b-203">Строка​</span><span class="sxs-lookup"><span data-stu-id="6954b-203">String</span></span>|<span data-ttu-id="6954b-204">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="6954b-204">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="6954b-205">Строка​</span><span class="sxs-lookup"><span data-stu-id="6954b-205">String</span></span>|<span data-ttu-id="6954b-206">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="6954b-206">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6954b-207">Требования</span><span class="sxs-lookup"><span data-stu-id="6954b-207">Requirements</span></span>

|<span data-ttu-id="6954b-208">Требование</span><span class="sxs-lookup"><span data-stu-id="6954b-208">Requirement</span></span>| <span data-ttu-id="6954b-209">Значение</span><span class="sxs-lookup"><span data-stu-id="6954b-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="6954b-210">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="6954b-210">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="6954b-211">1.0</span><span class="sxs-lookup"><span data-stu-id="6954b-211">1.0</span></span>|
|[<span data-ttu-id="6954b-212">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6954b-212">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="6954b-213">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6954b-213">Compose or read</span></span>|