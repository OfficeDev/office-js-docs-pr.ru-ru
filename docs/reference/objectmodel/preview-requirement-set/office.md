 

# <a name="office"></a><span data-ttu-id="521e3-101">Office</span><span class="sxs-lookup"><span data-stu-id="521e3-101">Office</span></span>

<span data-ttu-id="521e3-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="521e3-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="521e3-104">Требования</span><span class="sxs-lookup"><span data-stu-id="521e3-104">Requirements</span></span>

|<span data-ttu-id="521e3-105">Требование</span><span class="sxs-lookup"><span data-stu-id="521e3-105">Requirement</span></span>| <span data-ttu-id="521e3-106">Значение</span><span class="sxs-lookup"><span data-stu-id="521e3-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="521e3-107">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="521e3-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="521e3-108">1.0</span><span class="sxs-lookup"><span data-stu-id="521e3-108">1.0</span></span>|
|[<span data-ttu-id="521e3-109">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="521e3-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="521e3-110">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="521e3-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="521e3-111">Члены и методы</span><span class="sxs-lookup"><span data-stu-id="521e3-111">Members and methods</span></span>

| <span data-ttu-id="521e3-112">Член</span><span class="sxs-lookup"><span data-stu-id="521e3-112">Member</span></span> | <span data-ttu-id="521e3-113">Тип</span><span class="sxs-lookup"><span data-stu-id="521e3-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="521e3-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="521e3-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="521e3-115">Член</span><span class="sxs-lookup"><span data-stu-id="521e3-115">Member</span></span> |
| [<span data-ttu-id="521e3-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="521e3-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="521e3-117">Член</span><span class="sxs-lookup"><span data-stu-id="521e3-117">Member</span></span> |
| [<span data-ttu-id="521e3-118">EventType</span><span class="sxs-lookup"><span data-stu-id="521e3-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="521e3-119">Член</span><span class="sxs-lookup"><span data-stu-id="521e3-119">Member</span></span> |
| [<span data-ttu-id="521e3-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="521e3-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="521e3-121">Член</span><span class="sxs-lookup"><span data-stu-id="521e3-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="521e3-122">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="521e3-122">Namespaces</span></span>

<span data-ttu-id="521e3-123">[context](office.context.md). Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="521e3-123">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="521e3-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype). Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="521e3-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="521e3-125">Члены</span><span class="sxs-lookup"><span data-stu-id="521e3-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="521e3-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="521e3-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="521e3-127">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="521e3-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="521e3-128">Тип:</span><span class="sxs-lookup"><span data-stu-id="521e3-128">Type:</span></span>

*   <span data-ttu-id="521e3-129">String</span><span class="sxs-lookup"><span data-stu-id="521e3-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="521e3-130">Свойства:</span><span class="sxs-lookup"><span data-stu-id="521e3-130">Properties:</span></span>

|<span data-ttu-id="521e3-131">Имя</span><span class="sxs-lookup"><span data-stu-id="521e3-131">Name</span></span>| <span data-ttu-id="521e3-132">Тип</span><span class="sxs-lookup"><span data-stu-id="521e3-132">Type</span></span>| <span data-ttu-id="521e3-133">Описание</span><span class="sxs-lookup"><span data-stu-id="521e3-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="521e3-134">String</span><span class="sxs-lookup"><span data-stu-id="521e3-134">String</span></span>|<span data-ttu-id="521e3-135">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="521e3-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="521e3-136">String</span><span class="sxs-lookup"><span data-stu-id="521e3-136">String</span></span>|<span data-ttu-id="521e3-137">Вызов не удался.</span><span class="sxs-lookup"><span data-stu-id="521e3-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="521e3-138">Требования</span><span class="sxs-lookup"><span data-stu-id="521e3-138">Requirements</span></span>

|<span data-ttu-id="521e3-139">Требование</span><span class="sxs-lookup"><span data-stu-id="521e3-139">Requirement</span></span>| <span data-ttu-id="521e3-140">Значение</span><span class="sxs-lookup"><span data-stu-id="521e3-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="521e3-141">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="521e3-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="521e3-142">1.0</span><span class="sxs-lookup"><span data-stu-id="521e3-142">1.0</span></span>|
|[<span data-ttu-id="521e3-143">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="521e3-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="521e3-144">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="521e3-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="521e3-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="521e3-145">CoercionType :String</span></span>

<span data-ttu-id="521e3-146">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="521e3-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="521e3-147">Тип:</span><span class="sxs-lookup"><span data-stu-id="521e3-147">Type:</span></span>

*   <span data-ttu-id="521e3-148">String</span><span class="sxs-lookup"><span data-stu-id="521e3-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="521e3-149">Свойства:</span><span class="sxs-lookup"><span data-stu-id="521e3-149">Properties:</span></span>

|<span data-ttu-id="521e3-150">Имя</span><span class="sxs-lookup"><span data-stu-id="521e3-150">Name</span></span>| <span data-ttu-id="521e3-151">Тип</span><span class="sxs-lookup"><span data-stu-id="521e3-151">Type</span></span>| <span data-ttu-id="521e3-152">Описание</span><span class="sxs-lookup"><span data-stu-id="521e3-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="521e3-153">String</span><span class="sxs-lookup"><span data-stu-id="521e3-153">String</span></span>|<span data-ttu-id="521e3-154">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="521e3-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="521e3-155">String</span><span class="sxs-lookup"><span data-stu-id="521e3-155">String</span></span>|<span data-ttu-id="521e3-156">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="521e3-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="521e3-157">Требования</span><span class="sxs-lookup"><span data-stu-id="521e3-157">Requirements</span></span>

|<span data-ttu-id="521e3-158">Требование</span><span class="sxs-lookup"><span data-stu-id="521e3-158">Requirement</span></span>| <span data-ttu-id="521e3-159">Значение</span><span class="sxs-lookup"><span data-stu-id="521e3-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="521e3-160">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="521e3-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="521e3-161">1.0</span><span class="sxs-lookup"><span data-stu-id="521e3-161">1.0</span></span>|
|[<span data-ttu-id="521e3-162">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="521e3-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="521e3-163">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="521e3-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="521e3-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="521e3-164">EventType :String</span></span>

<span data-ttu-id="521e3-165">Указывает событие, связанное с обработчиком событий.</span><span class="sxs-lookup"><span data-stu-id="521e3-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="521e3-166">Тип:</span><span class="sxs-lookup"><span data-stu-id="521e3-166">Type:</span></span>

*   <span data-ttu-id="521e3-167">String</span><span class="sxs-lookup"><span data-stu-id="521e3-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="521e3-168">Свойства:</span><span class="sxs-lookup"><span data-stu-id="521e3-168">Properties:</span></span>

| <span data-ttu-id="521e3-169">Имя</span><span class="sxs-lookup"><span data-stu-id="521e3-169">Name</span></span> | <span data-ttu-id="521e3-170">Тип</span><span class="sxs-lookup"><span data-stu-id="521e3-170">Type</span></span> | <span data-ttu-id="521e3-171">Описание</span><span class="sxs-lookup"><span data-stu-id="521e3-171">Description</span></span> | <span data-ttu-id="521e3-172">Минимальный набор требований</span><span class="sxs-lookup"><span data-stu-id="521e3-172">Minimum mailbox requirement set version</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="521e3-173">String</span><span class="sxs-lookup"><span data-stu-id="521e3-173">String</span></span> | <span data-ttu-id="521e3-174">Дата или время выбранной встречи или серии была изменена.</span><span class="sxs-lookup"><span data-stu-id="521e3-174">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="521e3-175">1.7</span><span class="sxs-lookup"><span data-stu-id="521e3-175">17 </span></span> |
|`ItemChanged`| <span data-ttu-id="521e3-176">String</span><span class="sxs-lookup"><span data-stu-id="521e3-176">String</span></span> | <span data-ttu-id="521e3-177">Выбранный элемент изменился.</span><span class="sxs-lookup"><span data-stu-id="521e3-177">The selected item has changed.</span></span> | <span data-ttu-id="521e3-178">1.5</span><span class="sxs-lookup"><span data-stu-id="521e3-178">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="521e3-179">String</span><span class="sxs-lookup"><span data-stu-id="521e3-179">String</span></span> | <span data-ttu-id="521e3-180">Выбранный элемент изменился.</span><span class="sxs-lookup"><span data-stu-id="521e3-180">The selected item has changed.</span></span> | <span data-ttu-id="521e3-181">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="521e3-181">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="521e3-182">String</span><span class="sxs-lookup"><span data-stu-id="521e3-182">String</span></span> | <span data-ttu-id="521e3-183">Изменился список получателей в выбранном элементе или изменилось расположение встречи.</span><span class="sxs-lookup"><span data-stu-id="521e3-183">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="521e3-184">1.7</span><span class="sxs-lookup"><span data-stu-id="521e3-184">17 </span></span> |
|`RecurrenceChanged`| <span data-ttu-id="521e3-185">String</span><span class="sxs-lookup"><span data-stu-id="521e3-185">String</span></span> | <span data-ttu-id="521e3-186">Расписание повторения выбранной серии было изменено.</span><span class="sxs-lookup"><span data-stu-id="521e3-186">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="521e3-187">1.7</span><span class="sxs-lookup"><span data-stu-id="521e3-187">17 </span></span> |

##### <a name="requirements"></a><span data-ttu-id="521e3-188">Требования</span><span class="sxs-lookup"><span data-stu-id="521e3-188">Requirements</span></span>

|<span data-ttu-id="521e3-189">Требование</span><span class="sxs-lookup"><span data-stu-id="521e3-189">Requirement</span></span>| <span data-ttu-id="521e3-190">Значение</span><span class="sxs-lookup"><span data-stu-id="521e3-190">Value</span></span>|
|---|---|
|[<span data-ttu-id="521e3-191">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="521e3-191">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="521e3-192">1.5</span><span class="sxs-lookup"><span data-stu-id="521e3-192">1.5</span></span> |
|[<span data-ttu-id="521e3-193">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="521e3-193">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="521e3-194">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="521e3-194">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="521e3-195">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="521e3-195">SourceProperty :String</span></span>

<span data-ttu-id="521e3-196">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="521e3-196">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="521e3-197">Тип:</span><span class="sxs-lookup"><span data-stu-id="521e3-197">Type:</span></span>

*   <span data-ttu-id="521e3-198">String</span><span class="sxs-lookup"><span data-stu-id="521e3-198">String</span></span>

##### <a name="properties"></a><span data-ttu-id="521e3-199">Свойства:</span><span class="sxs-lookup"><span data-stu-id="521e3-199">Properties:</span></span>

|<span data-ttu-id="521e3-200">Имя</span><span class="sxs-lookup"><span data-stu-id="521e3-200">Name</span></span>| <span data-ttu-id="521e3-201">Тип</span><span class="sxs-lookup"><span data-stu-id="521e3-201">Type</span></span>| <span data-ttu-id="521e3-202">Описание</span><span class="sxs-lookup"><span data-stu-id="521e3-202">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="521e3-203">String</span><span class="sxs-lookup"><span data-stu-id="521e3-203">String</span></span>|<span data-ttu-id="521e3-204">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="521e3-204">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="521e3-205">String</span><span class="sxs-lookup"><span data-stu-id="521e3-205">String</span></span>|<span data-ttu-id="521e3-206">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="521e3-206">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="521e3-207">Требования</span><span class="sxs-lookup"><span data-stu-id="521e3-207">Requirements</span></span>

|<span data-ttu-id="521e3-208">Требование</span><span class="sxs-lookup"><span data-stu-id="521e3-208">Requirement</span></span>| <span data-ttu-id="521e3-209">Значение</span><span class="sxs-lookup"><span data-stu-id="521e3-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="521e3-210">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="521e3-210">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="521e3-211">1.0</span><span class="sxs-lookup"><span data-stu-id="521e3-211">1.0</span></span>|
|[<span data-ttu-id="521e3-212">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="521e3-212">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="521e3-213">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="521e3-213">Compose or read</span></span>|