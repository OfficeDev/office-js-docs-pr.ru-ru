 

# <a name="office"></a><span data-ttu-id="8e949-101">Office</span><span class="sxs-lookup"><span data-stu-id="8e949-101">Office</span></span>

<span data-ttu-id="8e949-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="8e949-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="8e949-104">Требования</span><span class="sxs-lookup"><span data-stu-id="8e949-104">Requirements</span></span>

|<span data-ttu-id="8e949-105">Requirement</span><span class="sxs-lookup"><span data-stu-id="8e949-105">Requirement</span></span>| <span data-ttu-id="8e949-106">Значение</span><span class="sxs-lookup"><span data-stu-id="8e949-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e949-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8e949-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8e949-108">1.0</span><span class="sxs-lookup"><span data-stu-id="8e949-108">1.0</span></span>|
|[<span data-ttu-id="8e949-109">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8e949-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8e949-110">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8e949-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="8e949-111">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="8e949-111">Members and methods</span></span>

| <span data-ttu-id="8e949-112">Член</span><span class="sxs-lookup"><span data-stu-id="8e949-112">Member</span></span> | <span data-ttu-id="8e949-113">Тип</span><span class="sxs-lookup"><span data-stu-id="8e949-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="8e949-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="8e949-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="8e949-115">Член</span><span class="sxs-lookup"><span data-stu-id="8e949-115">Member</span></span> |
| [<span data-ttu-id="8e949-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="8e949-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="8e949-117">Член</span><span class="sxs-lookup"><span data-stu-id="8e949-117">Member</span></span> |
| [<span data-ttu-id="8e949-118">EventType</span><span class="sxs-lookup"><span data-stu-id="8e949-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="8e949-119">Член</span><span class="sxs-lookup"><span data-stu-id="8e949-119">Member</span></span> |
| [<span data-ttu-id="8e949-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="8e949-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="8e949-121">Член</span><span class="sxs-lookup"><span data-stu-id="8e949-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="8e949-122">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="8e949-122">Namespaces</span></span>

<span data-ttu-id="8e949-123">[context](office.context.md). Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="8e949-123">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="8e949-124">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype). Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="8e949-124">[MailboxEnums](/javascript/api/outlook_1_7/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="8e949-125">Элементы</span><span class="sxs-lookup"><span data-stu-id="8e949-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="8e949-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="8e949-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="8e949-127">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="8e949-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="8e949-128">Тип:</span><span class="sxs-lookup"><span data-stu-id="8e949-128">Type:</span></span>

*   <span data-ttu-id="8e949-129">String</span><span class="sxs-lookup"><span data-stu-id="8e949-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8e949-130">Свойства:</span><span class="sxs-lookup"><span data-stu-id="8e949-130">Properties:</span></span>

|<span data-ttu-id="8e949-131">Имя</span><span class="sxs-lookup"><span data-stu-id="8e949-131">Name</span></span>| <span data-ttu-id="8e949-132">Тип</span><span class="sxs-lookup"><span data-stu-id="8e949-132">Type</span></span>| <span data-ttu-id="8e949-133">Описание</span><span class="sxs-lookup"><span data-stu-id="8e949-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="8e949-134">Для указания</span><span class="sxs-lookup"><span data-stu-id="8e949-134">String</span></span>|<span data-ttu-id="8e949-135">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="8e949-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="8e949-136">Для указания</span><span class="sxs-lookup"><span data-stu-id="8e949-136">String</span></span>|<span data-ttu-id="8e949-137">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="8e949-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e949-138">Требования</span><span class="sxs-lookup"><span data-stu-id="8e949-138">Requirements</span></span>

|<span data-ttu-id="8e949-139">Requirement</span><span class="sxs-lookup"><span data-stu-id="8e949-139">Requirement</span></span>| <span data-ttu-id="8e949-140">Значение</span><span class="sxs-lookup"><span data-stu-id="8e949-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e949-141">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8e949-141">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8e949-142">1.0</span><span class="sxs-lookup"><span data-stu-id="8e949-142">1.0</span></span>|
|[<span data-ttu-id="8e949-143">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8e949-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8e949-144">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8e949-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="8e949-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="8e949-145">CoercionType :String</span></span>

<span data-ttu-id="8e949-146">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="8e949-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="8e949-147">Тип:</span><span class="sxs-lookup"><span data-stu-id="8e949-147">Type:</span></span>

*   <span data-ttu-id="8e949-148">String</span><span class="sxs-lookup"><span data-stu-id="8e949-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8e949-149">Свойства:</span><span class="sxs-lookup"><span data-stu-id="8e949-149">Properties:</span></span>

|<span data-ttu-id="8e949-150">Имя</span><span class="sxs-lookup"><span data-stu-id="8e949-150">Name</span></span>| <span data-ttu-id="8e949-151">Тип</span><span class="sxs-lookup"><span data-stu-id="8e949-151">Type</span></span>| <span data-ttu-id="8e949-152">Описание</span><span class="sxs-lookup"><span data-stu-id="8e949-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="8e949-153">String</span><span class="sxs-lookup"><span data-stu-id="8e949-153">String</span></span>|<span data-ttu-id="8e949-154">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="8e949-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="8e949-155">String</span><span class="sxs-lookup"><span data-stu-id="8e949-155">String</span></span>|<span data-ttu-id="8e949-156">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="8e949-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e949-157">Требования</span><span class="sxs-lookup"><span data-stu-id="8e949-157">Requirements</span></span>

|<span data-ttu-id="8e949-158">Requirement</span><span class="sxs-lookup"><span data-stu-id="8e949-158">Requirement</span></span>| <span data-ttu-id="8e949-159">Значение</span><span class="sxs-lookup"><span data-stu-id="8e949-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e949-160">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8e949-160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8e949-161">1.0</span><span class="sxs-lookup"><span data-stu-id="8e949-161">1.0</span></span>|
|[<span data-ttu-id="8e949-162">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8e949-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8e949-163">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8e949-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="8e949-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="8e949-164">EventType :String</span></span>

<span data-ttu-id="8e949-165">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="8e949-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="8e949-166">Тип:</span><span class="sxs-lookup"><span data-stu-id="8e949-166">Type:</span></span>

*   <span data-ttu-id="8e949-167">String</span><span class="sxs-lookup"><span data-stu-id="8e949-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8e949-168">Свойства:</span><span class="sxs-lookup"><span data-stu-id="8e949-168">Properties:</span></span>

| <span data-ttu-id="8e949-169">Имя</span><span class="sxs-lookup"><span data-stu-id="8e949-169">Name</span></span> | <span data-ttu-id="8e949-170">Тип</span><span class="sxs-lookup"><span data-stu-id="8e949-170">Type</span></span> | <span data-ttu-id="8e949-171">Описание</span><span class="sxs-lookup"><span data-stu-id="8e949-171">Description</span></span> | <span data-ttu-id="8e949-172">Минимальный набор обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="8e949-172">Minimum mailbox requirement set version</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="8e949-173">String</span><span class="sxs-lookup"><span data-stu-id="8e949-173">String</span></span> | <span data-ttu-id="8e949-174">Произошло изменение даты или времени выбранной встречи либо ряда встреч.</span><span class="sxs-lookup"><span data-stu-id="8e949-174">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="8e949-175">1.7</span><span class="sxs-lookup"><span data-stu-id="8e949-175">ExcelApi 1.7 Beta</span></span> |
|`ItemChanged`| <span data-ttu-id="8e949-176">String</span><span class="sxs-lookup"><span data-stu-id="8e949-176">String</span></span> | <span data-ttu-id="8e949-177">Пока область задач закреплена, для просмотра выбран другой элемент Outlook.</span><span class="sxs-lookup"><span data-stu-id="8e949-177">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="8e949-178">1.5</span><span class="sxs-lookup"><span data-stu-id="8e949-178">1.5</span></span> |
|`RecipientsChanged`| <span data-ttu-id="8e949-179">String</span><span class="sxs-lookup"><span data-stu-id="8e949-179">String</span></span> | <span data-ttu-id="8e949-180">Произошло изменение списка получателей выбранного элемента или места встречи.</span><span class="sxs-lookup"><span data-stu-id="8e949-180">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="8e949-181">1.7</span><span class="sxs-lookup"><span data-stu-id="8e949-181">ExcelApi 1.7 Beta</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="8e949-182">String</span><span class="sxs-lookup"><span data-stu-id="8e949-182">String</span></span> | <span data-ttu-id="8e949-183">Расписание повторения выбранного ряда элементов изменилось.</span><span class="sxs-lookup"><span data-stu-id="8e949-183">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="8e949-184">1.7</span><span class="sxs-lookup"><span data-stu-id="8e949-184">ExcelApi 1.7 Beta</span></span> |

##### <a name="requirements"></a><span data-ttu-id="8e949-185">Требования</span><span class="sxs-lookup"><span data-stu-id="8e949-185">Requirements</span></span>

|<span data-ttu-id="8e949-186">Requirement</span><span class="sxs-lookup"><span data-stu-id="8e949-186">Requirement</span></span>| <span data-ttu-id="8e949-187">Значение</span><span class="sxs-lookup"><span data-stu-id="8e949-187">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e949-188">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="8e949-188">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8e949-189">1.5</span><span class="sxs-lookup"><span data-stu-id="8e949-189">1.5</span></span> |
|[<span data-ttu-id="8e949-190">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8e949-190">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8e949-191">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8e949-191">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="8e949-192">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="8e949-192">SourceProperty :String</span></span>

<span data-ttu-id="8e949-193">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="8e949-193">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="8e949-194">Тип:</span><span class="sxs-lookup"><span data-stu-id="8e949-194">Type:</span></span>

*   <span data-ttu-id="8e949-195">String</span><span class="sxs-lookup"><span data-stu-id="8e949-195">String</span></span>

##### <a name="properties"></a><span data-ttu-id="8e949-196">Свойства:</span><span class="sxs-lookup"><span data-stu-id="8e949-196">Properties:</span></span>

|<span data-ttu-id="8e949-197">Имя</span><span class="sxs-lookup"><span data-stu-id="8e949-197">Name</span></span>| <span data-ttu-id="8e949-198">Тип</span><span class="sxs-lookup"><span data-stu-id="8e949-198">Type</span></span>| <span data-ttu-id="8e949-199">Описание</span><span class="sxs-lookup"><span data-stu-id="8e949-199">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="8e949-200">String</span><span class="sxs-lookup"><span data-stu-id="8e949-200">String</span></span>|<span data-ttu-id="8e949-201">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="8e949-201">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="8e949-202">String</span><span class="sxs-lookup"><span data-stu-id="8e949-202">String</span></span>|<span data-ttu-id="8e949-203">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="8e949-203">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="8e949-204">Требования</span><span class="sxs-lookup"><span data-stu-id="8e949-204">Requirements</span></span>

|<span data-ttu-id="8e949-205">Requirement</span><span class="sxs-lookup"><span data-stu-id="8e949-205">Requirement</span></span>| <span data-ttu-id="8e949-206">Значение</span><span class="sxs-lookup"><span data-stu-id="8e949-206">Value</span></span>|
|---|---|
|[<span data-ttu-id="8e949-207">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8e949-207">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8e949-208">1.0</span><span class="sxs-lookup"><span data-stu-id="8e949-208">1.0</span></span>|
|[<span data-ttu-id="8e949-209">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8e949-209">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8e949-210">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8e949-210">Compose or read</span></span>|