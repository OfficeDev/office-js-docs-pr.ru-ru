 

# <a name="office"></a><span data-ttu-id="9bd70-101">Office</span><span class="sxs-lookup"><span data-stu-id="9bd70-101">Office</span></span>

<span data-ttu-id="9bd70-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="9bd70-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="9bd70-104">Требования</span><span class="sxs-lookup"><span data-stu-id="9bd70-104">Requirements</span></span>

|<span data-ttu-id="9bd70-105">Requirement</span><span class="sxs-lookup"><span data-stu-id="9bd70-105">Requirement</span></span>| <span data-ttu-id="9bd70-106">Значение</span><span class="sxs-lookup"><span data-stu-id="9bd70-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="9bd70-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9bd70-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9bd70-108">1.0</span><span class="sxs-lookup"><span data-stu-id="9bd70-108">1.0</span></span>|
|[<span data-ttu-id="9bd70-109">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9bd70-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9bd70-110">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9bd70-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="9bd70-111">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="9bd70-111">Members and methods</span></span>

| <span data-ttu-id="9bd70-112">Член</span><span class="sxs-lookup"><span data-stu-id="9bd70-112">Member</span></span> | <span data-ttu-id="9bd70-113">Тип</span><span class="sxs-lookup"><span data-stu-id="9bd70-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="9bd70-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="9bd70-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="9bd70-115">Член</span><span class="sxs-lookup"><span data-stu-id="9bd70-115">Member</span></span> |
| [<span data-ttu-id="9bd70-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="9bd70-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="9bd70-117">Член</span><span class="sxs-lookup"><span data-stu-id="9bd70-117">Member</span></span> |
| [<span data-ttu-id="9bd70-118">EventType</span><span class="sxs-lookup"><span data-stu-id="9bd70-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="9bd70-119">Член</span><span class="sxs-lookup"><span data-stu-id="9bd70-119">Member</span></span> |
| [<span data-ttu-id="9bd70-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="9bd70-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="9bd70-121">Член</span><span class="sxs-lookup"><span data-stu-id="9bd70-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="9bd70-122">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="9bd70-122">Namespaces</span></span>

<span data-ttu-id="9bd70-123">[context.](office.context.md) Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="9bd70-123">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="9bd70-124">[MailboxEnums.](/javascript/api/outlook/office.mailboxenums.attachmenttype) Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="9bd70-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="9bd70-125">Элементы</span><span class="sxs-lookup"><span data-stu-id="9bd70-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="9bd70-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="9bd70-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="9bd70-127">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="9bd70-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="9bd70-128">Тип:</span><span class="sxs-lookup"><span data-stu-id="9bd70-128">Type:</span></span>

*   <span data-ttu-id="9bd70-129">String</span><span class="sxs-lookup"><span data-stu-id="9bd70-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9bd70-130">Свойства:</span><span class="sxs-lookup"><span data-stu-id="9bd70-130">Properties:</span></span>

|<span data-ttu-id="9bd70-131">Имя</span><span class="sxs-lookup"><span data-stu-id="9bd70-131">Name</span></span>| <span data-ttu-id="9bd70-132">Тип</span><span class="sxs-lookup"><span data-stu-id="9bd70-132">Type</span></span>| <span data-ttu-id="9bd70-133">Описание</span><span class="sxs-lookup"><span data-stu-id="9bd70-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="9bd70-134">Для указания</span><span class="sxs-lookup"><span data-stu-id="9bd70-134">String</span></span>|<span data-ttu-id="9bd70-135">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="9bd70-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="9bd70-136">Для указания</span><span class="sxs-lookup"><span data-stu-id="9bd70-136">String</span></span>|<span data-ttu-id="9bd70-137">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="9bd70-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9bd70-138">Требования</span><span class="sxs-lookup"><span data-stu-id="9bd70-138">Requirements</span></span>

|<span data-ttu-id="9bd70-139">Requirement</span><span class="sxs-lookup"><span data-stu-id="9bd70-139">Requirement</span></span>| <span data-ttu-id="9bd70-140">Значение</span><span class="sxs-lookup"><span data-stu-id="9bd70-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="9bd70-141">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9bd70-141">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9bd70-142">1.0</span><span class="sxs-lookup"><span data-stu-id="9bd70-142">1.0</span></span>|
|[<span data-ttu-id="9bd70-143">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9bd70-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9bd70-144">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9bd70-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="9bd70-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="9bd70-145">CoercionType :String</span></span>

<span data-ttu-id="9bd70-146">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="9bd70-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="9bd70-147">Тип:</span><span class="sxs-lookup"><span data-stu-id="9bd70-147">Type:</span></span>

*   <span data-ttu-id="9bd70-148">String</span><span class="sxs-lookup"><span data-stu-id="9bd70-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9bd70-149">Свойства:</span><span class="sxs-lookup"><span data-stu-id="9bd70-149">Properties:</span></span>

|<span data-ttu-id="9bd70-150">Имя</span><span class="sxs-lookup"><span data-stu-id="9bd70-150">Name</span></span>| <span data-ttu-id="9bd70-151">Тип</span><span class="sxs-lookup"><span data-stu-id="9bd70-151">Type</span></span>| <span data-ttu-id="9bd70-152">Описание</span><span class="sxs-lookup"><span data-stu-id="9bd70-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="9bd70-153">String</span><span class="sxs-lookup"><span data-stu-id="9bd70-153">String</span></span>|<span data-ttu-id="9bd70-154">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="9bd70-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="9bd70-155">String</span><span class="sxs-lookup"><span data-stu-id="9bd70-155">String</span></span>|<span data-ttu-id="9bd70-156">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="9bd70-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9bd70-157">Требования</span><span class="sxs-lookup"><span data-stu-id="9bd70-157">Requirements</span></span>

|<span data-ttu-id="9bd70-158">Requirement</span><span class="sxs-lookup"><span data-stu-id="9bd70-158">Requirement</span></span>| <span data-ttu-id="9bd70-159">Значение</span><span class="sxs-lookup"><span data-stu-id="9bd70-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="9bd70-160">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9bd70-160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9bd70-161">1.0</span><span class="sxs-lookup"><span data-stu-id="9bd70-161">1.0</span></span>|
|[<span data-ttu-id="9bd70-162">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9bd70-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9bd70-163">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9bd70-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="9bd70-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="9bd70-164">EventType :String</span></span>

<span data-ttu-id="9bd70-165">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="9bd70-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="9bd70-166">Тип:</span><span class="sxs-lookup"><span data-stu-id="9bd70-166">Type:</span></span>

*   <span data-ttu-id="9bd70-167">String</span><span class="sxs-lookup"><span data-stu-id="9bd70-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9bd70-168">Свойства:</span><span class="sxs-lookup"><span data-stu-id="9bd70-168">Properties:</span></span>

| <span data-ttu-id="9bd70-169">Имя</span><span class="sxs-lookup"><span data-stu-id="9bd70-169">Name</span></span> | <span data-ttu-id="9bd70-170">Тип</span><span class="sxs-lookup"><span data-stu-id="9bd70-170">Type</span></span> | <span data-ttu-id="9bd70-171">Описание</span><span class="sxs-lookup"><span data-stu-id="9bd70-171">Description</span></span> | <span data-ttu-id="9bd70-172">Минимальный набор обязательных элементов</span><span class="sxs-lookup"><span data-stu-id="9bd70-172">Minimum mailbox requirement set version</span></span> |
|---|---|---|---|
|`AppointmentTimeChanged`| <span data-ttu-id="9bd70-173">String</span><span class="sxs-lookup"><span data-stu-id="9bd70-173">String</span></span> | <span data-ttu-id="9bd70-174">Произошло изменение даты или времени выбранной встречи либо ряда встреч.</span><span class="sxs-lookup"><span data-stu-id="9bd70-174">The date or time of the selected appointment or series has changed.</span></span> | <span data-ttu-id="9bd70-175">1.7</span><span class="sxs-lookup"><span data-stu-id="9bd70-175">ExcelApi 1.7 Beta</span></span> |
|`AttachmentsChanged`| <span data-ttu-id="9bd70-176">String</span><span class="sxs-lookup"><span data-stu-id="9bd70-176">String</span></span> | <span data-ttu-id="9bd70-177">Было добавлено или удалено вложение для элемента.</span><span class="sxs-lookup"><span data-stu-id="9bd70-177">An attachment has been added to or removed from the item.</span></span> | <span data-ttu-id="9bd70-178">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="9bd70-178">Preview</span></span> |
|`ItemChanged`| <span data-ttu-id="9bd70-179">String</span><span class="sxs-lookup"><span data-stu-id="9bd70-179">String</span></span> | <span data-ttu-id="9bd70-180">Пока область задач закреплена, для просмотра выбран другой элемент Outlook.</span><span class="sxs-lookup"><span data-stu-id="9bd70-180">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> | <span data-ttu-id="9bd70-181">1.5</span><span class="sxs-lookup"><span data-stu-id="9bd70-181">1.5</span></span> |
|`OfficeThemeChanged`| <span data-ttu-id="9bd70-182">String</span><span class="sxs-lookup"><span data-stu-id="9bd70-182">String</span></span> | <span data-ttu-id="9bd70-183">Тема Office в почтовом ящике была изменена.</span><span class="sxs-lookup"><span data-stu-id="9bd70-183">The Office theme on the mailbox has changed.</span></span> | <span data-ttu-id="9bd70-184">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="9bd70-184">Preview</span></span> |
|`RecipientsChanged`| <span data-ttu-id="9bd70-185">String</span><span class="sxs-lookup"><span data-stu-id="9bd70-185">String</span></span> | <span data-ttu-id="9bd70-186">Произошло изменение списка получателей выбранного элемента или места встречи.</span><span class="sxs-lookup"><span data-stu-id="9bd70-186">The recipient list of the selected item or appointment location has changed.</span></span> | <span data-ttu-id="9bd70-187">1.7</span><span class="sxs-lookup"><span data-stu-id="9bd70-187">ExcelApi 1.7 Beta</span></span> |
|`RecurrenceChanged`| <span data-ttu-id="9bd70-188">String</span><span class="sxs-lookup"><span data-stu-id="9bd70-188">String</span></span> | <span data-ttu-id="9bd70-189">Расписание повторения выбранного ряда элементов изменилось.</span><span class="sxs-lookup"><span data-stu-id="9bd70-189">The recurrence pattern of the selected series has changed.</span></span> | <span data-ttu-id="9bd70-190">1.7</span><span class="sxs-lookup"><span data-stu-id="9bd70-190">ExcelApi 1.7 Beta</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9bd70-191">Требования</span><span class="sxs-lookup"><span data-stu-id="9bd70-191">Requirements</span></span>

|<span data-ttu-id="9bd70-192">Requirement</span><span class="sxs-lookup"><span data-stu-id="9bd70-192">Requirement</span></span>| <span data-ttu-id="9bd70-193">Значение</span><span class="sxs-lookup"><span data-stu-id="9bd70-193">Value</span></span>|
|---|---|
|[<span data-ttu-id="9bd70-194">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9bd70-194">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9bd70-195">1.5</span><span class="sxs-lookup"><span data-stu-id="9bd70-195">1.5</span></span> |
|[<span data-ttu-id="9bd70-196">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9bd70-196">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9bd70-197">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9bd70-197">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="9bd70-198">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="9bd70-198">SourceProperty :String</span></span>

<span data-ttu-id="9bd70-199">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="9bd70-199">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="9bd70-200">Тип:</span><span class="sxs-lookup"><span data-stu-id="9bd70-200">Type:</span></span>

*   <span data-ttu-id="9bd70-201">String</span><span class="sxs-lookup"><span data-stu-id="9bd70-201">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9bd70-202">Свойства:</span><span class="sxs-lookup"><span data-stu-id="9bd70-202">Properties:</span></span>

|<span data-ttu-id="9bd70-203">Имя</span><span class="sxs-lookup"><span data-stu-id="9bd70-203">Name</span></span>| <span data-ttu-id="9bd70-204">Тип</span><span class="sxs-lookup"><span data-stu-id="9bd70-204">Type</span></span>| <span data-ttu-id="9bd70-205">Описание</span><span class="sxs-lookup"><span data-stu-id="9bd70-205">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="9bd70-206">String</span><span class="sxs-lookup"><span data-stu-id="9bd70-206">String</span></span>|<span data-ttu-id="9bd70-207">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="9bd70-207">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="9bd70-208">String</span><span class="sxs-lookup"><span data-stu-id="9bd70-208">String</span></span>|<span data-ttu-id="9bd70-209">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="9bd70-209">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9bd70-210">Требования</span><span class="sxs-lookup"><span data-stu-id="9bd70-210">Requirements</span></span>

|<span data-ttu-id="9bd70-211">Requirement</span><span class="sxs-lookup"><span data-stu-id="9bd70-211">Requirement</span></span>| <span data-ttu-id="9bd70-212">Значение</span><span class="sxs-lookup"><span data-stu-id="9bd70-212">Value</span></span>|
|---|---|
|[<span data-ttu-id="9bd70-213">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9bd70-213">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9bd70-214">1.0</span><span class="sxs-lookup"><span data-stu-id="9bd70-214">1.0</span></span>|
|[<span data-ttu-id="9bd70-215">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9bd70-215">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9bd70-216">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9bd70-216">Compose or read</span></span>|