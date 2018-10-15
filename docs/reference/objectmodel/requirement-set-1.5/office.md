# <a name="office"></a><span data-ttu-id="860df-101">Office</span><span class="sxs-lookup"><span data-stu-id="860df-101">Office</span></span>

<span data-ttu-id="860df-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="860df-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="860df-104">Требования</span><span class="sxs-lookup"><span data-stu-id="860df-104">Requirements</span></span>

|<span data-ttu-id="860df-105">Требование</span><span class="sxs-lookup"><span data-stu-id="860df-105">Requirement</span></span>| <span data-ttu-id="860df-106">Значение</span><span class="sxs-lookup"><span data-stu-id="860df-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="860df-107">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="860df-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="860df-108">1.0</span><span class="sxs-lookup"><span data-stu-id="860df-108">1.0</span></span>|
|[<span data-ttu-id="860df-109">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="860df-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="860df-110">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="860df-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="860df-111">Члены и методы</span><span class="sxs-lookup"><span data-stu-id="860df-111">Members and methods</span></span>

| <span data-ttu-id="860df-112">Член</span><span class="sxs-lookup"><span data-stu-id="860df-112">Member</span></span> | <span data-ttu-id="860df-113">Тип</span><span class="sxs-lookup"><span data-stu-id="860df-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="860df-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="860df-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="860df-115">Член</span><span class="sxs-lookup"><span data-stu-id="860df-115">Member</span></span> |
| [<span data-ttu-id="860df-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="860df-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="860df-117">Член</span><span class="sxs-lookup"><span data-stu-id="860df-117">Member</span></span> |
| [<span data-ttu-id="860df-118">EventType</span><span class="sxs-lookup"><span data-stu-id="860df-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="860df-119">Член</span><span class="sxs-lookup"><span data-stu-id="860df-119">Member</span></span> |
| [<span data-ttu-id="860df-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="860df-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="860df-121">Член</span><span class="sxs-lookup"><span data-stu-id="860df-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="860df-122">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="860df-122">Namespaces</span></span>

<span data-ttu-id="860df-123">[context](office.context.md) — предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="860df-123">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="860df-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype) — включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="860df-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="860df-125">Члены</span><span class="sxs-lookup"><span data-stu-id="860df-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="860df-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="860df-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="860df-127">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="860df-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="860df-128">Тип:</span><span class="sxs-lookup"><span data-stu-id="860df-128">Type:</span></span>

*   <span data-ttu-id="860df-129">String</span><span class="sxs-lookup"><span data-stu-id="860df-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="860df-130">Свойства:</span><span class="sxs-lookup"><span data-stu-id="860df-130">Properties:</span></span>

|<span data-ttu-id="860df-131">Имя</span><span class="sxs-lookup"><span data-stu-id="860df-131">Name</span></span>| <span data-ttu-id="860df-132">Тип</span><span class="sxs-lookup"><span data-stu-id="860df-132">Type</span></span>| <span data-ttu-id="860df-133">Описание</span><span class="sxs-lookup"><span data-stu-id="860df-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="860df-134">String</span><span class="sxs-lookup"><span data-stu-id="860df-134">String</span></span>|<span data-ttu-id="860df-135">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="860df-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="860df-136">String</span><span class="sxs-lookup"><span data-stu-id="860df-136">String</span></span>|<span data-ttu-id="860df-137">Вызов не удался.</span><span class="sxs-lookup"><span data-stu-id="860df-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="860df-138">Требования</span><span class="sxs-lookup"><span data-stu-id="860df-138">Requirements</span></span>

|<span data-ttu-id="860df-139">Требование</span><span class="sxs-lookup"><span data-stu-id="860df-139">Requirement</span></span>| <span data-ttu-id="860df-140">Значение</span><span class="sxs-lookup"><span data-stu-id="860df-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="860df-141">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="860df-141">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="860df-142">1.0</span><span class="sxs-lookup"><span data-stu-id="860df-142">1.0</span></span>|
|[<span data-ttu-id="860df-143">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="860df-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="860df-144">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="860df-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="860df-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="860df-145">CoercionType :String</span></span>

<span data-ttu-id="860df-146">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="860df-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="860df-147">Тип:</span><span class="sxs-lookup"><span data-stu-id="860df-147">Type:</span></span>

*   <span data-ttu-id="860df-148">String</span><span class="sxs-lookup"><span data-stu-id="860df-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="860df-149">Свойства:</span><span class="sxs-lookup"><span data-stu-id="860df-149">Properties:</span></span>

|<span data-ttu-id="860df-150">Имя</span><span class="sxs-lookup"><span data-stu-id="860df-150">Name</span></span>| <span data-ttu-id="860df-151">Тип</span><span class="sxs-lookup"><span data-stu-id="860df-151">Type</span></span>| <span data-ttu-id="860df-152">Описание</span><span class="sxs-lookup"><span data-stu-id="860df-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="860df-153">String</span><span class="sxs-lookup"><span data-stu-id="860df-153">String</span></span>|<span data-ttu-id="860df-154">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="860df-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="860df-155">String</span><span class="sxs-lookup"><span data-stu-id="860df-155">String</span></span>|<span data-ttu-id="860df-156">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="860df-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="860df-157">Требования</span><span class="sxs-lookup"><span data-stu-id="860df-157">Requirements</span></span>

|<span data-ttu-id="860df-158">Требование</span><span class="sxs-lookup"><span data-stu-id="860df-158">Requirement</span></span>| <span data-ttu-id="860df-159">Значение</span><span class="sxs-lookup"><span data-stu-id="860df-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="860df-160">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="860df-160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="860df-161">1.0</span><span class="sxs-lookup"><span data-stu-id="860df-161">1.0</span></span>|
|[<span data-ttu-id="860df-162">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="860df-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="860df-163">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="860df-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="860df-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="860df-164">EventType :String</span></span>

<span data-ttu-id="860df-165">Указывает событие, связанное с обработчиком событий.</span><span class="sxs-lookup"><span data-stu-id="860df-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="860df-166">Тип:</span><span class="sxs-lookup"><span data-stu-id="860df-166">Type:</span></span>

*   <span data-ttu-id="860df-167">String</span><span class="sxs-lookup"><span data-stu-id="860df-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="860df-168">Свойства:</span><span class="sxs-lookup"><span data-stu-id="860df-168">Properties:</span></span>

| <span data-ttu-id="860df-169">Имя</span><span class="sxs-lookup"><span data-stu-id="860df-169">Name</span></span> | <span data-ttu-id="860df-170">Тип</span><span class="sxs-lookup"><span data-stu-id="860df-170">Type</span></span> | <span data-ttu-id="860df-171">Описание</span><span class="sxs-lookup"><span data-stu-id="860df-171">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="860df-172">String</span><span class="sxs-lookup"><span data-stu-id="860df-172">String</span></span> | <span data-ttu-id="860df-173">Выбранный элемент изменился.</span><span class="sxs-lookup"><span data-stu-id="860df-173">The selected item has changed.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="860df-174">Требования</span><span class="sxs-lookup"><span data-stu-id="860df-174">Requirements</span></span>

|<span data-ttu-id="860df-175">Требование</span><span class="sxs-lookup"><span data-stu-id="860df-175">Requirement</span></span>| <span data-ttu-id="860df-176">Значение</span><span class="sxs-lookup"><span data-stu-id="860df-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="860df-177">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="860df-177">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="860df-178">1.5</span><span class="sxs-lookup"><span data-stu-id="860df-178">1.5</span></span> |
|[<span data-ttu-id="860df-179">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="860df-179">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="860df-180">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="860df-180">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="860df-181">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="860df-181">SourceProperty :String</span></span>

<span data-ttu-id="860df-182">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="860df-182">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="860df-183">Тип:</span><span class="sxs-lookup"><span data-stu-id="860df-183">Type:</span></span>

*   <span data-ttu-id="860df-184">String</span><span class="sxs-lookup"><span data-stu-id="860df-184">String</span></span>

##### <a name="properties"></a><span data-ttu-id="860df-185">Свойства:</span><span class="sxs-lookup"><span data-stu-id="860df-185">Properties:</span></span>

|<span data-ttu-id="860df-186">Имя</span><span class="sxs-lookup"><span data-stu-id="860df-186">Name</span></span>| <span data-ttu-id="860df-187">Тип</span><span class="sxs-lookup"><span data-stu-id="860df-187">Type</span></span>| <span data-ttu-id="860df-188">Описание</span><span class="sxs-lookup"><span data-stu-id="860df-188">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="860df-189">String</span><span class="sxs-lookup"><span data-stu-id="860df-189">String</span></span>|<span data-ttu-id="860df-190">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="860df-190">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="860df-191">String</span><span class="sxs-lookup"><span data-stu-id="860df-191">String</span></span>|<span data-ttu-id="860df-192">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="860df-192">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="860df-193">Требования</span><span class="sxs-lookup"><span data-stu-id="860df-193">Requirements</span></span>

|<span data-ttu-id="860df-194">Требование</span><span class="sxs-lookup"><span data-stu-id="860df-194">Requirement</span></span>| <span data-ttu-id="860df-195">Значение</span><span class="sxs-lookup"><span data-stu-id="860df-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="860df-196">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="860df-196">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="860df-197">1.0</span><span class="sxs-lookup"><span data-stu-id="860df-197">1.0</span></span>|
|[<span data-ttu-id="860df-198">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="860df-198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="860df-199">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="860df-199">Compose or read</span></span>|