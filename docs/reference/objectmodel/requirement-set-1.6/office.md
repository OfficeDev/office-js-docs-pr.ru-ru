 

# <a name="office"></a><span data-ttu-id="9da47-101">Office</span><span class="sxs-lookup"><span data-stu-id="9da47-101">Office</span></span>

<span data-ttu-id="9da47-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="9da47-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="9da47-104">Требования</span><span class="sxs-lookup"><span data-stu-id="9da47-104">Requirements</span></span>

|<span data-ttu-id="9da47-105">Требование</span><span class="sxs-lookup"><span data-stu-id="9da47-105">Requirement</span></span>| <span data-ttu-id="9da47-106">Значение</span><span class="sxs-lookup"><span data-stu-id="9da47-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="9da47-107">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="9da47-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9da47-108">1.0</span><span class="sxs-lookup"><span data-stu-id="9da47-108">1.0</span></span>|
|[<span data-ttu-id="9da47-109">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9da47-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9da47-110">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="9da47-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="9da47-111">Члены и методы</span><span class="sxs-lookup"><span data-stu-id="9da47-111">Members and methods</span></span>

| <span data-ttu-id="9da47-112">Член</span><span class="sxs-lookup"><span data-stu-id="9da47-112">Member</span></span> | <span data-ttu-id="9da47-113">Тип</span><span class="sxs-lookup"><span data-stu-id="9da47-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="9da47-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="9da47-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="9da47-115">Член</span><span class="sxs-lookup"><span data-stu-id="9da47-115">Member</span></span> |
| [<span data-ttu-id="9da47-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="9da47-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="9da47-117">Член</span><span class="sxs-lookup"><span data-stu-id="9da47-117">Member</span></span> |
| [<span data-ttu-id="9da47-118">EventType</span><span class="sxs-lookup"><span data-stu-id="9da47-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="9da47-119">Член</span><span class="sxs-lookup"><span data-stu-id="9da47-119">Member</span></span> |
| [<span data-ttu-id="9da47-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="9da47-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="9da47-121">Член</span><span class="sxs-lookup"><span data-stu-id="9da47-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="9da47-122">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="9da47-122">Namespaces</span></span>

<span data-ttu-id="9da47-123">[context](office.context.md) — предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="9da47-123">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="9da47-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype) — включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="9da47-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="9da47-125">Члены</span><span class="sxs-lookup"><span data-stu-id="9da47-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="9da47-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="9da47-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="9da47-127">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="9da47-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="9da47-128">Тип:</span><span class="sxs-lookup"><span data-stu-id="9da47-128">Type:</span></span>

*   <span data-ttu-id="9da47-129">Строка</span><span class="sxs-lookup"><span data-stu-id="9da47-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9da47-130">Свойства:</span><span class="sxs-lookup"><span data-stu-id="9da47-130">Properties:</span></span>

|<span data-ttu-id="9da47-131">Имя</span><span class="sxs-lookup"><span data-stu-id="9da47-131">Name</span></span>| <span data-ttu-id="9da47-132">Тип</span><span class="sxs-lookup"><span data-stu-id="9da47-132">Type</span></span>| <span data-ttu-id="9da47-133">Описание</span><span class="sxs-lookup"><span data-stu-id="9da47-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="9da47-134">Строка</span><span class="sxs-lookup"><span data-stu-id="9da47-134">String</span></span>|<span data-ttu-id="9da47-135">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="9da47-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="9da47-136">Строка</span><span class="sxs-lookup"><span data-stu-id="9da47-136">String</span></span>|<span data-ttu-id="9da47-137">Вызов не удался.</span><span class="sxs-lookup"><span data-stu-id="9da47-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9da47-138">Требования</span><span class="sxs-lookup"><span data-stu-id="9da47-138">Requirements</span></span>

|<span data-ttu-id="9da47-139">Требование</span><span class="sxs-lookup"><span data-stu-id="9da47-139">Requirement</span></span>| <span data-ttu-id="9da47-140">Значение</span><span class="sxs-lookup"><span data-stu-id="9da47-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="9da47-141">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="9da47-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9da47-142">1.0</span><span class="sxs-lookup"><span data-stu-id="9da47-142">1.0</span></span>|
|[<span data-ttu-id="9da47-143">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9da47-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9da47-144">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="9da47-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="9da47-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="9da47-145">CoercionType :String</span></span>

<span data-ttu-id="9da47-146">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="9da47-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="9da47-147">Тип:</span><span class="sxs-lookup"><span data-stu-id="9da47-147">Type:</span></span>

*   <span data-ttu-id="9da47-148">Строка</span><span class="sxs-lookup"><span data-stu-id="9da47-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9da47-149">Свойства:</span><span class="sxs-lookup"><span data-stu-id="9da47-149">Properties:</span></span>

|<span data-ttu-id="9da47-150">Имя</span><span class="sxs-lookup"><span data-stu-id="9da47-150">Name</span></span>| <span data-ttu-id="9da47-151">Тип</span><span class="sxs-lookup"><span data-stu-id="9da47-151">Type</span></span>| <span data-ttu-id="9da47-152">Описание</span><span class="sxs-lookup"><span data-stu-id="9da47-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="9da47-153">Строка</span><span class="sxs-lookup"><span data-stu-id="9da47-153">String</span></span>|<span data-ttu-id="9da47-154">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="9da47-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="9da47-155">Строка</span><span class="sxs-lookup"><span data-stu-id="9da47-155">String</span></span>|<span data-ttu-id="9da47-156">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="9da47-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9da47-157">Требования</span><span class="sxs-lookup"><span data-stu-id="9da47-157">Requirements</span></span>

|<span data-ttu-id="9da47-158">Требование</span><span class="sxs-lookup"><span data-stu-id="9da47-158">Requirement</span></span>| <span data-ttu-id="9da47-159">Значение</span><span class="sxs-lookup"><span data-stu-id="9da47-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="9da47-160">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="9da47-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9da47-161">1.0</span><span class="sxs-lookup"><span data-stu-id="9da47-161">1.0</span></span>|
|[<span data-ttu-id="9da47-162">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9da47-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9da47-163">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="9da47-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="9da47-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="9da47-164">EventType :String</span></span>

<span data-ttu-id="9da47-165">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="9da47-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="9da47-166">Тип:</span><span class="sxs-lookup"><span data-stu-id="9da47-166">Type:</span></span>

*   <span data-ttu-id="9da47-167">Строка</span><span class="sxs-lookup"><span data-stu-id="9da47-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9da47-168">Свойства:</span><span class="sxs-lookup"><span data-stu-id="9da47-168">Properties:</span></span>

| <span data-ttu-id="9da47-169">Имя</span><span class="sxs-lookup"><span data-stu-id="9da47-169">Name</span></span> | <span data-ttu-id="9da47-170">Тип</span><span class="sxs-lookup"><span data-stu-id="9da47-170">Type</span></span> | <span data-ttu-id="9da47-171">Описание</span><span class="sxs-lookup"><span data-stu-id="9da47-171">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="9da47-172">Строка</span><span class="sxs-lookup"><span data-stu-id="9da47-172">String</span></span> | <span data-ttu-id="9da47-173">Выбранный элемент изменился.</span><span class="sxs-lookup"><span data-stu-id="9da47-173">The selected item has changed.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="9da47-174">Требования</span><span class="sxs-lookup"><span data-stu-id="9da47-174">Requirements</span></span>

|<span data-ttu-id="9da47-175">Требование</span><span class="sxs-lookup"><span data-stu-id="9da47-175">Requirement</span></span>| <span data-ttu-id="9da47-176">Значение</span><span class="sxs-lookup"><span data-stu-id="9da47-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="9da47-177">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9da47-177">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9da47-178">1.5</span><span class="sxs-lookup"><span data-stu-id="9da47-178">1.5</span></span> |
|[<span data-ttu-id="9da47-179">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9da47-179">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9da47-180">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="9da47-180">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="9da47-181">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="9da47-181">SourceProperty :String</span></span>

<span data-ttu-id="9da47-182">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="9da47-182">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="9da47-183">Тип:</span><span class="sxs-lookup"><span data-stu-id="9da47-183">Type:</span></span>

*   <span data-ttu-id="9da47-184">Строка</span><span class="sxs-lookup"><span data-stu-id="9da47-184">String</span></span>

##### <a name="properties"></a><span data-ttu-id="9da47-185">Свойства:</span><span class="sxs-lookup"><span data-stu-id="9da47-185">Properties:</span></span>

|<span data-ttu-id="9da47-186">Имя</span><span class="sxs-lookup"><span data-stu-id="9da47-186">Name</span></span>| <span data-ttu-id="9da47-187">Тип</span><span class="sxs-lookup"><span data-stu-id="9da47-187">Type</span></span>| <span data-ttu-id="9da47-188">Описание</span><span class="sxs-lookup"><span data-stu-id="9da47-188">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="9da47-189">Строка</span><span class="sxs-lookup"><span data-stu-id="9da47-189">String</span></span>|<span data-ttu-id="9da47-190">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="9da47-190">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="9da47-191">Строка</span><span class="sxs-lookup"><span data-stu-id="9da47-191">String</span></span>|<span data-ttu-id="9da47-192">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="9da47-192">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="9da47-193">Требования</span><span class="sxs-lookup"><span data-stu-id="9da47-193">Requirements</span></span>

|<span data-ttu-id="9da47-194">Требование</span><span class="sxs-lookup"><span data-stu-id="9da47-194">Requirement</span></span>| <span data-ttu-id="9da47-195">Значение</span><span class="sxs-lookup"><span data-stu-id="9da47-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="9da47-196">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="9da47-196">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="9da47-197">1.0</span><span class="sxs-lookup"><span data-stu-id="9da47-197">1.0</span></span>|
|[<span data-ttu-id="9da47-198">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9da47-198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="9da47-199">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="9da47-199">Compose or read</span></span>|