# <a name="office"></a><span data-ttu-id="d4510-101">Office</span><span class="sxs-lookup"><span data-stu-id="d4510-101">Office</span></span>

<span data-ttu-id="d4510-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="d4510-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="d4510-104">Требования</span><span class="sxs-lookup"><span data-stu-id="d4510-104">Requirements</span></span>

|<span data-ttu-id="d4510-105">Требование</span><span class="sxs-lookup"><span data-stu-id="d4510-105">Requirement</span></span>| <span data-ttu-id="d4510-106">Значение</span><span class="sxs-lookup"><span data-stu-id="d4510-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4510-107">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d4510-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d4510-108">1.0</span><span class="sxs-lookup"><span data-stu-id="d4510-108">1.0</span></span>|
|[<span data-ttu-id="d4510-109">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d4510-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d4510-110">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="d4510-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="d4510-111">Члены и методы</span><span class="sxs-lookup"><span data-stu-id="d4510-111">Members and methods</span></span>

| <span data-ttu-id="d4510-112">Член</span><span class="sxs-lookup"><span data-stu-id="d4510-112">Member</span></span> | <span data-ttu-id="d4510-113">Тип</span><span class="sxs-lookup"><span data-stu-id="d4510-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="d4510-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="d4510-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="d4510-115">Член</span><span class="sxs-lookup"><span data-stu-id="d4510-115">Member</span></span> |
| [<span data-ttu-id="d4510-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="d4510-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="d4510-117">Член</span><span class="sxs-lookup"><span data-stu-id="d4510-117">Member</span></span> |
| [<span data-ttu-id="d4510-118">EventType</span><span class="sxs-lookup"><span data-stu-id="d4510-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="d4510-119">Член</span><span class="sxs-lookup"><span data-stu-id="d4510-119">Member</span></span> |
| [<span data-ttu-id="d4510-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="d4510-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="d4510-121">Член</span><span class="sxs-lookup"><span data-stu-id="d4510-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="d4510-122">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="d4510-122">Namespaces</span></span>

<span data-ttu-id="d4510-123">[context](office.context.md). Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="d4510-123">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="d4510-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype). Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="d4510-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="d4510-125">Члены</span><span class="sxs-lookup"><span data-stu-id="d4510-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="d4510-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="d4510-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="d4510-127">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="d4510-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="d4510-128">Тип:</span><span class="sxs-lookup"><span data-stu-id="d4510-128">Type:</span></span>

*   <span data-ttu-id="d4510-129">String</span><span class="sxs-lookup"><span data-stu-id="d4510-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d4510-130">Свойства:</span><span class="sxs-lookup"><span data-stu-id="d4510-130">Properties:</span></span>

|<span data-ttu-id="d4510-131">Name</span><span class="sxs-lookup"><span data-stu-id="d4510-131">Name</span></span>| <span data-ttu-id="d4510-132">Тип</span><span class="sxs-lookup"><span data-stu-id="d4510-132">Type</span></span>| <span data-ttu-id="d4510-133">Описание</span><span class="sxs-lookup"><span data-stu-id="d4510-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="d4510-134">Строка</span><span class="sxs-lookup"><span data-stu-id="d4510-134">String</span></span>|<span data-ttu-id="d4510-135">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="d4510-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="d4510-136">Строка</span><span class="sxs-lookup"><span data-stu-id="d4510-136">String</span></span>|<span data-ttu-id="d4510-137">Вызов не удался.</span><span class="sxs-lookup"><span data-stu-id="d4510-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d4510-138">Требования</span><span class="sxs-lookup"><span data-stu-id="d4510-138">Requirements</span></span>

|<span data-ttu-id="d4510-139">Требование</span><span class="sxs-lookup"><span data-stu-id="d4510-139">Requirement</span></span>| <span data-ttu-id="d4510-140">Значение</span><span class="sxs-lookup"><span data-stu-id="d4510-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4510-141">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d4510-141">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d4510-142">1.0</span><span class="sxs-lookup"><span data-stu-id="d4510-142">1.0</span></span>|
|[<span data-ttu-id="d4510-143">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d4510-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d4510-144">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="d4510-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="d4510-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="d4510-145">CoercionType :String</span></span>

<span data-ttu-id="d4510-146">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="d4510-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="d4510-147">Тип:</span><span class="sxs-lookup"><span data-stu-id="d4510-147">Type:</span></span>

*   <span data-ttu-id="d4510-148">Строка</span><span class="sxs-lookup"><span data-stu-id="d4510-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d4510-149">Свойства:</span><span class="sxs-lookup"><span data-stu-id="d4510-149">Properties:</span></span>

|<span data-ttu-id="d4510-150">Name</span><span class="sxs-lookup"><span data-stu-id="d4510-150">Name</span></span>| <span data-ttu-id="d4510-151">Тип</span><span class="sxs-lookup"><span data-stu-id="d4510-151">Type</span></span>| <span data-ttu-id="d4510-152">Описание</span><span class="sxs-lookup"><span data-stu-id="d4510-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="d4510-153">Строка</span><span class="sxs-lookup"><span data-stu-id="d4510-153">String</span></span>|<span data-ttu-id="d4510-154">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="d4510-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="d4510-155">Строка</span><span class="sxs-lookup"><span data-stu-id="d4510-155">String</span></span>|<span data-ttu-id="d4510-156">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="d4510-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d4510-157">Требования</span><span class="sxs-lookup"><span data-stu-id="d4510-157">Requirements</span></span>

|<span data-ttu-id="d4510-158">Требование</span><span class="sxs-lookup"><span data-stu-id="d4510-158">Requirement</span></span>| <span data-ttu-id="d4510-159">Значение</span><span class="sxs-lookup"><span data-stu-id="d4510-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4510-160">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d4510-160">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d4510-161">1.0</span><span class="sxs-lookup"><span data-stu-id="d4510-161">1.0</span></span>|
|[<span data-ttu-id="d4510-162">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d4510-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d4510-163">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="d4510-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="d4510-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="d4510-164">EventType :String</span></span>

<span data-ttu-id="d4510-165">Указывает событие, связанное с обработчиком событий.</span><span class="sxs-lookup"><span data-stu-id="d4510-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="d4510-166">Тип:</span><span class="sxs-lookup"><span data-stu-id="d4510-166">Type:</span></span>

*   <span data-ttu-id="d4510-167">Строка</span><span class="sxs-lookup"><span data-stu-id="d4510-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d4510-168">Свойства:</span><span class="sxs-lookup"><span data-stu-id="d4510-168">Properties:</span></span>

| <span data-ttu-id="d4510-169">Name</span><span class="sxs-lookup"><span data-stu-id="d4510-169">Name</span></span> | <span data-ttu-id="d4510-170">Тип</span><span class="sxs-lookup"><span data-stu-id="d4510-170">Type</span></span> | <span data-ttu-id="d4510-171">Описание</span><span class="sxs-lookup"><span data-stu-id="d4510-171">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="d4510-172">Строка</span><span class="sxs-lookup"><span data-stu-id="d4510-172">String</span></span> | <span data-ttu-id="d4510-173">Выбранный элемент изменился.</span><span class="sxs-lookup"><span data-stu-id="d4510-173">The selected item has changed.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="d4510-174">Требования</span><span class="sxs-lookup"><span data-stu-id="d4510-174">Requirements</span></span>

|<span data-ttu-id="d4510-175">Требование</span><span class="sxs-lookup"><span data-stu-id="d4510-175">Requirement</span></span>| <span data-ttu-id="d4510-176">Значение</span><span class="sxs-lookup"><span data-stu-id="d4510-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4510-177">Версия минимального набора требований для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="d4510-177">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d4510-178">1.5</span><span class="sxs-lookup"><span data-stu-id="d4510-178">1.5</span></span> |
|[<span data-ttu-id="d4510-179">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d4510-179">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d4510-180">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="d4510-180">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="d4510-181">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="d4510-181">SourceProperty :String</span></span>

<span data-ttu-id="d4510-182">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="d4510-182">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="d4510-183">Тип:</span><span class="sxs-lookup"><span data-stu-id="d4510-183">Type:</span></span>

*   <span data-ttu-id="d4510-184">Строка</span><span class="sxs-lookup"><span data-stu-id="d4510-184">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d4510-185">Свойства:</span><span class="sxs-lookup"><span data-stu-id="d4510-185">Properties:</span></span>

|<span data-ttu-id="d4510-186">Name</span><span class="sxs-lookup"><span data-stu-id="d4510-186">Name</span></span>| <span data-ttu-id="d4510-187">Тип</span><span class="sxs-lookup"><span data-stu-id="d4510-187">Type</span></span>| <span data-ttu-id="d4510-188">Описание</span><span class="sxs-lookup"><span data-stu-id="d4510-188">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="d4510-189">Строка</span><span class="sxs-lookup"><span data-stu-id="d4510-189">String</span></span>|<span data-ttu-id="d4510-190">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="d4510-190">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="d4510-191">Строка</span><span class="sxs-lookup"><span data-stu-id="d4510-191">String</span></span>|<span data-ttu-id="d4510-192">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="d4510-192">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d4510-193">Требования</span><span class="sxs-lookup"><span data-stu-id="d4510-193">Requirements</span></span>

|<span data-ttu-id="d4510-194">Требование</span><span class="sxs-lookup"><span data-stu-id="d4510-194">Requirement</span></span>| <span data-ttu-id="d4510-195">Значение</span><span class="sxs-lookup"><span data-stu-id="d4510-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="d4510-196">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d4510-196">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d4510-197">1.0</span><span class="sxs-lookup"><span data-stu-id="d4510-197">1.0</span></span>|
|[<span data-ttu-id="d4510-198">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d4510-198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d4510-199">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="d4510-199">Compose or read</span></span>|