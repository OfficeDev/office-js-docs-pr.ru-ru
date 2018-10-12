 

# <a name="office"></a><span data-ttu-id="c046f-101">Office</span><span class="sxs-lookup"><span data-stu-id="c046f-101">Office</span></span>

<span data-ttu-id="c046f-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="c046f-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="c046f-104">Требования</span><span class="sxs-lookup"><span data-stu-id="c046f-104">Requirements</span></span>

|<span data-ttu-id="c046f-105">Требование</span><span class="sxs-lookup"><span data-stu-id="c046f-105">Requirement</span></span>| <span data-ttu-id="c046f-106">Значение</span><span class="sxs-lookup"><span data-stu-id="c046f-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="c046f-107">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="c046f-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c046f-108">1.0</span><span class="sxs-lookup"><span data-stu-id="c046f-108">1.0</span></span>|
|[<span data-ttu-id="c046f-109">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c046f-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c046f-110">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="c046f-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="c046f-111">Члены и методы</span><span class="sxs-lookup"><span data-stu-id="c046f-111">Members and methods</span></span>

| <span data-ttu-id="c046f-112">Член</span><span class="sxs-lookup"><span data-stu-id="c046f-112">Member</span></span> | <span data-ttu-id="c046f-113">Тип</span><span class="sxs-lookup"><span data-stu-id="c046f-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="c046f-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="c046f-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="c046f-115">Член</span><span class="sxs-lookup"><span data-stu-id="c046f-115">Member</span></span> |
| [<span data-ttu-id="c046f-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="c046f-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="c046f-117">Член</span><span class="sxs-lookup"><span data-stu-id="c046f-117">Member</span></span> |
| [<span data-ttu-id="c046f-118">EventType</span><span class="sxs-lookup"><span data-stu-id="c046f-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="c046f-119">Член</span><span class="sxs-lookup"><span data-stu-id="c046f-119">Member</span></span> |
| [<span data-ttu-id="c046f-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="c046f-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="c046f-121">Член</span><span class="sxs-lookup"><span data-stu-id="c046f-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="c046f-122">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="c046f-122">Namespaces</span></span>

<span data-ttu-id="c046f-123">[context](office.context.md) — предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="c046f-123">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="c046f-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype) — включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="c046f-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="c046f-125">Члены</span><span class="sxs-lookup"><span data-stu-id="c046f-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="c046f-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="c046f-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="c046f-127">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="c046f-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="c046f-128">Тип:</span><span class="sxs-lookup"><span data-stu-id="c046f-128">Type:</span></span>

*   <span data-ttu-id="c046f-129">Строка​</span><span class="sxs-lookup"><span data-stu-id="c046f-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c046f-130">Свойства:</span><span class="sxs-lookup"><span data-stu-id="c046f-130">Properties:</span></span>

|<span data-ttu-id="c046f-131">Имя</span><span class="sxs-lookup"><span data-stu-id="c046f-131">Name</span></span>| <span data-ttu-id="c046f-132">Тип</span><span class="sxs-lookup"><span data-stu-id="c046f-132">Type</span></span>| <span data-ttu-id="c046f-133">Описание</span><span class="sxs-lookup"><span data-stu-id="c046f-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="c046f-134">Строка​</span><span class="sxs-lookup"><span data-stu-id="c046f-134">String</span></span>|<span data-ttu-id="c046f-135">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="c046f-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="c046f-136">Строка​</span><span class="sxs-lookup"><span data-stu-id="c046f-136">String</span></span>|<span data-ttu-id="c046f-137">Вызов не удался.</span><span class="sxs-lookup"><span data-stu-id="c046f-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c046f-138">Требования</span><span class="sxs-lookup"><span data-stu-id="c046f-138">Requirements</span></span>

|<span data-ttu-id="c046f-139">Требование</span><span class="sxs-lookup"><span data-stu-id="c046f-139">Requirement</span></span>| <span data-ttu-id="c046f-140">Значение</span><span class="sxs-lookup"><span data-stu-id="c046f-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="c046f-141">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="c046f-141">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c046f-142">1.0</span><span class="sxs-lookup"><span data-stu-id="c046f-142">1.0</span></span>|
|[<span data-ttu-id="c046f-143">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c046f-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c046f-144">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="c046f-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="c046f-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="c046f-145">CoercionType :String</span></span>

<span data-ttu-id="c046f-146">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="c046f-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="c046f-147">Тип:</span><span class="sxs-lookup"><span data-stu-id="c046f-147">Type:</span></span>

*   <span data-ttu-id="c046f-148">Строка​</span><span class="sxs-lookup"><span data-stu-id="c046f-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c046f-149">Свойства:</span><span class="sxs-lookup"><span data-stu-id="c046f-149">Properties:</span></span>

|<span data-ttu-id="c046f-150">Имя</span><span class="sxs-lookup"><span data-stu-id="c046f-150">Name</span></span>| <span data-ttu-id="c046f-151">Тип</span><span class="sxs-lookup"><span data-stu-id="c046f-151">Type</span></span>| <span data-ttu-id="c046f-152">Описание</span><span class="sxs-lookup"><span data-stu-id="c046f-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="c046f-153">Строка​</span><span class="sxs-lookup"><span data-stu-id="c046f-153">String</span></span>|<span data-ttu-id="c046f-154">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="c046f-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="c046f-155">Строка​</span><span class="sxs-lookup"><span data-stu-id="c046f-155">String</span></span>|<span data-ttu-id="c046f-156">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="c046f-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c046f-157">Требования</span><span class="sxs-lookup"><span data-stu-id="c046f-157">Requirements</span></span>

|<span data-ttu-id="c046f-158">Требование</span><span class="sxs-lookup"><span data-stu-id="c046f-158">Requirement</span></span>| <span data-ttu-id="c046f-159">Значение</span><span class="sxs-lookup"><span data-stu-id="c046f-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="c046f-160">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="c046f-160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c046f-161">1.0</span><span class="sxs-lookup"><span data-stu-id="c046f-161">1.0</span></span>|
|[<span data-ttu-id="c046f-162">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c046f-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c046f-163">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="c046f-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="c046f-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="c046f-164">EventType :String</span></span>

<span data-ttu-id="c046f-165">Указывает событие, связанное с обработчиком событий.</span><span class="sxs-lookup"><span data-stu-id="c046f-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="c046f-166">Тип:</span><span class="sxs-lookup"><span data-stu-id="c046f-166">Type:</span></span>

*   <span data-ttu-id="c046f-167">Строка​</span><span class="sxs-lookup"><span data-stu-id="c046f-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c046f-168">Свойства:</span><span class="sxs-lookup"><span data-stu-id="c046f-168">Properties:</span></span>

| <span data-ttu-id="c046f-169">Имя</span><span class="sxs-lookup"><span data-stu-id="c046f-169">Name</span></span> | <span data-ttu-id="c046f-170">Тип</span><span class="sxs-lookup"><span data-stu-id="c046f-170">Type</span></span> | <span data-ttu-id="c046f-171">Описание</span><span class="sxs-lookup"><span data-stu-id="c046f-171">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="c046f-172">Строка​</span><span class="sxs-lookup"><span data-stu-id="c046f-172">String</span></span> | <span data-ttu-id="c046f-173">Выбранный элемент изменился.</span><span class="sxs-lookup"><span data-stu-id="c046f-173">The selected item has changed.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="c046f-174">Требования</span><span class="sxs-lookup"><span data-stu-id="c046f-174">Requirements</span></span>

|<span data-ttu-id="c046f-175">Требование</span><span class="sxs-lookup"><span data-stu-id="c046f-175">Requirement</span></span>| <span data-ttu-id="c046f-176">Значение</span><span class="sxs-lookup"><span data-stu-id="c046f-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="c046f-177">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="c046f-177">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c046f-178">1.5</span><span class="sxs-lookup"><span data-stu-id="c046f-178">1.5</span></span> |
|[<span data-ttu-id="c046f-179">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c046f-179">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c046f-180">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="c046f-180">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="c046f-181">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="c046f-181">SourceProperty :String</span></span>

<span data-ttu-id="c046f-182">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="c046f-182">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="c046f-183">Тип:</span><span class="sxs-lookup"><span data-stu-id="c046f-183">Type:</span></span>

*   <span data-ttu-id="c046f-184">Строка​</span><span class="sxs-lookup"><span data-stu-id="c046f-184">String</span></span>

##### <a name="properties"></a><span data-ttu-id="c046f-185">Свойства:</span><span class="sxs-lookup"><span data-stu-id="c046f-185">Properties:</span></span>

|<span data-ttu-id="c046f-186">Имя</span><span class="sxs-lookup"><span data-stu-id="c046f-186">Name</span></span>| <span data-ttu-id="c046f-187">Тип</span><span class="sxs-lookup"><span data-stu-id="c046f-187">Type</span></span>| <span data-ttu-id="c046f-188">Описание</span><span class="sxs-lookup"><span data-stu-id="c046f-188">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="c046f-189">Строка​</span><span class="sxs-lookup"><span data-stu-id="c046f-189">String</span></span>|<span data-ttu-id="c046f-190">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="c046f-190">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="c046f-191">Строка​</span><span class="sxs-lookup"><span data-stu-id="c046f-191">String</span></span>|<span data-ttu-id="c046f-192">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="c046f-192">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="c046f-193">Требования</span><span class="sxs-lookup"><span data-stu-id="c046f-193">Requirements</span></span>

|<span data-ttu-id="c046f-194">Требование</span><span class="sxs-lookup"><span data-stu-id="c046f-194">Requirement</span></span>| <span data-ttu-id="c046f-195">Значение</span><span class="sxs-lookup"><span data-stu-id="c046f-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="c046f-196">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="c046f-196">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c046f-197">1.0</span><span class="sxs-lookup"><span data-stu-id="c046f-197">1.0</span></span>|
|[<span data-ttu-id="c046f-198">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c046f-198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c046f-199">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="c046f-199">Compose or read</span></span>|