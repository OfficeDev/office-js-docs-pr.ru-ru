# <a name="office"></a><span data-ttu-id="1261e-101">Office</span><span class="sxs-lookup"><span data-stu-id="1261e-101">Office</span></span>

<span data-ttu-id="1261e-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="1261e-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="1261e-104">Требования</span><span class="sxs-lookup"><span data-stu-id="1261e-104">Requirements</span></span>

|<span data-ttu-id="1261e-105">Requirement</span><span class="sxs-lookup"><span data-stu-id="1261e-105">Requirement</span></span>| <span data-ttu-id="1261e-106">Значение</span><span class="sxs-lookup"><span data-stu-id="1261e-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="1261e-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1261e-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1261e-108">1.0</span><span class="sxs-lookup"><span data-stu-id="1261e-108">1.0</span></span>|
|[<span data-ttu-id="1261e-109">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1261e-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1261e-110">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1261e-110">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="1261e-111">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="1261e-111">Members and methods</span></span>

| <span data-ttu-id="1261e-112">Член</span><span class="sxs-lookup"><span data-stu-id="1261e-112">Member</span></span> | <span data-ttu-id="1261e-113">Тип</span><span class="sxs-lookup"><span data-stu-id="1261e-113">Type</span></span> |
|--------|------|
| [<span data-ttu-id="1261e-114">AsyncResultStatus</span><span class="sxs-lookup"><span data-stu-id="1261e-114">AsyncResultStatus</span></span>](#asyncresultstatus-string) | <span data-ttu-id="1261e-115">Член</span><span class="sxs-lookup"><span data-stu-id="1261e-115">Member</span></span> |
| [<span data-ttu-id="1261e-116">CoercionType</span><span class="sxs-lookup"><span data-stu-id="1261e-116">CoercionType</span></span>](#coerciontype-string) | <span data-ttu-id="1261e-117">Член</span><span class="sxs-lookup"><span data-stu-id="1261e-117">Member</span></span> |
| [<span data-ttu-id="1261e-118">EventType</span><span class="sxs-lookup"><span data-stu-id="1261e-118">EventType</span></span>](#eventtype-string) | <span data-ttu-id="1261e-119">Член</span><span class="sxs-lookup"><span data-stu-id="1261e-119">Member</span></span> |
| [<span data-ttu-id="1261e-120">SourceProperty</span><span class="sxs-lookup"><span data-stu-id="1261e-120">SourceProperty</span></span>](#sourceproperty-string) | <span data-ttu-id="1261e-121">Член</span><span class="sxs-lookup"><span data-stu-id="1261e-121">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="1261e-122">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="1261e-122">Namespaces</span></span>

<span data-ttu-id="1261e-123">[context](office.context.md). Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="1261e-123">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="1261e-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype). Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="1261e-124">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="1261e-125">Элементы</span><span class="sxs-lookup"><span data-stu-id="1261e-125">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="1261e-126">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="1261e-126">AsyncResultStatus :String</span></span>

<span data-ttu-id="1261e-127">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="1261e-127">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="1261e-128">Тип:</span><span class="sxs-lookup"><span data-stu-id="1261e-128">Type:</span></span>

*   <span data-ttu-id="1261e-129">String</span><span class="sxs-lookup"><span data-stu-id="1261e-129">String</span></span>

##### <a name="properties"></a><span data-ttu-id="1261e-130">Свойства:</span><span class="sxs-lookup"><span data-stu-id="1261e-130">Properties:</span></span>

|<span data-ttu-id="1261e-131">Имя</span><span class="sxs-lookup"><span data-stu-id="1261e-131">Name</span></span>| <span data-ttu-id="1261e-132">Тип</span><span class="sxs-lookup"><span data-stu-id="1261e-132">Type</span></span>| <span data-ttu-id="1261e-133">Описание</span><span class="sxs-lookup"><span data-stu-id="1261e-133">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="1261e-134">Для указания</span><span class="sxs-lookup"><span data-stu-id="1261e-134">String</span></span>|<span data-ttu-id="1261e-135">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="1261e-135">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="1261e-136">Для указания</span><span class="sxs-lookup"><span data-stu-id="1261e-136">String</span></span>|<span data-ttu-id="1261e-137">Вызов завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="1261e-137">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1261e-138">Требования</span><span class="sxs-lookup"><span data-stu-id="1261e-138">Requirements</span></span>

|<span data-ttu-id="1261e-139">Requirement</span><span class="sxs-lookup"><span data-stu-id="1261e-139">Requirement</span></span>| <span data-ttu-id="1261e-140">Значение</span><span class="sxs-lookup"><span data-stu-id="1261e-140">Value</span></span>|
|---|---|
|[<span data-ttu-id="1261e-141">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1261e-141">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1261e-142">1.0</span><span class="sxs-lookup"><span data-stu-id="1261e-142">1.0</span></span>|
|[<span data-ttu-id="1261e-143">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1261e-143">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1261e-144">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1261e-144">Compose or read</span></span>|

---

####  <a name="coerciontype-string"></a><span data-ttu-id="1261e-145">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="1261e-145">CoercionType :String</span></span>

<span data-ttu-id="1261e-146">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="1261e-146">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="1261e-147">Тип:</span><span class="sxs-lookup"><span data-stu-id="1261e-147">Type:</span></span>

*   <span data-ttu-id="1261e-148">String</span><span class="sxs-lookup"><span data-stu-id="1261e-148">String</span></span>

##### <a name="properties"></a><span data-ttu-id="1261e-149">Свойства:</span><span class="sxs-lookup"><span data-stu-id="1261e-149">Properties:</span></span>

|<span data-ttu-id="1261e-150">Имя</span><span class="sxs-lookup"><span data-stu-id="1261e-150">Name</span></span>| <span data-ttu-id="1261e-151">Тип</span><span class="sxs-lookup"><span data-stu-id="1261e-151">Type</span></span>| <span data-ttu-id="1261e-152">Описание</span><span class="sxs-lookup"><span data-stu-id="1261e-152">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="1261e-153">String</span><span class="sxs-lookup"><span data-stu-id="1261e-153">String</span></span>|<span data-ttu-id="1261e-154">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="1261e-154">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="1261e-155">String</span><span class="sxs-lookup"><span data-stu-id="1261e-155">String</span></span>|<span data-ttu-id="1261e-156">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="1261e-156">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1261e-157">Требования</span><span class="sxs-lookup"><span data-stu-id="1261e-157">Requirements</span></span>

|<span data-ttu-id="1261e-158">Requirement</span><span class="sxs-lookup"><span data-stu-id="1261e-158">Requirement</span></span>| <span data-ttu-id="1261e-159">Значение</span><span class="sxs-lookup"><span data-stu-id="1261e-159">Value</span></span>|
|---|---|
|[<span data-ttu-id="1261e-160">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1261e-160">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1261e-161">1.0</span><span class="sxs-lookup"><span data-stu-id="1261e-161">1.0</span></span>|
|[<span data-ttu-id="1261e-162">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1261e-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1261e-163">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1261e-163">Compose or read</span></span>|

---

####  <a name="eventtype-string"></a><span data-ttu-id="1261e-164">EventType :String</span><span class="sxs-lookup"><span data-stu-id="1261e-164">EventType :String</span></span>

<span data-ttu-id="1261e-165">Указывает событие, связанное с обработчиком.</span><span class="sxs-lookup"><span data-stu-id="1261e-165">Specifies the event associated with an event handler.</span></span>

##### <a name="type"></a><span data-ttu-id="1261e-166">Тип:</span><span class="sxs-lookup"><span data-stu-id="1261e-166">Type:</span></span>

*   <span data-ttu-id="1261e-167">String</span><span class="sxs-lookup"><span data-stu-id="1261e-167">String</span></span>

##### <a name="properties"></a><span data-ttu-id="1261e-168">Свойства:</span><span class="sxs-lookup"><span data-stu-id="1261e-168">Properties:</span></span>

| <span data-ttu-id="1261e-169">Имя</span><span class="sxs-lookup"><span data-stu-id="1261e-169">Name</span></span> | <span data-ttu-id="1261e-170">Тип</span><span class="sxs-lookup"><span data-stu-id="1261e-170">Type</span></span> | <span data-ttu-id="1261e-171">Описание</span><span class="sxs-lookup"><span data-stu-id="1261e-171">Description</span></span> |
|---|---|---|
|`ItemChanged`| <span data-ttu-id="1261e-172">Строка</span><span class="sxs-lookup"><span data-stu-id="1261e-172">String</span></span> | <span data-ttu-id="1261e-173">Пока область задач закреплена, для просмотра выбран другой элемент Outlook.</span><span class="sxs-lookup"><span data-stu-id="1261e-173">A different Outlook item is selected for viewing while the task pane is pinned.</span></span> |

##### <a name="requirements"></a><span data-ttu-id="1261e-174">Требования</span><span class="sxs-lookup"><span data-stu-id="1261e-174">Requirements</span></span>

|<span data-ttu-id="1261e-175">Requirement</span><span class="sxs-lookup"><span data-stu-id="1261e-175">Requirement</span></span>| <span data-ttu-id="1261e-176">Значение</span><span class="sxs-lookup"><span data-stu-id="1261e-176">Value</span></span>|
|---|---|
|[<span data-ttu-id="1261e-177">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="1261e-177">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1261e-178">1.5</span><span class="sxs-lookup"><span data-stu-id="1261e-178">1.5</span></span> |
|[<span data-ttu-id="1261e-179">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1261e-179">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1261e-180">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1261e-180">Compose or read</span></span> |

---

####  <a name="sourceproperty-string"></a><span data-ttu-id="1261e-181">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="1261e-181">SourceProperty :String</span></span>

<span data-ttu-id="1261e-182">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="1261e-182">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="1261e-183">Тип:</span><span class="sxs-lookup"><span data-stu-id="1261e-183">Type:</span></span>

*   <span data-ttu-id="1261e-184">String</span><span class="sxs-lookup"><span data-stu-id="1261e-184">String</span></span>

##### <a name="properties"></a><span data-ttu-id="1261e-185">Свойства:</span><span class="sxs-lookup"><span data-stu-id="1261e-185">Properties:</span></span>

|<span data-ttu-id="1261e-186">Имя</span><span class="sxs-lookup"><span data-stu-id="1261e-186">Name</span></span>| <span data-ttu-id="1261e-187">Тип</span><span class="sxs-lookup"><span data-stu-id="1261e-187">Type</span></span>| <span data-ttu-id="1261e-188">Описание</span><span class="sxs-lookup"><span data-stu-id="1261e-188">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="1261e-189">String</span><span class="sxs-lookup"><span data-stu-id="1261e-189">String</span></span>|<span data-ttu-id="1261e-190">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="1261e-190">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="1261e-191">String</span><span class="sxs-lookup"><span data-stu-id="1261e-191">String</span></span>|<span data-ttu-id="1261e-192">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="1261e-192">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="1261e-193">Требования</span><span class="sxs-lookup"><span data-stu-id="1261e-193">Requirements</span></span>

|<span data-ttu-id="1261e-194">Requirement</span><span class="sxs-lookup"><span data-stu-id="1261e-194">Requirement</span></span>| <span data-ttu-id="1261e-195">Значение</span><span class="sxs-lookup"><span data-stu-id="1261e-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="1261e-196">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1261e-196">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="1261e-197">1.0</span><span class="sxs-lookup"><span data-stu-id="1261e-197">1.0</span></span>|
|[<span data-ttu-id="1261e-198">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1261e-198">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="1261e-199">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1261e-199">Compose or read</span></span>|