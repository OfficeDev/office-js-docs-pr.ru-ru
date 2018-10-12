 

# <a name="office"></a><span data-ttu-id="04927-101">Office</span><span class="sxs-lookup"><span data-stu-id="04927-101">Office</span></span>

<span data-ttu-id="04927-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="04927-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="04927-104">Требования</span><span class="sxs-lookup"><span data-stu-id="04927-104">Requirements</span></span>

|<span data-ttu-id="04927-105">Требование</span><span class="sxs-lookup"><span data-stu-id="04927-105">Requirement</span></span>| <span data-ttu-id="04927-106">Значение</span><span class="sxs-lookup"><span data-stu-id="04927-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="04927-107">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="04927-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04927-108">1.0</span><span class="sxs-lookup"><span data-stu-id="04927-108">1.0</span></span>|
|[<span data-ttu-id="04927-109">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="04927-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="04927-110">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="04927-110">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="04927-111">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="04927-111">Namespaces</span></span>

<span data-ttu-id="04927-112">[context](office.context.md) — предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="04927-112">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="04927-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype) — включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="04927-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="04927-114">Члены</span><span class="sxs-lookup"><span data-stu-id="04927-114">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="04927-115">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="04927-115">AsyncResultStatus :String</span></span>

<span data-ttu-id="04927-116">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="04927-116">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="04927-117">Тип:</span><span class="sxs-lookup"><span data-stu-id="04927-117">Type:</span></span>

*   <span data-ttu-id="04927-118">String</span><span class="sxs-lookup"><span data-stu-id="04927-118">String</span></span>

##### <a name="properties"></a><span data-ttu-id="04927-119">Свойства:</span><span class="sxs-lookup"><span data-stu-id="04927-119">Properties:</span></span>

|<span data-ttu-id="04927-120">Name</span><span class="sxs-lookup"><span data-stu-id="04927-120">Name</span></span>| <span data-ttu-id="04927-121">Тип</span><span class="sxs-lookup"><span data-stu-id="04927-121">Type</span></span>| <span data-ttu-id="04927-122">Описание</span><span class="sxs-lookup"><span data-stu-id="04927-122">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="04927-123">String</span><span class="sxs-lookup"><span data-stu-id="04927-123">String</span></span>|<span data-ttu-id="04927-124">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="04927-124">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="04927-125">String</span><span class="sxs-lookup"><span data-stu-id="04927-125">String</span></span>|<span data-ttu-id="04927-126">Вызов не удался.</span><span class="sxs-lookup"><span data-stu-id="04927-126">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="04927-127">Требования</span><span class="sxs-lookup"><span data-stu-id="04927-127">Requirements</span></span>

|<span data-ttu-id="04927-128">Требование</span><span class="sxs-lookup"><span data-stu-id="04927-128">Requirement</span></span>| <span data-ttu-id="04927-129">Значение</span><span class="sxs-lookup"><span data-stu-id="04927-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="04927-130">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="04927-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04927-131">1.0</span><span class="sxs-lookup"><span data-stu-id="04927-131">1.0</span></span>|
|[<span data-ttu-id="04927-132">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="04927-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="04927-133">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="04927-133">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="04927-134">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="04927-134">CoercionType :String</span></span>

<span data-ttu-id="04927-135">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="04927-135">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="04927-136">Тип:</span><span class="sxs-lookup"><span data-stu-id="04927-136">Type:</span></span>

*   <span data-ttu-id="04927-137">String</span><span class="sxs-lookup"><span data-stu-id="04927-137">String</span></span>

##### <a name="properties"></a><span data-ttu-id="04927-138">Свойства:</span><span class="sxs-lookup"><span data-stu-id="04927-138">Properties:</span></span>

|<span data-ttu-id="04927-139">Name</span><span class="sxs-lookup"><span data-stu-id="04927-139">Name</span></span>| <span data-ttu-id="04927-140">Тип</span><span class="sxs-lookup"><span data-stu-id="04927-140">Type</span></span>| <span data-ttu-id="04927-141">Описание</span><span class="sxs-lookup"><span data-stu-id="04927-141">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="04927-142">String</span><span class="sxs-lookup"><span data-stu-id="04927-142">String</span></span>|<span data-ttu-id="04927-143">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="04927-143">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="04927-144">String</span><span class="sxs-lookup"><span data-stu-id="04927-144">String</span></span>|<span data-ttu-id="04927-145">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="04927-145">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="04927-146">Требования</span><span class="sxs-lookup"><span data-stu-id="04927-146">Requirements</span></span>

|<span data-ttu-id="04927-147">Требование</span><span class="sxs-lookup"><span data-stu-id="04927-147">Requirement</span></span>| <span data-ttu-id="04927-148">Значение</span><span class="sxs-lookup"><span data-stu-id="04927-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="04927-149">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="04927-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04927-150">1.0</span><span class="sxs-lookup"><span data-stu-id="04927-150">1.0</span></span>|
|[<span data-ttu-id="04927-151">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="04927-151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="04927-152">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="04927-152">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="04927-153">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="04927-153">SourceProperty :String</span></span>

<span data-ttu-id="04927-154">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="04927-154">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="04927-155">Тип:</span><span class="sxs-lookup"><span data-stu-id="04927-155">Type:</span></span>

*   <span data-ttu-id="04927-156">String</span><span class="sxs-lookup"><span data-stu-id="04927-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="04927-157">Свойства:</span><span class="sxs-lookup"><span data-stu-id="04927-157">Properties:</span></span>

|<span data-ttu-id="04927-158">Name</span><span class="sxs-lookup"><span data-stu-id="04927-158">Name</span></span>| <span data-ttu-id="04927-159">Тип</span><span class="sxs-lookup"><span data-stu-id="04927-159">Type</span></span>| <span data-ttu-id="04927-160">Описание</span><span class="sxs-lookup"><span data-stu-id="04927-160">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="04927-161">String</span><span class="sxs-lookup"><span data-stu-id="04927-161">String</span></span>|<span data-ttu-id="04927-162">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="04927-162">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="04927-163">String</span><span class="sxs-lookup"><span data-stu-id="04927-163">String</span></span>|<span data-ttu-id="04927-164">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="04927-164">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="04927-165">Требования</span><span class="sxs-lookup"><span data-stu-id="04927-165">Requirements</span></span>

|<span data-ttu-id="04927-166">Требование</span><span class="sxs-lookup"><span data-stu-id="04927-166">Requirement</span></span>| <span data-ttu-id="04927-167">Значение</span><span class="sxs-lookup"><span data-stu-id="04927-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="04927-168">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="04927-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="04927-169">1.0</span><span class="sxs-lookup"><span data-stu-id="04927-169">1.0</span></span>|
|[<span data-ttu-id="04927-170">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="04927-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="04927-171">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="04927-171">Compose or read</span></span>|