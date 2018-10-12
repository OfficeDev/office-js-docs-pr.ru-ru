 

# <a name="office"></a><span data-ttu-id="d7391-101">Office</span><span class="sxs-lookup"><span data-stu-id="d7391-101">Office</span></span>

<span data-ttu-id="d7391-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="d7391-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="d7391-104">Требования</span><span class="sxs-lookup"><span data-stu-id="d7391-104">Requirements</span></span>

|<span data-ttu-id="d7391-105">Требование</span><span class="sxs-lookup"><span data-stu-id="d7391-105">Requirement</span></span>| <span data-ttu-id="d7391-106">Значение</span><span class="sxs-lookup"><span data-stu-id="d7391-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7391-107">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d7391-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7391-108">1.0</span><span class="sxs-lookup"><span data-stu-id="d7391-108">1.0</span></span>|
|[<span data-ttu-id="d7391-109">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d7391-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d7391-110">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d7391-110">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="d7391-111">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="d7391-111">Namespaces</span></span>

<span data-ttu-id="d7391-112">[context](office.context.md) — предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="d7391-112">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="d7391-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype) — включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="d7391-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="d7391-114">Члены</span><span class="sxs-lookup"><span data-stu-id="d7391-114">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="d7391-115">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="d7391-115">AsyncResultStatus :String</span></span>

<span data-ttu-id="d7391-116">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="d7391-116">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="d7391-117">Тип:</span><span class="sxs-lookup"><span data-stu-id="d7391-117">Type:</span></span>

*   <span data-ttu-id="d7391-118">String</span><span class="sxs-lookup"><span data-stu-id="d7391-118">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d7391-119">Свойства:</span><span class="sxs-lookup"><span data-stu-id="d7391-119">Properties:</span></span>

|<span data-ttu-id="d7391-120">Имя</span><span class="sxs-lookup"><span data-stu-id="d7391-120">Name</span></span>| <span data-ttu-id="d7391-121">Тип</span><span class="sxs-lookup"><span data-stu-id="d7391-121">Type</span></span>| <span data-ttu-id="d7391-122">Описание</span><span class="sxs-lookup"><span data-stu-id="d7391-122">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="d7391-123">String</span><span class="sxs-lookup"><span data-stu-id="d7391-123">String</span></span>|<span data-ttu-id="d7391-124">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="d7391-124">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="d7391-125">String</span><span class="sxs-lookup"><span data-stu-id="d7391-125">String</span></span>|<span data-ttu-id="d7391-126">Вызов не удался.</span><span class="sxs-lookup"><span data-stu-id="d7391-126">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7391-127">Требования</span><span class="sxs-lookup"><span data-stu-id="d7391-127">Requirements</span></span>

|<span data-ttu-id="d7391-128">Требование</span><span class="sxs-lookup"><span data-stu-id="d7391-128">Requirement</span></span>| <span data-ttu-id="d7391-129">Значение</span><span class="sxs-lookup"><span data-stu-id="d7391-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7391-130">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d7391-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7391-131">1.0</span><span class="sxs-lookup"><span data-stu-id="d7391-131">1.0</span></span>|
|[<span data-ttu-id="d7391-132">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d7391-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d7391-133">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d7391-133">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="d7391-134">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="d7391-134">CoercionType :String</span></span>

<span data-ttu-id="d7391-135">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="d7391-135">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="d7391-136">Тип:</span><span class="sxs-lookup"><span data-stu-id="d7391-136">Type:</span></span>

*   <span data-ttu-id="d7391-137">String</span><span class="sxs-lookup"><span data-stu-id="d7391-137">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d7391-138">Свойства:</span><span class="sxs-lookup"><span data-stu-id="d7391-138">Properties:</span></span>

|<span data-ttu-id="d7391-139">Имя</span><span class="sxs-lookup"><span data-stu-id="d7391-139">Name</span></span>| <span data-ttu-id="d7391-140">Тип</span><span class="sxs-lookup"><span data-stu-id="d7391-140">Type</span></span>| <span data-ttu-id="d7391-141">Описание</span><span class="sxs-lookup"><span data-stu-id="d7391-141">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="d7391-142">String</span><span class="sxs-lookup"><span data-stu-id="d7391-142">String</span></span>|<span data-ttu-id="d7391-143">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="d7391-143">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="d7391-144">String</span><span class="sxs-lookup"><span data-stu-id="d7391-144">String</span></span>|<span data-ttu-id="d7391-145">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="d7391-145">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7391-146">Требования</span><span class="sxs-lookup"><span data-stu-id="d7391-146">Requirements</span></span>

|<span data-ttu-id="d7391-147">Требование</span><span class="sxs-lookup"><span data-stu-id="d7391-147">Requirement</span></span>| <span data-ttu-id="d7391-148">Значение</span><span class="sxs-lookup"><span data-stu-id="d7391-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7391-149">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d7391-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7391-150">1.0</span><span class="sxs-lookup"><span data-stu-id="d7391-150">1.0</span></span>|
|[<span data-ttu-id="d7391-151">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d7391-151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d7391-152">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d7391-152">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="d7391-153">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="d7391-153">SourceProperty :String</span></span>

<span data-ttu-id="d7391-154">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="d7391-154">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="d7391-155">Тип:</span><span class="sxs-lookup"><span data-stu-id="d7391-155">Type:</span></span>

*   <span data-ttu-id="d7391-156">String</span><span class="sxs-lookup"><span data-stu-id="d7391-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="d7391-157">Свойства:</span><span class="sxs-lookup"><span data-stu-id="d7391-157">Properties:</span></span>

|<span data-ttu-id="d7391-158">Имя</span><span class="sxs-lookup"><span data-stu-id="d7391-158">Name</span></span>| <span data-ttu-id="d7391-159">Тип</span><span class="sxs-lookup"><span data-stu-id="d7391-159">Type</span></span>| <span data-ttu-id="d7391-160">Описание</span><span class="sxs-lookup"><span data-stu-id="d7391-160">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="d7391-161">String</span><span class="sxs-lookup"><span data-stu-id="d7391-161">String</span></span>|<span data-ttu-id="d7391-162">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="d7391-162">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="d7391-163">String</span><span class="sxs-lookup"><span data-stu-id="d7391-163">String</span></span>|<span data-ttu-id="d7391-164">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="d7391-164">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d7391-165">Требования</span><span class="sxs-lookup"><span data-stu-id="d7391-165">Requirements</span></span>

|<span data-ttu-id="d7391-166">Требование</span><span class="sxs-lookup"><span data-stu-id="d7391-166">Requirement</span></span>| <span data-ttu-id="d7391-167">Значение</span><span class="sxs-lookup"><span data-stu-id="d7391-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="d7391-168">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="d7391-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d7391-169">1.0</span><span class="sxs-lookup"><span data-stu-id="d7391-169">1.0</span></span>|
|[<span data-ttu-id="d7391-170">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d7391-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="d7391-171">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d7391-171">Compose or read</span></span>|