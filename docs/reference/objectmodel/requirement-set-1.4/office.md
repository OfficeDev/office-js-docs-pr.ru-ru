 

# <a name="office"></a><span data-ttu-id="badf6-101">Office</span><span class="sxs-lookup"><span data-stu-id="badf6-101">Office</span></span>

<span data-ttu-id="badf6-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="badf6-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="badf6-104">Требования</span><span class="sxs-lookup"><span data-stu-id="badf6-104">Requirements</span></span>

|<span data-ttu-id="badf6-105">Требование</span><span class="sxs-lookup"><span data-stu-id="badf6-105">Requirement</span></span>| <span data-ttu-id="badf6-106">Значение</span><span class="sxs-lookup"><span data-stu-id="badf6-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="badf6-107">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="badf6-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="badf6-108">1.0</span><span class="sxs-lookup"><span data-stu-id="badf6-108">1.0</span></span>|
|[<span data-ttu-id="badf6-109">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="badf6-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="badf6-110">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="badf6-110">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="badf6-111">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="badf6-111">Namespaces</span></span>

<span data-ttu-id="badf6-112">[context](Office.context.md) — предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="badf6-112">[context](Office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="badf6-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype) — включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="badf6-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="badf6-114">Члены</span><span class="sxs-lookup"><span data-stu-id="badf6-114">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="badf6-115">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="badf6-115">AsyncResultStatus :String</span></span>

<span data-ttu-id="badf6-116">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="badf6-116">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="badf6-117">Тип:</span><span class="sxs-lookup"><span data-stu-id="badf6-117">Type:</span></span>

*   <span data-ttu-id="badf6-118">String</span><span class="sxs-lookup"><span data-stu-id="badf6-118">String</span></span>

##### <a name="properties"></a><span data-ttu-id="badf6-119">Свойства:</span><span class="sxs-lookup"><span data-stu-id="badf6-119">Properties:</span></span>

|<span data-ttu-id="badf6-120">Имя</span><span class="sxs-lookup"><span data-stu-id="badf6-120">Name</span></span>| <span data-ttu-id="badf6-121">Тип</span><span class="sxs-lookup"><span data-stu-id="badf6-121">Type</span></span>| <span data-ttu-id="badf6-122">Описание</span><span class="sxs-lookup"><span data-stu-id="badf6-122">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="badf6-123">String</span><span class="sxs-lookup"><span data-stu-id="badf6-123">String</span></span>|<span data-ttu-id="badf6-124">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="badf6-124">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="badf6-125">String</span><span class="sxs-lookup"><span data-stu-id="badf6-125">String</span></span>|<span data-ttu-id="badf6-126">Вызов не удался.</span><span class="sxs-lookup"><span data-stu-id="badf6-126">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="badf6-127">Требования</span><span class="sxs-lookup"><span data-stu-id="badf6-127">Requirements</span></span>

|<span data-ttu-id="badf6-128">Требование</span><span class="sxs-lookup"><span data-stu-id="badf6-128">Requirement</span></span>| <span data-ttu-id="badf6-129">Значение</span><span class="sxs-lookup"><span data-stu-id="badf6-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="badf6-130">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="badf6-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="badf6-131">1.0</span><span class="sxs-lookup"><span data-stu-id="badf6-131">1.0</span></span>|
|[<span data-ttu-id="badf6-132">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="badf6-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="badf6-133">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="badf6-133">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="badf6-134">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="badf6-134">CoercionType :String</span></span>

<span data-ttu-id="badf6-135">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="badf6-135">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="badf6-136">Тип:</span><span class="sxs-lookup"><span data-stu-id="badf6-136">Type:</span></span>

*   <span data-ttu-id="badf6-137">String</span><span class="sxs-lookup"><span data-stu-id="badf6-137">String</span></span>

##### <a name="properties"></a><span data-ttu-id="badf6-138">Свойства:</span><span class="sxs-lookup"><span data-stu-id="badf6-138">Properties:</span></span>

|<span data-ttu-id="badf6-139">Имя</span><span class="sxs-lookup"><span data-stu-id="badf6-139">Name</span></span>| <span data-ttu-id="badf6-140">Тип</span><span class="sxs-lookup"><span data-stu-id="badf6-140">Type</span></span>| <span data-ttu-id="badf6-141">Описание</span><span class="sxs-lookup"><span data-stu-id="badf6-141">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="badf6-142">String</span><span class="sxs-lookup"><span data-stu-id="badf6-142">String</span></span>|<span data-ttu-id="badf6-143">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="badf6-143">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="badf6-144">String</span><span class="sxs-lookup"><span data-stu-id="badf6-144">String</span></span>|<span data-ttu-id="badf6-145">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="badf6-145">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="badf6-146">Требования</span><span class="sxs-lookup"><span data-stu-id="badf6-146">Requirements</span></span>

|<span data-ttu-id="badf6-147">Требование</span><span class="sxs-lookup"><span data-stu-id="badf6-147">Requirement</span></span>| <span data-ttu-id="badf6-148">Значение</span><span class="sxs-lookup"><span data-stu-id="badf6-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="badf6-149">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="badf6-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="badf6-150">1.0</span><span class="sxs-lookup"><span data-stu-id="badf6-150">1.0</span></span>|
|[<span data-ttu-id="badf6-151">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="badf6-151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="badf6-152">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="badf6-152">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="badf6-153">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="badf6-153">SourceProperty :String</span></span>

<span data-ttu-id="badf6-154">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="badf6-154">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="badf6-155">Тип:</span><span class="sxs-lookup"><span data-stu-id="badf6-155">Type:</span></span>

*   <span data-ttu-id="badf6-156">String</span><span class="sxs-lookup"><span data-stu-id="badf6-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="badf6-157">Свойства:</span><span class="sxs-lookup"><span data-stu-id="badf6-157">Properties:</span></span>

|<span data-ttu-id="badf6-158">Имя</span><span class="sxs-lookup"><span data-stu-id="badf6-158">Name</span></span>| <span data-ttu-id="badf6-159">Тип</span><span class="sxs-lookup"><span data-stu-id="badf6-159">Type</span></span>| <span data-ttu-id="badf6-160">Описание</span><span class="sxs-lookup"><span data-stu-id="badf6-160">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="badf6-161">String</span><span class="sxs-lookup"><span data-stu-id="badf6-161">String</span></span>|<span data-ttu-id="badf6-162">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="badf6-162">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="badf6-163">String</span><span class="sxs-lookup"><span data-stu-id="badf6-163">String</span></span>|<span data-ttu-id="badf6-164">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="badf6-164">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="badf6-165">Требования</span><span class="sxs-lookup"><span data-stu-id="badf6-165">Requirements</span></span>

|<span data-ttu-id="badf6-166">Требование</span><span class="sxs-lookup"><span data-stu-id="badf6-166">Requirement</span></span>| <span data-ttu-id="badf6-167">Значение</span><span class="sxs-lookup"><span data-stu-id="badf6-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="badf6-168">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="badf6-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="badf6-169">1.0</span><span class="sxs-lookup"><span data-stu-id="badf6-169">1.0</span></span>|
|[<span data-ttu-id="badf6-170">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="badf6-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="badf6-171">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="badf6-171">Compose or read</span></span>|