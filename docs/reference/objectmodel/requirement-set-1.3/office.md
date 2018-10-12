 

# <a name="office"></a><span data-ttu-id="0b53d-101">Office</span><span class="sxs-lookup"><span data-stu-id="0b53d-101">Office</span></span>

<span data-ttu-id="0b53d-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="0b53d-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="0b53d-104">Требования</span><span class="sxs-lookup"><span data-stu-id="0b53d-104">Requirements</span></span>

|<span data-ttu-id="0b53d-105">Требование</span><span class="sxs-lookup"><span data-stu-id="0b53d-105">Requirement</span></span>| <span data-ttu-id="0b53d-106">Значение</span><span class="sxs-lookup"><span data-stu-id="0b53d-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b53d-107">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0b53d-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0b53d-108">1.0</span><span class="sxs-lookup"><span data-stu-id="0b53d-108">1.0</span></span>|
|[<span data-ttu-id="0b53d-109">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b53d-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0b53d-110">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0b53d-110">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="0b53d-111">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="0b53d-111">Namespaces</span></span>

<span data-ttu-id="0b53d-112">[context](office.context.md) — предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="0b53d-112">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="0b53d-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype) — включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="0b53d-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="0b53d-114">Члены</span><span class="sxs-lookup"><span data-stu-id="0b53d-114">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="0b53d-115">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="0b53d-115">AsyncResultStatus :String</span></span>

<span data-ttu-id="0b53d-116">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="0b53d-116">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="0b53d-117">Тип:</span><span class="sxs-lookup"><span data-stu-id="0b53d-117">Type:</span></span>

*   <span data-ttu-id="0b53d-118">String</span><span class="sxs-lookup"><span data-stu-id="0b53d-118">String</span></span>

##### <a name="properties"></a><span data-ttu-id="0b53d-119">Свойства:</span><span class="sxs-lookup"><span data-stu-id="0b53d-119">Properties:</span></span>

|<span data-ttu-id="0b53d-120">Имя</span><span class="sxs-lookup"><span data-stu-id="0b53d-120">Name</span></span>| <span data-ttu-id="0b53d-121">Тип</span><span class="sxs-lookup"><span data-stu-id="0b53d-121">Type</span></span>| <span data-ttu-id="0b53d-122">Описание</span><span class="sxs-lookup"><span data-stu-id="0b53d-122">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="0b53d-123">String</span><span class="sxs-lookup"><span data-stu-id="0b53d-123">String</span></span>|<span data-ttu-id="0b53d-124">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="0b53d-124">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="0b53d-125">String</span><span class="sxs-lookup"><span data-stu-id="0b53d-125">String</span></span>|<span data-ttu-id="0b53d-126">Вызов не удался.</span><span class="sxs-lookup"><span data-stu-id="0b53d-126">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0b53d-127">Требования</span><span class="sxs-lookup"><span data-stu-id="0b53d-127">Requirements</span></span>

|<span data-ttu-id="0b53d-128">Требование</span><span class="sxs-lookup"><span data-stu-id="0b53d-128">Requirement</span></span>| <span data-ttu-id="0b53d-129">Значение</span><span class="sxs-lookup"><span data-stu-id="0b53d-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b53d-130">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0b53d-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0b53d-131">1.0</span><span class="sxs-lookup"><span data-stu-id="0b53d-131">1.0</span></span>|
|[<span data-ttu-id="0b53d-132">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b53d-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0b53d-133">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0b53d-133">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="0b53d-134">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="0b53d-134">CoercionType :String</span></span>

<span data-ttu-id="0b53d-135">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="0b53d-135">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="0b53d-136">Тип:</span><span class="sxs-lookup"><span data-stu-id="0b53d-136">Type:</span></span>

*   <span data-ttu-id="0b53d-137">String</span><span class="sxs-lookup"><span data-stu-id="0b53d-137">String</span></span>

##### <a name="properties"></a><span data-ttu-id="0b53d-138">Свойства:</span><span class="sxs-lookup"><span data-stu-id="0b53d-138">Properties:</span></span>

|<span data-ttu-id="0b53d-139">Имя</span><span class="sxs-lookup"><span data-stu-id="0b53d-139">Name</span></span>| <span data-ttu-id="0b53d-140">Тип</span><span class="sxs-lookup"><span data-stu-id="0b53d-140">Type</span></span>| <span data-ttu-id="0b53d-141">Описание</span><span class="sxs-lookup"><span data-stu-id="0b53d-141">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="0b53d-142">String</span><span class="sxs-lookup"><span data-stu-id="0b53d-142">String</span></span>|<span data-ttu-id="0b53d-143">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="0b53d-143">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="0b53d-144">Строка</span><span class="sxs-lookup"><span data-stu-id="0b53d-144">String</span></span>|<span data-ttu-id="0b53d-145">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="0b53d-145">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0b53d-146">Требования</span><span class="sxs-lookup"><span data-stu-id="0b53d-146">Requirements</span></span>

|<span data-ttu-id="0b53d-147">Требование</span><span class="sxs-lookup"><span data-stu-id="0b53d-147">Requirement</span></span>| <span data-ttu-id="0b53d-148">Значение</span><span class="sxs-lookup"><span data-stu-id="0b53d-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b53d-149">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0b53d-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0b53d-150">1.0</span><span class="sxs-lookup"><span data-stu-id="0b53d-150">1.0</span></span>|
|[<span data-ttu-id="0b53d-151">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b53d-151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0b53d-152">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0b53d-152">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="0b53d-153">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="0b53d-153">SourceProperty :String</span></span>

<span data-ttu-id="0b53d-154">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="0b53d-154">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="0b53d-155">Тип:</span><span class="sxs-lookup"><span data-stu-id="0b53d-155">Type:</span></span>

*   <span data-ttu-id="0b53d-156">String</span><span class="sxs-lookup"><span data-stu-id="0b53d-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="0b53d-157">Свойства:</span><span class="sxs-lookup"><span data-stu-id="0b53d-157">Properties:</span></span>

|<span data-ttu-id="0b53d-158">Имя</span><span class="sxs-lookup"><span data-stu-id="0b53d-158">Name</span></span>| <span data-ttu-id="0b53d-159">Тип</span><span class="sxs-lookup"><span data-stu-id="0b53d-159">Type</span></span>| <span data-ttu-id="0b53d-160">Описание</span><span class="sxs-lookup"><span data-stu-id="0b53d-160">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="0b53d-161">Строка</span><span class="sxs-lookup"><span data-stu-id="0b53d-161">String</span></span>|<span data-ttu-id="0b53d-162">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="0b53d-162">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="0b53d-163">String</span><span class="sxs-lookup"><span data-stu-id="0b53d-163">String</span></span>|<span data-ttu-id="0b53d-164">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="0b53d-164">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0b53d-165">Требования</span><span class="sxs-lookup"><span data-stu-id="0b53d-165">Requirements</span></span>

|<span data-ttu-id="0b53d-166">Требование</span><span class="sxs-lookup"><span data-stu-id="0b53d-166">Requirement</span></span>| <span data-ttu-id="0b53d-167">Значение</span><span class="sxs-lookup"><span data-stu-id="0b53d-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="0b53d-168">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0b53d-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0b53d-169">1.0</span><span class="sxs-lookup"><span data-stu-id="0b53d-169">1.0</span></span>|
|[<span data-ttu-id="0b53d-170">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0b53d-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0b53d-171">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0b53d-171">Compose or read</span></span>|