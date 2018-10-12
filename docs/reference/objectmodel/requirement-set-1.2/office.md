 

# <a name="office"></a><span data-ttu-id="f2833-101">Office</span><span class="sxs-lookup"><span data-stu-id="f2833-101">Office</span></span>

<span data-ttu-id="f2833-p101">Пространство имен Office содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office см. в статье [Общий API](/javascript/api/office).</span><span class="sxs-lookup"><span data-stu-id="f2833-p101">The Office namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office namespace, see the [Shared API](/javascript/api/office).</span></span>

##### <a name="requirements"></a><span data-ttu-id="f2833-104">Обязательные элементы</span><span class="sxs-lookup"><span data-stu-id="f2833-104">Requirements</span></span>

|<span data-ttu-id="f2833-105">Обязательный элемент</span><span class="sxs-lookup"><span data-stu-id="f2833-105">Requirement</span></span>| <span data-ttu-id="f2833-106">Значение</span><span class="sxs-lookup"><span data-stu-id="f2833-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="f2833-107">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="f2833-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f2833-108">1.0</span><span class="sxs-lookup"><span data-stu-id="f2833-108">1.0</span></span>|
|[<span data-ttu-id="f2833-109">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f2833-109">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f2833-110">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="f2833-110">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="f2833-111">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="f2833-111">Namespaces</span></span>

<span data-ttu-id="f2833-112">[context](office.context.md). Предоставляет общие интерфейсы из контекстного пространства имен API надстроек Office для использования в API надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="f2833-112">[context](office.context.md): Provides shared interfaces from the Office add-ins API's context namespace for use in the Outlook add-in API.</span></span>

<span data-ttu-id="f2833-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype). Включает перечисления ItemType, EntityType, AttachmentType, RecipientType, ResponseType и ItemNotificationMessageType.</span><span class="sxs-lookup"><span data-stu-id="f2833-113">[MailboxEnums](/javascript/api/outlook/office.mailboxenums.attachmenttype): Includes the ItemType, EntityType, AttachmentType, RecipientType, ResponseType, and ItemNotificationMessageType enumerations.</span></span>

### <a name="members"></a><span data-ttu-id="f2833-114">Члены</span><span class="sxs-lookup"><span data-stu-id="f2833-114">Members</span></span>

####  <a name="asyncresultstatus-string"></a><span data-ttu-id="f2833-115">AsyncResultStatus :String</span><span class="sxs-lookup"><span data-stu-id="f2833-115">AsyncResultStatus :String</span></span>

<span data-ttu-id="f2833-116">Указывает результат асинхронного вызова.</span><span class="sxs-lookup"><span data-stu-id="f2833-116">Specifies the result of an asynchronous call.</span></span>

##### <a name="type"></a><span data-ttu-id="f2833-117">Тип:</span><span class="sxs-lookup"><span data-stu-id="f2833-117">Type:</span></span>

*   <span data-ttu-id="f2833-118">String</span><span class="sxs-lookup"><span data-stu-id="f2833-118">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f2833-119">Свойства:</span><span class="sxs-lookup"><span data-stu-id="f2833-119">Properties:</span></span>

|<span data-ttu-id="f2833-120">Имя</span><span class="sxs-lookup"><span data-stu-id="f2833-120">Name</span></span>| <span data-ttu-id="f2833-121">Тип</span><span class="sxs-lookup"><span data-stu-id="f2833-121">Type</span></span>| <span data-ttu-id="f2833-122">Описание</span><span class="sxs-lookup"><span data-stu-id="f2833-122">Description</span></span>|
|---|---|---|
|`Succeeded`| <span data-ttu-id="f2833-123">String</span><span class="sxs-lookup"><span data-stu-id="f2833-123">String</span></span>|<span data-ttu-id="f2833-124">Вызов завершился успешно.</span><span class="sxs-lookup"><span data-stu-id="f2833-124">The call succeeded.</span></span>|
|`Failed`| <span data-ttu-id="f2833-125">String</span><span class="sxs-lookup"><span data-stu-id="f2833-125">String</span></span>|<span data-ttu-id="f2833-126">Вызов не удался.</span><span class="sxs-lookup"><span data-stu-id="f2833-126">The call failed.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f2833-127">Обязательные элементы</span><span class="sxs-lookup"><span data-stu-id="f2833-127">Requirements</span></span>

|<span data-ttu-id="f2833-128">Требование</span><span class="sxs-lookup"><span data-stu-id="f2833-128">Requirement</span></span>| <span data-ttu-id="f2833-129">Значение</span><span class="sxs-lookup"><span data-stu-id="f2833-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="f2833-130">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="f2833-130">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f2833-131">1.0</span><span class="sxs-lookup"><span data-stu-id="f2833-131">1.0</span></span>|
|[<span data-ttu-id="f2833-132">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f2833-132">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f2833-133">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="f2833-133">Compose or read</span></span>|
####  <a name="coerciontype-string"></a><span data-ttu-id="f2833-134">CoercionType :String</span><span class="sxs-lookup"><span data-stu-id="f2833-134">CoercionType :String</span></span>

<span data-ttu-id="f2833-135">Указывает способ приведения данных, возвращаемых или задаваемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="f2833-135">Specifies how to coerce data returned or set by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f2833-136">Тип:</span><span class="sxs-lookup"><span data-stu-id="f2833-136">Type:</span></span>

*   <span data-ttu-id="f2833-137">String</span><span class="sxs-lookup"><span data-stu-id="f2833-137">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f2833-138">Свойства:</span><span class="sxs-lookup"><span data-stu-id="f2833-138">Properties:</span></span>

|<span data-ttu-id="f2833-139">Имя</span><span class="sxs-lookup"><span data-stu-id="f2833-139">Name</span></span>| <span data-ttu-id="f2833-140">Тип</span><span class="sxs-lookup"><span data-stu-id="f2833-140">Type</span></span>| <span data-ttu-id="f2833-141">Описание</span><span class="sxs-lookup"><span data-stu-id="f2833-141">Description</span></span>|
|---|---|---|
|`Html`| <span data-ttu-id="f2833-142">String</span><span class="sxs-lookup"><span data-stu-id="f2833-142">String</span></span>|<span data-ttu-id="f2833-143">Запрашивает возврат данных в формате HTML.</span><span class="sxs-lookup"><span data-stu-id="f2833-143">Requests the data be returned in HTML format.</span></span>|
|`Text`| <span data-ttu-id="f2833-144">String</span><span class="sxs-lookup"><span data-stu-id="f2833-144">String</span></span>|<span data-ttu-id="f2833-145">Запрашивает возврат данных в формате текста.</span><span class="sxs-lookup"><span data-stu-id="f2833-145">Requests the data be returned in text format.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f2833-146">Обязательные элементы</span><span class="sxs-lookup"><span data-stu-id="f2833-146">Requirements</span></span>

|<span data-ttu-id="f2833-147">Требование</span><span class="sxs-lookup"><span data-stu-id="f2833-147">Requirement</span></span>| <span data-ttu-id="f2833-148">Значение</span><span class="sxs-lookup"><span data-stu-id="f2833-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="f2833-149">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="f2833-149">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f2833-150">1.0</span><span class="sxs-lookup"><span data-stu-id="f2833-150">1.0</span></span>|
|[<span data-ttu-id="f2833-151">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f2833-151">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f2833-152">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="f2833-152">Compose or read</span></span>|
####  <a name="sourceproperty-string"></a><span data-ttu-id="f2833-153">SourceProperty :String</span><span class="sxs-lookup"><span data-stu-id="f2833-153">SourceProperty :String</span></span>

<span data-ttu-id="f2833-154">Указывает источник данных, возвращаемых вызванным методом.</span><span class="sxs-lookup"><span data-stu-id="f2833-154">Specifies the source of the data returned by the invoked method.</span></span>

##### <a name="type"></a><span data-ttu-id="f2833-155">Тип:</span><span class="sxs-lookup"><span data-stu-id="f2833-155">Type:</span></span>

*   <span data-ttu-id="f2833-156">String</span><span class="sxs-lookup"><span data-stu-id="f2833-156">String</span></span>

##### <a name="properties"></a><span data-ttu-id="f2833-157">Свойства:</span><span class="sxs-lookup"><span data-stu-id="f2833-157">Properties:</span></span>

|<span data-ttu-id="f2833-158">Имя</span><span class="sxs-lookup"><span data-stu-id="f2833-158">Name</span></span>| <span data-ttu-id="f2833-159">Тип</span><span class="sxs-lookup"><span data-stu-id="f2833-159">Type</span></span>| <span data-ttu-id="f2833-160">Описание</span><span class="sxs-lookup"><span data-stu-id="f2833-160">Description</span></span>|
|---|---|---|
|`Body`| <span data-ttu-id="f2833-161">String</span><span class="sxs-lookup"><span data-stu-id="f2833-161">String</span></span>|<span data-ttu-id="f2833-162">Источник данных — текст сообщения.</span><span class="sxs-lookup"><span data-stu-id="f2833-162">The source of the data is from the body of a message.</span></span>|
|`Subject`| <span data-ttu-id="f2833-163">String</span><span class="sxs-lookup"><span data-stu-id="f2833-163">String</span></span>|<span data-ttu-id="f2833-164">Источник данных — тема сообщения.</span><span class="sxs-lookup"><span data-stu-id="f2833-164">The source of the data is from the subject of a message.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f2833-165">Обязательные элементы</span><span class="sxs-lookup"><span data-stu-id="f2833-165">Requirements</span></span>

|<span data-ttu-id="f2833-166">Требование</span><span class="sxs-lookup"><span data-stu-id="f2833-166">Requirement</span></span>| <span data-ttu-id="f2833-167">Значение</span><span class="sxs-lookup"><span data-stu-id="f2833-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="f2833-168">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="f2833-168">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f2833-169">1.0</span><span class="sxs-lookup"><span data-stu-id="f2833-169">1.0</span></span>|
|[<span data-ttu-id="f2833-170">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f2833-170">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="f2833-171">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="f2833-171">Compose or read</span></span>|