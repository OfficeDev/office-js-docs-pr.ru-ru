
# <a name="userprofile"></a><span data-ttu-id="3c57c-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="3c57c-101">userProfile</span></span>

### <span data-ttu-id="3c57c-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="3c57c-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="3c57c-104">Требования</span><span class="sxs-lookup"><span data-stu-id="3c57c-104">Requirements</span></span>

|<span data-ttu-id="3c57c-105">Требование</span><span class="sxs-lookup"><span data-stu-id="3c57c-105">Requirement</span></span>| <span data-ttu-id="3c57c-106">Значение</span><span class="sxs-lookup"><span data-stu-id="3c57c-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="3c57c-107">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="3c57c-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3c57c-108">1.0</span><span class="sxs-lookup"><span data-stu-id="3c57c-108">1.0</span></span>|
|[<span data-ttu-id="3c57c-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="3c57c-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3c57c-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3c57c-110">ReadItem</span></span>|
|[<span data-ttu-id="3c57c-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3c57c-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3c57c-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3c57c-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="3c57c-113">Члены</span><span class="sxs-lookup"><span data-stu-id="3c57c-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="3c57c-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="3c57c-114">displayName :String</span></span>

<span data-ttu-id="3c57c-115">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="3c57c-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="3c57c-116">Тип:</span><span class="sxs-lookup"><span data-stu-id="3c57c-116">Type:</span></span>

*   <span data-ttu-id="3c57c-117">String</span><span class="sxs-lookup"><span data-stu-id="3c57c-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3c57c-118">Требования</span><span class="sxs-lookup"><span data-stu-id="3c57c-118">Requirements</span></span>

|<span data-ttu-id="3c57c-119">Требование</span><span class="sxs-lookup"><span data-stu-id="3c57c-119">Requirement</span></span>| <span data-ttu-id="3c57c-120">Значение</span><span class="sxs-lookup"><span data-stu-id="3c57c-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="3c57c-121">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="3c57c-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3c57c-122">1.0</span><span class="sxs-lookup"><span data-stu-id="3c57c-122">1.0</span></span>|
|[<span data-ttu-id="3c57c-123">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="3c57c-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3c57c-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3c57c-124">ReadItem</span></span>|
|[<span data-ttu-id="3c57c-125">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3c57c-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3c57c-126">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3c57c-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="3c57c-127">Пример</span><span class="sxs-lookup"><span data-stu-id="3c57c-127">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="3c57c-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="3c57c-128">emailAddress :String</span></span>

<span data-ttu-id="3c57c-129">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="3c57c-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="3c57c-130">Тип:</span><span class="sxs-lookup"><span data-stu-id="3c57c-130">Type:</span></span>

*   <span data-ttu-id="3c57c-131">String</span><span class="sxs-lookup"><span data-stu-id="3c57c-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3c57c-132">Требования</span><span class="sxs-lookup"><span data-stu-id="3c57c-132">Requirements</span></span>

|<span data-ttu-id="3c57c-133">Требование</span><span class="sxs-lookup"><span data-stu-id="3c57c-133">Requirement</span></span>| <span data-ttu-id="3c57c-134">Значение</span><span class="sxs-lookup"><span data-stu-id="3c57c-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="3c57c-135">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="3c57c-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3c57c-136">1.0</span><span class="sxs-lookup"><span data-stu-id="3c57c-136">1.0</span></span>|
|[<span data-ttu-id="3c57c-137">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="3c57c-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3c57c-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3c57c-138">ReadItem</span></span>|
|[<span data-ttu-id="3c57c-139">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3c57c-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3c57c-140">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3c57c-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="3c57c-141">Пример</span><span class="sxs-lookup"><span data-stu-id="3c57c-141">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="3c57c-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="3c57c-142">timeZone :String</span></span>

<span data-ttu-id="3c57c-143">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="3c57c-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="3c57c-144">Тип:</span><span class="sxs-lookup"><span data-stu-id="3c57c-144">Type:</span></span>

*   <span data-ttu-id="3c57c-145">String</span><span class="sxs-lookup"><span data-stu-id="3c57c-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="3c57c-146">Требования</span><span class="sxs-lookup"><span data-stu-id="3c57c-146">Requirements</span></span>

|<span data-ttu-id="3c57c-147">Требование</span><span class="sxs-lookup"><span data-stu-id="3c57c-147">Requirement</span></span>| <span data-ttu-id="3c57c-148">Значение</span><span class="sxs-lookup"><span data-stu-id="3c57c-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="3c57c-149">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="3c57c-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="3c57c-150">1.0</span><span class="sxs-lookup"><span data-stu-id="3c57c-150">1.0</span></span>|
|[<span data-ttu-id="3c57c-151">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="3c57c-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="3c57c-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="3c57c-152">ReadItem</span></span>|
|[<span data-ttu-id="3c57c-153">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="3c57c-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="3c57c-154">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="3c57c-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="3c57c-155">Пример</span><span class="sxs-lookup"><span data-stu-id="3c57c-155">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```