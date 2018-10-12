
# <a name="userprofile"></a><span data-ttu-id="225e5-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="225e5-101">userProfile</span></span>

### <span data-ttu-id="225e5-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="225e5-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="225e5-104">Требования</span><span class="sxs-lookup"><span data-stu-id="225e5-104">Requirements</span></span>

|<span data-ttu-id="225e5-105">Требование</span><span class="sxs-lookup"><span data-stu-id="225e5-105">Requirement</span></span>| <span data-ttu-id="225e5-106">Значение</span><span class="sxs-lookup"><span data-stu-id="225e5-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="225e5-107">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="225e5-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="225e5-108">1.0</span><span class="sxs-lookup"><span data-stu-id="225e5-108">1.0</span></span>|
|[<span data-ttu-id="225e5-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="225e5-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="225e5-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="225e5-110">ReadItem</span></span>|
|[<span data-ttu-id="225e5-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="225e5-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="225e5-112">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="225e5-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="225e5-113">Члены</span><span class="sxs-lookup"><span data-stu-id="225e5-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="225e5-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="225e5-114">displayName :String</span></span>

<span data-ttu-id="225e5-115">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="225e5-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="225e5-116">Тип:</span><span class="sxs-lookup"><span data-stu-id="225e5-116">Type:</span></span>

*   <span data-ttu-id="225e5-117">Строка</span><span class="sxs-lookup"><span data-stu-id="225e5-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="225e5-118">Требования</span><span class="sxs-lookup"><span data-stu-id="225e5-118">Requirements</span></span>

|<span data-ttu-id="225e5-119">Требование</span><span class="sxs-lookup"><span data-stu-id="225e5-119">Requirement</span></span>| <span data-ttu-id="225e5-120">Значение</span><span class="sxs-lookup"><span data-stu-id="225e5-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="225e5-121">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="225e5-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="225e5-122">1.0</span><span class="sxs-lookup"><span data-stu-id="225e5-122">1.0</span></span>|
|[<span data-ttu-id="225e5-123">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="225e5-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="225e5-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="225e5-124">ReadItem</span></span>|
|[<span data-ttu-id="225e5-125">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="225e5-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="225e5-126">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="225e5-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="225e5-127">Пример</span><span class="sxs-lookup"><span data-stu-id="225e5-127">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="225e5-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="225e5-128">emailAddress :String</span></span>

<span data-ttu-id="225e5-129">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="225e5-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="225e5-130">Тип:</span><span class="sxs-lookup"><span data-stu-id="225e5-130">Type:</span></span>

*   <span data-ttu-id="225e5-131">Строка</span><span class="sxs-lookup"><span data-stu-id="225e5-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="225e5-132">Требования</span><span class="sxs-lookup"><span data-stu-id="225e5-132">Requirements</span></span>

|<span data-ttu-id="225e5-133">Требование</span><span class="sxs-lookup"><span data-stu-id="225e5-133">Requirement</span></span>| <span data-ttu-id="225e5-134">Значение</span><span class="sxs-lookup"><span data-stu-id="225e5-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="225e5-135">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="225e5-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="225e5-136">1.0</span><span class="sxs-lookup"><span data-stu-id="225e5-136">1.0</span></span>|
|[<span data-ttu-id="225e5-137">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="225e5-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="225e5-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="225e5-138">ReadItem</span></span>|
|[<span data-ttu-id="225e5-139">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="225e5-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="225e5-140">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="225e5-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="225e5-141">Пример</span><span class="sxs-lookup"><span data-stu-id="225e5-141">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="225e5-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="225e5-142">timeZone :String</span></span>

<span data-ttu-id="225e5-143">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="225e5-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="225e5-144">Тип:</span><span class="sxs-lookup"><span data-stu-id="225e5-144">Type:</span></span>

*   <span data-ttu-id="225e5-145">Строка</span><span class="sxs-lookup"><span data-stu-id="225e5-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="225e5-146">Требования</span><span class="sxs-lookup"><span data-stu-id="225e5-146">Requirements</span></span>

|<span data-ttu-id="225e5-147">Требование</span><span class="sxs-lookup"><span data-stu-id="225e5-147">Requirement</span></span>| <span data-ttu-id="225e5-148">Значение</span><span class="sxs-lookup"><span data-stu-id="225e5-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="225e5-149">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="225e5-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="225e5-150">1.0</span><span class="sxs-lookup"><span data-stu-id="225e5-150">1.0</span></span>|
|[<span data-ttu-id="225e5-151">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="225e5-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="225e5-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="225e5-152">ReadItem</span></span>|
|[<span data-ttu-id="225e5-153">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="225e5-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="225e5-154">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="225e5-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="225e5-155">Пример</span><span class="sxs-lookup"><span data-stu-id="225e5-155">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```