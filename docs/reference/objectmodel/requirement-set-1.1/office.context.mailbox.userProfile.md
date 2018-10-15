
# <a name="userprofile"></a><span data-ttu-id="0a7a9-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="0a7a9-101">userProfile</span></span>

### <span data-ttu-id="0a7a9-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="0a7a9-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="0a7a9-104">Требования</span><span class="sxs-lookup"><span data-stu-id="0a7a9-104">Requirements</span></span>

|<span data-ttu-id="0a7a9-105">Требование</span><span class="sxs-lookup"><span data-stu-id="0a7a9-105">Requirement</span></span>| <span data-ttu-id="0a7a9-106">Значение</span><span class="sxs-lookup"><span data-stu-id="0a7a9-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="0a7a9-107">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0a7a9-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0a7a9-108">1.0</span><span class="sxs-lookup"><span data-stu-id="0a7a9-108">1.0</span></span>|
|[<span data-ttu-id="0a7a9-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0a7a9-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0a7a9-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0a7a9-110">ReadItem</span></span>|
|[<span data-ttu-id="0a7a9-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0a7a9-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0a7a9-112">Compose или read</span><span class="sxs-lookup"><span data-stu-id="0a7a9-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="0a7a9-113">Члены</span><span class="sxs-lookup"><span data-stu-id="0a7a9-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="0a7a9-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="0a7a9-114">displayName :String</span></span>

<span data-ttu-id="0a7a9-115">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="0a7a9-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="0a7a9-116">Тип:</span><span class="sxs-lookup"><span data-stu-id="0a7a9-116">Type:</span></span>

*   <span data-ttu-id="0a7a9-117">String</span><span class="sxs-lookup"><span data-stu-id="0a7a9-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0a7a9-118">Требования</span><span class="sxs-lookup"><span data-stu-id="0a7a9-118">Requirements</span></span>

|<span data-ttu-id="0a7a9-119">Требование</span><span class="sxs-lookup"><span data-stu-id="0a7a9-119">Requirement</span></span>| <span data-ttu-id="0a7a9-120">Значение</span><span class="sxs-lookup"><span data-stu-id="0a7a9-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="0a7a9-121">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0a7a9-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0a7a9-122">1.0</span><span class="sxs-lookup"><span data-stu-id="0a7a9-122">1.0</span></span>|
|[<span data-ttu-id="0a7a9-123">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0a7a9-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0a7a9-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0a7a9-124">ReadItem</span></span>|
|[<span data-ttu-id="0a7a9-125">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0a7a9-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0a7a9-126">Compose или read</span><span class="sxs-lookup"><span data-stu-id="0a7a9-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="0a7a9-127">Пример</span><span class="sxs-lookup"><span data-stu-id="0a7a9-127">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="0a7a9-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="0a7a9-128">emailAddress :String</span></span>

<span data-ttu-id="0a7a9-129">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="0a7a9-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="0a7a9-130">Тип:</span><span class="sxs-lookup"><span data-stu-id="0a7a9-130">Type:</span></span>

*   <span data-ttu-id="0a7a9-131">String</span><span class="sxs-lookup"><span data-stu-id="0a7a9-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0a7a9-132">Требования</span><span class="sxs-lookup"><span data-stu-id="0a7a9-132">Requirements</span></span>

|<span data-ttu-id="0a7a9-133">Требование</span><span class="sxs-lookup"><span data-stu-id="0a7a9-133">Requirement</span></span>| <span data-ttu-id="0a7a9-134">Значение</span><span class="sxs-lookup"><span data-stu-id="0a7a9-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="0a7a9-135">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0a7a9-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0a7a9-136">1.0</span><span class="sxs-lookup"><span data-stu-id="0a7a9-136">1.0</span></span>|
|[<span data-ttu-id="0a7a9-137">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0a7a9-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0a7a9-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0a7a9-138">ReadItem</span></span>|
|[<span data-ttu-id="0a7a9-139">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0a7a9-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0a7a9-140">Compose или read</span><span class="sxs-lookup"><span data-stu-id="0a7a9-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="0a7a9-141">Пример</span><span class="sxs-lookup"><span data-stu-id="0a7a9-141">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="0a7a9-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="0a7a9-142">timeZone :String</span></span>

<span data-ttu-id="0a7a9-143">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="0a7a9-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="0a7a9-144">Тип:</span><span class="sxs-lookup"><span data-stu-id="0a7a9-144">Type:</span></span>

*   <span data-ttu-id="0a7a9-145">String</span><span class="sxs-lookup"><span data-stu-id="0a7a9-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0a7a9-146">Требования</span><span class="sxs-lookup"><span data-stu-id="0a7a9-146">Requirements</span></span>

|<span data-ttu-id="0a7a9-147">Требование</span><span class="sxs-lookup"><span data-stu-id="0a7a9-147">Requirement</span></span>| <span data-ttu-id="0a7a9-148">Значение</span><span class="sxs-lookup"><span data-stu-id="0a7a9-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="0a7a9-149">Версия минимального набора обязательных элементов для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="0a7a9-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0a7a9-150">1.0</span><span class="sxs-lookup"><span data-stu-id="0a7a9-150">1.0</span></span>|
|[<span data-ttu-id="0a7a9-151">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0a7a9-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0a7a9-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="0a7a9-152">ReadItem</span></span>|
|[<span data-ttu-id="0a7a9-153">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0a7a9-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="0a7a9-154">Compose или read</span><span class="sxs-lookup"><span data-stu-id="0a7a9-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="0a7a9-155">Пример</span><span class="sxs-lookup"><span data-stu-id="0a7a9-155">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```