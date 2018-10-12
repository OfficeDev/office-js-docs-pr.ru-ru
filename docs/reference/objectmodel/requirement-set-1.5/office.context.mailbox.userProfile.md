# <a name="userprofile"></a><span data-ttu-id="a68fe-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="a68fe-101">userProfile</span></span>

### <span data-ttu-id="a68fe-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="a68fe-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="a68fe-104">Требования</span><span class="sxs-lookup"><span data-stu-id="a68fe-104">Requirements</span></span>

|<span data-ttu-id="a68fe-105">Требование</span><span class="sxs-lookup"><span data-stu-id="a68fe-105">Requirement</span></span>| <span data-ttu-id="a68fe-106">Значение</span><span class="sxs-lookup"><span data-stu-id="a68fe-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="a68fe-107">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="a68fe-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a68fe-108">1.0</span><span class="sxs-lookup"><span data-stu-id="a68fe-108">1.0</span></span>|
|[<span data-ttu-id="a68fe-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a68fe-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a68fe-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a68fe-110">ReadItem</span></span>|
|[<span data-ttu-id="a68fe-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a68fe-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a68fe-112">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="a68fe-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="a68fe-113">Члены и методы</span><span class="sxs-lookup"><span data-stu-id="a68fe-113">Members and methods</span></span>

| <span data-ttu-id="a68fe-114">Член</span><span class="sxs-lookup"><span data-stu-id="a68fe-114">Member</span></span> | <span data-ttu-id="a68fe-115">Тип</span><span class="sxs-lookup"><span data-stu-id="a68fe-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="a68fe-116">displayName</span><span class="sxs-lookup"><span data-stu-id="a68fe-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="a68fe-117">Член</span><span class="sxs-lookup"><span data-stu-id="a68fe-117">Member</span></span> |
| [<span data-ttu-id="a68fe-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="a68fe-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="a68fe-119">Член</span><span class="sxs-lookup"><span data-stu-id="a68fe-119">Member</span></span> |
| [<span data-ttu-id="a68fe-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="a68fe-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="a68fe-121">Член</span><span class="sxs-lookup"><span data-stu-id="a68fe-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="a68fe-122">Члены</span><span class="sxs-lookup"><span data-stu-id="a68fe-122">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="a68fe-123">displayName :String</span><span class="sxs-lookup"><span data-stu-id="a68fe-123">displayName :String</span></span>

<span data-ttu-id="a68fe-124">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="a68fe-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="a68fe-125">Тип:</span><span class="sxs-lookup"><span data-stu-id="a68fe-125">Type:</span></span>

*   <span data-ttu-id="a68fe-126">String</span><span class="sxs-lookup"><span data-stu-id="a68fe-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a68fe-127">Требования</span><span class="sxs-lookup"><span data-stu-id="a68fe-127">Requirements</span></span>

|<span data-ttu-id="a68fe-128">Требование</span><span class="sxs-lookup"><span data-stu-id="a68fe-128">Requirement</span></span>| <span data-ttu-id="a68fe-129">Значение</span><span class="sxs-lookup"><span data-stu-id="a68fe-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="a68fe-130">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="a68fe-130">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a68fe-131">1.0</span><span class="sxs-lookup"><span data-stu-id="a68fe-131">1.0</span></span>|
|[<span data-ttu-id="a68fe-132">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a68fe-132">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a68fe-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a68fe-133">ReadItem</span></span>|
|[<span data-ttu-id="a68fe-134">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a68fe-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a68fe-135">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="a68fe-135">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a68fe-136">Пример</span><span class="sxs-lookup"><span data-stu-id="a68fe-136">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="a68fe-137">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="a68fe-137">emailAddress :String</span></span>

<span data-ttu-id="a68fe-138">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="a68fe-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="a68fe-139">Тип:</span><span class="sxs-lookup"><span data-stu-id="a68fe-139">Type:</span></span>

*   <span data-ttu-id="a68fe-140">String</span><span class="sxs-lookup"><span data-stu-id="a68fe-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a68fe-141">Требования</span><span class="sxs-lookup"><span data-stu-id="a68fe-141">Requirements</span></span>

|<span data-ttu-id="a68fe-142">Требование</span><span class="sxs-lookup"><span data-stu-id="a68fe-142">Requirement</span></span>| <span data-ttu-id="a68fe-143">Значение</span><span class="sxs-lookup"><span data-stu-id="a68fe-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="a68fe-144">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="a68fe-144">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a68fe-145">1.0</span><span class="sxs-lookup"><span data-stu-id="a68fe-145">1.0</span></span>|
|[<span data-ttu-id="a68fe-146">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a68fe-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a68fe-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a68fe-147">ReadItem</span></span>|
|[<span data-ttu-id="a68fe-148">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a68fe-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a68fe-149">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="a68fe-149">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a68fe-150">Пример</span><span class="sxs-lookup"><span data-stu-id="a68fe-150">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="a68fe-151">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="a68fe-151">timeZone :String</span></span>

<span data-ttu-id="a68fe-152">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="a68fe-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="a68fe-153">Тип:</span><span class="sxs-lookup"><span data-stu-id="a68fe-153">Type:</span></span>

*   <span data-ttu-id="a68fe-154">String</span><span class="sxs-lookup"><span data-stu-id="a68fe-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a68fe-155">Требования</span><span class="sxs-lookup"><span data-stu-id="a68fe-155">Requirements</span></span>

|<span data-ttu-id="a68fe-156">Требование</span><span class="sxs-lookup"><span data-stu-id="a68fe-156">Requirement</span></span>| <span data-ttu-id="a68fe-157">Значение</span><span class="sxs-lookup"><span data-stu-id="a68fe-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="a68fe-158">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="a68fe-158">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a68fe-159">1.0</span><span class="sxs-lookup"><span data-stu-id="a68fe-159">1.0</span></span>|
|[<span data-ttu-id="a68fe-160">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a68fe-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a68fe-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a68fe-161">ReadItem</span></span>|
|[<span data-ttu-id="a68fe-162">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a68fe-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a68fe-163">Cоздание или чтение</span><span class="sxs-lookup"><span data-stu-id="a68fe-163">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a68fe-164">Пример</span><span class="sxs-lookup"><span data-stu-id="a68fe-164">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```