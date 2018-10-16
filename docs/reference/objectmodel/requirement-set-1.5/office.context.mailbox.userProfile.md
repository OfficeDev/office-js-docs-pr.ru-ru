# <a name="userprofile"></a><span data-ttu-id="7c23f-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="7c23f-101">userProfile</span></span>

### <span data-ttu-id="7c23f-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="7c23f-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="7c23f-104">Требования</span><span class="sxs-lookup"><span data-stu-id="7c23f-104">Requirements</span></span>

|<span data-ttu-id="7c23f-105">Требование</span><span class="sxs-lookup"><span data-stu-id="7c23f-105">Requirement</span></span>| <span data-ttu-id="7c23f-106">Значение</span><span class="sxs-lookup"><span data-stu-id="7c23f-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c23f-107">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="7c23f-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c23f-108">1.0</span><span class="sxs-lookup"><span data-stu-id="7c23f-108">1.0</span></span>|
|[<span data-ttu-id="7c23f-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7c23f-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7c23f-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7c23f-110">ReadItem</span></span>|
|[<span data-ttu-id="7c23f-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7c23f-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7c23f-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="7c23f-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="7c23f-113">Члены и методы</span><span class="sxs-lookup"><span data-stu-id="7c23f-113">Members and methods</span></span>

| <span data-ttu-id="7c23f-114">Член</span><span class="sxs-lookup"><span data-stu-id="7c23f-114">Member</span></span> | <span data-ttu-id="7c23f-115">Тип</span><span class="sxs-lookup"><span data-stu-id="7c23f-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="7c23f-116">displayName</span><span class="sxs-lookup"><span data-stu-id="7c23f-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="7c23f-117">Член</span><span class="sxs-lookup"><span data-stu-id="7c23f-117">Member</span></span> |
| [<span data-ttu-id="7c23f-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="7c23f-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="7c23f-119">Член</span><span class="sxs-lookup"><span data-stu-id="7c23f-119">Member</span></span> |
| [<span data-ttu-id="7c23f-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="7c23f-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="7c23f-121">Член</span><span class="sxs-lookup"><span data-stu-id="7c23f-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="7c23f-122">Члены</span><span class="sxs-lookup"><span data-stu-id="7c23f-122">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="7c23f-123">displayName :String</span><span class="sxs-lookup"><span data-stu-id="7c23f-123">displayName :String</span></span>

<span data-ttu-id="7c23f-124">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="7c23f-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="7c23f-125">Тип:</span><span class="sxs-lookup"><span data-stu-id="7c23f-125">Type:</span></span>

*   <span data-ttu-id="7c23f-126">String</span><span class="sxs-lookup"><span data-stu-id="7c23f-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7c23f-127">Требования</span><span class="sxs-lookup"><span data-stu-id="7c23f-127">Requirements</span></span>

|<span data-ttu-id="7c23f-128">Требование</span><span class="sxs-lookup"><span data-stu-id="7c23f-128">Requirement</span></span>| <span data-ttu-id="7c23f-129">Значение</span><span class="sxs-lookup"><span data-stu-id="7c23f-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c23f-130">Версия минимального набора  требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="7c23f-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c23f-131">1.0</span><span class="sxs-lookup"><span data-stu-id="7c23f-131">1.0</span></span>|
|[<span data-ttu-id="7c23f-132">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7c23f-132">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7c23f-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7c23f-133">ReadItem</span></span>|
|[<span data-ttu-id="7c23f-134">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7c23f-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7c23f-135">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="7c23f-135">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="7c23f-136">Пример</span><span class="sxs-lookup"><span data-stu-id="7c23f-136">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="7c23f-137">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="7c23f-137">emailAddress :String</span></span>

<span data-ttu-id="7c23f-138">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="7c23f-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="7c23f-139">Тип:</span><span class="sxs-lookup"><span data-stu-id="7c23f-139">Type:</span></span>

*   <span data-ttu-id="7c23f-140">String</span><span class="sxs-lookup"><span data-stu-id="7c23f-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7c23f-141">Требования</span><span class="sxs-lookup"><span data-stu-id="7c23f-141">Requirements</span></span>

|<span data-ttu-id="7c23f-142">Требование</span><span class="sxs-lookup"><span data-stu-id="7c23f-142">Requirement</span></span>| <span data-ttu-id="7c23f-143">Значение</span><span class="sxs-lookup"><span data-stu-id="7c23f-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c23f-144">Версия минимального набора  требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="7c23f-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c23f-145">1.0</span><span class="sxs-lookup"><span data-stu-id="7c23f-145">1.0</span></span>|
|[<span data-ttu-id="7c23f-146">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7c23f-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7c23f-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7c23f-147">ReadItem</span></span>|
|[<span data-ttu-id="7c23f-148">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7c23f-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7c23f-149">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="7c23f-149">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="7c23f-150">Пример</span><span class="sxs-lookup"><span data-stu-id="7c23f-150">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="7c23f-151">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="7c23f-151">timeZone :String</span></span>

<span data-ttu-id="7c23f-152">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="7c23f-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="7c23f-153">Тип:</span><span class="sxs-lookup"><span data-stu-id="7c23f-153">Type:</span></span>

*   <span data-ttu-id="7c23f-154">String</span><span class="sxs-lookup"><span data-stu-id="7c23f-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7c23f-155">Требования</span><span class="sxs-lookup"><span data-stu-id="7c23f-155">Requirements</span></span>

|<span data-ttu-id="7c23f-156">Требование</span><span class="sxs-lookup"><span data-stu-id="7c23f-156">Requirement</span></span>| <span data-ttu-id="7c23f-157">Значение</span><span class="sxs-lookup"><span data-stu-id="7c23f-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="7c23f-158">Версия минимального набора  требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="7c23f-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="7c23f-159">1.0</span><span class="sxs-lookup"><span data-stu-id="7c23f-159">1.0</span></span>|
|[<span data-ttu-id="7c23f-160">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7c23f-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="7c23f-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="7c23f-161">ReadItem</span></span>|
|[<span data-ttu-id="7c23f-162">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7c23f-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="7c23f-163">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="7c23f-163">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="7c23f-164">Пример</span><span class="sxs-lookup"><span data-stu-id="7c23f-164">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```