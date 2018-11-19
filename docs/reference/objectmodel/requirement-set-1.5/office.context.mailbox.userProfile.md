# <a name="userprofile"></a><span data-ttu-id="a56df-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="a56df-101">userProfile</span></span>

### <span data-ttu-id="a56df-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="a56df-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="a56df-104">Требования</span><span class="sxs-lookup"><span data-stu-id="a56df-104">Requirements</span></span>

|<span data-ttu-id="a56df-105">Requirement</span><span class="sxs-lookup"><span data-stu-id="a56df-105">Requirement</span></span>| <span data-ttu-id="a56df-106">Значение</span><span class="sxs-lookup"><span data-stu-id="a56df-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="a56df-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a56df-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a56df-108">1.0</span><span class="sxs-lookup"><span data-stu-id="a56df-108">1.0</span></span>|
|[<span data-ttu-id="a56df-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a56df-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a56df-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a56df-110">ReadItem</span></span>|
|[<span data-ttu-id="a56df-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a56df-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a56df-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a56df-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="a56df-113">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="a56df-113">Members and methods</span></span>

| <span data-ttu-id="a56df-114">Член</span><span class="sxs-lookup"><span data-stu-id="a56df-114">Member</span></span> | <span data-ttu-id="a56df-115">Тип</span><span class="sxs-lookup"><span data-stu-id="a56df-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="a56df-116">displayName</span><span class="sxs-lookup"><span data-stu-id="a56df-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="a56df-117">Элемент</span><span class="sxs-lookup"><span data-stu-id="a56df-117">Member</span></span> |
| [<span data-ttu-id="a56df-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="a56df-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="a56df-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="a56df-119">Member</span></span> |
| [<span data-ttu-id="a56df-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="a56df-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="a56df-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="a56df-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="a56df-122">Элементы</span><span class="sxs-lookup"><span data-stu-id="a56df-122">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="a56df-123">displayName :String</span><span class="sxs-lookup"><span data-stu-id="a56df-123">displayName :String</span></span>

<span data-ttu-id="a56df-124">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="a56df-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="a56df-125">Тип:</span><span class="sxs-lookup"><span data-stu-id="a56df-125">Type:</span></span>

*   <span data-ttu-id="a56df-126">String</span><span class="sxs-lookup"><span data-stu-id="a56df-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a56df-127">Требования</span><span class="sxs-lookup"><span data-stu-id="a56df-127">Requirements</span></span>

|<span data-ttu-id="a56df-128">Requirement</span><span class="sxs-lookup"><span data-stu-id="a56df-128">Requirement</span></span>| <span data-ttu-id="a56df-129">Значение</span><span class="sxs-lookup"><span data-stu-id="a56df-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="a56df-130">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a56df-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a56df-131">1.0</span><span class="sxs-lookup"><span data-stu-id="a56df-131">1.0</span></span>|
|[<span data-ttu-id="a56df-132">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a56df-132">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a56df-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a56df-133">ReadItem</span></span>|
|[<span data-ttu-id="a56df-134">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a56df-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a56df-135">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a56df-135">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a56df-136">Пример</span><span class="sxs-lookup"><span data-stu-id="a56df-136">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="a56df-137">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="a56df-137">emailAddress :String</span></span>

<span data-ttu-id="a56df-138">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="a56df-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="a56df-139">Тип:</span><span class="sxs-lookup"><span data-stu-id="a56df-139">Type:</span></span>

*   <span data-ttu-id="a56df-140">String</span><span class="sxs-lookup"><span data-stu-id="a56df-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a56df-141">Требования</span><span class="sxs-lookup"><span data-stu-id="a56df-141">Requirements</span></span>

|<span data-ttu-id="a56df-142">Requirement</span><span class="sxs-lookup"><span data-stu-id="a56df-142">Requirement</span></span>| <span data-ttu-id="a56df-143">Значение</span><span class="sxs-lookup"><span data-stu-id="a56df-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="a56df-144">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a56df-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a56df-145">1.0</span><span class="sxs-lookup"><span data-stu-id="a56df-145">1.0</span></span>|
|[<span data-ttu-id="a56df-146">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a56df-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a56df-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a56df-147">ReadItem</span></span>|
|[<span data-ttu-id="a56df-148">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a56df-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a56df-149">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a56df-149">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a56df-150">Пример</span><span class="sxs-lookup"><span data-stu-id="a56df-150">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="a56df-151">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="a56df-151">timeZone :String</span></span>

<span data-ttu-id="a56df-152">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="a56df-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="a56df-153">Тип:</span><span class="sxs-lookup"><span data-stu-id="a56df-153">Type:</span></span>

*   <span data-ttu-id="a56df-154">String</span><span class="sxs-lookup"><span data-stu-id="a56df-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a56df-155">Требования</span><span class="sxs-lookup"><span data-stu-id="a56df-155">Requirements</span></span>

|<span data-ttu-id="a56df-156">Requirement</span><span class="sxs-lookup"><span data-stu-id="a56df-156">Requirement</span></span>| <span data-ttu-id="a56df-157">Значение</span><span class="sxs-lookup"><span data-stu-id="a56df-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="a56df-158">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a56df-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a56df-159">1.0</span><span class="sxs-lookup"><span data-stu-id="a56df-159">1.0</span></span>|
|[<span data-ttu-id="a56df-160">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a56df-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a56df-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a56df-161">ReadItem</span></span>|
|[<span data-ttu-id="a56df-162">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a56df-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a56df-163">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a56df-163">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a56df-164">Пример</span><span class="sxs-lookup"><span data-stu-id="a56df-164">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```