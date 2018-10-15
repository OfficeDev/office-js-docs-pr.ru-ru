
# <a name="userprofile"></a><span data-ttu-id="c14d4-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="c14d4-101">userProfile</span></span>

### <span data-ttu-id="c14d4-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="c14d4-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="c14d4-104">Требования</span><span class="sxs-lookup"><span data-stu-id="c14d4-104">Requirements</span></span>

|<span data-ttu-id="c14d4-105">Требование</span><span class="sxs-lookup"><span data-stu-id="c14d4-105">Requirement</span></span>| <span data-ttu-id="c14d4-106">Значение</span><span class="sxs-lookup"><span data-stu-id="c14d4-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="c14d4-107">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="c14d4-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c14d4-108">1.0</span><span class="sxs-lookup"><span data-stu-id="c14d4-108">1.0</span></span>|
|[<span data-ttu-id="c14d4-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="c14d4-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c14d4-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c14d4-110">ReadItem</span></span>|
|[<span data-ttu-id="c14d4-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c14d4-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c14d4-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c14d4-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="c14d4-113">Члены</span><span class="sxs-lookup"><span data-stu-id="c14d4-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="c14d4-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="c14d4-114">displayName :String</span></span>

<span data-ttu-id="c14d4-115">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="c14d4-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="c14d4-116">Тип:</span><span class="sxs-lookup"><span data-stu-id="c14d4-116">Type:</span></span>

*   <span data-ttu-id="c14d4-117">String</span><span class="sxs-lookup"><span data-stu-id="c14d4-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c14d4-118">Требования</span><span class="sxs-lookup"><span data-stu-id="c14d4-118">Requirements</span></span>

|<span data-ttu-id="c14d4-119">Требование</span><span class="sxs-lookup"><span data-stu-id="c14d4-119">Requirement</span></span>| <span data-ttu-id="c14d4-120">Значение</span><span class="sxs-lookup"><span data-stu-id="c14d4-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="c14d4-121">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="c14d4-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c14d4-122">1.0</span><span class="sxs-lookup"><span data-stu-id="c14d4-122">1.0</span></span>|
|[<span data-ttu-id="c14d4-123">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="c14d4-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c14d4-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c14d4-124">ReadItem</span></span>|
|[<span data-ttu-id="c14d4-125">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c14d4-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c14d4-126">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c14d4-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c14d4-127">Пример</span><span class="sxs-lookup"><span data-stu-id="c14d4-127">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="c14d4-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="c14d4-128">emailAddress :String</span></span>

<span data-ttu-id="c14d4-129">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="c14d4-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="c14d4-130">Тип:</span><span class="sxs-lookup"><span data-stu-id="c14d4-130">Type:</span></span>

*   <span data-ttu-id="c14d4-131">String</span><span class="sxs-lookup"><span data-stu-id="c14d4-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c14d4-132">Требования</span><span class="sxs-lookup"><span data-stu-id="c14d4-132">Requirements</span></span>

|<span data-ttu-id="c14d4-133">Требование</span><span class="sxs-lookup"><span data-stu-id="c14d4-133">Requirement</span></span>| <span data-ttu-id="c14d4-134">Значение</span><span class="sxs-lookup"><span data-stu-id="c14d4-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="c14d4-135">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="c14d4-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c14d4-136">1.0</span><span class="sxs-lookup"><span data-stu-id="c14d4-136">1.0</span></span>|
|[<span data-ttu-id="c14d4-137">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="c14d4-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c14d4-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c14d4-138">ReadItem</span></span>|
|[<span data-ttu-id="c14d4-139">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c14d4-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c14d4-140">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c14d4-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c14d4-141">Пример</span><span class="sxs-lookup"><span data-stu-id="c14d4-141">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="c14d4-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="c14d4-142">timeZone :String</span></span>

<span data-ttu-id="c14d4-143">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="c14d4-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="c14d4-144">Тип:</span><span class="sxs-lookup"><span data-stu-id="c14d4-144">Type:</span></span>

*   <span data-ttu-id="c14d4-145">String</span><span class="sxs-lookup"><span data-stu-id="c14d4-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c14d4-146">Требования</span><span class="sxs-lookup"><span data-stu-id="c14d4-146">Requirements</span></span>

|<span data-ttu-id="c14d4-147">Требование</span><span class="sxs-lookup"><span data-stu-id="c14d4-147">Requirement</span></span>| <span data-ttu-id="c14d4-148">Значение</span><span class="sxs-lookup"><span data-stu-id="c14d4-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="c14d4-149">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="c14d4-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c14d4-150">1.0</span><span class="sxs-lookup"><span data-stu-id="c14d4-150">1.0</span></span>|
|[<span data-ttu-id="c14d4-151">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="c14d4-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c14d4-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c14d4-152">ReadItem</span></span>|
|[<span data-ttu-id="c14d4-153">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c14d4-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c14d4-154">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c14d4-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c14d4-155">Пример</span><span class="sxs-lookup"><span data-stu-id="c14d4-155">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```