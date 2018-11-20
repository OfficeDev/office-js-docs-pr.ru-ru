
# <a name="userprofile"></a><span data-ttu-id="c0216-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="c0216-101">userProfile</span></span>

### <span data-ttu-id="c0216-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="c0216-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="c0216-104">Требования</span><span class="sxs-lookup"><span data-stu-id="c0216-104">Requirements</span></span>

|<span data-ttu-id="c0216-105">Requirement</span><span class="sxs-lookup"><span data-stu-id="c0216-105">Requirement</span></span>| <span data-ttu-id="c0216-106">Значение</span><span class="sxs-lookup"><span data-stu-id="c0216-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="c0216-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c0216-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c0216-108">1.0</span><span class="sxs-lookup"><span data-stu-id="c0216-108">1.0</span></span>|
|[<span data-ttu-id="c0216-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="c0216-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c0216-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c0216-110">ReadItem</span></span>|
|[<span data-ttu-id="c0216-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c0216-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c0216-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c0216-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="c0216-113">Элементы</span><span class="sxs-lookup"><span data-stu-id="c0216-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="c0216-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="c0216-114">displayName :String</span></span>

<span data-ttu-id="c0216-115">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="c0216-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="c0216-116">Тип:</span><span class="sxs-lookup"><span data-stu-id="c0216-116">Type:</span></span>

*   <span data-ttu-id="c0216-117">String</span><span class="sxs-lookup"><span data-stu-id="c0216-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c0216-118">Требования</span><span class="sxs-lookup"><span data-stu-id="c0216-118">Requirements</span></span>

|<span data-ttu-id="c0216-119">Requirement</span><span class="sxs-lookup"><span data-stu-id="c0216-119">Requirement</span></span>| <span data-ttu-id="c0216-120">Значение</span><span class="sxs-lookup"><span data-stu-id="c0216-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="c0216-121">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c0216-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c0216-122">1.0</span><span class="sxs-lookup"><span data-stu-id="c0216-122">1.0</span></span>|
|[<span data-ttu-id="c0216-123">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="c0216-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c0216-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c0216-124">ReadItem</span></span>|
|[<span data-ttu-id="c0216-125">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c0216-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c0216-126">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c0216-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c0216-127">Пример</span><span class="sxs-lookup"><span data-stu-id="c0216-127">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="c0216-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="c0216-128">emailAddress :String</span></span>

<span data-ttu-id="c0216-129">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="c0216-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="c0216-130">Тип:</span><span class="sxs-lookup"><span data-stu-id="c0216-130">Type:</span></span>

*   <span data-ttu-id="c0216-131">String</span><span class="sxs-lookup"><span data-stu-id="c0216-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c0216-132">Требования</span><span class="sxs-lookup"><span data-stu-id="c0216-132">Requirements</span></span>

|<span data-ttu-id="c0216-133">Requirement</span><span class="sxs-lookup"><span data-stu-id="c0216-133">Requirement</span></span>| <span data-ttu-id="c0216-134">Значение</span><span class="sxs-lookup"><span data-stu-id="c0216-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="c0216-135">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c0216-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c0216-136">1.0</span><span class="sxs-lookup"><span data-stu-id="c0216-136">1.0</span></span>|
|[<span data-ttu-id="c0216-137">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="c0216-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c0216-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c0216-138">ReadItem</span></span>|
|[<span data-ttu-id="c0216-139">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c0216-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c0216-140">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c0216-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c0216-141">Пример</span><span class="sxs-lookup"><span data-stu-id="c0216-141">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="c0216-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="c0216-142">timeZone :String</span></span>

<span data-ttu-id="c0216-143">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="c0216-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="c0216-144">Тип:</span><span class="sxs-lookup"><span data-stu-id="c0216-144">Type:</span></span>

*   <span data-ttu-id="c0216-145">String</span><span class="sxs-lookup"><span data-stu-id="c0216-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c0216-146">Требования</span><span class="sxs-lookup"><span data-stu-id="c0216-146">Requirements</span></span>

|<span data-ttu-id="c0216-147">Requirement</span><span class="sxs-lookup"><span data-stu-id="c0216-147">Requirement</span></span>| <span data-ttu-id="c0216-148">Значение</span><span class="sxs-lookup"><span data-stu-id="c0216-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="c0216-149">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c0216-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c0216-150">1.0</span><span class="sxs-lookup"><span data-stu-id="c0216-150">1.0</span></span>|
|[<span data-ttu-id="c0216-151">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="c0216-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c0216-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c0216-152">ReadItem</span></span>|
|[<span data-ttu-id="c0216-153">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c0216-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c0216-154">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c0216-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c0216-155">Пример</span><span class="sxs-lookup"><span data-stu-id="c0216-155">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```