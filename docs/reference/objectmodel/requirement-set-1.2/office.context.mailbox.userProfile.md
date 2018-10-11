
# <a name="userprofile"></a><span data-ttu-id="b2156-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="b2156-101">userProfile</span></span>

### <span data-ttu-id="b2156-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="b2156-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="b2156-104">Требования</span><span class="sxs-lookup"><span data-stu-id="b2156-104">Requirements</span></span>

|<span data-ttu-id="b2156-105">Требование</span><span class="sxs-lookup"><span data-stu-id="b2156-105">Requirement</span></span>| <span data-ttu-id="b2156-106">Значение</span><span class="sxs-lookup"><span data-stu-id="b2156-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="b2156-107">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="b2156-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b2156-108">1.0</span><span class="sxs-lookup"><span data-stu-id="b2156-108">1.0</span></span>|
|[<span data-ttu-id="b2156-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b2156-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b2156-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b2156-110">ReadItem</span></span>|
|[<span data-ttu-id="b2156-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b2156-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b2156-112">Создание или чтение​</span><span class="sxs-lookup"><span data-stu-id="b2156-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="b2156-113">Члены</span><span class="sxs-lookup"><span data-stu-id="b2156-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="b2156-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="b2156-114">displayName :String</span></span>

<span data-ttu-id="b2156-115">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="b2156-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="b2156-116">Тип:</span><span class="sxs-lookup"><span data-stu-id="b2156-116">Type:</span></span>

*   <span data-ttu-id="b2156-117">String</span><span class="sxs-lookup"><span data-stu-id="b2156-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b2156-118">Требования</span><span class="sxs-lookup"><span data-stu-id="b2156-118">Requirements</span></span>

|<span data-ttu-id="b2156-119">Требование</span><span class="sxs-lookup"><span data-stu-id="b2156-119">Requirement</span></span>| <span data-ttu-id="b2156-120">Значение</span><span class="sxs-lookup"><span data-stu-id="b2156-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="b2156-121">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="b2156-121">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b2156-122">1.0</span><span class="sxs-lookup"><span data-stu-id="b2156-122">1.0</span></span>|
|[<span data-ttu-id="b2156-123">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b2156-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b2156-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b2156-124">ReadItem</span></span>|
|[<span data-ttu-id="b2156-125">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b2156-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b2156-126">Создание или чтение​</span><span class="sxs-lookup"><span data-stu-id="b2156-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b2156-127">Пример</span><span class="sxs-lookup"><span data-stu-id="b2156-127">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="b2156-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="b2156-128">emailAddress :String</span></span>

<span data-ttu-id="b2156-129">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="b2156-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="b2156-130">Тип:</span><span class="sxs-lookup"><span data-stu-id="b2156-130">Type:</span></span>

*   <span data-ttu-id="b2156-131">String</span><span class="sxs-lookup"><span data-stu-id="b2156-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b2156-132">Требования</span><span class="sxs-lookup"><span data-stu-id="b2156-132">Requirements</span></span>

|<span data-ttu-id="b2156-133">Требование</span><span class="sxs-lookup"><span data-stu-id="b2156-133">Requirement</span></span>| <span data-ttu-id="b2156-134">Значение</span><span class="sxs-lookup"><span data-stu-id="b2156-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="b2156-135">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="b2156-135">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b2156-136">1.0</span><span class="sxs-lookup"><span data-stu-id="b2156-136">1.0</span></span>|
|[<span data-ttu-id="b2156-137">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b2156-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b2156-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b2156-138">ReadItem</span></span>|
|[<span data-ttu-id="b2156-139">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b2156-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b2156-140">Создание или чтение​</span><span class="sxs-lookup"><span data-stu-id="b2156-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b2156-141">Пример</span><span class="sxs-lookup"><span data-stu-id="b2156-141">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="b2156-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="b2156-142">timeZone :String</span></span>

<span data-ttu-id="b2156-143">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="b2156-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="b2156-144">Тип:</span><span class="sxs-lookup"><span data-stu-id="b2156-144">Type:</span></span>

*   <span data-ttu-id="b2156-145">String</span><span class="sxs-lookup"><span data-stu-id="b2156-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b2156-146">Требования</span><span class="sxs-lookup"><span data-stu-id="b2156-146">Requirements</span></span>

|<span data-ttu-id="b2156-147">Требование</span><span class="sxs-lookup"><span data-stu-id="b2156-147">Requirement</span></span>| <span data-ttu-id="b2156-148">Значение</span><span class="sxs-lookup"><span data-stu-id="b2156-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="b2156-149">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="b2156-149">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="b2156-150">1.0</span><span class="sxs-lookup"><span data-stu-id="b2156-150">1.0</span></span>|
|[<span data-ttu-id="b2156-151">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b2156-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="b2156-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="b2156-152">ReadItem</span></span>|
|[<span data-ttu-id="b2156-153">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b2156-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="b2156-154">Создание или чтение​</span><span class="sxs-lookup"><span data-stu-id="b2156-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="b2156-155">Пример</span><span class="sxs-lookup"><span data-stu-id="b2156-155">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```