
# <a name="userprofile"></a><span data-ttu-id="533bd-101">userProfile</span><span class="sxs-lookup"><span data-stu-id="533bd-101">userProfile</span></span>

### <span data-ttu-id="533bd-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span><span class="sxs-lookup"><span data-stu-id="533bd-p101">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md). userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="533bd-104">Обязательные элементы</span><span class="sxs-lookup"><span data-stu-id="533bd-104">Requirements</span></span>

|<span data-ttu-id="533bd-105">Обязательный элемент</span><span class="sxs-lookup"><span data-stu-id="533bd-105">Requirement</span></span>| <span data-ttu-id="533bd-106">Значение</span><span class="sxs-lookup"><span data-stu-id="533bd-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="533bd-107">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="533bd-107">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="533bd-108">1.0</span><span class="sxs-lookup"><span data-stu-id="533bd-108">1.0</span></span>|
|[<span data-ttu-id="533bd-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="533bd-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="533bd-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="533bd-110">ReadItem</span></span>|
|[<span data-ttu-id="533bd-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="533bd-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="533bd-112">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="533bd-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="533bd-113">Члены</span><span class="sxs-lookup"><span data-stu-id="533bd-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="533bd-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="533bd-114">displayName :String</span></span>

<span data-ttu-id="533bd-115">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="533bd-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="533bd-116">Тип:</span><span class="sxs-lookup"><span data-stu-id="533bd-116">Type:</span></span>

*   <span data-ttu-id="533bd-117">String</span><span class="sxs-lookup"><span data-stu-id="533bd-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="533bd-118">Обязательные элементы</span><span class="sxs-lookup"><span data-stu-id="533bd-118">Requirements</span></span>

|<span data-ttu-id="533bd-119">Требование</span><span class="sxs-lookup"><span data-stu-id="533bd-119">Requirement</span></span>| <span data-ttu-id="533bd-120">Значение</span><span class="sxs-lookup"><span data-stu-id="533bd-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="533bd-121">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="533bd-121">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="533bd-122">1.0</span><span class="sxs-lookup"><span data-stu-id="533bd-122">1.0</span></span>|
|[<span data-ttu-id="533bd-123">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="533bd-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="533bd-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="533bd-124">ReadItem</span></span>|
|[<span data-ttu-id="533bd-125">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="533bd-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="533bd-126">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="533bd-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="533bd-127">Пример</span><span class="sxs-lookup"><span data-stu-id="533bd-127">Example</span></span>

```
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="533bd-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="533bd-128">emailAddress :String</span></span>

<span data-ttu-id="533bd-129">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="533bd-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="533bd-130">Тип:</span><span class="sxs-lookup"><span data-stu-id="533bd-130">Type:</span></span>

*   <span data-ttu-id="533bd-131">String</span><span class="sxs-lookup"><span data-stu-id="533bd-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="533bd-132">Обязательные элементы</span><span class="sxs-lookup"><span data-stu-id="533bd-132">Requirements</span></span>

|<span data-ttu-id="533bd-133">Требование</span><span class="sxs-lookup"><span data-stu-id="533bd-133">Requirement</span></span>| <span data-ttu-id="533bd-134">Значение</span><span class="sxs-lookup"><span data-stu-id="533bd-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="533bd-135">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="533bd-135">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="533bd-136">1.0</span><span class="sxs-lookup"><span data-stu-id="533bd-136">1.0</span></span>|
|[<span data-ttu-id="533bd-137">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="533bd-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="533bd-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="533bd-138">ReadItem</span></span>|
|[<span data-ttu-id="533bd-139">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="533bd-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="533bd-140">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="533bd-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="533bd-141">Пример</span><span class="sxs-lookup"><span data-stu-id="533bd-141">Example</span></span>

```
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="533bd-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="533bd-142">timeZone :String</span></span>

<span data-ttu-id="533bd-143">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="533bd-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="533bd-144">Тип:</span><span class="sxs-lookup"><span data-stu-id="533bd-144">Type:</span></span>

*   <span data-ttu-id="533bd-145">String</span><span class="sxs-lookup"><span data-stu-id="533bd-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="533bd-146">Обязательные элементы</span><span class="sxs-lookup"><span data-stu-id="533bd-146">Requirements</span></span>

|<span data-ttu-id="533bd-147">Требование</span><span class="sxs-lookup"><span data-stu-id="533bd-147">Requirement</span></span>| <span data-ttu-id="533bd-148">Значение</span><span class="sxs-lookup"><span data-stu-id="533bd-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="533bd-149">Версия минимального набора требований для почтового ящика (mailbox)</span><span class="sxs-lookup"><span data-stu-id="533bd-149">Minimum mailbox requirement set version</span></span>](/javascript/office/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="533bd-150">1.0</span><span class="sxs-lookup"><span data-stu-id="533bd-150">1.0</span></span>|
|[<span data-ttu-id="533bd-151">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="533bd-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="533bd-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="533bd-152">ReadItem</span></span>|
|[<span data-ttu-id="533bd-153">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="533bd-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="533bd-154">Compose (создание) или read (чтение)</span><span class="sxs-lookup"><span data-stu-id="533bd-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="533bd-155">Пример</span><span class="sxs-lookup"><span data-stu-id="533bd-155">Example</span></span>

```
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```